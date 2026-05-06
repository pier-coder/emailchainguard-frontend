/* ============================================================
   EmailChainGuard v3.3 — taskpane.js
   Fix:
   - Dominio dell'utente escluso dalla lista
   - Primo contatto: verifica PRIMA di salvare
   - Nuovo CC funzionante
   ============================================================ */
'use strict';

const CONFIG = {
  BACKEND_URL: 'https://emailchainguard-backend.onrender.com',
  ECG_API_KEY: 'ecg-dev-key-2024', // <-- sostituisci con la tua chiave
  MAX_KNOWN_EMAILS: 1000,
  MAX_CONVERSATIONS: 200,
  FETCH_TIMEOUT_MS: 15000,
};

const KEY_KNOWN_SENDERS = 'ecg_known_senders_v3';
const KEY_KNOWN_CC      = 'ecg_known_cc_v3';
const _cacheSenders     = new Set();
const _cacheCC          = {};

function _roamingOk() {
  try { return !!(Office?.context?.roamingSettings); }
  catch { return false; }
}

function _storageGet(key) {
  if (_roamingOk()) {
    try { const v = Office.context.roamingSettings.get(key); if (v) return v; } catch {}
  }
  try { return localStorage.getItem(key); } catch {}
  return null;
}

function _storageSet(key, value) {
  if (_roamingOk()) {
    try {
      Office.context.roamingSettings.set(key, value);
      Office.context.roamingSettings.saveAsync(() => {});
    } catch {}
  }
  try { localStorage.setItem(key, value); } catch {}
}

function loadKnownSenders() {
  const raw = _storageGet(KEY_KNOWN_SENDERS);
  const stored = raw ? new Set(JSON.parse(raw)) : new Set();
  _cacheSenders.forEach(e => stored.add(e));
  return stored;
}

function saveKnownSenders(emails) {
  emails.forEach(e => _cacheSenders.add(e));
  const all = loadKnownSenders();
  emails.forEach(e => all.add(e));
  let arr = Array.from(all);
  if (arr.length > CONFIG.MAX_KNOWN_EMAILS) arr = arr.slice(arr.length - CONFIG.MAX_KNOWN_EMAILS);
  _storageSet(KEY_KNOWN_SENDERS, JSON.stringify(arr));
}

function loadKnownCC() {
  const raw = _storageGet(KEY_KNOWN_CC);
  return raw ? JSON.parse(raw) : {};
}

function getKnownCCForConversation(convId) {
  if (_cacheCC[convId]) return _cacheCC[convId];
  const all = loadKnownCC();
  return new Set(all[convId] || []);
}

function saveKnownCCForConversation(convId, emails) {
  if (!_cacheCC[convId]) _cacheCC[convId] = new Set();
  emails.forEach(e => _cacheCC[convId].add(e));
  const all = loadKnownCC();
  const existing = new Set(all[convId] || []);
  emails.forEach(e => existing.add(e));
  all[convId] = Array.from(existing);
  const keys = Object.keys(all);
  if (keys.length > CONFIG.MAX_CONVERSATIONS) delete all[keys[0]];
  _storageSet(KEY_KNOWN_CC, JSON.stringify(all));
}

// ── Office init ───────────────────────────────────────────────────────
Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    runScan();
    if (Office.context.mailbox.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(
        Office.EventType.ItemChanged,
        () => { resetUI(); runScan(); }
      );
    }
  }
});

function resetUI() {
  ['banner-new','banner-danger','banner-link','banner-cc'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.classList.remove('visible');
  });
  document.getElementById('advice').classList.remove('visible');
  document.getElementById('domains-section').classList.remove('visible');
  document.getElementById('domain-list').innerHTML = '';
  document.getElementById('footer-addr').textContent = '—';
  setDot('');
}

// ── Scansione ─────────────────────────────────────────────────────────
async function runScan() {
  const item = Office.context.mailbox.item;
  if (!item) return;

  setDot('scanning');
  setScanBar(true);
  document.getElementById('idle-state').style.display = 'none';

  try {
    const myEmail   = Office.context.mailbox.userProfile?.emailAddress?.toLowerCase() || null;
    const fromAddr  = item.from?.emailAddress?.toLowerCase() || null;
    const toAddrs   = (item.to  || []).map(a => a.emailAddress?.toLowerCase()).filter(Boolean);
    const ccAddrs   = (item.cc  || []).map(a => a.emailAddress?.toLowerCase()).filter(Boolean);
    const convId    = item.conversationId || null;

    // Tutti gli indirizzi tranne il mio
    const allAddrs = [fromAddr, ...toAddrs, ...ccAddrs]
      .filter(Boolean)
      .filter(e => e !== myEmail);

    // ── 1. Primo contatto — verifica PRIMA di salvare ──────────────
    const knownSenders = loadKnownSenders();
    const isNewSender  = fromAddr && fromAddr !== myEmail && !knownSenders.has(fromAddr);

    // Salva DOPO la verifica
    saveKnownSenders(allAddrs);

    // ── 2. Nuovo CC in conversazione ───────────────────────────────
    let newCCAddrs = [];
    const ccAddrsFiltrati = ccAddrs.filter(e => e !== myEmail);
    if (convId && ccAddrsFiltrati.length > 0) {
      const knownCC = getKnownCCForConversation(convId);
      if (knownCC.size > 0) {
        newCCAddrs = ccAddrsFiltrati.filter(e => !knownCC.has(e));
      }
      saveKnownCCForConversation(convId, ccAddrsFiltrati);
    } else if (convId) {
      saveKnownCCForConversation(convId, ccAddrsFiltrati);
    }

    // ── 3. Domini da analizzare — escludo il mio ──────────────────
    const myDomain = myEmail ? myEmail.split('@')[1] : null;
    const domains  = [...new Set(
      allAddrs.map(a => a.split('@')[1]).filter(Boolean).filter(d => d !== myDomain)
    )];

    if (domains.length === 0) {
      // Nessun dominio esterno — mostra solo primo contatto se necessario
      renderResults({ overall_label: 'ok', suspect_count: 0, domains: [] },
        isNewSender ? fromAddr : null, newCCAddrs);
      return;
    }

    const result = await callBackend(domains, convId);
    if (!result || !Array.isArray(result.domains)) throw new Error('Risposta backend non valida');

    renderResults(result, isNewSender ? fromAddr : null, newCCAddrs);

  } catch (err) {
    setDot('');
    const p = document.getElementById('idle-state').querySelector('p');
    p.textContent = 'Errore: ' + (err.message || 'impossibile completare la scansione');
    document.getElementById('idle-state').style.display = 'flex';
  } finally {
    setScanBar(false);
  }
}

async function callBackend(domains, conversationId) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), CONFIG.FETCH_TIMEOUT_MS);
  try {
    const resp = await fetch(`${CONFIG.BACKEND_URL}/analyze`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'X-ECG-Key': CONFIG.ECG_API_KEY },
      body: JSON.stringify({ domains, conversation_id: conversationId }),
      signal: controller.signal,
    });
    if (!resp.ok) {
      const err = await resp.json().catch(() => ({}));
      throw new Error(err.detail || `Errore backend HTTP ${resp.status}`);
    }
    return resp.json();
  } catch (err) {
    if (err.name === 'AbortError') throw new Error('Analisi scaduta (timeout 15s) — riprova');
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

function renderResults(data, newSenderEmail, newCCAddrs) {
  const { overall_label, domains } = data;

  if (overall_label === 'danger')       setDot('danger');
  else if (overall_label === 'warning') setDot('warning');
  else if (newCCAddrs.length > 0)       setDot('warning');
  else if (newSenderEmail)              setDot('new');
  else                                  setDot('ok');

  if (newSenderEmail) {
    document.getElementById('banner-new-email').textContent = newSenderEmail;
    showBanner('banner-new');
  }

  if (newCCAddrs.length > 0) {
    const el = document.getElementById('banner-cc-detail');
    if (el) el.textContent = newCCAddrs.join(', ');
    showBanner('banner-cc');
  }

  if (overall_label === 'danger' || overall_label === 'warning') {
    const details = domains
      .filter(d => d.is_suspect)
      .map(d => d.similar_to ? `@${d.domain} (simile a @${d.similar_to})` : `@${d.domain}`)
      .join(', ');
    document.getElementById('banner-danger-detail').textContent = details;
    showBanner('banner-danger');
    document.getElementById('advice').classList.add('visible');
  }

  if (domains.length > 0) {
    const list = document.getElementById('domain-list');
    list.innerHTML = '';
    domains.forEach((d, i) => {
      const card = buildDomainCard(d);
      card.style.animationDelay = `${i * 55}ms`;
      list.appendChild(card);
    });
    document.getElementById('domains-section').classList.add('visible');
    document.getElementById('footer-addr').textContent = `${domains.length} domini analizzati`;
  } else {
    document.getElementById('footer-addr').textContent = 'Analisi completata';
  }
}

function buildDomainCard(d) {
  const cardClass = d.is_suspect ? (d.risk_label === 'danger' ? 'suspect' : 'warning-card') : 'safe';
  const badgeText = d.is_suspect ? (d.risk_label === 'danger' ? 'PERICOLO' : 'ATTENZIONE') : 'OK';
  const scoreColor = d.risk_label === 'danger' ? 'var(--danger)' : d.risk_label === 'warning' ? 'var(--warn)' : 'var(--ok)';
  const card = document.createElement('div');
  card.className = `dcard ${cardClass}`;
  const header = document.createElement('div');
  header.className = 'dcard-header';
  header.innerHTML = `
    <div class="dc-dot"></div>
    <div class="dc-main"><div class="dc-domain">@${esc(d.domain)}</div></div>
    <div class="dc-right">
      ${d.is_suspect ? `<div class="dc-score" style="color:${scoreColor}">${Number(d.risk_score)||0}</div>` : ''}
      <div class="dc-badge">${badgeText}</div>
      ${d.is_suspect ? '<div class="dc-chevron">&#9660;</div>' : ''}
    </div>`;
  card.appendChild(header);
  if (d.is_suspect) {
    const detail = document.createElement('div');
    detail.className = 'dcard-detail';
    detail.innerHTML = buildDetailHTML(d);
    card.appendChild(detail);
    header.addEventListener('click', () => card.classList.toggle('open'));
  }
  return card;
}

function buildDetailHTML(d) {
  let h = '<div class="dcard-detail-inner">';
  if (d.similar_to) {
    const { a, b } = buildDiff(d.similar_to, d.domain);
    h += `<div class="diff-block">
      <div class="diff-block-title">Confronto carattere per carattere</div>
      <div class="diff-row-pair">
        <div class="diff-line legit"><span class="diff-lbl">LEGIT</span><span class="diff-chars">${a}</span></div>
        <div class="diff-line fake"><span class="diff-lbl">FAKE?</span><span class="diff-chars">${b}</span></div>
      </div></div>`;
  }
  if (d.whois) {
    const w = d.whois; const ac = w.risk_flag ? 'danger' : 'ok';
    h += `<div class="detail-block"><div class="detail-block-title">WHOIS</div>
      <div class="detail-row"><span class="detail-key">Registrato</span><span class="detail-val ${ac}">${w.creation_date ? new Date(w.creation_date).toLocaleDateString('it-IT') : '—'}</span></div>
      <div class="detail-row"><span class="detail-key">Eta</span><span class="detail-val ${ac}">${esc(w.age_label||'—')}</span></div>
      <div class="detail-row"><span class="detail-key">Registrar</span><span class="detail-val">${esc(w.registrar||'—')}</span></div></div>`;
  }
  if (d.dns) {
    const dns = d.dns;
    h += `<div class="detail-block"><div class="detail-block-title">DNS</div>
      <div class="detail-row"><span class="detail-key">MX</span><span class="detail-val ${dns.has_mx?'ok':'danger'}">${dns.has_mx?'Presente':'Assente'}</span></div>
      <div class="detail-row"><span class="detail-key">SPF</span><span class="detail-val ${dns.has_spf?'ok':'danger'}">${dns.has_spf?'Presente':'Assente'}</span></div>
      <div class="detail-row"><span class="detail-key">DMARC</span><span class="detail-val ${dns.has_dmarc?'ok':'danger'}">${dns.has_dmarc?'Presente':'Assente'}</span></div></div>`;
  }
  if (d.reputation) {
    const rep = d.reputation;
    const vtMal = Number(rep.vt_malicious)||0, vtSus = Number(rep.vt_suspicious)||0;
    const vtc = vtMal > 0 ? 'danger' : vtSus > 0 ? 'warn' : 'ok';
    h += `<div class="detail-block"><div class="detail-block-title">Reputazione</div>`;
    if (rep.vt_available) {
      h += `<div class="detail-row"><span class="detail-key">VirusTotal</span><span class="detail-val ${vtc}">${vtMal} malevoli · ${vtSus} sospetti</span></div>`;
      if (rep.vt_link && String(rep.vt_link).startsWith('https://www.virustotal.com/'))
        h += `<div class="detail-row"><span class="detail-key">Report</span><span class="detail-val"><a href="${esc(rep.vt_link)}" target="_blank" rel="noopener noreferrer" style="color:var(--accent)">Apri VirusTotal</a></span></div>`;
    } else {
      h += `<div class="detail-row"><span class="detail-key">VirusTotal</span><span class="detail-val" style="color:var(--muted)">API key non configurata</span></div>`;
    }
    if (rep.gsb_available) {
      const threats = Array.isArray(rep.gsb_threats) ? rep.gsb_threats.map(t=>esc(String(t))).join(', ') : '';
      h += `<div class="detail-row"><span class="detail-key">Google SB</span><span class="detail-val ${rep.gsb_threats?.length?'danger':'ok'}">${rep.gsb_threats?.length ? threats : 'Non segnalato'}</span></div>`;
    }
    h += '</div>';
  }
  h += '</div>';
  return h;
}

function buildDiff(legit, fake) {
  const max = Math.max(legit.length, fake.length);
  let a = '', b = '';
  for (let i = 0; i < max; i++) {
    const ca = i < legit.length ? legit[i] : '';
    const cb = i < fake.length  ? fake[i]  : '';
    if (ca === cb) { a += esc(ca); b += esc(cb); }
    else {
      a += ca ? `<span class="ch">${esc(ca)}</span>` : '';
      b += cb ? `<span class="ch">${esc(cb)}</span>` : '<span class="ch">_</span>';
    }
  }
  return { a, b };
}

function setDot(state) {
  document.getElementById('status-dot').className = 'status-dot' + (state ? ' ' + state : '');
}
function setScanBar(on) {
  document.getElementById('scan-bar').classList.toggle('active', on);
}
function showBanner(id) {
  const el = document.getElementById(id);
  if (el) el.classList.add('visible');
}
function esc(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function resetMemory() {
  _cacheSenders.clear();
  Object.keys(_cacheCC).forEach(k => delete _cacheCC[k]);
  try { localStorage.removeItem(KEY_KNOWN_SENDERS); } catch {}
  try { localStorage.removeItem(KEY_KNOWN_CC); } catch {}
  if (_roamingOk()) {
    try {
      Office.context.roamingSettings.set(KEY_KNOWN_SENDERS, '[]');
      Office.context.roamingSettings.set(KEY_KNOWN_CC, '{}');
      Office.context.roamingSettings.saveAsync(() => {});
    } catch {}
  }
  document.getElementById('footer-addr').textContent = 'Memoria resettata';
  setTimeout(() => runScan(), 500);
}
