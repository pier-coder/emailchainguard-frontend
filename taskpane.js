/* ============================================================
   EmailChainGuard v3.1 — taskpane.js
   Fix sicurezza:
   - Validazione URL VirusTotal (blocca javascript:)
   - Limite 1000 email in memoria
   - Timeout fetch 15 secondi
   - Validazione risposta backend
   ============================================================ */
'use strict';

const CONFIG = {
  BACKEND_URL: 'https://emailchainguard-backend-production.up.railway.app',
  ECG_API_KEY: 'MalibuStacy25_5', // <-- sostituisci con la tua chiave
  MAX_KNOWN_EMAILS: 1000,
  FETCH_TIMEOUT_MS: 15000,
};

// ── Storage email conosciute ──────────────────────────────────────────
const STORAGE_KEY = 'ecg_known_emails_v3';
const _memCache = new Set();

function _roamingOk() {
  try { return !!(Office?.context?.roamingSettings); }
  catch { return false; }
}

function loadKnownEmails() {
  if (_roamingOk()) {
    try {
      const raw = Office.context.roamingSettings.get(STORAGE_KEY);
      const stored = raw ? new Set(JSON.parse(raw)) : new Set();
      _memCache.forEach(e => stored.add(e));
      return stored;
    } catch { /* fallthrough */ }
  }
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    const stored = raw ? new Set(JSON.parse(raw)) : new Set();
    _memCache.forEach(e => stored.add(e));
    return stored;
  } catch { /* fallthrough */ }
  return new Set(_memCache);
}

function saveKnownEmails(emails) {
  emails.forEach(e => _memCache.add(e));
  const all = loadKnownEmails();
  emails.forEach(e => all.add(e));

  // Limite massimo 1000 email — rimuove le più vecchie
  let arr = Array.from(all);
  if (arr.length > CONFIG.MAX_KNOWN_EMAILS) {
    arr = arr.slice(arr.length - CONFIG.MAX_KNOWN_EMAILS);
  }

  const serialized = JSON.stringify(arr);
  if (_roamingOk()) {
    try {
      Office.context.roamingSettings.set(STORAGE_KEY, serialized);
      Office.context.roamingSettings.saveAsync(() => {});
    } catch { /* fallthrough */ }
  }
  try { localStorage.setItem(STORAGE_KEY, serialized); } catch { /* bloccato */ }
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

// ── Reset UI ──────────────────────────────────────────────────────────
function resetUI() {
  ['banner-new','banner-danger','banner-link'].forEach(id =>
    document.getElementById(id).classList.remove('visible')
  );
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
    const localAddrs = collectAddresses(item);
    const knownEmails = loadKnownEmails();
    const fromAddr = item.from?.emailAddress?.toLowerCase() || null;
    const isNewSender = fromAddr && !knownEmails.has(fromAddr);

    saveKnownEmails(localAddrs);

    const domains = [...new Set(localAddrs.map(a => a.split('@')[1]).filter(Boolean))];
    const result = await callBackend(domains, item.conversationId || null);

    // Validazione risposta backend
    if (!result || !Array.isArray(result.domains)) {
      throw new Error('Risposta backend non valida');
    }

    renderResults(result, isNewSender ? fromAddr : null);

  } catch (err) {
    setDot('');
    const p = document.getElementById('idle-state').querySelector('p');
    p.textContent = 'Errore: ' + (err.message || 'impossibile completare la scansione');
    document.getElementById('idle-state').style.display = 'flex';
  } finally {
    setScanBar(false);
  }
}

// ── Raccolta indirizzi ────────────────────────────────────────────────
function collectAddresses(item) {
  const addrs = new Set();
  const push = obj => { if (obj?.emailAddress) addrs.add(obj.emailAddress.toLowerCase()); };
  push(item.from);
  (item.to || []).forEach(push);
  (item.cc || []).forEach(push);
  return Array.from(addrs);
}

// ── Backend ───────────────────────────────────────────────────────────
async function callBackend(domains, conversationId) {
  // Timeout di 15 secondi
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), CONFIG.FETCH_TIMEOUT_MS);

  try {
    const resp = await fetch(`${CONFIG.BACKEND_URL}/analyze`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-ECG-Key': CONFIG.ECG_API_KEY,
      },
      body: JSON.stringify({ domains, conversation_id: conversationId }),
      signal: controller.signal,
    });

    if (!resp.ok) {
      const err = await resp.json().catch(() => ({}));
      throw new Error(err.detail || `Errore backend HTTP ${resp.status}`);
    }
    return resp.json();
  } catch (err) {
    if (err.name === 'AbortError') {
      throw new Error('Analisi scaduta (timeout 15s) — riprova');
    }
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

// ── Rendering ─────────────────────────────────────────────────────────
function renderResults(data, newSenderEmail) {
  const { overall_label, domains } = data;

  if (overall_label === 'danger')       setDot('danger');
  else if (overall_label === 'warning') setDot('warning');
  else if (newSenderEmail)              setDot('new');
  else                                  setDot('ok');

  if (newSenderEmail) {
    document.getElementById('banner-new-email').textContent = newSenderEmail;
    showBanner('banner-new');
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

  const list = document.getElementById('domain-list');
  list.innerHTML = '';
  domains.forEach((d, i) => {
    const card = buildDomainCard(d);
    card.style.animationDelay = `${i * 55}ms`;
    list.appendChild(card);
  });
  document.getElementById('domains-section').classList.add('visible');
  document.getElementById('footer-addr').textContent = `${domains.length} domini analizzati`;
}

// ── Domain card ───────────────────────────────────────────────────────
function buildDomainCard(d) {
  const cardClass = d.is_suspect
    ? (d.risk_label === 'danger' ? 'suspect' : 'warning-card')
    : 'safe';
  const badgeText = d.is_suspect
    ? (d.risk_label === 'danger' ? 'PERICOLO' : 'ATTENZIONE')
    : 'OK';
  const scoreColor = d.risk_label === 'danger' ? 'var(--danger)'
    : d.risk_label === 'warning' ? 'var(--warn)' : 'var(--ok)';

  const card = document.createElement('div');
  card.className = `dcard ${cardClass}`;

  const header = document.createElement('div');
  header.className = 'dcard-header';
  header.innerHTML = `
    <div class="dc-dot"></div>
    <div class="dc-main">
      <div class="dc-domain">@${esc(d.domain)}</div>
    </div>
    <div class="dc-right">
      ${d.is_suspect ? `<div class="dc-score" style="color:${scoreColor}">${Number(d.risk_score) || 0}</div>` : ''}
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
    const w = d.whois;
    const ac = w.risk_flag ? 'danger' : 'ok';
    h += `<div class="detail-block">
      <div class="detail-block-title">WHOIS</div>
      <div class="detail-row"><span class="detail-key">Registrato</span><span class="detail-val ${ac}">${w.creation_date ? new Date(w.creation_date).toLocaleDateString('it-IT') : '—'}</span></div>
      <div class="detail-row"><span class="detail-key">Eta</span><span class="detail-val ${ac}">${esc(w.age_label || '—')}</span></div>
      <div class="detail-row"><span class="detail-key">Registrar</span><span class="detail-val">${esc(w.registrar || '—')}</span></div>
    </div>`;
  }

  if (d.dns) {
    const dns = d.dns;
    h += `<div class="detail-block">
      <div class="detail-block-title">DNS</div>
      <div class="detail-row"><span class="detail-key">MX</span><span class="detail-val ${dns.has_mx ? 'ok' : 'danger'}">${dns.has_mx ? 'Presente' : 'Assente'}</span></div>
      <div class="detail-row"><span class="detail-key">SPF</span><span class="detail-val ${dns.has_spf ? 'ok' : 'danger'}">${dns.has_spf ? 'Presente' : 'Assente'}</span></div>
      <div class="detail-row"><span class="detail-key">DMARC</span><span class="detail-val ${dns.has_dmarc ? 'ok' : 'danger'}">${dns.has_dmarc ? 'Presente' : 'Assente'}</span></div>
    </div>`;
  }

  if (d.reputation) {
    const rep = d.reputation;
    const vtMal = Number(rep.vt_malicious) || 0;
    const vtSus = Number(rep.vt_suspicious) || 0;
    const vtc = vtMal > 0 ? 'danger' : vtSus > 0 ? 'warn' : 'ok';

    h += `<div class="detail-block"><div class="detail-block-title">Reputazione</div>`;

    if (rep.vt_available) {
      h += `<div class="detail-row"><span class="detail-key">VirusTotal</span><span class="detail-val ${vtc}">${vtMal} malevoli · ${vtSus} sospetti</span></div>`;

      // Validazione URL VirusTotal — blocca qualsiasi cosa non sia https://www.virustotal.com
      if (rep.vt_link && String(rep.vt_link).startsWith('https://www.virustotal.com/')) {
        h += `<div class="detail-row"><span class="detail-key">Report</span><span class="detail-val"><a href="${esc(rep.vt_link)}" target="_blank" rel="noopener noreferrer" style="color:var(--accent)">Apri VirusTotal</a></span></div>`;
      }
    } else {
      h += `<div class="detail-row"><span class="detail-key">VirusTotal</span><span class="detail-val" style="color:var(--muted)">API key non configurata</span></div>`;
    }

    if (rep.gsb_available) {
      const threats = Array.isArray(rep.gsb_threats)
        ? rep.gsb_threats.map(t => esc(String(t))).join(', ')
        : 'Non segnalato';
      h += `<div class="detail-row"><span class="detail-key">Google SB</span><span class="detail-val ${rep.gsb_threats?.length ? 'danger' : 'ok'}">${rep.gsb_threats?.length ? threats : 'Non segnalato'}</span></div>`;
    }

    h += '</div>';
  }

  h += '</div>';
  return h;
}

// ── Diff visuale ──────────────────────────────────────────────────────
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

// ── UI helpers ────────────────────────────────────────────────────────
function setDot(state) {
  document.getElementById('status-dot').className =
    'status-dot' + (state ? ' ' + state : '');
}
function setScanBar(on) {
  document.getElementById('scan-bar').classList.toggle('active', on);
}
function showBanner(id) {
  document.getElementById(id).classList.add('visible');
}
function esc(s) {
  return String(s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
