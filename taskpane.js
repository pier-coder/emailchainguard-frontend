/* ============================================================
   EmailChainGuard v2 – taskpane.js
   Microsoft Graph API + WHOIS + DNS + Reputation + Risk Score
   ============================================================ */
'use strict';

// ── Config – questi valori vengono sostituiti in produzione ───────
// In locale, AZURE_CLIENT_ID va inserito anche nel manifest.xml
const CONFIG = {
  AZURE_CLIENT_ID: window.AZURE_CLIENT_ID || 'IL-TUO-CLIENT-ID-QUI',
  BACKEND_URL:     window.BACKEND_URL     || 'https://emailchainguard-backend-production.up.railway.app',
  ECG_API_KEY:     window.ECG_API_KEY     || 'ecg-dev-key-2024',
  SCOPES:          ['Mail.Read', 'User.Read'],
};

// ── Memoria domini ───────────────────────────────────────────────
// Prova roamingSettings (Exchange/M365), altrimenti sessionStorage.
// In entrambi i casi i dati restano sul dispositivo dell'utente.

const ECG_KEY = 'ecg_known_domains';

function _roamingOk() {
  try { return !!(Office && Office.context && Office.context.roamingSettings); }
  catch { return false; }
}

function loadKnownDomains() {
  try {
    let raw = null;
    if (_roamingOk()) {
      raw = Office.context.roamingSettings.get(ECG_KEY);
    } else {
      raw = sessionStorage.getItem(ECG_KEY);
    }
    return raw ? new Set(JSON.parse(raw)) : new Set();
  } catch { return new Set(); }
}

function saveKnownDomains(newDomains) {
  try {
    const existing = loadKnownDomains();
    newDomains.forEach(d => existing.add(d));
    const serialized = JSON.stringify(Array.from(existing));
    if (_roamingOk()) {
      Office.context.roamingSettings.set(ECG_KEY, serialized);
      Office.context.roamingSettings.saveAsync(() => {});
    } else {
      sessionStorage.setItem(ECG_KEY, serialized);
    }
  } catch (e) { console.warn('ECG: storage error', e); }
}

// ── MSAL instance ────────────────────────────────────────────────
let msalInstance = null;
let graphToken   = null;

function initMsal() {
  if (!window.msal) return;
  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId:    CONFIG.AZURE_CLIENT_ID,
      authority:   'https://login.microsoftonline.com/common',
      redirectUri: window.location.origin + '/auth-callback.html',
    },
    cache: { cacheLocation: 'sessionStorage' },
  });
}

async function handleAuth() {
  if (graphToken) return;   // già autenticato
  if (!msalInstance) {
    setStatus('idle', '⚠', 'MSAL non caricato', 'Verifica la configurazione Azure nel README');
    return;
  }

  try {
    setStatus('scanning', '◌', 'Accesso in corso…', 'Finestra di login Microsoft');
    const resp = await msalInstance.loginPopup({ scopes: CONFIG.SCOPES });
    graphToken = resp.accessToken;
    setBadgeLoggedIn(resp.account.username);
    setStatus('idle', '◌', 'Autenticato', 'Premi Scansiona per analizzare la conversazione completa');
  } catch (err) {
    setStatus('idle', '⚠', 'Login annullato', err.message || '');
  }
}

function setBadgeLoggedIn(email) {
  const badge = document.getElementById('auth-badge');
  badge.className = 'logged-in';
  badge.onclick = null;
  document.getElementById('auth-label').textContent = '✓ ' + email.split('@')[0];
}


// ── Office.js init ───────────────────────────────────────────────
Office.initialize = function () {
  Office.onReady(info => {
    if (info.host === Office.HostType.Outlook) {
      initMsal();
      runScan();
    }
  });
};


// ════════════════════════════════════════════════════════════
//  ENTRY POINT
// ════════════════════════════════════════════════════════════
function resetUI() {
  document.getElementById('domains-section').classList.remove('visible');
  document.getElementById('risk-gauge-section').classList.remove('visible');
  document.getElementById('advice').classList.remove('visible');
  document.getElementById('first-time-box').classList.remove('visible');
  document.getElementById('domain-list').innerHTML = '';
  document.getElementById('footer-addr').textContent = '';
}

async function runScan() {
  const item = Office.context.mailbox.item;
  if (!item) {
    setStatus('idle', '◌', 'Nessuna email aperta', 'Apri un messaggio per avviare la scansione');
    return;
  }

  setStatus('scanning', '◌', 'Raccolta indirizzi…', 'Lettura mittenti dell\'email corrente');
  setScanBar(true);
  disableBtn(true);
  hideEmptyState();
  document.getElementById('domains-section').classList.remove('visible');
  document.getElementById('advice').classList.remove('visible');
  document.getElementById('risk-gauge-section').classList.remove('visible');

  try {
    // 1. Raccogli indirizzi dall'email corrente
    const localAddrs = collectLocalAddresses(item);
    const convId     = item.conversationId || null;

    // 2. Mostra badge Graph se token disponibile
    if (graphToken && convId) {
      showGraphBadge('Lettura conversazione completa via Graph API…');
    }

    // 3. Chiama il backend
    const result = await callBackend(localAddrs, convId);

    hideGraphBadge();
    renderResults(result);

  } catch (err) {
    hideGraphBadge();
    setStatus('idle', '⚠', 'Errore', err.message || 'Impossibile completare la scansione');
  } finally {
    setScanBar(false);
    disableBtn(false);
  }
}


// ════════════════════════════════════════════════════════════
//  RACCOLTA INDIRIZZI LOCALI
// ════════════════════════════════════════════════════════════
function collectLocalAddresses(item) {
  const addrs = new Set();

  const push = (obj) => {
    if (obj && obj.emailAddress) addrs.add(obj.emailAddress.toLowerCase());
  };
  const pushList = (arr) => (arr || []).forEach(push);

  push(item.from);
  pushList(item.to);
  pushList(item.cc);

  return Array.from(addrs);
}


// ════════════════════════════════════════════════════════════
//  CHIAMATA AL BACKEND
// ════════════════════════════════════════════════════════════
async function callBackend(addresses, conversationId) {
  const headers = {
    'Content-Type': 'application/json',
    'X-ECG-Key': CONFIG.ECG_API_KEY,
  };
  if (graphToken) headers['Authorization'] = `Bearer ${graphToken}`;

  // Manda solo domini, non email complete — minimizzazione dati GDPR
  const domains = [...new Set(addresses.map(a => a.split('@')[1]).filter(Boolean))];
  const body = JSON.stringify({ domains, conversation_id: conversationId });

  const resp = await fetch(`${CONFIG.BACKEND_URL}/analyze`, {
    method:  'POST',
    headers,
    body,
  });

  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    throw new Error(err.detail || `Backend HTTP ${resp.status}`);
  }

  return resp.json();
}


// ════════════════════════════════════════════════════════════
//  RENDERING
// ════════════════════════════════════════════════════════════
function renderResults(data) {
  // ── Overall status ──
  const { overall_label, max_risk_score, suspect_count, domains } = data;

  if (overall_label === 'danger') {
    setStatus('danger', '🚨', `${suspect_count} dominio/i sospetto/i`, 'Possibile spoofing — verifica prima di rispondere!');
    document.getElementById('advice').classList.add('visible');
  } else if (overall_label === 'warning') {
    setStatus('warning', '⚠', `${suspect_count} dominio/i da verificare`, 'Controlla i dettagli prima di rispondere');
    document.getElementById('advice').classList.add('visible');
  } else {
    setStatus('ok', '✓', 'Catena coerente', `${domains.length} dominio/i analizzato/i — nessuna anomalia`);
  }

  // ── Risk gauge ──
  renderGauge(max_risk_score, overall_label);

  // ── Prima volta (roamingSettings locali) ──
  const knownDomains = loadKnownDomains();
  const newDomains = domains.filter(d => !d.is_suspect && !knownDomains.has(d.domain));
  renderFirstTime(newDomains);
  // Salva i nuovi domini legittimi
  saveKnownDomains(newDomains.map(d => d.domain));

  // ── Domain list ──
  const list = document.getElementById('domain-list');
  list.innerHTML = '';
  domains.forEach((d, idx) => {
    const card = buildDomainCard(d, knownDomains);
    card.style.animationDelay = `${idx * 55}ms`;
    list.appendChild(card);
  });

  document.getElementById('domains-section').classList.add('visible');

  // ── Footer ──
  document.getElementById('footer-addr').textContent =
    `${domains.length} domini analizzati`;
}

// ── First time notice ─────────────────────────────────────────────
function renderFirstTime(newDomains) {
  const box = document.getElementById('first-time-box');
  const list = document.getElementById('first-time-list');
  if (!newDomains.length) {
    box.classList.remove('visible');
    return;
  }
  list.innerHTML = newDomains
    .map(d => `<div class="ft-domain">@${esc(d.domain)}</div>`)
    .join('');
  box.classList.add('visible');
}

// ── Gauge ────────────────────────────────────────────────────────
function renderGauge(score, label) {
  const sec    = document.getElementById('risk-gauge-section');
  const fill   = document.getElementById('gauge-fill');
  const scoreEl= document.getElementById('gauge-score');
  const labelEl= document.getElementById('gauge-label');

  sec.classList.add('visible');

  const colorClass = label === 'danger' ? 'fill-danger' : label === 'warning' ? 'fill-warning' : 'fill-ok';
  fill.className = `gauge-bar-fill ${colorClass}`;

  // Animated fill
  requestAnimationFrame(() => {
    setTimeout(() => { fill.style.width = `${score}%`; }, 50);
  });

  scoreEl.textContent = score;
  scoreEl.style.color = label === 'danger' ? 'var(--danger)' : label === 'warning' ? 'var(--warn)' : 'var(--ok)';

  const labels = { ok: 'SICURO', warning: 'ATTENZIONE', danger: 'PERICOLO' };
  const colors = { ok: 'var(--ok)', warning: 'var(--warn)', danger: 'var(--danger)' };
  labelEl.textContent = labels[label] || '—';
  labelEl.style.color = colors[label];
}

// ── Domain card ──────────────────────────────────────────────────
function buildDomainCard(d, knownDomains) {
  const isNew = !d.is_suspect && knownDomains && !knownDomains.has(d.domain);
  const cardClass = d.is_suspect
    ? (d.risk_label === 'danger' ? 'suspect' : 'warning-card')
    : 'safe';

  const scoreColor = d.risk_label === 'danger' ? 'var(--danger)' : d.risk_label === 'warning' ? 'var(--warn)' : 'var(--ok)';
  const badgeText  = d.is_suspect ? (d.risk_label === 'danger' ? 'PERICOLO' : 'ATTENZIONE') : 'OK';

  const card = document.createElement('div');
  card.className = `dcard ${cardClass}`;

  // Header
  const header = document.createElement('div');
  header.className = 'dcard-header';
  header.innerHTML = `
    <div class="dc-dot"></div>
    <div class="dc-main">
      <div class="dc-domain">@${esc(d.domain)}</div>
      <div class="dc-emails">@${esc(d.domain)}</div>
    </div>
    <div class="dc-right">
      ${d.is_suspect ? `<div class="dc-score" style="color:${scoreColor}">${d.risk_score}</div>` : ''}
      <div class="dc-badge">${badgeText}</div>
      ${d.is_suspect ? '<div class="dc-chevron">▼</div>' : ''}
    </div>
  `;

  card.appendChild(header);

  // Detail panel (solo per sospetti)
  if (d.is_suspect) {
    const detail = document.createElement('div');
    detail.className = 'dcard-detail';
    detail.innerHTML = buildDetailHTML(d);
    card.appendChild(detail);

    header.addEventListener('click', () => {
      card.classList.toggle('open');
    });
  }

  return card;
}

function buildDetailHTML(d) {
  let html = '<div class="dcard-detail-inner">';

  // ── Diff visuale ──
  if (d.similar_to) {
    const { htmlA, htmlB } = buildDiff(d.similar_to, d.domain);
    html += `
      <div class="diff-block">
        <div class="diff-block-title">⚠ Confronto carattere per carattere</div>
        <div class="diff-row-pair">
          <div class="diff-line legit">
            <span class="diff-lbl">LEGIT</span>
            <span class="diff-chars">${htmlA}</span>
          </div>
          <div class="diff-line fake">
            <span class="diff-lbl">FAKE?</span>
            <span class="diff-chars">${htmlB}</span>
          </div>
        </div>
      </div>`;
  }

  // ── WHOIS ──
  if (d.whois) {
    const w = d.whois;
    const ageClass = w.risk_flag ? 'danger' : 'ok';
    html += `
      <div class="detail-block">
        <div class="detail-block-title">🗓 WHOIS · Registrazione dominio</div>
        <div class="detail-row">
          <span class="detail-key">Registrato</span>
          <span class="detail-val ${ageClass}">${w.creation_date ? new Date(w.creation_date).toLocaleDateString('it-IT') : '—'}</span>
        </div>
        <div class="detail-row">
          <span class="detail-key">Età</span>
          <span class="detail-val ${ageClass}">${w.age_days != null ? w.age_days + ' giorni · ' + w.age_label : w.age_label}</span>
        </div>
        <div class="detail-row">
          <span class="detail-key">Registrar</span>
          <span class="detail-val">${esc(w.registrar || '—')}</span>
        </div>
      </div>`;
  }

  // ── DNS ──
  if (d.dns) {
    const dns = d.dns;
    html += `
      <div class="detail-block">
        <div class="detail-block-title">🌐 DNS · Record email</div>
        <div class="detail-row">
          <span class="detail-key">MX</span>
          <span class="detail-val ${dns.has_mx ? 'ok' : 'danger'}">${dns.has_mx ? '✓ Presente' : '✗ Assente'}</span>
        </div>
        <div class="detail-row">
          <span class="detail-key">SPF</span>
          <span class="detail-val ${dns.has_spf ? 'ok' : 'danger'}">${dns.has_spf ? '✓ Presente' : '✗ Assente'}</span>
        </div>
        <div class="detail-row">
          <span class="detail-key">DMARC</span>
          <span class="detail-val ${dns.has_dmarc ? 'ok' : 'danger'}">${dns.has_dmarc ? '✓ Presente' : '✗ Assente'}</span>
        </div>
        ${dns.risk_flags.length ? `<div class="detail-row"><span class="detail-key">Note</span><span class="detail-val warn">${dns.risk_flags.join(' · ')}</span></div>` : ''}
      </div>`;
  }

  // ── Reputation ──
  if (d.reputation) {
    const rep = d.reputation;
    const vtClass = rep.vt_malicious > 0 ? 'danger' : rep.vt_suspicious > 0 ? 'warn' : 'ok';
    html += `
      <div class="detail-block">
        <div class="detail-block-title">🛡 Reputazione</div>`;

    if (rep.vt_available) {
      html += `
        <div class="detail-row">
          <span class="detail-key">VirusTotal</span>
          <span class="detail-val ${vtClass}">${rep.vt_malicious} malevolo · ${rep.vt_suspicious} sospetti · ${rep.vt_clean} puliti</span>
        </div>`;
      if (rep.vt_link) {
        html += `<div class="detail-row"><span class="detail-key">Report</span><span class="detail-val"><a href="${rep.vt_link}" target="_blank" style="color:var(--accent)">Apri su VirusTotal ↗</a></span></div>`;
      }
    } else {
      html += `<div class="detail-row"><span class="detail-key">VirusTotal</span><span class="detail-val" style="color:var(--muted)">API key non configurata</span></div>`;
    }

    if (rep.gsb_available) {
      html += `
        <div class="detail-row">
          <span class="detail-key">Google SB</span>
          <span class="detail-val ${rep.gsb_threats.length ? 'danger' : 'ok'}">${rep.gsb_threats.length ? rep.gsb_threats.join(', ') : '✓ Non segnalato'}</span>
        </div>`;
    } else {
      html += `<div class="detail-row"><span class="detail-key">Google SB</span><span class="detail-val" style="color:var(--muted)">API key non configurata</span></div>`;
    }

    html += `</div>`;
  }

  html += '</div>';
  return html;
}


// ════════════════════════════════════════════════════════════
//  DIFF VISUALE
// ════════════════════════════════════════════════════════════
function buildDiff(legit, fake) {
  const max = Math.max(legit.length, fake.length);
  let htmlA = '', htmlB = '';
  for (let i = 0; i < max; i++) {
    const ca = i < legit.length ? legit[i] : '';
    const cb = i < fake.length  ? fake[i]  : '';
    if (ca === cb) {
      htmlA += esc(ca);
      htmlB += esc(cb);
    } else {
      htmlA += ca ? `<span class="ch">${esc(ca)}</span>` : '';
      htmlB += cb ? `<span class="ch">${esc(cb)}</span>` : '<span class="ch">_</span>';
    }
  }
  return { htmlA, htmlB };
}


// ════════════════════════════════════════════════════════════
//  UI HELPERS
// ════════════════════════════════════════════════════════════
function setStatus(type, icon, title, sub) {
  const el = document.getElementById('overall-status');
  el.className = type;
  document.getElementById('st-icon').textContent  = icon;
  document.getElementById('st-title').textContent = title;
  document.getElementById('st-sub').textContent   = sub;
}

function setScanBar(on) {
  document.getElementById('scan-bar').classList.toggle('active', on);
}

function disableBtn(d) {
  document.getElementById('btn-scan').disabled = d;
}

function hideEmptyState() {
  document.getElementById('empty-state').style.display = 'none';
}

function showGraphBadge(txt) {
  const b = document.getElementById('graph-badge');
  document.getElementById('graph-badge-text').textContent = txt;
  b.classList.add('visible');
}

function hideGraphBadge() {
  document.getElementById('graph-badge').classList.remove('visible');
}

function esc(s) {
  return String(s)
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;');
}
