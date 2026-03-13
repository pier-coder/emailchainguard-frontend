/* ============================================================
   EmailChainGuard v2.1 – taskpane.js
   Nuovo layout: banner + scansione automatica
   Primo contatto: email completa (non solo dominio)
   ============================================================ */
'use strict';

const CONFIG = {
  AZURE_CLIENT_ID: window.AZURE_CLIENT_ID || 'IL-TUO-CLIENT-ID-QUI',
  BACKEND_URL:     window.BACKEND_URL     || 'https://emailchainguard-backend-production.up.railway.app',
  ECG_API_KEY:     window.ECG_API_KEY     || 'Columbus_25_1',
  SCOPES:          ['Mail.Read', 'User.Read'],
};

// ── Memoria email conosciute ──────────────────────────────────────
// Salva email complete (non domini) — es. mario@gmail.com
// così mario@gmail.com e luigi@gmail.com sono distinti.
const ECG_KEY = 'ecg_known_emails_v2';

function _roamingOk() {
  try { return !!(Office && Office.context && Office.context.roamingSettings); }
  catch { return false; }
}

function loadKnownEmails() {
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

function saveKnownEmails(newEmails) {
  try {
    const existing = loadKnownEmails();
    newEmails.forEach(e => existing.add(e));
    const serialized = JSON.stringify(Array.from(existing));
    if (_roamingOk()) {
      Office.context.roamingSettings.set(ECG_KEY, serialized);
      Office.context.roamingSettings.saveAsync(() => {});
    } else {
      sessionStorage.setItem(ECG_KEY, serialized);
    }
  } catch (e) { console.warn('ECG: storage error', e); }
}

// ── MSAL ──────────────────────────────────────────────────────────
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

// ── Office init ───────────────────────────────────────────────────
Office.initialize = function () {};

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    initMsal();
    // Scansione automatica all'apertura
    runScan();
    // Rescan ad ogni cambio email
    if (Office.context.mailbox.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(
        Office.EventType.ItemChanged,
        () => { resetUI(); runScan(); }
      );
    }
  }
});


// ════════════════════════════════════════════════════════════
//  RESET UI
// ════════════════════════════════════════════════════════════
function resetUI() {
  hideBanner('banner-new');
  hideBanner('banner-danger');
  hideBanner('banner-link');
  document.getElementById('advice').classList.remove('visible');
  document.getElementById('domains-section').classList.remove('visible');
  document.getElementById('domain-list').innerHTML = '';
  document.getElementById('footer-addr').textContent = '—';
  setDot('');
}


// ════════════════════════════════════════════════════════════
//  SCANSIONE
// ════════════════════════════════════════════════════════════
async function runScan() {
  const item = Office.context.mailbox.item;
  if (!item) return;

  setDot('scanning');
  setScanBar(true);
  document.getElementById('idle-state').style.display = 'none';

  try {
    // 1. Raccogli indirizzi email completi dall'email corrente
    const localAddrs = collectLocalAddresses(item);

    // 2. Chiama il backend (manda solo domini per GDPR)
    const domains = [...new Set(localAddrs.map(a => a.split('@')[1]).filter(Boolean))];
    const result  = await callBackend(domains, item.conversationId || null);

    // 3. Controlla primo contatto (email completa, locale)
    const knownEmails   = loadKnownEmails();
    const newSenders    = localAddrs.filter(e => !knownEmails.has(e));
    const fromAddr      = item.from && item.from.emailAddress
      ? item.from.emailAddress.toLowerCase() : null;
    const isNewSender   = fromAddr && newSenders.includes(fromAddr);

    // Salva tutti i mittenti di questa email come conosciuti
    saveKnownEmails(localAddrs);

    // 4. Render
    renderResults(result, isNewSender ? fromAddr : null);

  } catch (err) {
    setDot('');
    document.getElementById('idle-state').style.display = 'flex';
    document.getElementById('idle-state').querySelector('p').textContent =
      'Errore: ' + (err.message || 'impossibile completare la scansione');
  } finally {
    setScanBar(false);
  }
}


// ════════════════════════════════════════════════════════════
//  RACCOLTA INDIRIZZI
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
//  BACKEND
// ════════════════════════════════════════════════════════════
async function callBackend(domains, conversationId) {
  const headers = {
    'Content-Type': 'application/json',
    'X-ECG-Key': CONFIG.ECG_API_KEY,
  };
  if (graphToken) headers['Authorization'] = `Bearer ${graphToken}`;

  const resp = await fetch(`${CONFIG.BACKEND_URL}/analyze`, {
    method:  'POST',
    headers,
    body:    JSON.stringify({ domains, conversation_id: conversationId }),
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
function renderResults(data, newSenderEmail) {
  const { overall_label, suspect_count, domains } = data;

  // ── Status dot ──
  if (overall_label === 'danger') {
    setDot('danger');
  } else if (overall_label === 'warning') {
    setDot('warning');
  } else if (newSenderEmail) {
    setDot('new');
  } else {
    setDot('ok');
  }

  // ── Banner primo contatto ──
  if (newSenderEmail && overall_label === 'ok') {
    document.getElementById('banner-new-email').textContent = newSenderEmail;
    showBanner('banner-new');
  }

  // ── Banner dominio sospetto ──
  if (overall_label === 'danger' || overall_label === 'warning') {
    const suspectDomains = domains
      .filter(d => d.is_suspect)
      .map(d => {
        if (d.similar_to) return `@${d.domain} (simile a @${d.similar_to})`;
        return `@${d.domain}`;
      })
      .join(', ');
    document.getElementById('banner-danger-detail').textContent = suspectDomains;
    showBanner('banner-danger');
    document.getElementById('advice').classList.add('visible');
  }

  // ── Domain list ──
  const list = document.getElementById('domain-list');
  list.innerHTML = '';
  const knownEmails = loadKnownEmails();

  domains.forEach((d, idx) => {
    const card = buildDomainCard(d);
    card.style.animationDelay = `${idx * 55}ms`;
    list.appendChild(card);
  });

  document.getElementById('domains-section').classList.add('visible');
  document.getElementById('footer-addr').textContent = `${domains.length} domini analizzati`;
}


// ── Domain card ───────────────────────────────────────────────────
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
      ${d.is_suspect ? `<div class="dc-score" style="color:${scoreColor}">${d.risk_score}</div>` : ''}
      <div class="dc-badge">${badgeText}</div>
      ${d.is_suspect ? '<div class="dc-chevron">▼</div>' : ''}
    </div>
  `;
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
  let html = '<div class="dcard-detail-inner">';

  if (d.similar_to) {
    const { htmlA, htmlB } = buildDiff(d.similar_to, d.domain);
    html += `
      <div class="diff-block">
        <div class="diff-block-title">⚠ Confronto carattere per carattere</div>
        <div class="diff-row-pair">
          <div class="diff-line legit"><span class="diff-lbl">LEGIT</span><span class="diff-chars">${htmlA}</span></div>
          <div class="diff-line fake"><span class="diff-lbl">FAKE?</span><span class="diff-chars">${htmlB}</span></div>
        </div>
      </div>`;
  }

  if (d.whois) {
    const w = d.whois;
    const ageClass = w.risk_flag ? 'danger' : 'ok';
    html += `
      <div class="detail-block">
        <div class="detail-block-title">WHOIS · Registrazione dominio</div>
        <div class="detail-row"><span class="detail-key">Registrato</span><span class="detail-val ${ageClass}">${w.creation_date ? new Date(w.creation_date).toLocaleDateString('it-IT') : '—'}</span></div>
        <div class="detail-row"><span class="detail-key">Età</span><span class="detail-val ${ageClass}">${w.age_days != null ? w.age_days + ' giorni · ' + w.age_label : w.age_label}</span></div>
        <div class="detail-row"><span class="detail-key">Registrar</span><span class="detail-val">${esc(w.registrar || '—')}</span></div>
      </div>`;
  }

  if (d.dns) {
    const dns = d.dns;
    html += `
      <div class="detail-block">
        <div class="detail-block-title">DNS · Record email</div>
        <div class="detail-row"><span class="detail-key">MX</span><span class="detail-val ${dns.has_mx ? 'ok' : 'danger'}">${dns.has_mx ? '✓ Presente' : '✗ Assente'}</span></div>
        <div class="detail-row"><span class="detail-key">SPF</span><span class="detail-val ${dns.has_spf ? 'ok' : 'danger'}">${dns.has_spf ? '✓ Presente' : '✗ Assente'}</span></div>
        <div class="detail-row"><span class="detail-key">DMARC</span><span class="detail-val ${dns.has_dmarc ? 'ok' : 'danger'}">${dns.has_dmarc ? '✓ Presente' : '✗ Assente'}</span></div>
        ${dns.risk_flags && dns.risk_flags.length ? `<div class="detail-row"><span class="detail-key">Note</span><span class="detail-val warn">${dns.risk_flags.join(' · ')}</span></div>` : ''}
      </div>`;
  }

  if (d.reputation) {
    const rep = d.reputation;
    const vtClass = rep.vt_malicious > 0 ? 'danger' : rep.vt_suspicious > 0 ? 'warn' : 'ok';
    html += `<div class="detail-block"><div class="detail-block-title">Reputazione</div>`;
    if (rep.vt_available) {
      html += `<div class="detail-row"><span class="detail-key">VirusTotal</span><span class="detail-val ${vtClass}">${rep.vt_malicious} malevolo · ${rep.vt_suspicious} sospetti · ${rep.vt_clean} puliti</span></div>`;
      if (rep.vt_link) html += `<div class="detail-row"><span class="detail-key">Report</span><span class="detail-val"><a href="${rep.vt_link}" target="_blank" style="color:var(--accent)">Apri su VirusTotal ↗</a></span></div>`;
    } else {
      html += `<div class="detail-row"><span class="detail-key">VirusTotal</span><span class="detail-val" style="color:var(--muted)">API key non configurata</span></div>`;
    }
    if (rep.gsb_available) {
      html += `<div class="detail-row"><span class="detail-key">Google SB</span><span class="detail-val ${rep.gsb_threats.length ? 'danger' : 'ok'}">${rep.gsb_threats.length ? rep.gsb_threats.join(', ') : '✓ Non segnalato'}</span></div>`;
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
function setDot(state) {
  const dot = document.getElementById('status-dot');
  dot.className = 'status-dot' + (state ? ' ' + state : '');
}

function setScanBar(on) {
  document.getElementById('scan-bar').classList.toggle('active', on);
}

function showBanner(id) {
  document.getElementById(id).classList.add('visible');
}

function hideBanner(id) {
  document.getElementById(id).classList.remove('visible');
}

function esc(s) {
  return String(s)
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;');
}
