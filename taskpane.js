/* ============================================================
   EmailChainGuard v4 — taskpane.js
   - Welcome screen al primo avvio
   - Banner nativi Outlook (NotificationMessage) + pannello
   - Graph API per lettura conversazione completa
   - Pagina Settings con domini propri, mute mittenti, lingua
   - Multi-tenant (escludere domini multipli)
   ============================================================ */
'use strict';

// ────────────────────────────────────────────────────────────
//  CONFIG
// ────────────────────────────────────────────────────────────
const CONFIG = {
  BACKEND_URL:  'https://emailchainguard-backend.onrender.com',
  ECG_API_KEY:  'MalibuStacy25_5', // <-- sostituisci con la tua chiave
  AZURE_CLIENT_ID: 'f32f2bbe-8140-41f4-bb3b-8bddc8a3f495',
  AZURE_TENANT:    'common', // 'common' = multi-tenant. Usa il tenant ID per single-tenant
  SCOPES: ['Mail.Read', 'User.Read'],
  MAX_KNOWN_EMAILS: 1000,
  MAX_CONVERSATIONS: 200,
  FETCH_TIMEOUT_MS: 15000,
};

// Storage keys
const KEY_WELCOME_SEEN  = 'ecg_welcome_seen_v4';
const KEY_KNOWN_SENDERS = 'ecg_known_senders_v4';
const KEY_KNOWN_CC      = 'ecg_known_cc_v4';
const KEY_OWN_DOMAINS   = 'ecg_own_domains_v4';
const KEY_MUTED_SENDERS = 'ecg_muted_senders_v4';
const KEY_LANG          = 'ecg_lang_v4';
const KEY_GRAPH_ENABLED = 'ecg_graph_enabled_v4';

// Cache memoria
const _cache = {
  senders: new Set(),
  cc: {},
  ownDomains: new Set(),
  muted: new Set(),
};

// Stato runtime
let _state = {
  lang: 'it',
  graphEnabled: false,
  graphToken: null,
  currentSender: null,
};

// Traduzioni
const I18N = {
  it: {
    'analyzing': 'Analisi in corso...',
    'all_safe': 'Email sicura',
    'domains_analyzed': 'domini analizzati',
    'domain_count': 'domini',
    'mute_sender': 'Mai più avvisarmi',
    'reset_done': 'Memoria cancellata',
    'graph_active': 'attivata',
    'graph_inactive': 'non attivata',
    'graph_enabling': 'autorizzazione in corso...',
    'graph_error': 'errore autorizzazione',
    'add_domain': 'Aggiungi',
    'no_muted': 'Nessun mittente silenziato',
    'no_own_domains': 'Nessun dominio configurato',
  },
  en: {
    'analyzing': 'Analyzing...',
    'all_safe': 'Email is safe',
    'domains_analyzed': 'domains analyzed',
    'domain_count': 'domains',
    'mute_sender': 'Never alert me again',
    'reset_done': 'Memory cleared',
    'graph_active': 'enabled',
    'graph_inactive': 'not enabled',
    'graph_enabling': 'authorizing...',
    'graph_error': 'authorization error',
    'add_domain': 'Add',
    'no_muted': 'No muted senders',
    'no_own_domains': 'No domains configured',
  },
};
function t(key) { return I18N[_state.lang]?.[key] || I18N.it[key] || key; }

// ────────────────────────────────────────────────────────────
//  STORAGE
// ────────────────────────────────────────────────────────────
function _roamingOk() {
  try { return !!(Office?.context?.roamingSettings); }
  catch { return false; }
}

function _storageGet(key) {
  if (_roamingOk()) {
    try { const v = Office.context.roamingSettings.get(key); if (v != null) return v; } catch {}
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

function _storageRemove(key) {
  if (_roamingOk()) {
    try {
      Office.context.roamingSettings.remove(key);
      Office.context.roamingSettings.saveAsync(() => {});
    } catch {}
  }
  try { localStorage.removeItem(key); } catch {}
}

// Mittenti conosciuti
function loadKnownSenders() {
  const raw = _storageGet(KEY_KNOWN_SENDERS);
  const stored = raw ? new Set(JSON.parse(raw)) : new Set();
  _cache.senders.forEach(e => stored.add(e));
  return stored;
}

function saveKnownSenders(emails) {
  emails.forEach(e => _cache.senders.add(e));
  const all = loadKnownSenders();
  emails.forEach(e => all.add(e));
  let arr = Array.from(all);
  if (arr.length > CONFIG.MAX_KNOWN_EMAILS) arr = arr.slice(arr.length - CONFIG.MAX_KNOWN_EMAILS);
  _storageSet(KEY_KNOWN_SENDERS, JSON.stringify(arr));
}

// CC per conversazione
function loadKnownCC() {
  const raw = _storageGet(KEY_KNOWN_CC);
  return raw ? JSON.parse(raw) : {};
}
function getKnownCCForConversation(convId) {
  const all = loadKnownCC();
  return new Set(all[convId] || []);
}
function saveKnownCCForConversation(convId, emails) {
  if (!_cache.cc[convId]) _cache.cc[convId] = new Set();
  emails.forEach(e => _cache.cc[convId].add(e));
  const all = loadKnownCC();
  const existing = new Set(all[convId] || []);
  emails.forEach(e => existing.add(e));
  all[convId] = Array.from(existing);
  const keys = Object.keys(all);
  if (keys.length > CONFIG.MAX_CONVERSATIONS) delete all[keys[0]];
  _storageSet(KEY_KNOWN_CC, JSON.stringify(all));
}

// Domini propri
function loadOwnDomains() {
  const raw = _storageGet(KEY_OWN_DOMAINS);
  const stored = raw ? new Set(JSON.parse(raw)) : new Set();
  _cache.ownDomains.forEach(d => stored.add(d));
  return stored;
}
function saveOwnDomains(set) {
  _cache.ownDomains = set;
  _storageSet(KEY_OWN_DOMAINS, JSON.stringify(Array.from(set)));
}

// Mittenti silenziati
function loadMuted() {
  const raw = _storageGet(KEY_MUTED_SENDERS);
  const stored = raw ? new Set(JSON.parse(raw)) : new Set();
  _cache.muted.forEach(e => stored.add(e));
  return stored;
}
function saveMuted(set) {
  _cache.muted = set;
  _storageSet(KEY_MUTED_SENDERS, JSON.stringify(Array.from(set)));
}

// ────────────────────────────────────────────────────────────
//  Graph API tramite Office Dialog (OAuth implicit flow)
// ────────────────────────────────────────────────────────────
const TOKEN_KEY = 'ecg_graph_token_v4';
const TOKEN_EXP_KEY = 'ecg_graph_token_exp_v4';

function _saveToken(token, expiresInSec) {
  const expAt = Date.now() + (expiresInSec * 1000) - 60000; // 60s di margine
  _storageSet(TOKEN_KEY, token);
  _storageSet(TOKEN_EXP_KEY, String(expAt));
  _state.graphToken = token;
}

function _loadToken() {
  if (_state.graphToken) {
    const exp = parseInt(_storageGet(TOKEN_EXP_KEY) || '0', 10);
    if (Date.now() < exp) return _state.graphToken;
  }
  const token = _storageGet(TOKEN_KEY);
  const exp = parseInt(_storageGet(TOKEN_EXP_KEY) || '0', 10);
  if (token && Date.now() < exp) {
    _state.graphToken = token;
    return token;
  }
  return null;
}

function _buildAuthUrl() {
  const params = new URLSearchParams({
    client_id: CONFIG.AZURE_CLIENT_ID,
    response_type: 'token',
    redirect_uri: 'https://pier-coder.github.io/emailchainguard-frontend/auth-callback.html',
    scope: CONFIG.SCOPES.join(' '),
    response_mode: 'fragment',
    prompt: 'select_account',
    nonce: Math.random().toString(36).slice(2),
  });
  return `https://login.microsoftonline.com/${CONFIG.AZURE_TENANT}/oauth2/v2.0/authorize?${params.toString()}`;
}

function enableGraph() {
  return new Promise((resolve, reject) => {
    if (!Office?.context?.ui?.displayDialogAsync) {
      reject(new Error('Office Dialog API non disponibile'));
      return;
    }
    // Apri la nostra pagina ponte che poi reindirizza a Microsoft
    // (Office Dialog non può aprire login.microsoftonline.com direttamente)
    const url = 'https://pier-coder.github.io/emailchainguard-frontend/auth-callback.html';
    Office.context.ui.displayDialogAsync(url, { height: 60, width: 30, promptBeforeOpen: false }, function(asyncResult) {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error(asyncResult.error?.message || 'Apertura dialog non riuscita'));
        return;
      }
      const dialog = asyncResult.value;
      let resolved = false;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
        try {
          const data = JSON.parse(arg.message);
          if (data.access_token) {
            _saveToken(data.access_token, parseInt(data.expires_in || '3600', 10));
            _state.graphEnabled = true;
            _storageSet(KEY_GRAPH_ENABLED, '1');
            resolved = true;
            dialog.close();
            resolve(true);
          } else if (data.error) {
            resolved = true;
            dialog.close();
            reject(new Error(data.error_description || data.error));
          }
        } catch (e) {
          resolved = true;
          dialog.close();
          reject(e);
        }
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, function(arg) {
        if (!resolved) reject(new Error('Dialog chiuso prima del completamento'));
      });
    });
  });
}

async function getGraphToken() {
  return _loadToken();
}

async function fetchConversationCC(conversationId, currentCreatedISO) {
  // Restituisce { ccs: Set, hasPrior: boolean }
  // hasPrior = c'erano email precedenti nella conversazione (anche se senza CC)
  const token = await getGraphToken();
  if (!token) return null;
  try {
    let filter = `conversationId eq '${conversationId.replace(/'/g, "''")}'`;
    if (currentCreatedISO) {
      filter += ` and receivedDateTime lt ${currentCreatedISO}`;
    }
    const url = `https://graph.microsoft.com/v1.0/me/messages?$filter=${encodeURIComponent(filter)}&$select=ccRecipients,receivedDateTime&$orderby=receivedDateTime asc&$top=50`;
    const resp = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!resp.ok) return null;
    const data = await resp.json();
    const messages = data.value || [];
    const ccs = new Set();
    messages.forEach(msg => {
      (msg.ccRecipients || []).forEach(r => {
        const a = r.emailAddress?.address?.toLowerCase();
        if (a) ccs.add(a);
      });
    });
    try { console.log('[ECG] Graph: ', messages.length, 'email precedenti,', ccs.size, 'CC unici'); } catch {}
    return { ccs, hasPrior: messages.length > 0 };
  } catch (e) {
    try { console.log('[ECG] Graph error:', e.message); } catch {}
    return null;
  }
}

// ────────────────────────────────────────────────────────────
//  OFFICE INIT
// ────────────────────────────────────────────────────────────
Office.onReady(info => {
  if (info.host !== Office.HostType.Outlook) return;
  loadSettings();
  setupUI();
  if (!_storageGet(KEY_WELCOME_SEEN)) {
    showScreen('welcome');
  } else {
    runScan();
  }
  if (Office.context.mailbox.addHandlerAsync) {
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      () => { if (_currentScreen === 'analysis' || _currentScreen === 'idle') { resetAnalysis(); runScan(); } }
    );
  }
});

function loadSettings() {
  _state.lang = _storageGet(KEY_LANG) || 'it';
  _state.graphEnabled = _storageGet(KEY_GRAPH_ENABLED) === '1';
}

// ────────────────────────────────────────────────────────────
//  UI / SCREEN MANAGEMENT
// ────────────────────────────────────────────────────────────
let _currentScreen = 'welcome';

function showScreen(name) {
  _currentScreen = name;
  ['welcome','idle','settings'].forEach(s => {
    const el = document.getElementById(s);
    if (el) el.classList.remove('visible');
  });
  document.getElementById('domains').classList.remove('visible');
  ['banner-new','banner-cc','banner-danger'].forEach(b => {
    document.getElementById(b).classList.remove('visible');
  });
  document.getElementById('advice').classList.remove('visible');

  if (name === 'welcome') document.getElementById('welcome').classList.add('visible');
  if (name === 'idle')    document.getElementById('idle').classList.add('visible');
  if (name === 'settings') {
    document.getElementById('settings').classList.add('visible');
    renderSettings();
  }
}

function setupUI() {
  // Welcome
  document.getElementById('btn-welcome-start').addEventListener('click', () => {
    _storageSet(KEY_WELCOME_SEEN, '1');
    runScan();
  });
  document.getElementById('btn-welcome-graph').addEventListener('click', async () => {
    try {
      await enableGraph();
      _storageSet(KEY_WELCOME_SEEN, '1');
      runScan();
    } catch (e) {
      alert('Autorizzazione non completata. Puoi attivarla in seguito da Impostazioni.');
    }
  });
  // Settings
  document.getElementById('btn-settings').addEventListener('click', () => showScreen('settings'));
  document.getElementById('btn-back').addEventListener('click', () => runScan());
  document.querySelectorAll('.lang-btn').forEach(b => {
    b.addEventListener('click', () => {
      _state.lang = b.dataset.lang;
      _storageSet(KEY_LANG, _state.lang);
      renderSettings();
    });
  });
  document.getElementById('btn-add-domain').addEventListener('click', addOwnDomain);
  document.getElementById('own-domain-input').addEventListener('keydown', e => {
    if (e.key === 'Enter') addOwnDomain();
  });
  document.getElementById('btn-reset-memory').addEventListener('click', resetMemory);
  document.getElementById('btn-graph-toggle').addEventListener('click', toggleGraph);
  // Mute
  document.getElementById('btn-mute-sender').addEventListener('click', muteSender);

  document.getElementById('info-backend').textContent = CONFIG.BACKEND_URL.replace('https://','').split('.')[0];
}

// ────────────────────────────────────────────────────────────
//  ANALYSIS
// ────────────────────────────────────────────────────────────
function resetAnalysis() {
  ['banner-new','banner-cc','banner-danger'].forEach(b =>
    document.getElementById(b).classList.remove('visible')
  );
  document.getElementById('advice').classList.remove('visible');
  document.getElementById('domains').classList.remove('visible');
  document.getElementById('domain-list').innerHTML = '';
  setDot('');
  setNativeBanner(null);
}

async function runScan() {
  const item = Office.context.mailbox.item;
  if (!item) return;

  showScreen('idle');
  document.getElementById('idle').classList.remove('visible');
  setDot('scanning');
  setScanBar(true);
  document.getElementById('foot-status').textContent = t('analyzing');

  try {
    const myEmail = Office.context.mailbox.userProfile?.emailAddress?.toLowerCase() || null;
    const fromAddr = item.from?.emailAddress?.toLowerCase() || null;
    const toAddrs  = (item.to || []).map(a => a.emailAddress?.toLowerCase()).filter(Boolean);
    const ccAddrs  = (item.cc || []).map(a => a.emailAddress?.toLowerCase()).filter(Boolean);
    const convId   = item.conversationId || null;

    _state.currentSender = fromAddr;

    const ownDomains = loadOwnDomains();
    if (myEmail) {
      const myDom = myEmail.split('@')[1];
      ownDomains.add(myDom);
    }

    const allAddrs = [fromAddr, ...toAddrs, ...ccAddrs]
      .filter(Boolean).filter(e => e !== myEmail);

    // 1. Primo contatto
    const knownSenders = loadKnownSenders();
    const muted = loadMuted();
    const isNewSender = fromAddr && fromAddr !== myEmail
      && !knownSenders.has(fromAddr) && !muted.has(fromAddr);
    saveKnownSenders(allAddrs);

    // 2. Nuovo CC
    let newCCAddrs = [];
    const ccFiltrati = ccAddrs.filter(e => e !== myEmail);
    let _ccDebug = `cc=${ccFiltrati.length}`;
    if (convId && ccFiltrati.length > 0) {
      let knownCC = null;
      let hasPriorEmails = false;

      if (_state.graphEnabled) {
        let createdISO = null;
        try {
          if (item.dateTimeCreated) createdISO = new Date(item.dateTimeCreated).toISOString();
        } catch {}
        _ccDebug += ` graphOn iso=${createdISO ? 'sì' : 'no'}`;
        const graphResult = await fetchConversationCC(convId, createdISO);
        if (graphResult) {
          knownCC = graphResult.ccs;
          hasPriorEmails = graphResult.hasPrior;
          _ccDebug += ` prior=${hasPriorEmails ? 'sì' : 'no'} knownCC=${knownCC.size}`;
        } else {
          _ccDebug += ' graph=null';
        }
      } else {
        _ccDebug += ' graphOff';
      }
      // Fallback storage locale
      if (!knownCC) {
        knownCC = getKnownCCForConversation(convId);
        hasPriorEmails = knownCC.size > 0;
        _ccDebug += ` fallback=${knownCC.size}`;
      }

      // Confronta solo se ci sono email precedenti (con o senza CC)
      if (hasPriorEmails) {
        newCCAddrs = ccFiltrati.filter(e => !knownCC.has(e));
        _ccDebug += ` new=${newCCAddrs.length}`;
      } else {
        _ccDebug += ' skip(noPrior)';
      }
      saveKnownCCForConversation(convId, ccFiltrati);
    } else if (convId) {
      saveKnownCCForConversation(convId, ccFiltrati);
    }
    // Memorizza il debug per il footer
    _state.lastCCDebug = _ccDebug;

    // 3. Domini esterni
    const domains = [...new Set(
      allAddrs.map(a => a.split('@')[1]).filter(Boolean)
        .filter(d => !ownDomains.has(d))
    )];

    let result = { overall_label: 'ok', suspect_count: 0, domains: [] };
    if (domains.length > 0) {
      result = await callBackend(domains, convId);
      if (!result || !Array.isArray(result.domains)) throw new Error('Risposta backend non valida');
    }

    renderResults(result, isNewSender ? fromAddr : null, newCCAddrs);
    showNativeBanner(result.overall_label, isNewSender ? fromAddr : null, newCCAddrs);

  } catch (err) {
    setDot('');
    document.getElementById('foot-status').textContent = 'Errore: ' + (err.message || 'analisi non riuscita');
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
      throw new Error(err.detail || `Backend HTTP ${resp.status}`);
    }
    return resp.json();
  } catch (err) {
    if (err.name === 'AbortError') throw new Error('Analisi scaduta (15s)');
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

// ────────────────────────────────────────────────────────────
//  RENDERING
// ────────────────────────────────────────────────────────────
function renderResults(data, newSenderEmail, newCCAddrs) {
  const { overall_label, domains } = data;

  // Status dot
  if (overall_label === 'danger')       setDot('danger');
  else if (overall_label === 'warning') setDot('warning');
  else if (newCCAddrs.length > 0)       setDot('warning');
  else if (newSenderEmail)              setDot('new');
  else                                  setDot('');

  let hasContent = false;

  if (newSenderEmail) {
    document.getElementById('banner-new-email').textContent = newSenderEmail;
    document.getElementById('banner-new').classList.add('visible');
    hasContent = true;
  }

  if (newCCAddrs.length > 0) {
    document.getElementById('banner-cc-detail').textContent = newCCAddrs.join(', ');
    document.getElementById('banner-cc').classList.add('visible');
    hasContent = true;
  }

  if (overall_label === 'danger' || overall_label === 'warning') {
    const details = domains.filter(d => d.is_suspect)
      .map(d => d.similar_to ? `@${d.domain} (simile a @${d.similar_to})` : `@${d.domain}`)
      .join(', ');
    document.getElementById('banner-danger-detail').textContent = details;
    document.getElementById('banner-danger').classList.add('visible');
    document.getElementById('advice').classList.add('visible');
    hasContent = true;
  }

  if (domains.length > 0) {
    const list = document.getElementById('domain-list');
    list.innerHTML = '';
    domains.forEach((d, i) => {
      const card = buildDomainCard(d);
      card.style.animationDelay = `${i * 50}ms`;
      list.appendChild(card);
    });
    document.getElementById('domains').classList.add('visible');
    hasContent = true;
  }

  if (!hasContent) {
    document.getElementById('idle').classList.add('visible');
    document.getElementById('idle-meta').textContent = `${domains.length} ${t('domain_count')}`;
  } else {
    document.getElementById('idle').classList.remove('visible');
  }

  document.getElementById('foot-status').textContent = `${domains.length} ${t('domains_analyzed')} · ${_state.lastCCDebug || ''}`;
  _currentScreen = 'analysis';
}

function buildDomainCard(d) {
  const cardClass = d.is_suspect
    ? (d.risk_label === 'danger' ? 'suspect' : 'warning-card') : 'safe';
  const badgeText = d.is_suspect
    ? (d.risk_label === 'danger' ? 'PERICOLO' : 'ATTENZIONE') : 'OK';
  const scoreColor = d.risk_label === 'danger' ? 'var(--danger)'
    : d.risk_label === 'warning' ? 'var(--warn)' : 'var(--ok)';

  const card = document.createElement('div');
  card.className = `dcard ${cardClass}`;
  const head = document.createElement('div');
  head.className = 'dcard-head';
  head.innerHTML = `
    <div class="dc-dot"></div>
    <div class="dc-domain">@${esc(d.domain)}</div>
    ${d.is_suspect ? `<div class="dc-score" style="color:${scoreColor}">${Number(d.risk_score)||0}</div>` : ''}
    <div class="dc-badge">${badgeText}</div>
    ${d.is_suspect ? '<div class="dc-chevron"><svg viewBox="0 0 24 24"><path d="M7 10l5 5 5-5z"/></svg></div>' : ''}
  `;
  card.appendChild(head);

  if (d.is_suspect) {
    const detail = document.createElement('div');
    detail.className = 'dcard-detail';
    detail.innerHTML = buildDetailHTML(d);
    card.appendChild(detail);
    head.addEventListener('click', () => card.classList.toggle('open'));
  }
  return card;
}

function buildDetailHTML(d) {
  let h = '<div class="dcard-detail-inner">';

  if (d.similar_to) {
    const { a, b } = buildDiff(d.similar_to, d.domain);
    h += `<div class="diff-block">
      <div class="detail-title">Confronto carattere per carattere</div>
      <div class="diff-pair">
        <div class="diff-line legit"><span class="diff-lbl">LEGIT</span><span class="diff-chars">${a}</span></div>
        <div class="diff-line fake"><span class="diff-lbl">FAKE?</span><span class="diff-chars">${b}</span></div>
      </div></div>`;
  }
  if (d.whois) {
    const w = d.whois; const ac = w.risk_flag ? 'danger' : 'ok';
    h += `<div class="detail-block"><div class="detail-title">WHOIS</div>
      <div class="detail-row"><span class="detail-key">Registrato</span><span class="detail-val ${ac}">${w.creation_date ? new Date(w.creation_date).toLocaleDateString('it-IT') : '—'}</span></div>
      <div class="detail-row"><span class="detail-key">Eta</span><span class="detail-val ${ac}">${esc(w.age_label||'—')}</span></div>
      <div class="detail-row"><span class="detail-key">Registrar</span><span class="detail-val">${esc(w.registrar||'—')}</span></div>
    </div>`;
  }
  if (d.dns) {
    const dns = d.dns;
    h += `<div class="detail-block"><div class="detail-title">DNS</div>
      <div class="detail-row"><span class="detail-key">MX</span><span class="detail-val ${dns.has_mx?'ok':'danger'}">${dns.has_mx?'Presente':'Assente'}</span></div>
      <div class="detail-row"><span class="detail-key">SPF</span><span class="detail-val ${dns.has_spf?'ok':'danger'}">${dns.has_spf?'Presente':'Assente'}</span></div>
      <div class="detail-row"><span class="detail-key">DMARC</span><span class="detail-val ${dns.has_dmarc?'ok':'danger'}">${dns.has_dmarc?'Presente':'Assente'}</span></div>
    </div>`;
  }
  if (d.reputation) {
    const rep = d.reputation;
    const vtMal = Number(rep.vt_malicious)||0, vtSus = Number(rep.vt_suspicious)||0;
    const vtc = vtMal > 0 ? 'danger' : vtSus > 0 ? 'warn' : 'ok';
    h += `<div class="detail-block"><div class="detail-title">Reputazione</div>`;
    if (rep.vt_available) {
      h += `<div class="detail-row"><span class="detail-key">VirusTotal</span><span class="detail-val ${vtc}">${vtMal} malevoli · ${vtSus} sospetti</span></div>`;
      if (rep.vt_link && String(rep.vt_link).startsWith('https://www.virustotal.com/'))
        h += `<div class="detail-row"><span class="detail-key">Report</span><span class="detail-val"><a href="${esc(rep.vt_link)}" target="_blank" rel="noopener noreferrer">Apri</a></span></div>`;
    } else {
      h += `<div class="detail-row"><span class="detail-key">VirusTotal</span><span class="detail-val" style="color:var(--muted)">non configurata</span></div>`;
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

// ────────────────────────────────────────────────────────────
//  NATIVE BANNER (Outlook NotificationMessage)
// ────────────────────────────────────────────────────────────
function setNativeBanner(detail) {
  try {
    const nm = Office.context.mailbox.item?.notificationMessages;
    if (!nm) return;
    if (!detail) {
      nm.removeAsync('ecg-status', () => {});
      return;
    }
    nm.replaceAsync('ecg-status', detail, () => {});
  } catch {}
}

function showNativeBanner(overall, newSenderEmail, newCCAddrs) {
  // Priorità: danger > cc > new
  if (overall === 'danger') {
    setNativeBanner({
      type: 'errorMessage',
      message: 'EmailChainGuard: dominio sospetto rilevato. Verifica il mittente.',
    });
  } else if (newCCAddrs.length > 0) {
    setNativeBanner({
      type: 'informationalMessage',
      message: `EmailChainGuard: nuovo CC nella conversazione (${newCCAddrs[0]})`,
      icon: 'Icon.16x16',
      persistent: false,
    });
  } else if (newSenderEmail) {
    setNativeBanner({
      type: 'informationalMessage',
      message: `EmailChainGuard: primo contatto da ${newSenderEmail}`,
      icon: 'Icon.16x16',
      persistent: false,
    });
  } else {
    setNativeBanner(null);
  }
}

// ────────────────────────────────────────────────────────────
//  SETTINGS
// ────────────────────────────────────────────────────────────
function renderSettings() {
  // Lang
  document.querySelectorAll('.lang-btn').forEach(b => {
    b.classList.toggle('active', b.dataset.lang === _state.lang);
  });
  // Domini propri
  const ownEl = document.getElementById('own-domains');
  ownEl.innerHTML = '';
  const own = loadOwnDomains();
  if (own.size === 0) {
    ownEl.innerHTML = `<div style="font-size:10px;color:var(--muted)">${t('no_own_domains')}</div>`;
  } else {
    own.forEach(d => {
      const tag = document.createElement('span');
      tag.className = 'domain-tag';
      tag.innerHTML = `${esc(d)}<span class="domain-tag-x" data-domain="${esc(d)}">×</span>`;
      tag.querySelector('.domain-tag-x').addEventListener('click', () => {
        const set = loadOwnDomains();
        set.delete(d);
        saveOwnDomains(set);
        renderSettings();
      });
      ownEl.appendChild(tag);
    });
  }
  // Mittenti silenziati
  const muteEl = document.getElementById('muted-senders');
  muteEl.innerHTML = '';
  const muted = loadMuted();
  if (muted.size === 0) {
    muteEl.innerHTML = `<div style="font-size:10px;color:var(--muted)">${t('no_muted')}</div>`;
  } else {
    muted.forEach(e => {
      const tag = document.createElement('span');
      tag.className = 'domain-tag';
      tag.innerHTML = `${esc(e)}<span class="domain-tag-x">×</span>`;
      tag.querySelector('.domain-tag-x').addEventListener('click', () => {
        const set = loadMuted();
        set.delete(e);
        saveMuted(set);
        renderSettings();
      });
      muteEl.appendChild(tag);
    });
  }
  // Graph status
  document.getElementById('graph-status').textContent = _state.graphEnabled ? t('graph_active') : t('graph_inactive');
}

function addOwnDomain() {
  const input = document.getElementById('own-domain-input');
  const val = input.value.trim().toLowerCase().replace(/^@/, '');
  if (!val.match(/^[a-z0-9]([a-z0-9\-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9\-]*[a-z0-9])?)+$/)) return;
  const set = loadOwnDomains();
  set.add(val);
  saveOwnDomains(set);
  input.value = '';
  renderSettings();
}

function muteSender() {
  if (!_state.currentSender) return;
  const set = loadMuted();
  set.add(_state.currentSender);
  saveMuted(set);
  document.getElementById('banner-new').classList.remove('visible');
}

function resetMemory() {
  _cache.senders.clear();
  _cache.cc = {};
  _storageRemove(KEY_KNOWN_SENDERS);
  _storageRemove(KEY_KNOWN_CC);
  document.getElementById('foot-status').textContent = t('reset_done');
}

async function toggleGraph() {
  if (_state.graphEnabled) {
    _state.graphEnabled = false;
    _storageRemove(KEY_GRAPH_ENABLED);
    renderSettings();
  } else {
    document.getElementById('graph-status').textContent = t('graph_enabling');
    try {
      await enableGraph();
      renderSettings();
    } catch {
      document.getElementById('graph-status').textContent = t('graph_error');
    }
  }
}

// ────────────────────────────────────────────────────────────
//  HELPERS
// ────────────────────────────────────────────────────────────
function setDot(state) {
  document.getElementById('status-dot').className = 'status-dot' + (state ? ' ' + state : '');
}
function setScanBar(on) {
  document.getElementById('scan-bar').classList.toggle('active', on);
}
function esc(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
