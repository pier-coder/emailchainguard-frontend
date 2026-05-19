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
  ECG_API_KEY:  'ecg-dev-key-2024', // <-- sostituisci con la tua chiave
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
const KEY_LANG          = 'ecg_lang_v4';
const KEY_GRAPH_ENABLED = 'ecg_graph_enabled_v4';

// Cache memoria
const _cache = {
  senders: new Set(),
  cc: {},
  ownDomains: new Set(),
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
    // Runtime
    'analyzing':         'Analisi in corso...',
    'domains_analyzed':  'domini analizzati',
    'domain_count':      'domini',
    'reset_done':        'Memoria cancellata',
    'graph_active':      'attivata',
    'graph_inactive':    'non attivata',
    'graph_enabling':    'autorizzazione in corso...',
    'graph_error':       'errore autorizzazione',
    'no_own_domains':    'Nessun dominio configurato',
    // Welcome
    'welcome_title':     'Benvenuto',
    'welcome_subtitle':  'EmailChainGuard protegge le tue conversazioni email da spoofing e impersonificazione',
    'feat_suspect_title':'Domini sospetti',
    'feat_suspect_desc': 'Rileva domini falsi che imitano quelli reali',
    'feat_new_title':    'Primo contatto',
    'feat_new_desc':     'Avvisa quando ricevi email da un mittente per la prima volta',
    'feat_cc_title':     'Nuovo CC nella conversazione',
    'feat_cc_desc':      'Segnala se qualcuno è stato aggiunto in CC durante la catena email',
    'btn_start':         'Inizia',
    'btn_enable_graph':  'Abilita lettura conversazione completa',
    'privacy_notice':    'I tuoi dati restano nel tuo profilo Outlook · Solo nomi di dominio vengono inviati al server di analisi',
    // Idle
    'idle_title':        'Email sicura',
    // Banner
    'banner_new_title':  'Primo contatto',
    'banner_new_sub':    'Non hai mai ricevuto email da questo mittente',
    'banner_cc_title':   'Nuovo partecipante in CC',
    'banner_cc_sub':     'Non era presente nelle email precedenti di questa conversazione',
    'banner_danger_title':'Dominio sospetto rilevato',
    'banner_danger_sub': 'Verifica l\'identità del mittente prima di rispondere',
    'advice_title':      'Come procedere',
    // Sections
    'sec_domains':       'Domini analizzati',
    'back_to_analysis':  'Torna all\'analisi',
    'set_language':      'Lingua',
    'set_own_domains':   'Domini della tua organizzazione',
    'set_own_desc':      'Domini esclusi dall\'analisi (es. la tua azienda)',
    'btn_add':           'Aggiungi',
    'set_memory':        'Memoria',
    'btn_reset':         'Cancella memoria mittenti',
    'set_graph':         'Lettura conversazione completa',
    'set_graph_desc':    'Permette al plugin di leggere le email precedenti per rilevare nuovi CC',
    'state_label':       'Stato',
    'set_info':          'Informazioni',
    'info_version':      'Versione',
    'info_backend':      'Backend',
    // Feedback & support
    'set_feedback':      'Segnalazioni e suggerimenti',
    'set_feedback_desc': 'Apre il tuo client email con un messaggio precompilato verso il nostro supporto',
    'btn_report_bug':    'Segnala un problema',
    'btn_suggest':       'Suggerisci una funzione',
    'mail_subject_bug':  '[EmailChainGuard] Segnalazione problema',
    'mail_subject_sug':  '[EmailChainGuard] Suggerimento',
    'mail_subject_contact': '[EmailChainGuard] Contatto',
    'mail_body_bug':     'Descrivi il problema riscontrato (cosa stavi facendo, cosa ti aspettavi, cosa è successo invece):',
    'mail_body_sug':     'Descrivi la funzione che vorresti vedere o come potremmo migliorare il plugin:',
    'mail_body_contact': 'Scrivi qui il tuo messaggio:',
    'mail_tech_label':   'Informazioni tecniche (non modificare):',
    'mail_tech_version': 'Versione plugin',
    'mail_tech_lang':    'Lingua',
    'mail_tech_platform':'Piattaforma',
    'mail_tech_graph':   'Graph attivo',
    'mail_tech_yes':     'sì',
    'mail_tech_no':      'no',
    'set_support':       'Supporto e contatti',
    'contact_email':     'Email',
    'contact_privacy':   'Privacy',
    'contact_terms':     'Termini',
    'link_privacy':      'Leggi la policy',
    'link_terms':        'Leggi i termini',
  },
  en: {
    'analyzing':         'Analyzing...',
    'domains_analyzed':  'domains analyzed',
    'domain_count':      'domains',
    'reset_done':        'Memory cleared',
    'graph_active':      'enabled',
    'graph_inactive':    'not enabled',
    'graph_enabling':    'authorizing...',
    'graph_error':       'authorization error',
    'no_own_domains':    'No domains configured',
    'welcome_title':     'Welcome',
    'welcome_subtitle':  'EmailChainGuard protects your email conversations from spoofing and impersonation',
    'feat_suspect_title':'Suspicious domains',
    'feat_suspect_desc': 'Detects fake domains mimicking real ones',
    'feat_new_title':    'First contact',
    'feat_new_desc':     'Alerts you when you receive an email from a sender for the first time',
    'feat_cc_title':     'New CC in the conversation',
    'feat_cc_desc':      'Flags when someone was added in CC during the email chain',
    'btn_start':         'Start',
    'btn_enable_graph':  'Enable full conversation reading',
    'privacy_notice':    'Your data stays in your Outlook profile · Only domain names are sent to the analysis server',
    'idle_title':        'Email is safe',
    'banner_new_title':  'First contact',
    'banner_new_sub':    'You have never received emails from this sender',
    'banner_cc_title':   'New CC participant',
    'banner_cc_sub':     'They were not in previous emails of this conversation',
    'banner_danger_title':'Suspicious domain detected',
    'banner_danger_sub': 'Verify the sender identity before replying',
    'advice_title':      'How to proceed',
    'sec_domains':       'Analyzed domains',
    'back_to_analysis':  'Back to analysis',
    'set_language':      'Language',
    'set_own_domains':   'Your organization domains',
    'set_own_desc':      'Domains excluded from analysis (e.g. your company)',
    'btn_add':           'Add',
    'set_memory':        'Memory',
    'btn_reset':         'Clear sender memory',
    'set_graph':         'Full conversation reading',
    'set_graph_desc':    'Allows the plugin to read previous emails to detect new CCs',
    'state_label':       'Status',
    'set_info':          'About',
    'info_version':      'Version',
    'info_backend':      'Backend',
    'set_feedback':      'Feedback and suggestions',
    'set_feedback_desc': 'Opens your email client with a pre-filled message to our support',
    'btn_report_bug':    'Report a problem',
    'btn_suggest':       'Suggest a feature',
    'mail_subject_bug':  '[EmailChainGuard] Bug report',
    'mail_subject_sug':  '[EmailChainGuard] Suggestion',
    'mail_subject_contact': '[EmailChainGuard] Contact',
    'mail_body_bug':     'Describe the issue you encountered (what you were doing, what you expected, what happened instead):',
    'mail_body_sug':     'Describe the feature you would like to see or how we could improve the plugin:',
    'mail_body_contact': 'Write your message here:',
    'mail_tech_label':   'Technical information (do not edit):',
    'mail_tech_version': 'Plugin version',
    'mail_tech_lang':    'Language',
    'mail_tech_platform':'Platform',
    'mail_tech_graph':   'Graph enabled',
    'mail_tech_yes':     'yes',
    'mail_tech_no':      'no',
    'set_support':       'Support and contacts',
    'contact_email':     'Email',
    'contact_privacy':   'Privacy',
    'contact_terms':     'Terms',
    'link_privacy':      'Read the policy',
    'link_terms':        'Read the terms',
  },
};

function applyI18n() {
  document.querySelectorAll('[data-i18n]').forEach(el => {
    const key = el.getAttribute('data-i18n');
    const txt = t(key);
    if (txt) el.textContent = txt;
  });
  // Aggiorna anche i link Privacy/Termini con la lingua corrente
  const fp = document.getElementById('foot-privacy');
  if (fp) fp.href = `https://pier-coder.github.io/emailchainguard-frontend/privacy.html?lang=${_state.lang}`;
  const lp = document.getElementById('link-privacy');
  if (lp) lp.href = `https://pier-coder.github.io/emailchainguard-frontend/privacy.html?lang=${_state.lang}`;
  const lt = document.getElementById('link-terms');
  if (lt) lt.href = `https://pier-coder.github.io/emailchainguard-frontend/terms.html?lang=${_state.lang}`;
}
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

// ────────────────────────────────────────────────────────────
//  Graph API tramite Office Dialog (OAuth implicit flow)
// ────────────────────────────────────────────────────────────
const TOKEN_KEY = 'ecg_graph_token_v4';
const TOKEN_EXP_KEY = 'ecg_graph_token_exp_v4';

function _saveToken(token, expiresInSec) {
  return new Promise((resolve) => {
    const expAt = Date.now() + (expiresInSec * 1000) - 60000;
    // Salva in memoria
    _state.graphToken = token;
    _state.graphTokenExp = expAt;
    // Salva in localStorage (sincrono, sempre disponibile come fallback)
    try {
      localStorage.setItem(TOKEN_KEY, token);
      localStorage.setItem(TOKEN_EXP_KEY, String(expAt));
    } catch {}
    // Salva in roamingSettings (asincrono - aspettiamo il completamento)
    if (_roamingOk()) {
      try {
        Office.context.roamingSettings.set(TOKEN_KEY, token);
        Office.context.roamingSettings.set(TOKEN_EXP_KEY, String(expAt));
        Office.context.roamingSettings.saveAsync((res) => {
          if (res.status !== Office.AsyncResultStatus.Succeeded) {
            _state.lastGraphError = 'roaming save: ' + (res.error?.message || 'fail');
          }
          resolve();
        });
      } catch (e) {
        _state.lastGraphError = 'roaming set err: ' + e.message;
        resolve();
      }
    } else {
      resolve();
    }
  });
}

function _loadToken() {
  // 1. Memoria runtime
  if (_state.graphToken && _state.graphTokenExp && Date.now() < _state.graphTokenExp) {
    return _state.graphToken;
  }
  // 2. roamingSettings (più affidabile)
  if (_roamingOk()) {
    try {
      const token = Office.context.roamingSettings.get(TOKEN_KEY);
      const exp = parseInt(Office.context.roamingSettings.get(TOKEN_EXP_KEY) || '0', 10);
      if (token && Date.now() < exp) {
        _state.graphToken = token;
        _state.graphTokenExp = exp;
        return token;
      }
      if (token && Date.now() >= exp) {
        _state.lastGraphError = 'token roaming scaduto';
        return null;
      }
    } catch (e) {
      _state.lastGraphError = 'roaming load err: ' + e.message;
    }
  }
  // 3. Fallback localStorage
  try {
    const token = localStorage.getItem(TOKEN_KEY);
    const exp = parseInt(localStorage.getItem(TOKEN_EXP_KEY) || '0', 10);
    if (token && Date.now() < exp) {
      _state.graphToken = token;
      _state.graphTokenExp = exp;
      return token;
    }
  } catch {}
  _state.lastGraphError = 'token assente in tutti gli storage';
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
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, async function(arg) {
        _state.lastGraphError = 'dialog msg ricevuto len=' + (arg.message ? arg.message.length : 0);
        try {
          const data = JSON.parse(arg.message);
          if (data.access_token) {
            _state.lastGraphError = 'token ricevuto, saving...';
            await _saveToken(data.access_token, parseInt(data.expires_in || '3600', 10));
            _state.graphEnabled = true;
            _storageSet(KEY_GRAPH_ENABLED, '1');
            _state.lastGraphError = 'token salvato OK';
            resolved = true;
            dialog.close();
            resolve(true);
          } else if (data.error) {
            _state.lastGraphError = 'dialog err: ' + (data.error_description || data.error).substring(0, 100);
            resolved = true;
            dialog.close();
            reject(new Error(data.error_description || data.error));
          } else {
            _state.lastGraphError = 'dialog msg senza token. keys=' + Object.keys(data).join(',');
          }
        } catch (e) {
          _state.lastGraphError = 'parse err: ' + e.message + ' | raw=' + (arg.message || '').substring(0, 80);
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
  let token = _loadToken();
  if (token) return token;
  // Token mancante o scaduto: prova refresh silenzioso (senza popup)
  if (_state.graphEnabled) {
    _state.lastGraphError = 'tentativo refresh silenzioso...';
    try {
      await _silentRefreshToken();
      token = _loadToken();
      if (token) return token;
    } catch (e) {
      _state.lastGraphError = 'refresh fallito: ' + (e.message || '?');
    }
  }
  return null;
}

function _silentRefreshToken() {
  return new Promise((resolve, reject) => {
    // Usa iframe nascosto invece di displayDialogAsync (che chiede consenso utente)
    const url = 'https://pier-coder.github.io/emailchainguard-frontend/auth-callback.html?silent=1';
    let iframe = null;
    let timer = null;
    let resolved = false;

    function cleanup() {
      try { if (iframe) iframe.remove(); } catch {}
      try { window.removeEventListener('message', onMessage); } catch {}
      if (timer) clearTimeout(timer);
    }

    function onMessage(ev) {
      if (resolved) return;
      // Accetta solo messaggi dal nostro dominio
      if (!ev.origin || !ev.origin.startsWith('https://pier-coder.github.io')) return;
      let data = ev.data;
      // Supporta sia oggetto diretto sia stringa JSON
      if (typeof data === 'string') {
        try { data = JSON.parse(data); } catch { return; }
      }
      if (!data) return;
      // Cerca i payload OAuth nei vari formati
      const payload = data.ecgAuth ? data.payload : data;
      const parsed = (typeof payload === 'string') ? (function(){ try{return JSON.parse(payload);}catch{return null;}})() : payload;
      if (!parsed) return;
      if (parsed.access_token) {
        resolved = true;
        cleanup();
        _saveToken(parsed.access_token, parseInt(parsed.expires_in || '3600', 10)).then(() => resolve(true));
      } else if (parsed.error) {
        resolved = true;
        cleanup();
        reject(new Error(parsed.error_description || parsed.error));
      }
    }

    timer = setTimeout(() => {
      if (!resolved) {
        resolved = true;
        cleanup();
        reject(new Error('refresh timeout (5s)'));
      }
    }, 5000);

    window.addEventListener('message', onMessage);

    iframe = document.createElement('iframe');
    iframe.style.cssText = 'display:none;width:0;height:0;border:0';
    iframe.src = url;
    document.body.appendChild(iframe);
  });
}

function _normalizeSubject(s) {
  if (!s) return '';
  let n = s.trim();
  // Rimuovi prefissi ricorsivi: Re:, RE:, R:, Fwd:, Fw:, F:, I:, AW:, ANTW:, TR:
  const prefixRe = /^(\s*(re|r|fwd|fw|f|i|aw|antw|tr)\s*:\s*)+/i;
  while (prefixRe.test(n)) {
    n = n.replace(prefixRe, '');
  }
  // Normalizza spazi multipli
  n = n.replace(/\s+/g, ' ').trim().toLowerCase();
  return n;
}

async function fetchConversationCC(conversationId, currentCreatedISO, currentSubject, currentMessageId) {
  // Restituisce { ccs: Set, hasPrior: boolean, source: 'convId'|'subject' }
  const token = await getGraphToken();
  if (!token) {
    _state.lastGraphError = 'token mancante';
    return null;
  }
  try {
    // Step 1: query per conversationId
    const filter = `conversationId eq '${conversationId.replace(/'/g, "''")}'`;
    const url = `https://graph.microsoft.com/v1.0/me/messages?$filter=${encodeURIComponent(filter)}&$select=ccRecipients,from,receivedDateTime,subject,id&$top=50`;
    const resp = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!resp.ok) {
      let errBody = '';
      try { errBody = (await resp.text()).substring(0, 200); } catch {}
      _state.lastGraphError = `HTTP${resp.status}: ${errBody}`;
      return null;
    }
    const data = await resp.json();
    let allMessages = data.value || [];

    // Filtra in JavaScript: tieni solo i messaggi PRECEDENTI a quello corrente
    const currentTime = currentCreatedISO ? new Date(currentCreatedISO).getTime() : null;
    let priorMessages = currentTime
      ? allMessages.filter(msg => {
          if (!msg.receivedDateTime) return false;
          return new Date(msg.receivedDateTime).getTime() < currentTime;
        })
      : allMessages.filter(msg => !currentMessageId || msg.id !== currentMessageId);

    let source = 'convId';

    // Step 2: fallback per SUBJECT normalizzato se conversationId non ha trovato thread
    if (priorMessages.length === 0 && currentSubject) {
      const normSubj = _normalizeSubject(currentSubject);
      if (normSubj.length >= 3) {
        // Recupera email recenti e filtra in JS per subject normalizzato
        // Limit alle ultime ~50 email per costi/velocità
        const subjUrl = `https://graph.microsoft.com/v1.0/me/messages?$select=ccRecipients,from,receivedDateTime,subject,id&$orderby=receivedDateTime desc&$top=50`;
        const subjResp = await fetch(subjUrl, { headers: { Authorization: `Bearer ${token}` } });
        if (subjResp.ok) {
          const subjData = await subjResp.json();
          const candidates = subjData.value || [];
          priorMessages = candidates.filter(msg => {
            if (msg.id === currentMessageId) return false;
            if (currentTime && msg.receivedDateTime) {
              if (new Date(msg.receivedDateTime).getTime() >= currentTime) return false;
            }
            return _normalizeSubject(msg.subject) === normSubj;
          });
          if (priorMessages.length > 0) source = 'subject';
        }
      }
    }

    const ccs = new Set();
    priorMessages.forEach(msg => {
      (msg.ccRecipients || []).forEach(r => {
        const a = r.emailAddress?.address?.toLowerCase();
        if (a) ccs.add(a);
      });
      const fromA = msg.from?.emailAddress?.address?.toLowerCase();
      if (fromA) ccs.add(fromA);
    });
    try { console.log('[ECG] Graph:', priorMessages.length, 'prior (' + source + '),', ccs.size, 'partecipanti'); } catch {}
    return { ccs, hasPrior: priorMessages.length > 0, source };
  } catch (e) {
    _state.lastGraphError = 'exception: ' + (e.message || e.toString()).substring(0, 200);
    return null;
  }
}

// ────────────────────────────────────────────────────────────
//  OFFICE INIT
// ────────────────────────────────────────────────────────────
Office.onReady(async info => {
  if (info.host !== Office.HostType.Outlook) return;
  loadSettings();
  applyI18n();
  setupUI();

  // Se Graph è attivo ma il token è scaduto, forza il refresh PRIMA della prima scansione
  if (_state.graphEnabled && !_state.graphToken) {
    try {
      await _silentRefreshToken();
      _state.bootTokenInfo = 'boot:refresh-OK';
    } catch (e) {
      _state.bootTokenInfo = 'boot:refresh-FAIL ' + (e.message || '').substring(0, 40);
    }
  }

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
  // Pre-carica il token: forza la lettura da roamingSettings/localStorage in memoria
  if (_state.graphEnabled) {
    const t = _loadToken();
    _state.bootTokenInfo = t ? 'boot:OK len=' + t.length : 'boot:NO ' + (_state.lastGraphError || '?');
  }
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
      applyI18n();
      renderSettings();
    });
  });
  document.getElementById('btn-add-domain').addEventListener('click', addOwnDomain);
  document.getElementById('own-domain-input').addEventListener('keydown', e => {
    if (e.key === 'Enter') addOwnDomain();
  });
  document.getElementById('btn-reset-memory').addEventListener('click', resetMemory);
  document.getElementById('btn-graph-toggle').addEventListener('click', toggleGraph);

  // Segnalazioni / suggerimenti / contatti — apre client email con mailto precompilato
  document.getElementById('btn-report-bug').addEventListener('click', () => openMailto('bug'));
  document.getElementById('btn-suggest').addEventListener('click', () => openMailto('suggestion'));
  document.getElementById('contact-email-link').addEventListener('click', (e) => {
    e.preventDefault();
    openMailto('contact');
  });

  document.getElementById('info-backend').textContent = CONFIG.BACKEND_URL.replace('https://','').split('.')[0];
}

function openMailto(type) {
  const supportEmail = 'support@emailchainguard.example'; // <-- sostituisci con l'indirizzo reale
  let subject, bodyIntro;

  if (type === 'bug') {
    subject = t('mail_subject_bug');
    bodyIntro = t('mail_body_bug') + '\n\n\n\n---\n';
  } else if (type === 'suggestion') {
    subject = t('mail_subject_sug');
    bodyIntro = t('mail_body_sug') + '\n\n\n\n---\n';
  } else {
    subject = t('mail_subject_contact');
    bodyIntro = t('mail_body_contact') + '\n\n\n\n---\n';
  }

  // Info tecniche
  let techInfo = t('mail_tech_label') + '\n';
  techInfo += `${t('mail_tech_version')}: 4.0.0\n`;
  techInfo += `${t('mail_tech_lang')}: ${_state.lang || 'it'}\n`;
  try {
    techInfo += `${t('mail_tech_platform')}: ${navigator.platform || 'n/d'}\n`;
    const ua = navigator.userAgent || '';
    techInfo += `User-Agent: ${ua.substring(0, 200)}\n`;
  } catch {}
  techInfo += `${t('mail_tech_graph')}: ${_state.graphEnabled ? t('mail_tech_yes') : t('mail_tech_no')}\n`;

  const body = bodyIntro + techInfo;
  const mailtoUrl = `mailto:${supportEmail}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;

  try {
    window.open(mailtoUrl, '_blank');
  } catch {
    window.location.href = mailtoUrl;
  }
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
    const isNewSender = fromAddr && fromAddr !== myEmail
      && !knownSenders.has(fromAddr);
    saveKnownSenders(allAddrs);

    // 2. Nuovo CC
    let newCCAddrs = [];
    // Considera CC + From (un nuovo From in una conversazione esistente è anche un caso da segnalare)
    const ccFiltrati = ccAddrs.filter(e => e !== myEmail);
    const fromForCheck = (fromAddr && fromAddr !== myEmail) ? [fromAddr] : [];
    const addrsToCheck = [...new Set([...ccFiltrati, ...fromForCheck])];
    let _ccDebug = `cc=${ccFiltrati.length}+from=${fromForCheck.length}`;
    if (convId && addrsToCheck.length > 0) {
      let knownCC = null;
      let hasPriorEmails = false;

      if (_state.graphEnabled) {
        let createdISO = null;
        try {
          if (item.dateTimeCreated) createdISO = new Date(item.dateTimeCreated).toISOString();
        } catch {}
        const subj = item.subject || '';
        const msgId = item.itemId || null;
        _ccDebug += ` graphOn iso=${createdISO ? 'sì' : 'no'}`;
        const graphResult = await fetchConversationCC(convId, createdISO, subj, msgId);
        if (graphResult) {
          knownCC = graphResult.ccs;
          hasPriorEmails = graphResult.hasPrior;
          _ccDebug += ` prior=${hasPriorEmails ? 'sì' : 'no'}(${graphResult.source}) knownCC=${knownCC.size}`;
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
        newCCAddrs = addrsToCheck.filter(e => !knownCC.has(e));
        _ccDebug += ` new=${newCCAddrs.length}`;
      } else {
        _ccDebug += ' skip(noPrior)';
      }
      saveKnownCCForConversation(convId, addrsToCheck);
    } else if (convId) {
      saveKnownCCForConversation(convId, addrsToCheck);
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

    // Logica raffinata: in conversazione esistente, "Nuovo partecipante" prevale su "Primo contatto"
    // (un attaccante in una thread esistente è un caso più rilevante del semplice nuovo mittente)
    let showNewSender = isNewSender;
    if (isNewSender && newCCAddrs.includes(fromAddr)) {
      // Il from è già coperto dal banner giallo "Nuovo partecipante"
      showNewSender = false;
    }

    renderResults(result, showNewSender ? fromAddr : null, newCCAddrs);
    showNativeBanner(result.overall_label, showNewSender ? fromAddr : null, newCCAddrs);

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

  const dbgExtra = _state.lastGraphError ? ` · ERR:${_state.lastGraphError}` : '';
  const bootInfo = _state.bootTokenInfo ? ` · ${_state.bootTokenInfo}` : '';
  document.getElementById('foot-status').textContent = `${domains.length} ${t('domains_analyzed')} · ${_state.lastCCDebug || ''}${dbgExtra}${bootInfo}`;
  _state.lastGraphError = null;
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
  // Aggiorna i link Privacy/Termini con la lingua corrente
  const linkP = document.getElementById('link-privacy');
  const linkT = document.getElementById('link-terms');
  if (linkP) linkP.href = `https://pier-coder.github.io/emailchainguard-frontend/privacy.html?lang=${_state.lang}`;
  if (linkT) linkT.href = `https://pier-coder.github.io/emailchainguard-frontend/terms.html?lang=${_state.lang}`;
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
    try { localStorage.removeItem(TOKEN_KEY); localStorage.removeItem(TOKEN_EXP_KEY); } catch {}
    _state.graphToken = null;
    _state.graphTokenExp = 0;
    renderSettings();
  } else {
    document.getElementById('graph-status').textContent = t('graph_enabling');
    try {
      await enableGraph();
      renderSettings();
      // Mostra dettaglio nel footer
      if (_state.lastGraphError) {
        document.getElementById('foot-status').textContent = 'AUTH: ' + _state.lastGraphError;
      }
    } catch (e) {
      const detail = _state.lastGraphError || e.message || 'unknown';
      document.getElementById('graph-status').textContent = t('graph_error') + ': ' + detail.substring(0, 60);
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
