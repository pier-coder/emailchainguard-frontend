/* ============================================================
   EmailChainGuard — sentry-init.js
   Wrapper di Sentry browser SDK con scrubbing aggressivo del PII.
   - Si attiva solo se ECGSentry.init() viene chiamato (default ON
     dal taskpane se consent !== 'denied').
   - Tutti gli event e breadcrumb passano per beforeSend / beforeBreadcrumb
     che strippano email, JWT, Bearer tokens, Graph IDs, e droppano
     interamente i request body di fetch/xhr.
   - Integration potenzialmente invasive (Replay, BrowserTracing,
     BrowserProfiling, CaptureConsole) sono filtrate via integrations().
   - URL di flow OAuth (login.microsoftonline.com, auth-callback.html)
     sono in denyUrls e qualunque event/breadcrumb relativo viene droppato.
   ============================================================ */
'use strict';

(function () {
  // ---- CONFIG ----
  // DSN pubblico (esposto al browser by design; non e' un segreto).
  // Region EU (Francoforte) — verificato dal segmento .ingest.de.sentry.io.
  const DSN         = 'https://14043c9ec1523e7ce73dc305cc42b874@o4511418146750464.ingest.de.sentry.io/4511418174734416';
  const RELEASE     = 'emailchainguard@4.0.0';
  const ENVIRONMENT = 'production';

  // Marker leggibili in dashboard al posto dei valori scrubati
  const REDACTED_EMAIL  = '[redacted-email]';
  const REDACTED_TOKEN  = '[redacted-token]';
  const REDACTED_BEARER = 'Bearer [redacted-token]';
  const REDACTED_GRAPH  = '[redacted-graph-id]';
  const REDACTED_CONV   = '[redacted-conv-id]';
  const REDACTED_FILTER = '[redacted-filter]';
  const REDACTED_ID     = '[redacted-id]';

  // Regex centralizzate (riutilizzate da event + breadcrumb).
  // Ordine di applicazione in scrubString: piu' specifiche prima delle generiche.
  const EMAIL_RE     = /[\w.+-]+@[\w-]+(?:\.[\w-]+)+/g;
  const JWT_RE       = /eyJ[\w-]+\.[\w-]+\.[\w-]+/g;
  const BEARER_RE    = /Bearer\s+[\w.\-+/=]+/gi;
  // Copre tre forme: "conversationId eq 'X'", "conversationId%20eq%20'X'"
  // (percent-encoded come appare nei breadcrumb fetch URL), "conversationId=X"
  const CONV_ID_RE   = /conversationId(?:\s+eq\s+|%20eq%20|=)'?[\w+/=%]+'?/gi;
  // Intero valore di un $filter OData (belt-and-suspenders: tronca tutto il filtro)
  const ODATA_FILTER_RE = /\$filter=[^&]+/g;
  const GRAPH_ID_RE  = /\b(messages|conversations|users|me)\/[A-Za-z0-9_\-=]+/g;
  // Token opachi base64-like lunghi (es. id message/conversation Exchange).
  // AGGRESSIVA: applicata SOLO agli URL dei breadcrumb (via scrubUrl), mai al
  // message principale, per non scrubare hash innocui (es. SHA Git) negli errori.
  const ID_BASE64_RE = /[A-Za-z0-9+/]{40,}={0,2}/g;

  // URL dei flussi OAuth: qualunque evento/breadcrumb che li tocca viene droppato
  const OAUTH_URL_RE = /login\.microsoftonline\.com|auth-callback\.html/i;

  // Scrub standard — usato ovunque (event message, exception, breadcrumb message,
  // valori dentro deepScrub). Ordine: piu' specifiche prima delle generiche.
  function scrubString(s) {
    if (typeof s !== 'string') return s;
    return s
      .replace(EMAIL_RE, REDACTED_EMAIL)                                 // 1
      .replace(JWT_RE, REDACTED_TOKEN)                                   // 2
      .replace(BEARER_RE, REDACTED_BEARER)                               // 3
      .replace(CONV_ID_RE, () => 'conversationId=' + REDACTED_CONV)      // 4
      .replace(ODATA_FILTER_RE, () => '$filter=' + REDACTED_FILTER)      // 5 (fn: niente interpretazione di $)
      .replace(GRAPH_ID_RE, (_m, seg) => seg + '/' + REDACTED_GRAPH);    // 6
  }

  // Scrub aggressivo per URL — usato SOLO su breadcrumb.data.url.
  // Applica scrubString (passi 1-6) poi ID_BASE64_RE (passo 7) che tronca
  // qualunque token opaco base64-like residuo nell'URL.
  function scrubUrl(s) {
    if (typeof s !== 'string') return s;
    return scrubString(s).replace(ID_BASE64_RE, REDACTED_ID);            // 7
  }

  function deepScrub(value, depth) {
    depth = depth || 0;
    if (depth > 6) return value;
    if (value == null) return value;
    if (typeof value === 'string') return scrubString(value);
    if (Array.isArray(value)) return value.map(v => deepScrub(v, depth + 1));
    if (typeof value === 'object') {
      const out = {};
      for (const k of Object.keys(value)) {
        // Drop totalmente Authorization e Cookie headers — niente redacted marker
        if (/^(authorization|cookie|set-cookie|x-ecg-key)$/i.test(k)) continue;
        out[k] = deepScrub(value[k], depth + 1);
      }
      return out;
    }
    return value;
  }

  function beforeSend(event) {
    try {
      // Drop completo se la request di contesto e' nel flusso OAuth
      if (event.request && event.request.url && OAUTH_URL_RE.test(event.request.url)) {
        return null;
      }
      if (event.message) event.message = scrubString(event.message);
      if (event.request) {
        if (event.request.url) event.request.url = scrubString(event.request.url);
        if (event.request.headers) event.request.headers = deepScrub(event.request.headers);
        if (event.request.data) event.request.data = deepScrub(event.request.data);
      }
      if (event.exception && event.exception.values) {
        event.exception.values = event.exception.values.map(v => {
          if (v.value) v.value = scrubString(v.value);
          return v;
        });
      }
      if (event.breadcrumbs) {
        event.breadcrumbs = event.breadcrumbs.map(b => {
          const nb = Object.assign({}, b);
          if (nb.message) nb.message = scrubString(nb.message);
          if (nb.data) nb.data = deepScrub(nb.data);
          return nb;
        });
      }
      if (event.extra)    event.extra    = deepScrub(event.extra);
      if (event.contexts) event.contexts = deepScrub(event.contexts);
      if (event.tags)     event.tags     = deepScrub(event.tags);
    } catch (_e) {
      // Se lo scrubbing fallisce, droppiamo l'evento per safety
      return null;
    }
    return event;
  }

  function beforeBreadcrumb(breadcrumb) {
    try {
      // Console breadcrumb: droppa interamente (potrebbero contenere
      // qualunque oggetto che gli sviluppatori abbiano log-ato)
      if (breadcrumb.category === 'console') return null;

      if (breadcrumb.data) {
        // Fetch/XHR: il body puo' contenere lista domini analizzati
        if (breadcrumb.category === 'fetch' || breadcrumb.category === 'xhr') {
          delete breadcrumb.data.body;
          delete breadcrumb.data.request_body_size;
          delete breadcrumb.data.response_body_size;
        }
        // URL nel flusso OAuth: droppa interamente il breadcrumb
        if (breadcrumb.data.url && OAUTH_URL_RE.test(breadcrumb.data.url)) return null;
        breadcrumb.data = deepScrub(breadcrumb.data);
        // URL: scrub aggressivo aggiuntivo (include ID_BASE64_RE), solo qui nei breadcrumb
        if (breadcrumb.data.url) breadcrumb.data.url = scrubUrl(breadcrumb.data.url);
      }
      if (breadcrumb.message) breadcrumb.message = scrubString(breadcrumb.message);
    } catch (_e) {
      return null;
    }
    return breadcrumb;
  }

  // Filtra le integration di default rimuovendo quelle potenzialmente invasive
  function filterIntegrations(defaultIntegrations) {
    const BLOCKED = new Set([
      'Replay',
      'ReplayCanvas',
      'BrowserTracing',
      'BrowserProfiling',
      'CaptureConsole',
    ]);
    return defaultIntegrations.filter(i => !BLOCKED.has(i.name));
  }

  // ---- API pubblica ----
  let _initialized = false;

  function init() {
    if (_initialized) return;
    if (typeof window.Sentry === 'undefined') return; // CDN non caricato (probabile SRI fail)
    try {
      window.Sentry.init({
        dsn:               DSN,
        release:           RELEASE,
        environment:       ENVIRONMENT,
        sendDefaultPii:    false,
        attachStacktrace:  true,
        autoSessionTracking: false,    // niente session tracking (privacy)
        tracesSampleRate:  0,          // belt-and-suspenders contro tracing
        denyUrls:          [OAUTH_URL_RE],
        integrations:      filterIntegrations,
        beforeSend:        beforeSend,
        beforeBreadcrumb:  beforeBreadcrumb,
        initialScope:      { tags: { component: 'taskpane' } },
      });
      _initialized = true;
    } catch (_e) {
      // Senza Sentry attivo non c'e' modo di loggare il fallimento — silenzio
    }
  }

  function captureException(err, context) {
    if (!_initialized || typeof window.Sentry === 'undefined') return;
    try {
      const opts = context ? { extra: deepScrub(context) } : undefined;
      window.Sentry.captureException(err, opts);
    } catch (_e) {}
  }

  function close() {
    if (!_initialized) return;
    try {
      if (window.Sentry && typeof window.Sentry.close === 'function') {
        // Flush in background; ignoriamo la Promise (cleanup best-effort)
        window.Sentry.close();
      }
    } catch (_e) {}
    _initialized = false;
  }

  window.ECGSentry = { init: init, captureException: captureException, close: close };
})();
