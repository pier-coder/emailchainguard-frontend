/* ============================================================
   EmailChainGuard — analytics.js
   Caricamento dinamico opt-in di Umami Analytics.
   Lo script Umami viene iniettato solo se l'utente ha
   esplicitamente concesso il consenso (Impostazioni > Aiutaci a migliorare).
   track() è no-op finche' load() non e' stato chiamato con successo.
   ============================================================ */
'use strict';

(function () {
  // Identificatore pubblico del website su Umami Cloud.
  // Non e' un segreto: lo script Umami lo invia in chiaro al browser di ogni visitatore.
  const UMAMI_SCRIPT_URL = 'https://cloud.umami.is/script.js';
  const UMAMI_WEBSITE_ID = '7aa52fdf-c209-4ab2-9dad-3a1c258528c7';
  const SCRIPT_ID        = 'ecg-umami-script';

  function load() {
    if (document.getElementById(SCRIPT_ID)) return;
    const s = document.createElement('script');
    s.id = SCRIPT_ID;
    s.defer = true;
    s.src = UMAMI_SCRIPT_URL;
    s.setAttribute('data-website-id', UMAMI_WEBSITE_ID);
    // Disabilita auto-track delle pageview: il pannello non naviga,
    // tracciamo solo eventi espliciti tramite track().
    s.setAttribute('data-auto-track', 'false');
    document.head.appendChild(s);
  }

  function unload() {
    const s = document.getElementById(SCRIPT_ID);
    if (s) s.remove();
    try { delete window.umami; } catch { window.umami = undefined; }
  }

  function track(event, props) {
    try {
      if (window.umami && typeof window.umami.track === 'function') {
        if (props && Object.keys(props).length > 0) window.umami.track(event, props);
        else window.umami.track(event);
      }
    } catch {}
  }

  window.ECGAnalytics = { load, unload, track };
})();
