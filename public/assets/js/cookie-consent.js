/**
 * Cookie Consent Manager — Domein Walbrugge
 * Blokkeert tracking scripts tot de bezoeker toestemming geeft.
 * Voldoet aan ePrivacy-richtlijn + GDPR.
 */
(function () {
  'use strict';

  var COOKIE_NAME = 'walbrugge_consent';
  var COOKIE_DAYS = 365;

  /* ── i18n teksten ── */
  var LANG = document.documentElement.lang || 'nl';
  var T = {
    nl: {
      text: 'Deze website gebruikt cookies voor analyse en om uw ervaring te verbeteren. Meer info in ons <a href="/privacy">privacybeleid</a>.',
      accept: 'Alle cookies aanvaarden',
      reject: 'Alleen noodzakelijke',
      settings: 'Instellingen'
    },
    fr: {
      text: 'Ce site utilise des cookies pour l\'analyse et pour améliorer votre expérience. Plus d\'informations dans notre <a href="/fr/privacy">politique de confidentialité</a>.',
      accept: 'Accepter tous les cookies',
      reject: 'Cookies essentiels uniquement',
      settings: 'Paramètres'
    },
    en: {
      text: 'This website uses cookies for analytics and to improve your experience. Learn more in our <a href="/en/privacy">privacy policy</a>.',
      accept: 'Accept all cookies',
      reject: 'Essential only',
      settings: 'Settings'
    },
    de: {
      text: 'Diese Website verwendet Cookies für Analysen und um Ihr Erlebnis zu verbessern. Mehr dazu in unserer <a href="/de/privacy">Datenschutzerklärung</a>.',
      accept: 'Alle Cookies akzeptieren',
      reject: 'Nur notwendige',
      settings: 'Einstellungen'
    }
  };
  var t = T[LANG] || T['nl'];

  /* ── Cookie helpers ── */
  function getCookie(name) {
    var v = document.cookie.match('(^|;)\\s*' + name + '\\s*=\\s*([^;]+)');
    return v ? v.pop() : null;
  }
  function setCookie(name, value, days) {
    var d = new Date();
    d.setTime(d.getTime() + days * 864e5);
    document.cookie = name + '=' + value + ';expires=' + d.toUTCString() + ';path=/;SameSite=Lax;Secure';
  }

  /* ── LinkedIn Insight Tag loader ── */
  function loadLinkedIn() {
    if (window._walbrugge_li_loaded) return;
    window._walbrugge_li_loaded = true;

    window._linkedin_partner_id = '9619338';
    window._linkedin_data_partner_ids = window._linkedin_data_partner_ids || [];
    window._linkedin_data_partner_ids.push(window._linkedin_partner_id);

    var s = document.getElementsByTagName('script')[0];
    var b = document.createElement('script');
    b.type = 'text/javascript';
    b.async = true;
    b.src = 'https://snap.licdn.com/li.lms-analytics/insight.min.js';
    s.parentNode.insertBefore(b, s);

    // noscript fallback pixel
    var img = document.createElement('img');
    img.height = 1;
    img.width = 1;
    img.style.display = 'none';
    img.alt = '';
    img.src = 'https://px.ads.linkedin.com/collect/?pid=9619338&fmt=gif';
    document.body.appendChild(img);
  }

  /* ── Consent handlers ── */
  function acceptAll() {
    setCookie(COOKIE_NAME, 'all', COOKIE_DAYS);
    hideBanner();
    loadLinkedIn();
  }
  function rejectOptional() {
    setCookie(COOKIE_NAME, 'essential', COOKIE_DAYS);
    hideBanner();
  }

  /* ── Banner UI ── */
  function hideBanner() {
    var el = document.getElementById('cookieConsentBanner');
    if (el) el.style.display = 'none';
  }

  function showBanner() {
    var banner = document.createElement('div');
    banner.id = 'cookieConsentBanner';
    banner.setAttribute('role', 'dialog');
    banner.setAttribute('aria-label', 'Cookie consent');
    banner.innerHTML =
      '<div class="cc-inner">' +
        '<p class="cc-text">' + t.text + '</p>' +
        '<div class="cc-buttons">' +
          '<button class="cc-btn cc-accept" id="ccAccept">' + t.accept + '</button>' +
          '<button class="cc-btn cc-reject" id="ccReject">' + t.reject + '</button>' +
        '</div>' +
      '</div>';

    // Inline styles (no external CSS needed)
    banner.style.cssText = 'position:fixed;bottom:0;left:0;right:0;z-index:99999;background:rgba(45,58,40,0.97);color:#fff;padding:1rem 1.5rem;font-family:Inter,sans-serif;font-size:0.9rem;line-height:1.5;box-shadow:0 -4px 20px rgba(0,0,0,0.3);';

    var inner = banner.querySelector('.cc-inner');
    inner.style.cssText = 'max-width:1100px;margin:0 auto;display:flex;align-items:center;gap:1.5rem;flex-wrap:wrap;';

    var text = banner.querySelector('.cc-text');
    text.style.cssText = 'flex:1;min-width:280px;margin:0;';
    // Style links in text
    var links = text.querySelectorAll('a');
    for (var i = 0; i < links.length; i++) {
      links[i].style.cssText = 'color:#C5A572;text-decoration:underline;';
    }

    var btns = banner.querySelector('.cc-buttons');
    btns.style.cssText = 'display:flex;gap:0.75rem;flex-shrink:0;flex-wrap:wrap;';

    var acceptBtn = banner.querySelector('#ccAccept');
    acceptBtn.style.cssText = 'padding:0.6rem 1.5rem;border:none;border-radius:6px;cursor:pointer;font-weight:600;font-size:0.85rem;background:#C5A572;color:#fff;transition:background 0.2s;';

    var rejectBtn = banner.querySelector('#ccReject');
    rejectBtn.style.cssText = 'padding:0.6rem 1.5rem;border:1px solid rgba(255,255,255,0.4);border-radius:6px;cursor:pointer;font-weight:500;font-size:0.85rem;background:transparent;color:#fff;transition:background 0.2s;';

    document.body.appendChild(banner);

    acceptBtn.addEventListener('click', acceptAll);
    rejectBtn.addEventListener('click', rejectOptional);
  }

  /* ── Init ── */
  function init() {
    var consent = getCookie(COOKIE_NAME);
    if (consent === 'all') {
      loadLinkedIn();
    } else if (consent === 'essential') {
      // Do nothing — no tracking
    } else {
      // No consent yet — show banner
      if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', showBanner);
      } else {
        showBanner();
      }
    }
  }

  init();
})();
