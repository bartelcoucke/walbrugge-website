// ═══════════════════════════════════════════════════════════════════════════
// WALBRUGGE - Rustieke Hoeve JavaScript
// Parallax, scroll-reveal, smooth interactions
// ═══════════════════════════════════════════════════════════════════════════

(function() {
  'use strict';

  // ── Navigation ─────────────────────────────────────────────────────────────
  const nav = document.getElementById('nav');
  const navBurger = document.getElementById('navBurger');
  const navLinks = document.getElementById('navLinks');
  const hero = document.querySelector('.hero');

  // Scroll effect with parallax
  let ticking = false;
  
  function updateOnScroll() {
    const currentScroll = window.pageYOffset;
    
    // Nav background
    if (currentScroll > 60) {
      nav.classList.add('scrolled');
    } else {
      nav.classList.remove('scrolled');
    }
    
    // Fade hero content on scroll
    const heroContent = document.querySelector('.hero-content');
    if (heroContent && currentScroll < window.innerHeight) {
      const opacity = 1 - (currentScroll / (window.innerHeight * 0.6));
      const translateY = currentScroll * 0.3;
      heroContent.style.opacity = Math.max(0, opacity);
      heroContent.style.transform = `translateY(${translateY}px)`;
    }
    
    ticking = false;
  }

  window.addEventListener('scroll', () => {
    if (!ticking) {
      requestAnimationFrame(updateOnScroll);
      ticking = true;
    }
  }, { passive: true });

  // Mobile menu toggle
  if (navBurger && navLinks) {
    navBurger.addEventListener('click', () => {
      navBurger.classList.toggle('active');
      navLinks.classList.toggle('active');
      navBurger.setAttribute('aria-expanded', 
        navBurger.classList.contains('active') ? 'true' : 'false'
      );
    });

    // Close menu when clicking a link
    navLinks.querySelectorAll('a').forEach(link => {
      link.addEventListener('click', () => {
        navBurger.classList.remove('active');
        navLinks.classList.remove('active');
        navBurger.setAttribute('aria-expanded', 'false');
      });
    });
  }

  // ── Smooth scroll for anchor links ─────────────────────────────────────────
  document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function(e) {
      const target = document.querySelector(this.getAttribute('href'));
      if (target) {
        e.preventDefault();
        const offset = nav ? nav.offsetHeight + 20 : 20;
        const top = target.getBoundingClientRect().top + window.pageYOffset - offset;
        window.scrollTo({ top, behavior: 'smooth' });
      }
    });
  });

  // ── Scroll Reveal Animation ────────────────────────────────────────────────
  const revealElements = document.querySelectorAll(
    '.keuze-card, .testimonial, .ruimte-item, .stat, .facil-item, .intro-block, .cta-block'
  );

  const revealObserver = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        entry.target.classList.add('visible');
        // Add staggered delay for siblings
        const siblings = entry.target.parentElement.children;
        const index = Array.from(siblings).indexOf(entry.target);
        entry.target.style.transitionDelay = `${index * 0.1}s`;
      }
    });
  }, {
    threshold: 0.15,
    rootMargin: '0px 0px -60px 0px'
  });

  revealElements.forEach(el => {
    el.classList.add('reveal');
    revealObserver.observe(el);
  });

  // ── Keuze Cards Enhanced Hover ─────────────────────────────────────────────
  document.querySelectorAll('.keuze-card').forEach(card => {
    card.addEventListener('mouseenter', function() {
      this.style.transform = 'translateY(-14px) scale(1.02)';
    });
    
    card.addEventListener('mouseleave', function() {
      this.style.transform = '';
    });
  });

  // ── Image lazy loading with fade ───────────────────────────────────────────
  document.querySelectorAll('img[data-src]').forEach(img => {
    img.style.opacity = '0';
    img.style.transition = 'opacity 0.5s ease';
    
    const loadImage = () => {
      img.src = img.dataset.src;
      img.onload = () => {
        img.style.opacity = '1';
      };
    };

    if ('IntersectionObserver' in window) {
      const imgObserver = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
          if (entry.isIntersecting) {
            loadImage();
            imgObserver.unobserve(img);
          }
        });
      });
      imgObserver.observe(img);
    } else {
      loadImage();
    }
  });

  // ── Testimonials auto-rotate (optional) ────────────────────────────────────
  const testimonials = document.querySelectorAll('.testimonial');
  if (testimonials.length > 1 && window.innerWidth < 768) {
    let currentTestimonial = 0;
    
    // Hide all except first on mobile
    testimonials.forEach((t, i) => {
      if (i > 0) t.style.display = 'none';
    });
    
    setInterval(() => {
      testimonials[currentTestimonial].style.display = 'none';
      currentTestimonial = (currentTestimonial + 1) % testimonials.length;
      testimonials[currentTestimonial].style.display = 'block';
      testimonials[currentTestimonial].style.animation = 'fadeInUp 0.5s ease';
    }, 5000);
  }

  // ── Smooth number counter for stats ────────────────────────────────────────
  const animateCounter = (element, target, decimals = 0, sep = '.') => {
    const duration = 2000;
    const start = 0;
    const startTime = performance.now();
    const fmt = (v) => decimals > 0 ? v.toFixed(decimals).replace('.', sep) : String(Math.floor(v));

    const updateCounter = (currentTime) => {
      const elapsed = currentTime - startTime;
      const progress = Math.min(elapsed / duration, 1);
      const easeOut = 1 - Math.pow(1 - progress, 3);
      const current = start + (target - start) * easeOut;

      element.textContent = fmt(current);

      if (progress < 1) {
        requestAnimationFrame(updateCounter);
      } else {
        element.textContent = fmt(target);
      }
    };

    requestAnimationFrame(updateCounter);
  };

  // Observe stat numbers (ondersteunt ook decimalen zoals "9,7")
  document.querySelectorAll('.stat-num').forEach(stat => {
    const raw = stat.textContent.trim();
    if (!/^\d+([.,]\d+)?$/.test(raw)) return;
    const sep = raw.includes(',') ? ',' : '.';
    const decimals = (raw.split(/[.,]/)[1] || '').length;
    const target = parseFloat(raw.replace(',', '.'));
    if (!isNaN(target)) {
      const observer = new IntersectionObserver((entries) => {
        if (entries[0].isIntersecting) {
          animateCounter(stat, target, decimals, sep);
          observer.unobserve(stat);
        }
      }, { threshold: 0.5 });
      observer.observe(stat);
    }
  });

  // ── Form validation with style ─────────────────────────────────────────────
  document.querySelectorAll('form[data-validate]').forEach(form => {
    form.addEventListener('submit', function(e) {
      let valid = true;
      const required = form.querySelectorAll('[required]');
      
      required.forEach(field => {
        // Remove previous error state
        field.classList.remove('error');
        field.style.borderColor = '';
        
        if (!field.value.trim()) {
          valid = false;
          field.style.borderColor = '#c9302c';
          field.style.animation = 'shake 0.5s ease';
        }
        
        // Email validation
        if (field.type === 'email' && field.value) {
          const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
          if (!emailRegex.test(field.value)) {
            valid = false;
            field.style.borderColor = '#c9302c';
          }
        }
      });

      if (!valid) {
        e.preventDefault();
      }
    });
    
    // Clear error on input
    form.querySelectorAll('input, textarea, select').forEach(field => {
      field.addEventListener('input', () => {
        field.style.borderColor = '';
        field.style.animation = '';
      });
    });
  });

  // ── Add shake animation ────────────────────────────────────────────────────
  const shakeStyle = document.createElement('style');
  shakeStyle.textContent = `
    @keyframes shake {
      0%, 100% { transform: translateX(0); }
      20%, 60% { transform: translateX(-8px); }
      40%, 80% { transform: translateX(8px); }
    }
  `;
  document.head.appendChild(shakeStyle);

  // ── Booking.com links in de taal van de pagina ─────────────────────────────
  const bookingLinks = document.querySelectorAll('#bookingBadge, .js-booking-link');
  if (bookingLinks.length) {
    const lang = (document.documentElement.lang || 'nl').split('-')[0];
    const bookingLangs = { nl: 'nl', fr: 'fr', de: 'de', en: 'en-gb' };
    const suffix = bookingLangs[lang];
    const url = suffix
      ? 'https://www.booking.com/hotel/be/walbrugge.' + suffix + '.html'
      : 'https://www.booking.com/hotel/be/walbrugge.html';
    bookingLinks.forEach(el => { el.href = url; });
  }

  // ── Console branding ───────────────────────────────────────────────────────
  console.log(
    '%c🏡 Domein Walbrugge',
    'font-size: 28px; font-weight: bold; color: #5c7a4a; font-family: Georgia;'
  );
  console.log(
    '%c« Een 18e-eeuwse vierkantshoeve aan de Tiegemberg »',
    'font-size: 14px; font-style: italic; color: #b8860b;'
  );
  console.log(
    '%cFeestzaal · B&B · Zakelijk\nhttps://walbrugge.be',
    'font-size: 12px; color: #7a6d5d;'
  );

  // ── Page loaded ────────────────────────────────────────────────────────────
  document.body.classList.add('loaded');

})();

/* ── Beoordelingsscores actueel houden ──────────────────────────────────
 * De waarden in de HTML zijn de laatst bekende stand. Dit haalt de actuele
 * scores op: Google komt automatisch uit de Places API, Booking.com en
 * Eventplanner beheert de eigenaar in het beheerpaneel.
 */
(function () {
  var velden = document.querySelectorAll('[data-score], [data-score-stars]');
  if (!velden.length) return;

  fetch('/api/scores')
    .then(function (r) { return r.ok ? r.json() : null; })
    .then(function (d) {
      if (!d || !d.ok) return;
      velden.forEach(function (el) {
        var naam = el.getAttribute('data-score') || el.getAttribute('data-score-stars');
        var bron = d.scores[naam];
        if (!bron || !bron.score) return;

        if (el.hasAttribute('data-score-stars')) {
          // Score kan 4,6 of 4.6 zijn, en soms 9,4 op een schaal van 10.
          var n = parseFloat(String(bron.score).replace(',', '.'));
          if (isNaN(n)) return;
          if (n > 5) n = n / 2;
          var vul = el.querySelector('.g-stars-fill');
          if (vul) vul.style.width = Math.max(0, Math.min(100, n / 5 * 100)).toFixed(1) + '%';
          el.setAttribute('aria-label',
            n.toFixed(1).replace('.', ',') + ' van 5 op Google' +
            (bron.aantal ? ', ' + bron.aantal + ' beoordelingen' : ''));
        } else {
          el.textContent = bron.score;
        }
      });
    })
    .catch(function () { /* stil falen: de HTML-waarde blijft staan */ });
})();

/* ── Serverside bezoekersstatistiek (zonder cookies) ────────────────────────
 * Stuurt enkele gebeurtenissen naar /api/telling: klik op de offerteknop,
 * offerte verstuurd, WhatsApp/Messenger/telefoon/e-mail, Boek B&B, klik op
 * een zaal, bladeren in een fotocarrousel, klik op een award of reviewlink.
 * Er wordt niets op het toestel bewaard. De server noteert bij elke melding wel
 * het IP-adres, zodat het beheer terugkerende bezoekers kan herkennen.
 */
(function () {
  var q = new URLSearchParams(location.search);
  var basis = {
    pad: location.pathname,
    taal: document.documentElement.lang || 'nl',
    // Ruwe verwijzer; de server herkent er sites, Android-apps en in-app-browsers in.
    ref: (document.referrer || '').slice(0, 300),
    utm_source: q.get('utm_source') || '',
    utm_medium: q.get('utm_medium') || '',
    utm_campaign: q.get('utm_campaign') || '',
    // Sleutel per paginabezoek: zo weet de server dat meerdere klikken bij hetzelfde
    // bezoek horen, ook als een melding later aankomt of het toestel intussen van
    // netwerk wisselde.
    pid: Math.random().toString(36).slice(2, 12)
  };
  function stuur(naam, detail, extra) {
    var body = JSON.stringify(Object.assign({ naam: naam, detail: detail || '' }, basis, extra || {}));
    try {
      if (navigator.sendBeacon && navigator.sendBeacon('/api/telling', new Blob([body], { type: 'application/json' }))) return;
    } catch (e) { /* val terug op fetch */ }
    try {
      fetch('/api/telling', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: body, keepalive: true })
        .catch(function () {});
    } catch (e) { /* stil */ }
  }
  // Eén levensteken per paginabezoek. Daaraan ziet de server dat er een echte browser
  // achter zit (een robot stuurt dit nooit). Haalde de browser de pagina vooraf op en
  // toont ze die nu pas, dan is het bezoek bij de server nog niet geteld; met de vlag
  // "vooraf" gebeurt dat alsnog, op het moment dat de bezoeker de pagina echt ziet.
  (function () {
    var vooraf = false;
    try {
      var nav = performance.getEntriesByType && performance.getEntriesByType('navigation')[0];
      vooraf = !!(nav && nav.activationStart > 0);
    } catch (e) { /* oudere browser */ }
    stuur('weergave', '', { vooraf: vooraf });
  })();

  var tekst = function (a) { return (a.textContent || '').trim().replace(/\s+/g, ' ').slice(0, 60); };

  // Fotocarrousels: de pijltjes en bolletjes van de zaaltegels (teams), de
  // B&B-kamers en de feestpagina. Geeft een vaste sleutel per carrousel, gelijk
  // in alle talen: zaal:<slug>, kamer:<naam>, feest:<naam>.
  var CARROUSEL_KNOP = '.zaal-slider-btn, .zaal-slider-dot, .carousel-btn, .carousel-dot, .slider-btn, .slider-dot';
  function carrouselId(el) {
    var kamer = el.closest('.room-carousel[data-room]');
    if (kamer) return 'kamer:' + kamer.getAttribute('data-room');
    var feest = el.closest('.slider-container[data-slider]');
    if (feest) return 'feest:' + feest.getAttribute('data-slider');
    var zaal = el.closest('.zaal-slider');
    if (zaal) {
      var link = zaal.closest('a[href]');
      var m = link ? /\/ruimtes\/([a-z0-9-]+)/i.exec(link.getAttribute('href') || '') : null;
      if (m) return 'zaal:' + m[1].toLowerCase();
      // Tegel zonder link (bv. het terras): sleutel uit de bestandsnaam van de eerste foto
      var img = zaal.querySelector('img');
      var naam = (img ? img.getAttribute('src') || '' : '').split('/').pop()
        .replace(/\.[a-z0-9]+$/i, '').replace(/-(c|carrousel-?)?\d+$/i, '').replace(/^(teams|feesten)-/, '');
      return 'zaal:' + (naam || 'onbekend');
    }
    return '';
  }

  // Awards en reviewlinks: Salino, Booking.com, Eventplanner en Google. De
  // award-kaarten onder "Erkend & gewaardeerd" tellen als 'tegel'; de badges in
  // de hero, tekstlinks en knoppen als 'knop'. Google telt via de zoeklink, g.page
  // en de Maps-vermelding (?cid= of /maps/place/, zoals "Alle beoordelingen op
  // Google" op de homepage); routelinks (/maps/dir, maps?q=) tellen niet mee.
  function awardVan(a, href) {
    var h = href.toLowerCase();
    var platform = h.indexOf('salino.be') !== -1 ? 'salino'
      : h.indexOf('booking.com') !== -1 ? 'booking'
      : h.indexOf('eventplanner.') !== -1 ? 'eventplanner'
      : /google\.[a-z.]+\/search|g\.page\/|maps\.google\.[a-z.]+\/\?cid=|google\.[a-z.]+\/maps\/place\//.test(h) ? 'google' : '';
    if (!platform) return '';
    return platform + ':' + (a.classList.contains('award-card') ? 'tegel' : 'knop');
  }

  document.addEventListener('click', function (ev) {
    var t = ev.target;
    if (!t || !t.closest) return;
    var knop = t.closest(CARROUSEL_KNOP);
    if (knop) { var id = carrouselId(knop); if (id) stuur('carrousel', id); return; }
    var a = t.closest('a[href]');
    if (!a) return;
    var href = a.getAttribute('href') || '';
    var award = awardVan(a, href);
    if (award) { stuur('award_click', award); return; }
    // Klik op een zaal: de zaaltegels op de teamspagina en de zaallinks op de
    // zaalpagina's zelf (niet de taalkeuze, die verwijst naar dezelfde zaal).
    var zaal = /^\/(?:fr\/|en\/|de\/)?ruimtes\/([a-z0-9-]+)/i.exec(href);
    if (zaal && !a.closest('.lang-switch')) stuur('zaal_click', zaal[1].toLowerCase());
    else if (/^\/(fr\/|en\/|de\/)?offerte(\?|#|$)/.test(href)) stuur('offerte_click', tekst(a));
    else if (href.indexOf('wa.me/') !== -1 || href.indexOf('whatsapp.com') !== -1) stuur('whatsapp_click');
    else if (href.indexOf('m.me/') !== -1 || href.indexOf('messenger.com') !== -1) stuur('messenger_click');
    else if (href.indexOf('bookingengine.mylighthouse.com') !== -1) stuur('booking_click', tekst(a));
    else if (href.indexOf('tel:') === 0) stuur('phone_click');
    else if (href.indexOf('mailto:') === 0) stuur('email_click');
  }, true);

  // Offerte verstuurd: de offertepagina's roepen walbruggeTrack('generate_lead') aan
  // (dat stuurt naar Google Analytics, enkel na toestemming). Hier haken we in
  // zodat de server het altijd telt — met of zonder cookies.
  function koppel() {
    var orig = window.walbruggeTrack;
    window.walbruggeTrack = function (naam, params) {
      if (naam === 'generate_lead') stuur('generate_lead', params && params.event_type ? String(params.event_type) : '');
      if (typeof orig === 'function') { try { orig(naam, params); } catch (e) { /* stil */ } }
    };
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', koppel);
  else koppel();
})();
