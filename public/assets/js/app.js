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
 * Stuurt enkele gebeurtenissen naar /api/track: klik op de offerteknop,
 * offerte verstuurd, WhatsApp/Messenger/telefoon/e-mail, Boek B&B. Er wordt
 * niets op het toestel bewaard; de server slaat geen IP of user-agent op.
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
    utm_campaign: q.get('utm_campaign') || ''
  };
  function stuur(naam, detail) {
    var body = JSON.stringify(Object.assign({ naam: naam, detail: detail || '' }, basis));
    try {
      if (navigator.sendBeacon && navigator.sendBeacon('/api/telling', new Blob([body], { type: 'application/json' }))) return;
    } catch (e) { /* val terug op fetch */ }
    try {
      fetch('/api/telling', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: body, keepalive: true })
        .catch(function () {});
    } catch (e) { /* stil */ }
  }
  var tekst = function (a) { return (a.textContent || '').trim().replace(/\s+/g, ' ').slice(0, 60); };

  document.addEventListener('click', function (ev) {
    var a = ev.target && ev.target.closest ? ev.target.closest('a[href]') : null;
    if (!a) return;
    var href = a.getAttribute('href') || '';
    if (/^\/(fr\/|en\/|de\/)?offerte(\?|#|$)/.test(href)) stuur('offerte_click', tekst(a));
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
