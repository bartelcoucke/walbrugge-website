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
  const animateCounter = (element, target) => {
    const duration = 2000;
    const start = 0;
    const startTime = performance.now();
    
    const updateCounter = (currentTime) => {
      const elapsed = currentTime - startTime;
      const progress = Math.min(elapsed / duration, 1);
      const easeOut = 1 - Math.pow(1 - progress, 3);
      const current = Math.floor(start + (target - start) * easeOut);
      
      element.textContent = current;
      
      if (progress < 1) {
        requestAnimationFrame(updateCounter);
      } else {
        element.textContent = target;
      }
    };
    
    requestAnimationFrame(updateCounter);
  };

  // Observe stat numbers
  document.querySelectorAll('.stat-num').forEach(stat => {
    const target = parseInt(stat.textContent);
    if (!isNaN(target)) {
      const observer = new IntersectionObserver((entries) => {
        if (entries[0].isIntersecting) {
          animateCounter(stat, target);
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
