/**
 * Walbrugge Backend Server
 * Express + SQLite + JWT Auth
 */

const express = require('express');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const jwt = require('jsonwebtoken');
const bcrypt = require('bcryptjs');
const Database = require('better-sqlite3');
const nodemailer = require('nodemailer');

const app = express();
const PORT = process.env.PORT || 8100;
const JWT_SECRET = process.env.JWT_SECRET || crypto.randomBytes(32).toString('hex');
const DATA_DIR = path.join(__dirname, '..', 'data');

// Ensure data directory exists
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

// Database setup
const db = new Database(path.join(DATA_DIR, 'walbrugge.db'));
db.pragma('journal_mode = WAL');

// Initialize tables
db.exec(`
  CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    email TEXT UNIQUE NOT NULL,
    password_hash TEXT,
    role TEXT DEFAULT 'guest',
    name TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
  );
  
  CREATE TABLE IF NOT EXISTS bookings (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    reference TEXT UNIQUE NOT NULL,
    guest_email TEXT NOT NULL,
    guest_name TEXT NOT NULL,
    guest_phone TEXT,
    check_in DATE,
    check_out DATE,
    room TEXT,
    type TEXT DEFAULT 'bb',
    persons INTEGER DEFAULT 2,
    status TEXT DEFAULT 'pending',
    notes TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
  );
  
  CREATE TABLE IF NOT EXISTS contacts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    naam TEXT NOT NULL,
    email TEXT NOT NULL,
    telefoon TEXT,
    bedrijf TEXT,
    type TEXT,
    personen INTEGER,
    datum TEXT,
    formule TEXT,
    bericht TEXT,
    status TEXT DEFAULT 'nieuw',
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
  );
  
  CREATE TABLE IF NOT EXISTS rooms (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL,
    slug TEXT UNIQUE NOT NULL,
    description TEXT,
    capacity INTEGER DEFAULT 2,
    price_base REAL,
    amenities TEXT,
    images TEXT,
    available INTEGER DEFAULT 1
  );
  
  CREATE TABLE IF NOT EXISTS blog_posts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    title TEXT NOT NULL,
    slug TEXT UNIQUE NOT NULL,
    excerpt TEXT,
    content TEXT NOT NULL,
    category TEXT,
    tags TEXT,
    featured_image TEXT,
    author TEXT DEFAULT 'Walbrugge',
    status TEXT DEFAULT 'draft',
    published_at DATETIME,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
  );
  
  CREATE TABLE IF NOT EXISTS blog_categories (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL,
    slug TEXT UNIQUE NOT NULL,
    description TEXT
  );
`);

// Insert admin user if not exists
const adminExists = db.prepare("SELECT id FROM users WHERE email = ?").get('admin@walbrugge.be');
if (!adminExists) {
  const hash = bcrypt.hashSync('walbrugge2024', 10);
  db.prepare("INSERT INTO users (email, password_hash, role, name) VALUES (?, ?, 'admin', 'Administrator')")
    .run('admin@walbrugge.be', hash);
  console.log('Admin user created: admin@walbrugge.be');
}

// Insert sample rooms if none exist
const roomCount = db.prepare("SELECT COUNT(*) as count FROM rooms").get().count;
if (roomCount === 0) {
  const rooms = [
    { name: 'Campanula', slug: 'campanula', description: 'Gezellige tweepersoonskamer op de begane grond', capacity: 2, price_base: 95 },
    { name: 'Chicory Ast', slug: 'chicory-ast', description: 'Romantische kamer met uitzicht op het voedselbos', capacity: 2, price_base: 105 },
    { name: 'Cornus Mas', slug: 'cornus-mas', description: 'Ruime familiekamer onder de dakbalken', capacity: 4, price_base: 135 },
    { name: 'Eucalyptus', slug: 'eucalyptus', description: 'Luxe suite met eigen badkamer en terras', capacity: 2, price_base: 125 },
    { name: 'Kardoen', slug: 'kardoen', description: 'Authentieke boerenkamer met originele elementen', capacity: 2, price_base: 95 }
  ];
  const insert = db.prepare("INSERT INTO rooms (name, slug, description, capacity, price_base) VALUES (?, ?, ?, ?, ?)");
  rooms.forEach(r => insert.run(r.name, r.slug, r.description, r.capacity, r.price_base));
  console.log('Sample rooms created');
}

// Middleware
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Dynamische sitemap: statische paginas + gepubliceerde blogartikels
app.get('/sitemap.xml', (req, res) => {
  const staticSitemap = fs.readFileSync(path.join(__dirname, '..', 'public', 'sitemap.xml'), 'utf-8');
  let blogUrls = '';
  try {
    const posts = db.prepare("SELECT slug, updated_at, published_at, created_at FROM blog_posts WHERE status = 'published' ORDER BY published_at DESC").all();
    blogUrls = posts.map(p => {
      const lastmod = (p.updated_at || p.published_at || p.created_at || '').slice(0, 10);
      return `  <url>\n    <loc>https://walbrugge.be/blog/${p.slug}</loc>\n${lastmod ? `    <lastmod>${lastmod}</lastmod>\n` : ''}    <changefreq>monthly</changefreq>\n    <priority>0.6</priority>\n  </url>`;
    }).join('\n');
  } catch (e) { /* blog tabel nog niet beschikbaar */ }
  const xml = blogUrls
    ? staticSitemap.replace('</urlset>', blogUrls + '\n</urlset>')
    : staticSitemap;
  res.type('application/xml').send(xml);
});

// Static files
app.use(express.static(path.join(__dirname, '..', 'public')));

// Logging
app.use((req, res, next) => {
  console.log(`${new Date().toISOString()} ${req.method} ${req.path}`);
  next();
});

// ═══════════════════════════════════════════════════════════════════════════
// AUTH MIDDLEWARE
// ═══════════════════════════════════════════════════════════════════════════

function authMiddleware(requiredRole = null) {
  return (req, res, next) => {
    const authHeader = req.headers.authorization;
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
      return res.status(401).json({ error: 'Geen authenticatie' });
    }
    
    try {
      const token = authHeader.split(' ')[1];
      const decoded = jwt.verify(token, JWT_SECRET);
      req.user = decoded;
      
      if (requiredRole && decoded.role !== requiredRole && decoded.role !== 'admin') {
        return res.status(403).json({ error: 'Onvoldoende rechten' });
      }
      
      next();
    } catch (e) {
      return res.status(401).json({ error: 'Ongeldige token' });
    }
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// AUTH ROUTES
// ═══════════════════════════════════════════════════════════════════════════

// Unified login (auto-detect guest vs admin)
app.post('/api/login', (req, res) => {
  const { email, password, reference } = req.body;
  
  // Guest login: email + booking reference
  if (email && reference && !password) {
    const booking = db.prepare(`
      SELECT b.*, u.id as user_id 
      FROM bookings b 
      LEFT JOIN users u ON u.email = b.guest_email 
      WHERE b.guest_email = ? AND b.reference = ?
    `).get(email, reference.toUpperCase());
    
    if (!booking) {
      return res.status(401).json({ error: 'Boeking niet gevonden. Controleer uw e-mail en referentie.' });
    }
    
    const token = jwt.sign({
      id: booking.user_id || 0,
      email: booking.guest_email,
      role: 'guest',
      bookingId: booking.id,
      reference: booking.reference
    }, JWT_SECRET, { expiresIn: '7d' });
    
    return res.json({
      ok: true,
      token,
      user: {
        email: booking.guest_email,
        name: booking.guest_name,
        role: 'guest'
      },
      booking: {
        reference: booking.reference,
        checkIn: booking.check_in,
        checkOut: booking.check_out,
        room: booking.room,
        status: booking.status
      },
      redirect: '/gasten/dashboard'
    });
  }
  
  // Admin login: email + password
  if (email && password) {
    const user = db.prepare("SELECT * FROM users WHERE email = ?").get(email);
    
    if (!user || !user.password_hash) {
      return res.status(401).json({ error: 'Ongeldige inloggegevens' });
    }
    
    if (!bcrypt.compareSync(password, user.password_hash)) {
      return res.status(401).json({ error: 'Ongeldige inloggegevens' });
    }
    
    const token = jwt.sign({
      id: user.id,
      email: user.email,
      role: user.role,
      name: user.name
    }, JWT_SECRET, { expiresIn: '24h' });
    
    return res.json({
      ok: true,
      token,
      user: {
        email: user.email,
        name: user.name,
        role: user.role
      },
      redirect: user.role === 'admin' ? '/admin' : '/gasten/dashboard'
    });
  }
  
  return res.status(400).json({ error: 'Vul e-mail en wachtwoord of boekingsreferentie in' });
});

// Token verification
app.get('/api/me', authMiddleware(), (req, res) => {
  res.json({ ok: true, user: req.user });
});

// ═══════════════════════════════════════════════════════════════════════════
// CONTACT / OFFERTE ROUTES
// ═══════════════════════════════════════════════════════════════════════════

app.post('/api/contact', (req, res) => {
  const { naam, email, telefoon, bedrijf, type, personen, datum, formule, bericht, website } = req.body;
  
  // Honeypot check
  if (website) {
    return res.status(400).json({ error: 'Spam gedetecteerd' });
  }
  
  if (!naam || !email) {
    return res.status(400).json({ error: 'Naam en e-mail zijn verplicht' });
  }
  
  try {
    db.prepare(`
      INSERT INTO contacts (naam, email, telefoon, bedrijf, type, personen, datum, formule, bericht)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).run(naam, email, telefoon || null, bedrijf || null, type || null, personen || null, datum || null, formule || null, bericht || null);
    
    console.log(`New contact: ${naam} <${email}> - ${type}`);
    
    res.json({ ok: true, message: 'Aanvraag ontvangen' });
  } catch (e) {
    console.error('Contact error:', e);
    res.status(500).json({ error: 'Opslaan mislukt' });
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// BOOKINGS ROUTES
// ═══════════════════════════════════════════════════════════════════════════

// Guest: get own booking details
app.get('/api/bookings/mine', authMiddleware('guest'), (req, res) => {
  if (!req.user.bookingId) {
    return res.status(404).json({ error: 'Geen boeking gevonden' });
  }
  
  const booking = db.prepare("SELECT * FROM bookings WHERE id = ?").get(req.user.bookingId);
  if (!booking) {
    return res.status(404).json({ error: 'Boeking niet gevonden' });
  }
  
  res.json({ ok: true, booking });
});

// Admin: list all bookings
app.get('/api/bookings', authMiddleware('admin'), (req, res) => {
  const bookings = db.prepare("SELECT * FROM bookings ORDER BY check_in DESC").all();
  res.json({ ok: true, bookings });
});

// Admin: create booking
app.post('/api/bookings', authMiddleware('admin'), (req, res) => {
  const { guest_email, guest_name, guest_phone, check_in, check_out, room, type, persons, notes } = req.body;
  
  if (!guest_email || !guest_name || !check_in) {
    return res.status(400).json({ error: 'Verplichte velden ontbreken' });
  }
  
  // Generate unique reference
  const reference = 'WB-' + Date.now().toString(36).toUpperCase() + '-' + crypto.randomBytes(2).toString('hex').toUpperCase();
  
  try {
    const result = db.prepare(`
      INSERT INTO bookings (reference, guest_email, guest_name, guest_phone, check_in, check_out, room, type, persons, notes)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).run(reference, guest_email, guest_name, guest_phone || null, check_in, check_out || null, room || null, type || 'bb', persons || 2, notes || null);
    
    res.json({ ok: true, id: result.lastInsertRowid, reference });
  } catch (e) {
    console.error('Booking error:', e);
    res.status(500).json({ error: 'Boeking aanmaken mislukt' });
  }
});

// Admin: update booking
app.put('/api/bookings/:id', authMiddleware('admin'), (req, res) => {
  const { id } = req.params;
  const { status, notes, check_in, check_out, room, persons } = req.body;
  
  const updates = [];
  const values = [];
  
  if (status !== undefined) { updates.push('status = ?'); values.push(status); }
  if (notes !== undefined) { updates.push('notes = ?'); values.push(notes); }
  if (check_in !== undefined) { updates.push('check_in = ?'); values.push(check_in); }
  if (check_out !== undefined) { updates.push('check_out = ?'); values.push(check_out); }
  if (room !== undefined) { updates.push('room = ?'); values.push(room); }
  if (persons !== undefined) { updates.push('persons = ?'); values.push(persons); }
  
  if (updates.length === 0) {
    return res.status(400).json({ error: 'Geen updates' });
  }
  
  updates.push('updated_at = CURRENT_TIMESTAMP');
  values.push(id);
  
  try {
    db.prepare(`UPDATE bookings SET ${updates.join(', ')} WHERE id = ?`).run(...values);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: 'Update mislukt' });
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// ROOMS ROUTES
// ═══════════════════════════════════════════════════════════════════════════

app.get('/api/rooms', (req, res) => {
  const rooms = db.prepare("SELECT * FROM rooms WHERE available = 1 ORDER BY name").all();
  res.json({ ok: true, rooms });
});

app.get('/api/rooms/:slug', (req, res) => {
  const room = db.prepare("SELECT * FROM rooms WHERE slug = ?").get(req.params.slug);
  if (!room) {
    return res.status(404).json({ error: 'Kamer niet gevonden' });
  }
  res.json({ ok: true, room });
});

// ═══════════════════════════════════════════════════════════════════════════
// ADMIN ROUTES
// ═══════════════════════════════════════════════════════════════════════════

// Contacts list
app.get('/api/contacts', authMiddleware('admin'), (req, res) => {
  const contacts = db.prepare("SELECT * FROM contacts ORDER BY created_at DESC LIMIT 100").all();
  res.json({ ok: true, contacts });
});

// Update contact status
app.put('/api/contacts/:id', authMiddleware('admin'), (req, res) => {
  const { status } = req.body;
  db.prepare("UPDATE contacts SET status = ? WHERE id = ?").run(status, req.params.id);
  res.json({ ok: true });
});

// Dashboard stats
app.get('/api/admin/stats', authMiddleware('admin'), (req, res) => {
  const today = new Date().toISOString().split('T')[0];
  
  const stats = {
    bookings: {
      total: db.prepare("SELECT COUNT(*) as c FROM bookings").get().c,
      upcoming: db.prepare("SELECT COUNT(*) as c FROM bookings WHERE check_in >= ?").get(today).c,
      pending: db.prepare("SELECT COUNT(*) as c FROM bookings WHERE status = 'pending'").get().c
    },
    contacts: {
      total: db.prepare("SELECT COUNT(*) as c FROM contacts").get().c,
      new: db.prepare("SELECT COUNT(*) as c FROM contacts WHERE status = 'nieuw'").get().c
    },
    rooms: {
      total: db.prepare("SELECT COUNT(*) as c FROM rooms").get().c
    }
  };
  
  res.json({ ok: true, stats });
});

// ═══════════════════════════════════════════════════════════════════════════
// BLOG API ROUTES
// ═══════════════════════════════════════════════════════════════════════════

// Public: Get published blog posts
app.get('/api/blog', (req, res) => {
  const { category, limit = 10, offset = 0 } = req.query;
  let query = "SELECT id, title, slug, excerpt, category, tags, featured_image, author, published_at FROM blog_posts WHERE status = 'published'";
  const params = [];
  
  if (category) {
    query += " AND category = ?";
    params.push(category);
  }
  
  query += " ORDER BY published_at DESC LIMIT ? OFFSET ?";
  params.push(parseInt(limit), parseInt(offset));
  
  const posts = db.prepare(query).all(...params);
  const total = db.prepare("SELECT COUNT(*) as c FROM blog_posts WHERE status = 'published'").get().c;
  
  res.json({ ok: true, posts, total });
});

// Public: Get single blog post by slug
app.get('/api/blog/:slug', (req, res) => {
  const post = db.prepare("SELECT * FROM blog_posts WHERE slug = ? AND status = 'published'").get(req.params.slug);
  if (!post) {
    return res.status(404).json({ error: 'Artikel niet gevonden' });
  }
  res.json({ ok: true, post });
});

// Public: Get blog categories
app.get('/api/blog-categories', (req, res) => {
  const categories = db.prepare("SELECT * FROM blog_categories ORDER BY name").all();
  res.json({ ok: true, categories });
});

// Admin: Get all blog posts (including drafts)
app.get('/api/admin/blog', authMiddleware('admin'), (req, res) => {
  const posts = db.prepare("SELECT * FROM blog_posts ORDER BY created_at DESC").all();
  res.json({ ok: true, posts });
});

// Admin: Get single blog post for editing
app.get('/api/admin/blog/:id', authMiddleware('admin'), (req, res) => {
  const post = db.prepare("SELECT * FROM blog_posts WHERE id = ?").get(req.params.id);
  if (!post) {
    return res.status(404).json({ error: 'Artikel niet gevonden' });
  }
  res.json({ ok: true, post });
});

// Admin: Create blog post
app.post('/api/admin/blog', authMiddleware('admin'), (req, res) => {
  const { title, slug, excerpt, content, category, tags, featured_image, status } = req.body;
  
  if (!title || !content) {
    return res.status(400).json({ error: 'Titel en inhoud zijn verplicht' });
  }
  
  // Generate slug if not provided
  const finalSlug = slug || title.toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-|-$/g, '');
  
  try {
    const published_at = status === 'published' ? new Date().toISOString() : null;
    const result = db.prepare(`
      INSERT INTO blog_posts (title, slug, excerpt, content, category, tags, featured_image, status, published_at)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).run(title, finalSlug, excerpt || null, content, category || null, tags || null, featured_image || null, status || 'draft', published_at);
    
    res.json({ ok: true, id: result.lastInsertRowid, slug: finalSlug });
  } catch (e) {
    if (e.message.includes('UNIQUE constraint')) {
      return res.status(400).json({ error: 'Deze slug bestaat al' });
    }
    res.status(500).json({ error: 'Opslaan mislukt' });
  }
});

// Admin: Update blog post
app.put('/api/admin/blog/:id', authMiddleware('admin'), (req, res) => {
  const { title, slug, excerpt, content, category, tags, featured_image, status } = req.body;
  
  const existing = db.prepare("SELECT * FROM blog_posts WHERE id = ?").get(req.params.id);
  if (!existing) {
    return res.status(404).json({ error: 'Artikel niet gevonden' });
  }
  
  // Set published_at when first publishing
  let published_at = existing.published_at;
  if (status === 'published' && !existing.published_at) {
    published_at = new Date().toISOString();
  }
  
  try {
    db.prepare(`
      UPDATE blog_posts 
      SET title = ?, slug = ?, excerpt = ?, content = ?, category = ?, tags = ?, 
          featured_image = ?, status = ?, published_at = ?, updated_at = CURRENT_TIMESTAMP
      WHERE id = ?
    `).run(title, slug, excerpt, content, category, tags, featured_image, status, published_at, req.params.id);
    
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: 'Update mislukt' });
  }
});

// Admin: Delete blog post
app.delete('/api/admin/blog/:id', authMiddleware('admin'), (req, res) => {
  db.prepare("DELETE FROM blog_posts WHERE id = ?").run(req.params.id);
  res.json({ ok: true });
});

// Admin: Manage categories
app.post('/api/admin/blog-categories', authMiddleware('admin'), (req, res) => {
  const { name } = req.body;
  const slug = name.toLowerCase().replace(/[^a-z0-9]+/g, '-');
  
  try {
    db.prepare("INSERT INTO blog_categories (name, slug) VALUES (?, ?)").run(name, slug);
    res.json({ ok: true });
  } catch (e) {
    res.status(400).json({ error: 'Categorie bestaat al' });
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// REDIRECTS (oude URLs naar nieuwe)
// ═══════════════════════════════════════════════════════════════════════════

app.get('/feestzaal', (req, res) => res.redirect(301, '/feesten'));
app.get('/zakelijk', (req, res) => res.redirect(301, '/teams'));
app.get('/privefeesten', (req, res) => res.redirect(301, '/feesten'));
app.get('/trouwfeest', (req, res) => res.redirect(301, '/feesten#trouwfeest'));
app.get('/communiefeest', (req, res) => res.redirect(301, '/feesten#familiefeest'));

// ═══════════════════════════════════════════════════════════════════════════
// PAGE ROUTES (SPA-style routing)
// ═══════════════════════════════════════════════════════════════════════════

// Serve specific HTML pages
const pages = ['feestzaal', 'bb', 'zakelijk', 'teams', 'feesten', 'over-ons', 'contact', 'offerte', 'privacy', 'algemene-voorwaarden', 'login', 'admin', 'blog'];

pages.forEach(page => {
  app.get(`/${page}`, (req, res) => {
    const file = path.join(__dirname, '..', 'public', `${page}.html`);
    if (fs.existsSync(file)) {
      res.sendFile(file);
    } else {
      // Fallback to index for SPA behavior
      res.sendFile(path.join(__dirname, '..', 'public', 'index.html'));
    }
  });
});

// Guest routes
app.get('/gasten/login', (req, res) => {
  const file = path.join(__dirname, '..', 'public', 'gasten', 'login.html');
  if (fs.existsSync(file)) {
    res.sendFile(file);
  } else {
    res.sendFile(path.join(__dirname, '..', 'public', 'login.html'));
  }
});

app.get('/gasten/dashboard', (req, res) => {
  const file = path.join(__dirname, '..', 'public', 'gasten', 'dashboard.html');
  if (fs.existsSync(file)) {
    res.sendFile(file);
  } else {
    res.status(404).send('Dashboard not found');
  }
});

// B&B room pages
app.get('/bb/:slug', (req, res) => {
  const file = path.join(__dirname, '..', 'public', 'bb', `${req.params.slug}.html`);
  if (fs.existsSync(file)) {
    res.sendFile(file);
  } else {
    res.sendFile(path.join(__dirname, '..', 'public', 'bb.html'));
  }
});

// Ruimtes pages
const ruimtes = ['panoramische-zaal', 'polyvalente-zaal', 'vergaderzaal', 'zolderzalen', 'pianobar'];
ruimtes.forEach(ruimte => {
  app.get(`/ruimtes/${ruimte}`, (req, res) => {
    const file = path.join(__dirname, '..', 'public', 'ruimtes', `${ruimte}.html`);
    if (fs.existsSync(file)) {
      res.sendFile(file);
    } else {
      res.sendFile(path.join(__dirname, '..', 'public', 'teams.html'));
    }
  });
});

// Blog article pages — server-side SEO meta injectie per artikel
app.get('/blog/:slug', (req, res) => {
  const templatePath = path.join(__dirname, '..', 'public', 'blog-article.html');
  const post = db.prepare("SELECT * FROM blog_posts WHERE slug = ? AND status = 'published'").get(req.params.slug);
  if (!post) {
    return res.sendFile(templatePath);
  }
  let html = fs.readFileSync(templatePath, 'utf-8');
  const esc = s => String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const title = esc(post.title) + ' · Blog · domein Walbrugge';
  const desc = esc(post.excerpt || (post.content || '').replace(/<[^>]*>/g, '').slice(0, 155));
  const url = `https://walbrugge.be/blog/${post.slug}`;
  const img = post.image ? (post.image.startsWith('http') ? post.image : 'https://walbrugge.be' + post.image) : 'https://walbrugge.be/assets/img/og-image.jpg';
  const articleLd = JSON.stringify({
    '@context': 'https://schema.org',
    '@type': 'BlogPosting',
    headline: post.title,
    description: post.excerpt || undefined,
    image: img,
    url,
    datePublished: post.published_at || post.created_at,
    dateModified: post.updated_at || post.published_at || post.created_at,
    author: { '@type': 'Organization', name: 'Domein Walbrugge', url: 'https://walbrugge.be' },
    publisher: { '@type': 'Organization', name: 'Domein Walbrugge', logo: { '@type': 'ImageObject', url: 'https://walbrugge.be/assets/img/logo.png' } },
    mainEntityOfPage: url
  });
  const seoBlock = `<title>${title}</title>
<meta name="description" content="${desc}">
<link rel="canonical" href="${url}">
<meta property="og:title" content="${title}">
<meta property="og:description" content="${desc}">
<meta property="og:image" content="${img}">
<meta property="og:url" content="${url}">
<meta property="og:type" content="article">
<meta property="og:locale" content="nl_BE">
<meta property="og:site_name" content="Domein Walbrugge">
<meta name="twitter:card" content="summary_large_image">
<meta name="twitter:title" content="${title}">
<meta name="twitter:description" content="${desc}">
<meta name="twitter:image" content="${img}">
<script type="application/ld+json">${articleLd}</script>`;
  // Vervang bestaande <title> en injecteer de rest vóór </head>
  html = html.replace(/<title>[\s\S]*?<\/title>/, '');
  html = html.replace(/<meta name="description"[^>]*>/, '');
  html = html.replace('</head>', seoBlock + '\n</head>');
  res.send(html);
});

// ═══════════════════════════════════════════════════════════════════════════
// TAALVERSIES (FR/EN) — vertaalde pagina's in public/fr en public/en
// ═══════════════════════════════════════════════════════════════════════════

const languages = ['fr', 'en', 'de'];

// Gepensioneerde URLs: in NL redirecten deze al naar de nieuwe pagina.
// Per taal hetzelfde gedrag, zodat /fr/feestzaal niet langer NL-content op 200 serveert.
const retiredPages = { feestzaal: 'feesten', zakelijk: 'teams' };

languages.forEach(lang => {
  // Serveer enkel de vertaalde pagina. Bestaat die niet, dan 301 naar de NL-versie:
  // nooit NL-content onder een anderstalige URL met status 200 (duplicate content).
  const serveLang = (relPath, nlPath) => (req, res) => {
    const file = path.join(__dirname, '..', 'public', lang, relPath);
    if (fs.existsSync(file)) {
      return res.sendFile(file);
    }
    return res.redirect(301, nlPath);
  };

  app.get(`/${lang}`, serveLang('index.html', '/'));

  // Eerst de gepensioneerde URLs: 301 naar het anderstalige equivalent.
  Object.entries(retiredPages).forEach(([from, to]) => {
    app.get(`/${lang}/${from}`, (req, res) => res.redirect(301, `/${lang}/${to}`));
  });

  pages.forEach(page => {
    if (retiredPages[page]) return; // al afgehandeld hierboven
    app.get(`/${lang}/${page}`, serveLang(`${page}.html`, `/${page}`));
  });

  ruimtes.forEach(ruimte => {
    app.get(`/${lang}/ruimtes/${ruimte}`, serveLang(path.join('ruimtes', `${ruimte}.html`), `/ruimtes/${ruimte}`));
  });
});

// Catch-all: serve index.html
app.get('*', (req, res) => {
  // Skip API routes
  if (req.path.startsWith('/api/')) {
    return res.status(404).json({ error: 'Route niet gevonden' });
  }
  
  res.sendFile(path.join(__dirname, '..', 'public', 'index.html'));
});

// ═══════════════════════════════════════════════════════════════════════════
// ERROR HANDLING
// ═══════════════════════════════════════════════════════════════════════════

app.use((err, req, res, next) => {
  console.error('Server error:', err);
  res.status(500).json({ error: 'Server error' });
});

// ═══════════════════════════════════════════════════════════════════════════
// START SERVER
// ═══════════════════════════════════════════════════════════════════════════

app.listen(PORT, '127.0.0.1', () => {
  console.log(`
  ╔════════════════════════════════════════════════╗
  ║  🏡 Walbrugge Backend Server                   ║
  ║  Running on http://127.0.0.1:${PORT}             ║
  ╚════════════════════════════════════════════════╝
  `);
});

// Graceful shutdown
process.on('SIGTERM', () => {
  console.log('Shutting down...');
  db.close();
  process.exit(0);
});
