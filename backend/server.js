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

// ═══════════════════════════════════════════════════════════════════════════
// OFFERTEAANVRAAG PER E-MAIL
// ═══════════════════════════════════════════════════════════════════════════
// Instellingen komen uit de omgeving (zie /etc/walbrugge.env op de server).
// Ontbreken de inloggegevens, dan wordt er niets verstuurd en blijft de
// aanvraag gewoon in de database staan.

const MAIL_TO = process.env.MAIL_TO || 'info@walbrugge.be';
const MAIL_FROM = process.env.MAIL_FROM || 'info@walbrugge.be';

// ── Verzendweg 1 (voorkeur): Microsoft Graph met OAuth2 ────────────────────
// Client credentials flow: de server authenticeert als toepassing, niet als
// gebruiker. Geen wachtwoord, geen MFA-omzeiling, intrekbaar in Entra ID.
// Vereist in /etc/walbrugge.env: GRAPH_TENANT_ID, GRAPH_CLIENT_ID,
// GRAPH_CLIENT_SECRET. Optioneel GRAPH_SENDER (standaard MAIL_FROM).
const GRAPH_TENANT_ID = process.env.GRAPH_TENANT_ID;
const GRAPH_CLIENT_ID = process.env.GRAPH_CLIENT_ID;
const GRAPH_CLIENT_SECRET = process.env.GRAPH_CLIENT_SECRET;
const GRAPH_SENDER = process.env.GRAPH_SENDER || MAIL_FROM;
const graphActief = !!(GRAPH_TENANT_ID && GRAPH_CLIENT_ID && GRAPH_CLIENT_SECRET);

// Tokens zijn ~1 uur geldig; we hergebruiken ze tot 60 s voor het verlopen.
let graphToken = null;
let graphTokenVervalt = 0;

async function graphAccessToken() {
  if (graphToken && Date.now() < graphTokenVervalt) return graphToken;

  const body = new URLSearchParams({
    client_id: GRAPH_CLIENT_ID,
    client_secret: GRAPH_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });

  const res = await fetch(
    'https://login.microsoftonline.com/' + encodeURIComponent(GRAPH_TENANT_ID) + '/oauth2/v2.0/token',
    { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body }
  );

  const data = await res.json().catch(() => ({}));
  if (!res.ok || !data.access_token) {
    throw new Error('token ophalen mislukt (' + res.status + '): ' +
      (data.error_description || data.error || 'onbekende fout'));
  }

  graphToken = data.access_token;
  graphTokenVervalt = Date.now() + ((data.expires_in || 3600) - 60) * 1000;
  return graphToken;
}

async function graphSendMail({ subject, text, html, replyToAdres, replyToNaam, to }) {
  const token = await graphAccessToken();

  const ontvangers = (to && to.length ? to : [MAIL_TO]);

  const bericht = {
    message: {
      subject,
      body: { contentType: 'HTML', content: html },
      toRecipients: ontvangers.map(a => ({ emailAddress: { address: a } }))
    },
    saveToSentItems: false
  };

  if (replyToAdres) {
    bericht.message.replyTo = [{
      emailAddress: replyToNaam
        ? { address: replyToAdres, name: replyToNaam }
        : { address: replyToAdres }
    }];
  }

  const res = await fetch(
    'https://graph.microsoft.com/v1.0/users/' + encodeURIComponent(GRAPH_SENDER) + '/sendMail',
    {
      method: 'POST',
      headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify(bericht)
    }
  );

  // Graph antwoordt met 202 Accepted en een lege body bij succes.
  if (!res.ok) {
    const fout = await res.text().catch(() => '');
    throw new Error('Graph sendMail gaf ' + res.status + ': ' + fout.slice(0, 300));
  }
}

// ── Verzendweg 2 (terugval): klassieke SMTP ────────────────────────────────
const mailer = (process.env.SMTP_USER && process.env.SMTP_PASS)
  ? nodemailer.createTransport({
      host: process.env.SMTP_HOST || 'smtp.office365.com',
      port: parseInt(process.env.SMTP_PORT || '587', 10),
      secure: false,
      requireTLS: true,
      auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS }
    })
  : null;

if (graphActief) {
  console.log('[mail] Microsoft Graph (OAuth2) actief — afzender ' + GRAPH_SENDER);
} else if (mailer) {
  console.log('[mail] SMTP actief — afzender ' + MAIL_FROM);
} else {
  console.warn('[mail] Geen Graph- of SMTP-gegevens — offerteaanvragen worden niet gemaild ' +
    '(ze blijven wel in de database staan).');
}

const TYPE_LABELS = {
  vergadering: 'Vergadering / meeting',
  teambuilding: 'Teambuilding',
  seminarie: 'Seminarie / workshop',
  strategiedag: 'Strategiedag',
  'exclusief-domein': 'Exclusief domein',
  trouwfeest: 'Trouwfeest',
  communie: 'Communie / lentefeest',
  verjaardag: 'Verjaardag / jubileum',
  bedrijfsfeest: 'Bedrijfsfeest',
  'ander-feest': 'Ander feest',
  andere: 'Andere'
};

const FORMULE_LABELS = {
  vergadermiddag: 'Vergadermiddag',
  teamdag: 'Teamdag',
  '24uur': '24-uur formule',
  exclusief: 'Exclusief domein',
  'op-maat': 'Op maat',
  receptie: 'Enkel receptie',
  'walking-dinner': 'Walking dinner',
  'zittend-diner': 'Zittend diner',
  'andere-formule': 'Andere formule'
};

function verstuurOfferteMail(c) {
  if (!graphActief && !mailer) return;

  const esc = s => String(s === null || s === undefined || s === '' ? '—' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  const velden = [
    ['Contactpersoon', c.naam],
    ['Bedrijf', c.bedrijf],
    ['E-mail', c.email],
    ['Telefoon', c.telefoon],
    ['Type evenement', TYPE_LABELS[c.type] || c.type],
    ['Aantal personen', c.personen],
    ['Gewenste datum', c.datum],
    ['Gewenste formule', FORMULE_LABELS[c.formule] || c.formule],
    ['Bericht', c.bericht],
    ['Taal / pagina', c.taal],
    ['Ontvangen op', new Date().toLocaleString('nl-BE', { timeZone: 'Europe/Brussels' })]
  ];

  const tekst = velden.map(([k, v]) => k + ': ' + (v || '—')).join('\n');
  const rijen = velden.map(([k, v]) =>
    '<tr><th align="left" style="padding:6px 14px 6px 0;vertical-align:top;white-space:nowrap;' +
    'font-weight:600;color:#5b5b5b;">' + esc(k) + '</th>' +
    '<td style="padding:6px 0;vertical-align:top;">' + esc(v).replace(/\n/g, '<br>') + '</td></tr>'
  ).join('');

  const html =
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222;">' +
    '<h2 style="margin:0 0 4px;font-size:18px;">Nieuwe offerteaanvraag</h2>' +
    '<p style="margin:0 0 16px;color:#777;">via walbrugge.be</p>' +
    '<table cellpadding="0" cellspacing="0">' + rijen + '</table>' +
    '<p style="margin-top:18px;color:#777;font-size:12px;">Antwoord op deze mail om ' +
    'rechtstreeks naar de aanvrager te mailen.</p></div>';

  const onderwerp = 'Offerteaanvraag — ' + (TYPE_LABELS[c.type] || c.type || 'algemeen') +
                    ' — ' + (c.naam || 'onbekend');

  const verzonden = graphActief
    ? graphSendMail({
        subject: onderwerp,
        text: tekst,
        html: html,
        replyToAdres: c.email || null,
        replyToNaam: c.naam || null
      }).then(() => 'Graph')
    : mailer.sendMail({
        from: '"Domein Walbrugge" <' + MAIL_FROM + '>',
        to: MAIL_TO,
        replyTo: c.email ? (c.naam ? '"' + c.naam + '" <' + c.email + '>' : c.email) : undefined,
        subject: onderwerp,
        text: tekst,
        html: html
      }).then(() => 'SMTP');

  verzonden.then(via => {
    console.log('[mail] Offerteaanvraag verstuurd naar ' + MAIL_TO + ' via ' + via);
  }).catch(err => {
    // Nooit de aanvraag laten sneuvelen op een mailfout: hij staat al in de database.
    console.error('[mail] Versturen mislukt:', err.message);
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// WAARSCHUWING: CLIENT SECRET VAN DE GRAPH-APP VERVALT
// ═══════════════════════════════════════════════════════════════════════════
// Zet in /etc/walbrugge.env de vervaldatum van het secret:
//   GRAPH_SECRET_EXPIRES=JJJJ-MM-DD
//   ALERT_TO=info@walbrugge.be,bartelcoucke@renoperfect.be   (optioneel)
//   GRAPH_SECRET_WARN_DAYS=14                                (optioneel)
// Vanaf 14 dagen voor die datum vertrekt er elke dag een mail, tot er een
// nieuw secret met een nieuwe vervaldatum is ingesteld. Eén mail per dag:
// data/secret-waarschuwing.txt onthoudt wanneer er laatst één vertrok.

const SECRET_VERVALT = (process.env.GRAPH_SECRET_EXPIRES || '').trim();
const WAARSCHUW_DAGEN = parseInt(process.env.GRAPH_SECRET_WARN_DAYS || '14', 10);
const ALERT_TO = (process.env.ALERT_TO || 'info@walbrugge.be,bartelcoucke@renoperfect.be')
  .split(',').map(s => s.trim()).filter(Boolean);

if (graphActief && !SECRET_VERVALT) {
  console.warn('[mail] GRAPH_SECRET_EXPIRES niet ingesteld — er komt geen waarschuwing ' +
    'wanneer het client secret vervalt.');
}

function dagenTot(datum) {
  const d = new Date(datum + 'T00:00:00Z');
  return isNaN(d.getTime()) ? null : Math.ceil((d.getTime() - Date.now()) / 86400000);
}

async function controleerSecretVervaldatum() {
  if (!graphActief || !SECRET_VERVALT) return;

  const dagen = dagenTot(SECRET_VERVALT);
  if (dagen === null) {
    console.warn('[mail] GRAPH_SECRET_EXPIRES is geen geldige datum (JJJJ-MM-DD): ' + SECRET_VERVALT);
    return;
  }
  if (dagen > WAARSCHUW_DAGEN) return;

  const bestand = path.join(__dirname, '..', 'data', 'secret-waarschuwing.txt');
  const vandaag = new Date().toISOString().slice(0, 10);
  try {
    if (fs.readFileSync(bestand, 'utf-8').trim() === vandaag) return;
  } catch (e) { /* vandaag nog niets verstuurd */ }

  const verlopen = dagen < 0;
  const onderwerp = verlopen
    ? 'VERLOPEN: het client secret van de website werkt niet meer'
    : 'Nog ' + dagen + ' dag' + (dagen === 1 ? '' : 'en') + ': client secret website vervalt';

  const esc = s => String(s === null || s === undefined || s === '' ? '—' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  const html =
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222;">' +
    '<h2 style="margin:0 0 4px;font-size:18px;color:' +
    (verlopen || dagen <= 3 ? '#b3261e' : '#8a6d00') + ';">' + esc(onderwerp) + '</h2>' +
    '<p style="margin:0 0 16px;color:#777;">App-registratie “Walbrugge website mail” in Microsoft Entra</p>' +
    '<p style="margin:0 0 6px;"><strong>Vervaldatum:</strong> ' + esc(SECRET_VERVALT) + '</p>' +
    '<p style="margin:0 0 16px;">' + (verlopen
      ? 'Offerteaanvragen worden op dit moment <strong>niet meer gemaild</strong>. Ze komen ' +
        'nog wel in het beheerpaneel terecht, maar zonder verwittiging.'
      : 'Zodra het secret vervalt, komen offerteaanvragen alleen nog in het beheerpaneel ' +
        'terecht, zonder verwittiging.') + '</p>' +
    '<p style="margin:0 0 6px;"><strong>Wat te doen</strong></p>' +
    '<ol style="margin:6px 0 0;padding-left:20px;line-height:1.6;">' +
    '<li>Entra → App-registraties → “Walbrugge website mail” → Certificaten en geheimen → Nieuw clientgeheim.</li>' +
    '<li>Kopieer de <em>Waarde</em> — die is maar één keer zichtbaar.</li>' +
    '<li>Zet op de server in <code>/etc/walbrugge.env</code> de nieuwe <code>GRAPH_CLIENT_SECRET</code> ' +
    'en de nieuwe <code>GRAPH_SECRET_EXPIRES</code> (JJJJ-MM-DD).</li>' +
    '<li><code>systemctl restart walbrugge</code></li>' +
    '<li>Verwijder daarna het oude geheim in Entra.</li>' +
    '</ol>' +
    '<p style="margin-top:18px;color:#777;font-size:12px;">Deze herinnering komt elke dag terug ' +
    'tot de nieuwe vervaldatum is ingesteld.</p></div>';

  const tekst = onderwerp + '\n\nVervaldatum: ' + SECRET_VERVALT +
    '\n\nNieuw clientgeheim maken in Entra, GRAPH_CLIENT_SECRET en GRAPH_SECRET_EXPIRES ' +
    'aanpassen in /etc/walbrugge.env, daarna: systemctl restart walbrugge';

  try {
    if (graphActief) {
      await graphSendMail({ subject: onderwerp, text: tekst, html: html, to: ALERT_TO });
    } else {
      await mailer.sendMail({
        from: '"Domein Walbrugge" <' + MAIL_FROM + '>',
        to: ALERT_TO.join(', '),
        subject: onderwerp,
        text: tekst,
        html: html
      });
    }
    console.log('[mail] Waarschuwing client secret verstuurd naar ' + ALERT_TO.join(', '));
  } catch (err) {
    console.error('[mail] Waarschuwing client secret mislukt:', err.message);
    return;
  }

  // Pas noteren nadat de mail effectief vertrokken is, zodat een mislukte
  // poging later op de dag opnieuw geprobeerd wordt.
  try {
    fs.mkdirSync(path.dirname(bestand), { recursive: true });
    fs.writeFileSync(bestand, vandaag);
  } catch (err) {
    console.error('[mail] Kon ' + bestand + ' niet schrijven:', err.message);
  }
}

// Eerste controle 30 s na de start, daarna om de zes uur.
setTimeout(controleerSecretVervaldatum, 30000);
setInterval(controleerSecretVervaldatum, 6 * 60 * 60 * 1000);


const app = express();
app.disable('x-powered-by');
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

  -- Tweestapsverificatie: toestellen die de beheerder al bevestigd heeft.
  -- We bewaren enkel een hash van het toestelgeheim, nooit het geheim zelf.
  CREATE TABLE IF NOT EXISTS trusted_devices (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    token_hash TEXT UNIQUE NOT NULL,
    label TEXT,
    user_agent TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    last_used_at DATETIME
  );

  -- Openstaande inlogcodes. Kortstondig: opgeruimd na gebruik of verval.
  CREATE TABLE IF NOT EXISTS login_codes (
    challenge TEXT PRIMARY KEY,
    user_id INTEGER NOT NULL,
    code_hash TEXT NOT NULL,
    expires_at INTEGER NOT NULL,
    attempts INTEGER DEFAULT 0,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
  );
`);

// Insert admin user if not exists
// Geen vast wachtwoord in de broncode: zet ADMIN_PASSWORD in de omgeving,
// of gebruik  node set-admin-password.js  om er zelf een te kiezen.
const startWachtwoord = process.env.ADMIN_PASSWORD || crypto.randomBytes(24).toString('hex');
if (!process.env.ADMIN_PASSWORD) {
  console.warn('[auth] Geen ADMIN_PASSWORD ingesteld. Een nieuwe beheerder krijgt een willekeurig wachtwoord; kies er zelf een met: node set-admin-password.js');
}
// Het beheerdersadres is info@walbrugge.be: dat postvak bestaat echt, en de
// inlogcode van de tweestapsverificatie moet aankomen. Bestaande installaties
// stonden op admin@walbrugge.be — dat postvak bestaat niet — dus die naam
// wordt hier omgezet met behoud van het wachtwoord.
const ADMIN_EMAIL = process.env.ADMIN_EMAIL || 'info@walbrugge.be';
const OUD_ADMIN_EMAIL = 'admin@walbrugge.be';

const oudeAdmin = db.prepare('SELECT id FROM users WHERE email = ?').get(OUD_ADMIN_EMAIL);
const nieuweAdmin = db.prepare('SELECT id FROM users WHERE email = ?').get(ADMIN_EMAIL);

if (oudeAdmin && !nieuweAdmin) {
  db.prepare('UPDATE users SET email = ? WHERE id = ?').run(ADMIN_EMAIL, oudeAdmin.id);
  console.log('[auth] Beheerdersadres omgezet naar ' + ADMIN_EMAIL + ' (wachtwoord blijft gelden)');
} else if (oudeAdmin && nieuweAdmin) {
  // Beide bestaan al: de oude is een restant en mag weg.
  db.prepare('DELETE FROM users WHERE id = ?').run(oudeAdmin.id);
  console.log('[auth] Oude beheerder ' + OUD_ADMIN_EMAIL + ' verwijderd');
}

if (!db.prepare('SELECT id FROM users WHERE email = ?').get(ADMIN_EMAIL)) {
  const hash = bcrypt.hashSync(startWachtwoord, 12);
  db.prepare("INSERT INTO users (email, password_hash, role, name) VALUES (?, ?, 'admin', 'Administrator')")
    .run(ADMIN_EMAIL, hash);
  console.log('Admin user created: ' + ADMIN_EMAIL);
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
app.use(express.static(path.join(__dirname, '..', 'public'), {
  setHeaders: (res) => {
    res.removeHeader('Cache-Control');  // Caddy regelt caching — voorkomt dubbele headers
  },
  etag: true,
  lastModified: true
}));

// Logging
app.use((req, res, next) => {
  console.log(`${new Date().toISOString()} ${req.method} ${req.path}`);
  next();
});

// ═══════════════════════════════════════════════════════════════════════════
// TWEESTAPSVERIFICATIE (2FA) VOOR BEHEERDERS
// ═══════════════════════════════════════════════════════════════════════════
// Beheerders krijgen na e-mail + wachtwoord een code van 6 cijfers per mail.
// Bevestigt de beheerder het toestel, dan krijgt dat toestel een geheim dat
// hier gehasht bewaard wordt; bij een volgende login vervalt de codestap.
// Het adres voor de code is standaard dat van de gebruiker zelf; met
// ADMIN_2FA_EMAIL in /etc/walbrugge.env kan een ander postvak gekozen worden.

const CODE_GELDIG_MS = 10 * 60 * 1000;   // code vervalt na 10 minuten
const CODE_MAX_POGINGEN = 5;             // daarna is de code verbrand

function sha256(waarde) {
  return crypto.createHash('sha256').update(String(waarde)).digest('hex');
}

// Vergelijking in constante tijd, zodat de duur van het antwoord niets prijsgeeft.
function hashesGelijk(a, b) {
  const bufA = Buffer.from(String(a), 'utf8');
  const bufB = Buffer.from(String(b), 'utf8');
  if (bufA.length !== bufB.length) return false;
  return crypto.timingSafeEqual(bufA, bufB);
}

function ruimVervallenCodesOp() {
  db.prepare('DELETE FROM login_codes WHERE expires_at < ?').run(Date.now());
}

// Geeft het toestel terug als het geheim klopt, anders null.
function zoekVertrouwdToestel(userId, deviceToken) {
  if (!deviceToken || typeof deviceToken !== 'string') return null;
  return db.prepare('SELECT * FROM trusted_devices WHERE user_id = ? AND token_hash = ?')
    .get(userId, sha256(deviceToken)) || null;
}

// Maakt een nieuw toestelgeheim aan. De onversleutelde waarde gaat één keer
// naar de browser; wij houden enkel de hash bij.
function onthoudToestel(userId, userAgent) {
  const geheim = crypto.randomBytes(32).toString('hex');
  db.prepare(`INSERT INTO trusted_devices (user_id, token_hash, label, user_agent, last_used_at)
              VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)`)
    .run(userId, sha256(geheim), toestelLabel(userAgent), (userAgent || '').slice(0, 255));
  return geheim;
}

// Korte, herkenbare omschrijving zodat de beheerder zijn toestellen uit elkaar houdt.
function toestelLabel(userAgent) {
  const ua = userAgent || '';
  const systeem = /Windows/i.test(ua) ? 'Windows'
    : /iPhone|iPad|iOS/i.test(ua) ? 'iPhone/iPad'
    : /Android/i.test(ua) ? 'Android'
    : /Mac OS X|Macintosh/i.test(ua) ? 'Mac'
    : /Linux/i.test(ua) ? 'Linux' : 'Onbekend toestel';
  const browser = /Edg\//i.test(ua) ? 'Edge'
    : /OPR\//i.test(ua) ? 'Opera'
    : /Chrome\//i.test(ua) ? 'Chrome'
    : /Safari\//i.test(ua) ? 'Safari'
    : /Firefox\//i.test(ua) ? 'Firefox' : 'browser';
  return systeem + ' · ' + browser;
}

async function stuurInlogcode(naarAdres, code) {
  const onderwerp = 'Inlogcode Walbrugge: ' + code;
  const tekst = 'Uw inlogcode voor het beheer van walbrugge.be is ' + code + '.\n' +
                'De code blijft 10 minuten geldig.\n\n' +
                'Hebt u zelf niet proberen in te loggen, wijzig dan uw wachtwoord.';
  const html =
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222;">' +
    '<h2 style="margin:0 0 4px;font-size:18px;">Inlogcode beheer</h2>' +
    '<p style="margin:0 0 16px;color:#777;">walbrugge.be</p>' +
    '<p style="font-size:30px;letter-spacing:6px;font-weight:bold;margin:0 0 12px;">' + code + '</p>' +
    '<p style="margin:0 0 16px;">De code blijft 10 minuten geldig.</p>' +
    '<p style="color:#777;font-size:12px;">Hebt u zelf niet proberen in te loggen, ' +
    'wijzig dan meteen uw wachtwoord.</p></div>';

  if (graphActief) {
    await graphSendMail({ subject: onderwerp, text: tekst, html, naar: naarAdres });
    return 'Graph';
  }
  if (mailer) {
    await mailer.sendMail({
      from: '"Domein Walbrugge" <' + MAIL_FROM + '>',
      to: naarAdres,
      subject: onderwerp,
      text: tekst,
      html
    });
    return 'SMTP';
  }
  throw new Error('geen mailweg ingesteld');
}

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

    // Beheerders doorlopen een tweede stap, tenzij dit toestel al bevestigd is.
    if (user.role === 'admin') {
      const toestel = zoekVertrouwdToestel(user.id, req.body.deviceToken);
      if (toestel) {
        db.prepare('UPDATE trusted_devices SET last_used_at = CURRENT_TIMESTAMP WHERE id = ?')
          .run(toestel.id);
      } else {
        return startTweedeStap(user, req, res);
      }
    }

    return res.json(maakLoginAntwoord(user));
  }

  return res.status(400).json({ error: 'Vul e-mail en wachtwoord of boekingsreferentie in' });
});

// Bouwt het gewone, geslaagde login-antwoord.
function maakLoginAntwoord(user, extra) {
  const token = jwt.sign({
    id: user.id,
    email: user.email,
    role: user.role,
    name: user.name
  }, JWT_SECRET, { expiresIn: '24h' });

  return Object.assign({
    ok: true,
    token,
    user: { email: user.email, name: user.name, role: user.role },
    redirect: user.role === 'admin' ? '/admin' : '/gasten/dashboard'
  }, extra || {});
}

// Maakt een code van 6 cijfers, mailt die en antwoordt met een challenge-id.
// Het wachtwoord is op dit punt al gecontroleerd; de challenge alleen is
// waardeloos zonder de code uit de mailbox.
function startTweedeStap(user, req, res) {
  ruimVervallenCodesOp();

  const naarAdres = process.env.ADMIN_2FA_EMAIL || user.email;
  const code = String(crypto.randomInt(0, 1000000)).padStart(6, '0');
  const challenge = crypto.randomBytes(24).toString('hex');

  db.prepare(`INSERT INTO login_codes (challenge, user_id, code_hash, expires_at)
              VALUES (?, ?, ?, ?)`)
    .run(challenge, user.id, sha256(code), Date.now() + CODE_GELDIG_MS);

  stuurInlogcode(naarAdres, code)
    .then(via => console.log('[2fa] Inlogcode verstuurd naar ' + naarAdres + ' via ' + via))
    .catch(err => {
      // Zonder deze regel zou een mailstoring de beheerder buitensluiten.
      // De logs zijn enkel voor root leesbaar; de code vervalt na 10 minuten.
      console.error('[2fa] Versturen mislukt:', err.message);
      console.error('[2fa] Noodcode voor ' + user.email + ': ' + code);
    });

  return res.json({
    ok: true,
    twofa: true,
    challenge,
    hint: maskeerAdres(naarAdres)
  });
}

// info@walbrugge.be → i••••@walbrugge.be
function maskeerAdres(adres) {
  const [naam, domein] = String(adres).split('@');
  if (!domein) return '';
  return naam.slice(0, 1) + '•'.repeat(Math.max(naam.length - 1, 1)) + '@' + domein;
}

// Tweede stap: challenge + code inwisselen voor een token.
app.post('/api/login/2fa', (req, res) => {
  const { challenge, code, remember } = req.body || {};

  if (!challenge || !code) {
    return res.status(400).json({ error: 'Vul de code in' });
  }

  ruimVervallenCodesOp();
  const rij = db.prepare('SELECT * FROM login_codes WHERE challenge = ?').get(challenge);

  if (!rij || rij.expires_at < Date.now()) {
    return res.status(401).json({ error: 'De code is verlopen. Log opnieuw in.' });
  }

  if (rij.attempts >= CODE_MAX_POGINGEN) {
    db.prepare('DELETE FROM login_codes WHERE challenge = ?').run(challenge);
    return res.status(429).json({ error: 'Te veel pogingen. Log opnieuw in.' });
  }

  if (!hashesGelijk(sha256(String(code).trim()), rij.code_hash)) {
    db.prepare('UPDATE login_codes SET attempts = attempts + 1 WHERE challenge = ?').run(challenge);
    const over = CODE_MAX_POGINGEN - (rij.attempts + 1);
    return res.status(401).json({
      error: over > 0 ? 'Verkeerde code. Nog ' + over + ' poging(en).' : 'Verkeerde code.'
    });
  }

  const user = db.prepare('SELECT * FROM users WHERE id = ?').get(rij.user_id);
  if (!user) {
    return res.status(401).json({ error: 'Gebruiker bestaat niet meer' });
  }

  // Code is eenmalig: meteen opruimen.
  db.prepare('DELETE FROM login_codes WHERE challenge = ?').run(challenge);

  const extra = {};
  if (remember) {
    extra.deviceToken = onthoudToestel(user.id, req.headers['user-agent']);
  }

  return res.json(maakLoginAntwoord(user, extra));
});

// Vertrouwde toestellen bekijken en intrekken.
app.get('/api/admin/devices', authMiddleware('admin'), (req, res) => {
  const toestellen = db.prepare(`
    SELECT id, label, created_at, last_used_at
    FROM trusted_devices WHERE user_id = ?
    ORDER BY last_used_at DESC, created_at DESC
  `).all(req.user.id);
  res.json({ ok: true, devices: toestellen });
});

app.delete('/api/admin/devices/:id', authMiddleware('admin'), (req, res) => {
  const info = db.prepare('DELETE FROM trusted_devices WHERE id = ? AND user_id = ?')
    .run(req.params.id, req.user.id);
  if (!info.changes) return res.status(404).json({ error: 'Toestel niet gevonden' });
  res.json({ ok: true });
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
    
    verstuurOfferteMail({ naam, email, telefoon, bedrijf, type, personen, datum, formule, bericht, taal: (req.headers.referer || '') });

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

// Update room (admin only)
app.put('/api/rooms/:id', authMiddleware('admin'), (req, res) => {
  const { name, description, capacity, price_base, available } = req.body;
  const room = db.prepare('SELECT * FROM rooms WHERE id = ?').get(req.params.id);
  if (!room) return res.status(404).json({ error: 'Kamer niet gevonden' });
  
  db.prepare(`UPDATE rooms SET 
    name = COALESCE(?, name),
    description = COALESCE(?, description),
    capacity = COALESCE(?, capacity),
    price_base = COALESCE(?, price_base),
    available = COALESCE(?, available)
    WHERE id = ?`
  ).run(
    name || null, description || null, 
    capacity != null ? capacity : null, 
    price_base != null ? price_base : null,
    available != null ? available : null,
    req.params.id
  );
  
  const updated = db.prepare('SELECT * FROM rooms WHERE id = ?').get(req.params.id);
  res.json({ ok: true, room: updated });
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
const pages = ['feestzaal', 'zakelijk', 'teams', 'feesten', 'over-ons', 'contact', 'offerte', 'privacy', 'algemene-voorwaarden', 'login', 'admin'];

// B&B pagina — prijzen server-side uit database
function serveBBPage(lang) {
  return (req, res) => {
    const dir = lang ? path.join(__dirname, '..', 'public', lang) : path.join(__dirname, '..', 'public');
    const file = path.join(dir, 'bb.html');
    if (!fs.existsSync(file)) {
      return lang ? res.redirect(301, '/bb') : res.status(404).sendFile(path.join(__dirname, '..', 'public', '404.html'));
    }
    let html = fs.readFileSync(file, 'utf-8');
    try {
      const rooms = db.prepare('SELECT slug, price_base FROM rooms').all();
      rooms.forEach(r => {
        const key = '{{PRICE_' + r.slug.toUpperCase().replace(/-/g, '_') + '}}';
        html = html.split(key).join(Math.round(r.price_base).toString());
      });
    } catch (e) {
      console.error('BB price injection error:', e.message);
    }
    res.send(html);
  };
}
app.get('/bb', serveBBPage(null));
app.get('/fr/bb', serveBBPage('fr'));
app.get('/en/bb', serveBBPage('en'));
app.get('/de/bb', serveBBPage('de'));

// llms.txt — plain text voor AI-bots
app.get('/llms.txt', (req, res) => {
  const file = path.join(__dirname, '..', 'public', 'llms.txt');
  if (fs.existsSync(file)) {
    res.type('text/plain').sendFile(file);
  } else {
    res.status(404).type('text/plain').send('Not found');
  }
});

// Blog overzicht — server-side rendering voor SEO
// De bestaande JavaScript voor dynamische filtering en load more blijft werken.
app.get('/blog', (req, res) => {
  const templatePath = path.join(__dirname, '..', 'public', 'blog.html');
  let html = fs.readFileSync(templatePath, 'utf-8');

  try {
    const posts = db.prepare(
      "SELECT id, title, slug, excerpt, category, featured_image, published_at FROM blog_posts WHERE status = 'published' ORDER BY published_at DESC"
    ).all();

    const esc = s => String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

    const cardsHtml = posts.map(post => {
      const date = new Date(post.published_at).toLocaleDateString('nl-BE', { year: 'numeric', month: 'long', day: 'numeric' });
      const img = esc(post.featured_image || '/assets/img/hero.jpg');
      return `<article class="blog-card">
          <a href="/blog/${post.slug}" class="blog-card-link">
            <div class="blog-card-image">
              <img src="${img}" alt="${esc(post.title)}" loading="lazy">
            </div>
            <div class="blog-card-content">
              ${post.category ? `<span class="blog-card-category">${esc(post.category)}</span>` : ''}
              <h3>${esc(post.title)}</h3>
              <p>${esc(post.excerpt || '')}</p>
              <time>${date}</time>
            </div>
          </a>
        </article>`;
    }).join('\n');

    const totalPosts = posts.length;
    const initialCount = Math.min(totalPosts, 9);

    // Vervang de laadindicator door server-side gerenderde posts
    html = html.replace(
      '<div class="blog-loading">\n        <p>Artikelen laden...</p>\n      </div>',
      cardsHtml || '<div class="blog-empty"><p>Nog geen artikelen.</p></div>'
    );

    // Injecteer SSR state zodat de JavaScript weet dat posts al geladen zijn
    const ssrScript = `<script>\n(function(){\nvar g=document.getElementById('blogGrid');\nif(g&&g.querySelector('.blog-card')){\nwindow.__ssrRendered=${initialCount};\nwindow.__ssrTotal=${totalPosts};\n}\n})();\n<\/script>`;
    html = html.replace('</head>', ssrScript + '\n</head>');

  } catch (e) {
    console.error('Blog SSR error:', e);
  }

  res.send(html);
});

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
  let seoBlock = `<title>${title}</title>
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

  // Extract FAQ structured data from content (h3 questions after h2 "Veelgestelde vragen")
  const faqMatch = (post.content || '').split(/veelgestelde\s+vragen/i);
  if (faqMatch.length > 1) {
    const faqSection = faqMatch[1];
    const faqItems = [];
    const qRegex = /<h3[^>]*>([^<]+)<\/h3>\s*<p>([\s\S]*?)<\/p>/gi;
    let m;
    while ((m = qRegex.exec(faqSection)) !== null) {
      faqItems.push({
        '@type': 'Question',
        name: m[1].replace(/<[^>]*>/g, '').trim(),
        acceptedAnswer: {
          '@type': 'Answer',
          text: m[2].replace(/<[^>]*>/g, '').trim()
        }
      });
    }
    if (faqItems.length > 0) {
      const faqLd = JSON.stringify({
        '@context': 'https://schema.org',
        '@type': 'FAQPage',
        mainEntity: faqItems
      });
      seoBlock += `\n<script type="application/ld+json">${faqLd}</script>`;
    }
  }

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

  // Blog taalroutes — geen vertaalde blog.html, dus redirect naar NL blog (SSR)
  app.get(`/${lang}/blog`, serveLang('blog.html', '/blog'));

  ruimtes.forEach(ruimte => {
    app.get(`/${lang}/ruimtes/${ruimte}`, serveLang(path.join('ruimtes', `${ruimte}.html`), `/ruimtes/${ruimte}`));
  });
});

// Catch-all: echte 404 voor onbekende paden
app.get('*', (req, res) => {
  // Skip API routes
  if (req.path.startsWith('/api/')) {
    return res.status(404).json({ error: 'Route niet gevonden' });
  }
  
  res.status(404).sendFile(path.join(__dirname, '..', 'public', '404.html'));
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
