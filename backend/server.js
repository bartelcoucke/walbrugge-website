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

  -- Zichtbare beoordelingsscores per bron. Google wordt automatisch
  -- bijgewerkt; Booking.com en Eventplanner beheert de beheerder zelf,
  -- omdat die partijen geen publieke koppeling aanbieden.
  CREATE TABLE IF NOT EXISTS site_scores (
    bron TEXT PRIMARY KEY,
    score TEXT NOT NULL,
    aantal INTEGER,
    url TEXT,
    automatisch INTEGER DEFAULT 0,
    bijgewerkt DATETIME DEFAULT CURRENT_TIMESTAMP
  );

  -- Openstaande herstelcodes voor een vergeten wachtwoord.
  CREATE TABLE IF NOT EXISTS password_resets (
    challenge TEXT PRIMARY KEY,
    user_id INTEGER NOT NULL,
    code_hash TEXT NOT NULL,
    expires_at INTEGER NOT NULL,
    attempts INTEGER DEFAULT 0,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
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
// Caddy staat ervoor en is de enige die de app rechtstreeks bereikt. Zonder
// deze regel is req.ip altijd 127.0.0.1 en zou de snelheidsbegrenzing alle
// bezoekers als één persoon tellen.
app.set('trust proxy', 1);

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// ═══════════════════════════════════════════════════════════════════════════
// BEZOEKERSSTATISTIEK (serverside, zonder cookies)
// ═══════════════════════════════════════════════════════════════════════════
// Telt paginaweergaven en enkele gebeurtenissen (offerteknop, formulier
// verstuurd, WhatsApp/Messenger/telefoon/e-mail, Boek B&B) op de server.
// Er wordt niets op het toestel van de bezoeker bewaard en er wordt geen
// IP-adres of user-agent opgeslagen. "Unieke bezoekers" per dag komen van een
// hash met een willekeurig dagzout dat alleen in het geheugen leeft: na
// middernacht (of een herstart) is niets meer te herleiden. Werkt dus ook voor
// bezoekers die de cookies weigeren. Zichtbaar in het beheerpaneel (Bezoekers).

db.exec(`
  CREATE TABLE IF NOT EXISTS visits (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ts DATETIME DEFAULT CURRENT_TIMESTAMP,
    dag TEXT NOT NULL,
    pad TEXT NOT NULL,
    taal TEXT,
    bron TEXT,
    utm_source TEXT,
    utm_medium TEXT,
    utm_campaign TEXT,
    toestel TEXT,
    bezoeker TEXT
  );
  CREATE INDEX IF NOT EXISTS idx_visits_dag ON visits(dag);

  CREATE TABLE IF NOT EXISTS events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ts DATETIME DEFAULT CURRENT_TIMESTAMP,
    dag TEXT NOT NULL,
    naam TEXT NOT NULL,
    pad TEXT,
    taal TEXT,
    bron TEXT,
    utm_source TEXT,
    utm_medium TEXT,
    utm_campaign TEXT,
    detail TEXT,
    bezoeker TEXT
  );
  CREATE INDEX IF NOT EXISTS idx_events_dag ON events(dag);
`);

// Bewaartermijn: 400 dagen (een jaar plus wat marge om jaar op jaar te vergelijken).
try {
  db.prepare("DELETE FROM visits WHERE ts < datetime('now', '-400 days')").run();
  db.prepare("DELETE FROM events WHERE ts < datetime('now', '-400 days')").run();
} catch (e) { console.warn('[stats] Opruimen mislukt:', e.message); }

const STATS_EIGEN_HOSTS = new Set(['walbrugge.be', 'www.walbrugge.be', '2.28.71.249', 'localhost', '127.0.0.1']);
const STATS_BOT = /bot|crawl|spider|slurp|preview|monitor|fetch|scan|curl|wget|python|java\/|headless|lighthouse|pingdom|uptime|facebookexternalhit|whatsapp|telegrambot|linkedinbot|semrush|ahrefs|mj12|dotbot|petalbot|bytespider|gptbot|claudebot|ccbot/i;
const STATS_EVENTS = new Set(['offerte_click', 'generate_lead', 'whatsapp_click', 'messenger_click', 'phone_click', 'email_click', 'booking_click', 'zaal_click', 'carrousel', 'award_click']);

function statsDag() {
  return new Date().toLocaleDateString('sv-SE', { timeZone: 'Europe/Brussels' });
}

// Dagzout: willekeurig, alleen in het geheugen, wisselt elke dag.
let statsZout = { dag: '', waarde: '' };
function statsBezoeker(req) {
  const dag = statsDag();
  if (statsZout.dag !== dag) statsZout = { dag, waarde: crypto.randomBytes(16).toString('hex') };
  const ip = String(req.headers['x-forwarded-for'] || '').split(',')[0].trim() || req.ip || '';
  const ua = String(req.headers['user-agent'] || '');
  return crypto.createHash('sha256').update(statsZout.waarde + '|' + ip + '|' + ua).digest('hex').slice(0, 16);
}

function statsIsBot(req) {
  const ua = String(req.headers['user-agent'] || '');
  return !ua || STATS_BOT.test(ua);
}

function statsToestel(req) {
  return /Mobi|Android|iPhone|iPad/i.test(String(req.headers['user-agent'] || '')) ? 'mobiel' : 'desktop';
}

// Android geeft bij een klik vanuit een app een verwijzer "android-app://<pakket>" mee.
const STATS_APPS = {
  'com.whatsapp': 'WhatsApp (app)',
  'com.whatsapp.w4b': 'WhatsApp Business (app)',
  'com.google.android.gm': 'Gmail (app)',
  'com.microsoft.office.outlook': 'Outlook (app)',
  'com.samsung.android.email.provider': 'Samsung Mail (app)',
  'com.google.android.apps.messaging': 'Berichten (app)',
  'com.facebook.katana': 'Facebook (app)',
  'com.facebook.lite': 'Facebook (app)',
  'm.facebook.com': 'Facebook (app)',
  'com.facebook.orca': 'Messenger (app)',
  'com.instagram.android': 'Instagram (app)',
  'com.linkedin.android': 'LinkedIn (app)',
  'org.telegram.messenger': 'Telegram (app)',
  'com.google.android.googlequicksearchbox': 'Google (app)',
  'com.pinterest': 'Pinterest (app)',
  'com.twitter.android': 'X (app)',
  'com.zhiliaoapp.musically': 'TikTok (app)',
  'com.snapchat.android': 'Snapchat (app)'
};

// In-app-browsers (ook op iPhone) zijn te herkennen aan de user-agent.
function statsInAppBrowser(ua) {
  const s = String(ua || '');
  if (!s) return null;
  if (/MessengerForiOS|FB_IAB\/MESSENGER|\bMessenger\b/i.test(s)) return 'Messenger (app)';
  if (/FBAN|FBAV|FB_IAB|FBIOS/i.test(s)) return 'Facebook (app)';
  if (/Instagram/i.test(s)) return 'Instagram (app)';
  if (/LinkedInApp/i.test(s)) return 'LinkedIn (app)';
  if (/musical_ly|TikTok|BytedanceWebview/i.test(s)) return 'TikTok (app)';
  if (/Pinterest/i.test(s)) return 'Pinterest (app)';
  if (/Snapchat/i.test(s)) return 'Snapchat (app)';
  if (/Twitter for iPhone|Twitter for Android/i.test(s)) return 'X (app)';
  if (/\bGSA\/|GoogleApp/i.test(s)) return 'Google (app)';
  return null;
}

// Externe herkomst: het domein van de verwijzende site, een app-naam (Android
// app-verwijzer of in-app-browser), of null als de bezoeker van de site zelf
// komt of er niets te herkennen valt.
function statsBron(ref, eigenHost, ua) {
  const r = String(ref || '').trim();
  const app = /^android-app:\/\/([^/?#]+)/i.exec(r);
  if (app) {
    const pakket = app[1].toLowerCase();
    return STATS_APPS[pakket] || ('app: ' + pakket.slice(0, 60));
  }
  if (r) {
    try {
      const host = new URL(r).hostname.toLowerCase();
      const eigen = !host || STATS_EIGEN_HOSTS.has(host) || host === String(eigenHost || '').toLowerCase().split(':')[0];
      if (!eigen) return host.slice(0, 120);
      if (eigen && host) return null; // van de site zelf: geen nieuwe herkomst
    } catch (e) { /* onbruikbare verwijzer, val terug op de user-agent */ }
  }
  return statsInAppBrowser(ua);
}

function statsTaal(pad) {
  const m = /^\/(fr|en|de)(\/|$)/.exec(pad || '');
  return m ? m[1] : 'nl';
}

function statsKort(v, n) {
  return v == null || v === '' ? null : String(v).slice(0, n || 80);
}

const statsInsertVisit = db.prepare(`
  INSERT INTO visits (dag, pad, taal, bron, utm_source, utm_medium, utm_campaign, toestel, bezoeker)
  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`);
const statsInsertEvent = db.prepare(`
  INSERT INTO events (dag, naam, pad, taal, bron, utm_source, utm_medium, utm_campaign, detail, bezoeker)
  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`);

// Paginaweergaven: alles wat als HTML met status 200 vertrekt, behalve het
// beheer en de inlog-/gastenpagina's.
app.use((req, res, next) => {
  if (req.method !== 'GET') return next();
  res.on('finish', () => {
    try {
      const pad = req.path;

      // Oude WordPress-URL's (301 uit OUDE_URLS): apart geteld als gebeurtenis,
      // zodat zichtbaar blijft hoeveel bezoekers nog via de oude site binnenkomen.
      // De nieuwe pagina zelf telt daarna gewoon als weergave (de browser volgt de 301).
      if (res.statusCode === 301 && req.route && typeof OUDE_URLS === 'object'
          && Object.prototype.hasOwnProperty.call(OUDE_URLS, req.route.path)) {
        if (statsIsBot(req)) return;
        statsInsertEvent.run(
          statsDag(), 'oude_url', statsKort(pad, 200), 'nl', statsBron(req.get('referer'), req.get('host'), req.headers['user-agent']),
          statsKort(req.query.utm_source), statsKort(req.query.utm_medium), statsKort(req.query.utm_campaign),
          statsKort(OUDE_URLS[req.route.path], 120), statsBezoeker(req)
        );
        return;
      }

      if (res.statusCode !== 200) return;
      if (!String(res.getHeader('content-type') || '').includes('text/html')) return;
      if (/^\/(api|admin|gasten|login)(\/|$)/.test(pad) || /^\/(fr|en|de)\/login$/.test(pad)) return;
      if (/^\/google[0-9a-f]+\.html$/.test(pad)) return; // Search Console-verificatie
      if (statsIsBot(req)) return;
      statsInsertVisit.run(
        statsDag(), statsKort(pad, 200), statsTaal(pad), statsBron(req.get('referer'), req.get('host'), req.headers['user-agent']),
        statsKort(req.query.utm_source), statsKort(req.query.utm_medium), statsKort(req.query.utm_campaign),
        statsToestel(req), statsBezoeker(req)
      );
    } catch (e) { /* statistiek mag nooit een pagina breken */ }
  });
  next();
});

// Gebeurtenissen vanuit de pagina's (navigator.sendBeacon in app.js).
// Het pad heet bewust niet "track": adblockers filteren zulke namen.
app.post('/api/telling', (req, res) => {
  res.status(204).end();
  try {
    if (statsIsBot(req)) return;
    const b = req.body || {};
    const naam = String(b.naam || '');
    if (!STATS_EVENTS.has(naam)) return;
    const pad = statsKort(b.pad, 200) || '/';
    statsInsertEvent.run(
      statsDag(), naam, pad, statsTaal(pad),
      statsBron(b.ref, req.get('host'), req.headers['user-agent']),
      statsKort(b.utm_source), statsKort(b.utm_medium), statsKort(b.utm_campaign),
      statsKort(b.detail, 120), statsBezoeker(req)
    );
  } catch (e) { /* stil */ }
});

// Overzicht voor het beheerpaneel.
app.get('/api/admin/bezoekers', authMiddleware('admin'), (req, res) => {
  const dagen = Math.min(Math.max(parseInt(req.query.dagen, 10) || 30, 1), 400);
  const van = new Date(Date.now() - (dagen - 1) * 864e5).toLocaleDateString('sv-SE', { timeZone: 'Europe/Brussels' });
  const tot = statsDag();
  const w = 'WHERE dag >= ? AND dag <= ?';
  const all = (sql, ...p) => db.prepare(sql).all(...p);
  const get = (sql, ...p) => db.prepare(sql).get(...p);

  const totaal = get(`SELECT COUNT(*) AS weergaven, COUNT(DISTINCT bezoeker || dag) AS bezoekers FROM visits ${w}`, van, tot);
  const events = all(`SELECT naam, COUNT(*) AS aantal FROM events ${w} GROUP BY naam ORDER BY aantal DESC`, van, tot);
  const perDag = all(`
    SELECT d.dag,
      (SELECT COUNT(*) FROM visits v WHERE v.dag = d.dag) AS weergaven,
      (SELECT COUNT(DISTINCT bezoeker) FROM visits v WHERE v.dag = d.dag) AS bezoekers,
      (SELECT COUNT(*) FROM events e WHERE e.dag = d.dag AND e.naam = 'generate_lead') AS leads
    FROM (SELECT DISTINCT dag FROM visits ${w} UNION SELECT DISTINCT dag FROM events ${w}) d
    ORDER BY d.dag DESC`, van, tot, van, tot);
  const bronnen = all(`
    SELECT COALESCE(v.bron, '') AS bron, COUNT(*) AS weergaven,
      (SELECT COUNT(*) FROM events e WHERE e.naam = 'generate_lead' AND COALESCE(e.bron, '') = COALESCE(v.bron, '') AND e.dag >= ? AND e.dag <= ?) AS leads,
      (SELECT COUNT(*) FROM events e WHERE e.naam = 'booking_click' AND COALESCE(e.bron, '') = COALESCE(v.bron, '') AND e.dag >= ? AND e.dag <= ?) AS boekingen
    FROM visits v ${w} GROUP BY COALESCE(v.bron, '') ORDER BY weergaven DESC LIMIT 20`, van, tot, van, tot, van, tot);
  const paginas = all(`SELECT pad, COUNT(*) AS aantal FROM visits ${w} GROUP BY pad ORDER BY aantal DESC LIMIT 20`, van, tot);
  const campagnes = all(`
    SELECT utm_source, utm_medium, utm_campaign, COUNT(*) AS weergaven,
      (SELECT COUNT(*) FROM events e WHERE e.naam = 'generate_lead' AND COALESCE(e.utm_campaign,'') = COALESCE(v.utm_campaign,'') AND COALESCE(e.utm_source,'') = COALESCE(v.utm_source,'') AND e.dag >= ? AND e.dag <= ?) AS leads
    FROM visits v ${w} AND (utm_source IS NOT NULL OR utm_campaign IS NOT NULL)
    GROUP BY utm_source, utm_medium, utm_campaign ORDER BY weergaven DESC LIMIT 20`, van, tot, van, tot);
  const talen = all(`SELECT taal, COUNT(*) AS aantal FROM visits ${w} GROUP BY taal ORDER BY aantal DESC`, van, tot);
  const toestellen = all(`SELECT toestel, COUNT(*) AS aantal FROM visits ${w} GROUP BY toestel ORDER BY aantal DESC`, van, tot);
  const eventsDetail = all(`
    SELECT naam, COALESCE(detail, '') AS detail, COALESCE(pad, '') AS pad, COUNT(*) AS aantal
    FROM events ${w} GROUP BY naam, detail, pad ORDER BY aantal DESC LIMIT 200`, van, tot);

  // Zalen: klikken op een zaaltegel of -link (event zaal_click, detail = slug),
  // weergaven van de zaalpagina (/ruimtes/<slug>, in elke taal) en klikken in
  // de fotocarrousel van die zaal (event carrousel, detail = 'zaal:<slug>').
  const zalenMap = new Map();
  const zaalRij = z => { if (!zalenMap.has(z)) zalenMap.set(z, { zaal: z, klikken: 0, weergaven: 0, carrousel: 0 }); return zalenMap.get(z); };
  all(`SELECT detail AS zaal, COUNT(*) AS n FROM events ${w} AND naam = 'zaal_click' AND detail != '' GROUP BY detail`, van, tot)
    .forEach(r => { zaalRij(r.zaal).klikken = r.n; });
  all(`SELECT rtrim(substr(pad, instr(pad, '/ruimtes/') + 9), '/') AS zaal, COUNT(*) AS n
       FROM visits ${w} AND instr(pad, '/ruimtes/') > 0 GROUP BY zaal`, van, tot)
    .forEach(r => { if (r.zaal) zaalRij(r.zaal).weergaven = r.n; });
  all(`SELECT substr(detail, 6) AS zaal, COUNT(*) AS n FROM events ${w} AND naam = 'carrousel' AND detail LIKE 'zaal:%' GROUP BY detail`, van, tot)
    .forEach(r => { zaalRij(r.zaal).carrousel = r.n; });
  const zalen = [...zalenMap.values()].sort((a, b) => b.klikken - a.klikken || b.weergaven - a.weergaven);

  // Fotocarrousels: klikken op pijltjes/bolletjes per carrousel, plus hoeveel
  // bezoekers dat deden, afgezet tegen alle bezoekers van de pagina waar de
  // carrousel staat (zaal → /teams, kamer → /bb, feest → /feesten).
  const CARROUSEL_PAGINA = { zaal: 'teams', kamer: 'bb', feest: 'feesten' };
  const paginaStats = {};
  for (const [type, pagina] of Object.entries(CARROUSEL_PAGINA)) {
    paginaStats[type] = get(`SELECT COUNT(*) AS weergaven, COUNT(DISTINCT bezoeker || dag) AS bezoekers
                             FROM visits ${w} AND (pad = ? OR pad LIKE ?)`, van, tot, '/' + pagina, '/__/' + pagina);
  }
  const carrousels = all(`SELECT detail, COUNT(*) AS klikken, COUNT(DISTINCT bezoeker || dag) AS bezoekers
                          FROM events ${w} AND naam = 'carrousel' GROUP BY detail ORDER BY klikken DESC LIMIT 40`, van, tot)
    .map(c => {
      const type = String(c.detail).split(':')[0];
      const p = paginaStats[type] || {};
      return { detail: c.detail, type, pagina: CARROUSEL_PAGINA[type] ? '/' + CARROUSEL_PAGINA[type] : '',
               klikken: c.klikken, bezoekers: c.bezoekers, paginaWeergaven: p.weergaven || 0, paginaBezoekers: p.bezoekers || 0 };
    });

  // Awards en reviewlinks: klikken op de Salino-, Booking.com-, Eventplanner-
  // en Google-links (event award_click, detail = '<platform>:<soort>', met soort
  // 'tegel' voor de award-kaarten en 'knop' voor al de rest). Per pagina, de
  // vier talen samengeteld; meerdere gelijke knoppen op één pagina tellen samen.
  const PAGINA_ZONDER_TAAL = "CASE WHEN pad IN ('/fr', '/en', '/de', '/fr/', '/en/', '/de/') THEN '/' WHEN pad LIKE '/fr/%' OR pad LIKE '/en/%' OR pad LIKE '/de/%' THEN substr(pad, 4) ELSE pad END";
  const awards = all(`SELECT ${PAGINA_ZONDER_TAAL} AS pagina, detail, COUNT(*) AS klikken, COUNT(DISTINCT bezoeker || dag) AS bezoekers
                      FROM events ${w} AND naam = 'award_click' GROUP BY pagina, detail ORDER BY klikken DESC LIMIT 200`, van, tot);
  const awardWeergaven = all(`SELECT ${PAGINA_ZONDER_TAAL} AS pagina, COUNT(*) AS weergaven FROM visits ${w} GROUP BY pagina`, van, tot);

  res.json({ ok: true, periode: { van, tot, dagen }, totaal, events, perDag, bronnen, paginas, campagnes, talen, toestellen, eventsDetail, zalen, carrousels, awards, awardWeergaven });
});


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
// ── Homepage met echte Google-beoordelingen ──────────────────────────────
// De reviews worden server-side meegerenderd, zodat zoekmachines en
// AI-crawlers ze zien, en het schema alleen claimt wat op de pagina staat.

const REVIEW_TEKST = {
  nl: { kop: 'Wat gasten schrijven', sub: 'Beoordelingen op Google',
        van: 'op Google', alles: 'Alle beoordelingen op Google' },
  fr: { kop: 'Ce que disent nos hôtes', sub: 'Avis sur Google',
        van: 'sur Google', alles: 'Tous les avis sur Google' },
  en: { kop: 'What guests write', sub: 'Reviews on Google',
        van: 'on Google', alles: 'All reviews on Google' },
  de: { kop: 'Was Gäste schreiben', sub: 'Bewertungen auf Google',
        van: 'auf Google', alles: 'Alle Bewertungen auf Google' }
};

function escHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function bouwReviewBlok(cache, lang) {
  const t = REVIEW_TEKST[lang] || REVIEW_TEKST.nl;
  const sterren = n => '★'.repeat(Math.round(n)) + '☆'.repeat(5 - Math.round(n));

  const kaarten = cache.reviews.map(r =>
    '<figure class="gr-card">' +
      '<div class="gr-stars" aria-label="' + r.score + ' van 5">' + sterren(r.score) + '</div>' +
      '<blockquote class="gr-text">' + escHtml(r.tekst) + '</blockquote>' +
      '<figcaption class="gr-meta">— <strong>' + escHtml(r.auteur) + '</strong>' +
        (r.wanneer ? ' · ' + escHtml(r.wanneer) : '') + ' · ' + t.van +
      '</figcaption>' +
    '</figure>'
  ).join('');

  // Volgorde: score-regel ("Google ★★★★★ · 5,0 / 5 · 40 beoordelingen"),
  // dan de link naar alle beoordelingen, dan pas de reviewkaarten.
  return (
    '<section class="section" id="google-reviews">' +
      '<div class="container">' +
        '<p class="kicker center">' + t.sub + '</p>' +
        '<h2 class="center">' + t.kop + '</h2>' +
        '<p class="center gr-score">' +
          'Google <span class="gr-stars gr-stars-inline" aria-label="' + cache.score + ' / 5">' +
            sterren(cache.score) + '</span> · ' +
          '<strong>' + String(cache.score).replace('.', ',') + '</strong> / 5 · ' +
          cache.aantal + ' ' + (lang === 'fr' ? 'avis' : lang === 'de' ? 'Bewertungen'
                               : lang === 'en' ? 'reviews' : 'beoordelingen') +
        '</p>' +
        '<p class="center gr-alle"><a href="' + escHtml(cache.kaartUrl) +
          '" target="_blank" rel="noopener">' + t.alles + ' ↗</a></p>' +
        '<div class="gr-grid">' + kaarten + '</div>' +
      '</div>' +
    '</section>'
  );
}

function bouwReviewSchema(cache) {
  return JSON.stringify({
    '@context': 'https://schema.org',
    '@type': 'LocalBusiness',
    '@id': 'https://walbrugge.be/#beoordelingen',
    name: 'Domein Walbrugge',
    url: 'https://walbrugge.be/',
    aggregateRating: {
      '@type': 'AggregateRating',
      ratingValue: String(cache.score),
      bestRating: '5',
      worstRating: '1',
      ratingCount: String(cache.aantal)
    },
    review: cache.reviews.map(r => ({
      '@type': 'Review',
      author: { '@type': 'Person', name: r.auteur },
      reviewRating: { '@type': 'Rating', ratingValue: String(r.score), bestRating: '5' },
      reviewBody: r.tekst,
      datePublished: r.datum || undefined,
      publisher: { '@type': 'Organization', name: 'Google' }
    }))
  });
}

function serveerHomepage(lang) {
  return (req, res, next) => {
    const map = lang
      ? path.join(__dirname, '..', 'public', lang, 'index.html')
      : path.join(__dirname, '..', 'public', 'index.html');
    if (!fs.existsSync(map)) return next();

    let html = fs.readFileSync(map, 'utf-8');

    // Sterren meteen op de juiste breedte, ongeacht of er reviews opgehaald
    // zijn: anders lichten er even vijf volle sterren op.
    const gScore = parseFloat(String((leesScores().google || {}).score || '').replace(',', '.'));
    if (!isNaN(gScore)) {
      const pct = Math.max(0, Math.min(100, (gScore > 5 ? gScore / 2 : gScore) / 5 * 100)).toFixed(1);
      html = html.replace(/(<span class="g-stars-fill" style="width:)100%(")/g, '$1' + pct + '%$2');
    }

    // Reviews in de taal van de pagina; is die cache er (nog) niet, dan de
    // Nederlandse, zodat de sectie niet wegvalt.
    const cache = leesReviewCache(lang || 'nl') || leesReviewCache('nl');

    // Zonder cache blijft de pagina exact zoals ze is.
    if (cache && cache.score && cache.reviews && cache.reviews.length) {
      const blok = bouwReviewBlok(cache, lang || 'nl');
      html = html.replace('<section class="section section-cta">',
                          blok + '\n<section class="section section-cta">');
      html = html.replace('</head>',
        '<script type="application/ld+json">' + bouwReviewSchema(cache) + '</script>\n</head>');
    }
    res.send(html);
  };
}

app.get('/', serveerHomepage(null));
app.get('/fr', serveerHomepage('fr'));
app.get('/en', serveerHomepage('en'));
app.get('/de', serveerHomepage('de'));

// ── Oude WordPress-URL's ────────────────────────────────────────────────
// Google heeft de adressen van de vorige site nog in zijn index. Elk daarvan
// krijgt hier één 301 naar zijn nieuwe tegenhanger — vóór de slash-middleware,
// zodat /bedrijven/ en /bedrijven allebei in één sprong aankomen in plaats
// van via twee omleidingen. Express matcht beide vormen (geen strict routing).
// De query string blijft behouden: advertentieklikken dragen een gclid.
const OUDE_URLS = {
  '/bedrijven':       '/teams',
  '/zakelijk':        '/teams',
  '/coworking':       '/teams',
  '/vergaderzaal':    '/teams',
  '/vergaderzalen':   '/teams',
  '/teambuilding':    '/teams',
  '/feestzaal':       '/feesten',
  '/privefeesten':    '/feesten',
  '/trouwfeest':      '/feesten#trouwfeest',
  '/communiefeest':   '/feesten#familiefeest',
  '/bed-breakfast':   '/bb',
  '/kamer':           '/bb',
  '/kamer/*':         '/bb',
  '/key-selling-point/*': '/bb',
  '/faciliteiten':    '/over-ons',
  '/huisreglement':   '/algemene-voorwaarden'
};
Object.entries(OUDE_URLS).forEach(([oud, nieuw]) => {
  app.get(oud, (req, res) => {
    const q = req.url.indexOf('?') !== -1 ? req.url.slice(req.url.indexOf('?')) : '';
    const [pad, hash] = nieuw.split('#');
    res.redirect(301, pad + q + (hash ? '#' + hash : ''));
  });
});

// Eén URL per pagina. /teams/ en /index.html gaven eerder gewoon 200 met
// dezelfde inhoud als de canonieke URL; dat is duplicate content. Deze
// middleware moet vóór express.static staan.
app.use((req, res, next) => {
  if (req.method !== 'GET' && req.method !== 'HEAD') return next();

  const [pad, ...rest] = req.originalUrl.split('?');
  const query = rest.length ? '?' + rest.join('?') : '';

  if (pad.endsWith('/index.html')) {
    return res.redirect(301, (pad.slice(0, -'/index.html'.length) || '/') + query);
  }
  if (pad.length > 1 && pad.endsWith('/')) {
    return res.redirect(301, pad.slice(0, -1) + query);
  }
  next();
});

// Caddy zet de caching-headers. express.static mag er zelf geen sturen:
// removeHeader in setHeaders werkte niet, omdat express de header pas ná die
// callback zet — en dan juist wél, omdat er dan geen header meer staat.
// Met cacheControl: false zwijgt express en blijft er één header over.
// redirect: false — anders vangt express.static de map /fr en stuurt hij naar
// /fr/, terwijl de sitemap en de canonical /fr gebruiken.
app.use(express.static(path.join(__dirname, '..', 'public'), {
  cacheControl: false,
  redirect: false,
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

// Tweestapsverificatie staat aan tenzij ADMIN_2FA op 'uit' / 'off' / '0' staat.
// De herstelweg voor een vergeten wachtwoord blijft altijd werken.
const TWEESTAPS_AAN = !['uit', 'off', 'false', '0', 'nee'].includes(
  String(process.env.ADMIN_2FA || 'aan').trim().toLowerCase());

console.log('[auth] Tweestapsverificatie beheerder: ' + (TWEESTAPS_AAN ? 'AAN' : 'UIT'));

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

async function stuurHerstelcode(naarAdres, code) {
  const onderwerp = 'Herstelcode Walbrugge: ' + code;
  const tekst = 'Uw herstelcode voor het beheer van walbrugge.be is ' + code + '.\n' +
                'De code blijft 10 minuten geldig.\n\n' +
                'Hebt u zelf geen nieuw wachtwoord aangevraagd, negeer deze mail dan: ' +
                'zonder de code verandert er niets.';
  const html =
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222;">' +
    '<h2 style="margin:0 0 4px;font-size:18px;">Nieuw wachtwoord instellen</h2>' +
    '<p style="margin:0 0 16px;color:#777;">walbrugge.be</p>' +
    '<p style="font-size:30px;letter-spacing:6px;font-weight:bold;margin:0 0 12px;">' + code + '</p>' +
    '<p style="margin:0 0 16px;">De code blijft 10 minuten geldig.</p>' +
    '<p style="color:#777;font-size:12px;">Hebt u zelf geen nieuw wachtwoord aangevraagd, ' +
    'negeer deze mail dan — zonder de code verandert er niets.</p></div>';

  return verstuurEenvoudig(naarAdres, onderwerp, tekst, html);
}

async function verstuurEenvoudig(naarAdres, onderwerp, tekst, html) {
  if (graphActief) {
    await graphSendMail({ subject: onderwerp, text: tekst, html, to: [naarAdres] });
    return 'Graph';
  }
  if (mailer) {
    await mailer.sendMail({
      from: '"Domein Walbrugge" <' + MAIL_FROM + '>',
      to: naarAdres, subject: onderwerp, text: tekst, html
    });
    return 'SMTP';
  }
  throw new Error('geen mailweg ingesteld');
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

  return verstuurEenvoudig(naarAdres, onderwerp, tekst, html);
}

// ═══════════════════════════════════════════════════════════════════════════
// SNELHEIDSBEGRENZING
// ═══════════════════════════════════════════════════════════════════════════
// Zonder externe pakketten: een venster per IP in het geheugen. Genoeg om
// een formulier of een inlogpagina te beschermen tegen geautomatiseerd
// afvuren; bij een herstart begint de teller opnieuw, wat hier volstaat.

function begrensSnelheid({ max, vensterMs, boodschap }) {
  const pogingen = new Map();

  // Voorkomt dat de map onbeperkt groeit.
  setInterval(() => {
    const nu = Date.now();
    for (const [sleutel, rij] of pogingen) {
      if (nu - rij.start > vensterMs) pogingen.delete(sleutel);
    }
  }, vensterMs).unref();

  return (req, res, next) => {
    const ip = req.ip || req.socket.remoteAddress || 'onbekend';
    const nu = Date.now();
    const rij = pogingen.get(ip);

    // Enkel mislukte pogingen tellen. Wie er ondertussen wél in geraakt,
    // begint weer met een schone lei: res.wisBegrenzer() doet dat.
    res.wisBegrenzer = () => pogingen.delete(ip);

    if (rij && nu - rij.start <= vensterMs && rij.aantal >= max) {
      const overSec = Math.ceil((vensterMs - (nu - rij.start)) / 1000);
      res.set('Retry-After', String(overSec));
      return res.status(429).json({ error: boodschap });
    }

    res.on('finish', () => {
      if (res.statusCode < 400) return;           // gelukt: niet meetellen
      const r = pogingen.get(ip);
      if (!r || Date.now() - r.start > vensterMs) pogingen.set(ip, { start: Date.now(), aantal: 1 });
      else r.aantal++;
    });
    next();
  };
}

const contactBegrenzer = begrensSnelheid({
  max: 5, vensterMs: 15 * 60 * 1000,
  boodschap: 'Te veel aanvragen na elkaar. Probeer het over een kwartier opnieuw.'
});
const loginBegrenzer = begrensSnelheid({
  max: 20, vensterMs: 15 * 60 * 1000,
  boodschap: 'Te veel mislukte inlogpogingen. Probeer het over een kwartier opnieuw, of gebruik "Wachtwoord vergeten?".'
});
// Aparte teller: anders sluit een reeks foute wachtwoorden ook de enige
// weg af om een nieuw wachtwoord aan te vragen.
const herstelBegrenzer = begrensSnelheid({
  max: 5, vensterMs: 15 * 60 * 1000,
  boodschap: 'Te veel herstelaanvragen. Probeer het over een kwartier opnieuw.'
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
app.post('/api/login', loginBegrenzer, (req, res) => {
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
    // Uit te zetten met ADMIN_2FA=uit in /etc/walbrugge.env; de code blijft
    // staan, zodat aanzetten later één regel is.
    if (user.role === 'admin' && TWEESTAPS_AAN) {
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
app.post('/api/login/2fa', loginBegrenzer, (req, res) => {
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

// ── Wachtwoord vergeten ───────────────────────────────────────────────────
// Stap 1: code aanvragen. Het antwoord is altijd hetzelfde, ook als het adres
// niet bestaat — anders verklapt de route welke adressen beheerder zijn.
app.post('/api/login/forgot', herstelBegrenzer, (req, res) => {
  const email = String((req.body && req.body.email) || '').trim().toLowerCase();
  const altijd = {
    ok: true,
    message: 'Bestaat er een beheerder met dit adres, dan is er een herstelcode verstuurd.'
  };
  if (!email) return res.status(400).json({ error: 'Vul uw e-mailadres in' });

  const user = db.prepare('SELECT * FROM users WHERE lower(email) = ? AND role = ?')
    .get(email, 'admin');
  if (!user) return res.json(altijd);

  db.prepare('DELETE FROM password_resets WHERE expires_at < ?').run(Date.now());

  const code = String(crypto.randomInt(0, 1000000)).padStart(6, '0');
  const challenge = crypto.randomBytes(24).toString('hex');
  db.prepare(`INSERT INTO password_resets (challenge, user_id, code_hash, expires_at)
              VALUES (?, ?, ?, ?)`)
    .run(challenge, user.id, sha256(code), Date.now() + CODE_GELDIG_MS);

  const naarAdres = process.env.ADMIN_2FA_EMAIL || user.email;
  stuurHerstelcode(naarAdres, code)
    .then(via => console.log('[herstel] Herstelcode verstuurd naar ' + naarAdres + ' via ' + via))
    .catch(err => {
      console.error('[herstel] Versturen mislukt:', err.message);
      console.error('[herstel] Noodcode voor ' + user.email + ': ' + code);
    });

  res.json(Object.assign({ challenge, hint: maskeerAdres(naarAdres) }, altijd));
});

// Stap 2: code plus nieuw wachtwoord.
app.post('/api/login/reset', herstelBegrenzer, (req, res) => {
  const { challenge, code, password } = req.body || {};
  if (!challenge || !code || !password) {
    return res.status(400).json({ error: 'Vul de code en een nieuw wachtwoord in' });
  }
  if (String(password).length < 12) {
    return res.status(400).json({ error: 'Het wachtwoord moet minstens 12 tekens lang zijn' });
  }

  db.prepare('DELETE FROM password_resets WHERE expires_at < ?').run(Date.now());
  const rij = db.prepare('SELECT * FROM password_resets WHERE challenge = ?').get(challenge);

  if (!rij || rij.expires_at < Date.now()) {
    return res.status(401).json({ error: 'De code is verlopen. Vraag een nieuwe aan.' });
  }
  if (rij.attempts >= CODE_MAX_POGINGEN) {
    db.prepare('DELETE FROM password_resets WHERE challenge = ?').run(challenge);
    return res.status(429).json({ error: 'Te veel pogingen. Vraag een nieuwe code aan.' });
  }
  if (!hashesGelijk(sha256(String(code).trim()), rij.code_hash)) {
    db.prepare('UPDATE password_resets SET attempts = attempts + 1 WHERE challenge = ?').run(challenge);
    const over = CODE_MAX_POGINGEN - (rij.attempts + 1);
    return res.status(401).json({
      error: over > 0 ? 'Verkeerde code. Nog ' + over + ' poging(en).' : 'Verkeerde code.'
    });
  }

  const user = db.prepare('SELECT * FROM users WHERE id = ?').get(rij.user_id);
  if (!user) return res.status(401).json({ error: 'Gebruiker bestaat niet meer' });

  db.prepare('UPDATE users SET password_hash = ? WHERE id = ?')
    .run(bcrypt.hashSync(String(password), 12), user.id);

  // Een herstel maakt alle bestaande sessies en toestellen ongeldig: wie het
  // wachtwoord kwijt was, wil niet dat een oud toestel toegang houdt.
  db.prepare('DELETE FROM password_resets WHERE challenge = ?').run(challenge);
  db.prepare('DELETE FROM trusted_devices WHERE user_id = ?').run(user.id);
  db.prepare('DELETE FROM login_codes WHERE user_id = ?').run(user.id);
  console.log('[herstel] Wachtwoord opnieuw ingesteld voor ' + user.email +
              ' — vertrouwde toestellen ingetrokken');

  res.json({ ok: true, message: 'Wachtwoord gewijzigd. U kunt nu inloggen.' });
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

// ═══════════════════════════════════════════════════════════════════════════
// TECHNISCH LOGBOEK
// ═══════════════════════════════════════════════════════════════════════════
// Toont het journal van de dienst in het beheerpaneel, zodat een herstart of
// een mislukte mail zichtbaar is zonder SSH.

const LOG_NIVEAUS = {
  fout: /error|failed|mislukt|exception|fatal|\[2fa\] Versturen/i,
  waarschuwing: /warn|waarschuwing|niet ingesteld/i
};

// Inlogcodes horen niet in een webpagina thuis, ook niet voor een beheerder.
function verbergCodes(regel) {
  return regel.replace(/(Noodcode voor [^:]+: )\d{6}/gi, '$1••••••');
}

function duidNiveau(regel) {
  if (LOG_NIVEAUS.fout.test(regel)) return 'fout';
  if (LOG_NIVEAUS.waarschuwing.test(regel)) return 'waarschuwing';
  return 'info';
}

app.get('/api/admin/logs', authMiddleware('admin'), (req, res) => {
  const aantal = Math.min(Math.max(parseInt(req.query.lines, 10) || 200, 10), 1000);
  const { execFile } = require('child_process');

  execFile('journalctl', ['-u', 'walbrugge', '-n', String(aantal), '--no-pager', '-o', 'short-iso'],
    { timeout: 10000, maxBuffer: 4 * 1024 * 1024 },
    (err, stdout) => {
      if (err && !stdout) {
        return res.status(500).json({
          error: 'Logboek niet leesbaar op deze server: ' + err.message
        });
      }

      const regels = String(stdout).split('\n')
        .filter(r => r.trim() && !/^-- (Logs|No entries)/.test(r))
        .map(r => {
          const m = r.match(/^(\S+)\s+\S+\s+\S+?:\s?(.*)$/);
          const tekst = verbergCodes(m ? m[2] : r);
          return { tijd: m ? m[1] : '', tekst, niveau: duidNiveau(tekst) };
        });

      res.json({ ok: true, lines: regels });
    });
});

// Token verification
app.get('/api/me', authMiddleware(), (req, res) => {
  res.json({ ok: true, user: req.user });
});

// ═══════════════════════════════════════════════════════════════════════════
// CONTACT / OFFERTE ROUTES
// ═══════════════════════════════════════════════════════════════════════════

app.post('/api/contact', contactBegrenzer, (req, res) => {
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

// De omleidingen van oude URL's staan in OUDE_URLS, hogerop, vóór de
// slash-middleware — zodat ook /bedrijven/ in één sprong op /teams komt.

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

// ═══════════════════════════════════════════════════════════════════════════
// GOOGLE-BEOORDELINGEN
// ═══════════════════════════════════════════════════════════════════════════
// Haalt de echte score en de laatste reviews op bij de Places API, één keer
// per dag. Zo staat er nooit een verouderd cijfer op de site en verzinnen we
// niets: elke review draagt de naam van de schrijver en verwijst naar Google.
//
// Nodig in /etc/walbrugge.env:
//   GOOGLE_PLACES_API_KEY=...
//   GOOGLE_PLACE_ID=...        (optioneel; wordt anders opgezocht op naam)

const PLACES_KEY = process.env.GOOGLE_PLACES_API_KEY;
const PLACES_MAX_LEEFTIJD = 24 * 60 * 60 * 1000;

// De reviews worden per taal opgehaald (Google vertaalt ze mee), dus ook per
// taal bewaard. Het Nederlandse bestand houdt zijn oude naam.
const REVIEW_TALEN = ['nl', 'fr', 'en', 'de'];
function placesCachePad(lang) {
  return path.join(DATA_DIR, lang && lang !== 'nl' ? 'google-reviews-' + lang + '.json'
                                                    : 'google-reviews.json');
}

function leesReviewCache(lang) {
  try {
    return JSON.parse(fs.readFileSync(placesCachePad(lang || 'nl'), 'utf-8'));
  } catch (e) {
    return null;
  }
}

// De API-sleutel is bij Google beperkt tot het IPv4-adres van deze server.
// De server heeft ook IPv6, en fetch kiest zelf. Kiest hij IPv6, dan weigert
// Google de aanvraag. Daarom deze calls expliciet over IPv4.
function haalJsonIPv4(url, opties) {
  const https = require('https');
  const u = new URL(url);
  return new Promise((klaar, fout) => {
    const verzoek = https.request({
      hostname: u.hostname,
      path: u.pathname + u.search,
      method: (opties && opties.method) || 'GET',
      headers: (opties && opties.headers) || {},
      family: 4,
      timeout: 15000
    }, res => {
      let body = '';
      res.on('data', d => { body += d; });
      res.on('end', () => {
        let data = {};
        try { data = JSON.parse(body); } catch (e) { /* niet-JSON antwoord */ }
        klaar({ ok: res.statusCode >= 200 && res.statusCode < 300, status: res.statusCode, data });
      });
    });
    verzoek.on('timeout', () => verzoek.destroy(new Error('tijd verstreken')));
    verzoek.on('error', fout);
    if (opties && opties.body) verzoek.write(opties.body);
    verzoek.end();
  });
}

async function zoekPlaceId() {
  if (process.env.GOOGLE_PLACE_ID) return process.env.GOOGLE_PLACE_ID;

  const res = await haalJsonIPv4('https://places.googleapis.com/v1/places:searchText', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'X-Goog-Api-Key': PLACES_KEY,
      'X-Goog-FieldMask': 'places.id,places.displayName'
    },
    body: JSON.stringify({ textQuery: 'Domein Walbrugge, Walbrugge 36, 8573 Anzegem' })
  });
  const data = res.data;
  if (!res.ok || !data.places || !data.places.length) {
    throw new Error('plaats niet gevonden: ' + JSON.stringify(data).slice(0, 200));
  }
  return data.places[0].id;
}

async function haalGoogleReviews(lang) {
  if (!PLACES_KEY) throw new Error('GOOGLE_PLACES_API_KEY niet ingesteld');
  lang = REVIEW_TALEN.includes(lang) ? lang : 'nl';

  const placeId = await zoekPlaceId();
  // languageCode: Google levert de reviewteksten (en "2 maanden geleden")
  // dan in de gevraagde taal, vertaald waar nodig.
  const res = await haalJsonIPv4(
    'https://places.googleapis.com/v1/places/' + encodeURIComponent(placeId) +
      '?languageCode=' + lang,
    {
      headers: {
        'X-Goog-Api-Key': PLACES_KEY,
        'X-Goog-FieldMask': 'rating,userRatingCount,googleMapsUri,reviews'
      }
    }
  );
  const data = res.data;
  if (!res.ok) throw new Error('Places gaf ' + res.status + ': ' + JSON.stringify(data).slice(0, 200));

  const reviews = (data.reviews || [])
    .filter(r => r.text && r.text.text && r.rating)
    .slice(0, 6)
    .map(r => ({
      auteur: (r.authorAttribution && r.authorAttribution.displayName) || 'Google-gebruiker',
      auteurUrl: (r.authorAttribution && r.authorAttribution.uri) || null,
      score: r.rating,
      tekst: r.text.text,
      wanneer: r.relativePublishTimeDescription || '',
      datum: r.publishTime || null
    }));

  const bewaard = {
    score: data.rating || null,
    aantal: data.userRatingCount || 0,
    kaartUrl: data.googleMapsUri || 'https://www.google.com/search?q=domein+walbrugge',
    reviews,
    opgehaald: Date.now()
  };

  const cachePad = placesCachePad(lang);
  fs.mkdirSync(path.dirname(cachePad), { recursive: true });
  fs.writeFileSync(cachePad, JSON.stringify(bewaard, null, 1));

  // De zichtbare score op de site mee bijwerken (score en aantal zijn in
  // elke taal gelijk; één keer volstaat).
  if (bewaard.score && lang === 'nl') {
    db.prepare(`UPDATE site_scores SET score = ?, aantal = ?, url = ?,
                bijgewerkt = CURRENT_TIMESTAMP WHERE bron = 'google'`)
      .run(String(bewaard.score).replace('.', ','), bewaard.aantal, bewaard.kaartUrl);
  }
  console.log('[google] ' + lang + ': score ' + bewaard.score + ' uit ' + bewaard.aantal +
              ' beoordelingen, ' + reviews.length + ' reviews bewaard');
  return bewaard;
}

async function ververGoogleReviews() {
  if (!PLACES_KEY) return;
  for (const lang of REVIEW_TALEN) {
    const cache = leesReviewCache(lang);
    if (cache && Date.now() - cache.opgehaald < PLACES_MAX_LEEFTIJD) continue;
    try {
      await haalGoogleReviews(lang);
    } catch (err) {
      // Bij een fout blijft de vorige cache staan: liever een cijfer van
      // gisteren dan geen cijfer.
      console.error('[google] Ophalen (' + lang + ') mislukt:', err.message);
    }
  }
}

// Bij de start en daarna elke zes uur kijken of de cache ververst moet worden.
setTimeout(ververGoogleReviews, 20000).unref();
setInterval(ververGoogleReviews, 6 * 60 * 60 * 1000).unref();

// ── Zichtbare scores ──────────────────────────────────────────────────────
// Startwaarden: wat er op de site stond. De beheerder past ze aan in het
// beheerpaneel; Google wordt automatisch overschreven zodra de API-sleutel
// ingesteld is.
const START_SCORES = [
  ['google', '5,0', null, 'https://www.google.com/search?q=domein+walbrugge', 1],
  ['booking', '9,7', null, 'https://www.booking.com/hotel/be/walbrugge.nl.html', 0],
  ['eventplanner', '10/10', null, 'https://www.eventplanner.be/directory/13607_walbrugge.html', 0]
];
START_SCORES.forEach(r => {
  db.prepare(`INSERT OR IGNORE INTO site_scores (bron, score, aantal, url, automatisch)
              VALUES (?, ?, ?, ?, ?)`).run(r[0], r[1], r[2], r[3], r[4]);
});

function leesScores() {
  const rijen = db.prepare('SELECT * FROM site_scores').all();
  const uit = {};
  rijen.forEach(r => {
    uit[r.bron] = { score: r.score, aantal: r.aantal, url: r.url,
                    automatisch: !!r.automatisch, bijgewerkt: r.bijgewerkt };
  });
  return uit;
}

app.get('/api/scores', (req, res) => {
  res.json({ ok: true, scores: leesScores() });
});

app.put('/api/admin/scores/:bron', authMiddleware('admin'), (req, res) => {
  const { score, aantal, url } = req.body || {};
  if (!score || !String(score).trim()) {
    return res.status(400).json({ error: 'Vul een score in' });
  }
  const bestaat = db.prepare('SELECT bron FROM site_scores WHERE bron = ?').get(req.params.bron);
  if (!bestaat) return res.status(404).json({ error: 'Onbekende bron' });

  db.prepare(`UPDATE site_scores SET score = ?, aantal = ?, url = COALESCE(?, url),
              bijgewerkt = CURRENT_TIMESTAMP WHERE bron = ?`)
    .run(String(score).trim(), aantal || null, url || null, req.params.bron);
  res.json({ ok: true, scores: leesScores() });
});

app.get('/api/reviews', (req, res) => {
  const lang = REVIEW_TALEN.includes(req.query.lang) ? req.query.lang : 'nl';
  const cache = leesReviewCache(lang) || leesReviewCache('nl');
  if (!cache) {
    return res.json({ ok: false, reden: PLACES_KEY ? 'nog niet opgehaald' : 'geen API-sleutel' });
  }
  res.json({
    ok: true,
    score: cache.score,
    aantal: cache.aantal,
    kaartUrl: cache.kaartUrl,
    reviews: cache.reviews,
    opgehaald: cache.opgehaald
  });
});

// security.txt — RFC 9116. express.static laat bestanden met een punt vooraan
// links liggen, dus een eigen route.
app.get(['/.well-known/security.txt', '/security.txt'], (req, res) => {
  const file = path.join(__dirname, '..', 'public', '.well-known', 'security.txt');
  if (!fs.existsSync(file)) return res.status(404).end();
  res.type('text/plain').sendFile(file);
});

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

  // Het artikel zelf server-side renderen. Voorheen stond er enkel
  // "Artikel laden..." in de HTML en kwam de tekst pas via JavaScript —
  // onzichtbaar voor zoekmachines en AI-crawlers, die geen JavaScript
  // uitvoeren. De client vult hierna enkel nog de gerelateerde artikels aan.
  const datum = post.published_at
    ? new Date(post.published_at).toLocaleDateString('nl-BE',
        { year: 'numeric', month: 'long', day: 'numeric' })
    : '';
  const kop = esc(post.featured_image || '/assets/img/hero.webp');
  const tagsHtml = post.tags
    ? '<div class="blog-tags">' + String(post.tags).split(',')
        .map(t => '<span class="blog-tag">' + esc(t.trim()) + '</span>').join('') + '</div>'
    : '';

  const artikelHtml =
    '<header class="blog-article-header" style="background-image: url(\'' + kop + '\')">' +
      '<div class="blog-article-header-overlay"></div>' +
      '<div class="blog-article-header-content">' +
        '<div class="container">' +
          '<a href="/blog" class="blog-back">← Terug naar blog</a>' +
          (post.category ? '<span class="blog-article-category">' + esc(post.category) + '</span>' : '') +
          '<h1>' + esc(post.title) + '</h1>' +
          '<div class="blog-article-meta">' +
            '<span class="blog-article-author">Door ' + esc(post.author || 'Walbrugge') + '</span>' +
            (datum ? '<time datetime="' + esc(post.published_at) + '">' + datum + '</time>' : '') +
          '</div>' +
        '</div>' +
      '</div>' +
    '</header>' +
    '<div class="blog-article-body"><div class="container container-narrow">' +
      (post.content || '') +
    '</div></div>' +
    '<footer class="blog-article-footer"><div class="container container-narrow">' +
      tagsHtml +
      '<div class="blog-share"><span>Delen:</span>' +
        '<a href="https://www.facebook.com/sharer/sharer.php?u=' + encodeURIComponent(url) +
          '" target="_blank" rel="noopener">Facebook</a>' +
        '<a href="https://www.linkedin.com/shareArticle?mini=true&url=' + encodeURIComponent(url) +
          '" target="_blank" rel="noopener">LinkedIn</a>' +
      '</div>' +
    '</div></footer>';

  html = html.replace(
    /<article class="blog-article" id="blogArticle">[\s\S]*?<\/article>/,
    '<article class="blog-article" id="blogArticle" data-ssr="1" data-category="' +
      esc(post.category || '') + '" data-post-id="' + post.id + '">' +
      artikelHtml + '</article>'
  );

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

const server = app.listen(PORT, '127.0.0.1', () => {
  console.log(`
  ╔════════════════════════════════════════════════╗
  ║  🏡 Walbrugge Backend Server                   ║
  ║  Running on http://127.0.0.1:${PORT}             ║
  ╚════════════════════════════════════════════════╝
  `);
});

// Netjes afsluiten bij een herstart (deploy): eerst geen nieuwe verbindingen
// meer aannemen en de lopende verzoeken afwerken, dan pas stoppen. Voorheen
// stopte het proces meteen, waardoor een verzoek dat net binnen was een 502
// kreeg — Caddy's lb_try_duration vangt enkel mislukte verbindingen op, geen
// afgebroken verzoeken.
function sluitNetjesAf(signaal) {
  console.log('Shutting down (' + signaal + ')...');
  const noodstop = setTimeout(() => {
    console.error('[shutdown] lopende verzoeken niet op tijd klaar, geforceerd gestopt');
    process.exit(1);
  }, 8000);
  noodstop.unref();
  server.close(() => {
    try { db.close(); } catch (e) { /* al dicht */ }
    process.exit(0);
  });
}
process.on('SIGTERM', () => sluitNetjesAf('SIGTERM'));
process.on('SIGINT', () => sluitNetjesAf('SIGINT'));
