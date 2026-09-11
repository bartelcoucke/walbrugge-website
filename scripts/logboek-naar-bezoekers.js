#!/usr/bin/env node
/**
 * logboek-naar-bezoekers.js — haalt de bezoekgeschiedenis per IP-adres terug uit het
 * toegangslogboek van Caddy.
 *
 * De eigen meting bewaart pas sinds 12 september 2026 een IP-adres. Het logboek van
 * Caddy noteert er al veel langer één bij elk verzoek, dus daaruit is te reconstrueren
 * welk adres op welke dag welke pagina's in welke volgorde bezocht, en via welke
 * verwijzer het binnenkwam. Klikken op knoppen en carrousels staan er NIET in: die
 * worden door de pagina zelf gemeld en beginnen dus pas vanaf de invoering.
 *
 * Werkwijze: voor elke dag die het logboek dekt en waarvoor nog geen enkele weergave
 * met een adres bestaat, worden de bestaande weergaven van die dag vervangen door de
 * regels uit het logboek. Zo blijft er precies één rij per paginaweergave en krijgt de
 * geschiedenis alsnog een adres.
 *
 * Gebruik (op de server, in de map van het project):
 *   node scripts/logboek-naar-bezoekers.js                 proefdraai, schrijft niets
 *   node scripts/logboek-naar-bezoekers.js --schrijf       voert de wijziging uit
 *   node scripts/logboek-naar-bezoekers.js --dagen=120     beperk tot de laatste 120 dagen
 *   CADDY_LOG=/pad/naar/walbrugge.log node scripts/...     ander logbestand
 */
'use strict';
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
const crypto = require('crypto');
const readline = require('readline');

const SCHRIJF = process.argv.includes('--schrijf');
const DAGEN = (() => {
  const a = process.argv.find(x => x.startsWith('--dagen='));
  return a ? Math.max(1, parseInt(a.slice(8), 10) || 400) : 400;
})();
const CADDY_LOG = process.env.CADDY_LOG || '/var/log/caddy/walbrugge.log';
const DB_PAD = process.env.STATS_DB || path.join(__dirname, '..', 'data', 'walbrugge.db');

let Database;
try {
  Database = require(path.join(__dirname, '..', 'backend', 'node_modules', 'better-sqlite3'));
} catch (e) {
  try { Database = require('better-sqlite3'); }
  catch (e2) { console.error('better-sqlite3 niet gevonden. Draai dit script vanuit de projectmap.'); process.exit(2); }
}

// ── Dezelfde regels als de meting in backend/server.js ──────────────────────
const STATS_BOT = /bot|crawl|spider|slurp|preview|monitor|fetch|scan|curl|wget|python|java\/|headless|lighthouse|pingdom|uptime|facebookexternalhit|whatsapp|telegrambot|linkedinbot|semrush|ahrefs|mj12|dotbot|petalbot|bytespider|gptbot|claudebot|ccbot/i;
const EIGEN_HOSTS = new Set(['walbrugge.be', 'www.walbrugge.be', '2.28.71.249', 'localhost', '127.0.0.1']);

function normaliseerIp(ruw) {
  let ip = String(ruw || '').trim();
  if (!ip) return null;
  if (/^\[/.test(ip)) ip = ip.replace(/^\[|\]$/g, '');        // [::1]:1234
  if (ip.startsWith('::ffff:')) ip = ip.slice(7);
  ip = ip.split('%')[0];
  if (ip.indexOf(':') < 0) return ip.split(':')[0].slice(0, 45);
  const [voor, na] = ip.split('::');
  const a = voor ? voor.split(':').filter(Boolean) : [];
  const b = na !== undefined ? (na ? na.split(':').filter(Boolean) : []) : null;
  let groepen = b === null ? a : a.concat(Array(Math.max(0, 8 - a.length - b.length)).fill('0'), b);
  if (groepen.length < 4) groepen = groepen.concat(Array(4 - groepen.length).fill('0'));
  return groepen.slice(0, 4).map(g => g.replace(/^0+(?=.)/, '')).join(':') + '::/64';
}
const eigenIp = ip => !ip || ip === '::1' || ip === '0:0:0:0::/64' || /^127\./.test(ip)
  || /^(10\.|192\.168\.|169\.254\.)/.test(ip) || /^172\.(1[6-9]|2[0-9]|3[01])\./.test(ip);

function taalVan(pad) {
  const m = /^\/(fr|en|de)(\/|$)/.exec(pad || '');
  return m ? m[1] : 'nl';
}
function bronVan(ref, host) {
  const r = String(ref || '').trim();
  const app = /^android-app:\/\/([^/?#]+)/i.exec(r);
  if (app) return 'app: ' + app[1].toLowerCase().slice(0, 60);
  if (!r) return null;
  try {
    const h = new URL(r).hostname.toLowerCase();
    if (!h || EIGEN_HOSTS.has(h) || h === String(host || '').toLowerCase().split(':')[0]) return null;
    return h.slice(0, 120);
  } catch (e) { return null; }
}
const kop = (headers, naam) => {
  if (!headers) return '';
  const k = Object.keys(headers).find(x => x.toLowerCase() === naam.toLowerCase());
  const v = k ? headers[k] : null;
  return Array.isArray(v) ? String(v[0] || '') : String(v || '');
};
const dagVan = ms => new Date(ms).toLocaleDateString('sv-SE', { timeZone: 'Europe/Brussels' });
const tsVan = ms => new Date(ms).toISOString().slice(0, 19).replace('T', ' ');

// ── Logbestanden zoeken, oudste eerst ───────────────────────────────────────
function logBestanden() {
  const dir = path.dirname(CADDY_LOG);
  const basis = path.basename(CADDY_LOG).replace(/\.log$/, '');
  let namen = [];
  try { namen = fs.readdirSync(dir); } catch (e) {
    console.error('Logmap niet leesbaar: ' + dir + ' (' + e.message + ')');
    return [];
  }
  return namen
    .filter(n => n === basis + '.log' || (n.startsWith(basis + '-') && /\.log(\.gz)?$/.test(n)))
    .map(n => { const p = path.join(dir, n); const st = fs.statSync(p); return { pad: p, naam: n, mtime: st.mtimeMs, grootte: st.size, gz: n.endsWith('.gz') }; })
    .sort((a, b) => a.mtime - b.mtime);
}

(async () => {
  const bestanden = logBestanden();
  if (!bestanden.length) {
    console.error('Geen logbestanden gevonden op ' + CADDY_LOG + '.');
    console.error('Controleer het pad, of geef het mee met CADDY_LOG=/pad/naar/walbrugge.log');
    process.exit(1);
  }
  console.log('Logbestanden: ' + bestanden.length + ' (' + Math.round(bestanden.reduce((s, b) => s + b.grootte, 0) / 1048576) + ' MB)');
  bestanden.forEach(b => console.log('   ' + b.naam.padEnd(40) + Math.round(b.grootte / 1024) + ' kB'));

  const db = new Database(DB_PAD);
  const heeftIp = db.prepare("SELECT name FROM pragma_table_info('visits') WHERE name = 'ip'").get();
  if (!heeftIp) { console.error('De kolom visits.ip bestaat nog niet. Start eerst de server één keer.'); process.exit(2); }
  if (!db.prepare("SELECT name FROM pragma_table_info('visits') WHERE name = 'uit_logboek'").get()) {
    if (SCHRIJF) db.exec('ALTER TABLE visits ADD COLUMN uit_logboek INTEGER DEFAULT 0');
    console.log(SCHRIJF ? 'Kolom visits.uit_logboek toegevoegd.' : 'Kolom visits.uit_logboek zou toegevoegd worden.');
  }

  // Dagen die de eigen meting al mét adres dekt: die blijven ongemoeid.
  const dagenMetIp = new Set(db.prepare('SELECT DISTINCT dag FROM visits WHERE ip IS NOT NULL').all().map(r => r.dag));
  const grens = dagVan(Date.now() - (DAGEN - 1) * 864e5);
  console.log('\nDagen die de eigen meting al met adres dekt: ' + (dagenMetIp.size || 'geen'));
  console.log('Grens: enkel dagen vanaf ' + grens + '\n');

  const geheim = (db.prepare("SELECT waarde FROM stats_geheim WHERE sleutel = 'toestel'").get() || {}).waarde
    || crypto.randomBytes(32).toString('hex');
  const toestelId = (ip, ua) => crypto.createHash('sha256').update(geheim + '|' + (ip || '') + '|' + (ua || '')).digest('hex').slice(0, 12);

  const perDag = new Map();          // dag -> rijen
  const t = { regels: 0, onleesbaar: 0, overgeslagen: 0, bots: 0, bewaard: 0 };

  for (const b of bestanden) {
    const stroom = b.gz ? fs.createReadStream(b.pad).pipe(zlib.createGunzip()) : fs.createReadStream(b.pad);
    const lezer = readline.createInterface({ input: stroom, crlfDelay: Infinity });
    for await (const regel of lezer) {
      if (!regel.trim()) continue;
      t.regels++;
      let e;
      try { e = JSON.parse(regel); } catch (err) { t.onleesbaar++; continue; }
      const q = e && e.request;
      if (!q) { t.overgeslagen++; continue; }

      let ms = typeof e.ts === 'number' ? e.ts * 1000 : Date.parse(e.ts);
      if (!ms || isNaN(ms)) { t.overgeslagen++; continue; }
      const dag = dagVan(ms);
      if (dag < grens || dagenMetIp.has(dag)) { t.overgeslagen++; continue; }

      if ((q.method || 'GET') !== 'GET') { t.overgeslagen++; continue; }
      if (e.status && e.status !== 200 && e.status !== 304) { t.overgeslagen++; continue; }
      const pad = String(q.uri || '').split('?')[0];
      if (/^\/(api|admin|gasten|login|media)(\/|$)/.test(pad) || /^\/(fr|en|de)\/login$/.test(pad)) { t.overgeslagen++; continue; }
      if (/^\/google[0-9a-f]+\.html$/.test(pad)) { t.overgeslagen++; continue; }

      const ct = kop(e.resp_headers, 'Content-Type');
      const isPagina = ct ? /text\/html/i.test(ct) : (!/\.[a-z0-9]{2,5}$/i.test(pad) || /\.html?$/i.test(pad));
      if (!isPagina) { t.overgeslagen++; continue; }

      const ua = kop(q.headers, 'User-Agent');
      if (!ua || STATS_BOT.test(ua)) { t.bots++; continue; }
      const doel = kop(q.headers, 'Sec-Purpose') || kop(q.headers, 'Purpose');
      if (/prefetch|preview/i.test(doel)) { t.overgeslagen++; continue; }

      const ip = normaliseerIp(q.remote_ip || q.client_ip || e.remote_ip || e.client_ip);
      if (eigenIp(ip)) { t.overgeslagen++; continue; }

      const zoek = String(q.uri || '').split('?')[1] || '';
      const p = new URLSearchParams(zoek);
      if (!perDag.has(dag)) perDag.set(dag, []);
      perDag.get(dag).push([
        dag, pad.slice(0, 200), taalVan(pad), bronVan(kop(q.headers, 'Referer'), q.host),
        p.get('utm_source') || null, p.get('utm_medium') || null, p.get('utm_campaign') || null,
        /Mobi|Android|iPhone|iPad/i.test(ua) ? 'mobiel' : 'desktop',
        null, ip, ms, null, toestelId(ip, ua), tsVan(ms),
      ]);
      t.bewaard++;
    }
  }

  console.log('Gelezen: ' + t.regels + ' regels | onleesbaar ' + t.onleesbaar + ' | overgeslagen ' + t.overgeslagen
    + ' | robots ' + t.bots + ' | bruikbaar ' + t.bewaard);

  const dagen = [...perDag.keys()].sort();
  if (!dagen.length) { console.log('\nNiets te herstellen: het logboek dekt geen dagen die nog geen adres hebben.'); process.exit(0); }

  console.log('\nDagen die hersteld kunnen worden: ' + dagen.length + ' (' + dagen[0] + ' t/m ' + dagen[dagen.length - 1] + ')');
  const bestaand = db.prepare('SELECT dag, COUNT(*) AS n FROM visits WHERE dag = ?');
  let totaalNu = 0, totaalLog = 0, adressen = new Set();
  for (const dag of dagen) {
    const nu = (bestaand.get(dag) || {}).n || 0;
    const log = perDag.get(dag).length;
    totaalNu += nu; totaalLog += log;
    perDag.get(dag).forEach(r => adressen.add(r[9]));
    if (dagen.length <= 40) console.log('   ' + dag + '  nu ' + String(nu).padStart(5) + ' weergaven  ->  uit logboek ' + String(log).padStart(5));
  }
  console.log('   ' + '-'.repeat(52));
  console.log('   totaal      nu ' + String(totaalNu).padStart(5) + ' weergaven  ->  uit logboek ' + String(totaalLog).padStart(5));
  console.log('   verschillende adressen: ' + adressen.size);

  if (!SCHRIJF) {
    console.log('\nProefdraai: er is niets gewijzigd. Voer uit met --schrijf om dit door te voeren.');
    console.log('De bestaande weergaven van die dagen worden dan vervangen door de regels uit het logboek.');
    process.exit(0);
  }

  const wis = db.prepare('DELETE FROM visits WHERE dag = ?');
  const zet = db.prepare(`INSERT INTO visits (dag, pad, taal, bron, utm_source, utm_medium, utm_campaign, toestel, bezoeker, ip, ms, pagina_id, toestel_id, ts, uit_logboek)
                          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)`);
  const toestelZet = db.prepare(`INSERT INTO stats_toestellen (toestel_id, ip, ua, toestel, eerste_ts, laatste_ts, js)
                                 VALUES (?, ?, NULL, ?, ?, ?, 0)
                                 ON CONFLICT(toestel_id) DO UPDATE SET
                                   eerste_ts = MIN(stats_toestellen.eerste_ts, excluded.eerste_ts),
                                   laatste_ts = MAX(stats_toestellen.laatste_ts, excluded.laatste_ts)`);
  const alles = db.transaction(() => {
    for (const dag of dagen) {
      wis.run(dag);
      for (const rij of perDag.get(dag)) {
        zet.run(...rij);
        toestelZet.run(rij[12], rij[9], rij[7], rij[13], rij[13]);
      }
    }
  });
  alles();
  console.log('\nKlaar: ' + totaalLog + ' weergaven over ' + dagen.length + ' dagen hersteld, ' + adressen.size + ' verschillende adressen.');
  console.log('Zichtbaar in het beheerpaneel onder Bezoekers. Klikken op knoppen en carrousels staan niet in');
  console.log('het logboek en ontbreken dus voor die dagen.');
})().catch(e => { console.error('Mislukt: ' + (e && e.stack || e)); process.exit(1); });
