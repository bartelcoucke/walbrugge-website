#!/usr/bin/env node
/**
 * stamp-assets.js — cache-busting voor lokale CSS/JS.
 *
 * Zet achter elke verwijzing naar /assets/css/*.css en /assets/js/*.js in de
 * HTML-pagina's een versieparameter (?v=<hash>) die afgeleid is van de inhoud
 * van het bestand. Verandert een stylesheet of script, dan verandert de hash
 * en halen browsers het bestand opnieuw op — ook al staat er op de server een
 * cache van een jaar ("immutable") op statische bestanden.
 *
 * Gebruik:  npm run stamp        (na elke wijziging aan CSS of JS, vóór commit)
 *           node scripts/stamp-assets.js --check   (alleen rapporteren, niets schrijven)
 */
'use strict';

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const ROOT = path.join(__dirname, '..', 'public');
const CHECK_ONLY = process.argv.includes('--check');

// href="/assets/css/x.css" of src="/assets/js/x.js", met of zonder bestaande ?v=
const RE = /\b(href|src)="(\/assets\/(?:css|js)\/[^"?#]+\.(?:css|js))(\?v=[^"#]*)?"/g;

const hashCache = new Map();
function hashOf(assetPath) {
  if (!hashCache.has(assetPath)) {
    const abs = path.join(ROOT, assetPath.replace(/^\//, ''));
    if (!fs.existsSync(abs)) { hashCache.set(assetPath, null); return null; }
    // Regeleinden normaliseren: op Windows staat het bestand met CRLF, op de
    // server met LF. Zo geeft dezelfde inhoud overal dezelfde hash.
    const inhoud = fs.readFileSync(abs, 'utf8').replace(/\r\n/g, '\n');
    const h = crypto.createHash('md5').update(inhoud).digest('hex').slice(0, 8);
    hashCache.set(assetPath, h);
  }
  return hashCache.get(assetPath);
}

function* walk(dir) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, entry.name);
    if (entry.isDirectory()) yield* walk(p);
    else if (entry.isFile() && p.toLowerCase().endsWith('.html')) yield p;
  }
}

let filesChanged = 0, refsUpdated = 0, missing = new Set();
for (const file of walk(ROOT)) {
  const before = fs.readFileSync(file, 'utf8');
  const after = before.replace(RE, (m, attr, asset, oldV) => {
    const h = hashOf(asset);
    if (!h) { missing.add(asset); return m; }
    const next = `${attr}="${asset}?v=${h}"`;
    if (next !== m) refsUpdated++;
    return next;
  });
  if (after !== before) {
    filesChanged++;
    if (!CHECK_ONLY) fs.writeFileSync(file, after, 'utf8');
  }
}

console.log(`${CHECK_ONLY ? '[check] ' : ''}${filesChanged} pagina's ${CHECK_ONLY ? 'zouden wijzigen' : 'bijgewerkt'}, ${refsUpdated} verwijzingen ${CHECK_ONLY ? 'te vernieuwen' : 'vernieuwd'}.`);
if (missing.size) console.log('Niet gevonden (ongewijzigd gelaten): ' + [...missing].join(', '));
if (CHECK_ONLY && filesChanged) process.exit(1);
