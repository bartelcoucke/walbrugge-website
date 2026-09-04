#!/usr/bin/env node
/**
 * Beheerderswachtwoord instellen of wijzigen.
 *
 *   cd /opt/walbrugge/backend && node set-admin-password.js
 *
 * Het wachtwoord wordt niet getoond terwijl je typt en komt nergens in een
 * logbestand of in de shell-geschiedenis terecht.
 */

const path = require('path');
const readline = require('readline');
const bcrypt = require('bcryptjs');
const Database = require('better-sqlite3');

const EMAIL = process.argv[2] || process.env.ADMIN_EMAIL || 'admin@walbrugge.be';
const db = new Database(path.join(__dirname, '..', 'data', 'walbrugge.db'));

function vraagVerborgen(vraag) {
  return new Promise(resolve => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    const schrijf = rl._writeToOutput.bind(rl);
    let stil = false;
    rl._writeToOutput = function (s) { if (!stil) schrijf(s); };
    rl.question(vraag, antwoord => {
      rl._writeToOutput = schrijf;
      rl.close();
      process.stdout.write('\n');
      resolve(antwoord);
    });
    stil = true;
  });
}

(async () => {
  const pw1 = await vraagVerborgen(`Nieuw wachtwoord voor ${EMAIL}: `);
  if (pw1.length < 12) {
    console.error('Te kort. Gebruik minstens 12 tekens.');
    process.exit(1);
  }
  const pw2 = await vraagVerborgen('Nogmaals ter bevestiging: ');
  if (pw1 !== pw2) {
    console.error('De twee wachtwoorden komen niet overeen.');
    process.exit(1);
  }

  const hash = bcrypt.hashSync(pw1, 12);
  const bestaat = db.prepare('SELECT id FROM users WHERE email = ?').get(EMAIL);
  if (bestaat) {
    db.prepare('UPDATE users SET password_hash = ?, role = ? WHERE email = ?')
      .run(hash, 'admin', EMAIL);
    console.log(`Wachtwoord gewijzigd voor ${EMAIL}.`);
  } else {
    db.prepare("INSERT INTO users (email, password_hash, role, name) VALUES (?, ?, 'admin', 'Beheerder')")
      .run(EMAIL, hash);
    console.log(`Beheerder ${EMAIL} aangemaakt.`);
  }
  db.close();
})();
