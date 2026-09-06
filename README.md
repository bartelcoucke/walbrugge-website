# Domein Walbrugge Website

Website voor Domein Walbrugge - Feestzaal, B&B & Vergaderlocatie te Tiegem.

## 🌐 Live
- **Website:** http://2.28.71.249
- **Admin:** http://2.28.71.249/admin

## 🚀 Deployment
Automatisch via GitHub webhook. Elke push naar `main` wordt automatisch gedeployed.

## 📁 Structuur
```
├── backend/          # Node.js Express server
│   ├── server.js     # API + routes
│   └── package.json
├── public/           # Frontend bestanden
│   ├── assets/       # CSS, JS, images
│   ├── index.html    # Homepage
│   ├── zakelijk.html # Zakelijke pagina
│   └── ...
├── Caddyfile         # Webserver config
└── walbrugge.service # Systemd service
```

## 🔐 Login
- **Admin:** info@walbrugge.be
- **Gasten:** email + boekingsreferentie

## 📞 Contact
- WhatsApp: 0499/523.325
- Email: info@walbrugge.be

## Cache-busting voor CSS en JS

De server geeft statische bestanden een cache van een jaar (`immutable`). Daarom
krijgt elke verwijzing naar `/assets/css/*.css` en `/assets/js/*.js` in de
HTML-pagina's een versieparameter op basis van de bestandsinhoud
(`style.css?v=9415e3ad`). Verandert het bestand, dan verandert de hash en halen
browsers het opnieuw op. De hash negeert regeleinden (CRLF/LF), zodat Windows en
Linux dezelfde waarde geven.

**Na elke wijziging aan een CSS- of JS-bestand, vóór het committen:**

```bash
cd backend && npm run stamp        # werkt alle verwijzingen bij
npm run stamp:check                # controleert alleen (exit 1 als er iets ontbreekt)
```

De cookie-toestemming staat los hiervan: die zit in de cookie `walbrugge_consent`
(365 dagen) en wordt door een nieuwe versieparameter niet gereset.
