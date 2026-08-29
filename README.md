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
- **Admin:** admin@walbrugge.be
- **Gasten:** email + boekingsreferentie

## 📞 Contact
- WhatsApp: 0499/523.325
- Email: info@walbrugge.be
