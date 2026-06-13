<div align="center">

# 📊 TR Portfolio Tracker

**Suivez et analysez votre portefeuille Trade Republic — 100% dans votre navigateur, zéro serveur.**

[![HTML](https://img.shields.io/badge/HTML-Standalone-orange?style=flat-square&logo=html5)](./Portfolio_Tracker.html)
[![React](https://img.shields.io/badge/React-18-61dafb?style=flat-square&logo=react)](https://react.dev)
[![Recharts](https://img.shields.io/badge/Recharts-Graphiques-22c55e?style=flat-square)](https://recharts.org)
[![Licence](https://img.shields.io/badge/Licence-MIT-blue?style=flat-square)](./LICENSE)
[![Confidentialité](https://img.shields.io/badge/Données-100%25%20locales-purple?style=flat-square&logo=shield)](.)

</div>

---

## ✨ Fonctionnalités

| | Fonctionnalité | Description |
|---|---|---|
| 📥 | **Import CSV** | Glissez-déposez l'export Trade Republic, c'est tout |
| 📈 | **Dashboard temps réel** | Valeur du portefeuille, P&L, rendement global |
| 🗂️ | **Secteurs automatiques** | Détection par ISIN et nom : Défense, Semi-conducteurs, Uranium, Tech, ETF… |
| 🎨 | **Graphiques interactifs** | Évolution, répartition, performance par actif (Recharts) |
| 💰 | **Prix manuels** | Mettez à jour les cours en un clic |
| 🔒 | **100% privé** | Aucune donnée ne quitte votre machine — tout en localStorage |

---

## 🚀 Démarrage en 30 secondes

> **Aucune installation requise.** Pas de Node.js, pas de serveur, pas de compte.

**1.** Téléchargez [`Portfolio_Tracker.html`](./Portfolio_Tracker.html)

**2.** Exportez vos transactions depuis Trade Republic :
> *Profil → Relevés → Exportation de transactions → `.csv`*

**3.** Ouvrez le fichier HTML dans votre navigateur et uploadez le CSV

C'est tout. 🎉

---

## 🗂️ Secteurs reconnus automatiquement

```
🔴 Défense          🔵 Semi-conducteurs     🟢 Uranium & Nucléaire
🟡 Or & Métaux      🟣 Crypto               🔵 ETF Monde / Émergents
⚡ Tech & Mega Caps  🟤 Commodities          🟠 Énergie
💳 Finance          💊 Santé & Pharma       💎 Luxe
```

Détection par **ISIN** en priorité, puis par **mots-clés** sur le nom de l'actif.
Tout actif non reconnu tombe dans *Autre* — et vous pouvez forcer le secteur manuellement.

---

## 🛠️ Stack technique

| Couche | Technologie |
|---|---|
| UI | React 18 (via CDN — aucun build nécessaire) |
| Graphiques | Recharts |
| Icônes | Lucide React |
| Fichiers Excel/CSV | SheetJS (xlsx) |
| Persistance | `localStorage` du navigateur |

---

## 📁 Structure du projet

```
TR-Portfolio-Tracker/
├── Portfolio_Tracker.html   # Application complète (ouvrir directement)
├── Portfolio_Tracker.jsx    # Code source React
├── .gitignore
└── README.md
```

> ⚠️ Les fichiers CSV Trade Republic contiennent des données financières personnelles
> et sont exclus du dépôt via `.gitignore`.

---

## 🤝 Contribuer

1. Forkez le projet
2. Créez votre branche (`git checkout -b feature/ma-fonctionnalite`)
3. Commitez (`git commit -m 'feat: description'`)
4. Poussez (`git push origin feature/ma-fonctionnalite`)
5. Ouvrez une Pull Request

---

<div align="center">

**Faites croître votre richesse, pas votre stress.** 📈

*Questions ou suggestions → [ouvrir une issue](https://github.com/JeremyGa2/TR-Portfolio-Tracker/issues)*

</div>
