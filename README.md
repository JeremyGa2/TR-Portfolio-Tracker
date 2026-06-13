<div align="center">

<img src="https://upload.wikimedia.org/wikipedia/commons/thumb/1/1c/Trade_Republic_Logo.svg/200px-Trade_Republic_Logo.svg.png" width="60" alt="Trade Republic"/>

# 📊 TR Portfolio Tracker

### *Votre portefeuille Trade Republic, visualisé comme jamais.*

<br/>

[![🚀 Live Demo](https://img.shields.io/badge/🚀_Live_Demo-jeremyga2.github.io-black?style=for-the-badge)](https://jeremyga2.github.io/TR-Portfolio-Tracker/)
[![HTML](https://img.shields.io/badge/Standalone-HTML-E34F26?style=for-the-badge&logo=html5&logoColor=white)](./Portfolio_Tracker.html)
[![React](https://img.shields.io/badge/React-18-61DAFB?style=for-the-badge&logo=react&logoColor=black)](https://react.dev)
[![No Install](https://img.shields.io/badge/Installation-Aucune-22c55e?style=for-the-badge&logo=checkmarx&logoColor=white)](#)
[![Privacy](https://img.shields.io/badge/Données-100%25_Locales-8b5cf6?style=for-the-badge&logo=privateinternetaccess&logoColor=white)](#)

<br/>

> **Importez votre CSV Trade Republic → obtenez un dashboard complet en 10 secondes.**
> Aucun compte. Aucun serveur. Aucune fuite de données.

</div>

---

## 🎬 Comment ça marche

```mermaid
flowchart LR
    A([🏦 Trade Republic\nApp]) -->|1. Exporter| B([📄 fichier .csv])
    B -->|2. Ouvrir| C([🌐 Portfolio_Tracker\n.html])
    C -->|3. Cliquer\n➕ Ajouter un fichier| D([📤 Upload CSV])
    D -->|Instantané| E([📊 Dashboard\nComplet])

    style A fill:#1a1a2e,color:#fff,stroke:#6366f1
    style B fill:#16213e,color:#fff,stroke:#6366f1
    style C fill:#0f3460,color:#fff,stroke:#6366f1
    style D fill:#533483,color:#fff,stroke:#6366f1
    style E fill:#22c55e,color:#fff,stroke:#16a34a
```

---

## 🚀 Guide pas à pas

### Étape 1 — Exporter depuis Trade Republic

Ouvrez l'application **Trade Republic** sur votre téléphone :

```
📱 Trade Republic App
 └── 👤 Profil  (icône en bas à droite)
      └── 📋 Relevés
           └── 📤 Exportation des transactions
                └── ✅ Télécharger le .csv
```

> Le fichier s'appelle quelque chose comme `Exportation de transactions.csv`

---

### Étape 2 — Ouvrir l'application

Téléchargez [`Portfolio_Tracker.html`](./Portfolio_Tracker.html) et ouvrez-le dans votre navigateur (double-clic suffit).

---

### Étape 3 — Importer votre fichier

<table>
<tr>
<td width="60px" align="center">🖱️</td>
<td>Cliquez sur le bouton <kbd>➕ Ajouter un fichier</kbd> dans l'interface</td>
</tr>
<tr>
<td align="center">📂</td>
<td>Sélectionnez votre fichier <code>.csv</code> exporté depuis Trade Republic</td>
</tr>
<tr>
<td align="center">⚡</td>
<td>Le dashboard se charge <strong>instantanément</strong> — toutes vos transactions sont analysées</td>
</tr>
</table>

---

### Étape 4 — Explorer votre portefeuille

| Onglet | Ce que vous verrez |
|--------|-------------------|
| 🏠 **Vue globale** | Valeur totale, P&L, rendement, répartition |
| 📈 **Performance** | Évolution du portefeuille dans le temps |
| 🗂️ **Par secteur** | Répartition défense / tech / crypto / ETF… |
| 💼 **Par actif** | Chaque position : prix moyen, gain/perte |
| 📅 **Historique** | Toutes vos transactions ligne par ligne |

---

## ✨ Fonctionnalités

<table>
<tr>
  <td>📥 <b>Import CSV natif</b></td>
  <td>Lit directement le format Trade Republic officiel</td>
</tr>
<tr>
  <td>🗂️ <b>Secteurs auto</b></td>
  <td>Détection par ISIN + mots-clés, 15+ secteurs reconnus</td>
</tr>
<tr>
  <td>📊 <b>Graphiques riches</b></td>
  <td>Courbes, camemberts, barres — tout interactif</td>
</tr>
<tr>
  <td>💰 <b>Prix actualisables</b></td>
  <td>Entrez les cours du jour en un clic</td>
</tr>
<tr>
  <td>👁️ <b>Mode discret</b></td>
  <td>Masquez les montants d'un clic (présentation en public)</td>
</tr>
<tr>
  <td>🔒 <b>Zéro fuite</b></td>
  <td>Tout reste dans votre navigateur — jamais envoyé ailleurs</td>
</tr>
<tr>
  <td>💾 <b>Mémoire automatique</b></td>
  <td>Vos données sont sauvegardées localement entre les sessions</td>
</tr>
</table>

---

## 🗂️ Secteurs reconnus

<div align="center">

| 🔴 Défense | 🔵 Semi-conducteurs | 🟢 Uranium & Nucléaire |
|:---:|:---:|:---:|
| Rheinmetall, Thales, Leonardo… | NVIDIA, ASML, TSMC, AMD… | Cameco, Kazatomprom… |

| 🟡 Or & Métaux précieux | 🟣 Crypto | 🔵 ETF Monde |
|:---:|:---:|:---:|
| Xetra-Gold, Physical Gold… | Bitcoin, Ethereum, 21Shares… | MSCI World, S&P 500, FTSE… |

| ⚡ Tech & Mega Caps | 🟤 Commodities | 🟠 Énergie |
|:---:|:---:|:---:|
| Apple, MSFT, Google, Meta… | Glencore, BHP, Rio Tinto… | TotalEnergies, Shell… |

| 💳 Finance | 💊 Santé & Pharma | 💎 Luxe |
|:---:|:---:|:---:|
| JPMorgan, Goldman, BNP… | Pfizer, Novartis, Sanofi… | LVMH, Hermès, Kering… |

</div>

> Actif non reconnu ? Il tombe dans *Autre* — vous pouvez forcer le secteur manuellement dans l'interface.

---

## 🛠️ Stack technique

```
📦 Portfolio_Tracker.html  (fichier unique autonome)
 ├── ⚛️  React 18          — UI (chargé via CDN, aucun build)
 ├── 📈  Recharts          — Graphiques interactifs
 ├── 🎨  Lucide React      — Icônes
 ├── 📊  SheetJS (xlsx)    — Lecture CSV / Excel
 └── 💾  localStorage      — Persistance locale des données
```

---

## 📁 Structure du projet

```
TR-Portfolio-Tracker/
├── 📄 Portfolio_Tracker.html   ← Ouvrir ça dans le navigateur
├── ⚛️  Portfolio_Tracker.jsx   ← Code source React
├── 🚫 .gitignore               ← CSV exclus (données privées)
└── 📖 README.md
```

---

<details>
<summary>🤝 <b>Contribuer au projet</b></summary>

<br/>

1. Forkez le projet
2. Créez votre branche : `git checkout -b feature/ma-fonctionnalite`
3. Commitez : `git commit -m 'feat: description'`
4. Poussez : `git push origin feature/ma-fonctionnalite`
5. Ouvrez une **Pull Request**

</details>

<details>
<summary>⚠️ <b>Note sur la confidentialité</b></summary>

<br/>

Les fichiers CSV Trade Republic contiennent vos données financières personnelles.
Ils sont explicitement exclus du dépôt via `.gitignore` — ils ne seront **jamais** committés par accident.

</details>

---

<div align="center">

**Faites croître votre richesse, pas votre stress.** 📈💡

*Questions ou suggestions → [ouvrir une issue](https://github.com/JeremyGa2/TR-Portfolio-Tracker/issues)*

<br/>

![GitHub last commit](https://img.shields.io/github/last-commit/JeremyGa2/TR-Portfolio-Tracker?style=flat-square&color=6366f1)
![GitHub repo size](https://img.shields.io/github/repo-size/JeremyGa2/TR-Portfolio-Tracker?style=flat-square&color=22c55e)
![GitHub stars](https://img.shields.io/github/stars/JeremyGa2/TR-Portfolio-Tracker?style=flat-square&color=f59e0b)

</div>
