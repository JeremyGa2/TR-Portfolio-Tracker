import { useState, useMemo, useCallback, useEffect, useRef, Fragment } from "react";
import * as XLSX from "xlsx";
import {
  LineChart, Line, BarChart, Bar, PieChart, Pie, Cell, XAxis, YAxis,
  CartesianGrid, Tooltip, Legend, ResponsiveContainer, Area, AreaChart,
  ComposedChart, ScatterChart, Scatter, ReferenceLine,
} from "recharts";
import {
  TrendingUp, TrendingDown, Upload, Briefcase, BarChart3, Calendar,
  Download, Trash2, Search, Settings as SettingsIcon, Activity,
  Target, Award, AlertCircle, Zap, Clock, Database, FileText,
  Check, RefreshCw, ArrowUpRight, ArrowDownRight, ChevronLeft, ChevronRight,
  DollarSign, Percent, Hash, PieChart as PieIcon, Flame, CalendarDays,
  Gift, Coins, Eye, EyeOff, Info, Sparkles, Wallet, CreditCard,
} from "lucide-react";

/* PORTFOLIO TRACKER PRO v3 — Parser Trade Republic officiel */

const STORAGE_KEYS = {
  TRANSACTIONS: "tr_transactions_v3",
  SECTOR_OVERRIDES: "tr_sector_overrides_v2",
  CURRENT_PRICES: "tr_current_prices_v2",
  HIDE_BALANCES: "tr_hide_balances_v1",
};

const storage = {
  get: async (key) => {
    try { if (typeof window !== "undefined" && window.storage?.get) return await window.storage.get(key); } catch (e) {}
    try { const v = typeof localStorage !== "undefined" ? localStorage.getItem(key) : null; return v ? { key, value: v } : null; } catch (e) { return null; }
  },
  set: async (key, value) => {
    try { if (typeof window !== "undefined" && window.storage?.set) return await window.storage.set(key, value); } catch (e) {}
    try { localStorage.setItem(key, value); } catch (e) {}
    return { key, value };
  },
};

// ═════ SECTEURS ═════
const ISIN_SECTOR = {
  "DE0007030009": "Défense", "FR0000121329": "Défense", "IT0003856405": "Défense",
  "FR0014004L86": "Défense", "IE000OJ5TQP4": "Défense", "IE000YYE6WK5": "Défense",
  "IE000U9ODG19": "Défense", "US5398301094": "Défense", "US6516391066": "Défense",
  "US67066G1040": "Semi-conducteurs", "NL0010273215": "Semi-conducteurs",
  "US0079031078": "Semi-conducteurs", "US4581401001": "Semi-conducteurs",
  "SE0005999836": "Semi-conducteurs", "US11135F1012": "Semi-conducteurs",
  "CA13321L1085": "Uranium & Nucléaire", "DE000SIE7J08": "Uranium & Nucléaire",
  "US9168961038": "Uranium & Nucléaire",
  "IE00B3RBWM25": "ETF Monde", "IE00B4L5Y983": "ETF Monde", "IE00B5BMR087": "ETF Monde",
  "IE00BKM4GZ66": "ETF Émergents", "IE00BGHQ0G80": "ETF Monde",
  "IE00B4ND3602": "Or & Métaux précieux", "DE000A0S9GB0": "Or & Métaux précieux",
  "DE000A27Z304": "Crypto", "DE000A3GK2N1": "Crypto",
  "US0378331005": "Tech & Mega Caps", "US5949181045": "Tech & Mega Caps",
  "US02079K3059": "Tech & Mega Caps", "US0231351067": "Tech & Mega Caps",
  "US30303M1027": "Tech & Mega Caps", "FR0014003TT8": "Tech & Mega Caps",
  "FR0000120271": "Énergie", "US30231G1022": "Énergie",
  "GB00BH4HKS39": "Commodities",
};

const KEYWORD_SECTOR = [
  [["rheinmetall","lockheed","thales","leonardo","northrop","raytheon","dassault avi","airbus","future of defence","defense usd","aerospace & defence","aerospace defence"], "Défense"],
  [["nvidia","amd","asml","tsmc","taiwan semi","intel","broadcom","qualcomm","micron","sivers","infineon","arm holdings","soitec"], "Semi-conducteurs"],
  [["cameco","uranium","kazatomprom","nuclear","nuscale","paladin","siemens energy"], "Uranium & Nucléaire"],
  [["bitcoin","ethereum","crypto","btc","eth","21shares"], "Crypto"],
  [["gold","xetra-gold","physical gold","silver","platinum"], "Or & Métaux précieux"],
  [["msci world","ftse all-world","s&p 500","sp500","core s&p","nasdaq 100","msci em","emerging"], "ETF Monde"],
  [["apple","microsoft","alphabet","google","amazon","meta","tesla","netflix","dassault systèmes","mercadolibre","vertiv","vusion"], "Tech & Mega Caps"],
  [["totalenergies","shell","exxon","chevron","bp plc","equinor"], "Énergie"],
  [["glencore","bhp","rio tinto","freeport","standard chartered","anglo american"], "Commodities"],
  [["jpmorgan","bank of america","goldman","hsbc","bnp paribas","deutsche bank"], "Finance & Banques"],
  [["pfizer","novartis","roche","merck","eli lilly","sanofi","onco-innovations"], "Santé & Pharma"],
  [["lvmh","hermes","kering","louis vuitton","loreal","l'oreal"], "Luxe"],
  [["nestle","procter","unilever","coca-cola","pepsi"], "Consommation"],
  [["siemens","abb","caterpillar","deere","honeywell"], "Industrie"],
  [["long ","short "], "Dérivés & Warrants"],
];

const SECTOR_COLORS = {
  "Défense": "#f43f5e", "Semi-conducteurs": "#06b6d4", "Uranium & Nucléaire": "#22c55e",
  "Commodities": "#f59e0b", "Énergie": "#eab308", "Or & Métaux précieux": "#fbbf24",
  "Crypto": "#a855f7", "ETF Monde": "#3b82f6", "ETF Émergents": "#8b5cf6",
  "Tech & Mega Caps": "#0ea5e9", "Finance & Banques": "#10b981", "Santé & Pharma": "#ec4899",
  "Luxe": "#d946ef", "Consommation": "#14b8a6", "Industrie": "#78716c",
  "Dérivés & Warrants": "#f97316", "Autre": "#64748b",
};

// ═════ UTILS ═════
const parseNumber = (v) => {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  let s = String(v).trim().replace(/[€$£\s\u00A0]/g, "");
  if (s.includes(",") && s.includes(".")) {
    if (s.lastIndexOf(",") > s.lastIndexOf(".")) s = s.replace(/\./g, "").replace(",", ".");
    else s = s.replace(/,/g, "");
  } else if (s.includes(",")) s = s.replace(",", ".");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
};

const parseDate = (v) => {
  if (!v) return null;
  if (v instanceof Date) return v;
  if (typeof v === "number") return new Date(Date.UTC(1899, 11, 30) + v * 86400000);
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?(Z|[+\-]\d{2}:?\d{2})?)?/);
  if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)));
  m = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/);
  if (m) { let y = +m[3]; if (y < 100) y += 2000; return new Date(Date.UTC(y, +m[2] - 1, +m[1])); }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
};

const fmtEUR = (n, d = 2) => {
  if (n === null || n === undefined || isNaN(n)) return "—";
  const sign = n < 0 ? "−" : "";
  return sign + Math.abs(n).toLocaleString("fr-FR", { minimumFractionDigits: d, maximumFractionDigits: d }) + " €";
};
const fmtEURCompact = (n) => {
  if (n === null || n === undefined || isNaN(n)) return "—";
  const abs = Math.abs(n), sign = n < 0 ? "−" : "";
  if (abs >= 1e6) return sign + (abs / 1e6).toLocaleString("fr-FR", { maximumFractionDigits: 2 }) + " M€";
  if (abs >= 1e3) return sign + (abs / 1e3).toLocaleString("fr-FR", { maximumFractionDigits: 1 }) + " k€";
  return sign + abs.toLocaleString("fr-FR", { maximumFractionDigits: 0 }) + " €";
};
const fmtNum = (n, d = 2) => (n === null || n === undefined || isNaN(n)) ? "—" : n.toLocaleString("fr-FR", { minimumFractionDigits: d, maximumFractionDigits: d });
const fmtPct = (n, d = 2) => (n === null || n === undefined || isNaN(n)) ? "—" : (n > 0 ? "+" : "") + n.toFixed(d) + " %";
const fmtDate = (d) => d ? d.toISOString().slice(0, 10) : "—";
const fmtDateShort = (d) => !d ? "—" : d.toLocaleDateString("fr-FR", { day: "2-digit", month: "short", year: "2-digit" });
const fmtDateLong = (d) => !d ? "—" : d.toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "long", year: "numeric" });
const classFor = (n) => n > 0 ? "text-emerald-400" : n < 0 ? "text-rose-400" : "text-slate-400";
const mask = (s, h) => h ? "•••••" : s;

const detectSector = (tx, overrides) => {
  if (!tx) return "Autre";
  const sym = tx.symbol || tx.isin;
  if (overrides[sym]) return overrides[sym];
  if (ISIN_SECTOR[sym]) return ISIN_SECTOR[sym];
  const name = (tx.name || "").toLowerCase();
  for (const [kws, sec] of KEYWORD_SECTOR) if (kws.some(k => name.includes(k))) return sec;
  if (tx.assetClass === "CRYPTO") return "Crypto";
  if (tx.assetClass === "DERIVATIVE") return "Dérivés & Warrants";
  return "Autre";
};

// ═════ PARSER CSV ═════
function parseCSVText(text) {
  const lines = []; let cur = [], field = "", inQuote = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuote) {
      if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else inQuote = false; }
      else field += c;
    } else {
      if (c === '"') inQuote = true;
      else if (c === ",") { cur.push(field); field = ""; }
      else if (c === "\n") { cur.push(field); lines.push(cur); cur = []; field = ""; }
      else if (c !== "\r") field += c;
    }
  }
  if (field.length || cur.length) { cur.push(field); lines.push(cur); }
  return lines.filter(r => r.some(c => c && c.length));
}

// ═════ PARSER TRADE REPUBLIC ═════
const TR_TYPE_MAP = {
  "BUY": "BUY", "SELL": "SELL", "DIVIDEND": "DIVIDEND",
  "INTEREST_PAYMENT": "INTEREST",
  "CUSTOMER_INBOUND": "DEPOSIT", "TRANSFER_INSTANT_INBOUND": "DEPOSIT",
  "CUSTOMER_INPAYMENT": "DEPOSIT",
  "CUSTOMER_OUTBOUND": "WITHDRAWAL", "TRANSFER_INSTANT_OUTBOUND": "WITHDRAWAL",
  "BENEFITS_SAVEBACK": "SAVEBACK", "BENEFITS_SPARE_CHANGE": "SAVEBACK",
  "WARRANT_EXERCISE": "WARRANT_EXERCISE", "TILG": "TILG",
  "CARD_TRANSACTION": "CARD", "CARD_REFUND": "CARD_REFUND",
  "GIFT": "GIFT", "TAX_OPTIMIZATION": "TAX_OPT", "PEA_MARKETING": "OTHER",
  "FREE_RECEIPT": "FREE_RECEIPT",
};

function parseTRRows(rows) {
  const txs = [];
  if (!rows.length) return txs;
  let headerIdx = 0;
  for (let i = 0; i < Math.min(5, rows.length); i++) {
    const r = rows[i];
    if (r.some(c => /^datetime$/i.test(String(c || ""))) && r.some(c => /^type$/i.test(String(c || "")))) {
      headerIdx = i; break;
    }
  }
  const headers = rows[headerIdx].map(h => String(h || "").trim().toLowerCase());
  const col = {};
  headers.forEach((h, i) => { col[h] = i; });
  if (col.type === undefined) return txs;
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.every(c => !String(c || "").trim())) continue;
    const rawType = String(r[col.type] || "").trim().toUpperCase();
    if (!rawType) continue;
    const kind = TR_TYPE_MAP[rawType] || "OTHER";
    const date = parseDate(r[col.datetime]) || parseDate(r[col.date]);
    if (!date) continue;
    const sharesRaw = parseNumber(r[col.shares]);
    const priceVal = parseNumber(r[col.price]);
    let amountRawVal = parseNumber(r[col.amount]);
    // Fallback : si amount vide pour BUY/SELL (ex : vente du jour non settlée), reconstruire depuis shares × price
    if (amountRawVal === 0 && sharesRaw !== 0 && priceVal !== 0 && (rawType === "BUY" || rawType === "SELL")) {
      amountRawVal = -sharesRaw * priceVal;
    }
    txs.push({
      date, kind, rawType,
      shares: Math.abs(sharesRaw), sharesRaw,
      price: priceVal,
      amount: Math.abs(amountRawVal),
      amountRaw: amountRawVal,
      fee: Math.abs(parseNumber(r[col.fee])),
      tax: Math.abs(parseNumber(r[col.tax])),
      symbol: String(r[col.symbol] || "").trim(),
      name: String(r[col.name] || "").trim(),
      assetClass: String(r[col.asset_class] || "").trim(),
      category: String(r[col.category] || "").trim(),
      accountType: String(r[col.account_type] || "DEFAULT").trim(),
      currency: String(r[col.currency] || "EUR").trim() || "EUR",
      description: String(r[col.description] || "").trim(),
      txId: String(r[col.transaction_id] || `${date.toISOString()}_${rawType}_${Math.random()}`).trim(),
    });
  }
  return txs;
}

function parseWorkbookTR(wb) {
  const txs = [];
  for (const name of wb.SheetNames) {
    const sheet = wb.Sheets[name];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
    txs.push(...parseTRRows(json));
  }
  return txs;
}

// ═════ FIFO ENGINE ═════
const computeFIFO = (transactions) => {
  const tilgsByKey = {};
  for (const t of transactions) {
    if (t.kind === "TILG") {
      const k = t.symbol || t.name;
      (tilgsByKey[k] = tilgsByKey[k] || []).push(t);
    }
  }
  const byAsset = {};
  for (const t of transactions) {
    if (t.kind !== "BUY" && t.kind !== "SELL" && t.kind !== "WARRANT_EXERCISE") continue;
    const key = t.symbol || t.name || "UNKNOWN";
    (byAsset[key] = byAsset[key] || []).push(t);
  }
  const realizedTrades = [], openLots = [];
  for (const [assetKey, txs] of Object.entries(byAsset)) {
    const sorted = [...txs].sort((a, b) => a.date - b.date);
    let lots = [];
    for (const t of sorted) {
      if (t.kind === "BUY") {
        const totalCost = t.amount + t.fee + t.tax;
        const costPerShare = t.shares > 0 ? totalCost / t.shares : t.price;
        lots.push({ date: t.date, shares: t.shares, costPerShare, symbol: t.symbol, name: t.name, ref: t });
      } else {
        let proceedsTotal;
        if (t.kind === "WARRANT_EXERCISE") {
          const tilgs = tilgsByKey[t.symbol || t.name] || [];
          const m = tilgs.find(tg => Math.abs(tg.date - t.date) < 3 * 86400000);
          proceedsTotal = m ? m.amount : 0;
        } else proceedsTotal = t.amount - t.fee - t.tax;
        const pps = t.shares > 0 ? proceedsTotal / t.shares : 0;
        let rem = t.shares;
        while (rem > 1e-9 && lots.length) {
          const lot = lots[0];
          const used = Math.min(lot.shares, rem);
          const cost = used * lot.costPerShare;
          const proceeds = used * pps;
          const pnl = proceeds - cost;
          realizedTrades.push({
            symbol: t.symbol, isin: t.symbol, name: t.name || lot.name,
            buyDate: lot.date, sellDate: t.date, shares: used,
            buyPrice: lot.costPerShare, sellPrice: pps,
            cost, proceeds, pnl,
            pnlPct: cost > 0 ? (pnl / cost) * 100 : 0,
            holdingDays: Math.max(0, Math.round((t.date - lot.date) / 86400000)),
            kind: t.kind, assetClass: t.assetClass || "",
            accountType: t.accountType || "DEFAULT",
          });
          lot.shares -= used; rem -= used;
          if (lot.shares < 1e-9) lots.shift();
        }
      }
    }
    for (const l of lots) if (l.shares > 1e-9) openLots.push({ ...l, assetKey });
  }
  return { realizedTrades, openLots };
};

// ═════ AGRÉGATIONS ═════

// Consolide les trades FIFO issus d'une même vente (TR divise parfois 1 vente en N lots FIFO)
// Regroupe par (symbol + sellDate au seconde près) - chaque vente TR = 1 ligne consolidée
const consolidateTrades = (trades) => {
  const groups = {};
  for (const t of trades) {
    const key = (t.symbol || t.name) + "|" + (t.sellDate instanceof Date ? t.sellDate.toISOString().slice(0, 19) : String(t.sellDate).slice(0, 19));
    if (!groups[key]) {
      groups[key] = {
        symbol: t.symbol, isin: t.symbol, name: t.name,
        sellDate: t.sellDate, buyDate: t.buyDate, // on prend la date d'achat du premier lot (le + ancien FIFO)
        shares: 0, cost: 0, proceeds: 0, pnl: 0,
        kind: t.kind, assetClass: t.assetClass || "",
        accountType: t.accountType || "DEFAULT",
        lotCount: 0, firstBuyDate: t.buyDate, lastBuyDate: t.buyDate,
      };
    }
    const g = groups[key];
    g.shares += t.shares; g.cost += t.cost; g.proceeds += t.proceeds; g.pnl += t.pnl;
    g.lotCount++;
    if (t.buyDate < g.firstBuyDate) g.firstBuyDate = t.buyDate;
    if (t.buyDate > g.lastBuyDate) g.lastBuyDate = t.buyDate;
  }
  return Object.values(groups).map(g => {
    const holdingDays = Math.max(0, Math.round((g.sellDate - g.firstBuyDate) / 86400000));
    const pnlPct = g.cost > 0 ? (g.pnl / g.cost) * 100 : 0;
    // Rendement ANNUALISÉ : (1 + r)^(365/jours) - 1. Si détenu < 7j, non significatif (null)
    let annualizedPct = null;
    if (g.cost > 0 && holdingDays >= 7) {
      const r = g.pnl / g.cost;
      if (r > -1) {
        const ann = (Math.pow(1 + r, 365 / holdingDays) - 1) * 100;
        if (isFinite(ann) && Math.abs(ann) < 100000) annualizedPct = ann;
      }
    }
    return {
      ...g,
      buyDate: g.firstBuyDate, // date du 1er lot (le + ancien)
      buyPrice: g.shares > 0 ? g.cost / g.shares : 0,
      sellPrice: g.shares > 0 ? g.proceeds / g.shares : 0,
      pnlPct,
      annualizedPct,
      holdingDays,
      holdingDaysLast: Math.max(0, Math.round((g.sellDate - g.lastBuyDate) / 86400000)), // durée du lot le + récent
    };
  });
};

const aggregateBySector = (trades, lots, prices, sectorFn) => {
  const m = {};
  const e = (s) => m[s] = m[s] || { sector: s, realizedPnL: 0, realizedCount: 0, investedOpen: 0, currentValue: 0, unrealizedPnL: 0, wins: 0, losses: 0, totalWins: 0, totalLosses: 0, assets: new Set() };
  for (const t of trades) {
    const s = sectorFn(t); const x = e(s);
    x.realizedPnL += t.pnl; x.realizedCount++;
    x.assets.add(t.symbol || t.name);
    if (t.pnl > 0) { x.wins++; x.totalWins += t.pnl; } else if (t.pnl < 0) { x.losses++; x.totalLosses += Math.abs(t.pnl); }
  }
  for (const l of lots) {
    const s = sectorFn({ symbol: l.symbol, name: l.name, assetClass: l.ref?.assetClass });
    const x = e(s);
    const invested = l.shares * l.costPerShare;
    const px = prices[l.symbol || l.name];
    const val = px ? l.shares * px : invested;
    x.investedOpen += invested; x.currentValue += val; x.unrealizedPnL += (val - invested);
    x.assets.add(l.symbol || l.name);
  }
  return Object.values(m).map(x => ({
    ...x, assetCount: x.assets.size,
    totalPnL: x.realizedPnL + x.unrealizedPnL,
    returnPct: x.investedOpen > 0 ? (x.unrealizedPnL / x.investedOpen) * 100 : 0,
    winRate: (x.wins + x.losses) > 0 ? (x.wins / (x.wins + x.losses)) * 100 : 0,
    profitFactor: x.totalLosses > 0 ? x.totalWins / x.totalLosses : (x.totalWins > 0 ? Infinity : 0),
  })).sort((a, b) => b.totalPnL - a.totalPnL);
};

const aggregateByAsset = (trades, lots, prices) => {
  const m = {};
  const e = (k) => m[k] = m[k] || { key: k, name: "", symbol: "", isin: "", realizedPnL: 0, trades: 0, openShares: 0, invested: 0, currentValue: 0, unrealizedPnL: 0, wins: 0, losses: 0, firstDate: null, lastDate: null, assetClass: "" };
  for (const t of trades) {
    const k = t.symbol || t.name; const x = e(k);
    x.name = t.name || x.name; x.symbol = t.symbol || x.symbol; x.isin = t.symbol || x.isin;
    x.assetClass = t.assetClass || x.assetClass;
    x.realizedPnL += t.pnl; x.trades++;
    if (t.pnl > 0) x.wins++; else if (t.pnl < 0) x.losses++;
    if (!x.firstDate || t.buyDate < x.firstDate) x.firstDate = t.buyDate;
    if (!x.lastDate || t.sellDate > x.lastDate) x.lastDate = t.sellDate;
  }
  for (const l of lots) {
    const k = l.symbol || l.name; const x = e(k);
    x.name = l.name || x.name; x.symbol = l.symbol || x.symbol; x.isin = l.symbol || x.isin;
    const invested = l.shares * l.costPerShare;
    const px = prices[l.symbol || l.name];
    const val = px ? l.shares * px : invested;
    x.openShares += l.shares; x.invested += invested; x.currentValue += val; x.unrealizedPnL += (val - invested);
    if (!x.firstDate || l.date < x.firstDate) x.firstDate = l.date;
  }
  return Object.values(m).map(x => ({
    ...x, totalPnL: x.realizedPnL + x.unrealizedPnL,
    returnPct: x.invested > 0 ? (x.unrealizedPnL / x.invested) * 100 : 0,
    winRate: (x.wins + x.losses) > 0 ? (x.wins / (x.wins + x.losses)) * 100 : 0,
    avgCostPerShare: x.openShares > 0 ? x.invested / x.openShares : 0,
  }));
};

const computeStats = (trades, dividends, interests, savebacks, gifts, consolidatedTrades) => {
  const divNet = dividends.reduce((s, d) => s + d.net, 0);
  const intNet = interests.reduce((s, i) => s + i.net, 0);
  const saveTotal = savebacks.reduce((s, d) => s + d.amount, 0);
  const giftTotal = gifts.reduce((s, g) => s + g.amount, 0);
  const n = trades.length;
  // Pour best/worst on utilise les trades CONSOLIDÉS (1 vente réelle = 1 ligne)
  const tradesForStats = consolidatedTrades || trades;
  if (!n) return { count: 0, winRate: 0, avgWin: 0, avgLoss: 0, profitFactor: 0, totalPnL: 0, bestTrade: null, worstTrade: null, avgHolding: 0, totalDividends: divNet, totalInterests: intNet, totalSavebacks: saveTotal, totalGifts: giftTotal, maxDrawdown: 0, cumSeries: [], pnlDistribution: [], totalInvested: 0, totalProceeds: 0, largestWinStreak: 0, largestLossStreak: 0, winCount: 0, lossCount: 0, consolidatedCount: 0 };
  const wins = trades.filter(t => t.pnl > 0);
  const losses = trades.filter(t => t.pnl < 0);
  const totalWins = wins.reduce((s, t) => s + t.pnl, 0);
  const totalLosses = Math.abs(losses.reduce((s, t) => s + t.pnl, 0));
  // Série de win/loss basée sur les trades CONSOLIDÉS (sinon TSMC = 6 lots = fausse 6W d'affilée)
  const sortedConsolidated = [...tradesForStats].sort((a, b) => a.sellDate - b.sellDate);
  let cum = 0, peak = 0, maxDD = 0;
  const cumSeries = sortedConsolidated.map(t => { cum += t.pnl; if (cum > peak) peak = cum; if (peak - cum > maxDD) maxDD = peak - cum; return { date: fmtDate(t.sellDate), cum, pnl: t.pnl }; });
  let curW = 0, curL = 0, maxW = 0, maxL = 0;
  for (const t of sortedConsolidated) { if (t.pnl > 0) { curW++; curL = 0; if (curW > maxW) maxW = curW; } else if (t.pnl < 0) { curL++; curW = 0; if (curL > maxL) maxL = curL; } }
  const dist = {};
  // Distribution basée sur les trades CONSOLIDÉS (TSMC = 1 barre à +641€, pas 6 barres entre +55€ et +145€)
  for (const t of tradesForStats) { const b = Math.floor(t.pnl / 50) * 50; dist[b] = (dist[b] || 0) + 1; }
  const pnlDistribution = Object.entries(dist).map(([b, c]) => ({ bucket: +b, count: c })).sort((a, b) => a.bucket - b.bucket);
  // Win rate basé sur les trades CONSOLIDÉS aussi (plus pertinent)
  const winsConsolidated = tradesForStats.filter(t => t.pnl > 0);
  const lossesConsolidated = tradesForStats.filter(t => t.pnl < 0);
  return {
    count: tradesForStats.length, // nombre de VENTES réelles
    lotCount: n, // nombre de lots FIFO traités
    winRate: tradesForStats.length ? (winsConsolidated.length / tradesForStats.length) * 100 : 0,
    avgWin: winsConsolidated.length ? winsConsolidated.reduce((s, t) => s + t.pnl, 0) / winsConsolidated.length : 0,
    avgLoss: lossesConsolidated.length ? Math.abs(lossesConsolidated.reduce((s, t) => s + t.pnl, 0)) / lossesConsolidated.length : 0,
    profitFactor: totalLosses > 0 ? totalWins / totalLosses : (totalWins > 0 ? Infinity : 0),
    totalPnL: totalWins - totalLosses,
    bestTrade: sortedConsolidated.reduce((b, t) => (!b || t.pnl > b.pnl) ? t : b, null),
    worstTrade: sortedConsolidated.reduce((w, t) => (!w || t.pnl < w.pnl) ? t : w, null),
    avgHolding: tradesForStats.length ? tradesForStats.reduce((s, t) => s + t.holdingDays, 0) / tradesForStats.length : 0,
    totalDividends: divNet, totalInterests: intNet, totalSavebacks: saveTotal, totalGifts: giftTotal,
    maxDrawdown: maxDD, cumSeries, pnlDistribution,
    totalInvested: trades.reduce((s, t) => s + t.cost, 0),
    totalProceeds: trades.reduce((s, t) => s + t.proceeds, 0),
    largestWinStreak: maxW, largestLossStreak: maxL,
    winCount: winsConsolidated.length, lossCount: lossesConsolidated.length,
    consolidatedCount: tradesForStats.length,
  };
};

// ═══ UI PRIMITIVES ═══
const ACCENT_BAR = { slate: "bg-slate-500", cyan: "bg-cyan-500", blue: "bg-blue-500", emerald: "bg-emerald-500", rose: "bg-rose-500", amber: "bg-amber-500", purple: "bg-purple-500", orange: "bg-orange-500", violet: "bg-violet-500" };
const ACCENT_ICON = { slate: "text-slate-400", cyan: "text-cyan-400", blue: "text-blue-400", emerald: "text-emerald-400", rose: "text-rose-400", amber: "text-amber-400", purple: "text-purple-400", orange: "text-orange-400", violet: "text-violet-400" };
const PILL_CLASS = {
  slate: "bg-slate-500/20 text-slate-300 border-slate-500/30", cyan: "bg-cyan-500/20 text-cyan-300 border-cyan-500/30",
  emerald: "bg-emerald-500/20 text-emerald-300 border-emerald-500/30", rose: "bg-rose-500/20 text-rose-300 border-rose-500/30",
  amber: "bg-amber-500/20 text-amber-300 border-amber-500/30", blue: "bg-blue-500/20 text-blue-300 border-blue-500/30",
  purple: "bg-purple-500/20 text-purple-300 border-purple-500/30", orange: "bg-orange-500/20 text-orange-300 border-orange-500/30",
  violet: "bg-violet-500/20 text-violet-300 border-violet-500/30",
};

// ⚡ HOVER bien visible
const ROW_HOVER = "hover:bg-slate-700/40 transition-colors";

// Libellés et couleurs des types de transactions
const KIND_LABELS = {
  BUY: "Achat", SELL: "Vente", DIVIDEND: "Dividende", INTEREST: "Intérêt",
  DEPOSIT: "Dépôt", WITHDRAWAL: "Retrait", SAVEBACK: "Saveback", GIFT: "Cadeau",
  CARD: "Carte", WARRANT_EXERCISE: "Exercice warrant", TILG: "Remb. warrant",
  FREE_RECEIPT: "Transfert", TAX_OPT: "Optim. fiscale", CARD_REFUND: "Remb. carte", OTHER: "Autre",
};
const KIND_COLORS = {
  BUY: "blue", SELL: "emerald", DIVIDEND: "amber", INTEREST: "cyan",
  DEPOSIT: "violet", WITHDRAWAL: "orange", SAVEBACK: "purple", GIFT: "purple",
  CARD: "rose", WARRANT_EXERCISE: "amber", TILG: "amber",
  FREE_RECEIPT: "slate", TAX_OPT: "slate", CARD_REFUND: "emerald", OTHER: "slate",
};

const KPICard = ({ label, value, sub, accent = "slate", icon: Icon, subClass, big = false }) => (
  <div className="bg-slate-900/70 border border-slate-800 rounded-xl p-4 relative overflow-hidden hover:border-slate-700 transition-colors">
    <div className={`absolute top-0 left-0 h-full w-1 ${ACCENT_BAR[accent] || ACCENT_BAR.slate}`} />
    <div className="flex items-start justify-between gap-2">
      <div className="min-w-0 flex-1">
        <div className="text-[10px] uppercase tracking-widest text-slate-500 font-medium mb-1">{label}</div>
        <div className={`font-mono font-semibold text-slate-100 tabular-nums break-words ${big ? "text-2xl md:text-3xl" : "text-lg md:text-xl"}`}>{value}</div>
        {sub !== undefined && <div className={`text-xs mt-1 font-mono tabular-nums ${subClass || "text-slate-400"}`}>{sub}</div>}
      </div>
      {Icon && <Icon className={`w-4 h-4 ${ACCENT_ICON[accent] || ACCENT_ICON.slate} shrink-0`} />}
    </div>
  </div>
);

const SectionTitle = ({ children, icon: Icon, subtitle, actions }) => (
  <div className="flex items-center justify-between mb-4 gap-2">
    <div className="flex items-center gap-2 min-w-0">
      {Icon && <Icon className="w-4 h-4 text-cyan-400 shrink-0" />}
      <div className="min-w-0">
        <h2 className="text-sm font-semibold uppercase tracking-widest text-slate-200 truncate">{children}</h2>
        {subtitle && <div className="text-[11px] text-slate-500 mt-0.5">{subtitle}</div>}
      </div>
    </div>
    {actions}
  </div>
);

const Pill = ({ children, color = "slate" }) => (
  <span className={`inline-block px-2 py-0.5 text-[10px] rounded font-medium border ${PILL_CLASS[color] || PILL_CLASS.slate}`}>{children}</span>
);

const EmptyState = ({ title, subtitle, icon: Icon = Database }) => (
  <div className="flex flex-col items-center justify-center py-12 text-center">
    <Icon className="w-10 h-10 text-slate-700 mb-3" />
    <div className="text-slate-400 font-semibold mb-1">{title}</div>
    {subtitle && <div className="text-slate-500 text-sm max-w-md">{subtitle}</div>}
  </div>
);

const Card = ({ children, className = "" }) => (
  <div className={`bg-slate-900/70 border border-slate-800 rounded-xl p-5 ${className}`}>{children}</div>
);

const ChartTooltip = ({ active, payload, label, formatter = fmtEUR, labelFormatter }) => {
  if (!active || !payload?.length) return null;
  return (
    <div className="bg-slate-950 border border-slate-700 rounded-lg p-2.5 shadow-xl">
      {label !== undefined && <div className="text-[10px] text-slate-500 uppercase tracking-widest mb-1">{labelFormatter ? labelFormatter(label) : label}</div>}
      {payload.map((e, i) => (
        <div key={i} className="text-xs font-mono flex items-center gap-2">
          <span className="w-2 h-2 rounded-full" style={{ background: e.color || e.fill }} />
          <span className="text-slate-400">{e.name}:</span>
          <span className="text-slate-100 font-semibold">{formatter(e.value)}</span>
        </div>
      ))}
    </div>
  );
};

// ═══ APP ═══
export default function App() {
  const [transactions, setTransactions] = useState([]);
  const [sectorOverrides, setSectorOverrides] = useState({});
  const [currentPrices, setCurrentPrices] = useState({});
  const [hideBalances, setHideBalances] = useState(false);
  const [activeTab, setActiveTab] = useState("dashboard");
  const [importMsg, setImportMsg] = useState(null);
  const [loaded, setLoaded] = useState(false);

  useEffect(() => {
    (async () => {
      try {
        const [tx, so, cp, hb] = await Promise.all([
          storage.get(STORAGE_KEYS.TRANSACTIONS), storage.get(STORAGE_KEYS.SECTOR_OVERRIDES),
          storage.get(STORAGE_KEYS.CURRENT_PRICES), storage.get(STORAGE_KEYS.HIDE_BALANCES),
        ]);
        if (tx?.value) setTransactions(JSON.parse(tx.value).map(t => ({ ...t, date: new Date(t.date) })));
        if (so?.value) setSectorOverrides(JSON.parse(so.value));
        if (cp?.value) setCurrentPrices(JSON.parse(cp.value));
        if (hb?.value) setHideBalances(hb.value === "true");
      } catch (e) {}
      setLoaded(true);
    })();
  }, []);

  const persistTx = useCallback(async (data) => { try { await storage.set(STORAGE_KEYS.TRANSACTIONS, JSON.stringify(data.map(t => ({ ...t, date: t.date.toISOString() })))); } catch (e) {} }, []);
  const persistSectors = useCallback(async (d) => { try { await storage.set(STORAGE_KEYS.SECTOR_OVERRIDES, JSON.stringify(d)); } catch (e) {} }, []);
  const persistPrices = useCallback(async (d) => { try { await storage.set(STORAGE_KEYS.CURRENT_PRICES, JSON.stringify(d)); } catch (e) {} }, []);

  const handleFileImport = async (files) => {
    setImportMsg({ type: "loading", text: "Analyse des fichiers…" });
    const allNew = [], report = [];
    for (const file of files) {
      try {
        const isCSV = file.name.toLowerCase().endsWith(".csv");
        let parsed = [];
        if (isCSV) {
          const text = await file.text();
          parsed = parseTRRows(parseCSVText(text));
        } else {
          const buf = await file.arrayBuffer();
          parsed = parseWorkbookTR(XLSX.read(buf, { type: "array", cellDates: false }));
        }
        allNew.push(...parsed);
        report.push(`${file.name}: ${parsed.length} lignes`);
      } catch (e) { report.push(`${file.name}: ERREUR (${e.message})`); }
    }
    const existing = new Set(transactions.map(t => t.txId));
    const fresh = allNew.filter(t => !existing.has(t.txId) && (existing.add(t.txId), true));
    const merged = [...transactions, ...fresh].sort((a, b) => a.date - b.date);
    setTransactions(merged); persistTx(merged);
    setImportMsg({ type: "success", text: `✓ ${fresh.length} nouvelle(s) · ${allNew.length - fresh.length} doublon(s) ignoré(s)`, detail: report.join(" · ") });
    setTimeout(() => setImportMsg(null), 9000);
  };

  const sectorFn = useCallback((t) => detectSector(t, sectorOverrides), [sectorOverrides]);

  const derived = useMemo(() => {
    const buysSells = transactions.filter(t => t.kind === "BUY" || t.kind === "SELL" || t.kind === "WARRANT_EXERCISE" || t.kind === "TILG");
    const dividends = transactions.filter(t => t.kind === "DIVIDEND").map(t => ({ date: t.date, gross: t.amount, tax: t.tax, net: t.amount - t.tax, symbol: t.symbol, name: t.name, accountType: t.accountType || "DEFAULT" }));
    const interests = transactions.filter(t => t.kind === "INTEREST").map(t => ({ date: t.date, gross: t.amount, tax: t.tax, net: t.amount - t.tax, accountType: t.accountType || "DEFAULT" }));
    const savebacks = transactions.filter(t => t.kind === "SAVEBACK").map(t => ({ date: t.date, amount: t.amount, name: t.name }));
    const gifts = transactions.filter(t => t.kind === "GIFT").map(t => ({ date: t.date, amount: t.amount, name: t.name }));
    const deposits = transactions.filter(t => t.kind === "DEPOSIT").map(t => ({ date: t.date, amount: t.amount }));
    const withdrawals = transactions.filter(t => t.kind === "WITHDRAWAL").map(t => ({ date: t.date, amount: t.amount }));
    const cardTx = transactions.filter(t => t.kind === "CARD").map(t => ({ date: t.date, amount: t.amount, amountRaw: t.amountRaw, name: t.name }));

    const { realizedTrades, openLots } = computeFIFO(buysSells);
    const consolidatedTrades = consolidateTrades(realizedTrades); // 1 ligne par vente réelle (pas par lot FIFO)
    const stats = computeStats(realizedTrades, dividends, interests, savebacks, gifts, consolidatedTrades);
    const bySector = aggregateBySector(realizedTrades, openLots, currentPrices, sectorFn);
    const byAsset = aggregateByAsset(realizedTrades, openLots, currentPrices);
    const openInvested = openLots.reduce((s, l) => s + l.shares * l.costPerShare, 0);
    const openValue = openLots.reduce((s, l) => { const px = currentPrices[l.symbol || l.name]; return s + (px ? l.shares * px : l.shares * l.costPerShare); }, 0);
    const unrealizedPnL = openValue - openInvested;
    const totalDeposits = deposits.reduce((s, d) => s + d.amount, 0);
    const totalWithdrawals = withdrawals.reduce((s, d) => s + d.amount, 0);
    const totalCardSpend = cardTx.reduce((s, d) => s - d.amountRaw, 0); // amountRaw < 0 pour dépense → on somme positivement la dépense NETTE

    // ═════ DÉTECTION ORDRES PLANIFIÉS vs SPOT + ESTIMATION FRAIS ═════
    // TR : plan d'épargne = 0€ frais. Ordre marché = 1€. Mais le CSV met 0€ partout (frais agrégés ailleurs).
    // Heuristique : un ordre est "planifié" si montant ROND (50, 100, 25, 200…) sur un actif récurrent ET sans frais déclaré.
    // Sinon c'est spot → on estime 1€ par ordre (BUY ou SELL).
    const FEE_PER_SPOT_ORDER = 1.0;
    const isLikelyPlannedAmount = (amt) => {
      const a = Math.abs(amt);
      if (a < 10 || a > 500) return false;
      // Multiples de 5€ "ronds" : 25, 50, 75, 100, 150, 200, 250, 300, 500…
      const rounded5 = Math.round(a / 5) * 5;
      return Math.abs(a - rounded5) < 0.01;
    };
    // Compte des achats par actif : si >= 4 BUY au même montant rond → c'est un plan d'épargne
    const buyAssetCounts = {};
    for (const t of transactions) {
      if (t.kind !== "BUY") continue;
      const k = (t.symbol || t.name) + "|" + Math.round(t.amount);
      buyAssetCounts[k] = (buyAssetCounts[k] || 0) + 1;
    }
    const annotateOrder = (t) => {
      const isBuy = t.kind === "BUY", isSell = t.kind === "SELL";
      if (!isBuy && !isSell) return null;
      const declaredFee = t.fee || 0;
      let isPlanned = false;
      if (isBuy && declaredFee === 0 && isLikelyPlannedAmount(t.amount)) {
        const k = (t.symbol || t.name) + "|" + Math.round(t.amount);
        if ((buyAssetCounts[k] || 0) >= 4) isPlanned = true;
      }
      // Une vente n'est jamais "planifiée" → toujours frais 1€
      const estimatedFee = (declaredFee > 0) ? declaredFee : (isPlanned ? 0 : FEE_PER_SPOT_ORDER);
      return { isPlanned, declaredFee, estimatedFee, isSpot: !isPlanned };
    };
    const orderAnnotations = transactions.map(annotateOrder);
    let totalEstimatedFees = 0, totalDeclaredFees = 0, plannedCount = 0, spotBuyCount = 0, spotSellCount = 0;
    transactions.forEach((t, i) => {
      const a = orderAnnotations[i];
      if (!a) return;
      totalEstimatedFees += a.estimatedFee;
      totalDeclaredFees += a.declaredFee;
      if (a.isPlanned) plannedCount++;
      else if (t.kind === "BUY") spotBuyCount++;
      else if (t.kind === "SELL") spotSellCount++;
    });
    const feesEstimation = {
      totalEstimated: totalEstimatedFees,
      totalDeclared: totalDeclaredFees,
      hiddenFees: Math.max(0, totalEstimatedFees - totalDeclaredFees),
      plannedCount, spotBuyCount, spotSellCount,
      avgPerOrder: (spotBuyCount + spotSellCount) > 0 ? FEE_PER_SPOT_ORDER : 0,
    };

    // ═════ DONNÉES FINANCE PERSO : MENSUEL ═════
    // Pour chaque mois : dépenses CB, P&L réalisée, dividendes, intérêts, savebacks → ratio couverture
    const monthlyMap = {};
    const monthKey = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    const ensureMonth = (k) => monthlyMap[k] = monthlyMap[k] || { month: k, cardSpend: 0, cardRefund: 0, cardNet: 0, realizedPnL: 0, dividends: 0, interests: 0, savebacks: 0, gifts: 0, deposits: 0, gains: 0, coverage: 0, surplus: 0 };
    for (const t of cardTx) {
      const m = ensureMonth(monthKey(t.date));
      if (t.amountRaw < 0) m.cardSpend += Math.abs(t.amountRaw);
      else m.cardRefund += t.amountRaw;
    }
    for (const rt of realizedTrades) ensureMonth(monthKey(rt.sellDate)).realizedPnL += rt.pnl;
    for (const dv of dividends) ensureMonth(monthKey(dv.date)).dividends += dv.net;
    for (const it of interests) ensureMonth(monthKey(it.date)).interests += it.net;
    for (const sb of savebacks) ensureMonth(monthKey(sb.date)).savebacks += sb.amount;
    for (const g of gifts) ensureMonth(monthKey(g.date)).gifts += g.amount;
    for (const d of deposits) ensureMonth(monthKey(d.date)).deposits += d.amount;
    const monthlyArr = Object.values(monthlyMap).sort((a, b) => a.month.localeCompare(b.month)).map(m => {
      const net = m.cardSpend - m.cardRefund;
      const gains = m.realizedPnL + m.dividends + m.interests + m.savebacks + m.gifts;
      return {
        ...m,
        cardNet: net,
        gains,
        coverage: net > 0 ? (gains / net) * 100 : (gains > 0 ? 100 : 0),
        surplus: gains - net,
      };
    });
    // Surplus cumulé mois après mois (l'épargne nette accumulée grâce aux gains)
    {
      let cum = 0;
      for (const m of monthlyArr) { cum += m.surplus; m.cumSurplus = cum; }
    }

    // ═════ DONNÉES HEBDOMADAIRES (lundi → dimanche, norme ISO) ═════
    const weekKeyOf = (dt) => {
      const d = new Date(dt);
      const day = (d.getDay() + 6) % 7; // 0 = lundi
      const monday = new Date(d); monday.setDate(d.getDate() - day); monday.setHours(0, 0, 0, 0);
      return `${monday.getFullYear()}-${String(monday.getMonth() + 1).padStart(2, "0")}-${String(monday.getDate()).padStart(2, "0")}`;
    };
    const weeklyMap = {};
    const ensureWeek = (k) => weeklyMap[k] = weeklyMap[k] || { week: k, cardSpend: 0, cardRefund: 0, realizedPnL: 0, dividends: 0, interests: 0, savebacks: 0, gifts: 0 };
    for (const t of cardTx) { const w = ensureWeek(weekKeyOf(t.date)); if (t.amountRaw < 0) w.cardSpend += Math.abs(t.amountRaw); else w.cardRefund += t.amountRaw; }
    for (const rt of realizedTrades) ensureWeek(weekKeyOf(rt.sellDate)).realizedPnL += rt.pnl;
    for (const dv of dividends) ensureWeek(weekKeyOf(dv.date)).dividends += dv.net;
    for (const it of interests) ensureWeek(weekKeyOf(it.date)).interests += it.net;
    for (const sb of savebacks) ensureWeek(weekKeyOf(sb.date)).savebacks += sb.amount;
    for (const g of gifts) ensureWeek(weekKeyOf(g.date)).gifts += g.amount;
    const weeklyArr = Object.values(weeklyMap).sort((a, b) => a.week.localeCompare(b.week)).map(w => {
      const net = w.cardSpend - w.cardRefund;
      const gains = w.realizedPnL + w.dividends + w.interests + w.savebacks + w.gifts;
      return { ...w, cardNet: net, gains, coverage: net > 0 ? (gains / net) * 100 : (gains > 0 ? 100 : 0), surplus: gains - net };
    });
    { let cum = 0; for (const w of weeklyArr) { cum += w.surplus; w.cumSurplus = cum; } }

    const totalRealizedPnL = realizedTrades.reduce((s, t) => s + t.pnl, 0);
    const totalGains = totalRealizedPnL + stats.totalDividends + stats.totalInterests + stats.totalSavebacks + stats.totalGifts;
    const cardSpendNet = cardTx.reduce((s, t) => s + (t.amountRaw < 0 ? Math.abs(t.amountRaw) : -t.amountRaw), 0);
    const lifetimeCoverage = cardSpendNet > 0 ? (totalGains / cardSpendNet) * 100 : 0;
    const lifetimeSurplus = totalGains - cardSpendNet;
    // Runway : combien de mois de dépenses ton surplus accumulé peut couvrir
    const monthsWithSpend = monthlyArr.filter(m => m.cardNet > 0).length;
    const avgMonthlySpendAllTime = monthsWithSpend > 0 ? cardSpendNet / monthsWithSpend : 0;
    const runwayMonths = avgMonthlySpendAllTime > 0 ? lifetimeSurplus / avgMonthlySpendAllTime : 0;

    // Catégorisation simple des dépenses CB par mots-clés (pour l'analyse)
    const categorizeMerchant = (name) => {
      const n = (name || "").toLowerCase();
      if (/uber|bolt|sncf|trainline|flixbus|blablacar|tgv|ouigo|ratp|jcdecaux|free.flotte|lime|tier|dott/i.test(n)) return "Transport";
      if (/airbnb|hotel|booking|hostel|expedia/i.test(n)) return "Hébergement";
      if (/restaurant|mcdo|burger|kfc|sushi|pizza|food|deliveroo|uber.eats|just.eat|frichti|too.good/i.test(n)) return "Restaurants";
      if (/carrefour|lidl|monoprix|leclerc|auchan|intermarche|casino|franprix|biocoop|naturalia|picard/i.test(n)) return "Courses";
      if (/amazon|cdiscount|fnac|darty|zalando|asos|h&m|zara|uniqlo|decathlon/i.test(n)) return "Shopping";
      if (/netflix|spotify|deezer|disney|prime.video|youtube|apple.tv|canal/i.test(n)) return "Abonnements";
      if (/orange|sfr|bouygues|free|sosh|iliad/i.test(n)) return "Télécom";
      if (/edf|enedis|engie|veolia/i.test(n)) return "Énergie";
      if (/pharmacie|doctolib|medecin|hospital|sante/i.test(n)) return "Santé";
      if (/cinema|ugc|pathe|gaumont|theatre|concert|fitnesspark|basicfit|ticketmaster/i.test(n)) return "Loisirs";
      if (/atm|withdraw|retrait/i.test(n)) return "Retrait DAB";
      return "Autres";
    };
    const cardByCategory = {};
    for (const t of cardTx) {
      if (t.amountRaw >= 0) continue; // skip remboursements
      const cat = categorizeMerchant(t.name);
      if (!cardByCategory[cat]) cardByCategory[cat] = { category: cat, total: 0, count: 0 };
      cardByCategory[cat].total += Math.abs(t.amountRaw);
      cardByCategory[cat].count++;
    }
    const cardCategories = Object.values(cardByCategory).sort((a, b) => b.total - a.total);

    // ═════ FISCALITÉ : taxe CTO 31,4% sur P&L réalisée + dividendes + intérêts (PEA exonéré) ═════
    const TAX_RATE_CTO = 0.314; // 31,4% PFU/flat tax France
    const realizedPnLCTO = realizedTrades.filter(t => t.accountType !== "PEA").reduce((s, t) => s + t.pnl, 0);
    const realizedPnLPEA = realizedTrades.filter(t => t.accountType === "PEA").reduce((s, t) => s + t.pnl, 0);
    const divsCTO = dividends.filter(d => d.accountType !== "PEA").reduce((s, d) => s + d.net, 0);
    const divsPEA = dividends.filter(d => d.accountType === "PEA").reduce((s, d) => s + d.net, 0);
    const intsCTO = interests.filter(i => i.accountType !== "PEA").reduce((s, i) => s + i.net, 0);
    const intsPEA = interests.filter(i => i.accountType === "PEA").reduce((s, i) => s + i.net, 0);
    // L'imposition ne s'applique qu'aux gains POSITIFS sur le CTO (pas de tax sur des pertes ; report des moins-values pas modélisé ici - approximation)
    const taxableGainsCTO = Math.max(0, realizedPnLCTO) + Math.max(0, divsCTO) + Math.max(0, intsCTO);
    const taxCTOEstimated = taxableGainsCTO * TAX_RATE_CTO;
    const taxation = {
      rate: TAX_RATE_CTO,
      taxableGainsCTO,
      taxCTOEstimated,
      realizedPnLCTO, realizedPnLPEA,
      divsCTO, divsPEA,
      intsCTO, intsPEA,
      hasPEA: realizedPnLPEA !== 0 || divsPEA !== 0 || intsPEA !== 0 || transactions.some(t => t.accountType === "PEA"),
    };

    // ═════ COMPORTEMENT D'INVESTISSEMENT ═════
    const buys = transactions.filter(t => t.kind === "BUY");
    const sells = transactions.filter(t => t.kind === "SELL");
    const allActivity = [...buys, ...sells, ...dividends.map(d => ({ ...d, kind: "DIVIDEND" }))];
    // Jours de la semaine préférés pour acheter (0=dim, 1=lun...)
    const DAY_NAMES = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"];
    const buysByDay = [0, 0, 0, 0, 0, 0, 0];
    for (const b of buys) buysByDay[b.date.getDay()]++;
    const favoriteDayIdx = buysByDay.indexOf(Math.max(...buysByDay));
    // Mois les plus actifs
    const MONTH_NAMES = ["Jan", "Fév", "Mar", "Avr", "Mai", "Juin", "Juil", "Août", "Sep", "Oct", "Nov", "Déc"];
    const buysByMonth = Array(12).fill(0);
    for (const b of buys) buysByMonth[b.date.getMonth()]++;
    const buysByDayData = DAY_NAMES.map((name, i) => ({ day: name.slice(0, 3), count: buysByDay[i], full: name }));
    const buysByMonthData = MONTH_NAMES.map((name, i) => ({ month: name, count: buysByMonth[i] }));
    // Ventes : jours et mois
    const sellsByDay = [0, 0, 0, 0, 0, 0, 0];
    for (const s of sells) sellsByDay[s.date.getDay()]++;
    const favoriteSellDayIdx = sellsByDay.indexOf(Math.max(...sellsByDay));
    const sellsByMonth = Array(12).fill(0);
    for (const s of sells) sellsByMonth[s.date.getMonth()]++;
    const sellsByDayData = DAY_NAMES.map((name, i) => ({ day: name.slice(0, 3), count: sellsByDay[i], full: name }));
    const sellsByMonthData = MONTH_NAMES.map((name, i) => ({ month: name, count: sellsByMonth[i] }));
    // Rythme : nombre de jours entre le 1er et dernier achat / nb achats
    const sortedBuys = [...buys].sort((a, b) => a.date - b.date);
    const daySpan = sortedBuys.length > 1 ? Math.max(1, Math.round((sortedBuys[sortedBuys.length - 1].date - sortedBuys[0].date) / 86400000)) : 0;
    const buysPerWeek = daySpan > 0 ? (buys.length / (daySpan / 7)) : 0;
    const avgBuyAmount = buys.length > 0 ? buys.reduce((s, t) => s + t.amount, 0) / buys.length : 0;
    const avgSellAmount = sells.length > 0 ? sells.reduce((s, t) => s + t.amount, 0) / sells.length : 0;
    // Rythme de vente
    const sortedSells = [...sells].sort((a, b) => a.date - b.date);
    const sellDaySpan = sortedSells.length > 1 ? Math.max(1, Math.round((sortedSells[sortedSells.length - 1].date - sortedSells[0].date) / 86400000)) : 0;
    const sellsPerWeek = sellDaySpan > 0 ? (sells.length / (sellDaySpan / 7)) : 0;
    const totalSellAmount = sells.reduce((s, t) => s + t.amount, 0);
    const capitalSoldPerWeek = sellDaySpan > 0 ? totalSellAmount / (sellDaySpan / 7) : 0;
    // Capital moyen investi par semaine
    const totalBuyAmount = buys.reduce((s, t) => s + t.amount, 0);
    const weeksSpan = daySpan > 0 ? daySpan / 7 : 1;
    const capitalPerWeek = totalBuyAmount / weeksSpan;
    // Ratio achat/vente (combien tu gardes vs vends)
    const buyToSellRatio = sells.length > 0 ? buys.length / sells.length : buys.length;
    // Première et dernière transaction
    const allDates = transactions.map(t => t.date).sort((a, b) => a - b);
    const firstTx = allDates[0];
    const lastTx = allDates[allDates.length - 1];
    const investingDays = firstTx && lastTx ? Math.max(1, Math.round((lastTx - firstTx) / 86400000)) : 0;
    const behaviorStats = {
      buysByDayData, buysByMonthData, favoriteDay: DAY_NAMES[favoriteDayIdx],
      sellsByDayData, sellsByMonthData, favoriteSellDay: DAY_NAMES[favoriteSellDayIdx],
      buysPerWeek, sellsPerWeek, avgBuyAmount, avgSellAmount, capitalPerWeek, capitalSoldPerWeek, buyToSellRatio,
      totalBuys: buys.length, totalSells: sells.length, daySpan, investingDays,
      firstTx, lastTx,
      mostActiveMonth: MONTH_NAMES[buysByMonth.indexOf(Math.max(...buysByMonth))],
      mostActiveSellMonth: MONTH_NAMES[sellsByMonth.indexOf(Math.max(...sellsByMonth))],
    };

    // ═════ XIRR : Taux de Rendement Interne annualisé (méthode Newton-Raphson) ═════
    // Flux : dépôts (−), retraits (+), valeur finale du portefeuille (+) aujourd'hui
    const computeXIRR = () => {
      const flows = [];
      for (const dp of deposits) flows.push({ date: dp.date, amount: -dp.amount });   // argent qui ENTRE chez TR = sortie de ta poche
      for (const wd of withdrawals) flows.push({ date: wd.date, amount: wd.amount }); // retrait = retour dans ta poche
      // Dépenses CB = sorties de valeur (équivalent retraits) ; remboursements = entrées
      for (const ct of cardTx) flows.push({ date: ct.date, amount: -ct.amountRaw > 0 ? Math.abs(ct.amountRaw) : -Math.abs(ct.amountRaw) });
      // Valeur terminale : valeur du portefeuille + cash estimé (dépôts - achats + ventes + gains...) — simplifié : valeur positions ouvertes
      // Approximation prudente : on ne valorise que les positions ouvertes (cash non modélisé)
      if (flows.length < 2 || openValue <= 0) return null;
      flows.push({ date: new Date(), amount: openValue });
      flows.sort((a, b) => a.date - b.date);
      const t0 = flows[0].date;
      const yearsFrom = (d) => (d - t0) / (365.25 * 86400000);
      const npv = (rate) => flows.reduce((s, f) => s + f.amount / Math.pow(1 + rate, yearsFrom(f.date)), 0);
      const dnpv = (rate) => flows.reduce((s, f) => { const y = yearsFrom(f.date); return s - y * f.amount / Math.pow(1 + rate, y + 1); }, 0);
      let rate = 0.1;
      for (let i = 0; i < 100; i++) {
        const v = npv(rate), dv = dnpv(rate);
        if (Math.abs(dv) < 1e-10) break;
        const next = rate - v / dv;
        if (!isFinite(next) || next <= -0.999) { rate = (rate - 0.999) / 2; continue; }
        if (Math.abs(next - rate) < 1e-8) { rate = next; break; }
        rate = next;
      }
      if (!isFinite(rate) || rate <= -0.999 || rate > 100) return null;
      return rate * 100;
    };
    const xirr = computeXIRR();

    // ═════ GAINS QUOTIDIENS (pour heatmap calendrier) ═════
    const dailyGains = {};
    const dayKey = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
    for (const rt of realizedTrades) { const k = dayKey(rt.sellDate); dailyGains[k] = (dailyGains[k] || 0) + rt.pnl; }
    for (const dv of dividends) { const k = dayKey(dv.date); dailyGains[k] = (dailyGains[k] || 0) + dv.net; }
    for (const it of interests) { const k = dayKey(it.date); dailyGains[k] = (dailyGains[k] || 0) + it.net; }
    for (const sb of savebacks) { const k = dayKey(sb.date); dailyGains[k] = (dailyGains[k] || 0) + sb.amount; }

    // ═════ RECORDS & MILESTONES ═════
    const dailyEntries = Object.entries(dailyGains).map(([k, v]) => ({ day: k, gain: v }));
    const bestDay = dailyEntries.reduce((b, e) => (!b || e.gain > b.gain) ? e : b, null);
    const worstDay = dailyEntries.reduce((w, e) => (!w || e.gain < w.gain) ? e : w, null);
    // Semaines
    const weeklyGains = {};
    const weekKey = (dstr) => { const d = new Date(dstr); const day = d.getDay() || 7; const monday = new Date(d); monday.setDate(d.getDate() - day + 1); return dayKey(monday); };
    for (const e of dailyEntries) { const wk = weekKey(e.day); weeklyGains[wk] = (weeklyGains[wk] || 0) + e.gain; }
    const weeklyEntries = Object.entries(weeklyGains).map(([k, v]) => ({ week: k, gain: v }));
    const bestWeek = weeklyEntries.reduce((b, e) => (!b || e.gain > b.gain) ? e : b, null);
    // Jours de gains consécutifs
    const sortedDays = dailyEntries.sort((a, b) => a.day.localeCompare(b.day));
    let curStreak = 0, maxPosStreak = 0;
    for (const e of sortedDays) { if (e.gain > 0) { curStreak++; if (curStreak > maxPosStreak) maxPosStreak = curStreak; } else curStreak = 0; }
    // Mois consécutifs couverts (gains >= dépenses CB)
    let curCovStreak = 0, maxCovStreak = 0;
    for (const m of monthlyArr) { if (m.cardNet > 0 && m.gains >= m.cardNet) { curCovStreak++; if (curCovStreak > maxCovStreak) maxCovStreak = curCovStreak; } else if (m.cardNet > 0) curCovStreak = 0; }
    const records = {
      bestDay, worstDay, bestWeek, maxPosStreak, maxCovStreak,
      profitableDays: dailyEntries.filter(e => e.gain > 0).length,
      totalActiveDays: dailyEntries.length,
      biggestTrade: consolidatedTrades.reduce((b, t) => (!b || t.pnl > b.pnl) ? t : b, null),
      avgDailyGain: dailyEntries.length > 0 ? dailyEntries.reduce((s, e) => s + e.gain, 0) / dailyEntries.length : 0,
    };

    // ═════ SCORE DE SANTÉ DU PORTEFEUILLE (0-100) ═════
    // Composite : diversification (25), win rate (25), couverture frais (15), régularité (20), profit factor (15)
    const healthComponents = [];
    {
      // Diversification : nb secteurs avec positions ouvertes (max 5+)
      const activeSectors = new Set(openLots.map(l => sectorFn(l.ref || l))).size;
      const divScore = Math.min(25, (activeSectors / 5) * 25);
      healthComponents.push({ label: "Diversification", score: divScore, max: 25, detail: `${activeSectors} secteur${activeSectors > 1 ? "s" : ""} actif${activeSectors > 1 ? "s" : ""}` });
      // Win rate (50% = 12.5pts, 70%+ = 25)
      const wrScore = stats.count > 0 ? Math.min(25, (stats.winRate / 70) * 25) : 0;
      healthComponents.push({ label: "Win rate", score: wrScore, max: 25, detail: stats.count > 0 ? `${stats.winRate.toFixed(0)}% de ventes gagnantes` : "Pas de ventes" });
      // Frais maîtrisés : ratio frais estimés / total investi (moins de 0.5% = max)
      const feeRatio = (stats.totalInvested + openInvested) > 0 ? totalEstimatedFees / (stats.totalInvested + openInvested) : 0;
      const feeScore = Math.max(0, Math.min(15, 15 * (1 - feeRatio / 0.005)));
      healthComponents.push({ label: "Frais maîtrisés", score: feeScore, max: 15, detail: `${(feeRatio * 100).toFixed(2)}% du capital en frais` });
      // Régularité : plans d'épargne vs spot (plus de planifié = mieux)
      const planRatio = (plannedCount + spotBuyCount) > 0 ? plannedCount / (plannedCount + spotBuyCount) : 0;
      const regScore = planRatio * 20;
      healthComponents.push({ label: "Régularité (DCA)", score: regScore, max: 20, detail: `${(planRatio * 100).toFixed(0)}% d'ordres planifiés` });
      // Profit factor (1.5+ = max)
      const pfScore = stats.profitFactor === Infinity ? 15 : Math.min(15, (stats.profitFactor / 1.5) * 15);
      healthComponents.push({ label: "Profit factor", score: pfScore, max: 15, detail: stats.profitFactor === Infinity ? "∞" : stats.profitFactor.toFixed(2) });
    }
    const healthScore = Math.round(healthComponents.reduce((s, c) => s + c.score, 0));

    // ═════ COMPARAISON MOIS COURANT vs MOIS PRÉCÉDENT ═════
    const nowD = new Date();
    const curKey = `${nowD.getFullYear()}-${String(nowD.getMonth() + 1).padStart(2, "0")}`;
    const prevD = new Date(nowD.getFullYear(), nowD.getMonth() - 1, 1);
    const prevKey = `${prevD.getFullYear()}-${String(prevD.getMonth() + 1).padStart(2, "0")}`;
    const curMonth = monthlyArr.find(m => m.month === curKey) || null;
    const prevMonth = monthlyArr.find(m => m.month === prevKey) || null;
    const monthComparison = { current: curMonth, previous: prevMonth,
      deltaGains: curMonth && prevMonth ? curMonth.gains - prevMonth.gains : null,
      deltaSpend: curMonth && prevMonth ? curMonth.cardNet - prevMonth.cardNet : null,
      deltaCoverage: curMonth && prevMonth ? curMonth.coverage - prevMonth.coverage : null,
    };

    // Courbes cumulées
    const events = [];
    for (const t of transactions) {
      if (t.kind === "BUY") events.push({ date: t.date, invested: t.amount + t.fee + t.tax, realized: 0, divInt: 0 });
      else if (t.kind === "DIVIDEND" || t.kind === "INTEREST") events.push({ date: t.date, invested: 0, realized: 0, divInt: t.amount - t.tax });
    }
    for (const rt of realizedTrades) events.push({ date: rt.sellDate, invested: -rt.cost, realized: rt.pnl, divInt: 0 });
    events.sort((a, b) => a.date - b.date);
    let ci = 0, cr = 0, cd = 0;
    const curveSeries = events.map(e => { ci += e.invested; cr += e.realized; cd += e.divInt; return { date: fmtDate(e.date), invested: ci, realized: cr, divInt: cd }; });

    return { transactions, orderAnnotations, dividends, interests, savebacks, gifts, deposits, withdrawals, cardTx, realizedTrades, consolidatedTrades, openLots, stats, bySector, byAsset, openInvested, openValue, unrealizedPnL, totalDeposits, totalWithdrawals, totalCardSpend, curveSeries, feesEstimation, monthlyArr, weeklyArr, lifetimeCoverage, lifetimeSurplus, cardCategories, cardSpendNet, totalGains, taxation, behaviorStats, xirr, dailyGains, records, healthScore, healthComponents, monthComparison, runwayMonths, avgMonthlySpendAllTime };
  }, [transactions, currentPrices, sectorFn]);

  if (!loaded) return <div className="min-h-screen bg-slate-950 flex items-center justify-center"><div className="text-slate-400 font-mono text-sm animate-pulse">Initialisation du terminal…</div></div>;

  const tabs = [
    { id: "dashboard", label: "Dashboard", icon: BarChart3 },
    { id: "finances", label: "Finances", icon: Wallet },
    { id: "periods", label: "Périodes", icon: Calendar },
    { id: "analyses", label: "Analyses", icon: Target },
    { id: "transactions", label: "Transactions", icon: Activity },
    { id: "settings", label: "Réglages", icon: SettingsIcon },
  ];

  const totalGain = derived.stats.totalPnL + derived.unrealizedPnL + derived.stats.totalDividends + derived.stats.totalInterests + derived.stats.totalSavebacks;

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100" style={{ fontFamily: "ui-sans-serif, -apple-system, BlinkMacSystemFont, system-ui, sans-serif" }}>
      <header className="bg-gradient-to-b from-slate-900 to-slate-950 border-b border-slate-800 sticky top-0 z-30 backdrop-blur">
        <div className="max-w-[1600px] mx-auto px-4 md:px-6 py-3">
          <div className="flex items-center justify-between gap-4 flex-wrap">
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 rounded-lg bg-gradient-to-br from-cyan-400 to-emerald-500 flex items-center justify-center shadow-lg shadow-cyan-500/20">
                <Zap className="w-5 h-5 text-slate-950" strokeWidth={2.5} />
              </div>
              <div>
                <div className="text-lg font-bold tracking-tight">Portfolio Tracker <span className="text-cyan-400">PRO</span></div>
                <div className="text-[10px] uppercase tracking-widest text-slate-500">Trade Republic Analytics</div>
              </div>
            </div>
            <div className="flex items-center gap-3 text-xs font-mono">
              <button onClick={() => { const n = !hideBalances; setHideBalances(n); storage.set(STORAGE_KEYS.HIDE_BALANCES, String(n)); }}
                className="p-2 bg-slate-800 hover:bg-slate-700 rounded text-slate-300 transition-colors" title={hideBalances ? "Afficher" : "Masquer"}>
                {hideBalances ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
              </button>
              <div className="hidden sm:flex flex-col items-end"><div className="text-[9px] uppercase text-slate-500 tracking-widest">Tx</div><div className="text-slate-200 font-semibold">{transactions.length}</div></div>
              <div className="flex flex-col items-end"><div className="text-[9px] uppercase text-slate-500 tracking-widest">Valeur</div><div className="text-slate-200 font-semibold">{mask(fmtEURCompact(derived.openValue), hideBalances)}</div></div>
              <div className="flex flex-col items-end"><div className="text-[9px] uppercase text-slate-500 tracking-widest">P&L</div><div className={`font-semibold ${classFor(totalGain)}`}>{mask(fmtEURCompact(totalGain), hideBalances)}</div></div>
            </div>
          </div>
          <div className="flex gap-1 mt-3 overflow-x-auto -mb-px scrollbar-thin">
            {tabs.map(t => (
              <button key={t.id} onClick={() => setActiveTab(t.id)}
                className={`flex items-center gap-2 px-3 md:px-4 py-2 text-xs font-medium uppercase tracking-wide whitespace-nowrap border-b-2 transition-colors ${activeTab === t.id ? "border-cyan-400 text-cyan-300" : "border-transparent text-slate-400 hover:text-slate-100 hover:border-slate-600"}`}>
                <t.icon className="w-3.5 h-3.5" /><span className="hidden sm:inline">{t.label}</span><span className="sm:hidden">{t.label.slice(0, 4)}</span>
              </button>
            ))}
          </div>
        </div>
      </header>

      {importMsg && (
        <div className="max-w-[1600px] mx-auto px-4 md:px-6 pt-4">
          <div className={`border rounded-lg p-3 text-sm flex items-start gap-3 ${importMsg.type === "success" ? "bg-emerald-500/10 border-emerald-500/30 text-emerald-200" : importMsg.type === "error" ? "bg-rose-500/10 border-rose-500/30 text-rose-200" : "bg-cyan-500/10 border-cyan-500/30 text-cyan-200"}`}>
            {importMsg.type === "success" ? <Check className="w-4 h-4 shrink-0 mt-0.5" /> : importMsg.type === "error" ? <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" /> : <RefreshCw className="w-4 h-4 shrink-0 mt-0.5 animate-spin" />}
            <div className="flex-1 min-w-0"><div>{importMsg.text}</div>{importMsg.detail && <div className="text-xs opacity-70 mt-1 font-mono break-all">{importMsg.detail}</div>}</div>
          </div>
        </div>
      )}

      <main className="max-w-[1600px] mx-auto px-4 md:px-6 py-6">
        {transactions.length === 0 && activeTab !== "settings" ? <Welcome onImport={handleFileImport} /> : (
          <>
            {activeTab === "dashboard" && <Dashboard d={derived} sectorFn={sectorFn} hide={hideBalances} prices={currentPrices} onPriceChange={(s, p) => { const np = { ...currentPrices, [s]: p }; if (!p) delete np[s]; setCurrentPrices(np); persistPrices(np); }} />}
            {activeTab === "finances" && <FinancesView d={derived} hide={hideBalances} />}
            {activeTab === "periods" && <PeriodsView d={derived} sectorFn={sectorFn} hide={hideBalances} />}
            {activeTab === "analyses" && <AnalysesView d={derived} sectorFn={sectorFn} hide={hideBalances} />}
            {activeTab === "transactions" && <TransactionsView d={derived} sectorFn={sectorFn} hide={hideBalances} />}
            {activeTab === "settings" && <Settings onImport={handleFileImport} transactions={transactions} onClear={() => { setTransactions([]); persistTx([]); }} sectorOverrides={sectorOverrides} onSectorOverride={(s, v) => { const n = { ...sectorOverrides }; if (v) n[s] = v; else delete n[s]; setSectorOverrides(n); persistSectors(n); }} derived={derived} sectorFn={sectorFn} />}
          </>
        )}
      </main>
      <footer className="max-w-[1600px] mx-auto px-6 py-6 text-center text-xs text-slate-600">Portfolio Tracker Pro v3 · FIFO · Trade Republic · Stockage local</footer>
    </div>
  );
}

function Welcome({ onImport }) {
  const inputRef = useRef(null);
  const [dragOver, setDragOver] = useState(false);
  return (
    <div className="max-w-3xl mx-auto mt-10">
      <div className="text-center mb-8">
        <h1 className="text-3xl font-bold mb-2">Bienvenue dans ton <span className="text-cyan-400">terminal d'analyse</span></h1>
        <p className="text-slate-400">Importe ton export Trade Republic (CSV ou XLSX) pour commencer.</p>
      </div>
      <div onClick={() => inputRef.current?.click()} onDragOver={(e) => { e.preventDefault(); setDragOver(true); }} onDragLeave={() => setDragOver(false)} onDrop={(e) => { e.preventDefault(); setDragOver(false); if (e.dataTransfer.files.length) onImport(Array.from(e.dataTransfer.files)); }}
        className={`border-2 border-dashed rounded-xl p-14 text-center cursor-pointer transition-all ${dragOver ? "border-cyan-400 bg-cyan-500/5" : "border-slate-700 hover:border-slate-500 bg-slate-900/50"}`}>
        <Upload className="w-12 h-12 mx-auto mb-4 text-slate-500" />
        <div className="text-lg font-semibold mb-1">Glisse ton fichier ici</div>
        <div className="text-sm text-slate-400">ou clique pour sélectionner</div>
        <input ref={inputRef} type="file" multiple accept=".xlsx,.xls,.csv" onChange={(e) => e.target.files?.length && onImport(Array.from(e.target.files))} className="hidden" />
      </div>
    </div>
  );
}

// ═══════════════ DASHBOARD REFONTE COMPLETE ═══════════════

function Dashboard({ d, sectorFn, hide, prices, onPriceChange }) {
  const { stats, bySector, openValue, unrealizedPnL, openInvested, realizedTrades, consolidatedTrades, curveSeries, totalDeposits, totalWithdrawals, totalCardSpend, byAsset, feesEstimation, taxation } = d;

  // ═════ CALCUL BRUT / FRAIS / TAXE / NET ═════
  // Niveau 1 : BRUT TOTAL = P&L réalisée + P&L latente + Dividendes + Intérêts + Saveback + Gifts
  const revenus = stats.totalDividends + stats.totalInterests + stats.totalSavebacks + stats.totalGifts;
  const gainTotalBrut = stats.totalPnL + unrealizedPnL + revenus;
  // Niveau 2 : APRÈS FRAIS = Brut − frais courtage estimés
  const totalFeesEst = feesEstimation?.totalEstimated || 0;
  const gainAfterFees = gainTotalBrut - totalFeesEst;
  // Niveau 3 : APRÈS IMPÔTS = Après-frais − taxe CTO 31,4% sur les gains imposables CTO
  const taxCTO = taxation?.taxCTOEstimated || 0;
  const gainNet = gainAfterFees - taxCTO;
  const totalInvAT = stats.totalInvested + openInvested;
  const pctBrut = totalInvAT > 0 ? (gainTotalBrut / totalInvAT) * 100 : 0;
  const pctNet = totalInvAT > 0 ? (gainNet / totalInvAT) * 100 : 0;

  const pieData = bySector.filter(s => s.investedOpen > 0.01 || s.currentValue > 0.01).map(s => ({ name: s.sector, value: Math.max(0, s.currentValue || s.investedOpen), color: SECTOR_COLORS[s.sector] || "#64748b" }));

  // Positions enrichies (récupéré de l'ancien Positions)
  const positions = useMemo(() => {
    const m = {};
    for (const l of d.openLots) {
      const k = l.symbol || l.name;
      if (!m[k]) m[k] = { symbol: l.symbol, name: l.name, shares: 0, invested: 0, oldestDate: l.date, lots: [], assetClass: l.ref?.assetClass || "" };
      m[k].shares += l.shares;
      m[k].invested += l.shares * l.costPerShare;
      if (l.date < m[k].oldestDate) m[k].oldestDate = l.date;
      m[k].lots.push(l);
    }
    return Object.values(m).map(p => {
      const px = prices[p.symbol] || prices[p.name];
      const avg = p.shares ? p.invested / p.shares : 0;
      const val = px ? p.shares * px : null;
      const pnl = val !== null ? val - p.invested : null;
      const pct = val !== null && p.invested > 0 ? (pnl / p.invested) * 100 : null;
      return { ...p, avgCost: avg, currentPrice: px || null, currentValue: val, unrealizedPnL: pnl, unrealizedPct: pct, sector: sectorFn(p), lotCount: p.lots.length };
    }).sort((a, b) => b.invested - a.invested);
  }, [d.openLots, prices, sectorFn]);

  const missingPx = positions.filter(p => p.currentPrice === null).length;

  return (
    <div className="space-y-6">
      {/* ═══════ HERO PRINCIPAL ═══════ */}
      <div className="relative overflow-hidden rounded-2xl border border-slate-800 bg-gradient-to-br from-slate-900 via-slate-900 to-slate-950">
        <div className="absolute -top-40 -right-40 w-96 h-96 rounded-full opacity-10 blur-3xl" style={{ background: gainNet >= 0 ? "#10b981" : "#f43f5e" }} />
        <div className="absolute -bottom-40 -left-40 w-96 h-96 rounded-full opacity-5 blur-3xl bg-cyan-500" />
        <div className="relative p-6 md:p-8">
          {/* Bandeau */}
          <div className="flex items-start justify-between mb-6 flex-wrap gap-4">
            <div>
              <div className="flex items-center gap-2 text-[10px] uppercase tracking-[0.2em] text-slate-500 font-medium mb-2">
                <Sparkles className="w-3 h-3 text-cyan-400" />Vue d'ensemble du portefeuille
              </div>
              <div className="text-xs text-slate-400">
                {positions.length > 0 && `${positions.length} positions · `}{stats.count} ventes · depuis {fmtDateShort(d.transactions[0]?.date)}
              </div>
            </div>
            <Pill color={gainNet >= 0 ? "emerald" : "rose"}>{gainNet >= 0 ? "📈 Portefeuille gagnant" : "📉 Portefeuille en perte"}</Pill>
          </div>

          {/* 3 chiffres principaux */}
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-6">
            <div>
              <div className="text-[10px] uppercase tracking-widest text-slate-500 mb-1 flex items-center gap-1"><Wallet className="w-3 h-3" /> Valeur du portefeuille</div>
              <div className="text-4xl md:text-5xl font-mono font-black tabular-nums text-slate-100 leading-none">{mask(fmtEUR(openValue, 2), hide)}</div>
              <div className="text-sm text-slate-400 mt-2 font-mono">Investi : {mask(fmtEUR(openInvested, 2), hide)}</div>
            </div>
            <div>
              <div className="text-[10px] uppercase tracking-widest text-slate-500 mb-1 flex items-center gap-1"><Flame className="w-3 h-3" /> Gain total all-time</div>
              <div className={`text-4xl md:text-5xl font-mono font-black tabular-nums leading-none ${classFor(gainTotalBrut)}`}>
                {mask((gainTotalBrut >= 0 ? "+" : "") + fmtEUR(gainTotalBrut, 2).replace("+", ""), hide)}
              </div>
              <div className="text-[10px] uppercase tracking-widest text-slate-500 mt-1">brut · {fmtPct(pctBrut, 2)}</div>
            </div>
            <div>
              <div className="text-[10px] uppercase tracking-widest text-slate-500 mb-1 flex items-center gap-1"><DollarSign className="w-3 h-3" /> Capital déposé net</div>
              <div className="text-4xl md:text-5xl font-mono font-black tabular-nums text-slate-100 leading-none">{mask(fmtEUR(totalDeposits - totalWithdrawals, 0), hide)}</div>
              <div className="text-sm text-slate-400 mt-2 font-mono">Dépenses CB : <span className="text-rose-400">{mask(fmtEUR(-totalCardSpend, 0), hide)}</span></div>
            </div>
          </div>

          {/* CASCADE Brut → Frais → Impôts → Net */}
          <div className="bg-slate-950/40 rounded-xl border border-slate-800/60 p-4 md:p-5">
            <div className="text-[10px] uppercase tracking-[0.2em] text-cyan-400 font-bold mb-3">💎 Du brut au net réel</div>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
              <CascadeStep label="Brut" value={gainTotalBrut} icon={Sparkles} color={classFor(gainTotalBrut)} hide={hide} sub="Tout cumulé" />
              <CascadeStep label="− Frais courtage" value={-totalFeesEst} icon={Coins} color="text-amber-400" hide={hide} sub={`${(feesEstimation?.spotBuyCount || 0) + (feesEstimation?.spotSellCount || 0)} ordres × 1€`} />
              <CascadeStep label="− Impôts CTO (31,4%)" value={-taxCTO} icon={Percent} color="text-rose-400" hide={hide} sub={taxation?.hasPEA ? "PEA exonéré" : "PFU France"} />
              <CascadeStep label="= Net réel" value={gainNet} icon={Flame} color={classFor(gainNet)} hide={hide} sub={`${fmtPct(pctNet, 2)}`} highlight />
            </div>
            {taxation && taxation.taxableGainsCTO > 0 && (
              <div className="mt-3 text-[11px] text-slate-500 font-mono">
                ⓘ Imposable CTO : {fmtEUR(taxation.taxableGainsCTO, 2)} (P&L positive {fmtEUR(Math.max(0, taxation.realizedPnLCTO), 2)} + div+int CTO {fmtEUR(Math.max(0, taxation.divsCTO + taxation.intsCTO), 2)})
                {taxation.hasPEA && <span> · PEA non taxé : {fmtEUR(taxation.realizedPnLPEA + taxation.divsPEA + taxation.intsPEA, 2)}</span>}
              </div>
            )}
          </div>

          {/* Breakdown détaillé */}
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mt-3 pt-4 border-t border-slate-800">
            <PnLBreakdown label="P&L latente" value={unrealizedPnL} pct={openInvested > 0 ? (unrealizedPnL / openInvested) * 100 : null} icon={TrendingUp} hide={hide} />
            <PnLBreakdown label="P&L réalisée" value={stats.totalPnL} pct={stats.totalInvested > 0 ? (stats.totalPnL / stats.totalInvested) * 100 : null} icon={Activity} hide={hide} />
            <PnLBreakdown label="Dividendes + Intérêts" value={stats.totalDividends + stats.totalInterests} icon={Coins} hide={hide} accent="amber" />
            <PnLBreakdown label="Bonus (Saveback + Gift)" value={stats.totalSavebacks + stats.totalGifts} icon={Gift} hide={hide} accent="violet" />
          </div>
        </div>
      </div>

      {/* ═══════ FRAIS DE COURTAGE ═══════ */}
      <FraisCourtageCard d={d} hide={hide} />

      {/* ═══════ GAINS RÉCENTS CLIQUABLES ═══════ */}
      <GainsRecents d={d} hide={hide} />

      {/* ═══════ COURBE D'ÉVOLUTION ═══════ */}
      {curveSeries.length > 0 && (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <Card>
            <SectionTitle icon={Wallet} subtitle="Somme mise au travail sur tes positions ouvertes">Capital investi cumulé</SectionTitle>
            <ResponsiveContainer width="100%" height={260}>
              <AreaChart data={curveSeries}>
                <defs><linearGradient id="gInv" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#06b6d4" stopOpacity={0.5} /><stop offset="100%" stopColor="#06b6d4" stopOpacity={0} /></linearGradient></defs>
                <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" vertical={false} />
                <XAxis dataKey="date" stroke="#64748b" fontSize={10} tickFormatter={(v) => v?.slice(5)} minTickGap={40} />
                <YAxis stroke="#64748b" fontSize={10} tickFormatter={(v) => `${(v / 1000).toFixed(0)}k`} />
                <Tooltip content={<ChartTooltip />} />
                <Area type="monotone" dataKey="invested" name="Capital investi net" stroke="#06b6d4" strokeWidth={2.5} fill="url(#gInv)" />
              </AreaChart>
            </ResponsiveContainer>
          </Card>
          <Card>
            <SectionTitle icon={TrendingUp} subtitle="Tes plus-values réalisées + dividendes encaissés">Gains cumulés dans le temps</SectionTitle>
            <ResponsiveContainer width="100%" height={260}>
              <ComposedChart data={curveSeries}>
                <defs>
                  <linearGradient id="gRz" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#10b981" stopOpacity={0.4} /><stop offset="100%" stopColor="#10b981" stopOpacity={0} /></linearGradient>
                  <linearGradient id="gDv" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#f59e0b" stopOpacity={0.3} /><stop offset="100%" stopColor="#f59e0b" stopOpacity={0} /></linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" vertical={false} />
                <XAxis dataKey="date" stroke="#64748b" fontSize={10} tickFormatter={(v) => v?.slice(5)} minTickGap={40} />
                <YAxis stroke="#64748b" fontSize={10} tickFormatter={(v) => fmtEUR(v, 0)} />
                <Tooltip content={<ChartTooltip />} />
                <ReferenceLine y={0} stroke="#475569" />
                <Area type="monotone" dataKey="realized" name="P&L réalisée cumulée" stroke="#10b981" strokeWidth={2.5} fill="url(#gRz)" />
                <Area type="monotone" dataKey="divInt" name="Dividendes+Intérêts" stroke="#f59e0b" strokeWidth={2} fill="url(#gDv)" />
                <Legend wrapperStyle={{ fontSize: 11, color: "#cbd5e1", paddingTop: 8 }} />
              </ComposedChart>
            </ResponsiveContainer>
          </Card>
        </div>
      )}

      {/* ═══════ MÉTRIQUES TRADING ═══════ */}
      <div>
        <div className="text-[10px] uppercase tracking-widest text-slate-500 font-medium mb-3">📊 Métriques clés</div>
        <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3">
          <KPICard label="Win rate" value={stats.count ? stats.winRate.toFixed(1) + "%" : "—"} icon={Target} accent={stats.winRate >= 50 ? "emerald" : "rose"} sub={stats.count ? `${stats.winCount}W · ${stats.lossCount}L` : ""} />
          <KPICard label="Profit factor" value={stats.profitFactor === Infinity ? "∞" : stats.profitFactor.toFixed(2)} icon={Zap} accent={stats.profitFactor >= 1.5 ? "emerald" : stats.profitFactor >= 1 ? "amber" : "rose"} sub="> 1,5 = excellent" />
          <KPICard label="Gain moyen" value={mask(fmtEUR(stats.avgWin, 0), hide)} icon={ArrowUpRight} accent="emerald" />
          <KPICard label="Perte moyenne" value={mask(fmtEUR(-stats.avgLoss, 0), hide)} icon={ArrowDownRight} accent="rose" />
          <KPICard label="Durée moyenne" value={stats.avgHolding ? `${Math.round(stats.avgHolding)} j` : "—"} icon={Clock} accent="slate" sub={`Série: ${stats.largestWinStreak}W max`} />
          <KPICard label="Max drawdown" value={mask(fmtEUR(-stats.maxDrawdown, 0), hide)} icon={TrendingDown} accent="rose" sub={`Série: ${stats.largestLossStreak}L max`} />
        </div>
      </div>

      {/* ═══════ SCORE DE SANTÉ + XIRR ═══════ */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <HealthScoreCard d={d} hide={hide} />
        <XirrCard d={d} hide={hide} />
      </div>

      {/* ═══════ HEATMAP CALENDRIER DES GAINS ═══════ */}
      <GainsHeatmap d={d} hide={hide} />

      {/* ═══════ RECORDS & MILESTONES ═══════ */}
      <RecordsSection d={d} hide={hide} />

      {/* ═══════ COMPORTEMENT D'INVESTISSEMENT ═══════ */}
      <BehaviorSection d={d} hide={hide} />

      {/* ═══════ RÉPARTITION SECTORIELLE (camembert cliquable → modal positions du secteur) ═══════ */}
      {pieData.length > 0 && (
        <SectorPieClickable pieData={pieData} positions={positions} hide={hide} onPriceChange={onPriceChange} missingPx={missingPx} sectorFn={sectorFn} />
      )}

      {/* ═══════ MEILLEURES POSITIONS ═══════ */}
      <MeilleuresPositions d={d} sectorFn={sectorFn} hide={hide} />
    </div>
  );
}

// Composant interne pour la cascade Brut → Net
const CascadeStep = ({ label, value, icon: Icon, color, hide, sub, highlight }) => (
  <div className={`relative rounded-lg p-3 ${highlight ? "bg-cyan-500/10 border-2 border-cyan-500/40" : "bg-slate-900/50 border border-slate-800"}`}>
    <div className="flex items-center gap-1.5 text-[9px] uppercase tracking-widest text-slate-500 font-bold">
      <Icon className="w-2.5 h-2.5" />{label}
    </div>
    <div className={`text-xl md:text-2xl font-mono font-black tabular-nums mt-1 ${color}`}>
      {mask((value > 0 && !label.startsWith("−") ? "+" : "") + fmtEUR(value, 0).replace("+", ""), hide)}
    </div>
    {sub && <div className="text-[9px] text-slate-500 font-mono mt-0.5">{sub}</div>}
  </div>
);

// Tableau positions (extrait de l'ancien composant Positions)
const PositionsTable = ({ positions, onPriceChange, hide }) => {
  const [editing, setEditing] = useState(null);
  const [editVal, setEditVal] = useState("");
  const [sort, setSort] = useState({ key: "invested", dir: "desc" });

  const sorted = useMemo(() => {
    const arr = [...positions];
    const dir = sort.dir === "asc" ? 1 : -1;
    arr.sort((a, b) => {
      const va = a[sort.key], vb = b[sort.key];
      if (va === null || va === undefined) return 1;
      if (vb === null || vb === undefined) return -1;
      if (typeof va === "string") return va.localeCompare(vb) * dir;
      return (va - vb) * dir;
    });
    return arr;
  }, [positions, sort]);

  const startEdit = (k, v) => { setEditing(k); setEditVal(v !== null ? String(v) : ""); };
  const commitEdit = (s, n) => { const v = parseNumber(editVal); onPriceChange(s || n, v || null); setEditing(null); setEditVal(""); };
  const totalInv = positions.reduce((s, p) => s + p.invested, 0);

  return (
    <div className="overflow-x-auto">
      <table className="w-full text-xs">
        <thead>
          <tr className="border-b border-slate-800 bg-slate-900">
            {[{ k: "name", l: "Actif" }, { k: "sector", l: "Secteur" }, { k: "invested", l: "Pondération", n: true }, { k: "shares", l: "Qté", n: true }, { k: "avgCost", l: "PRU", n: true }, { k: "currentPrice", l: "Cours", n: true }, { k: "currentValue", l: "Valeur", n: true }, { k: "unrealizedPnL", l: "P&L €", n: true }, { k: "unrealizedPct", l: "P&L %", n: true }].map(c => (
              <th key={c.k} className={`p-2 font-semibold text-slate-400 uppercase tracking-wider text-[10px] cursor-pointer hover:text-slate-100 ${c.n ? "text-right" : "text-left"}`} onClick={() => setSort(s => ({ key: c.k, dir: s.key === c.k && s.dir === "desc" ? "asc" : "desc" }))}>
                {c.l} {sort.key === c.k && (sort.dir === "desc" ? "↓" : "↑")}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {sorted.map((p, i) => {
            const key = p.symbol || p.name;
            const isEditing = editing === key;
            const weight = totalInv > 0 ? (p.invested / totalInv) * 100 : 0;
            return (
              <tr key={i} className={`border-b border-slate-800/60 ${ROW_HOVER}`}>
                <td className="p-2"><div className="font-semibold text-slate-100 truncate max-w-[200px]">{p.name || p.symbol}</div><div className="text-[9px] text-slate-500 font-mono">{p.symbol}</div></td>
                <td className="p-2"><span className="inline-flex items-center gap-1.5"><span className="w-2 h-2 rounded-full" style={{ background: SECTOR_COLORS[p.sector] || "#64748b" }} /><span className="text-slate-300 text-[11px]">{p.sector}</span></span></td>
                <td className="p-2 text-right">
                  <div className="flex items-center justify-end gap-2">
                    <div className="w-16 h-1.5 bg-slate-800 rounded-full overflow-hidden">
                      <div className="h-full rounded-full" style={{ width: `${weight}%`, background: SECTOR_COLORS[p.sector] || "#64748b" }} />
                    </div>
                    <span className="font-mono tabular-nums text-slate-300 text-[10px] w-10">{weight.toFixed(1)}%</span>
                  </div>
                  <div className="text-[9px] text-slate-500 font-mono mt-0.5">{mask(fmtEUR(p.invested, 0), hide)}</div>
                </td>
                <td className="p-2 text-right font-mono tabular-nums">{fmtNum(p.shares, 4)}</td>
                <td className="p-2 text-right font-mono tabular-nums text-slate-300">{fmtEUR(p.avgCost, 4)}</td>
                <td className="p-2 text-right font-mono tabular-nums">
                  {isEditing ? (
                    <input autoFocus type="text" value={editVal} onChange={(e) => setEditVal(e.target.value)} onBlur={() => commitEdit(p.symbol, p.name)} onKeyDown={(e) => { if (e.key === "Enter") commitEdit(p.symbol, p.name); if (e.key === "Escape") { setEditing(null); setEditVal(""); } }} className="bg-slate-800 border border-cyan-500 rounded px-2 py-0.5 text-right w-24 text-cyan-200" />
                  ) : (
                    <button onClick={() => startEdit(key, p.currentPrice)} className={`px-2 py-0.5 rounded hover:bg-slate-700 transition-colors ${p.currentPrice === null ? "text-amber-400 border border-amber-500/30 border-dashed" : "text-cyan-300"}`}>
                      {p.currentPrice !== null ? fmtEUR(p.currentPrice, 4) : "saisir…"}
                    </button>
                  )}
                </td>
                <td className="p-2 text-right font-mono tabular-nums text-slate-100 font-semibold">{p.currentValue !== null ? mask(fmtEUR(p.currentValue, 2), hide) : "—"}</td>
                <td className={`p-2 text-right font-mono tabular-nums font-semibold ${classFor(p.unrealizedPnL)}`}>{p.unrealizedPnL !== null ? mask(fmtEUR(p.unrealizedPnL, 2), hide) : "—"}</td>
                <td className={`p-2 text-right font-mono tabular-nums ${classFor(p.unrealizedPct)}`}>{p.unrealizedPct !== null ? fmtPct(p.unrealizedPct, 2) : "—"}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
};

// ═══════ Camembert sectoriel CLIQUABLE → modal positions du secteur ═══════
const SectorPieClickable = ({ pieData, positions, hide, onPriceChange, missingPx, sectorFn }) => {
  const [openSector, setOpenSector] = useState(null);
  const sortedPie = useMemo(() => [...pieData].sort((a, b) => b.value - a.value), [pieData]);
  const tot = pieData.reduce((s, x) => s + x.value, 0);

  return (
    <>
      <Card>
        <SectionTitle icon={PieIcon} subtitle="Clique sur un secteur pour voir tes positions détaillées" actions={
          <span className="text-[10px] text-slate-500 font-mono">{positions.length} positions · {positions.reduce((s, p) => s + p.lotCount, 0)} lots</span>
        }>
          🥧 Répartition sectorielle
        </SectionTitle>
        {missingPx > 0 && (
          <div className="bg-amber-500/10 border border-amber-500/30 rounded-lg p-2.5 text-[11px] text-amber-200 flex items-start gap-2 mb-3">
            <Info className="w-3.5 h-3.5 shrink-0 mt-0.5" />
            <span>{missingPx} position(s) sans cours saisi — clique sur un secteur puis sur le cours pour les compléter.</span>
          </div>
        )}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4 items-center">
          <ResponsiveContainer width="100%" height={280}>
            <PieChart>
              <Pie data={sortedPie} dataKey="value" nameKey="name" innerRadius={62} outerRadius={115} paddingAngle={3} stroke="#0f172a" strokeWidth={2} onClick={(e) => setOpenSector(e.name)} cursor="pointer">
                {sortedPie.map((e, i) => <Cell key={i} fill={e.color} />)}
              </Pie>
              <Tooltip content={<ChartTooltip />} />
            </PieChart>
          </ResponsiveContainer>
          <div className="space-y-1.5">
            {sortedPie.map((s, i) => {
              const pct = tot > 0 ? (s.value / tot) * 100 : 0;
              return (
                <button key={i} onClick={() => setOpenSector(s.name)}
                  className="w-full flex items-center gap-2 text-xs p-2 rounded hover:bg-slate-800/60 transition-colors cursor-pointer text-left group">
                  <span className="w-2.5 h-2.5 rounded-sm shrink-0" style={{ background: s.color }} />
                  <span className="flex-1 text-slate-300 truncate group-hover:text-slate-100">{s.name}</span>
                  <span className="font-mono text-slate-400 tabular-nums">{pct.toFixed(1)}%</span>
                  <span className="font-mono text-slate-100 font-semibold tabular-nums w-16 text-right">{mask(fmtEURCompact(s.value), hide)}</span>
                  <span className="text-cyan-400/0 group-hover:text-cyan-400/80 text-[9px] transition-colors">↗</span>
                </button>
              );
            })}
          </div>
        </div>
      </Card>
      {openSector && <SectorPositionsModal sector={openSector} color={sortedPie.find(s => s.name === openSector)?.color || "#64748b"} positions={positions.filter(p => p.sector === openSector)} onClose={() => setOpenSector(null)} onPriceChange={onPriceChange} hide={hide} />}
    </>
  );
};

const SectorPositionsModal = ({ sector, color, positions, onClose, onPriceChange, hide }) => {
  useEffect(() => {
    const orig = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => { document.body.style.overflow = orig; };
  }, []);
  const totalInv = positions.reduce((s, p) => s + p.invested, 0);
  const totalVal = positions.reduce((s, p) => s + (p.currentValue ?? p.invested), 0);
  const totalPnL = positions.reduce((s, p) => s + (p.unrealizedPnL ?? 0), 0);
  return (
    <div className="fixed inset-0 z-50 flex items-end md:items-center justify-center bg-black/70 backdrop-blur-sm p-0 md:p-4" onClick={onClose}>
      <div className="bg-slate-900 border border-slate-700 rounded-t-2xl md:rounded-2xl w-full md:max-w-5xl max-h-[92vh] overflow-hidden flex flex-col" onClick={e => e.stopPropagation()}>
        <div className="p-5 border-b border-slate-800" style={{ background: `linear-gradient(135deg, ${color}22, transparent)` }}>
          <div className="flex items-start justify-between gap-3">
            <div className="min-w-0">
              <div className="flex items-center gap-2 mb-1">
                <span className="w-3 h-3 rounded-full" style={{ background: color }} />
                <span className="text-[10px] uppercase tracking-widest text-slate-400">Secteur</span>
              </div>
              <div className="text-2xl md:text-3xl font-bold">{sector}</div>
              <div className="text-xs text-slate-400 font-mono mt-1">{positions.length} position{positions.length > 1 ? "s" : ""}</div>
            </div>
            <button onClick={onClose} className="p-2 bg-slate-800 hover:bg-slate-700 rounded text-slate-300 text-xl leading-none">✕</button>
          </div>
          <div className="mt-4 grid grid-cols-3 gap-3">
            <div className="bg-slate-950/60 rounded-lg p-2.5 border border-slate-800">
              <div className="text-[9px] uppercase tracking-widest text-slate-500">Investi</div>
              <div className="text-base font-mono font-bold tabular-nums">{mask(fmtEUR(totalInv, 0), hide)}</div>
            </div>
            <div className="bg-slate-950/60 rounded-lg p-2.5 border border-slate-800">
              <div className="text-[9px] uppercase tracking-widest text-slate-500">Valeur</div>
              <div className="text-base font-mono font-bold tabular-nums">{mask(fmtEUR(totalVal, 0), hide)}</div>
            </div>
            <div className="bg-slate-950/60 rounded-lg p-2.5 border border-slate-800">
              <div className="text-[9px] uppercase tracking-widest text-slate-500">P&L latente</div>
              <div className={`text-base font-mono font-bold tabular-nums ${classFor(totalPnL)}`}>{mask((totalPnL >= 0 ? "+" : "") + fmtEUR(totalPnL, 0).replace("+", ""), hide)}</div>
            </div>
          </div>
        </div>
        <div className="overflow-y-auto flex-1 p-2">
          <PositionsTable positions={positions} onPriceChange={onPriceChange} hide={hide} />
        </div>
      </div>
    </div>
  );
};

// Gains sur différentes périodes glissantes (1j / 7j / 30j / 90j / YTD)
// Cartes CLIQUABLES : ouvre un modal avec le détail des ventes
const GainsRecents = ({ d, hide }) => {
  const now = new Date();
  const [openBucket, setOpenBucket] = useState(null);
  const buckets = [
    { id: "1j", label: "Aujourd'hui", days: 1 },
    { id: "7j", label: "7 jours", days: 7 },
    { id: "30j", label: "30 jours", days: 30 },
    { id: "90j", label: "90 jours", days: 90 },
    { id: "ytd", label: "Depuis janv.", days: null, ytd: true },
    { id: "all", label: "All-time", days: null, all: true },
  ];
  const computeFor = (b) => {
    let from;
    if (b.all) from = new Date(0);
    else if (b.ytd) from = new Date(now.getFullYear(), 0, 1);
    else { from = new Date(now); from.setDate(from.getDate() - b.days); from.setHours(0, 0, 0, 0); }
    const tradesRaw = d.realizedTrades.filter(t => t.sellDate >= from && t.sellDate <= now);
    const trades = consolidateTrades(tradesRaw); // 1 vente = 1 ligne
    const divsList = d.dividends.filter(x => x.date >= from && x.date <= now);
    const intsList = d.interests.filter(x => x.date >= from && x.date <= now);
    const svbList = d.savebacks.filter(x => x.date >= from && x.date <= now);
    const divs = divsList.reduce((s, x) => s + x.net, 0);
    const ints = intsList.reduce((s, x) => s + x.net, 0);
    const svb = svbList.reduce((s, x) => s + x.amount, 0);
    const realizedBrut = trades.reduce((s, t) => s + t.pnl, 0);
    // Frais estimés sur les VENTES de cette période (1€/vente spot)
    const spotSells = trades.length;
    const feesOnSells = spotSells * 1.0;
    const realizedAfterFees = realizedBrut - feesOnSells;
    // TAXE CTO 31,4% sur P&L positive + dividendes + intérêts (PEA exonéré)
    const TAX = 0.314;
    const ctoTrades = tradesRaw.filter(t => t.accountType !== "PEA");
    const ctoPnL = ctoTrades.reduce((s, t) => s + t.pnl, 0);
    const ctoPnLAfterFees = ctoPnL - feesOnSells * (ctoTrades.length / Math.max(1, tradesRaw.length)); // au prorata
    const ctoDivs = divsList.filter(d => d.accountType !== "PEA").reduce((s, x) => s + x.net, 0);
    const ctoInts = intsList.filter(i => i.accountType !== "PEA").reduce((s, x) => s + x.net, 0);
    const taxableCTO = Math.max(0, ctoPnLAfterFees) + Math.max(0, ctoDivs) + Math.max(0, ctoInts);
    const taxCTO = taxableCTO * TAX;
    const brut = realizedBrut + divs + ints + svb;
    const afterFees = realizedAfterFees + divs + ints + svb;
    const net = afterFees - taxCTO;
    return { realizedBrut, realizedAfterFees, feesOnSells, taxCTO, divs, ints, svb,
             brut, afterFees, net,
             tradeCount: trades.length, trades, divsList, intsList, svbList,
             from, to: now, label: b.label };
  };
  return (
    <>
      <div>
        <div className="text-[10px] uppercase tracking-widest text-slate-500 font-medium mb-3 flex items-center gap-2"><Flame className="w-3 h-3 text-orange-400" />💸 Combien j'ai gagné récemment <span className="text-slate-600 normal-case tracking-normal">(clique pour voir le détail)</span></div>
        <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3">
          {buckets.map(b => {
            const r = computeFor(b);
            const clickable = r.tradeCount > 0 || r.divsList.length > 0 || r.intsList.length > 0 || r.svbList.length > 0;
            return (
              <button key={b.id}
                onClick={() => clickable && setOpenBucket({ ...r, bucketLabel: b.label })}
                disabled={!clickable}
                className={`relative overflow-hidden rounded-xl p-4 border text-left transition-all ${clickable ? "cursor-pointer hover:scale-[1.02] hover:border-cyan-500/50" : "cursor-default"} ${r.net > 0 ? "border-emerald-500/30 bg-emerald-500/5" : r.net < 0 ? "border-rose-500/30 bg-rose-500/5" : "border-slate-800 bg-slate-900/60"}`}>
                <div className="flex items-center justify-between">
                  <div className="text-[10px] uppercase tracking-widest text-slate-400 font-medium">{b.label}</div>
                  {clickable && <span className="text-[9px] text-cyan-400/60">↗</span>}
                </div>
                <div className={`text-2xl font-mono font-black tabular-nums mt-1 ${classFor(r.net)}`}>{mask((r.net >= 0 ? "+" : "") + fmtEUR(r.net, 0).replace("+", ""), hide)}</div>
                <div className="text-[10px] text-slate-500 font-mono mt-1 space-y-0.5">
                  <div>
                    {r.tradeCount > 0 ? `${r.tradeCount} vente${r.tradeCount > 1 ? "s" : ""}` : "Aucune vente"}
                    {r.divs + r.ints > 0 && ` · +${fmtEUR(r.divs + r.ints, 0)} div/int`}
                  </div>
                  {(r.feesOnSells > 0 || r.taxCTO > 0) && (
                    <div className="text-[9px] text-slate-600 leading-tight">
                      Brut {fmtEUR(r.brut, 0)}
                      {r.feesOnSells > 0 && <> · Frais −{fmtEUR(r.feesOnSells, 0)}</>}
                      {r.taxCTO > 0 && <> · Tax −{fmtEUR(r.taxCTO, 0)}</>}
                    </div>
                  )}
                </div>
              </button>
            );
          })}
        </div>
      </div>
      {openBucket && <GainsDetailModal data={openBucket} onClose={() => setOpenBucket(null)} hide={hide} />}
    </>
  );
};

// Modal de détail des gains pour une période donnée
const GainsDetailModal = ({ data, onClose, hide }) => {
  // Empêche le scroll body en arrière-plan
  useEffect(() => {
    const orig = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => { document.body.style.overflow = orig; };
  }, []);
  const { trades, divsList, intsList, svbList, brut, feesOnSells, taxCTO, net, afterFees, divs, ints, svb, bucketLabel, from, to, realizedBrut, realizedAfterFees } = data;
  return (
    <div className="fixed inset-0 z-50 flex items-end md:items-center justify-center bg-black/70 backdrop-blur-sm p-0 md:p-4" onClick={onClose}>
      <div className="bg-slate-900 border border-slate-700 rounded-t-2xl md:rounded-2xl w-full md:max-w-3xl max-h-[92vh] overflow-hidden flex flex-col" onClick={e => e.stopPropagation()}>
        {/* Header */}
        <div className={`p-5 border-b border-slate-800 ${net >= 0 ? "bg-gradient-to-br from-emerald-500/10 to-transparent" : "bg-gradient-to-br from-rose-500/10 to-transparent"}`}>
          <div className="flex items-start justify-between gap-3">
            <div className="min-w-0">
              <div className="text-[10px] uppercase tracking-widest text-slate-400 mb-1">📊 Détail · {bucketLabel}</div>
              <div className={`text-3xl md:text-4xl font-mono font-black tabular-nums ${classFor(net)}`}>
                {mask((net >= 0 ? "+" : "") + fmtEUR(net, 2).replace("+", ""), hide)}
              </div>
              <div className="text-xs text-slate-400 mt-1 font-mono">{fmtDateShort(from)} → {fmtDateShort(to)}</div>
            </div>
            <button onClick={onClose} className="p-2 bg-slate-800 hover:bg-slate-700 rounded text-slate-300 transition-colors text-xl leading-none">✕</button>
          </div>
          {/* Brut → Frais → Impôts → Net */}
          <div className="mt-4 grid grid-cols-2 md:grid-cols-4 gap-2">
            <div className="bg-slate-950/60 rounded-lg p-2.5 border border-slate-800">
              <div className="text-[9px] uppercase tracking-widest text-slate-500">Brut</div>
              <div className={`text-sm md:text-base font-mono font-bold tabular-nums ${classFor(brut)}`}>{mask(fmtEUR(brut, 2), hide)}</div>
            </div>
            <div className="bg-slate-950/60 rounded-lg p-2.5 border border-slate-800">
              <div className="text-[9px] uppercase tracking-widest text-slate-500">− Frais</div>
              <div className="text-sm md:text-base font-mono font-bold text-amber-400 tabular-nums">−{mask(fmtEUR(feesOnSells, 2), hide)}</div>
              <div className="text-[9px] text-slate-600 mt-0.5">{trades.length} × 1€</div>
            </div>
            <div className="bg-slate-950/60 rounded-lg p-2.5 border border-slate-800">
              <div className="text-[9px] uppercase tracking-widest text-slate-500">− Impôts CTO</div>
              <div className="text-sm md:text-base font-mono font-bold text-rose-400 tabular-nums">−{mask(fmtEUR(taxCTO, 2), hide)}</div>
              <div className="text-[9px] text-slate-600 mt-0.5">31,4% PFU</div>
            </div>
            <div className={`rounded-lg p-2.5 border-2 ${net >= 0 ? "bg-emerald-500/10 border-emerald-500/40" : "bg-rose-500/10 border-rose-500/40"}`}>
              <div className="text-[9px] uppercase tracking-widest text-slate-500">= Net réel</div>
              <div className={`text-sm md:text-base font-mono font-black tabular-nums ${classFor(net)}`}>{mask(fmtEUR(net, 2), hide)}</div>
            </div>
          </div>
        </div>
        {/* Body scrollable */}
        <div className="overflow-y-auto p-5 space-y-4 flex-1">
          {trades.length > 0 && (
            <div>
              <div className="text-[10px] uppercase tracking-widest text-slate-500 font-medium mb-2 flex items-center gap-2">
                <Activity className="w-3 h-3 text-emerald-400" /> Ventes ({trades.length}) · {classFor(realizedAfterFees) === "text-emerald-400" ? "+" : ""}{fmtEUR(realizedAfterFees, 2)} après frais
              </div>
              <div className="space-y-1.5">
                {[...trades].sort((a, b) => b.pnl - a.pnl).map((t, i) => {
                  const feeShare = 1; // 1€ par vente
                  const afterFees = t.pnl - feeShare;
                  // Taxe CTO : seulement si gain positif sur compte CTO (pas PEA)
                  const isPEA = t.accountType === "PEA";
                  const taxOnTrade = (afterFees > 0 && !isPEA) ? afterFees * 0.314 : 0;
                  const pnlNetReal = afterFees - taxOnTrade;
                  return (
                    <div key={i} className={`flex items-center gap-3 p-3 rounded-lg border bg-slate-950/60 ${t.pnl >= 0 ? "border-emerald-500/20" : "border-rose-500/20"}`}>
                      <div className={`w-1 h-12 rounded-full shrink-0 ${t.pnl >= 0 ? "bg-emerald-500" : "bg-rose-500"}`} />
                      <div className="flex-1 min-w-0">
                        <div className="text-sm font-semibold truncate">{t.name || t.symbol}</div>
                        <div className="flex items-center gap-2 text-[10px] text-slate-500 font-mono flex-wrap">
                          <span>{fmtDateShort(t.sellDate)}</span>
                          <span>·</span>
                          <span>{fmtNum(t.shares, 4)} parts</span>
                          <span>·</span>
                          <span>{t.holdingDays}j détenus</span>
                          {t.lotCount > 1 && <><span>·</span><span className="text-cyan-400">{t.lotCount} lots FIFO</span></>}
                        </div>
                        <div className="text-[10px] text-slate-600 font-mono mt-0.5">
                          Acheté {fmtEUR(t.buyPrice, 4)} → Vendu {fmtEUR(t.sellPrice, 4)}
                        </div>
                      </div>
                      <div className="text-right shrink-0">
                        <div className={`text-sm font-mono font-bold tabular-nums ${classFor(t.pnl)}`}>
                          {mask((t.pnl >= 0 ? "+" : "") + fmtEUR(t.pnl, 2).replace("+", ""), hide)}
                        </div>
                        <div className={`text-[9px] font-mono ${classFor(t.pnl)}`}>{fmtPct(t.pnlPct, 1)} brut</div>
                        <div className="text-[9px] font-mono text-slate-500 mt-0.5">net {fmtEUR(pnlNetReal, 2)}{isPEA && " (PEA)"}</div>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}
          {divsList.length > 0 && (
            <div>
              <div className="text-[10px] uppercase tracking-widest text-slate-500 font-medium mb-2 flex items-center gap-2">
                <Coins className="w-3 h-3 text-amber-400" /> Dividendes ({divsList.length}) · +{fmtEUR(divs, 2)} net
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
                {divsList.map((dv, i) => (
                  <div key={i} className="flex items-center justify-between p-2 bg-slate-950/60 rounded border border-amber-500/20">
                    <div className="min-w-0 flex-1">
                      <div className="text-xs font-semibold truncate">💰 {dv.name}</div>
                      <div className="text-[10px] text-slate-500 font-mono">{fmtDateShort(dv.date)}</div>
                    </div>
                    <div className="text-right">
                      <div className="text-sm font-mono font-bold text-amber-300 tabular-nums">{mask(fmtEUR(dv.net, 2), hide)}</div>
                      <div className="text-[9px] text-slate-600 font-mono">brut {fmtEUR(dv.gross, 2)}</div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
          {intsList.length > 0 && (
            <div>
              <div className="text-[10px] uppercase tracking-widest text-slate-500 font-medium mb-2 flex items-center gap-2">
                <Percent className="w-3 h-3 text-cyan-400" /> Intérêts ({intsList.length}) · +{fmtEUR(ints, 2)} net
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
                {intsList.map((it, i) => (
                  <div key={i} className="flex items-center justify-between p-2 bg-slate-950/60 rounded border border-cyan-500/20">
                    <div className="text-xs font-mono text-slate-400">{fmtDateShort(it.date)}</div>
                    <div className="text-right">
                      <div className="text-sm font-mono font-bold text-cyan-300 tabular-nums">{mask(fmtEUR(it.net, 2), hide)}</div>
                      <div className="text-[9px] text-slate-600 font-mono">brut {fmtEUR(it.gross, 2)}</div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
          {svbList.length > 0 && (
            <div>
              <div className="text-[10px] uppercase tracking-widest text-slate-500 font-medium mb-2 flex items-center gap-2">
                <Gift className="w-3 h-3 text-violet-400" /> Saveback ({svbList.length}) · +{fmtEUR(svb, 2)}
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
                {svbList.map((sb, i) => (
                  <div key={i} className="flex items-center justify-between p-2 bg-slate-950/60 rounded border border-violet-500/20">
                    <div className="min-w-0 flex-1">
                      <div className="text-xs font-semibold truncate">🎁 {sb.name}</div>
                      <div className="text-[10px] text-slate-500 font-mono">{fmtDateShort(sb.date)}</div>
                    </div>
                    <div className="text-sm font-mono font-bold text-violet-300 tabular-nums">{mask(fmtEUR(sb.amount, 2), hide)}</div>
                  </div>
                ))}
              </div>
            </div>
          )}
          {trades.length === 0 && divsList.length === 0 && intsList.length === 0 && svbList.length === 0 && (
            <div className="text-center text-slate-500 text-sm py-8">Aucune activité sur cette période</div>
          )}
        </div>
      </div>
    </div>
  );
};

// Meilleures positions : tableau complet par actif avec total gains (réalisés + latents + divs touchés)
// ═══════ Comportement d'investissement ═══════
// ═══════ Score de santé du portefeuille (0-100) ═══════
const HealthScoreCard = ({ d, hide }) => {
  const score = d.healthScore ?? 0;
  const comps = d.healthComponents || [];
  if (!comps.length) return null;
  const scoreColor = score >= 75 ? "#10b981" : score >= 50 ? "#f59e0b" : "#f43f5e";
  const scoreLabel = score >= 75 ? "Excellent" : score >= 60 ? "Bon" : score >= 40 ? "Moyen" : "À améliorer";
  // Donut SVG simple (pas de Recharts pour un gauge léger)
  const R = 52, C = 2 * Math.PI * R;
  const filled = (score / 100) * C * 0.75; // arc de 270°
  return (
    <Card>
      <SectionTitle icon={Gauge} subtitle="Score composite : diversification, win rate, frais, régularité, profit factor">🩺 Santé du portefeuille</SectionTitle>
      <div className="flex items-center gap-5 flex-wrap">
        <div className="relative w-32 h-32 shrink-0">
          <svg viewBox="0 0 128 128" className="w-full h-full -rotate-[225deg]">
            <circle cx="64" cy="64" r={R} fill="none" stroke="#1e293b" strokeWidth="10" strokeDasharray={`${C * 0.75} ${C}`} strokeLinecap="round" />
            <circle cx="64" cy="64" r={R} fill="none" stroke={scoreColor} strokeWidth="10" strokeDasharray={`${filled} ${C}`} strokeLinecap="round" />
          </svg>
          <div className="absolute inset-0 flex flex-col items-center justify-center">
            <div className="text-3xl font-mono font-black tabular-nums" style={{ color: scoreColor }}>{score}</div>
            <div className="text-[9px] uppercase tracking-widest text-slate-500">/100</div>
          </div>
        </div>
        <div className="flex-1 min-w-[220px] space-y-2">
          <div className="text-sm font-bold" style={{ color: scoreColor }}>{scoreLabel}</div>
          {comps.map((c, i) => (
            <div key={i}>
              <div className="flex justify-between text-[11px] mb-0.5">
                <span className="text-slate-400">{c.label}</span>
                <span className="font-mono text-slate-300">{c.score.toFixed(0)}/{c.max} <span className="text-slate-600">· {c.detail}</span></span>
              </div>
              <div className="h-1.5 bg-slate-800 rounded-full overflow-hidden">
                <div className="h-full rounded-full transition-all" style={{ width: `${(c.score / c.max) * 100}%`, background: c.score / c.max >= 0.75 ? "#10b981" : c.score / c.max >= 0.5 ? "#f59e0b" : "#f43f5e" }} />
              </div>
            </div>
          ))}
        </div>
      </div>
    </Card>
  );
};

// ═══════ XIRR : rendement annualisé ═══════
const XirrCard = ({ d, hide }) => {
  const xirr = d.xirr;
  const b = d.behaviorStats;
  return (
    <Card>
      <SectionTitle icon={Zap} subtitle="Taux de rendement interne : la performance annualisée de TON argent, flux par flux">⚡ Rendement annualisé (XIRR)</SectionTitle>
      {xirr !== null && isFinite(xirr) ? (
        <div className="flex items-center gap-6 flex-wrap">
          <div>
            <div className={`text-5xl font-mono font-black tabular-nums ${classFor(xirr)}`}>{xirr >= 0 ? "+" : ""}{xirr.toFixed(1)}%</div>
            <div className="text-xs text-slate-500 font-mono mt-1">par an, net de tes flux réels</div>
          </div>
          <div className="flex-1 min-w-[200px] space-y-2 text-xs">
            <div className="flex items-start gap-2 text-slate-400">
              <Info className="w-3.5 h-3.5 shrink-0 mt-0.5 text-cyan-400" />
              <span>Le XIRR pondère chaque euro par son temps de présence. C'est la métrique la plus juste pour comparer ta perf à un livret ({(3).toFixed(0)}%) ou au S&P 500 (~10%/an historique).</span>
            </div>
            <div className="grid grid-cols-2 gap-2 mt-2">
              <div className={`rounded-lg p-2 border ${xirr > 3 ? "bg-emerald-500/10 border-emerald-500/30" : "bg-slate-900 border-slate-800"}`}>
                <div className="text-[9px] uppercase tracking-widest text-slate-500">vs Livret A (3%)</div>
                <div className={`text-sm font-mono font-bold ${xirr > 3 ? "text-emerald-400" : "text-rose-400"}`}>{xirr > 3 ? "✓ Tu fais mieux" : "✗ En dessous"}</div>
              </div>
              <div className={`rounded-lg p-2 border ${xirr > 10 ? "bg-emerald-500/10 border-emerald-500/30" : "bg-slate-900 border-slate-800"}`}>
                <div className="text-[9px] uppercase tracking-widest text-slate-500">vs S&P 500 (~10%)</div>
                <div className={`text-sm font-mono font-bold ${xirr > 10 ? "text-emerald-400" : "text-amber-400"}`}>{xirr > 10 ? "✓ Tu bats le marché" : "≈ Marché"}</div>
              </div>
            </div>
          </div>
        </div>
      ) : (
        <div className="text-sm text-slate-500 py-4">
          Pas encore calculable — il faut des cours saisis sur tes positions ouvertes pour valoriser le portefeuille.
          <div className="text-[11px] text-slate-600 mt-1">Va dans le camembert sectoriel → clique un secteur → saisis les cours.</div>
        </div>
      )}
    </Card>
  );
};

// ═══════ Heatmap calendrier des gains (style GitHub) ═══════
const GainsHeatmap = ({ d, hide }) => {
  const daily = d.dailyGains || {};
  const entries = Object.keys(daily);
  if (entries.length === 0) return null;
  // 26 dernières semaines (~6 mois)
  const today = new Date();
  const weeks = [];
  const start = new Date(today);
  start.setDate(start.getDate() - start.getDay() + 1 - 25 * 7); // lundi il y a 25 semaines
  for (let w = 0; w < 26; w++) {
    const days = [];
    for (let dd = 0; dd < 7; dd++) {
      const cur = new Date(start);
      cur.setDate(start.getDate() + w * 7 + dd);
      const k = `${cur.getFullYear()}-${String(cur.getMonth() + 1).padStart(2, "0")}-${String(cur.getDate()).padStart(2, "0")}`;
      days.push({ key: k, date: new Date(cur), gain: daily[k] || 0, future: cur > today });
    }
    weeks.push(days);
  }
  // Échelle de couleurs à paliers : on calcule les seuils sur les valeurs réelles (gains et pertes séparément)
  const posVals = Object.values(daily).filter(v => v > 0).sort((a, b) => a - b);
  const negVals = Object.values(daily).filter(v => v < 0).map(Math.abs).sort((a, b) => a - b);
  const quantile = (arr, q) => arr.length ? arr[Math.min(arr.length - 1, Math.floor(q * arr.length))] : 0;
  // Paliers gains : du vert sombre/pâle au vert flashy
  const GREENS = ["#0e3d2e", "#10b981", "#34d399", "#00e676", "#00ff88"];
  const REDS = ["#3d1420", "#e11d48", "#f43f5e", "#ff1744", "#ff0040"];
  const posThresh = [quantile(posVals, 0.2), quantile(posVals, 0.45), quantile(posVals, 0.7), quantile(posVals, 0.9)];
  const negThresh = [quantile(negVals, 0.2), quantile(negVals, 0.45), quantile(negVals, 0.7), quantile(negVals, 0.9)];
  const levelFor = (val, thresh) => { for (let i = thresh.length - 1; i >= 0; i--) { if (val >= thresh[i]) return i + 1; } return 0; };
  const cellColor = (g, future) => {
    if (future) return "transparent";
    if (g === 0) return "#1e293b";
    if (g > 0) return GREENS[levelFor(g, posThresh)];
    return REDS[levelFor(Math.abs(g), negThresh)];
  };
  const [tooltip, setTooltip] = useState(null);
  const monthLabels = [];
  let lastMonth = -1;
  weeks.forEach((wk, i) => {
    const m = wk[0].date.getMonth();
    if (m !== lastMonth) { monthLabels.push({ idx: i, label: ["Jan", "Fév", "Mar", "Avr", "Mai", "Juin", "Juil", "Août", "Sep", "Oct", "Nov", "Déc"][m] }); lastMonth = m; }
  });
  return (
    <Card>
      <SectionTitle icon={CalendarDays} subtitle="Chaque case = un jour. Vert = gains encaissés, rouge = pertes. Plus c'est foncé, plus c'est gros.">🗓 Calendrier des gains (6 derniers mois)</SectionTitle>
      <div className="overflow-x-auto pb-2">
        <div className="inline-block min-w-full">
          <div className="flex gap-[3px] mb-1 ml-8">
            {weeks.map((_, i) => {
              const ml = monthLabels.find(m => m.idx === i);
              return <div key={i} className="w-[14px] text-[8px] text-slate-500 font-mono">{ml ? ml.label : ""}</div>;
            })}
          </div>
          <div className="flex gap-1">
            <div className="flex flex-col gap-[3px] mr-1">
              {["L", "M", "M", "J", "V", "S", "D"].map((l, i) => <div key={i} className="h-[14px] w-6 text-[8px] text-slate-500 font-mono flex items-center">{l}</div>)}
            </div>
            <div className="flex gap-[3px]">
              {weeks.map((wk, wi) => (
                <div key={wi} className="flex flex-col gap-[3px]">
                  {wk.map((day, di) => (
                    <div key={di}
                      className="w-[14px] h-[14px] rounded-[3px] cursor-pointer hover:ring-1 hover:ring-cyan-400 transition-shadow"
                      style={{ background: cellColor(day.gain, day.future) }}
                      onMouseEnter={() => !day.future && setTooltip(day)}
                      onMouseLeave={() => setTooltip(null)}
                      onClick={() => !day.future && setTooltip(tooltip?.key === day.key ? null : day)}
                    />
                  ))}
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
      <div className="flex items-center justify-between mt-2 flex-wrap gap-2">
        <div className="flex items-center gap-2 text-[9px] text-slate-500 font-mono">
          <span>Grosse perte</span>
          <div className="flex gap-[2px]">
            <span className="w-3 h-3 rounded-sm" style={{ background: "#ff0040" }} />
            <span className="w-3 h-3 rounded-sm" style={{ background: "#ff1744" }} />
            <span className="w-3 h-3 rounded-sm" style={{ background: "#f43f5e" }} />
            <span className="w-3 h-3 rounded-sm" style={{ background: "#e11d48" }} />
            <span className="w-3 h-3 rounded-sm" style={{ background: "#3d1420" }} />
            <span className="w-3 h-3 rounded-sm bg-slate-800" />
            <span className="w-3 h-3 rounded-sm" style={{ background: "#0e3d2e" }} />
            <span className="w-3 h-3 rounded-sm" style={{ background: "#10b981" }} />
            <span className="w-3 h-3 rounded-sm" style={{ background: "#34d399" }} />
            <span className="w-3 h-3 rounded-sm" style={{ background: "#00e676" }} />
            <span className="w-3 h-3 rounded-sm" style={{ background: "#00ff88" }} />
          </div>
          <span>Gros gain</span>
        </div>
        {tooltip && (
          <div className="text-xs font-mono bg-slate-950 border border-slate-700 rounded px-2.5 py-1">
            <span className="text-slate-400">{fmtDateShort(tooltip.date)}</span>
            <span className={`ml-2 font-bold ${classFor(tooltip.gain)}`}>{tooltip.gain === 0 ? "—" : mask((tooltip.gain > 0 ? "+" : "") + fmtEUR(tooltip.gain, 2).replace("+", ""), hide)}</span>
          </div>
        )}
      </div>
    </Card>
  );
};

// ═══════ Records & milestones ═══════
const RecordsSection = ({ d, hide }) => {
  const r = d.records;
  if (!r || r.totalActiveDays === 0) return null;
  const fmtDay = (k) => { if (!k) return "—"; const [y, m, dd] = k.split("-"); return `${dd}/${m}/${y.slice(2)}`; };
  const items = [
    r.bestDay && { emoji: "🏆", label: "Meilleur jour", value: mask("+" + fmtEUR(r.bestDay.gain, 0).replace("+", ""), hide), sub: fmtDay(r.bestDay.day), color: "emerald" },
    r.bestWeek && { emoji: "🚀", label: "Meilleure semaine", value: mask("+" + fmtEUR(r.bestWeek.gain, 0).replace("+", ""), hide), sub: `sem. du ${fmtDay(r.bestWeek.week)}`, color: "emerald" },
    r.worstDay && r.worstDay.gain < 0 && { emoji: "💥", label: "Pire jour", value: mask(fmtEUR(r.worstDay.gain, 0), hide), sub: fmtDay(r.worstDay.day), color: "rose" },
    { emoji: "🔥", label: "Série de jours gagnants", value: `${r.maxPosStreak}`, sub: "jours consécutifs", color: "amber" },
    r.maxCovStreak > 0 && { emoji: "🛡", label: "Mois couverts d'affilée", value: `${r.maxCovStreak}`, sub: "gains ≥ dépenses CB", color: "cyan" },
    { emoji: "📈", label: "Jours profitables", value: `${r.profitableDays}/${r.totalActiveDays}`, sub: `${((r.profitableDays / r.totalActiveDays) * 100).toFixed(0)}% des jours actifs`, color: "violet" },
  ].filter(Boolean);
  return (
    <Card>
      <SectionTitle icon={Trophy} subtitle="Tes meilleurs moments et tes séries en cours">🏅 Records & milestones</SectionTitle>
      <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3">
        {items.map((it, i) => {
          const colors = { emerald: "border-emerald-500/30 bg-emerald-500/5 text-emerald-300", rose: "border-rose-500/30 bg-rose-500/5 text-rose-300", amber: "border-amber-500/30 bg-amber-500/5 text-amber-300", cyan: "border-cyan-500/30 bg-cyan-500/5 text-cyan-300", violet: "border-violet-500/30 bg-violet-500/5 text-violet-300" };
          return (
            <div key={i} className={`rounded-xl border p-3 ${colors[it.color] || colors.cyan}`}>
              <div className="text-xl mb-1">{it.emoji}</div>
              <div className="text-[9px] uppercase tracking-widest text-slate-500">{it.label}</div>
              <div className="text-lg font-mono font-black tabular-nums">{it.value}</div>
              <div className="text-[9px] text-slate-500 font-mono mt-0.5">{it.sub}</div>
            </div>
          );
        })}
      </div>
    </Card>
  );
};

const BehaviorSection = ({ d, hide }) => {
  const b = d.behaviorStats;
  const [mode, setMode] = useState("buy"); // buy | sell
  if (!b || b.totalBuys === 0) return null;
  const isBuy = mode === "buy";
  const dayData = isBuy ? b.buysByDayData : b.sellsByDayData;
  const monthData = isBuy ? b.buysByMonthData : b.sellsByMonthData;
  const maxDay = Math.max(...dayData.map(x => x.count), 1);
  const maxMonth = Math.max(...monthData.map(x => x.count), 1);
  const accent = isBuy ? "#a855f7" : "#10b981";
  const accentBar = isBuy ? "#06b6d4" : "#10b981";
  const perWeek = isBuy ? b.buysPerWeek : b.sellsPerWeek;
  const capital = isBuy ? b.capitalPerWeek : b.capitalSoldPerWeek;
  const avgAmount = isBuy ? b.avgBuyAmount : b.avgSellAmount;
  const favDay = isBuy ? b.favoriteDay : b.favoriteSellDay;
  const total = isBuy ? b.totalBuys : b.totalSells;
  const verb = isBuy ? "achat" : "vente";
  return (
    <Card>
      <div className="flex items-center justify-between flex-wrap gap-3 mb-1">
        <SectionTitle icon={Activity} subtitle="Comprendre ton style et ta régularité d'investisseur">🧠 Mon comportement</SectionTitle>
        <div className="inline-flex rounded-lg border border-slate-700 overflow-hidden shrink-0">
          <button onClick={() => setMode("buy")} className={`px-4 py-1.5 text-xs font-semibold transition-colors ${isBuy ? "bg-cyan-500/20 text-cyan-200" : "bg-slate-950 text-slate-400 hover:text-slate-100"}`}>🛒 Achats</button>
          <button onClick={() => setMode("sell")} className={`px-4 py-1.5 text-xs font-semibold transition-colors ${!isBuy ? "bg-emerald-500/20 text-emerald-200" : "bg-slate-950 text-slate-400 hover:text-slate-100"}`}>💸 Ventes</button>
        </div>
      </div>

      {/* KPIs comportement */}
      <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3 mb-5">
        <div className="bg-slate-950/40 rounded-lg p-3 border border-slate-800">
          <div className="text-[9px] uppercase tracking-widest text-slate-500">Rythme de {verb}</div>
          <div className="text-xl font-mono font-black text-cyan-300 tabular-nums">{perWeek.toFixed(1)}</div>
          <div className="text-[9px] text-slate-500 font-mono">ordres / semaine</div>
        </div>
        <div className="bg-slate-950/40 rounded-lg p-3 border border-slate-800">
          <div className="text-[9px] uppercase tracking-widest text-slate-500">{isBuy ? "Capital investi" : "Capital vendu"} /sem.</div>
          <div className="text-xl font-mono font-black text-emerald-300 tabular-nums">{mask(fmtEURCompact(capital), hide)}</div>
          <div className="text-[9px] text-slate-500 font-mono">en moyenne</div>
        </div>
        <div className="bg-slate-950/40 rounded-lg p-3 border border-slate-800">
          <div className="text-[9px] uppercase tracking-widest text-slate-500">{verb} moyen{isBuy ? "" : "ne"}</div>
          <div className="text-xl font-mono font-black text-slate-100 tabular-nums">{mask(fmtEUR(avgAmount, 0), hide)}</div>
          <div className="text-[9px] text-slate-500 font-mono">par ordre</div>
        </div>
        <div className="bg-slate-950/40 rounded-lg p-3 border border-slate-800">
          <div className="text-[9px] uppercase tracking-widest text-slate-500">Jour préféré</div>
          <div className="text-xl font-bold text-violet-300">{favDay}</div>
          <div className="text-[9px] text-slate-500 font-mono">pour {isBuy ? "acheter" : "vendre"}</div>
        </div>
        <div className="bg-slate-950/40 rounded-lg p-3 border border-slate-800">
          <div className="text-[9px] uppercase tracking-widest text-slate-500">Ratio achat/vente</div>
          <div className="text-xl font-mono font-black text-amber-300 tabular-nums">{b.buyToSellRatio.toFixed(1)}×</div>
          <div className="text-[9px] text-slate-500 font-mono">{b.totalBuys} achats · {b.totalSells} ventes</div>
        </div>
        <div className="bg-slate-950/40 rounded-lg p-3 border border-slate-800">
          <div className="text-[9px] uppercase tracking-widest text-slate-500">Ancienneté</div>
          <div className="text-xl font-mono font-black text-cyan-300 tabular-nums">{Math.floor(b.investingDays / 30)}</div>
          <div className="text-[9px] text-slate-500 font-mono">mois d'activité</div>
        </div>
      </div>

      {/* Graphiques jour/mois */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div>
          <div className="text-[10px] uppercase tracking-widest text-slate-500 mb-2">{isBuy ? "Achats" : "Ventes"} par jour de la semaine</div>
          <div className="space-y-1.5">
            {dayData.map((day, i) => (
              <div key={i} className="flex items-center gap-2">
                <span className="text-[10px] text-slate-400 font-mono w-8">{day.day}</span>
                <div className="flex-1 h-5 bg-slate-800 rounded overflow-hidden relative">
                  <div className="h-full rounded transition-all" style={{ width: `${(day.count / maxDay) * 100}%`, background: day.count === maxDay && day.count > 0 ? accent : "#475569" }} />
                  <span className="absolute right-2 top-1/2 -translate-y-1/2 text-[10px] font-mono text-slate-300">{day.count}</span>
                </div>
              </div>
            ))}
          </div>
        </div>
        <div>
          <div className="text-[10px] uppercase tracking-widest text-slate-500 mb-2">{isBuy ? "Achats" : "Ventes"} par mois de l'année</div>
          <ResponsiveContainer width="100%" height={160}>
            <BarChart data={monthData}>
              <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" vertical={false} />
              <XAxis dataKey="month" stroke="#64748b" fontSize={9} interval={0} />
              <YAxis stroke="#64748b" fontSize={9} allowDecimals={false} />
              <Tooltip content={<ChartTooltip formatter={(v) => `${v} ${isBuy ? "achats" : "ventes"}`} />} />
              <Bar dataKey="count" radius={[3, 3, 0, 0]}>
                {monthData.map((m, i) => <Cell key={i} fill={m.count === maxMonth && m.count > 0 ? accentBar : "#334155"} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </Card>
  );
};

const MeilleuresPositions = ({ d, sectorFn, hide }) => {
  const { byAsset, dividends } = d;
  // Grouper les dividendes par actif
  const divsByAsset = {};
  for (const dv of dividends) { const k = dv.symbol || dv.name; divsByAsset[k] = (divsByAsset[k] || 0) + dv.net; }
  const enriched = byAsset.map(a => ({
    ...a,
    divReceived: divsByAsset[a.symbol || a.name] || 0,
    totalGain: a.totalPnL + (divsByAsset[a.symbol || a.name] || 0),
  }));
  // Meilleures = celles qui ont généré du cash (réalisé+divs) OU forte PV latente
  const best = [...enriched].sort((a, b) => b.totalGain - a.totalGain).slice(0, 10);
  const worst = [...enriched].sort((a, b) => a.totalGain - b.totalGain).slice(0, 10);
  const openBest = enriched.filter(a => a.openShares > 0 && a.invested > 0).sort((a, b) => b.returnPct - a.returnPct).slice(0, 8);
  const openWorst = enriched.filter(a => a.openShares > 0 && a.invested > 0).sort((a, b) => a.returnPct - b.returnPct).slice(0, 8);

  return (
    <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
      <Card>
        <SectionTitle icon={Award} subtitle="Positions où tu as déjà le plus gagné (réalisé + latent + dividendes)">💎 Meilleurs actifs all-time</SectionTitle>
        <PositionRows items={best} sectorFn={sectorFn} hide={hide} />
      </Card>
      <Card>
        <SectionTitle icon={TrendingDown} subtitle="Positions où tu as le plus perdu (cumulé)">⚠️ Pires actifs all-time</SectionTitle>
        <PositionRows items={worst} sectorFn={sectorFn} hide={hide} />
      </Card>
      {openBest.length > 0 && (
        <Card>
          <SectionTitle icon={TrendingUp} subtitle="Tes positions ouvertes les plus rentables (nécessite cours saisi)">🚀 Positions ouvertes gagnantes</SectionTitle>
          {openBest.filter(a => a.currentValue !== a.invested).length === 0 ? (
            <div className="text-xs text-amber-400/80 bg-amber-500/10 border border-amber-500/30 rounded p-3 flex items-start gap-2">
              <Info className="w-4 h-4 shrink-0 mt-0.5" />
              <span>Saisis les cours actuels dans l'onglet Positions pour voir ici les positions qui performent le mieux.</span>
            </div>
          ) : (
            <PositionRows items={openBest.filter(a => a.currentValue !== a.invested)} sectorFn={sectorFn} hide={hide} showReturn />
          )}
        </Card>
      )}
      {openWorst.length > 0 && openWorst.filter(a => a.currentValue !== a.invested).length > 0 && (
        <Card>
          <SectionTitle icon={AlertCircle} subtitle="Tes positions ouvertes en perte">📉 Positions ouvertes perdantes</SectionTitle>
          <PositionRows items={openWorst.filter(a => a.currentValue !== a.invested)} sectorFn={sectorFn} hide={hide} showReturn />
        </Card>
      )}
    </div>
  );
};

const PositionRows = ({ items, sectorFn, hide, showReturn }) => (
  <div className="space-y-1.5">
    {items.length === 0 && <div className="text-slate-500 text-sm text-center py-6">Aucune donnée</div>}
    {items.map((a, i) => {
      const sec = sectorFn(a);
      const gain = a.totalGain !== undefined ? a.totalGain : a.totalPnL;
      return (
        <div key={i} className={`flex items-center gap-3 px-3 py-2.5 bg-slate-950/60 rounded-lg border border-slate-800/60 ${ROW_HOVER}`}>
          <div className="w-1 h-12 rounded-full shrink-0" style={{ background: SECTOR_COLORS[sec] || "#64748b" }} />
          <div className="flex-1 min-w-0">
            <div className="text-sm font-semibold truncate">{a.name || a.symbol}</div>
            <div className="flex items-center gap-2 text-[10px] text-slate-500 font-mono flex-wrap">
              <span>{sec}</span>
              {a.trades > 0 && <><span>·</span><span>{a.trades} trade{a.trades > 1 ? "s" : ""}</span></>}
              {a.openShares > 0 && <><span>·</span><span className="text-cyan-400">ouvert</span></>}
              {a.divReceived > 0 && <><span>·</span><span className="text-amber-400">+{fmtEUR(a.divReceived, 0)} div</span></>}
            </div>
          </div>
          <div className="text-right shrink-0">
            <div className={`text-sm font-mono font-bold tabular-nums ${classFor(gain)}`}>{mask(fmtEUR(gain, 0), hide)}</div>
            {showReturn && a.invested > 0 ? (
              <div className={`text-[10px] font-mono ${classFor(a.returnPct)}`}>{fmtPct(a.returnPct, 1)} latent</div>
            ) : a.trades > 0 && a.totalPnL !== 0 && (
              <div className={`text-[10px] font-mono ${classFor(a.totalPnL)}`}>
                {mask(fmtEUR(a.realizedPnL, 0), hide)} réalisé
              </div>
            )}
          </div>
        </div>
      );
    })}
  </div>
);

const PnLBreakdown = ({ label, value, pct, icon: Icon, hide, accent, sub }) => (
  <div className="flex items-center gap-3">
    <div className={`w-9 h-9 rounded-lg flex items-center justify-center ${accent === "amber" ? "bg-amber-500/15" : accent === "violet" ? "bg-violet-500/15" : value >= 0 ? "bg-emerald-500/15" : "bg-rose-500/15"}`}>
      <Icon className={`w-4 h-4 ${accent === "amber" ? "text-amber-400" : accent === "violet" ? "text-violet-400" : value >= 0 ? "text-emerald-400" : "text-rose-400"}`} />
    </div>
    <div className="min-w-0 flex-1">
      <div className="text-[10px] uppercase tracking-widest text-slate-500">{label}</div>
      <div className={`text-base font-mono font-bold tabular-nums ${accent === "amber" ? "text-amber-300" : accent === "violet" ? "text-violet-300" : classFor(value)}`}>{mask(fmtEUR(value, 2), hide)}</div>
      {pct !== null && pct !== undefined && <div className={`text-[10px] font-mono ${classFor(pct)}`}>{fmtPct(pct, 1)}</div>}
      {sub && <div className="text-[10px] font-mono text-slate-500">{sub}</div>}
    </div>
  </div>
);

const TradeList = ({ trades, sectorFn, hide }) => (
  <div className="space-y-1.5">
    {trades.length === 0 && <div className="text-slate-500 text-sm text-center py-6">Aucun trade</div>}
    {trades.map((t, i) => {
      const sec = sectorFn(t);
      return (
        <div key={i} className={`flex items-center gap-3 px-3 py-2 bg-slate-950/60 rounded-lg border border-slate-800/60 ${ROW_HOVER}`}>
          <div className="w-1 h-10 rounded-full shrink-0" style={{ background: SECTOR_COLORS[sec] || "#64748b" }} />
          <div className="flex-1 min-w-0">
            <div className="text-sm font-semibold truncate">{t.name || t.symbol}</div>
            <div className="flex items-center gap-2 text-[10px] text-slate-500 font-mono flex-wrap">
              <span>{fmtDateShort(t.buyDate)} → {fmtDateShort(t.sellDate)}</span><span>·</span><span>{t.holdingDays}j</span><span>·</span><span>{sec}</span>
              {t.lotCount > 1 && <><span>·</span><span className="text-cyan-500/70">{t.lotCount} lots</span></>}
            </div>
          </div>
          <div className="text-right shrink-0">
            <div className={`text-sm font-mono font-bold ${classFor(t.pnl)}`}>{mask(fmtEUR(t.pnl, 0), hide)}</div>
            <div className={`text-[10px] font-mono ${classFor(t.pnl)}`}>{fmtPct(t.pnlPct, 1)}{t.annualizedPct !== null && t.annualizedPct !== undefined && <span className="text-slate-600"> · {t.annualizedPct >= 0 ? "+" : ""}{t.annualizedPct.toFixed(0)}%/an</span>}</div>
          </div>
        </div>
      );
    })}
  </div>
);

// ═══════════════ POSITIONS (refonte) ═══════════════

function Positions({ d, prices, onPriceChange, sectorFn, hide }) {
  const [sort, setSort] = useState({ key: "invested", dir: "desc" });
  const [editing, setEditing] = useState(null);
  const [editVal, setEditVal] = useState("");
  const [viewMode, setViewMode] = useState("cards");

  const positions = useMemo(() => {
    const m = {};
    for (const l of d.openLots) {
      const k = l.symbol || l.name;
      if (!m[k]) m[k] = { symbol: l.symbol, name: l.name, shares: 0, invested: 0, oldestDate: l.date, lots: [], assetClass: l.ref?.assetClass || "" };
      m[k].shares += l.shares; m[k].invested += l.shares * l.costPerShare;
      if (l.date < m[k].oldestDate) m[k].oldestDate = l.date;
      m[k].lots.push(l);
    }
    return Object.values(m).map(p => {
      const px = prices[p.symbol] || prices[p.name];
      const avg = p.shares ? p.invested / p.shares : 0;
      const val = px ? p.shares * px : null;
      const pnl = val !== null ? val - p.invested : null;
      const pct = val !== null && p.invested > 0 ? (pnl / p.invested) * 100 : null;
      const holdDays = Math.floor((new Date() - p.oldestDate) / 86400000);
      return { ...p, avgCost: avg, currentPrice: px || null, currentValue: val, unrealizedPnL: pnl, unrealizedPct: pct, holdDays, sector: sectorFn(p), lotCount: p.lots.length };
    });
  }, [d.openLots, prices, sectorFn]);

  const sorted = useMemo(() => {
    const arr = [...positions];
    const dir = sort.dir === "asc" ? 1 : -1;
    arr.sort((a, b) => {
      const va = a[sort.key], vb = b[sort.key];
      if (va === null || va === undefined) return 1;
      if (vb === null || vb === undefined) return -1;
      if (typeof va === "string") return va.localeCompare(vb) * dir;
      return (va - vb) * dir;
    });
    return arr;
  }, [positions, sort]);

  const startEdit = (k, v) => { setEditing(k); setEditVal(v !== null ? String(v) : ""); };
  const commitEdit = (s, n) => { const v = parseNumber(editVal); onPriceChange(s || n, v || null); setEditing(null); setEditVal(""); };

  const totalInvested = positions.reduce((s, p) => s + p.invested, 0);
  const totalValue = positions.reduce((s, p) => s + (p.currentValue || p.invested), 0);
  const totalPnL = totalValue - totalInvested;
  const missingPx = positions.filter(p => p.currentPrice === null).length;

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <KPICard label="Positions ouvertes" value={positions.length} icon={Briefcase} accent="cyan" sub={`${positions.reduce((s, p) => s + p.lotCount, 0)} lots FIFO`} />
        <KPICard label="Capital investi" value={mask(fmtEUR(totalInvested, 0), hide)} icon={DollarSign} accent="blue" />
        <KPICard label="Valeur estimée" value={mask(fmtEUR(totalValue, 0), hide)} sub={missingPx > 0 ? `${missingPx} cours à saisir` : "cours complets"} subClass={missingPx > 0 ? "text-amber-400" : "text-emerald-400"} icon={Briefcase} accent="emerald" />
        <KPICard label="P&L latente totale" value={mask(fmtEUR(totalPnL, 0), hide)} sub={fmtPct(totalInvested > 0 ? (totalPnL / totalInvested) * 100 : 0)} subClass={classFor(totalPnL)} icon={TrendingUp} accent={totalPnL >= 0 ? "emerald" : "rose"} big />
      </div>

      {missingPx > 0 && (
        <div className="bg-amber-500/10 border border-amber-500/30 rounded-lg p-3 text-xs text-amber-200 flex items-start gap-2">
          <Info className="w-4 h-4 shrink-0 mt-0.5" />
          <div><strong>{missingPx}</strong> position(s) sans cours saisi. Clique sur <span className="bg-amber-500/20 border border-amber-500/40 px-1.5 py-0.5 rounded text-amber-200 font-mono text-[10px]">saisir…</span> pour calculer la P&L latente.</div>
        </div>
      )}

      <div className="flex items-center gap-2 flex-wrap">
        <div className="text-[10px] uppercase tracking-widest text-slate-500 font-medium">Affichage :</div>
        <div className="flex gap-1 bg-slate-900 p-1 rounded-lg border border-slate-800">
          <button onClick={() => setViewMode("cards")} className={`px-3 py-1 text-xs font-semibold rounded transition-colors ${viewMode === "cards" ? "bg-cyan-500/20 text-cyan-200" : "text-slate-400 hover:text-slate-100"}`}>Cartes</button>
          <button onClick={() => setViewMode("table")} className={`px-3 py-1 text-xs font-semibold rounded transition-colors ${viewMode === "table" ? "bg-cyan-500/20 text-cyan-200" : "text-slate-400 hover:text-slate-100"}`}>Tableau</button>
        </div>
        <div className="ml-auto text-[10px] uppercase tracking-widest text-slate-500 font-medium">Trier :</div>
        <select value={sort.key} onChange={(e) => setSort({ ...sort, key: e.target.value })} className="bg-slate-900 border border-slate-800 rounded px-2 py-1 text-xs">
          <option value="invested">Capital investi</option><option value="currentValue">Valeur actuelle</option>
          <option value="unrealizedPnL">P&L latente</option><option value="unrealizedPct">P&L %</option>
          <option value="name">Nom</option><option value="holdDays">Ancienneté</option>
        </select>
        <button onClick={() => setSort({ ...sort, dir: sort.dir === "desc" ? "asc" : "desc" })} className="bg-slate-900 border border-slate-800 rounded px-2 py-1 text-xs text-slate-300 hover:text-slate-100">{sort.dir === "desc" ? "↓" : "↑"}</button>
      </div>

      {viewMode === "cards" ? (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-3">
          {sorted.map((p, i) => {
            const key = p.symbol || p.name;
            const isEditing = editing === key;
            const hasPx = p.currentPrice !== null;
            const weight = totalInvested > 0 ? (p.invested / totalInvested) * 100 : 0;
            return (
              <div key={i} className="bg-slate-900/70 border border-slate-800 rounded-xl p-4 hover:border-slate-700 hover:bg-slate-800/40 transition-colors">
                <div className="flex items-start justify-between gap-2 mb-2">
                  <div className="min-w-0 flex-1">
                    <div className="flex items-center gap-1.5 mb-1">
                      <span className="w-2 h-2 rounded-full shrink-0" style={{ background: SECTOR_COLORS[p.sector] || "#64748b" }} />
                      <span className="text-[10px] uppercase tracking-wide text-slate-400 truncate">{p.sector}</span>
                    </div>
                    <div className="font-semibold text-slate-100 truncate text-sm">{p.name || p.symbol}</div>
                    <div className="text-[10px] text-slate-500 font-mono">{p.symbol}</div>
                  </div>
                  <span className="text-[10px] bg-slate-800 border border-slate-700 text-slate-300 rounded px-1.5 py-0.5 font-mono shrink-0">{weight.toFixed(1)}%</span>
                </div>
                <div className="space-y-1 text-xs">
                  <div className="flex justify-between"><span className="text-slate-500">Qté</span><span className="font-mono text-slate-200 tabular-nums">{fmtNum(p.shares, 4)}</span></div>
                  <div className="flex justify-between"><span className="text-slate-500">PRU</span><span className="font-mono text-slate-200 tabular-nums">{fmtEUR(p.avgCost, 4)}</span></div>
                  <div className="flex justify-between items-center">
                    <span className="text-slate-500">Cours</span>
                    {isEditing ? (
                      <input autoFocus type="text" value={editVal} onChange={(e) => setEditVal(e.target.value)} onBlur={() => commitEdit(p.symbol, p.name)}
                        onKeyDown={(e) => { if (e.key === "Enter") commitEdit(p.symbol, p.name); if (e.key === "Escape") { setEditing(null); setEditVal(""); } }}
                        className="bg-slate-800 border border-cyan-500 rounded px-2 py-0.5 text-right w-24 text-cyan-200 font-mono" />
                    ) : (
                      <button onClick={() => startEdit(key, p.currentPrice)} className={`px-2 py-0.5 rounded font-mono tabular-nums hover:bg-slate-700 transition-colors ${hasPx ? "text-cyan-300" : "text-amber-400 border border-amber-500/30 border-dashed"}`}>
                        {hasPx ? fmtEUR(p.currentPrice, 4) : "saisir…"}
                      </button>
                    )}
                  </div>
                  <div className="flex justify-between"><span className="text-slate-500">Investi</span><span className="font-mono text-slate-100 font-semibold tabular-nums">{mask(fmtEUR(p.invested, 2), hide)}</span></div>
                  <div className="flex justify-between"><span className="text-slate-500">Valeur</span><span className="font-mono text-slate-100 font-semibold tabular-nums">{hasPx ? mask(fmtEUR(p.currentValue, 2), hide) : <span className="text-slate-600">—</span>}</span></div>
                </div>
                {hasPx && (
                  <div className="mt-3 pt-3 border-t border-slate-800 flex items-center justify-between">
                    <div className={`text-xs font-semibold uppercase tracking-wide ${classFor(p.unrealizedPnL)}`}>P&L latente</div>
                    <div className="text-right">
                      <div className={`text-sm font-mono font-bold tabular-nums ${classFor(p.unrealizedPnL)}`}>{mask(fmtEUR(p.unrealizedPnL, 2), hide)}</div>
                      <div className={`text-[10px] font-mono ${classFor(p.unrealizedPct)}`}>{fmtPct(p.unrealizedPct, 2)}</div>
                    </div>
                  </div>
                )}
                <div className="mt-2 text-[10px] text-slate-600 font-mono">{p.lotCount} lot{p.lotCount > 1 ? "s" : ""} · détenu {p.holdDays}j</div>
              </div>
            );
          })}
        </div>
      ) : (
        <div className="bg-slate-900/70 border border-slate-800 rounded-xl overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-xs">
              <thead>
                <tr className="border-b border-slate-800 bg-slate-900">
                  {[{ k: "name", l: "Actif" }, { k: "sector", l: "Secteur" }, { k: "shares", l: "Qté", n: true }, { k: "avgCost", l: "PRU", n: true }, { k: "currentPrice", l: "Cours", n: true }, { k: "invested", l: "Investi", n: true }, { k: "currentValue", l: "Valeur", n: true }, { k: "unrealizedPnL", l: "P&L €", n: true }, { k: "unrealizedPct", l: "P&L %", n: true }, { k: "holdDays", l: "Âge", n: true }].map(c => (
                    <th key={c.k} className={`p-2 font-semibold text-slate-400 uppercase tracking-wider text-[10px] cursor-pointer hover:text-slate-100 ${c.n ? "text-right" : "text-left"}`} onClick={() => setSort(s => ({ key: c.k, dir: s.key === c.k && s.dir === "desc" ? "asc" : "desc" }))}>
                      {c.l} {sort.key === c.k && (sort.dir === "desc" ? "↓" : "↑")}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {sorted.map((p, i) => {
                  const key = p.symbol || p.name;
                  const isEditing = editing === key;
                  return (
                    <tr key={i} className={`border-b border-slate-800/60 ${ROW_HOVER}`}>
                      <td className="p-2"><div className="font-semibold text-slate-100 truncate max-w-[200px]">{p.name || p.symbol}</div><div className="text-[9px] text-slate-500 font-mono">{p.symbol}</div></td>
                      <td className="p-2"><span className="inline-flex items-center gap-1.5"><span className="w-2 h-2 rounded-full" style={{ background: SECTOR_COLORS[p.sector] || "#64748b" }} /><span className="text-slate-300 text-[11px]">{p.sector}</span></span></td>
                      <td className="p-2 text-right font-mono tabular-nums">{fmtNum(p.shares, 4)}</td>
                      <td className="p-2 text-right font-mono tabular-nums text-slate-300">{fmtEUR(p.avgCost, 4)}</td>
                      <td className="p-2 text-right font-mono tabular-nums">
                        {isEditing ? (
                          <input autoFocus type="text" value={editVal} onChange={(e) => setEditVal(e.target.value)} onBlur={() => commitEdit(p.symbol, p.name)} onKeyDown={(e) => { if (e.key === "Enter") commitEdit(p.symbol, p.name); if (e.key === "Escape") { setEditing(null); setEditVal(""); } }} className="bg-slate-800 border border-cyan-500 rounded px-2 py-0.5 text-right w-24 text-cyan-200" />
                        ) : (
                          <button onClick={() => startEdit(key, p.currentPrice)} className={`px-2 py-0.5 rounded hover:bg-slate-700 transition-colors ${p.currentPrice === null ? "text-amber-400 border border-amber-500/30 border-dashed" : "text-cyan-300"}`}>
                            {p.currentPrice !== null ? fmtEUR(p.currentPrice, 4) : "saisir…"}
                          </button>
                        )}
                      </td>
                      <td className="p-2 text-right font-mono tabular-nums text-slate-300">{mask(fmtEUR(p.invested, 2), hide)}</td>
                      <td className="p-2 text-right font-mono tabular-nums text-slate-100 font-semibold">{p.currentValue !== null ? mask(fmtEUR(p.currentValue, 2), hide) : "—"}</td>
                      <td className={`p-2 text-right font-mono tabular-nums font-semibold ${classFor(p.unrealizedPnL)}`}>{p.unrealizedPnL !== null ? mask(fmtEUR(p.unrealizedPnL, 2), hide) : "—"}</td>
                      <td className={`p-2 text-right font-mono tabular-nums ${classFor(p.unrealizedPct)}`}>{p.unrealizedPct !== null ? fmtPct(p.unrealizedPct, 2) : "—"}</td>
                      <td className="p-2 text-right font-mono tabular-nums text-slate-400">{p.holdDays}j</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

// ═══════════════ PERIODS VIEW ═══════════════

function PeriodsView({ d, sectorFn, hide }) {
  const [mode, setMode] = useState("month");
  const [anchor, setAnchor] = useState(() => new Date());
  const [customFrom, setCustomFrom] = useState("");
  const [customTo, setCustomTo] = useState("");
  const toISODate = (dt) => `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, "0")}-${String(dt.getDate()).padStart(2, "0")}`;

  const range = useMemo(() => {
    if (mode === "custom") {
      if (!customFrom || !customTo) return { from: null, to: null, label: "Sélectionne deux dates" };
      return { from: new Date(customFrom + "T00:00:00"), to: new Date(customTo + "T23:59:59.999"), label: `Du ${fmtDateShort(new Date(customFrom))} au ${fmtDateShort(new Date(customTo))}` };
    }
    const ref = new Date(anchor);
    if (mode === "day") { const f = new Date(ref); f.setHours(0, 0, 0, 0); const t = new Date(ref); t.setHours(23, 59, 59, 999); return { from: f, to: t, label: fmtDateLong(f) }; }
    if (mode === "week") {
      const dow = (ref.getDay() + 6) % 7;
      const mon = new Date(ref); mon.setDate(ref.getDate() - dow); mon.setHours(0, 0, 0, 0);
      const sun = new Date(mon); sun.setDate(mon.getDate() + 6); sun.setHours(23, 59, 59, 999);
      const tmp = new Date(Date.UTC(mon.getFullYear(), mon.getMonth(), mon.getDate()));
      tmp.setUTCDate(tmp.getUTCDate() - ((tmp.getUTCDay() + 6) % 7) + 3);
      const ft = new Date(Date.UTC(tmp.getUTCFullYear(), 0, 4));
      const wn = 1 + Math.round(((tmp - ft) / 86400000 - 3 + ((ft.getUTCDay() + 6) % 7)) / 7);
      return { from: mon, to: sun, label: `Semaine ${wn} · ${fmtDateShort(mon)} → ${fmtDateShort(sun)}` };
    }
    if (mode === "month") { const f = new Date(ref.getFullYear(), ref.getMonth(), 1); const t = new Date(ref.getFullYear(), ref.getMonth() + 1, 0, 23, 59, 59, 999); return { from: f, to: t, label: ref.toLocaleDateString("fr-FR", { month: "long", year: "numeric" }).replace(/^./, c => c.toUpperCase()) }; }
    if (mode === "quarter") { const q = Math.floor(ref.getMonth() / 3); return { from: new Date(ref.getFullYear(), q * 3, 1), to: new Date(ref.getFullYear(), q * 3 + 3, 0, 23, 59, 59, 999), label: `Trimestre ${q + 1} ${ref.getFullYear()}` }; }
    if (mode === "year") return { from: new Date(ref.getFullYear(), 0, 1), to: new Date(ref.getFullYear(), 11, 31, 23, 59, 59, 999), label: `Année ${ref.getFullYear()}` };
    return { from: null, to: null, label: "" };
  }, [mode, anchor, customFrom, customTo]);

  const periodData = useMemo(() => {
    if (!range.from || !range.to) return null;
    const tradesRaw = d.realizedTrades.filter(t => t.sellDate >= range.from && t.sellDate <= range.to);
    const trades = consolidateTrades(tradesRaw); // consolidé : 1 vente = 1 ligne
    const dividends = d.dividends.filter(dv => dv.date >= range.from && dv.date <= range.to);
    const interests = d.interests.filter(it => it.date >= range.from && it.date <= range.to);
    const savebacks = d.savebacks.filter(s => s.date >= range.from && s.date <= range.to);
    const buys = d.transactions.filter(t => t.kind === "BUY" && t.date >= range.from && t.date <= range.to);
    const sells = d.transactions.filter(t => t.kind === "SELL" && t.date >= range.from && t.date <= range.to);
    const realizedPnL = trades.reduce((s, t) => s + t.pnl, 0);
    const divAmount = dividends.reduce((s, x) => s + x.net, 0);
    const intAmount = interests.reduce((s, x) => s + x.net, 0);
    const saveAmount = savebacks.reduce((s, x) => s + x.amount, 0);
    const wins = trades.filter(t => t.pnl > 0);
    const losses = trades.filter(t => t.pnl < 0);
    const best = trades.reduce((b, t) => !b || t.pnl > b.pnl ? t : b, null);
    const worst = trades.reduce((w, t) => !w || t.pnl < w.pnl ? t : w, null);
    const sorted = [...trades].sort((a, b) => a.sellDate - b.sellDate);
    let cum = 0;
    const curve = sorted.map(t => ({ date: toISODate(t.sellDate), cum: (cum += t.pnl), pnl: t.pnl, name: t.name }));
    const bySec = {};
    for (const t of trades) {
      const s = sectorFn(t);
      if (!bySec[s]) bySec[s] = { sector: s, pnl: 0, count: 0, wins: 0, losses: 0 };
      bySec[s].pnl += t.pnl; bySec[s].count++;
      if (t.pnl > 0) bySec[s].wins++; else if (t.pnl < 0) bySec[s].losses++;
    }
    const bySector = Object.values(bySec).sort((a, b) => b.pnl - a.pnl);
    return { trades: sorted, dividends, interests, savebacks, realizedPnL, divAmount, intAmount, saveAmount, wins, losses, winRate: trades.length ? (wins.length / trades.length) * 100 : 0, best, worst, volume: trades.reduce((s, t) => s + t.proceeds, 0), invested: buys.reduce((s, t) => s + t.amount + t.fee + t.tax, 0), cashedOut: sells.reduce((s, t) => s + t.amount - t.fee - t.tax, 0), curve, bySector, buys, sells };
  }, [d, range, sectorFn]);

  // ═════ Période PRÉCÉDENTE (même durée juste avant) pour comparaison ═════
  const prevPeriodData = useMemo(() => {
    if (!range.from || !range.to || mode === "custom") return null;
    const span = range.to - range.from;
    const prevTo = new Date(range.from.getTime() - 1);
    const prevFrom = new Date(prevTo.getTime() - span);
    const tradesRaw = d.realizedTrades.filter(t => t.sellDate >= prevFrom && t.sellDate <= prevTo);
    const trades = consolidateTrades(tradesRaw);
    const realizedPnL = trades.reduce((s, t) => s + t.pnl, 0);
    const divAmount = d.dividends.filter(dv => dv.date >= prevFrom && dv.date <= prevTo).reduce((s, x) => s + x.net, 0);
    const intAmount = d.interests.filter(it => it.date >= prevFrom && it.date <= prevTo).reduce((s, x) => s + x.net, 0);
    const saveAmount = d.savebacks.filter(s => s.date >= prevFrom && s.date <= prevTo).reduce((s, x) => s + x.amount, 0);
    return { realizedPnL, divAmount, intAmount, saveAmount, net: realizedPnL + divAmount + intAmount + saveAmount, tradeCount: trades.length };
  }, [d, range, mode]);

  const step = (dir) => {
    const n = new Date(anchor);
    if (mode === "day") n.setDate(n.getDate() + dir);
    else if (mode === "week") n.setDate(n.getDate() + 7 * dir);
    else if (mode === "month") n.setMonth(n.getMonth() + dir);
    else if (mode === "quarter") n.setMonth(n.getMonth() + 3 * dir);
    else if (mode === "year") n.setFullYear(n.getFullYear() + dir);
    setAnchor(n);
  };
  const modes = [{ id: "day", label: "Jour" }, { id: "week", label: "Semaine" }, { id: "month", label: "Mois" }, { id: "quarter", label: "Trimestre" }, { id: "year", label: "Année" }, { id: "custom", label: "Plage" }];
  const netPeriod = periodData ? periodData.realizedPnL + periodData.divAmount + periodData.intAmount + periodData.saveAmount : 0;

  return (
    <div className="space-y-4">
      <Card>
        <div className="flex flex-wrap gap-1.5 mb-3">
          {modes.map(m => (
            <button key={m.id} onClick={() => setMode(m.id)} className={`px-3 md:px-4 py-2 text-xs md:text-sm font-semibold rounded border transition-colors ${mode === m.id ? "bg-cyan-500/20 border-cyan-400 text-cyan-200" : "bg-slate-950 border-slate-700 text-slate-400 hover:text-slate-100 hover:border-slate-500"}`}>{m.label}</button>
          ))}
        </div>
        {mode !== "custom" ? (
          <div className="flex items-center gap-2 flex-wrap">
            <button onClick={() => step(-1)} className="p-2 bg-slate-800 hover:bg-slate-700 rounded text-slate-200 transition-colors"><ChevronLeft className="w-4 h-4" /></button>
            <button onClick={() => step(1)} className="p-2 bg-slate-800 hover:bg-slate-700 rounded text-slate-200 transition-colors"><ChevronRight className="w-4 h-4" /></button>
            <button onClick={() => setAnchor(new Date())} className="px-3 py-2 bg-slate-800 hover:bg-slate-700 rounded text-xs font-semibold text-slate-200 flex items-center gap-1.5 transition-colors"><CalendarDays className="w-3.5 h-3.5" /> Aujourd'hui</button>
            <div className="flex-1 text-center text-sm md:text-base font-semibold text-cyan-300 px-2">{range.label}</div>
            <input type="date" value={toISODate(anchor)} onChange={(e) => setAnchor(new Date(e.target.value + "T12:00:00"))} className="bg-slate-950 border border-slate-700 rounded px-3 py-1.5 text-xs text-slate-200 focus:outline-none focus:border-cyan-500" />
          </div>
        ) : (
          <div className="flex items-center gap-3 flex-wrap">
            <div className="flex items-center gap-2"><label className="text-[10px] uppercase tracking-widest text-slate-500">Du</label><input type="date" value={customFrom} onChange={(e) => setCustomFrom(e.target.value)} className="bg-slate-950 border border-slate-700 rounded px-3 py-1.5 text-xs text-slate-200 focus:outline-none focus:border-cyan-500" /></div>
            <div className="flex items-center gap-2"><label className="text-[10px] uppercase tracking-widest text-slate-500">Au</label><input type="date" value={customTo} onChange={(e) => setCustomTo(e.target.value)} className="bg-slate-950 border border-slate-700 rounded px-3 py-1.5 text-xs text-slate-200 focus:outline-none focus:border-cyan-500" /></div>
            <div className="flex-1 text-center text-sm font-semibold text-cyan-300">{range.label}</div>
          </div>
        )}
      </Card>

      {!periodData ? <EmptyState title="Sélectionne une plage" icon={Calendar} /> : (
        <>
          <div className={`relative overflow-hidden rounded-xl p-8 md:p-10 border-2 ${netPeriod > 0 ? "bg-gradient-to-br from-emerald-500/10 to-cyan-500/5 border-emerald-500/30" : netPeriod < 0 ? "bg-gradient-to-br from-rose-500/10 to-orange-500/5 border-rose-500/30" : "bg-slate-900/60 border-slate-800"}`}>
            <div className="absolute top-0 right-0 -mt-12 -mr-12 w-48 h-48 rounded-full opacity-10 blur-2xl" style={{ background: netPeriod > 0 ? "#10b981" : netPeriod < 0 ? "#f43f5e" : "#64748b" }} />
            <div className="relative">
              <div className="text-[10px] uppercase tracking-widest text-slate-400 mb-2">{netPeriod > 0 ? "📈 Résultat net" : netPeriod < 0 ? "📉 Résultat net" : "Bilan de la période"}</div>
              <div className={`text-4xl md:text-6xl font-mono font-black tabular-nums ${classFor(netPeriod)}`}>{mask(fmtEUR(netPeriod, 2), hide)}</div>
              <div className="mt-3 grid grid-cols-2 md:grid-cols-4 gap-3 text-xs">
                <div><div className="text-[9px] uppercase text-slate-500 tracking-widest">P&L réalisée</div><div className={`font-mono font-semibold ${classFor(periodData.realizedPnL)}`}>{mask(fmtEUR(periodData.realizedPnL, 2), hide)}</div></div>
                <div><div className="text-[9px] uppercase text-slate-500 tracking-widest">Dividendes</div><div className="font-mono font-semibold text-amber-300">{mask(fmtEUR(periodData.divAmount, 2), hide)}</div></div>
                <div><div className="text-[9px] uppercase text-slate-500 tracking-widest">Intérêts</div><div className="font-mono font-semibold text-amber-300">{mask(fmtEUR(periodData.intAmount, 2), hide)}</div></div>
                <div><div className="text-[9px] uppercase text-slate-500 tracking-widest">Saveback</div><div className="font-mono font-semibold text-violet-300">{mask(fmtEUR(periodData.saveAmount, 2), hide)}</div></div>
              </div>
              <div className="mt-4 flex flex-wrap gap-4 text-xs text-slate-300 font-mono">
                <div><span className="text-slate-500">Trades :</span> <span className="font-semibold">{periodData.trades.length}</span></div>
                <div><span className="text-emerald-400">{periodData.wins.length}W</span> / <span className="text-rose-400">{periodData.losses.length}L</span></div>
                {periodData.trades.length > 0 && <div><span className="text-slate-500">WR :</span> <span className="font-semibold">{periodData.winRate.toFixed(0)}%</span></div>}
                <div><span className="text-slate-500">Volume :</span> <span className="font-semibold">{mask(fmtEUR(periodData.volume, 0), hide)}</span></div>
              </div>
              {/* Comparaison vs période précédente */}
              {prevPeriodData && (prevPeriodData.net !== 0 || prevPeriodData.tradeCount > 0) && (() => {
                const delta = netPeriod - prevPeriodData.net;
                const deltaPct = prevPeriodData.net !== 0 ? (delta / Math.abs(prevPeriodData.net)) * 100 : null;
                return (
                  <div className="mt-4 inline-flex items-center gap-2 bg-slate-950/60 border border-slate-800 rounded-lg px-3 py-2">
                    <span className="text-[10px] uppercase tracking-widest text-slate-500">vs période précédente :</span>
                    <span className={`text-sm font-mono font-bold ${classFor(delta)}`}>
                      {delta >= 0 ? "↑ +" : "↓ "}{mask(fmtEUR(delta, 0).replace("+", ""), hide)}
                    </span>
                    {deltaPct !== null && isFinite(deltaPct) && Math.abs(deltaPct) < 1000 && (
                      <span className={`text-[10px] font-mono ${classFor(delta)}`}>({deltaPct >= 0 ? "+" : ""}{deltaPct.toFixed(0)}%)</span>
                    )}
                    <span className="text-[10px] text-slate-600 font-mono">précéd. : {mask(fmtEUR(prevPeriodData.net, 0), hide)}</span>
                  </div>
                );
              })()}
            </div>
          </div>

          <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-3">
            <KPICard label="Achats période" value={mask(fmtEUR(periodData.invested, 0), hide)} sub={`${periodData.buys.length} ordres`} icon={ArrowDownRight} accent="rose" />
            <KPICard label="Ventes période" value={mask(fmtEUR(periodData.cashedOut, 0), hide)} sub={`${periodData.sells.length} ordres`} icon={ArrowUpRight} accent="emerald" />
            <KPICard label="Gain moyen" value={periodData.wins.length ? mask(fmtEUR(periodData.wins.reduce((s, t) => s + t.pnl, 0) / periodData.wins.length, 0), hide) : "—"} icon={TrendingUp} accent="emerald" />
            <KPICard label="Perte moyenne" value={periodData.losses.length ? mask(fmtEUR(periodData.losses.reduce((s, t) => s + t.pnl, 0) / periodData.losses.length, 0), hide) : "—"} icon={TrendingDown} accent="rose" />
            <KPICard label="Dividendes" value={mask(fmtEUR(periodData.divAmount, 2), hide)} sub={`${periodData.dividends.length} vers.`} icon={Percent} accent="amber" />
            <KPICard label="Saveback" value={mask(fmtEUR(periodData.saveAmount, 2), hide)} sub={`${periodData.savebacks.length} bonus`} icon={Gift} accent="violet" />
          </div>

          {periodData.best && periodData.worst && (
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="bg-emerald-500/5 border border-emerald-500/30 rounded-xl p-4">
                <div className="text-[10px] uppercase tracking-widest text-emerald-400 mb-2">★ Meilleur trade</div>
                <div className="text-lg font-bold truncate">{periodData.best.name}</div>
                <div className="text-3xl font-mono font-bold text-emerald-400 mt-1">{mask(fmtEUR(periodData.best.pnl, 2), hide)}</div>
                <div className="text-sm text-emerald-300/70 font-mono">{fmtPct(periodData.best.pnlPct, 2)} · {periodData.best.holdingDays}j</div>
                <div className="text-[10px] text-slate-500 font-mono mt-2">Vendu le {fmtDateShort(periodData.best.sellDate)}</div>
              </div>
              <div className="bg-rose-500/5 border border-rose-500/30 rounded-xl p-4">
                <div className="text-[10px] uppercase tracking-widest text-rose-400 mb-2">⚠ Pire trade</div>
                <div className="text-lg font-bold truncate">{periodData.worst.name}</div>
                <div className="text-3xl font-mono font-bold text-rose-400 mt-1">{mask(fmtEUR(periodData.worst.pnl, 2), hide)}</div>
                <div className="text-sm text-rose-300/70 font-mono">{fmtPct(periodData.worst.pnlPct, 2)} · {periodData.worst.holdingDays}j</div>
                <div className="text-[10px] text-slate-500 font-mono mt-2">Vendu le {fmtDateShort(periodData.worst.sellDate)}</div>
              </div>
            </div>
          )}

          {periodData.curve.length > 0 && (
            <Card>
              <SectionTitle icon={Activity}>Courbe P&L cumulée sur la période</SectionTitle>
              <ResponsiveContainer width="100%" height={240}>
                <AreaChart data={periodData.curve}>
                  <defs><linearGradient id="gP" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={periodData.realizedPnL >= 0 ? "#10b981" : "#ef4444"} stopOpacity={0.35} /><stop offset="100%" stopColor={periodData.realizedPnL >= 0 ? "#10b981" : "#ef4444"} stopOpacity={0} /></linearGradient></defs>
                  <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                  <XAxis dataKey="date" stroke="#64748b" fontSize={10} />
                  <YAxis stroke="#64748b" fontSize={10} tickFormatter={(v) => fmtEUR(v, 0)} />
                  <Tooltip content={<ChartTooltip />} />
                  <ReferenceLine y={0} stroke="#475569" />
                  <Area type="monotone" dataKey="cum" stroke={periodData.realizedPnL >= 0 ? "#10b981" : "#ef4444"} strokeWidth={2} fill="url(#gP)" name="P&L cumulée" />
                </AreaChart>
              </ResponsiveContainer>
            </Card>
          )}

          {periodData.bySector.length > 0 && (
            <Card>
              <SectionTitle icon={PieIcon}>P&L par secteur sur la période</SectionTitle>
              <div className="overflow-x-auto">
                <table className="w-full text-xs">
                  <thead><tr className="border-b border-slate-800"><th className="p-2 text-left text-[10px] text-slate-500 uppercase">Secteur</th><th className="p-2 text-right text-[10px] text-slate-500 uppercase">Trades</th><th className="p-2 text-right text-[10px] text-slate-500 uppercase">W / L</th><th className="p-2 text-right text-[10px] text-slate-500 uppercase">P&L</th></tr></thead>
                  <tbody>{periodData.bySector.map((s, i) => (
                    <tr key={i} className={`border-b border-slate-800/60 ${ROW_HOVER}`}>
                      <td className="p-2"><span className="inline-flex items-center gap-2"><span className="w-2 h-2 rounded-full" style={{ background: SECTOR_COLORS[s.sector] || "#64748b" }} /><span className="font-semibold text-slate-100">{s.sector}</span></span></td>
                      <td className="p-2 text-right font-mono text-slate-300">{s.count}</td>
                      <td className="p-2 text-right font-mono text-[10px]"><span className="text-emerald-400">{s.wins}W</span> / <span className="text-rose-400">{s.losses}L</span></td>
                      <td className={`p-2 text-right font-mono font-semibold ${classFor(s.pnl)}`}>{mask(fmtEUR(s.pnl, 2), hide)}</td>
                    </tr>))}
                  </tbody>
                </table>
              </div>
            </Card>
          )}

          {periodData.trades.length > 0 && (
            <Card>
              <SectionTitle icon={Activity}>{periodData.trades.length} trades (triés par P&L)</SectionTitle>
              <div className="overflow-x-auto max-h-[500px] overflow-y-auto">
                <table className="w-full text-xs">
                  <thead className="sticky top-0 bg-slate-900 z-10">
                    <tr className="border-b border-slate-800"><th className="p-2 text-left text-[10px] text-slate-500 uppercase">Vente</th><th className="p-2 text-left text-[10px] text-slate-500 uppercase">Actif</th><th className="p-2 text-left text-[10px] text-slate-500 uppercase hidden md:table-cell">Secteur</th><th className="p-2 text-right text-[10px] text-slate-500 uppercase hidden sm:table-cell">Qté</th><th className="p-2 text-right text-[10px] text-slate-500 uppercase hidden sm:table-cell">Durée</th><th className="p-2 text-right text-[10px] text-slate-500 uppercase">P&L</th><th className="p-2 text-right text-[10px] text-slate-500 uppercase">%</th></tr>
                  </thead>
                  <tbody>
                    {[...periodData.trades].sort((a, b) => b.pnl - a.pnl).map((t, i) => {
                      const sec = sectorFn(t);
                      return (
                        <tr key={i} className={`border-b border-slate-800/60 ${ROW_HOVER}`}>
                          <td className="p-2 font-mono text-slate-400 whitespace-nowrap text-[10px]">{fmtDateShort(t.sellDate)}</td>
                          <td className="p-2"><div className="text-slate-100 font-semibold truncate max-w-[180px]">{t.name}</div><div className="text-[9px] text-slate-500 font-mono">{t.symbol}</div></td>
                          <td className="p-2 hidden md:table-cell"><span className="inline-flex items-center gap-1.5"><span className="w-2 h-2 rounded-full" style={{ background: SECTOR_COLORS[sec] || "#64748b" }} /><span className="text-[10px] text-slate-400">{sec}</span></span></td>
                          <td className="p-2 text-right font-mono text-slate-300 hidden sm:table-cell">{fmtNum(t.shares, 4)}</td>
                          <td className="p-2 text-right font-mono text-slate-400 hidden sm:table-cell">{t.holdingDays}j</td>
                          <td className={`p-2 text-right font-mono font-semibold ${classFor(t.pnl)}`}>{mask(fmtEUR(t.pnl, 2), hide)}</td>
                          <td className={`p-2 text-right font-mono ${classFor(t.pnlPct)}`}>{fmtPct(t.pnlPct, 1)}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </Card>
          )}

          {(periodData.dividends.length > 0 || periodData.savebacks.length > 0) && (
            <Card>
              <SectionTitle icon={Coins}>Revenus de la période</SectionTitle>
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-2">
                {periodData.dividends.map((dv, i) => (
                  <div key={`d${i}`} className={`flex items-center justify-between bg-slate-950/50 rounded-lg p-2 border border-slate-800/60 ${ROW_HOVER}`}>
                    <div className="min-w-0 flex-1"><div className="font-semibold text-slate-100 truncate text-xs">💰 {dv.name}</div><div className="text-[10px] text-slate-500 font-mono">{fmtDateShort(dv.date)} · dividende</div></div>
                    <div className="text-amber-400 font-mono font-semibold text-sm">{mask(fmtEUR(dv.net, 2), hide)}</div>
                  </div>
                ))}
                {periodData.savebacks.map((sb, i) => (
                  <div key={`s${i}`} className={`flex items-center justify-between bg-slate-950/50 rounded-lg p-2 border border-slate-800/60 ${ROW_HOVER}`}>
                    <div className="min-w-0 flex-1"><div className="font-semibold text-slate-100 truncate text-xs">🎁 {sb.name}</div><div className="text-[10px] text-slate-500 font-mono">{fmtDateShort(sb.date)} · saveback</div></div>
                    <div className="text-violet-400 font-mono font-semibold text-sm">{mask(fmtEUR(sb.amount, 2), hide)}</div>
                  </div>
                ))}
              </div>
            </Card>
          )}
        </>
      )}
    </div>
  );
}

// ═══════════════ SECTORS VIEW ═══════════════

// ═══════════════ ANALYSES (Sectors + Analytics fusionnés) ═══════════════

function AnalysesView({ d, sectorFn, hide }) {
  const { stats, realizedTrades, consolidatedTrades, byAsset, bySector } = d;
  const [view, setView] = useState("performance"); // performance | distribution | trades | secteurs | actifs | fiscalite

  const topAssets = [...byAsset].sort((a, b) => b.totalPnL - a.totalPnL).slice(0, 15);
  const flopAssets = [...byAsset].filter(a => a.totalPnL < 0).sort((a, b) => a.totalPnL - b.totalPnL).slice(0, 15);
  const scatter = consolidatedTrades.map(t => ({ x: t.holdingDays, y: t.pnlPct, pnl: t.pnl, name: t.name }));
  // Top/Flop trades (déplacés depuis le Dashboard) — sur trades consolidés
  const topTrades = [...consolidatedTrades].sort((a, b) => b.pnl - a.pnl).slice(0, 10);
  const flopTrades = [...consolidatedTrades].sort((a, b) => a.pnl - b.pnl).slice(0, 10);

  const tabs = [
    { id: "performance", label: "Performance", icon: Target },
    { id: "distribution", label: "Distribution", icon: BarChart3 },
    { id: "trades", label: "Meilleurs / Pires", icon: Award },
    { id: "secteurs", label: "Secteurs", icon: PieIcon },
    { id: "actifs", label: "Actifs", icon: Briefcase },
    { id: "fiscalite", label: "Fiscalité", icon: Percent },
  ];

  return (
    <div className="space-y-4">
      {/* KPIs en haut, toujours visibles */}
      <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3">
        <KPICard label="Total ventes" value={stats.count} icon={Hash} accent="cyan" sub={stats.lotCount ? `${stats.lotCount} lots FIFO` : ""} />
        <KPICard label="Win rate" value={stats.count ? stats.winRate.toFixed(1) + "%" : "—"} icon={Target} accent={stats.winRate >= 50 ? "emerald" : "rose"} sub={stats.count ? `${stats.winCount}W · ${stats.lossCount}L` : ""} />
        <KPICard label="Profit factor" value={stats.profitFactor === Infinity ? "∞" : stats.profitFactor.toFixed(2)} icon={Zap} accent={stats.profitFactor >= 1.5 ? "emerald" : stats.profitFactor >= 1 ? "amber" : "rose"} sub="> 1,5 = excellent" />
        <KPICard label="Expectancy" value={stats.count ? mask(fmtEUR(stats.totalPnL / stats.count, 0), hide) : "—"} subClass={classFor(stats.totalPnL)} icon={Flame} accent="amber" sub="par vente" />
        <KPICard label="Meilleure série" value={`${stats.largestWinStreak}W`} sub={`Pire: ${stats.largestLossStreak}L`} icon={Award} accent="emerald" />
        <KPICard label="Max drawdown" value={mask(fmtEUR(-stats.maxDrawdown, 0), hide)} icon={TrendingDown} accent="rose" />
      </div>

      {/* Sélecteur de vue */}
      <div className="flex gap-1 bg-slate-900 p-1 rounded-lg border border-slate-800 overflow-x-auto">
        {tabs.map(t => (
          <button key={t.id} onClick={() => setView(t.id)}
            className={`flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded transition-colors whitespace-nowrap ${view === t.id ? "bg-cyan-500/20 text-cyan-200" : "text-slate-400 hover:text-slate-100"}`}>
            <t.icon className="w-3.5 h-3.5" /> {t.label}
          </button>
        ))}
      </div>

      {/* ═══ VIEW: PERFORMANCE ═══ */}
      {view === "performance" && (
        <>
          {stats.bestTrade && stats.worstTrade && (
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="bg-emerald-500/5 border border-emerald-500/30 rounded-xl p-4">
                <div className="text-[10px] uppercase tracking-widest text-emerald-400 mb-2">⭐ Meilleur trade all-time</div>
                <div className="text-lg font-bold truncate">{stats.bestTrade.name}</div>
                <div className="text-3xl font-mono font-bold text-emerald-400 mt-1">{mask(fmtEUR(stats.bestTrade.pnl, 2), hide)}</div>
                <div className="text-sm text-emerald-300/70 font-mono">{fmtPct(stats.bestTrade.pnlPct, 2)} · {stats.bestTrade.holdingDays}j détenus</div>
                <div className="text-[10px] text-slate-500 font-mono mt-1">Vendu le {fmtDateShort(stats.bestTrade.sellDate)}</div>
              </div>
              <div className="bg-rose-500/5 border border-rose-500/30 rounded-xl p-4">
                <div className="text-[10px] uppercase tracking-widest text-rose-400 mb-2">⚠ Pire trade all-time</div>
                <div className="text-lg font-bold truncate">{stats.worstTrade.name}</div>
                <div className="text-3xl font-mono font-bold text-rose-400 mt-1">{mask(fmtEUR(stats.worstTrade.pnl, 2), hide)}</div>
                <div className="text-sm text-rose-300/70 font-mono">{fmtPct(stats.worstTrade.pnlPct, 2)} · {stats.worstTrade.holdingDays}j détenus</div>
                <div className="text-[10px] text-slate-500 font-mono mt-1">Vendu le {fmtDateShort(stats.worstTrade.sellDate)}</div>
              </div>
            </div>
          )}

          {scatter.length > 0 && (
            <Card>
              <SectionTitle icon={Activity} subtitle="Chaque point = une vente. En haut = gain, en bas = perte. À droite = détenue longtemps.">
                Durée de détention vs Rendement
              </SectionTitle>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-2 mb-3 text-[10px]">
                <div className="bg-emerald-500/10 border border-emerald-500/30 rounded p-2"><div className="text-emerald-300 font-semibold">↖ Gain rapide</div><div className="text-slate-400 mt-0.5">courte · positive</div></div>
                <div className="bg-emerald-500/10 border border-emerald-500/30 rounded p-2"><div className="text-emerald-300 font-semibold">↗ Gain long-terme</div><div className="text-slate-400 mt-0.5">longue · positive</div></div>
                <div className="bg-rose-500/10 border border-rose-500/30 rounded p-2"><div className="text-rose-300 font-semibold">↙ Perte rapide</div><div className="text-slate-400 mt-0.5">courte · négative</div></div>
                <div className="bg-rose-500/10 border border-rose-500/30 rounded p-2"><div className="text-rose-300 font-semibold">↘ Perte long-terme</div><div className="text-slate-400 mt-0.5">longue · négative</div></div>
              </div>
              <ResponsiveContainer width="100%" height={340}>
                <ScatterChart margin={{ top: 10, right: 20, bottom: 40, left: 10 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                  <XAxis type="number" dataKey="x" stroke="#64748b" fontSize={10} tickFormatter={(v) => v + "j"} label={{ value: "Durée (jours)", position: "insideBottom", offset: -15, fill: "#94a3b8", fontSize: 11 }} />
                  <YAxis type="number" dataKey="y" stroke="#64748b" fontSize={10} tickFormatter={(v) => `${v.toFixed(0)}%`} label={{ value: "Rendement (%)", angle: -90, position: "insideLeft", fill: "#94a3b8", fontSize: 11 }} />
                  <Tooltip content={({ active, payload }) => {
                    if (!active || !payload?.length) return null;
                    const d = payload[0].payload;
                    return (
                      <div className="bg-slate-950 border border-slate-700 rounded-lg p-2.5 shadow-xl max-w-[220px]">
                        <div className="text-xs font-semibold text-slate-100 mb-1 truncate">{d.name}</div>
                        <div className="text-[10px] font-mono text-slate-400">Durée : {d.x}j</div>
                        <div className={`text-[10px] font-mono ${d.y >= 0 ? "text-emerald-400" : "text-rose-400"}`}>Rendement : {fmtPct(d.y, 2)}</div>
                        <div className={`text-[10px] font-mono font-semibold ${d.pnl >= 0 ? "text-emerald-400" : "text-rose-400"}`}>P&L : {fmtEUR(d.pnl, 2)}</div>
                      </div>
                    );
                  }} />
                  <ReferenceLine y={0} stroke="#64748b" strokeDasharray="4 4" />
                  <Scatter data={scatter}>{scatter.map((e, i) => <Cell key={i} fill={e.y >= 0 ? "#10b981" : "#ef4444"} fillOpacity={0.7} />)}</Scatter>
                </ScatterChart>
              </ResponsiveContainer>
            </Card>
          )}
        </>
      )}

      {/* ═══ VIEW: SECTEURS ═══ */}
      {view === "secteurs" && (
        <>
          <Card>
            <SectionTitle icon={PieIcon}>Vue d'ensemble par secteur</SectionTitle>
            <div className="overflow-x-auto">
              <table className="w-full text-xs">
                <thead><tr className="border-b border-slate-800">{["Secteur", "Actifs", "Investi", "Valeur", "P&L latente", "P&L réalisée", "Rdt %", "WR", "Trades"].map((h, i) => (<th key={i} className={`p-2 font-semibold text-slate-400 uppercase tracking-wider text-[10px] ${i === 0 ? "text-left" : "text-right"}`}>{h}</th>))}</tr></thead>
                <tbody>
                  {bySector.map((s, i) => (
                    <tr key={s.sector} className={`border-b border-slate-800/60 ${ROW_HOVER}`}>
                      <td className="p-2"><span className="inline-flex items-center gap-2"><span className="w-2.5 h-2.5 rounded-full" style={{ background: SECTOR_COLORS[s.sector] || "#64748b" }} /><span className="font-semibold text-slate-100">{s.sector}</span></span></td>
                      <td className="p-2 text-right font-mono text-slate-300">{s.assetCount}</td>
                      <td className="p-2 text-right font-mono text-slate-300">{mask(fmtEUR(s.investedOpen, 0), hide)}</td>
                      <td className="p-2 text-right font-mono text-slate-100 font-semibold">{mask(fmtEUR(s.currentValue, 0), hide)}</td>
                      <td className={`p-2 text-right font-mono font-semibold ${classFor(s.unrealizedPnL)}`}>{mask(fmtEUR(s.unrealizedPnL, 0), hide)}</td>
                      <td className={`p-2 text-right font-mono font-semibold ${classFor(s.realizedPnL)}`}>{mask(fmtEUR(s.realizedPnL, 0), hide)}</td>
                      <td className={`p-2 text-right font-mono ${classFor(s.returnPct)}`}>{fmtPct(s.returnPct, 1)}</td>
                      <td className="p-2 text-right font-mono text-slate-300">{s.realizedCount > 0 ? `${s.winRate.toFixed(0)}%` : "—"}</td>
                      <td className="p-2 text-right font-mono text-slate-400">{s.realizedCount}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
            <Card>
              <SectionTitle icon={BarChart3} subtitle="Sur les positions ouvertes">Rendement % par secteur</SectionTitle>
              {bySector.filter(s => s.investedOpen > 0).length > 0 ? (
                <ResponsiveContainer width="100%" height={320}>
                  <BarChart data={bySector.filter(s => s.investedOpen > 0)} layout="vertical" margin={{ left: 20 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                    <XAxis type="number" stroke="#64748b" fontSize={10} tickFormatter={(v) => `${v.toFixed(0)}%`} />
                    <YAxis dataKey="sector" type="category" stroke="#94a3b8" fontSize={10} width={140} />
                    <Tooltip content={<ChartTooltip formatter={(v) => fmtPct(v, 2)} />} />
                    <ReferenceLine x={0} stroke="#475569" />
                    <Bar dataKey="returnPct">{bySector.filter(s => s.investedOpen > 0).map((s, i) => <Cell key={i} fill={s.returnPct >= 0 ? "#10b981" : "#ef4444"} />)}</Bar>
                  </BarChart>
                </ResponsiveContainer>
              ) : <EmptyState title="Pas de positions" icon={BarChart3} />}
            </Card>
            <Card>
              <SectionTitle icon={Target}>P&L réalisée par secteur</SectionTitle>
              {bySector.filter(s => s.realizedCount > 0).length > 0 ? (
                <ResponsiveContainer width="100%" height={320}>
                  <BarChart data={bySector.filter(s => s.realizedCount > 0).sort((a, b) => b.realizedPnL - a.realizedPnL)}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                    <XAxis dataKey="sector" stroke="#64748b" fontSize={9} angle={-25} textAnchor="end" height={70} interval={0} />
                    <YAxis stroke="#64748b" fontSize={10} tickFormatter={(v) => `${(v / 1000).toFixed(1)}k`} />
                    <Tooltip content={<ChartTooltip />} />
                    <ReferenceLine y={0} stroke="#475569" />
                    <Bar dataKey="realizedPnL">{bySector.filter(s => s.realizedCount > 0).map((s, i) => <Cell key={i} fill={s.realizedPnL >= 0 ? SECTOR_COLORS[s.sector] || "#10b981" : "#ef4444"} />)}</Bar>
                  </BarChart>
                </ResponsiveContainer>
              ) : <EmptyState title="Pas de trades" icon={Target} />}
            </Card>
          </div>
        </>
      )}

      {/* ═══ VIEW: ACTIFS ═══ */}
      {view === "actifs" && (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <Card><SectionTitle icon={Award}>Top 15 actifs par P&L totale</SectionTitle><AssetTable assets={topAssets} sectorFn={sectorFn} hide={hide} /></Card>
          <Card><SectionTitle icon={TrendingDown}>Flop 15 actifs (perdants)</SectionTitle><AssetTable assets={flopAssets} sectorFn={sectorFn} hide={hide} /></Card>
        </div>
      )}

      {/* ═══ VIEW: DISTRIBUTION ═══ */}
      {view === "distribution" && stats.pnlDistribution.length > 0 && (
        <Card>
          <SectionTitle icon={BarChart3} subtitle="Combien de ventes par tranche de gain/perte ?">Distribution des P&L</SectionTitle>
          <ResponsiveContainer width="100%" height={320}>
            <BarChart data={stats.pnlDistribution}>
              <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
              <XAxis dataKey="bucket" stroke="#64748b" fontSize={10} tickFormatter={(v) => fmtEUR(v, 0)} />
              <YAxis stroke="#64748b" fontSize={10} />
              <Tooltip content={<ChartTooltip formatter={(v) => `${v} ventes`} labelFormatter={(l) => `P&L ~ ${fmtEUR(l, 0)}`} />} />
              <ReferenceLine x={0} stroke="#475569" />
              <Bar dataKey="count">{stats.pnlDistribution.map((e, i) => <Cell key={i} fill={e.bucket >= 0 ? "#10b981" : "#ef4444"} />)}</Bar>
            </BarChart>
          </ResponsiveContainer>
        </Card>
      )}

      {/* ═══ VIEW: MEILLEURS / PIRES TRADES (déplacé depuis Dashboard) ═══ */}
      {view === "trades" && (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <Card><SectionTitle icon={Award} subtitle="Tes 10 plus belles plus-values réalisées">🏆 Meilleurs trades</SectionTitle><TradeList trades={topTrades} sectorFn={sectorFn} hide={hide} /></Card>
          <Card><SectionTitle icon={TrendingDown} subtitle="Tes 10 plus grosses moins-values réalisées">📉 Pires trades</SectionTitle><TradeList trades={flopTrades} sectorFn={sectorFn} hide={hide} /></Card>
        </div>
      )}

      {/* ═══ VIEW: FISCALITÉ ═══ */}
      {view === "fiscalite" && <FiscaliteView d={d} hide={hide} />}
    </div>
  );
}

// ═══════ Vue Fiscalité : simulation déclaration d'impôts ═══════
const FiscaliteView = ({ d, hide }) => {
  const tax = d.taxation;
  if (!tax) return <EmptyState title="Pas de données fiscales" icon={Percent} />;
  const [year, setYear] = useState(new Date().getFullYear());

  // P&L réalisée + dividendes + intérêts par année (CTO uniquement)
  const byYear = useMemo(() => {
    const m = {};
    const ensure = (y) => m[y] = m[y] || { year: y, pnlCTO: 0, pnlPEA: 0, divsCTO: 0, divsPEA: 0, intsCTO: 0, intsPEA: 0 };
    for (const t of d.realizedTrades) {
      const y = t.sellDate.getFullYear();
      const e = ensure(y);
      if (t.accountType === "PEA") e.pnlPEA += t.pnl; else e.pnlCTO += t.pnl;
    }
    for (const dv of d.dividends) {
      const y = dv.date.getFullYear();
      const e = ensure(y);
      if (dv.accountType === "PEA") e.divsPEA += dv.net; else e.divsCTO += dv.net;
    }
    for (const it of d.interests) {
      const y = it.date.getFullYear();
      const e = ensure(y);
      if (it.accountType === "PEA") e.intsPEA += it.net; else e.intsCTO += it.net;
    }
    return Object.values(m).sort((a, b) => b.year - a.year);
  }, [d.realizedTrades, d.dividends, d.interests]);

  const years = byYear.map(y => y.year);
  const selectedData = byYear.find(y => y.year === year) || byYear[0];
  if (!selectedData) return <EmptyState title="Aucune transaction taxable" icon={Percent} />;

  const TAX = 0.314;
  const taxableCTO = Math.max(0, selectedData.pnlCTO) + Math.max(0, selectedData.divsCTO) + Math.max(0, selectedData.intsCTO);
  const taxDue = taxableCTO * TAX;
  // Décomposition PFU : 12,8% IR + 17,2% prélèvements sociaux
  const irPart = taxableCTO * 0.128;
  const psPart = taxableCTO * 0.172;

  return (
    <div className="space-y-4">
      {/* Sélecteur année */}
      {years.length > 1 && (
        <div className="flex items-center gap-2 flex-wrap">
          <span className="text-[10px] uppercase tracking-widest text-slate-500">Année fiscale :</span>
          {years.map(y => (
            <button key={y} onClick={() => setYear(y)} className={`px-3 py-1.5 text-xs font-semibold rounded ${year === y ? "bg-cyan-500/20 text-cyan-200 border border-cyan-500/40" : "bg-slate-950 text-slate-400 border border-slate-700"}`}>{y}</button>
          ))}
        </div>
      )}

      {/* Hero : impôt estimé */}
      <div className="relative overflow-hidden rounded-2xl border-2 border-rose-500/30 bg-gradient-to-br from-rose-500/10 via-slate-900 to-slate-950">
        <div className="absolute -top-32 -right-32 w-64 h-64 rounded-full opacity-20 blur-3xl bg-rose-500" />
        <div className="relative p-6 md:p-8">
          <div className="text-[10px] uppercase tracking-[0.2em] text-slate-400 mb-2">📋 Impôt estimé sur compte-titres (CTO) · {selectedData.year}</div>
          <div className="flex items-end gap-6 flex-wrap">
            <div>
              <div className="text-5xl md:text-6xl font-mono font-black text-rose-400 tabular-nums leading-none">{mask(fmtEUR(taxDue, 0), hide)}</div>
              <div className="text-sm text-slate-400 mt-2 font-mono">à provisionner pour l'année {selectedData.year}</div>
            </div>
            <div className="flex-1 min-w-[200px]">
              <div className="space-y-2">
                <div className="flex justify-between text-sm"><span className="text-slate-400">Base imposable CTO</span><span className="font-mono font-semibold">{mask(fmtEUR(taxableCTO, 2), hide)}</span></div>
                <div className="flex justify-between text-xs"><span className="text-slate-500 pl-3">Impôt sur le revenu (12,8%)</span><span className="font-mono text-amber-300">{mask(fmtEUR(irPart, 2), hide)}</span></div>
                <div className="flex justify-between text-xs"><span className="text-slate-500 pl-3">Prélèvements sociaux (17,2%)</span><span className="font-mono text-amber-300">{mask(fmtEUR(psPart, 2), hide)}</span></div>
                <div className="flex justify-between text-sm pt-2 border-t border-slate-700 font-bold"><span className="text-rose-300">Total PFU (30%... 31,4% ici)</span><span className="font-mono text-rose-400">{mask(fmtEUR(taxDue, 2), hide)}</span></div>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Détail CTO vs PEA */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <Card>
          <SectionTitle icon={AlertCircle} subtitle="Imposable au PFU 31,4%">Compte-titres (CTO)</SectionTitle>
          <div className="space-y-2">
            <FiscalRow label="Plus-values réalisées" value={selectedData.pnlCTO} taxable hide={hide} />
            <FiscalRow label="Dividendes nets" value={selectedData.divsCTO} taxable hide={hide} />
            <FiscalRow label="Intérêts nets" value={selectedData.intsCTO} taxable hide={hide} />
            <div className="flex justify-between pt-2 border-t border-slate-700 text-sm font-bold">
              <span className="text-slate-300">Base imposable</span>
              <span className="font-mono text-rose-400">{mask(fmtEUR(taxableCTO, 2), hide)}</span>
            </div>
            <div className="flex justify-between text-sm">
              <span className="text-slate-400">Impôt dû (31,4%)</span>
              <span className="font-mono font-bold text-rose-400">{mask(fmtEUR(taxDue, 2), hide)}</span>
            </div>
          </div>
        </Card>
        <Card>
          <SectionTitle icon={Check} subtitle="Exonéré d'impôt (hors prélèvements sociaux à la sortie)">PEA</SectionTitle>
          {(selectedData.pnlPEA !== 0 || selectedData.divsPEA !== 0 || selectedData.intsPEA !== 0) ? (
            <div className="space-y-2">
              <FiscalRow label="Plus-values réalisées" value={selectedData.pnlPEA} exempt hide={hide} />
              <FiscalRow label="Dividendes nets" value={selectedData.divsPEA} exempt hide={hide} />
              <FiscalRow label="Intérêts nets" value={selectedData.intsPEA} exempt hide={hide} />
              <div className="flex justify-between pt-2 border-t border-slate-700 text-sm font-bold">
                <span className="text-slate-300">Total PEA</span>
                <span className="font-mono text-emerald-400">{mask(fmtEUR(selectedData.pnlPEA + selectedData.divsPEA + selectedData.intsPEA, 2), hide)}</span>
              </div>
              <div className="text-[11px] text-emerald-300/70 bg-emerald-500/5 border border-emerald-500/20 rounded p-2 mt-2">
                ✓ Aucun impôt sur le revenu tant que tu ne fais pas de retrait avant 5 ans. Les gains restent dans l'enveloppe.
              </div>
            </div>
          ) : (
            <EmptyState title="Pas d'activité PEA cette année" icon={Check} />
          )}
        </Card>
      </div>

      {/* Historique par année */}
      {byYear.length > 1 && (
        <Card>
          <SectionTitle icon={CalendarDays}>Historique fiscal par année</SectionTitle>
          <div className="overflow-x-auto">
            <table className="w-full text-xs">
              <thead><tr className="border-b border-slate-800">{["Année", "P&L CTO", "Div+Int CTO", "Base imposable", "Impôt estimé", "Gains PEA (exonérés)"].map((h, i) => <th key={i} className={`p-2 text-[10px] text-slate-500 uppercase ${i === 0 ? "text-left" : "text-right"}`}>{h}</th>)}</tr></thead>
              <tbody>
                {byYear.map((y, i) => {
                  const base = Math.max(0, y.pnlCTO) + Math.max(0, y.divsCTO) + Math.max(0, y.intsCTO);
                  const tx = base * TAX;
                  return (
                    <tr key={i} className={`border-b border-slate-800/60 ${ROW_HOVER} cursor-pointer`} onClick={() => setYear(y.year)}>
                      <td className="p-2 font-mono font-semibold text-slate-200">{y.year}</td>
                      <td className={`p-2 text-right font-mono ${classFor(y.pnlCTO)}`}>{mask(fmtEUR(y.pnlCTO, 0), hide)}</td>
                      <td className="p-2 text-right font-mono text-amber-400">{mask(fmtEUR(y.divsCTO + y.intsCTO, 0), hide)}</td>
                      <td className="p-2 text-right font-mono text-slate-300">{mask(fmtEUR(base, 0), hide)}</td>
                      <td className="p-2 text-right font-mono font-bold text-rose-400">{mask(fmtEUR(tx, 0), hide)}</td>
                      <td className="p-2 text-right font-mono text-emerald-400">{mask(fmtEUR(y.pnlPEA + y.divsPEA + y.intsPEA, 0), hide)}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </Card>
      )}

      <div className="text-[11px] text-slate-500 bg-slate-900/60 border border-slate-800 rounded-lg p-3">
        ⚠️ <strong>Estimation indicative.</strong> Le calcul applique le PFU (flat tax) de 31,4% sur les gains positifs du CTO. Il ne prend pas en compte : l'option pour le barème progressif, le report des moins-values antérieures (reportables 10 ans), l'abattement pour durée de détention (titres acquis avant 2018), ni le crédit d'impôt sur dividendes étrangers. Consulte un conseiller fiscal pour ta déclaration réelle (formulaire 2074).
      </div>
    </div>
  );
};

const FiscalRow = ({ label, value, taxable, exempt, hide }) => (
  <div className="flex justify-between items-center text-sm">
    <span className="text-slate-400">{label}</span>
    <div className="flex items-center gap-2">
      <span className={`font-mono font-semibold ${classFor(value)}`}>{mask(fmtEUR(value, 2), hide)}</span>
      {taxable && value > 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-rose-500/15 text-rose-300">imposable</span>}
      {taxable && value <= 0 && <span className="text-[9px] px-1.5 py-0.5 rounded bg-slate-700 text-slate-400">non imposé</span>}
      {exempt && <span className="text-[9px] px-1.5 py-0.5 rounded bg-emerald-500/15 text-emerald-300">exonéré</span>}
    </div>
  </div>
);


const AssetTable = ({ assets, sectorFn, hide }) => (
  <div className="overflow-x-auto">
    <table className="w-full text-xs">
      <thead><tr className="border-b border-slate-800"><th className="p-2 text-left text-[10px] text-slate-500 uppercase">Actif</th><th className="p-2 text-right text-[10px] text-slate-500 uppercase">P&L totale</th><th className="p-2 text-right text-[10px] text-slate-500 uppercase">%</th><th className="p-2 text-right text-[10px] text-slate-500 uppercase">Trades</th></tr></thead>
      <tbody>
        {assets.map((a, i) => (
          <tr key={i} className={`border-b border-slate-800/60 ${ROW_HOVER}`}>
            <td className="p-2"><div className="flex items-center gap-2"><span className="w-2 h-2 rounded-full" style={{ background: SECTOR_COLORS[sectorFn(a)] || "#64748b" }} /><div className="min-w-0"><div className="truncate font-semibold">{a.name || a.symbol}</div><div className="text-[9px] text-slate-500 font-mono">{a.symbol}</div></div></div></td>
            <td className={`p-2 text-right font-mono tabular-nums font-semibold ${classFor(a.totalPnL)}`}>{mask(fmtEUR(a.totalPnL, 0), hide)}</td>
            <td className={`p-2 text-right font-mono tabular-nums ${classFor(a.returnPct)}`}>{a.invested > 0 ? fmtPct(a.returnPct, 1) : "—"}</td>
            <td className="p-2 text-right font-mono tabular-nums text-slate-400">{a.trades}</td>
          </tr>
        ))}
      </tbody>
    </table>
  </div>
);
function TransactionsView({ d, sectorFn, hide }) {
  const [search, setSearch] = useState("");
  const [kindFilter, setKindFilter] = useState("ALL");
  const [sectorFilter, setSectorFilter] = useState("ALL");
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const [sort, setSort] = useState({ key: "date", dir: "desc" });
  const [detail, setDetail] = useState(null);

  const sectors = useMemo(() => ["ALL", ...new Set(d.transactions.map(t => sectorFn(t)))], [d.transactions, sectorFn]);
  const kinds = ["ALL", "BUY", "SELL", "DIVIDEND", "INTEREST", "SAVEBACK", "DEPOSIT", "WITHDRAWAL", "GIFT", "CARD", "WARRANT_EXERCISE", "TILG", "FREE_RECEIPT"];

  const filtered = useMemo(() => {
    const fromD = from ? new Date(from) : null;
    const toD = to ? new Date(to + "T23:59:59") : null;
    let arr = d.transactions.filter(t => {
      if (kindFilter !== "ALL" && t.kind !== kindFilter) return false;
      if (sectorFilter !== "ALL" && sectorFn(t) !== sectorFilter) return false;
      if (fromD && t.date < fromD) return false;
      if (toD && t.date > toD) return false;
      if (search) { const s = search.toLowerCase(); if (!(t.name || "").toLowerCase().includes(s) && !(t.symbol || "").toLowerCase().includes(s)) return false; }
      return true;
    });
    const dir = sort.dir === "asc" ? 1 : -1;
    arr.sort((a, b) => {
      let va = a[sort.key], vb = b[sort.key];
      if (sort.key === "date") { va = a.date.getTime(); vb = b.date.getTime(); }
      if (sort.key === "sector") { va = sectorFn(a); vb = sectorFn(b); }
      if (va === null || va === undefined || va === "") return 1;
      if (vb === null || vb === undefined || vb === "") return -1;
      if (typeof va === "string") return va.localeCompare(vb) * dir;
      return (va - vb) * dir;
    });
    return arr;
  }, [d.transactions, kindFilter, sectorFilter, from, to, search, sectorFn, sort]);

  const counts = useMemo(() => { const c = {}; for (const t of filtered) c[t.kind] = (c[t.kind] || 0) + 1; return c; }, [filtered]);

  // Stats riches sur le filtre courant
  const fstats = useMemo(() => {
    const buys = filtered.filter(t => t.kind === "BUY");
    const sells = filtered.filter(t => t.kind === "SELL");
    const totalBuy = buys.reduce((s, t) => s + t.amount + t.fee + t.tax, 0);
    const totalSell = sells.reduce((s, t) => s + t.amount - t.fee - t.tax, 0);
    const totalFees = filtered.reduce((s, t) => s + (t.fee || 0), 0);
    const totalTax = filtered.reduce((s, t) => s + (t.tax || 0), 0);
    const totalDiv = filtered.filter(t => t.kind === "DIVIDEND").reduce((s, t) => s + t.amount - t.tax, 0);
    const totalCard = filtered.filter(t => t.kind === "CARD").reduce((s, t) => s + (t.amountRaw < 0 ? Math.abs(t.amountRaw) : -t.amountRaw), 0);
    return { totalBuy, totalSell, totalFees, totalTax, totalDiv, totalCard, buyCount: buys.length, sellCount: sells.length };
  }, [filtered]);

  // Activité par mois (volume de transactions)
  const activityByMonth = useMemo(() => {
    const m = {};
    for (const t of filtered) {
      const k = `${t.date.getFullYear()}-${String(t.date.getMonth() + 1).padStart(2, "0")}`;
      if (!m[k]) m[k] = { month: k, buys: 0, sells: 0, other: 0, count: 0, volBuy: 0, volSell: 0, volOther: 0, volume: 0 };
      m[k].count++;
      // Volume = argent transité (valeur absolue du montant)
      const vol = Math.abs(t.amountRaw != null ? t.amountRaw : (t.amount || 0));
      m[k].volume += vol;
      if (t.kind === "BUY") { m[k].buys++; m[k].volBuy += vol; }
      else if (t.kind === "SELL") { m[k].sells++; m[k].volSell += vol; }
      else { m[k].other++; m[k].volOther += vol; }
    }
    return Object.values(m).sort((a, b) => a.month.localeCompare(b.month));
  }, [filtered]);
  const [activityMode, setActivityMode] = useState("count"); // count | volume

  // Répartition par type (pour donut)
  const byKind = useMemo(() => {
    const m = {};
    for (const t of filtered) {
      if (!m[t.kind]) m[t.kind] = { kind: t.kind, count: 0, volume: 0 };
      m[t.kind].count++;
      m[t.kind].volume += t.amount || 0;
    }
    return Object.values(m).sort((a, b) => b.count - a.count);
  }, [filtered]);

  const resetFilters = () => { setSearch(""); setKindFilter("ALL"); setSectorFilter("ALL"); setFrom(""); setTo(""); };
  const hasFilters = search || kindFilter !== "ALL" || sectorFilter !== "ALL" || from || to;

  const KIND_PALETTE = { BUY: "#3b82f6", SELL: "#10b981", DIVIDEND: "#f59e0b", INTEREST: "#06b6d4", DEPOSIT: "#8b5cf6", WITHDRAWAL: "#f97316", SAVEBACK: "#a855f7", GIFT: "#a855f7", CARD: "#f43f5e", WARRANT_EXERCISE: "#f59e0b", TILG: "#f59e0b", FREE_RECEIPT: "#64748b", OTHER: "#64748b" };

  const exportFilteredCSV = () => {
    const headers = ["Date", "Type", "Actif", "ISIN", "Secteur", "Quantité", "Prix", "Montant", "Frais", "Impôts", "Compte"];
    const rows = filtered.map(t => [fmtDate(t.date), KIND_LABELS[t.kind] || t.kind, t.name || "", t.symbol || "", sectorFn(t), t.shares || "", t.price || "", t.amountRaw || t.amount, t.fee || "", t.tax || "", t.accountType || "DEFAULT"]);
    const csv = [headers.join(","), ...rows.map(r => r.map(c => `"${String(c).replace(/"/g, '""')}"`).join(","))].join("\n");
    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href = url; a.download = `transactions_${new Date().toISOString().slice(0, 10)}.csv`; a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="space-y-4">
      {/* KPIs résumé du filtre */}
      <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-7 gap-3">
        <KPICard label="Transactions" value={filtered.length} icon={Activity} accent="cyan" sub={`/ ${d.transactions.length} total`} />
        <KPICard label="Achats" value={fstats.buyCount} icon={ArrowDownRight} accent="blue" sub={mask(fmtEUR(fstats.totalBuy, 0), hide)} />
        <KPICard label="Ventes" value={fstats.sellCount} icon={ArrowUpRight} accent="emerald" sub={mask(fmtEUR(fstats.totalSell, 0), hide)} />
        <KPICard label="Dividendes" value={mask(fmtEUR(fstats.totalDiv, 0), hide)} icon={Percent} accent="amber" sub={`${counts.DIVIDEND || 0} versements`} />
        <KPICard label="Dépenses CB" value={mask(fmtEUR(fstats.totalCard, 0), hide)} icon={CreditCard} accent="rose" sub={`${counts.CARD || 0} paiements`} />
        <KPICard label="Frais payés" value={mask(fmtEUR(fstats.totalFees, 0), hide)} icon={Coins} accent="orange" />
        <KPICard label="Impôts retenus" value={mask(fmtEUR(fstats.totalTax, 0), hide)} icon={Hash} accent="slate" />
      </div>

      {/* Graphiques activité */}
      {filtered.length > 0 && (
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
          <Card className="lg:col-span-2">
            <div className="flex items-center justify-between flex-wrap gap-2 mb-1">
              <SectionTitle icon={BarChart3} subtitle={activityMode === "count" ? "Nombre de transactions par mois" : "Argent transité par mois (volume des ordres)"}>📊 Activité mensuelle</SectionTitle>
              <div className="inline-flex rounded-lg border border-slate-700 overflow-hidden shrink-0">
                <button onClick={() => setActivityMode("count")} className={`px-3 py-1 text-[11px] font-semibold transition-colors ${activityMode === "count" ? "bg-cyan-500/20 text-cyan-200" : "bg-slate-950 text-slate-400 hover:text-slate-100"}`}># Nombre</button>
                <button onClick={() => setActivityMode("volume")} className={`px-3 py-1 text-[11px] font-semibold transition-colors ${activityMode === "volume" ? "bg-cyan-500/20 text-cyan-200" : "bg-slate-950 text-slate-400 hover:text-slate-100"}`}>€ Volume</button>
              </div>
            </div>
            <ResponsiveContainer width="100%" height={240}>
              <BarChart data={activityByMonth}>
                <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" vertical={false} />
                <XAxis dataKey="month" stroke="#64748b" fontSize={10} tickFormatter={(v) => v.slice(2)} />
                <YAxis stroke="#64748b" fontSize={10} tickFormatter={(v) => activityMode === "volume" ? `${(v / 1000).toFixed(0)}k` : v} />
                <Tooltip content={({ active, payload, label }) => {
                  if (!active || !payload?.length) return null;
                  const m = payload[0].payload;
                  return (
                    <div className="bg-slate-950 border border-slate-700 rounded-lg p-2.5 shadow-xl text-xs font-mono min-w-[160px]">
                      <div className="text-slate-400 uppercase tracking-widest text-[10px] mb-1">{label}</div>
                      {activityMode === "count" ? (
                        <>
                          <div className="flex justify-between gap-3"><span className="text-blue-400">Achats</span><span>{m.buys}</span></div>
                          <div className="flex justify-between gap-3"><span className="text-emerald-400">Ventes</span><span>{m.sells}</span></div>
                          <div className="flex justify-between gap-3"><span className="text-slate-400">Autres</span><span>{m.other}</span></div>
                          <div className="flex justify-between gap-3 pt-1 border-t border-slate-700 mt-1 font-bold"><span>Total</span><span>{m.count}</span></div>
                        </>
                      ) : (
                        <>
                          <div className="flex justify-between gap-3"><span className="text-blue-400">Achats</span><span>{mask(fmtEUR(m.volBuy, 0), hide)}</span></div>
                          <div className="flex justify-between gap-3"><span className="text-emerald-400">Ventes</span><span>{mask(fmtEUR(m.volSell, 0), hide)}</span></div>
                          <div className="flex justify-between gap-3"><span className="text-slate-400">Autres</span><span>{mask(fmtEUR(m.volOther, 0), hide)}</span></div>
                          <div className="flex justify-between gap-3 pt-1 border-t border-slate-700 mt-1 font-bold"><span>Volume total</span><span>{mask(fmtEUR(m.volume, 0), hide)}</span></div>
                        </>
                      )}
                    </div>
                  );
                }} />
                {activityMode === "count" ? (
                  <>
                    <Bar dataKey="buys" stackId="a" fill="#3b82f6" name="Achats" />
                    <Bar dataKey="sells" stackId="a" fill="#10b981" name="Ventes" />
                    <Bar dataKey="other" stackId="a" fill="#64748b" name="Autres" radius={[3, 3, 0, 0]} />
                  </>
                ) : (
                  <>
                    <Bar dataKey="volBuy" stackId="v" fill="#3b82f6" name="Achats €" />
                    <Bar dataKey="volSell" stackId="v" fill="#10b981" name="Ventes €" />
                    <Bar dataKey="volOther" stackId="v" fill="#64748b" name="Autres €" radius={[3, 3, 0, 0]} />
                  </>
                )}
              </BarChart>
            </ResponsiveContainer>
          </Card>
          <Card>
            <SectionTitle icon={PieIcon} subtitle="Répartition par type">🥧 Types</SectionTitle>
            <ResponsiveContainer width="100%" height={180}>
              <PieChart>
                <Pie data={byKind} dataKey="count" nameKey="kind" innerRadius={45} outerRadius={80} paddingAngle={2} stroke="#0f172a" strokeWidth={2}>
                  {byKind.map((e, i) => <Cell key={i} fill={KIND_PALETTE[e.kind] || "#64748b"} />)}
                </Pie>
                <Tooltip content={({ active, payload }) => {
                  if (!active || !payload?.length) return null;
                  const e = payload[0].payload;
                  return <div className="bg-slate-950 border border-slate-700 rounded-lg p-2 shadow-xl text-xs font-mono"><span className="text-slate-200">{KIND_LABELS[e.kind] || e.kind}</span> : {e.count} tx</div>;
                }} />
              </PieChart>
            </ResponsiveContainer>
            <div className="space-y-1 mt-2 max-h-[60px] overflow-y-auto">
              {byKind.slice(0, 5).map((e, i) => (
                <div key={i} className="flex items-center gap-2 text-[10px]">
                  <span className="w-2 h-2 rounded-sm" style={{ background: KIND_PALETTE[e.kind] || "#64748b" }} />
                  <span className="flex-1 text-slate-400">{KIND_LABELS[e.kind] || e.kind}</span>
                  <span className="font-mono text-slate-300">{e.count}</span>
                </div>
              ))}
            </div>
          </Card>
        </div>
      )}

      {/* Filtres */}
      <Card>
        <div className="grid grid-cols-1 md:grid-cols-5 gap-3">
          <div><label className="text-[10px] uppercase tracking-widest text-slate-500 block mb-1">Recherche</label><div className="relative"><Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-500" /><input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Nom ou ISIN…" className="w-full bg-slate-950 border border-slate-700 rounded pl-8 pr-3 py-1.5 text-sm focus:outline-none focus:border-cyan-500" /></div></div>
          <div><label className="text-[10px] uppercase tracking-widest text-slate-500 block mb-1">Type</label><select value={kindFilter} onChange={(e) => setKindFilter(e.target.value)} className="w-full bg-slate-950 border border-slate-700 rounded px-3 py-1.5 text-sm focus:outline-none focus:border-cyan-500">{kinds.map(t => <option key={t} value={t}>{t === "ALL" ? "Tous" : KIND_LABELS[t] || t}{counts[t] ? ` (${counts[t]})` : ""}</option>)}</select></div>
          <div><label className="text-[10px] uppercase tracking-widest text-slate-500 block mb-1">Secteur</label><select value={sectorFilter} onChange={(e) => setSectorFilter(e.target.value)} className="w-full bg-slate-950 border border-slate-700 rounded px-3 py-1.5 text-sm focus:outline-none focus:border-cyan-500">{sectors.map(s => <option key={s} value={s}>{s === "ALL" ? "Tous" : s}</option>)}</select></div>
          <div><label className="text-[10px] uppercase tracking-widest text-slate-500 block mb-1">Depuis</label><input type="date" value={from} onChange={(e) => setFrom(e.target.value)} className="w-full bg-slate-950 border border-slate-700 rounded px-3 py-1.5 text-sm focus:outline-none focus:border-cyan-500" /></div>
          <div><label className="text-[10px] uppercase tracking-widest text-slate-500 block mb-1">Jusqu'à</label><input type="date" value={to} onChange={(e) => setTo(e.target.value)} className="w-full bg-slate-950 border border-slate-700 rounded px-3 py-1.5 text-sm focus:outline-none focus:border-cyan-500" /></div>
        </div>
        <div className="mt-3 flex items-center justify-between flex-wrap gap-2">
          <div className="flex items-center gap-2">
            {hasFilters && <button onClick={resetFilters} className="text-[11px] px-2.5 py-1 bg-slate-800 hover:bg-slate-700 rounded text-slate-300 transition-colors flex items-center gap-1"><RefreshCw className="w-3 h-3" /> Réinitialiser</button>}
            <span className="text-xs text-slate-500 font-mono">{filtered.length} résultat{filtered.length > 1 ? "s" : ""}</span>
          </div>
          <button onClick={exportFilteredCSV} className="text-[11px] px-2.5 py-1 bg-emerald-500/10 border border-emerald-500/30 hover:bg-emerald-500/20 rounded text-emerald-200 transition-colors flex items-center gap-1"><Download className="w-3 h-3" /> Exporter CSV</button>
        </div>
      </Card>

      {/* Tableau */}
      <div className="bg-slate-900/70 border border-slate-800 rounded-xl overflow-hidden">
        <div className="overflow-x-auto max-h-[600px] overflow-y-auto">
          <table className="w-full text-xs">
            <thead className="sticky top-0 bg-slate-900 z-10">
              <tr className="border-b border-slate-800">
                {[{ k: "date", l: "Date" }, { k: "kind", l: "Type" }, { k: "name", l: "Actif" }, { k: "symbol", l: "ISIN" }, { k: "sector", l: "Secteur" }, { k: "shares", l: "Qté", n: true }, { k: "price", l: "Prix", n: true }, { k: "amount", l: "Montant", n: true }, { k: "fee", l: "Frais", n: true }, { k: "tax", l: "Impôts", n: true }].map(c => (
                  <th key={c.k} className={`p-2 font-semibold text-slate-400 uppercase tracking-wider text-[10px] cursor-pointer hover:text-slate-100 whitespace-nowrap ${c.n ? "text-right" : "text-left"}`} onClick={() => setSort(s => ({ key: c.k, dir: s.key === c.k && s.dir === "desc" ? "asc" : "desc" }))}>
                    {c.l} {sort.key === c.k && (sort.dir === "desc" ? "↓" : "↑")}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map((t, i) => {
                const sec = sectorFn(t);
                const color = KIND_COLORS[t.kind] || "slate";
                const showSector = t.kind === "BUY" || t.kind === "SELL" || t.kind === "DIVIDEND" || t.kind === "SAVEBACK" || t.kind === "WARRANT_EXERCISE";
                return (
                  <tr key={i} onClick={() => setDetail(t)} className={`border-b border-slate-800/60 cursor-pointer ${ROW_HOVER}`}>
                    <td className="p-2 font-mono text-slate-400 whitespace-nowrap">{fmtDateShort(t.date)}</td>
                    <td className="p-2"><span className={`inline-block px-1.5 py-0.5 rounded text-[10px] font-semibold border ${PILL_CLASS[color]}`}>{KIND_LABELS[t.kind] || t.kind}</span></td>
                    <td className="p-2 truncate max-w-[200px] text-slate-100">{t.name || "—"}</td>
                    <td className="p-2 font-mono text-slate-500 text-[10px]">{t.symbol || "—"}</td>
                    <td className="p-2">{showSector ? (<span className="inline-flex items-center gap-1"><span className="w-2 h-2 rounded-full" style={{ background: SECTOR_COLORS[sec] || "#64748b" }} /><span className="text-[10px] text-slate-400">{sec}</span></span>) : <span className="text-slate-600">—</span>}</td>
                    <td className="p-2 font-mono tabular-nums text-right text-slate-300">{t.shares ? fmtNum(t.shares, 4) : "—"}</td>
                    <td className="p-2 font-mono tabular-nums text-right text-slate-300">{t.price ? fmtEUR(t.price, 4) : "—"}</td>
                    <td className={`p-2 font-mono tabular-nums text-right font-semibold ${t.kind === "BUY" || (t.kind === "CARD" && t.amountRaw < 0) ? "text-rose-400" : t.kind === "SELL" || t.kind === "DEPOSIT" || t.kind === "DIVIDEND" || t.kind === "INTEREST" || t.kind === "SAVEBACK" || t.kind === "GIFT" ? "text-emerald-400" : "text-slate-200"}`}>{mask(fmtEUR(t.amount, 2), hide)}</td>
                    <td className="p-2 font-mono tabular-nums text-right text-slate-500">{t.fee ? fmtEUR(t.fee, 2) : "—"}</td>
                    <td className="p-2 font-mono tabular-nums text-right text-slate-500">{t.tax ? fmtEUR(t.tax, 2) : "—"}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          {filtered.length === 0 && <div className="p-10 text-center text-slate-500">Aucune transaction ne correspond aux filtres</div>}
        </div>
      </div>

      {/* Modal détail transaction */}
      {detail && <TransactionDetailModal tx={detail} sector={sectorFn(detail)} onClose={() => setDetail(null)} hide={hide} />}
    </div>
  );
}

const TransactionDetailModal = ({ tx, sector, onClose, hide }) => {
  useEffect(() => {
    const orig = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => { document.body.style.overflow = orig; };
  }, []);
  const color = KIND_COLORS[tx.kind] || "slate";
  const rows = [
    ["Type", KIND_LABELS[tx.kind] || tx.kind],
    ["Date & heure", tx.date.toLocaleString("fr-FR")],
    ["Actif", tx.name || "—"],
    ["ISIN / Symbole", tx.symbol || "—"],
    ["Secteur", sector],
    ["Classe d'actif", tx.assetClass || "—"],
    ["Compte", tx.accountType === "PEA" ? "PEA (exonéré d'impôts)" : "CTO (imposable 31,4%)"],
    ["Quantité", tx.shares ? fmtNum(tx.shares, 6) : "—"],
    ["Prix unitaire", tx.price ? fmtEUR(tx.price, 6) : "—"],
    ["Montant", mask(fmtEUR(tx.amount, 2), hide)],
    ["Frais", tx.fee ? fmtEUR(tx.fee, 2) : "0,00 €"],
    ["Impôts retenus", tx.tax ? fmtEUR(tx.tax, 2) : "0,00 €"],
    ["Devise", tx.currency || "EUR"],
  ];
  if (tx.description) rows.push(["Description", tx.description]);
  if (tx.txId) rows.push(["ID transaction", tx.txId]);

  return (
    <div className="fixed inset-0 z-50 flex items-end md:items-center justify-center bg-black/70 backdrop-blur-sm p-0 md:p-4" onClick={onClose}>
      <div className="bg-slate-900 border border-slate-700 rounded-t-2xl md:rounded-2xl w-full md:max-w-lg max-h-[92vh] overflow-hidden flex flex-col" onClick={e => e.stopPropagation()}>
        <div className="p-5 border-b border-slate-800">
          <div className="flex items-start justify-between gap-3">
            <div className="min-w-0">
              <span className={`inline-block px-2 py-0.5 rounded text-[10px] font-semibold border ${PILL_CLASS[color]} mb-2`}>{KIND_LABELS[tx.kind] || tx.kind}</span>
              <div className="text-xl font-bold truncate">{tx.name || KIND_LABELS[tx.kind] || "Transaction"}</div>
              <div className="text-xs text-slate-400 font-mono mt-0.5">{fmtDateLong(tx.date)}</div>
            </div>
            <button onClick={onClose} className="p-2 bg-slate-800 hover:bg-slate-700 rounded text-slate-300 text-xl leading-none shrink-0">✕</button>
          </div>
          <div className={`mt-3 text-3xl font-mono font-black tabular-nums ${tx.kind === "BUY" || (tx.kind === "CARD" && tx.amountRaw < 0) ? "text-rose-400" : tx.kind === "SELL" || tx.kind === "DEPOSIT" || tx.kind === "DIVIDEND" || tx.kind === "INTEREST" || tx.kind === "SAVEBACK" ? "text-emerald-400" : "text-slate-200"}`}>
            {mask(fmtEUR(tx.amount, 2), hide)}
          </div>
        </div>
        <div className="overflow-y-auto flex-1 p-3">
          <table className="w-full text-sm">
            <tbody>
              {rows.map(([k, v], i) => (
                <tr key={i} className="border-b border-slate-800/40">
                  <td className="py-2 px-2 text-slate-500 text-xs uppercase tracking-wider w-2/5">{k}</td>
                  <td className="py-2 px-2 text-slate-100 font-mono text-right break-all">{v}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// ═══════════════ SETTINGS ═══════════════

function Settings({ onImport, transactions, onClear, sectorOverrides, onSectorOverride, derived, sectorFn }) {
  const inputRef = useRef(null);
  const [confirmClear, setConfirmClear] = useState(false);

  const exportXLSX = () => {
    const wb = XLSX.utils.book_new();
    const txData = transactions.map(t => ({ Date: fmtDate(t.date), Type: t.kind, Nom: t.name, ISIN: t.symbol, Secteur: sectorFn(t), "Quantité": t.shares, "Prix": t.price, Montant: t.amount, Frais: t.fee, Impôts: t.tax, Devise: t.currency }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(txData), "Transactions");
    const trData = derived.realizedTrades.map(t => ({ Actif: t.name, ISIN: t.symbol, Secteur: sectorFn(t), "Achat": fmtDate(t.buyDate), "Vente": fmtDate(t.sellDate), "Durée (j)": t.holdingDays, "Qté": t.shares, "Prix achat": t.buyPrice, "Prix vente": t.sellPrice, "Coût": t.cost, "Produit": t.proceeds, "P&L": t.pnl, "P&L %": t.pnlPct }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(trData), "Trades réalisés");
    XLSX.writeFile(wb, `portfolio_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const exportJSON = () => {
    const data = { transactions: transactions.map(t => ({ ...t, date: t.date.toISOString() })), sectorOverrides, exportedAt: new Date().toISOString() };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href = url; a.download = `backup_${new Date().toISOString().slice(0, 10)}.json`; a.click();
    URL.revokeObjectURL(url);
  };

  const uniqueAssets = useMemo(() => {
    const m = {};
    for (const t of transactions) if ((t.kind === "BUY" || t.kind === "SELL") && t.symbol) m[t.symbol] = m[t.symbol] || { symbol: t.symbol, name: t.name, sector: sectorFn(t) };
    return Object.values(m).sort((a, b) => (a.name || "").localeCompare(b.name || ""));
  }, [transactions, sectorFn]);

  const allSectors = Object.keys(SECTOR_COLORS);

  return (
    <div className="space-y-6">
      <Card>
        <SectionTitle icon={Upload}>Import</SectionTitle>
        <div onClick={() => inputRef.current?.click()} className="border-2 border-dashed border-slate-700 rounded-lg p-8 text-center cursor-pointer hover:border-cyan-500 hover:bg-cyan-500/5 transition-colors">
          <Upload className="w-8 h-8 mx-auto mb-2 text-slate-500" />
          <div className="text-sm font-semibold">Ajouter des fichiers Trade Republic</div>
          <div className="text-xs text-slate-500 mt-1">Les doublons sont automatiquement ignorés</div>
          <input ref={inputRef} type="file" multiple accept=".xlsx,.xls,.csv" onChange={(e) => e.target.files?.length && onImport(Array.from(e.target.files))} className="hidden" />
        </div>
      </Card>

      <Card>
        <SectionTitle icon={Download}>Export</SectionTitle>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
          <button onClick={exportXLSX} className="flex items-center gap-3 px-4 py-3 bg-emerald-500/10 border border-emerald-500/30 hover:bg-emerald-500/20 rounded-lg text-emerald-200 transition-colors">
            <FileText className="w-5 h-5" /><div className="text-left"><div className="font-semibold text-sm">Analyse Excel (.xlsx)</div><div className="text-xs text-emerald-300/60">Transactions + Trades</div></div>
          </button>
          <button onClick={exportJSON} className="flex items-center gap-3 px-4 py-3 bg-cyan-500/10 border border-cyan-500/30 hover:bg-cyan-500/20 rounded-lg text-cyan-200 transition-colors">
            <Database className="w-5 h-5" /><div className="text-left"><div className="font-semibold text-sm">Sauvegarde (.json)</div><div className="text-xs text-cyan-300/60">Pour réimport</div></div>
          </button>
        </div>
      </Card>

      <Card>
        <SectionTitle icon={PieIcon}>Mapping secteurs</SectionTitle>
        <div className="text-xs text-slate-400 mb-3">Force la classification d'un ISIN.</div>
        <div className="max-h-[400px] overflow-y-auto">
          <table className="w-full text-xs">
            <thead className="sticky top-0 bg-slate-900"><tr className="border-b border-slate-800"><th className="p-2 text-left text-[10px] text-slate-500 uppercase">Actif</th><th className="p-2 text-left text-[10px] text-slate-500 uppercase">ISIN</th><th className="p-2 text-left text-[10px] text-slate-500 uppercase">Secteur</th></tr></thead>
            <tbody>
              {uniqueAssets.map((a, i) => (
                <tr key={i} className={`border-b border-slate-800/60 ${ROW_HOVER}`}>
                  <td className="p-2 text-slate-200 truncate max-w-[200px]">{a.name}</td>
                  <td className="p-2 font-mono text-slate-500 text-[10px]">{a.symbol}</td>
                  <td className="p-2"><select value={a.sector} onChange={(e) => onSectorOverride(a.symbol, e.target.value)} className="bg-slate-950 border border-slate-700 rounded px-2 py-1 text-xs">{allSectors.map(s => <option key={s} value={s}>{s}</option>)}</select></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>

      <div className="bg-rose-500/5 border border-rose-500/30 rounded-xl p-5">
        <SectionTitle icon={AlertCircle}>Zone dangereuse</SectionTitle>
        <div className="flex items-center justify-between gap-4 flex-wrap">
          <div className="text-xs text-slate-400">Supprime toutes les transactions ({transactions.length} lignes).</div>
          {!confirmClear ? (
            <button onClick={() => setConfirmClear(true)} className="flex items-center gap-2 px-4 py-2 bg-rose-500/20 border border-rose-500/40 text-rose-200 rounded hover:bg-rose-500/30 text-xs font-semibold"><Trash2 className="w-3.5 h-3.5" /> Tout effacer</button>
          ) : (
            <div className="flex gap-2">
              <button onClick={() => { onClear(); setConfirmClear(false); }} className="px-4 py-2 bg-rose-500 text-white rounded text-xs font-semibold hover:bg-rose-600">Confirmer</button>
              <button onClick={() => setConfirmClear(false)} className="px-4 py-2 bg-slate-700 text-slate-200 rounded text-xs">Annuler</button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ═══════════════ FRAIS DE COURTAGE (carte sur Dashboard) ═══════════════

function FraisCourtageCard({ d, hide }) {
  const { feesEstimation, transactions } = d;
  const f = feesEstimation;
  const totalOrders = f.plannedCount + f.spotBuyCount + f.spotSellCount;
  if (totalOrders === 0) return null;
  const economyVsAllSpot = totalOrders * 1.0; // ce que tu paierais si tout était spot
  const actuallyPaid = f.totalEstimated;
  const saved = economyVsAllSpot - actuallyPaid;

  return (
    <div className="bg-gradient-to-br from-slate-900 to-slate-950 border border-slate-800 rounded-xl p-5">
      <div className="flex items-center gap-2 mb-4">
        <Coins className="w-4 h-4 text-amber-400" />
        <h2 className="text-sm font-semibold uppercase tracking-widest text-slate-200">💰 Frais de courtage estimés</h2>
        <span className="text-[10px] text-slate-500 font-mono">(plan d'épargne 0€ · spot 1€)</span>
      </div>
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <div className="bg-amber-500/5 border border-amber-500/30 rounded-lg p-3">
          <div className="text-[10px] uppercase tracking-widest text-amber-300 mb-1">Frais payés</div>
          <div className="text-2xl font-mono font-black text-amber-300 tabular-nums">{mask(fmtEUR(actuallyPaid, 0), hide)}</div>
          <div className="text-[10px] text-slate-500 mt-1 font-mono">Sur {totalOrders} ordres</div>
        </div>
        <div className="bg-emerald-500/5 border border-emerald-500/30 rounded-lg p-3">
          <div className="text-[10px] uppercase tracking-widest text-emerald-300 mb-1">Économisé</div>
          <div className="text-2xl font-mono font-black text-emerald-400 tabular-nums">{mask(fmtEUR(saved, 0), hide)}</div>
          <div className="text-[10px] text-slate-500 mt-1 font-mono">Grâce au plan d'épargne</div>
        </div>
        <div className="bg-cyan-500/5 border border-cyan-500/30 rounded-lg p-3">
          <div className="text-[10px] uppercase tracking-widest text-cyan-300 mb-1">Plans d'épargne</div>
          <div className="text-2xl font-mono font-black text-cyan-300 tabular-nums">{f.plannedCount}</div>
          <div className="text-[10px] text-slate-500 mt-1 font-mono">Ordres gratuits</div>
        </div>
        <div className="bg-rose-500/5 border border-rose-500/30 rounded-lg p-3">
          <div className="text-[10px] uppercase tracking-widest text-rose-300 mb-1">Ordres spot</div>
          <div className="text-2xl font-mono font-black text-rose-300 tabular-nums">{f.spotBuyCount + f.spotSellCount}</div>
          <div className="text-[10px] text-slate-500 mt-1 font-mono">{f.spotBuyCount} achats · {f.spotSellCount} ventes</div>
        </div>
      </div>
      {f.totalDeclared < actuallyPaid - 0.5 && (
        <div className="mt-3 text-[11px] text-slate-400 bg-slate-950/50 border border-slate-800 rounded p-2 font-mono">
          ℹ Le CSV Trade Republic ne contient {fmtEUR(f.totalDeclared, 2)} de frais déclarés, mais d'après tes transactions tu as payé environ <span className="text-amber-300 font-semibold">{fmtEUR(actuallyPaid, 2)}</span> au total. La différence ({fmtEUR(f.hiddenFees, 2)}) correspond aux 1€ par ordre spot non listés explicitement dans le fichier.
        </div>
      )}
    </div>
  );
}

// ═══════════════ FINANCES VIEW : couverture CB par les gains ═══════════════

// ═══════════════ FINANCES VIEW V2 : ultra détaillé ═══════════════

function FinancesView({ d, hide }) {
  const { monthlyArr, weeklyArr, lifetimeCoverage, lifetimeSurplus, cardCategories, cardSpendNet, totalGains, cardTx, totalDeposits, taxation, feesEstimation, runwayMonths, avgMonthlySpendAllTime } = d;
  const [period, setPeriod] = useState("current");
  const [chartMode, setChartMode] = useState("comparison");
  const [granularity, setGranularity] = useState("month"); // month | week
  const [openCategory, setOpenCategory] = useState(null);
  const [view, setView] = useState("vue"); // vue | depenses | gains | projection

  // ═════ Données : 3 horizons clés ═════
  const horizons = useMemo(() => {
    if (!monthlyArr.length) return { current: null, last3: null, last12: null, lifetime: { coverage: 0, gains: 0, spend: 0, surplus: 0 } };
    const now = new Date();
    const currentKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
    const currentMonth = monthlyArr.find(m => m.month === currentKey);

    const sumRange = (months) => {
      const cutoff = new Date(now.getFullYear(), now.getMonth() - months + 1, 1);
      const filtered = monthlyArr.filter(m => {
        const [y, mo] = m.month.split("-").map(Number);
        return new Date(y, mo - 1, 1) >= cutoff;
      });
      const sum = filtered.reduce((acc, m) => ({
        gains: acc.gains + m.gains, spend: acc.spend + m.cardNet,
        realizedPnL: acc.realizedPnL + m.realizedPnL,
        dividends: acc.dividends + m.dividends,
        interests: acc.interests + m.interests,
        savebacks: acc.savebacks + m.savebacks,
      }), { gains: 0, spend: 0, realizedPnL: 0, dividends: 0, interests: 0, savebacks: 0 });
      return {
        ...sum,
        surplus: sum.gains - sum.spend,
        coverage: sum.spend > 0 ? (sum.gains / sum.spend) * 100 : (sum.gains > 0 ? 999 : 0),
        monthCount: filtered.length,
      };
    };
    return {
      current: currentMonth || { month: currentKey, gains: 0, cardNet: 0, surplus: 0, coverage: 0, realizedPnL: 0, dividends: 0, interests: 0, savebacks: 0 },
      last3: sumRange(3),
      last12: sumRange(12),
      lifetime: { coverage: lifetimeCoverage, gains: totalGains, spend: cardSpendNet, surplus: lifetimeSurplus, monthCount: monthlyArr.length },
    };
  }, [monthlyArr, lifetimeCoverage, lifetimeSurplus, totalGains, cardSpendNet]);

  const cur = horizons.current;
  const l3 = horizons.last3;
  const l12 = horizons.last12;
  const lt = horizons.lifetime;

  // ═════ Série active selon granularité (mois ou semaine), avec clé/label unifiés ═════
  const baseSeries = useMemo(() => {
    if (granularity === "week") {
      return weeklyArr.map(w => ({ ...w, key: w.week, dateRef: new Date(w.week) }));
    }
    return monthlyArr.map(m => { const [y, mo] = m.month.split("-").map(Number); return { ...m, key: m.month, dateRef: new Date(y, mo - 1, 1) }; });
  }, [granularity, monthlyArr, weeklyArr]);

  // ═════ Données filtrées sur la période sélectionnée ═════
  const filtered = useMemo(() => {
    if (!baseSeries.length) return [];
    const now = new Date();
    if (period === "all") return baseSeries;
    if (period === "current") {
      if (granularity === "week") {
        // semaine en cours
        const day = (now.getDay() + 6) % 7;
        const monday = new Date(now); monday.setDate(now.getDate() - day); monday.setHours(0, 0, 0, 0);
        const k = `${monday.getFullYear()}-${String(monday.getMonth() + 1).padStart(2, "0")}-${String(monday.getDate()).padStart(2, "0")}`;
        return baseSeries.filter(s => s.key === k);
      }
      const currentKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
      return baseSeries.filter(s => s.key === currentKey);
    }
    // périodes glissantes : on raisonne en nombre d'unités
    const unitsBack = granularity === "week"
      ? ({ "3m": 13, "6m": 26, "12m": 52 }[period] || 13)
      : ({ "3m": 3, "6m": 6, "12m": 12 }[period] || 3);
    return baseSeries.slice(-unitsBack);
  }, [baseSeries, period, granularity]);

  const totals = filtered.reduce((acc, m) => ({
    cardSpend: acc.cardSpend + m.cardSpend, cardRefund: acc.cardRefund + m.cardRefund, cardNet: acc.cardNet + m.cardNet,
    realizedPnL: acc.realizedPnL + m.realizedPnL, dividends: acc.dividends + m.dividends, interests: acc.interests + m.interests,
    savebacks: acc.savebacks + m.savebacks, gifts: acc.gifts + m.gifts, gains: acc.gains + m.gains, surplus: acc.surplus + m.surplus,
  }), { cardSpend: 0, cardRefund: 0, cardNet: 0, realizedPnL: 0, dividends: 0, interests: 0, savebacks: 0, gifts: 0, gains: 0, surplus: 0 });
  const periodCoverage = totals.cardNet > 0 ? (totals.gains / totals.cardNet) * 100 : (totals.gains > 0 ? 999 : 0);

  const monthsActive = filtered.filter(m => m.cardNet > 0).length;
  const avgMonthlySpend = monthsActive > 0 ? totals.cardNet / monthsActive : 0;
  const avgMonthlyGains = filtered.length > 0 ? totals.gains / filtered.length : 0;
  const monthsCovered = filtered.filter(m => m.gains >= m.cardNet && m.cardNet > 0).length;
  const bestMonth = filtered.reduce((b, m) => (!b || m.surplus > b.surplus) ? m : b, null);
  const worstMonth = filtered.reduce((w, m) => (!w || m.surplus < w.surplus) ? m : w, null);

  // ═════ Projection : combien il faut générer par mois pour atteindre 100% ═════
  const projection = useMemo(() => {
    const avgMonthlyGain12 = l12.monthCount > 0 ? l12.gains / l12.monthCount : 0;
    const avgMonthlySpend12 = l12.monthCount > 0 ? l12.spend / l12.monthCount : 0;
    const gap = avgMonthlySpend12 - avgMonthlyGain12;
    return { avgGain: avgMonthlyGain12, avgSpend: avgMonthlySpend12, gap };
  }, [l12]);

  // ═════ Label court selon granularité (mois "26-04" ou semaine "S.14") ═════
  const labelOf = (item) => {
    if (granularity === "week") {
      const d = item.dateRef || new Date(item.key);
      // numéro de semaine ISO
      const tmp = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
      tmp.setUTCDate(tmp.getUTCDate() - ((tmp.getUTCDay() + 6) % 7) + 3);
      const ft = new Date(Date.UTC(tmp.getUTCFullYear(), 0, 4));
      const wn = 1 + Math.round(((tmp - ft) / 86400000 - 3 + ((ft.getUTCDay() + 6) % 7)) / 7);
      return `S${wn}`;
    }
    return (item.month || item.key || "").slice(2);
  };
  const fullLabelOf = (item) => {
    if (granularity === "week") { const d = item.dateRef || new Date(item.key); const end = new Date(d); end.setDate(d.getDate() + 6); return `${fmtDateShort(d)} → ${fmtDateShort(end)}`; }
    return item.month || item.key;
  };

  // ═════ Cumulative data (surplus cumulé = courbe d'épargne nette) ═════
  const cumulativeData = useMemo(() => {
    let cumGain = 0, cumSpend = 0;
    return filtered.map(m => {
      cumGain += m.gains; cumSpend += m.cardNet;
      return { label: labelOf(m), gainsCumul: cumGain, spendCumul: cumSpend, soldeCumul: cumGain - cumSpend };
    });
  }, [filtered, granularity]);

  // ═════ Sources de gains : breakdown détaillé ═════
  const gainBreakdown = {
    "P&L réalisée": totals.realizedPnL, "Dividendes": totals.dividends, "Intérêts": totals.interests,
    "Saveback": totals.savebacks, "Gifts": totals.gifts,
  };
  const gainBreakdownArr = Object.entries(gainBreakdown).filter(([_, v]) => v !== 0).map(([k, v]) => ({ name: k, value: v }));

  // ═════ Top 10 dépenses ═════
  const top10Expenses = useMemo(() => [...cardTx].filter(t => t.amountRaw < 0).sort((a, b) => a.amountRaw - b.amountRaw).slice(0, 10), [cardTx]);

  // ═════ Catégories enrichies avec transactions détaillées par catégorie ═════
  const categorizeMerchant = (name) => {
    const n = (name || "").toLowerCase();
    if (/uber|bolt|sncf|trainline|flixbus|blablacar|tgv|ouigo|ratp|jcdecaux|free.flotte|lime|tier|dott/i.test(n)) return "Transport";
    if (/airbnb|hotel|booking|hostel|expedia/i.test(n)) return "Hébergement";
    if (/restaurant|mcdo|burger|kfc|sushi|pizza|food|deliveroo|uber.eats|just.eat|frichti|too.good/i.test(n)) return "Restaurants";
    if (/carrefour|lidl|monoprix|leclerc|auchan|intermarche|casino|franprix|biocoop|naturalia|picard/i.test(n)) return "Courses";
    if (/amazon|cdiscount|fnac|darty|zalando|asos|h&m|zara|uniqlo|decathlon/i.test(n)) return "Shopping";
    if (/netflix|spotify|deezer|disney|prime.video|youtube|apple.tv|canal/i.test(n)) return "Abonnements";
    if (/orange|sfr|bouygues|free|sosh|iliad/i.test(n)) return "Télécom";
    if (/edf|enedis|engie|veolia/i.test(n)) return "Énergie";
    if (/pharmacie|doctolib|medecin|hospital|sante/i.test(n)) return "Santé";
    if (/cinema|ugc|pathe|gaumont|theatre|concert|fitnesspark|basicfit|ticketmaster/i.test(n)) return "Loisirs";
    if (/atm|withdraw|retrait/i.test(n)) return "Retrait DAB";
    return "Autres";
  };
  const txByCategory = useMemo(() => {
    const m = {};
    for (const t of cardTx) {
      if (t.amountRaw >= 0) continue;
      const cat = categorizeMerchant(t.name);
      if (!m[cat]) m[cat] = { category: cat, total: 0, count: 0, transactions: [] };
      m[cat].total += Math.abs(t.amountRaw); m[cat].count++;
      m[cat].transactions.push(t);
    }
    for (const k of Object.keys(m)) m[k].transactions.sort((a, b) => a.amountRaw - b.amountRaw);
    return m;
  }, [cardTx]);
  const sortedCategories = useMemo(() => Object.values(txByCategory).sort((a, b) => b.total - a.total), [txByCategory]);

  // ═════ Évolution mois par mois par catégorie (heatmap mental) ═════
  const monthlyByCategory = useMemo(() => {
    const m = {};
    for (const t of cardTx) {
      if (t.amountRaw >= 0) continue;
      const cat = categorizeMerchant(t.name);
      const mkey = `${t.date.getFullYear()}-${String(t.date.getMonth() + 1).padStart(2, "0")}`;
      if (!m[mkey]) m[mkey] = { month: mkey };
      m[mkey][cat] = (m[mkey][cat] || 0) + Math.abs(t.amountRaw);
    }
    return Object.values(m).sort((a, b) => a.month.localeCompare(b.month));
  }, [cardTx]);

  const PALETTE = ["#06b6d4", "#a855f7", "#ec4899", "#10b981", "#f59e0b", "#ef4444", "#3b82f6", "#14b8a6", "#f97316", "#8b5cf6", "#64748b"];

  return (
    <div className="space-y-4">
      {/* ═══════ HERO : 3 cartes horizons ═══════ */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
        <HorizonCard label="★ Ce mois" sublabel={cur ? cur.month : "—"} data={cur && { coverage: cur.coverage, gains: cur.gains, spend: cur.cardNet, surplus: cur.surplus, monthCount: 1 }} accent="cyan" hide={hide} highlight />
        <HorizonCard label="3 derniers mois" sublabel={`${l3.monthCount} mois`} data={l3} accent="violet" hide={hide} />
        <HorizonCard label="All-time" sublabel={`${lt.monthCount} mois`} data={lt} accent="slate" hide={hide} />
      </div>

      {/* ═══════ COMPARAISON MOIS vs MOIS PRÉCÉDENT ═══════ */}
      <MonthComparisonBar comparison={d.monthComparison} hide={hide} />

      {/* ═══════ NAVIGATION INTERNE ═══════ */}
      <div className="flex gap-1 bg-slate-900 p-1 rounded-lg border border-slate-800 overflow-x-auto">
        {[
          { id: "vue", label: "Vue d'ensemble", icon: Activity },
          { id: "depenses", label: "Mes dépenses", icon: CreditCard },
          { id: "gains", label: "Mes gains", icon: Coins },
          { id: "projection", label: "Projection", icon: Target },
        ].map(t => (
          <button key={t.id} onClick={() => setView(t.id)}
            className={`flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded transition-colors whitespace-nowrap ${view === t.id ? "bg-cyan-500/20 text-cyan-200" : "text-slate-400 hover:text-slate-100"}`}>
            <t.icon className="w-3.5 h-3.5" /> {t.label}
          </button>
        ))}
      </div>

      {/* ═══════ VIEW : VUE D'ENSEMBLE ═══════ */}
      {view === "vue" && (
        <>
          {/* Sélecteur de période + granularité */}
          <Card>
            <div className="space-y-3">
              {/* Granularité semaine/mois */}
              <div className="flex items-center gap-2 flex-wrap">
                <span className="text-[10px] uppercase tracking-widest text-slate-500 mr-1">Granularité :</span>
                <div className="inline-flex rounded-lg border border-slate-700 overflow-hidden">
                  <button onClick={() => { setGranularity("month"); setPeriod("current"); }} className={`px-4 py-1.5 text-xs font-semibold transition-colors ${granularity === "month" ? "bg-cyan-500/20 text-cyan-200" : "bg-slate-950 text-slate-400 hover:text-slate-100"}`}>📅 Par mois</button>
                  <button onClick={() => { setGranularity("week"); setPeriod("current"); }} className={`px-4 py-1.5 text-xs font-semibold transition-colors ${granularity === "week" ? "bg-cyan-500/20 text-cyan-200" : "bg-slate-950 text-slate-400 hover:text-slate-100"}`}>🗓 Par semaine</button>
                </div>
              </div>
              <div className="flex items-center justify-between flex-wrap gap-3">
                <div className="flex items-center gap-1 flex-wrap">
                  <span className="text-[10px] uppercase tracking-widest text-slate-500 mr-2">Période :</span>
                  {[
                    { id: "current", label: granularity === "week" ? "Cette semaine" : "Ce mois" },
                    { id: "3m", label: granularity === "week" ? "13 sem." : "3 mois" },
                    { id: "6m", label: granularity === "week" ? "26 sem." : "6 mois" },
                    { id: "12m", label: granularity === "week" ? "52 sem." : "12 mois" },
                    { id: "all", label: "Tout" },
                  ].map(p => (
                    <button key={p.id} onClick={() => setPeriod(p.id)}
                      className={`px-3 py-1.5 text-xs font-semibold rounded transition-colors ${period === p.id ? "bg-cyan-500/20 text-cyan-200 border border-cyan-500/40" : "bg-slate-950 text-slate-400 hover:text-slate-100 border border-slate-700"}`}>
                      {p.label}
                    </button>
                  ))}
                </div>
                <div className="flex items-center gap-1">
                  <span className="text-[10px] uppercase tracking-widest text-slate-500 mr-2">Graphique :</span>
                  <button onClick={() => setChartMode("comparison")} className={`px-3 py-1.5 text-xs font-semibold rounded ${chartMode === "comparison" ? "bg-cyan-500/20 text-cyan-200 border border-cyan-500/40" : "bg-slate-950 text-slate-400 hover:text-slate-100 border border-slate-700"}`}>Barres</button>
                  <button onClick={() => setChartMode("cumulative")} className={`px-3 py-1.5 text-xs font-semibold rounded ${chartMode === "cumulative" ? "bg-cyan-500/20 text-cyan-200 border border-cyan-500/40" : "bg-slate-950 text-slate-400 hover:text-slate-100 border border-slate-700"}`}>Cumulé</button>
                </div>
              </div>
            </div>
          </Card>

          {/* Bandeau RUNWAY : combien de temps le surplus accumulé couvre les dépenses */}
          {runwayMonths !== undefined && lifetimeSurplus > 0 && avgMonthlySpendAllTime > 0 && (
            <div className="bg-gradient-to-br from-emerald-500/10 to-slate-900 border border-emerald-500/30 rounded-xl p-4">
              <div className="flex items-center gap-3 flex-wrap">
                <div className="text-3xl">🛟</div>
                <div className="flex-1 min-w-[200px]">
                  <div className="text-[10px] uppercase tracking-widest text-emerald-400 font-bold">Autonomie de ton surplus</div>
                  <div className="text-sm text-slate-300 mt-0.5">
                    Avec les <span className="text-emerald-300 font-semibold">{mask(fmtEUR(lifetimeSurplus, 0), hide)}</span> de surplus déjà accumulés, tes gains couvrent d'avance
                    <span className="text-emerald-300 font-bold"> {runwayMonths >= 12 ? `${(runwayMonths / 12).toFixed(1)} an${runwayMonths >= 24 ? "s" : ""}` : `${runwayMonths.toFixed(1)} mois`}</span> de dépenses carte
                    <span className="text-slate-500 font-mono"> (~{mask(fmtEUR(avgMonthlySpendAllTime, 0), hide)}/mois)</span>.
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* KPIs de la période */}
          <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-3">
            <KPICard label="Couverture" value={`${periodCoverage > 999 ? "∞" : periodCoverage.toFixed(0)}%`} icon={Target} accent={periodCoverage >= 100 ? "emerald" : periodCoverage >= 50 ? "amber" : "rose"} sub={`${monthsCovered}/${filtered.length} ${granularity === "week" ? "sem." : "mois"} OK`} />
            <KPICard label="Dépenses CB" value={mask(fmtEUR(totals.cardNet, 0), hide)} icon={ArrowDownRight} accent="rose" sub={`${monthsActive} ${granularity === "week" ? "sem." : "mois"} actif${monthsActive > 1 ? "s" : ""}`} />
            <KPICard label="Gains totaux" value={mask(fmtEUR(totals.gains, 0), hide)} icon={ArrowUpRight} accent="emerald" />
            <KPICard label="Solde net" value={mask(fmtEUR(totals.surplus, 0), hide)} subClass={classFor(totals.surplus)} icon={Flame} accent={totals.surplus >= 0 ? "emerald" : "rose"} big sub={totals.surplus >= 0 ? "✓ surplus" : "✗ déficit"} />
            <KPICard label={granularity === "week" ? "Dépense /sem." : "Dépense /mois"} value={mask(fmtEUR(avgMonthlySpend, 0), hide)} icon={CreditCard} accent="amber" />
            <KPICard label={granularity === "week" ? "Gain /sem." : "Gain /mois"} value={mask(fmtEUR(avgMonthlyGains, 0), hide)} icon={Coins} accent="cyan" />
          </div>

          {/* Graphique principal */}
          {filtered.length > 0 && (() => {
            const chartData = filtered.map(m => ({ ...m, label: labelOf(m) }));
            return (
            <Card>
              <SectionTitle icon={BarChart3} subtitle={chartMode === "comparison" ? `Gains vs dépenses par ${granularity === "week" ? "semaine" : "mois"}` : "Épargne nette accumulée dans le temps"}>
                {chartMode === "comparison" ? `📊 Bilan ${granularity === "week" ? "hebdomadaire" : "mensuel"}` : "📈 Surplus cumulé"}
              </SectionTitle>
              {chartMode === "comparison" ? (
                <ResponsiveContainer width="100%" height={360}>
                  <ComposedChart data={chartData}>
                    <defs>
                      <linearGradient id="gG" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#10b981" stopOpacity={0.9} /><stop offset="100%" stopColor="#10b981" stopOpacity={0.4} /></linearGradient>
                      <linearGradient id="gS" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#ef4444" stopOpacity={0.9} /><stop offset="100%" stopColor="#ef4444" stopOpacity={0.4} /></linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" vertical={false} />
                    <XAxis dataKey="label" stroke="#64748b" fontSize={10} />
                    <YAxis stroke="#64748b" fontSize={10} tickFormatter={(v) => fmtEUR(v, 0)} />
                    <Tooltip content={({ active, payload }) => {
                      if (!active || !payload?.length) return null;
                      const m = payload[0].payload;
                      return (
                        <div className="bg-slate-950 border border-slate-700 rounded-lg p-3 shadow-xl text-xs">
                          <div className="text-slate-400 uppercase tracking-widest text-[10px] mb-2">{fullLabelOf(m)}</div>
                          <div className="space-y-1 font-mono">
                            <div className="flex justify-between gap-4"><span className="text-emerald-400">↗ Gains :</span><span className="font-semibold">{fmtEUR(m.gains, 2)}</span></div>
                            <div className="pl-2 text-[10px] space-y-0.5">
                              <div className="flex justify-between gap-4"><span className="text-slate-500">P&L réalisée</span><span>{fmtEUR(m.realizedPnL, 2)}</span></div>
                              <div className="flex justify-between gap-4"><span className="text-slate-500">Dividendes</span><span>{fmtEUR(m.dividends, 2)}</span></div>
                              <div className="flex justify-between gap-4"><span className="text-slate-500">Intérêts</span><span>{fmtEUR(m.interests, 2)}</span></div>
                              <div className="flex justify-between gap-4"><span className="text-slate-500">Saveback</span><span>{fmtEUR(m.savebacks, 2)}</span></div>
                            </div>
                            <div className="flex justify-between gap-4 mt-1"><span className="text-rose-400">↘ Dépenses CB :</span><span className="font-semibold">{fmtEUR(m.cardNet, 2)}</span></div>
                            <div className={`flex justify-between gap-4 pt-1 border-t border-slate-700 mt-2 font-bold ${classFor(m.surplus)}`}><span>= Solde :</span><span>{fmtEUR(m.surplus, 2)}</span></div>
                            <div className="text-[10px] text-slate-500">{m.coverage >= 100 ? `✓ Couvert ×${(m.coverage / 100).toFixed(1)}` : `Couverture : ${m.coverage.toFixed(0)}%`}</div>
                          </div>
                        </div>
                      );
                    }} />
                    <Legend wrapperStyle={{ fontSize: 11, color: "#cbd5e1", paddingTop: 5 }} />
                    <ReferenceLine y={0} stroke="#475569" />
                    <Bar dataKey="gains" name="Gains (P&L + div + int + saveback)" fill="url(#gG)" radius={[4, 4, 0, 0]} />
                    <Bar dataKey="cardNet" name="Dépenses CB nettes" fill="url(#gS)" radius={[4, 4, 0, 0]} />
                    <Line type="monotone" dataKey="surplus" name="Solde net" stroke="#22d3ee" strokeWidth={2.5} dot={{ r: 3 }} />
                  </ComposedChart>
                </ResponsiveContainer>
              ) : (
                <ResponsiveContainer width="100%" height={360}>
                  <ComposedChart data={cumulativeData}>
                    <defs>
                      <linearGradient id="cSurplus" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#22d3ee" stopOpacity={0.4} /><stop offset="100%" stopColor="#22d3ee" stopOpacity={0} /></linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" vertical={false} />
                    <XAxis dataKey="label" stroke="#64748b" fontSize={10} />
                    <YAxis stroke="#64748b" fontSize={10} tickFormatter={(v) => fmtEUR(v, 0)} />
                    <Tooltip content={<ChartTooltip />} />
                    <ReferenceLine y={0} stroke="#475569" />
                    <Area type="monotone" dataKey="soldeCumul" name="Surplus cumulé (épargne nette)" stroke="#22d3ee" strokeWidth={3} fill="url(#cSurplus)" />
                    <Line type="monotone" dataKey="gainsCumul" name="Gains cumulés" stroke="#10b981" strokeWidth={2} dot={false} />
                    <Line type="monotone" dataKey="spendCumul" name="Dépenses cumulées" stroke="#ef4444" strokeWidth={2} dot={false} />
                    <Legend wrapperStyle={{ fontSize: 11, color: "#cbd5e1" }} />
                  </ComposedChart>
                </ResponsiveContainer>
              )}
            </Card>
            );
          })()}

          {/* Best/worst période */}
          {bestMonth && worstMonth && filtered.length > 1 && (
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="bg-emerald-500/5 border border-emerald-500/30 rounded-xl p-4">
                <div className="text-[10px] uppercase tracking-widest text-emerald-400 mb-2">⭐ {granularity === "week" ? "Meilleure semaine" : "Meilleur mois"}</div>
                <div className="text-xl font-bold">{fullLabelOf(bestMonth)}</div>
                <div className="text-3xl font-mono font-bold text-emerald-400 mt-1">{mask(fmtEUR(bestMonth.surplus, 0), hide)}</div>
                <div className="text-xs text-emerald-300/70 font-mono mt-1">{fmtEUR(bestMonth.gains, 0)} gains − {fmtEUR(bestMonth.cardNet, 0)} dépenses</div>
                <div className="text-[10px] text-slate-500 font-mono mt-1">{bestMonth.coverage >= 100 ? `✓ Couvert ×${(bestMonth.coverage / 100).toFixed(1)}` : `Couverture : ${bestMonth.coverage.toFixed(0)}%`}</div>
              </div>
              <div className="bg-rose-500/5 border border-rose-500/30 rounded-xl p-4">
                <div className="text-[10px] uppercase tracking-widest text-rose-400 mb-2">⚠ {granularity === "week" ? "Pire semaine" : "Pire mois"}</div>
                <div className="text-xl font-bold">{fullLabelOf(worstMonth)}</div>
                <div className="text-3xl font-mono font-bold text-rose-400 mt-1">{mask(fmtEUR(worstMonth.surplus, 0), hide)}</div>
                <div className="text-xs text-rose-300/70 font-mono mt-1">{fmtEUR(worstMonth.gains, 0)} gains − {fmtEUR(worstMonth.cardNet, 0)} dépenses</div>
                <div className="text-[10px] text-slate-500 font-mono mt-1">{worstMonth.coverage >= 100 ? `✓ Couvert ×${(worstMonth.coverage / 100).toFixed(1)}` : `Couverture : ${worstMonth.coverage.toFixed(0)}%`}</div>
              </div>
            </div>
          )}

          {/* Tableau mensuel détaillé */}
          {filtered.length > 0 && (
            <Card>
              <SectionTitle icon={CalendarDays}>📅 Détail {granularity === "week" ? "semaine par semaine" : "mois par mois"}</SectionTitle>
              <div className="overflow-x-auto max-h-[450px] overflow-y-auto">
                <table className="w-full text-xs">
                  <thead className="sticky top-0 bg-slate-900 z-10">
                    <tr className="border-b border-slate-800">
                      <th className="p-2 text-left text-[10px] text-slate-500 uppercase">Mois</th>
                      <th className="p-2 text-right text-[10px] text-slate-500 uppercase">Dépenses</th>
                      <th className="p-2 text-left text-[10px] text-slate-500 uppercase">{granularity === "week" ? "Semaine" : "Mois"}</th>
                      <th className="p-2 text-right text-[10px] text-slate-500 uppercase">Dépenses</th>
                      <th className="p-2 text-right text-[10px] text-slate-500 uppercase">P&L</th>
                      <th className="p-2 text-right text-[10px] text-slate-500 uppercase">Div+Int</th>
                      <th className="p-2 text-right text-[10px] text-slate-500 uppercase">Saveback</th>
                      <th className="p-2 text-right text-[10px] text-slate-500 uppercase">Total gains</th>
                      <th className="p-2 text-right text-[10px] text-slate-500 uppercase">Solde</th>
                      <th className="p-2 text-right text-[10px] text-slate-500 uppercase">Couv.</th>
                    </tr>
                  </thead>
                  <tbody>
                    {[...filtered].reverse().map((m, i) => (
                      <tr key={i} className={`border-b border-slate-800/60 ${ROW_HOVER}`}>
                        <td className="p-2 font-mono text-slate-200 font-semibold whitespace-nowrap">{granularity === "week" ? fullLabelOf(m) : m.month}</td>
                        <td className="p-2 text-right font-mono text-rose-400 tabular-nums">{mask(fmtEUR(m.cardNet, 0), hide)}</td>
                        <td className={`p-2 text-right font-mono tabular-nums ${classFor(m.realizedPnL)}`}>{m.realizedPnL !== 0 ? mask(fmtEUR(m.realizedPnL, 0), hide) : "—"}</td>
                        <td className="p-2 text-right font-mono text-amber-400 tabular-nums">{(m.dividends + m.interests) > 0 ? mask(fmtEUR(m.dividends + m.interests, 0), hide) : "—"}</td>
                        <td className="p-2 text-right font-mono text-violet-400 tabular-nums">{m.savebacks > 0 ? mask(fmtEUR(m.savebacks, 0), hide) : "—"}</td>
                        <td className="p-2 text-right font-mono text-emerald-400 font-semibold tabular-nums">{mask(fmtEUR(m.gains, 0), hide)}</td>
                        <td className={`p-2 text-right font-mono font-bold tabular-nums ${classFor(m.surplus)}`}>{mask(fmtEUR(m.surplus, 0), hide)}</td>
                        <td className="p-2 text-right">
                          <span className={`inline-block px-1.5 py-0.5 rounded text-[10px] font-bold ${m.coverage >= 100 ? "bg-emerald-500/20 text-emerald-300" : m.coverage >= 50 ? "bg-amber-500/20 text-amber-300" : m.coverage > 0 ? "bg-rose-500/20 text-rose-300" : "bg-slate-700 text-slate-400"}`}>
                            {m.coverage >= 100 ? `×${(m.coverage / 100).toFixed(1)}` : m.coverage.toFixed(0) + "%"}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot className="sticky bottom-0 bg-slate-900 border-t-2 border-slate-700">
                    <tr>
                      <td className="p-2 font-mono text-slate-200 font-bold uppercase tracking-widest text-[10px]">Total</td>
                      <td className="p-2 text-right font-mono text-rose-400 font-bold tabular-nums">{mask(fmtEUR(totals.cardNet, 0), hide)}</td>
                      <td className={`p-2 text-right font-mono font-bold tabular-nums ${classFor(totals.realizedPnL)}`}>{mask(fmtEUR(totals.realizedPnL, 0), hide)}</td>
                      <td className="p-2 text-right font-mono text-amber-400 font-bold tabular-nums">{mask(fmtEUR(totals.dividends + totals.interests, 0), hide)}</td>
                      <td className="p-2 text-right font-mono text-violet-400 font-bold tabular-nums">{mask(fmtEUR(totals.savebacks, 0), hide)}</td>
                      <td className="p-2 text-right font-mono text-emerald-400 font-bold tabular-nums">{mask(fmtEUR(totals.gains, 0), hide)}</td>
                      <td className={`p-2 text-right font-mono font-bold tabular-nums ${classFor(totals.surplus)}`}>{mask(fmtEUR(totals.surplus, 0), hide)}</td>
                      <td className="p-2 text-right"><span className={`inline-block px-1.5 py-0.5 rounded text-[10px] font-bold ${periodCoverage >= 100 ? "bg-emerald-500/20 text-emerald-300" : periodCoverage >= 50 ? "bg-amber-500/20 text-amber-300" : "bg-rose-500/20 text-rose-300"}`}>{periodCoverage >= 100 ? `×${(periodCoverage / 100).toFixed(1)}` : periodCoverage.toFixed(0) + "%"}</span></td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            </Card>
          )}
        </>
      )}

      {/* ═══════ VIEW : DÉPENSES (catégories cliquables → modal détail) ═══════ */}
      {view === "depenses" && (
        <>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
            <KPICard label="Total dépensé" value={mask(fmtEUR(cardSpendNet, 0), hide)} icon={CreditCard} accent="rose" big sub={`${cardTx.filter(t => t.amountRaw < 0).length} transactions`} />
            <KPICard label="Catégories" value={sortedCategories.length} icon={PieIcon} accent="cyan" sub={sortedCategories[0] ? `Top: ${sortedCategories[0].category}` : ""} />
            <KPICard label="Dépense moy./tx" value={cardTx.filter(t => t.amountRaw < 0).length > 0 ? mask(fmtEUR(cardSpendNet / cardTx.filter(t => t.amountRaw < 0).length, 0), hide) : "—"} icon={Hash} accent="amber" />
            <KPICard label="Plus grosse dépense" value={top10Expenses[0] ? mask(fmtEUR(Math.abs(top10Expenses[0].amountRaw), 0), hide) : "—"} icon={Flame} accent="rose" sub={top10Expenses[0]?.name ? top10Expenses[0].name.slice(0, 18) : ""} />
          </div>

          {/* Catégories cliquables */}
          <Card>
            <SectionTitle icon={PieIcon} subtitle="Clique sur une catégorie pour voir toutes les transactions">💳 Dépenses par catégorie</SectionTitle>
            {sortedCategories.length > 0 ? (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                {sortedCategories.map((c, i) => {
                  const pct = cardSpendNet > 0 ? (c.total / cardSpendNet) * 100 : 0;
                  const color = PALETTE[i % PALETTE.length];
                  return (
                    <button key={i} onClick={() => setOpenCategory(c)}
                      className="text-left bg-slate-900/60 border border-slate-800 rounded-lg p-3 hover:border-cyan-500/50 hover:bg-slate-800/60 transition-colors cursor-pointer">
                      <div className="flex items-center justify-between mb-2">
                        <div className="flex items-center gap-2">
                          <span className="w-2.5 h-2.5 rounded-sm" style={{ background: color }} />
                          <span className="text-sm font-semibold">{c.category}</span>
                        </div>
                        <span className="text-[10px] text-cyan-400/80">↗</span>
                      </div>
                      <div className="flex items-baseline justify-between mb-2">
                        <div>
                          <div className="text-2xl font-mono font-black tabular-nums">{mask(fmtEUR(c.total, 0), hide)}</div>
                          <div className="text-[10px] text-slate-500 font-mono">{c.count} tx · moy. {fmtEUR(c.total / c.count, 0)}</div>
                        </div>
                        <div className="text-right">
                          <div className="text-base font-mono font-bold text-slate-300">{pct.toFixed(1)}%</div>
                          <div className="text-[10px] text-slate-500">du total</div>
                        </div>
                      </div>
                      <div className="h-1.5 bg-slate-800 rounded-full overflow-hidden">
                        <div className="h-full rounded-full" style={{ width: `${pct}%`, background: color }} />
                      </div>
                    </button>
                  );
                })}
              </div>
            ) : <EmptyState title="Aucune dépense" icon={CreditCard} />}
          </Card>

          {/* Évolution des dépenses dans le temps (stacked) */}
          {monthlyByCategory.length > 1 && sortedCategories.length > 0 && (
            <Card>
              <SectionTitle icon={BarChart3} subtitle="Comment tes dépenses évoluent mois par mois">📈 Évolution des dépenses par catégorie</SectionTitle>
              <ResponsiveContainer width="100%" height={320}>
                <BarChart data={monthlyByCategory}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" vertical={false} />
                  <XAxis dataKey="month" stroke="#64748b" fontSize={10} tickFormatter={(v) => v.slice(2)} />
                  <YAxis stroke="#64748b" fontSize={10} tickFormatter={(v) => fmtEUR(v, 0)} />
                  <Tooltip content={({ active, payload, label }) => {
                    if (!active || !payload?.length) return null;
                    const tot = payload.reduce((s, p) => s + (p.value || 0), 0);
                    return (
                      <div className="bg-slate-950 border border-slate-700 rounded-lg p-2.5 shadow-xl text-xs max-w-xs">
                        <div className="text-slate-400 uppercase tracking-widest text-[10px] mb-1">{label}</div>
                        <div className="font-mono space-y-0.5">
                          {[...payload].sort((a, b) => (b.value || 0) - (a.value || 0)).map((p, i) => (
                            <div key={i} className="flex justify-between gap-2"><span className="text-slate-400">{p.name}</span><span>{fmtEUR(p.value || 0, 0)}</span></div>
                          ))}
                          <div className="flex justify-between gap-2 pt-1 border-t border-slate-700 mt-1 font-bold text-rose-400"><span>Total</span><span>{fmtEUR(tot, 2)}</span></div>
                        </div>
                      </div>
                    );
                  }} />
                  {sortedCategories.slice(0, 8).map((c, i) => (
                    <Bar key={c.category} dataKey={c.category} stackId="cat" fill={PALETTE[i % PALETTE.length]} />
                  ))}
                </BarChart>
              </ResponsiveContainer>
              <div className="flex flex-wrap gap-2 mt-2 text-[10px]">
                {sortedCategories.slice(0, 8).map((c, i) => (
                  <div key={c.category} className="flex items-center gap-1.5">
                    <span className="w-2 h-2 rounded-sm" style={{ background: PALETTE[i % PALETTE.length] }} />
                    <span className="text-slate-400">{c.category}</span>
                  </div>
                ))}
              </div>
            </Card>
          )}

          {/* Top 10 dépenses */}
          {top10Expenses.length > 0 && (
            <Card>
              <SectionTitle icon={Flame}>🔥 Top 10 plus grosses dépenses</SectionTitle>
              <div className="space-y-1.5">
                {top10Expenses.map((t, i) => (
                  <div key={i} className={`flex items-center gap-3 px-3 py-2 bg-slate-950/60 rounded-lg border border-slate-800/60 ${ROW_HOVER}`}>
                    <div className="w-7 h-7 rounded-full bg-rose-500/20 border border-rose-500/40 flex items-center justify-center text-xs font-bold text-rose-300 shrink-0">{i + 1}</div>
                    <div className="flex-1 min-w-0">
                      <div className="text-sm font-semibold truncate">{t.name || "Carte"}</div>
                      <div className="text-[10px] text-slate-500 font-mono">{fmtDateShort(t.date)} · {categorizeMerchant(t.name)}</div>
                    </div>
                    <div className="text-base font-mono font-bold text-rose-400 tabular-nums shrink-0">{mask(fmtEUR(-Math.abs(t.amountRaw), 2), hide)}</div>
                  </div>
                ))}
              </div>
            </Card>
          )}
        </>
      )}

      {/* ═══════ VIEW : GAINS ═══════ */}
      {view === "gains" && (
        <>
          <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
            <KPICard label="Total gains" value={mask(fmtEUR(totalGains, 0), hide)} icon={Coins} accent="emerald" big />
            <KPICard label="P&L réalisée" value={mask(fmtEUR(d.stats.totalPnL, 0), hide)} subClass={classFor(d.stats.totalPnL)} icon={Activity} accent={d.stats.totalPnL >= 0 ? "emerald" : "rose"} sub={`${d.stats.count} ventes`} />
            <KPICard label="Dividendes" value={mask(fmtEUR(d.stats.totalDividends, 0), hide)} icon={Percent} accent="amber" sub={`${d.dividends.length} versements`} />
            <KPICard label="Intérêts" value={mask(fmtEUR(d.stats.totalInterests, 0), hide)} icon={DollarSign} accent="cyan" sub={`${d.interests.length} versements`} />
            <KPICard label="Bonus" value={mask(fmtEUR(d.stats.totalSavebacks + d.stats.totalGifts, 0), hide)} icon={Gift} accent="violet" sub="Saveback + Gift" />
          </div>

          {/* Sources de gains visualisées */}
          {gainBreakdownArr.length > 0 && (
            <Card>
              <SectionTitle icon={Coins} subtitle="Détail de ce qui compose tes gains sur la période sélectionnée">💰 Sources de gains ({period === "current" ? "ce mois" : period === "all" ? "all-time" : period})</SectionTitle>
              <div className="space-y-2.5">
                {gainBreakdownArr.sort((a, b) => Math.abs(b.value) - Math.abs(a.value)).map((s, i) => {
                  const colors = { "P&L réalisée": "#10b981", "Dividendes": "#f59e0b", "Intérêts": "#06b6d4", "Saveback": "#a855f7", "Gifts": "#ec4899" };
                  const totAbs = gainBreakdownArr.reduce((sum, x) => sum + Math.abs(x.value), 0);
                  const pct = totAbs > 0 ? (Math.abs(s.value) / totAbs) * 100 : 0;
                  return (
                    <div key={i}>
                      <div className="flex items-center justify-between mb-1">
                        <div className="flex items-center gap-2">
                          <span className="w-2 h-2 rounded-full" style={{ background: colors[s.name] || "#64748b" }} />
                          <span className="text-sm text-slate-200 font-semibold">{s.name}</span>
                        </div>
                        <div className="flex items-center gap-2">
                          <span className="text-[10px] text-slate-500 font-mono">{pct.toFixed(1)}%</span>
                          <span className={`text-sm font-mono font-bold tabular-nums ${classFor(s.value)}`}>{mask(fmtEUR(s.value, 0), hide)}</span>
                        </div>
                      </div>
                      <div className="h-2 bg-slate-800 rounded-full overflow-hidden">
                        <div className="h-full rounded-full" style={{ width: `${pct}%`, background: colors[s.name] || "#64748b", opacity: s.value < 0 ? 0.4 : 1 }} />
                      </div>
                    </div>
                  );
                })}
              </div>
            </Card>
          )}

          {/* Dividendes vs Intérêts récents */}
          {d.dividends.length > 0 && (
            <Card>
              <SectionTitle icon={Percent}>💎 Derniers dividendes encaissés</SectionTitle>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-2 max-h-[300px] overflow-y-auto">
                {[...d.dividends].sort((a, b) => b.date - a.date).slice(0, 20).map((dv, i) => (
                  <div key={i} className={`flex items-center justify-between p-2 bg-slate-950/60 rounded border border-amber-500/20 ${ROW_HOVER}`}>
                    <div className="min-w-0 flex-1">
                      <div className="text-xs font-semibold truncate">💰 {dv.name}</div>
                      <div className="text-[10px] text-slate-500 font-mono">{fmtDateShort(dv.date)}</div>
                    </div>
                    <div className="text-right">
                      <div className="text-sm font-mono font-bold text-amber-300 tabular-nums">{mask(fmtEUR(dv.net, 2), hide)}</div>
                      <div className="text-[9px] text-slate-600 font-mono">brut {fmtEUR(dv.gross, 2)}</div>
                    </div>
                  </div>
                ))}
              </div>
            </Card>
          )}
        </>
      )}

      {/* ═══════ VIEW : PROJECTION ═══════ */}
      {view === "projection" && (
        <>
          {/* Hero projection */}
          <div className={`relative overflow-hidden rounded-2xl border-2 ${projection.gap <= 0 ? "border-emerald-500/40 bg-gradient-to-br from-emerald-500/10 via-slate-900 to-slate-950" : "border-amber-500/40 bg-gradient-to-br from-amber-500/10 via-slate-900 to-slate-950"}`}>
            <div className="absolute -top-32 -right-32 w-64 h-64 rounded-full opacity-20 blur-3xl" style={{ background: projection.gap <= 0 ? "#10b981" : "#f59e0b" }} />
            <div className="relative p-6 md:p-8">
              <div className="text-[10px] uppercase tracking-[0.2em] text-slate-400 mb-2">🎯 Objectif : couvrir mes dépenses CB chaque mois</div>
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 items-end">
                <div>
                  <div className="text-[10px] uppercase tracking-widest text-slate-500 mb-1">Tu génères en moyenne</div>
                  <div className="text-3xl md:text-4xl font-mono font-black text-cyan-300 tabular-nums">{mask(fmtEUR(projection.avgGain, 0), hide)}</div>
                  <div className="text-xs text-slate-500 font-mono">par mois (12 derniers)</div>
                </div>
                <div className="text-center hidden lg:block">
                  <span className="text-slate-500 text-3xl">→</span>
                </div>
                <div>
                  <div className="text-[10px] uppercase tracking-widest text-slate-500 mb-1">Pour couvrir tes dépenses</div>
                  <div className="text-3xl md:text-4xl font-mono font-black text-emerald-300 tabular-nums">{mask(fmtEUR(projection.avgSpend, 0), hide)}</div>
                  <div className="text-xs text-slate-500 font-mono">par mois (12 derniers)</div>
                </div>
              </div>
              <div className="mt-5">
                <div className="flex items-center justify-between text-[10px] text-slate-500 font-mono mb-1">
                  <span>0€</span><span>Objectif : {mask(fmtEUR(projection.avgSpend, 0), hide)}</span>
                </div>
                <div className="h-3 bg-slate-800 rounded-full overflow-hidden">
                  <div className={`h-full transition-all ${projection.gap <= 0 ? "bg-emerald-500" : "bg-amber-500"}`} style={{ width: `${projection.avgSpend > 0 ? Math.min(100, (projection.avgGain / projection.avgSpend) * 100) : 0}%` }} />
                </div>
              </div>
              <div className="mt-4 text-sm">
                {projection.gap <= 0 ? (
                  <div className="text-emerald-300">🎉 <strong>Objectif atteint en moyenne !</strong> Tu génères {mask(fmtEUR(-projection.gap, 0), hide)} de surplus mensuel en moyenne.</div>
                ) : (
                  <div className="text-amber-200">
                    Il te manque <strong className="text-amber-300">{mask(fmtEUR(projection.gap, 0), hide)}/mois</strong> en moyenne pour couvrir entièrement tes dépenses CB.
                  </div>
                )}
              </div>
            </div>
          </div>

          {/* Scénarios de projection */}
          <Card>
            <SectionTitle icon={Target} subtitle="À ce rythme, voilà combien tu auras généré">📊 Scénarios à court/moyen terme</SectionTitle>
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-3">
              {[
                { months: 3, label: "Dans 3 mois" },
                { months: 6, label: "Dans 6 mois" },
                { months: 12, label: "Dans 1 an" },
                { months: 24, label: "Dans 2 ans" },
              ].map(s => {
                const gainsProj = projection.avgGain * s.months;
                const spendProj = projection.avgSpend * s.months;
                const surplus = gainsProj - spendProj;
                return (
                  <div key={s.months} className="bg-slate-900/60 border border-slate-800 rounded-xl p-4">
                    <div className="text-[10px] uppercase tracking-widest text-slate-500 mb-2">{s.label}</div>
                    <div className="space-y-1.5">
                      <div className="flex justify-between text-xs"><span className="text-emerald-400">Gains projetés</span><span className="font-mono font-semibold tabular-nums">{mask(fmtEUR(gainsProj, 0), hide)}</span></div>
                      <div className="flex justify-between text-xs"><span className="text-rose-400">Dépenses projetées</span><span className="font-mono font-semibold tabular-nums">{mask(fmtEUR(spendProj, 0), hide)}</span></div>
                      <div className={`flex justify-between text-sm pt-1 border-t border-slate-800 font-bold ${classFor(surplus)}`}>
                        <span>Solde</span><span className="font-mono tabular-nums">{mask((surplus >= 0 ? "+" : "") + fmtEUR(surplus, 0).replace("+", ""), hide)}</span>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
            <div className="mt-3 text-[10px] text-slate-500 font-mono">ⓘ Projection linéaire basée sur les 12 derniers mois. Ne prend pas en compte les variations de marché ni les changements de comportement.</div>
          </Card>

          {/* Combien il faut générer de plus ? */}
          {projection.gap > 0 && (
            <Card>
              <SectionTitle icon={Flame}>🚀 Comment combler le manque ?</SectionTitle>
              <div className="space-y-3">
                <div className="text-sm text-slate-300">
                  Pour atteindre <strong className="text-emerald-300">100% de couverture mensuelle</strong>, il te faudrait générer <strong className="text-amber-300">{mask(fmtEUR(projection.gap, 0), hide)}/mois</strong> en plus. Voici quelques pistes :
                </div>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-3 mt-3">
                  <div className="bg-emerald-500/5 border border-emerald-500/30 rounded-lg p-3">
                    <div className="text-[10px] uppercase tracking-widest text-emerald-400 mb-1">📈 Plus de plus-values</div>
                    <div className="text-sm text-slate-300">
                      Vendre <strong>{mask(fmtEUR(projection.gap, 0), hide)}</strong> de plus-values par mois (avant impôts CTO).
                    </div>
                  </div>
                  <div className="bg-amber-500/5 border border-amber-500/30 rounded-lg p-3">
                    <div className="text-[10px] uppercase tracking-widest text-amber-400 mb-1">💰 Plus de dividendes</div>
                    <div className="text-sm text-slate-300">
                      Investir <strong>{mask(fmtEUR((projection.gap * 12) / 0.04, 0), hide)}</strong> de plus dans des actifs à rendement 4%/an.
                    </div>
                  </div>
                  <div className="bg-rose-500/5 border border-rose-500/30 rounded-lg p-3">
                    <div className="text-[10px] uppercase tracking-widest text-rose-400 mb-1">✂ Réduire dépenses</div>
                    <div className="text-sm text-slate-300">
                      Couper <strong>{mask(fmtEUR(projection.gap, 0), hide)}/mois</strong> dans tes dépenses CB.
                    </div>
                  </div>
                </div>
              </div>
            </Card>
          )}
        </>
      )}

      {/* Modal détail catégorie */}
      {openCategory && <CategoryDetailModal category={openCategory} onClose={() => setOpenCategory(null)} hide={hide} />}
    </div>
  );
}

// ═══ Sous-composants Finances ═══

// ═══════ Comparaison mois courant vs mois précédent ═══════
const MonthComparisonBar = ({ comparison, hide }) => {
  if (!comparison || !comparison.current || !comparison.previous) return null;
  const { current, previous, deltaGains, deltaSpend, deltaCoverage } = comparison;
  const DeltaPill = ({ delta, invert, suffix = "€" }) => {
    if (delta === null || delta === undefined) return null;
    // invert : pour les dépenses, une BAISSE est positive
    const isGood = invert ? delta < 0 : delta > 0;
    const isNeutral = Math.abs(delta) < 0.5;
    const cls = isNeutral ? "bg-slate-700/50 text-slate-400" : isGood ? "bg-emerald-500/15 text-emerald-300" : "bg-rose-500/15 text-rose-300";
    const arrow = isNeutral ? "→" : delta > 0 ? "↑" : "↓";
    return (
      <span className={`inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-[10px] font-bold font-mono ${cls}`}>
        {arrow} {suffix === "€" ? mask(fmtEUR(Math.abs(delta), 0), hide) : Math.abs(delta).toFixed(0) + suffix}
      </span>
    );
  };
  return (
    <Card>
      <div className="flex items-center justify-between flex-wrap gap-3">
        <div className="flex items-center gap-2">
          <RefreshCw className="w-3.5 h-3.5 text-cyan-400" />
          <span className="text-xs font-semibold uppercase tracking-widest text-slate-300">Tendance vs mois dernier</span>
          <span className="text-[10px] text-slate-500 font-mono">({previous.month} → {current.month})</span>
        </div>
        <div className="flex items-center gap-4 flex-wrap text-xs">
          <div className="flex items-center gap-1.5">
            <span className="text-slate-400">Gains :</span>
            <DeltaPill delta={deltaGains} />
          </div>
          <div className="flex items-center gap-1.5">
            <span className="text-slate-400">Dépenses :</span>
            <DeltaPill delta={deltaSpend} invert />
          </div>
          <div className="flex items-center gap-1.5">
            <span className="text-slate-400">Couverture :</span>
            <DeltaPill delta={deltaCoverage} suffix=" pts" />
          </div>
        </div>
      </div>
    </Card>
  );
};

const HorizonCard = ({ label, sublabel, data, accent, hide, highlight }) => {
  const accentColor = { cyan: "text-cyan-400", violet: "text-violet-400", slate: "text-slate-400" }[accent] || "text-slate-400";
  if (!data) return (
    <div className={`relative overflow-hidden rounded-xl border ${highlight ? "border-2 border-cyan-500/40" : "border-slate-800"} bg-slate-900/60 p-4`}>
      <div className={`text-[10px] uppercase tracking-[0.2em] ${accentColor} mb-1 font-bold`}>{label}</div>
      <div className="text-xl text-slate-500">—</div>
    </div>
  );
  const covered = data.coverage >= 100;
  const covColor = covered ? "text-emerald-400" : data.coverage >= 50 ? "text-amber-400" : data.coverage > 0 ? "text-rose-400" : "text-slate-500";
  const barColor = covered ? "bg-emerald-500" : data.coverage >= 50 ? "bg-amber-500" : "bg-rose-500";
  // % affiché : plafonné à un format lisible. Au-delà de 100% on parle de "surplus", pas de gros %.
  const pctLabel = data.spend <= 0 ? (data.gains > 0 ? "—" : "—") : data.coverage > 999 ? ">999%" : Math.round(data.coverage) + "%";
  return (
    <div className={`relative overflow-hidden rounded-xl border ${highlight ? "border-2 border-cyan-500/40 bg-gradient-to-br from-cyan-500/10 to-slate-900" : "border-slate-800 bg-slate-900/60"} p-4`}>
      <div className="flex items-center justify-between mb-1">
        <div>
          <div className={`text-[10px] uppercase tracking-[0.2em] ${accentColor} font-bold`}>{label}</div>
          <div className="text-[9px] text-slate-500 font-mono">{sublabel}</div>
        </div>
        {/* Badge couverture : statut clair plutôt qu'un % qui explose */}
        <div className={`px-2 py-0.5 rounded-full text-[10px] font-bold ${covered ? "bg-emerald-500/15 text-emerald-300 border border-emerald-500/30" : data.coverage >= 50 ? "bg-amber-500/15 text-amber-300 border border-amber-500/30" : data.spend > 0 ? "bg-rose-500/15 text-rose-300 border border-rose-500/30" : "bg-slate-700/50 text-slate-400"}`}>
          {covered ? "✓ Couvert" : data.spend > 0 ? pctLabel : "—"}
        </div>
      </div>
      {/* Chiffre principal = le SURPLUS (ou manque) en € */}
      <div className="mt-2">
        <div className="text-[9px] uppercase tracking-widest text-slate-500">{data.surplus >= 0 ? "💰 Surplus net" : "⚠ Reste à couvrir"}</div>
        <div className={`text-3xl md:text-4xl font-mono font-black tabular-nums leading-none ${classFor(data.surplus)}`}>
          {mask((data.surplus >= 0 ? "+" : "−") + fmtEUR(Math.abs(data.surplus), 0).replace("+", "").replace("-", ""), hide)}
        </div>
      </div>
      <div className="grid grid-cols-2 gap-2 mt-3">
        <div>
          <div className="text-[9px] uppercase text-slate-500 tracking-widest">↗ Gains</div>
          <div className="text-xs font-mono font-bold text-emerald-400 tabular-nums">{mask(fmtEUR(data.gains, 0), hide)}</div>
        </div>
        <div>
          <div className="text-[9px] uppercase text-slate-500 tracking-widest">↘ Dépenses CB</div>
          <div className="text-xs font-mono font-bold text-rose-400 tabular-nums">{mask(fmtEUR(data.spend, 0), hide)}</div>
        </div>
      </div>
      {/* Barre : remplie jusqu'à 100%, avec halo si dépassé */}
      <div className="mt-2.5 h-2 bg-slate-800 rounded-full overflow-hidden relative">
        <div className={`h-full transition-all ${barColor}`} style={{ width: `${Math.min(100, data.coverage)}%` }} />
        {covered && data.coverage > 100 && (
          <div className="absolute inset-0 flex items-center justify-end pr-1.5">
            <span className="text-[8px] font-bold text-emerald-100">×{(data.coverage / 100).toFixed(1)}</span>
          </div>
        )}
      </div>
      <div className="text-[9px] text-slate-500 font-mono mt-1">
        {covered ? `Tu couvres ${(data.coverage / 100).toFixed(1)}× tes dépenses` : data.spend > 0 ? `${pctLabel} des dépenses couvertes` : "Pas de dépenses sur la période"}
      </div>
    </div>
  );
};

const CategoryDetailModal = ({ category, onClose, hide }) => {
  useEffect(() => {
    const orig = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => { document.body.style.overflow = orig; };
  }, []);
  const total = category.total;
  const avg = category.count > 0 ? total / category.count : 0;
  const max = category.transactions[0] ? Math.abs(category.transactions[0].amountRaw) : 0;
  return (
    <div className="fixed inset-0 z-50 flex items-end md:items-center justify-center bg-black/70 backdrop-blur-sm p-0 md:p-4" onClick={onClose}>
      <div className="bg-slate-900 border border-slate-700 rounded-t-2xl md:rounded-2xl w-full md:max-w-3xl max-h-[92vh] overflow-hidden flex flex-col" onClick={e => e.stopPropagation()}>
        <div className="p-5 border-b border-slate-800 bg-gradient-to-br from-rose-500/10 to-transparent">
          <div className="flex items-start justify-between gap-3">
            <div>
              <div className="text-[10px] uppercase tracking-widest text-slate-400 mb-1">Catégorie</div>
              <div className="text-2xl md:text-3xl font-bold">{category.category}</div>
              <div className="text-xs text-slate-400 font-mono mt-1">{category.count} transactions</div>
            </div>
            <button onClick={onClose} className="p-2 bg-slate-800 hover:bg-slate-700 rounded text-slate-300 text-xl leading-none">✕</button>
          </div>
          <div className="mt-4 grid grid-cols-3 gap-3">
            <div className="bg-slate-950/60 rounded-lg p-2.5 border border-slate-800">
              <div className="text-[9px] uppercase tracking-widest text-slate-500">Total</div>
              <div className="text-base font-mono font-bold text-rose-400 tabular-nums">{mask(fmtEUR(total, 2), hide)}</div>
            </div>
            <div className="bg-slate-950/60 rounded-lg p-2.5 border border-slate-800">
              <div className="text-[9px] uppercase tracking-widest text-slate-500">Moy./tx</div>
              <div className="text-base font-mono font-bold tabular-nums">{mask(fmtEUR(avg, 2), hide)}</div>
            </div>
            <div className="bg-slate-950/60 rounded-lg p-2.5 border border-slate-800">
              <div className="text-[9px] uppercase tracking-widest text-slate-500">Max</div>
              <div className="text-base font-mono font-bold tabular-nums">{mask(fmtEUR(max, 2), hide)}</div>
            </div>
          </div>
        </div>
        <div className="overflow-y-auto flex-1 p-3">
          <div className="space-y-1">
            {category.transactions.map((t, i) => (
              <div key={i} className={`flex items-center gap-3 px-3 py-2 bg-slate-950/60 rounded-lg border border-slate-800/60 ${ROW_HOVER}`}>
                <div className="flex-1 min-w-0">
                  <div className="text-sm font-semibold truncate">{t.name || "Carte"}</div>
                  <div className="text-[10px] text-slate-500 font-mono">{fmtDateShort(t.date)}</div>
                </div>
                <div className="text-sm font-mono font-bold text-rose-400 tabular-nums">{mask(fmtEUR(t.amountRaw, 2), hide)}</div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};
