/**
 * Netlify Serverless Function — TV Nopi Braga
 * GET /.netlify/functions/dados
 */

const fetch = require("node-fetch");
const XLSX  = require("xlsx");

// ─────────────────────────────────────────────────────────────────────────────
// CONFIG
// ─────────────────────────────────────────────────────────────────────────────
const EXCEL_URL = [
  "https://www.dropbox.com/scl/fi/y4i9m6v4q8snd2m3qljoh/Motherboard-2026.xlsx",
  "?rlkey=4px2hpxbg8p6fot2l65bkdamg&st=4h2vu72e&dl=1",
].join("");

const MONTH_OFFSETS = [0, 17, 30, 42, 54, 66, 78, 90, 102, 114, 126, 138];
const MONTH_NAMES   = [
  "Janeiro","Fevereiro","Março","Abril","Maio","Junho",
  "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro",
];

const SKIP = new Set(["brg","cg","ag","fp","cm"]);
const STOP = new Set(["total geral","cessados"]);

const CORS = {
  "Access-Control-Allow-Origin":  "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

// ─────────────────────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────────────────────
function toInt(v) {
  if (v == null || v === "") return 0;
  const n = parseFloat(String(v));
  return Number.isFinite(n) ? Math.round(n) : 0;
}

function toNum(v) {
  if (v == null || v === "") return 0;
  const n = parseFloat(String(v).replace(",", "."));
  return Number.isFinite(n) ? n : 0;
}

/** Converte valor de célula Excel para string de data dd/mm/aaaa */
function fmtDate(v) {
  if (v instanceof Date && !isNaN(v)) {
    const p = (n) => String(n).padStart(2, "0");
    return `${p(v.getDate())}/${p(v.getMonth() + 1)}/${v.getFullYear()}`;
  }
  if (typeof v === "number" && v > 0) {
    // serial Excel → JS Date (UTC)
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    const p = (n) => String(n).padStart(2, "0");
    return `${p(d.getUTCDate())}/${p(d.getUTCMonth() + 1)}/${d.getUTCFullYear()}`;
  }
  return String(v || "");
}

/** Ordena datas (Date | number | string) de forma descendente */
function dateTs(v) {
  if (v instanceof Date) return v.getTime();
  if (typeof v === "number") return v; // serial Excel — ordem relativa correcta
  return 0;
}

function json(statusCode, body, extra = {}) {
  return {
    statusCode,
    headers: { ...CORS, "Content-Type": "application/json; charset=utf-8", ...extra },
    body: JSON.stringify(body),
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// HANDLER
// ─────────────────────────────────────────────────────────────────────────────
exports.handler = async (event) => {
  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 204, headers: CORS, body: "" };
  }

  try {
    // 1. Descarregar Excel
    const res = await fetch(EXCEL_URL, { timeout: 45_000 });
    if (!res.ok) throw new Error(`Dropbox respondeu HTTP ${res.status}`);
    const buf = await res.buffer();

    // 2. Parsear — cellDates:true converte datas para objetos JS Date
    const wb = XLSX.read(buf, { type: "buffer", cellDates: true });

    // ── Folha RC ──────────────────────────────────────────────────────────────
    const ws = wb.Sheets["RC"];
    if (!ws) throw new Error('Folha "RC" não encontrada no ficheiro Excel.');
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

    const now = new Date();
    const mi  = now.getMonth();
    const off = MONTH_OFFSETS[mi];

    // Totais BRG (linha índice 66)
    const brgRow    = rows[66] || [];
    const totalAng  = toInt(brgRow[off + 2]);
    const totalCont = toInt(brgRow[off + 4]);

    // Consultores (a partir da linha índice 67)
    const consultores = [];
    for (let i = 67; i < rows.length; i++) {
      const row   = rows[i];
      const name  = String(row[off] ?? "").trim();
      if (!name) continue;
      const lower = name.toLowerCase();
      if (STOP.has(lower)) break;
      if (SKIP.has(lower)) continue;
      const ang  = toInt(row[off + 2]);
      const cont = toInt(row[off + 4]);
      if (ang > 0 || cont > 0) consultores.push({ nome: name, ang, cont });
    }

    // ── Folha MOTHER — últimas angariações BRG/ANG/VO ───────────────────────────
    const wsMother = wb.Sheets["MOTHER"];
    const motherRows = wsMother
      ? XLSX.utils.sheet_to_json(wsMother, { header: 1, defval: "" })
      : [];
    const ultimasAngariações = motherRows
      .filter(row =>
        String(row[55] ?? "").trim() === "BRG" &&
        String(row[57] ?? "").trim() === "ANG" &&
        String(row[60] ?? "").trim() === "VO"
      )
      .slice(-5)
      .reverse()
      .map(row => ({
        ref:        String(row[61] ?? "").trim(),
        localidade: String(row[58] ?? "").trim(),
        consultor:  String(row[62] ?? "").trim(),
        valor:      toNum(row[67]),
        data:       fmtDate(row[59]),
      }));

    return json(
      200,
      { mes: MONTH_NAMES[mi], ano: now.getFullYear(), totalAng, totalCont, consultores, ultimasAngariações },
      { "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0" }
    );

  } catch (err) {
    console.error("[dados]", err.message);
    return json(500, { erro: err.message });
  }
};
