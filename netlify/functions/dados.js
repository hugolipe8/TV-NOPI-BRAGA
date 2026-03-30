/**
 * Netlify Serverless Function — TV Nopi Braga
 * GET /.netlify/functions/dados
 *
 * Descarrega o Excel da Dropbox, parseia a folha RC e devolve
 * os dados do mês corrente como JSON.
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

// Linhas a saltar (subtotais, cabeçalhos de equipa)
const SKIP = new Set(["brg","cg","ag","fp","cm"]);
// Linhas que marcam o fim dos consultores activos
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
  // Pre-flight CORS
  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 204, headers: CORS, body: "" };
  }

  try {
    // 1. Descarregar Excel
    const res = await fetch(EXCEL_URL, { timeout: 45_000 });
    if (!res.ok) throw new Error(`Dropbox respondeu HTTP ${res.status}`);
    const buf = await res.buffer();

    // 2. Parsear com SheetJS
    const wb = XLSX.read(buf, { type: "buffer" });
    const ws = wb.Sheets["RC"];
    if (!ws) throw new Error('Folha "RC" não encontrada no ficheiro Excel.');

    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

    // 3. Determinar mês corrente
    const now = new Date();
    const mi  = now.getMonth();          // 0 = Janeiro
    const off = MONTH_OFFSETS[mi];

    // 4. Totais BRG (linha índice 66)
    const brgRow   = rows[66] || [];
    const totalAng  = toInt(brgRow[off + 2]);
    const totalCont = toInt(brgRow[off + 4]);

    // 5. Consultores (a partir da linha índice 67)
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

    const pad = (n) => String(n).padStart(2, "0");

    return json(
      200,
      {
        mes:        MONTH_NAMES[mi],
        ano:        now.getFullYear(),
        totalAng,
        totalCont,
        consultores,
      },
      { "Cache-Control": "no-store, no-cache, must-revalidate" }
    );

  } catch (err) {
    console.error("[dados]", err.message);
    return json(500, { erro: err.message });
  }
};
