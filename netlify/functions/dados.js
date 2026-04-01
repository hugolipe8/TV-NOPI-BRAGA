/**
 * Netlify Serverless Function — TV Nopi Braga
 * GET /.netlify/functions/dados
 */

const fetch = require("node-fetch");
const XLSX  = require("xlsx");

const EXCEL_URL = [
  "https://www.dropbox.com/scl/fi/y4i9m6v4q8snd2m3qljoh/Motherboard-2026.xlsx",
  "?rlkey=4px2hpxbg8p6fot2l65bkdamg&st=4h2vu72e&dl=1",
].join("");

const MONTH_OFFSETS = [0, 17, 30, 42, 54, 66, 78, 90, 102, 114, 126, 138];
const MONTH_NAMES   = [
  "Janeiro","Fevereiro","Março","Abril","Maio","Junho",
  "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro",
];

// Estrutura BRG mensal (quando há dados): O,R,O,R,O,R,O,R,O,R
// off+1=FAT-O  off+2=FAT-R
// off+3=ANG-O  off+4=ANG-R  ← angariações reais
// off+5=PROP-O off+6=PROP-R
// off+7=CONT-O off+8=CONT-R ← contratos reais
const ANG_COL  = 4;
const CONT_COL = 8;

const SKIP = new Set(["brg","cg","ag","fp","cm"]);
const STOP = new Set(["total geral","cessados"]);

const CORS = {
  "Access-Control-Allow-Origin":  "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

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

function fmtDate(v) {
  if (v instanceof Date && !isNaN(v)) {
    const p = (n) => String(n).padStart(2, "0");
    return `${p(v.getDate())}/${p(v.getMonth() + 1)}/${v.getFullYear()}`;
  }
  if (typeof v === "number" && v > 0) {
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    const p = (n) => String(n).padStart(2, "0");
    return `${p(d.getUTCDate())}/${p(d.getUTCMonth() + 1)}/${d.getUTCFullYear()}`;
  }
  return String(v || "");
}

function json(statusCode, body, extra = {}) {
  return {
    statusCode,
    headers: { ...CORS, "Content-Type": "application/json; charset=utf-8", ...extra },
    body: JSON.stringify(body),
  };
}

exports.handler = async (event) => {
  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 204, headers: CORS, body: "" };
  }

  try {
    const res = await fetch(EXCEL_URL, { timeout: 45_000 });
    if (!res.ok) throw new Error(`Dropbox respondeu HTTP ${res.status}`);
    const buf = await res.buffer();

    const wb = XLSX.read(buf, { type: "buffer", cellDates: true });

    // ── Folha RC ──────────────────────────────────────────────────────────────
    const ws = wb.Sheets["RC"];
    if (!ws) throw new Error('Folha "RC" não encontrada no ficheiro Excel.');
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

    const now = new Date();
    const mi  = now.getMonth();
    const off = MONTH_OFFSETS[mi];

    // Encontrar linha BRG dinamicamente (a partir da row 50 para ignorar
    // a tabela anual que também tem BRG mas aparece antes, na row 11)
    let brgRowIdx = -1;
    for (let i = 50; i < rows.length; i++) {
      if (String(rows[i][off] ?? "").trim().toUpperCase() === "BRG") {
        brgRowIdx = i;
        break;
      }
    }

    // Se não há dados para este mês, devolve zeros
    if (brgRowIdx === -1) {
      return json(200,
        { mes: MONTH_NAMES[mi], ano: now.getFullYear(), totalAng: 0, totalCont: 0, consultores: [], ultimasAngariações: [] },
        { "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0" }
      );
    }

    // Totais BRG — usa colunas R (reais): ANG-R=off+4, CONT-R=off+8
    const brgRow    = rows[brgRowIdx];
    const totalAng  = toInt(brgRow[off + ANG_COL]);
    const totalCont = toInt(brgRow[off + CONT_COL]);

    // Consultores — lê dinamicamente até encontrar linha de paragem
    const consultores = [];
    for (let i = brgRowIdx + 1; i < rows.length; i++) {
      const row   = rows[i];
      const name  = String(row[off] ?? "").trim();
      if (!name) continue;
      const lower = name.toLowerCase();
      if (STOP.has(lower)) break;
      if (SKIP.has(lower)) continue;
      const ang  = toInt(row[off + ANG_COL]);
      const cont = toInt(row[off + CONT_COL]);
      if (ang > 0 || cont > 0) consultores.push({ nome: name, ang, cont });
    }

    // ── Folha MOTHER — últimas angariações BRG/ANG/VO ────────────────────────
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
        tipo:       String(row[65] ?? "").trim(),
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
