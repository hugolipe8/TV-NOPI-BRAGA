/**
 * Netlify Serverless Function — TV Nopi Braga
 * GET /.netlify/functions/gestao
 *
 * Lê a folha RE do Excel Motherboard-2026 e devolve KPIs
 * de faturação, angariação e contratação para global + 3 lojas.
 */

const fetch = require("node-fetch");
const XLSX  = require("xlsx");

const EXCEL_URL =
  "https://www.dropbox.com/scl/fi/y4i9m6v4q8snd2m3qljoh/Motherboard-2026.xlsx" +
  "?rlkey=4px2hpxbg8p6fot2l65bkdamg&st=4h2vu72e&dl=1";

const HEADERS = {
  "Access-Control-Allow-Origin":  "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
  "Cache-Control": "no-store",
  "Content-Type":  "application/json",
};

/** Lê o valor numérico de uma célula pelo índice de linha (1-based) e coluna (0-based A=0). */
function val(sheet, row, col) {
  const addr = XLSX.utils.encode_cell({ r: row - 1, c: col });
  const cell = sheet[addr];
  if (!cell || cell.v === undefined || cell.v === null) return 0;
  const n = Number(cell.v);
  return isNaN(n) ? 0 : n;
}

exports.handler = async () => {
  // OPTIONS preflight
  if (arguments[0] && arguments[0].httpMethod === "OPTIONS") {
    return { statusCode: 204, headers: HEADERS, body: "" };
  }

  try {
    // 1. Download do Excel
    const resp = await fetch(EXCEL_URL);
    if (!resp.ok) throw new Error(`HTTP ${resp.status} ao descarregar Excel`);
    const buffer = await resp.buffer();

    // 2. Leitura da folha RE
    const wb = XLSX.read(buffer, { type: "buffer" });
    const sh = wb.Sheets["RE"];
    if (!sh) throw new Error('Folha "RE" não encontrada no Excel');

    // 3. Mês actual (1-12)
    const mes = new Date().getMonth() + 1; // 1 = Janeiro … 12 = Dezembro

    // Colunas: B=1, C=2, D=3, E=4, F=5, G=6
    const C = 2, E = 4, G = 6;
    const B = 1, D = 3, F = 5;

    const result = {
      mes,

      global: {
        mes_fat: val(sh, 6 + mes, C),
        mes_ang: val(sh, 6 + mes, E),
        mes_con: val(sh, 6 + mes, G),
        ano_fat: val(sh, 20, C),
        ano_ang: val(sh, 20, E),
        ano_con: val(sh, 20, G),
      },

      brg: {
        mes_fat:     val(sh, 26 + mes, C),
        mes_ang:     val(sh, 26 + mes, E),
        mes_con:     val(sh, 26 + mes, G),
        ano_fat:     val(sh, 27, C),
        ano_ang:     val(sh, 27, E),
        ano_con:     val(sh, 27, G),
        meta_fat:    Math.abs(val(sh, 28, B)),
        meta_ang:    Math.abs(val(sh, 28, D)),
        meta_con:    Math.abs(val(sh, 28, F)),
        meta_ano_fat: Math.abs(val(sh, 27, B)),
        meta_ano_ang: Math.abs(val(sh, 27, D)),
        meta_ano_con: Math.abs(val(sh, 27, F)),
      },

      bcl: {
        mes_fat:     val(sh, 39 + mes, C),
        mes_ang:     val(sh, 39 + mes, E),
        mes_con:     val(sh, 39 + mes, G),
        ano_fat:     val(sh, 40, C),
        ano_ang:     val(sh, 40, E),
        ano_con:     val(sh, 40, G),
        meta_fat:    Math.abs(val(sh, 41, B)),
        meta_ang:    Math.abs(val(sh, 41, D)),
        meta_con:    Math.abs(val(sh, 41, F)),
        meta_ano_fat: Math.abs(val(sh, 40, B)),
        meta_ano_ang: Math.abs(val(sh, 40, D)),
        meta_ano_con: Math.abs(val(sh, 40, F)),
      },

      bgc: {
        mes_fat:     val(sh, 52 + mes, C),
        mes_ang:     val(sh, 52 + mes, E),
        mes_con:     val(sh, 52 + mes, G),
        ano_fat:     val(sh, 53, C),
        ano_ang:     val(sh, 53, E),
        ano_con:     val(sh, 53, G),
        meta_fat:    Math.abs(val(sh, 54, B)),
        meta_ang:    Math.abs(val(sh, 54, D)),
        meta_con:    Math.abs(val(sh, 54, F)),
        meta_ano_fat: Math.abs(val(sh, 53, B)),
        meta_ano_ang: Math.abs(val(sh, 53, D)),
        meta_ano_con: Math.abs(val(sh, 53, F)),
      },
    };

    return {
      statusCode: 200,
      headers: HEADERS,
      body: JSON.stringify(result),
    };
  } catch (err) {
    return {
      statusCode: 500,
      headers: HEADERS,
      body: JSON.stringify({ error: err.message }),
    };
  }
};
