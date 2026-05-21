import * as XLSX from "xlsx";

export const EQUIPES = [
  "HIDRÁULICA",
  "PINTURA",
  "ELÉTRICA",
  "CIVIL",
  "MARCENARIA",
  "SERRALHERIA",
  "SERVIÇOS GERAIS",
] as const;

export type Equipe = (typeof EQUIPES)[number];

export interface ApontamentoRow {
  numero_os: string;
  tecnico_codigo: string;
  tecnico_nome: string;
  especialidade: string;
  equipe: string;
  tipo_solicitacao: string;
  unidade_negocio: string;
  data_inicio: string | null;
  data_fim: string | null;
  horas_decimais: number;
  preco_hora: number;
  valor_mao_obra: number;
  mes_ano: string;
  semana_do_mes: number;
}

/** "8:30:00" or "08:30" -> 8.5 */
export function timeToDecimal(v: unknown): number {
  if (v == null || v === "") return 0;

  if (v instanceof Date) {
    // O XLSX cria o Date representando a duração localmente a partir de 1899-12-30.
    const year = v.getFullYear();
    const month = v.getMonth();
    const date = v.getDate();
    const hours = v.getHours();
    const minutes = v.getMinutes();
    const seconds = v.getSeconds();
    
    // Calcula a diferença de dias locais com relação a 1899-12-30
    const baseDate = new Date(1899, 11, 30);
    const currentDate = new Date(year, month, date);
    const diffDays = Math.round((currentDate.getTime() - baseDate.getTime()) / (1000 * 60 * 60 * 24));
    
    const totalHours = (diffDays * 24) + hours + (minutes / 60) + (seconds / 3600);
    return Math.round(totalHours * 100) / 100;
  }

  if (typeof v === "number") {
    // Excel stores duration as fraction of a day
    if (v < 10) return Math.round(v * 24 * 100) / 100;
    return v;
  }
  const s = String(v).trim();
  if (!s.includes(":")) {
    const num = parseFloat(s.replace(",", "."));
    if (!isNaN(num)) return Math.round(num * 100) / 100;
    return 0;
  }
  const parts = s.split(":").map((p) => parseInt(p, 10));
  if (parts.some(Number.isNaN)) return 0;
  const [h = 0, m = 0, sec = 0] = parts;
  return Math.round((h + m / 60 + sec / 3600) * 100) / 100;
}

function excelDateToISO(v: unknown): string | null {
  if (!v) return null;
  if (v instanceof Date) return v.toISOString();
  if (typeof v === "number") {
    const d = XLSX.SSF.parse_date_code(v);
    if (!d) return null;
    return new Date(Date.UTC(d.y, d.m - 1, d.d, d.H, d.M, d.S)).toISOString();
  }
  const d = new Date(String(v));
  return Number.isNaN(d.getTime()) ? null : d.toISOString();
}

function weekOfMonth(d: Date): number {
  const day = d.getDate();
  return Math.min(4, Math.ceil(day / 7));
}

function strictNormalize(str: string): string {
  return str
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

function pick(row: Record<string, unknown>, ...keys: string[]): unknown {
  const rowKeys = Object.keys(row);

  // 1. Tenta correspondência exata (case insensitive e trim simples)
  for (const k of keys) {
    const target = k.trim().toLowerCase();
    for (const rk of rowKeys) {
      if (rk.trim().toLowerCase() === target) {
        return row[rk];
      }
    }
  }

  // 2. Tenta correspondência estritamente grudada (remove espaços, pontos, acentos, etc.)
  for (const k of keys) {
    const targetNormalized = strictNormalize(k);
    for (const rk of rowKeys) {
      if (strictNormalize(rk) === targetNormalized) {
        return row[rk];
      }
    }
  }

  // 3. Fallback inteligente para horas
  const isHorasQuery = keys.some(k => k.toLowerCase().includes("tempo") || k.toLowerCase().includes("horas"));
  if (isHorasQuery) {
    for (const rk of rowKeys) {
      const norm = strictNormalize(rk);
      if (
        (norm.includes("tempo") && norm.includes("trabalho")) ||
        (norm.includes("tempo") && norm.includes("feedback")) ||
        norm.includes("horastrabalhadas") ||
        norm.includes("tempotrabalho") ||
        norm.includes("tempotrabalhofeedback") ||
        norm.includes("tempotrabalhofeedbackmobra") ||
        norm === "horas" ||
        norm === "hh"
      ) {
        return row[rk];
      }
    }
  }

  // 4. Fallback inteligente para número de OS
  const isOsQuery = keys.some(k => k.toLowerCase().includes("os") || k.toLowerCase().includes("ordem"));
  if (isOsQuery) {
    for (const rk of rowKeys) {
      const norm = strictNormalize(rk);
      if (
        (norm.includes("numero") && norm.includes("os")) ||
        norm === "os" ||
        norm === "numeroos" ||
        norm.includes("numeroordem") ||
        norm.includes("numos")
      ) {
        return row[rk];
      }
    }
  }

  return undefined;
}

export async function parseExcel(file: File): Promise<ApontamentoRow[]> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: "" });

  return rows
    .map((r) => {
      const dataInicio = excelDateToISO(pick(r, "Data Início Trabalho", "Data Inicio Trabalho"));
      const horas = timeToDecimal(pick(r, "Tempo Trabalho Feedback M.Obra", "Tempo Trabalho"));
      const precoHora = Number(pick(r, "Preço Hora", "Preco Hora")) || 0;
      const valorMO = Number(pick(r, "Valor Mão-de-Obra", "Valor Mao-de-Obra", "Valor Mao de Obra")) || horas * precoHora;
      const dateObj = dataInicio ? new Date(dataInicio) : new Date();
      const mes_ano = `${dateObj.getFullYear()}-${String(dateObj.getMonth() + 1).padStart(2, "0")}`;

      return {
        numero_os: String(pick(r, "Número OS", "Numero OS") ?? "").trim(),
        tecnico_codigo: String(pick(r, "Técnico", "Tecnico") ?? "").trim(),
        tecnico_nome: String(pick(r, "Denominação Técnico", "Denominacao Tecnico") ?? "").trim(),
        especialidade: String(pick(r, "Denominação Especialidade", "Denominacao Especialidade") ?? "").trim(),
        equipe: String(pick(r, "Denominação Oficina", "Denominacao Oficina") ?? "").trim().toUpperCase(),
        tipo_solicitacao: String(pick(r, "Denominação Tipo Solicitação Serviço", "Denominacao Tipo Solicitacao Servico") ?? "").trim(),
        unidade_negocio: String(pick(r, "Denominação Unidade Negócio", "Denominacao Unidade Negocio") ?? "").trim().toUpperCase(),
        data_inicio: dataInicio,
        data_fim: excelDateToISO(pick(r, "Data Final Trabalho", "Data Fim Trabalho")),
        horas_decimais: horas,
        preco_hora: precoHora,
        valor_mao_obra: valorMO,
        mes_ano,
        semana_do_mes: weekOfMonth(dateObj),
      } satisfies ApontamentoRow;
    })
    .filter((r) => r.tecnico_nome && r.equipe);
}

export function currentMonthKey(d = new Date()): string {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

export function currentWeekOfMonth(d = new Date()): number {
  return Math.min(4, Math.ceil(d.getDate() / 7));
}

export function exportToCSV(rows: Record<string, unknown>[], filename: string) {
  if (!rows.length) return;
  const ws = XLSX.utils.json_to_sheet(rows);
  const csv = XLSX.utils.sheet_to_csv(ws);
  const blob = new Blob(["\ufeff" + csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}
