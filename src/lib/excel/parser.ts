import * as XLSX from "xlsx";

export type SemanticField =
  | "UUID"
  | "FOLIO"
  | "MONTO"
  | "RETENCIONES"
  | "CONCEPTO"
  | "FECHA"
  | "NINGUNO";

export const SEMANTIC_FIELDS: SemanticField[] = [
  "NINGUNO",
  "UUID",
  "FOLIO",
  "MONTO",
  "RETENCIONES",
  "CONCEPTO",
  "FECHA",
];

export type FileRole = "PRINCIPAL" | "COMPLEMENTARIO";

export type CellValue = string | number | boolean | null;

export interface SheetData {
  name: string;
  rows: CellValue[][];
}

export interface NormalizedRow {
  sourceRow: number; // fila real dentro de la hoja (1-based)
  uuid: string;
  values: Record<string, string | number | null>;
}

export interface ParsedFile {
  id: string;
  fileName: string;
  role: FileRole;
  sheets: SheetData[];
  sheetName: string;
  headerRow: number; // índice 0-based dentro de rows
  mapping: Record<string, SemanticField>; // encabezado -> significado
}

export function cellText(cell: CellValue): string {
  if (cell === null || cell === undefined) return "";
  return String(cell).trim();
}

export function readWorkbook(buffer: ArrayBuffer): SheetData[] {
  const workbook = XLSX.read(new Uint8Array(buffer), { type: "array" });
  return workbook.SheetNames.map((name) => {
    const worksheet = workbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: null,
      raw: true,
      blankrows: true,
    }) as CellValue[][];
    return { name, rows };
  });
}

const HEADER_KEYWORDS =
  /uuid|folio|monto|importe|retenc|concepto|fecha|rfc|raz[oó]n|clave/i;

export function detectHeaderRow(rows: CellValue[][], maxScan = 15): number {
  const limit = Math.min(maxScan, rows.length);
  let bestRow = 0;
  let bestScore = -1;
  for (let i = 0; i < limit; i++) {
    const row = rows[i] || [];
    const nonEmpty = row.filter((c) => cellText(c) !== "");
    if (nonEmpty.length === 0) continue;
    let score = 0;
    for (const cell of nonEmpty) {
      if (typeof cell === "string") {
        score += 2;
        if (HEADER_KEYWORDS.test(cell)) score += 3;
      } else {
        score -= 1; // los números suelen ser datos, no encabezados
      }
    }
    if (score > bestScore) {
      bestScore = score;
      bestRow = i;
    }
  }
  return bestRow;
}

export function getHeaders(rows: CellValue[][], headerRow: number): string[] {
  const row = rows[headerRow] || [];
  const seen: Record<string, number> = {};
  const headers: string[] = [];
  for (let i = 0; i < row.length; i++) {
    const text = cellText(row[i]);
    let header = text !== "" ? text : "Columna " + (i + 1);
    if (seen[header] !== undefined) {
      seen[header] += 1;
      header = header + " (" + (seen[header] + 1) + ")";
    } else {
      seen[header] = 0;
    }
    headers.push(header);
  }
  return headers;
}

export function suggestField(header: string): SemanticField {
  const h = header.toLowerCase();
  if (h.includes("uuid") || h.includes("folio fiscal")) return "UUID";
  if (h.includes("retenc")) return "RETENCIONES";
  if (h.includes("monto") || h.includes("importe") || h.includes("total")) return "MONTO";
  if (h.includes("folio")) return "FOLIO";
  if (h.includes("concepto") || h.includes("descrip")) return "CONCEPTO";
  if (h.includes("fecha")) return "FECHA";
  return "NINGUNO";
}

export function buildAutoMapping(headers: string[]): Record<string, SemanticField> {
  const used = new Set<SemanticField>();
  const mapping: Record<string, SemanticField> = {};
  for (const header of headers) {
    let field = suggestField(header);
    if (field !== "NINGUNO" && used.has(field)) field = "NINGUNO";
    if (field !== "NINGUNO") used.add(field);
    mapping[header] = field;
  }
  return mapping;
}

export function toNumber(cell: CellValue): number | null {
  if (typeof cell === "number") return cell;
  const text = cellText(cell).replace(/[$,\s]/g, "");
  if (text === "") return null;
  const value = Number(text);
  return Number.isFinite(value) ? value : null;
}

export function normalizeRows(file: ParsedFile): NormalizedRow[] {
  const sheet = file.sheets.find((s) => s.name === file.sheetName);
  if (!sheet) return [];
  const headers = getHeaders(sheet.rows, file.headerRow);
  const uuidHeader = headers.find((h) => file.mapping[h] === "UUID");
  if (!uuidHeader) return [];
  const uuidIndex = headers.indexOf(uuidHeader);
  const result: NormalizedRow[] = [];
  for (let i = file.headerRow + 1; i < sheet.rows.length; i++) {
    const row = sheet.rows[i] || [];
    if (!row.some((c) => cellText(c) !== "")) continue; // fila en blanco
    const uuid = cellText(row[uuidIndex]).toUpperCase();
    if (uuid === "") continue; // fila sin UUID no sirve para el cruce
    const values: Record<string, string | number | null> = {};
    headers.forEach((header, index) => {
      const field = file.mapping[header];
      if (field === "MONTO" || field === "RETENCIONES") {
        values[header] = toNumber(row[index] ?? null);
      } else {
        values[header] = cellText(row[index] ?? null);
      }
    });
    result.push({ sourceRow: i + 1, uuid, values });
  }
  return result;
}

export function uuidColumnOf(file: ParsedFile): string | null {
  const sheet = file.sheets.find((s) => s.name === file.sheetName);
  if (!sheet) return null;
  const headers = getHeaders(sheet.rows, file.headerRow);
  return headers.find((h) => file.mapping[h] === "UUID") ?? null;
}