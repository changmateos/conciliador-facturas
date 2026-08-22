import type { FileRole, ParsedFile, SemanticField, SheetData, UuidMode } from "@/lib/excel/parser";
import {
  buildAutoMapping,
  buildAutoVisible,
  detectHeaderRow,
  getHeaders,
} from "@/lib/excel/parser";

// v2: se ignora la memoria anterior para re-aplicar el default de hoja "NO SIF"
const MEMORY_KEY = "conciliador-diccionario-v2";

export interface StoredConfig {
  uuidMode: UuidMode;
  sheetName: string | null;
  headerRow: number;
  mapping: Record<string, SemanticField>;
  visible: Record<string, boolean>;
}

export function loadAllMemory(): Record<string, StoredConfig> {
  try {
    const raw = localStorage.getItem(MEMORY_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw) as unknown;
    if (typeof parsed === "object" && parsed !== null) {
      return parsed as Record<string, StoredConfig>;
    }
    return {};
  } catch {
    return {};
  }
}

export function rememberFile(fileName: string, file: ParsedFile): void {
  try {
    const all = loadAllMemory();
    all[fileName] = {
      uuidMode: file.uuidMode,
      sheetName: file.sheetName,
      headerRow: file.headerRow,
      mapping: file.mapping,
      visible: file.visible,
    };
    localStorage.setItem(MEMORY_KEY, JSON.stringify(all));
  } catch {
    // sin almacenamiento disponible: continuar sin memoria
  }
}

// Hoja predeterminada: la primera que contenga "NO SIF"; si no, la de más filas.
function defaultSheetOf(sheets: SheetData[]): SheetData {
  let bestSheet = sheets[0];
  for (const s of sheets) {
    if (s.rows.length > bestSheet.rows.length) bestSheet = s;
  }
  const preferred = sheets.find((s) => s.name.toUpperCase().includes("NO SIF"));
  return preferred ?? bestSheet;
}

export function buildFileDraft(
  fileName: string,
  role: FileRole,
  sheets: SheetData[],
  autoUuidMode: UuidMode
): ParsedFile {
  const defaultSheet = defaultSheetOf(sheets);

  const stored = loadAllMemory()[fileName];

  const sheetName =
    stored && stored.sheetName && sheets.some((s) => s.name === stored.sheetName)
      ? stored.sheetName
      : defaultSheet.name;
  const sheet = sheets.find((s) => s.name === sheetName) ?? defaultSheet;

  const maxRow = Math.max(0, Math.min(14, sheet.rows.length - 1));
  const headerRow =
    stored &&
    Number.isInteger(stored.headerRow) &&
    stored.headerRow >= 0 &&
    stored.headerRow <= maxRow
      ? stored.headerRow
      : detectHeaderRow(sheet.rows);

  const headers = getHeaders(sheet.rows, headerRow);
  const autoMapping = buildAutoMapping(headers);
  const autoVisible = buildAutoVisible(headers, autoMapping);

  const mapping: Record<string, SemanticField> = {};
  const visible: Record<string, boolean> = {};
  for (const h of headers) {
    mapping[h] =
      stored && stored.mapping && h in stored.mapping
        ? stored.mapping[h]
        : autoMapping[h];
    visible[h] =
      stored && stored.visible && h in stored.visible
        ? Boolean(stored.visible[h])
        : autoVisible[h];
  }

  // garantizar a lo más una columna como UUID
  let uuidSeen = false;
  for (const h of headers) {
    if (mapping[h] === "UUID") {
      if (uuidSeen) mapping[h] = "NINGUNO";
      else uuidSeen = true;
    }
  }

  const uuidMode =
    stored && (stored.uuidMode === "TEXTO" || stored.uuidMode === "EXACTO")
      ? stored.uuidMode
      : autoUuidMode;

  return {
    id:
      fileName + "-" + Date.now().toString(36) + "-" + Math.random().toString(36).slice(2, 7),
    fileName,
    role,
    sheets,
    sheetName,
    headerRow,
    mapping,
    visible,
    uuidMode,
  };
}