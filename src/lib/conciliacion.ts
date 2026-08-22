import type { SupabaseClient } from "@supabase/supabase-js";
import type { NormalizedRow, ParsedFile } from "@/lib/excel/parser";

export type ResultStatus = "HISTORICO" | "COMPLEMENTARIO" | "NO_ENCONTRADA";

export interface ResultRow {
  uuid: string;
  status: ResultStatus;
  location: string;
  extraLocations: string[];
  aparicionesPrincipal: number;
  sourceRow: number;
  values: Record<string, string | number | null>;
}

export interface HistInfo {
  mes: number | null;
  anio: number | null;
  origen: string | null;
}

export interface ComplementaryHit {
  fileName: string;
  sheetName: string;
  sourceRow: number;
  detalle: string;
  mappedValues: Record<string, string | number | null>;
  // v4: todos los valores de la fila del complementario (para JOIN y selector)
  values: Record<string, string | number | null>;
}

const DETALLE_KEYS = /fecha|concepto|folio|raz[oó]n|rfc/i;

function buildDetalle(values: Record<string, string | number | null>): string {
  const parts: string[] = [];
  for (const [k, v] of Object.entries(values)) {
    if (k.toLowerCase().includes("uuid")) continue;
    if (v === null || v === undefined || String(v).trim() === "") continue;
    if (DETALLE_KEYS.test(k)) parts.push(k + ": " + String(v));
    if (parts.length >= 3) break;
  }
  return parts.join(" · ");
}

export function buildComplementIndex(
  comps: { file: ParsedFile; rows: NormalizedRow[] }[]
): Map<string, ComplementaryHit[]> {
  const index = new Map<string, ComplementaryHit[]>();
  for (const { file, rows } of comps) {
    for (const r of rows) {
      const mappedValues: Record<string, string | number | null> = {};
      for (const [header, field] of Object.entries(file.mapping)) {
        if (field === "NINGUNO" || field === "UUID") continue;
        mappedValues[field] = r.values[header] ?? null;
      }
      const hit: ComplementaryHit = {
        fileName: file.fileName,
        sheetName: file.sheetName,
        sourceRow: r.sourceRow,
        detalle: buildDetalle(r.values),
        mappedValues,
        values: r.values,
      };
      const list = index.get(r.uuid);
      if (list) list.push(hit);
      else index.set(r.uuid, [hit]);
    }
  }
  return index;
}

export async function fetchHistoricalMap(
  supabase: SupabaseClient,
  uuids: string[]
): Promise<Map<string, HistInfo>> {
  const found = new Map<string, HistInfo>();
  const chunk = 100;
  for (let i = 0; i < uuids.length; i += chunk) {
    const part = uuids.slice(i, i + chunk);
    if (part.length === 0) continue;
    const { data, error } = await supabase
      .from("historical_invoices")
      .select("uuid_fiscal, mes_periodo, anio_periodo, origen_archivo")
      .in("uuid_fiscal", part);
    if (error) {
      throw new Error("Error consultando histórico: " + error.message);
    }
    for (const row of (data ?? []) as {
      uuid_fiscal: string;
      mes_periodo: number | null;
      anio_periodo: number | null;
      origen_archivo: string | null;
    }[]) {
      found.set(String(row.uuid_fiscal).toUpperCase(), {
        mes: row.mes_periodo,
        anio: row.anio_periodo,
        origen: row.origen_archivo,
      });
    }
  }
  return found;
}

export function reconcile(
  principalRows: NormalizedRow[],
  historical: Map<string, HistInfo>,
  compIndex: Map<string, ComplementaryHit[]>
): ResultRow[] {
  const counts = new Map<string, number>();
  for (const r of principalRows) {
    counts.set(r.uuid, (counts.get(r.uuid) ?? 0) + 1);
  }

  const results: ResultRow[] = [];
  for (const r of principalRows) {
    const inHistorical = historical.has(r.uuid);
    const compHits = compIndex.get(r.uuid) ?? [];
    let status: ResultStatus = "NO_ENCONTRADA";
    let location = "";
    let extra: string[] = [];

    if (inHistorical) {
      status = "HISTORICO";
      location = "Encontrada en histórico acumulado";
      extra = compHits.map(
        (h) => h.fileName + " · " + h.sheetName + " · fila " + h.sourceRow
      );
    } else if (compHits.length > 0) {
      status = "COMPLEMENTARIO";
      const first = compHits[0];
      location = first.fileName + " · " + first.sheetName + " · fila " + first.sourceRow;
      extra = compHits
        .slice(1)
        .map((h) => h.fileName + " · " + h.sheetName + " · fila " + h.sourceRow);
    }

    results.push({
      uuid: r.uuid,
      status,
      location,
      extraLocations: extra,
      aparicionesPrincipal: counts.get(r.uuid) ?? 1,
      sourceRow: r.sourceRow,
      values: r.values,
    });
  }
  return results;
}