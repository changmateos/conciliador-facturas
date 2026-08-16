import type { SupabaseClient } from "@supabase/supabase-js";
import type { NormalizedRow, ParsedFile } from "@/lib/excel/parser";

export type ResultStatus = "HISTORICO" | "COMPLEMENTARIO" | "NO_ENCONTRADA";

export interface ResultRow {
  uuid: string;
  status: ResultStatus;
  location: string;
  extraLocations: string[]; // otras ubicaciones donde también aparece
  aparicionesPrincipal: number; // veces que aparece en el archivo principal
  sourceRow: number;
  values: Record<string, string | number | null>;
}

export interface ComplementaryHit {
  fileName: string;
  sheetName: string;
  sourceRow: number;
}

export function buildComplementIndex(
  comps: { file: ParsedFile; rows: NormalizedRow[] }[]
): Map<string, ComplementaryHit[]> {
  const index = new Map<string, ComplementaryHit[]>();
  for (const { file, rows } of comps) {
    for (const r of rows) {
      const hit: ComplementaryHit = {
        fileName: file.fileName,
        sheetName: file.sheetName,
        sourceRow: r.sourceRow,
      };
      const list = index.get(r.uuid);
      if (list) list.push(hit);
      else index.set(r.uuid, [hit]);
    }
  }
  return index;
}

// Consulta el histórico en lotes de 100 UUID para no saturar a Supabase
export async function fetchHistoricalUuids(
  supabase: SupabaseClient,
  uuids: string[]
): Promise<Set<string>> {
  const found = new Set<string>();
  const chunk = 100;
  for (let i = 0; i < uuids.length; i += chunk) {
    const part = uuids.slice(i, i + chunk);
    if (part.length === 0) continue;
    const { data, error } = await supabase
      .from("historical_invoices")
      .select("uuid_fiscal")
      .in("uuid_fiscal", part);
    if (error) {
      throw new Error("Error consultando histórico: " + error.message);
    }
    for (const row of data ?? []) {
      found.add(String((row as { uuid_fiscal: string }).uuid_fiscal).toUpperCase());
    }
  }
  return found;
}

export function reconcile(
  principalRows: NormalizedRow[],
  historical: Set<string>,
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