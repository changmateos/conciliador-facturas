"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import type { ComplementaryHit, ResultRow } from "@/lib/conciliacion";

interface SegmentOption {
  key: string;
  label: string;
  count: number;
}

interface ResultsTableProps {
  results: ResultRow[];
  visibleCols: string[];
  segmentMap: Record<string, string[]>;
  groupKeysMap: Record<string, string[]>;
  segmentOptions: SegmentOption[];
  compIndex: Map<string, ComplementaryHit[]> | null;
  sourceFields: string[];
}

type Filter = "TODAS" | "HISTORICO" | "COMPLEMENTARIO" | "NO_ENCONTRADA" | "SENALADAS";

const PAGE = 200;

function segKey(r: ResultRow): string {
  return r.uuid + "|" + r.sourceRow;
}

function applyStatus(list: ResultRow[], filter: Filter): ResultRow[] {
  if (filter === "HISTORICO") return list.filter((r) => r.status === "HISTORICO");
  if (filter === "COMPLEMENTARIO") return list.filter((r) => r.status === "COMPLEMENTARIO");
  if (filter === "NO_ENCONTRADA") return list.filter((r) => r.status === "NO_ENCONTRADA");
  if (filter === "SENALADAS")
    return list.filter((r) => r.aparicionesPrincipal > 1 || r.extraLocations.length > 0);
  return list;
}

function applySegment(
  list: ResultRow[],
  segment: string,
  groupKeysMap: Record<string, string[]>
): ResultRow[] {
  if (segment === "TODOS") return list;
  if (segment === "SIN_REGLA")
    return list.filter((r) => (groupKeysMap[segKey(r)] ?? []).length === 0);
  return list.filter((r) => (groupKeysMap[segKey(r)] ?? []).includes(segment));
}

export default function ResultsTable({
  results,
  visibleCols,
  segmentMap,
  groupKeysMap,
  segmentOptions,
  compIndex,
  sourceFields,
}: ResultsTableProps) {
  const [filter, setFilter] = useState<Filter>("TODAS");
  const [segment, setSegment] = useState("TODOS");
  const [query, setQuery] = useState("");
  const [limit, setLimit] = useState(PAGE);

  const byQuery = useMemo(() => {
    const q = query.trim().toUpperCase();
    if (q === "") return results;
    return results.filter((r) => r.uuid.includes(q));
  }, [results, query]);

  const tabList = useMemo(
    () => applySegment(byQuery, segment, groupKeysMap),
    [byQuery, segment, groupKeysMap]
  );

  const counts = useMemo(
    () => ({
      TODAS: tabList.length,
      HISTORICO: tabList.filter((r) => r.status === "HISTORICO").length,
      COMPLEMENTARIO: tabList.filter((r) => r.status === "COMPLEMENTARIO").length,
      NO_ENCONTRADA: tabList.filter((r) => r.status === "NO_ENCONTRADA").length,
      SENALADAS: tabList.filter(
        (r) => r.aparicionesPrincipal > 1 || r.extraLocations.length > 0
      ).length,
    }),
    [tabList]
  );

  const segList = useMemo(() => applyStatus(byQuery, filter), [byQuery, filter]);

  const segOptionsDynamic = useMemo(() => {
    const opts: SegmentOption[] = [
      { key: "TODOS", label: "Todos los segmentos", count: segList.length },
    ];
    for (const o of segmentOptions) {
      opts.push({
        key: o.key,
        label: o.label,
        count: applySegment(segList, o.key, groupKeysMap).length,
      });
    }
    return opts;
  }, [segList, segmentOptions, groupKeysMap]);

  const filtered = useMemo(() => applyStatus(tabList, filter), [tabList, filter]);

  const shown = filtered.slice(0, limit);

  function firstHit(r: ResultRow): ComplementaryHit | null {
    if (!compIndex) return null;
    const hits = compIndex.get(r.uuid);
    return hits && hits.length > 0 ? hits[0] : null;
  }

  function exportXlsx() {
    const rows = filtered.map((r) => {
      const base: Record<string, string | number> = {
        UUID: r.uuid,
        Estatus:
          r.status === "HISTORICO"
            ? "ENCONTRADA EN HISTÓRICO"
            : r.status === "COMPLEMENTARIO"
              ? "ENCONTRADA EN COMPLEMENTARIO"
              : "NO ENCONTRADA",
        Segmentacion: (segmentMap[segKey(r)] ?? []).join(" | "),
        Ubicacion: r.status === "NO_ENCONTRADA" ? "" : r.location,
        DuplicadoEnPrincipal:
          r.aparicionesPrincipal > 1 ? "SI x" + r.aparicionesPrincipal : "",
        TambienEn: r.extraLocations.join(" | "),
        FilaOrigen: r.sourceRow,
      };
      const hit = firstHit(r);
      for (const sf of sourceFields) {
        const v = hit && hit.mappedValues ? hit.mappedValues[sf] : null;
        base["Fuente · " + sf] = v === null || v === undefined ? "" : v;
      }
      for (const c of visibleCols) {
        const v = r.values[c];
        base[c] = v === null || v === undefined ? "" : v;
      }
      return base;
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Resultados");
    XLSX.writeFile(wb, "conciliacion_resultados.xlsx");
  }

  const tabs: { key: Filter; label: string }[] = [
    { key: "TODAS", label: "Todas" },
    { key: "HISTORICO", label: "En histórico" },
    { key: "COMPLEMENTARIO", label: "En complementarios" },
    { key: "NO_ENCONTRADA", label: "No encontradas" },
    { key: "SENALADAS", label: "Señaladas" },
  ];

  return (
    <div className="bg-white border border-gray-200 rounded-xl overflow-hidden">
      <div className="flex flex-wrap items-center justify-between gap-3 px-5 py-4 border-b border-gray-200 bg-gray-50">
        <div className="flex flex-wrap gap-2">
          {tabs.map((t) => (
            <button
              key={t.key}
              onClick={() => { setFilter(t.key); setLimit(PAGE); }}
              className={
                "text-[12px] font-semibold rounded-full px-3 py-1.5 border transition-colors " +
                (filter === t.key
                  ? "bg-blue-700 text-white border-blue-700"
                  : "bg-white text-gray-600 border-gray-300 hover:border-blue-500")
              }
            >
              {t.label} · {counts[t.key]}
            </button>
          ))}
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <select
            value={segment}
            onChange={(e) => { setSegment(e.target.value); setLimit(PAGE); }}
            className="rounded-lg border border-gray-300 bg-white px-3 py-1.5 text-[12px] font-semibold text-gray-900 max-w-[280px]"
            title="Filtrar por grupo de regla (los conteos respetan la pestaña activa)"
          >
            {segOptionsDynamic.map((o) => (
              <option key={o.key} value={o.key}>
                {o.label} · {o.count}
              </option>
            ))}
          </select>
          <input
            value={query}
            onChange={(e) => { setQuery(e.target.value); setLimit(PAGE); }}
            placeholder="Buscar UUID…"
            className="rounded-lg border border-gray-300 px-3 py-1.5 text-[12px] w-44 text-gray-900"
          />
          <button
            onClick={exportXlsx}
            className="text-[12px] font-semibold rounded-lg bg-gray-900 text-white px-3.5 py-1.5 hover:bg-gray-700"
          >
            Exportar .xlsx
          </button>
        </div>
      </div>

      <div className="overflow-auto max-h-[560px]">
        <table className="w-full text-[12px]">
          <thead className="bg-gray-50 sticky top-0">
            <tr>
              <th className="px-3 py-2 text-left font-semibold text-gray-500">Estatus</th>
              <th className="px-3 py-2 text-left font-semibold text-gray-500">UUID</th>
              <th className="px-3 py-2 text-left font-semibold text-gray-500">Segmentación</th>
              <th className="px-3 py-2 text-left font-semibold text-gray-500">Señalamientos</th>
              <th className="px-3 py-2 text-left font-semibold text-gray-500">Ubicación exacta</th>
              {sourceFields.map((sf) => (
                <th key={sf} className="px-3 py-2 text-left font-semibold text-blue-700 bg-blue-50/50">
                  Fuente · {sf}
                </th>
              ))}
              {visibleCols.map((c) => (
                <th key={c} className="px-3 py-2 text-left font-semibold text-gray-500">{c}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {shown.map((r, i) => {
              const segs = segmentMap[segKey(r)] ?? [];
              const hit = firstHit(r);
              return (
                <tr key={r.uuid + "-" + r.sourceRow + "-" + i} className="border-t border-gray-100 hover:bg-gray-50/60">
                  <td className="px-3 py-2">
                    {r.status === "HISTORICO" ? (
                      <span className="text-[10px] font-bold rounded-full px-2 py-0.5 bg-green-100 text-green-700">HISTÓRICO</span>
                    ) : r.status === "COMPLEMENTARIO" ? (
                      <span className="text-[10px] font-bold rounded-full px-2 py-0.5 bg-green-100 text-green-700 border border-green-300">COMPLEMENTARIO</span>
                    ) : (
                      <span className="text-[10px] font-bold rounded-full px-2 py-0.5 bg-red-100 text-red-700">NO ENCONTRADA</span>
                    )}
                  </td>
                  <td className="px-3 py-2 font-mono text-gray-800">{r.uuid}</td>
                  <td className="px-3 py-2">
                    {segs.length > 0 ? (
                      <span className="flex flex-wrap gap-1">
                        {segs.map((s) => (
                          <span key={s} className="text-[10px] font-semibold bg-blue-50 text-blue-700 border border-blue-200 rounded-full px-2 py-0.5">
                            {s}
                          </span>
                        ))}
                      </span>
                    ) : (
                      <span className="text-gray-300">—</span>
                    )}
                  </td>
                  <td className="px-3 py-2">
                    <span className="flex flex-wrap gap-1">
                      {r.aparicionesPrincipal > 1 ? (
                        <span
                          title={"Aparece " + r.aparicionesPrincipal + " veces en el archivo principal"}
                          className="text-[10px] font-bold rounded-full px-2 py-0.5 bg-amber-100 text-amber-700"
                        >
                          DUP ×{r.aparicionesPrincipal}
                        </span>
                      ) : null}
                      {r.extraLocations.length > 0 ? (
                        <span
                          title={"También en: " + r.extraLocations.join(" | ")}
                          className="text-[10px] font-bold rounded-full px-2 py-0.5 bg-amber-100 text-amber-700"
                        >
                          +{r.extraLocations.length} ubicación{r.extraLocations.length > 1 ? "es" : ""}
                        </span>
                      ) : null}
                      {r.aparicionesPrincipal <= 1 && r.extraLocations.length === 0 ? (
                        <span className="text-gray-300">—</span>
                      ) : null}
                    </span>
                  </td>
                  <td className="px-3 py-2 text-gray-600 max-w-[280px]">
                    <span
                      className="block truncate"
                      title={r.location + (r.extraLocations.length > 0 ? " | " + r.extraLocations.join(" | ") : "")}
                    >
                      {r.status === "NO_ENCONTRADA" ? "—" : r.location}
                    </span>
                  </td>
                  {sourceFields.map((sf) => {
                    const v = hit && hit.mappedValues ? hit.mappedValues[sf] : null;
                    return (
                      <td key={sf} className="px-3 py-2 text-gray-700 max-w-[220px] truncate bg-blue-50/30">
                        {v === null || v === undefined ? "—" : String(v)}
                      </td>
                    );
                  })}
                  {visibleCols.map((c) => (
                    <td key={c} className="px-3 py-2 text-gray-600 max-w-[200px] truncate">
                      {r.values[c] === null || r.values[c] === undefined ? "" : String(r.values[c])}
                    </td>
                  ))}
                </tr>
              );
            })}
            {shown.length === 0 ? (
              <tr>
                <td colSpan={5 + sourceFields.length + visibleCols.length} className="px-3 py-6 text-center text-gray-400">
                  Sin resultados para esta combinación de filtros.
                </td>
              </tr>
            ) : null}
          </tbody>
        </table>
      </div>

      {filtered.length > limit ? (
        <div className="px-5 py-3 border-t border-gray-200 bg-gray-50">
          <button
            onClick={() => setLimit(limit + PAGE)}
            className="text-[12px] font-semibold text-blue-700 hover:underline"
          >
            Mostrar 200 más ({filtered.length - limit} restantes)
          </button>
        </div>
      ) : null}
    </div>
  );
}