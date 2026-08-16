"use client";

import type { NormalizedRow, ParsedFile, SemanticField } from "@/lib/excel/parser";
import { SEMANTIC_FIELDS, getHeaders } from "@/lib/excel/parser";

interface FileCardProps {
  file: ParsedFile;
  rows: NormalizedRow[];
  onSheetChange: (sheetName: string) => void;
  onHeaderRowChange: (rowIndex: number) => void;
  onMappingChange: (header: string, field: SemanticField) => void;
  onRemove: () => void;
}

export default function FileCard({
  file,
  rows,
  onSheetChange,
  onHeaderRowChange,
  onMappingChange,
  onRemove,
}: FileCardProps) {
  const sheet = file.sheets.find((s) => s.name === file.sheetName);
  const headers = sheet ? getHeaders(sheet.rows, file.headerRow) : [];
  const hasUuid = Object.values(file.mapping).includes("UUID");
  const ready = hasUuid && rows.length > 0;
  const extraCols = headers
    .filter((h) => file.mapping[h] !== "NINGUNO" && file.mapping[h] !== "UUID")
    .slice(0, 4);
  const headerOptions = Math.min(15, sheet ? sheet.rows.length : 15);

  return (
    <div className="bg-white border border-gray-200 rounded-xl shadow-sm overflow-hidden">
      <div className="flex flex-wrap items-center justify-between gap-3 px-5 py-3.5 bg-gray-50 border-b border-gray-200">
        <div className="flex items-center gap-2.5 min-w-0">
          <span
            className={
              "text-[10px] font-bold rounded-full px-2.5 py-1 " +
              (file.role === "PRINCIPAL" ? "bg-blue-700 text-white" : "bg-gray-200 text-gray-700")
            }
          >
            {file.role === "PRINCIPAL" ? "PRINCIPAL" : "COMPLEMENTARIO"}
          </span>
          <p className="text-sm font-semibold text-gray-900 truncate" title={file.fileName}>
            {file.fileName}
          </p>
        </div>
        <div className="flex items-center gap-2">
          <span
            className={
              "text-[11px] font-semibold rounded-full px-2.5 py-1 " +
              (ready ? "bg-green-100 text-green-700" : "bg-amber-100 text-amber-700")
            }
          >
            {ready ? "LISTO · " + rows.length + " filas" : "PENDIENTE DE MAPEO"}
          </span>
          <button
            onClick={onRemove}
            title="Quitar archivo"
            className="w-7 h-7 rounded-lg border border-gray-300 text-gray-500 hover:bg-red-50 hover:text-red-600 hover:border-red-300 text-sm leading-none"
          >
            ×
          </button>
        </div>
      </div>

      <div className="p-5 space-y-5">
        <div className="grid sm:grid-cols-2 gap-4">
          <label className="block">
            <span className="text-[12px] font-semibold text-gray-600">Hoja del libro</span>
            <select
              value={file.sheetName}
              onChange={(e) => onSheetChange(e.target.value)}
              className="mt-1 w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm"
            >
              {file.sheets.map((s) => (
                <option key={s.name} value={s.name}>
                  {s.name} · {s.rows.length} filas
                </option>
              ))}
            </select>
          </label>
          <label className="block">
            <span className="text-[12px] font-semibold text-gray-600">
              Fila de encabezados (detectada automáticamente, ajustable)
            </span>
            <select
              value={file.headerRow}
              onChange={(e) => onHeaderRowChange(Number(e.target.value))}
              className="mt-1 w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm"
            >
              {Array.from({ length: headerOptions }, (_, i) => (
                <option key={i} value={i}>Fila {i + 1}</option>
              ))}
            </select>
          </label>
        </div>

        <div>
          <p className="text-[11px] font-bold tracking-wide text-gray-500 uppercase">
            Diccionario de datos · ¿qué significa cada columna?
          </p>
          <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3 mt-2">
            {headers.map((header) => (
              <div key={header} className="rounded-lg border border-gray-200 bg-gray-50 p-3">
                <p className="text-[12px] font-semibold text-gray-700 truncate" title={header}>
                  [{header}]
                </p>
                <select
                  value={file.mapping[header] ?? "NINGUNO"}
                  onChange={(e) => onMappingChange(header, e.target.value as SemanticField)}
                  className={
                    "mt-1.5 w-full rounded-lg border px-2.5 py-1.5 text-[12px] font-semibold " +
                    (file.mapping[header] === "UUID"
                      ? "border-blue-600 bg-blue-50 text-blue-700"
                      : file.mapping[header] !== "NINGUNO"
                        ? "border-gray-300 bg-white text-gray-800"
                        : "border-gray-200 bg-white text-gray-400")
                  }
                >
                  {SEMANTIC_FIELDS.map((f) => (
                    <option key={f} value={f}>{f}</option>
                  ))}
                </select>
              </div>
            ))}
          </div>
          {!hasUuid ? (
            <p className="mt-2 text-[12px] text-amber-800 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">
              Marca con "UUID" la columna que contiene el folio fiscal (por ejemplo "Concepto del Movimiento") para activar este archivo.
            </p>
          ) : null}
        </div>

        <div>
          <p className="text-[11px] font-bold tracking-wide text-gray-500 uppercase">
            Vista previa normalizada ·{" "}
            {rows.length > 0
              ? "primeras " + Math.min(8, rows.length) + " de " + rows.length + " filas"
              : "sin filas válidas todavía"}
          </p>
          {rows.length > 0 ? (
            <div className="mt-2 overflow-x-auto rounded-lg border border-gray-200">
              <table className="w-full text-[12px]">
                <thead className="bg-gray-50">
                  <tr>
                    <th className="px-3 py-2 text-left font-semibold text-gray-500">Fila</th>
                    <th className="px-3 py-2 text-left font-semibold text-gray-500">UUID</th>
                    {extraCols.map((c) => (
                      <th key={c} className="px-3 py-2 text-left font-semibold text-gray-500">{c}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {rows.slice(0, 8).map((r) => (
                    <tr key={r.sourceRow + "-" + r.uuid} className="border-t border-gray-100">
                      <td className="px-3 py-2 text-gray-400">{r.sourceRow}</td>
                      <td className="px-3 py-2 font-mono text-gray-800">{r.uuid}</td>
                      {extraCols.map((c) => (
                        <td key={c} className="px-3 py-2 text-gray-600 max-w-[220px] truncate">
                          {r.values[c] === null ? "" : String(r.values[c])}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <p className="mt-2 text-[12px] text-gray-500 bg-gray-50 border border-gray-200 rounded-lg px-3 py-2">
              Aún no hay filas con UUID. Revisa la fila de encabezados o el mapeo de columnas.
            </p>
          )}
        </div>
      </div>
    </div>
  );
}