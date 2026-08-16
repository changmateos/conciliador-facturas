"use client";

import { useState } from "react";
import type { NormalizedRow, ParsedFile, SemanticField, UuidMode } from "@/lib/excel/parser";
import { SEMANTIC_FIELDS, getHeaders } from "@/lib/excel/parser";

interface FileCardProps {
  file: ParsedFile;
  rows: NormalizedRow[];
  onSheetChange: (sheetName: string) => void;
  onHeaderRowChange: (rowIndex: number) => void;
  onMappingChange: (header: string, field: SemanticField) => void;
  onVisibleChange: (header: string, visible: boolean) => void;
  onUuidModeChange: (mode: UuidMode) => void;
  onRemove: () => void;
}

export default function FileCard({
  file,
  rows,
  onSheetChange,
  onHeaderRowChange,
  onMappingChange,
  onVisibleChange,
  onUuidModeChange,
  onRemove,
}: FileCardProps) {
  const [collapsed, setCollapsed] = useState(true);
  const sheet = file.sheets.find((s) => s.name === file.sheetName);
  const headers = sheet ? getHeaders(sheet.rows, file.headerRow) : [];
  const hasUuid = Object.values(file.mapping).includes("UUID");
  const ready = hasUuid && rows.length > 0;
  const previewCols = headers
    .filter((h) => file.visible[h] && file.mapping[h] !== "UUID")
    .slice(0, 6);
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
            onClick={() => setCollapsed(!collapsed)}
            className="text-[11px] font-semibold text-gray-600 border border-gray-300 rounded-lg px-2.5 py-1 hover:bg-gray-100"
          >
            {collapsed ? "Expandir ▼" : "Colapsar ▲"}
          </button>
          <button
            onClick={onRemove}
            title="Quitar archivo"
            className="w-7 h-7 rounded-lg border border-gray-300 text-gray-500 hover:bg-red-50 hover:text-red-600 hover:border-red-300 text-sm leading-none"
          >
            ×
          </button>
        </div>
      </div>

      {collapsed ? null : (
        <div className="p-5 space-y-5">
          <div className="grid sm:grid-cols-3 gap-4">
            <label className="block">
              <span className="text-[12px] font-semibold text-gray-600">Hoja del libro</span>
              <select
                value={file.sheetName}
                onChange={(e) => onSheetChange(e.target.value)}
                className="mt-1 w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
              >
                {file.sheets.map((s) => (
                  <option key={s.name} value={s.name}>
                    {s.name} · {s.rows.length} filas
                  </option>
                ))}
              </select>
            </label>
            <label className="block">
              <span className="text-[12px] font-semibold text-gray-600">Fila de encabezados</span>
              <select
                value={file.headerRow}
                onChange={(e) => onHeaderRowChange(Number(e.target.value))}
                className="mt-1 w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
              >
                {Array.from({ length: headerOptions }, (_, i) => (
                  <option key={i} value={i}>Fila {i + 1}</option>
                ))}
              </select>
            </label>
            <label className="block">
              <span className="text-[12px] font-semibold text-gray-600">Modo de extracción de UUID</span>
              <select
                value={file.uuidMode}
                onChange={(e) => onUuidModeChange(e.target.value as UuidMode)}
                className="mt-1 w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
              >
                <option value="EXACTO">Celda completa = UUID</option>
                <option value="TEXTO">Buscar UUID dentro del texto</option>
              </select>
            </label>
          </div>

          {file.uuidMode === "TEXTO" ? (
            <p className="text-[12px] text-blue-800 bg-blue-50 border border-blue-200 rounded-lg px-3 py-2">
              Se buscará el patrón 8-4-4-4-12 dentro de cada celda (ej. "Concepto del Movimiento").
              El texto completo se conserva como un campo más del diccionario y también se usa para la regla de palabras prohibidas; las filas sin UUID detectado se descartan del cruce.
            </p>
          ) : null}

          <div>
            <p className="text-[11px] font-bold tracking-wide text-gray-500 uppercase">
              Diccionario de datos · significado y visibilidad de cada columna
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
                          ? "border-gray-300 bg-white text-gray-900"
                          : "border-gray-300 bg-white text-gray-700")
                    }
                  >
                    {SEMANTIC_FIELDS.map((f) => (
                      <option key={f} value={f}>{f}</option>
                    ))}
                  </select>
                  <label className="mt-2 flex items-center gap-2 text-[11px] font-semibold text-gray-600 select-none cursor-pointer">
                    <input
                      type="checkbox"
                      checked={file.visible[header] ?? false}
                      onChange={(e) => onVisibleChange(header, e.target.checked)}
                      className="accent-blue-700 w-3.5 h-3.5"
                    />
                    Ver en tabla del dashboard
                  </label>
                </div>
              ))}
            </div>
            {!hasUuid ? (
              <p className="mt-2 text-[12px] text-amber-800 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">
                Marca con "UUID" la columna que contiene el folio fiscal (o la frase que lo contiene, si el modo es "dentro del texto") para activar este archivo.
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
                      {previewCols.map((c) => (
                        <th key={c} className="px-3 py-2 text-left font-semibold text-gray-500">{c}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {rows.slice(0, 8).map((r) => (
                      <tr key={r.sourceRow + "-" + r.uuid} className="border-t border-gray-100">
                        <td className="px-3 py-2 text-gray-400">{r.sourceRow}</td>
                        <td className="px-3 py-2 font-mono text-gray-800">{r.uuid}</td>
                        {previewCols.map((c) => (
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
                Aún no hay filas con UUID. Revisa la fila de encabezados, el mapeo de columnas o el modo de extracción.
              </p>
            )}
          </div>
        </div>
      )}
    </div>
  );
}