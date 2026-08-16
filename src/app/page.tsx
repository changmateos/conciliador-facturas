"use client";

import { useEffect, useMemo, useState } from "react";
import { createClient } from "@/lib/supabase/client";
import LogoutButton from "@/components/logout-button";
import Dropzone from "@/components/dropzone";
import FileCard from "@/components/file-card";
import type { FileRole, NormalizedRow, ParsedFile, SemanticField, UuidMode } from "@/lib/excel/parser";
import {
  buildAutoMapping,
  buildAutoVisible,
  detectHeaderRow,
  getHeaders,
  normalizeRows,
  readWorkbook,
  uuidColumnOf,
} from "@/lib/excel/parser";

export default function DashboardPage() {
  const supabase = useMemo(() => createClient(), []);
  const [email, setEmail] = useState("");
  const [files, setFiles] = useState<ParsedFile[]>([]);
  const [reading, setReading] = useState(false);

  useEffect(() => {
    supabase.auth.getUser().then(({ data }) => {
      setEmail(data.user?.email ?? "");
    });
  }, [supabase]);

  const rowsById = useMemo(() => {
    const map: Record<string, NormalizedRow[]> = {};
    for (const f of files) map[f.id] = normalizeRows(f);
    return map;
  }, [files]);

  const principal = files.find((f) => f.role === "PRINCIPAL") ?? null;
  const complementarios = files.filter((f) => f.role === "COMPLEMENTARIO");
  const readyCount = files.filter(
    (f) => (rowsById[f.id] ?? []).length > 0 && Object.values(f.mapping).includes("UUID")
  ).length;
  const principalRows = principal ? (rowsById[principal.id] ?? []).length : 0;
  const compRows = complementarios.reduce((acc, f) => acc + (rowsById[f.id] ?? []).length, 0);

  async function addFiles(list: File[], role: FileRole) {
    setReading(true);
    const added: ParsedFile[] = [];
    for (const file of list) {
      try {
        const buffer = await file.arrayBuffer();
        const sheets = readWorkbook(buffer);
        if (sheets.length === 0) continue;
        let bestSheet = sheets[0];
        for (const s of sheets) {
          if (s.rows.length > bestSheet.rows.length) bestSheet = s;
        }
        const headerRow = detectHeaderRow(bestSheet.rows);
        const headers = getHeaders(bestSheet.rows, headerRow);
        const mapping = buildAutoMapping(headers);
        added.push({
          id:
            file.name + "-" + Date.now().toString(36) + "-" + Math.random().toString(36).slice(2, 7),
          fileName: file.name,
          role,
          sheets,
          sheetName: bestSheet.name,
          headerRow,
          mapping,
          visible: buildAutoVisible(headers, mapping),
          uuidMode: /mayor/i.test(file.name) ? "TEXTO" : "EXACTO",
        });
      } catch {
        window.alert("No se pudo leer el archivo: " + file.name);
      }
    }
    setFiles((prev) =>
      role === "PRINCIPAL"
        ? [...prev.filter((f) => f.role !== "PRINCIPAL"), ...added]
        : [...prev, ...added]
    );
    setReading(false);
  }

  function updateFile(id: string, updater: (f: ParsedFile) => ParsedFile) {
    setFiles((prev) => prev.map((f) => (f.id === id ? updater(f) : f)));
  }

  function handleSheetChange(id: string, sheetName: string) {
    updateFile(id, (f) => {
      const sheet = f.sheets.find((s) => s.name === sheetName);
      if (!sheet) return f;
      const headerRow = detectHeaderRow(sheet.rows);
      const headers = getHeaders(sheet.rows, headerRow);
      const mapping = buildAutoMapping(headers);
      return { ...f, sheetName, headerRow, mapping, visible: buildAutoVisible(headers, mapping) };
    });
  }

  function handleHeaderRowChange(id: string, headerRow: number) {
    updateFile(id, (f) => {
      const sheet = f.sheets.find((s) => s.name === f.sheetName);
      if (!sheet) return f;
      const headers = getHeaders(sheet.rows, headerRow);
      const mapping = buildAutoMapping(headers);
      return { ...f, headerRow, mapping, visible: buildAutoVisible(headers, mapping) };
    });
  }

  function handleMappingChange(id: string, header: string, field: SemanticField) {
    updateFile(id, (f) => {
      const mapping = { ...f.mapping };
      if (field !== "NINGUNO") {
        for (const key of Object.keys(mapping)) {
          if (mapping[key] === field) mapping[key] = "NINGUNO";
        }
      }
      mapping[header] = field;
      const visible = { ...f.visible };
      if (field !== "NINGUNO") visible[header] = true;
      return { ...f, mapping, visible };
    });
  }

  function handleVisibleChange(id: string, header: string, visible: boolean) {
    updateFile(id, (f) => ({ ...f, visible: { ...f.visible, [header]: visible } }));
  }

  function handleUuidModeChange(id: string, mode: UuidMode) {
    updateFile(id, (f) => ({ ...f, uuidMode: mode }));
  }

  function removeFile(id: string) {
    setFiles((prev) => prev.filter((f) => f.id !== id));
  }

  const principalUuidCol = principal ? uuidColumnOf(principal) : null;

  return (
    <main className="min-h-screen bg-gray-100">
      <header className="bg-white border-b border-gray-200 sticky top-0 z-10">
        <div className="max-w-6xl mx-auto px-6 py-4 flex items-center justify-between gap-4">
          <div>
            <h1 className="font-bold text-gray-900">Conciliador de Facturas</h1>
            <p className="text-[11px] text-gray-500">Fase 2 · Carga y diccionario de datos (ajustado)</p>
          </div>
          <div className="flex items-center gap-4">
            <span className="text-sm text-gray-600 hidden sm:inline">{email}</span>
            <LogoutButton />
          </div>
        </div>
      </header>

      <div className="max-w-6xl mx-auto px-6 py-8 space-y-6">
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
          <div className="bg-white border border-gray-200 rounded-xl p-4">
            <p className="text-[11px] font-bold text-gray-500 uppercase tracking-wide">Archivos cargados</p>
            <p className="text-2xl font-bold text-gray-900 mt-1">{files.length}</p>
          </div>
          <div className="bg-white border border-gray-200 rounded-xl p-4">
            <p className="text-[11px] font-bold text-gray-500 uppercase tracking-wide">Listos para cruce</p>
            <p className={"text-2xl font-bold mt-1 " + (readyCount > 0 ? "text-green-600" : "text-amber-500")}>{readyCount}</p>
          </div>
          <div className="bg-white border border-gray-200 rounded-xl p-4">
            <p className="text-[11px] font-bold text-gray-500 uppercase tracking-wide">Filas · principal</p>
            <p className="text-2xl font-bold text-blue-700 mt-1">{principalRows}</p>
          </div>
          <div className="bg-white border border-gray-200 rounded-xl p-4">
            <p className="text-[11px] font-bold text-gray-500 uppercase tracking-wide">Filas · complementarios</p>
            <p className="text-2xl font-bold text-gray-900 mt-1">{compRows}</p>
          </div>
        </div>

        <div className="grid lg:grid-cols-2 gap-4">
          <Dropzone
            title="Archivo principal (lista a evaluar)"
            subtitle="Arrastra tu Excel A o haz clic para elegirlo (.xlsx, .xls, .csv). Si cargas otro, lo reemplaza."
            multiple={false}
            onFiles={(list) => addFiles(list, "PRINCIPAL")}
          />
          <Dropzone
            title="Archivos complementarios (fuentes del mes)"
            subtitle="Arrastra B1, B2, B3… puedes seleccionar varios a la vez."
            multiple={true}
            onFiles={(list) => addFiles(list, "COMPLEMENTARIO")}
          />
        </div>

        {reading ? (
          <p className="text-sm text-gray-600 bg-white border border-gray-200 rounded-xl px-4 py-3">
            Leyendo archivos…
          </p>
        ) : null}

        {files.length > 0 ? (
          <div className="space-y-5">
            {principal ? (
              <FileCard
                file={principal}
                rows={rowsById[principal.id] ?? []}
                onSheetChange={(s) => handleSheetChange(principal.id, s)}
                onHeaderRowChange={(r) => handleHeaderRowChange(principal.id, r)}
                onMappingChange={(h, field) => handleMappingChange(principal.id, h, field)}
                onVisibleChange={(h, v) => handleVisibleChange(principal.id, h, v)}
                onUuidModeChange={(m) => handleUuidModeChange(principal.id, m)}
                onRemove={() => removeFile(principal.id)}
              />
            ) : null}
            {complementarios.map((f) => (
              <FileCard
                key={f.id}
                file={f}
                rows={rowsById[f.id] ?? []}
                onSheetChange={(s) => handleSheetChange(f.id, s)}
                onHeaderRowChange={(r) => handleHeaderRowChange(f.id, r)}
                onMappingChange={(h, field) => handleMappingChange(f.id, h, field)}
                onVisibleChange={(h, v) => handleVisibleChange(f.id, h, v)}
                onUuidModeChange={(m) => handleUuidModeChange(f.id, m)}
                onRemove={() => removeFile(f.id)}
              />
            ))}
          </div>
        ) : (
          <p className="text-sm text-gray-500 text-center py-6">
            Aún no has cargado archivos. Cuando lo hagas, aquí aparecerá la tarjeta de mapeo de cada uno.
          </p>
        )}

        {files.length > 0 ? (
          <div className="bg-white border border-gray-200 rounded-xl p-6 space-y-4">
            <div>
              <h2 className="text-sm font-bold text-gray-900">Diccionario de datos del mes</h2>
              <p className="text-[12px] text-gray-500 mt-0.5">
                Todos estos campos se guardarán en historical_invoices.datos_json al consolidar el mes (Fase 5).
              </p>
            </div>
            {files.map((f) => {
              const sheet = f.sheets.find((s) => s.name === f.sheetName);
              const headers = sheet ? getHeaders(sheet.rows, f.headerRow) : [];
              const col = uuidColumnOf(f);
              return (
                <div key={f.id} className="border border-gray-200 rounded-lg p-4">
                  <p className="text-[12px] font-bold text-gray-800 font-mono truncate">
                    {f.fileName}
                    <span className="text-gray-400 font-normal">
                      {" "}· hoja "{f.sheetName}" · {f.uuidMode === "TEXTO" ? "UUID dentro de texto" : "UUID exacto"}
                    </span>
                  </p>
                  <ul className="mt-2 grid sm:grid-cols-2 gap-x-6 gap-y-1 text-[12px] font-mono text-gray-600">
                    {headers.map((h) => (
                      <li key={h} className="flex items-center gap-2 min-w-0">
                        {f.visible[h] ? (
                          <span className="text-[9px] font-bold bg-gray-200 text-gray-600 rounded px-1 shrink-0">TABLA</span>
                        ) : (
                          <span className="text-[9px] font-bold bg-gray-50 text-gray-300 rounded px-1 shrink-0">—</span>
                        )}
                        <span className="truncate">[{h}]</span>
                        <span className={f.mapping[h] !== "NINGUNO" ? "text-blue-700 font-semibold shrink-0" : "text-gray-400 shrink-0"}>
                          → {f.mapping[h] !== "NINGUNO" ? f.mapping[h] : "sin asignar"}
                        </span>
                        {h === col ? (
                          <span className="text-green-700 font-semibold shrink-0">· llave de cruce</span>
                        ) : null}
                      </li>
                    ))}
                  </ul>
                </div>
              );
            })}
            {principal && principalUuidCol ? (
              <div>
                <h3 className="text-[11px] font-bold tracking-wide text-gray-500 uppercase">
                  Equivalencias de cruce
                </h3>
                <ul className="mt-2 space-y-1.5 text-[13px] font-mono">
                  {complementarios.map((c) => {
                    const col = uuidColumnOf(c);
                    return (
                      <li key={c.id} className={col ? "text-green-700" : "text-amber-700"}>
                        {principal.fileName} [{principalUuidCol}] = {c.fileName} [{col ?? "sin UUID"}]
                      </li>
                    );
                  })}
                </ul>
              </div>
            ) : null}
          </div>
        ) : null}

        <div className="flex flex-wrap items-center justify-between gap-4 bg-white border border-gray-200 rounded-xl p-6">
          <div>
            <h2 className="text-sm font-bold text-gray-900">Ejecutar Conciliación</h2>
            <p className="text-[12px] text-gray-500 mt-0.5">
              Se habilita en la Fase 3: cruce de UUID contra el histórico de Supabase y los complementarios.
            </p>
          </div>
          <button
            disabled
            className="rounded-lg bg-gray-200 text-gray-500 text-sm font-semibold px-5 py-2.5 cursor-not-allowed"
          >
            Ejecutar Conciliación · Fase 3
          </button>
        </div>

        <p className="text-[12px] text-gray-500">
          Nota: en esta fase todo el procesamiento ocurre en tu navegador; todavía no se escribe nada en Supabase.
        </p>
      </div>
    </main>
  );
}