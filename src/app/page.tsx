"use client";

import { useEffect, useMemo, useState } from "react";
import Link from "next/link";
import * as XLSX from "xlsx";
import { createClient } from "@/lib/supabase/client";
import LogoutButton from "@/components/logout-button";
import Dropzone from "@/components/dropzone";
import FileCard from "@/components/file-card";
import ResultsTable from "@/components/results-table";
import type { FileRole, NormalizedRow, ParsedFile, SemanticField, UuidMode } from "@/lib/excel/parser";
import { buildAutoMapping, buildAutoVisible, detectHeaderRow, getHeaders, normalizeRows, readWorkbook, uuidColumnOf } from "@/lib/excel/parser";
import { buildFileDraft, rememberFile } from "@/lib/excel/memory";
import type { ComplementaryHit, HistInfo, ResultRow } from "@/lib/conciliacion";
import { buildComplementIndex, fetchHistoricalMap, reconcile } from "@/lib/conciliacion";
import type { RuleRow } from "@/lib/reglas";
import { classifyRows, ruleDescription } from "@/lib/reglas";

function preFromHeaders(headers: string[], mapping: Record<string, SemanticField>) {
  const pre: Record<string, boolean> = {};
  for (const h of headers) pre[h] = mapping[h] !== "NINGUNO";
  return pre;
}

export default function DashboardPage() {
  const supabase = useMemo(() => createClient(), []);
  const [email, setEmail] = useState("");
  const [userId, setUserId] = useState("");
  const [role, setRole] = useState("");
  const [files, setFiles] = useState<ParsedFile[]>([]);
  const [preById, setPreById] = useState<Record<string, Record<string, boolean>>>({});
  const [reading, setReading] = useState(false);
  const [running, setRunning] = useState(false);
  const [runError, setRunError] = useState("");
  const [results, setResults] = useState<ResultRow[] | null>(null);
  const [histMap, setHistMap] = useState<Map<string, HistInfo> | null>(null);
  const [compIdx, setCompIdx] = useState<Map<string, ComplementaryHit[]> | null>(null);
  const [showDict, setShowDict] = useState(false);
  const [showFieldPanel, setShowFieldPanel] = useState(true);
  const [fieldSel, setFieldSel] = useState<Record<string, boolean>>({});
  const [reportBusy, setReportBusy] = useState(false);
  const [reportMsg, setReportMsg] = useState("");

  const [uma, setUma] = useState(117.31);
  const [rules, setRules] = useState<RuleRow[]>([]);

  const [consolidating, setConsolidating] = useState(false);
  const [consolidateMsg, setConsolidateMsg] = useState("");

  useEffect(() => {
    supabase.auth.getUser().then(async ({ data }) => {
      setEmail(data.user?.email ?? "");
      setUserId(data.user?.id ?? "");
      if (data.user) {
        const { data: prof } = await supabase
          .from("profiles")
          .select("role")
          .eq("id", data.user.id)
          .maybeSingle();
        setRole(prof ? (prof as { role: string }).role : "");
      }
    });
  }, [supabase]);

  useEffect(() => {
    (async () => {
      const { data: s } = await supabase.from("settings").select("uma_valor").eq("id", 1).maybeSingle();
      if (s) setUma(Number((s as { uma_valor: number }).uma_valor));
      const { data: r } = await supabase
        .from("rules")
        .select("id, nombre, tipo_condicion, valor_limite, etiqueta, palabras_prohibidas, campo_nombre")
        .order("created_at", { ascending: true });
      if (r) setRules(r as RuleRow[]);
    })();
  }, [supabase]);

  useEffect(() => {
    for (const f of files) rememberFile(f.fileName, f);
  }, [files]);

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
  const principalReady = principal !== null && principalRows > 0;

  const principalHeaders = useMemo(() => {
    if (!principal) return [] as string[];
    const sheet = principal.sheets.find((s) => s.name === principal.sheetName);
    if (!sheet) return [] as string[];
    return getHeaders(sheet.rows, principal.headerRow);
  }, [principal]);

  // Del principal SOLO Concepto y Total (UUID va siempre como llave)
  const principalConceptoCol = useMemo(() => {
    if (!principal) return null;
    const mapped =
      Object.keys(principal.mapping).find((h) => principal.mapping[h] === "CONCEPTO") ?? null;
    if (mapped) return mapped;
    return principalHeaders.find((h) => h.toLowerCase().includes("concepto")) ?? null;
  }, [principal, principalHeaders]);

  const principalTotalCol = useMemo(() => {
    if (!principal) return null;
    const mapped =
      Object.keys(principal.mapping).find((h) => principal.mapping[h] === "MONTO") ?? null;
    if (mapped) return mapped;
    return principalHeaders.find((h) => h.toLowerCase().includes("total")) ?? null;
  }, [principal, principalHeaders]);

  const principalCols = useMemo(() => {
    const out: string[] = [];
    if (principalConceptoCol) out.push(principalConceptoCol);
    if (principalTotalCol) out.push(principalTotalCol);
    return out;
  }, [principalConceptoCol, principalTotalCol]);

  // TODAS las columnas de los complementarios (excepto su columna UUID)
  const compAllCols = useMemo(() => {
    const out: string[] = [];
    const seen = new Set<string>();
    for (const f of complementarios) {
      const sheet = f.sheets.find((s) => s.name === f.sheetName);
      if (!sheet) continue;
      for (const h of getHeaders(sheet.rows, f.headerRow)) {
        if (f.mapping[h] === "UUID") continue;
        const low = h.toLowerCase();
        if (!seen.has(low)) {
          seen.add(low);
          out.push(h);
        }
      }
    }
    return out;
  }, [complementarios]);

  useEffect(() => {
    setFieldSel((prev) => {
      const next = { ...prev };
      for (const c of compAllCols) {
        if (!(c in next)) next[c] = true;
      }
      return next;
    });
  }, [compAllCols]);

  const compColsChosen = useMemo(
    () => compAllCols.filter((c) => fieldSel[c]),
    [compAllCols, fieldSel]
  );

  const conceptHeaders = useMemo(
    () =>
      principalHeaders.filter(
        (h) => principal !== null && (principal.mapping[h] === "CONCEPTO" || principal.mapping[h] === "UUID")
      ),
    [principalHeaders, principal]
  );

  const complementHeaders = useMemo(
    () => principalHeaders.filter((h) => h.toLowerCase().includes("complemento")),
    [principalHeaders]
  );

  const montoCol = useMemo(() => {
    if (!principal) return null;
    return Object.keys(principal.mapping).find((h) => principal.mapping[h] === "MONTO") ?? null;
  }, [principal]);
  const retCol = useMemo(() => {
    if (!principal) return null;
    return Object.keys(principal.mapping).find((h) => principal.mapping[h] === "RETENCIONES") ?? null;
  }, [principal]);
  const folioCol = useMemo(() => {
    if (!principal) return null;
    return Object.keys(principal.mapping).find((h) => principal.mapping[h] === "FOLIO") ?? null;
  }, [principal]);

  const groups = useMemo(() => {
    if (!results) return null;
    return classifyRows(results, rules, uma, (r) => ({
      monto:
        montoCol !== null && typeof r.values[montoCol] === "number"
          ? (r.values[montoCol] as number)
          : null,
      retenciones:
        retCol !== null && typeof r.values[retCol] === "number"
          ? (r.values[retCol] as number)
          : null,
      concepto: conceptHeaders.map((h) => String(r.values[h] ?? "")).join(" ").trim(),
      complemento: complementHeaders.map((h) => String(r.values[h] ?? "")).join(" ").trim(),
      values: r.values,
    }));
  }, [results, rules, uma, montoCol, retCol, conceptHeaders, complementHeaders]);

  const segmentInfo = useMemo(() => {
    const labelMap: Record<string, string[]> = {};
    const keysMap: Record<string, string[]> = {};
    const options: { key: string; label: string; count: number }[] = [];
    if (!groups) return { labelMap, keysMap, options };
    for (const g of groups) {
      options.push({
        key: g.key,
        label: g.rule ? g.rule.etiqueta : "Sin regla aplicada",
        count: g.rows.length,
      });
      if (!g.rule) continue;
      for (const r of g.rows) {
        const k = r.uuid + "|" + r.sourceRow;
        if (!labelMap[k]) labelMap[k] = [];
        if (!keysMap[k]) keysMap[k] = [];
        labelMap[k].push(g.rule.etiqueta);
        keysMap[k].push(g.key);
      }
    }
    return { labelMap, keysMap, options };
  }, [groups]);

  const stats = useMemo(() => {
    if (!results) return null;
    const encontradas = results.filter((r) => r.status !== "NO_ENCONTRADA").length;
    const senaladas = results.filter(
      (r) => r.aparicionesPrincipal > 1 || r.extraLocations.length > 0
    ).length;
    return {
      evaluadas: results.length,
      encontradas,
      faltantes: results.length - encontradas,
      senaladas,
    };
  }, [results]);

  async function addFiles(list: File[], role: FileRole) {
    setReading(true);
    const added: ParsedFile[] = [];
    for (const file of list) {
      try {
        const buffer = await file.arrayBuffer();
        const sheets = readWorkbook(buffer);
        if (sheets.length === 0) continue;
        added.push(
          buildFileDraft(file.name, role, sheets, /mayor/i.test(file.name) ? "TEXTO" : "EXACTO")
        );
      } catch {
        window.alert("No se pudo leer el archivo: " + file.name);
      }
    }
    setPreById((prev) => {
      const next = { ...prev };
      for (const d of added) {
        const sheet = d.sheets.find((s) => s.name === d.sheetName);
        const headers = sheet ? getHeaders(sheet.rows, d.headerRow) : [];
        next[d.id] = preFromHeaders(headers, d.mapping);
      }
      return next;
    });
    setFiles((prev) =>
      role === "PRINCIPAL"
        ? [...prev.filter((f) => f.role !== "PRINCIPAL"), ...added]
        : [...prev, ...added]
    );
    setResults(null);
    setReading(false);
  }

  function updateFile(id: string, updater: (f: ParsedFile) => ParsedFile, clear: boolean) {
    setFiles((prev) => prev.map((f) => (f.id === id ? updater(f) : f)));
    if (clear) setResults(null);
  }

  function handleSheetChange(id: string, sheetName: string) {
    const f = files.find((x) => x.id === id);
    if (!f) return;
    const sheet = f.sheets.find((s) => s.name === sheetName);
    if (!sheet) return;
    const headerRow = detectHeaderRow(sheet.rows);
    const headers = getHeaders(sheet.rows, headerRow);
    const mapping = buildAutoMapping(headers);
    const visible = buildAutoVisible(headers, mapping);
    setFiles((prev) => prev.map((x) => (x.id === id ? { ...x, sheetName, headerRow, mapping, visible } : x)));
    setPreById((prev) => ({ ...prev, [id]: preFromHeaders(headers, mapping) }));
    setResults(null);
  }

  function handleHeaderRowChange(id: string, headerRow: number) {
    const f = files.find((x) => x.id === id);
    if (!f) return;
    const sheet = f.sheets.find((s) => s.name === f.sheetName);
    if (!sheet) return;
    const headers = getHeaders(sheet.rows, headerRow);
    const mapping = buildAutoMapping(headers);
    const visible = buildAutoVisible(headers, mapping);
    setFiles((prev) => prev.map((x) => (x.id === id ? { ...x, headerRow, mapping, visible } : x)));
    setPreById((prev) => ({ ...prev, [id]: preFromHeaders(headers, mapping) }));
    setResults(null);
  }

  function handleMappingChange(id: string, header: string, field: SemanticField) {
    const f = files.find((x) => x.id === id);
    const wasUuid = f ? f.mapping[header] === "UUID" : false;
    const pre = preById[id] ?? {};
    const clear = wasUuid || field === "UUID" || pre[header] !== true;
    updateFile(
      id,
      (cur) => {
        const mapping = { ...cur.mapping };
        if (field !== "NINGUNO") {
          for (const key of Object.keys(mapping)) {
            if (mapping[key] === field) mapping[key] = "NINGUNO";
          }
        }
        mapping[header] = field;
        const visible = { ...cur.visible };
        if (field !== "NINGUNO") visible[header] = true;
        return { ...cur, mapping, visible };
      },
      clear
    );
  }

  function handleUuidModeChange(id: string, mode: UuidMode) {
    updateFile(id, (f) => ({ ...f, uuidMode: mode }), true);
  }

  function removeFile(id: string) {
    setFiles((prev) => prev.filter((f) => f.id !== id));
    setResults(null);
  }

  async function runConciliacion() {
    if (!principal) return;
    setRunning(true);
    setRunError("");
    setResults(null);
    try {
      const principalRowsList = rowsById[principal.id] ?? [];
      const uniqueUuids = Array.from(new Set(principalRowsList.map((r) => r.uuid)));
      const historical = await fetchHistoricalMap(supabase, uniqueUuids);
      const compIndex = buildComplementIndex(
        complementarios.map((f) => ({ file: f, rows: rowsById[f.id] ?? [] }))
      );
      setHistMap(historical);
      setCompIdx(compIndex);
      setResults(reconcile(principalRowsList, historical, compIndex));
    } catch (err) {
      setRunError(err instanceof Error ? err.message : "Error inesperado al conciliar");
    } finally {
      setRunning(false);
    }
  }

  async function generarReporteFinal() {
    if (!principal || !results || !histMap || !compIdx) return;
    setReportBusy(true);
    setReportMsg("");
    try {
      const principalRowsList = rowsById[principal.id] ?? [];
      const seen = new Set<string>();
      const rowsOut: Record<string, string | number>[] = [];
      for (const r of principalRowsList) {
        if (seen.has(r.uuid)) continue;
        seen.add(r.uuid);
        let estatus = "";
        let ubicacion = "";
        let observacion = "";

        const h = histMap.get(r.uuid);
        const hits = compIdx.get(r.uuid);
        const hit = hits && hits.length > 0 ? hits[0] : null;

        if (h) {
          estatus = "Encontrada en histórico";
          ubicacion = "Histórico acumulado";
          observacion =
            "Ya registrada en el histórico acumulado" +
            (h.mes ? " (periodo " + h.mes + "/" + (h.anio ?? "") + ")" : "") +
            (h.origen ? "; origen: " + h.origen : "") +
            ".";
        } else if (hits && hits.length > 0 && hit) {
          estatus = "Encontrada en archivo del mes";
          ubicacion = hit.fileName + " · " + hit.sheetName + " · fila " + hit.sourceRow;
          observacion =
            "Ubicada en el archivo del mes: " + hit.fileName + " (hoja " + hit.sheetName + ", fila " + hit.sourceRow + ")." +
            (hit.detalle ? " Datos: " + hit.detalle + "." : "") +
            (hits.length > 1 ? " También aparece en otros " + (hits.length - 1) + " archivo(s)." : "");
        } else {
          estatus = "No encontrada";
          ubicacion = "";
          observacion = "No ubicada este mes; queda pendiente para el siguiente cierre.";
        }

        const base: Record<string, string | number> = {
          UUID: r.uuid,
          Estatus: estatus,
          Ubicacion: ubicacion,
          Observacion: observacion.trim(),
        };
        for (const pc of principalCols) {
          const v = r.values[pc] ?? null;
          base[pc] = v === null || v === undefined ? "" : v;
        }
        for (const cc of compColsChosen) {
          const v = hit ? hit.values[cc] ?? null : null;
          base["Fuente · " + cc] = v === null || v === undefined ? "" : v;
        }
        rowsOut.push(base);
      }

      const ws = XLSX.utils.json_to_sheet(rowsOut);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Reporte final");
      XLSX.writeFile(wb, "reporte_final_facturas.xlsx");
      setReportMsg("Reporte final generado con " + rowsOut.length + " facturas y sus observaciones.");
    } catch (err) {
      setReportMsg("Error al generar el reporte: " + (err instanceof Error ? err.message : "error inesperado"));
    }
    setReportBusy(false);
  }

  async function consolidarMes() {
    if (!principal || !results) return;
    setConsolidating(true);
    setConsolidateMsg("");
    try {
      const principalRowsList = rowsById[principal.id] ?? [];
      const uniqueUuids = Array.from(new Set(principalRowsList.map((r) => r.uuid)));
      const historical = await fetchHistoricalMap(supabase, uniqueUuids);
      const compIndex = buildComplementIndex(
        complementarios.map((f) => ({ file: f, rows: rowsById[f.id] ?? [] }))
      );

      const now = new Date();
      const seen = new Set<string>();
      const toInsert: Record<string, unknown>[] = [];
      let alreadyHistorical = 0;
      let leftPending = 0;
      for (const r of principalRowsList) {
        if (seen.has(r.uuid)) continue;
        seen.add(r.uuid);
        if (historical.has(r.uuid)) {
          alreadyHistorical += 1;
          continue;
        }
        const hits = compIndex.get(r.uuid);
        if (!hits || hits.length === 0) {
          leftPending += 1;
          continue;
        }
        const datos: Record<string, unknown> = {
          ...r.values,
          __origen: principal.fileName,
          __hoja: principal.sheetName,
          __fila: r.sourceRow,
          __via: "COMPLEMENTARIO",
          __fuentes: hits.map((h) => h.fileName + " · " + h.sheetName + " · fila " + h.sourceRow),
        };
        for (const h of hits) {
          for (const [k, v] of Object.entries(h.values)) {
            datos["[" + h.fileName + "] " + k] = v;
          }
        }
        toInsert.push({
          uuid_fiscal: r.uuid,
          folio: folioCol !== null && typeof r.values[folioCol] === "string" ? (r.values[folioCol] as string) : null,
          monto: montoCol !== null && typeof r.values[montoCol] === "number" ? (r.values[montoCol] as number) : null,
          retenciones: retCol !== null && typeof r.values[retCol] === "number" ? (r.values[retCol] as number) : null,
          mes_periodo: now.getMonth() + 1,
          anio_periodo: now.getFullYear(),
          origen_archivo: principal.fileName,
          datos_json: datos,
        });
      }

      if (toInsert.length > 0) {
        const { error: insErr } = await supabase
          .from("historical_invoices")
          .upsert(toInsert, { onConflict: "uuid_fiscal", ignoreDuplicates: true });
        if (insErr) {
          window.alert("Error al consolidar: " + insErr.message);
          return;
        }
      }

      setConsolidateMsg(
        "Consolidadas " + toInsert.length + " facturas con todos sus campos (JOIN de archivos). " +
        "Ya estaban en histórico: " + alreadyHistorical + " (omitidas). " +
        "Quedan pendientes para el siguiente mes: " + leftPending + "."
      );
    } catch (err) {
      window.alert(err instanceof Error ? err.message : "Error al consolidar");
    } finally {
      setConsolidating(false);
    }
  }

  const principalUuidCol = principal ? uuidColumnOf(principal) : null;

  return (
    <main className="min-h-screen bg-gray-100 text-gray-900">
      <header className="bg-white border-b border-gray-200 sticky top-0 z-10">
        <div className="max-w-6xl mx-auto px-6 py-4 flex items-center justify-between gap-4">
          <div>
            <h1 className="font-bold text-gray-900">Conciliador de Facturas</h1>
            <p className="text-[11px] text-gray-500">Panel del Contador Principal</p>
          </div>
          <div className="flex items-center gap-3">
            <nav className="flex items-center gap-1 text-[12px] font-semibold">
              <span className="rounded-lg bg-blue-50 text-blue-700 px-3 py-1.5">Panel</span>
              <Link href="/reglas" className="rounded-lg px-3 py-1.5 text-gray-600 hover:bg-gray-100">Reglas y UMA</Link>
              {role === "SUPER_USUARIO" ? (
                <Link href="/admin" className="rounded-lg px-3 py-1.5 text-gray-600 hover:bg-gray-100">Mantenimiento</Link>
              ) : null}
            </nav>
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
          <p className="text-sm text-gray-600 bg-white border border-gray-200 rounded-xl px-4 py-3">Leyendo archivos…</p>
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
          <div className="bg-white border border-gray-200 rounded-xl overflow-hidden">
            <button
              onClick={() => setShowDict(!showDict)}
              className="w-full flex items-center justify-between px-6 py-4 bg-white hover:bg-gray-50 transition-colors"
            >
              <span className="text-left">
                <span className="block text-sm font-bold text-gray-900">Diccionario de datos del mes</span>
                <span className="block text-[12px] text-gray-500 mt-0.5">
                  Todos estos campos se conservarán en el histórico acumulado al consolidar el mes.
                </span>
              </span>
              <span className="text-[11px] font-semibold text-gray-500 shrink-0">
                {showDict ? "Colapsar ▲" : "Expandir ▼"}
              </span>
            </button>
            {showDict ? (
              <div className="px-6 pb-6 space-y-4 border-t border-gray-100">
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
                        {headers.map((h) => {
                          const inReport =
                            f.role === "PRINCIPAL"
                              ? h === col || principalCols.includes(h)
                              : fieldSel[h] ?? false;
                          return (
                            <li key={h} className="flex items-center gap-2 min-w-0">
                              {inReport ? (
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
                          );
                        })}
                      </ul>
                    </div>
                  );
                })}
                {principal && principalUuidCol ? (
                  <div>
                    <h3 className="text-[11px] font-bold tracking-wide text-gray-500 uppercase">Equivalencias de cruce</h3>
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
          </div>
        ) : null}

        <div className="bg-white border border-gray-200 rounded-xl p-6 space-y-4">
          <div className="flex flex-wrap items-center justify-between gap-4">
            <div>
              <h2 className="text-sm font-bold text-gray-900">Ejecutar Conciliación</h2>
              <p className="text-[12px] text-gray-500 mt-0.5">
                Orden de búsqueda: 1) histórico acumulado · 2) archivos del mes.
              </p>
            </div>
            <button
              onClick={runConciliacion}
              disabled={!principalReady || reading || running}
              className={
                "rounded-lg text-sm font-semibold px-5 py-2.5 transition-colors " +
                (principalReady && !running
                  ? "bg-blue-700 text-white hover:bg-blue-800"
                  : "bg-gray-200 text-gray-500 cursor-not-allowed")
              }
            >
              {running ? "Conciliando…" : "Ejecutar Conciliación"}
            </button>
          </div>

          {!principalReady ? (
            <p className="text-[12px] text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">
              Para conciliar necesitas un archivo principal LISTO (con su columna UUID marcada).
            </p>
          ) : null}

          {runError ? (
            <p className="text-[12px] text-red-700 bg-red-50 border border-red-200 rounded-lg px-3 py-2">{runError}</p>
          ) : null}

          {stats ? (
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
              <div className="rounded-xl border border-gray-200 p-4">
                <p className="text-[11px] font-bold text-gray-500 uppercase tracking-wide">Evaluadas</p>
                <p className="text-2xl font-bold text-gray-900 mt-1">{stats.evaluadas}</p>
              </div>
              <div className="rounded-xl border border-green-200 bg-green-50 p-4">
                <p className="text-[11px] font-bold text-green-700 uppercase tracking-wide">Encontradas</p>
                <p className="text-2xl font-bold text-green-700 mt-1">{stats.encontradas}</p>
              </div>
              <div className="rounded-xl border border-red-200 bg-red-50 p-4">
                <p className="text-[11px] font-bold text-red-700 uppercase tracking-wide">No encontradas</p>
                <p className="text-2xl font-bold text-red-700 mt-1">{stats.faltantes}</p>
              </div>
              <div className="rounded-xl border border-amber-200 bg-amber-50 p-4">
                <p className="text-[11px] font-bold text-amber-700 uppercase tracking-wide">Señaladas</p>
                <p className="text-2xl font-bold text-amber-700 mt-1">{stats.senaladas}</p>
              </div>
            </div>
          ) : null}

          {results ? (
            <div className="border border-gray-200 rounded-lg bg-gray-50">
              <button
                onClick={() => setShowFieldPanel(!showFieldPanel)}
                className="w-full flex items-center justify-between px-4 py-3 hover:bg-gray-100 transition-colors"
              >
                <span className="text-[12px] font-bold text-gray-700">
                  Campos de los complementarios a mostrar · {compColsChosen.length} de {compAllCols.length}
                </span>
                <span className="text-[11px] font-semibold text-gray-500">
                  {showFieldPanel ? "Colapsar ▲" : "Expandir ▼"}
                </span>
              </button>
              {showFieldPanel ? (
                <div className="px-4 pb-4">
                  <p className="text-[11px] text-gray-500 mb-2">
                    Del principal se muestran fijos: UUID{principalConceptoCol ? ", " + principalConceptoCol : ""}{principalTotalCol ? ", " + principalTotalCol : ""}. Marca los campos de complementarios que quieres ver y exportar.
                  </p>
                  <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-2">
                    {compAllCols.map((c) => (
                      <label
                        key={c}
                        className="flex items-center gap-2 text-[12px] font-semibold text-gray-700 bg-white border border-gray-200 rounded-lg px-3 py-2 cursor-pointer hover:border-blue-400"
                      >
                        <input
                          type="checkbox"
                          checked={fieldSel[c] ?? false}
                          onChange={(e) => setFieldSel((p) => ({ ...p, [c]: e.target.checked }))}
                          className="accent-blue-700 w-3.5 h-3.5"
                        />
                        <span className="truncate">[{c}]</span>
                      </label>
                    ))}
                    {compAllCols.length === 0 ? (
                      <p className="text-[12px] text-gray-500">Carga complementarios para ver sus campos.</p>
                    ) : null}
                  </div>
                </div>
              ) : null}
            </div>
          ) : null}

          {results ? (
            <ResultsTable
              results={results}
              segmentMap={segmentInfo.labelMap}
              groupKeysMap={segmentInfo.keysMap}
              segmentOptions={segmentInfo.options}
              compIndex={compIdx}
              principalCols={principalCols}
              compCols={compColsChosen}
            />
          ) : null}

          {results ? (
            <div className="flex flex-wrap items-center justify-between gap-3 border border-gray-200 rounded-lg px-4 py-3 bg-gray-50">
              <p className="text-[12px] text-gray-600 max-w-xl">
                <strong>Reporte final de facturas:</strong> un .xlsx con estatus, ubicación, observación automática, Concepto y Total del principal, y los campos de complementarios marcados.
              </p>
              <button
                onClick={generarReporteFinal}
                disabled={reportBusy}
                className="rounded-lg bg-green-600 text-white text-[12px] font-semibold px-4 py-2 hover:bg-green-700 disabled:opacity-50"
              >
                {reportBusy ? "Generando…" : "Reporte final (.xlsx)"}
              </button>
            </div>
          ) : null}
          {reportMsg ? (
            <p className="text-[12px] text-green-700 bg-green-50 border border-green-200 rounded-lg px-3 py-2">{reportMsg}</p>
          ) : null}
        </div>

        <div className="bg-white border border-gray-200 rounded-xl p-6 space-y-3">
          <h2 className="text-sm font-bold text-gray-900">Cierre mensual</h2>
          <p className="text-[12px] text-gray-600 leading-relaxed">
            <strong>Consolidar Mes</strong> envía al histórico acumulado las facturas encontradas en los archivos del mes,
            con un <strong>consolidado tipo JOIN</strong>: cada UUID guarda sus campos del principal más todas las columnas de
            cada complementario donde aparece (prefijadas por archivo) y la lista de fuentes. Las que ya estaban en el
            histórico se omiten y las no encontradas quedan para el siguiente mes. Nunca se duplica: cada UUID se consolida una sola vez.
          </p>
          <div className="flex flex-wrap items-center gap-3">
            <button
              onClick={consolidarMes}
              disabled={!results || consolidating}
              className={
                "rounded-lg text-sm font-semibold px-5 py-2.5 transition-colors " +
                (results && !consolidating
                  ? "bg-green-600 text-white hover:bg-green-700"
                  : "bg-gray-200 text-gray-500 cursor-not-allowed")
              }
            >
              {consolidating ? "Consolidando…" : "Consolidar Mes"}
            </button>
            {!results ? (
              <span className="text-[12px] text-gray-500">Primero ejecuta la conciliación del mes.</span>
            ) : null}
          </div>
          {consolidateMsg ? (
            <p className="text-[12px] text-green-700 bg-green-50 border border-green-200 rounded-lg px-3 py-2">{consolidateMsg}</p>
          ) : null}
        </div>
      </div>
    </main>
  );
}