"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import Link from "next/link";
import { useRouter } from "next/navigation";
import { createClient } from "@/lib/supabase/client";
import LogoutButton from "@/components/logout-button";
import Dropzone from "@/components/dropzone";
import FileCard from "@/components/file-card";
import ResultsTable from "@/components/results-table";
import InvoiceModal from "@/components/invoice-modal";
import type { LocalInvoiceInfo } from "@/components/invoice-modal";
import type { FileRole, NormalizedRow, ParsedFile, SemanticField, UuidMode } from "@/lib/excel/parser";
import { buildAutoMapping, buildAutoVisible, detectHeaderRow, getHeaders, normalizeRows, readWorkbook, uuidColumnOf } from "@/lib/excel/parser";
import { buildFileDraft, rememberFile } from "@/lib/excel/memory";
import type { ResultRow } from "@/lib/conciliacion";
import { buildComplementIndex, fetchHistoricalUuids, reconcile } from "@/lib/conciliacion";
import type { RuleRow } from "@/lib/reglas";
import { classifyRows, ruleDescription } from "@/lib/reglas";
import type { BatchRow } from "@/lib/batches";
import { BATCH_STATUS_LABEL, ITEM_STATUSES, addEvent, fetchBatches } from "@/lib/batches";

interface Colab {
  id: string;
  email: string;
}

function preFromHeaders(headers: string[], mapping: Record<string, SemanticField>) {
  const pre: Record<string, boolean> = {};
  for (const h of headers) pre[h] = mapping[h] !== "NINGUNO";
  return pre;
}

function batchBadge(status: string) {
  if (status === "COMPLETADO")
    return <span className="text-[10px] font-bold rounded-full px-2.5 py-1 bg-green-100 text-green-700">COMPLETADO</span>;
  if (status === "EN_PROCESO")
    return <span className="text-[10px] font-bold rounded-full px-2.5 py-1 bg-blue-100 text-blue-700">EN PROCESO</span>;
  return <span className="text-[10px] font-bold rounded-full px-2.5 py-1 bg-amber-100 text-amber-700">PENDIENTE</span>;
}

export default function DashboardPage() {
  const supabase = useMemo(() => createClient(), []);
  const router = useRouter();
  const [email, setEmail] = useState("");
  const [userId, setUserId] = useState("");
  const [role, setRole] = useState("");
  const [files, setFiles] = useState<ParsedFile[]>([]);
  const [preById, setPreById] = useState<Record<string, Record<string, boolean>>>({});
  const [reading, setReading] = useState(false);
  const [running, setRunning] = useState(false);
  const [runError, setRunError] = useState("");
  const [results, setResults] = useState<ResultRow[] | null>(null);
  const [showCols, setShowCols] = useState(false);
  const [showDict, setShowDict] = useState(false);
  const [showSeg, setShowSeg] = useState(false);
  const [modalUuid, setModalUuid] = useState<string | null>(null);

  const [uma, setUma] = useState(117.31);
  const [rules, setRules] = useState<RuleRow[]>([]);
  const [colabs, setColabs] = useState<Colab[]>([]);
  const [assignSel, setAssignSel] = useState<Record<string, string>>({});
  const [batchedBy, setBatchedBy] = useState<Record<string, string>>({});
  const [assigning, setAssigning] = useState("");

  const [batchesAdmin, setBatchesAdmin] = useState<BatchRow[] | null>(null);
  const [showLogAdmin, setShowLogAdmin] = useState<Record<string, boolean>>({});
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
        const r = prof ? (prof as { role: string }).role : "";
        setRole(r);
        if (r === "COLABORADOR_CONTADOR") router.replace("/mis-batches");
      }
    });
  }, [supabase, router]);

  const refreshBatches = useCallback(async () => {
    try {
      setBatchesAdmin(await fetchBatches(supabase, null));
    } catch {
      setBatchesAdmin([]);
    }
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
      const { data: p } = await supabase.from("profiles").select("id, email, role");
      if (p) {
        setColabs(
          (p as { id: string; email: string; role: string }[])
            .filter((x) => x.role === "COLABORADOR_CONTADOR")
            .map((x) => ({ id: x.id, email: x.email }))
        );
      }
      await refreshBatches();
    })();
  }, [supabase, refreshBatches]);

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

  const localByUuid = useMemo(() => {
    const map: Record<string, LocalInvoiceInfo> = {};
    const pr = files.find((f) => f.role === "PRINCIPAL") ?? null;
    const comps = files.filter((f) => f.role === "COMPLEMENTARIO");
    const ordered = pr !== null ? [pr, ...comps] : comps;
    for (const f of ordered) {
      for (const r of rowsById[f.id] ?? []) {
        if (!map[r.uuid]) {
          map[r.uuid] = {
            fileName: f.fileName,
            sheetName: f.sheetName,
            sourceRow: r.sourceRow,
            values: r.values,
          };
        }
      }
    }
    return map;
  }, [files, rowsById]);
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

  const principalVisibleCols = useMemo(
    () => principalHeaders.filter((h) => principal !== null && principal.visible[h] && principal.mapping[h] !== "UUID"),
    [principalHeaders, principal]
  );

  const principalAllCols = useMemo(
    () => principalHeaders.filter((h) => principal !== null && principal.mapping[h] !== "UUID"),
    [principalHeaders, principal]
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

  function handleVisibleChange(id: string, header: string, visible: boolean) {
    const pre = preById[id] ?? {};
    const clear = pre[header] !== true;
    updateFile(id, (f) => ({ ...f, visible: { ...f.visible, [header]: visible } }), clear);
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
      const historical = await fetchHistoricalUuids(supabase, uniqueUuids);
      const compIndex = buildComplementIndex(
        complementarios.map((f) => ({ file: f, rows: rowsById[f.id] ?? [] }))
      );
      setResults(reconcile(principalRowsList, historical, compIndex));
    } catch (err) {
      setRunError(err instanceof Error ? err.message : "Error inesperado al conciliar");
    } finally {
      setRunning(false);
    }
  }

  async function assignBatch(groupKey: string, titulo: string, rows: ResultRow[]) {
    const colabId = assignSel[groupKey];
    if (!colabId) {
      window.alert("Selecciona un colaborador");
      return;
    }
    setAssigning(groupKey);
    try {
      const { data: batch, error } = await supabase
        .from("batches")
        .insert({ titulo, assigned_to: colabId, created_by: userId, status: "PENDIENTE" })
        .select()
        .single();
      if (error || !batch) {
        window.alert(error ? error.message : "No se pudo crear el batch");
        return;
      }
      const batchId = (batch as { id: string }).id;
      const items = rows.map((r) => ({ batch_id: batchId, uuid_fiscal: r.uuid }));
      const { error: errItems } = await supabase.from("batch_items").insert(items);
      if (errItems) {
        window.alert(errItems.message);
        return;
      }
      await addEvent(supabase, batchId, null, email, "ASIGNACION", "PENDIENTE", "Batch creado y asignado desde el panel");
      const colab = colabs.find((c) => c.id === colabId);
      setBatchedBy((prev) => {
        const next = { ...prev };
        next[groupKey] = colab ? colab.email : "colaborador";
        return next;
      });
      await refreshBatches();
    } finally {
      setAssigning("");
    }
  }

  async function consolidarMes() {
    if (!principal || !results) return;
    setConsolidating(true);
    setConsolidateMsg("");
    try {
      const principalRowsList = rowsById[principal.id] ?? [];
      const uniqueUuids = Array.from(new Set(principalRowsList.map((r) => r.uuid)));
      const historical = await fetchHistoricalUuids(supabase, uniqueUuids);
      const compIndex = buildComplementIndex(
        complementarios.map((f) => ({ file: f, rows: rowsById[f.id] ?? [] }))
      );

      const notFound = uniqueUuids.filter((u) => !historical.has(u) && !compIndex.has(u));
      const subidaSet = new Set<string>();
      if (notFound.length > 0) {
        const { data: items } = await supabase
          .from("batch_items")
          .select("uuid_fiscal")
          .in("uuid_fiscal", notFound)
          .eq("status", "SUBIDA_SISTEMA");
        for (const it of items ?? []) subidaSet.add((it as { uuid_fiscal: string }).uuid_fiscal);
      }

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
        const inComp = compIndex.has(r.uuid);
        const subida = subidaSet.has(r.uuid);
        if (!inComp && !subida) {
          leftPending += 1;
          continue;
        }
        toInsert.push({
          uuid_fiscal: r.uuid,
          folio: folioCol !== null && typeof r.values[folioCol] === "string" ? (r.values[folioCol] as string) : null,
          monto: montoCol !== null && typeof r.values[montoCol] === "number" ? (r.values[montoCol] as number) : null,
          retenciones: retCol !== null && typeof r.values[retCol] === "number" ? (r.values[retCol] as number) : null,
          mes_periodo: now.getMonth() + 1,
          anio_periodo: now.getFullYear(),
          origen_archivo: principal.fileName,
          datos_json: {
            ...r.values,
            __origen: principal.fileName,
            __hoja: principal.sheetName,
            __fila: r.sourceRow,
            __via: inComp ? "COMPLEMENTARIO" : "SUBIDA_SISTEMA",
          },
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

      const consolidatedSet = new Set(toInsert.map((t) => String(t.uuid_fiscal)));
      if (consolidatedSet.size > 0) {
        const { data: allItems } = await supabase.from("batch_items").select("batch_id, uuid_fiscal");
        const perBatch = new Map<string, number>();
        for (const it of (allItems ?? []) as { batch_id: string; uuid_fiscal: string }[]) {
          if (consolidatedSet.has(it.uuid_fiscal)) {
            perBatch.set(it.batch_id, (perBatch.get(it.batch_id) ?? 0) + 1);
          }
        }
        for (const [bid, count] of perBatch) {
          await addEvent(
            supabase,
            bid,
            null,
            email,
            "CONSOLIDACION",
            null,
            "Consolidar Mes: " + count + " factura(s) de este batch pasaron al histórico"
          );
        }
      }

      await refreshBatches();
      setConsolidateMsg(
        "Consolidadas " + toInsert.length + " facturas con todos sus campos. " +
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
            {role !== "COLABORADOR_CONTADOR" ? (
              <nav className="flex items-center gap-1 text-[12px] font-semibold">
                <span className="rounded-lg bg-blue-50 text-blue-700 px-3 py-1.5">Panel</span>
                <Link href="/reglas" className="rounded-lg px-3 py-1.5 text-gray-600 hover:bg-gray-100">Reglas y UMA</Link>
                <Link href="/mis-batches" className="rounded-lg px-3 py-1.5 text-gray-600 hover:bg-gray-100">Mis Batches</Link>
                {role === "SUPER_USUARIO" ? (
                  <Link href="/admin" className="rounded-lg px-3 py-1.5 text-gray-600 hover:bg-gray-100">Mantenimiento</Link>
                ) : null}
              </nav>
            ) : null}
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

          {results && principal ? (
            <div className="border border-gray-200 rounded-lg overflow-hidden">
              <button
                onClick={() => setShowCols(!showCols)}
                className="w-full flex items-center justify-between px-4 py-3 bg-gray-50 hover:bg-gray-100 transition-colors"
              >
                <span className="text-[12px] font-bold text-gray-700">
                  Columnas del reporte · {principalVisibleCols.length} visibles
                </span>
                <span className="text-[11px] font-semibold text-gray-500">
                  {showCols ? "Colapsar ▲" : "Expandir ▼"}
                </span>
              </button>
              {showCols ? (
                <div className="p-4 grid sm:grid-cols-2 lg:grid-cols-3 gap-2 bg-white">
                  {principalAllCols.map((h) => (
                    <label
                      key={h}
                      className="flex items-center gap-2 text-[12px] font-semibold text-gray-700 bg-gray-50 border border-gray-200 rounded-lg px-3 py-2 cursor-pointer hover:border-blue-400"
                    >
                      <input
                        type="checkbox"
                        checked={principal.visible[h] ?? false}
                        onChange={(e) => handleVisibleChange(principal.id, h, e.target.checked)}
                        className="accent-blue-700 w-3.5 h-3.5"
                      />
                      <span className="truncate">[{h}]</span>
                    </label>
                  ))}
                </div>
              ) : null}
            </div>
          ) : null}

          {results ? (
            <ResultsTable
              results={results}
              visibleCols={principalVisibleCols}
              segmentMap={segmentInfo.labelMap}
              groupKeysMap={segmentInfo.keysMap}
              segmentOptions={segmentInfo.options}
            />
          ) : null}
        </div>

        {groups ? (
          <div className="bg-white border border-gray-200 rounded-xl p-6 space-y-4">
            <div>
              <h2 className="text-sm font-bold text-gray-900">Clasificación por reglas y asignación de batches</h2>
              <p className="text-[12px] text-gray-500 mt-0.5">
                UMA actual: {"$"}{uma.toFixed(2)} · Cada regla segmenta su propio grupo; una factura puede aparecer en varios grupos y en la columna “Segmentación” del reporte.
              </p>
            </div>
            {groups.map((g) => (
              <div key={g.key} className="border border-gray-200 rounded-lg p-4">
                <div className="flex flex-wrap items-center justify-between gap-3">
                  <div className="min-w-0">
                    <p className="text-[13px] font-bold text-gray-900">
                      {g.rule ? g.rule.etiqueta : "Sin regla aplicada"}
                      <span className="ml-2 text-[11px] font-bold text-gray-500 bg-gray-100 rounded-full px-2 py-0.5">
                        {g.rows.length} facturas
                      </span>
                    </p>
                    <p className="text-[12px] text-gray-500 mt-0.5">
                      {g.rule ? ruleDescription(g.rule, uma) : "Facturas que no cumplieron ninguna regla."}
                    </p>
                  </div>
                  <div className="flex items-center gap-2">
                    {batchedBy[g.key] ? (
                      <span className="text-[11px] font-semibold bg-green-100 text-green-700 rounded-full px-3 py-1">
                        Batch asignado a {batchedBy[g.key]}
                      </span>
                    ) : (
                      <>
                        <select
                          value={assignSel[g.key] ?? ""}
                          onChange={(e) => setAssignSel((p) => ({ ...p, [g.key]: e.target.value }))}
                          disabled={g.rows.length === 0 || colabs.length === 0}
                          className="rounded-lg border border-gray-300 bg-white px-3 py-1.5 text-[12px] text-gray-900"
                        >
                          <option value="">Selecciona colaborador…</option>
                          {colabs.map((c) => (
                            <option key={c.id} value={c.id}>{c.email}</option>
                          ))}
                        </select>
                        <button
                          onClick={() =>
                            assignBatch(
                              g.key,
                              (g.rule ? g.rule.etiqueta : "Sin regla") + " · " + g.rows.length + " facturas",
                              g.rows
                            )
                          }
                          disabled={g.rows.length === 0 || !assignSel[g.key] || assigning === g.key}
                          className="rounded-lg bg-blue-700 text-white text-[12px] font-semibold px-3.5 py-1.5 hover:bg-blue-800 disabled:opacity-50"
                        >
                          {assigning === g.key ? "Asignando…" : "Crear batch y asignar"}
                        </button>
                      </>
                    )}
                  </div>
                </div>
              </div>
            ))}
            {colabs.length === 0 ? (
              <p className="text-[12px] text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">
                 No hay colaboradores registrados. Créalos en la consola de usuarios del sistema y asígnales el rol de colaborador.
              </p>
            ) : null}
          </div>
        ) : null}

        <div className="bg-white border border-gray-200 rounded-xl overflow-hidden">
          <button
            onClick={() => setShowSeg(!showSeg)}
            className="w-full flex items-center justify-between px-6 py-4 bg-white hover:bg-gray-50 transition-colors"
          >
            <span className="text-left">
              <span className="text-sm font-bold text-gray-900">Seguimiento de batches</span>
              <span className="block text-[12px] text-gray-500 mt-0.5">
                Estatus, contadores por factura y bitácora de todos los batches asignados.
              </span>
            </span>
            <span className="text-[11px] font-semibold text-gray-500 shrink-0">
              {showSeg ? "Colapsar ▲" : "Expandir ▼"}
            </span>
          </button>
          {showSeg ? (
            <div className="px-6 pb-6 space-y-4 border-t border-gray-100">
              {(batchesAdmin ?? []).length === 0 ? (
                <p className="text-[13px] text-gray-500">Aún no hay batches creados.</p>
              ) : (
                (batchesAdmin ?? []).map((b) => {
                  const counts: Record<string, number> = {};
                  for (const it of b.items) counts[it.status] = (counts[it.status] ?? 0) + 1;
                  return (
                    <div key={b.id} className="border border-gray-200 rounded-lg p-4">
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <p className="text-[13px] font-bold text-gray-900">{b.titulo}</p>
                        <div className="flex items-center gap-2">
                          {batchBadge(b.status)}
                          <button
                            onClick={() => setShowLogAdmin((p) => ({ ...p, [b.id]: !p[b.id] }))}
                            className="text-[11px] font-semibold text-gray-600 border border-gray-300 rounded-lg px-2.5 py-1 hover:bg-gray-100"
                          >
                            {showLogAdmin[b.id] ? "Ocultar bitácora ▲" : "Bitácora ▼"}
                          </button>
                        </div>
                      </div>
                      <p className="text-[12px] text-gray-500 mt-1">
                        Asignado a <strong>{b.assigned_email}</strong> · {new Date(b.created_at).toLocaleString()}
                      </p>
                      <div className="flex flex-wrap gap-2 mt-2">
                        {ITEM_STATUSES.map((s) => (
                          <span key={s.value} className="text-[11px] font-semibold bg-gray-100 text-gray-600 rounded-full px-2.5 py-1">
                            {s.label}: {counts[s.value] ?? 0}
                          </span>
                        ))}
                        <span className="text-[11px] font-semibold bg-gray-100 text-gray-600 rounded-full px-2.5 py-1">
                          Total: {b.items.length}
                        </span>
                      </div>
                      {b.nota_final ? (
                        <p className="text-[12px] text-green-700 bg-green-50 border border-green-200 rounded-lg px-3 py-2 mt-2">
                          <strong>Nota de finalización:</strong> {b.nota_final}
                        </p>
                      ) : null}
                      {showLogAdmin[b.id] ? (
                        <ul className="mt-3 space-y-1.5 text-[12px] text-gray-600 border-t border-gray-100 pt-3">
                          {b.events.map((e) => {
                            const itemUuid = e.item_id
                              ? b.items.find((i) => i.id === e.item_id)?.uuid_fiscal ?? null
                              : null;
                            return (
                              <li key={e.id} className="flex flex-wrap items-baseline gap-x-2 gap-y-0.5">
                                <span className="text-gray-400 shrink-0">{new Date(e.created_at).toLocaleString()}</span>
                                <span className="font-semibold text-gray-800 shrink-0">[{e.tipo}]</span>
                                {e.status_nuevo ? <span className="text-blue-700 font-semibold shrink-0">→ {e.status_nuevo}</span> : null}
                                {itemUuid ? (
                                  <button
                                    onClick={() => setModalUuid(itemUuid)}
                                    className="font-mono text-blue-700 hover:underline shrink-0"
                                  >
                                    {itemUuid}
                                  </button>
                                ) : null}
                                <span className="text-gray-500 shrink-0">{e.actor_email}</span>
                                {e.detalle ? <span>· {e.detalle}</span> : null}
                              </li>
                            );
                          })}
                          {b.events.length === 0 ? <li className="text-gray-400">Sin eventos.</li> : null}
                        </ul>
                      ) : null}
                    </div>
                  );
                })
              )}
            </div>
          ) : null}
        </div>

        <div className="bg-white border border-gray-200 rounded-xl p-6 space-y-3">
          <h2 className="text-sm font-bold text-gray-900">Cierre mensual</h2>
          <p className="text-[12px] text-gray-600 leading-relaxed">
            <strong>Consolidar Mes</strong> envía al histórico acumulado (con todos sus campos) únicamente:
            (1) las facturas encontradas en archivos complementarios este mes, y
            (2) las no encontradas que el colaborador marcó como <strong>“Subida al sistema”</strong>.
            Las que ya estaban en el histórico se omiten y las pendientes se quedan para el siguiente mes.
            El cierre nunca duplica: cada UUID se consolida una sola vez. A cada batch tocado se le escribe un evento de CONSOLIDACIÓN en su bitácora.
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

      <InvoiceModal
        uuid={modalUuid}
        localInfo={modalUuid ? localByUuid[modalUuid] ?? null : null}
        onClose={() => setModalUuid(null)}
      />
    </main>
  );
}