"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import Link from "next/link";
import { createClient } from "@/lib/supabase/client";
import LogoutButton from "@/components/logout-button";
import InvoiceModal from "@/components/invoice-modal";
import type { BatchRow } from "@/lib/batches";
import {
  BATCH_STATUS_LABEL,
  ITEM_STATUSES,
  fetchBatches,
  finalizeBatch,
  setItemStatus,
  startBatch,
} from "@/lib/batches";

function batchBadge(status: string) {
  if (status === "COMPLETADO")
    return <span className="text-[10px] font-bold rounded-full px-2.5 py-1 bg-green-100 text-green-700">COMPLETADO</span>;
  if (status === "EN_PROCESO")
    return <span className="text-[10px] font-bold rounded-full px-2.5 py-1 bg-blue-100 text-blue-700">EN PROCESO</span>;
  return <span className="text-[10px] font-bold rounded-full px-2.5 py-1 bg-amber-100 text-amber-700">PENDIENTE</span>;
}

function itemBadge(status: string) {
  if (status === "SUBIDA_SISTEMA")
    return <span className="text-[10px] font-bold rounded-full px-2 py-0.5 bg-green-100 text-green-700">SUBIDA AL SISTEMA</span>;
  if (status === "PENDIENTE_TERCIERO")
    return <span className="text-[10px] font-bold rounded-full px-2 py-0.5 bg-amber-100 text-amber-700">PENDIENTE POR TERCERO</span>;
  if (status === "NO_CORRESPONDE")
    return <span className="text-[10px] font-bold rounded-full px-2 py-0.5 bg-gray-200 text-gray-600">NO CORRESPONDE</span>;
  return <span className="text-[10px] font-bold rounded-full px-2 py-0.5 bg-amber-50 text-amber-600 border border-amber-200">PENDIENTE</span>;
}

export default function MisBatchesPage() {
  const supabase = useMemo(() => createClient(), []);
  const [email, setEmail] = useState("");
  const [userId, setUserId] = useState("");
  const [role, setRole] = useState("");
  const [batches, setBatches] = useState<BatchRow[] | null>(null);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState("");
  const [showLog, setShowLog] = useState<Record<string, boolean>>({});
  const [itemStatusSel, setItemStatusSel] = useState<Record<string, string>>({});
  const [itemComment, setItemComment] = useState<Record<string, string>>({});
  const [finalComment, setFinalComment] = useState<Record<string, string>>({});
  const [modalUuid, setModalUuid] = useState<string | null>(null);

  const refresh = useCallback(
    async (uid: string) => {
      try {
        setError("");
        const list = await fetchBatches(supabase, uid);
        setBatches(list);
      } catch (e) {
        setError(e instanceof Error ? e.message : "Error cargando batches");
      }
    },
    [supabase]
  );

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
        refresh(data.user.id);
      }
    });
  }, [supabase, refresh]);

  async function handleItem(batch: BatchRow, itemId: string) {
    const st = itemStatusSel[itemId] ?? "PENDIENTE";
    const comment = (itemComment[itemId] ?? "").trim();
    if (st === "PENDIENTE_TERCIERO" && comment === "") {
      window.alert("Escribe el motivo por el que queda pendiente por tercero.");
      return;
    }
    if (st === "SUBIDA_SISTEMA" && comment === "") {
      window.alert("Escribe con qué folio o póliza se subió la factura al sistema.");
      return;
    }
    const item = batch.items.find((i) => i.id === itemId);
    if (!item) return;
    setBusy(itemId);
    try {
      await setItemStatus(supabase, item, st, comment, email);
      await refresh(userId);
    } catch (e) {
      window.alert(e instanceof Error ? e.message : "Error al guardar");
    }
    setBusy("");
  }

  async function handleStart(batch: BatchRow) {
    setBusy(batch.id);
    try {
      await startBatch(supabase, batch.id, email);
      await refresh(userId);
    } catch (e) {
      window.alert(e instanceof Error ? e.message : "Error al actualizar");
    }
    setBusy("");
  }

  async function handleFinalize(batch: BatchRow) {
    const comment = (finalComment[batch.id] ?? "").trim();
    if (comment === "") {
      window.alert("El comentario de finalización es obligatorio: describe las acciones realizadas con este batch.");
      return;
    }
    setBusy(batch.id);
    try {
      await finalizeBatch(supabase, batch.id, comment, email);
      await refresh(userId);
    } catch (e) {
      window.alert(e instanceof Error ? e.message : "Error al finalizar");
    }
    setBusy("");
  }

  return (
    <main className="min-h-screen bg-gray-100 text-gray-900">
      <header className="bg-white border-b border-gray-200 sticky top-0 z-10">
        <div className="max-w-5xl mx-auto px-6 py-4 flex items-center justify-between gap-4">
          <div>
            <h1 className="font-bold text-gray-900">Conciliador de Facturas</h1>
            <p className="text-[11px] text-gray-500">Mis lotes de trabajo · panel del colaborador</p>
          </div>
          <div className="flex items-center gap-3">
            <nav className="flex items-center gap-1 text-[12px] font-semibold">
              {role === "ADMIN_CONTADOR" || role === "SUPER_USUARIO" ? (
                <>
                  <Link href="/" className="rounded-lg px-3 py-1.5 text-gray-600 hover:bg-gray-100">Panel</Link>
                  <Link href="/reglas" className="rounded-lg px-3 py-1.5 text-gray-600 hover:bg-gray-100">Reglas y UMA</Link>
                  <span className="rounded-lg bg-blue-50 text-blue-700 px-3 py-1.5">Mis lotes</span>
                </>
              ) : (
                <span className="rounded-lg bg-blue-50 text-blue-700 px-3 py-1.5">Mis Batches</span>
              )}
            </nav>
            <span className="text-sm text-gray-600 hidden sm:inline">{email}</span>
            <LogoutButton />
          </div>
        </div>
      </header>

      <div className="max-w-5xl mx-auto px-6 py-8 space-y-6">
        {error ? (
          <p className="text-[12px] text-red-700 bg-red-50 border border-red-200 rounded-lg px-3 py-2">{error}</p>
        ) : null}

        {batches === null ? (
          <p className="text-sm text-gray-500">Cargando tus batches…</p>
        ) : batches.length === 0 ? (
          <p className="text-sm text-gray-500 bg-white border border-gray-200 rounded-xl px-4 py-6 text-center">
            No tienes lotes de trabajo asignados todavía. El Contador Principal te los asignará desde el Panel.
          </p>
        ) : (
          batches.map((b) => (
            <div key={b.id} className="bg-white border border-gray-200 rounded-xl overflow-hidden">
              <div className="flex flex-wrap items-center justify-between gap-3 px-5 py-4 bg-gray-50 border-b border-gray-200">
                <div className="min-w-0">
                  <p className="text-sm font-bold text-gray-900 truncate">{b.titulo}</p>
                  <p className="text-[11px] text-gray-500 mt-0.5">
                    Asignado a {b.assigned_email} · {new Date(b.created_at).toLocaleString()}
                  </p>
                </div>
                <div className="flex items-center gap-2">
                  {batchBadge(b.status)}
                  <button
                    onClick={() => setShowLog((p) => ({ ...p, [b.id]: !p[b.id] }))}
                    className="text-[11px] font-semibold text-gray-600 border border-gray-300 rounded-lg px-2.5 py-1 hover:bg-gray-100"
                  >
                    {showLog[b.id] ? "Ocultar bitácora ▲" : "Ver bitácora ▼"}
                  </button>
                </div>
              </div>

              <div className="p-5 space-y-4">
                <div className="overflow-x-auto rounded-lg border border-gray-200">
                  <table className="w-full text-[12px]">
                    <thead className="bg-gray-50">
                      <tr>
                        <th className="px-3 py-2 text-left font-semibold text-gray-500">UUID</th>
                        <th className="px-3 py-2 text-left font-semibold text-gray-500">Estatus</th>
                        <th className="px-3 py-2 text-left font-semibold text-gray-500">Última nota</th>
                        <th className="px-3 py-2 text-left font-semibold text-gray-500">Cambiar estatus</th>
                      </tr>
                    </thead>
                    <tbody>
                      {b.items.map((it) => (
                        <tr key={it.id} className="border-t border-gray-100 align-top">
                          <td className="px-3 py-2">
                            <button
                              onClick={() => setModalUuid(it.uuid_fiscal)}
                              title="Ver información de la factura"
                              className="font-mono text-blue-700 hover:underline"
                            >
                              {it.uuid_fiscal}
                            </button>
                          </td>
                          <td className="px-3 py-2">{itemBadge(it.status)}</td>
                          <td className="px-3 py-2 text-gray-600 max-w-[220px]">{it.nota ?? "—"}</td>
                          <td className="px-3 py-2">
                            {b.status === "COMPLETADO" ? (
                              <span className="text-gray-400">Lote cerrado</span>
                            ) : (
                              <div className="flex flex-wrap items-center gap-2">
                                <select
                                  value={itemStatusSel[it.id] ?? it.status}
                                  onChange={(e) => setItemStatusSel((p) => ({ ...p, [it.id]: e.target.value }))}
                                  className="rounded-lg border border-gray-300 bg-white px-2.5 py-1.5 text-[12px] text-gray-900"
                                >
                                  {ITEM_STATUSES.map((s) => (
                                    <option key={s.value} value={s.value}>{s.label}</option>
                                  ))}
                                </select>
                                <input
                                  value={itemComment[it.id] ?? ""}
                                  onChange={(e) => setItemComment((p) => ({ ...p, [it.id]: e.target.value }))}
                                  placeholder="Comentario (folio/póliza o motivo)"
                                  className="rounded-lg border border-gray-300 px-2.5 py-1.5 text-[12px] text-gray-900 w-56"
                                />
                                <button
                                  onClick={() => handleItem(b, it.id)}
                                  disabled={busy === it.id}
                                  className="rounded-lg bg-blue-700 text-white text-[12px] font-semibold px-3 py-1.5 hover:bg-blue-800 disabled:opacity-50"
                                >
                                  Registrar
                                </button>
                              </div>
                            )}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {b.status !== "COMPLETADO" ? (
                  <div className="border border-gray-200 rounded-lg p-4 bg-gray-50 space-y-3">
                    <div className="flex flex-wrap items-center gap-2">
                      {b.status === "PENDIENTE" ? (
                        <button
                          onClick={() => handleStart(b)}
                          disabled={busy === b.id}
                          className="rounded-lg border border-blue-300 text-blue-700 text-[12px] font-semibold px-3.5 py-2 hover:bg-blue-50 disabled:opacity-50"
                        >
                          Marcar en proceso
                        </button>
                      ) : null}
                      <p className="text-[12px] text-gray-600">
                        Para finalizar, describe las acciones realizadas con este lote (obligatorio):
                      </p>
                    </div>
                    <textarea
                      value={finalComment[b.id] ?? ""}
                      onChange={(e) => setFinalComment((p) => ({ ...p, [b.id]: e.target.value }))}
                      rows={2}
                      placeholder="Ej. 12 facturas subidas con póliza P-204 a P-215; 3 quedan pendientes por tercero (faltan acuses)."
                      className="w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
                    />
                    <button
                      onClick={() => handleFinalize(b)}
                      disabled={busy === b.id}
                      className="rounded-lg bg-green-600 text-white text-[12px] font-semibold px-4 py-2 hover:bg-green-700 disabled:opacity-50"
                    >
                      Finalizar lote
                    </button>
                  </div>
                ) : (
                  <p className="text-[12px] text-green-700 bg-green-50 border border-green-200 rounded-lg px-3 py-2">
                    <strong>Nota de finalización:</strong> {b.nota_final ?? "—"}
                  </p>
                )}

                {showLog[b.id] ? (
                  <div className="border border-gray-200 rounded-lg p-4">
                    <p className="text-[11px] font-bold tracking-wide text-gray-500 uppercase mb-2">Bitácora del lote</p>
                    <ul className="space-y-2 text-[12px] text-gray-600">
                      {b.events.map((e) => {
                        const itemUuid = e.item_id
                          ? b.items.find((i) => i.id === e.item_id)?.uuid_fiscal ?? null
                          : null;
                        return (
                          <li key={e.id} className="flex flex-wrap items-baseline gap-x-2 gap-y-0.5">
                            <span className="text-gray-400 shrink-0">{new Date(e.created_at).toLocaleString()}</span>
                            <span className="font-semibold text-gray-800 shrink-0">[{e.tipo}]</span>
                            {e.status_nuevo ? (
                              <span className="text-blue-700 font-semibold shrink-0">→ {e.status_nuevo}</span>
                            ) : null}
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
                      {b.events.length === 0 ? <li className="text-gray-400">Sin eventos aún.</li> : null}
                    </ul>
                  </div>
                ) : null}
              </div>
            </div>
          ))
        )}
      </div>

      <InvoiceModal uuid={modalUuid} localInfo={null} onClose={() => setModalUuid(null)} />
    </main>
  );
}