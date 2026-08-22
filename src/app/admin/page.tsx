"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import Link from "next/link";
import { useRouter } from "next/navigation";
import { createClient } from "@/lib/supabase/client";
import LogoutButton from "@/components/logout-button";

const TABLES = ["settings", "rules", "historical_invoices", "batches", "batch_items", "batch_events"] as const;

export default function AdminPage() {
  const supabase = useMemo(() => createClient(), []);
  const router = useRouter();
  const [email, setEmail] = useState("");
  const [allowed, setAllowed] = useState(false);
  const [msg, setMsg] = useState("");
  const [busy, setBusy] = useState("");
  const [cleanBefore, setCleanBefore] = useState(true);
  const [pending, setPending] = useState<Record<string, unknown> | null>(null);
  const fileRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    supabase.auth.getUser().then(async ({ data }) => {
      setEmail(data.user?.email ?? "");
      if (!data.user) return;
      const { data: prof } = await supabase
        .from("profiles")
        .select("role")
        .eq("id", data.user.id)
        .maybeSingle();
      const r = prof ? (prof as { role: string }).role : "";
      if (r === "SUPER_USUARIO") setAllowed(true);
      else router.replace("/");
    });
  }, [supabase, router]);

  async function generarRespaldo() {
    setBusy("backup");
    setMsg("");
    try {
      const backup: Record<string, unknown> = {};
      for (const t of TABLES) {
        const { data, error } = await supabase.from(t).select("*");
        if (error) throw new Error(t + ": " + error.message);
        backup[t] = data ?? [];
      }
      backup._meta = { generado: new Date().toISOString(), correo: email, version: 2 };
      const blob = new Blob([JSON.stringify(backup, null, 2)], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      const stamp = new Date().toISOString().slice(0, 16).replace(/[:T]/g, "-");
      a.href = url;
      a.download = "respaldo_conciliador_" + stamp + ".json";
      a.click();
      URL.revokeObjectURL(url);
      setMsg("Respaldo generado y descargado. Guárdalo en un lugar seguro.");
    } catch (e) {
      setMsg("Error al respaldar: " + (e instanceof Error ? e.message : "error inesperado"));
    }
    setBusy("");
  }

  function onFile(e: React.ChangeEvent<HTMLInputElement>) {
    const f = e.target.files?.[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(String(reader.result)) as Record<string, unknown>;
        setPending(parsed);
        setMsg(
          "Respaldo cargado (" + String((parsed._meta as { generado?: string } | undefined)?.generado ?? "sin fecha") + "): " +
            TABLES.map((t) => t + " " + (((parsed[t] as unknown[]) ?? []).length)).join(" · ")
        );
      } catch {
        setMsg("El archivo no es un respaldo válido.");
        setPending(null);
      }
    };
    reader.readAsText(f);
    e.target.value = "";
  }

  async function restaurar() {
    if (!pending) return;
    if (!window.confirm("Esto reemplazará los datos actuales con los del respaldo. ¿Continuar?")) return;
    setBusy("restore");
    setMsg("");
    try {
      if (cleanBefore) {
        await supabase.from("batch_events").delete().not("id", "is", null);
        await supabase.from("batch_items").delete().not("id", "is", null);
        await supabase.from("batches").delete().not("id", "is", null);
        await supabase.from("historical_invoices").delete().not("id", "is", null);
        await supabase.from("rules").delete().not("id", "is", null);
        await supabase.from("settings").delete().not("id", "is", null);
      }
      const insertar = async (t: (typeof TABLES)[number]) => {
        const rows = (pending[t] ?? []) as Record<string, unknown>[];
        for (let i = 0; i < rows.length; i += 200) {
          const part = rows.slice(i, i + 200);
          const { error } = await supabase.from(t).upsert(part, { onConflict: "id" });
          if (error) throw new Error(t + ": " + error.message);
        }
      };
      await insertar("settings");
      await insertar("rules");
      await insertar("historical_invoices");
      await insertar("batches");
      await insertar("batch_items");
      await insertar("batch_events");
      setMsg("Restauración completada con éxito.");
      setPending(null);
    } catch (e) {
      setMsg("Error al restaurar: " + (e instanceof Error ? e.message : "error inesperado"));
    }
    setBusy("");
  }

  async function vaciarHistorico() {
    if (!window.confirm("Se vaciará TODO el histórico acumulado. Esta acción es delicada. ¿Continuar?")) return;
    if (!window.confirm("Segunda confirmación: ¿vaciar histórico acumulado?")) return;
    setBusy("vacia");
    const { error, count } = await supabase
      .from("historical_invoices")
      .delete({ count: "exact" })
      .not("id", "is", null);
    setMsg(error ? "Error: " + error.message : "Histórico vaciado. Registros eliminados: " + String(count ?? 0));
    setBusy("");
  }

  if (!allowed) {
    return (
      <main className="min-h-screen bg-gray-100 text-gray-900 flex items-center justify-center">
        <p className="text-sm text-gray-500">Verificando permisos de super usuario…</p>
      </main>
    );
  }

  return (
    <main className="min-h-screen bg-gray-100 text-gray-900">
      <header className="bg-white border-b border-gray-200 sticky top-0 z-10">
        <div className="max-w-5xl mx-auto px-6 py-4 flex items-center justify-between gap-4">
          <div>
            <h1 className="font-bold text-gray-900">Conciliador de Facturas</h1>
            <p className="text-[11px] text-gray-500">Mantenimiento · exclusivo super usuario</p>
          </div>
          <div className="flex items-center gap-3">
            <nav className="flex items-center gap-1 text-[12px] font-semibold">
              <Link href="/" className="rounded-lg px-3 py-1.5 text-gray-600 hover:bg-gray-100">Panel</Link>
              <Link href="/reglas" className="rounded-lg px-3 py-1.5 text-gray-600 hover:bg-gray-100">Reglas y UMA</Link>
              <span className="rounded-lg bg-blue-50 text-blue-700 px-3 py-1.5">Mantenimiento</span>
            </nav>
            <span className="text-sm text-gray-600 hidden sm:inline">{email}</span>
            <LogoutButton />
          </div>
        </div>
      </header>

      <div className="max-w-5xl mx-auto px-6 py-8 space-y-6">
        {msg ? (
          <p className="text-[12px] text-gray-700 bg-white border border-gray-200 rounded-lg px-3 py-2">{msg}</p>
        ) : null}

        <div className="bg-white border border-gray-200 rounded-xl p-6 space-y-3">
          <h2 className="text-sm font-bold text-gray-900">Respaldos de la base de datos</h2>
          <p className="text-[12px] text-gray-600 leading-relaxed">
            Genera un archivo .json con el estado completo de la operación (histórico acumulado, reglas, configuración y registros históricos de lotes).
            Recomendado: antes de cada “Consolidar Mes” y antes de cualquier limpieza.
          </p>
          <button
            onClick={generarRespaldo}
            disabled={busy !== ""}
            className="rounded-lg bg-blue-700 text-white text-[12px] font-semibold px-4 py-2 hover:bg-blue-800 disabled:opacity-50"
          >
            {busy === "backup" ? "Generando…" : "Generar respaldo (.json)"}
          </button>
        </div>

        <div className="bg-white border border-gray-200 rounded-xl p-6 space-y-3">
          <h2 className="text-sm font-bold text-gray-900">Restaurar desde un respaldo</h2>
          <p className="text-[12px] text-gray-600 leading-relaxed">
            Selecciona el archivo .json de un respaldo previo. Con “Limpiar tablas antes” el sistema queda exactamente como estaba en ese punto.
          </p>
          <div className="flex flex-wrap items-center gap-3">
            <input ref={fileRef} type="file" accept=".json,application/json" onChange={onFile} className="text-[12px] text-gray-600" />
            <label className="flex items-center gap-2 text-[12px] font-semibold text-gray-700 select-none cursor-pointer">
              <input
                type="checkbox"
                checked={cleanBefore}
                onChange={(e) => setCleanBefore(e.target.checked)}
                className="accent-blue-700 w-3.5 h-3.5"
              />
              Limpiar tablas antes de restaurar
            </label>
            <button
              onClick={restaurar}
              disabled={!pending || busy !== ""}
              className="rounded-lg bg-amber-600 text-white text-[12px] font-semibold px-4 py-2 hover:bg-amber-700 disabled:opacity-50"
            >
              {busy === "restore" ? "Restaurando…" : "Restaurar respaldo"}
            </button>
          </div>
        </div>

        <div className="bg-white border border-gray-200 rounded-xl p-6 space-y-3">
          <h2 className="text-sm font-bold text-gray-900">Limpiezas de mantenimiento</h2>
          <button
            onClick={vaciarHistorico}
            disabled={busy !== ""}
            className="rounded-lg border border-red-300 text-red-600 text-[12px] font-semibold px-4 py-2 hover:bg-red-50 disabled:opacity-50"
          >
            Vaciar histórico acumulado
          </button>
          <p className="text-[11px] text-gray-500">
            Todas las acciones piden confirmación y quedan disponibles solo para tu rol SUPER_USUARIO gracias a RLS.
          </p>
        </div>
      </div>
    </main>
  );
}