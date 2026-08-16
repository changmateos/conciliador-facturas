"use client";

import { useEffect, useMemo, useState } from "react";
import Link from "next/link";
import { createClient } from "@/lib/supabase/client";
import LogoutButton from "@/components/logout-button";
import { loadAllMemory } from "@/lib/excel/memory";
import type { RuleRow } from "@/lib/reglas";
import { RULE_TYPES } from "@/lib/reglas";

const SIN_VALOR = ["CON_RETENCION", "SIN_RETENCION", "CONCEPTO_PROHIBIDO", "COMPLEMENTO_INE", "CAMPO_CONTIENE"];
const ES_CAMPO = ["CAMPO_CONTIENE", "CAMPO_MAYOR"];

interface RuleCardProps {
  rule: RuleRow;
  dictFields: string[];
  onSave: (patch: Partial<RuleRow>) => void;
  onDelete: () => void;
}

function RuleCard({ rule, dictFields, onSave, onDelete }: RuleCardProps) {
  const [valor, setValor] = useState(rule.valor_limite === null ? "" : String(rule.valor_limite));
  const [etiqueta, setEtiqueta] = useState(rule.etiqueta);
  const [palabras, setPalabras] = useState(rule.palabras_prohibidas ?? "");
  const [campo, setCampo] = useState(rule.campo_nombre ?? "");
  const tipoLabel =
    RULE_TYPES.find((t) => t.value === rule.tipo_condicion)?.label ?? rule.tipo_condicion;
  const sinValor = SIN_VALOR.includes(rule.tipo_condicion);
  const esCampo = ES_CAMPO.includes(rule.tipo_condicion);

  function save() {
    const patch: Partial<RuleRow> = { etiqueta: etiqueta.trim() || rule.etiqueta };
    if (rule.tipo_condicion === "CONCEPTO_PROHIBIDO" || rule.tipo_condicion === "CAMPO_CONTIENE") {
      patch.palabras_prohibidas = palabras;
    }
    if (esCampo) {
      if (!campo) {
        window.alert("Selecciona un campo del diccionario");
        return;
      }
      patch.campo_nombre = campo;
    }
    if (!sinValor) {
      const n = Number(valor);
      if (!Number.isFinite(n) || n <= 0) {
        window.alert("Valor límite inválido");
        return;
      }
      patch.valor_limite = n;
    }
    onSave(patch);
  }

  return (
    <div className="border border-gray-200 rounded-lg p-4 bg-white">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <p className="text-[13px] font-bold text-gray-900">
          {rule.nombre}
          {esCampo && rule.campo_nombre ? (
            <span className="ml-2 text-[11px] font-semibold text-blue-700 bg-blue-50 rounded-full px-2 py-0.5">[{rule.campo_nombre}]</span>
          ) : null}
        </p>
        <span className="text-[11px] font-semibold text-gray-600 bg-gray-100 rounded-full px-2.5 py-1">
          {tipoLabel}
        </span>
      </div>
      <div className="grid sm:grid-cols-3 gap-3 mt-3">
        <label className="block">
          <span className="text-[11px] font-semibold text-gray-600">Valor límite</span>
          <input
            value={valor}
            onChange={(e) => setValor(e.target.value)}
            disabled={sinValor}
            className="mt-1 w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900 disabled:bg-gray-100 disabled:text-gray-500"
          />
        </label>
        <label className="block">
          <span className="text-[11px] font-semibold text-gray-600">Etiqueta del grupo</span>
          <input
            value={etiqueta}
            onChange={(e) => setEtiqueta(e.target.value)}
            className="mt-1 w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
          />
        </label>
        <div className="flex items-end gap-2">
          <button
            onClick={save}
            className="rounded-lg bg-blue-700 text-white text-[12px] font-semibold px-3.5 py-2 hover:bg-blue-800"
          >
            Guardar cambios
          </button>
          <button
            onClick={onDelete}
            className="rounded-lg border border-red-200 text-red-600 text-[12px] font-semibold px-3.5 py-2 hover:bg-red-50"
          >
            Eliminar
          </button>
        </div>
      </div>
      {esCampo ? (
        <label className="block mt-3">
          <span className="text-[11px] font-semibold text-gray-600">Campo del diccionario</span>
          <select
            value={campo}
            onChange={(e) => setCampo(e.target.value)}
            className="mt-1 w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
          >
            <option value="">Selecciona campo…</option>
            {dictFields.map((f) => (
              <option key={f} value={f}>[{f}]</option>
            ))}
          </select>
        </label>
      ) : null}
      {rule.tipo_condicion === "CONCEPTO_PROHIBIDO" || rule.tipo_condicion === "CAMPO_CONTIENE" ? (
        <label className="block mt-3">
          <span className="text-[11px] font-semibold text-gray-600">
            Palabras a buscar (separadas por coma o por renglón, sin importar mayúsculas)
          </span>
          <textarea
            value={palabras}
            onChange={(e) => setPalabras(e.target.value)}
            rows={3}
            className="mt-1 w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
          />
        </label>
      ) : null}
    </div>
  );
}

export default function ReglasPage() {
  const supabase = useMemo(() => createClient(), []);
  const [email, setEmail] = useState("");
  const [uma, setUma] = useState("117.31");
  const [umaSaved, setUmaSaved] = useState(false);
  const [rules, setRules] = useState<RuleRow[]>([]);

  const [nombre, setNombre] = useState("");
  const [tipo, setTipo] = useState("MONTO_MAYOR_UMAS");
  const [valor, setValor] = useState("90");
  const [etiqueta, setEtiqueta] = useState("");
  const [palabras, setPalabras] = useState("");
  const [campo, setCampo] = useState("");

  // Campos recordados del diccionario de datos (memoria del navegador)
  const dictFields = useMemo(() => {
    const all = loadAllMemory();
    const set = new Set<string>();
    for (const key of Object.keys(all)) {
      const cfg = all[key];
      if (cfg && cfg.mapping) {
        for (const h of Object.keys(cfg.mapping)) set.add(h);
      }
    }
    return Array.from(set).sort();
  }, []);

  useEffect(() => {
    load();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [supabase]);

  async function load() {
    const { data: sess } = await supabase.auth.getUser();
    setEmail(sess.user?.email ?? "");
    const { data: s } = await supabase.from("settings").select("uma_valor").eq("id", 1).maybeSingle();
    if (s) setUma(String((s as { uma_valor: number }).uma_valor));
    const { data: r } = await supabase
      .from("rules")
      .select("id, nombre, tipo_condicion, valor_limite, etiqueta, palabras_prohibidas, campo_nombre")
      .order("created_at", { ascending: true });
    if (r) setRules(r as RuleRow[]);
  }

  async function saveUma() {
    const n = Number(uma);
    if (!Number.isFinite(n) || n <= 0) {
      window.alert("Valor de UMA inválido");
      return;
    }
    const { error } = await supabase
      .from("settings")
      .upsert({ id: 1, uma_valor: n, updated_at: new Date().toISOString() });
    if (error) {
      window.alert(error.message);
      return;
    }
    setUmaSaved(true);
    setTimeout(() => setUmaSaved(false), 1500);
  }

  const sinValorNuevo = SIN_VALOR.includes(tipo);
  const esCampoNuevo = ES_CAMPO.includes(tipo);

  async function createRule() {
    if (!nombre.trim() || !etiqueta.trim()) {
      window.alert("Nombre y etiqueta son obligatorios");
      return;
    }
    let valorNum: number | null = null;
    let palabrasTxt: string | null = null;
    let campoTxt: string | null = null;
    if (tipo === "CONCEPTO_PROHIBIDO" || tipo === "CAMPO_CONTIENE") {
      if (palabras.trim() === "") {
        window.alert("Escribe al menos una palabra de búsqueda");
        return;
      }
      palabrasTxt = palabras;
    }
    if (esCampoNuevo) {
      if (!campo) {
        window.alert("Selecciona un campo del diccionario");
        return;
      }
      campoTxt = campo;
    }
    if (!sinValorNuevo) {
      valorNum = Number(valor);
      if (!Number.isFinite(valorNum) || valorNum <= 0) {
        window.alert("Valor límite inválido");
        return;
      }
    }
    const { error } = await supabase.from("rules").insert({
      nombre: nombre.trim(),
      tipo_condicion: tipo,
      valor_limite: valorNum,
      etiqueta: etiqueta.trim(),
      palabras_prohibidas: palabrasTxt,
      campo_nombre: campoTxt,
    });
    if (error) {
      window.alert(error.message);
      return;
    }
    setNombre("");
    setValor("90");
    setEtiqueta("");
    setPalabras("");
    setCampo("");
    await load();
  }

  async function updateRule(rule: RuleRow, patch: Partial<RuleRow>) {
    const { error } = await supabase.from("rules").update(patch).eq("id", rule.id);
    if (error) window.alert(error.message);
    await load();
  }

  async function deleteRule(rule: RuleRow) {
    if (!window.confirm("¿Eliminar la regla " + rule.nombre + "?")) return;
    const { error } = await supabase.from("rules").delete().eq("id", rule.id);
    if (error) window.alert(error.message);
    await load();
  }

  return (
    <main className="min-h-screen bg-gray-100 text-gray-900">
      <header className="bg-white border-b border-gray-200 sticky top-0 z-10">
        <div className="max-w-5xl mx-auto px-6 py-4 flex items-center justify-between gap-4">
          <div>
            <h1 className="font-bold text-gray-900">Conciliador de Facturas</h1>
            <p className="text-[11px] text-gray-500">Gestor de Reglas Fiscales</p>
          </div>
          <div className="flex items-center gap-3">
            <nav className="flex items-center gap-1 text-[12px] font-semibold">
              <Link href="/" className="rounded-lg px-3 py-1.5 text-gray-600 hover:bg-gray-100">
                Panel
              </Link>
              <span className="rounded-lg bg-blue-50 text-blue-700 px-3 py-1.5">Reglas y UMA</span>
            </nav>
            <span className="text-sm text-gray-600 hidden sm:inline">{email}</span>
            <LogoutButton />
          </div>
        </div>
      </header>

      <div className="max-w-5xl mx-auto px-6 py-8 space-y-6">
        <div className="bg-white border border-gray-200 rounded-xl p-6">
          <h2 className="text-sm font-bold text-gray-900">Valor de la UMA (MXN, diario)</h2>
          <p className="text-[12px] text-gray-500 mt-0.5">
            Se usa en todas las reglas de tipo “UMA × factor”. Puedes cambiarlo cuando se publique la UMA del año.
          </p>
          <div className="flex items-center gap-3 mt-4">
            <input
              value={uma}
              onChange={(e) => setUma(e.target.value)}
              className="w-40 rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
            />
            <button
              onClick={saveUma}
              className="rounded-lg bg-blue-700 text-white text-[12px] font-semibold px-4 py-2 hover:bg-blue-800"
            >
              {umaSaved ? "✓ Guardado" : "Guardar UMA"}
            </button>
          </div>
        </div>

        <div className="bg-white border border-gray-200 rounded-xl p-6">
          <h2 className="text-sm font-bold text-gray-900">Crear nueva regla</h2>
          <p className="text-[12px] text-gray-500 mt-0.5">
            Cada regla segmenta su propio grupo y aparece como insignia en la columna “Segmentación” del reporte. Los tipos “Campo del diccionario…” te dejan crear reglas múltiples sobre cualquier campo mapeado.
          </p>
          <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-3 mt-4">
            <label className="block">
              <span className="text-[11px] font-semibold text-gray-600">Nombre</span>
              <input
                value={nombre}
                onChange={(e) => setNombre(e.target.value)}
                placeholder="Ej. Pagos en una sola exhibición"
                className="mt-1 w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
              />
            </label>
            <label className="block">
              <span className="text-[11px] font-semibold text-gray-600">Tipo de condición</span>
              <select
                value={tipo}
                onChange={(e) => setTipo(e.target.value)}
                className="mt-1 w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
              >
                {RULE_TYPES.map((t) => (
                  <option key={t.value} value={t.value}>{t.label}</option>
                ))}
              </select>
            </label>
            <label className="block">
              <span className="text-[11px] font-semibold text-gray-600">Valor límite</span>
              <input
                value={valor}
                onChange={(e) => setValor(e.target.value)}
                disabled={sinValorNuevo}
                className="mt-1 w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900 disabled:bg-gray-100 disabled:text-gray-500"
              />
            </label>
            <label className="block">
              <span className="text-[11px] font-semibold text-gray-600">Etiqueta del grupo</span>
              <input
                value={etiqueta}
                onChange={(e) => setEtiqueta(e.target.value)}
                placeholder="Ej. Conceptos prohibidos"
                className="mt-1 w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
              />
            </label>
          </div>
          {esCampoNuevo ? (
            <label className="block mt-3">
              <span className="text-[11px] font-semibold text-gray-600">Campo del diccionario de datos</span>
              <select
                value={campo}
                onChange={(e) => setCampo(e.target.value)}
                className="mt-1 w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
              >
                <option value="">Selecciona campo…</option>
                {dictFields.map((f) => (
                  <option key={f} value={f}>[{f}]</option>
                ))}
              </select>
              {dictFields.length === 0 ? (
                <span className="block mt-1 text-[11px] text-amber-700">
                  Aún no hay campos recordados: carga y mapea tus archivos una vez en el Panel y recarga esta página.
                </span>
              ) : null}
            </label>
          ) : null}
          {tipo === "CONCEPTO_PROHIBIDO" || tipo === "CAMPO_CONTIENE" ? (
            <label className="block mt-3">
              <span className="text-[11px] font-semibold text-gray-600">
                Palabras a buscar (separadas por coma o por renglón, sin importar mayúsculas)
              </span>
              <textarea
                value={palabras}
                onChange={(e) => setPalabras(e.target.value)}
                rows={3}
                placeholder={"RENTA, ANTICIPO, PRESTAMO"}
                className="mt-1 w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm text-gray-900"
              />
            </label>
          ) : null}
          <button
            onClick={createRule}
            className="mt-4 rounded-lg bg-blue-700 text-white text-[12px] font-semibold px-4 py-2 hover:bg-blue-800"
          >
            Agregar regla
          </button>
        </div>

        <div className="space-y-4">
          <h2 className="text-sm font-bold text-gray-900">Reglas existentes ({rules.length})</h2>
          {rules.map((rule) => (
            <RuleCard
              key={rule.id}
              rule={rule}
              dictFields={dictFields}
              onSave={(patch) => updateRule(rule, patch)}
              onDelete={() => deleteRule(rule)}
            />
          ))}
          {rules.length === 0 ? (
            <p className="text-[13px] text-gray-500 bg-white border border-gray-200 rounded-xl px-4 py-3">
              Aún no hay reglas. Crea la primera con el formulario de arriba.
            </p>
          ) : null}
        </div>
      </div>
    </main>
  );
}