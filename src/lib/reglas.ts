import type { ResultRow } from "@/lib/conciliacion";

export interface RuleRow {
  id: string;
  nombre: string;
  tipo_condicion: string;
  valor_limite: number | null;
  etiqueta: string;
  palabras_prohibidas: string | null;
  campo_nombre: string | null;
}

export const RULE_TYPES: { value: string; label: string }[] = [
  { value: "MONTO_MAYOR_UMAS", label: "Monto total mayor a (UMA × factor)" },
  { value: "MONTO_MAYOR_PESOS", label: "Monto total mayor a pesos" },
  { value: "CON_RETENCION", label: "Con retenciones (monto > 0)" },
  { value: "SIN_RETENCION", label: "Sin retención (RESICO)" },
  { value: "CONCEPTO_PROHIBIDO", label: "Concepto contiene palabras prohibidas" },
  { value: "COMPLEMENTO_INE", label: "Con complemento INE" },
  { value: "CAMPO_CONTIENE", label: "Campo del diccionario contiene palabras" },
  { value: "CAMPO_MAYOR", label: "Campo del diccionario mayor a valor" },
];

// Comparación en centavos: "estrictamente mayor".
const cents = (n: number) => Math.round(n * 100);

export function parsePalabras(texto: string | null): string[] {
  if (!texto) return [];
  return texto
    .split(/[\n,;]+/)
    .map((p) => p.trim().toLowerCase())
    .filter((p) => p.length > 0);
}

function toNumberLike(v: string | number | null | undefined): number | null {
  if (v === null || v === undefined) return null;
  if (typeof v === "number") return v;
  const t = String(v).trim().replace(/[$,\s]/g, "");
  if (t === "") return null;
  const n = Number(t);
  return Number.isFinite(n) ? n : null;
}

export interface RuleInput {
  monto: number | null;
  retenciones: number | null;
  concepto: string;
  complemento: string;
  values: Record<string, string | number | null>;
}

export function evaluateRule(rule: RuleRow, input: RuleInput, uma: number): boolean {
  if (rule.tipo_condicion === "CON_RETENCION") {
    return input.retenciones !== null && input.retenciones > 0;
  }
  if (rule.tipo_condicion === "SIN_RETENCION") {
    return input.retenciones === null || input.retenciones <= 0;
  }
  if (rule.tipo_condicion === "CONCEPTO_PROHIBIDO") {
    const palabras = parsePalabras(rule.palabras_prohibidas);
    if (palabras.length === 0 || input.concepto === "") return false;
    const concepto = input.concepto.toLowerCase();
    return palabras.some((p) => concepto.includes(p));
  }
  if (rule.tipo_condicion === "COMPLEMENTO_INE") {
    // palabra completa "INE" en columnas de complemento y/o concepto
    const text = (input.complemento + " " + input.concepto).toLowerCase();
    if (text.trim() === "") return false;
    return /\bine\b/.test(text);
  }
  if (rule.tipo_condicion === "CAMPO_CONTIENE") {
    const campo = rule.campo_nombre;
    if (!campo) return false;
    const v = input.values[campo];
    const text = v === null || v === undefined ? "" : String(v).toLowerCase();
    if (text === "") return false;
    return parsePalabras(rule.palabras_prohibidas).some((p) => text.includes(p));
  }
  if (rule.tipo_condicion === "CAMPO_MAYOR") {
    const campo = rule.campo_nombre;
    if (!campo) return false;
    const n = toNumberLike(input.values[campo]);
    if (n === null) return false;
    return cents(n) > cents(rule.valor_limite ?? 0);
  }
  if (input.monto === null) return false;
  if (rule.tipo_condicion === "MONTO_MAYOR_UMAS") {
    return cents(input.monto) > cents((rule.valor_limite ?? 0) * uma);
  }
  if (rule.tipo_condicion === "MONTO_MAYOR_PESOS") {
    return cents(input.monto) > cents(rule.valor_limite ?? 0);
  }
  return false;
}

export function ruleDescription(rule: RuleRow, uma: number): string {
  if (rule.tipo_condicion === "CON_RETENCION") return "Retenciones > 0";
  if (rule.tipo_condicion === "SIN_RETENCION") return "Retenciones = 0 o sin retención (RESICO)";
  if (rule.tipo_condicion === "CONCEPTO_PROHIBIDO") {
    const palabras = parsePalabras(rule.palabras_prohibidas);
    return "Concepto contiene: " + (palabras.length > 0 ? palabras.join(", ") : "(sin palabras)");
  }
  if (rule.tipo_condicion === "COMPLEMENTO_INE") return "Contiene complemento INE";
  if (rule.tipo_condicion === "CAMPO_CONTIENE") {
    const palabras = parsePalabras(rule.palabras_prohibidas);
    return "[" + (rule.campo_nombre ?? "?") + "] contiene: " + (palabras.length > 0 ? palabras.join(", ") : "(sin palabras)");
  }
  if (rule.tipo_condicion === "CAMPO_MAYOR") {
    return "[" + (rule.campo_nombre ?? "?") + "] > " + (rule.valor_limite ?? 0);
  }
  if (rule.tipo_condicion === "MONTO_MAYOR_UMAS") {
    const limite = (rule.valor_limite ?? 0) * uma;
    return "Monto total > " + (rule.valor_limite ?? 0) + " UMAs (" + limite.toFixed(2) + " MXN)";
  }
  return "Monto total > " + (rule.valor_limite ?? 0) + " MXN";
}

export interface RuleGroup {
  key: string;
  rule: RuleRow | null; // null = grupo "Sin regla aplicada"
  rows: ResultRow[];
}

// Cada regla segmenta su propio grupo; una factura puede caer en varios.
export function classifyRows(
  results: ResultRow[],
  rules: RuleRow[],
  uma: number,
  inputOf: (r: ResultRow) => RuleInput
): RuleGroup[] {
  const groups: RuleGroup[] = rules.map((rule) => ({
    key: rule.id,
    rule,
    rows: results.filter((r) => evaluateRule(rule, inputOf(r), uma)),
  }));

  const matched = new Set<string>();
  for (const g of groups) {
    for (const r of g.rows) matched.add(r.uuid + "|" + r.sourceRow);
  }

  groups.push({
    key: "SIN_REGLA",
    rule: null,
    rows: results.filter((r) => !matched.has(r.uuid + "|" + r.sourceRow)),
  });
  return groups;
}