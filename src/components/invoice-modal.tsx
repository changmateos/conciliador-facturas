"use client";

import { useEffect, useMemo, useState } from "react";
import { createClient } from "@/lib/supabase/client";

export interface LocalInvoiceInfo {
  fileName: string;
  sheetName: string;
  sourceRow: number;
  values: Record<string, string | number | null>;
}

interface InvoiceModalProps {
  uuid: string | null;
  localInfo: LocalInvoiceInfo | null;
  onClose: () => void;
}

export default function InvoiceModal({ uuid, localInfo, onClose }: InvoiceModalProps) {
  const supabase = useMemo(() => createClient(), []);
  const [histRow, setHistRow] = useState<Record<string, unknown> | null>(null);
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    if (!uuid) return;
    let alive = true;
    setLoading(true);
    setHistRow(null);
    (async () => {
      const { data } = await supabase
        .from("historical_invoices")
        .select("uuid_fiscal, folio, monto, retenciones, mes_periodo, anio_periodo, origen_archivo, datos_json")
        .eq("uuid_fiscal", uuid)
        .maybeSingle();
      if (alive) {
        setHistRow((data as Record<string, unknown> | null) ?? null);
        setLoading(false);
      }
    })();
    return () => {
      alive = false;
    };
  }, [uuid, supabase]);

  if (!uuid) return null;

  const datos =
    histRow && typeof histRow.datos_json === "object" && histRow.datos_json !== null
      ? (histRow.datos_json as Record<string, unknown>)
      : null;

  const entries: [string, unknown][] = datos
    ? Object.entries(datos)
    : localInfo
      ? Object.entries(localInfo.values)
      : [];

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
      onClick={onClose}
    >
      <div
        className="bg-white rounded-xl shadow-xl max-w-lg w-full max-h-[80vh] overflow-auto"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between gap-3 px-5 py-4 border-b border-gray-200 bg-gray-50 sticky top-0">
          <div className="min-w-0">
            <p className="text-[11px] font-bold tracking-wide text-gray-500 uppercase">Factura</p>
            <p className="text-[12px] font-mono text-gray-900 break-all">{uuid}</p>
          </div>
          <button
            onClick={onClose}
            className="w-8 h-8 rounded-lg border border-gray-300 text-gray-500 hover:bg-gray-100 shrink-0"
            title="Cerrar"
          >
            ×
          </button>
        </div>

        <div className="p-5 space-y-3">
          {loading ? (
            <p className="text-[12px] text-gray-500">Buscando en el histórico…</p>
          ) : histRow ? (
            <p className="text-[12px] font-semibold text-green-700 bg-green-50 border border-green-200 rounded-lg px-3 py-2">
              En histórico · periodo {String(histRow.mes_periodo ?? "")}/{String(histRow.anio_periodo ?? "")} · origen: {String(histRow.origen_archivo ?? "—")}
            </p>
          ) : (
            <p className="text-[12px] font-semibold text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">
              Aún no está en el histórico de Supabase.
            </p>
          )}

          {localInfo ? (
            <p className="text-[12px] font-semibold text-blue-700 bg-blue-50 border border-blue-200 rounded-lg px-3 py-2">
              En archivos del mes: {localInfo.fileName} · hoja "{localInfo.sheetName}" · fila {localInfo.sourceRow}
            </p>
          ) : null}

          {entries.length > 0 ? (
            <div className="overflow-hidden rounded-lg border border-gray-200">
              <table className="w-full text-[12px]">
                <tbody>
                  {entries.map(([k, v]) => (
                    <tr key={k} className="border-t border-gray-100 first:border-t-0">
                      <td className="px-3 py-1.5 bg-gray-50 font-semibold text-gray-600 w-[45%] align-top">{k}</td>
                      <td className="px-3 py-1.5 text-gray-800 break-all">
                        {v === null || v === undefined ? "" : String(v)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : !loading ? (
            <p className="text-[12px] text-gray-500">
              Sin datos adicionales todavía: no está en el histórico ni en los archivos cargados en esta sesión.
            </p>
          ) : null}
        </div>
      </div>
    </div>
  );
}