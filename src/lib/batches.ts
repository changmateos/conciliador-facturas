import type { SupabaseClient } from "@supabase/supabase-js";

export interface BatchEventRow {
  id: string;
  batch_id: string;
  item_id: string | null;
  actor_email: string | null;
  tipo: string;
  status_nuevo: string | null;
  detalle: string | null;
  created_at: string;
}

export interface BatchItemRow {
  id: string;
  batch_id: string;
  uuid_fiscal: string;
  status: string;
  nota: string | null;
}

export interface BatchRow {
  id: string;
  titulo: string;
  status: string;
  nota_final: string | null;
  created_at: string;
  assigned_email: string;
  items: BatchItemRow[];
  events: BatchEventRow[];
}

export const ITEM_STATUSES: { value: string; label: string; consolidable: boolean }[] = [
  { value: "PENDIENTE", label: "Pendiente", consolidable: false },
  { value: "PENDIENTE_TERCIERO", label: "Pendiente por tercero", consolidable: false },
  { value: "SUBIDA_SISTEMA", label: "Subida al sistema", consolidable: true },
  { value: "NO_CORRESPONDE", label: "No corresponde", consolidable: false },
];

export const BATCH_STATUS_LABEL: Record<string, string> = {
  PENDIENTE: "Pendiente",
  EN_PROCESO: "En proceso",
  COMPLETADO: "Completado",
};

export async function fetchBatches(
  supabase: SupabaseClient,
  onlyUserId: string | null
): Promise<BatchRow[]> {
  const { data: batchRows, error } = await supabase
    .from("batches")
    .select("id, titulo, status, nota_final, created_at, assigned_to")
    .order("created_at", { ascending: false });
  if (error) throw new Error(error.message);

  let list = (batchRows ?? []) as {
    id: string;
    titulo: string;
    status: string;
    nota_final: string | null;
    created_at: string;
    assigned_to: string | null;
  }[];
  if (onlyUserId) list = list.filter((b) => b.assigned_to === onlyUserId);
  if (list.length === 0) return [];

  const ids = list.map((b) => b.id);
  const { data: itemRows } = await supabase
    .from("batch_items")
    .select("id, batch_id, uuid_fiscal, status, nota")
    .in("batch_id", ids)
    .order("created_at", { ascending: true });
  const { data: eventRows } = await supabase
    .from("batch_events")
    .select("id, batch_id, item_id, actor_email, tipo, status_nuevo, detalle, created_at")
    .in("batch_id", ids)
    .order("created_at", { ascending: true });
  const { data: profileRows } = await supabase.from("profiles").select("id, email");

  const emailById: Record<string, string> = {};
  for (const p of (profileRows ?? []) as { id: string; email: string }[]) {
    emailById[p.id] = p.email;
  }

  return list.map((b) => ({
    id: b.id,
    titulo: b.titulo,
    status: b.status,
    nota_final: b.nota_final,
    created_at: b.created_at,
    assigned_email: b.assigned_to ? emailById[b.assigned_to] ?? "—" : "—",
    items: ((itemRows ?? []) as BatchItemRow[]).filter((i) => i.batch_id === b.id),
    events: ((eventRows ?? []) as BatchEventRow[]).filter((e) => e.batch_id === b.id),
  }));
}

export async function addEvent(
  supabase: SupabaseClient,
  batchId: string,
  itemId: string | null,
  actorEmail: string,
  tipo: string,
  statusNuevo: string | null,
  detalle: string
): Promise<void> {
  const { error } = await supabase.from("batch_events").insert({
    batch_id: batchId,
    item_id: itemId,
    actor_email: actorEmail,
    tipo,
    status_nuevo: statusNuevo,
    detalle,
  });
  if (error) throw new Error(error.message);
}

export async function setItemStatus(
  supabase: SupabaseClient,
  item: BatchItemRow,
  newStatus: string,
  comment: string,
  actorEmail: string
): Promise<void> {
  const { error } = await supabase
    .from("batch_items")
    .update({ status: newStatus, nota: comment !== "" ? comment : item.nota })
    .eq("id", item.id);
  if (error) throw new Error(error.message);
  await addEvent(supabase, item.batch_id, item.id, actorEmail, "STATUS_ITEM", newStatus, comment);
}

export async function startBatch(
  supabase: SupabaseClient,
  batchId: string,
  actorEmail: string
): Promise<void> {
  const { error } = await supabase
    .from("batches")
    .update({ status: "EN_PROCESO" })
    .eq("id", batchId);
  if (error) throw new Error(error.message);
  await addEvent(supabase, batchId, null, actorEmail, "STATUS_BATCH", "EN_PROCESO", "Inicio de atención del batch");
}

export async function finalizeBatch(
  supabase: SupabaseClient,
  batchId: string,
  comment: string,
  actorEmail: string
): Promise<void> {
  const { error } = await supabase
    .from("batches")
    .update({ status: "COMPLETADO", nota_final: comment })
    .eq("id", batchId);
  if (error) throw new Error(error.message);
  await addEvent(supabase, batchId, null, actorEmail, "STATUS_BATCH", "COMPLETADO", comment);
}