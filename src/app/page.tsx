import { redirect } from "next/navigation";
import { createClient } from "@/lib/supabase/server";
import LogoutButton from "@/components/logout-button";

export const dynamic = "force-dynamic";

export default async function HomePage() {
  const supabase = await createClient();
  const { data, error } = await supabase.auth.getUser();

  if (error || !data.user) {
    redirect("/login");
  }

  return (
    <main className="min-h-screen bg-gray-100">
      <header className="bg-white border-b border-gray-200">
        <div className="max-w-5xl mx-auto px-6 py-4 flex items-center justify-between">
          <h1 className="font-bold text-gray-900">Conciliador de Facturas</h1>
          <div className="flex items-center gap-4">
            <span className="text-sm text-gray-600">{data.user.email}</span>
            <LogoutButton />
          </div>
        </div>
      </header>

      <section className="max-w-5xl mx-auto px-6 py-10">
        <div className="bg-white border border-gray-200 rounded-xl p-8">
          <span className="inline-flex items-center rounded-full bg-green-100 text-green-700 text-xs font-semibold px-3 py-1">
            Conexión exitosa
          </span>
          <h2 className="text-2xl font-bold text-gray-900 mt-4">
            Fase 1 completada
          </h2>
          <p className="text-gray-600 mt-2 max-w-xl text-sm leading-relaxed">
            La autenticación con Supabase funciona correctamente. En la Fase 2
            agregaremos la carga de archivos Excel y el motor de conciliación.
          </p>
        </div>
      </section>
    </main>
  );
}