'use client'

import { useState } from 'react'

export default function DashboardPage() {
  const [mainFile, setMainFile] = useState<File | null>(null)
  const [complementaryFiles, setComplementaryFiles] = useState<File[]>([])
  const [result, setResult] = useState<any>(null)
  const [loading, setLoading] = useState(false)

  const handleMainFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setMainFile(e.target.files[0])
    }
  }

  const handleComplementaryChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setComplementaryFiles(Array.from(e.target.files))
    }
  }

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    if (!mainFile) return alert('Selecciona el archivo principal')

    const formData = new FormData()
    formData.append('mainFile', mainFile)
    complementaryFiles.forEach(file => {
      formData.append('complementaryFiles', file)
    })

    setLoading(true)
    try {
      const res = await fetch('/api/upload', {
        method: 'POST',
        body: formData,
      })
      const data = await res.json()
      setResult(data)
    } catch (error) {
      console.error(error)
      alert('Error al procesar los archivos')
    } finally {
      setLoading(false)
    }
  }

  return (
    <main className="min-h-screen p-8 bg-gray-50">
      <h1 className="text-2xl font-bold mb-6">Dashboard - Conciliador de Facturas</h1>
      <form onSubmit={handleSubmit} className="space-y-4 max-w-2xl">
        <div>
          <label className="block text-sm font-medium text-gray-700">Archivo Principal (Facturas a buscar)</label>
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={handleMainFileChange}
            className="mt-1 block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
            required
          />
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700">Archivos Complementarios (hasta 4)</label>
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            multiple
            onChange={handleComplementaryChange}
            className="mt-1 block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-green-50 file:text-green-700 hover:file:bg-green-100"
          />
          <p className="text-xs text-gray-500 mt-1">Selecciona hasta 4 archivos (mantén Ctrl/Cmd para múltiple)</p>
        </div>
        <button
          type="submit"
          disabled={loading}
          className="bg-blue-600 text-white py-2 px-6 rounded-md hover:bg-blue-700 disabled:opacity-50"
        >
          {loading ? 'Procesando...' : 'Ejecutar Conciliación'}
        </button>
      </form>

      {result && (
        <div className="mt-8">
          <h2 className="text-xl font-semibold mb-2">Resultado (vista previa)</h2>
          <pre className="bg-white p-4 rounded shadow overflow-auto max-h-96 text-sm">
            {JSON.stringify(result, null, 2)}
          </pre>
        </div>
      )}
    </main>
  )
}