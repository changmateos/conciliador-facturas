import { NextRequest, NextResponse } from 'next/server'
import * as XLSX from 'xlsx'

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData()
    const mainFile = formData.get('mainFile') as File
    const complementaryFiles = formData.getAll('complementaryFiles') as File[]

    if (!mainFile) {
      return NextResponse.json({ error: 'Archivo principal no proporcionado' }, { status: 400 })
    }

    // Leer archivo principal
    const mainBuffer = Buffer.from(await mainFile.arrayBuffer())
    const mainWorkbook = XLSX.read(mainBuffer, { type: 'buffer' })
    const mainData: any[] = []

    // Procesar ambas hojas (SIF y NO SIF)
    mainWorkbook.SheetNames.forEach(sheetName => {
      const sheet = mainWorkbook.Sheets[sheetName]
      const json = XLSX.utils.sheet_to_json(sheet)
      // Mapear columnas
      const mapped = json.map((row: any) => ({
        uuid: row['UUID'] || row['uuid'] || '',
        folio: row['Folio'] || row['folio'] || '',
        total: parseFloat(row['Total'] || row['total'] || 0),
        retenciones: parseFloat(row['Total Imp Ret'] || row['totalImpRet'] || 0),
        fecha: row['Fecha Emisión'] || '',
        contabilidadId: row['SIF IdContabilidad'] || '',
        ambito: row['Ambito'] || '',
        // Guardar también la hoja de origen
        _sheet: sheetName,
        _raw: row
      }))
      mainData.push(...mapped)
    })

    // Leer archivos complementarios (solo procesamos el primero para mostrar ejemplo)
    const complementaryData: any[] = []
    for (const file of complementaryFiles) {
      const buffer = Buffer.from(await file.arrayBuffer())
      const workbook = XLSX.read(buffer, { type: 'buffer' })
      workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName]
        const json = XLSX.utils.sheet_to_json(sheet)
        // Aquí asumo que los complementarios tienen al menos UUID y/o Folio
        const mapped = json.map((row: any) => ({
          uuid: row['UUID'] || row['uuid'] || '',
          folio: row['Folio'] || row['folio'] || '',
          total: parseFloat(row['Total'] || row['total'] || 0),
          retenciones: parseFloat(row['Total Imp Ret'] || row['totalImpRet'] || 0),
          _sheet: sheetName,
          _raw: row,
          _file: file.name
        }))
        complementaryData.push(...mapped)
      })
    }

    return NextResponse.json({
      mainFile: {
        name: mainFile.name,
        totalRows: mainData.length,
        sample: mainData.slice(0, 10) // muestra primeros 10
      },
      complementaryFiles: complementaryFiles.map(f => f.name),
      complementaryTotalRows: complementaryData.length,
      complementarySample: complementaryData.slice(0, 10)
    })

  } catch (error: any) {
    console.error(error)
    return NextResponse.json({ error: error.message }, { status: 500 })
  }
}