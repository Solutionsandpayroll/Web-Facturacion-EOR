import { useState, useRef, useEffect } from 'react'
import * as XLSX from 'xlsx'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import './App.css'

const CLIENTES = [...new Set([
  'C1007 - NZD',
  'C1024 - MKD',
  'C1032 - FLEXCO',
  'C1038 - ONCEHUB',
  'C1041 - EDRINGTON',
  'C1042 - EPDM',
  'C1043 - NEO',
  'C1050 - HEMMERSBACH',
  'C1051 - YONYOU',
  'C1052 - BUBBLE BPM INC',
  'C1053 - GLOBAL EXPANSION',
  'C1055 - RIVERMATE',
  'C1037 - REMOFIRST',
  'C1029 - INSIDER',
  'C1036 - ACTION AD',
  'C1056 - EUROPORTAGE',
  'C1022 - Root Capital',
  'C1058 - POC PHARMA',
  'C1059 - SIFFI',
])]

// fee: coeficiente para la fórmula FEE (porcentaje como '6%' o valor fijo como '120')
// iva: true = fórmula =$BC{fila}*19%, false = vacío
// banking: true = fórmula =SUMA($AF{fila};$AN{fila};$AS{fila};$AZ{fila})*0,4%, false = vacío
const CONFIG_CLIENTES = {
  'C1007 - NZD':               { fee: '6%',     iva: false, banking: false },
  'C1024 - MKD':               { fee: '11%',    iva: false, banking: false },
  'C1032 - FLEXCO':            { fee: '9%',     iva: false, banking: false },
  'C1038 - ONCEHUB':           { fee: '100.84', iva: true,  banking: true  },
  'C1041 - EDRINGTON':         { fee: '5.5%',   iva: false, banking: false },
  'C1042 - EPDM':              { fee: '10%',    iva: false, banking: false },
  'C1043 - NEO':               { fee: '8%',     iva: false, banking: false },
  'C1050 - HEMMERSBACH':       { fee: '10%',    iva: false, banking: false },
  'C1051 - YONYOU':            { fee: '210',    iva: true,  banking: true  },
  'C1052 - BUBBLE BPM INC':    { fee: '11%',    iva: false, banking: false },
  'C1053 - GLOBAL EXPANSION':  { fee: '150',    iva: true,  banking: true  },
  'C1055 - RIVERMATE':         { fee: '150',    iva: false, banking: true  },
  'C1037 - REMOFIRST':         { fee: '120',    iva: true,  banking: true  },
  'C1029 - INSIDER':           { fee: '190',    iva: true,  banking: true  },
  'C1036 - ACTION AD':         { fee: '11%',    iva: false, banking: false },
  'C1056 - EUROPORTAGE':       { fee: '200',    iva: true,  banking: true  },
  'C1022 - Root Capital':      { fee: '9%',     iva: false, banking: false },
  'C1058 - POC PHARMA':        { fee: '160',    iva: false, banking: true  },
  'C1059 - SIFFI':             { fee: '10%',    iva: false, banking: true  },
}

function App() {
  const [isHelpExpanded, setIsHelpExpanded] = useState(false)
  const [baseEmpleados, setBaseEmpleados] = useState(null)
  const [reporteNovasoft, setReporteNovasoft] = useState(null)
  const [dragBase, setDragBase] = useState(false)
  const [dragNova, setDragNova] = useState(false)
  const inputBase = useRef(null)
  const inputNova = useRef(null)
  const [clientesSeleccionados, setClientesSeleccionados] = useState([])
  const [dropdownOpen, setDropdownOpen] = useState(false)
  const [busqueda, setBusqueda] = useState('')
  const dropdownRef = useRef(null)
  const [periodo, setPeriodo] = useState('')
  const [generando, setGenerando] = useState(false)
  const [error, setError] = useState(null)
  const [exitoCount, setExitoCount] = useState(0)

  useEffect(() => {
    const handleClickOutside = (e) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target)) {
        setDropdownOpen(false)
      }
    }
    document.addEventListener('mousedown', handleClickOutside)
    return () => document.removeEventListener('mousedown', handleClickOutside)
  }, [])

  const toggleCliente = (cliente) => {
    setClientesSeleccionados(prev =>
      prev.includes(cliente) ? prev.filter(c => c !== cliente) : [...prev, cliente]
    )
  }

  const clientesFiltrados = CLIENTES.filter(c =>
    c.toLowerCase().includes(busqueda.toLowerCase())
  )

  const formatSize = (bytes) => {
    if (bytes < 1024) return bytes + ' B'
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB'
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB'
  }

  const handleFileDrop = (e, setter, setDrag) => {
    e.preventDefault()
    setDrag(false)
    const file = e.dataTransfer.files[0]
    if (file) setter(file)
  }

  const handleFileChange = (e, setter) => {
    const file = e.target.files[0]
    if (file) setter(file)
  }

  const generar = async () => {
    setError(null)
    setExitoCount(0)

    if (!baseEmpleados) return setError('Sube el archivo de base de empleados.')
    if (!reporteNovasoft) return setError('Sube el archivo de reporte Novasoft.')
    if (clientesSeleccionados.length === 0) return setError('Selecciona al menos un cliente.')
    if (!periodo.trim()) return setError('Escribe el periodo de facturación.')

    setGenerando(true)
    try {
      // --- 1. Leer Base de empleados ---
      const baseBuffer = await baseEmpleados.arrayBuffer()
      const baseWb = XLSX.read(baseBuffer, { type: 'array' })
      const baseSheet = baseWb.Sheets[baseWb.SheetNames[0]]
      const baseData = XLSX.utils.sheet_to_json(baseSheet, { header: 1 })

      const headerBase = (baseData[0] || []).map(h => String(h ?? '').trim().toUpperCase())
      const colEmp = headerBase.indexOf('CODIGO EMPLEADO')
      const colCC  = headerBase.indexOf('CENTRO COSTOS')

      console.log('=== BASE DE EMPLEADOS ===')
      console.log('Encabezados encontrados:', headerBase)
      console.log('Columna CODIGO EMPLEADO (índice):', colEmp)
      console.log('Columna CENTRO COSTOS (índice):', colCC)
      console.log('Total filas (incluyendo encabezado):', baseData.length)

      if (colEmp === -1) throw new Error('No se encontró la columna "CODIGO EMPLEADO" en la base de empleados.')
      if (colCC  === -1) throw new Error('No se encontró la columna "CENTRO COSTOS" en la base de empleados.')

      // Mapa: códigoEmpleado → centroCostos
      const mapaCC = new Map()
      for (let i = 1; i < baseData.length; i++) {
        const codigo = String(baseData[i][colEmp] ?? '').trim()
        const cc     = String(baseData[i][colCC]  ?? '').trim()
        if (codigo) mapaCC.set(codigo, cc)
      }

      console.log('Primeras 10 entradas del mapa código→CC:',
        [...mapaCC.entries()].slice(0, 10)
      )

      // --- 2. Leer Reporte Novasoft → hoja "nomplcol" ---
      const novaBuffer = await reporteNovasoft.arrayBuffer()
      const novaWb = XLSX.read(novaBuffer, { type: 'array', raw: false })

      console.log('=== NOVASOFT ===')
      console.log('Hojas disponibles:', novaWb.SheetNames)

      const novaSheet = novaWb.Sheets['nomplcol']
      if (!novaSheet) throw new Error('No se encontró la hoja "nomplcol" en el reporte Novasoft.')

      const novaData = XLSX.utils.sheet_to_json(novaSheet, { header: 1, raw: false, defval: '' })

      // Fila 22 = índice 21 → encabezados principales. Buscar columna "Codigo Empl"
      const headerNova = (novaData[21] || []).map(h => String(h ?? '').trim())
      console.log('Encabezados fila 22 (nomplcol):', headerNova)
      const colCodigo = headerNova.findIndex(h => h.toUpperCase().includes('CODIGO') && h.toUpperCase().includes('EMPL'))
      console.log('Columna "Codigo Empl" (índice):', colCodigo)
      if (colCodigo === -1) throw new Error('No se encontró la columna "Codigo Empl" en la fila 22 de la hoja nomplcol.')

      // Fila 23 = índice 22 → sub-encabezados de conceptos
      const conceptosNova = (novaData[22] || []).map(h => String(h ?? '').trim())
      console.log('Sub-encabezados fila 23 (conceptos):', conceptosNova.filter(Boolean))

      // Definición de grupos de conceptos → columna destino en plantilla
      const GRUPOS = [
        {
          columnaDestino: 'SALARY',
          conceptos: [
            '000050 Días hábiles en Vacacio',
            '000051 Días No hábiles en Vaca',
            '001050 Salario',
            '001051 Salario Integral',
            '001055 Descanso Remunerado',
            '001053 Apoyo Sostenimiento',
            '100322 Salario Ord Retroactivo',
            '100321 Sal Integral Retroactiv',
            '001070 Descanso compensatorio',
            '001090 Pago Descanso Dominical',
            '001092 Pago Dom Liquidacion',
            '001170 Licencia Remunerada',
            '001190 Honorarios',
            '100102 Sobresueldo',
            '100103 Salario Variable',
            '100101 Salario Pendiente',
            '101171 Licencia Remunerada x h',
          ],
        },
        {
          columnaDestino: '"Sick" Leave',
          conceptos: [
            '001150 Inc. por Enfermedad Com',
            '001151 Inc. Enf Comun Asumida',
            '001160 Incapacidad Acc Trabajo',
            '001161 Inc. Acc Trabajo Asumid',
            '001177 Aux. Inc. Enfermedad Co',
            '001178 Factor Pres Inc. Sal In',
          ],
        },
        {
          columnaDestino: '13th Salary',
          conceptos: [
            '008407 Provision Prima',
          ],
        },
        {
          columnaDestino: '14th Salary',
          conceptos: [
            '008412 Provision Cesantias',
          ],
        },
        {
          columnaDestino: 'Alloawance 1 (Car Allowance)',
          conceptos: [
            '001430 Aux Extralegal Transpor',
            '101314 Auxilio de Transporte',
            '100207 Auxilio Movilizacion',
            '101337 Aux Extralegal de Auto',
          ],
        },
        {
          columnaDestino: 'Alloawance 2 (Mobile & Internet Allowance)',
          conceptos: [
            '101308 Auxilio de celular',
            '100213 Auxilio Conectividad',
            '101306-Auxilio Telecomunicacio',
            '101335 Auxilio Telefonia',
          ],
        },
        {
          columnaDestino: 'Alloawance 4 (Other allowances)',
          conceptos: [
            '101315 Gross Up',
            '100205 Auxilio Vivienda',
            '100209 Aux Extralegal No Sal',
            '100206 Auxilio AFC',
            '101750 Viaticos No Salariales',
            '101336 Aux Extralegal No Sal',
            '101338 AuxExtralegal por Edu',
          ],
        },
        {
          columnaDestino: 'Bonus/Commission',
          conceptos: [
            '001072 Comisiones',
            '001073 Comisiones Salario Inte',
            '100201 Bonificacion No Salaria',
            '111072 Comisiones Honorarios',
            '001751-Bonificacion  Salarial',
            '100212 Prima de Extralegal',
            '001750 BonExtralegal Liberidad',
          ],
        },
        {
          columnaDestino: 'Deduction or Gross Amount adjustments prevous month',
          conceptos: [
            '200050 Descuento Anticipo de N',
            '200007 Descuento Autorizado',
            '200006 Descuento Prestamo',
          ],
        },
        {
          columnaDestino: 'Expenses',
          conceptos: [
            '300002 Rembolso de Gastos',
            '300101 Gastos Tarjeta Credito',
          ],
        },
        {
          columnaDestino: 'Family Fund Cost',
          conceptos: [
            '002910 Aporte Caja Compensació',
          ],
        },
        {
          columnaDestino: 'Food allowance',
          conceptos: [
            '001401 Auxilio de Alimentacion',
          ],
        },
        {
          columnaDestino: 'Health Cost',
          conceptos: [
            '002220 Salud Patrono',
          ],
        },
        {
          columnaDestino: 'Health Insurance',
          conceptos: [
            '300003 Poliza Med Sura',
          ],
        },
        {
          columnaDestino: 'Home/Remote work allowance',
          conceptos: [
            '101319 Auxilio de Computador',
          ],
        },
        {
          columnaDestino: 'ICBF cost',
          conceptos: [
            '002915 Aportes I.C.B.F.',
          ],
        },
        {
          columnaDestino: 'Interest on 14th Salary',
          conceptos: [
            '008415 Provision Int Ces',
          ],
        },
        {
          columnaDestino: 'Labor Risk Cost',
          conceptos: [
            '002222 Riesgo Profesional',
          ],
        },
        {
          columnaDestino: 'Medical Test',
          conceptos: [
            '300009 Gasto Exa Medico',
            'MEDICAL TEST',
          ],
        },
        {
          columnaDestino: 'On Call/ Plus Disponibilidad',
          conceptos: [
            '101751 Standby Salarial',
          ],
        },
        {
          columnaDestino: 'Overtime',
          conceptos: [
            '001061 Horas Extras Diurnas',
            '001062 Horas Extras Nocturnas',
            '001063 Hora Extra Dom o Fes',
            '001064 Hora Extra Dom Fes Noct',
            '001066 Recargo Noct Dom Fes',
            '001067 Hora Dom o Fest Ordinar',
          ],
        },
        {
          columnaDestino: 'Paternity/ Maternity leave',
          conceptos: [
            '001153 Licencia Paternidad',
            '001154 Aux.Lic. Paternidad',
            '001155 Licencia Maternidad',
            '001156 Aux.Lic. Maternidad',
          ],
        },
        {
          columnaDestino: 'Pension Cost',
          conceptos: [
            '002221 Pensión Patrono',
          ],
        },
        {
          columnaDestino: 'Rectroactive payment/Plus Compensation',
          conceptos: [
            '101050 Retro Apoyo Sostenimien',
          ],
        },
        {
          columnaDestino: 'SENA Cost',
          conceptos: [
            '002905 Apropiación SENA',
          ],
        },
        {
          columnaDestino: 'Severance Pay (Taxable)',
          conceptos: [
            '001600 Indemnización',
            '101600 Suma Transaccional',
            '100203 Bono Retiro No Salarial',
            '001610 Bonificación Por Retiro',
          ],
        },
        {
          columnaDestino: 'Sign-on Bonus',
          conceptos: [
            '100202 Bonificacion Firma',
          ],
        },
        {
          columnaDestino: 'Transport allowance',
          conceptos: [
            '001310 Auxilio Conectividad',
            '001300 Subsidio de Transporte',
            '101300 Retro Subsidio Transpor',
          ],
        },
        {
          columnaDestino: 'Unused Holidays',
          conceptos: [
            '001145 Vacaciones en Dinero',
            '001146 Vacaciones en Liq de Co',
            '101146 Vac Adicion LiqContrato',
          ],
        },
        {
          columnaDestino: 'Wellness Allowance',
          conceptos: [
            '101304 Auxilio de Bienestar',
            '101302 Auxilio de salud',
            '101307 Auxilio Gym',
            '101317 Auxilio Medicina Prepag',
            '101318 Auxilio Poliza de Vida',
          ],
        },
      ]

      // Para cada grupo, encontrar los índices de columna en fila 23
      const gruposConIndices = GRUPOS.map(grupo => ({
        columnaDestino: grupo.columnaDestino,
        indices: grupo.conceptos
          .map(concepto => {
            const idx = conceptosNova.findIndex(c =>
              c.trim().toUpperCase() === concepto.trim().toUpperCase()
            )
            if (idx === -1) return idx // concepto no presente en este reporte, se ignora
            return idx
          })
          .filter(idx => idx !== -1),
      }))

      console.log('=== DIAGNÓSTICO CONCEPTOS ===')
      console.log('Fila 23 completa (conceptosNova):', conceptosNova)
      console.log('Gruposcon índices resueltos:', JSON.stringify(gruposConIndices))
      // Muestra la fila 25 completa para ver valores reales
      console.log('Fila 25 (primer dato):', novaData[24])
      console.log('Fila 26:', novaData[25])

      // Construir mapa: codigoEmpleado → { SALARY: número, ... }
      // Datos desde fila 25 = índice 24
      const mapaValores = new Map() // código → { grupoKey: suma }

      const codigosNova = []
      for (let i = 24; i < novaData.length; i++) {
        const fila = novaData[i]
        const codigo = String(fila[colCodigo] ?? '').trim()
        if (!codigo) continue
        if (!codigosNova.includes(codigo)) codigosNova.push(codigo)

        if (!mapaValores.has(codigo)) {
          const entry = {}
          gruposConIndices.forEach(g => { entry[g.columnaDestino] = 0 })
          mapaValores.set(codigo, entry)
        }

        const entry = mapaValores.get(codigo)
        gruposConIndices.forEach(grupo => {
          grupo.indices.forEach(idx => {
            const raw = fila[idx]
            const val = parseFloat(String(raw ?? '').replace(/[$,]/g, '')) || 0
            if (codigosNova.indexOf(codigo) < 3 && grupo.indices.indexOf(idx) === 0) {
              console.log(`  [${codigo}] col ${idx} raw="${raw}" → val=${val}`)
            }
            entry[grupo.columnaDestino] += val
          })
        })
      }

      console.log('Primeros 10 códigos de nomplcol:', codigosNova.slice(0, 10))
      console.log('Ejemplo mapaValores (primer empleado):',
        codigosNova[0], mapaValores.get(codigosNova[0])
      )

      // --- 3. Cargar plantilla una sola vez y limpiar shared formulas ---
      const tplResponse = await fetch('/Formato facturacion - final.xlsx')
      if (!tplResponse.ok) throw new Error('No se pudo cargar la plantilla "Formato facturacion - final.xlsx" desde public/.')
      const tplBuffer = await tplResponse.arrayBuffer()

      // Extraer fórmulas de fila 5 con SheetJS ANTES de que JSZip elimine los clones
      // SheetJS resuelve shared formulas correctamente, así recuperamos todas las fórmulas
      const wbTpl = XLSX.read(tplBuffer, { cellFormula: true })
      const wsTpl = wbTpl.Sheets[wbTpl.SheetNames[0]]
      const formulasRow5 = {} // colNumber (1-based) → formula string
      Object.keys(wsTpl).forEach(cellAddr => {
        if (cellAddr.startsWith('!')) return
        const match = cellAddr.match(/^([A-Z]+)(\d+)$/)
        if (!match || parseInt(match[2]) !== 5) return
        const cell = wsTpl[cellAddr]
        if (cell && cell.f) {
          const colNum = XLSX.utils.decode_col(match[1]) + 1 // convierte a 1-based
          formulasRow5[colNum] = cell.f
        }
      })

      // Pre-procesar: eliminar fórmulas compartidas que bloquean a ExcelJS
      const zip = await JSZip.loadAsync(tplBuffer)
      const sheetNames = Object.keys(zip.files).filter(f => f.startsWith('xl/worksheets/') && f.endsWith('.xml'))
      for (const sheetName of sheetNames) {
        let xml = await zip.files[sheetName].async('string')
        // Convertir fórmulas compartidas master → fórmula normal
        xml = xml.replace(/<f t="shared"([^>]*)>([^<]*)<\/f>/g, '<f>$2</f>')
        // Eliminar clones de fórmula compartida (sin contenido)
        xml = xml.replace(/<f t="shared"[^>]*\/>/g, '')
        zip.file(sheetName, xml)
      }
      const cleanTplBuffer = await zip.generateAsync({ type: 'arraybuffer' })

      const periodoLimpio = periodo.trim().replace(/[\/\\?%*:|"<>]/g, '-')
      let archivosGenerados = 0

      // --- 4. Generar un archivo por cada cliente seleccionado ---
      for (const cliente of clientesSeleccionados) {
        // Filtrar códigos para este cliente
        const codigosFiltrados = codigosNova.filter(codigo => {
          const cc = mapaCC.get(codigo)
          return cc && cc.trim() === cliente.trim()
        })

        console.log(`Cliente "${cliente}": ${codigosFiltrados.length} códigos`, codigosFiltrados.slice(0, 5))

        if (codigosFiltrados.length === 0) {
          console.warn(`Sin empleados para ${cliente}, se omite.`)
          continue
        }

        // Cargar plantilla fresca (ya limpia) para cada cliente
        const workbook = new ExcelJS.Workbook()
        await workbook.xlsx.load(cleanTplBuffer.slice(0))
        const worksheet = workbook.worksheets[0]

        // Encontrar columna "EMPLOYEE CODE" y columnas destino de grupos en fila 3
        const headerRow = worksheet.getRow(3)
        let colEC = -1
        // mapa: nombreColumnaDestino → número de columna en plantilla
        const colsDestino = {}
        gruposConIndices.forEach(g => { colsDestino[g.columnaDestino] = -1 })

        // Columnas TOTAL: col → { formula: string, style: object }
        const colsTotal = {}
        // Columnas con fórmula variable por cliente
        let colFee = -1, colBanking = -1, colIva = -1

        headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const val = String(cell.value ?? '').trim().toUpperCase()
          if (val === 'EMPLOYEE CODE') colEC = colNumber
          gruposConIndices.forEach(g => {
            if (val === g.columnaDestino.toUpperCase()) colsDestino[g.columnaDestino] = colNumber
          })
          // Columnas de fórmula por cliente (manejadas aparte, NO en colsTotal)
          if (val === 'FEE')         { colFee     = colNumber; return }
          if (val === 'BANKING TAX') { colBanking  = colNumber; return }
          if (val === 'IVA')         { colIva      = colNumber; return }
          // Resto de columnas con fórmula replicada genéricamente
          if (['TOTAL', 'PAYMENTS', 'FEE USD', 'VAT'].some(kw => val.includes(kw))) {
            colsTotal[colNumber] = null // se rellena abajo con la fórmula de fila 4
          }
        })

        console.log('colEC:', colEC, '| colsDestino:', colsDestino)
        if (colEC === -1) throw new Error('No se encontró la columna "EMPLOYEE CODE" en la fila 3 de la plantilla.')

        // Capturar fórmulas y estilos de fila 4 para columnas TOTAL
        Object.keys(colsTotal).forEach(colNum => {
          const c = worksheet.getRow(4).getCell(Number(colNum))
          const v = c.value
          let formula = null
          if (v && typeof v === 'object' && v.formula) formula = v.formula
          else if (typeof v === 'string' && v.startsWith('=')) formula = v.slice(1)
          colsTotal[colNum] = {
            formula,
            style: c.style ? JSON.parse(JSON.stringify(c.style)) : {},
          }
        })

        // Capturar estilos de referencia (fila 4) para cada columna que se va a escribir
        const estilosRef = {}
        const celdaRefEC = worksheet.getRow(4).getCell(colEC)
        estilosRef[colEC] = celdaRefEC.style ? JSON.parse(JSON.stringify(celdaRefEC.style)) : {}
        gruposConIndices.forEach(g => {
          const col = colsDestino[g.columnaDestino]
          if (col !== -1) {
            const c = worksheet.getRow(4).getCell(col)
            estilosRef[col] = c.style ? JSON.parse(JSON.stringify(c.style)) : {}
          }
        })
        // Capturar estilos fila 4 para columnas de fórmula por cliente
        ;[colFee, colBanking, colIva].forEach(col => {
          if (col !== -1) {
            const c = worksheet.getRow(4).getCell(col)
            estilosRef[col] = c.style ? JSON.parse(JSON.stringify(c.style)) : {}
          }
        })

        // Guardar contenido completo de la fila 5 (fórmulas de totales) antes de limpiar
        // Usamos SheetJS como fuente de verdad para las fórmulas (ExcelJS pierde los clones)
        const filaFormulas = []
        const row5 = worksheet.getRow(5)
        row5.eachCell({ includeEmpty: true }, (cell, colNum) => {
          // Obtener fórmula: primero ExcelJS, si no tiene usar SheetJS como fallback
          let value = cell.value
          const formulaSheetJS = formulasRow5[colNum]
          if (formulaSheetJS) {
            // Preferir siempre la fórmula de SheetJS (garantiza tenerla aunque ExcelJS la haya perdido)
            value = { formula: formulaSheetJS }
          }
          filaFormulas[colNum] = {
            value,
            style: cell.style ? JSON.parse(JSON.stringify(cell.style)) : {},
          }
        })
        // Añadir celdas de formulasRow5 que eachCell no visitó (celdas vacías no iteradas)
        Object.entries(formulasRow5).forEach(([colNum, formula]) => {
          const col = Number(colNum)
          if (!filaFormulas[col]) {
            const cell = row5.getCell(col)
            filaFormulas[col] = {
              value: { formula },
              style: cell.style ? JSON.parse(JSON.stringify(cell.style)) : {},
            }
          }
        })

        // Limpiar la fila 5 original por completo (valores + estilos) en TODAS las columnas
        // para que no queden restos cuando los empleados la sobreescriban parcialmente
        const totalCols = worksheet.columnCount || filaFormulas.length
        for (let colNum = 1; colNum <= totalCols; colNum++) {
          const c = row5.getCell(colNum)
          c.value = null
          c.style = {}
        }
        row5.commit()

        // Limpiar datos previos desde fila 4 en todas las columnas destino
        const todasLasCols = [colEC, ...Object.values(colsDestino).filter(c => c !== -1)]
        for (let r = 4; r <= worksheet.rowCount; r++) {
          todasLasCols.forEach(col => { worksheet.getRow(r).getCell(col).value = null })
        }

        // Escribir datos desde fila 4
        codigosFiltrados.forEach((codigo, i) => {
          const row = worksheet.getRow(4 + i)
          const valores = mapaValores.get(codigo) || {}

          // Employee code
          const cellEC = row.getCell(colEC)
          cellEC.value = codigo
          if (estilosRef[colEC]) cellEC.style = JSON.parse(JSON.stringify(estilosRef[colEC]))

          // Valores de cada grupo
          gruposConIndices.forEach(g => {
            const col = colsDestino[g.columnaDestino]
            if (col === -1) return
            const cell = row.getCell(col)
            const suma = valores[g.columnaDestino] ?? 0
            cell.value = suma !== 0 ? parseFloat(suma.toFixed(2)) : null
            if (estilosRef[col]) cell.style = JSON.parse(JSON.stringify(estilosRef[col]))
          })

          // Replicar fórmulas de columnas TOTAL ajustando referencias de fila
          const filaActual = 4 + i
          Object.entries(colsTotal).forEach(([colNum, data]) => {
            if (!data || !data.formula) return
            // Reemplaza referencias a fila 4 por la fila actual (ej. B4 → B7)
            const formulaAdaptada = data.formula.replace(
              /([A-Za-z]+)(\d+)/g,
              (match, col, row) => parseInt(row) === 4 ? `${col}${filaActual}` : match
            )
            const cell = row.getCell(Number(colNum))
            cell.value = { formula: formulaAdaptada }
            if (data.style) cell.style = JSON.parse(JSON.stringify(data.style))
          })

          // Fórmulas específicas por cliente: FEE, BANKING TAX, IVA
          const cfg = CONFIG_CLIENTES[cliente] || {}
          if (colFee !== -1) {
            const cell = row.getCell(colFee)
            if (cfg.fee) {
              cell.value = { formula: `${cfg.fee}*$BG${filaActual}` }
              if (estilosRef[colFee]) cell.style = JSON.parse(JSON.stringify(estilosRef[colFee]))
            } else {
              cell.value = null
            }
          }
          if (colBanking !== -1) {
            const cell = row.getCell(colBanking)
            if (cfg.banking) {
              cell.value = { formula: `SUM($AF${filaActual},$AN${filaActual},$AS${filaActual},$AZ${filaActual})*0.004` }
              if (estilosRef[colBanking]) cell.style = JSON.parse(JSON.stringify(estilosRef[colBanking]))
            } else {
              cell.value = null
            }
          }
          if (colIva !== -1) {
            const cell = row.getCell(colIva)
            if (cfg.iva) {
              cell.value = { formula: `$BC${filaActual}*0.19` }
              if (estilosRef[colIva]) cell.style = JSON.parse(JSON.stringify(estilosRef[colIva]))
            } else {
              cell.value = null
            }
          }

          row.commit()
        })

        // Restaurar la fila de fórmulas al final de los datos, actualizando rangos dinámicamente
        const filaTotal = 4 + codigosFiltrados.length
        const ultimaFilaDatos = filaTotal - 1  // última fila con empleados

        // Reescribe rangos tipo "B4:B4" o "B4:B5" para que cubran todos los datos
        const actualizarRango = (formula) => {
          return String(formula).replace(
            /([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)/g,
            (match, col1, row1, col2, row2) => {
              if (parseInt(row1) === 4) {
                return `${col1}${row1}:${col2}${ultimaFilaDatos}`
              }
              return match
            }
          )
        }

        const rowFormulas = worksheet.getRow(filaTotal)
        filaFormulas.forEach((data, colNum) => {
          if (!data) return
          const cell = rowFormulas.getCell(colNum)
          let valor = data.value
          // Si la celda tiene fórmula, actualizar el rango para abarcar todos los empleados
          if (valor && typeof valor === 'object' && valor.formula) {
            valor = { formula: actualizarRango(valor.formula) }
          } else if (typeof valor === 'string' && valor.startsWith('=')) {
            valor = { formula: actualizarRango(valor.slice(1)) }
          }
          cell.value = valor
          if (data.style) cell.style = JSON.parse(JSON.stringify(data.style))
        })
        rowFormulas.commit()

        // Descargar: nombre = "Facturación EOR - {Cliente} - {Periodo}.xlsx"
        const clienteLimpio = cliente.replace(/[\/\\?%*:|"<>]/g, '-')
        const outBuffer = await workbook.xlsx.writeBuffer()
        const blob = new Blob([outBuffer], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        })
        const url = URL.createObjectURL(blob)
        const a = document.createElement('a')
        a.href = url
        a.download = `Facturación EOR - ${clienteLimpio} - ${periodoLimpio}.xlsx`
        document.body.appendChild(a)
        a.click()
        document.body.removeChild(a)
        URL.revokeObjectURL(url)
        archivosGenerados++

        // Pequeña pausa entre descargas para que el navegador no las bloquee
        await new Promise(r => setTimeout(r, 400))
      }

      if (archivosGenerados === 0)
        throw new Error('No se encontraron empleados para ninguno de los clientes seleccionados. Verifica los archivos.')

      setExitoCount(archivosGenerados)
      setTimeout(() => setExitoCount(0), 5000)

    } catch (err) {
      setError(err.message)
    } finally {
      setGenerando(false)
    }
  }

  return (
    <div className="app">
      {/* Header Corporativo Solutions & Payroll */}
      <header className="header">
        <div className="container">
          <div className="header-content">
            <div className="logo-container">
              <div className="logo">
                <img 
                  src="/Logo syp.png" 
                  alt="Solutions & Payroll Logo" 
                  width="60" 
                  height="60"
                />
              </div>
              <div className="header-text">
                <h1>Solutions & Payroll</h1>
                <p className="subtitle">Facturación EOR</p>
              </div>
            </div>
            <div className="welcome-box">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/>
                <circle cx="12" cy="7" r="4"/>
              </svg>
              <span>Bienvenido, Usuario</span>
            </div>
          </div>
        </div>
      </header>

      {/* Contenido Principal */}
      <main className="main-content">
        <div className="container">
          
          {/* Sección de ayuda colapsable (opcional - puedes eliminarla si no la necesitas) */}
          <div className="help-section">
            <button 
              className="help-toggle"
              onClick={() => setIsHelpExpanded(!isHelpExpanded)}
              aria-expanded={isHelpExpanded}
            >
              <div className="help-toggle-header">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10"/>
                  <line x1="12" y1="16" x2="12" y2="12"/>
                  <line x1="12" y1="8" x2="12.01" y2="8"/>
                </svg>
                <span>¿Cómo usar esta aplicación?</span>
              </div>
              <svg 
                className={`chevron ${isHelpExpanded ? 'expanded' : ''}`}
                width="20" 
                height="20" 
                viewBox="0 0 24 24" 
                fill="none" 
                stroke="currentColor" 
                strokeWidth="2"
              >
                <polyline points="6 9 12 15 18 9"/>
              </svg>
            </button>
            <div className={`help-content ${isHelpExpanded ? 'expanded' : ''}`}>
              <ol className="help-list">
                <li>
                  <span className="step-number">1</span>
                  <div>
                    <strong>Paso 1</strong>
                    <p>Descripción del primer paso</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">2</span>
                  <div>
                    <strong>Paso 2</strong>
                    <p>Descripción del segundo paso</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">3</span>
                  <div>
                    <strong>Paso 3</strong>
                    <p>Descripción del tercer paso</p>
                  </div>
                </li>
              </ol>
            </div>
          </div>

          {/* Card Principal - Aquí va tu contenido específico */}
          <div className="card">
            <div className="card-header">
              <h2>Facturación EOR</h2>
              <p className="description">
                Sistema de facturación - Solutions & Payroll
              </p>
            </div>

            <div className="card-body">
              <div className="form-section">
                
                {/* Ejemplo de campo de formulario */}
                <div className="form-group">
                  <label className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <rect x="3" y="3" width="18" height="18" rx="2"/>
                      <path d="M9 3v18M15 3v18M3 9h18M3 15h18"/>
                    </svg>
                    Periodo de facturación
                  </label>
                  <input
                    type="text"
                    placeholder="Ej: Enero 2026 — se usará en el nombre de los archivos generados"
                    className="select-input"
                    value={periodo}
                    onChange={e => { setPeriodo(e.target.value); setError(null) }}
                  />
                </div>

                {/* Drop zone: Base de empleados */}
                <div className="form-group">
                  <label className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/>
                      <circle cx="9" cy="7" r="4"/>
                      <path d="M23 21v-2a4 4 0 0 0-3-3.87"/>
                      <path d="M16 3.13a4 4 0 0 1 0 7.75"/>
                    </svg>
                    Base de empleados
                  </label>
                  <input
                    ref={inputBase}
                    type="file"
                    accept=".xlsx,.xls"
                    className="file-input"
                    onChange={(e) => handleFileChange(e, setBaseEmpleados)}
                  />
                  <div
                    className={`drop-zone ${dragBase ? 'drag-active' : ''} ${baseEmpleados ? 'has-file' : ''}`}
                    onClick={() => inputBase.current.click()}
                    onDragOver={(e) => { e.preventDefault(); setDragBase(true) }}
                    onDragLeave={() => setDragBase(false)}
                    onDrop={(e) => handleFileDrop(e, setBaseEmpleados, setDragBase)}
                  >
                    {baseEmpleados ? (
                      <div className="file-preview">
                        <div className="file-icon">
                          <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                            <polyline points="14 2 14 8 20 8"/>
                            <line x1="8" y1="13" x2="16" y2="13"/>
                            <line x1="8" y1="17" x2="16" y2="17"/>
                          </svg>
                        </div>
                        <div className="file-details">
                          <p className="file-name">{baseEmpleados.name}</p>
                          <p className="file-size">{formatSize(baseEmpleados.size)}</p>
                        </div>
                        <button
                          className="btn-remove"
                          onClick={(e) => { e.stopPropagation(); setBaseEmpleados(null); inputBase.current.value = '' }}
                          title="Quitar archivo"
                        >
                          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <line x1="18" y1="6" x2="6" y2="18"/>
                            <line x1="6" y1="6" x2="18" y2="18"/>
                          </svg>
                        </button>
                      </div>
                    ) : (
                      <div className="drop-zone-content">
                        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                          <polyline points="17 8 12 3 7 8"/>
                          <line x1="12" y1="3" x2="12" y2="15"/>
                        </svg>
                        <div className="drop-zone-text">
                          <p className="drop-zone-title">Sube la base de empleados</p>
                          <p className="drop-zone-subtitle">Arrastra aquí o haz clic para seleccionar</p>
                        </div>
                        <p className="drop-zone-hint">.xlsx / .xls</p>
                      </div>
                    )}
                  </div>
                </div>

                {/* Drop zone: Reporte Novasoft */}
                <div className="form-group">
                  <label className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <rect x="3" y="3" width="18" height="18" rx="2"/>
                      <path d="M9 3v18M3 9h18M3 15h18"/>
                    </svg>
                    Reporte Novasoft
                  </label>
                  <input
                    ref={inputNova}
                    type="file"
                    accept=".xlsx,.xls"
                    className="file-input"
                    onChange={(e) => handleFileChange(e, setReporteNovasoft)}
                  />
                  <div
                    className={`drop-zone ${dragNova ? 'drag-active' : ''} ${reporteNovasoft ? 'has-file' : ''}`}
                    onClick={() => inputNova.current.click()}
                    onDragOver={(e) => { e.preventDefault(); setDragNova(true) }}
                    onDragLeave={() => setDragNova(false)}
                    onDrop={(e) => handleFileDrop(e, setReporteNovasoft, setDragNova)}
                  >
                    {reporteNovasoft ? (
                      <div className="file-preview">
                        <div className="file-icon">
                          <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                            <polyline points="14 2 14 8 20 8"/>
                            <line x1="8" y1="13" x2="16" y2="13"/>
                            <line x1="8" y1="17" x2="16" y2="17"/>
                          </svg>
                        </div>
                        <div className="file-details">
                          <p className="file-name">{reporteNovasoft.name}</p>
                          <p className="file-size">{formatSize(reporteNovasoft.size)}</p>
                        </div>
                        <button
                          className="btn-remove"
                          onClick={(e) => { e.stopPropagation(); setReporteNovasoft(null); inputNova.current.value = '' }}
                          title="Quitar archivo"
                        >
                          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <line x1="18" y1="6" x2="6" y2="18"/>
                            <line x1="6" y1="6" x2="18" y2="18"/>
                          </svg>
                        </button>
                      </div>
                    ) : (
                      <div className="drop-zone-content">
                        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                          <polyline points="17 8 12 3 7 8"/>
                          <line x1="12" y1="3" x2="12" y2="15"/>
                        </svg>
                        <div className="drop-zone-text">
                          <p className="drop-zone-title">Sube el reporte Novasoft</p>
                          <p className="drop-zone-subtitle">Arrastra aquí o haz clic para seleccionar</p>
                        </div>
                        <p className="drop-zone-hint">.xlsx / .xls</p>
                      </div>
                    )}
                  </div>
                </div>

                {/* Multiselect: Clientes a facturar */}
                <div className="form-group">
                  <label className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/>
                      <circle cx="9" cy="7" r="4"/>
                      <path d="M23 21v-2a4 4 0 0 0-3-3.87"/>
                      <path d="M16 3.13a4 4 0 0 1 0 7.75"/>
                    </svg>
                    Selecciona los clientes a facturar
                  </label>

                  <div className="multiselect-container" ref={dropdownRef}>
                    <button
                      type="button"
                      className={`multiselect-trigger ${dropdownOpen ? 'open' : ''}`}
                      onClick={() => setDropdownOpen(o => !o)}
                    >
                      <span className="multiselect-trigger-text">
                        {clientesSeleccionados.length === 0
                          ? 'Selecciona uno o más clientes...'
                          : `${clientesSeleccionados.length} cliente${clientesSeleccionados.length > 1 ? 's' : ''} seleccionado${clientesSeleccionados.length > 1 ? 's' : ''}`
                        }
                      </span>
                      <svg
                        className={`multiselect-chevron ${dropdownOpen ? 'open' : ''}`}
                        width="16" height="16" viewBox="0 0 24 24"
                        fill="none" stroke="currentColor" strokeWidth="2"
                      >
                        <polyline points="6 9 12 15 18 9"/>
                      </svg>
                    </button>

                    {dropdownOpen && (
                      <div className="multiselect-dropdown">
                        <div className="multiselect-search-wrap">
                          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <circle cx="11" cy="11" r="8"/>
                            <line x1="21" y1="21" x2="16.65" y2="16.65"/>
                          </svg>
                          <input
                            type="text"
                            className="multiselect-search"
                            placeholder="Buscar cliente..."
                            value={busqueda}
                            onChange={e => setBusqueda(e.target.value)}
                            autoFocus
                          />
                        </div>
                        <ul className="multiselect-list">
                          {clientesFiltrados.length > 0 ? clientesFiltrados.map(cliente => (
                            <li key={cliente}>
                              <label className={`multiselect-option ${clientesSeleccionados.includes(cliente) ? 'selected' : ''}`}>
                                <input
                                  type="checkbox"
                                  checked={clientesSeleccionados.includes(cliente)}
                                  onChange={() => toggleCliente(cliente)}
                                />
                                <span className="multiselect-check">
                                  {clientesSeleccionados.includes(cliente) && (
                                    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3">
                                      <polyline points="20 6 9 17 4 12"/>
                                    </svg>
                                  )}
                                </span>
                                {cliente}
                              </label>
                            </li>
                          )) : (
                            <li className="multiselect-empty">Sin resultados</li>
                          )}
                        </ul>
                        {clientesSeleccionados.length > 0 && (
                          <div className="multiselect-footer">
                            <button type="button" className="multiselect-clear" onClick={() => setClientesSeleccionados([])}>
                              Limpiar selección
                            </button>
                          </div>
                        )}
                      </div>
                    )}
                  </div>

                  {/* Tags de clientes seleccionados */}
                  {clientesSeleccionados.length > 0 && (
                    <div className="multiselect-tags">
                      {clientesSeleccionados.map(cliente => (
                        <span key={cliente} className="multiselect-tag">
                          {cliente}
                          <button type="button" onClick={() => toggleCliente(cliente)}>
                            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                              <line x1="18" y1="6" x2="6" y2="18"/>
                              <line x1="6" y1="6" x2="18" y2="18"/>
                            </svg>
                          </button>
                        </span>
                      ))}
                    </div>
                  )}
                </div>

                {/* Mensajes de error / éxito */}
                {error && (
                  <div className="alert alert-error">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <circle cx="12" cy="12" r="10"/>
                      <line x1="12" y1="8" x2="12" y2="12"/>
                      <line x1="12" y1="16" x2="12.01" y2="16"/>
                    </svg>
                    {error}
                  </div>
                )}
                {exitoCount > 0 && (
                  <div className="alert alert-success">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
                      <polyline points="22 4 12 14.01 9 11.01"/>
                    </svg>
                    {exitoCount === 1
                      ? '¡Archivo generado y descargado correctamente!'
                      : `¡${exitoCount} archivos generados y descargados correctamente!`
                    }
                  </div>
                )}

                {/* Botón Generar */}
                <button
                  className="btn-primary"
                  onClick={generar}
                  disabled={generando}
                >
                  {generando ? (
                    <>
                      <svg className="spinner" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
                      </svg>
                      Generando...
                    </>
                  ) : (
                    <>
                      <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                        <polyline points="7 10 12 15 17 10"/>
                        <line x1="12" y1="15" x2="12" y2="3"/>
                      </svg>
                      Generar archivo
                    </>
                  )}
                </button>

              </div>
            </div>
          </div>
        </div>
      </main>

      {/* Footer */}
      <footer className="footer">
        <div className="container">
          <p>&copy; {new Date().getFullYear()} Solutions & Payroll. Todos los derechos reservados.</p>
        </div>
      </footer>
    </div>
  )
}

export default App
