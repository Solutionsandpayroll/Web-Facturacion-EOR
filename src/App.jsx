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
  const [tasaCambio, setTasaCambio] = useState('')
  const [tasaCambioEur, setTasaCambioEur] = useState('')
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
    if (!tasaCambio.trim()) return setError('Escribe la tasa de cambio USD → COP.')
    if (clientesSeleccionados.includes('C1055 - RIVERMATE') && !tasaCambioEur.trim())
      return setError('Escribe la tasa de cambio EUR → COP para RIVERMATE.')

    setGenerando(true)
    try {
      // --- 1. Leer Base de empleados ---
      const baseBuffer = await baseEmpleados.arrayBuffer()
      const baseWb = XLSX.read(baseBuffer, { type: 'array', cellDates: true })
      const baseSheet = baseWb.Sheets[baseWb.SheetNames[0]]
      const baseData = XLSX.utils.sheet_to_json(baseSheet, { header: 1 })

      const headerBase = (baseData[0] || []).map(h => String(h ?? '').trim().toUpperCase())
      const colEmp    = headerBase.indexOf('CODIGO EMPLEADO')
      const colCC     = headerBase.indexOf('CENTRO COSTOS')
      const colAlt    = headerBase.indexOf('CODIGO ALTERNO')
      const colNombre = headerBase.indexOf('NOMBRE')
      const colFIng   = headerBase.indexOf('F_INGRESO')
      const colFRet   = headerBase.indexOf('F_RETIRO')
      const colSubcli = headerBase.indexOf('SUBCLIENTE')

      console.log('=== BASE DE EMPLEADOS ===')
      console.log('Encabezados encontrados:', headerBase)
      console.log('Columna CODIGO EMPLEADO (índice):', colEmp)
      console.log('Columna CENTRO COSTOS (índice):', colCC)
      console.log('Total filas (incluyendo encabezado):', baseData.length)

      if (colEmp === -1) throw new Error('No se encontró la columna "CODIGO EMPLEADO" en la base de empleados.')
      if (colCC  === -1) throw new Error('No se encontró la columna "CENTRO COSTOS" en la base de empleados.')

      // Mapa: códigoEmpleado → { cc, alt, nombre, fIngreso, fRetiro, subcli, clasif }
      const mapaCC = new Map()
      const mapaEmpleados = new Map()
      for (let i = 1; i < baseData.length; i++) {
        const fila   = baseData[i]
        const codigo = String(fila[colEmp] ?? '').trim()
        if (!codigo) continue
        const cc = String(fila[colCC] ?? '').trim()
        mapaCC.set(codigo, cc)

        // SUBCLIENTE: quitar prefijo "NN - " si no es "0 - NO APLICA"
        let subcliRaw = String(fila[colSubcli] ?? '').trim()
        let subcliVal = ''
        if (subcliRaw && !/^0\s*-\s*NO APLICA/i.test(subcliRaw)) {
          // Eliminar "XX - " del inicio (número + guión + espacio)
          subcliVal = subcliRaw.replace(/^\d+\s*-\s*/, '')
        }

        // Customer ID: tomar solo el prefijo numérico de SUBCLIENTE (ej. "06 - ..." => "06")
        // Si es "0 - NO APLICA", dejar vacío.
        let clasifVal = ''
        if (subcliRaw && !/^0\s*-\s*NO APLICA/i.test(subcliRaw)) {
          const m = subcliRaw.match(/^(\d+)\s*-/)
          clasifVal = m ? m[1] : ''
        }

        mapaEmpleados.set(codigo, {
          alt:      colAlt    !== -1 ? String(fila[colAlt]    ?? '').trim() : '',
          nombre:   colNombre !== -1 ? String(fila[colNombre] ?? '').trim() : '',
          fIngreso: colFIng   !== -1 ? fila[colFIng]  ?? null : null,
          fRetiro:  colFRet   !== -1 ? fila[colFRet]  ?? null : null,
          subcli:   subcliVal,
          clasif:   clasifVal,
        })
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
          columnaDestino: '13th Salary Alt',
          conceptos: [
            '001500 Prima de Servicios',
          ],
        },
        {
          columnaDestino: '14th Salary',
          conceptos: [
            '008412 Provision Cesantias',
          ],
        },
        {
          columnaDestino: '14th Salary Alt',
          conceptos: [
            '001560 Cesantias',
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
            '101310-Peoplepass Alimentacion',
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
          columnaDestino: 'Interest on 14th Salary Alt',
          conceptos: [
            '001565 Int Cesantias',
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
            '100210 Peoplepass Bienestar',
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

      // Extraer fórmulas de filas 4 y 5 con SheetJS ANTES de que JSZip elimine los clones
      // SheetJS resuelve shared formulas correctamente, así recuperamos todas las fórmulas
      const wbTpl = XLSX.read(tplBuffer, { cellFormula: true })
      const wsTpl = wbTpl.Sheets[wbTpl.SheetNames[0]]
      const formulasRow5 = {} // colNumber (1-based) → formula string
      const formulasRow4 = {} // colNumber (1-based) → formula string (fallback para colsTotal)
      Object.keys(wsTpl).forEach(cellAddr => {
        if (cellAddr.startsWith('!')) return
        const match = cellAddr.match(/^([A-Z]+)(\d+)$/)
        if (!match) return
        const cell = wsTpl[cellAddr]
        if (!cell || !cell.f) return
        const colNum = XLSX.utils.decode_col(match[1]) + 1
        const rowNum = parseInt(match[2])
        if (rowNum === 5) formulasRow5[colNum] = cell.f
        if (rowNum === 4) formulasRow4[colNum] = cell.f
      })

      console.log('=== formulasRow4 (SheetJS) ===', JSON.stringify(
        Object.fromEntries(Object.entries(formulasRow4).map(([k, v]) => [k + '(' + String.fromCharCode(64 + Number(k) % 26 || 26) + ')', v]))
      ))

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

      // Eliminar hojas extra (sheet2+) directamente del zip ANTES de que ExcelJS las cargue.
      // removeWorksheet() de ExcelJS no borra el archivo interno, por eso sale el error
      // "Registros quitados: Fórmula de /xl/worksheets/sheet2.xml".
      const hojasOrdenadas = [...sheetNames].sort()
      if (hojasOrdenadas.length > 1) {
        const extrasZip = hojasOrdenadas.slice(1)
        extrasZip.forEach(path => zip.remove(path))

        // xl/workbook.xml → quitar entradas <sheet> con sheetId > 1
        if (zip.files['xl/workbook.xml']) {
          let wbXml = await zip.files['xl/workbook.xml'].async('string')
          wbXml = wbXml.replace(/<sheet\b[^>]+\bsheetId="([^"]+)"[^>]*\/>/g,
            (m, id) => parseInt(id) > 1 ? '' : m)
          zip.file('xl/workbook.xml', wbXml)
        }

        // xl/_rels/workbook.xml.rels → quitar relaciones a las hojas eliminadas
        const relsPath = 'xl/_rels/workbook.xml.rels'
        if (zip.files[relsPath]) {
          let relsXml = await zip.files[relsPath].async('string')
          extrasZip.forEach(sheetPath => {
            const relTarget = sheetPath.replace(/^xl\//, '')
            const esc = relTarget.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
            relsXml = relsXml.replace(new RegExp(`<Relationship[^>]*Target="${esc}"[^>]*\\/>`, 'g'), '')
          })
          zip.file(relsPath, relsXml)
        }

        // [Content_Types].xml → quitar Override para las hojas eliminadas
        if (zip.files['[Content_Types].xml']) {
          let ctXml = await zip.files['[Content_Types].xml'].async('string')
          extrasZip.forEach(sheetPath => {
            const partName = ('/' + sheetPath).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
            ctXml = ctXml.replace(new RegExp(`<Override[^>]*PartName="${partName}"[^>]*\\/>`, 'g'), '')
          })
          zip.file('[Content_Types].xml', ctXml)
        }
      }

      const cleanTplBuffer = await zip.generateAsync({ type: 'arraybuffer' })

      // Verificar que el zip limpio no tiene sheet2
      JSZip.loadAsync(cleanTplBuffer).then(z => {
        console.log('=== Hojas en cleanTplBuffer (tras JSZip) ===',
          Object.keys(z.files).filter(f => f.startsWith('xl/worksheets/') && f.endsWith('.xml')))
      })

      const periodoLimpio = periodo.trim().replace(/[\/\\?%*:|"<>]/g, '-')
      let archivosGenerados = 0

      // --- 4. Generar un archivo por cada cliente seleccionado ---
      for (const cliente of clientesSeleccionados) {
        // Clientes que usan conceptos de liquidación (no provisiones) para 13th/14th/Interest
        const CLIENTES_LIQUIDACION = ['C1038 - ONCEHUB', 'C1029 - INSIDER']
        const esLiquidacion = CLIENTES_LIQUIDACION.includes(cliente)

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

        // Eliminar hojas extra del modelo ExcelJS (doble seguro junto al JSZip)
        console.log('=== Hojas en ExcelJS tras load ===', workbook.worksheets.map(ws => ws.name))
        workbook.worksheets.slice(1).forEach(ws => workbook.removeWorksheet(ws.id))
        console.log('=== Hojas en ExcelJS tras removeWorksheet ===', workbook.worksheets.map(ws => ws.name))

        const worksheet = workbook.worksheets[0]

        // Encontrar columna "EMPLOYEE CODE" y columnas destino de grupos en fila 3
        const headerRow = worksheet.getRow(3)
        let colEC = -1
        // mapa: nombreColumnaDestino → número de columna en plantilla
        const colsDestino = {}
        gruposConIndices.forEach(g => { colsDestino[g.columnaDestino] = -1 })

        // Columnas de fórmula: detectadas por nombre de encabezado, escritas post-splice
        let colPayments = -1
        const totalesHeaders = []  // cols con encabezado exacto "TOTAL", en orden de aparición
        let colTotalLegal = -1, colTotalCOP = -1, colTotalEmpCostUSD = -1
        let colFeeUSD = -1, colVAT = -1, colTotalUSD = -1
        // Columnas con fórmula variable por cliente
        let colFee = -1, colBanking = -1, colIva = -1
        // Columnas con datos de base de empleados o valores fijos
        let colRfWid = -1, colName = -1, colOnboard = -1, colOffboard = -1
        let colCustName = -1, colCustId = -1, colSvcType = -1, colPayMonth = -1, colCountry = -1
        let colErSs = -1, colEeStatus = -1, colExpenses = -1, colTotalEmpCost = -1, colExchangeRate = -1

        headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const val = String(cell.value ?? '').trim().toUpperCase()
          if (val === 'EMPLOYEE CODE')                  colEC        = colNumber
          if (val === 'EE RF WID')                      colRfWid     = colNumber
          if (val === 'NAME')                           colName      = colNumber
          if (val === 'ONBOARDING DATE')                colOnboard   = colNumber
          if (val === 'OFFBOARDING DATE')               colOffboard  = colNumber
          if (val === 'CUSTOMER NAME')                  colCustName  = colNumber
          if (val === 'CUSTOMER ID')                    colCustId    = colNumber
          if (val === 'SERVICE TYPE / INVOICE TYPE')    colSvcType   = colNumber
          if (val === 'PAYROLL MONTH')                  colPayMonth  = colNumber
          if (val === 'COUNTRY')                        colCountry   = colNumber
          if (val === 'ER SS RATE %')                   colErSs      = colNumber
          if (val === 'EE STATUS (ONBOARDING/ACTIVE/OFFBOARDING)') colEeStatus = colNumber
          if (val === 'EXPENSES') colExpenses = colNumber
          if (val === 'TOTAL EMPLOYEE COST') colTotalEmpCost = colNumber
          if (val === 'EXCHANGE RATE') colExchangeRate = colNumber
          gruposConIndices.forEach(g => {
            if (val === g.columnaDestino.toUpperCase()) colsDestino[g.columnaDestino] = colNumber
          })
          // Columnas de fórmula por cliente
          if (val === 'FEE')         { colFee     = colNumber; return }
          if (val === 'BANKING TAX') { colBanking  = colNumber; return }
          if (val === 'IVA')         { colIva      = colNumber; return }
          // Columnas de fórmula calculadas post-splice
          if (val === 'PAYMENTS')                      colPayments        = colNumber
          if (val === 'TOTAL')                         totalesHeaders.push(colNumber)
          if (val === 'TOTAL COST AND LEGAL BENEFITS') colTotalLegal      = colNumber
          if (val === 'TOTAL COP')                     colTotalCOP        = colNumber
          if (val === 'TOTAL EMPLOYEE COST USD')       colTotalEmpCostUSD = colNumber
          if (val === 'FEE USD')                       colFeeUSD          = colNumber
          if (val === 'VAT')                           colVAT             = colNumber
          if (val === 'TOTAL USD')                     colTotalUSD        = colNumber
        })

        const [colTotalSS, colTotalProv, colTotalOther] = [...totalesHeaders, -1, -1, -1]
        const cfg = CONFIG_CLIENTES[cliente] || {}

        console.log('colEC:', colEC, '| colsDestino:', colsDestino)
        if (colEC === -1) throw new Error('No se encontró la columna "EMPLOYEE CODE" en la fila 3 de la plantilla.')

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
        // Capturar estilos fila 4 para todas las columnas que se escribirán
        ;[colFee, colBanking, colIva,
          colRfWid, colName, colOnboard, colOffboard,
          colCustName, colCustId, colSvcType, colPayMonth, colCountry, colErSs, colEeStatus, colExchangeRate,
          colPayments, ...totalesHeaders, colTotalLegal, colTotalEmpCost, colTotalCOP, colTotalEmpCostUSD, colFeeUSD, colVAT, colTotalUSD,
        ].forEach(col => {
          if (col !== -1) {
            const c = worksheet.getRow(4).getCell(col)
            estilosRef[col] = c.style ? JSON.parse(JSON.stringify(c.style)) : {}
          }
        })
        // Capturar estilos de fila 5 (totales) antes de limpiarla → se restauran post-splice
        const estilosRow5 = {}
        worksheet.getRow(5).eachCell({ includeEmpty: true }, (cell, colNum) => {
          if (cell.style && Object.keys(cell.style).length > 0)
            estilosRow5[colNum] = JSON.parse(JSON.stringify(cell.style))
        })

        // Limpiar la fila 5 original (estilos ya capturados en estilosRow5)
        // Los formularios se reconstruirán dinámicamente post-splice
        const row5 = worksheet.getRow(5)
        const totalCols = worksheet.columnCount || 300
        for (let colNum = 1; colNum <= totalCols; colNum++) {
          const c = row5.getCell(colNum)
          c.value = null
          c.style = {}
        }
        row5.commit()

        // Limpiar datos previos desde fila 4 en todas las columnas destino
        // También limpiar las columnas de fórmula del template (TOTAL, PAYMENTS, etc.)
        // para que colsConDatos no las detecte como "con datos" debido a la fórmula original de fila 4.
        const todasLasCols = [
          colEC,
          ...Object.values(colsDestino).filter(c => c !== -1),
          colPayments, ...totalesHeaders, colTotalLegal, colTotalEmpCost,
          colFee, colBanking, colIva,
          colTotalCOP, colTotalEmpCostUSD, colFeeUSD, colVAT, colTotalUSD,
          colErSs, colEeStatus,
        ].filter((c, i, arr) => c !== -1 && arr.indexOf(c) === i)
        for (let r = 4; r <= worksheet.rowCount; r++) {
          todasLasCols.forEach(col => { worksheet.getRow(r).getCell(col).value = null })
        }
        // Limpieza defensiva: eliminar cualquier fórmula residual del template en filas de datos.
        // Evita referencias circulares en columnas que no reconstruimos explícitamente.
        for (let r = 4; r <= worksheet.rowCount; r++) {
          const row = worksheet.getRow(r)
          row.eachCell({ includeEmpty: false }, (cell) => {
            const v = cell.value
            if (v && typeof v === 'object' && v.formula) cell.value = null
          })
        }

        const sinUSD = ['C1042 - EPDM', 'C1055 - RIVERMATE', 'C1058 - POC PHARMA'].includes(cliente)
        const esRemofirst = cliente === 'C1037 - REMOFIRST'

        // Helpers de conversión columna ↔ letra Excel
        const numToLetter = (n) => {
          let s = ''
          while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26) }
          return s
        }
        const colLetterToNum = (str) => {
          let n = 0
          for (let i = 0; i < str.length; i++) n = n * 26 + str.toUpperCase().charCodeAt(i) - 64
          return n
        }

        // Escribir datos desde fila 4
        codigosFiltrados.forEach((codigo, i) => {
          const row = worksheet.getRow(4 + i)
          const valores = mapaValores.get(codigo) || {}

          // Employee code
          const cellEC = row.getCell(colEC)
          cellEC.value = codigo
          if (estilosRef[colEC]) cellEC.style = JSON.parse(JSON.stringify(estilosRef[colEC]))

          // Datos adicionales de la base de empleados
          const empData = mapaEmpleados.get(codigo) || {}
          const escribirCelda = (col, valor) => {
            if (col === -1) return
            const c = row.getCell(col)
            c.value = valor || null
            if (estilosRef[col]) c.style = JSON.parse(JSON.stringify(estilosRef[col]))
          }
          escribirCelda(colRfWid,    empData.alt    || null)
          escribirCelda(colName,     empData.nombre || null)
          escribirCelda(colOnboard,  empData.fIngreso)
          escribirCelda(colOffboard, empData.fRetiro)
          if (esRemofirst) {
            escribirCelda(colCustName, empData.subcli || null)
            escribirCelda(colCustId,   empData.clasif || null)
          }
          // Valores fijos por fila
          const MESES_EN = ['January','February','March','April','May','June','July','August','September','October','November','December']
          const mesActual = MESES_EN[new Date().getMonth()]
          escribirCelda(colSvcType,  'Monthly Payroll')
          escribirCelda(colPayMonth, mesActual)
          escribirCelda(colCountry,  'Colombia')
          if (!sinUSD && colExchangeRate !== -1) {
            const esRivermate = cliente === 'C1055 - RIVERMATE'
            const tasaStr = esRivermate ? tasaCambioEur : tasaCambio
            const tasaNum = parseFloat(tasaStr.replace(/,/g, '.'))
            escribirCelda(colExchangeRate, isNaN(tasaNum) ? null : tasaNum)
          }

          // EE Status: onboarding / active / offboarding
          if (colEeStatus !== -1) {
            const ahora = new Date()
            const mesHoy  = ahora.getMonth()
            const anioHoy = ahora.getFullYear()
            const toDate = (v) => {
              if (!v) return null
              if (v instanceof Date) return v
              const d = new Date(v)
              return isNaN(d) ? null : d
            }
            const fIng = toDate(empData.fIngreso)
            const fRet = toDate(empData.fRetiro)
            let status = 'active'
            if (fRet && (fRet.getFullYear() < anioHoy || (fRet.getFullYear() === anioHoy && fRet.getMonth() <= mesHoy))) {
              status = 'offboarding'
            } else if (fIng && fIng.getMonth() === mesHoy && fIng.getFullYear() === anioHoy) {
              status = 'onboarding'
            }
            escribirCelda(colEeStatus, status)
          }

          // Valores de cada grupo
          gruposConIndices.forEach(g => {
            const col = colsDestino[g.columnaDestino]
            if (col === -1) return
            const cell = row.getCell(col)
            // Para ONCEHUB/INSIDER usar conceptos de liquidación; otros usan provisiones
            let sumaKey = g.columnaDestino
            if (esLiquidacion && g.columnaDestino === '13th Salary') sumaKey = '13th Salary Alt'
            if (esLiquidacion && g.columnaDestino === '14th Salary') sumaKey = '14th Salary Alt'
            if (esLiquidacion && g.columnaDestino === 'Interest on 14th Salary') sumaKey = 'Interest on 14th Salary Alt'
            const suma = valores[sumaKey] ?? 0
            cell.value = suma !== 0 ? parseFloat(suma.toFixed(2)) : null
            if (estilosRef[col]) cell.style = JSON.parse(JSON.stringify(estilosRef[col]))
          })

          row.commit()
        })

        const filaTotal = 4 + codigosFiltrados.length
        const ultimaFilaDatos = filaTotal - 1  // última fila con empleados

        // ── Eliminar columnas vacías ────────────────────────────────────────────
        // (Las hojas extra ya se eliminaron del zip en el pre-procesado de la plantilla)

        // Verificar si hay datos de provisiones (13th/14th/Interest).
        // Si no los hay, el 2° TOTAL y las columnas de provisiones se eliminarán automáticamente.
        const tieneProvisiones = codigosFiltrados.some(codigo => {
          const vals = mapaValores.get(codigo) || {}
          if (esLiquidacion) {
            return (vals['13th Salary Alt']||0) !== 0 || (vals['14th Salary Alt']||0) !== 0 || (vals['Interest on 14th Salary Alt']||0) !== 0
          }
          return (vals['13th Salary']||0) !== 0 || (vals['14th Salary']||0) !== 0 || (vals['Interest on 14th Salary']||0) !== 0
        })

        // 1. Columnas de fórmula que deben sobrevivir aunque estén vacías en datos.
        //    El 2° TOTAL (provisiones = totalesHeaders[1]) solo se protege si hay datos de provisiones.
        const protectedCols = new Set([
          colPayments,
          totalesHeaders[0],
          tieneProvisiones ? (totalesHeaders[1] ?? -1) : -1,
          totalesHeaders[2],
          tieneProvisiones ? colTotalLegal : -1,
          colTotalEmpCost,
          colFee,
          cfg.banking ? colBanking : -1,
          cfg.iva     ? colIva     : -1,
          colTotalCOP,
          sinUSD ? -1 : colTotalEmpCostUSD,
          sinUSD ? -1 : colFeeUSD,
          cfg.iva && !sinUSD ? colVAT : -1,
          sinUSD ? -1 : colTotalUSD,
          sinUSD ? -1 : colExchangeRate,
        ].filter(c => c !== -1))

        // 2. Detectar columnas con datos en filas 4…filaTotal-1
        const colsConDatos = new Set()
        const maxCols = worksheet.columnCount || 300
        worksheet.eachRow((row, rowNum) => {
          if (rowNum < 4 || rowNum >= filaTotal) return
          row.eachCell({ includeEmpty: false }, (cell, colNum) => {
            const v = cell.value
            if (v !== null && v !== undefined && v !== '') colsConDatos.add(colNum)
          })
        })

        // 2.1 Eliminaciones condicionales explícitas (sin depender de "columna vacía")
        const forcedRemoveCols = new Set([
          // Configuración por cliente
          cfg.banking ? -1 : colBanking,
          cfg.iva ? -1 : colIva,
          cfg.iva && !sinUSD ? -1 : colVAT,

          // Clientes que no usan bloque USD
          sinUSD ? colExchangeRate : -1,
          sinUSD ? colTotalEmpCostUSD : -1,
          sinUSD ? colFeeUSD : -1,
          sinUSD ? colTotalUSD : -1,

          // Customer fields: solo REMOFIRST
          esRemofirst ? -1 : colCustName,
          esRemofirst ? -1 : colCustId,

          // Sin provisiones: eliminar columnas base + 2° TOTAL + TOTAL LEGAL
          tieneProvisiones ? -1 : colsDestino['13th Salary'],
          tieneProvisiones ? -1 : colsDestino['14th Salary'],
          tieneProvisiones ? -1 : colsDestino['Interest on 14th Salary'],
          tieneProvisiones ? -1 : (totalesHeaders[1] ?? -1),
          tieneProvisiones ? -1 : colTotalLegal,
        ].filter(c => c !== -1))

        const emptyColSet = new Set()
        for (let col = 1; col <= maxCols; col++) {
          // Regla global desactivada por solicitud:
          // if (!colsConDatos.has(col) && !protectedCols.has(col)) emptyColSet.add(col)
          if (forcedRemoveCols.has(col) && !protectedCols.has(col)) emptyColSet.add(col)
        }

        // 3. Mapa de columna original → nueva posición tras las eliminaciones
        const colShiftMap = {}
        let nEliminadas = 0
        for (let col = 1; col <= maxCols; col++) {
          if (emptyColSet.has(col)) { nEliminadas++; continue }
          colShiftMap[col] = col - nEliminadas
        }
        // Mapa inverso: nueva posición → columna original (para recuperar estilos)
        const colOrigFromNew = {}
        Object.entries(colShiftMap).forEach(([orig, newC]) => { colOrigFromNew[newC] = parseInt(orig) })

        // 4. Guardar y desunir merges antes del splice
        const savedMerges = []
        const mergeModel = (worksheet.model && worksheet.model.merges) || []
        // Primero, recopilar info de todos los merges y desunirlos
        mergeModel.forEach(rangeStr => {
          const [tlStr, brStr] = String(rangeStr).split(':')
          const parseAddr = addr => {
            const m = String(addr || '').match(/^([A-Z]+)(\d+)$/)
            return m ? { col: colLetterToNum(m[1]), row: parseInt(m[2]) } : null
          }
          const tl = parseAddr(tlStr), br = parseAddr(brStr || tlStr)
          if (tl && br) {
            // Leer valor y estilo de la top-left ANTES de desunir (mientras aún es master)
            const tlCell = worksheet.getRow(tl.row).getCell(tl.col)
            const tlValue = tlCell.value ?? null
            const tlStyle = tlCell.style ? JSON.parse(JSON.stringify(tlCell.style)) : {}
            savedMerges.push({ tl, br, tlValue, tlStyle })
            try { worksheet.unMergeCells(rangeStr) } catch (e) {}
          }
        })
        // Después de desunir todo: si la top-left va a ser eliminada, propagar
        // valor y estilo a la primera celda superviviente del rango (ya no son esclavas)
        savedMerges.forEach(({ tl, br, tlValue, tlStyle }) => {
          // Solo necesitamos propagar contenido visual de encabezados.
          // En filas de datos podría copiar fórmulas residuales y causar referencias circulares.
          if (tl.row > 3) return
          if (!emptyColSet.has(tl.col)) return  // top-left sobrevive, no hace falta propagar
          if (tlValue === null && Object.keys(tlStyle).length === 0) return  // nada que propagar
          for (let c = tl.col + 1; c <= br.col; c++) {
            if (!emptyColSet.has(c)) {
              const cell = worksheet.getRow(tl.row).getCell(c)
              if (tlValue !== null) cell.value = tlValue
              cell.style = JSON.parse(JSON.stringify(tlStyle))
              break
            }
          }
        })

        // 5. Eliminar columnas de derecha a izquierda
        const colsAEliminar = [...emptyColSet].sort((a, b) => b - a)
        for (const col of colsAEliminar) {
          worksheet.spliceColumns(col, 1)
        }

        // 6. Insertar columnas separadoras en blanco (ANTES de restaurar merges)
        // Calcular posiciones post-eliminación de los encabezados separadores via colShiftMap
        const sepColsPostDelete = [colPayments, totalesHeaders[0], colTotalLegal, totalesHeaders[2]]
          .filter(c => c !== -1 && colShiftMap[c] !== undefined)
          .map(c => colShiftMap[c])
          .filter((c, i, arr) => arr.indexOf(c) === i)
          .sort((a, b) => a - b) // ascendente para calcular el shift correctamente

        // Cuántos separadores se insertan ANTES de una columna post-eliminación dada
        const sepShiftFor = (col) => sepColsPostDelete.filter(s => s < col).length

        // Insertar de derecha a izquierda para no alterar índices
        for (const s of [...sepColsPostDelete].sort((a, b) => b - a)) {
          worksheet.spliceColumns(s + 1, 0, [])
        }

        // 7. Restaurar merges aplicando colShiftMap + sepShiftFor en un solo paso
        savedMerges.forEach(({ tl, br }) => {
          let newStart = null, newEnd = null
          for (let col = tl.col; col <= br.col; col++) {
            if (!emptyColSet.has(col) && colShiftMap[col] !== undefined) {
              const finalCol = colShiftMap[col] + sepShiftFor(colShiftMap[col])
              if (newStart === null) newStart = finalCol
              newEnd = finalCol
            }
          }
          if (newStart === null || newEnd === null) return
          if (newStart === newEnd && tl.row === br.row) return
          try {
            worksheet.mergeCells(`${numToLetter(newStart)}${tl.row}:${numToLetter(newEnd)}${br.row}`)
          } catch (e) {}
        })

        // Reconstruir colOrigFromNew considerando también el shift de los separadores
        // colOrigFromNew[finalCol] = originalCol, donde finalCol = colShiftMap[orig] + sepShiftFor(colShiftMap[orig])
        Object.entries(colShiftMap).forEach(([orig, postDeleteCol]) => {
          const finalCol = postDeleteCol + sepShiftFor(postDeleteCol)
          colOrigFromNew[finalCol] = parseInt(orig)
        })

        // 8. El límite a columna CA se aplica después de escribir las fórmulas (ver más abajo)

        // ── Fórmulas post-splice ──────────────────────────────────────────────────────────────────────
        // Re-escanear fila 3 para obtener las posiciones REALES de columnas tras el splice
        const totalesPost = []
        let pcPayments = -1, pcTotalLegal = -1, pcTotalEmpCost = -1
        let pcFee = -1, pcBanking = -1, pcIva = -1
        let pcTotalCOP = -1, pcTotalEmpCostUSD = -1, pcFeeUSD = -1, pcVAT = -1, pcTotalUSD = -1
        let pcExchangeRate = -1, pcErSs = -1
        const pcDestino = {}  // grupo.columnaDestino → col post-splice
        const postCols  = {}  // header (uppercase) → col post-splice

        worksheet.getRow(3).eachCell({ includeEmpty: true }, (cell, n) => {
          const v = String(cell.value ?? '').trim().toUpperCase()
          postCols[v] = n
          if (v === 'PAYMENTS')                      pcPayments        = n
          if (v === 'TOTAL')                         totalesPost.push(n)
          if (v === 'TOTAL COST AND LEGAL BENEFITS') pcTotalLegal      = n
          if (v === 'TOTAL EMPLOYEE COST')           pcTotalEmpCost    = n
          if (v === 'FEE')                           pcFee             = n
          if (v === 'BANKING TAX')                   pcBanking         = n
          if (v === 'IVA')                           pcIva             = n
          if (v === 'TOTAL COP')                     pcTotalCOP        = n
          if (v === 'TOTAL EMPLOYEE COST USD')       pcTotalEmpCostUSD = n
          if (v === 'FEE USD')                       pcFeeUSD          = n
          if (v === 'VAT')                           pcVAT             = n
          if (v === 'TOTAL USD')                     pcTotalUSD        = n
          if (v === 'EXCHANGE RATE')                 pcExchangeRate    = n
          if (v === 'ER SS RATE %')                  pcErSs            = n
          gruposConIndices.forEach(g => {
            if (v === g.columnaDestino.toUpperCase()) pcDestino[g.columnaDestino] = n
          })
        })
        // Asignar los TOTAL post-splice según si existen provisiones:
        // Con provisiones:    totalesPost = [1°TOTAL, 2°TOTAL, 3°TOTAL]
        // Sin provisiones:    totalesPost = [1°TOTAL, 3°TOTAL]  (2°TOTAL eliminado)
        const pcTotalSS   = totalesPost[0] ?? -1
        const pcTotalProv = tieneProvisiones ? (totalesPost[1] ?? -1) : -1
        const pcTotalOther = tieneProvisiones ? (totalesPost[2] ?? -1) : (totalesPost[1] ?? -1)

        // Helper: estilo fila 4 (pre-splice) para columna post-splice
        const sty  = (pc) => { const s = estilosRef[colOrigFromNew[pc]];  return s ? JSON.parse(JSON.stringify(s)) : {} }
        const sty5 = (pc) => { const orig = colOrigFromNew[pc]; const s = (orig ? (estilosRow5[orig] || estilosRef[orig]) : null); return s ? JSON.parse(JSON.stringify(s)) : {} }
        // Helper: SUM de lista de columnas para una fila
        const L = (c) => numToLetter(c)
        const sumCols = (cols, rowNum, targetCol = null) => {
          const valid = [...new Set(cols)]
            .filter(c => c && c !== -1)
            .filter(c => targetCol === null ? true : c !== targetCol)
          if (valid.length === 0) return null
          return `SUM(${valid.map(c => `${L(c)}${rowNum}`).join(',')})`
        }

        // Grupos de columnas post-splice para cada fórmula
        const PAYMENT_GRUPOS = [
          'SALARY', '"Sick" Leave', 'Alloawance 1 (Car Allowance)', 'Unused Holidays',
          'Overtime', 'Bonus/Commission', 'Deduction or Gross Amount adjustments prevous month',
          'Alloawance 2 (Mobile & Internet Allowance)', 'Food allowance',
          'Alloawance 3 (Home/Remote work allowance)', 'Home/Remote work allowance', 'Alloawance 4 (Other allowances)',
          'Sign-on Bonus', 'Transport allowance', 'Wellness Allowance',
          'On Call/ Plus Disponibilidad', 'Severance Pay (Taxable)',
          'Rectroactive payment/Plus Compensation', 'Paternity/ Maternity leave', 'Lieu of Notice',
        ]
        const TOTAL_SS_GRUPOS    = ['Family Fund Cost', 'Health Cost', 'ICBF cost', 'Labor Risk Cost', 'SENA Cost', 'Pension Cost']
        const TOTAL_PROV_GRUPOS  = ['13th Salary', '14th Salary', 'Interest on 14th Salary']
        const TOTAL_OTHER_GRUPOS = ['Expenses', 'Health Insurance', 'Medical Test']
        const TOTAL_OTHER_EXTRA  = ['LEAVE PAID REFOUND', 'PARKING COST', 'LIEU OF NOTICE']

        const PAYMENT_EXTRA_HEADERS = [
          'ALLOAWANCE 3 (HOME/REMOTE WORK ALLOWANCE)',
          'LIEU OF NOTICE',
        ]
        const paymentCols = [
          ...PAYMENT_GRUPOS.map(g => pcDestino[g]),
          ...PAYMENT_EXTRA_HEADERS.map(h => postCols[h]),
        ].filter(c => c && c !== -1)
        const ssCols      = TOTAL_SS_GRUPOS.map(g => pcDestino[g]).filter(c => c && c !== -1)
        const provCols    = TOTAL_PROV_GRUPOS.map(g => pcDestino[g]).filter(c => c && c !== -1)
        const otherCols   = [
          ...TOTAL_OTHER_GRUPOS.map(g => pcDestino[g]),
          ...TOTAL_OTHER_EXTRA.map(h => postCols[h]),
        ].filter(c => c && c !== -1)

        // ER SS Rate %: en el template original era la col AL (=38 en base 1)
        const pcErSsRef = colShiftMap[38]

        // Escribe las fórmulas de una fila concreta (datos + cálculos)
        const writeRowFormulas = (rowNum) => {
          const row = worksheet.getRow(rowNum)

          if (pcPayments !== -1) {
            const f = sumCols(paymentCols, rowNum, pcPayments)
            if (f) { row.getCell(pcPayments).value = { formula: f }; row.getCell(pcPayments).style = sty(pcPayments) }
          }
          if (pcTotalSS !== -1) {
            const f = sumCols(ssCols, rowNum, pcTotalSS)
            if (f) { row.getCell(pcTotalSS).value = { formula: f }; row.getCell(pcTotalSS).style = sty(pcTotalSS) }
          }
          if (pcTotalProv !== -1) {
            const f = sumCols(provCols, rowNum, pcTotalProv)
            if (f) { row.getCell(pcTotalProv).value = { formula: f }; row.getCell(pcTotalProv).style = sty(pcTotalProv) }
          }
          if (pcTotalLegal !== -1) {
            const f = sumCols([pcTotalSS, pcTotalProv].filter(c => c !== -1), rowNum, pcTotalLegal)
            if (f) { row.getCell(pcTotalLegal).value = { formula: f }; row.getCell(pcTotalLegal).style = sty(pcTotalLegal) }
          }
          if (pcTotalOther !== -1) {
            const f = sumCols(otherCols, rowNum, pcTotalOther)
            if (f) { row.getCell(pcTotalOther).value = { formula: f }; row.getCell(pcTotalOther).style = sty(pcTotalOther) }
          }
          // BANKING TAX
          if (pcBanking !== -1 && cfg.banking) {
            const base = [pcPayments, pcTotalSS, pcTotalProv, pcTotalOther].filter(c => c !== -1 && c !== pcBanking)
            if (base.length > 0) {
              row.getCell(pcBanking).value = { formula: `(${base.map(c => `${L(c)}${rowNum}`).join('+')})*0.004` }
              row.getCell(pcBanking).style = sty(pcBanking)
            }
          }
          // TOTAL EMPLOYEE COST
          if (pcTotalEmpCost !== -1) {
            const parts = [pcPayments, pcTotalSS, pcTotalProv, pcTotalOther,
              cfg.banking && pcBanking !== -1 ? pcBanking : -1].filter(c => c !== -1)
            const f = sumCols(parts, rowNum, pcTotalEmpCost)
            if (f) { row.getCell(pcTotalEmpCost).value = { formula: f }; row.getCell(pcTotalEmpCost).style = sty(pcTotalEmpCost) }
          }
          // FEE
          if (pcFee !== -1 && cfg.fee) {
            let feeFormula
            const esPorcentaje = cfg.fee.endsWith('%')
            if (esPorcentaje && pcTotalEmpCost !== -1 && pcFee !== pcTotalEmpCost) {
              feeFormula = cfg.banking && pcBanking !== -1 && pcBanking !== pcFee
                ? `${cfg.fee}*(${L(pcTotalEmpCost)}${rowNum}+${L(pcBanking)}${rowNum})`
                : `${cfg.fee}*${L(pcTotalEmpCost)}${rowNum}`
            } else if (cliente === 'C1055 - RIVERMATE') {
              const tasaEurNum = parseFloat(tasaCambioEur.replace(/,/g, '.'))
              feeFormula = `${cfg.fee}*${isNaN(tasaEurNum) ? 1 : tasaEurNum}`
            } else if (cliente === 'C1058 - POC PHARMA') {
              const tasaUsdNum = parseFloat(tasaCambio.replace(/,/g, '.'))
              feeFormula = `${cfg.fee}*${isNaN(tasaUsdNum) ? 1 : tasaUsdNum}`
            } else if (pcExchangeRate !== -1) {
              feeFormula = `${cfg.fee}*${L(pcExchangeRate)}${rowNum}`
            }
            if (feeFormula) { row.getCell(pcFee).value = { formula: feeFormula }; row.getCell(pcFee).style = sty(pcFee) }
          }
          // IVA
          if (pcIva !== -1 && cfg.iva && pcFee !== -1 && pcIva !== pcFee) {
            row.getCell(pcIva).value = { formula: `${L(pcFee)}${rowNum}*0.19` }
            row.getCell(pcIva).style = sty(pcIva)
          }
          // TOTAL COP
          if (pcTotalCOP !== -1) {
            const parts = [pcTotalEmpCost, pcFee, cfg.iva && pcIva !== -1 ? pcIva : -1].filter(c => c !== -1)
            const f = sumCols(parts, rowNum, pcTotalCOP)
            if (f) { row.getCell(pcTotalCOP).value = { formula: f }; row.getCell(pcTotalCOP).style = sty(pcTotalCOP) }
          }
          // TOTAL EMPLOYEE COST USD
          if (!sinUSD && pcTotalEmpCostUSD !== -1 && pcTotalEmpCost !== -1 && pcExchangeRate !== -1 && pcTotalEmpCostUSD !== pcTotalEmpCost && pcTotalEmpCostUSD !== pcExchangeRate) {
            row.getCell(pcTotalEmpCostUSD).value = { formula: `ROUND(${L(pcTotalEmpCost)}${rowNum}/${L(pcExchangeRate)}${rowNum},2)` }
            row.getCell(pcTotalEmpCostUSD).style = sty(pcTotalEmpCostUSD)
          }
          // FEE USD
          if (!sinUSD && pcFeeUSD !== -1 && pcFee !== -1 && pcExchangeRate !== -1 && pcFeeUSD !== pcFee && pcFeeUSD !== pcExchangeRate) {
            row.getCell(pcFeeUSD).value = { formula: `ROUND(${L(pcFee)}${rowNum}/${L(pcExchangeRate)}${rowNum},2)` }
            row.getCell(pcFeeUSD).style = sty(pcFeeUSD)
          }
          // VAT
          if (!sinUSD && pcVAT !== -1 && cfg.iva && pcIva !== -1 && pcExchangeRate !== -1 && pcVAT !== pcIva && pcVAT !== pcExchangeRate) {
            row.getCell(pcVAT).value = { formula: `ROUND(${L(pcIva)}${rowNum}/${L(pcExchangeRate)}${rowNum},2)` }
            row.getCell(pcVAT).style = sty(pcVAT)
          }
          // TOTAL USD
          if (!sinUSD && pcTotalUSD !== -1) {
            const parts = [pcTotalEmpCostUSD, pcFeeUSD, cfg.iva && pcVAT !== -1 ? pcVAT : -1].filter(c => c !== -1)
            const f = sumCols(parts, rowNum, pcTotalUSD)
            if (f) { row.getCell(pcTotalUSD).value = { formula: f }; row.getCell(pcTotalUSD).style = sty(pcTotalUSD) }
          }
          // ER SS Rate % — solo REMOFIRST
          if (pcErSs !== -1 && cliente === 'C1037 - REMOFIRST' && pcErSsRef && pcErSsRef !== pcErSs) {
            row.getCell(pcErSs).value = { formula: `IF(${L(pcErSsRef)}${rowNum}>0,31.94,14.44)` }
            row.getCell(pcErSs).style = sty(pcErSs)
          }

          row.commit()
        }

        // Escribir fórmulas en cada fila de empleado
        for (let i = 0; i < codigosFiltrados.length; i++) writeRowFormulas(4 + i)

        // Fila de totales: SUM de cada columna numérica
        const rowTotals = worksheet.getRow(filaTotal)
        const todasNumCols = new Set([
          ...Object.values(pcDestino).filter(c => c && c !== -1),
          ...PAYMENT_EXTRA_HEADERS.map(h => postCols[h]).filter(c => c && c !== -1),
          pcPayments, pcTotalSS, pcTotalProv, pcTotalLegal, pcTotalOther, pcTotalEmpCost,
          cfg.fee     && pcFee     !== -1 ? pcFee     : -1,
          cfg.banking && pcBanking !== -1 ? pcBanking : -1,
          cfg.iva     && pcIva     !== -1 ? pcIva     : -1,
          pcTotalCOP, pcTotalEmpCostUSD, pcFeeUSD,
          cfg.iva     && pcVAT     !== -1 ? pcVAT     : -1,
          pcTotalUSD,
        ].filter(c => c && c !== -1))
        todasNumCols.forEach(col => {
          const cell = rowTotals.getCell(col)
          cell.value = { formula: `SUM(${L(col)}4:${L(col)}${ultimaFilaDatos})` }
          cell.style = sty5(col)
        })
        rowTotals.commit()

        // Limitar a columna CA (79): eliminar todo lo que quede más allá.
        // Se usa un número fijo grande porque worksheet.columnCount puede devolver
        // el extent original del template (XEA, XDX...) aunque esté vacío.
        worksheet.spliceColumns(80, 20000)

        // Descargar: nombre = "Facturación EOR - {Cliente} - {Periodo}.xlsx"
        const clienteLimpio = cliente.replace(/[\/\\?%*:|"<>]/g, '-')
        let outBuffer = await workbook.xlsx.writeBuffer()

        // Post-proceso: ExcelJS regenera sheet2 al hacer spliceColumns incluso si
        // la plantilla ya no lo tenía. Eliminarlo del buffer de salida.
        {
          const outZip = await JSZip.loadAsync(outBuffer)
          const outSheets = Object.keys(outZip.files)
            .filter(f => f.startsWith('xl/worksheets/') && f.endsWith('.xml'))
            .sort()
          if (outSheets.length > 1) {
            const extras = outSheets.slice(1)
            extras.forEach(p => outZip.remove(p))
            if (outZip.files['xl/workbook.xml']) {
              let wx = await outZip.files['xl/workbook.xml'].async('string')
              wx = wx.replace(/<sheet\b[^>]+\bsheetId="([^"]+)"[^>]*\/>/g,
                (m, id) => parseInt(id) > 1 ? '' : m)
              outZip.file('xl/workbook.xml', wx)
            }
            const rp = 'xl/_rels/workbook.xml.rels'
            if (outZip.files[rp]) {
              let rx = await outZip.files[rp].async('string')
              extras.forEach(sp => {
                const rt = sp.replace(/^xl\//, '')
                const esc = rt.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
                rx = rx.replace(new RegExp(`<Relationship[^>]*Target="${esc}"[^>]*\\/>`, 'g'), '')
              })
              outZip.file(rp, rx)
            }
            if (outZip.files['[Content_Types].xml']) {
              let cx = await outZip.files['[Content_Types].xml'].async('string')
              extras.forEach(sp => {
                const pn = ('/' + sp).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
                cx = cx.replace(new RegExp(`<Override[^>]*PartName="${pn}"[^>]*\\/>`, 'g'), '')
              })
              outZip.file('[Content_Types].xml', cx)
            }
            outBuffer = await outZip.generateAsync({ type: 'arraybuffer' })
          }
        }

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
                    <strong>Selecciona el periodo</strong>
                    <p>Escribe el periodo de facturación (ej: <em>Enero 2026</em>). Este texto se usará en el nombre del archivo generado.</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">2</span>
                  <div>
                    <strong>Carga la base de empleados</strong>
                    <p>Arrastra o selecciona el archivo Excel de base de empleados (contiene nombres, fechas de ingreso/retiro, subclienta, etc.).</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">3</span>
                  <div>
                    <strong>Carga el reporte de Novasoft</strong>
                    <p>Arrastra o selecciona el archivo Excel exportado desde Novasoft con los conceptos de nómina del periodo.</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">4</span>
                  <div>
                    <strong>Selecciona los clientes a facturar</strong>
                    <p>Marca uno o varios clientes de la lista. Se generará un archivo Excel independiente por cada cliente seleccionado.</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">5</span>
                  <div>
                    <strong>Ingresa la tasa de cambio USD → COP</strong>
                    <p>Escribe el valor de la tasa de cambio del dólar a peso colombiano vigente para el periodo.</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">6</span>
                  <div>
                    <strong>Genera los archivos</strong>
                    <p>Haz clic en <em>Generar Excel</em>. El sistema procesará cada cliente y descargará automáticamente un archivo Excel con la facturación lista, incluyendo fórmulas, formatos y columnas ajustadas según el cliente.</p>
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

                {/* Tasa de cambio USD→COP — visible sólo cuando hay clientes seleccionados */}
                {clientesSeleccionados.length > 0 && (
                  <div className="form-group">
                    <label className="label">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <line x1="12" y1="1" x2="12" y2="23"/>
                        <path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>
                      </svg>
                      Tasa de cambio (USD → COP)
                    </label>
                    <input
                      type="number"
                      placeholder="Ej: 4200"
                      className="select-input"
                      value={tasaCambio}
                      onChange={e => { setTasaCambio(e.target.value); setError(null) }}
                    />
                  </div>
                )}

                {/* Tasa de cambio EUR→COP — solo si RIVERMATE está seleccionado */}
                {clientesSeleccionados.includes('C1055 - RIVERMATE') && (
                  <div className="form-group">
                    <label className="label">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <line x1="12" y1="1" x2="12" y2="23"/>
                        <path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>
                      </svg>
                      Tasa de cambio RIVERMATE (EUR → COP)
                    </label>
                    <input
                      type="number"
                      placeholder="Ej: 4850"
                      className="select-input"
                      value={tasaCambioEur}
                      onChange={e => { setTasaCambioEur(e.target.value); setError(null) }}
                    />
                  </div>
                )}

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
