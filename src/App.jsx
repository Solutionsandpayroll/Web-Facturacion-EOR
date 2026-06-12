import { useState, useRef, useEffect, useCallback } from 'react'
import * as XLSX from 'xlsx'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { neon } from '@neondatabase/serverless'
import './App.css'

const sql = neon(import.meta.env.VITE_DATABASE_URL)

// Configuración por defecto (hardcoded como fallback y para el botón "Restablecer")
// fee: '6%' = porcentaje, '120' = valor fijo; sin_usd = oculta bloque USD; es_liquidacion = usa conceptos de liquidación
const DEFAULT_CONFIG_CLIENTES_GENERAL = {
  'C1007 - NZD':               { fee: '6%',     iva: false, banking: false, sin_usd: false, es_liquidacion: false },
  'C1024 - MKD':               { fee: '11%',    iva: false, banking: false, sin_usd: false, es_liquidacion: false },
  'C1032 - FLEXCO':            { fee: '9%',     iva: false, banking: false, sin_usd: false, es_liquidacion: false },
  'C1038 - ONCEHUB':           { fee: '100.84', iva: true,  banking: true,  sin_usd: false, es_liquidacion: true  },
  'C1041 - EDRINGTON':         { fee: '5.5%',   iva: false, banking: false, sin_usd: false, es_liquidacion: false },
  'C1042 - EPDM':              { fee: '10%',    iva: false, banking: false, sin_usd: true,  es_liquidacion: false },
  'C1043 - NEO':               { fee: '8%',     iva: false, banking: false, sin_usd: false, es_liquidacion: false },
  'C1050 - HEMMERSBACH':       { fee: '10%',    iva: false, banking: false, sin_usd: false, es_liquidacion: false },
  'C1051 - YONYOU':            { fee: '210',    iva: true,  banking: true,  sin_usd: false, es_liquidacion: false },
  'C1052 - BUBBLE BPM INC':    { fee: '11%',    iva: false, banking: false, sin_usd: false, es_liquidacion: false },
  'C1053 - GLOBAL EXPANSION':  { fee: '150',    iva: true,  banking: true,  sin_usd: false, es_liquidacion: false },
  'C1055 - RIVERMATE':         { fee: '150',    iva: false, banking: true,  sin_usd: true,  es_liquidacion: false },
  'C1037 - REMOFIRST':         { fee: '120',    iva: true,  banking: true,  sin_usd: false, es_liquidacion: false },
  'C1029 - INSIDER':           { fee: '190',    iva: true,  banking: true,  sin_usd: false, es_liquidacion: true  },
  'C1036 - ACTION AD':         { fee: '11%',    iva: false, banking: false, sin_usd: false, es_liquidacion: false },
  'C1056 - EUROPORTAGE':       { fee: '200',    iva: true,  banking: true,  sin_usd: false, es_liquidacion: false },
  'C1022 - Root Capital':      { fee: '9%',     iva: false, banking: false, sin_usd: false, es_liquidacion: false },
  'C1058 - POC PHARMA':        { fee: '160',    iva: false, banking: true,  sin_usd: true,  es_liquidacion: false },
  'C1059 - SIFFI':             { fee: '10%',    iva: false, banking: true,  sin_usd: false, es_liquidacion: false },
  'C1060 - BETTEBUNA':         { fee: '10%',    iva: false, banking: true,  sin_usd: false, es_liquidacion: false },
}

const DEFAULT_CONFIG_CLIENTES_COSTA_RICA = {
  'CR010 - BIPO SANYl':        { fee: '135',    iva: false, banking: true,  sin_usd: true,  es_liquidacion: false },
  'CR010 - BIPO CATL':         { fee: '135',    iva: false, banking: true,  sin_usd: false, es_liquidacion: false },
  'CR010 - BIPO DONCHENG':     { fee: '135',    iva: false, banking: true,  sin_usd: true,  es_liquidacion: false },
  'CR018 - BUBBLE BPM INC':    { fee: '11%',    iva: false, banking: false, sin_usd: true,  es_liquidacion: false },
  'CR009 - REMOFIRST INC':     { fee: '135',    iva: false, banking: true,  sin_usd: false, es_liquidacion: false },
  'CR017 - EUROPORTAGE':       { fee: '200',    iva: false, banking: true,  sin_usd: false, es_liquidacion: false },
}

const DEFAULT_CONFIG_CLIENTES = {
  ...DEFAULT_CONFIG_CLIENTES_GENERAL,
  ...DEFAULT_CONFIG_CLIENTES_COSTA_RICA,
}

const PROCESS_CONFIG = {
  general: {
    key: 'general',
    label: 'General',
    countryLabel: 'Colombia',
    clientPrefix: 'C',
    templatePath: '/Formato facturacion - final.xlsx',
    reportSheetName: 'nomplcol',
    useBaseEmployees: true,
    baseRequired: true,
    requiresEurRate: true,
    currencyLabel: 'USD → COP',
    exchangeRateHint: 'Escribe la tasa de cambio USD → COP.',
  },
  'costa-rica': {
    key: 'costa-rica',
    label: 'Costa Rica',
    countryLabel: 'Costa Rica',
    clientPrefix: 'CR',
    templatePath: '/Formato facturacion final - costa rica.xlsx',
    reportSheetName: 'rpt_InformeNominaMes',
    useBaseEmployees: false,
    baseRequired: false,
    requiresEurRate: false,
    currencyLabel: 'CRC → USD',
    exchangeRateHint: 'Escribe la tasa de cambio CRC → USD.',
  },
}

const DEFAULT_REMOFIRST_SUBCLIENTES = [
  { ccosto: 'CR004', descCcosto: 'ANGL RF', nombreCompleto: '' },
  { ccosto: 'CR006', descCcosto: 'CZARNIKOW', nombreCompleto: '' },
  { ccosto: 'CR007', descCcosto: 'FLORES & PELAEZ-PRADA PLLC', nombreCompleto: '' },
  { ccosto: 'CR020', descCcosto: 'UPTOWN MOOSE LLC (RF)', nombreCompleto: '' },
  { ccosto: 'CR003', descCcosto: 'SSS - RF', nombreCompleto: '' },
  { ccosto: 'CR016', descCcosto: 'SUOL INNOVATIONS LTD (RF)', nombreCompleto: '' },
  { ccosto: 'CR012', descCcosto: 'REMOFIRST - SKELLIG AUTOMATION US LLC', nombreCompleto: '' },
  { ccosto: 'CR011', descCcosto: 'REMOFIRST-IPPC TECHNOLOGIES', nombreCompleto: '' },
  { ccosto: 'CR019', descCcosto: 'COMRISE INC (RF)', nombreCompleto: '' },
]

const DEFAULT_EMPLEADOS_CONFIG_COSTA_RICA = [
  { documento: '112280205', nombre: 'TATIANA ARAYA POCHET' },
  { documento: '112850849', nombre: 'BERNAL FALLAS BARBOZA' },
  { documento: '503990169', nombre: 'JONATHAN ALONSO SALAZAR GARCIA' },
  { documento: '603250181', nombre: 'ANGEL ALBERTO TORRES GONZALEZ' },
  { documento: '110930288', nombre: 'JORGE ESTEBAN VALVERDE ESPINOZA' },
  { documento: '701910926', nombre: 'OSCAR JOSUE PEREZ CASCANTE' },
  { documento: '205940804', nombre: 'LUIS ANGEL VEGA DELGADO' },
  { documento: '155825766303', nombre: 'ORIA BRITO VANEGA' },
]

const MONTH_OPTIONS = [
  'ENERO',
  'FEBRERO',
  'MARZO',
  'ABRIL',
  'MAYO',
  'JUNIO',
  'JULIO',
  'AGOSTO',
  'SEPTIEMBRE',
  'OCTUBRE',
  'NOVIEMBRE',
  'DICIEMBRE',
]

function getDefaultConfigByProcess(processKey) {
  return processKey === 'costa-rica'
    ? DEFAULT_CONFIG_CLIENTES_COSTA_RICA
    : DEFAULT_CONFIG_CLIENTES_GENERAL
}

function getClientesTableByProcess(processKey) {
  return processKey === 'costa-rica' ? 'clientes_config_costa_rica' : 'clientes_config'
}

const REMO_MATCH_AUTOMATCH_THRESHOLD = 0.8
const REMO_MATCH_REVIEW_THRESHOLD = 0.55

function normalizeMatchText(value) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Z0-9\s]/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase()
}

function tokenizeMatchText(value) {
  const normalized = normalizeMatchText(value)
  return normalized ? normalized.split(' ').filter(Boolean) : []
}

function isTokenSubsequence(shortTokens, longTokens) {
  if (shortTokens.length === 0 || shortTokens.length > longTokens.length) return false
  let longIndex = 0
  for (const token of shortTokens) {
    while (longIndex < longTokens.length && longTokens[longIndex] !== token) longIndex += 1
    if (longIndex >= longTokens.length) return false
    longIndex += 1
  }
  return true
}

function scoreMatchText(leftValue, rightValue) {
  const leftTokens = tokenizeMatchText(leftValue)
  const rightTokens = tokenizeMatchText(rightValue)
  if (leftTokens.length === 0 || rightTokens.length === 0) return 0

  const leftNormalized = leftTokens.join(' ')
  const rightNormalized = rightTokens.join(' ')
  if (leftNormalized === rightNormalized) return 1

  const leftCompact = leftTokens.join('')
  const rightCompact = rightTokens.join('')
  const compactIncludes = leftCompact.includes(rightCompact) || rightCompact.includes(leftCompact)

  const rightTokenSet = new Set(rightTokens)
  const commonTokens = leftTokens.filter(token => rightTokenSet.has(token))
  if (commonTokens.length === 0) {
    return compactIncludes ? 0.65 : 0
  }

  const tokenBalance = (2 * commonTokens.length) / (leftTokens.length + rightTokens.length)
  const coverage = Math.max(commonTokens.length / leftTokens.length, commonTokens.length / rightTokens.length)
  const subsequenceBonus = isTokenSubsequence(
    leftTokens.length <= rightTokens.length ? leftTokens : rightTokens,
    leftTokens.length <= rightTokens.length ? rightTokens : leftTokens,
  ) ? 0.1 : 0
  const compactBonus = compactIncludes ? 0.08 : 0

  return Math.min(1, (tokenBalance * 0.6) + (coverage * 0.3) + subsequenceBonus + compactBonus)
}

function getRemoMatchCandidates(employeeName, codesEntries, limit = 5) {
  return codesEntries
    .map(entry => ({
      ...entry,
      score: scoreMatchText(employeeName, entry.fullName),
    }))
    .filter(entry => entry.score > 0)
    .sort((a, b) => b.score - a.score || a.fullName.localeCompare(b.fullName))
    .slice(0, limit)
}

function App() {
  const [activeProcess, setActiveProcess] = useState('general')
  const [isHelpExpanded, setIsHelpExpanded] = useState(false)
  const [baseEmpleados, setBaseEmpleados] = useState(null)
  const [reporteNovasoft, setReporteNovasoft] = useState(null)
  const [codigosRemo, setCodigosRemo] = useState(null)
  const [dragBase, setDragBase] = useState(false)
  const [dragNova, setDragNova] = useState(false)
  const [reporteValoresHealth, setReporteValoresHealth] = useState(null)
  const [dragHealth, setDragHealth] = useState(false)
  const [dragCodes, setDragCodes] = useState(false)
  const inputBase = useRef(null)
  const inputNova = useRef(null)
  const inputHealth = useRef(null)
  const inputCodes = useRef(null)
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

  // ── Panel de configuración de clientes (Neon DB) ──
  const [clientesConfig, setClientesConfig] = useState(DEFAULT_CONFIG_CLIENTES)
  const [adminOpen, setAdminOpen] = useState(false)
  const [configLoading, setConfigLoading] = useState(false)
  const [configError, setConfigError] = useState(null)
  const [editando, setEditando] = useState(null)   // code del cliente en edición
  const [editForm, setEditForm] = useState({})     // valores del formulario activo
  const [savingConfig, setSavingConfig] = useState(false)
  const [resetting, setResetting] = useState(false)
  const [activeTab, setActiveTab] = useState('facturacion')
  const [showAddForm, setShowAddForm] = useState(false)
  const [addForm, setAddForm] = useState({ code: '', fee: '10%', iva: false, banking: false, sin_usd: false, es_liquidacion: false })
  const [addingClient, setAddingClient] = useState(false)
  const [remoSubclientes, setRemoSubclientes] = useState([])
  const [subclientesLoading, setSubclientesLoading] = useState(false)
  const [subclientesError, setSubclientesError] = useState(null)
  const [showRemoSubForm, setShowRemoSubForm] = useState(false)
  const [newRemoSubcliente, setNewRemoSubcliente] = useState({ ccosto: '', descCcosto: '', nombreCompleto: '' })
  const [addingRemoSubcliente, setAddingRemoSubcliente] = useState(false)
  const [deletingRemoSubcliente, setDeletingRemoSubcliente] = useState('')
  const [editingRemoSubcliente, setEditingRemoSubcliente] = useState('')
  const [editRemoSubclienteForm, setEditRemoSubclienteForm] = useState({ ccosto: '', descCcosto: '', nombreCompleto: '' })
  const [savingRemoSubcliente, setSavingRemoSubcliente] = useState(false)
  const [empleadosConfig, setEmpleadosConfig] = useState([])
  const [empleadosConfigLoading, setEmpleadosConfigLoading] = useState(false)
  const [empleadosConfigError, setEmpleadosConfigError] = useState(null)
  const [showEmpleadoForm, setShowEmpleadoForm] = useState(false)
  const [newEmpleado, setNewEmpleado] = useState({ documento: '', nombre: '' })
  const [addingEmpleado, setAddingEmpleado] = useState(false)
  const [deletingEmpleado, setDeletingEmpleado] = useState('')
  const [editingEmpleado, setEditingEmpleado] = useState('')
  const [editEmpleadoForm, setEditEmpleadoForm] = useState({ documento: '', nombre: '' })
  const [savingEmpleado, setSavingEmpleado] = useState(false)
  const [costaRicaExchangeRate, setCostaRicaExchangeRate] = useState('')
  const [healthInsuranceMonth, setHealthInsuranceMonth] = useState('')
  const [nameMatchReview, setNameMatchReview] = useState(null)
  const nameMatchResolveRef = useRef(null)
  const currentProcess = PROCESS_CONFIG[activeProcess] || PROCESS_CONFIG.general

  useEffect(() => {
    const handleClickOutside = (e) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target)) {
        setDropdownOpen(false)
      }
    }
    document.addEventListener('mousedown', handleClickOutside)
    return () => document.removeEventListener('mousedown', handleClickOutside)
  }, [])

  useEffect(() => {
    setClientesSeleccionados([])
    setBusqueda('')
    setDropdownOpen(false)
    setBaseEmpleados(null)
    setReporteNovasoft(null)
    setReporteValoresHealth(null)
    setCodigosRemo(null)
    setNameMatchReview(null)
    if (nameMatchResolveRef.current) {
      nameMatchResolveRef.current(null)
      nameMatchResolveRef.current = null
    }
    setError(null)
    setExitoCount(0)
    setTasaCambio('')
    setTasaCambioEur('')
    setCostaRicaExchangeRate('')
    setHealthInsuranceMonth('')
    setDragHealth(false)
    if (inputBase.current) inputBase.current.value = ''
    if (inputNova.current) inputNova.current.value = ''
    if (inputHealth.current) inputHealth.current.value = ''
    if (inputCodes.current) inputCodes.current.value = ''
  }, [activeProcess])

  const openNameMatchReview = useCallback((review) => new Promise(resolve => {
    nameMatchResolveRef.current = resolve
    setNameMatchReview(review)
  }), [])

  const closeNameMatchReview = useCallback((selection = null) => {
    const resolve = nameMatchResolveRef.current
    nameMatchResolveRef.current = null
    setNameMatchReview(null)
    if (resolve) resolve(selection)
  }, [])

  // ── Cargar configuración desde Neon DB ──────────────────────────────────────
  const loadConfig = useCallback(async () => {
    setConfigLoading(true)
    setConfigError(null)
    try {
      const targetTable = getClientesTableByProcess(activeProcess)
      const rows = targetTable === 'clientes_config_costa_rica'
        ? await sql`
            SELECT code, fee, iva, banking, sin_usd, es_liquidacion
            FROM clientes_config_costa_rica
            WHERE activo = true
            ORDER BY code
          `
        : await sql`
            SELECT code, fee, iva, banking, sin_usd, es_liquidacion
            FROM clientes_config
            WHERE activo = true
            ORDER BY code
          `
      const cfg = {}
      rows.forEach(row => {
        cfg[row.code] = {
          fee:           row.fee,
          iva:           row.iva,
          banking:       row.banking,
          sin_usd:       row.sin_usd,
          es_liquidacion: row.es_liquidacion,
        }
      })

      setClientesConfig(cfg)
    } catch (err) {
      console.error('Error al cargar configuración desde DB:', err)
      setConfigError('No se pudo conectar con la base de datos para cargar la configuración de clientes.')
      setClientesConfig({})
    } finally {
      setConfigLoading(false)
    }
  }, [activeProcess])

  useEffect(() => { loadConfig() }, [loadConfig])

  const loadRemoSubclientes = useCallback(async () => {
    if (activeProcess !== 'costa-rica') {
      setRemoSubclientes([])
      setSubclientesError(null)
      return
    }
    setSubclientesLoading(true)
    setSubclientesError(null)
    try {
      const rows = await sql`
        SELECT ccosto, desc_ccosto, nombre_completo
        FROM remofirst_subclientes_config
        WHERE activo = true
        ORDER BY ccosto, desc_ccosto
      `
      setRemoSubclientes(rows.map(row => ({
        ccosto: String(row.ccosto ?? '').trim(),
        descCcosto: String(row.desc_ccosto ?? '').trim(),
        nombreCompleto: String(row.nombre_completo ?? '').trim(),
      })))
    } catch (err) {
      console.error('Error al cargar subclientes Remofirst:', err)
      setSubclientesError('No se pudo cargar la configuración de subclientes Remofirst.')
      setRemoSubclientes([])
    } finally {
      setSubclientesLoading(false)
    }
  }, [activeProcess])

  useEffect(() => {
    if (activeProcess === 'costa-rica') {
      loadRemoSubclientes()
    }
  }, [activeProcess, loadRemoSubclientes])

  const normalizeDocumento = (value) => {
    const raw = String(value ?? '').trim()
    const digits = raw.replace(/\D/g, '')
    if (digits) return digits
    return raw.toUpperCase().replace(/\s+/g, '')
  }

  const loadEmpleadosConfig = useCallback(async () => {
    if (activeProcess !== 'costa-rica') {
      setEmpleadosConfig([])
      setEmpleadosConfigError(null)
      return
    }
    setEmpleadosConfigLoading(true)
    setEmpleadosConfigError(null)
    try {
      const rows = await sql`
        SELECT documento, nombre
        FROM empleados_config_costa_rica
        WHERE activo = true
        ORDER BY documento
      `
      setEmpleadosConfig(rows.map(row => ({
        documento: String(row.documento ?? '').trim(),
        nombre: String(row.nombre ?? '').trim(),
      })))
    } catch (err) {
      console.error('Error al cargar configuración de empleados CR:', err)
      setEmpleadosConfigError('No se pudo cargar la configuración de empleados de Costa Rica.')
      setEmpleadosConfig([])
    } finally {
      setEmpleadosConfigLoading(false)
    }
  }, [activeProcess])

  useEffect(() => {
    if (activeProcess === 'costa-rica') {
      loadEmpleadosConfig()
    }
  }, [activeProcess, loadEmpleadosConfig])

  // ── Guardar un cliente en DB ────────────────────────────────────────────────
  const saveClienteConfig = async (code, cfg) => {
    setSavingConfig(true)
    setConfigError(null)
    try {
      const targetTable = getClientesTableByProcess(activeProcess)
      if (targetTable === 'clientes_config_costa_rica') {
        await sql`
          UPDATE clientes_config_costa_rica
          SET fee          = ${cfg.fee},
              iva          = ${cfg.iva},
              banking      = ${cfg.banking},
              sin_usd      = ${cfg.sin_usd},
              es_liquidacion = ${cfg.es_liquidacion}
          WHERE code = ${code}
        `
      } else {
        await sql`
          UPDATE clientes_config
          SET fee          = ${cfg.fee},
              iva          = ${cfg.iva},
              banking      = ${cfg.banking},
              sin_usd      = ${cfg.sin_usd},
              es_liquidacion = ${cfg.es_liquidacion}
          WHERE code = ${code}
        `
      }
      await loadConfig()
      setEditando(null)
    } catch (err) {
      console.error('Error al guardar configuración:', err)
      setConfigError(`Error al guardar "${code}": ${err.message}`)
    } finally {
      setSavingConfig(false)
    }
  }

  // ── Restablecer todos los valores por defecto en DB ─────────────────────────
  const resetToDefaults = async () => {
    if (!window.confirm('¿Restablecer todos los clientes a la configuración inicial? Esto sobreescribirá los cambios guardados en la base de datos.')) return
    setResetting(true)
    setConfigError(null)
    try {
      const defaultsByProcess = getDefaultConfigByProcess(activeProcess)
      const targetTable = getClientesTableByProcess(activeProcess)
      for (const [code, cfg] of Object.entries(defaultsByProcess)) {
        if (targetTable === 'clientes_config_costa_rica') {
          await sql`
            INSERT INTO clientes_config_costa_rica (code, name, fee, iva, banking, sin_usd, es_liquidacion, activo)
            VALUES (${code}, ${code}, ${cfg.fee}, ${cfg.iva}, ${cfg.banking}, ${cfg.sin_usd}, ${cfg.es_liquidacion}, true)
            ON CONFLICT (code) DO UPDATE
            SET fee            = EXCLUDED.fee,
                iva            = EXCLUDED.iva,
                banking        = EXCLUDED.banking,
                sin_usd        = EXCLUDED.sin_usd,
                es_liquidacion = EXCLUDED.es_liquidacion,
                activo         = true
          `
        } else {
          await sql`
            INSERT INTO clientes_config (code, name, fee, iva, banking, sin_usd, es_liquidacion, activo)
            VALUES (${code}, ${code}, ${cfg.fee}, ${cfg.iva}, ${cfg.banking}, ${cfg.sin_usd}, ${cfg.es_liquidacion}, true)
            ON CONFLICT (code) DO UPDATE
            SET fee            = EXCLUDED.fee,
                iva            = EXCLUDED.iva,
                banking        = EXCLUDED.banking,
                sin_usd        = EXCLUDED.sin_usd,
                es_liquidacion = EXCLUDED.es_liquidacion,
                activo         = true
          `
        }
      }
      await loadConfig()
    } catch (err) {
      console.error('Error al restablecer configuración:', err)
      setConfigError(`Error al restablecer: ${err.message}`)
    } finally {
      setResetting(false)
    }
  }

  const addCliente = async () => {
    const code = addForm.code.trim().toUpperCase()
    if (!code) return setConfigError('El código del cliente es obligatorio.')
    if (!code.startsWith(currentProcess.clientPrefix.toUpperCase())) {
      return setConfigError(`Para ${currentProcess.label}, el código debe iniciar por "${currentProcess.clientPrefix}".`)
    }
    if (clientesConfig[code]) return setConfigError(`El cliente "${code}" ya existe.`)
    setAddingClient(true)
    setConfigError(null)
    try {
      const targetTable = getClientesTableByProcess(activeProcess)
      if (targetTable === 'clientes_config_costa_rica') {
        await sql`
          INSERT INTO clientes_config_costa_rica (code, name, fee, iva, banking, sin_usd, es_liquidacion, activo)
          VALUES (${code}, ${code}, ${addForm.fee || '10%'}, ${addForm.iva}, ${addForm.banking}, ${addForm.sin_usd}, ${addForm.es_liquidacion}, true)
        `
      } else {
        await sql`
          INSERT INTO clientes_config (code, name, fee, iva, banking, sin_usd, es_liquidacion, activo)
          VALUES (${code}, ${code}, ${addForm.fee || '10%'}, ${addForm.iva}, ${addForm.banking}, ${addForm.sin_usd}, ${addForm.es_liquidacion}, true)
        `
      }
      await loadConfig()
      setAddForm({ code: '', fee: '10%', iva: false, banking: false, sin_usd: false, es_liquidacion: false })
      setShowAddForm(false)
    } catch (err) {
      setConfigError(`Error al agregar cliente: ${err.message}`)
    } finally {
      setAddingClient(false)
    }
  }

  const addRemoSubcliente = async () => {
    const ccosto = newRemoSubcliente.ccosto.trim().toUpperCase()
    const descCcosto = newRemoSubcliente.descCcosto.trim()
    const nombreCompleto = newRemoSubcliente.nombreCompleto.trim()
    if (!ccosto) {
      setSubclientesError('El campo CCOSTO es obligatorio.')
      return
    }
    if (!descCcosto) {
      setSubclientesError('El campo DESC.CCOSTO es obligatorio.')
      return
    }
    const exists = remoSubclientes.some(item => item.ccosto.toUpperCase() === ccosto && item.descCcosto.toUpperCase() === descCcosto.toUpperCase())
    if (exists) {
      setSubclientesError('Ese subcliente ya existe en la configuración.')
      return
    }

    setAddingRemoSubcliente(true)
    setSubclientesError(null)
    try {
      await sql`
        INSERT INTO remofirst_subclientes_config (ccosto, desc_ccosto, nombre_completo, activo)
        VALUES (${ccosto}, ${descCcosto}, ${nombreCompleto || null}, true)
        ON CONFLICT (ccosto, desc_ccosto) DO UPDATE
        SET nombre_completo = EXCLUDED.nombre_completo,
            activo = true
      `
      await loadRemoSubclientes()
      setNewRemoSubcliente({ ccosto: '', descCcosto: '', nombreCompleto: '' })
      setShowRemoSubForm(false)
    } catch (err) {
      console.error('Error al agregar subcliente Remofirst:', err)
      setSubclientesError(`Error al guardar subcliente: ${err.message}`)
    } finally {
      setAddingRemoSubcliente(false)
    }
  }

  const removeRemoSubcliente = async (ccosto, descCcosto) => {
    const key = `${ccosto} | ${descCcosto}`
    if (!window.confirm(`¿Eliminar el subcliente ${key}?`)) return
    setDeletingRemoSubcliente(key)
    setSubclientesError(null)
    try {
      await sql`
        UPDATE remofirst_subclientes_config
        SET activo = false
        WHERE ccosto = ${ccosto} AND desc_ccosto = ${descCcosto}
      `
      await loadRemoSubclientes()
    } catch (err) {
      console.error('Error al eliminar subcliente Remofirst:', err)
      setSubclientesError(`Error al eliminar subcliente: ${err.message}`)
    } finally {
      setDeletingRemoSubcliente('')
    }
  }

  const startEditRemoSubcliente = (item) => {
    const key = `${item.ccosto}|${item.descCcosto}`
    setEditingRemoSubcliente(key)
    setEditRemoSubclienteForm({
      ccosto: item.ccosto,
      descCcosto: item.descCcosto,
      nombreCompleto: item.nombreCompleto || '',
    })
    setSubclientesError(null)
  }

  const cancelEditRemoSubcliente = () => {
    setEditingRemoSubcliente('')
    setEditRemoSubclienteForm({ ccosto: '', descCcosto: '', nombreCompleto: '' })
  }

  const saveRemoSubcliente = async (originalCcosto, originalDescCcosto) => {
    const ccosto = editRemoSubclienteForm.ccosto.trim().toUpperCase()
    const descCcosto = editRemoSubclienteForm.descCcosto.trim()
    const nombreCompleto = editRemoSubclienteForm.nombreCompleto.trim()

    if (!ccosto) return setSubclientesError('El campo CCOSTO es obligatorio.')
    if (!descCcosto) return setSubclientesError('El campo DESC.CCOSTO es obligatorio.')

    const originalKey = `${originalCcosto}|${originalDescCcosto}`
    const newKey = `${ccosto}|${descCcosto}`
    const duplicated = remoSubclientes.some(item => {
      const key = `${item.ccosto}|${item.descCcosto}`
      return key !== originalKey && key.toUpperCase() === newKey.toUpperCase()
    })
    if (duplicated) {
      setSubclientesError('Ya existe otro subcliente con ese CCOSTO y DESC.CCOSTO.')
      return
    }

    setSavingRemoSubcliente(true)
    setSubclientesError(null)
    try {
      await sql`
        UPDATE remofirst_subclientes_config
        SET ccosto = ${ccosto},
            desc_ccosto = ${descCcosto},
            nombre_completo = ${nombreCompleto || null}
        WHERE ccosto = ${originalCcosto} AND desc_ccosto = ${originalDescCcosto}
      `
      await loadRemoSubclientes()
      cancelEditRemoSubcliente()
    } catch (err) {
      console.error('Error al editar subcliente Remofirst:', err)
      setSubclientesError(`Error al editar subcliente: ${err.message}`)
    } finally {
      setSavingRemoSubcliente(false)
    }
  }

  const addEmpleadoConfig = async () => {
    const documento = normalizeDocumento(newEmpleado.documento)
    const nombre = newEmpleado.nombre.trim().toUpperCase()
    if (!documento) return setEmpleadosConfigError('El documento es obligatorio.')
    if (!nombre) return setEmpleadosConfigError('El nombre es obligatorio.')

    const exists = empleadosConfig.some(emp => normalizeDocumento(emp.documento) === documento)
    if (exists) return setEmpleadosConfigError(`Ya existe un empleado con documento ${documento}.`)

    setAddingEmpleado(true)
    setEmpleadosConfigError(null)
    try {
      await sql`
        INSERT INTO empleados_config_costa_rica (documento, nombre, activo)
        VALUES (${documento}, ${nombre}, true)
        ON CONFLICT (documento) DO UPDATE
        SET nombre = EXCLUDED.nombre,
            activo = true
      `
      await loadEmpleadosConfig()
      setNewEmpleado({ documento: '', nombre: '' })
      setShowEmpleadoForm(false)
    } catch (err) {
      console.error('Error al agregar empleado CR:', err)
      setEmpleadosConfigError(`Error al guardar empleado: ${err.message}`)
    } finally {
      setAddingEmpleado(false)
    }
  }

  const removeEmpleadoConfig = async (documento) => {
    if (!window.confirm(`¿Eliminar el empleado con documento ${documento}?`)) return
    setDeletingEmpleado(documento)
    setEmpleadosConfigError(null)
    try {
      await sql`
        UPDATE empleados_config_costa_rica
        SET activo = false
        WHERE documento = ${documento}
      `
      await loadEmpleadosConfig()
    } catch (err) {
      console.error('Error al eliminar empleado CR:', err)
      setEmpleadosConfigError(`Error al eliminar empleado: ${err.message}`)
    } finally {
      setDeletingEmpleado('')
    }
  }

  const startEditEmpleadoConfig = (empleado) => {
    setEditingEmpleado(empleado.documento)
    setEditEmpleadoForm({ documento: empleado.documento, nombre: empleado.nombre || '' })
    setEmpleadosConfigError(null)
  }

  const cancelEditEmpleadoConfig = () => {
    setEditingEmpleado('')
    setEditEmpleadoForm({ documento: '', nombre: '' })
  }

  const saveEmpleadoConfig = async (originalDocumento) => {
    const documento = normalizeDocumento(editEmpleadoForm.documento)
    const nombre = editEmpleadoForm.nombre.trim().toUpperCase()
    if (!documento) return setEmpleadosConfigError('El documento es obligatorio.')
    if (!nombre) return setEmpleadosConfigError('El nombre es obligatorio.')

    const duplicated = empleadosConfig.some(emp => {
      const doc = normalizeDocumento(emp.documento)
      return doc !== normalizeDocumento(originalDocumento) && doc === documento
    })
    if (duplicated) return setEmpleadosConfigError(`Ya existe un empleado con documento ${documento}.`)

    setSavingEmpleado(true)
    setEmpleadosConfigError(null)
    try {
      await sql`
        UPDATE empleados_config_costa_rica
        SET documento = ${documento},
            nombre = ${nombre}
        WHERE documento = ${originalDocumento}
      `
      await loadEmpleadosConfig()
      cancelEditEmpleadoConfig()
    } catch (err) {
      console.error('Error al editar empleado CR:', err)
      setEmpleadosConfigError(`Error al editar empleado: ${err.message}`)
    } finally {
      setSavingEmpleado(false)
    }
  }

  const toggleCliente = (cliente) => {
    setClientesSeleccionados(prev =>
      prev.includes(cliente) ? prev.filter(c => c !== cliente) : [...prev, cliente]
    )
  }

  const clientesFiltrados = Object.keys(clientesConfig).filter(c =>
    c.toUpperCase().startsWith(currentProcess.clientPrefix.toUpperCase()) && c.toLowerCase().includes(busqueda.toLowerCase())
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

    if (currentProcess.baseRequired && !baseEmpleados) return setError('Sube el archivo de base de empleados.')
    if (!reporteNovasoft) return setError('Sube el archivo de reporte Novasoft.')
    if (clientesSeleccionados.length === 0) return setError('Selecciona al menos un cliente.')
    if (!periodo.trim()) return setError('Escribe el periodo de facturación.')
    if (activeProcess === 'general') {
      if (!tasaCambio.trim()) return setError('Escribe la tasa de cambio USD → COP.')
      if (clientesSeleccionados.includes('C1055 - RIVERMATE') && !tasaCambioEur.trim())
        return setError('Escribe la tasa de cambio EUR → COP para RIVERMATE.')
    } else {
      if (!costaRicaExchangeRate.trim()) return setError('Escribe la tasa de cambio CRC → USD.')
    }

    setGenerando(true)
    try {
      let archivosGenerados = 0
      // --- 1. Leer Base de empleados ---
      let baseData = []
      let mapaCC = new Map()
      let mapaEmpleados = new Map()
      let headerBase = []
      let colEmp = -1
      let colCC = -1
      let colAlt = -1
      let colNombre = -1
      let colFIng = -1
      let colFRet = -1
      let colSubcli = -1

      if (currentProcess.useBaseEmployees) {
        const baseBuffer = await baseEmpleados.arrayBuffer()
        const baseWb = XLSX.read(baseBuffer, { type: 'array', cellDates: true })
        const baseSheet = baseWb.Sheets[baseWb.SheetNames[0]]
        baseData = XLSX.utils.sheet_to_json(baseSheet, { header: 1 })

        headerBase = (baseData[0] || []).map(h => String(h ?? '').trim().toUpperCase())
        colEmp    = headerBase.indexOf('CODIGO EMPLEADO')
        colCC     = headerBase.indexOf('CENTRO COSTOS')
        colAlt    = headerBase.indexOf('CODIGO ALTERNO')
        colNombre = headerBase.indexOf('NOMBRE')
        colFIng   = headerBase.indexOf('F_INGRESO')
        colFRet   = headerBase.indexOf('F_RETIRO')
        colSubcli = headerBase.indexOf('SUBCLIENTE')

        console.log('=== BASE DE EMPLEADOS ===')
        console.log('Encabezados encontrados:', headerBase)
        console.log('Columna CODIGO EMPLEADO (índice):', colEmp)
        console.log('Columna CENTRO COSTOS (índice):', colCC)
        console.log('Total filas (incluyendo encabezado):', baseData.length)

        if (colEmp === -1) throw new Error('No se encontró la columna "CODIGO EMPLEADO" en la base de empleados.')
        if (colCC  === -1) throw new Error('No se encontró la columna "CENTRO COSTOS" en la base de empleados.')

        // Mapa: códigoEmpleado → { cc, alt, nombre, fIngreso, fRetiro, subcli, clasif }
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
      }

      // --- 2. Leer Reporte Novasoft ---
      const novaBuffer = await reporteNovasoft.arrayBuffer()
      const novaWb = XLSX.read(novaBuffer, { type: 'array', raw: false })

      console.log('=== NOVASOFT ===')
      console.log('Hojas disponibles:', novaWb.SheetNames)

      const novaSheet = novaWb.Sheets[currentProcess.reportSheetName]
      if (!novaSheet) throw new Error(`No se encontró la hoja "${currentProcess.reportSheetName}" en el reporte Novasoft.`)

      const novaData = XLSX.utils.sheet_to_json(novaSheet, { header: 1, raw: false, defval: '' })

      if (activeProcess === 'costa-rica') {
        const normalizeText = (value) => normalizeMatchText(value)
        const normalizeCostaRicaKey = (value) => normalizeMatchText(value)
        const normalizeDocumentoCr = (value) => {
          const raw = String(value ?? '').trim()
          const digits = raw.replace(/\D/g, '')
          if (digits) return digits
          return raw.toUpperCase().replace(/\s+/g, '')
        }
        const stripLeadingZeros = (value) => {
          const text = String(value ?? '').trim()
          if (!text) return ''
          const stripped = text.replace(/^0+/, '')
          return stripped || '0'
        }
        const getDocumentoCandidates = (...values) => {
          const keys = new Set()
          values.forEach((value) => {
            const normalized = normalizeDocumentoCr(value)
            if (!normalized) return
            keys.add(normalized)
            const noLeadingZeros = stripLeadingZeros(normalized)
            if (noLeadingZeros) keys.add(noLeadingZeros)
          })
          return [...keys]
        }

        const headerRowIndex = 2
        const dataStartRowIndex = 3
        const crHeaders = (novaData[headerRowIndex] || []).map(h => String(h ?? '').trim())
        const crHeaderMap = new Map(crHeaders.map((h, idx) => [h.toUpperCase(), idx]))
        const getCrIndex = (name) => crHeaderMap.get(String(name).trim().toUpperCase()) ?? -1
        const getCrValue = (row, name) => {
          const idx = getCrIndex(name)
          return idx === -1 ? '' : row[idx]
        }
        const getCrValueAny = (row, names) => {
          for (const name of names) {
            const value = getCrValue(row, name)
            if (String(value ?? '').trim()) return value
          }
          return ''
        }
        const toNumber = (value) => {
          const parsed = parseFloat(String(value ?? '').replace(/[$,\s]/g, '').replace(/,/g, '.'))
          return Number.isFinite(parsed) ? parsed : 0
        }
        const getClientCode = (value) => String(value ?? '').trim().toUpperCase()
        const getSelectedClientCode = (cliente) => getClientCode(String(cliente).split(' - ')[0])
        const getSelectedClientDesc = (cliente) => {
          const parts = String(cliente).split(' - ')
          return parts.slice(1).join(' - ').trim()
        }
        const remofirstSubclientSet = new Set(
          remoSubclientes
            .map(item => `${normalizeCostaRicaKey(item.ccosto)}|${normalizeCostaRicaKey(item.descCcosto)}`)
            .filter(key => key !== '|')
        )
        const remofirstCustomerNameMap = new Map(
          remoSubclientes
            .map(item => [
              `${normalizeCostaRicaKey(item.ccosto)}|${normalizeCostaRicaKey(item.descCcosto)}`,
              String(item.nombreCompleto ?? '').trim(),
            ])
        )
        const documentosConTasaUno = new Set(
          empleadosConfig.map(emp => normalizeDocumentoCr(emp.documento)).filter(Boolean)
        )

        const requiresCodesFile = clientesSeleccionados.includes('CR009 - REMOFIRST INC')
        if (requiresCodesFile && !codigosRemo) {
          throw new Error('Para CR009 - REMOFIRST INC debes subir el archivo Codigos Remofirst.')
        }

        let codesEntries = []
        let codesExactMap = new Map()
        if (codigosRemo) {
          const codesBuffer = await codigosRemo.arrayBuffer()
          const codesWb = XLSX.read(codesBuffer, { type: 'array', raw: false, cellDates: true })
          const codesSheet = codesWb.Sheets[codesWb.SheetNames[0]]
          if (!codesSheet) throw new Error('No se encontró la hoja principal en el archivo Codigos Remofirst.')
          const codesData = XLSX.utils.sheet_to_json(codesSheet, { header: 1, defval: '' })
          const codesHeaders = (codesData[0] || []).map(h => String(h ?? '').trim())
          const headerIndex = (headers, name) => headers.findIndex(h => normalizeText(h) === normalizeText(name))
          const contractIdx = headerIndex(codesHeaders, 'Contract ID')
          const fullNameIdx = headerIndex(codesHeaders, 'Full name')
          const companyIdx = headerIndex(codesHeaders, 'Company ID')
          if (contractIdx === -1) throw new Error('No se encontró la columna "Contract ID" en el archivo Codigos Remofirst.')
          if (fullNameIdx === -1) throw new Error('No se encontró la columna "Full name" en el archivo Codigos Remofirst.')
          if (companyIdx === -1) throw new Error('No se encontró la columna "Company ID" en el archivo Codigos Remofirst.')

          for (let i = 1; i < codesData.length; i++) {
            const fila = codesData[i] || []
            const fullName = String(fila[fullNameIdx] ?? '').trim()
            const normalizedName = normalizeText(fullName)
            if (!normalizedName) continue
            const entry = {
              fullName,
              normalizedName,
              contractId: String(fila[contractIdx] ?? '').trim(),
              companyId: String(fila[companyIdx] ?? '').trim(),
            }
            codesEntries.push(entry)
            if (!codesExactMap.has(normalizedName)) {
              codesExactMap.set(normalizedName, entry)
            }
          }
        }

        const resolveRemoIdentity = async (employeeName) => {
          const normalizedEmployeeName = normalizeText(employeeName)
          if (!normalizedEmployeeName) return null

          const exactMatch = codesExactMap.get(normalizedEmployeeName)
          if (exactMatch) return exactMatch

          const candidates = getRemoMatchCandidates(employeeName, codesEntries, 5)
          if (candidates.length === 0) return null

          const bestCandidate = candidates[0]
          const secondCandidate = candidates[1]
          const shouldAutoMatch = bestCandidate.score >= REMO_MATCH_AUTOMATCH_THRESHOLD
            || (bestCandidate.score >= 0.72 && (!secondCandidate || (bestCandidate.score - secondCandidate.score) >= 0.12))

          if (shouldAutoMatch) return bestCandidate

          if (bestCandidate.score >= REMO_MATCH_REVIEW_THRESHOLD) {
            const selectedCandidate = await openNameMatchReview({
              employeeName,
              candidates,
              suggestedIndex: 0,
            })
            return selectedCandidate || null
          }

          return null
        }

        const crConceptMap = {
          SALARY: ['A000 - SALARIO ORDINARIO', 'A500 - DESCUENTO SALARIO ORDINARIO', 'A300 - RETROACTIVO SALARIO ORDINARIO'],
          'Transport allowance': ['A023 - AUXILIO DE TRANSPORTE', 'A102 - AUXILIO TRANSPORTE'],
          'Alloawance 2 (Mobile & Internet Allowance)': ['A024 - AUXILIO DE TELECOMUNICACIONES', 'A101 - AUXILIO TELEFONO'],
          'Bonus/Commission': ['A025 - BONO SALARIAL', 'A019 - COMISIONES'],
          'Alloawance 4 (Other allowances)': ['A028 - AUXILIO DE SALUD'],
          Overtime: ['C001 - HORA EXTRA DIURNA'],
          Holidays: ['G006 - VACACIONES DISFRUTADAS HABILES', 'G022 -VACACIONES DEFINITIVAS', 'B007 - LICENCIA PATERNIDAD', 'B001- LICENCIA REMUNERADA'],
          'X100 - APORTE SEG SOCIAL EMPLEADOR': ['H100 - APORTE SEG SOCIAL EMPLEADOR', 'X100 - APORTE SEG SOCIAL EMPLEADOR'],
          'X200 - APORTE INS EMPLEADOR': ['X200 - APORTE INS EMPLEADOR'],
          '13TH MONTH (AGUINALDO)': ['PROVISION AGUINALDO', 'G013 - AGUINALDO'],
        }

        const reportRows = []
        for (let i = dataStartRowIndex; i < novaData.length; i++) {
          const row = novaData[i]
          const code = String(getCrValue(row, 'EMPLEADO') ?? '').trim()
          if (!code) continue
          reportRows.push({
            row,
            code,
            clientCode: getClientCode(getCrValue(row, 'CCOSTO')),
            clientDesc: String(getCrValue(row, 'DESC.CCOSTO') ?? '').trim(),
            sourceRowNumber: i + 1,
          })
        }

        const sourceNovasoftWorkbook = new ExcelJS.Workbook()
        await sourceNovasoftWorkbook.xlsx.load(novaBuffer)
        const sourceNovasoftSheet = sourceNovasoftWorkbook.getWorksheet(currentProcess.reportSheetName)
        if (!sourceNovasoftSheet) {
          throw new Error(`No se encontró la hoja "${currentProcess.reportSheetName}" para copiar estilos en el reporte Novasoft.`)
        }

        const clonePlain = (value) => {
          if (value === null || value === undefined) return value
          if (value instanceof Date) return new Date(value.getTime())
          if (typeof value !== 'object') return value
          return JSON.parse(JSON.stringify(value))
        }

        const getSafeCopiedCellValue = (cellValue) => {
          if (cellValue === null || cellValue === undefined) return cellValue
          if (typeof cellValue !== 'object') return cellValue

          // Evita errores de ExcelJS al clonar formulas (incluidas shared formulas)
          // al copiar la hoja de Novasoft: conservar solo el resultado visible.
          if (Object.prototype.hasOwnProperty.call(cellValue, 'formula')) {
            if (Object.prototype.hasOwnProperty.call(cellValue, 'result')) {
              return clonePlain(cellValue.result)
            }
            return null
          }

          // Backward/edge compatibility for alternate shared-formula payloads.
          if (
            Object.prototype.hasOwnProperty.call(cellValue, 'sharedFormula')
            || Object.prototype.hasOwnProperty.call(cellValue, 'shareType')
          ) {
            if (Object.prototype.hasOwnProperty.call(cellValue, 'result')) {
              return clonePlain(cellValue.result)
            }
            return null
          }

          return clonePlain(cellValue)
        }

        const copyNovasoftSheetStyled = (targetWorkbook, sheetName) => {
          const targetSheet = targetWorkbook.addWorksheet(sheetName)
          targetSheet.properties = { ...sourceNovasoftSheet.properties }
          targetSheet.pageSetup = { ...sourceNovasoftSheet.pageSetup }
          targetSheet.headerFooter = { ...sourceNovasoftSheet.headerFooter }
          targetSheet.views = (sourceNovasoftSheet.views || []).map(v => ({ ...v }))
          targetSheet.state = sourceNovasoftSheet.state

          if (sourceNovasoftSheet.autoFilter) {
            targetSheet.autoFilter = clonePlain(sourceNovasoftSheet.autoFilter)
          }

          sourceNovasoftSheet.columns.forEach((col, idx) => {
            const targetCol = targetSheet.getColumn(idx + 1)
            targetCol.width = col.width
            targetCol.hidden = col.hidden
            targetCol.outlineLevel = col.outlineLevel
            targetCol.style = clonePlain(col.style) || {}
          })

          sourceNovasoftSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            const targetRow = targetSheet.getRow(rowNumber)
            targetRow.height = row.height
            targetRow.hidden = row.hidden
            targetRow.outlineLevel = row.outlineLevel
            targetRow.style = clonePlain(row.style) || {}

            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
              const targetCell = targetRow.getCell(colNumber)
              targetCell.value = getSafeCopiedCellValue(cell.value)
              targetCell.style = clonePlain(cell.style) || {}
              if (cell.numFmt) targetCell.numFmt = cell.numFmt
              if (cell.protection) targetCell.protection = clonePlain(cell.protection)
              if (cell.note) targetCell.note = clonePlain(cell.note)
            })
            targetRow.commit()
          })

          const mergeRanges = (sourceNovasoftSheet.model && sourceNovasoftSheet.model.merges) || []
          mergeRanges.forEach((range) => {
            try { targetSheet.mergeCells(range) } catch (e) {}
          })

          if (typeof sourceNovasoftSheet.getImages === 'function') {
            const sourceImages = sourceNovasoftSheet.getImages() || []
            sourceImages.forEach((img) => {
              try {
                const media = typeof sourceNovasoftWorkbook.getImage === 'function'
                  ? sourceNovasoftWorkbook.getImage(img.imageId)
                  : null
                if (!media) return
                const imageId = targetWorkbook.addImage(media)
                targetSheet.addImage(imageId, clonePlain(img.range))
              } catch (e) {}
            })
          }

          return targetSheet
        }

        const logFormulaObjects = (sheet, label) => {
          try {
            const samples = []
            sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
              row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                const value = cell.value
                if (!value || typeof value !== 'object') return
                if (Object.prototype.hasOwnProperty.call(value, 'formula') || Object.prototype.hasOwnProperty.call(value, 'sharedFormula') || Object.prototype.hasOwnProperty.call(value, 'shareType')) {
                  samples.push({
                    cell: `${XLSX.utils.encode_col(colNumber - 1)}${rowNumber}`,
                    keys: Object.keys(value),
                  })
                }
              })
            })
            if (samples.length > 0) {
              console.log(`[FormulaScan] ${label}:`, samples.slice(0, 10))
            }
          } catch (e) {}
        }

        const sanitizeSharedFormulaCells = (sheet, label) => {
          try {
            let sanitized = 0
            sheet.eachRow({ includeEmpty: false }, (row) => {
              row.eachCell({ includeEmpty: false }, (cell) => {
                const value = cell.value
                if (!value || typeof value !== 'object') return

                const isSharedFormula =
                  Object.prototype.hasOwnProperty.call(value, 'sharedFormula')
                  || Object.prototype.hasOwnProperty.call(value, 'shareType')
                  || (
                    Object.prototype.hasOwnProperty.call(value, 'formula')
                    && Object.prototype.hasOwnProperty.call(value, 'ref')
                  )

                if (!isSharedFormula) return

                cell.value = Object.prototype.hasOwnProperty.call(value, 'result')
                  ? clonePlain(value.result)
                  : null
                sanitized++
              })
            })

            if (sanitized > 0) {
              console.log(`[FormulaSanitize] ${label}:`, sanitized)
            }
          } catch (e) {}
        }

        const extractConcept = (row, conceptList) => conceptList.reduce((sum, conceptName) => {
          const idx = getCrIndex(conceptName)
          if (idx === -1) return sum
          return sum + toNumber(row[idx])
        }, 0)

        const safeNumber = (value) => {
          const parsed = parseFloat(String(value ?? '').replace(/,/g, '.'))
          return Number.isFinite(parsed) ? parsed : null
        }

        const getSheetMaxUsedCol = (sheet, maxRow = sheet.rowCount || 1) => {
          let maxUsedCol = 1
          for (let r = 1; r <= maxRow; r++) {
            sheet.getRow(r).eachCell({ includeEmpty: false }, (cell, colNumber) => {
              const value = cell.value
              if (value === null || value === undefined || value === '') return
              maxUsedCol = Math.max(maxUsedCol, colNumber)
            })
          }
          return maxUsedCol
        }

        const shouldUseHealthInsurance = Boolean(reporteValoresHealth)
        let healthInsuranceMap = new Map()
        if (shouldUseHealthInsurance) {
          if (!healthInsuranceMonth.trim()) {
            throw new Error('Selecciona el mes para la plantilla Valores Health.')
          }

          const healthBuffer = await reporteValoresHealth.arrayBuffer()
          const healthWb = XLSX.read(healthBuffer, { type: 'array', raw: false, defval: '' })
          const healthSheet = healthWb.Sheets['DETALLE']
          if (!healthSheet) {
            throw new Error('No se encontró la hoja "DETALLE" en la plantilla Valores Health.')
          }

          const healthData = XLSX.utils.sheet_to_json(healthSheet, { header: 1, raw: false, defval: '' })
          const healthHeaders = (healthData[0] || []).map(h => String(h ?? '').trim().toUpperCase())
          const documentCol = healthHeaders.indexOf('NRO_DOCUMENTO')
          const monthCol = healthHeaders.indexOf(healthInsuranceMonth.trim().toUpperCase())
          if (documentCol === -1) {
            throw new Error('No se encontró el encabezado "NRO_DOCUMENTO" en la plantilla Valores Health.')
          }
          if (monthCol === -1) {
            throw new Error(`No se encontró el mes "${healthInsuranceMonth}" en la plantilla Valores Health.`)
          }

          const normalizeHealthValue = (value) => {
            const text = String(value ?? '').trim()
            if (!text) return null
            const normalized = text
              .replace(/\s+/g, '')
              .replace(/[$€₡]/g, '')
              .replace(/,/g, '.')
            const parsed = parseFloat(normalized)
            return Number.isFinite(parsed) ? parsed : text
          }

          healthInsuranceMap = new Map()
          for (let i = 1; i < healthData.length; i++) {
            const row = healthData[i] || []
            const docCandidates = getDocumentoCandidates(row[documentCol])
            if (docCandidates.length === 0) continue
            const value = normalizeHealthValue(row[monthCol])
            if (value === null || value === '') continue
            docCandidates.forEach((docKey) => {
              healthInsuranceMap.set(docKey, value)
            })
          }
          console.log('[Health] registros cargados:', healthInsuranceMap.size, 'mes:', healthInsuranceMonth, 'muestra:', [...healthInsuranceMap.keys()].slice(0, 8))
        }

        const tplResponse = await fetch(currentProcess.templatePath)
        if (!tplResponse.ok) throw new Error(`No se pudo cargar la plantilla "${currentProcess.templatePath}" desde public/.`)
        const tplBuffer = await tplResponse.arrayBuffer()
        const workbook = new ExcelJS.Workbook()
        await workbook.xlsx.load(tplBuffer)
        const worksheet = workbook.worksheets[0]

        const headerRow = worksheet.getRow(3)
        const headerToCol = {}
        headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const key = String(cell.value ?? '').trim().toUpperCase()
          if (key) headerToCol[key] = colNumber
        })

        const fieldNames = {
          employeeCode: 'EMPLOYEE CODE',
          eeRfWid: 'EE RF WID',
          name: 'NAME',
          onboard: 'Onboarding Date',
          offboard: 'Offboarding Date',
          eeStatus: 'EE Status (onboarding/active/offboarding)',
          country: 'Country',
          customerName: 'Customer Name',
          customerId: 'Customer ID',
          payrollMonth: 'Payroll Month',
          serviceType: 'Service Type / Invoice Type',
          erSsRate: 'ER SS rate %',
          salary: 'SALARY',
          retro: 'Rectroactive payment/Plus Compensation',
          sickLeave: '"Sick" Leave',
          allowance1: 'Alloawance 1 (Car Allowance)',
          unusedHolidays: 'Unused Holidays',
          overtime: 'Overtime',
          bonus: 'Bonus/Commission',
          deduction: 'Deduction or Gross Amount adjustments prevous month ',
          allowance2: 'Alloawance 2 (Mobile & Internet Allowance)',
          foodAllowance: 'Food allowance',
          allowance3: 'Alloawance 3 (Home/Remote work allowance)',
          allowance4: 'Alloawance 4 (Other allowances)',
          holidays: 'Holidays',
          signOnBonus: 'Sign-on Bonus',
          transportAllowance: 'Transport allowance',
          wellnessAllowance: 'Wellness Allowance',
          onCall: 'On Call/ Plus Disponibilidad',
          severance: 'Severance Pay (Taxable)',
          lieuOfNotice: 'Lieu of Notice',
          maternity: 'Paternity/ Maternity leave',
          payments: 'PAYMENTS',
          basic: 'BASIC',
          x100: 'X100 - APORTE SEG SOCIAL EMPLEADOR',
          x200: 'X200 - APORTE INS EMPLEADOR',
          thirteenthMonth: '13TH MONTH (AGUINALDO)',
          total1: 'TOTAL',
          other1: 'OTHER',
          other2: 'OTHER',
          other3: 'OTHER',
          other4: 'OTHER',
          healthInsurance: 'HEALT INSURANCE',
          expensesReimbursement: 'EXPENSES REIMBURSEMENT',
          total2: 'TOTAL',
          totalEmployeeCost: 'TOTAL EMPLOYEE COST',
          fee: 'FEE',
          bankingTax: 'BANKING TAX',
          total3: 'TOTAL',
          exchangeRate: 'EXCHANGE RATE',
          totalEmployeeCostUsd: 'TOTAL EMPLOYEE COST USD',
          feeUsd: 'FEE USD',
          totalUsd: 'TOTAL USD',
        }

        for (const cliente of clientesSeleccionados) {
          const cfg = clientesConfig[cliente] || DEFAULT_CONFIG_CLIENTES[cliente] || {}
          const selectedCode = getSelectedClientCode(cliente)
          const selectedDesc = getSelectedClientDesc(cliente)
          const isRemofirst = selectedCode === 'CR009'
          const rowsCliente = reportRows.filter(item => {
            if (item.clientCode === selectedCode) {
              // Only apply DESC.CCOSTO matching for CCOSTO CR010 as requested
              if (selectedCode === 'CR010' && selectedDesc) {
                const normItemDesc = normalizeCostaRicaKey(item.clientDesc)
                const normSelectedDesc = normalizeCostaRicaKey(selectedDesc)
                if (normItemDesc && normSelectedDesc) return normItemDesc === normSelectedDesc
                // fallback to code-only match if descriptions can't be normalized
              }
              return true
            }
            if (!isRemofirst) return false
            const composedKey = `${normalizeCostaRicaKey(item.clientCode)}|${normalizeCostaRicaKey(item.clientDesc)}`
            return remofirstSubclientSet.has(composedKey)
          })
          if (rowsCliente.length === 0) continue

          const localWorkbook = new ExcelJS.Workbook()
          await localWorkbook.xlsx.load(tplBuffer)
          const localSheet = localWorkbook.worksheets[0]
          sanitizeSharedFormulaCells(localSheet, 'TemplateLocal')

          // En esta plantilla la fila 5 es la de totales.
          // Evitamos duplicateRow porque puede romper con shared formulas del template.
          // Insertamos filas vacías y copiamos solo estilo desde la fila detalle (4).
          if (rowsCliente.length > 1) {
            const insertCount = rowsCliente.length - 1
            localSheet.spliceRows(5, 0, ...Array.from({ length: insertCount }, () => []))

            const templateDetailRow = localSheet.getRow(4)
            const maxStyledCol = XLSX.utils.decode_col('BG') + 1

            for (let rowNum = 5; rowNum < 5 + insertCount; rowNum++) {
              const newRow = localSheet.getRow(rowNum)
              newRow.height = templateDetailRow.height
              for (let col = 1; col <= maxStyledCol; col++) {
                const sourceCell = templateDetailRow.getCell(col)
                const targetCell = newRow.getCell(col)
                targetCell.style = clonePlain(sourceCell.style) || {}
                if (sourceCell.numFmt) targetCell.numFmt = sourceCell.numFmt
                if (sourceCell.protection) targetCell.protection = clonePlain(sourceCell.protection)
                if (sourceCell.note) targetCell.note = clonePlain(sourceCell.note)
                targetCell.value = null
              }
              newRow.commit()
            }
          }

          const localHeaderRow = localSheet.getRow(3)
          const localHeaderToCol = {}
          const totalCols = []
          const otherCols = []
          localHeaderRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            const key = String(cell.value ?? '').trim().toUpperCase()
            if (key) localHeaderToCol[key] = colNumber
            if (key === 'TOTAL') totalCols.push(colNumber)
            if (key === 'OTHER') otherCols.push(colNumber)
          })
          const setByHeader = (row, header, value) => {
            const col = localHeaderToCol[String(header).toUpperCase()]
            if (!col) return
            row.getCell(col).value = value ?? null
          }

          const colLetter = (colNumber) => XLSX.utils.encode_col(colNumber - 1)
          const paymentsStartCol = localHeaderToCol['SALARY']
          const paymentsEndCol = localHeaderToCol['PATERNITY/ MATERNITY LEAVE']
          const firstTotalStartCol = localHeaderToCol['X100 - APORTE SEG SOCIAL EMPLEADOR']
          const firstTotalEndCol = localHeaderToCol['13TH MONTH (AGUINALDO)']
          const secondTotalStartCol = otherCols[0]
          const secondTotalEndCol = localHeaderToCol['EXPENSES REIMBURSEMENT']
          const exchangeRateCol = localHeaderToCol['EXCHANGE RATE']
          const bankingTaxCol = localHeaderToCol['BANKING TAX']
          const totalEmployeeCostCol = localHeaderToCol['TOTAL EMPLOYEE COST']

          let outputRow = 4
          for (const item of rowsCliente) {
            const row = localSheet.getRow(outputRow++)
            const rowNumber = row.number
            const source = item.row
            const salary = extractConcept(source, crConceptMap.SALARY)
            const allowanceTransport = extractConcept(source, crConceptMap['Transport allowance'])
            const allowance2 = extractConcept(source, crConceptMap['Alloawance 2 (Mobile & Internet Allowance)'])
            const bonus = extractConcept(source, crConceptMap['Bonus/Commission'])
            const allowance4 = extractConcept(source, crConceptMap['Alloawance 4 (Other allowances)'])
            const overtime = extractConcept(source, crConceptMap.Overtime)
            const holidays = extractConcept(source, crConceptMap.Holidays)
            const x100 = extractConcept(source, crConceptMap['X100 - APORTE SEG SOCIAL EMPLEADOR'])
            const x200 = extractConcept(source, crConceptMap['X200 - APORTE INS EMPLEADOR'])
            const thirteenthMonth = extractConcept(source, crConceptMap['13TH MONTH (AGUINALDO)'])

            const clientName = String(getCrValue(source, 'DESC.CCOSTO') ?? '').trim()
            const clientId = String(getCrValue(source, 'CCOSTO') ?? '').trim()
            const remofirstKey = `${normalizeCostaRicaKey(clientId)}|${normalizeCostaRicaKey(clientName)}`
            const configuredCustomerName = remofirstCustomerNameMap.get(remofirstKey)
            const customerNameValue = cliente === 'CR009 - REMOFIRST INC' && configuredCustomerName
              ? configuredCustomerName
              : clientName
            const employeeCode = String(getCrValue(source, 'EMPLEADO') ?? '').trim()
            const employeeName = String(getCrValue(source, 'NOMBRE COMPLETO') ?? '').trim()
            const employeeDocumentRaw = getCrValueAny(source, [
              'CEDULA', 'CÉDULA', 'NO. CEDULA', 'NO CEDULA', 'N° CEDULA',
              'DOCUMENTO', 'NUMERO DOCUMENTO', 'NRO DOCUMENTO', 'IDENTIFICACION', 'IDENTIFICACIÓN',
              'DOCUMENTO IDENTIDAD', 'DOC IDENTIDAD', 'NO DOCUMENTO IDENTIDAD',
            ])
            const employeeDocument = normalizeDocumentoCr(employeeDocumentRaw)
            const onboardingDate = getCrValue(source, 'FECHA ANTIGUEDAD')
            const exchangeRateValue = safeNumber(costaRicaExchangeRate)
            const rowExchangeRate = documentosConTasaUno.has(employeeDocument) ? 1 : (exchangeRateValue || 0)
            const healthInsuranceDocCandidates = getDocumentoCandidates(employeeCode)
            const healthInsuranceValue = shouldUseHealthInsurance
              ? (healthInsuranceDocCandidates.map(key => healthInsuranceMap.get(key)).find(value => value !== undefined) ?? null)
              : null
            if (shouldUseHealthInsurance) {
              console.log('[Health] match', {
                employeeCode,
                employeeDocumentRaw,
                employeeDocument,
                healthInsuranceDocCandidates,
                healthInsuranceValue,
                hasKey: healthInsuranceDocCandidates.some(key => healthInsuranceMap.has(key)),
              })
            }
            const payments = [salary, allowanceTransport, allowance2, bonus, allowance4, overtime, holidays].reduce((sum, value) => sum + value, 0)
            const bankingTaxValue = cfg.banking ? (25 * rowExchangeRate) : 0
            const totalEmployeeCost = payments + x100 + x200 + thirteenthMonth
            const feeValue = cfg.fee && String(cfg.fee).endsWith('%')
              ? totalEmployeeCost * (parseFloat(String(cfg.fee).replace('%', '')) / 100)
              : (safeNumber(cfg.fee) || 0) * rowExchangeRate
            const usdTotalEmployeeCost = (cfg.sin_usd || !rowExchangeRate) ? null : (totalEmployeeCost / rowExchangeRate)
            const usdFee = (cfg.sin_usd || !rowExchangeRate) ? null : (feeValue / rowExchangeRate)
            const identityMatch = cliente === 'CR009 - REMOFIRST INC'
              ? await resolveRemoIdentity(employeeName)
              : null

            setByHeader(row, fieldNames.employeeCode, employeeCode)
            setByHeader(row, fieldNames.eeRfWid, cliente === 'CR009 - REMOFIRST INC' ? (identityMatch?.contractId || '') : '')
            setByHeader(row, fieldNames.name, employeeName)
            setByHeader(row, fieldNames.onboard, onboardingDate)
            setByHeader(row, fieldNames.offboard, null)
            setByHeader(row, fieldNames.eeStatus, 'onboarding')
            setByHeader(row, fieldNames.country, currentProcess.countryLabel)
            setByHeader(row, fieldNames.customerName, customerNameValue)
            setByHeader(row, fieldNames.customerId, cliente === 'CR009 - REMOFIRST INC' ? (identityMatch?.companyId || '') : '')
            setByHeader(row, fieldNames.payrollMonth, periodo.trim())
            setByHeader(row, fieldNames.serviceType, 'Monthly Payroll')
            setByHeader(row, fieldNames.erSsRate, '27,83%')
            setByHeader(row, fieldNames.healthInsurance, healthInsuranceValue)
            setByHeader(row, fieldNames.salary, salary || null)
            setByHeader(row, fieldNames.transportAllowance, allowanceTransport || null)
            setByHeader(row, fieldNames.allowance2, allowance2 || null)
            setByHeader(row, fieldNames.bonus, bonus || null)
            setByHeader(row, fieldNames.allowance4, allowance4 || null)
            setByHeader(row, fieldNames.overtime, overtime || null)
            setByHeader(row, fieldNames.holidays, holidays || null)
            setByHeader(row, fieldNames.x100, x100 || null)
            setByHeader(row, fieldNames.x200, x200 || null)
            setByHeader(row, fieldNames.thirteenthMonth, thirteenthMonth || null)
            if (paymentsStartCol && paymentsEndCol && localHeaderToCol[fieldNames.payments.toUpperCase()]) {
              row.getCell(localHeaderToCol[fieldNames.payments.toUpperCase()]).value = {
                formula: `SUM(${colLetter(paymentsStartCol)}${rowNumber}:${colLetter(paymentsEndCol)}${rowNumber})`,
              }
            } else {
              setByHeader(row, fieldNames.payments, payments || null)
            }
            setByHeader(row, fieldNames.basic, salary || null)
            if (totalCols[0] && firstTotalStartCol && firstTotalEndCol) {
              row.getCell(totalCols[0]).value = {
                formula: `SUM(${colLetter(firstTotalStartCol)}${rowNumber}:${colLetter(firstTotalEndCol)}${rowNumber})`,
              }
            } else {
              setByHeader(row, fieldNames.total1, (salary + x100 + x200 + thirteenthMonth) || null)
            }
            if (totalCols[1] && secondTotalStartCol && secondTotalEndCol) {
              row.getCell(totalCols[1]).value = {
                formula: `SUM(${colLetter(secondTotalStartCol)}${rowNumber}:${colLetter(secondTotalEndCol)}${rowNumber})`,
              }
            } else {
              setByHeader(row, fieldNames.total2, totalEmployeeCost || null)
            }
            if (bankingTaxCol) {
              if (cfg.banking) {
                if (cfg.sin_usd) {
                  row.getCell(bankingTaxCol).value = {
                    formula: `25*${rowExchangeRate}`,
                  }
                } else {
                  row.getCell(bankingTaxCol).value = exchangeRateCol
                    ? { formula: `25*${colLetter(exchangeRateCol)}${rowNumber}` }
                    : { formula: `25*${rowExchangeRate}` }
                }
              } else {
                row.getCell(bankingTaxCol).value = null
              }
            } else {
              setByHeader(row, fieldNames.bankingTax, bankingTaxValue || null)
            }
            if (totalEmployeeCostCol) {
              const paymentsColLetter = localHeaderToCol[fieldNames.payments.toUpperCase()] ? colLetter(localHeaderToCol[fieldNames.payments.toUpperCase()]) : null
              const firstTotalColLetter = totalCols[0] ? colLetter(totalCols[0]) : null
              const secondTotalColLetter = totalCols[1] ? colLetter(totalCols[1]) : null
              const bankingTaxColLetter = bankingTaxCol ? colLetter(bankingTaxCol) : null
              if (paymentsColLetter && firstTotalColLetter && secondTotalColLetter && bankingTaxColLetter) {
                row.getCell(totalEmployeeCostCol).value = {
                  formula: `${paymentsColLetter}${rowNumber}+${firstTotalColLetter}${rowNumber}+${secondTotalColLetter}${rowNumber}+${bankingTaxColLetter}${rowNumber}`,
                }
              } else {
                setByHeader(row, fieldNames.totalEmployeeCost, totalEmployeeCost || null)
              }
            } else {
              setByHeader(row, fieldNames.totalEmployeeCost, totalEmployeeCost || null)
            }
            if (totalCols[2] && totalEmployeeCostCol && localHeaderToCol[fieldNames.fee.toUpperCase()]) {
              row.getCell(totalCols[2]).value = {
                formula: `${colLetter(totalEmployeeCostCol)}${rowNumber}+${colLetter(localHeaderToCol[fieldNames.fee.toUpperCase()])}${rowNumber}`,
              }
            } else {
              setByHeader(row, fieldNames.total3, (totalEmployeeCost + feeValue) || null)
            }
            const feeCol = localHeaderToCol[fieldNames.fee.toUpperCase()]
            if (feeCol) {
              const feeSource = String(cfg.fee ?? '').trim()
              if (feeSource.endsWith('%') && totalEmployeeCostCol) {
                const feeRate = (parseFloat(feeSource.replace('%', '')) || 0) / 100
                row.getCell(feeCol).value = {
                  formula: `${colLetter(totalEmployeeCostCol)}${rowNumber}*${feeRate}`,
                }
              } else {
                const feeFixed = safeNumber(feeSource) || 0
                if (!cfg.sin_usd && exchangeRateCol) {
                  row.getCell(feeCol).value = {
                    formula: `${feeFixed}*${colLetter(exchangeRateCol)}${rowNumber}`,
                  }
                } else {
                  row.getCell(feeCol).value = {
                    formula: `${feeFixed}*${rowExchangeRate}`,
                  }
                }
              }
            } else {
              setByHeader(row, fieldNames.fee, feeValue || null)
            }
            setByHeader(row, fieldNames.exchangeRate, rowExchangeRate || null)
            if (!cfg.sin_usd) {
              if (localHeaderToCol[fieldNames.totalEmployeeCostUsd.toUpperCase()] && exchangeRateCol) {
                row.getCell(localHeaderToCol[fieldNames.totalEmployeeCostUsd.toUpperCase()]).value = {
                  formula: `ROUND(${colLetter(totalEmployeeCostCol)}${rowNumber}/${colLetter(exchangeRateCol)}${rowNumber},2)`,
                }
              } else {
                setByHeader(row, fieldNames.totalEmployeeCostUsd, usdTotalEmployeeCost)
              }
              if (localHeaderToCol[fieldNames.feeUsd.toUpperCase()] && exchangeRateCol) {
                row.getCell(localHeaderToCol[fieldNames.feeUsd.toUpperCase()]).value = {
                  formula: `${colLetter(localHeaderToCol[fieldNames.fee.toUpperCase()])}${rowNumber}/${colLetter(exchangeRateCol)}${rowNumber}`,
                }
              } else {
                setByHeader(row, fieldNames.feeUsd, usdFee)
              }
              if (localHeaderToCol[fieldNames.totalUsd.toUpperCase()]) {
                const totalUsdCol = localHeaderToCol[fieldNames.totalUsd.toUpperCase()]
                const totalEmployeeCostUsdCol = localHeaderToCol[fieldNames.totalEmployeeCostUsd.toUpperCase()]
                const feeUsdCol = localHeaderToCol[fieldNames.feeUsd.toUpperCase()]
                if (totalEmployeeCostUsdCol && feeUsdCol) {
                  row.getCell(totalUsdCol).value = {
                    formula: `SUM(${colLetter(totalEmployeeCostUsdCol)}${rowNumber}:${colLetter(feeUsdCol)}${rowNumber})`,
                  }
                } else {
                  setByHeader(row, fieldNames.totalUsd, (usdTotalEmployeeCost || 0) + (usdFee || 0))
                }
              }
            }
            row.commit()
          }

          const totalRow = localSheet.getRow(outputRow)
          const detailStartRow = 4
          const detailEndRow = outputRow - 1
          const normalizeTotalSumFormula = (formula, colLetter) => {
            if (!formula) return null
            const rawFormula = String(formula).replace(/^=/, '').replace(/^\+/, '').trim()
            if (!/(^|[^A-Z])(SUM|SUMA)\s*\(/i.test(rawFormula)) return null
            return `SUM(${colLetter}${detailStartRow}:${colLetter}${detailEndRow})`
          }

          ;['PAYMENTS', 'BASIC', 'X100 - APORTE SEG SOCIAL EMPLEADOR', 'X200 - APORTE INS EMPLEADOR', '13TH MONTH (AGUINALDO)', 'TOTAL EMPLOYEE COST', 'FEE', 'BANKING TAX', 'TOTAL EMPLOYEE COST USD', 'FEE USD', 'TOTAL USD'].forEach(header => {
            const col = localHeaderToCol[header]
            if (!col) return
            const colLetter = XLSX.utils.encode_col(col - 1)
            totalRow.getCell(col).value = { formula: `SUM(${colLetter}${detailStartRow}:${colLetter}${detailEndRow})` }
          })

          totalRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            const currentValue = cell.value
            const currentFormula = currentValue && typeof currentValue === 'object' && 'formula' in currentValue
              ? currentValue.formula
              : null
            const colLetter = XLSX.utils.encode_col(colNumber - 1)
            const normalizedFormula = normalizeTotalSumFormula(currentFormula, colLetter)
            if (normalizedFormula) {
              cell.value = { formula: normalizedFormula }
            }
          })

          totalRow.commit()

          if (cfg.sin_usd) {
            const removeCols = [
              'EXCHANGE RATE',
              'TOTAL EMPLOYEE COST USD',
              'FEE USD',
              'TOTAL USD',
            ]
              .filter(Boolean)
              .map(name => localHeaderToCol[name.toUpperCase()])
              .filter(Boolean)
              .sort((a, b) => b - a)
            removeCols.forEach(col => localSheet.spliceColumns(col, 1))

            const headerAfterSplice = {}
            localSheet.getRow(3).eachCell({ includeEmpty: true }, (cell, colNumber) => {
              const key = String(cell.value ?? '').trim().toUpperCase()
              if (key) headerAfterSplice[key] = colNumber
            })

            const totalsStartCol = headerAfterSplice['TOTAL EMPLOYEE COST']
            const totalsEndCol = [
              headerAfterSplice['TOTAL EMPLOYEE COST'],
              headerAfterSplice['FEE'],
              headerAfterSplice['BANKING TAX'],
              headerAfterSplice['TOTAL'],
            ]
              .filter(Boolean)
              .sort((a, b) => a - b)
              .pop()

            if (totalsStartCol) {
              const parseCellRef = (ref) => {
                const match = String(ref || '').match(/^([A-Z]+)(\d+)$/)
                if (!match) return null
                return { col: XLSX.utils.decode_col(match[1]) + 1, row: parseInt(match[2], 10) }
              }
              const mergeRanges = (localSheet.model && localSheet.model.merges) || []
              mergeRanges.forEach((rangeStr) => {
                const [startRef, endRefRaw] = String(rangeStr).split(':')
                const start = parseCellRef(startRef)
                const end = parseCellRef(endRefRaw || startRef)
                if (!start || !end) return
                if (start.row <= 2 && end.row >= 2 && end.col >= totalsStartCol && start.col <= (totalsEndCol || totalsStartCol)) {
                  try { localSheet.unMergeCells(rangeStr) } catch (e) {}
                }
              })

              const totalsRow = localSheet.getRow(2)
              totalsRow.getCell(totalsStartCol).value = 'TOTALS'
              if (totalsEndCol && totalsEndCol > totalsStartCol) {
                try {
                  localSheet.mergeCells(
                    `${XLSX.utils.encode_col(totalsStartCol - 1)}2:${XLSX.utils.encode_col(totalsEndCol - 1)}2`,
                  )
                } catch (e) {}
              }
              totalsRow.commit()
            }
          }

          // Agrega una copia del reporte Novasoft con la misma estructura visual
          // y elimina filas que no pertenecen al cliente seleccionado.
          const novasoftSheetName = 'rpt_InformeNominaMes'
          const existingNovasoftSheet = localWorkbook.getWorksheet(novasoftSheetName)
          if (existingNovasoftSheet) {
            localWorkbook.removeWorksheet(existingNovasoftSheet.id)
          }
          const novasoftSheet = copyNovasoftSheetStyled(localWorkbook, novasoftSheetName)
          logFormulaObjects(novasoftSheet, 'NovasoftCopiada')
          const selectedSourceRows = new Set(rowsCliente.map(item => item.sourceRowNumber))
          const firstDetailExcelRow = dataStartRowIndex + 1
          for (let excelRow = novasoftSheet.rowCount; excelRow >= firstDetailExcelRow; excelRow--) {
            if (!selectedSourceRows.has(excelRow)) {
              novasoftSheet.spliceRows(excelRow, 1)
            }
          }

          const novasoftMaxUsedCol = getSheetMaxUsedCol(novasoftSheet)
          novasoftSheet.spliceColumns(novasoftMaxUsedCol + 1, 20000)

          // Recortar columnas finales realmente vacías (evita "colas" de columnas innecesarias)
          const maxUsedCol = getSheetMaxUsedCol(localSheet, outputRow)
          if (localSheet.columnCount > maxUsedCol) {
            localSheet.spliceColumns(maxUsedCol + 1, localSheet.columnCount - maxUsedCol)
          }

          // En Costa Rica la plantilla termina en BG; nunca dejar columnas adicionales.
          const maxTemplateCol = XLSX.utils.decode_col('BG') + 1
          if (localSheet.columnCount > maxTemplateCol) {
            localSheet.spliceColumns(maxTemplateCol + 1, localSheet.columnCount - maxTemplateCol)
          }

          // Intento de insertar una imagen del proyecto en la hoja Novasoft (A1)
          try {
            const tryImagePaths = [
              '/Imagen%20syp.jpg',
              '/Imagen syp.jpg',
              '/Logo%20syp.png',
              '/Logo syp.png',
            ]

            const arrayBufferToBase64 = (buffer) => {
              const bytes = new Uint8Array(buffer)
              const chunkSize = 0x8000
              let binary = ''
              for (let i = 0; i < bytes.length; i += chunkSize) {
                const chunk = bytes.subarray(i, i + chunkSize)
                binary += String.fromCharCode.apply(null, chunk)
              }
              return btoa(binary)
            }

            for (const imgPath of tryImagePaths) {
              try {
                const resp = await fetch(imgPath)
                if (!resp.ok) continue
                const ab = await resp.arrayBuffer()
                const lower = imgPath.toLowerCase()
                const ext = lower.endsWith('.png') ? 'png' : (lower.endsWith('.jpg') || lower.endsWith('.jpeg') ? 'jpeg' : 'png')
                const base64 = arrayBufferToBase64(ab)
                try {
                  const imageId = localWorkbook.addImage({ base64, extension: ext })
                  // Coloca la imagen en A1 (tl col 0,row 0). Ajusta tamaño si es necesario.
                  // imagen un poco más pequeña (ancho x alto)
                  novasoftSheet.addImage(imageId, { tl: { col: 0, row: 0 }, ext: { width: 135, height: 36 } })
                  console.log('[ImageInsert] inserted', imgPath)
                  break
                } catch (e) {
                  console.warn('[ImageInsert] addImage failed for', imgPath, e)
                }
              } catch (e) {
                // sigue probando otras rutas
              }
            }
          } catch (e) {
            console.warn('[ImageInsert] unexpected error', e)
          }

          let outBuffer = await localWorkbook.xlsx.writeBuffer()

          // Excel puede conservar definiciones <col max="16380"> aunque la data termine en BG.
          // Esto hace que visualmente aparezcan columnas vacías hasta XFD.
          // Se normaliza el XML de cada hoja para truncar los <col> a su ancho real.
          const clampSheetXmlColumns = (sheetXml, maxCol) => {
            let xml = String(sheetXml ?? '')
            xml = xml.replace(/<cols>[\s\S]*?<\/cols>/g, (colsBlock) => {
              const colTags = [...colsBlock.matchAll(/<col[^>]*\/>/g)].map(m => m[0])
              const kept = []
              colTags.forEach((tag) => {
                const minMatch = tag.match(/min="(\d+)"/)
                const maxMatch = tag.match(/max="(\d+)"/)
                if (!minMatch || !maxMatch) {
                  kept.push(tag)
                  return
                }
                const min = parseInt(minMatch[1], 10)
                const max = parseInt(maxMatch[1], 10)
                if (min > maxCol) return
                const clampedMax = Math.min(max, maxCol)
                kept.push(tag.replace(/max="\d+"/, `max="${clampedMax}"`))
              })
              return kept.length > 0 ? `<cols>${kept.join('')}</cols>` : ''
            })
            return xml
          }

          const zipWorkbook = await JSZip.loadAsync(outBuffer)
          const sheetXmlMaxCols = {
            'xl/worksheets/sheet1.xml': maxTemplateCol,
            'xl/worksheets/sheet2.xml': novasoftMaxUsedCol,
          }
          for (const [sheetPath, maxCol] of Object.entries(sheetXmlMaxCols)) {
            if (!zipWorkbook.file(sheetPath)) continue
            const xml = await zipWorkbook.file(sheetPath).async('string')
            zipWorkbook.file(sheetPath, clampSheetXmlColumns(xml, maxCol))
          }
          outBuffer = await zipWorkbook.generateAsync({ type: 'arraybuffer' })

          const blob = new Blob([outBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
          const url = URL.createObjectURL(blob)
          const a = document.createElement('a')
          a.href = url
          a.download = `Facturación EOR CR - ${cliente.replace(/[\/\?%*:|"<>]/g, '-')}- ${periodo.trim().replace(/[\/\?%*:|"<>]/g, '-')}.xlsx`
          document.body.appendChild(a)
          a.click()
          document.body.removeChild(a)
          URL.revokeObjectURL(url)
          archivosGenerados++
          await new Promise(r => setTimeout(r, 400))
        }

        if (archivosGenerados === 0) throw new Error('No se encontraron filas para ninguno de los clientes seleccionados en el reporte de Costa Rica.')

        setExitoCount(archivosGenerados)
        setTimeout(() => setExitoCount(0), 5000)
        setGenerando(false)
        return
      }

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
          columnaDestino: 'Alloawance 3 (Home/Remote work allowance)',
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
      const tplResponse = await fetch(currentProcess.templatePath)
      if (!tplResponse.ok) throw new Error(`No se pudo cargar la plantilla "${currentProcess.templatePath}" desde public/.`)
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

      // --- 4. Generar un archivo por cada cliente seleccionado ---
      for (const cliente of clientesSeleccionados) {
        const cfg = clientesConfig[cliente] || DEFAULT_CONFIG_CLIENTES[cliente] || {}
        // Flags derivados de la configuración (DB o defecto)
        const esLiquidacion = !!cfg.es_liquidacion

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

        const sinUSD = !!cfg.sin_usd
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
          ...['LEAVE PAID REFOUND', 'PARKING COST'].map(h => postCols[h]).filter(c => c && c !== -1),
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

      <section className="process-switcher">
        <div className="container">
          <div className="process-switcher-inner">
            <div>
              <p className="process-switcher-label">Selecciona el proceso</p>
              <h2 className="process-switcher-title">General - German / Costa Rica - Tatiana</h2>
            </div>
            <div className="process-switcher-tabs">
              <button
                type="button"
                className={`process-switch-btn ${activeProcess === 'general' ? 'active' : ''}`}
                onClick={() => setActiveProcess('general')}
              >
                General - German
              </button>
              <button
                type="button"
                className={`process-switch-btn ${activeProcess === 'costa-rica' ? 'active' : ''}`}
                onClick={() => setActiveProcess('costa-rica')}
              >
                Costa Rica - Tatiana
              </button>
            </div>
          </div>
        </div>
      </section>


      {nameMatchReview && (
        <div className="modal-overlay" role="dialog" aria-modal="true" aria-labelledby="name-match-review-title">
          <div className="modal-content match-review">
            <div className="match-review-header">
              <div className="modal-icon review">
                <svg width="42" height="42" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M12 2v20"/>
                  <path d="M17 6H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>
                </svg>
              </div>
              <div>
                <h3 className="modal-title" id="name-match-review-title">Revisar coincidencia de nombre</h3>
                <p className="modal-message">
                  Se encontró una coincidencia parcial para <strong>{nameMatchReview.employeeName}</strong>.
                  Confirma si corresponde o elige otra opción antes de seguir generando el archivo.
                </p>
              </div>
            </div>

            <div className="match-review-list">
              {nameMatchReview.candidates.map((candidate, index) => (
                <label
                  key={`${candidate.fullName}-${index}`}
                  className={`match-review-option ${nameMatchReview.selectedIndex === index ? 'selected' : ''}`}
                >
                  <input
                    type="radio"
                    name="remo-name-match"
                    checked={nameMatchReview.selectedIndex === index}
                    onChange={() => setNameMatchReview(prev => prev ? { ...prev, selectedIndex: index } : prev)}
                  />
                  <div className="match-review-option-main">
                    <div className="match-review-option-title">{candidate.fullName}</div>
                    <div className="match-review-option-meta">
                      Confianza {Math.round(candidate.score * 100)}% · {candidate.contractId || 'Contract ID vacío'}
                    </div>
                  </div>
                </label>
              ))}
              {nameMatchReview.candidates.length === 0 && (
                <div className="match-review-empty">
                  No hay candidatos para comparar.
                </div>
              )}
            </div>

            <div className="match-review-actions">
              <button type="button" className="modal-button secondary" onClick={() => closeNameMatchReview(null)}>
                Omitir
              </button>
              <button
                type="button"
                className="modal-button"
                onClick={() => closeNameMatchReview(nameMatchReview.candidates[nameMatchReview.selectedIndex] || null)}
                disabled={nameMatchReview.candidates.length === 0}
              >
                Usar coincidencia
              </button>
            </div>
          </div>
        </div>
      )}
      {/* Navegación por pestañas */}
      <nav className="tab-nav">
        <div className="container">
          <div className="tab-bar">
            <button
              className={`tab-btn ${activeTab === 'facturacion' ? 'active' : ''}`}
              onClick={() => setActiveTab('facturacion')}
            >
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                <polyline points="14 2 14 8 20 8"/>
                <line x1="16" y1="13" x2="8" y2="13"/>
                <line x1="16" y1="17" x2="8" y2="17"/>
              </svg>
              Facturación
            </button>
            <button
              className={`tab-btn ${activeTab === 'config' ? 'active' : ''}`}
              onClick={() => setActiveTab('config')}
            >
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="12" cy="12" r="3"/>
                <path d="M19.07 4.93a10 10 0 0 1 0 14.14M4.93 4.93a10 10 0 0 0 0 14.14"/>
                <path d="M12 2v2M12 20v2M2 12h2M20 12h2"/>
              </svg>
              Configuración de Clientes
              {configLoading && <span className="tab-badge">Cargando…</span>}
              {!configLoading && !configError && <span className="tab-badge ok">DB</span>}
              {configError && <span className="tab-badge err">!</span>}
            </button>
          </div>
        </div>
      </nav>

      {activeTab === 'facturacion' && (
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
                    Periodo de facturación ({currentProcess.label})
                  </label>
                  <input
                    type="text"
                    placeholder="Ej: Enero 2026 — se usará en el nombre de los archivos generados"
                    className="select-input"
                    value={periodo}
                    onChange={e => { setPeriodo(e.target.value); setError(null) }}
                  />
                </div>

                {currentProcess.useBaseEmployees && (
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
                )}

                {!currentProcess.useBaseEmployees && (
                <div className="form-group">
                  <label className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                      <polyline points="17 8 12 3 7 8"/>
                      <line x1="12" y1="3" x2="12" y2="15"/>
                    </svg>
                    Codigos Remofirst
                  </label>
                  <input
                    ref={inputCodes}
                    type="file"
                    accept=".xlsx,.xls"
                    className="file-input"
                    onChange={(e) => handleFileChange(e, setCodigosRemo)}
                  />
                  <div
                    className={`drop-zone ${dragCodes ? 'drag-active' : ''} ${codigosRemo ? 'has-file' : ''}`}
                    onClick={() => inputCodes.current.click()}
                    onDragOver={(e) => { e.preventDefault(); setDragCodes(true) }}
                    onDragLeave={() => setDragCodes(false)}
                    onDrop={(e) => handleFileDrop(e, setCodigosRemo, setDragCodes)}
                  >
                    {codigosRemo ? (
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
                          <p className="file-name">{codigosRemo.name}</p>
                          <p className="file-size">{formatSize(codigosRemo.size)}</p>
                        </div>
                        <button
                          className="btn-remove"
                          onClick={(e) => { e.stopPropagation(); setCodigosRemo(null); if (inputCodes.current) inputCodes.current.value = '' }}
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
                          <p className="drop-zone-title">Sube el archivo Codigos Remofirst</p>
                          <p className="drop-zone-subtitle">Se usa para CR009 - REMOFIRST INC</p>
                        </div>
                        <p className="drop-zone-hint">.xlsx / .xls</p>
                      </div>
                    )}
                  </div>
                </div>
                )}

                {/* Drop zone: Reporte Novasoft */}
                <div className="form-group">
                  <label className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <rect x="3" y="3" width="18" height="18" rx="2"/>
                      <path d="M9 3v18M3 9h18M3 15h18"/>
                    </svg>
                    Reporte Novasoft ({currentProcess.label})
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

                {activeProcess === 'costa-rica' && (
                  <div className="form-group">
                    <label className="label">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <rect x="3" y="3" width="18" height="18" rx="2"/>
                        <path d="M3 9h18M9 3v18"/>
                      </svg>
                      Plantilla Valores Health (opcional)
                    </label>
                    <input
                      ref={inputHealth}
                      type="file"
                      accept=".xlsx,.xls"
                      className="file-input"
                      onChange={(e) => handleFileChange(e, setReporteValoresHealth)}
                    />
                    <div
                      className={`drop-zone ${dragHealth ? 'drag-active' : ''} ${reporteValoresHealth ? 'has-file' : ''}`}
                      onClick={() => inputHealth.current.click()}
                      onDragOver={(e) => { e.preventDefault(); setDragHealth(true) }}
                      onDragLeave={() => setDragHealth(false)}
                      onDrop={(e) => handleFileDrop(e, setReporteValoresHealth, setDragHealth)}
                    >
                      {reporteValoresHealth ? (
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
                            <p className="file-name">{reporteValoresHealth.name}</p>
                            <p className="file-size">{formatSize(reporteValoresHealth.size)}</p>
                          </div>
                          <button
                            className="btn-remove"
                            onClick={(e) => { e.stopPropagation(); setReporteValoresHealth(null); setHealthInsuranceMonth(''); setDragHealth(false); if (inputHealth.current) inputHealth.current.value = '' }}
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
                            <p className="drop-zone-title">Sube la plantilla Valores Health</p>
                            <p className="drop-zone-subtitle">Opcional. Cruza EMPLEADO con NRO_DOCUMENTO para llenar HEALT INSURANCE</p>
                          </div>
                          <p className="drop-zone-hint">.xlsx / .xls</p>
                        </div>
                      )}
                    </div>
                    <label className="label" style={{ marginTop: 14 }}>
                      Mes para Health Insurance
                    </label>
                    <select
                      className="select-input"
                      value={healthInsuranceMonth}
                      onChange={(e) => setHealthInsuranceMonth(e.target.value)}
                      disabled={!reporteValoresHealth}
                    >
                      <option value="">Selecciona un mes</option>
                      {MONTH_OPTIONS.map(month => (
                        <option key={month} value={month}>{month}</option>
                      ))}
                    </select>
                  </div>
                )}

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
                      Tasa de cambio ({currentProcess.currencyLabel})
                    </label>
                    <input
                      type="number"
                      placeholder="Ej: 4200"
                      className="select-input"
                      value={activeProcess === 'general' ? tasaCambio : costaRicaExchangeRate}
                      onChange={e => {
                        if (activeProcess === 'general') setTasaCambio(e.target.value)
                        else setCostaRicaExchangeRate(e.target.value)
                        setError(null)
                      }}
                    />
                  </div>
                )}

                {/* Tasa de cambio EUR→COP — solo si RIVERMATE está seleccionado */}
                {activeProcess === 'general' && clientesSeleccionados.includes('C1055 - RIVERMATE') && (
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
      )}

      {activeTab === 'config' && (
      <section className="admin-section config-tab">
        <div className="container">

          {/* ── Glosario de columnas ── */}
          <div className="admin-glossary">
            <h3 className="glossary-title">
              <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="12" cy="12" r="10"/>
                <line x1="12" y1="16" x2="12" y2="12"/>
                <line x1="12" y1="8" x2="12.01" y2="8"/>
              </svg>
              Guía de columnas — ¿qué significa cada configuración?
            </h3>
            <div className="glossary-grid">

              <div className="glossary-item">
                <div className="glossary-col-name"><span className="badge-fee">Fee</span></div>
                <p>Es el cargo de gestión que S&amp;P cobra al cliente por empleado. Puede ser un <strong>porcentaje</strong> del costo total del empleado (ej: <code>10%</code>) o un <strong>valor fijo en USD</strong> que se multiplica por la tasa de cambio del periodo (ej: <code>120</code> = USD 120 × TRM). Si el cliente no necesita conversión USD, el valor fijo se multiplica directamente por la tasa EUR si aplica.</p>
              </div>

              <div className="glossary-item">
                <div className="glossary-col-name">
                  <span className="dot dot-on glossary-dot"/><span className="dot dot-off glossary-dot"/> IVA
                </div>
                <p><strong><span className="dot dot-on glossary-dot"/> Verde (activo):</strong> se genera la columna <em>IVA</em> en el Excel con la fórmula <code>Fee × 19%</code>, y además la columna <em>VAT</em> (IVA convertido a USD). <strong><span className="dot dot-off glossary-dot"/> Gris (inactivo):</strong> ambas columnas se eliminan del archivo generado.</p>
              </div>

              <div className="glossary-item">
                <div className="glossary-col-name">
                  <span className="dot dot-on glossary-dot"/><span className="dot dot-off glossary-dot"/> Banking Tax
                </div>
                <p><strong><span className="dot dot-on glossary-dot"/> Verde (activo):</strong> se calcula el impuesto bancario GMF (gravamen 4×1000) como <code>(Pagos + Costos SS + Provisiones + Otros) × 0.4%</code> y aparece como columna en el Excel. <strong><span className="dot dot-off glossary-dot"/> Gris (inactivo):</strong> la columna Banking Tax no se genera.</p>
              </div>

              <div className="glossary-item">
                <div className="glossary-col-name">
                  <span className="dot dot-on glossary-dot"/><span className="dot dot-off glossary-dot"/> Sin USD
                </div>
                <p><strong><span className="dot dot-on glossary-dot"/> Verde (activo):</strong> el cliente factura <em>solo en COP</em> — se eliminan del Excel las columnas <em>Exchange Rate</em>, <em>Total Employee Cost USD</em>, <em>Fee USD</em>, <em>VAT</em> y <em>Total USD</em>. <strong><span className="dot dot-off glossary-dot"/> Gris (inactivo):</strong> se incluyen todas las columnas en USD usando la tasa de cambio ingresada al generar.</p>
              </div>

              <div className="glossary-item">
                <div className="glossary-col-name">
                  <span className="dot dot-on glossary-dot"/><span className="dot dot-off glossary-dot"/> Es Liquidación
                </div>
                <p><strong><span className="dot dot-on glossary-dot"/> Verde (activo):</strong> las columnas de <em>Legal Benefits</em> (13th Salary, 14th Salary e Interest on 14th Salary) <strong>no aparecen</strong> en el Excel generado. <strong><span className="dot dot-off glossary-dot"/> Gris (inactivo):</strong> las columnas de <em>Legal Benefits</em> sí se incluyen en el archivo.</p>
              </div>

            </div>
          </div>

          {/* ── Tabla de configuración ── */}
          <div className="admin-card">
            <div className="admin-body">

              {configError && (
                <div className="alert alert-error" style={{marginBottom:'1rem'}}>
                  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
                  </svg>
                  {configError}
                </div>
              )}

              <div className="admin-toolbar">
                <p className="admin-hint">
                  Haz clic en <strong>Editar</strong> para modificar la configuración de un cliente. Los cambios se guardan automáticamente en la base de datos.
                </p>
                <div className="admin-toolbar-actions">
                  <button
                    className="btn-add-client"
                    onClick={() => { setShowAddForm(f => !f); setConfigError(null) }}
                  >
                    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                      <line x1="12" y1="5" x2="12" y2="19"/>
                      <line x1="5" y1="12" x2="19" y2="12"/>
                    </svg>
                    Agregar cliente
                  </button>
                  <button
                    className="btn-reset-defaults"
                    onClick={resetToDefaults}
                    disabled={resetting || configLoading}
                    title="Vuelve todos los valores a los originales del sistema"
                  >
                    {resetting ? (
                      <>
                        <svg className="spinner" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
                        </svg>
                        Restableciendo...
                      </>
                    ) : (
                      <>
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <polyline points="1 4 1 10 7 10"/>
                          <path d="M3.51 15a9 9 0 1 0 .49-4.95"/>
                        </svg>
                        Restablecer valores por defecto
                      </>
                    )}
                  </button>
                </div>
              </div>

              {showAddForm && (
              <div className="add-client-form">
                <h4 className="add-client-title">
                  <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                    <line x1="12" y1="5" x2="12" y2="19"/>
                    <line x1="5" y1="12" x2="19" y2="12"/>
                  </svg>
                  Nuevo cliente
                </h4>
                <div className="add-client-fields">
                  <div className="add-field add-field-code">
                    <label>
                      Código del cliente
                      <span className="add-field-hint">
                        {activeProcess === 'general'
                          ? '(debe coincidir exactamente con el valor en la columna CENTRO COSTOS de la base de empleados)'
                          : '(debe iniciar por CR y coincidir con el código usado en el reporte de Costa Rica)'}
                      </span>
                    </label>
                    <input
                      className="admin-input"
                      placeholder={activeProcess === 'general' ? 'Ej: C1061 - ACME CORP' : 'Ej: CR019 - CLIENTE NUEVO'}
                      value={addForm.code}
                      onChange={e => setAddForm(f => ({ ...f, code: e.target.value }))}
                      autoFocus
                    />
                  </div>
                  <div className="add-field add-field-fee">
                    <label>Fee</label>
                    <input
                      className="admin-input"
                      placeholder="Ej: 10% ó 150"
                      value={addForm.fee}
                      onChange={e => setAddForm(f => ({ ...f, fee: e.target.value }))}
                    />
                  </div>
                  <div className="add-field add-field-checks">
                    <label className="add-check-label">
                      <input type="checkbox" checked={addForm.iva} onChange={e => setAddForm(f => ({ ...f, iva: e.target.checked }))} />
                      IVA
                    </label>
                    <label className="add-check-label">
                      <input type="checkbox" checked={addForm.banking} onChange={e => setAddForm(f => ({ ...f, banking: e.target.checked }))} />
                      Banking Tax
                    </label>
                    <label className="add-check-label">
                      <input type="checkbox" checked={addForm.sin_usd} onChange={e => setAddForm(f => ({ ...f, sin_usd: e.target.checked }))} />
                      Sin USD
                    </label>
                    <label className="add-check-label">
                      <input type="checkbox" checked={addForm.es_liquidacion} onChange={e => setAddForm(f => ({ ...f, es_liquidacion: e.target.checked }))} />
                      Es Liquidación
                    </label>
                  </div>
                </div>
                <div className="add-client-actions">
                  <button
                    className="btn-save-row"
                    onClick={addCliente}
                    disabled={addingClient || !addForm.code.trim()}
                  >
                    {addingClient ? (
                      <>
                        <svg className="spinner" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
                        </svg>
                        Guardando...
                      </>
                    ) : 'Agregar cliente'}
                  </button>
                  <button className="btn-cancel-row" onClick={() => setShowAddForm(false)}>
                    Cancelar
                  </button>
                </div>
              </div>
              )}

              <div className="admin-table-wrap">
                <table className="admin-table">
                  <thead>
                    <tr>
                      <th>Cliente</th>
                      <th>Fee</th>
                      <th>IVA</th>
                      <th>Banking Tax</th>
                      <th>Sin USD</th>
                      <th>Es Liquidación</th>
                      <th>Acciones</th>
                    </tr>
                  </thead>
                  <tbody>
                    {Object.keys(clientesConfig).filter(code => code.toUpperCase().startsWith(currentProcess.clientPrefix.toUpperCase())).map(code => {
                      const cfg = clientesConfig[code] || DEFAULT_CONFIG_CLIENTES[code] || {}
                      const isEditing = editando === code
                      return (
                        <tr key={code} className={isEditing ? 'editing' : ''}>
                          <td className="col-code">{code}</td>

                          {isEditing ? (
                            <>
                              <td>
                                <input
                                  className="admin-input"
                                  value={editForm.fee}
                                  onChange={e => setEditForm(f => ({ ...f, fee: e.target.value }))}
                                  placeholder="ej: 10% ó 150"
                                />
                              </td>
                              <td className="col-bool">
                                <input type="checkbox" checked={editForm.iva} onChange={e => setEditForm(f => ({ ...f, iva: e.target.checked }))} />
                              </td>
                              <td className="col-bool">
                                <input type="checkbox" checked={editForm.banking} onChange={e => setEditForm(f => ({ ...f, banking: e.target.checked }))} />
                              </td>
                              <td className="col-bool">
                                <input type="checkbox" checked={editForm.sin_usd} onChange={e => setEditForm(f => ({ ...f, sin_usd: e.target.checked }))} />
                              </td>
                              <td className="col-bool">
                                <input type="checkbox" checked={editForm.es_liquidacion} onChange={e => setEditForm(f => ({ ...f, es_liquidacion: e.target.checked }))} />
                              </td>
                              <td className="col-actions">
                                <button
                                  className="btn-save-row"
                                  onClick={() => saveClienteConfig(code, editForm)}
                                  disabled={savingConfig}
                                >
                                  {savingConfig ? '...' : 'Guardar'}
                                </button>
                                <button className="btn-cancel-row" onClick={() => setEditando(null)}>
                                  Cancelar
                                </button>
                              </td>
                            </>
                          ) : (
                            <>
                              <td><span className="badge-fee">{cfg.fee}</span></td>
                              <td className="col-bool">{cfg.iva     ? <span className="dot dot-on"/> : <span className="dot dot-off"/>}</td>
                              <td className="col-bool">{cfg.banking ? <span className="dot dot-on"/> : <span className="dot dot-off"/>}</td>
                              <td className="col-bool">{cfg.sin_usd ? <span className="dot dot-on"/> : <span className="dot dot-off"/>}</td>
                              <td className="col-bool">{cfg.es_liquidacion ? <span className="dot dot-on"/> : <span className="dot dot-off"/>}</td>
                              <td className="col-actions">
                                <button
                                  className="btn-edit-row"
                                  onClick={() => {
                                    setEditando(code)
                                    setEditForm({ fee: cfg.fee, iva: !!cfg.iva, banking: !!cfg.banking, sin_usd: !!cfg.sin_usd, es_liquidacion: !!cfg.es_liquidacion })
                                  }}
                                >
                                  Editar
                                </button>
                              </td>
                            </>
                          )}
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>

            </div>
          </div>

          {activeProcess === 'costa-rica' && (
            <div className="admin-card remofirst-subclientes-card">
              <div className="admin-body">
                <div className="remofirst-subclientes-header">
                  <h3>Configuración subclientes - Remofirst</h3>
                  <button
                    className="btn-add-client"
                    onClick={() => { setShowRemoSubForm(v => !v); setSubclientesError(null) }}
                  >
                    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                      <line x1="12" y1="5" x2="12" y2="19"/>
                      <line x1="5" y1="12" x2="19" y2="12"/>
                    </svg>
                    Agregar subcliente
                  </button>
                </div>

                <div className="remofirst-subclientes-help">
                  Aquí debes registrar los subclientes exactamente como aparecen en el archivo de Novasoft. Cuando se seleccione <strong>CR009 - REMOFIRST INC</strong>, la facturación incluirá filas con <strong>CCOSTO = CR009</strong> y también filas que coincidan con estos pares <strong>CCOSTO + DESC.CCOSTO</strong>. Si diligencias <strong>Nombre completo</strong>, ese texto reemplaza el valor que se pone en <strong>Customer Name</strong>.
                </div>

                {subclientesError && (
                  <div className="alert alert-error" style={{marginBottom:'1rem'}}>
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
                    </svg>
                    {subclientesError}
                  </div>
                )}

                {showRemoSubForm && (
                  <div className="add-client-form remofirst-subform">
                    <h4 className="add-client-title">Nuevo subcliente Remofirst</h4>
                    <div className="add-client-fields remofirst-subfields">
                      <div className="add-field remofirst-subfield-ccosto">
                        <label>CCOSTO</label>
                        <input
                          className="admin-input"
                          placeholder="Ej: CR004"
                          value={newRemoSubcliente.ccosto}
                          onChange={e => setNewRemoSubcliente(s => ({ ...s, ccosto: e.target.value }))}
                        />
                      </div>
                      <div className="add-field remofirst-subfield-desc">
                        <label>DESC.CCOSTO</label>
                        <input
                          className="admin-input"
                          placeholder="Ej: ANGL RF"
                          value={newRemoSubcliente.descCcosto}
                          onChange={e => setNewRemoSubcliente(s => ({ ...s, descCcosto: e.target.value }))}
                        />
                      </div>
                      <div className="add-field remofirst-subfield-name">
                        <label>Nombre completo</label>
                        <input
                          className="admin-input"
                          placeholder="Opcional: nombre para Customer Name"
                          value={newRemoSubcliente.nombreCompleto}
                          onChange={e => setNewRemoSubcliente(s => ({ ...s, nombreCompleto: e.target.value }))}
                        />
                      </div>
                    </div>
                    <div className="add-client-actions">
                      <button
                        className="btn-save-row"
                        onClick={addRemoSubcliente}
                        disabled={addingRemoSubcliente || !newRemoSubcliente.ccosto.trim() || !newRemoSubcliente.descCcosto.trim()}
                      >
                        {addingRemoSubcliente ? 'Guardando...' : 'Agregar subcliente'}
                      </button>
                      <button className="btn-cancel-row" onClick={() => setShowRemoSubForm(false)}>
                        Cancelar
                      </button>
                    </div>
                  </div>
                )}

                <div className="admin-table-wrap">
                  <table className="admin-table remofirst-subclientes-table">
                    <thead>
                      <tr>
                        <th>CCOSTO</th>
                        <th>DESC.CCOSTO</th>
                        <th>Nombre completo</th>
                        <th>Acciones</th>
                      </tr>
                    </thead>
                    <tbody>
                      {subclientesLoading ? (
                        <tr>
                          <td colSpan="4">Cargando subclientes...</td>
                        </tr>
                      ) : remoSubclientes.length === 0 ? (
                        <tr>
                          <td colSpan="4">No hay subclientes configurados para Remofirst.</td>
                        </tr>
                      ) : (
                        remoSubclientes.map(item => {
                          const rowKey = `${item.ccosto}|${item.descCcosto}`
                          const isEditing = editingRemoSubcliente === rowKey
                          return (
                            <tr key={rowKey}>
                              {isEditing ? (
                                <>
                                  <td>
                                    <input
                                      className="admin-input"
                                      value={editRemoSubclienteForm.ccosto}
                                      onChange={e => setEditRemoSubclienteForm(s => ({ ...s, ccosto: e.target.value }))}
                                    />
                                  </td>
                                  <td>
                                    <input
                                      className="admin-input"
                                      value={editRemoSubclienteForm.descCcosto}
                                      onChange={e => setEditRemoSubclienteForm(s => ({ ...s, descCcosto: e.target.value }))}
                                    />
                                  </td>
                                  <td>
                                    <input
                                      className="admin-input"
                                      value={editRemoSubclienteForm.nombreCompleto}
                                      onChange={e => setEditRemoSubclienteForm(s => ({ ...s, nombreCompleto: e.target.value }))}
                                    />
                                  </td>
                                </>
                              ) : (
                                <>
                                  <td className="col-code">{item.ccosto}</td>
                                  <td>{item.descCcosto}</td>
                                  <td>{item.nombreCompleto || <span className="subcliente-empty">(vacío)</span>}</td>
                                </>
                              )}
                              <td className="col-actions">
                                {isEditing ? (
                                  <>
                                    <button
                                      className="btn-save-row"
                                      onClick={() => saveRemoSubcliente(item.ccosto, item.descCcosto)}
                                      disabled={savingRemoSubcliente}
                                    >
                                      {savingRemoSubcliente ? 'Guardando...' : 'Guardar'}
                                    </button>
                                    <button className="btn-cancel-row" onClick={cancelEditRemoSubcliente}>
                                      Cancelar
                                    </button>
                                  </>
                                ) : (
                                  <>
                                    <button
                                      className="btn-edit-row"
                                      onClick={() => startEditRemoSubcliente(item)}
                                    >
                                      Editar
                                    </button>
                                    <button
                                      className="btn-cancel-row"
                                      onClick={() => removeRemoSubcliente(item.ccosto, item.descCcosto)}
                                      disabled={deletingRemoSubcliente === `${item.ccosto} | ${item.descCcosto}`}
                                    >
                                      {deletingRemoSubcliente === `${item.ccosto} | ${item.descCcosto}` ? 'Eliminando...' : 'Eliminar'}
                                    </button>
                                  </>
                                )}
                              </td>
                            </tr>
                          )
                        })
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {activeProcess === 'costa-rica' && (
            <div className="admin-card remofirst-subclientes-card">
              <div className="admin-body">
                <div className="remofirst-subclientes-header">
                  <h3>Configuracion empleados</h3>
                  <button
                    className="btn-add-client"
                    onClick={() => { setShowEmpleadoForm(v => !v); setEmpleadosConfigError(null) }}
                  >
                    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                      <line x1="12" y1="5" x2="12" y2="19"/>
                      <line x1="5" y1="12" x2="19" y2="12"/>
                    </svg>
                    Agregar empleado
                  </button>
                </div>

                <div className="remofirst-subclientes-help">
                  Este listado controla excepciones de tasa para Costa Rica. Si el documento (cédula) del reporte Novasoft coincide con un empleado de esta tabla, en el Excel final la columna <strong>EXCHANGE RATE</strong> se fija en <strong>1</strong> para ese empleado, sin importar la tasa ingresada en la web.
                </div>

                {empleadosConfigError && (
                  <div className="alert alert-error" style={{marginBottom:'1rem'}}>
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
                    </svg>
                    {empleadosConfigError}
                  </div>
                )}

                {showEmpleadoForm && (
                  <div className="add-client-form remofirst-subform">
                    <h4 className="add-client-title">Nuevo empleado</h4>
                    <div className="add-client-fields remofirst-subfields">
                      <div className="add-field remofirst-subfield-ccosto">
                        <label>Documento</label>
                        <input
                          className="admin-input"
                          placeholder="Ej: 112280205"
                          value={newEmpleado.documento}
                          onChange={e => setNewEmpleado(s => ({ ...s, documento: e.target.value }))}
                        />
                      </div>
                      <div className="add-field remofirst-subfield-desc">
                        <label>Nombre</label>
                        <input
                          className="admin-input"
                          placeholder="Ej: TATIANA ARAYA POCHET"
                          value={newEmpleado.nombre}
                          onChange={e => setNewEmpleado(s => ({ ...s, nombre: e.target.value }))}
                        />
                      </div>
                    </div>
                    <div className="add-client-actions">
                      <button
                        className="btn-save-row"
                        onClick={addEmpleadoConfig}
                        disabled={addingEmpleado || !newEmpleado.documento.trim() || !newEmpleado.nombre.trim()}
                      >
                        {addingEmpleado ? 'Guardando...' : 'Agregar empleado'}
                      </button>
                      <button className="btn-cancel-row" onClick={() => setShowEmpleadoForm(false)}>
                        Cancelar
                      </button>
                    </div>
                  </div>
                )}

                <div className="admin-table-wrap">
                  <table className="admin-table remofirst-subclientes-table">
                    <thead>
                      <tr>
                        <th>Documento</th>
                        <th>Nombre</th>
                        <th>Acciones</th>
                      </tr>
                    </thead>
                    <tbody>
                      {empleadosConfigLoading ? (
                        <tr><td colSpan="3">Cargando empleados...</td></tr>
                      ) : empleadosConfig.length === 0 ? (
                        <tr><td colSpan="3">No hay empleados configurados.</td></tr>
                      ) : (
                        empleadosConfig.map(emp => {
                          const isEditing = editingEmpleado === emp.documento
                          return (
                            <tr key={emp.documento}>
                              {isEditing ? (
                                <>
                                  <td>
                                    <input
                                      className="admin-input"
                                      value={editEmpleadoForm.documento}
                                      onChange={e => setEditEmpleadoForm(s => ({ ...s, documento: e.target.value }))}
                                    />
                                  </td>
                                  <td>
                                    <input
                                      className="admin-input"
                                      value={editEmpleadoForm.nombre}
                                      onChange={e => setEditEmpleadoForm(s => ({ ...s, nombre: e.target.value }))}
                                    />
                                  </td>
                                </>
                              ) : (
                                <>
                                  <td className="col-code">{emp.documento}</td>
                                  <td>{emp.nombre}</td>
                                </>
                              )}
                              <td className="col-actions">
                                {isEditing ? (
                                  <>
                                    <button
                                      className="btn-save-row"
                                      onClick={() => saveEmpleadoConfig(emp.documento)}
                                      disabled={savingEmpleado}
                                    >
                                      {savingEmpleado ? 'Guardando...' : 'Guardar'}
                                    </button>
                                    <button className="btn-cancel-row" onClick={cancelEditEmpleadoConfig}>
                                      Cancelar
                                    </button>
                                  </>
                                ) : (
                                  <>
                                    <button className="btn-edit-row" onClick={() => startEditEmpleadoConfig(emp)}>
                                      Editar
                                    </button>
                                    <button
                                      className="btn-cancel-row"
                                      onClick={() => removeEmpleadoConfig(emp.documento)}
                                      disabled={deletingEmpleado === emp.documento}
                                    >
                                      {deletingEmpleado === emp.documento ? 'Eliminando...' : 'Eliminar'}
                                    </button>
                                  </>
                                )}
                              </td>
                            </tr>
                          )
                        })
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

        </div>
      </section>
      )}

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
