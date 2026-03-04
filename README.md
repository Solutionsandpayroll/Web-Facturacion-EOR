# Web Facturación EOR — Solutions & Payroll

Aplicación web interna desarrollada en **React + Vite** para la generación automática de archivos Excel de facturación por cliente, a partir de los reportes exportados desde **Novasoft** y la base de empleados.

---

## ¿Qué hace esta aplicación?

1. El usuario carga dos archivos Excel:
   - **Base de empleados**: contiene la relación entre código de empleado y centro de costos (cliente).
   - **Reporte Novasoft** (`nomplcol`): contiene los conceptos de nómina liquidados por empleado en el periodo.
2. El usuario selecciona uno o varios clientes y escribe el periodo de facturación.
3. La app genera **un archivo Excel por cliente**, basado en la plantilla `Formato facturacion - final.xlsx`, con:
   - Los códigos de empleado filtrados por cliente.
   - Los valores de cada concepto sumados y asignados a la columna correspondiente.
   - Fórmulas de TOTAL y PAYMENTS replicadas dinámicamente fila por fila.
   - Fórmulas de FEE, BANKING TAX e IVA calculadas según la tarifa configurada para ese cliente.
   - Fila de totales (fila de fórmulas de la plantilla) reposicionada al final de los datos con rangos actualizados dinámicamente.

---

## Estructura del proyecto

```
Web facturación/
├── public/
│   ├── Formato facturacion - final.xlsx   # Plantilla Excel de salida
│   └── Logo syp.png                        # Logo corporativo
├── src/
│   ├── App.jsx                             # Componente principal + lógica de procesamiento
│   └── App.css                             # Estilos corporativos S&P
├── index.html
├── package.json
└── vite.config.js
```

---

## Instalación y arranque

### Requisitos
- Node.js 18+
- npm

### Pasos

```bash
# Instalar dependencias
npm install

# Iniciar servidor de desarrollo
npm run dev
```

> **Nota:** El script `dev` usa `node node_modules/vite/bin/vite.js` en lugar del alias `vite` estándar, para evitar un problema de la terminal de Windows con el carácter `&` en la ruta de la carpeta `S&P`.

---

## Dependencias principales

| Paquete | Uso |
|---|---|
| `react` + `react-dom` | Interfaz de usuario |
| `vite` | Bundler y servidor de desarrollo |
| `xlsx` (SheetJS) | Lectura de archivos Excel de entrada |
| `exceljs` | Escritura del Excel de salida preservando estilos |
| `jszip` | Pre-procesamiento del XML del template para corregir shared formulas |

---

## Configuración de clientes (`CONFIG_CLIENTES`)

Ubicada al inicio de `src/App.jsx`, esta constante define las reglas de facturación por cliente:

```js
const CONFIG_CLIENTES = {
  'C1037 - REMOFIRST': { fee: '120', iva: true, banking: true },
  'C1007 - NZD':       { fee: '6%',  iva: false, banking: false },
  // ...
}
```

| Campo | Tipo | Descripción |
|---|---|---|
| `fee` | `string` | Valor del coeficiente para la fórmula FEE. Si es porcentaje (ej. `'6%'`), se multiplica directamente por `$BG{fila}`. Si es monto fijo (ej. `'120'`), ídem. |
| `iva` | `boolean` | `true` → genera fórmula `=$BC{fila}*0.19`. `false` → celda vacía. |
| `banking` | `boolean` | `true` → genera fórmula `=SUM($AF,$AN,$AS,$AZ)*0.004`. `false` → celda vacía. |

### Tarifa FEE completa por cliente

| Cliente | FEE | IVA | 4x1000 |
|---|---|---|---|
| C1007 - NZD | 6% | Incluido | No |
| C1024 - MKD | 11% | Incluido | No |
| C1032 - FLEXCO | 9% | Incluido | No |
| C1038 - ONCEHUB | USD 100,84 | 19% | Sí |
| C1041 - EDRINGTON | 5,5% | Incluido | No |
| C1042 - EPDM | 10% | Incluido | No |
| C1043 - NEO | 8% | Incluido | No |
| C1050 - HEMMERSBACH | 10% | Incluido | No |
| C1051 - YONYOU | USD 210,00 | 19% | Sí |
| C1052 - BUBBLE BPM INC | 11% | Incluido | No |
| C1053 - GLOBAL EXPANSION | USD 150,00 | 19% | Sí |
| C1055 - RIVERMATE | EUR 150,00 | Incluido | Sí |
| C1037 - REMOFIRST | USD 120,00 | 19% | Sí |
| C1029 - INSIDER | USD 190,00 | 19% | Sí |
| C1036 - ACTION AD | 11% | Incluido | No |
| C1056 - EUROPORTAGE | USD 200,00 | 19% | Sí |
| C1022 - Root Capital | 9% | Incluido | No |
| C1058 - POC PHARMA | USD 160,00 | Incluido | Sí |
| C1059 - SIFFI | 10% | Incluido | Sí |

---

## Mapeo de conceptos Novasoft → columnas de salida (`GRUPOS`)

Definido en `src/App.jsx` dentro de la función `generar()`. Cada grupo especifica:
- `columnaDestino`: nombre exacto del encabezado en la plantilla Excel.
- `conceptos`: lista de conceptos de la fila 23 del Novasoft cuya suma va a esa columna.

La comparación es **case-insensitive** y con `.trim()`.

Grupos configurados:

| Columna destino | Descripción |
|---|---|
| `SALARY` | Salario base, vacaciones, honorarios, retroactivos, etc. |
| `"Sick" Leave` | Incapacidades por enfermedad común y accidente |
| `13th Salary` | Provisión prima |
| `14th Salary` | Provisión cesantías |
| `Alloawance 1 (Car Allowance)` | Auxilios de transporte extralegal y movilización |
| `Alloawance 2 (Mobile & Internet Allowance)` | Auxilios de celular, conectividad, telecomunicaciones |
| `Alloawance 4 (Other allowances)` | Gross Up, vivienda, AFC, viáticos, educación |
| `Bonus/Commission` | Comisiones, bonificaciones, primas extralegales |
| `Deduction or Gross Amount adjustments prevous month` | Descuentos, anticipos, préstamos |
| `Expenses` | Reembolsos y gastos tarjeta crédito |
| `Family Fund Cost` | Aporte caja de compensación |
| `Food allowance` | Auxilio de alimentación |
| `Health Cost` | Salud patrono |
| `Health Insurance` | Póliza médica |
| `Home/Remote work allowance` | Auxilio de computador |
| `ICBF cost` | Aportes ICBF |
| `Interest on 14th Salary` | Provisión intereses cesantías |
| `Labor Risk Cost` | Riesgo profesional (ARL) |
| `Medical Test` | Gastos examen médico |
| `On Call/ Plus Disponibilidad` | Standby salarial |
| `Overtime` | Horas extras diurnas, nocturnas, dominicales |
| `Paternity/ Maternity leave` | Licencias de paternidad y maternidad |
| `Pension Cost` | Pensión patrono |
| `Rectroactive payment/Plus Compensation` | Retro apoyo sostenimiento |
| `SENA Cost` | Apropiación SENA |
| `Severance Pay (Taxable)` | Indemnización, sumas transaccionales, bonos de retiro |
| `Sign-on Bonus` | Bonificación de firma |
| `Transport allowance` | Subsidio de transporte y retroactivos |
| `Unused Holidays` | Vacaciones en dinero y liquidación |
| `Wellness Allowance` | Bienestar, salud, gym, medicina prepagada, póliza de vida |

---

## Estructura del Novasoft esperado

| Fila | Contenido |
|---|---|
| 22 | Encabezados principales (`Codigo Empl`, `Nombre`, secciones DEVENGO/DEDUCCION…) |
| 23 | Sub-encabezados de conceptos (texto exacto que se usa para el mapeo) |
| 24 | Fila vacía / separador |
| 25+ | Datos por empleado |

La hoja debe llamarse **`nomplcol`**.

---

## Plantilla Excel de salida

Archivo: `public/Formato facturacion - final.xlsx`

| Fila | Contenido |
|---|---|
| 1-2 | Encabezado corporativo (logos, título) |
| 3 | Encabezados de columnas (`EMPLOYEE CODE`, `SALARY`, `FEE`, `IVA`, etc.) |
| 4 | Primera fila de datos (estilos de referencia + fórmulas TOTAL/PAYMENTS replicadas) |
| 5 | Fila de totales con fórmulas `=SUM(X4:X4)` → se mueve dinámicamente al final |

---

## Nombre del archivo generado

```
Facturación EOR - {Cliente} - {Periodo}.xlsx
```

Ejemplo: `Facturación EOR - C1037 - REMOFIRST - Enero 2026.xlsx`

---

## Notas técnicas

- El template usa **shared formulas** en Excel que ExcelJS no soporta directamente. Se resuelve con JSZip pre-procesando el XML antes de cargarlo.
- Las fórmulas de la fila de totales se recuperan con SheetJS (antes del pre-procesamiento) para garantizar que todas las columnas tengan su fórmula completa.
- ExcelJS requiere fórmulas en **formato inglés** (`SUM`, `,` como separador, decimales con `.`). Excel las convierte al locale del usuario al abrir.
