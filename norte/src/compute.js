// Cálculos derivados del estado: patrimonio, datos por mes y Pulso (salud financiera).
import { state } from './store.js'
import { ym } from './format.js'

export const monthsSorted = () => [...new Set(state.movimientos.map(m => ym(m.fecha)))].sort()

export function monthData(key) {
  const ms = state.movimientos.filter(m => ym(m.fecha) === key)
  const ingresos = ms.filter(m => m.importe > 0).reduce((a, b) => a + b.importe, 0)
  const gastos = ms.filter(m => m.importe < 0).reduce((a, b) => a + Math.abs(b.importe), 0)
  const porCat = {}
  ms.filter(m => m.importe < 0).forEach(m => { porCat[m.categoria] = (porCat[m.categoria] || 0) + Math.abs(m.importe) })
  return { ingresos, gastos, porCat, ms }
}

export const patrimonio = () => state.cuentas.reduce((a, b) => a + b.saldo, 0)

// Pulso 0–100: parte de la tasa de ahorro del último mes cerrado más un pequeño peso de la liquidez.
export function pulso() {
  const months = monthsSorted()
  const cerrado = months.filter(k => k < '2026-11').slice(-1)[0] || months.slice(-1)[0]
  const d = cerrado ? monthData(cerrado) : { ingresos: 0, gastos: 0 }
  const tasa = d.ingresos > 0 ? (d.ingresos - d.gastos) / d.ingresos : 0
  let s = 50 + Math.max(-25, Math.min(30, tasa * 60)) + Math.min(20, patrimonio() / 1000)
  return Math.round(Math.max(12, Math.min(96, s)))
}
