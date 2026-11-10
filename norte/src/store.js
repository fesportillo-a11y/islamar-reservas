// Estado en memoria + sincronización con Dexie. Las vistas leen de `state`.
import { db } from './db.js'

export const state = { cuentas: [], movimientos: [] }

export async function reload() {
  state.cuentas = await db.cuentas.toArray()
  state.movimientos = (await db.movimientos.toArray()).sort((a, b) => (a.fecha < b.fecha ? -1 : 1))
}

export async function addMovimientos(rows) {
  let base = Date.now()
  const nuevos = rows.map(r => ({ id: 'i' + (base++), cuenta: 'c1', ...r }))
  await db.movimientos.bulkAdd(nuevos)
  await reload()
  return nuevos.length
}

export async function loadDemo(seed) {
  await db.transaction('rw', db.cuentas, db.movimientos, async () => {
    await db.movimientos.clear()
    await db.cuentas.clear()
    await db.cuentas.bulkAdd(seed.cuentas)
    await db.movimientos.bulkAdd(seed.movimientos)
  })
  await reload()
}

export async function resetAll() {
  await db.transaction('rw', db.cuentas, db.movimientos, async () => {
    await db.movimientos.clear()
    for (const c of state.cuentas) await db.cuentas.put({ ...c, saldo: 0 })
  })
  await reload()
}
