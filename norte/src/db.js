// Base de datos local con Dexie (IndexedDB). Todo vive en el navegador del usuario.
import Dexie from 'dexie'

export const db = new Dexie('norte')
db.version(1).stores({
  cuentas: 'id',
  movimientos: 'id, fecha, categoria, cuenta',
})

// Si no hay nada guardado, siembra los datos de ejemplo.
export async function ensureSeed(seed) {
  const n = await db.cuentas.count()
  if (n === 0) {
    await db.transaction('rw', db.cuentas, db.movimientos, async () => {
      await db.cuentas.bulkAdd(seed.cuentas)
      await db.movimientos.bulkAdd(seed.movimientos)
    })
  }
}
