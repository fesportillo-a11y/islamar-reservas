// Datos de ejemplo: 4 meses de movimientos y 3 cuentas. Nada de marcas reales.
import { guessCat } from './categorias.js'

export function seedData() {
  const cuentas = [
    { id: 'c1', nombre: 'Cuenta Nómina', banco: 'BBVA', tipo: 'corriente', saldo: 3184.55 },
    { id: 'c2', nombre: 'Cuenta Ahorro', banco: 'ING', tipo: 'ahorro', saldo: 12650.00 },
    { id: 'c3', nombre: 'Tarjeta Crédito', banco: 'CaixaBank', tipo: 'tarjeta', saldo: -842.30 },
  ]
  const movs = []
  let id = 1
  const mk = (fecha, concepto, imp, cuenta = 'c1') =>
    movs.push({ id: 'm' + (id++), fecha, concepto, importe: +imp.toFixed(2), categoria: guessCat(concepto, imp), cuenta })

  const month = (year, m, extra = 0) => {
    const d = n => `${year}-${String(m + 1).padStart(2, '0')}-${String(n).padStart(2, '0')}`
    mk(d(1), 'Alquiler piso', -850, 'c1'); mk(d(2), 'Suscripción vídeo', -13.99, 'c3'); mk(d(3), 'Endesa Luz', -(58 + extra), 'c1')
    mk(d(4), 'Mercadona', -62.40, 'c1'); mk(d(5), 'Nómina', 1850, 'c1'); mk(d(6), 'Spotify', -10.99, 'c3')
    mk(d(7), 'Repsol Gasolinera', -55.10, 'c3'); mk(d(9), 'Restaurante La Brasa', -(34 + extra * 0.6), 'c3'); mk(d(11), 'Amazon', -27.80, 'c3')
    mk(d(12), 'Emasesa Agua', -31.20, 'c1'); mk(d(14), 'Mercadona', -71.15, 'c1'); mk(d(16), 'Farmacia', -18.40, 'c1')
    mk(d(18), 'Glovo', -23.50, 'c3'); mk(d(20), 'Decathlon', -45.90, 'c3'); mk(d(22), 'Movistar Fibra+Móvil', -49.90, 'c1')
    mk(d(24), 'Carrefour', -58.30, 'c1'); mk(d(26), 'Cine Nervión', -16.00, 'c3'); mk(d(28), 'Transferencia a Ahorro', -300, 'c1')
  }
  month(2026, 7, 6); month(2026, 8, 2); month(2026, 9, 0)

  // Noviembre parcial (hasta el día 9).
  const d = n => `2026-11-${String(n).padStart(2, '0')}`
  mk(d(1), 'Alquiler piso', -850, 'c1'); mk(d(2), 'Suscripción vídeo', -13.99, 'c3'); mk(d(3), 'Endesa Luz', -71.30, 'c1')
  mk(d(4), 'Mercadona', -54.20, 'c1'); mk(d(5), 'Nómina', 1850, 'c1'); mk(d(6), 'Spotify', -10.99, 'c3')
  mk(d(7), 'Repsol Gasolinera', -60.00, 'c3'); mk(d(8), 'Restaurante Casa Paco', -78.50, 'c3'); mk(d(9), 'Glovo', -19.90, 'c3'); mk(d(9), 'Amazon', -34.95, 'c3')

  return { cuentas, movimientos: movs }
}
