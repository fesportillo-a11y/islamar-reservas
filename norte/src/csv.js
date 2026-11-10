// Lectura de CSV de extractos bancarios. Detecta separador, cabecera y columnas.
import { CATS, guessCat } from './categorias.js'

export function parseCsv(text) {
  const lines = text.split(/\r?\n/).filter(l => l.trim())
  if (!lines.length) return []
  const sep = (lines[0].match(/;/g) || []).length > (lines[0].match(/,/g) || []).length ? ';' : ','
  const head = splitLine(lines[0], sep).map(h => h.trim().toLowerCase())
  const hasHeader = head.some(h => /fecha|concepto|import|descr|categor/.test(h))
  const idx = {
    fecha: findCol(head, ['fecha', 'date', 'día', 'dia']),
    concepto: findCol(head, ['concepto', 'descripcion', 'descripción', 'description', 'detalle']),
    importe: findCol(head, ['importe', 'amount', 'cantidad', 'euros', 'import']),
    categoria: findCol(head, ['categoria', 'categoría', 'category', 'tipo']),
  }
  const rows = (hasHeader ? lines.slice(1) : lines).map(l => splitLine(l, sep))
  return rows.map(r => {
    const get = (k, fb) => (idx[k] >= 0 ? r[idx[k]] : r[fb])
    const fecha = normDate(get('fecha', 0))
    const concepto = (get('concepto', 1) || '').trim()
    const importe = normNum(get('importe', 2))
    let categoria = (get('categoria', 3) || '').trim()
    if (!categoria || !CATS[categoria]) categoria = guessCat(concepto, importe)
    return { fecha, concepto, importe, categoria }
  }).filter(r => r.fecha && r.concepto && !isNaN(r.importe))
}

function splitLine(l, sep) {
  const out = []
  let cur = '', q = false
  for (let i = 0; i < l.length; i++) {
    const ch = l[i]
    if (ch === '"') { q = !q; continue }
    if (ch === sep && !q) { out.push(cur); cur = ''; continue }
    cur += ch
  }
  out.push(cur)
  return out
}
const findCol = (head, names) => { for (const n of names) { const i = head.indexOf(n); if (i >= 0) return i } return -1 }

function normDate(s) {
  if (!s) return ''
  s = s.trim()
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/)
  if (m) return `${m[1]}-${m[2]}-${m[3]}`
  m = s.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})/)
  if (m) { let [, d, mo, y] = m; if (y.length === 2) y = '20' + y; return `${y}-${mo.padStart(2, '0')}-${d.padStart(2, '0')}` }
  return ''
}
function normNum(s) {
  if (s == null) return NaN
  s = ('' + s).replace(/[€\s]/g, '')
  if (s.includes(',') && s.includes('.')) s = s.replace(/\./g, '').replace(',', '.')
  else if (s.includes(',')) s = s.replace(',', '.')
  return parseFloat(s)
}
