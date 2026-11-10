// Pantalla Movimientos: lista categorizada con buscador y filtro por categoría.
import { state } from '../store.js'
import { CATS, catOf, catSVG } from '../categorias.js'
import { eur, fmtDay, esc } from '../format.js'

const $ = s => document.querySelector(s)
let filterCat = 'Todas'
let query = ''

export function init() {
  $('#q').addEventListener('input', e => { query = e.target.value; render() })
  // El filtro por categoría se delega desde main.js (data-cat).
}

export function setFilter(cat) { filterCat = cat; render() }
export function setQuery(q) { query = q; render() }

export function render() {
  const cats = ['Todas', ...Object.keys(CATS).filter(c => state.movimientos.some(m => m.categoria === c))]
  $('#catFilter').innerHTML = cats.map(c => `<button class="pill ${c === filterCat ? 'on' : ''}" data-cat="${c}">${c}</button>`).join('')

  let list = [...state.movimientos].sort((a, b) => (b.fecha < a.fecha ? -1 : 1))
  if (filterCat !== 'Todas') list = list.filter(m => m.categoria === filterCat)
  if (query) list = list.filter(m => m.concepto.toLowerCase().includes(query.toLowerCase()))

  const box = $('#movList')
  if (!list.length) { box.innerHTML = '<div class="empty">No hay movimientos que coincidan.</div>'; return }

  let html = '', lastDay = ''
  for (const m of list) {
    if (m.fecha !== lastDay) { lastDay = m.fecha; html += `<div class="daterow">${fmtDay(m.fecha)}</div>` }
    const c = catOf(m.categoria), pos = m.importe > 0
    html += `<div class="row"><div class="ic" style="background:${c.col}1f">${catSVG(m.categoria)}</div>
      <div class="info"><div class="n">${esc(m.concepto)}</div>
      <div class="s"><span class="chip">${m.categoria}</span><span>${cuentaNom(m.cuenta)}</span></div></div>
      <div class="amt ${pos ? 'up' : ''}">${pos ? '+' : ''}${eur(m.importe)}</div></div>`
  }
  box.innerHTML = html
}

const cuentaNom = id => { const c = state.cuentas.find(x => x.id === id); return c ? c.nombre : '—' }
