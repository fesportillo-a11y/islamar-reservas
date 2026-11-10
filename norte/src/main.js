// Arranque de Norte: abre la base de datos, siembra si hace falta, monta las vistas y el router.
import { ensureSeed } from './db.js'
import { reload } from './store.js'
import { seedData } from './seed.js'
import { setRenderAll } from './bus.js'

import * as resumen from './views/resumen.js'
import * as movimientos from './views/movimientos.js'
import * as analisis from './views/analisis.js'
import * as cuentas from './views/cuentas.js'
import * as importar from './views/importar.js'

const $ = s => document.querySelector(s)
const $$ = s => [...document.querySelectorAll(s)]

const VIEWS = [
  ['resumen', 'Resumen', '<path d="M3 10 12 3l9 7"/><path d="M5 9v11h14V9"/>'],
  ['movimientos', 'Movimientos', '<path d="M4 7h16M4 12h16M4 17h10"/>'],
  ['analisis', 'Análisis', '<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>'],
  ['cuentas', 'Cuentas', '<path d="M3 7h18v12H3z"/><path d="M3 11h18"/><path d="M7 15h3"/>'],
  ['importar', 'Importar', '<path d="M12 15V3M8 7l4-4 4 4"/><path d="M4 15v4a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-4"/>'],
]
const MODULES = { resumen, movimientos, analisis, cuentas, importar }

function renderAll() { for (const k in MODULES) MODULES[k].render && MODULES[k].render() }

function goto(v) {
  VIEWS.forEach(([id]) => { $('#v-' + id).hidden = id !== v })
  $$('.tab').forEach(t => t.classList.toggle('on', t.dataset.v === v))
  window.scrollTo(0, 0)
  if (v === 'analisis') analisis.render() // el canvas necesita recalcularse al mostrarse
}

function buildNav() {
  $('#nav').innerHTML = VIEWS.map(([id, label, p]) =>
    `<button class="tab" data-v="${id}"><svg class="i" viewBox="0 0 24 24" stroke="currentColor">${p}</svg>${label}</button>`
  ).join('')
}

function wireEvents() {
  document.addEventListener('click', e => {
    const t = e.target.closest('[data-v]'); if (t) { goto(t.dataset.v); return }
    const g = e.target.closest('[data-go]'); if (g) { goto(g.dataset.go); return }
    const c = e.target.closest('[data-cat]'); if (c) { movimientos.setFilter(c.dataset.cat); return }
  })

  const root = document.documentElement
  $('#themeBtn').addEventListener('click', () => {
    const dark = matchMedia('(prefers-color-scheme:dark)').matches
    const curr = root.getAttribute('data-theme') || (dark ? 'dark' : 'light')
    root.setAttribute('data-theme', curr === 'dark' ? 'light' : 'dark')
    try { localStorage.setItem('norte.theme', root.getAttribute('data-theme')) } catch (e) {}
    analisis.render()
  })
  try { const th = localStorage.getItem('norte.theme'); if (th) root.setAttribute('data-theme', th) } catch (e) {}
}

async function boot() {
  setRenderAll(renderAll)
  buildNav()
  wireEvents()
  movimientos.init()
  importar.init(goto)

  await ensureSeed(seedData())
  await reload()
  renderAll()
  goto('resumen')
}

boot()
