// Pantalla Importar: carga CSV, datos de ejemplo y borrado. El núcleo de la app.
import { parseCsv } from '../csv.js'
import { addMovimientos, loadDemo, resetAll } from '../store.js'
import { seedData } from '../seed.js'
import * as bus from '../bus.js'

const $ = s => document.querySelector(s)

export function init(goto) {
  const drop = $('#drop'), fileIn = $('#file')

  drop.addEventListener('click', e => { if (!e.target.closest('#loadDemo')) fileIn.click() })
  fileIn.addEventListener('change', e => { const f = e.target.files[0]; if (f) readCsv(f, goto); fileIn.value = '' })
  ;['dragover', 'dragenter'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('hot') }))
  ;['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('hot') }))
  drop.addEventListener('drop', e => { const f = e.dataTransfer.files[0]; if (f) readCsv(f, goto) })

  $('#loadDemo').addEventListener('click', async e => {
    e.stopPropagation()
    await loadDemo(seedData())
    bus.renderAll(); goto('resumen'); bus.toast('Datos de ejemplo cargados')
  })

  const reset = $('#reset')
  reset.addEventListener('click', async () => {
    if (reset.dataset.c) {
      await resetAll()
      bus.renderAll(); bus.toast('Datos borrados')
      reset.textContent = 'Borrar todos mis datos'; delete reset.dataset.c
    } else {
      reset.dataset.c = '1'; reset.textContent = '¿Seguro? Pulsa otra vez para borrar'
    }
  })
}

export function render() { /* estática: nada que recalcular */ }

function readCsv(file, goto) {
  const rd = new FileReader()
  rd.onload = async () => {
    try {
      const rows = parseCsv(rd.result)
      if (!rows.length) { bus.toast('No encontré movimientos válidos'); return }
      const n = await addMovimientos(rows)
      bus.renderAll(); goto('movimientos'); bus.toast(`Importados ${n} movimientos`)
    } catch (err) {
      bus.toast('No pude leer el CSV')
    }
  }
  rd.readAsText(file, 'utf-8')
}
