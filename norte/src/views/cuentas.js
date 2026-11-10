// Pantalla Cuentas: productos del usuario y suma de patrimonio.
import { state } from '../store.js'
import { patrimonio } from '../compute.js'
import { eur } from '../format.js'
import { acctRow } from './resumen.js'

const $ = s => document.querySelector(s)

export function render() {
  $('#fullAcct').innerHTML = state.cuentas.map(acctRow).join('') +
    `<div class="daterow">Patrimonio total</div>
     <div class="row"><div class="ic" style="background:var(--brand-soft)"><svg class="i" viewBox="0 0 24 24" stroke="var(--brand)"><path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/></svg></div>
     <div class="info"><div class="n">Suma de cuentas</div><div class="s">${state.cuentas.length} productos</div></div>
     <div class="amt">${eur(patrimonio())}</div></div>`
}
