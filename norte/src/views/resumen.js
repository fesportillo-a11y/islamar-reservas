// Pantalla Resumen: patrimonio, Pulso, cuentas, avisos y próximos recibos.
import { state } from '../store.js'
import { patrimonio, pulso, monthsSorted, monthData } from '../compute.js'
import { eur } from '../format.js'
import { catOf, catSVG, acctSVG } from '../categorias.js'

const $ = s => document.querySelector(s)

export function render() {
  $('#net').textContent = eur(patrimonio())

  const months = monthsSorted()
  const prev = months.filter(k => k < '2026-11').slice(-1)[0]
  if (prev) {
    const d = monthData(prev)
    const diff = d.ingresos - d.gastos
    const el = $('#netDelta')
    el.className = 'delta ' + (diff >= 0 ? 'up' : 'down')
    el.textContent = (diff >= 0 ? '▲ ' : '▼ ') + eur(Math.abs(diff)) + ' el mes pasado'
  }

  const s = pulso()
  $('#pulseNum').textContent = s
  const arc = $('#pulseArc')
  arc.style.strokeDashoffset = 314 * (1 - s / 100)
  let col, txt, msg
  if (s >= 70) { col = 'var(--good)'; txt = 'Vas bien'; msg = 'Ahorras de forma estable. Mantén el ritmo.' }
  else if (s >= 45) { col = 'var(--warn)'; txt = 'Vigila'; msg = 'Tus gastos se comen casi todo lo que entra.' }
  else { col = 'var(--bad)'; txt = 'En riesgo'; msg = 'Gastas más de lo que ingresas.' }
  arc.setAttribute('stroke', col)
  const tag = $('#pulseTag'); tag.textContent = txt; tag.style.background = col
  $('#pulseMsg').textContent = msg

  $('#miniAcct').innerHTML = state.cuentas.map(acctRow).join('')

  const al = buildAlerts()
  $('#alertCount').textContent = al.length ? al.length + (al.length > 1 ? ' avisos' : ' aviso') : ''
  $('#alerts').innerHTML = al.length
    ? al.map(a => `<div class="alert"><div class="bar" style="background:${a.col}"></div><div><div class="t">${a.t}</div><div class="d">${a.d}</div></div></div>`).join('')
    : '<div class="empty">Todo en orden, sin avisos.</div>'

  $('#bills').innerHTML = buildBills().map(b =>
    `<div class="row"><div class="ic" style="background:${catOf(b.cat).col}1f">${catSVG(b.cat)}</div><div class="info"><div class="n">${b.n}</div><div class="s">${b.d}</div></div><div class="amt down">-${eur(b.imp).replace('-', '')}</div></div>`
  ).join('')
}

export function acctRow(a) {
  const neg = a.saldo < 0
  return `<div class="row"><div class="ic" style="background:${(a.saldo < 0 ? '#e5533d' : '#0d6e5c')}1f">${acctSVG(a.tipo)}</div>
    <div class="info"><div class="n">${a.nombre}</div><div class="s">${a.banco} · ${a.tipo}</div></div>
    <div class="amt ${neg ? 'down' : ''}">${eur(a.saldo)}</div></div>`
}

function buildAlerts() {
  const out = []
  const months = monthsSorted()
  const curr = months.slice(-1)[0], prev = months.slice(-2)[0]
  if (curr && prev) {
    const dc = monthData(curr), dp = monthData(prev)
    const luz = ms => ms.filter(m => /luz|endesa|iberdrola/i.test(m.concepto)).reduce((a, b) => a + Math.abs(b.importe), 0)
    const luzC = luz(dc.ms), luzP = luz(dp.ms)
    if (luzC > luzP * 1.15 && luzP > 0) out.push({ col: 'var(--warn)', t: 'La luz ha subido', d: `Has pagado ${eur(luzC)}, un ${Math.round((luzC / luzP - 1) * 100)}% más que el mes pasado.` })
    const rc = dc.porCat['Restaurantes'] || 0, rp = dp.porCat['Restaurantes'] || 0
    if (rc > rp * 1.4 && rc > 40) out.push({ col: 'var(--brand)', t: 'Mucho en Restaurantes', d: `Llevas ${eur(rc)} este mes. Vas por encima de tu media.` })
  }
  const nom = state.cuentas.find(c => c.tipo === 'corriente')
  if (nom && nom.saldo < 500) out.push({ col: 'var(--bad)', t: 'Ojo con el saldo', d: `Te quedan ${eur(nom.saldo)} en ${nom.nombre} y aún entran recibos.` })
  const tj = state.cuentas.find(c => c.tipo === 'tarjeta')
  if (tj && tj.saldo < -500) out.push({ col: 'var(--warn)', t: 'Deuda en ' + tj.nombre, d: `Pendiente de liquidar ${eur(Math.abs(tj.saldo))} este mes.` })
  return out
}

function buildBills() {
  return [
    { n: 'Alquiler piso', cat: 'Recibos', imp: 850, d: 'aprox. 1 dic' },
    { n: 'Movistar Fibra+Móvil', cat: 'Recibos', imp: 49.90, d: 'aprox. 22 nov' },
    { n: 'Endesa Luz', cat: 'Recibos', imp: 71.30, d: 'aprox. 3 dic' },
    { n: 'Suscripción vídeo', cat: 'Suscripciones', imp: 13.99, d: 'aprox. 2 dic' },
  ]
}
