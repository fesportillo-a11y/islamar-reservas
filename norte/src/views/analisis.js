// Pantalla Análisis: ingresos/gastos del mes, dónut por categoría y barras de evolución.
import { monthsSorted, monthData } from '../compute.js'
import { eur, eur0, MES, MESL } from '../format.js'
import { catOf } from '../categorias.js'

const $ = s => document.querySelector(s)

export function render() {
  const months = monthsSorted()
  const key = months.slice(-1)[0]
  if (!key) { $('#legend').innerHTML = '<div class="empty">Sin datos.</div>'; return }

  const d = monthData(key)
  const [y, mo] = key.split('-')
  $('#anMonth').textContent = `${MESL[+mo - 1]} ${y}`
  $('#anIn').textContent = '+' + eur(d.ingresos)
  $('#anOut').textContent = '-' + eur(d.gastos)
  $('#donutTot').textContent = eur0(d.gastos)
  drawDonut(d.porCat)

  const entries = Object.entries(d.porCat).sort((a, b) => b[1] - a[1])
  $('#legend').innerHTML = entries.map(([k, v]) =>
    `<div class="leg"><span class="sw" style="background:${catOf(k).col}"></span><span class="ln">${k}</span><span class="lv">${eur(v)}</span></div>`
  ).join('') || '<div class="empty">Sin gastos este mes.</div>'

  const last = months.slice(-6)
  const vals = last.map(k => monthData(k).gastos)
  const max = Math.max(...vals, 1)
  $('#bars').innerHTML = last.map((k, i) => {
    const h = Math.max(6, Math.round(vals[i] / max * 100))
    const mm = k.split('-')[1]
    const curr = i === last.length - 1
    return `<div class="bar"><div class="col" style="height:${h}%;${curr ? '' : 'opacity:.5'}"></div><div class="bl">${MES[+mm - 1]}</div></div>`
  }).join('')
}

function drawDonut(porCat) {
  const cv = $('#donut'), ctx = cv.getContext('2d')
  const dpr = window.devicePixelRatio || 1, size = 160
  cv.width = size * dpr; cv.height = size * dpr; cv.style.width = size + 'px'; cv.style.height = size + 'px'
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0); ctx.clearRect(0, 0, size, size)
  const entries = Object.entries(porCat).sort((a, b) => b[1] - a[1])
  const tot = entries.reduce((a, b) => a + b[1], 0)
  const cx = size / 2, cy = size / 2, R = 70, r = 46
  if (tot === 0) {
    ctx.beginPath(); ctx.arc(cx, cy, (R + r) / 2, 0, 7)
    ctx.strokeStyle = getCss('--surface-2'); ctx.lineWidth = R - r; ctx.stroke(); return
  }
  let a = -Math.PI / 2
  for (const [k, v] of entries) {
    const ang = v / tot * Math.PI * 2
    ctx.beginPath(); ctx.arc(cx, cy, (R + r) / 2, a, a + ang)
    ctx.strokeStyle = catOf(k).col; ctx.lineWidth = R - r; ctx.lineCap = 'butt'; ctx.stroke()
    a += ang
  }
}
const getCss = v => getComputedStyle(document.documentElement).getPropertyValue(v) || '#ccc'
