// Puente para que cualquier vista pueda pedir un re-render completo sin dependencias circulares.
export let renderAll = () => {}
export function setRenderAll(fn) { renderAll = fn }

let toastTimer
export function toast(msg) {
  const el = document.getElementById('toast')
  if (!el) return
  el.textContent = msg
  el.classList.add('show')
  clearTimeout(toastTimer)
  toastTimer = setTimeout(() => el.classList.remove('show'), 2200)
}
