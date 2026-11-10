// Categorías, iconos de línea propios y deducción automática de categoría.

export const CATS = {
  'Alimentación': { col: '#0d6e5c', p: '<path d="M6 7h14l-1.2 9a2 2 0 0 1-2 1.7H9.2a2 2 0 0 1-2-1.7L6 7Z"/><path d="M9 7V5.5a3 3 0 0 1 6 0V7"/>' },
  'Restaurantes': { col: '#e8941a', p: '<path d="M6 3v7a2 2 0 0 0 4 0V3M8 10v11"/><path d="M16 3c-1.5 0-2.5 2-2.5 5s1 4 2.5 4v9"/>' },
  'Transporte': { col: '#0ea5a4', p: '<path d="M5 11l1.5-4.5A2 2 0 0 1 8.4 5h7.2a2 2 0 0 1 1.9 1.5L19 11"/><path d="M5 11h14v6H5z"/><circle cx="8" cy="17" r="1.3"/><circle cx="16" cy="17" r="1.3"/>' },
  'Hogar': { col: '#7c6cf0', p: '<path d="M3 10 12 3l9 7"/><path d="M5 9v11h14V9"/><path d="M10 20v-6h4v6"/>' },
  'Suscripciones': { col: '#e85d9b', p: '<path d="M3 5h18v11H3z"/><path d="M8 20h8M12 16v4"/>' },
  'Recibos': { col: '#2d9cdb', p: '<path d="M6 3h12v18l-3-2-3 2-3-2-3 2Z"/><path d="M9 8h6M9 12h6"/>' },
  'Ocio': { col: '#56b870', p: '<path d="M4 5h16v14H4z"/><path d="m10 9 5 3-5 3Z"/>' },
  'Salud': { col: '#eb5757', p: '<path d="M3 12h4l2-5 4 10 2-5h6"/>' },
  'Compras': { col: '#9b8cff', p: '<path d="M6 4h12v16l-6-3-6 3Z"/>' },
  'Nómina': { col: '#19c37d', p: '<path d="M12 2v20"/><path d="M17 6.5c0-2-2.2-3-5-3s-5 1-5 3 2.2 2.6 5 3 5 1.2 5 3-2.2 3-5 3-5-1-5-3"/>' },
  'Ingresos': { col: '#12a970', p: '<path d="M12 2v20"/><path d="M17 6.5c0-2-2.2-3-5-3s-5 1-5 3 2.2 2.6 5 3 5 1.2 5 3-2.2 3-5 3-5-1-5-3"/>' },
  'Otros': { col: '#5a6b65', p: '<circle cx="6" cy="12" r="1.4"/><circle cx="12" cy="12" r="1.4"/><circle cx="18" cy="12" r="1.4"/>' },
}
export const catOf = n => CATS[n] || CATS['Otros']
export const catSVG = (n, stroke) => {
  const c = catOf(n)
  return `<svg class="i" viewBox="0 0 24 24" stroke="${stroke || c.col}">${c.p}</svg>`
}

export const ACCT = {
  corriente: { col: '#0d6e5c', p: '<path d="M3 10 12 4l9 6"/><path d="M5 10v9h14v-9"/><path d="M9 19v-5h6v5"/>' },
  ahorro: { col: '#12a970', p: '<path d="M4 9a7 5 0 0 1 14 0v4a7 5 0 0 1-14 0Z"/><circle cx="14.5" cy="10" r="1"/><path d="M9 6c0-2 2-3 4-2"/>' },
  tarjeta: { col: '#7c6cf0', p: '<path d="M3 6h18v12H3z"/><path d="M3 10h18"/><path d="M7 15h3"/>' },
}
export const acctSVG = tipo => {
  const a = ACCT[tipo] || ACCT.corriente
  return `<svg class="i" viewBox="0 0 24 24" stroke="${a.col}">${a.p}</svg>`
}

export function guessCat(concepto, importe) {
  const c = (concepto || '').toLowerCase()
  if (importe > 0) return /nómina|nomina|salario/.test(c) ? 'Nómina' : 'Ingresos'
  const reglas = [
    ['Alimentación', /mercadona|carrefour|lidl|dia |aldi|supercor|super|aliment|fruter|panader/],
    ['Restaurantes', /restaur|bar |cafet|cafe|burger|pizza|glovo|just eat|uber eats|telepizza|starbucks|mcdon/],
    ['Transporte', /gasolin|repsol|cepsa|bp |parking|renfe|metro|bus|taxi|cabify|uber|peaje|itv|tussam/],
    ['Suscripciones', /netflix|spotify|hbo|disney|amazon prime|youtube|icloud|google|microsoft|dazn|movistar\+|apple\.com|suscrip/],
    ['Recibos', /endesa|iberdrola|naturgy|luz|agua|emasesa|gas |orange|vodafone|movistar|jazztel|masmovil|seguro|comunidad|hipotec|alquiler/],
    ['Hogar', /ikea|leroy|brico|ferret|muebl/],
    ['Salud', /farmac|clinic|dental|hospital|optica|sanitas|adeslas/],
    ['Ocio', /cine|teatro|gym|gimnas|decathlon|padel|museo/],
    ['Compras', /amazon|zara|amzn|aliexpress|mediamarkt|el corte|primark|pccomponentes|shein/],
  ]
  for (const [k, re] of reglas) if (re.test(c)) return k
  return 'Otros'
}
