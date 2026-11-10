# Norte

Finanzas personales que funcionan en tu equipo, sin dar tus claves a nadie.
Importas el extracto de tu banco (CSV), Norte lo ordena por categorías y te enseña
patrimonio, movimientos, análisis de gasto y un indicador de salud financiera ("Pulso").

Todo se guarda **en tu navegador** (IndexedDB). Nada sale de tu dispositivo, no hay servidor
ni cuentas. Es gratis de arrancar: solo necesitas Node.js.

## Requisitos

- [Node.js](https://nodejs.org) 18 o superior (incluye `npm`).

## Arrancar en tu ordenador

```bash
cd norte
npm install      # solo la primera vez
npm run dev
```

Abre la dirección **Local** que imprime (normalmente `http://localhost:5173`).

## Abrirla en el móvil

El servidor ya se expone en tu red local. Con el móvil **en el mismo WiFi** que el ordenador:

1. Arranca `npm run dev`.
2. Mira la línea **Network** que aparece en la terminal, por ejemplo `http://192.168.1.40:5173`.
3. Escribe esa dirección en el navegador del móvil.

Si no carga, suele ser el cortafuegos del ordenador bloqueando el puerto 5173: permite el acceso
a Node en la red privada. Para que se vea como una app, en el móvil usa "Añadir a pantalla de inicio".

## Build de producción

```bash
npm run build      # genera la carpeta dist/
npm run preview    # sirve dist/ para comprobarlo (también en la red local)
```

Lo de `dist/` es estático: puedes subirlo tal cual a cualquier hosting gratuito (Netlify, GitHub Pages,
Cloudflare Pages) el día que quieras, sin backend.

## Datos

- Al abrir por primera vez se cargan **datos de ejemplo** (4 meses) para que veas la app funcionando.
- En **Importar** puedes subir tu propio CSV o cargar/borrar los datos de ejemplo.
- Formato del CSV: `fecha,concepto,importe,categoria` (importe negativo = gasto, positivo = ingreso;
  la categoría es opcional). Detecta separador coma o punto y coma y varias cabeceras habituales de bancos.
- Para empezar de cero: botón **Borrar todos mis datos** en Importar.

## Estructura

```
norte/
  index.html            # estructura de las 5 pantallas
  src/
    main.js             # arranque + router + eventos
    db.js               # base de datos local (Dexie/IndexedDB)
    store.js            # estado en memoria y sincronización con la BD
    seed.js             # datos de ejemplo
    compute.js          # patrimonio, datos por mes, Pulso
    categorias.js       # categorías, iconos y deducción automática
    csv.js              # lector de extractos CSV
    format.js           # moneda y fechas en español
    bus.js              # re-render global + avisos (toast)
    styles.css          # identidad visual Norte (claro/oscuro)
    views/
      resumen.js        # patrimonio, Pulso, cuentas, avisos, recibos
      movimientos.js    # lista con buscador y filtro
      analisis.js       # ingresos/gastos, dónut, evolución
      cuentas.js        # cuentas y tarjetas
      importar.js       # importar CSV / ejemplo / borrar
```

## Qué falta para una "app de verdad" (y cuándo cuesta dinero)

Lo de ahora es gratis y local. Lo siguiente, por orden:

- **Norma 43 (CSB43)**: el extracto estándar de los bancos españoles, para no depender del CSV.
- **Editar categorías** de un movimiento y que aprenda reglas.
- **Presupuestos** por categoría.
- **PWA instalable** y backup/restaurar a archivo.

Y ya con coste, cuando la app te resulte útil:

- **Conexión bancaria (PSD2)**: requiere licencia de agregación o un proveedor (Tink, GoCardless…).
- **Sincronizar móvil ↔ ordenador**: requiere un pequeño servidor.
