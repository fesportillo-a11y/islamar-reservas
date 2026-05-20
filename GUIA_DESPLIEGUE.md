# 🚀 Guía de despliegue — Apartamentos Islamar

Sigue estos 4 pasos en orden. Cada uno tarda entre 5 y 10 minutos.

---

## PASO 1 · Crear la base de datos en Supabase (GRATIS)

1. Ve a **https://supabase.com** y crea una cuenta gratuita.
2. Haz clic en **"New project"**, ponle de nombre `islamar` y elige una contraseña segura.
3. Espera ~2 minutos a que se cree el proyecto.
4. En el menú de la izquierda, haz clic en **"SQL Editor"**.
5. Pega este código y pulsa **"Run"**:

```sql
CREATE TABLE reservas (
  id            BIGSERIAL PRIMARY KEY,
  nro_reserva   TEXT,
  fuente        TEXT,
  mes           TEXT,
  mes_num       INTEGER,
  nombre        TEXT,
  dormitorios   TEXT,
  entrada       TEXT,
  salida        TEXT,
  noches        INTEGER,
  personas      TEXT,
  precio        TEXT,
  pago_cta      TEXT,
  fecha_ingreso TEXT,
  resto_pdte    TEXT,
  estado_pago   TEXT,
  comentarios   TEXT,
  created_at    TIMESTAMPTZ DEFAULT NOW()
);
```

6. Ve a **"Project Settings" → "API"** y copia:
   - **Project URL** (algo como `https://abcdefgh.supabase.co`)
   - **anon public key** (una clave larga que empieza por `eyJ...`)

---

## PASO 2 · Cargar los datos iniciales

1. Abre el archivo `cargar_datos_iniciales.py` con el Bloc de notas.
2. Reemplaza `SUPABASE_URL` y `SUPABASE_KEY` con los valores que copiaste.
3. Abre una ventana de comandos (CMD) en la carpeta `islamar-app` y ejecuta:

```
pip install supabase
python cargar_datos_iniciales.py
```

Verás algo como:
```
  Lote 1: 50 reservas insertadas
  Lote 2: 50 reservas insertadas
  Lote 3: 24 reservas insertadas
✅ Carga completa: 124 reservas en Supabase
```

---

## PASO 3 · Subir la app a GitHub

1. Ve a **https://github.com** y crea una cuenta gratuita.
2. Haz clic en **"New repository"**, nómbralo `islamar-reservas`, márcalo como **Privado** y pulsa **"Create repository"**.
3. En tu ordenador, instala **GitHub Desktop**: https://desktop.github.com
4. Abre GitHub Desktop → **"Add an Existing Repository"** → selecciona la carpeta `islamar-app`.
5. Haz clic en **"Publish repository"** y súbela a GitHub.

---

## PASO 4 · Publicar la app en Streamlit Cloud

1. Ve a **https://streamlit.io/cloud** e inicia sesión con tu cuenta de GitHub.
2. Haz clic en **"New app"**.
3. Selecciona el repositorio `islamar-reservas` y como archivo principal `app.py`.
4. Antes de desplegar, haz clic en **"Advanced settings"** → **"Secrets"** y pega esto
   (sustituyendo con tus credenciales reales):

```toml
SUPABASE_URL = "https://XXXXXXXXXXXXXXXX.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9..."
```

5. Pulsa **"Deploy"** y espera ~1 minuto.

¡Listo! Recibirás una URL del tipo `https://islamar-reservas.streamlit.app`
que puedes compartir con todo tu equipo.

---

## ✅ Resultado final

- La app funciona en cualquier navegador, móvil u ordenador.
- Los datos se guardan automáticamente en Supabase.
- Cualquier cambio que haga uno del equipo lo ven todos al instante.
- Puedes descargar el Excel actualizado en cualquier momento desde la propia app.

---

## ❓ ¿Problemas?

Si algo no funciona, manda una captura de pantalla del error a tu asistente de Claude
y te ayudará a resolverlo en minutos.
