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

## PASO 5 · Cerrar la app con usuario y contraseña

La app está protegida con login. Sin esta configuración, **nadie puede entrar** (la
app muestra un aviso y se bloquea). Sigue estos pasos UNA sola vez por cada usuario
que quieras dar de alta.

### 5.1 Generar el hash de la contraseña

En tu ordenador, dentro de la carpeta del repo:

```
pip install bcrypt
python tools/hash_password.py
```

Te pedirá la contraseña dos veces (no se ve mientras la escribes) y devolverá algo
como esto — **es el hash**, no la contraseña original:

```
password = "$2b$12$abcDEFghij...largo y único..."
```

Copia esa línea. Repite por cada usuario.

### 5.2 Añadir la sección [auth] a los Secrets de Streamlit Cloud

En Streamlit Cloud → tu app → **⋮ → Settings → Secrets**, **AÑADE** esto debajo de
lo que ya tienes (no borres `SUPABASE_URL` ni `SUPABASE_KEY`):

```toml
[auth.cookie]
name        = "islamar_auth_cookie"
key         = "PEGA_AQUÍ_UNA_CADENA_LARGA_Y_ALEATORIA"
expiry_days = 30

[auth.credentials.usernames.festeban]
email    = "festeban@esportillo.es"
name     = "Francisco Esteban"
password = "$2b$12$...el hash que generaste..."

# Para añadir un segundo usuario, copia el bloque anterior con otro nombre:
# [auth.credentials.usernames.juana]
# email    = "juana@ejemplo.com"
# name     = "Juana López"
# password = "$2b$12$..."
```

- `name`: lo que se mostrará "👤 Francisco Esteban" en el sidebar.
- El **nombre detrás de `usernames.`** (ej. `festeban`) es el USUARIO con el que
  inicia sesión.
- `key` de la cookie: cadena aleatoria larga. Si la cambias, todos los usuarios
  tendrán que volver a iniciar sesión. Puedes generar una con:

  ```
  python -c "import secrets; print(secrets.token_urlsafe(48))"
  ```

### 5.3 Guardar y verificar

Pulsa **Save** en Secrets. Streamlit Cloud reinicia la app en ~30 segundos.
Al entrar, ahora verás la pantalla **"Iniciar sesión"**. Introduce el usuario y la
contraseña ORIGINAL (no el hash). Si todo va bien, entras y aparece tu nombre y un
botón "🚪 Cerrar sesión" en el sidebar.

### 5.4 Añadir / quitar usuarios después

Tienes dos formas de hacerlo, según preferencia:

**Forma fácil (recomendada): desde la propia app.**
Una vez tengas creada la tabla `usuarios` (ver PASO 6), entras en la app como
admin y vas a **👥 Usuarios** en el menú lateral. Desde ahí das de alta, cambias
contraseñas o eliminas usuarios sin tocar Secrets.

**Forma manual (Streamlit Cloud → Misterios):**
Edita los Secrets:
- **Añadir usuario**: copia el bloque `[auth.credentials.usernames.nuevo]` y pega
  el hash de su contraseña.
- **Quitar usuario**: borra su bloque entero.
- **Cambiar contraseña**: genera nuevo hash con `tools/hash_password.py` y reemplaza
  el valor de `password` del usuario.

Los usuarios definidos en Misterios son "admins de rescate": siempre pueden entrar
aunque la BD falle, y NO se pueden modificar desde la pantalla 👥 Usuarios.

---

## PASO 6 · Habilitar la gestión de usuarios desde la app

Para poder gestionar usuarios desde la pantalla **👥 Usuarios** dentro de la app
(en vez de tocar Misterios cada vez), hace falta crear una tabla en Supabase.
Se hace UNA SOLA VEZ.

### 6.1 Crear la tabla `usuarios`

1. Entra en **https://supabase.com** y abre tu proyecto.
2. Menú izquierdo → **"SQL Editor"** → **"New query"**.
3. Pega esto y pulsa **"Run"**:

```sql
CREATE TABLE IF NOT EXISTS usuarios (
  id            BIGSERIAL PRIMARY KEY,
  username      TEXT UNIQUE NOT NULL,
  nombre        TEXT,
  email         TEXT,
  password_hash TEXT NOT NULL,
  rol           TEXT NOT NULL DEFAULT 'usuario',
  activo        BOOLEAN NOT NULL DEFAULT TRUE,
  created_at    TIMESTAMPTZ DEFAULT NOW()
);
```

Verás algo tipo `Success. No rows returned`.

### 6.2 Usar la pantalla 👥 Usuarios

Recarga tu app. Si estás logueado como admin (los usuarios de Misterios son
admins automáticamente), te aparece **👥 Usuarios** en la barra lateral. Desde
ahí:

- **➕ Dar de alta**: usuario, nombre, email, rol (`usuario` o `admin`) y contraseña.
- **📋 Lista de usuarios**: edita datos, cambia contraseñas, activa/desactiva
  o elimina usuarios.

La app no te dejará:
- Borrarte a ti mismo.
- Quedarte sin ningún administrador activo.

Si quieres que cualquiera quite o añada usuarios sin necesidad de tocar Secrets,
crea un primer usuario con `rol = admin` desde **👥 Usuarios** y pásale las
credenciales por mensaje.

---

## PASO 7 · Habilitar el desglose Adultos / Niños (opcional)

El Listado Raquel muestra el número total de personas alojadas. Si quieres ver
**también** el desglose de adultos y niños (ej. `5 (3 ad + 2 niños)`), hay que
añadir dos columnas a la tabla `reservas`. **Es opcional**: sin esto, el listado
seguirá mostrando solo el total.

### 7.1 Crear las columnas

1. En Supabase → **SQL Editor → New query**.
2. Pega esto y pulsa **Run**:

```sql
ALTER TABLE reservas ADD COLUMN IF NOT EXISTS adultos INTEGER;
ALTER TABLE reservas ADD COLUMN IF NOT EXISTS ninos   INTEGER;
NOTIFY pgrst, 'reload schema';
```

Verás "Success. No rows returned".

### 7.2 Reimportar las reservas de Booking

Las reservas que importes a partir de ahora ya guardarán el desglose. Para que
las **ya existentes** muestren el desglose, vuelve a importar el Excel de
Booking más reciente: el importador detectará las que ya están y solo
actualizará los campos que vienen del Excel (sin tocar lo que has editado a
mano), aplicando ahora también `adultos` y `ninos`.

---

## ✅ Resultado final

- La app funciona en cualquier navegador, móvil u ordenador.
- Solo entran los usuarios registrados en los Secrets.
- Los datos se guardan automáticamente en Supabase.
- Cualquier cambio que haga uno del equipo lo ven todos al instante.
- Puedes descargar el Excel actualizado en cualquier momento desde la propia app.

---

## ❓ ¿Problemas?

Si algo no funciona, manda una captura de pantalla del error a tu asistente de Claude
y te ayudará a resolverlo en minutos.
