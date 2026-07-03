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

## PASO 7.2 · Habilitar la columna Estado en Listado Raquel (opcional)

El Listado Raquel marca las reservas con etiquetas:
* 🚫 **CANCELADA** — funciona sin SQL extra (se detecta por estado_pago).
* ✨ **NUEVA** — funciona sin SQL extra (usa la columna `created_at` que ya existe).
* 🔄 **MODIFICADA** — necesita una columna `updated_at` para detectar cuándo
  se ha editado una reserva.

Para activar el detector de modificaciones:

```sql
ALTER TABLE reservas ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ DEFAULT NOW();
NOTIFY pgrst, 'reload schema';
```

Sin esto, las CANCELADAS y NUEVAS siguen apareciendo; solo MODIFICADA se queda
en blanco.

---

## PASO 7.3 · Habilitar el campo "Teléfono" en Listado Raquel (opcional)

En los formularios de Nueva/Editar reserva aparece un campo "Teléfono de
contacto". Para que el valor se guarde y se muestre en el Listado Raquel y su
PDF, añade la columna:

```sql
ALTER TABLE reservas ADD COLUMN IF NOT EXISTS telefono TEXT;
NOTIFY pgrst, 'reload schema';
```

Sin este paso, el campo aparece pero el valor no se persistirá (la app
detecta que la columna no existe y guarda el resto de campos sin error).

---

## PASO 7.5 · Seguridad — bloqueo tras 5 intentos fallidos

Protección contra ataques de fuerza bruta. Un usuario que falla la
contraseña 5 veces en 15 minutos queda bloqueado 15 minutos adicionales.

### 1. Crear la tabla en Supabase

Ejecuta este SQL en **Supabase → SQL Editor → New query → Run**:

```sql
CREATE TABLE IF NOT EXISTS public.login_attempts (
  id            BIGSERIAL PRIMARY KEY,
  username      TEXT NOT NULL,
  success       BOOLEAN NOT NULL,
  attempted_at  TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_login_attempts_username_time
  ON public.login_attempts (username, attempted_at DESC);

-- Row Level Security con políticas abiertas (la app usa la anon key
-- y toda la seguridad real la aporta el login + este bloqueo).
ALTER TABLE public.login_attempts ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS login_attempts_select_all ON public.login_attempts;
DROP POLICY IF EXISTS login_attempts_insert_all ON public.login_attempts;

CREATE POLICY login_attempts_select_all
  ON public.login_attempts FOR SELECT
  TO anon, authenticated USING (true);
CREATE POLICY login_attempts_insert_all
  ON public.login_attempts FOR INSERT
  TO anon, authenticated WITH CHECK (true);

GRANT SELECT, INSERT ON public.login_attempts TO anon, authenticated;
GRANT USAGE, SELECT ON SEQUENCE login_attempts_id_seq TO anon, authenticated;

NOTIFY pgrst, 'reload schema';
```

### 2. Verificar

Con esta tabla creada, la app:

- Registra CADA intento de login (éxito o fallo) en `login_attempts`.
- Si un usuario acumula **5 fallos en 15 minutos**, queda bloqueado
  durante 15 min más. La app muestra un mensaje claro con el tiempo
  restante.
- Si al llegar al 3er o 4º fallo el usuario ya está cerca del límite,
  se le avisa: "Te quedan N intento(s) antes del bloqueo".
- Los intentos exitosos reinician el contador.

### 3. Ver el histórico (opcional)

Desde el propio Supabase:

```sql
SELECT username, success, attempted_at
FROM login_attempts
WHERE attempted_at > NOW() - INTERVAL '1 day'
ORDER BY attempted_at DESC;
```

Sin esta tabla la app sigue funcionando, pero SIN protección contra
brute-force (todos los intentos pasan sin registrarse ni bloquearse).

---

## PASO 7.8 · Facturas formales tipo empresa

Además del **documento de reserva** para particulares (PASO 7.4),
la app puede emitir **facturas formales** con la estructura tabular,
IVA, referencia de obra y demás datos que piden los clientes empresariales.

### 1. Añadir columnas a la tabla `reservas`

```sql
ALTER TABLE public.reservas ADD COLUMN IF NOT EXISTS cliente_direccion    TEXT;
ALTER TABLE public.reservas ADD COLUMN IF NOT EXISTS cliente_cp_localidad TEXT;
ALTER TABLE public.reservas ADD COLUMN IF NOT EXISTS ref_obra             TEXT;
ALTER TABLE public.reservas ADD COLUMN IF NOT EXISTS nro_factura_emp      TEXT;
ALTER TABLE public.reservas ADD COLUMN IF NOT EXISTS fecha_factura_emp    TEXT;
ALTER TABLE public.reservas ADD COLUMN IF NOT EXISTS proveedor            TEXT;
ALTER TABLE public.reservas ADD COLUMN IF NOT EXISTS pedido               TEXT;
ALTER TABLE public.reservas ADD COLUMN IF NOT EXISTS iva_porcentaje       NUMERIC DEFAULT 0;
NOTIFY pgrst, 'reload schema';
```

### 2. Cómo usarlo

En el menú lateral aparece **📋 Facturas**:

1. Busca al cliente por nombre.
2. Selecciona la reserva concreta.
3. Rellena:
   - **Dirección**, **CP + Localidad**, **CIF/NIF** del cliente.
   - **Nº factura** (por defecto siguiente correlativo `NNN-AP26` — en 2026
     arranca en 003).
   - **Fecha** (por defecto hoy).
   - **IVA %**: 0 = exenta (añade automáticamente la nota "Factura exenta
     de IVA según apartado 23 del artículo 20.1 de la Ley 37/1992").
   - Opcional: **REF. OBRA**, **Proveedor**, **Pedido**.
4. Pulsa **📄 Emitir factura** para persistir los datos en BD.
5. Pulsa **⬇️ Descargar PDF factura**.

El PDF sigue el formato del modelo tipo `002-AP26`:
cabecera con logo y datos registrales, cliente a la derecha, cuadro de
Fecha/Factura/Proveedor/Pedido, banner azul "FACTURA", tabla
Descripción/Cantidad/P.Unitario/TOTAL y pie con Base Imp / IVA / Ret /
Total Factura.

Coexiste con la sección **📄 Documento de reserva** (particulares) —
usa la que corresponda según el tipo de cliente.

---

## PASO 7.7 · Seguridad — 2FA para administradores

Verificación en dos pasos con **Google Authenticator / Authy / Microsoft
Authenticator**. Obligatorio para usuarios con `rol='admin'`, opcional para
los demás.

### 1. Añadir columnas a la tabla `usuarios` en Supabase

Ejecuta en **SQL Editor → New query → Run**:

```sql
ALTER TABLE public.usuarios ADD COLUMN IF NOT EXISTS mfa_enabled BOOLEAN DEFAULT FALSE;
ALTER TABLE public.usuarios ADD COLUMN IF NOT EXISTS mfa_secret  TEXT;
NOTIFY pgrst, 'reload schema';
```

### 2. Primer login como admin tras el despliegue

En la primera entrada, la app pedirá configurar el 2FA:

1. Introduces usuario + contraseña como siempre.
2. Aparece una pantalla con un **código QR** y unas instrucciones.
3. Abres **Google Authenticator** (o Authy / Microsoft Authenticator) en
   el móvil → **"+"** → **"Escanear código QR"** → apuntas al QR.
4. La app del móvil añade una entrada "ISLAMAR · ESTEASUR 2015" con un
   código de 6 dígitos que cambia cada 30 s.
5. Copias ese código en el campo de la web y pulsas **"Activar 2FA"**.
6. A partir de ahí, cada vez que hagas login te pedirá el código actual.

### 3. Recuperación si se pierde el móvil (admin de rescate)

Si un admin pierde el móvil, otro admin puede desactivarle el 2FA con
este SQL directo:

```sql
UPDATE public.usuarios
SET mfa_enabled = FALSE, mfa_secret = NULL
WHERE username = 'usuario_afectado';
```

Después el usuario podrá loguearse solo con contraseña y volver a
configurar 2FA en el siguiente intento.

Los **admins de rescate** definidos en Streamlit Secrets (`festeban` de
`[auth.credentials.usernames]`) NO están en la tabla `usuarios`, por lo
que a ellos NO se les pide 2FA (son la última red de seguridad). Si en
algún momento quieres quitar esa exención, avísame.

Sin esta configuración la app sigue funcionando pero SIN 2FA.

---

## PASO 7.6 · Seguridad — log de auditoría

Registro completo de qué usuario crea/edita/borra qué reserva y cuándo.
Panel "🔒 Auditoría" en el menú lateral solo para admins.

### 1. Crear la tabla en Supabase

Ejecuta este SQL en **Supabase → SQL Editor → New query → Run**:

```sql
CREATE TABLE IF NOT EXISTS public.auditoria (
  id          BIGSERIAL PRIMARY KEY,
  usuario     TEXT NOT NULL,
  nombre      TEXT,
  accion      TEXT NOT NULL,               -- CREAR / EDITAR / ELIMINAR
  id_reserva  BIGINT,
  detalles    JSONB,
  creado_en   TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_auditoria_creado_en
  ON public.auditoria (creado_en DESC);
CREATE INDEX IF NOT EXISTS idx_auditoria_usuario
  ON public.auditoria (usuario);
CREATE INDEX IF NOT EXISTS idx_auditoria_id_reserva
  ON public.auditoria (id_reserva);

ALTER TABLE public.auditoria ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS auditoria_select_all ON public.auditoria;
DROP POLICY IF EXISTS auditoria_insert_all ON public.auditoria;

CREATE POLICY auditoria_select_all
  ON public.auditoria FOR SELECT
  TO anon, authenticated USING (true);
CREATE POLICY auditoria_insert_all
  ON public.auditoria FOR INSERT
  TO anon, authenticated WITH CHECK (true);

GRANT SELECT, INSERT ON public.auditoria TO anon, authenticated;
GRANT USAGE, SELECT ON SEQUENCE auditoria_id_seq TO anon, authenticated;

NOTIFY pgrst, 'reload schema';
```

### 2. Uso

Después, entra a la app como admin y verás en el menú lateral la nueva
sección **"🔒 Auditoría"**. Ahí puedes:

- Filtrar por rango de fechas, tipo de acción y usuario.
- Ver el detalle completo de cada cambio (cliente, apartamento, fechas,
  precio, etc.).
- Descargar el histórico como CSV para archivo o análisis externo.

Sin esta tabla la app sigue funcionando, pero los cambios NO se registran.

---

## PASO 7.4 · Habilitar la emisión del documento de reserva (opcional)

Para que la app pueda emitir el documento de reserva PDF tipo ESTEASUR
(con número correlativo automático tipo `040-26`), añade estas tres
columnas a la tabla `reservas`:

```sql
ALTER TABLE reservas ADD COLUMN IF NOT EXISTS nif           TEXT;
ALTER TABLE reservas ADD COLUMN IF NOT EXISTS nro_factura   TEXT;
ALTER TABLE reservas ADD COLUMN IF NOT EXISTS fecha_factura TEXT;
NOTIFY pgrst, 'reload schema';
```

Después, en **✏️ Editar reserva** verás una sección "📄 Documento de
reserva" con:

- Campo **N.I.F. del cliente** (obligatorio para emitir).
- Campos **Nº de documento** y **Fecha de emisión** editables. El nº se
  pre-rellena con el próximo correlativo automático (`NNN-YY` donde `YY`
  son los 2 últimos dígitos del año) y la fecha con hoy, pero puedes
  poner los valores que quieras.
- Botón **📄 Emitir documento** → guarda el nº y la fecha en BD.
- Botón **⬇️ Descargar PDF documento** → genera el PDF con el formato
  tipo ESTEASUR (logo, datos del emisor + cliente, concepto descriptivo
  de la reserva, base imponible, pago a cuenta, resto pendiente, datos
  bancarios y nota inferior).

Sin estos campos en BD la sección aparece pero no podrá persistir el nº
de documento ni la fecha de emisión.

---

## PASO 7.1 · Habilitar el campo "Forma de pago" (opcional)

En el formulario de **Nueva reserva** y **Editar reserva** hay un desplegable
"Forma de pago" con opciones Bankinter / Santander / La Caixa. Para que el
valor se guarde en la BD, añade la columna:

```sql
ALTER TABLE reservas ADD COLUMN IF NOT EXISTS forma_pago TEXT;
NOTIFY pgrst, 'reload schema';
```

Si no haces este paso, el desplegable sigue apareciendo y se puede usar, pero
el valor no se persistirá (la app detecta que la columna no existe y guarda el
resto de campos sin error).

---

## PASO 8 · Instalar la app en el móvil (como si fuera nativa)

La app está adaptada a móvil. Si la usas desde el navegador de tu teléfono, ya
se ve bien (sidebar plegada por defecto, botones grandes, tablas con scroll
lateral, etc.). Pero puedes ir un paso más: **instalarla como icono en la
pantalla de inicio del móvil**, igual que una app de la AppStore o Play Store.
Es gratis y no hace falta instalar nada raro.

### iPhone / iPad (Safari)

1. Abre la URL de tu app en **Safari** (no en Chrome — solo Safari permite
   instalar apps web en iOS).
2. Pulsa el icono de **Compartir** (un cuadrado con una flecha hacia arriba)
   en la barra inferior.
3. Desliza hacia abajo en el menú hasta **"Añadir a pantalla de inicio"**.
4. Te aparece una vista previa con el nombre **"ISLAMAR"**. Pulsa **"Añadir"**.

Listo: en tu pantalla de inicio aparecerá un icono nuevo. Cuando lo pulses,
la app se abre **a pantalla completa, sin barra de Safari**, como cualquier
otra app.

### Android (Chrome / Edge / Samsung Internet)

1. Abre la URL de tu app en **Chrome** (o el navegador que uses).
2. Pulsa el menú de tres puntos arriba a la derecha **⋮**.
3. Elige **"Instalar aplicación"** o **"Añadir a la pantalla principal"**.
4. Confirma. El icono aparece en el lanzador de apps.

Al pulsarlo se abre fullscreen, igual que en iOS.

### Cosas a saber

- **No es una "app nativa"** propiamente dicha, es una web disfrazada de app
  (lo llaman "PWA"). Para el día a día se comporta exactamente como una app
  normal.
- Para entrar siempre tienes que **iniciar sesión** con tu usuario y
  contraseña. La cookie de sesión dura 30 días por defecto.
- Si quieres cambiar el **nombre que aparece bajo el icono**, edítalo en la
  pantalla de "Añadir a inicio" antes de confirmar.

---

## ✅ Resultado final

- La app funciona en cualquier navegador, móvil u ordenador.
- Se puede **instalar en el móvil** como si fuera una app.
- Solo entran los usuarios registrados en los Secrets.
- Los datos se guardan automáticamente en Supabase.
- Cualquier cambio que haga uno del equipo lo ven todos al instante.
- Puedes descargar el Excel actualizado en cualquier momento desde la propia app.

---

## ❓ ¿Problemas?

Si algo no funciona, manda una captura de pantalla del error a tu asistente de Claude
y te ayudará a resolverlo en minutos.
