# Email Automator

Herramienta de línea de comandos (CLI) en Python para **contactar empresas de forma masiva y trazable** a partir de un archivo Excel/CSV:

- **Correo electrónico automático** con plantilla HTML + texto, adjuntos, pausas inteligentes y reintentos.
- **WhatsApp semiautomático** mediante enlaces `wa.me` (el programa abre cada chat con el mensaje listo; tú pulsas enviar).
- **Validación previa** de teléfonos venezolanos y de correos (sintaxis + dominio con DNS).
- **Detección de rebotes** (correos inexistentes) revisando tu bandeja por IMAP.
- **Trazabilidad total**: cada evento queda registrado con su número de fila del Excel.
- **Antifallos**: reanudación sin duplicados, reconexión SMTP, aborto de seguridad y alertas por correo.

> ⚠️ Úsalo de forma responsable: envía solo a contactos profesionales relevantes, personaliza el mensaje y respeta a quien pida no recibir más correos.

---

## Contenido

1. [Estructura del proyecto](#estructura-del-proyecto)
2. [Instalación](#instalación)
3. [Configuración](#configuración)
4. [Uso: correos (`main.py`)](#uso-correos-mainpy)
5. [Uso: WhatsApp (`wa.py`)](#uso-whatsapp-wapy)
6. [Validaciones](#validaciones)
7. [Archivos de registro y trazabilidad](#archivos-de-registro-y-trazabilidad)
8. [Códigos de salida](#códigos-de-salida)
9. [Solución de problemas](#solución-de-problemas)
10. [Buenas prácticas y límites](#buenas-prácticas-y-límites)

---

## Estructura del proyecto

```
email_automator/
├── main.py                  # Envío de correos, validación, rebotes
├── wa.py                    # WhatsApp semiautomático (wa.me)
├── checks.py                # Validación de teléfonos/correos y lectura de rebotes (IMAP)
├── config.toml              # Toda la configuración
├── .env                     # Credenciales (NO subir a Git)
├── .env.example             # Plantilla del .env
├── requirements.txt         # Dependencias
├── data/
│   └── contactos.xlsx       # Tu lista de contactos
├── templates/
│   ├── plantilla.html       # Cuerpo HTML del correo
│   ├── plantilla.txt        # Versión en texto plano del correo
│   └── whatsapp.txt         # Mensaje de WhatsApp
└── logs/                    # Se crea solo: registros e historiales
```

---

## Instalación

**Requisitos:** Python **3.11 o superior**.

### CachyOS / Arch Linux (terminal Fish)

```fish
sudo pacman -S --needed python python-pip git
cd /ruta/a/email_automator
python -m venv .venv
source .venv/bin/activate.fish
pip install -r requirements.txt
```

### Linux/macOS (bash/zsh)

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Windows (PowerShell)

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

> Cada vez que abras una terminal nueva debes **activar el entorno virtual** antes de ejecutar los comandos. Alternativa sin activar: `.venv/bin/python main.py --dry-run`.

**Dependencias:** `pandas`, `openpyxl`, `python-dotenv`, `email-validator`.

> Si el proyecto está en una carpeta de root (p. ej. `/srv/http`), dale permisos: `sudo chown -R $USER:$USER /srv/http/email_automator`.

---

## Configuración

### 1. Credenciales: `.env`

Copia `.env.example` como `.env` y complétalo:

```ini
SMTP_USER=tucorreo@gmail.com
SMTP_PASSWORD=xxxx xxxx xxxx xxxx
SENDER_EMAIL=
```

| Variable | Descripción |
|---|---|
| `SMTP_USER` | Cuenta con la que se inicia sesión en el servidor SMTP/IMAP. |
| `SMTP_PASSWORD` | Contraseña. En Gmail usa una **contraseña de aplicación** (requiere verificación en 2 pasos): <https://myaccount.google.com/apppasswords>. |
| `SENDER_EMAIL` | Opcional. Remitente distinto a `SMTP_USER`. Vacío = usa `SMTP_USER`. |

> Nunca compartas ni subas el `.env`. Ya está incluido en `.gitignore`.

### 2. Contactos: `data/contactos.xlsx`

La **fila 1** son los encabezados; los datos empiezan en la fila 2. Formatos admitidos: `.xlsx`, `.xls`, `.csv`.

| Empresa | Correo | Telefono |
|---|---|---|
| Acme Corp | contacto@acme.com | 0412-1234567 |
| Tech Solutions | info@techsolutions.com | +58 414 1234567 |

- **`Correo`** es obligatoria (el nombre se cambia en `excel.email_column`).
- **`Telefono`** es obligatoria si `validation.validate_phone = true` (si tu Excel no tiene teléfonos, pon `validate_phone = false`).
- Puedes añadir **cualquier columna** (`Contacto`, `Ciudad`, `Rubro`…) y usarla en asunto/plantillas como `{{Columna}}`. El nombre debe coincidir exactamente (mayúsculas, tildes y sin espacios sobrantes).
- Una fila puede tener solo correo, solo teléfono o ambos. Las filas sin correo se rechazan en el envío de emails, pero sí se usan en WhatsApp.

### 3. Ajustes: `config.toml`

Las rutas relativas se resuelven desde la carpeta de `config.toml`.

#### `[smtp]` — servidor de correo saliente

| Clave | Por defecto | Descripción |
|---|---|---|
| `host` | `smtp.gmail.com` | Servidor SMTP (Outlook: `smtp.office365.com`, Zoho: `smtp.zoho.com`). |
| `port` | `587` | `587` (starttls), `465` (ssl), `25` (none). |
| `security` | `starttls` | `"starttls"`, `"ssl"` o `"none"`. |
| `timeout` | `30` | Segundos de espera de red. |

#### `[excel]` — origen de datos

| Clave | Por defecto | Descripción |
|---|---|---|
| `file` | `data/contactos.xlsx` | Ruta del archivo (`.xlsx`, `.xls`, `.csv`). |
| `sheet` | `0` | Hoja: índice (0 = primera) o nombre (`"Hoja1"`). |
| `email_column` | `Correo` | Columna con el correo. |

#### `[message]` — contenido del correo

| Clave | Descripción |
|---|---|
| `subject` | Asunto. Admite `{{Columna}}`. |
| `html_template` | Ruta de la plantilla HTML. |
| `text_template` | Ruta de la plantilla de texto. `""` = se genera automáticamente desde el HTML. |
| `from_name` | Nombre visible del remitente. |
| `reply_to` | Dirección de respuesta (opcional). |
| `list_unsubscribe` | Cabecera de baja, p. ej. `"mailto:baja@tudominio.com"` (recomendado). |
| `attachments` | Lista de adjuntos, p. ej. `["data/CV.pdf"]`. Si un archivo no existe, el programa se detiene antes de enviar. |

#### `[sending]` — ritmo y resiliencia

| Clave | Por defecto | Descripción |
|---|---|---|
| `delay_min` / `delay_max` | `3` / `8` | Pausa aleatoria (segundos) entre correos. |
| `batch_size` | `20` | Cada N correos hace una pausa larga (`0` = desactivar). |
| `batch_pause` | `60` | Duración de la pausa larga (seg). |
| `max_per_run` | `0` | Máximo de correos por ejecución (`0` = sin límite). Gmail gratis ≈ 500/día. |
| `max_retries` | `3` | Reintentos ante errores temporales. |
| `retry_wait` | `10` | Espera base entre reintentos (se multiplica por el nº de intento). |
| `max_consecutive_errors` | `5` | Aborta si fallan N envíos seguidos (protege tu cuenta). |

#### `[validation]` — validación de contactos

| Clave | Por defecto | Descripción |
|---|---|---|
| `check_email_dns` | `true` | Comprueba que el dominio del correo exista y reciba correo (requiere internet). |
| `dns_timeout` | `10` | Segundos de espera por consulta DNS. |
| `validate_phone` | `true` | Valida teléfonos venezolanos. |
| `phone_column` | `Telefono` | Columna del teléfono. |
| `phone_required` | `false` | `true` = rechaza filas sin teléfono. |
| `allow_landline` | `true` | `false` = solo móviles. |
| `on_invalid_phone` | `skip` | `"skip"` = no envía a filas con teléfono inválido; `"warn"` = envía y avisa. |

#### `[files]` — registros

| Clave | Por defecto | Descripción |
|---|---|---|
| `log` | `logs/mailer.log` | Log general. |
| `sent` | `logs/enviados.csv` | Historial de correos (evita duplicados). |
| `trace` | `logs/trazabilidad.csv` | Registro de todos los eventos. |
| `rejected` | `logs/rechazados.csv` | Filas a corregir en el Excel. |

#### `[imap]` — revisión de rebotes

| Clave | Por defecto | Descripción |
|---|---|---|
| `host` | `imap.gmail.com` | Servidor IMAP. |
| `port` | `993` | Puerto IMAP SSL. |
| `days` | `3` | Cuántos días atrás buscar avisos de rebote. |

#### `[alerts]` — notificaciones por correo

| Clave | Por defecto | Descripción |
|---|---|---|
| `mode` | `problems` | `"problems"` = te escribe si hubo fallos, rechazos, rebotes o aborto; `"always"` = siempre; `"never"` = nunca. |
| `to` | `""` | Destinatario de las alertas. Vacío = el correo remitente. |

> Las alertas no se envían en `--dry-run` ni en `--test`.

#### `[whatsapp]` — mensajes por WhatsApp (`wa.py`)

| Clave | Por defecto | Descripción |
|---|---|---|
| `template` | `templates/whatsapp.txt` | Plantilla del mensaje. |
| `name_column` | `Empresa` | Columna mostrada como nombre del contacto. |
| `mobile_only` | `true` | Solo móviles (rechaza fijos). |
| `only_without_email` | `false` | `true` = solo filas **sin** correo. |
| `max_per_run` | `0` | Máximo de chats por ejecución (`0` = sin límite). |
| `history` | `logs/whatsapp_enviados.csv` | Historial de chats confirmados. |
| `links_file` | `logs/whatsapp_links.html` | Página generada por `--list`. |

### 4. Plantillas

- **`plantilla.html` / `plantilla.txt`**: cuerpo del correo en HTML y en texto plano (respaldo). Edita **ambas** y mantenlas coherentes.
- **`whatsapp.txt`**: mensaje de WhatsApp. Admite `*negrita*` y `_cursiva_`. Mantenlo corto.

Dentro de cualquier plantilla (y del asunto) escribe `{{NombreColumna}}` para insertar datos del Excel. En el HTML los valores se escapan automáticamente. Si una plantilla usa una columna que no existe, el programa avisa **antes** de enviar nada.

---

## Uso: correos (`main.py`)

### Flujo recomendado

```fish
python main.py --dry-run                              # 1. Validar y simular (no envía nada)
python main.py --test tucorreo@gmail.com              # 2. Probar con 1 correo real a ti mismo
python main.py --limit 10                             # 3. (Opcional) tanda pequeña
python main.py                                        # 4. Envío completo
python main.py --check-bounces                        # 5. 5-10 min después: revisar rebotes
```

1. **`--dry-run`**: muestra qué filas se rechazan y por qué, y deja `logs/preview.eml` para ver el primer correo. Corrige tu Excel con ayuda de `logs/rechazados.csv` hasta que quede limpio.
2. **`--test`**: envía solo el primer contacto válido a tu correo, con `[PRUEBA]` en el asunto. No se guarda en el historial. Revisa bandeja, spam y móvil.
3. **Envío real**: puedes interrumpir con `Ctrl+C` y volver a ejecutar; **no se repite a quien ya recibió**.
4. **`--check-bounces`**: busca en tu bandeja los avisos de entrega fallida y marca esos correos como `bounced`.

### Flags de `main.py`

| Flag | Descripción |
|---|---|
| `-c`, `--config RUTA` | Usa otro archivo de configuración (por defecto `config.toml` junto a `main.py`). |
| `--dry-run` | Valida contactos, renderiza los mensajes y simula. No conecta al SMTP ni envía. |
| `--test EMAIL` | Envía solo el primer contacto válido a `EMAIL` (asunto con `[PRUEBA]`). No registra en el historial. |
| `--limit N` | Máximo de correos en esta ejecución (tiene prioridad sobre `sending.max_per_run`). |
| `--ignore-history` | Ignora el historial y permite reenviar a quienes ya figuran como enviados/rebotados/rechazados. |
| `--check-bounces` | No envía nada: revisa el buzón por IMAP y marca los rebotes. |
| `-h`, `--help` | Muestra la ayuda. |

### Qué pasa durante el envío

- Cada fila se **valida**: correo (sintaxis + DNS) y teléfono (formato venezolano). Las inválidas se omiten y se anotan.
- Se **deduplican** correos repetidos en el Excel y se saltan los ya procesados.
- Entre correos hay pausas aleatorias; cada `batch_size` correos, una pausa larga.
- Ante **errores temporales** (red caída, límites 4xx) se reintenta con espera creciente y se reconecta automáticamente.
- Si el servidor **rechaza al destinatario** (buzón inexistente) se marca como `rechazado` y no se reintenta.
- Tras `max_consecutive_errors` fallos seguidos, el programa **aborta** para proteger tu cuenta y te avisa.
- Al terminar (o abortar) recibes un **correo resumen** si hubo problemas (según `alerts.mode`).

---

## Uso: WhatsApp (`wa.py`)

WhatsApp no permite el envío 100 % automático sin su API oficial de pago, y simular clics en WhatsApp Web suele terminar en bloqueo del número. Por eso `wa.py` es **semiautomático**: abre cada chat con el mensaje ya escrito y tú pulsas enviar.

Requisitos: sesión iniciada en **WhatsApp Web** (o la app de escritorio) en tu navegador predeterminado.

```fish
python wa.py --dry-run     # muestra los enlaces sin abrir nada
python wa.py               # modo interactivo: abre cada chat
python wa.py --list        # genera una página HTML con todos los enlaces
```

En el modo interactivo, tras abrirse cada chat verás:

```
[1/2] Empresa 6 +584125551041  ->  Enter = ya lo envié | s = saltar | q = salir:
```

| Tecla | Efecto |
|---|---|
| `Enter` | Confirmas que enviaste el mensaje: se guarda en el historial y no se repite. |
| `s` | Saltar este contacto (volverá a aparecer en la próxima ejecución). |
| `q` | Salir. La próxima ejecución continúa donde quedaste. |

### Flags de `wa.py`

| Flag | Descripción |
|---|---|
| `-c`, `--config RUTA` | Usa otro archivo de configuración. |
| `--dry-run` | Muestra los enlaces `wa.me` generados sin abrir el navegador ni registrar nada. |
| `--list` | Genera `logs/whatsapp_links.html` con todos los enlaces y lo abre. No registra envíos en el historial. |
| `--limit N` | Máximo de chats en esta ejecución. |
| `--ignore-history` | Incluye teléfonos ya confirmados como enviados. |
| `-h`, `--help` | Muestra la ayuda. |

Solo se usan filas con teléfono válido. Con `mobile_only = true` se rechazan los fijos, y con `only_without_email = true` solo se contacta a las filas sin correo.

---

## Validaciones

### Teléfonos (Venezuela)

Se aceptan formatos como `0412-1234567`, `+58 414 1234567`, `584261234567`, `04241234567`, `0058 416 1234567` y se normalizan a `+58XXXXXXXXXX`.

- **Móviles válidos:** prefijos `412`, `414`, `416`, `422`, `424`, `426`.
- **Fijos válidos:** códigos de área `2xx` (si `allow_landline = true`).
- Se rechazan caracteres no permitidos, longitudes incorrectas y prefijos que no existen.

### Correos

1. **Sintaxis** correcta del correo.
2. **Dominio con servidor de correo** (registros DNS). El resultado se guarda en caché por dominio para no repetir consultas.

> **Límite importante:** ningún sistema puede comprobar con certeza, antes de enviar, que un buzón concreto exista ni que el mensaje llegue a la bandeja principal. Lo que sí detecta la herramienta es: el dominio inexistente (antes de enviar), el rechazo inmediato del servidor y los **rebotes posteriores** (con `--check-bounces`).

Si tu red bloquea el DNS, pon `check_email_dns = false`.

### Rebotes

`--check-bounces` abre `INBOX` por IMAP, busca mensajes de `mailer-daemon`/`postmaster` de los últimos `imap.days` días (ignora los avisos de simple demora) y cruza las direcciones con los correos enviados. Los que rebotaron se marcan `bounced`, no se volverán a enviar y recibes un aviso.

---

## Archivos de registro y trazabilidad

Todos se crean automáticamente en `logs/`.

| Archivo | Contenido |
|---|---|
| `mailer.log` | Log general en orden cronológico, con errores inesperados y su traceback. |
| `trazabilidad.csv` | **Un registro por cada evento** (ver columnas abajo). Ábrelo con LibreOffice/Excel. |
| `rechazados.csv` | Filas del Excel que debes corregir: `fila_excel, correo, telefono, motivo`. Se reescribe en cada corrida. |
| `enviados.csv` | Historial de correos: `timestamp, email, status, detail`. |
| `whatsapp_enviados.csv` | Historial de WhatsApp (en la columna `email` se guarda el teléfono). |
| `preview.eml` | Vista previa del primer correo generado por `--dry-run`. |
| `whatsapp_links.html` | Página de enlaces generada por `wa.py --list`. |

**Estados en `enviados.csv`:**

| Estado | Significado | ¿Se reenvía en la próxima corrida? |
|---|---|---|
| `ok` | Aceptado por el servidor SMTP (el `detail` guarda el Message-ID). | No |
| `bounced` | Rebotó (detectado con `--check-bounces`). | No |
| `rechazado` | El servidor rechazó al destinatario. | No |
| `error` | Falló tras los reintentos (problema temporal u otro). | Sí |

**Columnas de `trazabilidad.csv`:** `timestamp, run_id, fila_excel, correo, telefono_original, telefono_normalizado, etapa, estado, detalle, message_id`.

- **`run_id`**: identifica cada ejecución (`AAAAMMDD-HHMMSS`; termina en `-sim` para `--dry-run` y `-wa` para WhatsApp).
- **`etapa`**: `validacion`, `simulacion`, `envio`, `rebote`, `whatsapp`, `sistema`.
- **`estado`**: `ok`, `rechazado`, `omitido`, `advertencia`, `enviado`, `error`, `rebotado`, `abierto`, `abortado`, `interrumpido`, `listado`.
- **`fila_excel`**: número de fila real en tu Excel (la fila 1 son los encabezados).

---

## Códigos de salida

Útiles para automatizar o depurar. Aplican a `main.py` y `wa.py`.

| Código | Significado |
|---|---|
| `0` | Todo correcto. |
| `1` | Error de configuración, de conexión o de autenticación. |
| `2` | Terminó, pero hubo filas rechazadas, fallos o rebotes (revisa los logs). |
| `3` | Fallo inesperado (el traceback queda en `logs/mailer.log`). |
| `130` | Interrumpido con `Ctrl+C`. |

---

## Solución de problemas

| Síntoma | Causa y solución |
|---|---|
| `ModuleNotFoundError: No module named 'pandas'` | Entorno virtual sin activar o dependencias sin instalar: `source .venv/bin/activate.fish` y `pip install -r requirements.txt`. |
| `Autenticación rechazada` | Contraseña de aplicación incorrecta o revocada. Genera una nueva. |
| No aparece la opción de contraseñas de aplicación en Google | Activa la verificación en 2 pasos (con un método distinto a solo llaves de seguridad). No disponible con Protección Avanzada ni en algunas cuentas de trabajo/escuela. Alternativa: un SMTP como Brevo (cambia `host`, `port` y las credenciales). |
| `Faltan columnas…` / `Falta la columna…` | El nombre en la plantilla o en `config.toml` no coincide exactamente con el encabezado del Excel. |
| `Falta la columna de teléfono` | Añade la columna o pon `validation.validate_phone = false`. |
| `Adjunto no encontrado` | Revisa la ruta en `message.attachments` (relativa a `config.toml`). |
| Todas las filas se rechazan por dominio | Tu red bloquea el DNS: pon `check_email_dns = false`. |
| Solo llegó el correo y no el mensaje por teléfono | El teléfono se valida pero **no se mensajea** desde `main.py`. Usa `python wa.py`. |
| Filas con teléfono y sin correo no se envían | Es lo esperado en `main.py` (requiere correo). Para ellas usa `wa.py`. |
| Quiero reenviar a todos | `python main.py --ignore-history`. |
| Se abortó por errores consecutivos | Revisa `logs/mailer.log` (p. ej. límite diario de Gmail alcanzado) y reintenta más tarde. |
| `--check-bounces` falla por IMAP | Verifica `[imap]` y que la cuenta permita IMAP con la contraseña de aplicación. |

---

## Buenas prácticas y límites

- **Gmail gratis:** unos 500 correos al día. Usa `max_per_run` (p. ej. 400) y continúa al día siguiente; el historial evita duplicados.
- **Calienta tu cuenta:** empieza con pocos correos y aumenta gradualmente.
- **Personaliza** el mensaje y evita lenguaje de spam. Configura `list_unsubscribe` y respeta las bajas: elimina del Excel a quien lo pida.
- **Limpia tu lista:** revisa `rechazados.csv` y ejecuta `--check-bounces` después de cada envío para no insistir en direcciones inexistentes.
- **WhatsApp:** escribe a pocas empresas por día. Los envíos masivos no solicitados pueden llevar al reporte y bloqueo de tu número.
- **Seguridad:** no subas `.env` ni `data/` ni `logs/` a repositorios públicos (el `.gitignore` incluido ya los excluye).
- **Responsabilidad:** eres responsable del cumplimiento de la normativa de comunicaciones y protección de datos aplicable a tus envíos y a los términos de tu proveedor de correo y de WhatsApp.
