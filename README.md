# zabbix-inventory

Genera un inventario en Excel (`.xlsx`) de una instalación **Zabbix 7.0.4**
usando únicamente la API JSON-RPC oficial (sin Ansible, sin SDKs de terceros).

Cubre en esta primera versión: hosts, estado, interfaces, proxy, host
groups, templates (directos), tags y descripción. **No** incluye todavía
LLD/Discovery Rules, host inventory nativo, items, triggers, disponibilidad
ni macros — el código está estructurado para añadir esto después sin
rehacerlo (ver [Próximas ampliaciones](#próximas-ampliaciones)).

## Requisitos

- Python 3.9 o superior
- Un usuario o token de API de Zabbix con permisos de lectura sobre hosts,
  host groups, templates y proxies

## Instalación

```bash
git clone <este-repo>
cd zabbix-inventory
python3 -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## Configuración (`.env`)

Copia `.env.example` a `.env` y rellénalo:

```bash
cp .env.example .env
```

```env
ZABBIX_URL=https://zabbix.example.com
ZABBIX_TOKEN=
ZABBIX_USER=
ZABBIX_PASSWORD=
ZABBIX_OUTPUT=inventario.xlsx
ZABBIX_VERIFY_SSL=true
```

El script usa, por este orden:

1. **`ZABBIX_TOKEN`** si está definido (método preferido).
2. Si no hay token, **`ZABBIX_USER` + `ZABBIX_PASSWORD`**.
3. Si no hay ningún método válido, el script termina con un error claro y
   código de salida distinto de cero.

`ZABBIX_URL` admite tanto la URL base del frontend
(`https://zabbix.example.com`) como la URL completa del endpoint
(`https://zabbix.example.com/api_jsonrpc.php`); el script normaliza
cualquiera de las dos formas.

### Cómo crear un API token

En el frontend de Zabbix: **Administration → General → API tokens →
Create API token**. Asigna el token a un usuario con permisos de solo
lectura sobre los objetos que quieras inventariar. Copia el valor generado
en `ZABBIX_TOKEN` — Zabbix solo lo muestra una vez.

### Certificados TLS autofirmados

Por defecto se valida el certificado del servidor (`ZABBIX_VERIFY_SSL=true`).
Si tu Zabbix usa un certificado interno/autofirmado y no puedes instalar
una CA válida, pon `ZABBIX_VERIFY_SSL=false` explícitamente en `.env`. El
tráfico sigue yendo cifrado por HTTPS; lo único que se deja de comprobar es
la identidad del servidor, así que hazlo solo en entornos de confianza.

## Ejecución

```bash
python zabbix_inventory.py
```

Con salida personalizada:

```bash
python zabbix_inventory.py --output inventario.xlsx
```

Sobrescribiendo la URL del `.env`:

```bash
python zabbix_inventory.py --url https://zabbix.example.com
```

Los parámetros de línea de comandos siempre tienen prioridad sobre los
valores de `.env`.

Código de salida: `0` éxito, `2` error de configuración, `3` error de la
API de Zabbix, `1` error inesperado.

## Estructura del Excel generado

| Hoja | Contenido |
|---|---|
| **Hosts** | Una fila por host, con resúmenes legibles de grupos, templates, tags e interfaces en una única celda cada uno |
| **Interfaces** | Una fila por interfaz configurada |
| **Host Groups** | Una fila por relación host↔grupo |
| **Templates** | Una fila por relación host↔template |
| **Tags** | Una fila por tag (tag + valor) |

Todas las hojas tienen: autofiltro, primera fila congelada, ancho de
columna automático, `wrap_text` en las celdas con múltiples valores y
altura de fila ajustada cuando corresponde. Los hosts se listan ordenados
alfabéticamente por su nombre técnico (`host`).

### Columnas de la hoja `Hosts`

`Host ID, Host, Visible Name, Status, Proxy, Host Groups, Templates, Tags,
Interfaces, Description`

- **Status**: `Enabled` / `Disabled` (nunca el valor numérico interno).
- **Proxy**: nombre del proxy, o `Direct` si el host se monitoriza
  directamente desde el Zabbix Server.
- **Interfaces** (celda resumen): una línea por interfaz, p.ej.
  `Zabbix Agent: 10.10.10.20 (srv01.example.local):10050 [Main]`.

## Decisiones tomadas frente a ambigüedades de la API

- **Host Groups**: se usa `selectHostGroups` (no `selectGroups`, que está
  **deprecado desde Zabbix 6.2** y devuelve una propiedad `groups` distinta
  a `hostgroups`).
- **Templates**: se usa `selectParentTemplates`, que devuelve los
  templates **enlazados directamente** al host. No se resuelven templates
  heredados de otros templates (templates de templates); si en el futuro
  se necesita esa cadena completa, habría que consultar además
  `template.get` con `selectParentTemplates` de forma recursiva.
- **Proxy**: Zabbix 7.0 introduce los *proxy groups*. Si un host se
  monitoriza por un proxy individual (`monitored_by=1`), se muestra su
  nombre real (resuelto con una única llamada `proxy.get`). Si se
  monitoriza por un *proxy group* (`monitored_by=2`), se muestra
  `Proxy group (ID: <id>)`, ya que la resolución de nombres de proxy group
  queda fuera del alcance de esta primera versión.
- **Tags sin valor**: se muestran como `tag_name` (sin `=`) en vez de
  `tag_name=` para que se lean con naturalidad.
- **Paginación**: no se aplica ningún `limit` en `host.get`, por lo que la
  API devuelve el conjunto completo de hosts en una sola llamada. No es
  necesaria paginación manual salvo que tu instalación imponga límites
  artificiales a nivel de proxy/firewall.

## Rendimiento

Todo el inventario se obtiene con **2 llamadas HTTP** en total,
independientemente del número de hosts:

1. `host.get` — una única llamada con `selectInterfaces`,
   `selectHostGroups`, `selectParentTemplates` y `selectTags`, para traer
   todos los hosts y sus relaciones de una vez.
2. `proxy.get` — una única llamada para resolver todos los nombres de
   proxy.

(Más `apiinfo.version`, `user.login`/token y `user.logout` para el ciclo
de conexión.) No se hace ninguna llamada por host.

## Errores comunes

| Síntoma | Causa probable |
|---|---|
| `No Zabbix URL configured` | Falta `ZABBIX_URL` en `.env` y no se pasó `--url` |
| `No valid authentication method configured` | No hay `ZABBIX_TOKEN` ni `ZABBIX_USER`+`ZABBIX_PASSWORD` |
| `TLS certificate validation failed` | Certificado autofirmado/interno; revisa `ZABBIX_VERIFY_SSL` |
| `Authentication check failed` | Token inválido/caducado o usuario/contraseña incorrectos |
| `Zabbix API error on 'host.get'` | El usuario/token no tiene permisos suficientes |
| `Could not connect to ...` | URL incorrecta, firewall, o Zabbix caído |

## Próximas ampliaciones

El código separa configuración, comunicación con la API, procesamiento y
generación de Excel en funciones independientes para poder añadir, sin
reescribir el proyecto:

- LLD / Discovery Rules
- Host Inventory nativo (`selectInventory`)
- Items, Triggers, Availability
- Macros

## Estructura del proyecto

```text
zabbix-inventory/
├── zabbix_inventory.py
├── requirements.txt
├── .env.example
└── README.md
```
# zabbix-inventory-script
