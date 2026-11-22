# Guía de referencia — Excel Template Builder v12

> Documentación técnica completa: instalación, variables de entorno, conceptos, zonas repetibles, automatización por
> herramienta, lotes, validación, API, colaboración, métricas, seguridad y solución de problemas.
> Para una vista general con capturas de pantalla consulta el [README](../README.md).

---

## Índice
1. [Requisitos](#requisitos)
2. [Instalación y ejecución](#instalación-y-ejecución)
3. [Variables de entorno](#variables-de-entorno)
4. [Carpeta de datos](#carpeta-de-datos)
5. [Conceptos clave](#conceptos-clave)
6. [Filas y columnas repetibles](#filas-y-columnas-repetibles)
7. [Automatiza en 3 pasos](#automatiza-en-3-pasos)
8. [Ejemplos por herramienta](#ejemplos-por-herramienta)
9. [Lotes → ZIP](#lotes--zip)
10. [Validación y modo estricto](#validación-y-modo-estricto)
11. [Modo base64](#modo-base64)
12. [OpenAPI y Postman](#openapi-y-postman)
13. [Referencia de rutas](#referencia-de-rutas)
14. [Colaboración en tiempo real](#colaboración-en-tiempo-real)
15. [Actividad y métricas de uso](#actividad-y-métricas-de-uso)
16. [Seguridad](#seguridad)
17. [Solución de problemas](#solución-de-problemas)
18. [Para desarrolladores](#para-desarrolladores)
19. [Changelog v12](#changelog-v12)

---

## Requisitos
- **Node.js 18 o superior** (probado con Node 25).
- Un navegador moderno (Edge, Chrome, Firefox o Safari).
- No necesitas tener Excel instalado para generar archivos.

## Instalación y ejecución
```bash
npm install          # instala las dependencias (solo la primera vez o al actualizar)
npm start            # inicia el servidor (puerto 3001 por defecto)
npm run dev          # modo desarrollo: se reinicia al guardar archivos (node --watch)
npm test             # ejecuta las pruebas automáticas (node:test)
```
Después abre **http://localhost:3001** en el navegador. Al arrancar, la consola muestra las direcciones útiles
(salud, OpenAPI, uso y tiempo real).

En Linux/macOS también puedes usar los scripts incluidos; `run.sh` localiza Node automáticamente (incluso si usas nvm):
```bash
./scripts/install.sh      # npm ci --omit=dev (instalación para producción)
./scripts/run.sh          # puerto por defecto
./scripts/run.sh 8080     # puerto como primer argumento
```

## Variables de entorno
| Variable | Predeterminado | Descripción |
|---|---|---|
| `PORT` | `3001` | Puerto HTTP. También se acepta como primer argumento: `node main.js 8080` (el argumento tiene prioridad). |
| `HOST` | `0.0.0.0` | Interfaz donde escucha. Usa `127.0.0.1` para aceptar solo conexiones del mismo equipo. |
| `ETB_DATA_DIR` | `./data` | Carpeta donde se guardan plantillas y métricas de uso. |
| `CORS_ORIGINS` | *(vacío = CORS desactivado)* | Orígenes de otros sitios que pueden llamar a la API desde el navegador, separados por comas, p. ej. `https://intranet.miempresa.mx,https://portal.miempresa.mx`. No afecta a cURL, scripts ni Power Automate. |
| `TRUST_PROXY` | `loopback` | Valor de `trust proxy` de Express. Si publicas detrás de un proxy inverso ajústalo (p. ej. `1`) para que los límites de tasa y las métricas vean la IP real del cliente. |

Ejemplos:
```bash
PORT=8080 ETB_DATA_DIR=/srv/etb-datos npm start                 # bash
```
```powershell
$env:PORT = 8080; $env:ETB_DATA_DIR = 'D:\etb-datos'; npm start   # PowerShell
```

## Carpeta de datos
```
<ETB_DATA_DIR>/
├── templates/
│   ├── <id>.json      # una plantilla por archivo (incluye "revision", "createdAt", "updatedAt", "updatedBy")
│   └── …
└── usage/
    └── usage.json     # métricas de uso (escritura atómica, agrupada cada ~2 s)
```
- **Respaldar** = copiar esta carpeta (de preferencia con el servidor detenido).
- Las plantillas de versiones anteriores (sin `revision`) se cargan tal cual; obtienen `revision` en su siguiente guardado.
- Las escrituras son atómicas (archivo temporal + renombrar), así que un corte de luz no deja archivos a medias.

---

## Conceptos clave

**Plantilla.** Un libro con una o varias hojas. Se identifica por un `id` (letras, números, `_` y `-`, máximo 64 caracteres).
Lo ves en el asistente **Automatizar** o en `GET /api/templates`.

**Variables `{{nombre}}`.** Marcan dónde van tus datos. Una celda se declara como variable desde el panel *Variables* con:
- **tipo**: `text`, `number`, `boolean`, `date` o `image`;
- **obligatoria** sí/no, **pista** (texto de ayuda) y **valor por defecto** (se usa en el cuerpo de ejemplo).

| Tipo | Qué enviar | Ejemplo | Cómo queda en Excel |
|---|---|---|---|
| `text` | texto, número o booleano | `"Ana López"` | texto tal cual |
| `number` | número o texto numérico | `1160.5` o `"1160.5"` | número (si no es numérico, `0`) |
| `boolean` | `true` / `false` (también se aceptan `"true"`, `"false"`, `0`, `1`) | `true` | VERDADERO/FALSO |
| `date` | fecha reconocible, de preferencia ISO | `"2026-01-15"` | fecha de Excel |
| `image` | data URL en base64: PNG, JPEG o GIF | `"data:image/png;base64,iVBOR…"` | imagen de 150×100 px sobre la celda |

> 💡 Para booleanos envía `true`/`false` de JSON. Los textos `"true"`/`"false"` pasan la validación, pero el generador
> convierte cualquier texto no vacío en VERDADERO.

**Variables implícitas.** Cualquier texto o fórmula puede contener `{{nombre}}` aunque la celda no esté declarada como
variable, por ejemplo `Atendió: {{vendedor}}` o `=B5*{{tasa}}`. Se reemplazan igual que las declaradas. En el esquema de
automatización aparecen con `implicit: true`, tipo `text` y no obligatorias; si el mismo nombre también está declarado, gana la
declaración explícita.

**El cuerpo (body) se organiza por nombre de hoja.** Cada clave es el nombre exacto de una hoja (respeta mayúsculas y acentos):
```json
{ "Factura": { "cliente": "Ana López", "total": 1160 } }
```

**Hojas repetibles.** Si una hoja está marcada como repetible, envía un **arreglo**: se crea una hoja por elemento
(`Pedido 1`, `Pedido 2`…; con un solo elemento conserva su nombre). Si mandas un objeto se usa como un único elemento (con
advertencia) y un arreglo vacío omite la hoja.
```json
{ "Pedido": [ { "folio": "P-1" }, { "folio": "P-2" } ] }
```

**Filas y columnas repetibles (zonas).** Un rango de filas (p. ej. de la 5 a la 7) o de columnas (p. ej. B a D) que se copia
por cada elemento de un arreglo cuyo nombre define la plantilla (`arrayVariable`, p. ej. `items` o `meses`). Una hoja puede
tener varias zonas, en ambas direcciones. Las variables de la zona van **dentro de cada elemento**; lo que está después se
desplaza y las fórmulas se ajustan solas. Detalles y ejemplos en [Filas y columnas repetibles](#filas-y-columnas-repetibles).
```json
{ "Factura": { "cliente": "Ana", "total": 1160, "items": [
  { "descripcion": "Servicio", "cantidad": 2, "precio": 500 },
  { "descripcion": "Envío",    "cantidad": 1, "precio": 160 } ] } }
```

---

## Filas y columnas repetibles

Las plantillas ya no son estáticas: cualquier bloque de **filas** o de **columnas** puede repetirse tantas veces como
elementos traiga un arreglo del body. Ejemplos típicos: renglones de una factura (filas), meses de un reporte (columnas), o
ambos en la misma hoja.

### Cómo se crea en el editor
1. Selecciona las filas o columnas (clic en los encabezados; arrastra para varias).
2. Pulsa **Repetir** en la cinta (columnas completas → columnas; si no, filas) o usa el menú contextual
   **Repetir filas…** / **Repetir columnas…**.
3. Escribe el **nombre del arreglo** (p. ej. `meses`). La ventana muestra qué se va a copiar y el JSON exacto que debes enviar.

La zona se ve en la cuadrícula como una banda punteada con su nombre (naranja para filas, morado para columnas) y en el
panel lateral **Repetición**, donde puedes seleccionarla, editarla o quitarla. Las variables que están dentro de una zona
muestran su arreglo (`meses[]`). También se administran desde **Configuración de hoja**. Al insertar o eliminar filas/columnas
las zonas se recorren solas, y **Deshacer/Rehacer** también cubre las zonas.

Reglas: dos zonas de la misma dirección no pueden superponerse; una zona de filas y una de columnas pueden convivir siempre
que no haya celdas con contenido en su cruce. El nombre del arreglo usa letras, números y `_`.

### Cómo se envían los datos
Cada zona es un arreglo de objetos dentro de los datos de su hoja. Ejemplo: la columna **B** repite `meses` y la fila **5**
repite `productos`:

| | A | B | C |
|---|---|---|---|
| **1** | Concepto | `{{mes}}` | Total |
| **2** | Ventas | `{{ventas}}` | `=SUMA(B2:B2)` |
| **5** | `{{producto}}` | | |
| **7** | Productos: | `=CONTARA(A5:A5)` | |

```json
{
  "Ventas": {
    "nota": "Q1",
    "meses": [
      { "mes": "Ene", "ventas": 100 }, { "mes": "Feb", "ventas": 250 },
      { "mes": "Mar", "ventas": 175 }, { "mes": "Abr", "ventas": 300 }
    ],
    "productos": [ { "producto": "Tornillos" }, { "producto": "Tuercas" }, { "producto": "Rondanas" } ]
  }
}
```
Resultado: los meses quedan en **B1:E1**, `Total` se recorre a **F1** con `=SUM(B2:E2)`; los productos ocupan **A5:A7** y el
conteo baja a **B9** con `=COUNTA(A5:A7)`.

### Qué se ajusta automáticamente
- **Celdas, estilos, alto de filas / ancho de columnas, combinaciones, filas/columnas ocultas y niveles de esquema**: se copian
  en cada repetición; lo que está después de la zona se desplaza.
- **Fórmulas**: dentro de la zona, las referencias relativas a la propia zona avanzan con cada copia (`=B5*C5` → `=B6*C6`…);
  las referencias con `$` a la zona se quedan fijas; las que apuntan después de la zona se desplazan; y **los rangos que cubren
  la zona completa crecen** para abarcar todas las copias (`=SUMA(C5:C7)` → `=SUMA(C5:C13)`). Así los totales funcionan solos.
  Las referencias a otras hojas (`Hoja2!B5`) no se modifican.
- **Imágenes flotantes, formas, paneles inmovilizados y área de impresión** que están después de la zona se desplazan.
- **Datos**: las variables de la zona leen de cada elemento; si un elemento no trae un valor se usa el de la hoja
  (útil para una moneda o una tasa común).
- **Sin elementos** (arreglo vacío o ausente): la zona queda **una vez**, con sus variables vacías, y nada se desplaza
  (igual que la antigua «zona de filas»). Si la clave existe pero no es un arreglo se genera igual y se reporta una advertencia.

### Formato en la plantilla (JSON)
```json
"repeatZones": [
  { "id": "meses", "direction": "cols", "start": 2, "end": 2, "arrayVariable": "meses" },
  { "id": "prod",  "direction": "rows", "start": 5, "end": 5, "arrayVariable": "productos" }
]
```
`start`/`end` son números de fila o de columna (1 = A); para columnas también se aceptan letras (`"start": "B"`).
**Compatibilidad:** las plantillas antiguas con `rowTemplate: { startRow, endRow, arrayVariable }` siguen funcionando igual (se
leen como una zona de filas). Si la hoja tiene `repeatZones`, esa lista manda; el editor guarda además `rowTemplate` como espejo
de la primera zona de filas para herramientas anteriores.

### En la automatización
- `GET /api/templates/:id/schema` incluye por hoja `repeatZones` (con `label`, p. ej. «columnas B–D») y en cada variable
  `zone: { arrayVariable, direction }` (o `null`). `inRowTemplate` sigue presente y vale `true` para cualquier variable de zona.
- `GET /example` trae un arreglo de 2 elementos por zona.
- `POST /validate` revisa cada elemento con rutas como `Ventas.meses[1].ventas` y avisa (`code: "zone"`) si hay zonas
  superpuestas, inválidas o cruzadas. Si mandas una variable de zona a nivel de hoja, la advertencia indica dónde va.
- Al generar, las advertencias de zonas se devuelven en el encabezado `X-Render-Warnings` (JSON codificado con
  `encodeURIComponent`) y, en modo base64, en el campo `warnings`.

---

## Automatiza en 3 pasos
1. **Abre la plantilla y pulsa «Automatizar»** (barra superior del editor, botón de la tarjeta o «Automatización» en el
   inicio). Verás el ID, las variables, un cuerpo de ejemplo editable con validación y fragmentos listos para copiar.
2. **Obtén el cuerpo de ejemplo** con `GET /api/templates/<id>/example` y cambia los valores por los tuyos.
3. **Genera el Excel** con `POST /api/render/<id>` enviando ese JSON. La respuesta es el archivo `.xlsx`.

En los ejemplos usamos `BASE=http://localhost:3001` y la plantilla `fixture-factura` (hoja `Factura` con zona de filas `items`).

## Ejemplos por herramienta

### cURL (bash)
```bash
BASE=http://localhost:3001
# Ver el cuerpo de ejemplo
curl -s "$BASE/api/templates/fixture-factura/example"

# Generar y guardar el archivo
curl -X POST "$BASE/api/render/fixture-factura?strict=1" \
  -H "Content-Type: application/json" \
  -d '{"Factura":{"cliente":"Ana López","total":1160,"items":[{"descripcion":"Servicio","cantidad":2,"precio":500}]}}' \
  --fail-with-body -o Factura.xlsx
```
Con `-D -` verás también los encabezados `Content-Disposition` (nombre sugerido) y `X-Render-Ms` (milisegundos que tardó).

### PowerShell (Windows PowerShell 5.1 o PowerShell 7+)
```powershell
$base = 'http://localhost:3001'
$body = @{ Factura = @{ cliente = 'Ana López'; total = 1160
           items = @(@{ descripcion = 'Servicio'; cantidad = 2; precio = 500 }) } } | ConvertTo-Json -Depth 10
$utf8 = [Text.Encoding]::UTF8.GetBytes($body)   # envía acentos correctamente también en PowerShell 5.1

# 1) Descargar el archivo directamente
Invoke-WebRequest -Method Post -Uri "$base/api/render/fixture-factura" `
  -ContentType 'application/json; charset=utf-8' -Body $utf8 -OutFile 'Factura.xlsx'

# 2) Modo base64 (útil si después quieres adjuntarlo o subirlo a otro sistema)
$r = Invoke-RestMethod -Method Post -Uri "$base/api/render/fixture-factura?format=base64" `
  -ContentType 'application/json; charset=utf-8' -Body $utf8
[IO.File]::WriteAllBytes((Join-Path $PWD $r.fileName), [Convert]::FromBase64String($r.base64))
```

### Python (requests)
```python
import requests

BASE = "http://localhost:3001"
body = {"Factura": {"cliente": "Ana López", "total": 1160,
                    "items": [{"descripcion": "Servicio", "cantidad": 2, "precio": 500}]}}

r = requests.post(f"{BASE}/api/render/fixture-factura", params={"strict": 1}, json=body, timeout=60)
if r.status_code == 422:                      # datos que no cumplen el esquema
    for e in r.json()["errors"]:
        print(e["path"], "→", e["message"])
r.raise_for_status()
with open("Factura.xlsx", "wb") as f:
    f.write(r.content)
```

### Node.js (fetch, Node 18+)
Guárdalo como `generar.mjs` y ejecútalo con `node generar.mjs`:
```js
import { writeFile } from 'node:fs/promises';

const BASE = 'http://localhost:3001';
const body = { Factura: { cliente: 'Ana López', total: 1160, items: [{ descripcion: 'Servicio', cantidad: 2, precio: 500 }] } };

const res = await fetch(`${BASE}/api/render/fixture-factura?strict=1`, {
  method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body),
});
if (!res.ok) {
  const err = await res.json();
  throw new Error(`${res.status} ${err.error}\n${JSON.stringify(err.errors ?? [], null, 2)}`);
}
await writeFile('Factura.xlsx', Buffer.from(await res.arrayBuffer()));
console.log('Generado en', res.headers.get('x-render-ms'), 'ms');
```

### C# (.NET 6+, HttpClient)
```csharp
using System.Net.Http.Json;

var http = new HttpClient { BaseAddress = new Uri("http://localhost:3001") };
var body = new {
    Factura = new {
        cliente = "Ana López", total = 1160,
        items = new[] { new { descripcion = "Servicio", cantidad = 2, precio = 500 } }
    }
};
var res = await http.PostAsJsonAsync("/api/render/fixture-factura?strict=1", body);
if (!res.IsSuccessStatusCode)
    throw new Exception($"{(int)res.StatusCode}: {await res.Content.ReadAsStringAsync()}");
await File.WriteAllBytesAsync("Factura.xlsx", await res.Content.ReadAsByteArrayAsync());
```

### Power Automate
1. Acción **HTTP**: Método `POST`, URI `https://tu-servidor/api/render/fixture-factura?format=base64`,
   encabezado `Content-Type: application/json` y en Cuerpo tu JSON (puedes usar contenido dinámico de pasos anteriores).
2. Acción **Crear archivo** (OneDrive para la Empresa o SharePoint):
   - Nombre de archivo: `@{body('HTTP')?['fileName']}`
   - Contenido del archivo: `@{base64ToBinary(body('HTTP')?['base64'])}`
3. (Opcional) Para enviarlo por correo, en **Enviar un correo electrónico (V2)** agrega un adjunto con el mismo nombre y
   contenido `base64ToBinary(...)`.

> El servidor debe ser accesible desde la nube de Microsoft (URL pública protegida) o usar la **puerta de enlace de datos local**.

### n8n
Nodo **HTTP Request**:
- *Method* `POST`, *URL* `http://servidor:3001/api/render/fixture-factura`
- *Send Body* → *JSON* con tu cuerpo (expresiones como `{{ $json.cliente }}` funcionan dentro del JSON).
- *Options → Response → Response Format* → **File** (queda en la propiedad binaria `data`).

Conéctalo a *Write Binary File*, Google Drive, Gmail, etc. Si prefieres JSON, agrega `?format=base64` y usa
*Response Format: JSON*. El asistente **Automatizar** incluye un nodo listo para pegar (Ctrl+V) en el lienzo de n8n.

### Enlace directo (GET)
Para hojas **no repetibles**, cada parámetro de la URL que coincide con una variable se usa como valor
(los tipos `number` y `boolean` se convierten; para booleanos se aceptan `true/false`, `1/0`, `sí/no`):
```
http://localhost:3001/api/render/fixture-factura?cliente=Ana%20L%C3%B3pez&total=1160&filename=Factura%20Ana
```
- Parámetros reservados (nunca se toman como variables): `format`, `filename`, `strict`, `download`.
- Los parámetros que no son variables se ignoran.
- Las variables de zonas de filas y de hojas repetibles no se pueden enviar por enlace: usa `POST`.
- Ideal para un botón «Descargar» en una intranet, un correo o un hipervínculo en otra hoja de cálculo.

## Lotes → ZIP
```bash
curl -X POST "$BASE/api/render/fixture-factura/batch?filename=Facturas%20noviembre" \
  -H "Content-Type: application/json" \
  -d '{"fileName":"Factura {{cliente}}","items":[{"Factura":{"cliente":"Ana","total":1}},{"Factura":{"cliente":"Luis","total":2}}]}' \
  --fail-with-body -o facturas.zip
```
- `items`: de 1 a 200 cuerpos (cada uno con la misma forma que en `/api/render/<id>`). También puedes enviar el arreglo directo.
- `fileName`: patrón para el nombre de cada archivo dentro del ZIP; `{{variable}}` toma el primer valor con ese nombre que
  encuentre en el elemento. Si se omite, se usa «Nombre de plantilla 1», «… 2». Los nombres repetidos se numeran: `Factura Ana (2).xlsx`.
- `?filename=` define el nombre del ZIP; `?strict=1` valida cada elemento (los errores incluyen `item`, índice base 0).
- `?format=base64` → `{ "fileName", "mimeType": "application/zip", "size", "base64", "count" }`.
- El lote cuenta como una solicitud para el límite de tasa, así que es la forma recomendada de generar muchos archivos.

## Validación y modo estricto
**Validar sin generar:** `POST /api/templates/<id>/validate` con el cuerpo. Siempre responde 200; revisa `valid`:
```json
{
  "valid": false,
  "errors": [
    { "sheet": "Factura", "path": "Factura.items[1].cantidad", "variable": "cantidad", "code": "type",
      "message": "\"Factura.items[1].cantidad\" debe ser un número (p. ej. 12.5); se recibió string." },
    { "sheet": "Factura", "path": "Factura.total", "variable": "total", "code": "missing",
      "message": "Falta la variable obligatoria \"total\" (Factura.total)." }
  ],
  "warnings": [
    { "sheet": null, "path": "Facturas", "code": "unknown",
      "message": "La hoja \"Facturas\" no existe en la plantilla; se ignorará." }
  ]
}
```
| Código | Significado |
|---|---|
| `missing` | Variable obligatoria ausente o vacía (en zonas de filas se revisa cada elemento: `Factura.items[1].cantidad`). |
| `type` | El valor no corresponde al tipo (número, booleano, fecha, imagen o texto). |
| `shape` | Estructura incorrecta: objeto donde va arreglo, elemento que no es objeto, cuerpo que no es JSON de objeto… |
| `unknown` | *Advertencia*: hoja o variable que la plantilla no tiene (se ignora). Suele indicar un error de escritura. |

**Modo estricto:** agrega `?strict=1` a `/api/render/<id>` (o al lote). Si hay errores responde **422** y no genera nada:
```json
{ "error": "Los datos no cumplen el esquema de la plantilla", "code": "VALIDATION",
  "errors": [ { "sheet": "Factura", "path": "Factura.total", "variable": "total", "code": "missing", "message": "…" } ],
  "warnings": [] }
```
Sin `strict` el archivo se genera siempre: los valores faltantes quedan vacíos y los números inválidos como `0`.
Recomendación: usa `strict=1` en integraciones automáticas para enterarte de inmediato si cambió la plantilla.

## Modo base64
Agrega `?format=base64` a `POST /api/render/<id>`, `GET /api/render/<id>` o al lote:
```json
{ "fileName": "Factura.xlsx",
  "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "size": 6789, "base64": "UEsDBBQABgAIAAAAIQ…" }
```
Útil en Power Automate, Logic Apps, webhooks o cualquier cliente que no maneje bien respuestas binarias.

## OpenAPI y Postman
`GET /api/openapi.json` devuelve una especificación **OpenAPI 3.1** en español con todas las rutas, parámetros, ejemplos y
respuestas de error (400, 404, 409, 422, 429). Etiquetas: *Plantillas*, *Generación*, *Automatización*, *Tiempo real* y *Uso*.
- **Postman:** *Import → Link* → pega `http://localhost:3001/api/openapi.json` (o descarga el archivo e impórtalo).
  Se crea una colección con todas las solicitudes y ejemplos.
- También sirve para Swagger UI, Insomnia, Bruno o generadores de clientes (`openapi-generator`).

## Referencia de rutas
| Método y ruta | Respuesta |
|---|---|
| `GET /api/health` | `{ ok, version, uptime }` |
| `GET /api/templates` | Lista de plantillas: `{ id, name, sheetCount, updatedAt, createdAt, revision, updatedBy, thumbnail, sheetNames, variables }[]`, más reciente primero |
| `GET /api/templates/:id` | Plantilla completa (incluye `revision`) |
| `POST /api/templates[?force=1]` | Crear/guardar → `{ success, id, revision, updatedAt }` · **409** `{ error, code: "CONFLICT", current: { revision, updatedAt, updatedBy } }` si alguien guardó antes |
| `PUT /api/templates/:id/thumbnail` | Body `{ thumbnail: "data:image/png;base64,…" }` (PNG/JPEG/WebP ≤ 2 MB) → `{ success }` |
| `POST /api/templates/:id/duplicate` | `{ success, id, name }` (nombre + « (copia)») |
| `DELETE /api/templates/:id` | `{ success }` |
| `GET /api/templates/:id/schema` | `{ id, name, revision, schema: [{ name, isRepeatable, repeatVariable, rowTemplate, variables }] }` |
| `GET /api/templates/:id/example` | Cuerpo de ejemplo listo para generar |
| `POST /api/templates/:id/validate` | `{ valid, errors, warnings }` (siempre 200) |
| `POST /api/import` | Multipart con campo `file` (.xlsx) → plantilla en JSON (se guarda al enviarla a `POST /api/templates`) |
| `POST /api/render/:id` | `.xlsx` · `?format=base64` · `?filename=` · `?strict=1` (422) · encabezado `X-Render-Ms` |
| `GET /api/render/:id?variable=valor…` | Igual, con variables en la URL |
| `POST /api/render/:id/batch` | ZIP con un `.xlsx` por elemento · `?format=base64` · `?filename=` · `?strict=1` |
| `GET /api/openapi.json` | Especificación OpenAPI 3.1 |
| `GET /api/live/stream?clientId=&userId=&name=&color=` | Flujo SSE en tiempo real |
| `GET /api/live/presence` · `POST /api/live/presence` · `POST /api/live/leave` | Presencia: consultar, actualizar (204) y salir (204) |
| `GET /api/usage` · `GET /api/usage/events?limit=100&templateId=` · `GET /api/usage/:id` | Métricas de uso |

Encabezados opcionales para identificar a quien llama: `X-Client-Id`, `X-User-Id` y `X-Client-Name`
(codificado con `encodeURIComponent`). Si los envías, los cambios y las generaciones aparecen con tu nombre en *Actividad*.
Las generaciones con `X-Client-Id` cuentan como origen `ui`; sin él, como `api`; las de `GET` como `link`.

---

## Colaboración en tiempo real
- **Presencia:** la primera vez se te pide un nombre y un color (puedes omitirlo y serás `Invitado-XXXX`; también puedes abrir
  la app con `?user=TuNombre`). En la barra superior ves quién está conectado y en cada tarjeta de inicio quién ve o edita
  esa plantilla. Haz clic en tu avatar para cambiar tu nombre o color.
- **Cursores remotos:** las celdas o rangos seleccionados por otras personas en la misma hoja se marcan con su color y nombre.
- **Sincronización casi en vivo:** si alguien más tiene abierta la misma plantilla y tú haces cambios, se guardan solos tras
  ~3 s sin actividad. Si tú no tienes cambios pendientes, la plantilla se recarga sola cuando otra persona guarda
  («Actualizado por X»); si los tienes, aparece un aviso «X guardó cambios» con **Recargar** / **Ignorar**.
- **Conflictos:** cada guardado incrementa `revision`. Si intentas guardar sobre una versión vieja, el servidor responde **409**
  y la app te ofrece:
  - **Recargar del servidor** — descarta tus cambios y toma la versión actual;
  - **Sobrescribir** — guarda tu versión encima (equivale a `?force=1`);
  - **Guardar como copia** — conserva ambas versiones.
  El guardado automático nunca reintenta en bucle ante un conflicto.
- **Por API:** si guardas plantillas desde un script, envía la `revision` que leíste; ante un 409 vuelve a leer con
  `GET /api/templates/:id`, aplica tus cambios y reintenta (o usa `?force=1` si sabes que quieres sobrescribir).
- **Cómo funciona:** *Server-Sent Events* en `/api/live/stream` con reconexión automática. Eventos: `hello`, `presence`,
  `template:saved`, `template:created`, `template:deleted`, `usage` y `ping` (cada 25 s). Una persona se marca como
  inactiva tras 2 min sin actividad.

## Actividad y métricas de uso
La pestaña **Actividad** (o el botón *Actividad* de la barra superior) muestra en vivo:
- archivos generados hoy y en la última hora, errores, vistas, guardados y bytes generados;
- gráficas de 24 horas y 7 días por plantilla, y detalle de 48 horas / 30 días;
- el origen de cada generación (`ui`, `api`, `link`), los clientes que más generan y un historial de eventos.

Cada tarjeta del inicio muestra además un resumen de uso. Los datos vienen de `GET /api/usage` (resumen),
`GET /api/usage/:id` (detalle) y `GET /api/usage/events?limit=100&templateId=` (eventos, más reciente primero).
Las vistas se cuentan una vez cada 10 minutos por persona y plantilla. Se conservan 8 días de datos por hora, 120 días por
día y los últimos 1000 eventos. Las horas usan la zona horaria del servidor.

---

## Seguridad
- **Sin autenticación incorporada.** Cualquiera que alcance el puerto puede leer, modificar y borrar plantillas. Para usarlo
  fuera de tu equipo publícalo **detrás de un proxy inverso con autenticación** (IIS/ARR, Nginx, Azure App Proxy, el inicio de
  sesión de tu portal o una VPN), o limita el acceso con `HOST=127.0.0.1`. Los nombres de usuario de la presencia son
  informativos, no una identidad verificada.
- **CSP** estricta con Helmet: sin scripts en atributos HTML, imágenes solo de `self`/`data:`/`blob:`, conexiones solo al
  mismo origen. Todo texto que viene de otras personas (nombres, plantillas, hojas) se escapa antes de mostrarse.
- **Límites de tasa:** 1200 solicitudes/min por IP en general (excepto tiempo real) y 120/min en `/api/render`; al
  superarlos responde **429** `{ error, code: "RATE_LIMIT" }` con los encabezados estándar `RateLimit` y `RateLimit-Policy`.
- **CORS:** desactivado salvo para los orígenes listados en `CORS_ORIGINS`.
- **Entradas:** los IDs se validan contra `^[A-Za-z0-9_-]{1,64}$` (sin recorrido de rutas); los nombres de archivo se sanean
  (sin `/ \ : * ? " < > |`, caracteres de control ni nombres reservados de Windows, ≤ 120 caracteres; los acentos se conservan);
  la validación ignora claves peligrosas como `__proto__`. Cuerpos JSON de hasta 50 MB.

## Solución de problemas
| Síntoma | Causa y solución |
|---|---|
| `EADDRINUSE` al iniciar | El puerto está ocupado. Usa otro (`PORT=3002 npm start` o `node main.js 3002`) o cierra el proceso: Windows `netstat -ano \| findstr :3001` y `taskkill /F /PID <pid>`; Linux/macOS `lsof -i :3001` y `kill <pid>`. |
| La presencia no se actualiza o aparece «Sin conexión» detrás de un proxy | El proxy está acumulando (bufferizando) el flujo SSE. El servidor ya envía `X-Accel-Buffering: no` y `Cache-Control: no-transform`; en Nginx agrega además `proxy_buffering off;`, `proxy_http_version 1.1;` y `proxy_read_timeout 120s;` para `/api/live/`. En IIS/ARR desactiva el *response buffering* y la compresión dinámica para esa ruta. En Cloudflare u otros CDN, excluye `/api/live/` de la caché. |
| **409** al guardar | Otra persona guardó antes que tú. En la app elige Recargar, Sobrescribir o Guardar como copia. Por API, vuelve a leer la plantilla y envía su `revision` actual (o `?force=1`). |
| **422** al generar | Usaste `strict=1` y los datos no cumplen el esquema. Revisa `errors[].path` y `message`, o prueba primero en `/api/templates/<id>/validate`. |
| **429** | Superaste el límite de tasa. Espera un minuto, espacia las llamadas o usa `/batch` para muchos archivos. |
| **400** «ID inválido» | El ID tiene caracteres no permitidos (espacios, acentos, `/`). Cópialo desde el asistente *Automatizar*. |
| Celdas vacías en el Excel | El nombre de la hoja o de la variable no coincide (mayúsculas, acentos, espacios) o pusiste una variable de zona de filas fuera de `items`. Revisa las advertencias `unknown` de `/validate`. |
| Un booleano sale VERDADERO aunque mandé `"false"` | Envía `false` de JSON (sin comillas). |
| La imagen no aparece | Debe ser un data URL `data:image/png;base64,…` (también `jpeg` o `gif`). Las URLs `http(s)` no se descargan. |
| Acentos rotos al llamar desde PowerShell 5.1 | Envía el cuerpo como bytes UTF-8 (ver ejemplo) e indica `charset=utf-8`. |

---

## Para desarrolladores
```
main.js                 Servidor Express: middleware, rutas y cableado
Library/store.js        Persistencia de plantillas (revisiones, conflictos, escritura atómica)
Library/render.js       Plantilla + datos → .xlsx (ExcelJS + formas)
Library/repeat-zones.js Zonas repetibles: normalización/validación, reescritura de referencias en fórmulas y expansión
Library/importer.js     .xlsx → plantilla
Library/live.js         Tiempo real (SSE, presencia)
Library/usage.js        Métricas de uso
Library/automation.js   Esquema, ejemplo, validación, enlaces GET, nombres de archivo y OpenAPI (funciones puras)
Public/                 Interfaz (sin build): index.html + js/ y css/ por módulo
test/                   Pruebas con node:test
```
`Library/automation.js` exporta funciones puras sin E/S, útiles también desde otros scripts de Node:
```js
const A = require('./Library/automation');
A.buildSchema(template);        // { id, name, revision, sheets: [{ name, isRepeatable, rowTemplate, repeatZones, zoneIssues, variables: [{ name, type, required, hint, defaultValue, addr, inRowTemplate, zone, implicit }] }] }
A.buildExample(template);       // cuerpo de ejemplo (hojas repetibles y cada zona repetible con 2 elementos)
A.validateBody(template, body); // { valid, errors, warnings } — nunca lanza
A.bodyFromQuery(template, q);   // parámetros de URL → cuerpo
A.safeFileName('Factura: Ana'); // 'Factura_ Ana.xlsx'
A.fillPattern('Factura {{cliente}}', body);
A.buildOpenApi({ baseUrl: 'http://localhost:3001' });
```
`renderWorkbook(template, body, { warnings: [] })` (Library/render.js) llena `warnings` con los avisos de zonas.
`Library/repeat-zones.js` exporta `normalizeZones`, `expandSheet`, `mapFormulaRefs` y `adjustFormulaCols` (contraparte de
`adjustFormulaRows`, ambas también disponibles desde `Library/excel-utils.js`).

Las pruebas nunca usan la carpeta `data/` real: define `ETB_DATA_DIR` con una carpeta temporal. Para correr un solo
archivo: `node --test test/automation.test.js`.

---

## Changelog v12
- **Tiempo real:** presencia de usuarios, cursores remotos, indicador de conexión, recarga automática y aviso de cambios
  remotos, guardado rápido colaborativo.
- **Control de versiones:** `revision`, `updatedBy` y `createdAt` en cada plantilla; detección de conflictos (409) con
  Recargar / Sobrescribir / Guardar como copia.
- **Uso de cada Excel:** registro de generaciones, vistas, guardados, importaciones y errores; panel *Actividad* con gráficas
  y eventos en vivo; API `/api/usage`.
- **Automatización:** asistente «Automatizar» con ejemplos para cURL, PowerShell, Python, Node.js, C#, Power Automate y n8n;
  variables implícitas y zonas de filas en el esquema de automatización; `/example`, `/validate`, modo estricto (422),
  modo base64, enlaces GET, lotes ZIP con patrón de nombre, encabezado `X-Render-Ms` y OpenAPI 3.1.
- **Filas y columnas repetibles:** varias zonas por hoja en ambas direcciones (`repeatZones`), con totales que crecen solos,
  desplazamiento de combinaciones, anchos/altos, imágenes y formas; editor con botón **Repetir**, menú contextual, panel
  lateral, banda visual en la cuadrícula y deshacer; compatibilidad total con `rowTemplate`.
- **Interfaz:** nuevos íconos, animaciones con desenfoque (respetan *reducir movimiento*), tarjetas con actividad en vivo.
- **Plataforma:** almacenamiento con caché y escritura atómica, servidor dividido en módulos, compresión (excepto SSE),
  `crypto.randomUUID` en lugar de `uuid`, `multer` 2, `npm test` con `node:test`, cierre ordenado del servidor.

### Correcciones posteriores
- **Importación:** los nombres de archivo con acentos (p. ej. `Cotización.xlsx`) ya no se convierten en `CotizaciÃ³n`; las
  celdas ocultas de una combinación ya no copian el valor de la celda principal (evitaba falsos cruces entre zonas y
  variables duplicadas).
- **Editor:** el cálculo de fórmulas ya no se congela en hojas con fórmulas encadenadas (los manejadores del analizador se
  acumulaban en cada evaluación); se admite `IFERROR`/`SI.ERROR`; las fórmulas se calculan desde que se abre la plantilla y
  las que dependen de una `{{variable}}` muestran la fórmula en lugar de `#VALUE!`; las fechas respetan formatos simples
  como `dd/mm/yyyy`.
- **Variables:** el tipo de texto se guarda como `text` (antes `string`), igual que en la API y el esquema; las plantillas
  existentes con `string` siguen funcionando.
