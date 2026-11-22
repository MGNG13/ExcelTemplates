<div align="center">

# Excel Template Builder

**Diseña plantillas de Excel en el navegador y genera archivos `.xlsx` llenos de datos desde cualquier sistema.**

[![Node.js](https://img.shields.io/badge/Node.js-%E2%89%A518-339933?logo=node.js&logoColor=white)](https://nodejs.org/)
[![Express](https://img.shields.io/badge/Express-4-000000?logo=express&logoColor=white)](https://expressjs.com/)
[![ExcelJS](https://img.shields.io/badge/ExcelJS-xlsx-217346?logo=microsoftexcel&logoColor=white)](https://github.com/exceljs/exceljs)
[![OpenAPI](https://img.shields.io/badge/OpenAPI-3.1-6BA539?logo=openapiinitiative&logoColor=white)](#api)
[![Tests](https://img.shields.io/badge/tests-node%3Atest-5FA04E)](#desarrollo)
[![Licencia: MIT](https://img.shields.io/badge/licencia-MIT-blue.svg)](LICENSE)

</div>

Excel Template Builder es un editor visual de plantillas de Excel con servidor Node.js. Importas un libro **con sus
estilos, combinaciones, fórmulas, imágenes y formas**, marcas los datos que cambian con `{{variables}}`, defines **filas
y columnas que se repiten** por cada elemento de una lista y, a partir de ahí, generas reportes, cotizaciones o pedidos
con un `POST` de JSON desde cURL, Python, Power Automate, n8n o cualquier otra herramienta. Incluye colaboración en
tiempo real, un panel de actividad con métricas de uso y un asistente que escribe el código de integración por ti.

<p align="center">
  <img src="Docs/screenshots/07-zones-grid.png" alt="Editor con una plantilla de reporte de ventas: zona de columnas «meses» y zona de filas «productos»" width="900">
</p>

> Las capturas usan datos inventados de la empresa ficticia *Comercial Aurora S.A. de C.V.*

---

## Contenido
- [Características](#características)
- [Recorrido por la aplicación](#recorrido-por-la-aplicación)
- [Inicio rápido](#inicio-rápido)
- [Automatiza en 3 pasos](#automatiza-en-3-pasos)
- [API](#api)
- [Arquitectura](#arquitectura)
- [Tecnologías](#tecnologías)
- [Desarrollo](#desarrollo)
- [Limitaciones y hoja de ruta](#limitaciones-y-hoja-de-ruta)
- [Licencia](#licencia)

📘 La referencia técnica completa (todas las opciones, códigos de error, seguridad, solución de problemas) está en
**[Docs/GUIA.md](Docs/GUIA.md)**.

---

## Características

| | |
|---|---|
| 📥 **Importación fiel de Excel** | Estilos, fuentes, rellenos, bordes, formatos numéricos, combinaciones, fórmulas, imágenes, formas, filas/columnas ocultas, paneles inmovilizados y configuración de página. |
| ✏️ **Editor tipo hoja de cálculo** | Cinta con formato, bordes, combinar, ajustar texto, formas e imágenes; varias hojas y pestañas de documentos; deshacer/rehacer; zoom. |
| `{{ }}` **Variables tipadas** | `text`, `number`, `boolean`, `date` e `image`, con obligatoriedad, pista y valor por defecto. También variables implícitas dentro de cualquier texto o fórmula. |
| 🔁 **Filas y columnas repetibles** | Zonas que se copian por cada elemento de un arreglo, en ambas direcciones y varias por hoja; los totales (`=SUMA(...)`), combinaciones, anchos y objetos se ajustan solos. |
| 📑 **Hojas repetibles** | Una hoja por elemento (`Pedido 1`, `Pedido 2`, …). |
| 👥 **Colaboración en tiempo real** | Presencia, cursores remotos, sincronización casi en vivo, aviso de cambios y resolución de conflictos por revisión. |
| 📊 **Actividad** | KPIs en vivo, uso por plantilla con gráficas, origen de cada generación (app, API, enlace), clientes principales e historial. |
| ⚡ **Automatizar** | Esquema, cuerpo de ejemplo editable con validación y fragmentos listos para cURL, PowerShell, Python, Node.js, C#, Power Automate, n8n y enlace directo. |
| 📦 **Generación por API** | `.xlsx` binario o base64, modo estricto (422), lotes en ZIP, enlaces GET y especificación **OpenAPI 3.1**. |

---

## Recorrido por la aplicación

### Inicio y presencia
Las plantillas aparecen como tarjetas con miniatura, número de hojas y variables, generaciones del día con una mini
gráfica y **quién las está viendo o editando en este momento**. La barra superior muestra a las personas conectadas.

<img src="Docs/screenshots/01-home.png" alt="Pantalla de inicio con tarjetas de plantillas, usuarios conectados e insignia «2 viendo»" width="900">

La primera vez se pide un nombre y un color; así el equipo sabe quién hace qué (también se puede abrir con `?user=Nombre`).

<img src="Docs/screenshots/02-welcome.png" alt="Ventana de bienvenida para elegir nombre y color" width="700">

### Importar un Excel con estilos
**Importar Excel** convierte cualquier `.xlsx` en una plantilla editable conservando su diseño: banda de título combinada,
colores, bordes, formatos de moneda, fórmulas y el logotipo incrustado.

<img src="Docs/screenshots/03-import-editor.png" alt="Reporte de ventas importado con estilos, logotipo y fórmulas" width="900">

### Variables
Selecciona una celda y pulsa **Variable** para declarar su nombre, tipo, si es obligatoria, una pista y un valor por
defecto. Si la celda está dentro de una zona repetible, la ventana lo indica: su valor vendrá de cada elemento del arreglo.

<img src="Docs/screenshots/04-variable-modal.png" alt="Ventana para definir la variable «cantidad» de tipo número dentro de la zona «productos»" width="800">

### Filas y columnas repetibles
Selecciona filas o columnas, pulsa **Repetir** y escribe el nombre del arreglo. La ventana explica qué se copiará y
muestra **el JSON exacto** que habrá que enviar.

| Columnas (`meses`) | Filas (`productos`) |
|---|---|
| <img src="Docs/screenshots/05-zone-cols-modal.png" alt="Ventana Repetir columnas para la columna C" width="440"> | <img src="Docs/screenshots/06-zone-rows-modal.png" alt="Ventana Repetir filas para la fila 15" width="440"> |

En la cuadrícula cada zona se dibuja como una banda punteada (naranja para filas, morada para columnas) y aparece en el
panel lateral **Repetición**, desde donde se edita o elimina. Las variables de una zona muestran su arreglo (`meses[]`).
Insertar o eliminar filas/columnas desplaza las zonas, y deshacer/rehacer también las incluye.

<img src="Docs/screenshots/07-zones-grid.png" alt="Cuadrícula con la zona de columnas «meses» y la zona de filas «productos», panel lateral Repetición y lista de variables" width="900">

### El resultado: el mismo diseño, con todos los datos
Con 4 meses y 5 productos, la columna **C** se convierte en **C:F** (Enero–Abril) y `Total` se desplaza a **G** con
`=SUM(C9:F9)`; la fila **15** se convierte en **15:19** y el subtotal pasa a `=SUM(I15:I19)`. Las combinaciones de título
crecen, los anchos se copian y el logotipo se conserva. La hoja repetible `Pedido` genera `Pedido 1`, `Pedido 2` y `Pedido 3`.

| Plantilla | Excel generado |
|---|---|
| <img src="Docs/screenshots/07-zones-grid.png" alt="Plantilla" width="440"> | <img src="Docs/screenshots/20-result-xlsx.png" alt="Excel generado con 4 meses y 5 productos" width="440"> |

<details>
<summary>Ver el Excel generado a tamaño completo y una de las hojas «Pedido»</summary>

<img src="Docs/screenshots/20-result-xlsx.png" alt="Excel generado abierto en la aplicación: meses Enero a Abril, 5 productos, totales calculados" width="900">

<img src="Docs/screenshots/21-result-pedido.png" alt="Hoja «Pedido 2» generada a partir de la hoja repetible" width="900">

</details>

### Formas, imágenes y hojas repetibles
Dibuja formas (rectángulos, flechas, conectores, texto…) y ajusta relleno, línea y texto desde el panel lateral; se
exportan como formas nativas de Excel. En **Configuración de hoja** se marca una hoja como repetible y se administran
sus zonas.

| Formas | Hoja repetible |
|---|---|
| <img src="Docs/screenshots/08-shapes.png" alt="Forma «PAGADO» con su panel de propiedades" width="440"> | <img src="Docs/screenshots/09-sheet-settings.png" alt="Configuración de hoja con «Hoja repetible» activado" width="440"> |

### Colaboración en tiempo real
Cada persona ve los avatares de quienes tienen abierta la plantilla y sus selecciones con nombre y color. Los cambios se
sincronizan casi en vivo; si alguien guarda mientras tú tienes cambios pendientes aparece un aviso, y si intentas guardar
sobre una versión vieja el servidor responde **409** y puedes **recargar**, **sobrescribir** o **guardar como copia**.

<img src="Docs/screenshots/10-collab-presence.png" alt="Avatar de Beto Ramírez en la barra y su selección C9:C11 marcada en azul" width="900">

| Aviso de cambios remotos | Conflicto de versiones |
|---|---|
| <img src="Docs/screenshots/11-remote-banner.png" alt="Aviso «Ana López guardó cambios» con Recargar e Ignorar" width="440"> | <img src="Docs/screenshots/12-conflict-modal.png" alt="Ventana de conflicto: Recargar, Sobrescribir o Guardar como copia" width="440"> |

### Actividad
Un tablero en vivo con archivos generados hoy, errores, personas conectadas, plantillas abiertas y generaciones de la
última hora; una tabla de uso por plantilla con gráficas de 24 h y un historial de eventos. Al hacer clic en una
plantilla se abre su detalle: 48 horas, 30 días, origen de las generaciones y clientes principales.

<img src="Docs/screenshots/13-dashboard.png" alt="Panel Actividad con KPIs, uso por plantilla, conectados y actividad reciente" width="900">

<img src="Docs/screenshots/14-dashboard-drawer.png" alt="Detalle de uso del Reporte de Ventas 2026" width="900">

### Automatizar
El asistente **Automatizar** reúne todo lo necesario para integrar una plantilla: endpoint, variables por hoja (incluidas
las de cada zona), opciones de respuesta, nombre de archivo con variables y un cuerpo JSON de prueba que puedes
**validar** o **generar** ahí mismo.

<img src="Docs/screenshots/15-automation.png" alt="Asistente Automatizar: endpoint, variables y cuerpo JSON de prueba" width="900">

La validación indica la ruta exacta de cada error (`Ventas.meses[0].ventas`) y advierte de hojas o variables desconocidas.
Debajo, el código listo para copiar en 8 variantes.

| Validación | Python | Enlace directo |
|---|---|---|
| <img src="Docs/screenshots/16-automation-validate.png" alt="Siete errores de validación con su ruta" width="290"> | <img src="Docs/screenshots/17-automation-python.png" alt="Fragmento de Python con requests" width="290"> | <img src="Docs/screenshots/18-automation-link.png" alt="URL GET para descargar el Excel" width="290"> |

<details>
<summary>Ver la guía para Power Automate</summary>

<img src="Docs/screenshots/19-automation-powerautomate.png" alt="Pasos para Power Automate" width="900">

</details>

---

## Inicio rápido

Requisitos: **Node.js 18+** (probado con Node 25) y un navegador moderno. No hace falta tener Excel instalado.

```bash
git clone https://github.com/<tu-usuario>/excel-template-builder.git
cd excel-template-builder
npm install
npm start            # http://localhost:3001
```

| Comando | Qué hace |
|---|---|
| `npm start` | Inicia el servidor (puerto 3001 por defecto; también `node main.js 8080`). |
| `npm run dev` | Modo desarrollo con recarga automática (`node --watch`). |
| `npm test` | Pruebas automáticas con `node:test`. |
| `./scripts/install.sh` · `./scripts/run.sh [puerto]` | Instalación y arranque en Linux/macOS (localiza Node incluso con nvm). |

### Variables de entorno

| Variable | Predeterminado | Descripción |
|---|---|---|
| `PORT` | `3001` | Puerto HTTP (el argumento de línea de comandos tiene prioridad). |
| `HOST` | `0.0.0.0` | Interfaz de escucha; `127.0.0.1` para aceptar solo conexiones locales. |
| `ETB_DATA_DIR` | `./data` | Carpeta de plantillas (`templates/*.json`) y métricas (`usage/usage.json`). |
| `CORS_ORIGINS` | *(vacío)* | Orígenes permitidos para llamar a la API desde otro sitio web, separados por comas. |
| `TRUST_PROXY` | `loopback` | Valor de `trust proxy` de Express cuando se publica detrás de un proxy inverso. |

```powershell
$env:PORT = 8080; $env:ETB_DATA_DIR = 'D:\etb-datos'; npm start   # PowerShell
```

> ⚠️ **No incluye autenticación.** Para usarlo fuera de tu equipo publícalo detrás de un proxy inverso con inicio de
> sesión o una VPN, o limítalo con `HOST=127.0.0.1`. Detalles en [Seguridad](Docs/GUIA.md#seguridad).

---

## Automatiza en 3 pasos

1. Abre la plantilla y pulsa **Automatizar** para ver su ID, variables y un cuerpo de ejemplo.
2. Obtén el cuerpo de ejemplo con `GET /api/templates/<id>/example` y cambia los valores.
3. Envía el JSON a `POST /api/render/<id>`: la respuesta es el `.xlsx`.

El cuerpo se organiza **por nombre de hoja**. Cada zona repetible es un arreglo de objetos dentro de su hoja y cada hoja
repetible es un arreglo de objetos:

```json
{
  "Ventas": {
    "periodo": "Enero – Abril",
    "cliente": "Distribuidora del Bajío S.A. de C.V.",
    "fecha": "2026-04-30",
    "meses": [
      { "mes": "Enero",   "ventas": 412350.50, "pedidos": 198 },
      { "mes": "Febrero", "ventas": 468920.00, "pedidos": 221 }
    ],
    "productos": [
      { "clave": "AUR-001", "producto": "Cuaderno profesional 100 h", "cantidad": 1200, "precio": 38.5 },
      { "clave": "AUR-014", "producto": "Bolígrafo gel azul (caja 12)", "cantidad": 640, "precio": 96 }
    ]
  },
  "Pedido": [
    { "folio": "P-1041", "unidades": 300, "precio_unitario": 38.5 },
    { "folio": "P-1042", "unidades": 120, "precio_unitario": 119 }
  ]
}
```

```bash
BASE=http://localhost:3001
ID=<id-de-la-plantilla>

# Validar sin generar (siempre 200; revisa "valid", "errors" y "warnings")
curl -s -X POST "$BASE/api/templates/$ID/validate" -H "Content-Type: application/json" -d @datos.json

# Generar en modo estricto: 422 con la lista de errores si los datos no cumplen el esquema
curl -X POST "$BASE/api/render/$ID?strict=1&filename=Reporte%20Bajio" \
  -H "Content-Type: application/json" -d @datos.json \
  --fail-with-body -o Reporte.xlsx

# Muchos archivos en una sola llamada → ZIP (1 a 200 elementos)
curl -X POST "$BASE/api/render/$ID/batch" -H "Content-Type: application/json" \
  -d '{"fileName":"Reporte {{cliente}}","items":[{"Ventas":{"cliente":"Ana"}},{"Ventas":{"cliente":"Luis"}}]}' \
  -o reportes.zip
```

Agrega `?format=base64` para recibir `{ fileName, mimeType, size, base64 }` (útil en Power Automate o webhooks). Ejemplos
completos para PowerShell, Python, Node.js, C#, Power Automate, n8n y enlaces GET en la
[guía](Docs/GUIA.md#ejemplos-por-herramienta).

---

## API

| Método y ruta | Descripción |
|---|---|
| `GET /api/health` | Estado: `{ ok, version, uptime }` |
| `GET /api/templates` · `GET /api/templates/:id` | Listar plantillas · obtener una (incluye `revision`) |
| `POST /api/templates[?force=1]` | Crear o guardar; **409** si alguien guardó una revisión más reciente |
| `POST /api/templates/:id/duplicate` · `DELETE /api/templates/:id` | Duplicar · eliminar |
| `POST /api/import` | Multipart `file` (.xlsx) → plantilla JSON |
| `GET /api/templates/:id/schema` · `/example` | Esquema de variables y zonas · cuerpo de ejemplo |
| `POST /api/templates/:id/validate` | `{ valid, errors, warnings }` |
| `POST /api/render/:id` | Genera el `.xlsx` (`?strict=1`, `?format=base64`, `?filename=`) |
| `GET /api/render/:id?variable=valor` | Generación por enlace (hojas no repetibles) |
| `POST /api/render/:id/batch` | ZIP con un archivo por elemento |
| `GET /api/usage` · `/api/usage/:id` · `/api/usage/events` | Métricas de uso |
| `GET /api/live/stream` · `/api/live/presence` | Tiempo real (SSE) y presencia |
| `GET /api/openapi.json` | Especificación **OpenAPI 3.1** (importable en Postman, Insomnia o Swagger UI) |

Referencia completa de parámetros, respuestas y códigos de error: [Docs/GUIA.md](Docs/GUIA.md#referencia-de-rutas).

---

## Arquitectura

```
excel-template-builder/
├── main.js                  Servidor Express: middleware, rutas y cableado
├── Library/
│   ├── store.js             Persistencia de plantillas (revisiones, conflictos, escritura atómica)
│   ├── importer.js          .xlsx → plantilla JSON
│   ├── render.js            Plantilla + datos → .xlsx (ExcelJS)
│   ├── repeat-zones.js      Zonas repetibles: validación, expansión y reescritura de fórmulas
│   ├── shapes.js            Lectura/escritura de formas e imágenes del dibujo (DrawingML)
│   ├── excel-utils.js       Utilidades de colores, medidas y referencias
│   ├── automation.js        Esquema, ejemplo, validación, enlaces y OpenAPI (funciones puras)
│   ├── live.js              Tiempo real: SSE y presencia
│   └── usage.js             Métricas de uso
├── Public/                  Interfaz sin compilación: index.html + js/ y css/ por módulo
├── test/                    Pruebas con node:test
├── scripts/                 install.sh · run.sh
└── Docs/                    GUIA.md y capturas de pantalla
```

```mermaid
flowchart LR
  UI[Editor web] -- REST + SSE --> S[Servidor Express]
  INT[cURL · Python · Power Automate · n8n] -- POST JSON --> S
  S --> ST[(templates/*.json)]
  S --> U[(usage/usage.json)]
  S -- ExcelJS + DrawingML --> X[.xlsx / .zip]
```

Los datos se guardan en archivos JSON con escritura atómica; respaldar es copiar la carpeta `ETB_DATA_DIR`.

---

## Tecnologías

- **Servidor:** Node.js 18+, Express 4, ExcelJS, fast-xml-parser y adm-zip (formas e imágenes del dibujo), multer,
  helmet (CSP estricta), compression, express-rate-limit, cors.
- **Cliente:** HTML, CSS y JavaScript sin framework ni paso de compilación; hot-formula-parser para el cálculo de fórmulas
  en el editor.
- **Tiempo real:** Server-Sent Events con reconexión automática.
- **Pruebas:** `node:test` y `node:assert`.

---

## Desarrollo

```bash
npm run dev                               # recarga automática
npm test                                  # todas las pruebas
node --test test/automation.test.js       # un solo archivo
```

Las pruebas nunca tocan la carpeta `data/` real (usan `ETB_DATA_DIR` temporal). `Library/automation.js` exporta funciones
puras reutilizables desde otros scripts (`buildSchema`, `buildExample`, `validateBody`, `buildOpenApi`…); consulta
[Para desarrolladores](Docs/GUIA.md#para-desarrolladores).

---

## Limitaciones y hoja de ruta

**Limitaciones actuales**
- Sin autenticación ni permisos por usuario: la identidad de la presencia es informativa.
- Almacenamiento en archivos JSON pensado para un solo servidor (sin base de datos ni réplicas).
- El editor calcula las fórmulas más comunes para mostrarlas; el cálculo definitivo lo hace Excel al abrir el archivo.
- Las imágenes de variables deben enviarse como data URL en base64 (no se descargan URLs externas).
- El enlace directo (GET) no admite zonas repetibles ni hojas repetibles: para eso usa `POST`.

**Ideas para próximas versiones**
- [ ] Inicio de sesión (OIDC / Entra ID) y roles por plantilla.
- [ ] Historial de versiones con comparación y restauración.
- [ ] Almacenamiento opcional en base de datos u objeto (S3/Azure Blob).
- [ ] Exportación a PDF.
- [ ] Webhooks al terminar generaciones por lote.

---

## Licencia

Distribuido bajo la licencia **MIT**. Consulta [LICENSE](LICENSE).
