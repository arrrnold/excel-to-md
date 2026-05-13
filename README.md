# Excel → Markdown | Geist Engine

Herramienta profesional ultra-ligera para la conversión de datos tabulares a GitHub Flavored Markdown (GFM).

## 🛠 Arquitectura Superior
- **Motor Modular:** Lógica separada en `converter.js` para facilitar el testeo y mantenimiento.
- **Web Worker:** Los archivos grandes se procesan en un hilo independiente (`worker.js`).
- **Parsing de Alta Fidelidad:**
  - **TSV RFC 4180:** Soporte completo para comillas y saltos de línea internos en celdas.
  - **HTML Table DOM:** Preservación de negritas, itálicas y enlaces mediante análisis del árbol DOM.
  - **Auto-Alineación:** Detección inteligente de columnas numéricas.
  - **Merged Cells:** Propagación automática de valores en celdas combinadas.

## 🚀 Despliegue en GitHub Pages
1. Asegúrate de tener los archivos: `index.html`, `.nojekyll`, `css/`, `js/`, `vendor/`.
2. En GitHub: `Settings` -> `Pages` -> `Deploy from a branch` (main).

## ✅ Test Cases Cubiertos (Probado contra:)
1. [x] **Excel Desktop:** Celdas multi-línea (`Alt+Enter`).
2. [x] **Excel Desktop:** Celdas con pipes (`|`) escapados automáticamente.
3. [x] **Google Sheets:** Copiado directo de rangos con hyperlinks y formato.
4. [x] **Múltiples Hojas:** Archivos .xlsx con varias pestañas (genera encabezados H2).
5. [x] **Celdas Combinadas:** Propagación de datos desde la celda top-left.
6. [x] **Fechas:** Conversión automática a formato ISO (YYYY-MM-DD).
7. [x] **Archivos Grandes:** Procesamiento de archivos de hasta 20MB vía Web Worker.

## 🔒 Privacidad
Ningún dato sale del navegador. No hay trackers, ni analytics, ni backend.
