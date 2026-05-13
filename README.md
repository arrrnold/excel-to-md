# to-markdown

Suite de herramientas para convertir documentos a Markdown. 100% client-side, sin servidores, privado por diseno.

## Herramientas

| Herramienta       | Ruta     | Descripcion                                     |
| :---------------- | :------- | :---------------------------------------------- |
| Excel to Markdown | `/`      | Pega desde Excel/Sheets o sube .xlsx/.csv        |
| Word to Markdown  | `/word/` | Pega desde Word o sube .docx                     |

## Tecnologias

- **UI:** [Pico CSS](https://picocss.com) (vendorizado)
- **Excel:** [SheetJS](https://sheetjs.com) (vendorizado)
- **Word:** [mammoth.js](https://github.com/mwilliamson/mammoth.js) (vendorizado)
- **HTML to MD:** [Turndown.js](https://github.com/mixmark-io/turndown) (vendorizado)
- **Despliegue:** GitHub Pages (archivos estaticos, sin build)

## Uso

1. Abre la herramienta que necesites
2. Pega contenido con `Ctrl+V` / `Cmd+V`
3. Presiona `Ctrl+Enter` para convertir
4. Copia o descarga el Markdown

## Licencia

MIT - Arnold Torres Maldonado
