/**
 * word-app.js — Orquestador Word/Rich-Text to Markdown
 *
 * Flujo:
 *   1. Usuario pega contenido de Word (rich HTML) o sube un .docx
 *   2. Presiona Ctrl+Enter (o boton) para convertir
 *   3. mammoth.js convierte .docx → HTML
 *   4. Turndown.js convierte HTML → Markdown
 *   5. Se muestra el resultado listo para copiar/descargar
 */

const WordApp = {
    state: {
        lastPastedHtml: null,
        docxArrayBuffer: null
    },

    init() {
        this.cache();
        this.bind();
    },

    // ═══════ DOM Cache ═══════
    cache() {
        this.pasteZone     = document.getElementById('paste-zone');
        this.convertBtn    = document.getElementById('convert-btn');
        this.clearBtn      = document.getElementById('clear-btn');
        this.fileInput     = document.getElementById('file-input');
        this.markdownArea  = document.getElementById('markdown-output');
        this.previewArea   = document.getElementById('preview-container');
        this.status        = document.getElementById('status-msg');
        this.copyBtn       = document.getElementById('copy-btn');
        this.downloadBtn   = document.getElementById('download-btn');
        this.outputSection = document.getElementById('output-section');
    },

    // ═══════ Event Binding ═══════
    bind() {
        // Paste: guardar HTML del clipboard
        this.pasteZone.addEventListener('paste', (e) => {
            this.state.lastPastedHtml = e.clipboardData.getData('text/html');
            this.state.docxArrayBuffer = null;
            this.setStatus('Contenido pegado. Presiona Ctrl+Enter para convertir.');
        });

        // Edicion manual invalida cache
        this.pasteZone.addEventListener('input', () => {
            this.state.lastPastedHtml = null;
        });

        // Ctrl+Enter / Cmd+Enter
        this.pasteZone.addEventListener('keydown', (e) => {
            if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
                e.preventDefault();
                this.convert();
            }
        });

        // Botones
        this.convertBtn.addEventListener('click', () => this.convert());

        this.clearBtn.addEventListener('click', () => {
            this.pasteZone.innerHTML = '';
            this.outputSection.classList.add('hidden');
            this.markdownArea.value = '';
            this.previewArea.innerHTML = '';
            this.state.lastPastedHtml = null;
            this.state.docxArrayBuffer = null;
            this.setStatus('Listo');
        });

        // Archivo .docx
        this.fileInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;
            this.setStatus(`Cargando ${file.name}...`);
            const reader = new FileReader();
            reader.onload = (ev) => {
                this.state.docxArrayBuffer = ev.target.result;
                this.state.lastPastedHtml = null;
                this.pasteZone.innerHTML = `<em>Archivo cargado: ${file.name}</em>`;
                this.setStatus(`${file.name} listo. Presiona Ctrl+Enter para convertir.`);
            };
            reader.readAsArrayBuffer(file);
        });

        // Copiar y descargar
        this.copyBtn.addEventListener('click', () => this.copyToClipboard());
        this.downloadBtn.addEventListener('click', () => this.downloadFile());
    },

    // ═══════ Conversion ═══════
    async convert() {
        const turndownOpts = this.getTurndownOptions();

        // Ruta 1: Archivo .docx cargado
        if (this.state.docxArrayBuffer) {
            this.setStatus('Convirtiendo .docx...');
            try {
                const result = await mammoth.convertToHtml(
                    { arrayBuffer: this.state.docxArrayBuffer },
                    { styleMap: [
                        "p[style-name='Title'] => h1:fresh",
                        "p[style-name='Heading 1'] => h1:fresh",
                        "p[style-name='Heading 2'] => h2:fresh",
                        "p[style-name='Heading 3'] => h3:fresh",
                        "p[style-name='Heading 4'] => h4:fresh",
                        "p[style-name='Subtitle'] => h2:fresh"
                    ]}
                );
                this.processHtml(result.value, turndownOpts);

                if (result.messages.length > 0) {
                    const warnings = result.messages.filter(m => m.type === 'warning').length;
                    if (warnings > 0) {
                        this.setStatus(`Convertido con ${warnings} advertencia(s). Revisa el resultado.`);
                    }
                }
            } catch (err) {
                this.setStatus(`Error: ${err.message}`);
            }
            return;
        }

        // Ruta 2: HTML pegado desde Word/navegador
        const html = this.state.lastPastedHtml || this.pasteZone.innerHTML;
        if (!html || html.trim() === '' || html === '<br>') {
            this.setStatus('Pega contenido de Word o sube un archivo .docx primero.');
            return;
        }

        this.processHtml(html, turndownOpts);
    },

    processHtml(html, opts) {
        // Limpiar HTML de basura de Office
        const cleaned = this.cleanOfficeHtml(html);

        // Configurar Turndown
        const td = new TurndownService({
            headingStyle:    opts.headingStyle,
            bulletListMarker: opts.bulletMarker,
            codeBlockStyle:  'fenced',
            fence:           '```',
            emDelimiter:     '*',
            strongDelimiter: '**',
            linkStyle:       opts.linkStyle,
            linkReferenceStyle: 'full',
            br:              opts.preserveBr ? '  \n' : '\n'
        });

        // Reglas GFM: tablas y strikethrough
        if (opts.gfm) {
            td.addRule('strikethrough', {
                filter: ['del', 's', 'strike'],
                replacement: (content) => `~~${content}~~`
            });

            td.addRule('table', {
                filter: 'table',
                replacement: (content, node) => {
                    return this.tableToMarkdown(node) + '\n\n';
                }
            });

            td.addRule('tableCell', {
                filter: ['th', 'td'],
                replacement: (content) => content.replace(/\n/g, ' ').trim()
            });

            td.addRule('tableRow', {
                filter: 'tr',
                replacement: () => ''
            });

            td.addRule('tableSection', {
                filter: ['thead', 'tbody', 'tfoot'],
                replacement: (content) => content
            });
        }

        // Convertir
        let md = td.turndown(cleaned);

        // Post-procesamiento: limpiar lineas vacias excesivas
        md = md.replace(/\n{4,}/g, '\n\n\n');
        md = md.trim() + '\n';

        this.markdownArea.value = md;
        this.previewArea.innerHTML = cleaned;
        this.outputSection.classList.remove('hidden');
        this.setStatus(`Convertido. ${md.split('\n').length} lineas de Markdown.`);
        this.outputSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
    },

    // ═══════ Tabla HTML → Markdown ═══════
    tableToMarkdown(tableNode) {
        const rows = [];
        for (const tr of tableNode.querySelectorAll('tr')) {
            const cells = [];
            for (const cell of tr.querySelectorAll('td, th')) {
                cells.push(cell.textContent.replace(/\|/g, '\\|').replace(/\n/g, ' ').trim());
            }
            if (cells.length > 0) rows.push(cells);
        }

        if (rows.length === 0) return '';

        // Calcular anchos para alineacion
        const numCols = Math.max(...rows.map(r => r.length));
        const widths = new Array(numCols).fill(3);
        for (const row of rows) {
            for (let i = 0; i < row.length; i++) {
                widths[i] = Math.max(widths[i], row[i].length);
            }
        }

        const fmtRow = (cells) => {
            const padded = cells.map((c, i) => (c || '').padEnd(widths[i] || 3));
            return '| ' + padded.join(' | ') + ' |';
        };

        const header = fmtRow(rows[0]);
        const sep = '| ' + widths.map(w => '-'.repeat(w)).join(' | ') + ' |';
        const body = rows.slice(1).map(r => fmtRow(r));

        return [header, sep, ...body].join('\n');
    },

    // ═══════ Limpiar HTML de Office ═══════
    cleanOfficeHtml(html) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');

        // Eliminar elementos basura de Microsoft Office
        doc.querySelectorAll('meta, style, link, script, xml, o\\:p, o\\:SmartTagType').forEach(n => n.remove());

        // Eliminar comentarios condicionales de IE
        const walker = doc.createTreeWalker(doc.body, NodeFilter.SHOW_COMMENT);
        const comments = [];
        while (walker.nextNode()) comments.push(walker.currentNode);
        comments.forEach(c => c.parentNode.removeChild(c));

        // Eliminar atributos de estilo de Office (class="MsoNormal", etc.)
        doc.querySelectorAll('[class^="Mso"]').forEach(el => {
            el.removeAttribute('class');
        });

        // Limpiar spans vacios
        doc.querySelectorAll('span').forEach(span => {
            if (!span.textContent.trim() && span.querySelectorAll('img, br').length === 0) {
                span.remove();
            }
        });

        return doc.body.innerHTML;
    },

    // ═══════ Opciones de Turndown ═══════
    getTurndownOptions() {
        const headingStyle = document.querySelector('input[name="heading"]:checked')?.value === 'setext' ? 'setext' : 'atx';
        const bulletMarker = document.querySelector('input[name="bullet"]:checked')?.value || '-';
        const linkStyle = document.querySelector('input[name="linkStyle"]:checked')?.value || 'inlined';
        const gfm = document.getElementById('opt-gfm')?.checked ?? true;
        const preserveBr = document.getElementById('opt-br')?.checked ?? false;

        return { headingStyle, bulletMarker, linkStyle, gfm, preserveBr };
    },

    // ═══════ Acciones ═══════
    copyToClipboard() {
        navigator.clipboard.writeText(this.markdownArea.value).then(() => {
            const orig = this.copyBtn.innerHTML;
            this.copyBtn.textContent = 'Copiado';
            setTimeout(() => this.copyBtn.innerHTML = orig, 2000);
        });
    },

    downloadFile() {
        const blob = new Blob([this.markdownArea.value], { type: 'text/markdown' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `documento_${Date.now()}.md`;
        a.click();
        URL.revokeObjectURL(url);
    },

    setStatus(msg) {
        this.status.textContent = msg;
    }
};

document.addEventListener('DOMContentLoaded', () => WordApp.init());
