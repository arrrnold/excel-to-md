/**
 * app.js — Orquestador UI
 *
 * Flujo: el usuario pega datos → los ve en la zona de pegado →
 * presiona Ctrl+Enter (o el botón) → se genera el Markdown.
 * Los archivos se procesan automáticamente al cargarlos.
 */

const App = {
    state: {
        sheets: [],
        lastPastedHtml: null,
        lastPastedText: null,
        options: {
            hasHeader: true,
            compact: false,
            escapePipeHtml: false,
            repeatMerged: true,
            preserveBlankRows: false,
            dateFormat: 'YYYY-MM-DD'
        }
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

        this.optHeader     = document.getElementById('opt-header');
        this.optCompact    = document.getElementById('opt-compact');
        this.optMerges     = document.getElementById('opt-merges');
        this.optBlanks     = document.getElementById('opt-blanks');
        this.optPipeHtml   = document.getElementById('opt-pipe-html');
        this.optDateFormat = document.getElementById('opt-date-format');

        this.sheetContainer  = document.getElementById('sheet-selector-container');
        this.sheetCheckboxes = document.getElementById('sheet-checkboxes');
    },

    // ═══════ Event Binding ═══════
    bind() {
        // Paste: dejar que el contenteditable muestre los datos, guardar clipboard crudo
        this.pasteZone.addEventListener('paste', (e) => {
            this.state.lastPastedHtml = e.clipboardData.getData('text/html');
            this.state.lastPastedText = e.clipboardData.getData('text/plain');
            this.setStatus('Datos pegados. Presiona Ctrl+Enter para convertir.');
        });

        // Si el usuario edita manualmente, invalidar caché
        this.pasteZone.addEventListener('input', () => {
            this.state.lastPastedHtml = null;
            this.state.lastPastedText = null;
        });

        // Ctrl+Enter / Cmd+Enter dentro del paste zone
        this.pasteZone.addEventListener('keydown', (e) => {
            if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
                e.preventDefault();
                this.convertFromPasteZone();
            }
        });

        // Botón "Convertir"
        this.convertBtn.addEventListener('click', () => this.convertFromPasteZone());

        // Botón "Limpiar"
        this.clearBtn.addEventListener('click', () => {
            this.pasteZone.innerHTML = '';
            this.outputSection.classList.add('hidden');
            this.state.sheets = [];
            this.state.lastPastedHtml = null;
            this.state.lastPastedText = null;
            this.setStatus('Listo');
        });

        // Archivo
        this.fileInput.addEventListener('change', (e) => this.handleFile(e.target.files[0]));

        // Copiar y Descargar
        this.copyBtn.addEventListener('click', () => this.copyToClipboard());
        this.downloadBtn.addEventListener('click', () => this.downloadFile());

        // Reactividad de opciones
        [this.optHeader, this.optCompact, this.optMerges, this.optBlanks, this.optPipeHtml, this.optDateFormat]
            .forEach(el => el.addEventListener('change', () => { this.syncOptions(); this.render(); }));
    },


    // ═══════ Conversión desde Paste Zone ═══════
    convertFromPasteZone() {
        const html = this.state.lastPastedHtml || this.pasteZone.innerHTML;
        const text = this.state.lastPastedText || this.pasteZone.innerText;

        // Prioridad 1: HTML con tabla
        if (html && html.includes('<table')) {
            const result = Converter.parseTableHtml(html);
            if (result && result.rows.length > 0) {
                this.processResult({
                    sheets: [{ name: 'Paste', rows: result.rows, selected: true, alignments: result.alignments }]
                }, 'HTML');
                return;
            }
        }

        // Prioridad 2: TSV (tabs)
        if (text && text.includes('\t')) {
            const { rows } = Converter.parseTsvString(text);
            if (rows.length > 0) {
                this.processResult({ sheets: [{ name: 'Paste', rows, selected: true }] }, 'TSV');
                return;
            }
        }

        // Fallback: texto plano separado por líneas
        if (text && text.trim()) {
            const rows = text.trim().split('\n').map(r => r.split('\t'));
            this.processResult({ sheets: [{ name: 'Manual', rows, selected: true }] }, 'Texto');
        } else {
            this.setStatus('Pega datos de Excel primero.');
        }
    },

    // ═══════ Archivo ═══════
    handleFile(file) {
        if (!file) return;
        this.setStatus(`Procesando ${file.name}…`);

        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            if (file.size > 2 * 1024 * 1024) this.processViaWorker(data);
            else this.processInline(data);
        };
        reader.readAsArrayBuffer(file);
    },

    processInline(data) {
        try {
            const wb = XLSX.read(data, { type: 'array', cellDates: true, cellNF: false, cellText: false, raw: false });
            const sheets = wb.SheetNames.map((name, i) => {
                const ws = wb.Sheets[name];
                // Propagar merges
                if (this.state.options.repeatMerged && ws['!merges']) {
                    for (const m of ws['!merges']) {
                        const val = ws[XLSX.utils.encode_cell(m.s)];
                        if (!val) continue;
                        for (let r = m.s.r; r <= m.e.r; r++)
                            for (let c = m.s.c; c <= m.e.c; c++) {
                                const ref = XLSX.utils.encode_cell({ r, c });
                                if (!ws[ref]) ws[ref] = { ...val };
                            }
                    }
                }
                return {
                    name,
                    rows: XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: this.state.options.preserveBlankRows }),
                    selected: i === 0
                };
            });
            this.processResult({ sheets }, 'Excel');
        } catch (err) {
            this.setStatus(`Error: ${err.message}`);
        }
    },

    processViaWorker(data) {
        const worker = new Worker('js/worker.js');
        worker.postMessage({ data, options: this.state.options });
        worker.onmessage = (e) => {
            if (e.data.success) {
                const sheets = e.data.result.sheets.map((s, i) => ({ ...s, selected: i === 0 }));
                this.processResult({ sheets }, 'Worker');
            } else {
                this.setStatus(`Error Worker: ${e.data.error}`);
            }
            worker.terminate();
        };
    },

    // ═══════ Pipeline común ═══════
    processResult(result, source) {
        this.state.sheets = result.sheets.filter(s => s.rows.length > 0);

        if (!this.state.sheets.length) {
            this.setStatus('No se encontraron datos.');
            return;
        }

        const totalRows = this.state.sheets.reduce((n, s) => n + s.rows.length, 0);
        this.setStatus(`${this.state.sheets.length} hoja(s), ${totalRows} filas — via ${source}`);
        this.outputSection.classList.remove('hidden');
        this.renderSheetSelector();
        this.render();

        // Scroll suave al resultado
        this.outputSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
    },

    // ═══════ Sheet selector ═══════
    renderSheetSelector() {
        if (this.state.sheets.length <= 1) {
            this.sheetContainer.classList.add('hidden');
            return;
        }
        this.sheetContainer.classList.remove('hidden');
        this.sheetCheckboxes.innerHTML = '';
        this.state.sheets.forEach((sheet, i) => {
            const lbl = document.createElement('label');
            lbl.innerHTML = `<input type="checkbox" ${sheet.selected ? 'checked' : ''}> ${sheet.name}`;
            lbl.querySelector('input').addEventListener('change', (e) => {
                this.state.sheets[i].selected = e.target.checked;
                this.render();
            });
            this.sheetCheckboxes.appendChild(lbl);
        });
    },

    // ═══════ Render ═══════
    render() {
        const selected = this.state.sheets.filter(s => s.selected);
        if (!selected.length) { this.markdownArea.value = ''; this.previewArea.innerHTML = ''; return; }

        const mdParts = selected.map(sheet => {
            const heading = selected.length > 1 ? `## ${sheet.name}\n\n` : '';
            return heading + Converter.sheetToMarkdown(sheet.rows, {
                ...this.state.options,
                alignments: sheet.alignments || []
            });
        });

        this.markdownArea.value = mdParts.join('\n\n');
        this.previewArea.innerHTML = this._previewHtml(selected[0].rows);
    },

    _previewHtml(rows) {
        const data = JSON.parse(JSON.stringify(rows));
        const hdr = this.state.options.hasHeader ? data.shift() : data[0].map((_, i) => `Col ${i + 1}`);
        const head = '<thead><tr>' + hdr.map(h => `<th>${h}</th>`).join('') + '</tr></thead>';
        const body = '<tbody>' + data.map(r => '<tr>' + r.map(c => `<td>${c}</td>`).join('') + '</tr>').join('') + '</tbody>';
        return `<table role="grid">${head}${body}</table>`;
    },

    // ═══════ Actions ═══════
    copyToClipboard() {
        navigator.clipboard.writeText(this.markdownArea.value).then(() => {
            const orig = this.copyBtn.innerHTML;
            this.copyBtn.textContent = '¡Copiado!';
            setTimeout(() => this.copyBtn.innerHTML = orig, 2000);
        });
    },

    downloadFile() {
        const blob = new Blob([this.markdownArea.value], { type: 'text/markdown' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `tabla_${Date.now()}.md`;
        a.click();
        URL.revokeObjectURL(url);
    },

    // ═══════ Helpers ═══════
    syncOptions() {
        this.state.options = {
            hasHeader:        this.optHeader.checked,
            compact:          this.optCompact.checked,
            repeatMerged:     this.optMerges.checked,
            preserveBlankRows: this.optBlanks.checked,
            escapePipeHtml:   this.optPipeHtml.checked,
            dateFormat:       this.optDateFormat.value
        };
    },

    setStatus(msg) {
        this.status.textContent = msg;
    }
};

document.addEventListener('DOMContentLoaded', () => App.init());
