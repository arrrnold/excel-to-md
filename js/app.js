/**
 * app.js - Senior Frontend Orchestrator (Pretty-Print Version)
 */
const App = {
    state: {
        sheets: [],
        alignments: [],
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
        this.initTheme();
    },

    cache() {
        this.body = document.body;
        this.themeToggle = document.getElementById('theme-toggle');
        this.pasteZone = document.getElementById('paste-zone');
        this.fileInput = document.getElementById('file-input');
        this.markdownArea = document.getElementById('markdown-output');
        this.previewArea = document.getElementById('preview-container');
        this.status = document.getElementById('status-msg');
        this.copyBtn = document.getElementById('copy-btn');
        this.downloadBtn = document.getElementById('download-btn');
        this.outputSection = document.getElementById('output-section');
        
        // Config Elements
        this.optHeader = document.getElementById('opt-header');
        this.optCompact = document.getElementById('opt-compact');
        this.optMerges = document.getElementById('opt-merges');
        this.optBlanks = document.getElementById('opt-blanks');
        this.optPipeHtml = document.getElementById('opt-pipe-html');
        this.optDateFormat = document.getElementById('opt-date-format');
        
        // Sheet Selector
        this.sheetSelectorContainer = document.getElementById('sheet-selector-container');
        this.sheetCheckboxes = document.getElementById('sheet-checkboxes');
    },

    bind() {
        this.themeToggle.addEventListener('click', () => this.toggleTheme());
        window.addEventListener('paste', (e) => this.handlePaste(e), true);
        this.fileInput.addEventListener('change', (e) => this.handleFile(e.target.files[0]));
        this.copyBtn.addEventListener('click', () => this.copyToClipboard());
        this.downloadBtn.addEventListener('click', () => this.downloadFile());

        // Reactividad de Opciones
        const configEls = [this.optHeader, this.optCompact, this.optMerges, this.optBlanks, this.optPipeHtml, this.optDateFormat];
        configEls.forEach(el => {
            el.addEventListener('change', () => {
                this.updateOptions();
                this.render();
            });
        });
    },

    updateOptions() {
        this.state.options = {
            hasHeader: this.optHeader.checked,
            compact: this.optCompact.checked,
            repeatMerged: this.optMerges.checked,
            preserveBlankRows: this.optBlanks.checked,
            escapePipeHtml: this.optPipeHtml.checked,
            dateFormat: this.optDateFormat.value
        };
    },

    handlePaste(e) {
        const text = e.clipboardData.getData('text/plain');
        const html = e.clipboardData.getData('text/html');

        if (text && text.includes('\t')) {
            e.preventDefault();
            const { rows } = Converter.parseTsvString(text);
            this.processResult({ sheets: [{ name: 'Pasted', rows, selected: true }] }, 'TSV');
            return;
        }

        if (html && html.includes('<table')) {
            e.preventDefault();
            const result = Converter.parseTableHtml(html);
            if (result && result.rows.length > 0) {
                this.processResult({ sheets: [{ name: 'Pasted HTML', rows: result.rows, selected: true, alignments: result.alignments }] }, 'HTML Rico');
                return;
            }
        }
    },

    handleFile(file) {
        if (!file) return;
        this.status.innerText = `Cargando ${file.name}...`;
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            if (file.size > 2 * 1024 * 1024) this.useWorker(data);
            else this.useInline(data);
        };
        reader.readAsArrayBuffer(file);
    },

    useInline(data) {
        try {
            const workbook = XLSX.read(data, { type: 'array', cellDates: true, cellNF: false, cellText: false, raw: false });
            const sheets = workbook.SheetNames.map((name, i) => {
                const sheet = workbook.Sheets[name];
                if (this.state.options.repeatMerged && sheet['!merges']) {
                    sheet['!merges'].forEach(m => {
                        const val = sheet[XLSX.utils.encode_cell(m.s)];
                        if (val) {
                            for (let r = m.s.r; r <= m.e.r; r++) {
                                for (let c = m.s.c; c <= m.e.c; c++) {
                                    const ref = XLSX.utils.encode_cell({r, c});
                                    if (!sheet[ref]) sheet[ref] = { ...val };
                                }
                            }
                        }
                    });
                }
                return {
                    name,
                    rows: XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: this.state.options.preserveBlankRows }),
                    selected: i === 0
                };
            });
            this.processResult({ sheets }, 'Excel');
        } catch (err) {
            this.status.innerText = "Error: " + err.message;
        }
    },

    useWorker(data) {
        const worker = new Worker('js/worker.js');
        worker.postMessage({ data, options: this.state.options });
        worker.onmessage = (e) => {
            if (e.data.success) {
                const sheets = e.data.result.sheets.map((s, i) => ({ ...s, selected: i === 0 }));
                this.processResult({ sheets }, 'Worker');
            } else this.status.innerText = "Error Worker: " + e.data.error;
            worker.terminate();
        };
    },

    processResult(result, source) {
        this.state.sheets = result.sheets.filter(s => s.rows.length > 0);
        this.status.innerText = `Éxito. ${this.state.sheets.length} hoja(s) procesada(s) via ${source}.`;
        this.outputSection.classList.remove('hidden');
        this.renderSheetSelector();
        this.render();
    },

    renderSheetSelector() {
        if (this.state.sheets.length <= 1) {
            this.sheetSelectorContainer.classList.add('hidden');
            return;
        }
        this.sheetSelectorContainer.classList.remove('hidden');
        this.sheetCheckboxes.innerHTML = '';
        this.state.sheets.forEach((sheet, i) => {
            const label = document.createElement('label');
            label.className = 'checkbox-label';
            label.innerHTML = `<input type="checkbox" ${sheet.selected ? 'checked' : ''}> ${sheet.name}`;
            label.querySelector('input').addEventListener('change', (e) => {
                this.state.sheets[i].selected = e.target.checked;
                this.render();
            });
            this.sheetCheckboxes.appendChild(label);
        });
    },

    render() {
        const selected = this.state.sheets.filter(s => s.selected);
        if (!selected.length) {
            this.markdownArea.value = '';
            return;
        }

        const mdParts = selected.map(sheet => {
            const title = selected.length > 1 ? `## ${sheet.name}\n\n` : '';
            return title + Converter.sheetToMarkdown(sheet.rows, {
                ...this.state.options,
                alignments: sheet.alignments || []
            });
        });

        this.markdownArea.value = mdParts.join('\n\n');
        this.previewArea.innerHTML = this.renderPreview(selected[0].rows);
    },

    renderPreview(rows) {
        const data = JSON.parse(JSON.stringify(rows));
        const headers = this.state.options.hasHeader ? data.shift() : data[0].map((_, i) => `Col ${i + 1}`);
        let html = '<table><thead><tr>' + headers.map(h => `<th>${h}</th>`).join('') + '</tr></thead><tbody>';
        data.slice(0, 10).forEach(row => {
            html += '<tr>' + row.map(c => `<td>${c}</td>`).join('') + '</tr>';
        });
        html += '</tbody></table>';
        if (data.length > 10) html += `<p style="font-size:12px; margin-top:10px; color:var(--accents-5)">Mostrando 10 de ${data.length} filas...</p>`;
        return html;
    },

    copyToClipboard() {
        navigator.clipboard.writeText(this.markdownArea.value).then(() => {
            const oldText = this.copyBtn.innerText;
            this.copyBtn.innerText = "¡Copiado!";
            setTimeout(() => this.copyBtn.innerText = oldText, 2000);
        });
    },

    downloadFile() {
        const blob = new Blob([this.markdownArea.value], { type: 'text/markdown' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `excel_export_${new Date().getTime()}.md`;
        a.click();
        URL.revokeObjectURL(url);
    },

    initTheme() {
        const dark = localStorage.getItem('theme') === 'dark' || (!localStorage.getItem('theme') && window.matchMedia('(prefers-color-scheme: dark)').matches);
        if (dark) this.body.classList.add('dark');
        this.updateThemeIcons(dark);
    },

    toggleTheme() {
        const dark = this.body.classList.toggle('dark');
        localStorage.setItem('theme', dark ? 'dark' : 'light');
        this.updateThemeIcons(dark);
    },

    updateThemeIcons(dark) {
        document.getElementById('sun-icon').classList.toggle('hidden', !dark);
        document.getElementById('moon-icon').classList.toggle('hidden', dark);
    }
};

document.addEventListener('DOMContentLoaded', () => App.init());
