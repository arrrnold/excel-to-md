/**
 * converter.js — Motor de conversión Excel → Markdown
 *
 * Genera tablas con alineación visual perfecta:
 * las barras | quedan verticalmente alineadas
 * y los delimitadores se expanden con guiones.
 */

const Converter = {

    // ════════════════════════════════════════
    // PARSER TSV — Máquina de estados RFC 4180
    // ════════════════════════════════════════
    parseTsvString(text) {
        const rows = [];
        let row = [];
        let cell = '';
        let inQuotes = false;

        // Normalizar: quitar BOM y unificar saltos de línea
        const src = text.replace(/^\uFEFF/, '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');

        for (let i = 0; i < src.length; i++) {
            const ch = src[i];
            const next = src[i + 1];

            if (inQuotes) {
                if (ch === '"' && next === '"') { cell += '"'; i++; }
                else if (ch === '"')             { inQuotes = false; }
                else                             { cell += ch; }
            } else {
                if (ch === '"' && cell === '') { inQuotes = true; }
                else if (ch === '\t')          { row.push(cell); cell = ''; }
                else if (ch === '\n')          { row.push(cell); rows.push(row); row = []; cell = ''; }
                else                           { cell += ch; }
            }
        }
        // Última celda/fila
        if (cell !== '' || row.length > 0) { row.push(cell); rows.push(row); }
        return { rows };
    },

    // ════════════════════════════════════════
    // PARSER CSV — Máquina de estados RFC 4180
    // ════════════════════════════════════════
    parseCsvString(text) {
        const rows = [];
        let row = [];
        let cell = '';
        let inQuotes = false;

        // Normalizar: quitar BOM y unificar saltos de línea
        const src = text.replace(/^\uFEFF/, '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');

        for (let i = 0; i < src.length; i++) {
            const ch = src[i];
            const next = src[i + 1];

            if (inQuotes) {
                if (ch === '"' && next === '"') { cell += '"'; i++; }
                else if (ch === '"')             { inQuotes = false; }
                else                             { cell += ch; }
            } else {
                if (ch === '"' && cell === '') { inQuotes = true; }
                else if (ch === ',')           { row.push(cell); cell = ''; }
                else if (ch === '\n')          { row.push(cell); rows.push(row); row = []; cell = ''; }
                else                           { cell += ch; }
            }
        }
        // Última celda/fila
        if (cell !== '' || row.length > 0) { row.push(cell); rows.push(row); }
        return { rows };
    },

    // ════════════════════════════════════════
    // PARSER HTML — Extrae tabla del clipboard
    // ════════════════════════════════════════
    parseTableHtml(htmlString) {
        const doc = new DOMParser().parseFromString(htmlString, 'text/html');
        const table = doc.querySelector('table');
        if (!table) return null;

        const rows = [];
        const colAligns = [];

        for (const tr of table.querySelectorAll('tr')) {
            const row = [];
            let colIdx = 0;
            for (const cell of tr.querySelectorAll('td, th')) {
                const colspan = parseInt(cell.getAttribute('colspan') || 1);
                const content = this._richText(cell);

                // Capturar alineación de la primera fila
                if (rows.length === 0) {
                    const s = (cell.getAttribute('style') || '').replace(/\s/g, '');
                    const a = cell.getAttribute('align') || '';
                    if (s.includes('text-align:right')  || a === 'right')  colAligns[colIdx] = 'r';
                    else if (s.includes('text-align:center') || a === 'center') colAligns[colIdx] = 'c';
                    else colAligns[colIdx] = 'l';
                }
                for (let k = 0; k < colspan; k++) { row.push(content); colIdx++; }
            }
            if (row.length) rows.push(row);
        }
        return { rows, alignments: colAligns };
    },

    /** Extrae texto rico preservando negritas, links, etc. */
    _richText(cell) {
        const el = cell.cloneNode(true);
        // Eliminar basura de MSO / Office
        el.querySelectorAll('meta, style, link, script, o\\:p').forEach(n => n.remove());

        const walk = (node) => {
            if (node.nodeType === 3) return node.textContent;
            if (node.nodeType !== 1) return '';
            const inner = Array.from(node.childNodes).map(walk).join('');
            const tag = node.tagName.toLowerCase();
            const sty = node.getAttribute('style') || '';
            if (tag === 'b' || tag === 'strong' || sty.includes('font-weight:bold') || sty.includes('font-weight:700'))
                return inner.trim() ? `**${inner.trim()}**` : '';
            if (tag === 'i' || tag === 'em' || sty.includes('font-style:italic'))
                return inner.trim() ? `*${inner.trim()}*` : '';
            if (tag === 's' || tag === 'del' || tag === 'strike' || sty.includes('line-through'))
                return inner.trim() ? `~~${inner.trim()}~~` : '';
            if (tag === 'a') {
                const href = node.getAttribute('href') || '';
                if (href && !href.startsWith('javascript:') && !href.startsWith('data:') && inner.trim())
                    return `[${inner.trim()}](${href})`;
            }
            if (tag === 'code' || tag === 'tt' || tag === 'kbd') return `\`${inner.trim()}\``;
            if (tag === 'br') return '\n';
            return inner;
        };
        return Array.from(el.childNodes).map(walk).join('').trim();
    },

    // ════════════════════════════════════════════
    // GENERADOR MARKDOWN — Pretty-print alineado
    // ════════════════════════════════════════════
    sheetToMarkdown(rows, opts = {}) {
        if (!rows || !rows.length) return '';

        // 1. Normalizar: cada celda a string escapado
        const escaped = rows.map(r => r.map(c => this._escape(c, opts)));

        // 2. Separar headers de data
        let headers, data;
        if (opts.hasHeader !== false) {
            headers = escaped[0];
            data = escaped.slice(1);
        } else {
            headers = escaped[0].map((_, i) => `Column ${i + 1}`);
            data = escaped;
        }

        const numCols = headers.length;

        // 3. Alineación por columna
        const align = headers.map((_, i) => {
            if (opts.alignments && opts.alignments[i]) return opts.alignments[i];
            // Heurística: >60% numérico → derecha
            let numCount = 0;
            for (const row of data) {
                const v = (row[i] || '').replace(/[$€£,% ]/g, '');
                if (v && !isNaN(v)) numCount++;
            }
            return data.length && (numCount / data.length) > 0.6 ? 'r' : 'l';
        });

        // 4. Calcular ancho de cada columna
        //    Mínimo 3 caracteres para que quepan los delimitadores
        const widths = new Array(numCols).fill(3);
        for (let i = 0; i < numCols; i++) {
            widths[i] = Math.max(widths[i], headers[i].length);
            for (const row of data) {
                widths[i] = Math.max(widths[i], (row[i] || '').length);
            }
        }

        // 5. Construir delimitadores con alineación expandida
        const delims = align.map((a, i) => {
            const w = widths[i];
            if (a === 'r') return '-'.repeat(w - 1) + ':';
            if (a === 'c') return ':' + '-'.repeat(w - 2) + ':';
            return ':' + '-'.repeat(w - 1);
        });

        // 6. Funciones de formato
        const pad = (s, i) => {
            if (opts.compact) return s;
            const w = widths[i];
            return align[i] === 'r'
                ? s.padStart(w)
                : s.padEnd(w);
        };

        const fmtRow = (cells) => {
            const parts = cells.map((c, i) => ` ${pad(c || '', i)} `);
            return '|' + parts.join('|') + '|';
        };

        const fmtDelim = () => {
            if (opts.compact) return '| ' + delims.join(' | ') + ' |';
            const parts = delims.map((d, i) => ` ${d.padEnd(widths[i])} `);
            return '|' + parts.join('|') + '|';
        };

        // 7. Ensamblar
        const lines = [
            fmtRow(headers),
            fmtDelim(),
            ...data.map(r => fmtRow(r))
        ];

        return lines.join('\n');
    },

    /** Escapa una celda para GFM */
    _escape(val, opts) {
        if (val === null || val === undefined) return '';

        // Fechas
        if (val instanceof Date && !isNaN(val)) {
            const y = val.getUTCFullYear();
            const m = String(val.getUTCMonth() + 1).padStart(2, '0');
            const d = String(val.getUTCDate()).padStart(2, '0');
            const fmt = opts.dateFormat || 'YYYY-MM-DD';
            if (fmt === 'DD/MM/YYYY') return `${d}/${m}/${y}`;
            if (fmt === 'MM/DD/YYYY') return `${m}/${d}/${y}`;
            return `${y}-${m}-${d}`;
        }

        let s = String(val);

        // Saltos de línea → <br>
        s = s.replace(/\r?\n/g, ' <br> ');

        // Pipes
        s = s.replace(/\|/g, opts.escapePipeHtml ? '&#124;' : '\\|');

        // Bug GFM: backslash al final de celda escapa el pipe siguiente
        if (s.endsWith('\\')) s += ' ';

        // Tabs → espacios
        s = s.replace(/\t/g, '    ');

        // Eliminar caracteres de control excepto los ya tratados
        s = s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');

        return s.trim();
    }
};

if (typeof module !== 'undefined') module.exports = Converter;
else window.Converter = Converter;
