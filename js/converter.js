/**
 * converter.js - Ultra-High Fidelity Edition
 * Genera Markdown visualmente alineado (Pretty-Print).
 */

const Converter = {
    /**
     * Parser TSV RFC 4180 (Máquina de estados)
     */
    parseTsvString(text) {
        const rows = [];
        let currentRow = [];
        let currentCell = '';
        let inQuotes = false;
        const cleanText = text.replace(/^\uFEFF/, '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');

        for (let i = 0; i < cleanText.length; i++) {
            const char = cleanText[i];
            const nextChar = cleanText[i + 1];

            if (inQuotes) {
                if (char === '"' && nextChar === '"') {
                    currentCell += '"'; i++;
                } else if (char === '"') inQuotes = false;
                else currentCell += char;
            } else {
                if (char === '"' && currentCell === '') inQuotes = true;
                else if (char === '\t') { currentRow.push(currentCell); currentCell = ''; }
                else if (char === '\n') { 
                    currentRow.push(currentCell); rows.push(currentRow);
                    currentRow = []; currentCell = '';
                } else currentCell += char;
            }
        }
        if (currentCell !== '' || currentRow.length > 0) { currentRow.push(currentCell); rows.push(currentRow); }
        return { rows };
    },

    parseTableHtml(htmlString) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(htmlString, 'text/html');
        const table = doc.querySelector('table');
        if (!table) return null;

        const rows = [];
        const alignments = [];
        const trs = Array.from(table.querySelectorAll('tr'));

        trs.forEach((tr, rowIndex) => {
            const row = [];
            const cells = Array.from(tr.querySelectorAll('td, th'));
            cells.forEach((cell, colIndex) => {
                const colspan = parseInt(cell.getAttribute('colspan') || 1);
                const content = this._extractRichText(cell);
                
                if (rowIndex === 0) {
                    const style = cell.getAttribute('style') || '';
                    const alignAttr = cell.getAttribute('align') || '';
                    if (style.includes('text-align:right') || alignAttr === 'right') alignments[colIndex] = 'r';
                    else if (style.includes('text-align:center') || alignAttr === 'center') alignments[colIndex] = 'c';
                    else alignments[colIndex] = 'l';
                }
                for (let k = 0; k < colspan; k++) row.push(content);
            });
            if (row.length > 0) rows.push(row);
        });
        return { rows, alignments };
    },

    _extractRichText(cell) {
        const clone = cell.cloneNode(true);
        clone.querySelectorAll('meta, style, link, script').forEach(n => n.remove());

        const processNode = (node) => {
            if (node.nodeType === 3) return node.textContent;
            if (node.nodeType !== 1) return '';
            let content = Array.from(node.childNodes).map(processNode).join('');
            const tag = node.tagName.toLowerCase();
            const style = node.getAttribute('style') || '';

            if (tag === 'b' || tag === 'strong' || style.includes('font-weight:bold')) return `**${content.trim()}**`;
            if (tag === 'i' || tag === 'em' || style.includes('font-style:italic')) return `*${content.trim()}*`;
            if (tag === 'a' && node.href && !node.href.startsWith('javascript:')) return `[${content.trim()}](${node.href})`;
            if (tag === 'br') return '\n';
            return content;
        };
        return Array.from(clone.childNodes).map(processNode).join('').trim();
    },

    /**
     * Generador Markdown con Alineación Vertical de Pipes
     */
    sheetToMarkdown(rows, options = {}) {
        if (!rows || rows.length === 0) return '';
        
        // 1. Normalizar y Escapar
        let data = rows.map(row => row.map(cell => this.escapeCell(cell, options)));
        
        // 2. Manejo de Headers
        let headers = [];
        if (options.hasHeader) headers = data.shift();
        else headers = data[0].map((_, i) => `Column ${i + 1}`);

        // 3. Determinar Alineaciones
        const finalAlignments = headers.map((h, i) => {
            if (options.alignments && options.alignments[i]) return options.alignments[i];
            const numericCount = data.filter(row => row[i] && !isNaN(parseFloat(row[i].replace(/[$,]/g, ''))) && isFinite(row[i].replace(/[$,]/g, ''))).length;
            return (numericCount / (data.length || 1)) > 0.6 ? 'r' : 'l';
        });

        // 4. Calcular anchos de columna (incluyendo headers y delimitadores)
        const colWidths = headers.map((h, i) => {
            let max = String(h).length;
            data.forEach(row => {
                if (row[i]) max = Math.max(max, String(row[i]).length);
            });
            return Math.max(max, 3); // Mínimo 3 para '---'
        });

        // 5. Construir Delimitador Alineado
        const delimiterRow = finalAlignments.map((a, i) => {
            const width = colWidths[i];
            if (a === 'r') return '-'.repeat(width) + ':';
            if (a === 'c') return ':' + '-'.repeat(width - 1) + ':';
            return ':' + '-'.repeat(width);
        });

        // 6. Formatear Filas
        const formatRow = (row, isHeader = false) => {
            const padded = row.map((c, i) => {
                const s = String(c);
                const align = isHeader ? 'l' : finalAlignments[i];
                const width = colWidths[i] + (isHeader ? 0 : (finalAlignments[i] === 'l' ? 1 : 0));
                
                if (options.compact) return s;
                
                // Si es numérico (derecha), padStart. Si no, padEnd.
                return align === 'r' ? s.padStart(colWidths[i], ' ') : s.padEnd(colWidths[i], ' ');
            });
            return `| ${padded.join(' | ')} |`;
        };

        const formatDelimiter = (row) => {
            return `| ${row.join(' | ')} |`;
        };

        const result = [
            formatRow(headers, true),
            formatDelimiter(delimiterRow),
            ...data.map(row => formatRow(row))
        ];

        return result.join('\n');
    },

    escapeCell(val, options) {
        if (val === null || val === undefined) return '';
        if (val instanceof Date) {
            const y = val.getUTCFullYear(), m = String(val.getUTCMonth()+1).padStart(2,'0'), d = String(val.getUTCDate()).padStart(2,'0');
            if (options.dateFormat === 'DD/MM/YYYY') return `${d}/${m}/${y}`;
            return `${y}-${m}-${d}`;
        }
        let str = String(val).replace(/\r?\n/g, ' <br> ');
        const pipe = options.escapePipeHtml ? '&#124;' : '\\|';
        str = str.replace(/\|/g, pipe);
        if (str.endsWith('\\')) str += ' ';
        return str.trim();
    }
};

if (typeof module !== 'undefined') module.exports = Converter;
else window.Converter = Converter;
