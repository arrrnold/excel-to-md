/**
 * worker.js
 * Procesa archivos Excel en segundo plano.
 */
importScripts('../vendor/xlsx.full.min.js');

self.onmessage = function(e) {
    const { data, options } = e.data;
    try {
        const workbook = XLSX.read(data, {
            type: 'array',
            cellDates: true,
            cellNF: false,
            cellText: false,
            raw: false
        });

        const result = { sheets: [] };

        workbook.SheetNames.forEach(name => {
            const sheet = workbook.Sheets[name];
            
            // Lógica de Merged Cells
            if (sheet['!merges']) {
                sheet['!merges'].forEach(merge => {
                    const val = sheet[XLSX.utils.encode_cell(merge.s)];
                    if (val) {
                        for (let r = merge.s.r; r <= merge.e.r; r++) {
                            for (let c = merge.s.c; c <= merge.e.c; c++) {
                                const ref = XLSX.utils.encode_cell({r, c});
                                if (!sheet[ref]) sheet[ref] = { ...val };
                            }
                        }
                    }
                });
            }

            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
            if (rows.length > 0) {
                result.sheets.push({ name, rows });
            }
        });

        self.postMessage({ success: true, result });
    } catch (err) {
        self.postMessage({ success: false, error: err.message });
    }
};
