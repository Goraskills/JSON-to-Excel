window.function = function (jsonData) {
    // On récupère la chaîne JSON de Glide
    const jsonString = jsonData.value ?? '[]';

    try {
        // On s'assure que le JSON est un tableau d'objets
        const data = JSON.parse(jsonString);
        if (!Array.isArray(data)) {
            return "Erreur: Le JSON doit être un tableau d'objets (commençant par [).";
        }

        const worksheet = XLSX.utils.json_to_sheet(data);

        // AMÉLIORATION : On force le type des cellules numériques
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
            for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
                const cellAddress = XLSX.utils.encode_cell({ r: rowNum, c: colNum });
                const cell = worksheet[cellAddress];
                if (cell && typeof cell.v === 'number') {
                    cell.t = 'n';
                }
            }
        }
        
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Données');

        const base64Excel = XLSX.write(workbook, { bookType: 'xlsx', type: 'base64' });

        const dataUrl = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${base64Excel}`;

        return dataUrl;

    } catch (error) {
        return `Erreur de conversion: ${error.message}`;
    }
}
