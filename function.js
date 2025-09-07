window.function = function (jsonData) {
    // On récupère la chaîne JSON de Glide
    const jsonString = jsonData.value ?? '[]';

    try {
        // On s'assure que le JSON est un tableau d'objets
        const data = JSON.parse(jsonString);
        if (!Array.isArray(data)) {
            return "Erreur: Le JSON doit être un tableau d'objets (commençant par [).";
        }

        // --- Début de la conversion ---
        
        // 1. SheetJS convertit le JSON en une feuille de calcul
        const worksheet = XLSX.utils.json_to_sheet(data);

        // 2. On crée un nouveau classeur
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Données');

        // 3. On génère le fichier Excel en format Base64
        const base64Excel = XLSX.write(workbook, { bookType: 'xlsx', type: 'base64' });

        // 4. On construit la "Data URL", qui est un lien contenant le fichier
        const dataUrl = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${base64Excel}`;

        // 5. On retourne le lien directement à Glide
        return dataUrl;

    } catch (error) {
        // En cas d'erreur, on renvoie le message d'erreur
        return `Erreur de conversion: ${error.message}`;
    }
}
