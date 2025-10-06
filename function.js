window.function = function (jsonData, fileName) {
    // --- On récupère les valeurs de Glide ---
    const jsonString = jsonData.value ?? '[]';
    const name = fileName.value ?? 'export';

    // --- On crée le HTML qui sera affiché dans le Web Embed ---
    const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            body { font-family: sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background-color: #f0f2f5; }
            button { background-color: #1D6F42; color: white; border: none; padding: 12px 24px; border-radius: 6px; font-size: 16px; cursor: pointer; transition: background-color 0.3s; }
            button:hover { background-color: #144E2E; }
            button:disabled { background-color: #ccc; }
        </style>
    </head>
    <body>
        <button id="downloadButton">Générer et Télécharger Excel</button>

        <script>
            document.getElementById('downloadButton').addEventListener('click', function() {
                const button = this;
                button.innerText = 'Génération en cours...';
                button.disabled = true;

                try {
                    const data = JSON.parse(${JSON.stringify(jsonString)});
                    const worksheet = XLSX.utils.json_to_sheet(data);
                    const workbook = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(workbook, worksheet, 'Données');

                    // On déclenche le téléchargement
                    XLSX.writeFile(workbook, \\${name}.xlsx\);

                    // --- CORRECTION : On réinitialise le bouton immédiatement ---
                    // On change le texte pour confirmer le succès
                    button.innerText = 'Terminé ! Téléchargement lancé.';
                    
                    // On réactive le bouton après un court instant pour éviter les doubles clics
                    setTimeout(() => {
                        button.innerText = 'Générer et Télécharger Excel';
                        button.disabled = false;
                    }, 2000);


                } catch (error) {
                    button.innerText = 'Erreur ! Vérifiez le JSON.';
                    // En cas d'erreur, on réactive aussi le bouton
                    button.disabled = false; 
                    console.error(error);
                }
            });
        <\/script>
    </body>
    </html>
    `;

    // On retourne le HTML encodé pour que Glide puisse l'afficher
    return "data:text/html;charset=utf-8," + encodeURIComponent(htmlContent);
}
