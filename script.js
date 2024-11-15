document.getElementById('processCsv').addEventListener('click', () => {
    const fileInput = document.getElementById('csvFile');
    if (fileInput.files.length === 0) {
        alert("Veuillez sélectionner un fichier CSV !");
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = function (event) {
        const csvContent = event.target.result;

        // Traitement du CSV
        const rows = csvContent.split("\n").map(row => row.split(","));
        rows[0].push("Nouvelle colonne"); // Ajouter une colonne d'entête
        for (let i = 1; i < rows.length; i++) {
            rows[i].push("Valeur " + i); // Ajouter des valeurs
        }

        const newCsvContent = rows.map(row => row.join(",")).join("\n");

        // Générer un fichier téléchargeable
        const blob = new Blob([newCsvContent], { type: 'text/csv' });
        const url = URL.createObjectURL(blob);

        const downloadLink = document.getElementById('downloadLink');
        downloadLink.href = url;
        downloadLink.download = "fichier_modifié.csv";
        downloadLink.style.display = "inline-block";
        downloadLink.textContent = "Télécharger le CSV modifié";
    };

    reader.readAsText(file);
});
