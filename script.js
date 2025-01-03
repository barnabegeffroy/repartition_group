function processFile() {
  const fileInput = document.getElementById("tsvFile");
  if (fileInput.files.length === 0) {
    alert("Veuillez sélectionner un fichier TSV.");
    return;
  }

  const file = fileInput.files[0];
  const reader = new FileReader();

  reader.onload = function (event) {
    const tsvContent = event.target.result;

    // Traitement du contenu TSV :
    // 1. Séparer par ligne.
    // 2. Séparer chaque ligne par la tabulation.
    // 3. Enlever les guillemets autour des cellules.
    const rows = tsvContent
      .split("\n") // Sépare le fichier par ligne
      .map(
        (row) =>
          row
            .split("\t") // Sépare la ligne par tabulation
            .map((cell) => cell.replace(/(^"|"$)/g, "").trim()) // Retire les guillemets autour des cellules et enlève les espaces inutiles
      );

    findEvents(rows); // Traitez les données après les avoir séparées
  };

  reader.readAsText(file);
}

function findEvents(rows) {
  const eventsStartIndex = rows[1].indexOf("events");
  const nbrvoeuxIndex = rows[2].indexOf("nbr_voeux");
  const NBR_VOEUX = rows[3][nbrvoeuxIndex];
  const nbrEventsIndex = rows[2].indexOf("nbr_events");
  const NBR_SORTIES = rows[3][nbrEventsIndex];
  const headers = rows[2];
  const nameIndex = headers.indexOf("nom");
  const eventsBrut = headers.slice(eventsStartIndex);

  const eventsList = processEventList(eventsBrut);

  // Variables pour stocker les classements
  const eventsOrdonnes = {};

  // Récupérer les préférences de chaque utilisateurice
  for (let i = 3; i < rows.length; i++) {
    const row = rows[i];
    if (row.length > 0) {
      const webformSerial = row[0];

      var voeuxList = [];

      const voeuxBrut = row
        .slice(autresStartIndex)
        .map((item) => (item === "" ? "-1-" : parseInt(item, 10)));

      for (let j = 0; j < NBR_VOEUX; j++) {
        var autreEventIndex = voeuxBrut.indexOf(j);
        if (autreEventIndex != -1) {
          voeuxList.push(eventsList[autreEventIndex]);
        }
      }
      eventsOrdonnes[webformSerial] = voeuxList;
    }
  }

  // Récupérer les infos de chaque utilisateur
  const personnes = processUserDetails(rows, nameIndex, eventsStartIndex);


  let scoreAutres = {};
  let assignmentsAutres = {};
  let waitingListAutres = {};

  eventsList.forEach((event) => {
    waitingListAutres[event.name] = []; // Initialisation des listes d'attente pour chaque événement
  });

  // Initialiser les scores et les assignations pour "autres"
  for (let participantId in eventsOrdonnes) {
    scoreAutres[participantId] = 0; // Score initial des participants
    assignmentsAutres[participantId] = []; // Assignation des événements "autres"
  }

  // Parcours des voeux de chaque participant pour "autres"
  for (let voeuxNum = 0; voeuxNum < NBR_VOEUX; voeuxNum++) {
    const sortedParticipants = Object.keys(scoreAutres)
      .filter(
        (participantId) => assignmentsAutres[participantId].length < NBR_SORTIES
      )
      .sort((a, b) => scoreAutres[a] - scoreAutres[b]);

    for (let participantId of sortedParticipants) {
      try {
        let event = eventsOrdonnes[participantId][voeuxNum];
        let eventName = event.name;

        let eventCapacity = eventsList[event.index].capacity;
        let minimumCapacity = personnes[participantId].binome == 1 ? 1 : 0;
        if (eventCapacity > minimumCapacity) {
          if (assignmentsAutres[participantId].length < NBR_SORTIES) {
            assignmentsAutres[participantId].push(eventName);
            eventsList[event.index].capacity -=
              personnes[participantId].binome == 1 ? 2 : 1;
            scoreAutres[participantId] += NBR_VOEUX - voeuxNum;
          }
        } else {
          waitingListAutres[eventName].push(participantId);
        }
      } catch (error) {
        // Gérer les erreurs si nécessaire
      }
    }
  }

  let participantsByEventAutres = {};

  Object.keys(assignmentsAutres).forEach((participantId) => {
    assignmentsAutres[participantId].forEach((eventName) => {
      if (!participantsByEventAutres[eventName]) {
        participantsByEventAutres[eventName] = [];
      }
      participantsByEventAutres[eventName].push(participantId);
    });
  });

  const headersXLS = [
    "Nom",
    "Prénom",
    "Coordonnées",
    "Binôme",
    "Coordonnées Binôme",
  ];

  for (let i = 1; i <= NBR_SORTIES; i++) {
    headersXLS.push(`Choix - ${i}`);
  }

  const data = [headersXLS]; // Première ligne : les en-têtes

  Object.entries(personnes).forEach(([id, personne]) => {
    const row = [
      personne.nom,
      personne.prenom,
      personne.coordonnees,
      personne.binome == 1 ? "Oui" : "Non",
      personne.binome == 1 ? personne.coordonnees_binome : "",
      ...(assignmentsAutres[id] || [])
        .concat(Array(NBR_SORTIES).fill(""))
        .slice(0, NBR_SORTIES),
    ];
    data.push(row);
  });

  generateExcelFile(data, "repartitions_par_personnes.xlsx", "downloadLink");

  generateRecapInscriptionExcel(
    participantsByEventAutres,
    "recap_inscriptions_par_evenements.xlsx",
    "downloadLinkEventRecap"
  );

  generateRecapInscriptionExcel(
    ...waitingListAutres,
    "liste_attente.xlsx",
    "downloadLinkWaitingList"
  );

  function generateRecapInscriptionExcel(list, fileName, id) {
    const data = []; // Initialisation des données pour le fichier Excel

    // Ajouter les titres des événements sur la première ligne
    const headerRow = Object.keys(list);
    data.push(headerRow);

    // Préparer les informations des participants
    const maxParticipantsPerEvent = Math.max(
      ...Object.values(list).map((participants) => participants.length)
    );

    // Pour chaque ligne, ajouter les participants sous chaque événement
    for (let i = 0; i < maxParticipantsPerEvent; i++) {
      const row = headerRow.map((eventName) => {
        const participants = list[eventName] || [];
        if (participants[i]) {
          const person = personnes[participants[i]];
          return formatParticipantInfo(person);
        }
        return ""; // Cellule vide si aucun participant
      });
      data.push(row);
    }

    // Générer le fichier Excel
    generateExcelFile(data, fileName, id);
  }
}

function processEventList(eventsBrut) {
  return eventsBrut.map((eventBrut) => {
    const [index, name, capacity] = eventBrut.split("-");
    return {
      index: parseInt(index) - 1,
      name,
      capacity: parseInt(capacity),
    };
  });
}

function processUserPreferences(preferences, events) {
  return preferences.map((pref) => {
    if (pref === "") return { name: "-1-" };
    const eventIndex = parseInt(pref, 10);
    return events[eventIndex] || {};
  });
}

function processUserDetails(rows, nameIndex, eventsStartIndex) {
  const personnes = {};
  for (let i = 3; i < rows.length; i++) {
    const row = rows[i];
    if (row.length > 0) {
      const webformSerial = row[0];
      const [nom, prenom, coordonnees, binome, coordonnees_binome] = row.slice(
        nameIndex,
        eventsStartIndex
      );
      personnes[webformSerial] = {
        nom,
        prenom,
        coordonnees,
        binome,
        coordonnees_binome,
      };
    }
  }
  return personnes;
}

// Fonction utilitaire pour formater les informations des participants
function formatParticipantInfo(person) {
  const binomeInfo =
    person.binome == 1
      ? ` \nBinôme : ${person.binome} \nCoordonnées Binôme : ${person.coordonnees_binome}`
      : "";
  return `${person.nom} ${person.prenom} \nCoordonnées : ${person.coordonnees}${binomeInfo}`;
}

function generateExcelFile(data, fileName, downloadLinkId) {
  const ExcelJS = window.ExcelJS; // Charger ExcelJS
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Feuille 1");

  // Ajouter des données et des styles
  data.forEach((row, rowIndex) => {
    const rowObj = worksheet.addRow(row);
    rowObj.eachCell((cell) => {
      if (
        cell.value &&
        typeof cell.value === "string" &&
        cell.value.includes("\n")
      ) {
        // Remplacer \n par sa version compatible Excel
        cell.value = cell.value.replace(/\n/g, "\n"); // Garder tel quel, ExcelJS gère cela
        cell.alignment = { wrapText: true }; // Activer le wrapText pour afficher correctement
      }
    });

    if (rowIndex === 0) {
      // Style pour la première ligne (en-têtes)
      rowObj.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: "FFFFFF" } };
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "4CAF50" },
        };
        cell.alignment = {
          vertical: "middle",
          horizontal: "center",
          wrapText: true,
        };
      });
    }
  });

  // Ajuster la largeur des colonnes automatiquement
  worksheet.columns.forEach((column) => {
    let maxLength = 0;
    column.eachCell({ includeEmpty: true }, (cell) => {
      const cellValue = cell.value ? cell.value.toString() : "";
      if (cellValue.includes("\n")) {
        // Diviser en lignes et prendre la longueur maximale
        const lines = cellValue.split("\n");
        const longestLine = Math.max(...lines.map((line) => line.length));
        maxLength = Math.max(maxLength, longestLine);
      } else {
        maxLength = Math.max(maxLength, cellValue.length);
      }
    });
    column.width = maxLength + 2; // +2 pour une marge visuelle
  });

  // Figer les en-têtes
  worksheet.views = [{ state: "frozen", ySplit: 1 }];

  // Télécharger le fichier Excel
  workbook.xlsx.writeBuffer().then((buffer) => {
    const blob = new Blob([buffer], { type: "application/octet-stream" });
    const downloadLink = document.getElementById(downloadLinkId);
    downloadLink.setAttribute("href", URL.createObjectURL(blob));
    downloadLink.setAttribute("download", fileName);
    downloadLink.style.display = "block";
  });
}
