const NBR_VOEUX = 4;
const NBR_SORTIES = 3;

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
  const premiumStartIndex = rows[1].indexOf("premium");
  const autresStartIndex = rows[1].indexOf("autres");
  const headers = rows[2];
  const nameIndex = headers.indexOf("nom");
  const premiumBrut = headers.slice(premiumStartIndex, autresStartIndex);
  const autresBrut = headers.slice(autresStartIndex);

  var premiumEvents = [];
  const autresEvents = [];

  // Créer les objets d'événements pour premium
  premiumBrut.forEach((eventBrut, _) => {
    const [index, name, capacity] = eventBrut.split("-");

    const eventObject = {
      index: parseInt(index) - 1,
      name,
      capacity: capacity,
    };
    premiumEvents.push(eventObject);
  });

  // Créer les objets d'événements pour autres
  autresBrut.forEach((eventBrut, _) => {
    const [index, name, capacity] = eventBrut.split("-");
    const eventObject = {
      index: parseInt(index) - 1,
      name,
      capacity: parseInt(capacity),
    };
    autresEvents.push(eventObject);
  });

  // Variables pour stocker les classements
  const premiumEventsOrdonnes = {};
  const autresEventsOrdonnes = {};

  // Récupérer les préférences de chaque utilisateur pour premium
  for (let i = 3; i < rows.length; i++) {
    const row = rows[i];
    if (row.length > 0) {
      const webformSerial = row[0];

      var voeuxPremium = [];
      var voeuxAutres = [];

      const voeuxPremiumBrut = row
        .slice(premiumStartIndex, autresStartIndex)
        .map((item) => (item === "" ? "-1-" : parseInt(item, 10)));

      const voeuxAutresBrut = row
        .slice(autresStartIndex)
        .map((item) => (item === "" ? "-1-" : parseInt(item, 10)));

      for (let j = 0; j < NBR_VOEUX; j++) {
        var premiumEventIndex = voeuxPremiumBrut.indexOf(j);
        var autreEventIndex = voeuxAutresBrut.indexOf(j);
        if (premiumEventIndex != -1) {
          voeuxPremium.push(premiumEvents[premiumEventIndex]);
          voeuxAutres.push(autresEvents[autreEventIndex]);
        }
      }

      // Remplacer l'utilisation de la Map par un objet JSON
      premiumEventsOrdonnes[webformSerial] = voeuxPremium;
      autresEventsOrdonnes[webformSerial] = voeuxAutres;
    }
  }

  // Récupérer les infos de chaque utilisateur
  var personnes = {};
  for (let i = 3; i < rows.length; i++) {
    const row = rows[i];
    if (row.length > 0) {
      const webformSerial = row[0];
      const [nom, prenom, coordonnees, binome, coordonnees_binome] = row.slice(
        nameIndex,
        premiumStartIndex
      );

      personnes[webformSerial] = {
        nom: nom,
        prenom: prenom,
        coordonnees: coordonnees,
        binome: binome,
        coordonnees_binome: coordonnees_binome,
      };
    }
  }

  // Assignation

  let scorePremium = {};
  let assignmentsPremium = {};
  let waitingListPremium = {};

  premiumEvents.forEach((event) => {
    waitingListPremium[event.name] = []; // Initialisation des listes d'attente pour chaque événement
  });

  // Initialiser les scores et les assignations
  for (let participantId in premiumEventsOrdonnes) {
    scorePremium[participantId] = 0; // Score initial des participants
    assignmentsPremium[participantId] = []; // Assignation des événements aux participants
  }

  // Parcours des voeux de chaque participant, de 1 à NBR_VOEUX
  for (let voeuxNum = 0; voeuxNum < NBR_VOEUX; voeuxNum++) {
    // Trier les participants par leur score (prioriser les plus petits scores)
    const sortedParticipants = Object.keys(scorePremium)
      .filter(
        (participantId) =>
          assignmentsPremium[participantId].length < NBR_SORTIES
      ) // Sélectionner ceux qui n'ont pas encore de score et n'ont pas atteint la limite de sorties
      .sort((a, b) => scorePremium[a] - scorePremium[b]); // Trier par score croissant

    for (let participantId of sortedParticipants) {
      try {
        let event = premiumEventsOrdonnes[participantId][voeuxNum];
        let eventName = event.name;

        // Vérifier la capacité de l'événement
        let eventCapacity = premiumEvents[event.index].capacity;
        let minimumCapacity = personnes[participantId].binome == 1 ? 1 : 0;

        if (eventCapacity > minimumCapacity) {
          // Si le participant n'a pas encore de sortie, l'assigner à cet événement
          if (assignmentsPremium[participantId].length < NBR_SORTIES) {
            assignmentsPremium[participantId].push(eventName);
            premiumEvents[event.index].capacity -=
              personnes[participantId].binome == 1 ? 2 : 1; // Réduire la capacité de l'événement
            scorePremium[participantId] += NBR_VOEUX - voeuxNum; // Calculer le score
          }
        } else {
          // Si l'événement est plein, l'ajouter à la liste d'attente
          waitingListPremium[eventName].push(participantId);
        }
      } catch (error) {
        // Gérer les erreurs si nécessaire
      }
    }
  }
  let participantsByEventPremium = {};

  // Parcourir les assignations pour remplir `participantsByEvent`
  Object.keys(assignmentsPremium).forEach((participantId) => {
    assignmentsPremium[participantId].forEach((eventName) => {
      if (!participantsByEventPremium[eventName]) {
        participantsByEventPremium[eventName] = [];
      }
      participantsByEventPremium[eventName].push(participantId);
    });
  });

  let scoreAutres = {};
  let assignmentsAutres = {};
  let waitingListAutres = {};

  autresEvents.forEach((event) => {
    waitingListAutres[event.name] = []; // Initialisation des listes d'attente pour chaque événement
  });

  // Initialiser les scores et les assignations pour "autres"
  for (let participantId in autresEventsOrdonnes) {
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
        let event = autresEventsOrdonnes[participantId][voeuxNum];
        let eventName = event.name;

        let eventCapacity = autresEvents[event.index].capacity;
        let minimumCapacity = personnes[participantId].binome == 1 ? 1 : 0;
        if (eventCapacity > minimumCapacity) {
          if (assignmentsAutres[participantId].length < NBR_SORTIES) {
            assignmentsAutres[participantId].push(eventName);
            autresEvents[event.index].capacity -=
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
    headersXLS.push(`Premium - ${i}`);
  }
  for (let i = 1; i <= NBR_SORTIES; i++) {
    headersXLS.push(`Autres - ${i}`);
  }

  const data = [headersXLS]; // Première ligne : les en-têtes

  Object.entries(personnes).forEach(([id, personne]) => {
    const row = [
      personne.nom,
      personne.prenom,
      personne.coordonnees,
      personne.binome == 1 ? "Oui" : "Non",
      personne.binome == 1 ? personne.coordonnees_binome : "",
      ...(assignmentsPremium[id] || [])
        .concat(Array(NBR_SORTIES).fill(""))
        .slice(0, NBR_SORTIES),
      ...(assignmentsAutres[id] || [])
        .concat(Array(NBR_SORTIES).fill(""))
        .slice(0, NBR_SORTIES),
    ];
    data.push(row);
  });

  generateExcelFile(data, "fichier5.xlsx", "downloadLink");

  generateRecapInscriptionExcel(
    { ...participantsByEventPremium, ...participantsByEventAutres },
    "recap_inscriptions.xlsx",
    "downloadLinkEventRecap"
  );

  generateRecapInscriptionExcel(
    { ...waitingListPremium, ...waitingListAutres },
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

  // Fonction utilitaire pour formater les informations des participants
  function formatParticipantInfo(person) {
    const lineBreak = `
`;
    const binomeInfo =
      person.binome == 1
        ? ` \nBinôme : ${person.binome} \nCoordonnées Binôme : ${person.coordonnees_binome}`
        : "";
    return `${person.nom} ${person.prenom} \nCoordonnées : ${person.coordonnees}${binomeInfo}`;
  }
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
