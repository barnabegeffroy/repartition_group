function processFile() {
  try {
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
      // 2. Séparer chaque ligne par tabulation.
      // 3. Enlever les guillemets et espaces superflus autour des cellules.
      const rows = tsvContent
        .split("\n") // Sépare le fichier par ligne
        .map(
          (row) =>
            row
              .split("\t") // Sépare la ligne par tabulation
              .map((cell) => cell.replace(/(^"|"$)/g, "").trim()) // Retire les guillemets autour des cellules et enlève les espaces inutiles
        );
      try {
        findEvents(rows); // Traitez les données après les avoir séparées
      } catch (error) {
        displayError(error);
      }
    };

    reader.readAsText(file);
  } catch (error) {
    displayError(error);
  }
}

// Fonction utilitaire pour afficher les erreurs
function displayError(error) {
  const errorMessage = document.getElementById("errorMessage");
  errorMessage.textContent = error.message;
  errorMessage.style.display = "block";
  console.error(error);
}

function findEvents(rows) {
  const eventsStartIndex = rows[1].indexOf("events");
  const nbrVoeuxIndex = rows[2].indexOf("nbr_voeux");
  const NBR_VOEUX = rows[3][nbrVoeuxIndex];
  const nbrEventsIndex = rows[2].indexOf("nbr_events");
  const NBR_SORTIES = rows[3][nbrEventsIndex];
  const headers = rows[2];
  const nameIndex = headers.indexOf("nom");
  const eventsBrut = headers.slice(eventsStartIndex);

  const eventsList = processEventList(eventsBrut);

  // Variables pour stocker les classements
  const eventsOrdonnes = {};

  // Récupérer les infos de chaque utilisateur
  const personnes = processUserDetails(rows, nameIndex, eventsStartIndex);

  // Récupérer les préférences de chaque utilisateurice
  for (let i = 3; i < rows.length; i++) {
    const row = rows[i];
    if (row.length > 0) {
      const webformSerial = row[0];

      // Vérifier si le webformSerial existe dans personnes
      if (personnes[webformSerial]) {
        var voeuxList = [];

        const voeuxBrut = row
          .slice(eventsStartIndex)
          .map((item) => (item === "" ? "-1-" : parseInt(item, 10)));

        for (let j = 0; j < NBR_VOEUX; j++) {
          var autreEventIndex = voeuxBrut.indexOf(j);
          if (autreEventIndex != -1) {
            voeuxList.push(eventsList[autreEventIndex]);
          }
        }
        // Ajouter les voeux à eventsOrdonnes uniquement si personnes[webformSerial] existe
        eventsOrdonnes[webformSerial] = voeuxList;
      }
    }
  }

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
      if (eventsOrdonnes[participantId][voeuxNum]) {
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
      }
    }
  }

  // Tri de la liste d'attente par score pour chaque événement
  Object.keys(waitingListAutres).forEach((eventName) => {
    // Trier les participants dans la liste d'attente par leur score
    waitingListAutres[eventName].sort((a, b) => {
      return scoreAutres[a] - scoreAutres[b]; // Tri croissant par score
    });
  });

  // Parcours des participants pour afficher ceux ayant un score de 0
  const participantsAvecScoreZero = Object.keys(scoreAutres).filter(
    (participantId) => scoreAutres[participantId] === 0
  );

  console.log("Participants avec un score de 0 :");
  participantsAvecScoreZero.forEach((participantId) => {
    console.log(personnes[participantId]);
  });

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
    "Email",
    "Type de chambre",
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
      personne.email,
      personne.chambre,
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
    waitingListAutres,
    "liste_attente.xlsx",
    "downloadLinkWaitingList"
  );
  generateMailingListExcel(
    participantsByEventAutres,
    "mailing_list.xlsx",
    "downloadLinkMailingList"
  );

  function generateRecapInscriptionExcel(list, fileName, id) {
    const data = []; // Initialisation des données pour le fichier Excel

    // Fonction pour extraire la date du nom de l'événement et la convertir en un objet Date
    function extractDateFromEventName(eventName) {
      const datePattern = /^(\d{1,2})\/(\d{1,2})/; // Modèle pour extraire la date (jour/mois)
      var name =
        eventName.charAt(0) === " "
          ? (eventName = eventName.substring(1))
          : eventName;
      const match = name.match(datePattern);
      if (match) {
        const day = match[1].padStart(2, "0"); // Ajoute un zéro si le jour est à un seul chiffre
        const month = match[2].padStart(2, "0"); // Ajoute un zéro si le mois est à un seul chiffre
        // Crée un objet Date à partir de la date extraite
        return new Date(`${month}/${day}/2025`); // On fixe l'année pour simplifier la comparaison
      }
      return null; // Retourne null si aucun match
    }

    // Trier les événements par date
    const headerRow = Object.keys(list).sort((eventA, eventB) => {
      const dateA = extractDateFromEventName(eventA);
      const dateB = extractDateFromEventName(eventB);

      // Si l'une des dates est invalide, on la place après les autres
      if (!dateA) return 1;
      if (!dateB) return -1;

      return dateA - dateB; // Comparaison chronologique des objets Date
    });

    data.push(headerRow); // Ajouter les titres triés des événements

    // Préparer les informations des participants
    const maxParticipantsPerEvent = Math.max(
      ...Object.values(list).map((participants) => participants.length)
    );

    // Ajouter le nombre de places restantes
    data.push(
      headerRow.map((eventName) => {
        const event = eventsList.find((event) => event.name === eventName);
        if (event) return "Places restantes : " + event.capacity;
        return "";
      })
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

  function generateMailingListExcel(list, fileName, id) {
    const data = []; // Initialisation des données pour le fichier Excel

    // Trier les événements par date
    const headerRow = Object.keys(list);

    data.push(headerRow); // Ajouter les titres triés des événements

    // Préparer les informations des participants
    const maxParticipantsPerEvent = Math.max(
      ...Object.values(list).map((participants) => participants.length)
    );

    // Ajouter le nombre de places restantes
    data.push(
      headerRow.map((eventName) => {
        const event = eventsList.find((event) => event.name === eventName);
        if (event) return "Places restantes : " + event.capacity;
        return "";
      })
    );

    // Pour chaque ligne, ajouter les participants sous chaque événement
    for (let i = 0; i < maxParticipantsPerEvent; i++) {
      const row = headerRow.map((eventName) => {
        const participants = list[eventName] || [];
        if (participants[i]) {
          const person = personnes[participants[i]];
          return person.email === ""
            ? formatParticipantInfo(person, false)
            : person.email;
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
  const uniqueEntries = new Set();
  const namePrenomMap = {};
  const conflicts = [];

  for (let i = 3; i < rows.length; i++) {
    const row = rows[i];
    if (row.length > 0) {
      const webformSerial = row[0];

      if (!webformSerial.trim()) {
        continue;
      }

      const [
        nom,
        prenom,
        coordonnees,
        email,
        chambre,
        binome,
        coordonnees_binome,
      ] = row.slice(nameIndex, eventsStartIndex);

      // Normalisation des champs
      const normalizedNom = normalizeString(nom.split(/[\s-]/)[0]);
      const normalizedPrenom = normalizeString(prenom).charAt(0);
      const normalizedCoordonnees = coordonnees.replace(/\s+/g, ""); // Suppression des espaces
      const normalizedBinome = parseInt(binome, 10); // Conversion en entier

      const uniqueKey = `${normalizedNom}|${normalizedPrenom}|${normalizedCoordonnees}|${normalizedBinome}`;
      const namePrenomKey = `${normalizedNom}|${normalizedPrenom}`;

      // Vérifier si cette entrée est déjà présente
      if (uniqueEntries.has(uniqueKey)) {
        console.log(`Doublon détecté : ${nom} ${prenom}`);
        continue; // Ignorer cette ligne si elle est un doublon exact
      }

      // Vérifier si le nom et prénom existent déjà avec des coordonnées ou binômes différents
      if (namePrenomMap[namePrenomKey]) {
        const existingEntry = namePrenomMap[namePrenomKey];

        if (
          existingEntry.coordonnees !== normalizedCoordonnees ||
          existingEntry.binome !== normalizedBinome
        ) {
          console.warn(
            `Conflit détecté pour ${nom} ${prenom} : différentes coordonnées ou binômes.`
          );
          // Ajout du conflit à la liste
          conflicts.push(
            `Conflit pour ${nom} ${prenom}: coordonnées ou binôme différents.`
          );
        }
      } else {
        // Ajouter cette entrée pour suivi futur
        namePrenomMap[namePrenomKey] = {
          coordonnees: normalizedCoordonnees,
          binome: normalizedBinome,
        };
      }

      // Ajouter à la liste des entrées uniques
      uniqueEntries.add(uniqueKey);

      // Stocker les détails de la personne
      personnes[webformSerial] = {
        nom,
        prenom,
        coordonnees,
        email,
        chambre,
        binome,
        coordonnees_binome,
      };
    }
  }

  // Si des conflits ont été détectés, on les affiche dans l'HTML
  if (conflicts.length > 0) {
    const conflictList = document.getElementById("conflictList");
    conflictList.innerHTML = conflicts
      .map((conflict) => `<li>${conflict}</li>`)
      .join("");
    document.getElementById("conflictMessages").style.display = "block";
  }

  return personnes;
}

// Fonction utilitaire pour normaliser une chaîne (sans accents et sans casse)
const normalizeString = (str) =>
  str
    .toLowerCase() // Minuscule
    .normalize("NFD") // Décompose les accents
    .replace(/[\u0300-\u036f]/g, "") // Supprime les accents
    .replace(/[\s-]+/g, ""); // Supprime les espaces et les tirets

// Fonction utilitaire pour formater les informations des participants
function formatParticipantInfo(person, displayDetails = true) {
  const binomeInfo =
    person.binome == 1
      ? ` \nBinôme : ${person.binome} \nCoordonnées Binôme : ${person.coordonnees_binome}`
      : "";
  const details = displayDetails
    ? `\n Email :${person.email}${binomeInfo}\n Type chambre ${person.chambre}`
    : "";
  return `${person.nom} ${person.prenom} \nCoordonnées : ${person.coordonnees} ${details}`;
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
