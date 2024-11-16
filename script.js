const NBR_VOEUX = 4;

document.getElementById("processCsv").addEventListener("click", () => {
  const fileInput = document.getElementById("csvFile");
  if (fileInput.files.length === 0) {
    alert("Veuillez sélectionner un fichier CSV.");
    return;
  }

  const file = fileInput.files[0];
  const reader = new FileReader();

  reader.onload = function (event) {
    const csvContent = event.target.result;
    const rows = csvContent
      .split("\n")
      .map((row) => row.split(",").map((cell) => cell.trim()));

    findEvents(rows);
  };

  reader.readAsText(file);
});

function findEvents(rows) {
  const premiumStartIndex = rows[1].indexOf("premium");
  const autresStartIndex = rows[1].indexOf("autres");
  const headers = rows[2];
  const nameIndex = headers.indexOf("nom");
  const premiumBrut = headers.slice(premiumStartIndex, autresStartIndex);
  const autresBrut = headers.slice(autresStartIndex);

  var premiumEvents = [];
  const autresEvents = [];
  var users = {};

  // Créer les objets d'événements pour premium
  premiumBrut.forEach((eventBrut, _) => {
    const [index, name, capacity] = eventBrut.split("-");
    const eventObject = {
      index: parseInt(index),
      name,
      capacity: parseInt(capacity),
    };
    premiumEvents.push(eventObject);
  });

  console.log(JSON.stringify(premiumEvents));

  // Créer les objets d'événements pour autres
  autresBrut.forEach((eventBrut, _) => {
    const [index, name, capacity] = eventBrut.split("-");
    const eventObject = {
      index: parseInt(index),
      name,
      capacity: parseInt(capacity),
    };
    autresEvents.push(eventObject);
  });

  // Récupérer les infos de chaque utilisateur
  for (let i = 3; i < rows.length; i++) {
    const row = rows[i];
    if (row.length > 0) {
      const webformSerial = row[0];
      const [nom, prenom, coordonnees, binome, coordonnees_binome] = row.slice(
        nameIndex,
        premiumStartIndex
      );

      users[webformSerial] = {
        nom: nom,
        prenom: prenom,
        coordonnees: coordonnees,
        binome: binome,
        coordonnees_binome: coordonnees_binome,
      };
    }
  }

  // Variables pour stocker les classements
  const autresEventsOrder = {};
  const premiumEventsOrder = {};

  // Récupérer les préférences de chaque utilisateur pour premium
  for (let i = 3; i < rows.length; i++) {
    const row = rows[i];
    if (row.length > 0) {
      const webformSerial = row[0];

      var voeuxPremium = [];

      const voeuxPremiumBrut = row
        .slice(premiumStartIndex, autresStartIndex)
        .map((item) => (item === "" ? "-1-" : parseInt(item, 10)));

      for (let j = 0; j < NBR_VOEUX; j++) {
        var eventIndex = voeuxPremiumBrut.indexOf(j);
        if (eventIndex != -1) {
          voeuxPremium.push(premiumEvents[eventIndex]);
        }
      }

      // Remplacer l'utilisation de la Map par un objet JSON
      premiumEventsOrder[webformSerial] = voeuxPremium;
    }
  }
  const NBR_SORTIES = 3; // Limite du nombre de sorties par participant

  let score = {};
  let assignments = {};
  let waitingList = {};

  premiumEvents.forEach((event) => {
    waitingList[event.name] = []; // Initialisation des listes d'attente pour chaque événement
  });

  // Initialiser les scores et les assignations
  for (let participantId in premiumEventsOrder) {
    score[participantId] = 0; // Score initial des participants
    assignments[participantId] = []; // Assignation des événements aux participants
  }

  // Parcours des voeux de chaque participant, de 1 à NBR_VOEUX
  for (let voeuxNum = 0; voeuxNum < NBR_VOEUX; voeuxNum++) {
    // Trier les participants par leur score (prioriser les plus petits scores)
    const sortedParticipants = Object.keys(score)
      .filter(
        (participantId) => assignments[participantId].length < NBR_SORTIES
      ) // Sélectionner ceux qui n'ont pas encore de score et n'ont pas atteint la limite de sorties
      .sort((a, b) => score[a] - score[b]); // Trier par score croissant

    for (let participantId of sortedParticipants) {
      try {
        let event = premiumEventsOrder[participantId][voeuxNum];
        let eventName = event.name;

        // Vérifier la capacité de l'événement
        let eventCapacity = premiumEvents[event.index].capacity;
        let minimumCapacity = users[participantId].binome == 1 ? 1 : 0;
        if (eventCapacity > minimumCapacity) {
          // Si le participant n'a pas encore de sortie, l'assigner à cet événement
          if (assignments[participantId].length < NBR_SORTIES) {
            assignments[participantId].push(eventName);
            premiumEvents[event.index].capacity -=
              users[participantId].binome == 1 ? 2 : 1; // Réduire la capacité de l'événement
            score[participantId] += 4 - voeuxNum; // Calculer le score
          }
        } else {
          // Si l'événement est plein, l'ajouter à la liste d'attente
          waitingList[eventName].push(participantId);
        }
      } catch (error) {
        // Gérer les erreurs si nécessaire
      }
    }
  }
  let participantsByEvent = {};

  // Parcourir les assignations pour remplir `participantsByEvent`
  Object.keys(assignments).forEach((participantId) => {
    assignments[participantId].forEach((eventName) => {
      if (!participantsByEvent[eventName]) {
        participantsByEvent[eventName] = [];
      }
      participantsByEvent[eventName].push(participantId);
    });
  });

  //   console.log(JSON.stringify(users));
  //   console.log(JSON.stringify(participantsByEvent));
  //   console.log(JSON.stringify(assignments));
  //   console.log(JSON.stringify(waitingList));
  //   console.log(JSON.stringify(score));

  // Construction des lignes du CSV
  const csvHeader = [
    "Nom",
    "Prénom",
    "Coordonnées",
    "Binôme",
    "Coordonnées Binôme",
  ];
  for (let i = 1; i <= NBR_SORTIES; i++) {
    csvHeader.push(`Visite ${i}`);
  }
  const csvRows = [csvHeader];

  Object.entries(users).forEach(([id, personne]) => {
    const row = [
      personne.nom,
      personne.prenom,
      personne.coordonnees,
      personne.binome,
      personne.coordonnees_binome,
      ...(assignments[id] || [])
        .concat(Array(NBR_SORTIES).fill(""))
        .slice(0, NBR_SORTIES),
    ];
    csvRows.push(row);
  });

  // Conversion en format CSV
  const csvContent = csvRows.map((row) => row.join(",")).join("\n");

  // Ajout du lien de téléchargement
  const downloadLink = document.getElementById("downloadLink");
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  downloadLink.setAttribute("href", url);
  downloadLink.setAttribute("download", "personnes_sorties.csv");
  downloadLink.style.display = "block";
}
