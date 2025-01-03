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

    // Traitement du contenu TSV
    const rows = tsvContent
      .split("\n")
      .map((row) =>
        row
          .split("\t")
          .map((cell) => cell.replace(/(^"|"$)/g, "").trim())
      );

    try {
      findEvents(rows); // Traitez les données après les avoir séparées
    } catch (error) {
    }
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

  const premiumEvents = processEventList(premiumBrut);
  console.log(premiumEvents)
  const autresEvents = processEventList(autresBrut);

  const premiumEventsOrdonnes = {};
  const autresEventsOrdonnes = {};

  // Récupérer les préférences des utilisateurs pour premium et autres
  for (let i = 3; i < rows.length; i++) {
    const row = rows[i];
    if (row.length > 0) {
      const webformSerial = row[0];
      premiumEventsOrdonnes[webformSerial] = processUserPreferences(row.slice(premiumStartIndex, autresStartIndex), premiumEvents);
      autresEventsOrdonnes[webformSerial] = processUserPreferences(row.slice(autresStartIndex), autresEvents);
    }
  }

  // Traitement des données des utilisateurs
  const personnes = processUserDetails(rows, nameIndex, premiumStartIndex);
  
//   // Logique de distribution des événements aux utilisateurs
  const { scorePremium, assignmentsPremium, waitingListPremium } = assignEventsToUsers(premiumEvents, premiumEventsOrdonnes, personnes, 'premium');
  const { scoreAutres, assignmentsAutres, waitingListAutres } = assignEventsToUsers(autresEvents, autresEventsOrdonnes, personnes, 'autres');

  // Créez et exportez les fichiers Excel
  exportToExcel(personnes, assignmentsPremium, assignmentsAutres, waitingListPremium, waitingListAutres);
}

function processEventList(eventsBrut) {
  return eventsBrut.map((eventBrut) => {
    const [index, name, capacity] = eventBrut.split("-");
    return {
      index: parseInt(index) - 1,
      name,
      capacity: parseInt(capacity)
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

function processUserDetails(rows, nameIndex, premiumStartIndex) {
  const personnes = {};
  for (let i = 3; i < rows.length; i++) {
    const row = rows[i];
    if (row.length > 0) {
      const webformSerial = row[0];
      const [nom, prenom, coordonnees, binome, coordonnees_binome] = row.slice(nameIndex, premiumStartIndex);
      personnes[webformSerial] = { nom, prenom, coordonnees, binome, coordonnees_binome };
    }
  }
  return personnes;
}

function assignEventsToUsers(events, eventsOrdonnes, personnes, type) {
  const score = {};
  const assignments = {};
  const waitingList = {};

  events.forEach((event) => {
    waitingList[event.name] = []; // Initialisation des listes d'attente
  });

  Object.keys(eventsOrdonnes).forEach((participantId) => {
    score[participantId] = 0;
    assignments[participantId] = [];
  });

  for (let voeuxNum = 0; voeuxNum < NBR_VOEUX; voeuxNum++) {
    const sortedParticipants = Object.keys(score)
      .filter((participantId) => assignments[participantId].length < NBR_SORTIES)
      .sort((a, b) => score[a] - score[b]);

    for (let participantId of sortedParticipants) {
      try {
        const event = eventsOrdonnes[participantId][voeuxNum];
        const eventName = event.name;
        const eventCapacity = events[event.index].capacity;

        if (eventCapacity > 0) {
          if (assignments[participantId].length < NBR_SORTIES) {
            assignments[participantId].push(eventName);
            events[event.index].capacity--; // Décrémenter la capacité de l'événement
            score[participantId] += NBR_VOEUX - voeuxNum;
          }
        } else {
          waitingList[eventName].push(participantId);
        }
      } catch (error) {
      }
    }
  }

  return { score, assignments, waitingList };
}

function exportToExcel(personnes, assignmentsPremium, assignmentsAutres, waitingListPremium, waitingListAutres) {
  const headersXLS = ["Nom", "Prénom", "Coordonnées", "Binôme", "Coordonnées Binôme"];
  
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
      ...(assignmentsPremium[id] || []).concat(Array(NBR_SORTIES).fill("")).slice(0, NBR_SORTIES),
      ...(assignmentsAutres[id] || []).concat(Array(NBR_SORTIES).fill("")).slice(0, NBR_SORTIES),
    ];
    data.push(row);
  });

  generateExcelFile(data, "fichier5.xlsx", "downloadLink");
  generateRecapInscriptionExcel({ ...waitingListPremium, ...waitingListAutres }, "liste_attente.xlsx", "downloadLinkWaitingList");
}