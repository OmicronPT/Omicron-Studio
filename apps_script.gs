// ============================================================
//  OMICRON STUDIO — Sistema Prenotazioni
//  Google Apps Script — Backend Completo
//  Versione 1.0
// ============================================================
//
//  SETUP INIZIALE:
//  1. Apri il tuo Google Sheet
//  2. Vai su Estensioni > Apps Script
//  3. Incolla questo codice
//  4. Modifica le costanti nella sezione CONFIGURAZIONE
//  5. Salva e clicca su "Esegui" > setupSheet() per creare i fogli
//  6. Vai su Deploy > Nuova distribuzione > Web App
//     - Esegui come: Me
//     - Chi ha accesso: Chiunque
//  7. Copia l'URL generato — è il tuo backend endpoint
//
// ============================================================

// ============================================================
//  CONFIGURAZIONE — modifica questi valori
// ============================================================

const CONFIG = {
  SPREADSHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId(),
  
  // Il tuo numero WhatsApp per ricevere notifiche admin (formato internazionale, es. 393331234567)
  ADMIN_WHATSAPP: "39XXXXXXXXXX",
  
  // API Key CallMeBot — registrati su https://www.callmebot.com/blog/free-api-whatsapp-messages/
  CALLMEBOT_API_KEY: "XXXXXXXX",
  
  // Nome del tuo studio
  STUDIO_NAME: "Omicron Studio",
  
  // Ore prima dell'allenamento per inviare il reminder
  ORE_REMINDER: 20,
  
  // Soglia lezioni rimanenti per avviso scadenza pacchetto
  SOGLIA_AVVISO_LEZIONI: 2,
};

// ============================================================
//  NOMI DEI FOGLI
// ============================================================

const SHEETS = {
  CLIENTI: "Clienti",
  SESSIONI: "Sessioni",
  PRENOTAZIONI: "Prenotazioni",
  PACCHETTI: "Pacchetti",
};

// ============================================================
//  SETUP — crea la struttura del Google Sheet
// ============================================================

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Crea o ottieni i fogli
  _getOrCreateSheet(ss, SHEETS.CLIENTI, [
    "ID", "Nome", "Cognome", "Telefono", "Email",
    "Tipo Pacchetto", "Lezioni Totali", "Lezioni Rimanenti",
    "Data Inizio", "Data Scadenza", "Stato", "Token Prenotazione"
  ]);

  _getOrCreateSheet(ss, SHEETS.SESSIONI, [
    "ID", "Data", "Ora Inizio", "Ora Fine",
    "Posti Totali", "Posti Occupati", "Stato", "Note"
  ]);

  _getOrCreateSheet(ss, SHEETS.PRENOTAZIONI, [
    "ID", "ID Cliente", "Nome Cliente", "ID Sessione",
    "Data Sessione", "Ora", "Data Prenotazione", "Stato", "Lezione Scalata"
  ]);

  _getOrCreateSheet(ss, SHEETS.PACCHETTI, [
    "ID", "Nome", "N° Lezioni", "Durata Giorni", "Prezzo", "Note"
  ]);

  // Inserisce pacchetti di esempio
  const shPacchetti = ss.getSheetByName(SHEETS.PACCHETTI);
  if (shPacchetti.getLastRow() <= 1) {
    const esempi = [
      ["PKG001", "Singola", 1, 30, 40, "Lezione singola"],
      ["PKG002", "Pacchetto 10", 10, 90, 350, "10 lezioni, validità 90 giorni"],
      ["PKG003", "Pacchetto 20", 20, 180, 600, "20 lezioni, validità 6 mesi"],
      ["PKG004", "Abbonamento Mensile", 12, 30, 180, "Circa 3 sessioni/settimana"],
    ];
    shPacchetti.getRange(2, 1, esempi.length, esempi[0].length).setValues(esempi);
  }

  SpreadsheetApp.getUi().alert("Setup completato! Fogli creati correttamente.");
}

function _getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground("#1a1a2e")
      .setFontColor("#ffffff")
      .setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ============================================================
//  WEB APP — gestisce le richieste HTTP dal frontend
// ============================================================

function doGet(e) {
  const action = e.parameter.action;
  const token = e.parameter.token;

  try {
    switch (action) {
      case "getSessioni":
        return _jsonResponse(getSessioniDisponibili());
      case "getCliente":
        return _jsonResponse(getClienteByToken(token));
      case "getPrenotazioniCliente":
        return _jsonResponse(getPrenotazioniCliente(token));
      default:
        return _jsonResponse({ error: "Azione non riconosciuta" });
    }
  } catch (err) {
    return _jsonResponse({ error: err.message });
  }
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const action = data.action;

  try {
    switch (action) {
      case "prenota":
        return _jsonResponse(prenotaSessione(data.token, data.idSessione));
      case "cancella":
        return _jsonResponse(cancellaPrenotazione(data.token, data.idPrenotazione));
      default:
        return _jsonResponse({ error: "Azione non riconosciuta" });
    }
  } catch (err) {
    return _jsonResponse({ error: err.message });
  }
}

function _jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  SESSIONI
// ============================================================

function getSessioniDisponibili() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.SESSIONI);
  const data = sheet.getDataRange().getValues();
  const oggi = new Date();
  oggi.setHours(0, 0, 0, 0);

  const sessioni = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dataSessione = new Date(row[1]);
    const postiOccupati = parseInt(row[5]) || 0;
    const postiTotali = parseInt(row[4]) || 3;
    const stato = row[6];

    if (dataSessione >= oggi && stato !== "Cancellata" && postiOccupati < postiTotali) {
      sessioni.push({
        id: row[0],
        data: Utilities.formatDate(dataSessione, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        dataLeggibile: Utilities.formatDate(dataSessione, Session.getScriptTimeZone(), "EEEE d MMMM yyyy"),
        oraInizio: row[2],
        oraFine: row[3],
        postiLiberi: postiTotali - postiOccupati,
        postiTotali: postiTotali,
      });
    }
  }
  return sessioni;
}

// Aggiunge una nuova sessione (usato dal pannello admin)
function aggiungiSessione(data, oraInizio, oraFine, postiTotali, note) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.SESSIONI);
  const id = "SES" + Date.now();
  sheet.appendRow([id, data, oraInizio, oraFine, postiTotali || 3, 0, "Disponibile", note || ""]);
  return { success: true, id };
}

// ============================================================
//  CLIENTI
// ============================================================

function getClienteByToken(token) {
  const clienti = _getAllClienti();
  const cliente = clienti.find(c => c.token === token);
  if (!cliente) return { error: "Token non valido" };
  return cliente;
}

function _getAllClienti() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.CLIENTI);
  const data = sheet.getDataRange().getValues();
  const clienti = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    clienti.push({
      id: row[0],
      nome: row[1],
      cognome: row[2],
      telefono: row[3],
      email: row[4],
      tipoPacchetto: row[5],
      lezioniTotali: row[6],
      lezioniRimanenti: row[7],
      dataInizio: row[8] ? Utilities.formatDate(new Date(row[8]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
      dataScadenza: row[9] ? Utilities.formatDate(new Date(row[9]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
      stato: row[10],
      token: row[11],
      _riga: i + 1,
    });
  }
  return clienti;
}

// Genera un token univoco per il link di prenotazione del cliente
function generaTokenCliente(idCliente) {
  return Utilities.base64Encode(idCliente + "_" + Date.now()).replace(/[^a-zA-Z0-9]/g, "").substring(0, 16);
}

// Aggiunge un nuovo cliente (usato dal pannello admin)
function aggiungiCliente(nome, cognome, telefono, email, idPacchetto) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shClienti = ss.getSheetByName(SHEETS.CLIENTI);
  const shPacchetti = ss.getSheetByName(SHEETS.PACCHETTI);

  const pacchetti = shPacchetti.getDataRange().getValues();
  const pacchetto = pacchetti.find(p => p[0] === idPacchetto);
  if (!pacchetto) return { error: "Pacchetto non trovato" };

  const id = "CLI" + Date.now();
  const token = generaTokenCliente(id);
  const oggi = new Date();
  const scadenza = new Date();
  scadenza.setDate(scadenza.getDate() + parseInt(pacchetto[3]));

  shClienti.appendRow([
    id, nome, cognome, telefono, email,
    pacchetto[1], pacchetto[2], pacchetto[2],
    Utilities.formatDate(oggi, Session.getScriptTimeZone(), "yyyy-MM-dd"),
    Utilities.formatDate(scadenza, Session.getScriptTimeZone(), "yyyy-MM-dd"),
    "Attivo", token
  ]);

  return { success: true, id, token };
}

// ============================================================
//  PRENOTAZIONI
// ============================================================

function prenotaSessione(token, idSessione) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cliente = getClienteByToken(token);
  if (cliente.error) return cliente;

  // Controlla stato cliente
  if (cliente.stato !== "Attivo") return { error: "Il tuo abbonamento non è attivo." };
  if (cliente.lezioniRimanenti <= 0) return { error: "Non hai lezioni rimanenti nel tuo pacchetto." };

  const oggi = new Date();
  if (cliente.dataScadenza && new Date(cliente.dataScadenza) < oggi) {
    return { error: "Il tuo pacchetto è scaduto. Contatta lo studio per rinnovarlo." };
  }

  // Controlla sessione
  const shSessioni = ss.getSheetByName(SHEETS.SESSIONI);
  const dataSessioni = shSessioni.getDataRange().getValues();
  let sessRiga = -1;
  let sessData = null;

  for (let i = 1; i < dataSessioni.length; i++) {
    if (dataSessioni[i][0] === idSessione) {
      sessRiga = i + 1;
      sessData = dataSessioni[i];
      break;
    }
  }

  if (sessRiga === -1) return { error: "Sessione non trovata." };

  const postiOccupati = parseInt(sessData[5]) || 0;
  const postiTotali = parseInt(sessData[4]) || 3;
  if (postiOccupati >= postiTotali) return { error: "Sessione al completo." };
  if (sessData[6] === "Cancellata") return { error: "Questa sessione è stata cancellata." };

  // Controlla se già prenotato
  const shPrenotazioni = ss.getSheetByName(SHEETS.PRENOTAZIONI);
  const dataPrenotazioni = shPrenotazioni.getDataRange().getValues();
  for (let i = 1; i < dataPrenotazioni.length; i++) {
    if (dataPrenotazioni[i][1] === cliente.id && dataPrenotazioni[i][3] === idSessione && dataPrenotazioni[i][7] !== "Cancellata") {
      return { error: "Sei già prenotato per questa sessione." };
    }
  }

  // Tutto ok — crea prenotazione
  const idPrenotazione = "PRE" + Date.now();
  const dataSessione = new Date(sessData[1]);
  const dataLeggibile = Utilities.formatDate(dataSessione, Session.getScriptTimeZone(), "EEEE d MMMM");

  shPrenotazioni.appendRow([
    idPrenotazione, cliente.id, `${cliente.nome} ${cliente.cognome}`,
    idSessione, sessData[1], sessData[2],
    Utilities.formatDate(oggi, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"),
    "Confermata", "Sì"
  ]);

  // Scala lezione
  const shClienti = ss.getSheetByName(SHEETS.CLIENTI);
  shClienti.getRange(cliente._riga, 8).setValue(cliente.lezioniRimanenti - 1);

  // Aggiorna posti occupati
  shSessioni.getRange(sessRiga, 6).setValue(postiOccupati + 1);

  // Manda WhatsApp di conferma
  const msgCliente = `✅ *${CONFIG.STUDIO_NAME}*\nCiao ${cliente.nome}! La tua prenotazione è confermata.\n📅 ${dataLeggibile}\n🕐 ${sessData[2]} - ${sessData[3]}\nLezioni rimanenti: ${cliente.lezioniRimanenti - 1}`;
  _inviaWhatsApp(cliente.telefono, msgCliente);

  // Avviso scadenza pacchetto
  if (cliente.lezioniRimanenti - 1 <= CONFIG.SOGLIA_AVVISO_LEZIONI && cliente.lezioniRimanenti - 1 > 0) {
    const msgAvviso = `⚠️ *${CONFIG.STUDIO_NAME}*\nCiao ${cliente.nome}, ti ricordiamo che hai solo *${cliente.lezioniRimanenti - 1} lezioni* rimanenti nel tuo pacchetto. Contattaci per rinnovarlo!`;
    _inviaWhatsApp(cliente.telefono, msgAvviso);
  }

  return { success: true, idPrenotazione, lezioniRimanenti: cliente.lezioniRimanenti - 1 };
}

function cancellaPrenotazione(token, idPrenotazione) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cliente = getClienteByToken(token);
  if (cliente.error) return cliente;

  const shPrenotazioni = ss.getSheetByName(SHEETS.PRENOTAZIONI);
  const data = shPrenotazioni.getDataRange().getValues();
  let pRiga = -1;
  let pData = null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === idPrenotazione && data[i][1] === cliente.id) {
      pRiga = i + 1;
      pData = data[i];
      break;
    }
  }

  if (pRiga === -1) return { error: "Prenotazione non trovata." };
  if (pData[7] === "Cancellata") return { error: "Prenotazione già cancellata." };

  // Controlla limite cancellazione (es. almeno 2 ore prima)
  const dataSessione = new Date(pData[4]);
  const ora = pData[5].toString();
  const [h, m] = ora.split(":").map(Number);
  dataSessione.setHours(h, m, 0, 0);
  const diff = (dataSessione - new Date()) / (1000 * 60 * 60);
  if (diff < 2) return { error: "Non è possibile cancellare meno di 2 ore prima della sessione." };

  // Cancella
  shPrenotazioni.getRange(pRiga, 8).setValue("Cancellata");

  // Restituisci lezione
  const shClienti = ss.getSheetByName(SHEETS.CLIENTI);
  shClienti.getRange(cliente._riga, 8).setValue(cliente.lezioniRimanenti + 1);

  // Riduci posti occupati nella sessione
  const shSessioni = ss.getSheetByName(SHEETS.SESSIONI);
  const dataSessioni = shSessioni.getDataRange().getValues();
  for (let i = 1; i < dataSessioni.length; i++) {
    if (dataSessioni[i][0] === pData[3]) {
      const occupati = Math.max(0, parseInt(dataSessioni[i][5]) - 1);
      shSessioni.getRange(i + 1, 6).setValue(occupati);
      break;
    }
  }

  const msg = `❌ *${CONFIG.STUDIO_NAME}*\nCiao ${cliente.nome}, la tua prenotazione del ${Utilities.formatDate(new Date(pData[4]), Session.getScriptTimeZone(), "d MMMM")} alle ${pData[5]} è stata cancellata.\nLa lezione è stata riaccreditata al tuo pacchetto.`;
  _inviaWhatsApp(cliente.telefono, msg);

  return { success: true, lezioniRimanenti: cliente.lezioniRimanenti + 1 };
}

function getPrenotazioniCliente(token) {
  const cliente = getClienteByToken(token);
  if (cliente.error) return cliente;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PRENOTAZIONI);
  const data = sheet.getDataRange().getValues();
  const oggi = new Date();
  oggi.setHours(0, 0, 0, 0);

  const prenotazioni = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === cliente.id && data[i][7] !== "Cancellata") {
      const dataS = new Date(data[i][4]);
      if (dataS >= oggi) {
        prenotazioni.push({
          id: data[i][0],
          dataSessione: Utilities.formatDate(dataS, Session.getScriptTimeZone(), "yyyy-MM-dd"),
          dataLeggibile: Utilities.formatDate(dataS, Session.getScriptTimeZone(), "EEEE d MMMM yyyy"),
          ora: data[i][5],
          stato: data[i][7],
        });
      }
    }
  }
  return prenotazioni;
}

// ============================================================
//  REMINDER AUTOMATICI — da attivare come trigger giornaliero
// ============================================================

function inviaReminderGiornalieri() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPrenotazioni = ss.getSheetByName(SHEETS.PRENOTAZIONI);
  const shClienti = ss.getSheetByName(SHEETS.CLIENTI);

  const prenotazioni = shPrenotazioni.getDataRange().getValues();
  const clienti = _getAllClienti();
  const domani = new Date();
  domani.setDate(domani.getDate() + 1);
  domani.setHours(0, 0, 0, 0);
  const dopodomani = new Date(domani);
  dopodomani.setDate(dopodomani.getDate() + 1);

  for (let i = 1; i < prenotazioni.length; i++) {
    const row = prenotazioni[i];
    if (row[7] !== "Confermata") continue;

    const dataS = new Date(row[4]);
    dataS.setHours(0, 0, 0, 0);

    if (dataS >= domani && dataS < dopodomani) {
      const cliente = clienti.find(c => c.id === row[1]);
      if (!cliente) continue;

      const msg = `🔔 *Reminder ${CONFIG.STUDIO_NAME}*\nCiao ${cliente.nome}! Ti ricordiamo il tuo allenamento di domani:\n📅 ${row[5]} - ${row[5]}\nA domani! 💪`;
      _inviaWhatsApp(cliente.telefono, msg);
    }
  }
}

// Controlla pacchetti in scadenza (da eseguire settimanalmente)
function controllaPackettiInScadenza() {
  const clienti = _getAllClienti();
  const oggi = new Date();
  const tra7giorni = new Date();
  tra7giorni.setDate(oggi.getDate() + 7);

  clienti.forEach(c => {
    if (c.stato !== "Attivo") return;
    const scadenza = new Date(c.dataScadenza);
    if (scadenza <= tra7giorni && scadenza >= oggi) {
      const msg = `⚠️ *${CONFIG.STUDIO_NAME}*\nCiao ${cliente.nome}! Il tuo pacchetto scade il *${c.dataScadenza}*.\nHai ancora *${c.lezioniRimanenti} lezioni* da utilizzare.\nContattaci per rinnovare! 💪`;
      _inviaWhatsApp(c.telefono, msg);
    }
  });
}

// ============================================================
//  WHATSAPP — CallMeBot
// ============================================================

function _inviaWhatsApp(numero, messaggio) {
  try {
    const url = `https://api.callmebot.com/whatsapp.php?phone=${numero}&text=${encodeURIComponent(messaggio)}&apikey=${CONFIG.CALLMEBOT_API_KEY}`;
    UrlFetchApp.fetch(url);
  } catch (e) {
    Logger.log("Errore WhatsApp: " + e.message);
  }
}

// ============================================================
//  SETUP TRIGGER — esegui una volta sola dopo il deploy
// ============================================================

function setupTriggers() {
  // Rimuovi trigger esistenti
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Reminder giornaliero alle 18:00
  ScriptApp.newTrigger("inviaReminderGiornalieri")
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .create();

  // Controllo pacchetti ogni lunedì mattina
  ScriptApp.newTrigger("controllaPackettiInScadenza")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  SpreadsheetApp.getUi().alert("Trigger configurati correttamente!");
}
