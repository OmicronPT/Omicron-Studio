// ============================================================
//  OMICRON STUDIO — Sistema Prenotazioni
//  Google Apps Script — Backend v5.0
//  Slot automatici 30min, 9:00-20:00, lun-ven
// ============================================================

const CONFIG = {
  STUDIO_NAME:        "Omicron Studio",
  ADMIN_WHATSAPP:     "39XXXXXXXXXX",
  CALLMEBOT_API_KEY:  "",
  ORE_REMINDER:       18,
  SOGLIA_AVVISO:      2,
  MAX_CONTEMPORANEI:  3,     // max persone contemporanee
  DURATA_SLOT_MIN:    30,    // durata slot in minuti
  DURATA_SESSION_MIN: 60,    // durata sessione in minuti (= 2 slot)
  ORE_CANCELLAZIONE:  2,
  ORA_INIZIO:         9,     // 09:00
  ORA_FINE:           20,    // 20:00 (ultimo slot inizia alle 19:00 per durata 60min)
  SETTIMANE_AVANTI:   4,     // quante settimane mostrare
  ADMIN_PASSWORD:     "omicron2024",
};

// ──────────────────────────────────────────────────────────
//  SETUP
// ──────────────────────────────────────────────────────────
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _creaFoglio(ss, "Clienti", [
    "ID","Nome","Cognome","Telefono","Email",
    "Pacchetto","Lezioni Totali","Lezioni Rimanenti",
    "Data Inizio","Data Scadenza","Stato","Token",
    "Data Nascita","Sesso","Indirizzo","Instagram","Facebook","Note Anamnesi"
  ]);
  _creaFoglio(ss, "Prenotazioni", [
    "ID","ID Cliente","Nome Cliente","Data","Ora Inizio","Ora Fine",
    "Data Prenotazione","Stato"
  ]);
  _creaFoglio(ss, "Blocchi", [
    "ID","Data","Ora Inizio","Ora Fine","Motivo","Creato Il"
  ]);
  _creaFoglio(ss, "Pacchetti", [
    "ID","Nome","Lezioni","Durata Giorni","Prezzo","Note"
  ]);
  const shP = ss.getSheetByName("Pacchetti");
  if (shP.getLastRow() <= 1) {
    shP.getRange(2,1,4,6).setValues([
      ["PKG001","Singola",1,30,40,""],
      ["PKG002","Pacchetto 10",10,90,350,""],
      ["PKG003","Pacchetto 20",20,180,600,""],
      ["PKG004","Mensile",12,30,180,"~3 sessioni/settimana"],
    ]);
  }
  SpreadsheetApp.getUi().alert("Setup completato!");
}

function _creaFoglio(ss, nome, headers) {
  let sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  if (sh.getLastRow() === 0) {
    const r = sh.getRange(1,1,1,headers.length);
    r.setValues([headers]);
    r.setBackground("#1a1a2e").setFontColor("#ffffff").setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

// ──────────────────────────────────────────────────────────
//  WEB APP
// ──────────────────────────────────────────────────────────
function doGet(e) {
  const action   = e.parameter.action   || "";
  const token    = e.parameter.token    || "";
  const callback = e.parameter.callback || "";
  const pw       = e.parameter.pw       || "";

  let result;
  try {
    switch (action) {
      case "getSlotDisponibili":   result = getSlotDisponibili(token);       break;
      case "getCliente":           result = getClienteByToken(token);        break;
      case "getPrenotazioni":      result = getPrenotazioniCliente(token);   break;
      case "adminLogin":           result = adminLogin(pw);                  break;
      case "adminDashboard":       result = adminDashboard(pw);              break;
      case "adminClienti":         result = adminClienti(pw);                break;
      case "adminCalendario":      result = adminCalendario(pw);             break;
      case "adminBlocchi":         result = adminBlocchi(pw);                break;
      case "adminPacchetti":       result = getPacchetti();                  break;
      default: result = { error: "Azione non riconosciuta: " + action };
    }
  } catch(err) {
    result = { error: err.message };
  }

  const json = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback+"("+json+")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let data = {};
  try { data = JSON.parse(e.postData.contents); } catch(_) {}
  const pw = data.pw || "";

  let result;
  try {
    switch (data.action) {
      case "prenota":                   result = prenotaSlot(data.token, data.data, data.oraInizio);    break;
      case "cancella":                  result = cancellaPrenotazione(data.token, data.idPrenotazione); break;
      case "adminAddBlocco":            result = adminAddBlocco(pw, data);                              break;
      case "adminDelBlocco":            result = adminDelBlocco(pw, data.id);                           break;
      case "adminAddCliente":           result = adminAddCliente(pw, data);                             break;
      case "adminEditCliente":          result = adminEditCliente(pw, data);                            break;
      case "adminDelCliente":           result = adminDelCliente(pw, data.id);                          break;
      case "adminCancellaPrenotazione": result = adminCancellaPrenotazione(pw, data.id);                break;
      default: result = { error: "Azione non riconosciuta" };
    }
  } catch(err) {
    result = { error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ──────────────────────────────────────────────────────────
//  LOGICA SLOT
// ──────────────────────────────────────────────────────────

// Converte "HH:MM" in minuti dall'inizio della giornata
function _oreToMin(ora) {
  const [h,m] = ora.split(":").map(Number);
  return h * 60 + m;
}

// Converte minuti in "HH:MM"
function _minToOre(min) {
  return String(Math.floor(min/60)).padStart(2,"0") + ":" + String(min%60).padStart(2,"0");
}

// Genera tutti gli slot teorici del giorno (array di "HH:MM")
function _slotDelGiorno() {
  const slots = [];
  const fine = (CONFIG.ORA_FINE - CONFIG.DURATA_SESSION_MIN / 60) * 60; // ultimo inizio possibile
  for (let m = CONFIG.ORA_INIZIO * 60; m <= fine; m += CONFIG.DURATA_SLOT_MIN) {
    slots.push(_minToOre(m));
  }
  return slots;
}

// Verifica se una data è un giorno feriale (lun-ven)
function _isFeriale(dateStr) {
  const d = new Date(dateStr + "T12:00:00");
  const dow = d.getDay(); // 0=dom, 6=sab
  return dow >= 1 && dow <= 5;
}

// Conta quante prenotazioni attive si sovrappongono a un dato slot
// Un slot occupa [oraInizio, oraInizio + DURATA_SESSION_MIN)
function _conteggioSovrapposti(dateStr, oraInizio, prenotazioni) {
  const startMin = _oreToMin(oraInizio);
  const endMin   = startMin + CONFIG.DURATA_SESSION_MIN;

  return prenotazioni.filter(p => {
    if (p.data !== dateStr || p.stato === "Cancellata") return false;
    const pStart = _oreToMin(p.oraInizio);
    const pEnd   = pStart + CONFIG.DURATA_SESSION_MIN;
    // Sovrapposizione se gli intervalli si intersecano
    return pStart < endMin && pEnd > startMin;
  }).length;
}

// Verifica se uno slot è bloccato
function _isBloccat(dateStr, oraInizio, blocchi) {
  const startMin = _oreToMin(oraInizio);
  const endMin   = startMin + CONFIG.DURATA_SESSION_MIN;

  return blocchi.some(b => {
    if (b.data !== dateStr) return false;
    // Blocco giornata intera
    if (!b.oraInizio && !b.oraFine) return true;
    const bStart = _oreToMin(b.oraInizio || "00:00");
    const bEnd   = b.oraFine ? _oreToMin(b.oraFine) : 24*60;
    return bStart < endMin && bEnd > startMin;
  });
}

// Restituisce gli slot disponibili per le prossime SETTIMANE_AVANTI settimane
// Se token è passato, esclude slot già prenotati da quel cliente
function getSlotDisponibili(token) {
  const tz   = Session.getScriptTimeZone();
  const oggi = new Date();
  oggi.setHours(0,0,0,0);

  const fine = new Date(oggi);
  fine.setDate(fine.getDate() + CONFIG.SETTIMANE_AVANTI * 7);

  const prenotazioni = _tuttePrenotazioni();
  const blocchi      = _tuttiBlocchi();
  const slotsGiorno  = _slotDelGiorno();

  // Prenotazioni del cliente corrente (per marcarle)
  let preClienteSet = new Set();
  if (token) {
    const cliente = getClienteByToken(token);
    if (!cliente.error) {
      prenotazioni
        .filter(p => p.idCliente === cliente.id && p.stato === "Confermata")
        .forEach(p => preClienteSet.add(p.data + "_" + p.oraInizio));
    }
  }

  const risultato = [];
  for (let d = new Date(oggi); d < fine; d.setDate(d.getDate()+1)) {
    const dateStr = d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");
    if (!_isFeriale(dateStr)) continue;

    const slotsGiornata = [];
    for (const ora of slotsGiorno) {
      if (_isBloccat(dateStr, ora, blocchi)) continue;
      const count = _conteggioSovrapposti(dateStr, ora, prenotazioni);
      if (count >= CONFIG.MAX_CONTEMPORANEI) continue;

      slotsGiornata.push({
        ora,
        oraFine: _minToOre(_oreToMin(ora) + CONFIG.DURATA_SESSION_MIN),
        postiLiberi: CONFIG.MAX_CONTEMPORANEI - count,
        prenotato: preClienteSet.has(dateStr + "_" + ora),
      });
    }

    if (slotsGiornata.length > 0) {
      risultato.push({
        data: dateStr,
        dataLabel: Utilities.formatDate(d, tz, "EEEE d MMMM yyyy"),
        slots: slotsGiornata,
      });
    }
  }
  return risultato;
}

// ──────────────────────────────────────────────────────────
//  PRENOTAZIONI
// ──────────────────────────────────────────────────────────
function prenotaSlot(token, data, oraInizio) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const cliente = getClienteByToken(token);
  if (cliente.error)              return cliente;
  if (cliente.stato !== "Attivo") return { error: "Abbonamento non attivo." };
  if (cliente.lezioniRim <= 0)    return { error: "Nessuna lezione rimanente." };
  if (cliente.dataScad && new Date(cliente.dataScad) < new Date()) return { error: "Pacchetto scaduto." };

  const prenotazioni = _tuttePrenotazioni();
  const blocchi      = _tuttiBlocchi();

  if (_isBloccat(data, oraInizio, blocchi)) return { error: "Orario non disponibile." };

  const count = _conteggioSovrapposti(data, oraInizio, prenotazioni);
  if (count >= CONFIG.MAX_CONTEMPORANEI) return { error: "Slot al completo." };

  // Già prenotato?
  const già = prenotazioni.find(p =>
    p.idCliente === cliente.id && p.data === data &&
    p.oraInizio === oraInizio && p.stato !== "Cancellata"
  );
  if (già) return { error: "Hai già prenotato questo slot." };

  const id     = "PRE" + Date.now();
  const tz     = Session.getScriptTimeZone();
  const oraFine= _minToOre(_oreToMin(oraInizio) + CONFIG.DURATA_SESSION_MIN);
  const now    = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm");

  ss.getSheetByName("Prenotazioni").appendRow([
    id, cliente.id, cliente.nome+" "+cliente.cognome,
    data, oraInizio, oraFine, now, "Confermata"
  ]);

  // Scala lezione
  ss.getSheetByName("Clienti").getRange(cliente._riga, 8).setValue(cliente.lezioniRim - 1);

  // WhatsApp
  const dl = Utilities.formatDate(new Date(data+"T12:00:00"), tz, "EEEE d MMMM");
  _wa(cliente.telefono, `✅ *${CONFIG.STUDIO_NAME}*\nCiao ${cliente.nome}! Prenotazione confermata.\n📅 ${dl} ore ${oraInizio}\nLezioni rimanenti: ${cliente.lezioniRim-1}`);
  if (cliente.lezioniRim-1 <= CONFIG.SOGLIA_AVVISO && cliente.lezioniRim-1 > 0)
    _wa(cliente.telefono, `⚠️ *${CONFIG.STUDIO_NAME}*\nCiao ${cliente.nome}, rimangono solo *${cliente.lezioniRim-1} lezioni*. Contattaci!`);

  return { ok: true, idPrenotazione: id, lezioniRimanenti: cliente.lezioniRim-1 };
}

function cancellaPrenotazione(token, idPrenotazione) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const cliente = getClienteByToken(token);
  if (cliente.error) return cliente;

  const shPre = ss.getSheetByName("Prenotazioni");
  const rows  = shPre.getDataRange().getValues();
  let pRiga=-1, p=null;
  for (let i=1; i<rows.length; i++) {
    if (rows[i][0]===idPrenotazione && rows[i][1]===cliente.id) { pRiga=i+1; p=rows[i]; break; }
  }
  if (!p)                  return { error: "Prenotazione non trovata." };
  if (p[7]==="Cancellata") return { error: "Già cancellata." };

  const dataS = new Date(p[3]+"T"+p[4]);
  if ((dataS - new Date()) < CONFIG.ORE_CANCELLAZIONE * 3600000)
    return { error: `Impossibile cancellare meno di ${CONFIG.ORE_CANCELLAZIONE}h prima.` };

  shPre.getRange(pRiga, 8).setValue("Cancellata");
  ss.getSheetByName("Clienti").getRange(cliente._riga, 8).setValue(cliente.lezioniRim + 1);

  const tz = Session.getScriptTimeZone();
  _wa(cliente.telefono, `❌ *${CONFIG.STUDIO_NAME}*\nPrenotazione del ${Utilities.formatDate(new Date(p[3]+"T12:00:00"),tz,"d MMMM")} alle ${p[4]} cancellata. Lezione riaccreditata.`);
  return { ok: true, lezioniRimanenti: cliente.lezioniRim+1 };
}

function getPrenotazioniCliente(token) {
  const cliente = getClienteByToken(token);
  if (cliente.error) return cliente;

  const tz   = Session.getScriptTimeZone();
  const oggi = new Date(); oggi.setHours(0,0,0,0);

  return _tuttePrenotazioni()
    .filter(p => p.idCliente===cliente.id && p.stato!=="Cancellata" && new Date(p.data+"T12:00:00")>=oggi)
    .map(p => ({
      id: p.id,
      data: p.data,
      dataLabel: Utilities.formatDate(new Date(p.data+"T12:00:00"), tz, "EEEE d MMMM yyyy"),
      oraInizio: p.oraInizio,
      oraFine:   p.oraFine,
    }));
}

function _tuttePrenotazioni() {
  const rows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Prenotazioni").getDataRange().getValues();
  const out  = [];
  for (let i=1; i<rows.length; i++) {
    const r = rows[i]; if (!r[0]) continue;
    out.push({
      id: r[0], idCliente: r[1], nomeCliente: r[2],
      data: r[3] ? _dateToStr(new Date(r[3])) : "",
      oraInizio: r[4], oraFine: r[5],
      stato: r[7], _riga: i+1,
    });
  }
  return out;
}

function _dateToStr(d) {
  return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");
}

// ──────────────────────────────────────────────────────────
//  BLOCCHI
// ──────────────────────────────────────────────────────────
function _tuttiBlocchi() {
  const rows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Blocchi").getDataRange().getValues();
  const out  = [];
  for (let i=1; i<rows.length; i++) {
    const r = rows[i]; if (!r[0]) continue;
    out.push({
      id: r[0],
      data: r[1] ? _dateToStr(new Date(r[1])) : "",
      oraInizio: r[2]||"", oraFine: r[3]||"",
      motivo: r[4]||"", _riga: i+1,
    });
  }
  return out;
}

function adminBlocchi(pw) {
  _checkAdmin(pw);
  return _tuttiBlocchi();
}

function adminAddBlocco(pw, data) {
  _checkAdmin(pw);
  if (!data.data) return { error: "Data obbligatoria" };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const id = "BLK" + Date.now();
  const tz = Session.getScriptTimeZone();
  ss.getSheetByName("Blocchi").appendRow([
    id, data.data, data.oraInizio||"", data.oraFine||"",
    data.motivo||"",
    Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm")
  ]);
  return { ok: true, id };
}

function adminDelBlocco(pw, id) {
  _checkAdmin(pw);
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Blocchi"), rows=sh.getDataRange().getValues();
  for (let i=1; i<rows.length; i++) {
    if (rows[i][0]===id) { sh.deleteRow(i+1); return { ok: true }; }
  }
  return { error: "Blocco non trovato" };
}

// ──────────────────────────────────────────────────────────
//  ADMIN — DASHBOARD
// ──────────────────────────────────────────────────────────
function adminDashboard(pw) {
  _checkAdmin(pw);
  const tz   = Session.getScriptTimeZone();
  const oggi = new Date(); oggi.setHours(0,0,0,0);
  const tra7 = new Date(oggi); tra7.setDate(tra7.getDate()+7);

  const clienti      = _tuttiClienti();
  const prenotazioni = _tuttePrenotazioni();

  const oggiStr  = _dateToStr(oggi);
  const preOggi  = prenotazioni.filter(p => p.data===oggiStr && p.stato==="Confermata");
  const inScadenza   = clienti.filter(c => { if(c.stato!=="Attivo")return false; const s=new Date(c.dataScad); return s>=oggi&&s<=tra7; });
  const pocheLezioni = clienti.filter(c => c.stato==="Attivo"&&c.lezioniRim<=CONFIG.SOGLIA_AVVISO&&c.lezioniRim>0);

  // Conta slot unici oggi
  const slotsOggi = new Set(preOggi.map(p=>p.oraInizio)).size;

  return {
    ok: true,
    stats: {
      clientiAttivi:    clienti.filter(c=>c.stato==="Attivo").length,
      slotsOggi,
      prenotazioniOggi: preOggi.length,
      inScadenza:       inScadenza.length,
    },
    prenotazioniOggi: preOggi,
    inScadenza,
    pocheLezioni,
  };
}

// ──────────────────────────────────────────────────────────
//  ADMIN — CALENDARIO
// ──────────────────────────────────────────────────────────
function adminCalendario(pw) {
  _checkAdmin(pw);
  const tz   = Session.getScriptTimeZone();
  const oggi = new Date(); oggi.setHours(0,0,0,0);
  const fine = new Date(oggi); fine.setDate(fine.getDate() + CONFIG.SETTIMANE_AVANTI * 7);

  const prenotazioni = _tuttePrenotazioni();
  const blocchi      = _tuttiBlocchi();
  const slotsGiorno  = _slotDelGiorno();
  const risultato    = [];

  for (let d = new Date(oggi); d < fine; d.setDate(d.getDate()+1)) {
    const dateStr = _dateToStr(d);
    if (!_isFeriale(dateStr)) continue;

    const bloccatoTutto = blocchi.some(b => b.data===dateStr && !b.oraInizio && !b.oraFine);
    const preGiorno = prenotazioni.filter(p => p.data===dateStr && p.stato==="Confermata");

    // Raggruppa prenotazioni per slot
    const slotsInfo = slotsGiorno.map(ora => {
      const count     = _conteggioSovrapposti(dateStr, ora, prenotazioni);
      const bloccato  = _isBloccat(dateStr, ora, blocchi);
      const preSlot   = prenotazioni.filter(p =>
        p.data===dateStr && p.stato==="Confermata" &&
        _oreToMin(p.oraInizio) <= _oreToMin(ora) &&
        _oreToMin(p.oraInizio) + CONFIG.DURATA_SESSION_MIN > _oreToMin(ora)
      );
      return { ora, count, bloccato, clienti: preSlot.map(p=>p.nomeCliente) };
    }).filter(s => s.count > 0 || s.bloccato);

    risultato.push({
      data: dateStr,
      dataLabel: Utilities.formatDate(new Date(dateStr+"T12:00:00"), tz, "EEEE d MMMM yyyy"),
      bloccatoTutto,
      nPrenotazioni: preGiorno.length,
      slots: slotsInfo,
      blocchi: blocchi.filter(b=>b.data===dateStr),
    });
  }
  return risultato;
}

// ──────────────────────────────────────────────────────────
//  AUTH
// ──────────────────────────────────────────────────────────
function _checkAdmin(pw) {
  if (pw !== CONFIG.ADMIN_PASSWORD) throw new Error("Password non valida");
}
function adminLogin(pw) {
  return pw===CONFIG.ADMIN_PASSWORD ? {ok:true} : {error:"Password non valida"};
}

// ──────────────────────────────────────────────────────────
//  CLIENTI
// ──────────────────────────────────────────────────────────
function adminClienti(pw) { _checkAdmin(pw); return _tuttiClienti(); }

function adminAddCliente(pw, data) {
  _checkAdmin(pw);
  if (!data.nome||!data.cognome||!data.idPacchetto) return { error: "Dati mancanti" };
  return aggiungiCliente(data);
}

function adminEditCliente(pw, data) {
  _checkAdmin(pw);
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Clienti"), rows=sh.getDataRange().getValues();
  for (let i=1; i<rows.length; i++) {
    if (rows[i][0]===data.id) {
      if (data.nome         !== undefined) sh.getRange(i+1,2).setValue(data.nome);
      if (data.cognome      !== undefined) sh.getRange(i+1,3).setValue(data.cognome);
      if (data.telefono     !== undefined) sh.getRange(i+1,4).setValue(data.telefono);
      if (data.email        !== undefined) sh.getRange(i+1,5).setValue(data.email);
      if (data.lezioniRim   !== undefined) sh.getRange(i+1,8).setValue(data.lezioniRim);
      if (data.dataScad     !== undefined) sh.getRange(i+1,10).setValue(data.dataScad);
      if (data.stato        !== undefined) sh.getRange(i+1,11).setValue(data.stato);
      if (data.dataNascita  !== undefined) sh.getRange(i+1,13).setValue(data.dataNascita);
      if (data.sesso        !== undefined) sh.getRange(i+1,14).setValue(data.sesso);
      if (data.indirizzo    !== undefined) sh.getRange(i+1,15).setValue(data.indirizzo);
      if (data.instagram    !== undefined) sh.getRange(i+1,16).setValue(data.instagram);
      if (data.facebook     !== undefined) sh.getRange(i+1,17).setValue(data.facebook);
      if (data.noteAnamnesi !== undefined) sh.getRange(i+1,18).setValue(data.noteAnamnesi);
      return { ok: true };
    }
  }
  return { error: "Cliente non trovato" };
}

function adminDelCliente(pw, id) {
  _checkAdmin(pw);
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Clienti"), rows=sh.getDataRange().getValues();
  for (let i=1; i<rows.length; i++) { if(rows[i][0]===id){ sh.getRange(i+1,11).setValue("Eliminato"); return {ok:true}; } }
  return { error: "Cliente non trovato" };
}

function adminCancellaPrenotazione(pw, id) {
  _checkAdmin(pw);
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Prenotazioni"), rows=sh.getDataRange().getValues();
  for (let i=1; i<rows.length; i++) {
    if (rows[i][0]===id) {
      sh.getRange(i+1,8).setValue("Cancellata");
      const c=_tuttiClienti().find(x=>x.id===rows[i][1]);
      if (c) ss.getSheetByName("Clienti").getRange(c._riga,8).setValue(c.lezioniRim+1);
      return { ok: true };
    }
  }
  return { error: "Prenotazione non trovata" };
}

function getClienteByToken(token) {
  if (!token) return { error: "Token mancante" };
  const c = _tuttiClienti().find(x=>x.token===token);
  return c || { error: "Token non valido" };
}

function _tuttiClienti() {
  const ss=SpreadsheetApp.getActiveSpreadsheet(), rows=ss.getSheetByName("Clienti").getDataRange().getValues(), tz=Session.getScriptTimeZone(), out=[];
  for (let i=1; i<rows.length; i++) {
    const r=rows[i]; if(!r[0]) continue;
    out.push({
      id:r[0], nome:r[1], cognome:r[2], telefono:r[3], email:r[4],
      pacchetto:r[5], lezioniTot:parseInt(r[6])||0, lezioniRim:parseInt(r[7])||0,
      dataInizio:r[8]?Utilities.formatDate(new Date(r[8]),tz,"yyyy-MM-dd"):"",
      dataScad:r[9]?Utilities.formatDate(new Date(r[9]),tz,"yyyy-MM-dd"):"",
      stato:r[10], token:r[11],
      dataNascita:r[12]||"", sesso:r[13]||"", indirizzo:r[14]||"",
      instagram:r[15]||"", facebook:r[16]||"", noteAnamnesi:r[17]||"",
      _riga:i+1,
    });
  }
  return out;
}

function getPacchetti() {
  const rows=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pacchetti").getDataRange().getValues(), out=[];
  for (let i=1; i<rows.length; i++) { if(!rows[i][0])continue; out.push({id:rows[i][0],nome:rows[i][1],lezioni:rows[i][2],durata:rows[i][3],prezzo:rows[i][4]}); }
  return out;
}

function aggiungiCliente(data) {
  const ss=SpreadsheetApp.getActiveSpreadsheet(), shC=ss.getSheetByName("Clienti"), shP=ss.getSheetByName("Pacchetti"), tz=Session.getScriptTimeZone();
  const pkg=shP.getDataRange().getValues().find(r=>r[0]===data.idPacchetto);
  if (!pkg) return { error:"Pacchetto non trovato: "+data.idPacchetto };
  const id="CLI"+Date.now();
  const token=Utilities.base64Encode(id+Math.random()).replace(/[^a-zA-Z0-9]/g,"").substring(0,20);
  const oggi=new Date(), scad=new Date(); scad.setDate(scad.getDate()+parseInt(pkg[3]));
  shC.appendRow([
    id, data.nome, data.cognome, data.telefono||"", data.email||"",
    pkg[1], pkg[2], pkg[2],
    Utilities.formatDate(oggi,tz,"yyyy-MM-dd"),
    Utilities.formatDate(scad,tz,"yyyy-MM-dd"),
    "Attivo", token,
    data.dataNascita||"", data.sesso||"", data.indirizzo||"",
    data.instagram||"", data.facebook||"", data.noteAnamnesi||""
  ]);
  const link="https://OmicronPT.github.io/Omicron-Studio/cliente.html?t="+token;
  Logger.log("Cliente: "+data.nome+" "+data.cognome+" | Link: "+link);
  return { ok:true, id, token, link };
}

// ──────────────────────────────────────────────────────────
//  TRIGGER
// ──────────────────────────────────────────────────────────
function inviaReminder() {
  const tz     = Session.getScriptTimeZone();
  const domani = new Date(); domani.setDate(domani.getDate()+1);
  const domaniStr = _dateToStr(domani);
  const clienti   = _tuttiClienti();
  _tuttePrenotazioni()
    .filter(p => p.data===domaniStr && p.stato==="Confermata")
    .forEach(p => {
      const c = clienti.find(x=>x.id===p.idCliente);
      if (c) _wa(c.telefono, `🔔 *Reminder ${CONFIG.STUDIO_NAME}*\nCiao ${c.nome}! Ti aspettiamo domani alle ${p.oraInizio} 💪`);
    });
}

function controllaScadenze() {
  const oggi=new Date(), tra7=new Date(); tra7.setDate(oggi.getDate()+7);
  _tuttiClienti().forEach(c => {
    if(c.stato!=="Attivo")return;
    const scad=new Date(c.dataScad);
    if(scad>=oggi&&scad<=tra7)
      _wa(c.telefono,`⚠️ *${CONFIG.STUDIO_NAME}*\nCiao ${c.nome}! Pacchetto in scadenza il *${c.dataScad}* con ${c.lezioniRim} lezioni. Contattaci!`);
  });
}

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("inviaReminder").timeBased().everyDays(1).atHour(CONFIG.ORE_REMINDER).create();
  ScriptApp.newTrigger("controllaScadenze").timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
  SpreadsheetApp.getUi().alert("Trigger attivati!");
}

// ──────────────────────────────────────────────────────────
//  WHATSAPP
// ──────────────────────────────────────────────────────────
function _wa(numero, msg) {
  if(!CONFIG.CALLMEBOT_API_KEY)return;
  try { UrlFetchApp.fetch("https://api.callmebot.com/whatsapp.php?phone="+numero+"&text="+encodeURIComponent(msg)+"&apikey="+CONFIG.CALLMEBOT_API_KEY); }
  catch(e){ Logger.log("WA error: "+e.message); }
}
