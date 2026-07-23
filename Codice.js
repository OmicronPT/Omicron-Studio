// ============================================================
//  OMICRON STUDIO — Sistema Prenotazioni
//  Google Apps Script — Backend v5.0
//  Slot automatici 30min, 9:00-20:00, lun-ven
// ============================================================

const CONFIG = {
  STUDIO_NAME:          "Omicron Studio",
  ADMIN_WHATSAPP:       "393514880234",   // tuo numero WhatsApp
  WAAPI_INSTANCE_ID:    PropertiesService.getScriptProperties().getProperty("WAAPI_INSTANCE_ID") || "",  // letto dalle Proprieta dello script
  WAAPI_TOKEN:          PropertiesService.getScriptProperties().getProperty("WAAPI_TOKEN") || "",         // letto dalle Proprieta dello script
  WAAPI_URL:            "https://waapi.app/api/v1/instances/",    // base URL WAAPI (instanceId e azione vengono aggiunti in _waSend)
  ORE_REMINDER:         18,               // ora invio reminder giornaliero
  SOGLIA_AVVISO:        2,                // avvisa quando rimangono N lezioni
  MAX_CONTEMPORANEI:    3,                // max persone per slot
  DURATA_SLOT_MIN:      30,               // durata slot in minuti
  DURATA_SESSION_MIN:   60,               // durata sessione in minuti
  MAX_SETTIMANE_RICORRENTE: 4,
  ORE_CANCELLAZIONE:    24,               // soglia (ore) sotto la quale una cancellazione è "tardiva" — vedi _valutaCancellazione (policy preavviso+jolly, dal 23/7/2026 non è più un blocco rigido)
  ORE_PRENOTAZIONE_MIN: 4,                // ore minime per prenotare
  LISTA_ATTESA_MAX_NOTIFICATI: 3,         // quante persone avvisare insieme quando si libera un posto
  LISTA_ATTESA_TIMEOUT_MIN: 15,           // minuti di tempo per prenotare dopo l'avviso
  ORA_INIZIO:           8.5,              // 08:30
  ORA_FINE:             20,               // ultimo slot alle 19:00
  SETTIMANE_AVANTI:     2,                // settimane visibili
  ADMIN_PASSWORD:       PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD") || "",  // letta dalle Proprieta dello script, NON piu scritta qui
  CALENDAR_ID:          "debe52a9ce43a5851dd0b630435f83c5774c8062c4029121e6e69984ed798cf3@group.calendar.google.com",
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
    "Data Nascita","Sesso","Indirizzo","Instagram","Facebook","Note Anamnesi",
    "Scadenza Certificato"
  ]);
  _creaFoglio(ss, "Prenotazioni", [
    "ID","ID Cliente","Nome Cliente","Data","Ora Inizio","Ora Fine",
    "Data Prenotazione","Stato"
  ]);
  _creaFoglio(ss, "Blocchi", [
    "ID","Data","Ora Inizio","Ora Fine","Motivo","Creato Il","Gruppo ID"
  ]);
  _creaFoglio(ss, "Pacchetti", [
    "ID","Nome","Lezioni","Durata Giorni","Prezzo","Note"
  ]);
  _creaFoglio(ss, "ListaAttesa", [
    "ID","ID Cliente","Nome Cliente","Data","Ora Inizio","Data Iscrizione","Stato","Notificato Il"
  ]);
  _creaFoglio(ss, "Log", [
    "ID","ID Cliente","Nome Cliente","Tipo","Descrizione","Data Evento"
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
  const action    = e.parameter.action    || "";
  const token     = e.parameter.token     || "";
  const callback  = e.parameter.callback  || "";
  const pw        = e.parameter.pw        || "";
  const idCliente = e.parameter.idCliente || "";

  let result;
  try {
    switch (action) {
      case "getSlotDisponibili":   result = getSlotDisponibili(token);       break;
      case "getInitData":          result = getInitData(token);              break;
      case "getCliente":           result = getClienteByToken(token);        break;
      case "getPrenotazioni":      result = getPrenotazioniCliente(token);   break;
      case "getListaAttesa":       result = getListaAttesaCliente(token);   break;
      case "getSchedeCliente":     result = getSchedeCliente(token);        break;
      case "getLogCliente":        result = getLogCliente(token);           break;
      case "adminDashboard":       result = adminDashboard(pw);              break;
      case "adminClienti":         result = adminClienti(pw);                break;
      case "adminCalendario":      result = adminCalendario(pw);             break;
      case "adminBlocchi":         result = adminBlocchi(pw);                break;
      case "adminStoricoCanc":     result = adminStoricoCanc(pw, idCliente); break;
      case "adminGetLog":          result = adminGetLog(pw, idCliente);      break;
      case "adminSchede":          result = adminSchede(pw, idCliente);      break;
      case "adminGetLimiti":  result = handleGetLimiti();  break;
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
      case "prenotaRicorrente": result = prenotaRicorrente(data.token, data.data, data.oraInizio, data.settimane); break;
      case "cancella":                  result = cancellaPrenotazione(data.token, data.idPrenotazione); break;
      case "adminLogin":                result = adminLogin(pw);                                        break;
      case "adminAddBlocco":            result = adminAddBlocco(pw, data);                              break;
      case "adminDelBlocco":            result = adminDelBlocco(pw, data.id);                           break;
      case "adminAddCliente":           result = adminAddCliente(pw, data);                             break;
      case "adminEditCliente":          result = adminEditCliente(pw, data);                            break;
      case "adminDelCliente":           result = adminDelCliente(pw, data.id);                          break;
      case "adminRigeneraToken":        result = adminRigeneraToken(pw, data.id);                       break;
      case "adminCancellaPrenotazione": result = adminCancellaPrenotazione(pw, data.id);                break;
      case "adminSegnaAssenza":        result = adminSegnaAssenza(pw, data.id);                        break;
      case "adminRinnovaAbbonamento":    result = adminRinnovaAbbonamento(pw, data);                  break;
      case "adminAnnullaAbbonamento":    result = adminAnnullaAbbonamento(pw, data.id);               break;
      case "adminDelBloccoGruppo":       result = adminDelBloccoGruppo(pw, data.gruppoId);           break;
      case "richiestaRinnovo":           result = richiestaRinnovo(data.token, data.idPacchetto);    break;
      case "iscriviListaAttesa":         result = iscriviListaAttesa(data.token, data.data, data.oraInizio); break;
      case "cancellaListaAttesa":        result = cancellaListaAttesa(data.token, data.id);                  break;
      case "adminUploadScheda":          result = adminUploadScheda(pw, data);                               break;
      case "adminDelScheda":             result = adminDelScheda(pw, data.id);                               break;

      // ── LIMITI CAPIENZA ──────────────────────────────────────────
      case "adminGetLimiti":             result = handleGetLimiti();                                         break;
      case "adminAddLimite":             result = handleAggiungiLimite(data);                                break;
      case "adminToggleLimite":          result = handleToggleLimite(data);                                  break;
      case "adminDelLimite":             result = handleEliminaLimite(data);                                 break;

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
// Calcola la data di Pasqua per un dato anno (algoritmo di Butcher)
function _pasqua(anno) {
  const a = anno % 19;
  const b = Math.floor(anno / 100);
  const c = anno % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const mese = Math.floor((h + l - 7 * m + 114) / 31);
  const giorno = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(anno, mese - 1, giorno);
}

// Restituisce set di stringhe "MM-DD" per festività fisse
// e "YYYY-MM-DD" per festività mobili dell'anno dato
function _festivita(anno) {
  const fisse = new Set([
    "01-01", // Capodanno
    "01-06", // Epifania
    "04-25", // Festa della Liberazione
    "05-01", // Festa del Lavoro
    "06-02", // Festa della Repubblica
    "08-15", // Ferragosto
    "11-01", // Ognissanti
    "12-08", // Immacolata Concezione
    "12-25", // Natale
    "12-26", // Santo Stefano
    "09-19", // San Gennaro - Patrono di Napoli
  ]);

  // Festività mobili: Pasqua e Pasquetta
  const pasqua = _pasqua(anno);
  const pasquetta = new Date(pasqua);
  pasquetta.setDate(pasquetta.getDate() + 1);

  const mobili = new Set([
    pasqua.getFullYear()+"-"+String(pasqua.getMonth()+1).padStart(2,"0")+"-"+String(pasqua.getDate()).padStart(2,"0"),
    pasquetta.getFullYear()+"-"+String(pasquetta.getMonth()+1).padStart(2,"0")+"-"+String(pasquetta.getDate()).padStart(2,"0"),
  ]);

  return { fisse, mobili };
}

function _isFestivo(dateStr) {
  const d = new Date(dateStr + "T12:00:00");
  const anno = d.getFullYear();
  const { fisse, mobili } = _festivita(anno);
  const mmdd = dateStr.substring(5); // "MM-DD"
  return fisse.has(mmdd) || mobili.has(dateStr);
}

function _isFeriale(dateStr) {
  const d = new Date(dateStr + "T12:00:00");
  const dow = d.getDay(); // 0=dom, 6=sab
  if (dow < 1 || dow > 5) return false; // weekend
  if (_isFestivo(dateStr)) return false;  // festività
  return true;
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
function getInitData(token) {
  return {
    cliente:      getClienteByToken(token),
    slot:         getSlotDisponibili(token),
    prenotazioni: getPrenotazioniCliente(token),
    listaAttesa:  getListaAttesaCliente(token),
    schede:       getSchedeCliente(token),
    log:          getLogCliente(token),
    pacchetti:    getPacchetti()
  };
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
  const limiti = getLimitiAttivi();
  
  for (let d = new Date(oggi); d < fine; d.setDate(d.getDate()+1)) {
    const dateStr = d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");
    if (!_isFeriale(dateStr)) continue;

    const slotsGiornata = [];
    for (const ora of slotsGiorno) {
      if (_isBloccat(dateStr, ora, blocchi)) continue;
      const count = _conteggioSovrapposti(dateStr, ora, prenotazioni);

      // Capienza dinamica: tiene conto di eventuali limiti personalizzati attivi
      const capienzaSlot = getCapienzaSlot(d, ora, limiti);
      const pieno = count >= capienzaSlot;

      const prenotato = preClienteSet.has(dateStr + "_" + ora);

      // Controllo prenotazione minima (non mostrare slot troppo vicini)
      const oraSlotCheck = new Date(dateStr + "T" + ora + ":00");
      if ((oraSlotCheck - new Date()) < CONFIG.ORE_PRENOTAZIONE_MIN * 3600000) continue;

      // Includi slot pieni per mostrare il pulsante lista attesa
      // ma solo se il cliente non è già prenotato
      slotsGiornata.push({
        ora,
        oraFine: _minToOre(_oreToMin(ora) + CONFIG.DURATA_SESSION_MIN),
        postiLiberi: Math.max(0, capienzaSlot - count),
        prenotato: prenotato,
        pieno: pieno,
      });
    }

    if (slotsGiornata.length > 0) {
      risultato.push({
        data: dateStr,
        dataLabel: _formatDataIT(d, "EEEE d MMMM yyyy"),
        slots: slotsGiornata,
      });
    }
  }
  return risultato;
}

// ──────────────────────────────────────────────────────────
//  PRENOTAZIONI
// ──────────────────────────────────────────────────────────
function prenotaSlot(token, data, oraInizio, silent, ctx) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const cliente = ctx ? ctx.cliente : getClienteByToken(token);
  if (cliente.error)              return cliente;
  if (cliente.stato !== "Attivo") return { error: "Abbonamento non attivo." };
  if (cliente.lezioniRim <= 0)    return { error: "Nessuna lezione rimanente." };
  if (cliente.dataScad && new Date(cliente.dataScad) < new Date()) return { error: "Pacchetto scaduto." };

  const prenotazioni = ctx ? ctx.prenotazioni : _tuttePrenotazioni();
  const blocchi      = ctx ? ctx.blocchi : _tuttiBlocchi();

  if (_isBloccat(data, oraInizio, blocchi)) return { error: "Orario non disponibile." };

  // Controllo limite settimanale in base al pacchetto
  const pacchetti = ctx ? ctx.pacchetti : getPacchetti();
  const pkg = pacchetti.find(p => p.nome === cliente.pacchetto);
  const limiteSettimana = pkg ? (parseInt(pkg.lezioniSettimana) > 0 ? parseInt(pkg.lezioniSettimana) : 3) : 3;

  const dataSlot = new Date(data + "T12:00:00");
  const dow = (dataSlot.getDay() + 6) % 7; // lun=0, dom=6
  const lunedi = new Date(dataSlot); lunedi.setDate(dataSlot.getDate() - dow); lunedi.setHours(0,0,0,0);
  const domenica = new Date(lunedi); domenica.setDate(lunedi.getDate() + 6); domenica.setHours(23,59,59,999);
  const preSettimana = prenotazioni.filter(p => {
    if (p.idCliente !== cliente.id || p.stato === "Cancellata") return false;
    const ds = new Date(p.data + "T12:00:00");
    return ds >= lunedi && ds <= domenica;
  });
  if (preSettimana.length >= limiteSettimana) {
    return { error: `Hai raggiunto il limite di ${limiteSettimana} prenotazioni settimanali per il pacchetto ${cliente.pacchetto}.` };
  }

  const count = _conteggioSovrapposti(data, oraInizio, prenotazioni);
  const dataObjSlot = new Date(data + "T12:00:00");
  const limitiSlot = ctx ? ctx.limiti : getLimitiAttivi();
  const capienzaSlot = getCapienzaSlot(dataObjSlot, oraInizio, limitiSlot);
  if (count >= capienzaSlot) return { error: "Slot al completo." };

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

  // La creazione dell'evento Calendar non è più sincrona: si scrive subito la riga
  // (colonna 10 = eventId, vuota per ora) e si accoda la creazione su CodaCal.
  // drainCodaCal() (trigger a tempo) creerà l'evento e scriverà l'eventId in J entro ~1 minuto.
  ss.getSheetByName("Prenotazioni").appendRow([
    id, cliente.id, cliente.nome+" "+cliente.cognome,
    data, oraInizio, oraFine, now, "Confermata", "", ""
  ]);

  _calEnqueue(id, cliente.nome+" "+cliente.cognome, data, oraInizio, oraFine);

  const tz2 = Session.getScriptTimeZone();
  _logEvento(cliente.id, cliente.nome+" "+cliente.cognome, "Prenotazione",
    "Prenotazione confermata per il " + _formatDataIT(new Date(data+"T12:00:00"), "EEEE d MMMM") + " ore " + oraInizio);

  // Aggiorna la lista d'attesa se questo slot era anche in attesa (vedi _gestisciPrenotazioneListaAttesa)
  const slotOraPieno = (count + 1) >= capienzaSlot;
  _gestisciPrenotazioneListaAttesa(cliente.id, data, oraInizio, slotOraPieno);

  // Scala lezione
  const nuovoRim = cliente.lezioniRim - 1;
  ss.getSheetByName("Clienti").getRange(cliente._riga, 8).setValue(nuovoRim);

  // WhatsApp (soppresso quando chiamato in modalità silent, es. ricorrente)
  if (!silent) {
    const dl = _formatDataIT(new Date(data+"T12:00:00"), "EEEE d MMMM");
    _wa(cliente.telefono, `✅ *${CONFIG.STUDIO_NAME}*\nCiao ${cliente.nome}! Prenotazione confermata.\n📅 ${dl} ore ${oraInizio}\nLezioni rimanenti: ${nuovoRim}`);
    if (nuovoRim <= CONFIG.SOGLIA_AVVISO && nuovoRim > 0)
      _wa(cliente.telefono, `⚠️ *${CONFIG.STUDIO_NAME}*\nCiao ${cliente.nome}, rimangono solo *${nuovoRim} lezioni*. Contattaci!`);
    _wa(CONFIG.ADMIN_WHATSAPP, `🔔 *Nuova prenotazione*\n👤 ${cliente.nome} ${cliente.cognome}\n📅 ${dl} ore ${oraInizio}\n💪 Lezioni rimanenti: ${nuovoRim}`);
  }

  // In modalità ctx (ricorrente): aggiorna lo stato condiviso in memoria
  // per rendere corrette le iterazioni successive
  if (ctx) {
    cliente.lezioniRim = nuovoRim;
    ctx.prenotazioni.push({
      id: id, idCliente: cliente.id, nomeCliente: cliente.nome + " " + cliente.cognome,
      data: data, oraInizio: oraInizio, oraFine: oraFine, stato: "Confermata", _riga: null
    });
  }

  return { ok: true, idPrenotazione: id, lezioniRimanenti: nuovoRim };
}
function prenotaRicorrente(token, data, oraInizio, settimane) {
  settimane = parseInt(settimane) || 1;
  if (settimane < 1) settimane = 1;
  if (settimane > CONFIG.MAX_SETTIMANE_RICORRENTE) settimane = CONFIG.MAX_SETTIMANE_RICORRENTE;

  const cliente = getClienteByToken(token);
  if (cliente.error) return cliente;


  // Carica UNA sola volta i dati condivisi e passali a prenotaSlot nel loop.
  const ctx = {
    cliente:      cliente,
    prenotazioni: _tuttePrenotazioni(),
    blocchi:      _tuttiBlocchi(),
    pacchetti:    getPacchetti(),
    limiti:       getLimitiAttivi()
  };
  const tz        = Session.getScriptTimeZone();
  const prenotate = [];   // occorrenze riuscite
  const saltate   = [];   // occorrenze non prenotabili + motivo

  for (let i = 0; i < settimane; i++) {
    const d = new Date(data + "T12:00:00");
    d.setDate(d.getDate() + i * 7);
    const dataStr   = d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0") + "-" + String(d.getDate()).padStart(2,"0");
    const dataLabel = _formatDataIT(d, "EEEE d MMMM");

    // Salta i giorni festivi (prenotaSlot non li controlla)
    if (!_isFeriale(dataStr)) {
      saltate.push({ data: dataStr, dataLabel: dataLabel, motivo: "Giorno festivo" });
      continue;
    }

    // Prenota in modalità silent: nessun WhatsApp per la singola
    const res = prenotaSlot(token, dataStr, oraInizio, true, ctx);
    if (res && res.ok) {
      prenotate.push({ data: dataStr, oraInizio: oraInizio, dataLabel: dataLabel });
    } else {
      saltate.push({ data: dataStr, dataLabel: dataLabel, motivo: (res && res.error) ? res.error : "Non prenotabile" });
    }
  }

  // Se non è stato prenotato nulla, restituisci errore
  if (prenotate.length === 0) {
    return { error: (saltate[0] && saltate[0].motivo) || "Nessuno slot prenotabile.", prenotate: [], saltate: saltate };
  }

  // Rileggo il cliente per le lezioni rimanenti aggiornate
  const clienteAgg = getClienteByToken(token);
  const lezioniRim = clienteAgg.error ? null : clienteAgg.lezioniRim;

  // ── WhatsApp: un solo riepilogo ──
  const elenco = prenotate.map(p => `• ${p.dataLabel} ore ${p.oraInizio}`).join("\n");
  let msgCli = `✅ *${CONFIG.STUDIO_NAME}*\nCiao ${cliente.nome}! Prenotate ${prenotate.length} sessioni:\n${elenco}`;
  if (lezioniRim !== null) msgCli += `\nLezioni rimanenti: ${lezioniRim}`;
  if (saltate.length) msgCli += `\n\n⚠️ Non prenotate (${saltate.length}): ` + saltate.map(s => s.dataLabel).join(", ");
  _wa(cliente.telefono, msgCli);

  if (lezioniRim !== null && lezioniRim <= CONFIG.SOGLIA_AVVISO && lezioniRim > 0)
    _wa(cliente.telefono, `⚠️ *${CONFIG.STUDIO_NAME}*\nCiao ${cliente.nome}, rimangono solo *${lezioniRim} lezioni*. Contattaci!`);

  _wa(CONFIG.ADMIN_WHATSAPP, `🔔 *Nuova prenotazione ricorrente*\n👤 ${cliente.nome} ${cliente.cognome}\n${elenco}\n💪 Lezioni rimanenti: ${lezioniRim}`);

  return { ok: true, prenotate: prenotate, saltate: saltate, lezioniRimanenti: lezioniRim };
}

// Policy di cancellazione "preavviso + jolly" (decisa con Fabrizio il 23/7/2026,
// sostituisce il vecchio limite fisso di 3 cancellazioni gratuite al mese e il
// vecchio blocco rigido sotto le 24h):
// - Cancellazione con >= CONFIG.ORE_CANCELLAZIONE ore di anticipo sulla lezione:
//   SEMPRE gratuita, lezione riaccreditata, nessun limite di quante volte.
// - Cancellazione con meno anticipo ("tardiva"): la lezione viene riaccreditata
//   SOLO se il cliente ha ancora disponibile il "jolly" del ciclo corrente (1 ogni
//   30 giorni, calcolati da dataInizio del cliente — la data di inizio/rinnovo del
//   pacchetto, stessa colonna sia per abbonamenti mensili che per pacchetti a
//   ingressi). Se il jolly di questo ciclo e' gia' stato usato, la lezione va persa.
// Il jolly usato si registra sulla riga di Prenotazioni, colonna K (indice 10),
// cosi' da poter verificare in futuro se e' gia' stato consumato nel ciclo corrente.
// rows: dump completo (getDataRange().getValues()) del foglio Prenotazioni, usato
// per cercare eventuali jolly gia' usati dal cliente nel ciclo attuale.
function _valutaCancellazione(idCliente, dataInizioCliente, dataPre, oraPre, rows) {
  const oraCancellazione = new Date();
  const dataLezione = new Date(dataPre + "T" + oraPre);
  const tardiva = (dataLezione - oraCancellazione) < CONFIG.ORE_CANCELLAZIONE * 3600000;

  if (!tardiva) {
    return { lezioneRecuperata: true, jollyUsato: false, tardiva: false };
  }

  // Calcola l'inizio del ciclo di 30 giorni corrente, a partire da dataInizio del cliente.
  // Se per qualche motivo dataInizio manca (non dovrebbe mai succedere), si ricade su
  // "nessun jolly disponibile" per sicurezza (meglio negare un riaccredito raro che
  // introdurre un jolly infinito per errore).
  if (!dataInizioCliente) {
    return { lezioneRecuperata: false, jollyUsato: false, tardiva: true };
  }
  const inizio = new Date(dataInizioCliente + "T00:00:00");
  const giorniPassati = Math.floor((oraCancellazione - inizio) / 86400000);
  const cicliPassati = Math.floor(Math.max(0, giorniPassati) / 30);
  const cicloInizio = new Date(inizio.getTime() + cicliPassati * 30 * 86400000);
  const cicloFine   = new Date(cicloInizio.getTime() + 30 * 86400000);

  // Cerca se il cliente ha già usato il jolly in questo ciclo (colonna K = indice 10)
  const jollyGiaUsato = rows.some(r => {
    if (r[1] !== idCliente || r[7] !== "Cancellata" || r[10] !== true) return false;
    const dataCanc = r[8] ? new Date(r[8]) : null;
    if (!dataCanc) return false;
    return dataCanc >= cicloInizio && dataCanc < cicloFine;
  });

  return jollyGiaUsato
    ? { lezioneRecuperata: false, jollyUsato: false, tardiva: true }
    : { lezioneRecuperata: true, jollyUsato: true, tardiva: true };
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

  // Normalizza data/ora: Google Sheets a volte auto-converte queste celle in
  // oggetti Data/Ora anche se scritte come testo, producendo date/orari sbagliati
  // nei messaggi WhatsApp (o addirittura errori nei calcoli) se usate cosi' come sono.
  const dataPre = _valStr(p[3], "yyyy-MM-dd");
  const oraPre  = _valStr(p[4], "HH:mm");

  // Guardia di sicurezza: non si può cancellare una lezione già passata (prima non
  // c'era nessun controllo su questo, dato che il vecchio blocco rigido delle 24h
  // copriva implicitamente anche questo caso; ora che quel blocco è stato rimosso
  // serve un controllo esplicito).
  if (new Date(dataPre+"T"+oraPre) < new Date())
    return { error: "Non puoi cancellare una lezione già passata." };

  // Policy "preavviso + jolly" (vedi _valutaCancellazione). Sostituisce sia il
  // vecchio blocco rigido sotto le 24h, sia il vecchio limite di 3 cancellazioni
  // gratuite al mese.
  const esito = _valutaCancellazione(cliente.id, cliente.dataInizio, dataPre, oraPre, rows);

  // Cancella prenotazione e registra la data/ora reale di cancellazione (colonna I)
  const oggi = new Date();
  const oraCancellazione = Utilities.formatDate(oggi, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  shPre.getRange(pRiga, 8).setValue("Cancellata");
  shPre.getRange(pRiga, 9).setValue(oraCancellazione);
  if (esito.jollyUsato) shPre.getRange(pRiga, 11).setValue(true); // colonna K = Jolly Usato

  // Rimuovi dal calendario Google
  _calDelPrenotazione(p[9]||null); // colonna 10 = eventId (se presente)

  // Log
  let descrLog = "Prenotazione del " + _formatDataIT(new Date(dataPre+"T12:00:00"), "EEEE d MMMM") + " ore " + oraPre;
  if (esito.jollyUsato)            descrLog += " — lezione riaccreditata (jolly cancellazione tardiva usato)";
  else if (esito.lezioneRecuperata) descrLog += " — lezione riaccreditata";
  else                              descrLog += " — lezione persa (cancellazione tardiva, jolly già usato questo ciclo)";
  _logEvento(cliente.id, cliente.nome+" "+cliente.cognome, "Cancellazione", descrLog);

  if (esito.lezioneRecuperata) {
    ss.getSheetByName("Clienti").getRange(cliente._riga, 8).setValue(cliente.lezioniRim + 1);
  }

  const dlCanc = _formatDataIT(new Date(dataPre+"T12:00:00"), "d MMMM");

  if (esito.jollyUsato) {
    _wa(cliente.telefono, `❌ *${CONFIG.STUDIO_NAME}*\nPrenotazione del ${dlCanc} alle ${oraPre} cancellata.\nLezione riaccreditata. ✅\n⚠️ Hai usato il tuo "jolly" per le cancellazioni sotto le ${CONFIG.ORE_CANCELLAZIONE}h: non sarà di nuovo disponibile per 30 giorni.`);
  } else if (esito.lezioneRecuperata) {
    _wa(cliente.telefono, `❌ *${CONFIG.STUDIO_NAME}*\nPrenotazione del ${dlCanc} alle ${oraPre} cancellata.\nLezione riaccreditata. ✅`);
  } else {
    _wa(cliente.telefono, `❌ *${CONFIG.STUDIO_NAME}*\nPrenotazione del ${dlCanc} alle ${oraPre} cancellata.\n⚠️ Cancellazione a meno di ${CONFIG.ORE_CANCELLAZIONE}h dalla lezione e hai già usato il "jolly" di questo mese — la lezione non viene riaccreditata.`);
  }

  // Notifica admin
  _wa(CONFIG.ADMIN_WHATSAPP, `⚠️ *Cancellazione*\n👤 ${cliente.nome} ${cliente.cognome}\n📅 ${dlCanc} ore ${oraPre}\n${esito.lezioneRecuperata ? 'Lezione riaccreditata ✅'+(esito.jollyUsato?' (jolly usato)':'') : 'Lezione persa (cancellazione tardiva, jolly già usato) ❌'}`);

  // Notifica lista d'attesa se c'è qualcuno in coda
  _notificaListaAttesa(dataPre, oraPre);

  return { ok: true, recreditata: esito.lezioneRecuperata, lezioniRimanenti: esito.lezioneRecuperata ? cliente.lezioniRim+1 : cliente.lezioniRim };
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
      dataLabel: _formatDataIT(new Date(p.data+"T12:00:00"), "EEEE d MMMM yyyy"),
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
      oraInizio: r[4] ? _valStr(r[4], "HH:mm") : "",
      oraFine:   r[5] ? _valStr(r[5], "HH:mm") : "",
      stato: r[7], _riga: i+1,
    });
  }
  return out;
}

function _dateToStr(d) {
  return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");
}

// Converte in stringa un valore letto da un foglio (data o ora) che Google Sheets
// potrebbe aver auto-convertito in un oggetto Date, anche se scritto come testo
// (stesso problema gia' documentato per CodaCal, qui applicato in modo generico
// a qualsiasi foglio/colonna). Se v e' gia' una stringa, la restituisce cosi' com'e'.
function _valStr(v, formato) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), formato);
  return String(v);
}

// Formatta una data con giorno della settimana e/o mese SEMPRE in italiano,
// indipendentemente dalla lingua impostata sul progetto Apps Script (che di
// default e' inglese e faceva uscire "Wednesday 15 July" invece di "mercoledì 15 luglio"
// nei messaggi WhatsApp). Usa un elenco di nomi propri invece di affidarsi alla
// localizzazione automatica di Utilities.formatDate.
function _formatDataIT(d, formato) {
  const GIORNI = ["domenica","lunedì","martedì","mercoledì","giovedì","venerdì","sabato"];
  const MESI   = ["gennaio","febbraio","marzo","aprile","maggio","giugno","luglio","agosto","settembre","ottobre","novembre","dicembre"];
  const giorno = GIORNI[d.getDay()];
  const mese   = MESI[d.getMonth()];
  const gg     = d.getDate();
  const aaaa   = d.getFullYear();
  switch (formato) {
    case "EEEE d MMMM yyyy": return `${giorno} ${gg} ${mese} ${aaaa}`;
    case "EEEE d MMMM":      return `${giorno} ${gg} ${mese}`;
    case "d MMMM yyyy":      return `${gg} ${mese} ${aaaa}`;
    case "d MMMM":           return `${gg} ${mese}`;
    default:                 return `${gg} ${mese} ${aaaa}`;
  }
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
      oraInizio: r[2] ? _valStr(r[2], "HH:mm") : "", oraFine: r[3] ? _valStr(r[3], "HH:mm") : "",
      motivo: r[4]||"", gruppoId: r[6]||"", _riga: i+1,
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
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName("Blocchi");
  const tz  = Session.getScriptTimeZone();
  const now = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm");

  // Ripetizione: none, weeks, forever
  const ripetizione = data.ripetizione || "none";
  const nSettimane  = parseInt(data.nSettimane) || 1;
  const giorni      = data.giorni || "single"; // single, weekday, allFeriali

  // Calcola tutte le date da bloccare
  const dateDaBloccare = _calcolaDateBlocco(data.data, ripetizione, nSettimane, giorni);

  const ids = [];
  // Usa un gruppo ID per collegare i blocchi ricorrenti
  const gruppoId = "GRP" + Date.now();
  dateDaBloccare.forEach((dateStr, idx) => {
    const id = "BLK" + Date.now() + "_" + idx;
    sh.appendRow([id, dateStr, data.oraInizio||"", data.oraFine||"", data.motivo||"", now, gruppoId]);
    ids.push(id);
    // Aggiungi blocco al calendario
    _calAddBlocco(dateStr, data.oraInizio||"", data.oraFine||"");
    Utilities.sleep(10); // evita ID duplicati
  });

  return { ok: true, ids, totale: ids.length, gruppoId };
}

// Calcola le date in cui creare i blocchi
function _calcolaDateBlocco(dataBase, ripetizione, nSettimane, giorni) {
  const date = [];
  const base = new Date(dataBase + "T12:00:00");
  const dowBase = base.getDay(); // 0=dom, 6=sab

  // Numero di settimane da coprire
  // "forever" = 52 settimane (1 anno)
  const settimane = ripetizione === "none" ? 1 : ripetizione === "forever" ? 52 : nSettimane;

  for (let w = 0; w < settimane; w++) {
    if (giorni === "single") {
      // Solo il giorno della settimana uguale alla data base
      const d = new Date(base);
      d.setDate(d.getDate() + w * 7);
      date.push(_dateToStr(d));
    } else if (giorni === "weekday") {
      // Stesso giorno della settimana ogni settimana
      const d = new Date(base);
      d.setDate(d.getDate() + w * 7);
      date.push(_dateToStr(d));
    } else if (giorni === "allFeriali") {
      // Tutti i giorni feriali della settimana
      const lunedi = new Date(base);
      // Vai al lunedì della settimana w
      const diffToMon = (dowBase === 0 ? -6 : 1 - dowBase);
      lunedi.setDate(base.getDate() + diffToMon + w * 7);
      for (let g = 0; g < 5; g++) { // lun-ven
        const day = new Date(lunedi);
        day.setDate(lunedi.getDate() + g);
        date.push(_dateToStr(day));
      }
    }
  }
  return [...new Set(date)].sort(); // rimuovi duplicati e ordina
}

// Rimuove tutti i blocchi di un gruppo ricorrente
function adminDelBloccoGruppo(pw, gruppoId) {
  _checkAdmin(pw);
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Blocchi");
  const rows=sh.getDataRange().getValues();
  // Elimina dall'ultima riga verso l'alto per non spostare gli indici
  for (let i=rows.length-1; i>=1; i--) {
    if (rows[i][6]===gruppoId) sh.deleteRow(i+1);
  }
  return { ok: true };
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
      return { ora, count, bloccato, clienti: preSlot.map(p=>p.nomeCliente),
        prenotazioni: preSlot.map(p=>({id:p.id, idCliente:p.idCliente, nome:p.nomeCliente})) };
    }).filter(s => s.count > 0 || s.bloccato);

    risultato.push({
      data: dateStr,
      dataLabel: _formatDataIT(new Date(dateStr+"T12:00:00"), "EEEE d MMMM yyyy"),
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
// Durata massima di una sessione admin, in secondi (6 ore = limite massimo consentito da CacheService)
const ADMIN_SESSION_DURATA_SEC = 21600;

function _checkAdmin(pw) {
  // Da qui in poi "pw" e' in realta' il gettone di sessione ricevuto al login, non la password vera.
  const sessioneValida = CacheService.getScriptCache().get("admin_sess_" + pw);
  if (!sessioneValida) throw new Error("Sessione scaduta, effettua di nuovo il login");
}
function adminLogin(pw) {
  if (pw !== CONFIG.ADMIN_PASSWORD) return { error: "Password non valida" };
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put("admin_sess_" + token, "1", ADMIN_SESSION_DURATA_SEC);
  return { ok: true, token: token };
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
      if (data.scadCertificato !== undefined) sh.getRange(i+1,19).setValue(data.scadCertificato);
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

function adminRigeneraToken(pw, id) {
  _checkAdmin(pw);
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName("Clienti"), rows=sh.getDataRange().getValues();
  for (let i=1; i<rows.length; i++) {
    if (rows[i][0]===id) {
      const nuovoToken=Utilities.getUuid().replace(/-/g,"");
      sh.getRange(i+1,12).setValue(nuovoToken); // colonna 12 = token
      const link="https://OmicronPT.github.io/Omicron-Studio/cliente.html?t="+nuovoToken;
      return { ok:true, token:nuovoToken, link:link };
    }
  }
  return { error: "Cliente non trovato" };
}

function adminCancellaPrenotazione(pw, id) {
  _checkAdmin(pw);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Prenotazioni");
  const rows = sh.getDataRange().getValues();

  let pRiga = -1, p = null;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) { pRiga = i + 1; p = rows[i]; break; }
  }
  if (!p) return { error: "Prenotazione non trovata" };
  if (p[7] === "Cancellata") return { error: "Già cancellata." };

  const c = _tuttiClienti().find(x => x.id === p[1]);

  // Stessa policy "preavviso + jolly" di cancellaPrenotazione (vedi _valutaCancellazione),
  // per restare allineati e non ripetere il disallineamento già capitato in passato tra
  // le due funzioni. NB: qui NON c'è nessuna guardia "lezione già passata" — a differenza
  // della versione cliente, l'admin deve poter sempre cancellare/correggere una
  // prenotazione, anche a lezione già iniziata o conclusa (decisione presa con
  // Fabrizio il 23/7/2026). Una cancellazione ammin di una lezione già passata rientra
  // comunque nel ramo "tardiva" della policy, quindi segue la stessa logica del jolly.
  const dataPreAdmin = _valStr(p[3], "yyyy-MM-dd");
  const oraPreAdmin  = _valStr(p[4], "HH:mm");
  const esito = _valutaCancellazione(p[1], c ? c.dataInizio : "", dataPreAdmin, oraPreAdmin, rows);

  const oggi = new Date();
  const oraCancellazione = Utilities.formatDate(oggi, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  sh.getRange(pRiga, 8).setValue("Cancellata");
  sh.getRange(pRiga, 9).setValue(oraCancellazione);
  if (esito.jollyUsato) sh.getRange(pRiga, 11).setValue(true); // colonna K = Jolly Usato

  // Rimuovi dal calendario Google (mancava del tutto in precedenza: lasciava eventi orfani)
  _calDelPrenotazione(p[9] || null);

  if (c && esito.lezioneRecuperata) {
    ss.getSheetByName("Clienti").getRange(c._riga, 8).setValue(c.lezioniRim + 1);
  }

  return { ok: true, lezioneRecuperata: esito.lezioneRecuperata };
}

// Segna una prenotazione passata come "Assente" (cliente non presentato).
// Decisioni prese con Fabrizio: la lezione NON viene riaccreditata (resta consumata,
// il posto era comunque riservato), e NON viene inviato alcun messaggio WhatsApp al
// cliente — resta visibile solo internamente nel pannello admin e nello storico Log.
function adminSegnaAssenza(pw, id) {
  _checkAdmin(pw);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Prenotazioni");
  const rows = sh.getDataRange().getValues();

  let pRiga = -1, p = null;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) { pRiga = i + 1; p = rows[i]; break; }
  }
  if (!p) return { error: "Prenotazione non trovata" };
  if (p[7] === "Cancellata") return { error: "Prenotazione già cancellata, non può essere segnata come assente." };
  if (p[7] === "Assente")    return { error: "Già segnata come assente." };

  sh.getRange(pRiga, 8).setValue("Assente");

  const c = _tuttiClienti().find(x => x.id === p[1]);
  _logEvento(p[1], c ? (c.nome + " " + c.cognome) : p[2], "Assenza",
    "Non presentato alla lezione del " + p[3] + " ore " + p[4]);

  return { ok: true };
}

// Annulla l'abbonamento di un cliente: il cliente NON viene eliminato (resta nella
// lista clienti), ma il suo abbonamento viene disattivato (stato "Annullato",
// lezioni rimanenti azzerate) e tutte le sue prenotazioni future ancora attive
// vengono cancellate automaticamente (stesso percorso di adminCancellaPrenotazione,
// quindi con rimozione dell'evento Calendar per ciascuna).
function adminAnnullaAbbonamento(pw, id) {
  _checkAdmin(pw);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shC = ss.getSheetByName("Clienti");
  const c = _tuttiClienti().find(x => x.id === id);
  if (!c) return { error: "Cliente non trovato" };
  if (c.stato === "Annullato") return { error: "Abbonamento già annullato." };

  // Cancella tutte le prenotazioni future non ancora cancellate di questo cliente
  const oggiStr = _dateToStr(new Date());
  const future = _tuttePrenotazioni().filter(p =>
    p.idCliente === id && p.stato !== "Cancellata" && p.data >= oggiStr
  );
  future.forEach(p => adminCancellaPrenotazione(pw, p.id));

  // Disattiva l'abbonamento: stato "Annullato" e lezioni rimanenti a zero.
  // (Il riaccredito eventuale fatto sopra da adminCancellaPrenotazione viene
  // sovrascritto qui, quindi non ha effetto sul risultato finale.)
  shC.getRange(c._riga, 11).setValue("Annullato"); // colonna 11 = Stato
  shC.getRange(c._riga, 8).setValue(0);            // colonna 8 = Lezioni Rimanenti

  _logEvento(c.id, c.nome + " " + c.cognome, "Annullamento",
    "Abbonamento " + c.pacchetto + " annullato" +
    (future.length > 0 ? " — " + future.length + " prenotazione/i futura/e cancellata/e automaticamente" : ""));

  return { ok: true, prenotazioniCancellate: future.length };
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
      dataNascita:r[12]?_valStr(r[12],"yyyy-MM-dd"):"", sesso:r[13]||"", indirizzo:r[14]||"",
      instagram:r[15]||"", facebook:r[16]||"", noteAnamnesi:r[17]||"",
      scadCertificato:r[18]?_valStr(r[18],"yyyy-MM-dd"):"",
      _riga:i+1,
    });
  }
  return out;
}

function getPacchetti() {
  const rows=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pacchetti").getDataRange().getValues(), out=[];
  for (let i=1; i<rows.length; i++) {
    if(!rows[i][0])continue;
    out.push({
      id:rows[i][0], nome:rows[i][1], lezioni:rows[i][2],
      durata:rows[i][3], prezzo:rows[i][4],
      tipo:rows[i][5]||"ingresso", lezioniSettimana:parseInt(rows[i][6])||0,
      note:rows[i][7]||"", descrizione:rows[i][8]||""
    });
  }
  return out;
}

// Crea il foglio Log se non esiste
function aggiornaLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _creaFoglio(ss, "Log", [
    "ID","ID Cliente","Nome Cliente","Tipo","Descrizione","Data Evento"
  ]);
  Logger.log("Foglio Log pronto!");
}

// Crea il foglio ListaAttesa se non esiste
function aggiornaListaAttesa() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _creaFoglio(ss, "ListaAttesa", [
    "ID","ID Cliente","Nome Cliente","Data","Ora Inizio","Data Iscrizione","Stato","Notificato Il"
  ]);
  _creaFoglio(ss, "Log", [
    "ID","ID Cliente","Nome Cliente","Tipo","Descrizione","Data Evento"
  ]);
  Logger.log("Foglio ListaAttesa pronto!");
}

// Aggiorna il foglio Blocchi aggiungendo la colonna Gruppo ID se mancante
function aggiornaBlocchi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Blocchi");
  if (!sh) { Logger.log("Foglio Blocchi non trovato"); return; }
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  if (!headers.includes("Gruppo ID")) {
    const col = sh.getLastColumn() + 1;
    sh.getRange(1, col).setValue("Gruppo ID")
      .setBackground("#1a1a2e").setFontColor("#ffffff").setFontWeight("bold");
    Logger.log("Colonna Gruppo ID aggiunta al foglio Blocchi");
  } else {
    Logger.log("Colonna Gruppo ID già presente");
  }
}

// Aggiorna il foglio Pacchetti con gli abbonamenti mensili Omicron Studio
function setupPacchetti() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Pacchetti");
  if (!sh) sh = ss.insertSheet("Pacchetti");
  sh.clearContents();
  const headers = ["ID","Nome","Lezioni","Durata Giorni","Prezzo","Tipo","Lezioni/Settimana","Note","Descrizione"];
  const r = sh.getRange(1,1,1,headers.length);
  r.setValues([headers]);
  r.setBackground("#1a1a2e").setFontColor("#ffffff").setFontWeight("bold");
  sh.setFrozenRows(1);
  sh.getRange(2,1,7,9).setValues([
    ["ABB001","Pulse",  4, 30, 120, "mensile",  1, "1 lezione/settimana",  "Abbonamento mensile · 1 volta a settimana · 4 lezioni/mese"],
    ["ABB002","Beat",   8, 30, 215, "mensile",  2, "2 lezioni/settimana",  "Abbonamento mensile · 2 volte a settimana · 8 lezioni/mese"],
    ["ABB003","Vibe",  12, 30, 290, "mensile",  3, "3 lezioni/settimana",  "Abbonamento mensile · 3 volte a settimana · 12 lezioni/mese"],
    ["PKG001","Rise",  10, 60, 300, "ingressi", 0, "10 ingressi, validità 2 mesi", "10 ingressi · validità 2 mesi"],
    ["PKG002","Grow",  15, 90, 420, "ingressi", 0, "15 ingressi, validità 3 mesi", "15 ingressi · validità 3 mesi"],
    ["PKG003","Flow",  20,120, 520, "ingressi", 0, "20 ingressi, validità 4 mesi", "20 ingressi · validità 4 mesi"],
    ["PKG004","Peak",  30,180, 720, "ingressi", 0, "30 ingressi, validità 6 mesi", "30 ingressi · validità 6 mesi"],
  ]);
  Logger.log("Pacchetti aggiornati: Pulse, Beat, Vibe, Rise, Grow, Flow, Peak!");
  return { ok: true };
}

// Rinnova abbonamento mensile con eventuale aggiunta lezioni recuperabili
function adminRinnovaAbbonamento(pw, data) {
  _checkAdmin(pw);
  const ss=SpreadsheetApp.getActiveSpreadsheet(), shC=ss.getSheetByName("Clienti"), tz=Session.getScriptTimeZone();
  const clienti=_tuttiClienti(), c=clienti.find(x=>x.id===data.id);
  if (!c) return { error:"Cliente non trovato" };
  const pkg=getPacchetti().find(p=>p.id===data.idPacchetto);
  if (!pkg) return { error:"Pacchetto non trovato" };
  const recuperabili=parseInt(data.lezioniRecuperabili)||0;
  const nuoveLezioni=parseInt(pkg.lezioni)+recuperabili;
  const oggi=new Date(), nuovaScad=new Date(); nuovaScad.setDate(oggi.getDate()+parseInt(pkg.durata));
  shC.getRange(c._riga,6).setValue(pkg.nome);
  shC.getRange(c._riga,7).setValue(nuoveLezioni);
  shC.getRange(c._riga,8).setValue(nuoveLezioni);
  shC.getRange(c._riga,9).setValue(Utilities.formatDate(oggi,tz,"yyyy-MM-dd"));
  shC.getRange(c._riga,10).setValue(Utilities.formatDate(nuovaScad,tz,"yyyy-MM-dd"));
  shC.getRange(c._riga,11).setValue("Attivo");
  _wa(c.telefono,`✅ *${CONFIG.STUDIO_NAME}*\nCiao ${c.nome}! Abbonamento *${pkg.nome}* rinnovato.\n📅 Valido fino al ${_formatDataIT(nuovaScad, "d MMMM yyyy")}\n💪 ${nuoveLezioni} lezioni${recuperabili>0?' (incluse '+recuperabili+' da recupero)':''}`);

  // Log rinnovo
  _logEvento(c.id, c.nome+" "+c.cognome, "Rinnovo",
    "Abbonamento " + pkg.nome + " rinnovato — " + nuoveLezioni + " lezioni" +
    (recuperabili > 0 ? " (incluse " + recuperabili + " da recupero)" : "") +
    " — scadenza " + _formatDataIT(nuovaScad, "d MMMM yyyy"));

  return { ok:true, nuoveLezioni, nuovaScad:Utilities.formatDate(nuovaScad,tz,"yyyy-MM-dd") };
}

// Storico cancellazioni cliente ultimo mese
function adminStoricoCanc(pw, idCliente) {
  _checkAdmin(pw);
  const tz=Session.getScriptTimeZone(), oggi=new Date(), unMeseFa=new Date();
  unMeseFa.setDate(oggi.getDate()-30);
  return _tuttePrenotazioni()
    .filter(p=>p.idCliente===idCliente&&p.stato==="Cancellata"&&new Date(p.data+"T12:00:00")>=unMeseFa)
    .map(p=>({id:p.id,data:p.data,dataLabel:_formatDataIT(new Date(p.data+"T12:00:00"), "EEEE d MMMM"),oraInizio:p.oraInizio}));
}

function aggiungiCliente(data) {
  const ss=SpreadsheetApp.getActiveSpreadsheet(), shC=ss.getSheetByName("Clienti"), tz=Session.getScriptTimeZone();

  let nomePkg, lezioni, durata;

  if (data.idPacchetto === "CUSTOM") {
    // Pacchetto personalizzato
    if (!data.pkgCustomNome || !data.pkgCustomLezioni || !data.pkgCustomDurata)
      return { error: "Dati pacchetto personalizzato mancanti" };
    nomePkg = data.pkgCustomNome;
    lezioni  = parseInt(data.pkgCustomLezioni);
    durata   = parseInt(data.pkgCustomDurata);
  } else {
    const shP = ss.getSheetByName("Pacchetti");
    const pkg = shP.getDataRange().getValues().find(r=>r[0]===data.idPacchetto);
    if (!pkg) return { error:"Pacchetto non trovato: "+data.idPacchetto };
    nomePkg = pkg[1];
    lezioni  = parseInt(pkg[2]);
    durata   = parseInt(pkg[3]);
  }

  const id="CLI"+Date.now();
  const token=Utilities.getUuid().replace(/-/g,""); // codice lungo e imprevedibile (32 caratteri), non più basato su orario+numero casuale
  const oggi=new Date(), scad=new Date(); scad.setDate(scad.getDate()+durata);
  shC.appendRow([
    id, data.nome, data.cognome, data.telefono||"", data.email||"",
    nomePkg, lezioni, lezioni,
    Utilities.formatDate(oggi,tz,"yyyy-MM-dd"),
    Utilities.formatDate(scad,tz,"yyyy-MM-dd"),
    "Attivo", token,
    data.dataNascita||"", data.sesso||"", data.indirizzo||"",
    data.instagram||"", data.facebook||"", data.noteAnamnesi||"",
    data.scadCertificato||""
  ]);
  const link="https://OmicronPT.github.io/Omicron-Studio/cliente.html?t="+token;
  Logger.log("Cliente: "+data.nome+" "+data.cognome+" | Link: "+link);

  // Messaggio di benvenuto automatico con link personale e istruzioni per salvare
  // la pagina sulla schermata Home del telefono (icona come un'app vera).
  if (data.telefono) {
    _wa(data.telefono, `🎉 *Benvenuto in ${CONFIG.STUDIO_NAME}, ${data.nome}!*\nEcco il tuo link personale per prenotare le lezioni:\n${link}\n\n📱 Salvalo sulla schermata Home del telefono, così lo trovi come un'app, senza cercarlo su WhatsApp ogni volta:\n\n*Se hai un iPhone:*\n1. Apri il link (si apre in Safari)\n2. Tocca l'icona di condivisione in basso (quadrato con freccia)\n3. Scorri e tocca "Aggiungi a Home"\n4. Tocca "Aggiungi" in alto a destra\n\n*Se hai un Android:*\n1. Apri il link (si apre in Chrome)\n2. Tocca i tre puntini in alto a destra\n3. Tocca "Aggiungi a schermata Home" (o "Installa app")\n4. Conferma\n\nA presto! 💪`);
  }

  return { ok:true, id, token, link };
}

// ──────────────────────────────────────────────────────────
//  RICHIESTA RINNOVO
// ──────────────────────────────────────────────────────────

function richiestaRinnovo(token, idPacchetto) {
  const cliente = getClienteByToken(token);
  if (cliente.error) return cliente;

  const tz  = Session.getScriptTimeZone();
  const now = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm");

  // Trova nome pacchetto richiesto
  let nomePacchetto = "stesso pacchetto";
  if (idPacchetto) {
    const pkg = getPacchetti().find(p => p.id === idPacchetto);
    if (pkg) nomePacchetto = pkg.nome + " (€" + pkg.prezzo + ")";
  }

  // Notifica admin WhatsApp
  _wa(CONFIG.ADMIN_WHATSAPP,
    `🔄 *Richiesta rinnovo*
👤 ${cliente.nome} ${cliente.cognome}
📦 Pacchetto richiesto: *${nomePacchetto}*
📅 Scadenza attuale: ${cliente.dataScad}
💪 Lezioni rimanenti: ${cliente.lezioniRim}
⏰ Richiesta alle ${now}`
  );

  // Conferma al cliente
  _wa(cliente.telefono,
    `✅ *${CONFIG.STUDIO_NAME}*
Ciao ${cliente.nome}! La tua richiesta di rinnovo per il pacchetto *${nomePacchetto}* è stata inviata.
Ti contatteremo presto per completare il rinnovo! 💪`
  );

  // Log
  _logEvento(cliente.id, cliente.nome+" "+cliente.cognome, "Rinnovo",
    "Richiesta rinnovo inviata — pacchetto: " + nomePacchetto);

  return { ok: true };
}

// ──────────────────────────────────────────────────────────
//  LOG ATTIVITÀ
// ──────────────────────────────────────────────────────────

function _logEvento(idCliente, nomeCliente, tipo, descrizione) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const sh  = ss.getSheetByName("Log");
    if (!sh) return;
    const id  = "LOG" + Date.now();
    const tz  = Session.getScriptTimeZone();
    const now = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm");
    sh.appendRow([id, idCliente, nomeCliente, tipo, descrizione, now]);
  } catch(e) {
    Logger.log("Log error: " + e.message);
  }
}

function getLogCliente(token) {
  const cliente = getClienteByToken(token);
  if (cliente.error) return cliente;
  return _getLogByCliente(cliente.id);
}

function adminGetLog(pw, idCliente) {
  _checkAdmin(pw);
  if (!idCliente) return [];
  return _getLogByCliente(idCliente);
}

function _getLogByCliente(idCliente) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sh   = ss.getSheetByName("Log");
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  const out  = [];
  for (let i=1; i<rows.length; i++) {
    const r = rows[i];
    if (!r[0] || r[1] !== idCliente) continue;
    out.push({
      id: r[0], tipo: r[3], descrizione: r[4], data: r[5]
    });
  }
  // Più recenti prima
  return out.reverse();
}

// ──────────────────────────────────────────────────────────
//  GOOGLE CALENDAR
// ──────────────────────────────────────────────────────────

function _getCalendar() {
  return CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
}

// Aggiunge evento prenotazione sul calendario
function _calAddPrenotazione(idPrenotazione, nomeCliente, data, oraInizio, oraFine) {
  try {
    const cal  = _getCalendar();
    const tz   = Session.getScriptTimeZone();
    const [hi, mi] = oraInizio.split(":").map(Number);
    const [hf, mf] = oraFine.split(":").map(Number);
    const start = new Date(data + "T" + oraInizio + ":00");
    const end   = new Date(data + "T" + oraFine   + ":00");

    const evento = cal.createEvent(
      "💪 " + nomeCliente,
      start, end,
      {
        description: "Prenotazione #" + idPrenotazione + "Cliente: " + nomeCliente,
        colorId: "9", // Blu
      }
    );
    // Salva eventId nel foglio Prenotazioni per poterlo aggiornare/cancellare
    return evento.getId();
  } catch(e) {
    Logger.log("Calendar add error: " + e.message);
    return null;
  }
}

// Cancella evento dal calendario
function _calDelPrenotazione(eventId) {
  try {
    if (!eventId) return;
    const cal   = _getCalendar();
    const event = cal.getEventById(eventId);
    if (event) event.deleteEvent();
  } catch(e) {
    Logger.log("Calendar del error: " + e.message);
  }
}

// Accoda la creazione dell'evento Calendar: l'evento vero e proprio viene creato
// in background da drainCodaCal (trigger a tempo), cosi' prenotaSlot non aspetta
// piu' la chiamata a CalendarApp.createEvent.
function _calEnqueue(idPrenotazione, nomeCliente, data, oraInizio, oraFine) {
  try {
    getFogliCodaCal().appendRow([idPrenotazione, nomeCliente, data, oraInizio, oraFine, "In coda", 0]);
  } catch(e) {
    Logger.log("Cal enqueue error: " + e.message);
  }
}

// Foglio coda Calendar (creato al volo se non esiste)
function getFogliCodaCal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("CodaCal");
  if (!sh) {
    sh = ss.insertSheet("CodaCal");
    sh.getRange(1, 1, 1, 7).setValues([["idPrenotazione", "nome", "data", "oraInizio", "oraFine", "Stato", "Tentativi"]]);
    // Forza colonne C,D,E (data/oraInizio/oraFine) come testo puro: altrimenti Sheets
    // le auto-converte in valori Data/Ora e _calAddPrenotazione riceve oggetti Date
    // invece di stringhe "yyyy-MM-dd" / "HH:mm", producendo orari sbagliati sul Calendar.
    sh.getRange(2, 3, sh.getMaxRows() - 1, 3).setNumberFormat("@");
  }
  return sh;
}

// Converte in stringa "yyyy-MM-dd"/"HH:mm" un valore letto da CodaCal, che Sheets
// potrebbe aver convertito in Date nonostante la formattazione testo (righe già
// esistenti prima del fix, o edge case di autoconversione).
function _codaCalStr(v, formato) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), formato);
  return String(v);
}

// Svuota la coda Calendar: crea gli eventi "In coda" e scrive l'eventId in colonna J
// della prenotazione corrispondente. Eseguito da un trigger a tempo (ogni minuto),
// stesso pattern di drainCodaWA -> fuori dal percorso critico.
function drainCodaCal() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return; // un altro drain e' gia' in corso
  try {
    const sh   = getFogliCodaCal();
    const rows = sh.getDataRange().getValues();
    const MAX_PER_RUN   = 50;
    const MAX_TENTATIVI = 3;
    let processate = 0;

    const shPre  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Prenotazioni");
    const rowsPre = shPre.getDataRange().getValues();

    for (let i = 1; i < rows.length && processate < MAX_PER_RUN; i++) {
      if (rows[i][5] !== "In coda") continue;
      processate++;

      const idPrenotazione = rows[i][0];

      // Trova la prenotazione corrispondente nel foglio Prenotazioni
      let pRiga = -1, statoAttuale = null;
      for (let j = 1; j < rowsPre.length; j++) {
        if (rowsPre[j][0] === idPrenotazione) { pRiga = j + 1; statoAttuale = rowsPre[j][7]; break; }
      }

      // Prenotazione non trovata o già cancellata (es. cancellata entro il minuto,
      // prima che l'evento venisse creato) -> salta, nessun evento orfano.
      if (pRiga === -1 || statoAttuale === "Cancellata") {
        sh.getRange(i + 1, 6).setValue("Saltata");
        continue;
      }

      const dataStr      = _codaCalStr(rows[i][2], "yyyy-MM-dd");
      const oraInizioStr = _codaCalStr(rows[i][3], "HH:mm");
      const oraFineStr   = _codaCalStr(rows[i][4], "HH:mm");
      const eventId = _calAddPrenotazione(idPrenotazione, rows[i][1], dataStr, oraInizioStr, oraFineStr);
      if (eventId) {
        shPre.getRange(pRiga, 10).setValue(eventId); // colonna J
        sh.getRange(i + 1, 6).setValue("Creato");
      } else {
        const tentativi = (Number(rows[i][6]) || 0) + 1;
        sh.getRange(i + 1, 7).setValue(tentativi);
        if (tentativi >= MAX_TENTATIVI) sh.getRange(i + 1, 6).setValue("Errore");
      }
    }
  } finally {
    lock.releaseLock();
  }
}

// Aggiunge blocco sul calendario
function _calAddBlocco(data, oraInizio, oraFine) {
  try {
    const cal = _getCalendar();
    let start, end;
    if (!oraInizio) {
      // Giorno intero
      start = new Date(data + "T00:00:00");
      end   = new Date(data + "T23:59:59");
    } else {
      start = new Date(data + "T" + oraInizio + ":00");
      end   = new Date(data + "T" + oraFine   + ":00");
    }
    const evento = cal.createEvent(
      "🔒 Blocco",
      start, end,
      { colorId: "11" } // Rosso
    );
    return evento.getId();
  } catch(e) {
    Logger.log("Calendar blocco error: " + e.message);
    return null;
  }
}

// Sincronizza tutte le prenotazioni future sul calendario (da eseguire una volta)
function sincronizzaCalendario() {
  const cal   = _getCalendar();
  const tz    = Session.getScriptTimeZone();
  const oggi  = new Date(); oggi.setHours(0,0,0,0);
  const pre   = _tuttePrenotazioni().filter(p => p.stato === "Confermata" && new Date(p.data+"T12:00:00") >= oggi);
  const blocchi = _tuttiBlocchi();

  let count = 0;
  pre.forEach(p => {
    const start = new Date(p.data + "T" + p.oraInizio + ":00");
    const end   = new Date(p.data + "T" + p.oraFine   + ":00");
    cal.createEvent("💪 " + p.nomeCliente, start, end, {
      description: "Prenotazione #" + p.id,
      colorId: "9"
    });
    count++;
  });

  blocchi.forEach(b => {
    if (!b.data) return;
    let start, end;
    if (!b.oraInizio) {
      start = new Date(b.data+"T00:00:00"); end = new Date(b.data+"T23:59:59");
    } else {
      start = new Date(b.data+"T"+b.oraInizio+":00"); end = new Date(b.data+"T"+b.oraFine+":00");
    }
    cal.createEvent("🔒 Blocco", start, end, { colorId: "11" });
    count++;
  });

  Logger.log("Sincronizzati " + count + " eventi sul calendario Omicron Studio!");
}

// ──────────────────────────────────────────────────────────
//  SCHEDE DI ALLENAMENTO
// ──────────────────────────────────────────────────────────

// Crea cartella cliente su Drive se non esiste
function _getCartellaCliente(idCliente, nomeCliente) {
  const rootName = "Omicron Studio - Schede";
  let rootFolder;

  // Cerca o crea cartella root
  const rootFolders = DriveApp.getFoldersByName(rootName);
  if (rootFolders.hasNext()) {
    rootFolder = rootFolders.next();
  } else {
    rootFolder = DriveApp.createFolder(rootName);
  }

  // Cerca o crea cartella cliente
  const clienteFolders = rootFolder.getFoldersByName(idCliente);
  if (clienteFolders.hasNext()) {
    return clienteFolders.next();
  } else {
    return rootFolder.createFolder(idCliente + " - " + nomeCliente);
  }
}

function adminSchede(pw, idCliente) {
  _checkAdmin(pw);
  if (!idCliente) return [];

  const clienti = _tuttiClienti();
  const c = clienti.find(x => x.id === idCliente);
  if (!c) return { error: "Cliente non trovato" };

  try {
    const folder = _getCartellaCliente(idCliente, c.nome+" "+c.cognome);
    const files = folder.getFiles();
    const out = [];
    while (files.hasNext()) {
      const f = files.next();
      out.push({
        id:       f.getId(),
        nome:     f.getName(),
        url:      f.getDownloadUrl(),
        viewUrl:  "https://drive.google.com/file/d/"+f.getId()+"/view",
        data:     Utilities.formatDate(f.getDateCreated(), Session.getScriptTimeZone(), "dd/MM/yyyy"),
      });
    }
    return out.sort((a,b) => b.data.localeCompare(a.data));
  } catch(e) {
    return { error: e.message };
  }
}

function adminUploadScheda(pw, data) {
  _checkAdmin(pw);
  // data: { idCliente, nome, base64, mimeType }
  if (!data.idCliente || !data.base64 || !data.nome) return { error: "Dati mancanti" };

  const clienti = _tuttiClienti();
  const c = clienti.find(x => x.id === data.idCliente);
  if (!c) return { error: "Cliente non trovato" };

  try {
    const folder   = _getCartellaCliente(data.idCliente, c.nome+" "+c.cognome);
    const decoded  = Utilities.base64Decode(data.base64);
    const blob     = Utilities.newBlob(decoded, data.mimeType || "application/pdf", data.nome);
    const file     = folder.createFile(blob);

    // Rendi il file accessibile con il link
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return {
      ok:      true,
      id:      file.getId(),
      nome:    file.getName(),
      viewUrl: "https://drive.google.com/file/d/"+file.getId()+"/view",
    };
  } catch(e) {
    return { error: e.message };
  }
}

function adminDelScheda(pw, id) {
  _checkAdmin(pw);
  try {
    DriveApp.getFileById(id).setTrashed(true);
    return { ok: true };
  } catch(e) {
    return { error: e.message };
  }
}

function getSchedeCliente(token) {
  const cliente = getClienteByToken(token);
  if (cliente.error) return cliente;

  try {
    const folder = _getCartellaCliente(cliente.id, cliente.nome+" "+cliente.cognome);
    const files  = folder.getFiles();
    const out    = [];
    while (files.hasNext()) {
      const f = files.next();
      out.push({
        id:      f.getId(),
        nome:    f.getName(),
        viewUrl: "https://drive.google.com/file/d/"+f.getId()+"/view",
        data:    Utilities.formatDate(f.getDateCreated(), Session.getScriptTimeZone(), "dd/MM/yyyy"),
      });
    }
    return out.sort((a,b) => b.data.localeCompare(a.data));
  } catch(e) {
    // Cartella non ancora creata = nessuna scheda
    return [];
  }
}

// ──────────────────────────────────────────────────────────
//  LISTA D'ATTESA
// ──────────────────────────────────────────────────────────

function _tuttiListaAttesa() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sh   = ss.getSheetByName("ListaAttesa");
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  const tz   = Session.getScriptTimeZone();
  const out  = [];
  for (let i=1; i<rows.length; i++) {
    const r = rows[i]; if (!r[0]) continue;
    out.push({
      id: r[0], idCliente: r[1], nomeCliente: r[2],
      data: r[3] ? _dateToStr(new Date(r[3])) : "",
      oraInizio: r[4] ? _valStr(r[4], "HH:mm") : "",
      dataIscrizione: r[5] ? Utilities.formatDate(new Date(r[5]), tz, "yyyy-MM-dd HH:mm") : "",
      stato: r[6] || "Attivo",
      notificatoIl: r[7] || "",
      _riga: i+1,
    });
  }
  return out;
}

function getListaAttesaCliente(token) {
  const cliente = getClienteByToken(token);
  if (cliente.error) return cliente;
  const lista = _tuttiListaAttesa();
  return lista.filter(l => l.idCliente === cliente.id && l.stato === "Attivo")
    .map(l => ({ id: l.id, data: l.data, oraInizio: l.oraInizio }));
}

function iscriviListaAttesa(token, data, oraInizio) {
  const cliente = getClienteByToken(token);
  if (cliente.error) return cliente;
  if (cliente.stato !== "Attivo") return { error: "Abbonamento non attivo." };
  if (cliente.lezioniRim <= 0)    return { error: "Nessuna lezione rimanente." };

  const lista = _tuttiListaAttesa();

  // Verifica che non sia già in lista per un altro slot
  const giàInLista = lista.find(l => l.idCliente === cliente.id && l.stato === "Attivo");
  if (giàInLista) return { error: "Sei già in lista d'attesa per un altro slot. Cancellati prima di iscriverti a un nuovo slot." };

  // Verifica che non sia già prenotato per questo slot
  const prenotazioni = _tuttePrenotazioni();
  const giàPrenotato = prenotazioni.find(p =>
    p.idCliente === cliente.id && p.data === data &&
    p.oraInizio === oraInizio && p.stato !== "Cancellata"
  );
  if (giàPrenotato) return { error: "Sei già prenotato per questo slot." };

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName("ListaAttesa");
  const id  = "ATT" + Date.now();
  const tz  = Session.getScriptTimeZone();
  const now = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm");

  sh.appendRow([id, cliente.id, cliente.nome+" "+cliente.cognome, data, oraInizio, now, "Attivo", ""]);

  // Conta posizione in lista
  const posizione = lista.filter(l => l.data === data && l.oraInizio === oraInizio && l.stato === "Attivo").length + 1;

  const tz2 = Session.getScriptTimeZone();
  const dl  = _formatDataIT(new Date(data+"T12:00:00"), "EEEE d MMMM");
  _wa(cliente.telefono, `⏳ *${CONFIG.STUDIO_NAME}*\nCiao ${cliente.nome}! Sei in lista d'attesa per:\n📅 ${dl} ore ${oraInizio}\nSei il numero ${posizione} in lista. Ti avviseremo subito se si libera un posto!`);

  return { ok: true, id, posizione };
}

function cancellaListaAttesa(token, id) {
  const cliente = getClienteByToken(token);
  if (cliente.error) return cliente;

  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sh   = ss.getSheetByName("ListaAttesa");
  const rows = sh.getDataRange().getValues();

  for (let i=1; i<rows.length; i++) {
    if (rows[i][0] === id && rows[i][1] === cliente.id) {
      sh.getRange(i+1, 7).setValue("Cancellato");
      return { ok: true };
    }
  }
  return { error: "Iscrizione non trovata." };
}

// Chiamata da prenotaSlot() dopo ogni prenotazione riuscita, per tenere pulita
// la lista d'attesa quando lo slot prenotato coincide con uno slot in attesa:
//  - se chi ha prenotato era lui stesso in lista d'attesa per questo slot (avvisato o no),
//    la sua voce viene marcata "Prenotato" (non deve più essere considerata).
//  - se lo slot ora è pieno, gli altri eventuali clienti avvisati per lo stesso slot
//    tornano "Attivo" (pronti per il prossimo posto libero) e ricevono un messaggio
//    che il posto è stato preso da un altro cliente — invece di restare bloccati per
//    sempre come "Notificato" o ricevere in seguito un avviso su un posto già occupato.
function _gestisciPrenotazioneListaAttesa(idCliente, data, oraInizio, slotOraPieno) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("ListaAttesa");
  if (!sh) return;

  const lista = _tuttiListaAttesa();
  const perSlot = lista.filter(l =>
    l.data === data && l.oraInizio === oraInizio && (l.stato === "Attivo" || l.stato === "Notificato")
  );
  if (!perSlot.length) return;

  const clienti = _tuttiClienti();
  const dl = _formatDataIT(new Date(data+"T12:00:00"), "EEEE d MMMM");

  perSlot.forEach(l => {
    if (l.idCliente === idCliente) {
      // Chi ha appena prenotato: la sua voce in lista d'attesa non serve più
      sh.getRange(l._riga, 7).setValue("Prenotato");
    } else if (slotOraPieno && l.stato === "Notificato") {
      // Altri clienti avvisati per lo stesso slot, ma il posto è stato preso: tornano in lista attiva
      sh.getRange(l._riga, 7).setValue("Attivo");
      sh.getRange(l._riga, 8).setValue("");
      const c = clienti.find(x => x.id === l.idCliente);
      if (c) _wa(c.telefono, `ℹ️ *${CONFIG.STUDIO_NAME}*\nCiao ${c.nome}, il posto per ${dl} ore ${oraInizio} è stato preso da un altro cliente.\nResti in lista d'attesa: ti avviseremo al prossimo posto libero.`);
    }
  });
}

// Chiamata quando uno slot si libera (dalla cancellazione prenotazione)
// Avvisa fino a CONFIG.LISTA_ATTESA_MAX_NOTIFICATI persone insieme (non solo la prima):
// vale il principio "primo che prenota, primo servito".
function _notificaListaAttesa(data, oraInizio) {
  const lista    = _tuttiListaAttesa();
  const clienti  = _tuttiClienti();
  const tz       = Session.getScriptTimeZone();
  const dl       = _formatDataIT(new Date(data+"T12:00:00"), "EEEE d MMMM");

  // Trova i primi N in lista attivi per questo slot (i più vecchi iscritti prima)
  const attesi = lista
    .filter(l => l.data === data && l.oraInizio === oraInizio && l.stato === "Attivo")
    .sort((a,b) => a.dataIscrizione.localeCompare(b.dataIscrizione))
    .slice(0, CONFIG.LISTA_ATTESA_MAX_NOTIFICATI);

  if (!attesi.length) return;

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName("ListaAttesa");
  const now = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm");
  const avvisoCondiviso = attesi.length > 1
    ? `\n⚡ Il posto è stato offerto anche ad altri clienti in lista: vale chi prenota per primo.`
    : "";

  attesi.forEach(l => {
    const cliente = clienti.find(c => c.id === l.idCliente);
    if (!cliente) return;

    // Marca come notificato
    sh.getRange(l._riga, 7).setValue("Notificato");
    sh.getRange(l._riga, 8).setValue(now);

    // Manda WhatsApp con link diretto
    const link = "https://OmicronPT.github.io/Omicron-Studio/cliente.html?t=" + cliente.token;
    _wa(cliente.telefono, `🎉 *${CONFIG.STUDIO_NAME}*\nCiao ${cliente.nome}! Si è liberato un posto per:\n📅 ${dl} ore ${oraInizio}\n\nHai *${CONFIG.LISTA_ATTESA_TIMEOUT_MIN} minuti* per prenotare:\n${link}${avvisoCondiviso}\n\nDopo ${CONFIG.LISTA_ATTESA_TIMEOUT_MIN} minuti, se nessuno ha prenotato, il posto passerà al prossimo in lista.`);
  });

  // Nota: la scadenza NON viene gestita qui con un trigger dedicato
  // (rimosso perché puntava a una funzione inesistente e creava trigger "fantasma"
  // mai ripuliti, con rischio di esaurire il limite di 20 trigger per progetto).
  // La scadenza è già gestita correttamente da controllaScadenzeListaAttesa(),
  // che gira ogni 10 minuti tramite trigger periodico (vedi setupTriggers()).
}

// Controlla scadenze lista attesa (trigger ogni 10 min, timeout CONFIG.LISTA_ATTESA_TIMEOUT_MIN)
function controllaScadenzeListaAttesa() {
  const lista = _tuttiListaAttesa();
  const ora   = new Date();

  const scaduti = lista.filter(l => {
    if (l.stato !== "Notificato" || !l.notificatoIl) return false;
    return (ora - new Date(l.notificatoIl)) >= CONFIG.LISTA_ATTESA_TIMEOUT_MIN * 60 * 1000;
  });
  if (!scaduti.length) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("ListaAttesa");

  // Marca tutti gli scaduti, raccogliendo gli slot distinti da ri-notificare
  // (con più persone avvisate insieme, più righe scadute possono riguardare
  // lo stesso slot: va ri-notificato una sola volta per slot, non una per riga)
  const slotDaRinotificare = new Set();
  scaduti.forEach(l => {
    sh.getRange(l._riga, 7).setValue("Scaduto");
    slotDaRinotificare.add(l.data + "|" + l.oraInizio);
  });

  slotDaRinotificare.forEach(chiave => {
    const [data, oraInizio] = chiave.split("|");
    _notificaListaAttesa(data, oraInizio);
  });
}

// ──────────────────────────────────────────────────────────
//  TRIGGER
// ──────────────────────────────────────────────────────────
function inviaReminder() {
  const tz        = Session.getScriptTimeZone();
  const domani    = new Date(); domani.setDate(domani.getDate()+1);
  const domaniStr = _dateToStr(domani);
  const domaniLabel = _formatDataIT(domani, "EEEE d MMMM");
  const clienti   = _tuttiClienti();
  const predomani = _tuttePrenotazioni().filter(p => p.data===domaniStr && p.stato==="Confermata");

  // Reminder ai clienti
  predomani.forEach(p => {
    const c = clienti.find(x=>x.id===p.idCliente);
    if (c) _wa(c.telefono, `🔔 *Reminder ${CONFIG.STUDIO_NAME}*\nCiao ${c.nome}! Ti aspettiamo domani alle ${p.oraInizio} 💪`);
  });

  // Riepilogo giornaliero all'admin
  if (predomani.length > 0) {
    const righe = predomani
      .sort((a,b) => a.oraInizio.localeCompare(b.oraInizio))
      .map(p => {
        const c = clienti.find(x=>x.id===p.idCliente);
        return `🕐 ${p.oraInizio} — ${c ? c.nome+" "+c.cognome : p.nomeCliente}`;
      }).join("\n");
    _wa(CONFIG.ADMIN_WHATSAPP, `📋 *Agenda di domani — ${domaniLabel}*\n\n${righe}\n\nTotale: ${predomani.length} sessioni`);
  } else {
    _wa(CONFIG.ADMIN_WHATSAPP, `📋 *Agenda di domani — ${domaniLabel}*\n\nNessuna sessione in programma.`);
  }

  // Auguri di compleanno automatici: confronta giorno/mese di dataNascita con oggi.
  // NB: confronto fatto SOLO su testo (substring), senza passare da new Date(dataNascita).
  // Motivo: new Date("yyyy-MM-dd") viene interpretata come mezzanotte UTC, e poi
  // getDate()/getMonth() la riconvertono nel fuso orario locale: a seconda del fuso
  // impostato sul progetto Apps Script questo puo' far scivolare il giorno letto
  // indietro di una unita' (bug riscontrato il 23/7: messaggio arrivato un giorno
  // prima del vero compleanno). dataNascita e' gia' una stringa "yyyy-MM-dd" (vedi
  // _tuttiClienti/_valStr), quindi mese e giorno si leggono direttamente dal testo.
  const oggiTz  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM-dd");
  clienti.forEach(c => {
    if (!c.dataNascita || c.dataNascita.length < 10) return; // vuota o non nel formato atteso, salta senza errori
    const meseGiornoNascita = c.dataNascita.substring(5, 10); // es. "2026-07-22" -> "07-22"
    if (meseGiornoNascita === oggiTz) {
      _wa(c.telefono, `🎉 *${CONFIG.STUDIO_NAME}*\nTanti auguri di buon compleanno, ${c.nome}! 🎂\nDa tutto il team ti auguriamo una splendida giornata! 💪`);
    }
  });
}

function controllaScadenze() {
  const oggi=new Date(), tra7=new Date(); tra7.setDate(oggi.getDate()+7);
  const tz=Session.getScriptTimeZone();
  _tuttiClienti().forEach(c => {
    if(c.stato!=="Attivo")return;
    const scad=new Date(c.dataScad);
    if(scad>=oggi&&scad<=tra7) {
      const dataLabel=_formatDataIT(scad, "d MMMM yyyy");
      _wa(c.telefono,`⚠️ *${CONFIG.STUDIO_NAME}*\nCiao ${c.nome}! Il tuo abbonamento *${c.pacchetto}* scade il *${dataLabel}*.\nHai ancora ${c.lezioniRim} lezioni disponibili.\nContattaci per rinnovare! 💪`);
    }
    // Promemoria scadenza certificato medico (stesso schema di 7 giorni prima, solo avviso — non blocca le prenotazioni)
    if (c.scadCertificato) {
      const scadCert=new Date(c.scadCertificato);
      if (scadCert>=oggi && scadCert<=tra7) {
        const dataCertLabel=_formatDataIT(scadCert, "d MMMM yyyy");
        _wa(c.telefono,`🏥 *${CONFIG.STUDIO_NAME}*\nCiao ${c.nome}! Il tuo certificato medico sportivo scade il *${dataCertLabel}*.\nRicordati di rinnovarlo e di portarci la copia aggiornata. 📋`);
      }
    }
  });
}

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("inviaReminder").timeBased().everyDays(1).atHour(CONFIG.ORE_REMINDER).create();
  ScriptApp.newTrigger("controllaScadenze").timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
  ScriptApp.newTrigger("controllaScadenzeListaAttesa").timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger("drainCodaWA").timeBased().everyMinutes(1).create();   // <-- NUOVA
  ScriptApp.newTrigger("drainCodaCal").timeBased().everyMinutes(1).create(); // <-- aggiunto 15/7: mancava, esisteva solo come trigger creato a mano
  Logger.log("Trigger attivati! Reminder giornaliero alle " + CONFIG.ORE_REMINDER + ":00, controllo scadenze ogni lunedì, lista attesa ogni 10 minuti.");
}

// ──────────────────────────────────────────────────────────
//  WHATSAPP
// ──────────────────────────────────────────────────────────
function _wa(numero, msg) {
  // Accoda il messaggio: l'invio vero avviene in background (drainCodaWA via trigger),
  // cosi' la risposta all'utente non aspetta piu' la chiamata HTTP a WAAPI.
  if (!CONFIG.WAAPI_TOKEN) return;
  try {
    getFogliCodaWA().appendRow([new Date(), String(numero), msg, "In coda", 0, ""]);
  } catch(e) {
    Logger.log("WA enqueue error: " + e.message);
  }
}

// Foglio coda WhatsApp (creato al volo se non esiste)
function getFogliCodaWA() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("CodaWA");
  if (!sh) {
    sh = ss.insertSheet("CodaWA");
    sh.getRange(1, 1, 1, 6).setValues([["Timestamp", "Telefono", "Messaggio", "Stato", "Tentativi", "Inviato"]]);
  }
  return sh;
}

// Invio effettivo del messaggio WhatsApp (chiamato SOLO dal drain in background)
function _waSend(numero, msg) {
  if (!CONFIG.WAAPI_TOKEN || !CONFIG.WAAPI_INSTANCE_ID) return false;
  try {
    const tel = String(numero).replace(/^\+/, "");
    const url = CONFIG.WAAPI_URL + CONFIG.WAAPI_INSTANCE_ID + "/client/action/send-message";
    const payload = JSON.stringify({ chatId: tel + "@c.us", message: msg });
    const resp = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + CONFIG.WAAPI_TOKEN },
      payload: payload,
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    return code >= 200 && code < 300;
  } catch(e) {
    Logger.log("WA send error: " + e.message);
    return false;
  }
}

// Svuota la coda WhatsApp: spedisce i messaggi "In coda".
// Eseguito da un trigger a tempo (ogni minuto) -> fuori dal percorso critico.
function drainCodaWA() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return; // un altro drain e' gia' in corso
  try {
    const sh = getFogliCodaWA();
    const rows = sh.getDataRange().getValues();
    const MAX_PER_RUN = 24; // con la pausa di 15s tra invii (richiesta da WAAPI), 24 msg/run restano dentro il limite di 6 min di Apps Script
    const MAX_TENTATIVI = 3;
    let processate = 0;
    for (let i = 1; i < rows.length && processate < MAX_PER_RUN; i++) {
      if (rows[i][3] !== "In coda") continue;
      const ok = _waSend(rows[i][1], rows[i][2]);
      processate++;
      if (ok) {
        sh.getRange(i + 1, 4).setValue("Inviato");
        sh.getRange(i + 1, 6).setValue(new Date());
      } else {
        const tentativi = (Number(rows[i][4]) || 0) + 1;
        sh.getRange(i + 1, 5).setValue(tentativi);
        if (tentativi >= MAX_TENTATIVI) sh.getRange(i + 1, 4).setValue("Errore");
      }
      // WAAPI raccomanda almeno 15s tra un messaggio e l'altro per non rischiare il ban del numero
      if (processate < MAX_PER_RUN) Utilities.sleep(15000);
    }
  } finally {
    lock.releaseLock();
  }
}
// ============================================================
// LIMITI CAPIENZA PERSONALIZZATI
// Da aggiungere in apps_script.gs
// ============================================================

// ------------------------------------------------------------
// Recupera il foglio Limiti (lo crea se non esiste)
// ------------------------------------------------------------
function getFogliLimiti() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let foglio = ss.getSheetByName("Limiti");
  if (!foglio) {
    foglio = ss.insertSheet("Limiti");
    foglio.getRange(1, 1, 1, 9).setValues([[
      "ID", "Tipo", "GiornoSettimana", "Data",
      "OraInizio", "OraFine", "MaxPersone", "Motivo", "Attivo"
    ]]);
  }
  return foglio;
}

// ------------------------------------------------------------
// Legge tutti i limiti attivi dal foglio
// ------------------------------------------------------------
function getLimitiAttivi() {
  const foglio = getFogliLimiti();
  const dati = foglio.getDataRange().getValues();
  if (dati.length <= 1) return [];

  return dati.slice(1)
    .filter(r => r[8] === true || r[8] === "TRUE" || r[8] === 1)
    .map(r => ({
      id:              r[0],
      tipo:            r[1],           // "ricorrente" | "una_tantum"
      giornoSettimana: r[2],           // 1=lun … 5=ven (solo per ricorrente)
      data:            r[3],           // oggetto Date o stringa (solo per una_tantum)
      oraInizio:       r[4],           // "09:00"
      oraFine:         r[5],           // "11:00"
      maxPersone:      parseInt(r[6]),
      motivo:          r[7],
      attivo:          true
    }));
}

// ------------------------------------------------------------
// Dato uno slot (dataObj Date, oraInizio "HH:MM"),
// restituisce la capienza massima applicabile.
// Se nessun limite copre quello slot → restituisce CONFIG.MAX_CONTEMPORANEI
// Se più limiti si sovrappongono → prende il più restrittivo
// ------------------------------------------------------------
function getCapienzaSlot(dataObj, oraInizioSlot, limiti) {
  if (!limiti || limiti.length === 0) return CONFIG.MAX_CONTEMPORANEI;
  

  // Giorno della settimana: 1=lun, 2=mar … 5=ven
  const giornoJS = dataObj.getDay(); // 0=dom, 1=lun … 6=sab
  const giornoIT = giornoJS === 0 ? 7 : giornoJS; // converti in 1-7

  // Normalizza orario slot in minuti dalla mezzanotte
  const slotMin = orarioInMinuti(oraInizioSlot);

  let capienzaMinima = CONFIG.MAX_CONTEMPORANEI;

  for (const l of limiti) {
    const inizioLimite = orarioInMinuti(l.oraInizio);
    const fineLimite   = orarioInMinuti(l.oraFine);

    // Lo slot deve essere DENTRO la fascia oraria del limite
    if (slotMin < inizioLimite || slotMin >= fineLimite) continue;

    let applicabile = false;

    if (l.tipo === "ricorrente") {
      applicabile = (parseInt(l.giornoSettimana) === giornoIT);
    } else if (l.tipo === "una_tantum") {
      const dataLimite = new Date(l.data);
      applicabile = (
        dataLimite.getFullYear() === dataObj.getFullYear() &&
        dataLimite.getMonth()    === dataObj.getMonth()    &&
        dataLimite.getDate()     === dataObj.getDate()
      );
    }

    if (applicabile) {
      capienzaMinima = Math.min(capienzaMinima, l.maxPersone);
    }
  }

  return capienzaMinima;
}

// Utility: converte "HH:MM" in minuti dalla mezzanotte
function orarioInMinuti(orario) {
  const parti = String(orario).split(":");
  return parseInt(parti[0]) * 60 + parseInt(parti[1]);
}

// ------------------------------------------------------------
// ENDPOINT — chiamato dall'admin per gestire i limiti
// Aggiungere nel doPost() / doGet() esistente come nuovo'action'
// ------------------------------------------------------------

// action: "getLimiti"
function handleGetLimiti() {
  const foglio = getFogliLimiti();
  const dati = foglio.getDataRange().getValues();
  if (dati.length <= 1) return { limiti: [] };

  const limiti = dati.slice(1).map((r, i) => ({
    riga:            i + 2,
    id:              r[0],
    tipo:            r[1],
    giornoSettimana: r[2],
    data:            r[3] ? Utilities.formatDate(new Date(r[3]), "Europe/Rome", "yyyy-MM-dd") : "",
    oraInizio:       r[4],
    oraFine:         r[5],
    maxPersone:      r[6],
    motivo:          r[7],
    attivo:          r[8] === true || r[8] === "TRUE" || r[8] === 1
  }));

  return { limiti };
}

// action: "aggiungiLimite"
function handleAggiungiLimite(params) {
  const foglio = getFogliLimiti();
  const id = "LIM_" + new Date().getTime();
  foglio.appendRow([
    id,
    params.tipo,
    params.tipo === "ricorrente" ? parseInt(params.giornoSettimana) : "",
    params.tipo === "una_tantum" ? params.data : "",
    params.oraInizio,
    params.oraFine,
    parseInt(params.maxPersone),
    params.motivo || "",
    true
  ]);
  return { success: true, id };
}

// action: "toggleLimite"
function handleToggleLimite(params) {
  const foglio = getFogliLimiti();
  const dati = foglio.getDataRange().getValues();
  for (let i = 1; i < dati.length; i++) {
    if (dati[i][0] === params.id) {
      const nuovoStato = !(dati[i][8] === true || dati[i][8] === "TRUE" || dati[i][8] === 1);
      foglio.getRange(i + 1, 9).setValue(nuovoStato);
      return { success: true, attivo: nuovoStato };
    }
  }
  return { success: false, errore: "Limite non trovato" };
}

// action: "eliminaLimite"
function handleEliminaLimite(params) {
  const foglio = getFogliLimiti();
  const dati = foglio.getDataRange().getValues();
  for (let i = 1; i < dati.length; i++) {
    if (dati[i][0] === params.id) {
      foglio.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, errore: "Limite non trovato" };
}
function testWaapi() {
  const tel = String(CONFIG.ADMIN_WHATSAPP).replace(/^\+/, "");
  const url = CONFIG.WAAPI_URL + CONFIG.WAAPI_INSTANCE_ID + "/client/action/send-message";
  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: { "Authorization": "Bearer " + CONFIG.WAAPI_TOKEN },
    payload: JSON.stringify({ chatId: tel + "@c.us", message: "Test WAAPI " + new Date() }),
    muteHttpExceptions: true
  });
  Logger.log("HTTP " + resp.getResponseCode());
  Logger.log(resp.getContentText());
}