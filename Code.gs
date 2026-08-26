const SHEET_NAME = "Sheet1";
const CONFIG_SHEET_NAME = "Config";
const HEADER_ROW = 1;
const ADMIN_EMAILS = ['filmast@gmail.com', 'anna.raimondi@hotmail.it'];
const DASHBOARD_URL = 'https://filmast-babysitter.netlify.app/';

// ── CORS ──────────────────────────────────────────────────────────────────────

function doOptions(e) {
  var output = ContentService.createTextOutput("");
  output.setMimeType(ContentService.MimeType.TEXT);
  return output;
}

// ── GET ───────────────────────────────────────────────────────────────────────

function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : null;

  // Endpoint: verifica password al login
  if (action === 'checkPassword') {
    return handleCheckPassword(e.parameter.email, e.parameter.password);
  }

  // Default: restituisce i record del foglio Sheet1
  try {
    var sheet = SpreadsheetApp.openById(getSheetId()).getSheetByName(SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    var records = data.slice(HEADER_ROW);
    var output = ContentService.createTextOutput(JSON.stringify({
      status: 'ok',
      records: records
    }));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
  } catch (error) {
    return jsonError(error.toString());
  }
}

// ── POST ──────────────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);

    if (payload.action === 'append') {
      return handleAppend(payload);
    } else if (payload.action === 'updateStatus') {
      return handleUpdateStatus(payload);
    } else if (payload.action === 'changePassword') {
      return handleChangePassword(payload);
    }

    return jsonError('Azione non riconosciuta');
  } catch (error) {
    return jsonError(error.toString());
  }
}

// ── AUTH ──────────────────────────────────────────────────────────────────────

function handleCheckPassword(email, password) {
  try {
    if (!email || !password) {
      return jsonResult({ status: 'error', message: 'Email e password richiesti' });
    }

    var configSheet = SpreadsheetApp.openById(getSheetId()).getSheetByName(CONFIG_SHEET_NAME);
    if (!configSheet) {
      return jsonResult({ status: 'error', message: 'Foglio Config non trovato. Contatta l\'amministratore.' });
    }

    var data = configSheet.getDataRange().getValues();
    // Riga 1 = intestazioni (Email | Password | Ruolo | Nome)
    // Righe 2+ = utenti
    for (var i = 1; i < data.length; i++) {
      var rowEmail = (data[i][0] || '').toString().trim().toLowerCase();
      var rowPassword = (data[i][1] || '').toString().trim();
      var rowRole = (data[i][2] || '').toString().trim().toLowerCase();
      var rowName = (data[i][3] || '').toString().trim();

      if (rowEmail === email.trim().toLowerCase() && rowPassword === password) {
        return jsonResult({
          status: 'ok',
          role: rowRole === 'admin' ? 'admin' : 'babysitter',
          displayName: rowName || email.split('@')[0]
        });
      }
    }

    return jsonResult({ status: 'error', message: 'Email o password non corretti' });
  } catch (error) {
    return jsonResult({ status: 'error', message: error.toString() });
  }
}

function handleChangePassword(payload) {
  try {
    var email = (payload.email || '').trim().toLowerCase();
    var currentPassword = payload.currentPassword || '';
    var newPassword = payload.newPassword || '';

    if (!email || !currentPassword || !newPassword) {
      return jsonResult({ status: 'error', message: 'Dati mancanti' });
    }

    var configSheet = SpreadsheetApp.openById(getSheetId()).getSheetByName(CONFIG_SHEET_NAME);
    if (!configSheet) {
      return jsonResult({ status: 'error', message: 'Foglio Config non trovato' });
    }

    var data = configSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var rowEmail = (data[i][0] || '').toString().trim().toLowerCase();
      var rowPassword = (data[i][1] || '').toString().trim();

      if (rowEmail === email && rowPassword === currentPassword) {
        // Aggiorna la password nella colonna B (indice 1, colonna 2)
        configSheet.getRange(i + 1, 2).setValue(newPassword);
        return jsonResult({ status: 'ok' });
      }
    }

    return jsonResult({ status: 'error', message: 'Password attuale non corretta' });
  } catch (error) {
    return jsonResult({ status: 'error', message: error.toString() });
  }
}

// ── APPEND ────────────────────────────────────────────────────────────────────

function handleAppend(payload) {
  var sheet = SpreadsheetApp.openById(getSheetId()).getSheetByName(SHEET_NAME);
  // Colonne: Data | Giorno | Tipo | Ore Cont. | Ore Eff. | Ore Extra | Note | Status | Data Invio | Motivo Rifiuto | Email Babysitter
  sheet.appendRow([
    payload.data.date,
    payload.data.dayOfWeek,
    payload.data.type,
    payload.data.contractHours,
    payload.data.hours,
    payload.data.extraHours,
    payload.data.notes,
    payload.data.status,
    payload.data.submittedAt,
    payload.data.rejectionReason || '',
    payload.data.babysitterEmail || ''
  ]);

  if (payload.data.status === 'pending_approval') {
    sendAdminNotification(payload.data);
  }

  return jsonResult({ status: 'ok' });
}

// ── UPDATE STATUS ─────────────────────────────────────────────────────────────

function handleUpdateStatus(payload) {
  var sheet = SpreadsheetApp.openById(getSheetId()).getSheetByName(SHEET_NAME);
  var rowIndex = payload.rowIndex + HEADER_ROW + 1;
  var newStatus = payload.status;
  var rejectionReason = payload.rejectionReason || '';
  var babysitterEmail = payload.babysitterEmail || '';
  var babysitterDate = payload.babysitterDate || '';

  sheet.getRange(rowIndex, 8).setValue(newStatus);
  if (rejectionReason) sheet.getRange(rowIndex, 10).setValue(rejectionReason);
  if (babysitterEmail) sheet.getRange(rowIndex, 11).setValue(babysitterEmail);

  if (babysitterEmail && (newStatus === 'approved' || newStatus === 'rejected')) {
    sendBabysitterNotification(babysitterEmail, newStatus, babysitterDate, rejectionReason);
  }

  return jsonResult({ status: 'ok' });
}

// ── NOTIFICHE EMAIL ───────────────────────────────────────────────────────────

function sendAdminNotification(data) {
  try {
    var typeLabel = data.type === 'ore' ? 'Servizio' : data.type === 'ferie' ? 'Ferie' : data.type === 'malattia' ? 'Malattia' : 'Altro';
    var subject = 'Cartellino da approvare - ' + data.date + ' (' + typeLabel + ')';
    var body = 'Nuova registrazione da approvare:\n\n' +
      'Data: ' + data.date + ' (' + data.dayOfWeek + ')\n' +
      'Tipo: ' + typeLabel + '\n' +
      (data.type === 'ore' ? 'Ore effettive: ' + data.hours + '\n' : '') +
      (data.extraHours > 0 ? 'Ore extra: ' + data.extraHours + '\n' : '') +
      (data.notes ? 'Note: ' + data.notes + '\n' : '') +
      '\nVai alla dashboard:\n' + DASHBOARD_URL +
      '\n\nInviato il: ' + data.submittedAt;

    ADMIN_EMAILS.forEach(function(email) {
      MailApp.sendEmail(email, subject, body);
    });
  } catch (error) {
    Logger.log('Errore email admin: ' + error);
  }
}

function sendBabysitterNotification(babysitterEmail, status, date, rejectionReason) {
  try {
    var formattedDate = date;
    if (date && date.includes('-')) {
      var parts = date.split('-');
      if (parts.length === 3) formattedDate = parts[2] + '.' + parts[1] + '.' + parts[0];
    }

    var subject, body;
    if (status === 'approved') {
      subject = 'Registrazione approvata - ' + formattedDate;
      body = 'Ciao,\n\nFilippo e Anna hanno approvato la registrazione del ' + formattedDate + '.\n\nVai alla dashboard:\n' + DASHBOARD_URL;
    } else if (status === 'rejected') {
      subject = 'Registrazione rifiutata - ' + formattedDate;
      body = 'Ciao,\n\nFilippo e Anna hanno rifiutato la registrazione del ' + formattedDate + '.\n\nMotivazione: ' + (rejectionReason || 'nessuna motivazione specificata') + '\n\nVai alla dashboard:\n' + DASHBOARD_URL;
    }

    if (subject && body) MailApp.sendEmail(babysitterEmail, subject, body);
  } catch (error) {
    Logger.log('Errore email babysitter: ' + error);
  }
}

// ── UTILS ─────────────────────────────────────────────────────────────────────

function jsonResult(obj) {
  var output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function jsonError(message) {
  return jsonResult({ status: 'error', message: message });
}

function getSheetId() {
  return SpreadsheetApp.getActiveSheet().getParent().getId();
}
