// =========================================================
// RC-Snackbar eSG - Google Apps Script Backend
// Version: Beta v1.0.8 (Fix: Keine undefined-Einträge & GET Auth)
// =========================================================

// GET-Request: Verifizierung, Produktliste oder manuelle Archivierung
function doGet(e) {
  try {
    var action = e && e.parameter ? e.parameter.action : null;

    // A) Schüler-Verifizierung über GET (Schreibt niemals in die Tabelle)
    if (action === "verifyStudent") {
      var username = e.parameter.username;
      var userClass = e.parameter.userClass;
      var verifyResult = checkStudentCredentials(username, userClass);
      return ContentService.createTextOutput(JSON.stringify(verifyResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // B) Manuelle Archivierung per URL (?action=archive&row=2)
    if (action === "archive") {
      var rowParam = e.parameter.row;
      var row = parseInt(rowParam, 10);
      return archiveOrderRowManually(row);
    }

    // C) Standard: Produktliste abrufen
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Produkte");
    
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ error: "Tabellenblatt 'Produkte' nicht gefunden." }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var data = sheet.getDataRange().getValues();
    var products = [];

    for (var i = 1; i < data.length; i++) {
      var rowData = data[i];
      if (rowData[0] !== "" && rowData[0] !== undefined) {
        products.push({
          name: rowData[0],
          price: rowData[1],
          image: rowData[2] || "",
          stock: rowData[3] !== "" ? rowData[3] : 99,
          isAction: rowData[4] === true || rowData[4] === "TRUE" || rowData[4] === "true"
        });
      }
    }

    return ContentService.createTextOutput(JSON.stringify(products))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// POST-Request: AUSSCHLIESSLICH für echte Bestellungen
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Keine Daten empfangen." }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var data = JSON.parse(e.postData.contents);

    // Strikte Abweisung: Wenn Name, Klasse oder Artikel fehlen oder undefined sind, KEIN Tabelleneintrag!
    if (!data.studentName || data.studentName === "undefined" || !data.items || data.items === "undefined" || (Array.isArray(data.items) && data.items.length === 0)) {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "error", 
        message: "Ungültige Bestellung. Keine Speicherung durchgeführt." 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Bestellungen");

    if (!sheet) {
      sheet = ss.insertSheet("Bestellungen");
      sheet.appendRow(["Zeitstempel", "Name", "Klasse", "Pause", "Artikel", "Gesamtsumme", "Status"]);
    }

    var timestamp = new Date();
    var itemsFormatted = Array.isArray(data.items) ? data.items.join(", ") : data.items;
    var pickupInfo = (data.pickupDate || "") + (data.pickupPause ? " (" + data.pickupPause + ")" : "");

    sheet.appendRow([
      timestamp,
      data.studentName,
      data.studentClass || "",
      pickupInfo,
      itemsFormatted,
      (data.totalPrice || "0,00") + " €",
      "Offen"
    ]);

    autoArchivieren();
    saveOrderToDriveFolder(data, timestamp);

    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Prüft Name und Klasse aus Spalte K (Name) und L (Klasse) im Tabellenblatt 'Produkte'
function checkStudentCredentials(inputName, inputClass) {
  if (!inputName || !inputClass) {
    return { success: false, error: "Bitte Name und Klasse vollständig angeben." };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Produkte");
  if (!sheet) {
    return { success: false, error: "Tabellenblatt 'Produkte' nicht gefunden." };
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    return { success: false, error: "Schülerliste ist leer." };
  }

  // Liest gezielt Spalte K (Index 0) und L (Index 1)
  var studentData = sheet.getRange("K1:L" + lastRow).getValues();
  var searchName = String(inputName).trim().toLowerCase();
  var searchClass = String(inputClass).trim().toLowerCase();

  for (var i = 0; i < studentData.length; i++) {
    var studentNameInSheet = String(studentData[i][0] || "").trim().toLowerCase();
    var studentClassInSheet = String(studentData[i][1] || "").trim().toLowerCase();

    if (studentNameInSheet === searchName && studentClassInSheet === searchClass) {
      return { 
        success: true, 
        officialName: String(studentData[i][0]).trim(), 
        officialClass: String(studentData[i][1]).trim() 
      };
    }
  }

  return { 
    success: false, 
    error: "Name oder Klasse nicht in der Schülerliste gefunden." 
  };
}

// Automatische E-Archivierung
function autoArchivieren() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bestellungenSheet = ss.getSheetByName("Bestellungen");
  if (!bestellungenSheet) return;

  var archivSheet = ss.getSheetByName("Archiv");
  if (!archivSheet) {
    archivSheet = ss.insertSheet("Archiv");
    archivSheet.appendRow(["Zeitstempel", "Name", "Klasse", "Pause", "Artikel", "Gesamtsumme", "Status"]);
  }

  var data = bestellungenSheet.getDataRange().getValues();

  for (var i = data.length - 1; i >= 1; i--) {
    var status = String(data[i][6] || "").trim();

    if (status.toUpperCase() === "E" || status.toLowerCase() === "erledigt") {
      var rowToMove = data[i];
      rowToMove[6] = "Erledigt";
      archivSheet.appendRow(rowToMove);
      bestellungenSheet.deleteRow(i + 1);
    }
  }
}

// Manuelle Archivierung per URL
function archiveOrderRowManually(rowIndex) {
  if (isNaN(rowIndex) || rowIndex === null || rowIndex < 2) {
    return ContentService.createTextOutput("Fehler: Ungültige Zeilennummer.").setMimeType(ContentService.MimeType.TEXT);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("Bestellungen");
  if (!sourceSheet) return ContentService.createTextOutput("Fehler: Blatt 'Bestellungen' nicht gefunden.").setMimeType(ContentService.MimeType.TEXT);

  var targetSheet = ss.getSheetByName("Archiv");
  if (!targetSheet) {
    targetSheet = ss.insertSheet("Archiv");
    var headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
    targetSheet.appendRow(headers);
  }

  var numCols = sourceSheet.getLastColumn();
  var rowData = sourceSheet.getRange(rowIndex, 1, 1, numCols).getValues()[0];
  rowData[6] = "Erledigt";

  targetSheet.appendRow(rowData);
  sourceSheet.deleteRow(rowIndex);

  return ContentService.createTextOutput("Zeile " + rowIndex + " erfolgreich archiviert!").setMimeType(ContentService.MimeType.TEXT);
}

function saveOrderToDriveFolder(data, timestamp) {
  try {
    var folderName = "RC-Snackbar_Bestellungen";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    var orderId = "RC-" + Math.floor(100000 + Math.random() * 900000);
    var fileName = "Bestellung_" + String(data.studentName).replace(/\s+/g, "_") + "_" + orderId + ".txt";
    
    var content = "RC-SNACKBAR eSG - BESTELLUNG\n" +
                  "===============================\n" +
                  "Bestell-ID: " + orderId + "\n" +
                  "Eingegangen: " + timestamp.toLocaleString("de-DE") + "\n" +
                  "Name: " + data.studentName + "\n" +
                  "Klasse: " + data.studentClass + "\n" +
                  "-------------------------------\n" +
                  "Artikel: " + (Array.isArray(data.items) ? data.items.join(", ") : data.items) + "\n" +
                  "Gesamtsumme: " + data.totalPrice + " €\n" +
                  "===============================";

    folder.createFile(fileName, content, MimeType.PLAIN_TEXT);
  } catch (e) {
    Logger.log("Drive-Speicherfehler: " + e.toString());
  }
}
