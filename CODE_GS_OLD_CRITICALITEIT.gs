//
// V 1.5 Oversteyns Robrecht 9 3903
//
// ---------------- CONFIG ----------------
const SHEET_NAME = "Onderdelen";

// ---------------- HELPERS ----------------
function getSheet() {
  try {
    // Gebruik de ACTIEVE spreadsheet 
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);
    
    if (!sh) {
      throw new Error("Tabblad '" + SHEET_NAME + "' niet gevonden. Zorg dat het tabblad 'Onderdelen' bestaat.");
    }
    
    return sh;
  } catch (e) {
    Logger.log("ERROR in getSheet: " + e.toString());
    throw new Error("Kan tabblad niet openen: " + e.message);
  }
}

function ensureHeaders(sheet) {
  const expected = [
    "Datum","Naam","RemaxNummer","Kosten","Levertijd",
    "FalenFreq","FalenSeverity","MachineCrit",
    "AankoopFactor","Risico","FunctieCrit","Categorie","Gebruiker"
  ];
  
  const firstRow = sheet.getRange(1,1,1, expected.length).getValues()[0];
  const allEmpty = firstRow.every(c => c === "" || c === null);
  
  if (allEmpty) {
    sheet.getRange(1,1,1,expected.length).setValues([expected]);
  } else {
    // Check of de eerste kolom "Datum" is
    if (firstRow[0] !== "Datum") {
      sheet.insertRowBefore(1);
      sheet.getRange(1,1,1,expected.length).setValues([expected]);
    }
  }
}

// ---------------- POPUP VOOR SHEET ----------------
function showFormPopup() {
  const html = HtmlService.createHtmlOutputFromFile("FormPopup")
    .setWidth(500)
    .setHeight(650)
    .setTitle("Nieuw onderdeel toevoegen");
  
  SpreadsheetApp.getUi().showModalDialog(html, "Spare Parts QS Manager");
}

// Functie om vanuit het menu te openen
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 Spare Parts')
    .addItem('Nieuw onderdeel toevoegen', 'showFormPopup')
    .addToUi();
}

// ---------------- USER INFO ----------------
function getUserInfo3() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    return {
      email: userEmail,
      name: userEmail.split('@')[0] // Gebruik eerste deel van email als naam
    };
  } catch (e) {
    Logger.log("ERROR in getUserInfo: " + e.toString());
    return {
      email: "Onbekend",
      name: "Onbekend"
    };
  }
}

// ---------------- CLASSIFICATION LOGICA (TABLE LOOKUPS) ----------------
function classifyPart(data) {
  // Converteer kosten tekst naar numerieke waarde voor berekening
  let cost = 0;
  const kostenText = (data.Kosten || "").toString();
  
  if (kostenText === "Minder dan €1000") {
    cost = 999;
  } else if (kostenText === "Tussen €1000 en €5000") {
    cost = 2500;
  } else if (kostenText === "Tussen €5000 en €10000") {
    cost = 7500;
  } else if (kostenText === "Meer dan €10000") {
    cost = 15000;
  } else {
    // Fallback voor oude numerieke waarden 
    cost = parseFloat(data.Kosten) || 0;
  }
  
  const lead = parseFloat(data.Levertijd) || 0;
  const freq = (data.FalenFreq || "").toString();
  const impact = parseFloat(data.FalenSeverity) || 0;
  let crit = (data.MachineCrit || "").toString().toUpperCase();
  if (!crit) crit = "C";

  Logger.log("classifyPart input: cost=" + cost + " (van: " + kostenText + "), lead=" + lead + ", freq=" + freq + ", impact=" + impact + ", crit=" + crit);

  // --- TABEL 1: cost vs delivery (index 1..4)
  function costBucket(c) {
    if (c < 1000) return 0;
    if (c >= 1000 && c < 5000) return 1;
    if (c >= 5000 && c <= 10000) return 2;
    return 3;
  }
  
  function deliveryBucket(w) {
    if (w < 2) return 0;
    if (w >= 2 && w < 6) return 1;
    if (w >= 6 && w < 10) return 2;
    return 3;
  }

  const TABLE1 = [
    [2,2,1,1],
    [2,2,2,1],
    [3,3,3,2],
    [4,3,3,3]
  ];

  const cIdx = costBucket(cost);
  const dIdx = deliveryBucket(lead);
  const table1Index = TABLE1[dIdx][cIdx];
  Logger.log("TABLE1: costBucket=" + cIdx + ", deliveryBucket=" + dIdx + ", table1Index=" + table1Index);

  // --- TABEL Risicofactor
  let riskFactor = "C";
  const f = freq.toLowerCase();
  
  if (impact === 0) {
    riskFactor = "C";
  } else if (impact > 8) {
    if (f === ">1/week" || f === ">1/maand") riskFactor = "AA";
    else if (f === ">1/jaar") riskFactor = "A";
    else riskFactor = "B";
  } else if (impact >= 2 && impact <= 8) {
    if (f === ">1/week") riskFactor = "AA";
    else if (f === ">1/maand") riskFactor = "A";
    else if (f === ">1/jaar") riskFactor = "B";
    else riskFactor = "C";
  } else if (impact < 2) {
    if (f === ">1/week" || f === ">1/maand") riskFactor = "B";
    else if (f === ">1/jaar") riskFactor = "B";
    else riskFactor = "C";
  }
  
  Logger.log("Risicofactor: " + riskFactor);

  // --- TABEL 2: functie-criticaliteit
  const funcMatrix = {
    "AA": { "AA":4, "A":4, "B":3, "C":2 },
    "A":  { "AA":4, "A":4, "B":3, "C":2 },
    "B":  { "AA":3, "A":3, "B":2, "C":1 },
    "C":  { "AA":2, "A":2, "B":1, "C":1 }
  };

  const validCrit = funcMatrix[crit] ? crit : "C";
  const validRisk = funcMatrix[validCrit][riskFactor] ? riskFactor : "C";
  const functieCrit = funcMatrix[validCrit][validRisk] || 1;
  Logger.log("Functie-criticaliteit: " + functieCrit);

  // --- TABEL 3: QS categorie
  const TABLE3 = [
    ["QS0","QS0","QS0","QS0"],
    ["QS0","QS0","QS1","QS1"],
    ["QS1","QS1","QS2","QS2"],
    ["QS1","QS2","QS3","QS3"]
  ];

  const row = Math.max(1, Math.min(4, functieCrit)) - 1;
  const col = Math.max(1, Math.min(4, table1Index)) - 1;
  const category = TABLE3[row][col] || "QS0";
  Logger.log("QS category: " + category);

  return {
    category: category,
    table1Index: table1Index,
    risk: riskFactor,
    table2Index: functieCrit
  };
}

// ---------------- CRUD ----------------
function savePart(data) {
  try {
    const sheet = getSheet();
    if (!sheet) throw new Error("Tabblad 'Onderdelen' niet gevonden");
    
    ensureHeaders(sheet);

    const result = classifyPart(data);
    
    // Haal gebruikersinfo op
    const userInfo = getUserInfo();

    // Volgorde headers (zie ensureHeaders)
    const row = [
      new Date(),
      data.Naam || "",
      data.RemaxNummer || "",
      data.Kosten || "",
      data.Levertijd || "",
      data.FalenFreq || "",
      data.FalenSeverity || "",
      data.MachineCrit || "",
      result.table1Index,
      result.risk,
      result.table2Index,
      result.category,
      userInfo.email // Gebruiker email toevoegen
    ];

    sheet.appendRow(row);
    Logger.log("Part succesvol opgeslagen: " + data.Naam + " door " + userInfo.email);
    return result;
  } catch (e) {
    Logger.log("ERROR in savePart: " + e.toString());
    throw e;
  }
}

function getParts() {
  try {
    Logger.log("=== getParts START ===");
    
    const sheet = getSheet();
    if (!sheet) {
      Logger.log("WARN: Tabblad niet gevonden, return lege array");
      return [];
    }
    
    const lastRow = sheet.getLastRow();
    Logger.log("LastRow: " + lastRow);
    
    // Als er alleen een header is of helemaal niets
    if (lastRow <= 1) {
      Logger.log("Geen data rijen (alleen header of leeg)");
      return [];
    }
    
    const lastCol = 13;
    const dataRowCount = lastRow - 1;
    
    Logger.log("Ophalen range: 2,1," + dataRowCount + "," + lastCol);
    const values = sheet.getRange(2, 1, dataRowCount, lastCol).getValues();
    Logger.log("Rijen opgehaald: " + values.length);
    
    const parts = [];
    for (var i = 0; i < values.length; i++) {
      var r = values[i];
      // Filter lege rijen
      if (!r[1] || r[1] === '') continue;
      
      parts.push({
        rowIndex: i + 2, // Rijnummer in sheet (header = 1, data start bij 2)
        Datum: r[0] ? r[0].toString() : '',
        Naam: r[1] ? r[1].toString() : '',
        RemaxNummer: r[2] ? r[2].toString() : '',
        Kosten: r[3] ? r[3].toString() : '',
        Levertijd: r[4] ? r[4].toString() : '',
        FalenFreq: r[5] ? r[5].toString() : '',
        FalenSeverity: r[6] ? r[6].toString() : '',
        MachineCrit: r[7] ? r[7].toString() : '',
        AankoopFactor: r[8] ? r[8].toString() : '',
        Risico: r[9] ? r[9].toString() : '',
        FunctieCrit: r[10] ? r[10].toString() : '',
        Categorie: r[11] ? r[11].toString() : '',
        Gebruiker: r[12] ? r[12].toString() : ''
      });
    }
    
    Logger.log("Parts na filtering: " + parts.length);
    Logger.log("=== getParts END ===");
    return parts;
    
  } catch (e) {
    Logger.log("ERROR in getParts: " + e.toString());
    Logger.log("Stack: " + e.stack);
    throw new Error("Fout bij ophalen onderdelen: " + e.message);
  }
}

function deletePart(rowIndex) {
  try {
    Logger.log("=== deletePart START voor rij " + rowIndex + " ===");
    
    const sheet = getSheet();
    if (!sheet) throw new Error("Tabblad 'Onderdelen' niet gevonden");
    
    // Validatie
    if (!rowIndex || rowIndex < 2) {
      throw new Error("Ongeldig rijnummer: " + rowIndex);
    }
    
    const lastRow = sheet.getLastRow();
    if (rowIndex > lastRow) {
      throw new Error("Rijnummer bestaat niet: " + rowIndex);
    }
    
    // Haal de naam op voor logging
    const naam = sheet.getRange(rowIndex, 2).getValue();
    
    // Verwijder de rij
    sheet.deleteRow(rowIndex);
    
    Logger.log("Rij " + rowIndex + " succesvol verwijderd (onderdeel: " + naam + ")");
    
    return {
      success: true,
      message: "Onderdeel '" + naam + "' succesvol verwijderd"
    };
    
  } catch (e) {
    Logger.log("ERROR in deletePart: " + e.toString());
    throw new Error("Fout bij verwijderen: " + e.message);
  }
}

// 🧪 TEST FUNCTIES
function testGetParts() {
  Logger.log("=== TEST START ===");
  try {
    const parts = getParts();
    Logger.log("Resultaat: " + JSON.stringify(parts));
    Logger.log("Aantal: " + parts.length);
  } catch (e) {
    Logger.log("TEST FAILED: " + e.toString());
  }
  Logger.log("=== TEST END ===");
}

function testSheetAccess() {
  Logger.log("=== TEST SHEET ACCESS ===");
  try {
    const sheet = getSheet();
    Logger.log("Sheet naam: " + sheet.getName());
    Logger.log("Laatste rij: " + sheet.getLastRow());
    Logger.log("Laatste kolom: " + sheet.getLastColumn());
    Logger.log("Sheet toegankelijk: JA");
  } catch (e) {
    Logger.log("Sheet NIET toegankelijk: " + e.toString());
  }
}

function getUserInfo() {
  try {
    // Probeer eerst de actieve gebruiker
    let userEmail = Session.getActiveUser().getEmail();
    
    // Als dat niet werkt, gebruik effectieve gebruiker
    if (!userEmail || userEmail === "" || userEmail === "Onbekend") {
      userEmail = Session.getEffectiveUser().getEmail();
    }
    
    // Als nog steeds leeg, gebruik eigenaar van spreadsheet
    if (!userEmail || userEmail === "") {
      userEmail = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail();
    }
    
    Logger.log("Gebruiker email gevonden: " + userEmail);
    
    return {
      email: userEmail,
      name: userEmail.split('@')[0] // Gebruik eerste deel van email als naam
    };
  } catch (e) {
    Logger.log("ERROR in getUserInfo: " + e.toString());
    // Fallback: gebruik de eigenaar van de spreadsheet
    try {
      const ownerEmail = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail();
      Logger.log("Fallback naar eigenaar: " + ownerEmail);
      return {
        email: ownerEmail,
        name: ownerEmail.split('@')[0]
      };
    } catch (e2) {
      Logger.log("ERROR bij fallback: " + e2.toString());
      return {
        email: "Onbekend",
        name: "Onbekend"
      };
    }
  }
}


function testSavePart() {
  Logger.log("=== TEST SAVE PART ===");
  try {
    const testData = {
      Naam: "Test Onderdeel",
      RemaxNummer: "TEST-123",
      Kosten: "500",
      Levertijd: "4",
      FalenFreq: ">1/maand",
      FalenSeverity: "6",
      MachineCrit: "A"
    };
    
    const result = savePart(testData);
    Logger.log("Test onderdeel opgeslagen: " + JSON.stringify(result));
  } catch (e) {
    Logger.log("TEST FAILED: " + e.toString());
  }
}