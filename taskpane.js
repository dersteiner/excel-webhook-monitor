// Excel Webhook Monitor - Mit Auto-Minimize für Excel Online
// Minimiert Panel automatisch nach Start

console.log("🔥 taskpane.js wird geladen...");

const PROXY_URL = "https://autumn-sea-2657.daniel-steiner-mail.workers.dev";
const API_KEY = "akdsadhoiadoiwoqi8wd";

const ALLOWED_FILES = [
  "Tracking.xlsx",
  "Projektliste.xlsx",
  "KOPIE",
  "2025",
	"Anfragen"
];

console.log("🔥 Konfiguration geladen:", { PROXY_URL, ALLOWED_FILES });

Office.onReady((info) => {
  console.log("🔥 Office.onReady aufgerufen!", info);
  
  if (info.host === Office.HostType.Excel) {
    console.log("✅ Excel Host erkannt");
    addLog("✅ Excel Webhook Monitor geladen", "success");
    
    // Prüfe und starte Monitoring
    checkAndStartMonitoring();
    
    // WICHTIG: Minimiere Panel nach 3 Sekunden (nur bei Autostart)
    // User kann es manuell wieder öffnen wenn er den Status sehen will
    setTimeout(() => {
      try {
        // Versuche Panel zu minimieren (funktioniert nicht in allen Szenarien)
        if (Office.context.ui && Office.context.ui.closeContainer) {
          console.log("💡 Minimiere Panel automatisch");
          addLog("💡 Panel minimiert - Monitoring läuft im Hintergrund");
          // Office.context.ui.closeContainer(); // Würde komplett schließen
        }
      } catch (e) {
        console.log("ℹ️ Konnte Panel nicht minimieren (normal bei manuellem Öffnen)");
      }
    }, 3000);
    
  } else {
    console.log("⚠️ Kein Excel Host:", info.host);
    addLog("⚠️ Nicht in Excel geöffnet", "error");
  }
});

console.log("🔥 Office.onReady registriert");

async function checkAndStartMonitoring() {
  console.log("🔍 Starte checkAndStartMonitoring()");
  addLog("🔍 Prüfe Dateinamen...");
  
  try {
    await Excel.run(async (context) => {
      console.log("📊 Excel.run gestartet");
      
      const workbook = context.workbook;
      workbook.load("name");
      await context.sync();
      
      const fileName = workbook.name;
      console.log("📄 Geöffnete Datei:", fileName);
      addLog("📄 Datei: " + fileName);
      
      console.log("🔍 Prüfe gegen Liste:", ALLOWED_FILES);
      let matchFound = false;
      
      for (const allowedFile of ALLOWED_FILES) {
        const matches = fileName.toLowerCase().includes(allowedFile.toLowerCase());
        console.log(`  - "${allowedFile}" → ${matches ? "✅ MATCH" : "❌ kein Match"}`);
        if (matches) matchFound = true;
      }
      
      console.log("🎯 Match gefunden:", matchFound);
      
      if (matchFound) {
        console.log("✅ Diese Datei wird überwacht!");
        addLog("✅ Webhook Monitor aktiv für: " + fileName, "success");
        addLog("🔍 Überwache Spalte G...");
        addLog("💡 Du kannst dieses Panel schließen - Monitoring läuft im Hintergrund", "info");
        
        await startMonitoring();
      } else {
        console.log("⏸️ Diese Datei wird NICHT überwacht");
        addLog("⏸️ Webhook Monitor inaktiv für diese Datei");
        addLog("📋 Überwachte Dateien: " + ALLOWED_FILES.join(", "));
        addLog("💡 Dateiname muss einen dieser Strings enthalten");
      }
    });
  } catch (error) {
    console.error("❌ Fehler in checkAndStartMonitoring:", error);
    addLog("❌ Fehler beim Prüfen: " + error.message, "error");
    
    if (error.stack) {
      console.error("Stack trace:", error.stack);
    }
  }
}

async function startMonitoring() {
  console.log("🚀 Starte startMonitoring()");
  
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      console.log("📝 Registriere onChanged Handler...");
      sheet.onChanged.add(handleCellChange);
      
      await context.sync();
      console.log("✅ Event-Handler erfolgreich registriert");
      addLog("✅ Bereit - Monitoring läuft im Hintergrund!", "success");
    });
  } catch (error) {
    console.error("❌ Fehler in startMonitoring:", error);
    addLog("❌ Fehler beim Starten: " + error.message, "error");
  }
}


async function handleCellChange(event) {
  console.log("🔔 handleCellChange aufgerufen:", event);
  
  try {
    await Excel.run(async (context) => {
      const match = event.address.match(/([A-Z]+)(\d+)/);
      if (!match) {
        console.log("⚠️ Konnte Adresse nicht parsen:", event.address);
        return;
      }
      
      const column = match[1];
      const row = parseInt(match[2]);
      
      console.log(`📍 Änderung in Spalte ${column}, Zeile ${row}`);
      
      if (column !== "G") {
        console.log(`⏭️ Ignoriere Spalte ${column}`);
        return;
      }
      
      console.log("✅ Spalte G betroffen!");
      addLog(`📝 Änderung in Spalte G: Zeile ${row}`);
      
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Hole Header (Zeile 1) UND Datenzeile
      const headerRange = sheet.getRange("A1:P1");
      const dataRange = sheet.getRange(`A${row}:P${row}`);
      
      headerRange.load("values");
      dataRange.load("values");
      
      await context.sync();
      
      if (!dataRange.values || !dataRange.values[0]) {
        console.error("❌ Keine Daten gefunden");
        addLog("❌ Fehler: Zeile enthält keine Daten", "error");
        return;
      }
      
      const headers = headerRange.values[0];
      const rowData = dataRange.values[0];
      
      // Erstelle Objekt mit Spaltennamen
      const rowObject = {};
      headers.forEach((header, index) => {
        const colLetter = String.fromCharCode(65 + index); // A, B, C, ...
        const key = header || `Spalte_${colLetter}`;
        const value = rowData[index];
        rowObject[key] = (value === "" || value === undefined) ? null : value;
      });
      
      console.log(`📊 Zeile ${row} als Objekt:`, rowObject);
      
      await sendWebhook(row, rowObject);
    });
  } catch (error) {
    console.error("❌ Fehler in handleCellChange:", error);
    console.error("❌ Stack:", error.stack);
    addLog("❌ Fehler: " + error.message, "error");
  }
}


async function sendWebhook(rowNumber, rowData) {
  console.log("📤 Sende Webhook...");
  console.log("🔍 rowData type:", typeof rowData);
  console.log("🔍 rowData:", rowData);
  
  // Prüfe ob rowData ein Objekt oder Array ist
  let payload;
  
  if (Array.isArray(rowData)) {
    // Array-Format (A-P)
    console.log("✅ Array-Format erkannt");
    
    if (rowData.length === 0) {
      console.error("❌ Array ist leer");
      addLog("❌ Fehler: Keine Daten in der Zeile", "error");
      return;
    }
    
    payload = {
      row: rowNumber,
      value: rowData[6],  // Spalte G (Index 6)
      data: rowData,
      timestamp: new Date().toISOString()
    };
    
  } else if (typeof rowData === 'object' && rowData !== null) {
    // Objekt-Format (mit Spaltennamen)
    console.log("✅ Objekt-Format erkannt");
    
    const keys = Object.keys(rowData);
    if (keys.length === 0) {
      console.error("❌ Objekt ist leer");
      addLog("❌ Fehler: Keine Daten in der Zeile", "error");
      return;
    }
    
    // Finde den Wert von Spalte G
    // Der Key könnte "Spalte_G" oder der Header-Name sein
    const columnGValue = rowData['Spalte_G'] || Object.values(rowData)[6] || null;
    
    payload = {
      row: rowNumber,
      value: columnGValue,
      data: rowData,
      timestamp: new Date().toISOString()
    };
    
  } else {
    console.error("❌ rowData hat ungültiges Format:", rowData);
    addLog("❌ Fehler: Ungültiges Datenformat", "error");
    return;
  }
  
  console.log("📦 Payload:", JSON.stringify(payload, null, 2));
  addLog(`📤 Sende Webhook: Zeile ${rowNumber}`);
  
  if (!PROXY_URL || PROXY_URL.includes("DEIN-SUBDOMAIN")) {
    console.error("❌ PROXY_URL nicht konfiguriert!");
    addLog("❌ Fehler: PROXY_URL nicht konfiguriert!", "error");
    return;
  }
  
  if (!PROXY_URL.startsWith("https://")) {
    console.error("❌ PROXY_URL muss mit https:// beginnen!");
    addLog("❌ Fehler: PROXY_URL braucht https://", "error");
    return;
  }
  
  try {
    console.log("🌐 Fetch zu:", PROXY_URL);
    
    const response = await fetch(PROXY_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-API-Key': API_KEY
      },
      body: JSON.stringify(payload)
    });
    
    console.log("📨 Response Status:", response.status);
    
    const result = await response.json();
    console.log("📨 Response Body:", result);
    
    if (response.ok && result.success) {
      addLog(`✅ Webhook erfolgreich gesendet!`, "success");
      console.log("✅ Webhook erfolgreich!");
    } else {
      addLog(`⚠️ Webhook-Fehler: ${result.error || result.message}`, "error");
      console.error("⚠️ Webhook-Fehler:", result);
    }
  } catch (error) {
    console.error("❌ Fetch-Fehler:", error);
    addLog(`❌ Netzwerkfehler: ${error.message}`, "error");
    
    if (error.message.includes("Failed to fetch")) {
      addLog("💡 Prüfe: CORS, HTTPS, Worker-URL", "info");
    }
  }
}

function addLog(message, type = "") {
  console.log(`[LOG ${type}]`, message);
  
  const logDiv = document.getElementById("log");
  if (!logDiv) {
    console.warn("⚠️ Log-Div nicht gefunden!");
    return;
  }
  
  const entry = document.createElement("div");
  entry.className = "log-entry " + type;
  
  const timestamp = new Date().toLocaleTimeString("de-DE");
  entry.textContent = `[${timestamp}] ${message}`;
  
  logDiv.insertBefore(entry, logDiv.firstChild);
  
  while (logDiv.children.length > 50) {
    logDiv.removeChild(logDiv.lastChild);
  }
}

console.log("🔥 taskpane.js vollständig geladen");
console.log("💡 Öffne die Console (F12) für detaillierte Logs");
