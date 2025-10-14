// Excel Webhook Monitor - Nur für bestimmte Dateien
// Mit Cloudflare Worker Proxy

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Excel Webhook Monitor bereit");
    
    // Prüfe zuerst ob diese Datei überwacht werden soll
    checkAndStartMonitoring();
  }
});


const PROXY_URL = "https://autumn-sea-2657.daniel-steiner-mail.workers.dev";

// 2. Dein API-Key (muss mit Worker übereinstimmen!)
const API_KEY = "akdsadhoiadoiwoqi8wd";


// 3. DEFINIERE HIER: Welche Dateien sollen überwacht werden?
const ALLOWED_FILES = [
  "Tracking.xlsx",           // Exakter Dateiname
  "Projektliste.xlsx",       // Exakter Dateiname
  "KOPIE",             // Teilstring - alle Dateien mit "Kundendaten" im Namen
  "2025",                    // Alle Dateien mit "2025" im Namen
];

// Prüfe ob diese Datei überwacht werden soll
async function checkAndStartMonitoring() {
  await Excel.run(async (context) => {
    try {
      // Hole Dateiname
      const workbook = context.workbook;
      workbook.load("name");
      await context.sync();
      
      const fileName = workbook.name;
      console.log("📄 Geöffnete Datei:", fileName);
      
      // Prüfe ob Dateiname in der Whitelist ist
      const shouldMonitor = ALLOWED_FILES.some(allowedFile => 
        fileName.toLowerCase().includes(allowedFile.toLowerCase())
      );
      
      if (shouldMonitor) {
        console.log("✅ Diese Datei wird überwacht!");
        addLog("✅ Webhook Monitor aktiv für: " + fileName, "success");
        addLog("🔍 Überwache Spalte G...");
        
        // Starte Monitoring
        await startMonitoring();
      } else {
        console.log("⏸️ Diese Datei wird NICHT überwacht");
        addLog("⏸️ Webhook Monitor inaktiv für diese Datei");
        addLog("📋 Überwachte Dateien: " + ALLOWED_FILES.join(", "));
        addLog("💡 Tipp: Dateiname muss einen dieser Strings enthalten", "info");
      }
      
    } catch (error) {
      console.error("Fehler beim Prüfen:", error);
      addLog("⚠️ Fehler beim Prüfen des Dateinamens: " + error.message, "error");
    }
  });
}

// Starte die Überwachung
async function startMonitoring() {
  await Excel.run(async (context) => {
    try {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Registriere Event-Handler für Zelländerungen
      sheet.onChanged.add(handleCellChange);
      
      await context.sync();
      console.log("✅ Event-Handler registriert");
      addLog("✅ Bereit - Warte auf Änderungen in Spalte G...", "success");
      
    } catch (error) {
      console.error("Fehler beim Starten:", error);
      addLog("❌ Fehler beim Starten: " + error.message, "error");
    }
  });
}

// Handler für Zelländerungen
async function handleCellChange(event) {
  await Excel.run(async (context) => {
    try {
      // Extrahiere Spalte und Zeile aus der Adresse (z.B. "G5")
      const match = event.address.match(/([A-Z]+)(\d+)/);
      if (!match) return;
      
      const column = match[1];
      const row = parseInt(match[2]);
      
      // Nur Spalte G überwachen
      if (column !== "G") {
        return; // Ignoriere alle anderen Spalten
      }
      
      addLog(`📝 Änderung in Spalte G erkannt: Zeile ${row}`);
      
      // Hole den neuen Wert aus der Zelle
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(event.address);
      range.load("values");
      await context.sync();
      
      const cellValue = range.values[0][0];
      console.log(`Zeile ${row}, Neuer Wert: ${cellValue}`);
      
      // Sende Webhook
      await sendWebhook(row, cellValue);
      
    } catch (error) {
      console.error("Fehler beim Verarbeiten:", error);
      addLog("❌ Fehler: " + error.message, "error");
    }
  });
}

// Webhook senden via Cloudflare Worker
async function sendWebhook(rowNumber, cellContent) {
  const payload = {
    row: rowNumber,
    value: cellContent,
    timestamp: new Date().toISOString()
  };
  
  addLog(`📤 Sende Webhook via Proxy: Zeile ${rowNumber}, Wert: "${cellContent}"`);
  
  try {
    // Sende an Cloudflare Worker (NICHT direkt an Make.com!)
    const response = await fetch(PROXY_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-API-Key': API_KEY  // Authentifizierung
      },
      body: JSON.stringify(payload)
    });
    
    const result = await response.json();
    
    if (response.ok && result.success) {
      addLog(`✅ Webhook erfolgreich gesendet!`, "success");
      console.log("Proxy response:", result);
    } else {
      addLog(`⚠️ Webhook-Fehler: ${result.error || result.message}`, "error");
      console.error("Proxy error:", result);
    }
    
  } catch (error) {
    console.error("Fetch-Fehler:", error);
    addLog(`❌ Netzwerkfehler: ${error.message}`, "error");
  }
}

// Log-Eintrag hinzufügen
function addLog(message, type = "") {
  const logDiv = document.getElementById("log");
  if (!logDiv) return;
  
  const entry = document.createElement("div");
  entry.className = "log-entry " + type;
  
  const timestamp = new Date().toLocaleTimeString("de-DE");
  entry.textContent = `[${timestamp}] ${message}`;
  
  // Neuste Einträge oben
  logDiv.insertBefore(entry, logDiv.firstChild);
  
  // Maximal 50 Einträge behalten
  while (logDiv.children.length > 50) {
    logDiv.removeChild(logDiv.lastChild);
  }
}
