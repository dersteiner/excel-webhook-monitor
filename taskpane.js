// Excel Webhook Monitor - Taskpane Script
// Überwacht Spalte G und sendet Events an Make.com

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Excel Webhook Monitor bereit");
    addLog("✅ Webhook Monitor gestartet", "success");
    
    // Automatisch Event-Handler registrieren
    startMonitoring();
  }
});

// Event-Handler für Zelländerungen
async function startMonitoring() {
  await Excel.run(async (context) => {
    try {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Event-Handler registrieren
      sheet.onChanged.add(handleCellChange);
      
      await context.sync();
      console.log("Event-Handler registriert");
      addLog("🔍 Überwachung aktiv - Warte auf Änderungen in Spalte G...");
    } catch (error) {
      console.error("Fehler beim Starten:", error);
      addLog("❌ Fehler: " + error.message, "error");
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
      
      // Prüfe ob Spalte G betroffen ist
      if (column !== "G") {
        return; // Ignoriere andere Spalten
      }
      
      addLog(`📝 Änderung erkannt: Zeile ${row}`);
      
      // Hole den neuen Wert aus der Zelle
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(event.address);
      range.load("values");
      await context.sync();
      
      const cellValue = range.values[0][0];
      console.log(`Zeile ${row}, Wert: ${cellValue}`);
      
      // Sende Webhook
      await sendWebhook(row, cellValue);
      
    } catch (error) {
      console.error("Fehler beim Verarbeiten:", error);
      addLog("❌ Fehler: " + error.message, "error");
    }
  });
}

// Webhook an Make.com senden
async function sendWebhook(rowNumber, cellContent) {
  const webhookUrl = "https://hook.eu2.make.com/df669kpkbrssyytnmm8turcwy7cql219";
  
  const payload = {
    row: rowNumber,
    value: cellContent,
    timestamp: new Date().toISOString()
  };
  
  addLog(`📤 Sende Webhook: Zeile ${rowNumber}, Wert: "${cellContent}"`);
  
  try {
    const response = await fetch(webhookUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });
    
    if (response.ok) {
      addLog(`✅ Webhook erfolgreich gesendet!`, "success");
    } else {
      addLog(`⚠️ Webhook Fehler: ${response.status}`, "error");
    }
  } catch (error) {
    console.error("Fetch-Fehler:", error);
    addLog(`❌ Netzwerkfehler: ${error.message}`, "error");
  }
}

// Log-Eintrag hinzufügen
function addLog(message, type = "") {
  const logDiv = document.getElementById("log");
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
