// Excel Webhook Monitor - Mit Cloudflare Worker Proxy

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Excel Webhook Monitor bereit");
    addLog("✅ Webhook Monitor gestartet", "success");
    startMonitoring();
  }
});

// ÄNDERE DIESE WERTE:
// 1. Deine Cloudflare Worker URL
const PROXY_URL = "autumn-sea-2657.daniel-steiner-mail.workers.dev
";

// 2. Dein API-Key (muss mit Worker übereinstimmen!)
const API_KEY = "akdsadhoiadoiwoqi8wd";

async function startMonitoring() {
  await Excel.run(async (context) => {
    try {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
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

async function handleCellChange(event) {
  await Excel.run(async (context) => {
    try {
      const match = event.address.match(/([A-Z]+)(\d+)/);
      if (!match) return;
      
      const column = match[1];
      const row = parseInt(match[2]);
      
      if (column !== "G") {
        return;
      }
      
      addLog(`📝 Änderung erkannt: Zeile ${row}`);
      
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(event.address);
      range.load("values");
      await context.sync();
      
      const cellValue = range.values[0][0];
      console.log(`Zeile ${row}, Wert: ${cellValue}`);
      
      await sendWebhook(row, cellValue);
      
    } catch (error) {
      console.error("Fehler beim Verarbeiten:", error);
      addLog("❌ Fehler: " + error.message, "error");
    }
  });
}

async function sendWebhook(rowNumber, cellContent) {
  const payload = {
    row: rowNumber,
    value: cellContent,
    timestamp: new Date().toISOString()
  };
  
  addLog(`📤 Sende via Proxy: Zeile ${rowNumber}`);
  
  try {
    // Sende an Cloudflare Worker (nicht direkt an Make.com!)
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
      addLog(`⚠️ Fehler: ${result.error || result.message}`, "error");
      console.error("Proxy error:", result);
    }
  } catch (error) {
    console.error("Fetch-Fehler:", error);
    addLog(`❌ Netzwerkfehler: ${error.message}`, "error");
  }
}

function addLog(message, type = "") {
  const logDiv = document.getElementById("log");
  if (!logDiv) return;
  
  const entry = document.createElement("div");
  entry.className = "log-entry " + type;
  
  const timestamp = new Date().toLocaleTimeString("de-DE");
  entry.textContent = `[${timestamp}] ${message}`;
  
  logDiv.insertBefore(entry, logDiv.firstChild);
  
  while (logDiv.children.length > 50) {
    logDiv.removeChild(logDiv.lastChild);
  }
}
