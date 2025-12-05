console.log("🚀 taskpane.js Version: 19:45 Uhr - Bereinigt ohne Duplikate");

// KONFIGURATION - HIER DEINE URLs EINTRAGEN
const PROXY_URL = "https://autumn-sea-2657.daniel-steiner-mail.workers.dev";
const API_KEY = "akdsadhoiadoiwoqi8wd";


// GLOBALE VARIABLEN - NUR EINMAL DEKLARIERT
let isMonitoringActive = false;
let eventHandlerContext = null;
let handlerRegistered = false;


// ===== INITIALISIERUNG =====
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("✅ Excel Webhook Monitor geladen");
    addLog("✅ Excel Webhook Monitor geladen");
    
    initializeMonitor();
  }
});

async function initializeMonitor() {
  console.log("🔍 Prüfe Dateinamen...");
  
  try {
    await Excel.run(async (context) => {
      const file = context.workbook.properties;
      file.load("name");
      await context.sync();
      
      const fileName = file.name;
      console.log("📄 Datei:", fileName);
      addLog(`📄 Datei: ${fileName}`);
      
      // Prüfe ob Monitoring vorher aktiv war
      const wasActive = localStorage.getItem('monitoringActive') === 'true';
      const lastFileName = localStorage.getItem('lastFileName');
      
      console.log("🔍 War aktiv?", wasActive);
      console.log("🔍 Letzte Datei:", lastFileName);
      console.log("🔍 Aktuelle Datei:", fileName);
      
      if (wasActive && lastFileName === fileName) {
        console.log("🔄 Auto-Start: Monitoring war vorher aktiv für diese Datei");
        addLog("🔄 Starte Monitoring automatisch...", "info");
        
        // Automatisch starten nach kurzer Verzögerung
        setTimeout(() => {
          startMonitoring();
        }, 1000);
      } else {
        console.log("⚪ Monitoring muss manuell gestartet werden");
        updateStatusUI(false);
        addLog("⚪ Klicke unten auf 'START' um Monitoring zu aktivieren", "info");
      }
      
      // Speichere aktuelle Datei
      localStorage.setItem('lastFileName', fileName);
    });
  } catch (error) {
    console.error("❌ Fehler in initializeMonitor:", error);
    addLog("❌ Fehler: " + error.message, "error");
  }
  
  // Initialisiere UI
  initializeUI();
}

function initializeUI() {
  console.log("🎨 Initialisiere UI...");
  
  const button = document.getElementById("toggleButton");
  
  if (button) {
    console.log("✅ Button gefunden, füge Event-Listener hinzu");
    button.addEventListener("click", toggleMonitoring);
    addLog("🎨 UI initialisiert", "info");
  } else {
    console.error("❌ toggleButton nicht gefunden!");
    addLog("❌ Fehler: Button nicht gefunden", "error");
  }
}

// ===== MONITORING STEUERUNG =====
async function toggleMonitoring() {
  console.log("=== toggleMonitoring aufgerufen ===");
  
  const button = document.getElementById("toggleButton");
  
  if (button) {
    button.disabled = true;
    button.style.opacity = "0.6";
    button.style.cursor = "wait";
  }
  
  try {
    if (isMonitoringActive) {
      await stopMonitoring();
    } else {
      await startMonitoring();
    }
  } finally {
    if (button) {
      button.disabled = false;
      button.style.opacity = "1";
      button.style.cursor = "pointer";
    }
  }
}


async function startMonitoring() {
  console.log("=== startMonitoring aufgerufen ===");
  console.log("  isMonitoringActive:", isMonitoringActive);
  console.log("  handlerRegistered:", handlerRegistered);
  
  if (isMonitoringActive) {
    console.log("⚠️⚠️⚠️ ABBRUCH: Monitoring läuft bereits!");
    addLog("⚠️ Monitoring läuft bereits", "info");
    return;
  }
  
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // NUR registrieren wenn noch nicht registriert
      if (!handlerRegistered) {
        console.log("📝 Registriere Handler zum ersten Mal...");
        eventHandlerContext = sheet.onChanged.add(handleCellChange);
        await context.sync();
        handlerRegistered = true;
        console.log("✅ Handler erfolgreich registriert");
      } else {
        console.log("ℹ️ Handler bereits registriert, überspringe Registrierung");
      }
      
      isMonitoringActive = true;
      localStorage.setItem('monitoringActive', 'true');
      
      updateStatusUI(true);
      addLog("✅ Bereit - Monitoring läuft im Hintergrund!", "success");
      addLog("💡 Du kannst dieses Panel schließen", "info");
      console.log("🔍 Überwache Spalte G...");
    });
  } catch (error) {
    console.error("❌ Fehler in startMonitoring:", error);
    addLog("❌ Fehler beim Starten: " + error.message, "error");
    isMonitoringActive = false;
  }
  
  console.log("=== startMonitoring beendet ===");
}

async function stopMonitoring() {
  console.log("=== stopMonitoring aufgerufen ===");
  
  if (!isMonitoringActive) {
    console.log("⚠️ Monitoring ist bereits gestoppt");
    return;
  }
  
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      if (handlerRegistered && eventHandlerContext) {
        console.log("🗑️ Entferne Event-Handler...");
        eventHandlerContext.remove();
        await context.sync();
        
        handlerRegistered = false;
        eventHandlerContext = null;
        console.log("✅ Event-Handler entfernt");
      }
    });
  } catch (error) {
    console.error("❌ Fehler beim Entfernen:", error);
  }
  
  isMonitoringActive = false;
  localStorage.setItem('monitoringActive', 'false');
  
  updateStatusUI(false);
  addLog("⏸️ Monitoring gestoppt", "info");
}







// ===== UI UPDATE =====
function updateStatusUI(isActive) {
  const indicator = document.getElementById("statusIndicator");
  const statusBar = document.getElementById("statusBar");
  const button = document.getElementById("toggleButton");
  
  if (isActive) {
    // GRÜN - AKTIV
    if (indicator) {
      indicator.classList.remove("status-inactive");
      indicator.classList.add("status-active");
    }
    
    if (statusBar) {
      statusBar.style.backgroundColor = "#4CAF50";
      statusBar.innerHTML = "🟢 MONITORING AKTIV";
    }
    
    if (button) {
      button.textContent = "⏸️ MONITORING STOPPEN";
      button.style.background = "linear-gradient(135deg, #f093fb 0%, #f5576c 100%)";
    }
  } else {
    // ROT - INAKTIV
    if (indicator) {
      indicator.classList.remove("status-active");
      indicator.classList.add("status-inactive");
    }
    
    if (statusBar) {
      statusBar.style.backgroundColor = "#f44336";
      statusBar.innerHTML = "🔴 MONITORING INAKTIV";
    }
    
    if (button) {
      button.textContent = "🚀 MONITORING STARTEN";
      button.style.background = "linear-gradient(135deg, #667eea 0%, #764ba2 100%)";
    }
  }
}

// ===== EVENT HANDLER =====
let lastProcessedCell = null;
let lastProcessedTime = 0;

async function handleCellChange(event) {
  console.log("🔔 handleCellChange aufgerufen:", event);
  
  if (!isMonitoringActive) {
    console.log("⚠️ Monitoring ist inaktiv, ignoriere Event");
    return;
  }
  
  // DEBOUNCE: Ignoriere gleiche Zelle innerhalb von 1 Sekunde
  const currentTime = Date.now();
  const cellKey = event.address;
  
  if (cellKey === lastProcessedCell && (currentTime - lastProcessedTime) < 1000) {
    console.log("⏭️ Überspringe doppeltes Event für", cellKey);
    return;
  }
  
  lastProcessedCell = cellKey;
  lastProcessedTime = currentTime;

  // ===== NEU: Prüfe die Quelle der Änderung =====
  console.log("🔍 Event Source:", event.source);
  console.log("🔍 Event Change Type:", event.changeType);
  
  // Wenn die Änderung NICHT vom User kommt, ignorieren
  if (event.source !== Excel.EventSource.local) {
    console.log(`⏭️ Ignoriere Änderung - Quelle ist: ${event.source} (nicht local)`);
    addLog(`⏭️ API-Änderung ignoriert (Source: ${event.source})`);
    return;
  }
  
  console.log("✅ Änderung vom User bestätigt!");
  // ===== Ende der Source-Prüfung =====
  

  
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
      
      if (column !== "H") {
        console.log(`⏭️ Ignoriere Spalte ${column}`);
        return;
      }
      
      console.log("✅ Spalte H betroffen!");
      addLog(`📝 Änderung in Spalte H: Zeile ${row}`);
      
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Hole Spalten A bis P (16 Spalten)
      const rowRange = sheet.getRange(`A${row}:P${row}`);
      rowRange.load("values");
      await context.sync();
      
      if (!rowRange.values || !rowRange.values[0]) {
        console.error("❌ Keine Daten gefunden");
        addLog("❌ Fehler: Zeile enthält keine Daten", "error");
        return;
      }
      
      const rowData = rowRange.values[0];
      
      console.log(`📊 Gesamte Zeile ${row} (A-P):`, rowData);
      console.log(`📊 Anzahl Spalten: ${rowData.length}`);
      
      await sendWebhook(row, rowData);
    });
  } catch (error) {
    console.error("❌ Fehler in handleCellChange:", error);
    addLog("❌ Fehler: " + error.message, "error");
  }
}
// ===== WEBHOOK SENDEN =====
async function sendWebhook(rowNumber, rowData) {
  console.log("📤 Sende Webhook...");
  
  if (!Array.isArray(rowData) || rowData.length === 0) {
    console.error("❌ rowData ist ungültig:", rowData);
    addLog("❌ Fehler: Keine Daten in der Zeile", "error");
    return;
  }
  
  const payload = {
    row: rowNumber,
    value: rowData[6],  // Spalte G (Index 6)
    data: rowData,      // Komplette Zeile A-P
    timestamp: new Date().toISOString()
  };
  
  console.log("📦 Payload:", JSON.stringify(payload, null, 2));
  addLog(`📤 Sende Webhook: Zeile ${rowNumber} mit ${rowData.length} Spalten`);
  
  if (!PROXY_URL || PROXY_URL.includes("DEIN") || PROXY_URL.includes("dein")) {
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

// ===== LOGGING =====
function addLog(message, type = "") {
  console.log(`[LOG ${type}]`, message);
  
  const logDiv = document.getElementById("log");
  if (!logDiv) {
    console.warn("⚠️ Log-Div nicht gefunden!");
    return;
  }
  
  // Entferne Platzhalter beim ersten echten Log
  const placeholder = logDiv.querySelector('.log-entry');
  if (placeholder && placeholder.textContent.includes('Warte auf Initialisierung')) {
    logDiv.innerHTML = '';
  }
  
  const entry = document.createElement("div");
  entry.className = "log-entry " + type;
  
  const timestamp = new Date().toLocaleTimeString("de-DE");
  entry.textContent = `[${timestamp}] ${message}`;
  
  logDiv.insertBefore(entry, logDiv.firstChild);
  
  while (logDiv.children.length > 50) {
    logDiv.removeChild(logDiv.lastChild);
  }
  
  logDiv.scrollTop = 0;
}

console.log("🔥 taskpane.js vollständig geladen");
console.log("💡 Öffne die Console (F12) für detaillierte Logs");
