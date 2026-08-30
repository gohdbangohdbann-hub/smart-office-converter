const { app, BrowserWindow, shell } = require("electron");
const path = require("node:path");
const http = require("node:http");
const { pathToFileURL } = require("node:url");

const localPort = Number(process.env.NAWA_DESKTOP_PORT || 4317);

function waitForServer(url, timeoutMs = 15000) {
  const started = Date.now();
  return new Promise((resolve, reject) => {
    const probe = () => {
      const request = http.get(url, response => { response.resume(); if (response.statusCode && response.statusCode < 500) return resolve(); setTimeout(probe, 200); });
      request.on("error", () => { if (Date.now() - started > timeoutMs) reject(new Error("Desktop server did not start")); else setTimeout(probe, 200); });
    };
    probe();
  });
}

async function createWindow() {
  const projectRoot = app.isPackaged ? process.resourcesPath : path.resolve(__dirname, "..");
  const serverEntry = path.join(projectRoot, "dist", "index.js");
  process.env.NODE_ENV = "production";
  process.env.PORT = String(localPort);
  await import(pathToFileURL(serverEntry).href);
  await waitForServer(`http://127.0.0.1:${localPort}/`);
  const window = new BrowserWindow({ width: 1180, height: 820, minWidth: 860, minHeight: 620, title: "NumberToWords", webPreferences: { contextIsolation: true, nodeIntegration: false, sandbox: true } });
  window.webContents.setWindowOpenHandler(({ url }) => { void shell.openExternal(url); return { action: "deny" }; });
  await window.loadURL(`http://127.0.0.1:${localPort}/`);
}

app.whenReady().then(() => createWindow().catch(error => { console.error(error); app.quit(); }));
app.on("window-all-closed", () => { if (process.platform !== "darwin") app.quit(); });
