"use strict";

const http = require("http");
const fs = require("fs");
const path = require("path");
const os = require("os");
const { execFileSync } = require("child_process");
const dns = require("dns");
const { MongoClient } = require("mongodb");
try { dns.setServers(["8.8.8.8", "8.8.4.4"]); } catch (e) {}

const ROOT = __dirname;
const PORT = Number(process.env.PORT) || 8080;
const LIVE = path.join(ROOT, "Chess_Tournament_Live.xlsx");
const SOURCE = path.join(ROOT, "Chess_Tournament_Tracker.xlsx");
const XLSX = require("./xlsx.full.min.js");

function registrationPath() {
  const desktop = path.join(process.env.USERPROFILE || "", "OneDrive", "Desktop", "Asia_By_Gram_Form_Registrations_2026-10-17 (2).xlsx");
  if (fs.existsSync(desktop)) return desktop;
  const local = path.join(ROOT, "Registrations.xlsx");
  if (fs.existsSync(local)) return local;
  return "";
}
const MAX_BYTES = 20 * 1024 * 1024;

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".txt": "text/plain; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
};

function workbookPath() {
  return fs.existsSync(LIVE) ? LIVE : SOURCE;
}

function revisionOf(file) {
  try {
    return String(fs.statSync(file).mtimeMs);
  } catch {
    return "";
  }
}

function lanUrls(port) {
  const urls = [];
  const nets = os.networkInterfaces();
  for (const list of Object.values(nets)) {
    for (const net of list || []) {
      const v4 = net.family === "IPv4" || net.family === 4;
      if (v4 && !net.internal) urls.push("http://" + net.address + ":" + port);
    }
  }
  return urls;
}

function sendJson(res, code, obj) {
  const body = Buffer.from(JSON.stringify(obj));
  res.writeHead(code, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": body.length,
    "Cache-Control": "no-store"
  });
  res.end(body);
}

function safeFile(urlPath) {
  let decoded;
  try {
    decoded = decodeURIComponent(urlPath.split("?")[0]);
  } catch {
    return null;
  }
  const rel = decoded === "/" ? "index.html" : decoded.replace(/^\/+/, "");
  if (!rel || rel.includes("\0")) return null;
  const full = path.resolve(ROOT, rel);
  if (full !== ROOT && !full.startsWith(ROOT + path.sep)) return null;
  return full;
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > MAX_BYTES) {
        reject(Object.assign(new Error("too large"), { status: 413 }));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

let chain = Promise.resolve();
function enqueue(fn) {
  const run = chain.then(fn, fn);
  chain = run.then(() => {}, () => {});
  return run;
}

function isLocked(err) {
  return err && (err.code === "EBUSY" || err.code === "EPERM" || err.code === "EACCES");
}

function runPowerShell(script, env) {
  execFileSync("powershell.exe", ["-NoProfile", "-NonInteractive", "-Command", script], {
    timeout: 25000,
    windowsHide: true,
    env: Object.assign({}, process.env, env || {})
  });
}

function closeLiveWorkbook() {
  try {
    runPowerShell(`
$ErrorActionPreference = 'Stop'
try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }
catch { exit 3 }
$found = $false
foreach ($wb in @($excel.Workbooks)) {
  if ($wb.Name -eq 'Chess_Tournament_Live.xlsx') { $wb.Close($false); $found = $true }
}
if (-not $found) { exit 2 }
`);
    return true;
  } catch (err) {
    return false;
  }
}

function openLiveWorkbook() {
  runPowerShell(`
$ErrorActionPreference = 'Stop'
try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }
catch { $excel = New-Object -ComObject Excel.Application }
$excel.Visible = $true
$excel.DisplayAlerts = $false
$opened = $excel.Workbooks.Open($env:LIVE_XLSX)
$opened.Activate()
$excel.Visible = $true
`, { LIVE_XLSX: LIVE });
}

function copyOverLive(tmp) {
  let lastError;
  for (let attempt = 0; attempt < 6; attempt++) {
    try {
      fs.copyFileSync(tmp, LIVE);
      return;
    } catch (err) {
      lastError = err;
      if (!isLocked(err)) throw err;
      if (attempt === 0) closeLiveWorkbook();
      Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 400);
    }
  }
  throw lastError;
}

function roundRows(wb, round) {
  const sheet = wb.Sheets["Round " + round];
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, { defval: "" });
}

function filled(value) {
  return String(value == null ? "" : value).trim();
}

function mergeWorkbooks(currentBuf, incomingBuf) {
  const current = XLSX.read(currentBuf, { type: "buffer" });
  const incoming = XLSX.read(incomingBuf, { type: "buffer" });
  const resultKey = "Result (Winner / Draw)";
  const roundNumbers = new Set();
  for (const name of current.SheetNames.concat(incoming.SheetNames)) {
    const match = /^Round (\d+)$/.exec(name);
    if (match) roundNumbers.add(Number(match[1]));
  }
  for (const round of [...roundNumbers].sort((a, b) => a - b)) {
    const byTable = new Map();
    for (const row of roundRows(current, round)) byTable.set(Number(row.Table), row);
    for (const row of roundRows(incoming, round)) {
      const table = Number(row.Table);
      const previous = byTable.get(table);
      if (!previous) {
        byTable.set(table, row);
        continue;
      }
      const merged = Object.assign({}, previous, row);
      if (filled(previous[resultKey]) && !filled(row[resultKey])) merged[resultKey] = previous[resultKey];
      for (const key of ["Player 1", "Player 2"]) {
        if (filled(previous[key]) && !filled(row[key])) merged[key] = previous[key];
      }
      byTable.set(table, merged);
    }
    const rows = [...byTable.values()].filter((row) => [6, 7, 8, 9, 10].includes(Number(row.Table)));
    rows.sort((a, b) => Number(a.Table) - Number(b.Table));
    if (rows.length) incoming.Sheets["Round " + round] = XLSX.utils.json_to_sheet(rows);
  }
  return XLSX.write(incoming, { bookType: "xlsx", type: "buffer" });
}

function saveBuffer(buf, clientRevision) {
  if (buf.length < 4 || buf[0] !== 0x50 || buf[1] !== 0x4b) {
    return { code: 400, body: { error: "That was not an Excel file." } };
  }
  let merged = false;
  if (fs.existsSync(LIVE)) {
    const currentRevision = revisionOf(LIVE);
    if (clientRevision && clientRevision !== currentRevision) {
      try {
        buf = mergeWorkbooks(fs.readFileSync(LIVE), buf);
        merged = true;
      } catch (err) {
        console.log("Kept the newest score file because the two copies could not be combined: " + err.message);
      }
    }
  }
  const tmp = path.join(ROOT, "Chess_Tournament_Live.incoming.xlsx");
  let excelWasOpen = false;
  try {
    const fd = fs.openSync(LIVE, "r+");
    fs.closeSync(fd);
  } catch (err) {
    if (isLocked(err)) {
      closeLiveWorkbook();
      excelWasOpen = true;
    }
  }
  try {
    fs.writeFileSync(tmp, buf);
    copyOverLive(tmp);
    if (excelWasOpen) {
      try { openLiveWorkbook(); } catch (err) {
        console.log("Saved the workbook, but Excel did not reopen it: " + err.message);
      }
    }
  } catch (err) {
    return {
      code: isLocked(err) ? 423 : 500,
      body: {
        error: isLocked(err)
          ? "Excel has Chess_Tournament_Live.xlsx open. Close it in Excel and the page will save again."
          : "Could not save the workbook on this PC."
      }
    };
  } finally {
    try { fs.rmSync(tmp, { force: true }); } catch {}
  }
  return {
    code: 200,
    body: { ok: true, revision: revisionOf(LIVE), merged, file: "Chess_Tournament_Live.xlsx" }
  };
}

function mongoUri() {
  const file = path.join(process.env.USERPROFILE || "", "OneDrive", "Desktop", "asia by gram", "ASIA_BY_GRAM", ".env");
  try {
    const text = fs.readFileSync(file, "utf8");
    const found = text.match(/^MONGODB_URI=(.*)$/m);
    return found ? found[1].trim() : "";
  } catch (err) {
    return "";
  }
}

let mongoClient = null;
async function competitionCollection() {
  const uri = mongoUri();
  if (!uri) throw new Error("The database address is not set.");
  if (!mongoClient) mongoClient = new MongoClient(uri, { serverSelectionTimeoutMS: 8000, family: 4 });
  await mongoClient.connect();
  return mongoClient.db("test").collection("competition");
}

function sheetRows(wb, name) {
  const sheet = wb.Sheets[name];
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, { defval: "" });
}

async function syncCompetition() {
  const wb = XLSX.read(fs.readFileSync(LIVE), { type: "buffer" });
  const rounds = {};
  for (const name of wb.SheetNames) {
    const found = /^Round (\d+)$/.exec(name);
    if (found) rounds[found[1]] = sheetRows(wb, name);
  }
  const doc = {
    kind: "tournament",
    competition: "Chess Championship 2026",
    updatedAt: new Date(),
    players: sheetRows(wb, "Players"),
    rounds,
    matches: sheetRows(wb, "All Matches"),
    standings: sheetRows(wb, "Standings"),
    settings: sheetRows(wb, "Settings")
  };
  const collection = await competitionCollection();
  await collection.replaceOne({ _id: "live-tournament" }, doc, { upsert: true });
  console.log("Saved the live tournament into the competition collection.");
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, "http://localhost");

    if (req.method === "GET" && url.pathname === "/api/revision") {
      sendJson(res, 200, { revision: revisionOf(workbookPath()) });
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/info") {
      const port = server.address().port;
      sendJson(res, 200, {
        liveFile: "Chess_Tournament_Live.xlsx",
        usingLive: fs.existsSync(LIVE),
        revision: revisionOf(workbookPath()),
        urls: ["http://localhost:" + port, ...lanUrls(port)]
      });
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/registrations") {
      const file = registrationPath();
      if (!file) {
        sendJson(res, 404, { error: "Registration file not found." });
        return;
      }
      const wb = XLSX.read(fs.readFileSync(file), { type: "buffer" });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
      sendJson(res, 200, rows);
      return;
    }

    if ((req.method === "GET" || req.method === "HEAD") && url.pathname === "/api/workbook") {
      const file = workbookPath();
      if (!fs.existsSync(file)) {
        sendJson(res, 404, { error: "No workbook on this PC yet." });
        return;
      }
      const data = fs.readFileSync(file);
      res.writeHead(200, {
        "Content-Type": MIME[".xlsx"],
        "Content-Length": data.length,
        "Cache-Control": "no-store",
        "X-Workbook-Revision": revisionOf(file),
        "X-Workbook-Name": path.basename(file)
      });
      if (req.method === "HEAD") res.end();
      else res.end(data);
      return;
    }

    if (req.method === "PUT" && url.pathname === "/api/workbook") {
      const buf = await readBody(req);
      const result = await enqueue(async () => {
        const saved = saveBuffer(buf, req.headers["x-workbook-revision"] || "");
        if (saved.code === 200) {
          try {
            await syncCompetition();
            saved.body.database = true;
          } catch (err) {
            saved.body.database = false;
            console.log("Database save failed: " + err.message);
          }
        }
        return saved;
      });
      console.log("Save from " + (req.socket.remoteAddress || "a device") + " -> " + result.code);
      sendJson(res, result.code, result.body);
      return;
    }

    if (req.method !== "GET" && req.method !== "HEAD") {
      sendJson(res, 405, { error: "Method not allowed" });
      return;
    }

    const file = safeFile(url.pathname);
    if (!file || !fs.existsSync(file) || !fs.statSync(file).isFile()) {
      sendJson(res, 404, { error: "Not found" });
      return;
    }
    const data = fs.readFileSync(file);
    const ext = path.extname(file).toLowerCase();
    res.writeHead(200, {
      "Content-Type": MIME[ext] || "application/octet-stream",
      "Content-Length": data.length,
      "Cache-Control": ext === ".html" || ext === ".js" ? "no-store" : "public, max-age=3600"
    });
    if (req.method === "HEAD") res.end();
    else res.end(data);
  } catch (err) {
    if (!res.headersSent) sendJson(res, err.status || 500, { error: "Server error" });
  }
});

server.on("error", (err) => {
  if (err.code === "EADDRINUSE") {
    console.error("Port " + PORT + " is already in use. Close the other host window, or start again with a different port.");
  } else {
    console.error(err);
  }
  process.exit(1);
});

server.listen(PORT, "0.0.0.0", () => {
  const port = server.address().port;
  const others = lanUrls(port);
  console.log("");
  console.log("Chess Tournament is hosted on this PC.");
  console.log("On this computer:  http://localhost:" + port);
  if (others.length) {
    console.log("On other devices connected to the same Wi-Fi, open:");
    for (const url of others) console.log("  " + url);
  } else {
    console.log("No network address found. Connect this PC to Wi-Fi, then start the host again.");
  }
  console.log("");
  console.log("Scores are written on this PC to: Chess_Tournament_Live.xlsx");
  console.log("The original Chess_Tournament_Tracker.xlsx is left unchanged.");
  console.log("Close Chess_Tournament_Live.xlsx in Excel while people are entering scores.");
  console.log("If Windows asks whether to allow Node.js on private networks, choose Allow.");
  console.log("Press Ctrl+C to stop hosting.");
  console.log("");
  syncCompetition().catch((err) => console.log("Database save failed: " + err.message));
});
