"use strict";

const http = require("http");
const fs = require("fs");
const path = require("path");
const os = require("os");
const dns = require("dns");
try { dns.setServers(["8.8.8.8", "8.8.4.4"]); } catch (e) {}
const store = require("./lib/store");

const ROOT = __dirname;
const PORT = Number(process.env.PORT) || 8080;

const MAX_BYTES = 20 * 1024 * 1024;

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".txt": "text/plain; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
};

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

async function saveToDatabase(buf, clientRevision) {
  if (buf.length < 4 || buf[0] !== 0x50 || buf[1] !== 0x4b) {
    return { code: 400, body: { error: "That was not an Excel file." } };
  }
  let merged = false;
  const current = await store.readWorkbook();
  if (current && clientRevision && clientRevision !== current.revision) {
    try {
      buf = store.mergeWorkbooks(current.buffer, buf);
      merged = true;
    } catch (err) {
      console.log("Kept the newest score file because the two copies could not be combined: " + err.message);
    }
  }
  const saved = await store.syncBuffer(buf);
  return {
    code: 200,
    body: { ok: true, revision: saved.revision, merged, database: true, storage: "mongodb", file: "competition" }
  };
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, "http://localhost");

    if (req.method === "GET" && url.pathname === "/api/revision") {
      const meta = await store.readMeta();
      sendJson(res, 200, { revision: meta.revision || "" });
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/info") {
      const port = server.address().port;
      const meta = await store.readMeta();
      sendJson(res, 200, {
        liveFile: "competition",
        usingLive: true,
        database: true,
        storage: "mongodb",
        revision: meta.revision || "",
        rounds: meta.rounds,
        players: meta.players,
        urls: ["http://localhost:" + port, ...lanUrls(port)]
      });
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/registrations") {
      sendJson(res, 200, await store.readRegistrations());
      return;
    }

    if ((req.method === "GET" || req.method === "HEAD") && url.pathname === "/api/workbook") {
      const saved = await store.readWorkbook();
      if (!saved) {
        sendJson(res, 404, { error: "No tournament is stored in the database yet." });
        return;
      }
      res.writeHead(200, {
        "Content-Type": MIME[".xlsx"],
        "Content-Length": saved.buffer.length,
        "Cache-Control": "no-store",
        "X-Workbook-Revision": saved.revision,
        "X-Workbook-Name": "Chess_Tournament_Live.xlsx"
      });
      if (req.method === "HEAD") res.end();
      else res.end(saved.buffer);
      return;
    }

    if (req.method === "PUT" && url.pathname === "/api/workbook") {
      const buf = await readBody(req);
      const result = await enqueue(() => saveToDatabase(buf, req.headers["x-workbook-revision"] || ""));
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
  console.log("Scores are stored in the database. Excel is only for import and export.");
  console.log("If Windows asks whether to allow Node.js on private networks, choose Allow.");
  console.log("Press Ctrl+C to stop hosting.");
  console.log("");
  store.readMeta().then((meta) => {
    console.log("Database is ready. Rounds stored: " + meta.rounds + ".");
  }).catch((err) => console.log("Database read failed: " + err.message));
});
