"use strict";

const fs = require("fs");
const path = require("path");
const { MongoClient, Binary } = require("mongodb");
const XLSX = require(path.join(__dirname, "..", "xlsx.full.min.js"));

const DOC_ID = "live-tournament";

function loadEnv() {
  if (process.env.MONGODB_URI) return;
  const file = path.join(__dirname, "..", ".env");
  let text = "";
  try {
    text = fs.readFileSync(file, "utf8");
  } catch (err) {
    return;
  }
  for (const line of text.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const match = trimmed.match(/^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$/);
    if (!match || process.env[match[1]]) continue;
    let value = match[2].trim();
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    }
    process.env[match[1]] = value;
  }
}

let client = null;
async function collection() {
  loadEnv();
  const uri = process.env.MONGODB_URI;
  if (!uri) throw new Error("MONGODB_URI is not set. Add it to .env for this project, and to the Vercel project settings when you deploy.");
  if (!client) client = new MongoClient(uri, { serverSelectionTimeoutMS: 8000, family: 4 });
  await client.connect();
  return client.db("test").collection("competition");
}

function sheetRows(wb, name) {
  const sheet = wb.Sheets[name];
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, { defval: "" });
}

function roundRows(wb, round) {
  return sheetRows(wb, "Round " + round);
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

function toBuffer(value) {
  if (!value) return null;
  if (Buffer.isBuffer(value)) return value;
  if (value instanceof Uint8Array) return Buffer.from(value);
  if (typeof value.value === "function") return Buffer.from(value.value(true));
  if (value.buffer) return Buffer.from(value.buffer);
  return null;
}

async function syncBuffer(buf, registrations) {
  const bytes = Buffer.from(buf);
  const wb = XLSX.read(bytes, { type: "buffer" });
  const rounds = {};
  for (const name of wb.SheetNames) {
    const found = /^Round (\d+)$/.exec(name);
    if (found) rounds[found[1]] = sheetRows(wb, name);
  }
  const col = await collection();
  const existing = await col.findOne({ _id: DOC_ID }, { projection: { registrations: 1 } });
  const revision = String(Date.now());
  const doc = {
    _id: DOC_ID,
    kind: "tournament",
    competition: "Chess Championship 2026",
    updatedAt: new Date(),
    revision,
    players: sheetRows(wb, "Players"),
    rounds,
    matches: sheetRows(wb, "All Matches"),
    standings: sheetRows(wb, "Standings"),
    settings: sheetRows(wb, "Settings"),
    registrations: Array.isArray(registrations) ? registrations : (existing && existing.registrations) || [],
    workbook: new Binary(bytes)
  };
  await col.replaceOne({ _id: DOC_ID }, doc, { upsert: true });
  return { revision, rounds: Object.keys(rounds).length, players: doc.players.length, matches: doc.matches.length };
}

async function syncFile(file, registrations) {
  return syncBuffer(fs.readFileSync(file), registrations);
}

async function readWorkbook() {
  const col = await collection();
  const doc = await col.findOne({ _id: DOC_ID }, { projection: { workbook: 1, revision: 1, updatedAt: 1 } });
  if (!doc) return null;
  const buffer = toBuffer(doc.workbook);
  if (!buffer || buffer.length < 4) return null;
  return {
    buffer,
    revision: doc.revision || String(new Date(doc.updatedAt).getTime())
  };
}

async function readMeta() {
  const col = await collection();
  const doc = await col.findOne({ _id: DOC_ID }, { projection: { revision: 1, updatedAt: 1, rounds: 1, players: 1 } });
  if (!doc) return { revision: "", rounds: 0, players: 0 };
  return {
    revision: doc.revision || String(new Date(doc.updatedAt).getTime()),
    rounds: doc.rounds ? Object.keys(doc.rounds).length : 0,
    players: Array.isArray(doc.players) ? doc.players.length : 0
  };
}

async function readRegistrations() {
  const col = await collection();
  const doc = await col.findOne({ _id: DOC_ID }, { projection: { registrations: 1 } });
  return (doc && doc.registrations) || [];
}

module.exports = {
  loadEnv,
  mergeWorkbooks,
  syncBuffer,
  syncFile,
  readWorkbook,
  readMeta,
  readRegistrations
};
