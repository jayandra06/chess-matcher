"use strict";

const store = require("../lib/store");

function readRaw(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (chunk) => chunks.push(chunk));
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

module.exports = async function handler(req, res) {
  try {
    if (req.method === "GET" || req.method === "HEAD") {
      const saved = await store.readWorkbook();
      if (!saved) {
        res.status(404).json({ error: "No tournament is stored in the database yet." });
        return;
      }
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Length", saved.buffer.length);
      res.setHeader("Cache-Control", "no-store");
      res.setHeader("X-Workbook-Revision", saved.revision);
      res.setHeader("X-Workbook-Name", "Chess_Tournament_Live.xlsx");
      res.status(200);
      if (req.method === "HEAD") res.end();
      else res.end(saved.buffer);
      return;
    }

    if (req.method === "PUT") {
      let buf = await readRaw(req);
      if (buf.length < 4 || buf[0] !== 0x50 || buf[1] !== 0x4b) {
        res.status(400).json({ error: "That was not an Excel file." });
        return;
      }
      const current = await store.readWorkbook();
      const clientRevision = req.headers["x-workbook-revision"] || "";
      let merged = false;
      if (current && clientRevision && clientRevision !== current.revision) {
        try {
          buf = store.mergeWorkbooks(current.buffer, buf);
          merged = true;
        } catch (err) {
          console.log("Kept the newest score file because the two copies could not be combined: " + err.message);
        }
      }
      await store.syncBuffer(buf);
      res.setHeader("Cache-Control", "no-store");
      res.status(200).json({ ok: true, database: true, storage: "mongodb", merged, file: "competition" });
      return;
    }

    res.status(405).json({ error: "Method not allowed" });
  } catch (err) {
    console.log(err.message);
    if (!res.headersSent) res.status(500).json({ error: "Could not reach the database." });
  }
};

module.exports.config = { api: { bodyParser: false } };
