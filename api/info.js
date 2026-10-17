"use strict";

const store = require("../lib/store");

module.exports = async function handler(req, res) {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }
  try {
    const meta = await store.readMeta();
    res.setHeader("Cache-Control", "no-store");
    res.status(200).json({
      liveFile: "competition",
      usingLive: true,
      database: true,
      storage: "mongodb",
      revision: meta.revision || "",
      rounds: meta.rounds,
      players: meta.players,
      urls: []
    });
  } catch (err) {
    res.status(500).json({ error: "Could not reach the database." });
  }
};
