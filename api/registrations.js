"use strict";

const store = require("../lib/store");

module.exports = async function handler(req, res) {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }
  try {
    const rows = await store.readRegistrations();
    res.setHeader("Cache-Control", "no-store");
    res.status(200).json(rows);
  } catch (err) {
    res.status(500).json({ error: "Could not reach the database." });
  }
};
