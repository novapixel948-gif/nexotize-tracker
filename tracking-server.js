const express = require("express");
const fs = require("fs");
const app = express();
const PORT = process.env.PORT || 3000;
let events = [];
const EVENTS_FILE = "./events.json";
if (fs.existsSync(EVENTS_FILE)) {
  try { events = JSON.parse(fs.readFileSync(EVENTS_FILE, "utf8")); }
  catch (e) { events = []; }
}
function saveEvents() {
  fs.writeFileSync(EVENTS_FILE, JSON.stringify(events, null, 2));
}
const PIXEL = Buffer.from("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7","base64");
app.get("/track/open", (req, res) => {
  const { id } = req.query;
  if (id) {
    events.push({ id, type: "open", timestamp: new Date().toISOString() });
    saveEvents();
  }
  res.set({ "Content-Type": "image/gif", "Cache-Control": "no-store" });
  res.send(PIXEL);
});
app.get("/track/click", (req, res) => {
  const { id, url } = req.query;
  if (id && url) {
    events.push({ id, type: "click", timestamp: new Date().toISOString(), destination: decodeURIComponent(url) });
    saveEvents();
  }
  res.redirect(302, url ? decodeURIComponent(url) : "https://nexotizemedia.com");
});
app.get("/api/events", (req, res) => res.json(events));
app.get("/api/summary", (req, res) => {
  const summary = {};
  events.forEach(e => {
    if (!summary[e.id]) summary[e.id] = { opens: 0, clicks: 0 };
    if (e.type === "open") summary[e.id].opens++;
    if (e.type === "click") summary[e.id].clicks++;
  });
  res.json(summary);
});
app.get("/health", (req, res) => res.json({ status: "ok", totalEvents: events.length }));
app.listen(PORT, () => console.log("Tracker running on port " + PORT));
