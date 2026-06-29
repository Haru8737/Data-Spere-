/**
 * Fleet Performance Dashboard — Express Server
 * Serves the static client-side application.
 * All processing (Excel parsing, charts, PDF) runs in the browser.
 */

const express = require('express');
const path    = require('path');

const app  = express();
const PORT = process.env.PORT || 3000;

// Serve everything in /public as static files
app.use(express.static(path.join(__dirname, 'public')));

// Catch-all: always return index.html for any route (SPA)
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
  console.log(`Fleet Dashboard running → http://localhost:${PORT}`);
});
