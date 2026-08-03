// Minimal static HTTPS server for local Office Add-in dev — no webpack, no
// bundler, matching the rest of this project's "no build step beyond one
// small script" approach. Office requires HTTPS even for localhost, so this
// reads the trusted dev certificate `npm run dev-certs` installs.
//
// UNVERIFIED: written against office-addin-dev-certs' documented conventional
// output location (~/.office-addin-dev-certs/localhost.{crt,key}). Could not
// confirm the exact path against a real install in this environment — if it
// doesn't find the files, run `npx office-addin-dev-certs install` first and
// check where it actually wrote them.
const fs = require('fs');
const path = require('path');
const os = require('os');
const https = require('https');
const { readFile } = require('fs/promises');

const CERT_DIR = path.join(os.homedir(), '.office-addin-dev-certs');
const CERT_PATH = path.join(CERT_DIR, 'localhost.crt');
const KEY_PATH = path.join(CERT_DIR, 'localhost.key');

if (!fs.existsSync(CERT_PATH) || !fs.existsSync(KEY_PATH)) {
  console.error(`Dev certificate not found at ${CERT_DIR}.`);
  console.error('Run: npx office-addin-dev-certs install');
  process.exit(1);
}

const PORT = 3000;
const ROOT = __dirname;
const MIME = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.png': 'image/png', '.xml': 'application/xml' };

const server = https.createServer(
  { cert: fs.readFileSync(CERT_PATH), key: fs.readFileSync(KEY_PATH) },
  async (req, res) => {
    const urlPath = req.url === '/' ? '/taskpane.html' : req.url.split('?')[0];
    const filePath = path.join(ROOT, decodeURIComponent(urlPath));
    if (!filePath.startsWith(ROOT)) { res.writeHead(403); res.end('Forbidden'); return; }
    try {
      const data = await readFile(filePath);
      const ext = path.extname(filePath);
      res.writeHead(200, { 'Content-Type': MIME[ext] || 'application/octet-stream' });
      res.end(data);
    } catch (err) {
      res.writeHead(404);
      res.end('Not found: ' + urlPath);
    }
  }
);

server.listen(PORT, () => console.log(`Serving ${ROOT} at https://localhost:${PORT}/`));
