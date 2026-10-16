// Static HTTPS server for the add-in. Listens on loopback only.
const https = require("https");
const fs = require("fs");
const path = require("path");

const PORT = Number(process.env.PORT) || 3021;
const ROOT = path.join(__dirname, "public");
const VENDOR = path.join(__dirname, "node_modules", "sql.js", "dist");
const CERTS = path.join(__dirname, "certs");

const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".wasm": "application/wasm",
  ".png": "image/png",
  ".xml": "application/xml",
};

if (!fs.existsSync(path.join(CERTS, "localhost.pfx"))) {
  console.error("Missing certs/localhost.pfx. Run scripts\\install.ps1 first.");
  process.exit(1);
}

const tls = {
  pfx: fs.readFileSync(path.join(CERTS, "localhost.pfx")),
  passphrase: fs.readFileSync(path.join(CERTS, "pfx-pass.txt"), "utf8").trim(),
};

function handler(req, res) {
  let urlPath;
  try {
    urlPath = decodeURIComponent(new URL(req.url, "https://localhost").pathname);
  } catch {
    res.writeHead(400).end();
    return;
  }
  const base = urlPath.startsWith("/vendor/") ? VENDOR : ROOT;
  const rel = urlPath.startsWith("/vendor/") ? urlPath.slice("/vendor/".length) : urlPath === "/" ? "taskpane.html" : urlPath;
  const file = path.resolve(base, "." + path.sep + rel);
  if (!file.startsWith(base + path.sep)) {
    res.writeHead(403).end();
    return;
  }
  fs.readFile(file, (err, body) => {
    if (err) {
      res.writeHead(404, { "Content-Type": "text/plain" }).end("Not found");
      return;
    }
    res.writeHead(200, {
      "Content-Type": TYPES[path.extname(file).toLowerCase()] || "application/octet-stream",
      "Cache-Control": "no-cache",
    });
    res.end(body);
  });
}

for (const host of ["127.0.0.1", "::1"]) {
  https
    .createServer(tls, handler)
    .on("error", (e) => {
      if (e.code === "EADDRINUSE") console.error(`Port ${PORT} on ${host} is already in use (server already running?).`);
      else if (e.code !== "EADDRNOTAVAIL") console.error(e);
    })
    .listen(PORT, host, () => console.log(`Excel SQL add-in served at https://localhost:${PORT} (${host})`));
}
