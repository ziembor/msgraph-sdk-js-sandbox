const http = require("http");
const fs = require("fs");
const path = require("path");

const publicDir = path.join(__dirname, "public");
const port = Number(process.env.PORT) || 3000;

const contentTypes = {
  ".html": "text/html; charset=UTF-8",
  ".js": "text/javascript; charset=UTF-8",
  ".css": "text/css; charset=UTF-8",
  ".json": "application/json; charset=UTF-8",
  ".ico": "image/x-icon",
};

function resolvePath(urlPath) {
  const normalized = path.normalize(urlPath).replace(/^(\.\.[/\\])+/, "");
  const filePath = path.join(publicDir, normalized);
  if (fs.existsSync(filePath) && fs.statSync(filePath).isFile()) {
    return filePath;
  }
  return path.join(publicDir, "index.html");
}

http
  .createServer((req, res) => {
    const reqPath = req.url === "/" ? "/index.html" : req.url.split("?")[0];
    const filePath = resolvePath(reqPath);
    const ext = path.extname(filePath).toLowerCase();
    const contentType = contentTypes[ext] || "application/octet-stream";

    fs.readFile(filePath, (error, content) => {
      if (error) {
        res.writeHead(500, { "Content-Type": "text/plain; charset=UTF-8" });
        res.end("Internal Server Error");
        return;
      }
      res.writeHead(200, { "Content-Type": contentType });
      res.end(content);
    });
  })
  .listen(port, () => {
    console.log(`Server running on http://localhost:${port}`);
  });
