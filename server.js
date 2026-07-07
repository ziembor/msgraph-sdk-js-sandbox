const http = require("http");
const fs = require("fs");
const path = require("path");

const publicDir = path.join(__dirname, "public");
const indexPath = path.join(publicDir, "index.html");
const port = Number(process.env.PORT) || 3000;

const contentTypes = {
  ".html": "text/html; charset=UTF-8",
  ".js": "text/javascript; charset=UTF-8",
  ".css": "text/css; charset=UTF-8",
  ".json": "application/json; charset=UTF-8",
  ".ico": "image/x-icon",
};

function resolvePath(urlPath) {
  const relativePath = urlPath === "/" ? "/index.html" : urlPath;
  const safePath = path.resolve(publicDir, `.${relativePath}`);
  if (safePath === publicDir || safePath.startsWith(`${publicDir}${path.sep}`)) {
    return safePath;
  }
  return null;
}

function sendResponse(res, statusCode, contentType, content) {
  res.writeHead(statusCode, { "Content-Type": contentType });
  res.end(content);
}

function readAndSendFile(res, filePath, reqPath) {
  fs.readFile(filePath, (error, content) => {
    if (error) {
      if (error.code === "ENOENT") {
        const isAssetRequest = path.extname(reqPath) !== "";
        if (!isAssetRequest && filePath !== indexPath) {
          readAndSendFile(res, indexPath, "/index.html");
          return;
        }
        sendResponse(res, 404, "text/plain; charset=UTF-8", "Not Found");
        return;
      }
      sendResponse(res, 500, "text/plain; charset=UTF-8", "Internal Server Error");
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = contentTypes[ext] || "application/octet-stream";
    sendResponse(res, 200, contentType, content);
  });
}

http
  .createServer((req, res) => {
    const reqPath = (req.url || "/").split("?")[0];
    const filePath = resolvePath(reqPath);
    if (!filePath) {
      sendResponse(res, 403, "text/plain; charset=UTF-8", "Forbidden");
      return;
    }
    readAndSendFile(res, filePath, reqPath);
  })
  .listen(port, () => {
    console.log(`Server running on http://localhost:${port}`);
  });
