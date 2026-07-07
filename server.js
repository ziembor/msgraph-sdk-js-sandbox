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

function indexPublicFiles(dir, relativeDir = "", fileMap = {}) {
  const entries = fs.readdirSync(dir, { withFileTypes: true });
  for (const entry of entries) {
    const absolutePath = path.join(dir, entry.name);
    const relativePath = path.posix.join(relativeDir, entry.name);
    if (entry.isDirectory()) {
      indexPublicFiles(absolutePath, relativePath, fileMap);
    } else {
      fileMap[`/${relativePath}`] = absolutePath;
    }
  }
  return fileMap;
}

const publicFiles = indexPublicFiles(publicDir);
publicFiles["/"] = indexPath;

function normalizeRequestPath(rawUrl) {
  const pathOnly = (rawUrl || "/").split("?")[0];
  let decodedPath;
  try {
    decodedPath = decodeURIComponent(pathOnly);
  } catch (error) {
    return null;
  }

  const normalizedPath = path.posix.normalize(decodedPath.replace(/\\/g, "/"));
  if (normalizedPath.includes("\0")) {
    return null;
  }

  let absolutePath = normalizedPath.startsWith("/") ? normalizedPath : `/${normalizedPath}`;
  if (absolutePath === "/.") {
    absolutePath = "/";
  }
  if (absolutePath.split("/").includes("..")) {
    return null;
  }

  return absolutePath;
}

function sendResponse(res, statusCode, contentType, content) {
  res.writeHead(statusCode, { "Content-Type": contentType });
  res.end(content);
}

function readAndSendFile(res, filePath) {
  fs.readFile(filePath, (error, content) => {
    if (error) {
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
    const reqPath = normalizeRequestPath(req.url);
    if (!reqPath) {
      sendResponse(res, 403, "text/plain; charset=UTF-8", "Forbidden");
      return;
    }

    const filePath =
      publicFiles[reqPath] || (path.extname(reqPath) === "" ? indexPath : null);
    if (!filePath) {
      sendResponse(res, 404, "text/plain; charset=UTF-8", "Not Found");
      return;
    }

    readAndSendFile(res, filePath);
  })
  .listen(port, () => {
    console.log(`Server running on http://localhost:${port}`);
  });
