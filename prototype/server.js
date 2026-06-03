const http = require("http");
const fs = require("fs/promises");
const path = require("path");
const { fileURLToPath } = require("url");

const PORT = Number(process.env.PORT || 4173);
const ROOT_DIR = __dirname;
const STORAGE_DIR = path.join(ROOT_DIR, "storage");
const STATE_FILE = path.join(STORAGE_DIR, "roadside-status-state.json");

const MIME_TYPES = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
};

function sendJson(response, status, payload) {
  response.writeHead(status, { "Content-Type": "application/json; charset=utf-8" });
  response.end(JSON.stringify(payload));
}

function sendNoContent(response) {
  response.writeHead(204, { "Cache-Control": "no-store" });
  response.end();
}

async function readRequestJson(request) {
  const chunks = [];
  for await (const chunk of request) chunks.push(chunk);
  const text = Buffer.concat(chunks).toString("utf8");
  return text ? JSON.parse(text) : null;
}

function validateStatePayload(payload) {
  return Boolean(
    payload &&
      typeof payload.currentDate === "string" &&
      Array.isArray(payload.currentRows) &&
      Array.isArray(payload.archives),
  );
}

async function handleRoadsideState(request, response) {
  if (request.method === "GET") {
    try {
      const data = await fs.readFile(STATE_FILE, "utf8");
      response.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      response.end(data);
    } catch (error) {
      if (error.code === "ENOENT") return sendJson(response, 404, { error: "not_found" });
      console.error(error);
      return sendJson(response, 500, { error: "read_failed" });
    }
    return;
  }

  if (request.method === "POST") {
    try {
      const payload = await readRequestJson(request);
      if (!validateStatePayload(payload)) return sendJson(response, 400, { error: "invalid_payload" });
      await fs.mkdir(STORAGE_DIR, { recursive: true });
      const saved = {
        currentDate: payload.currentDate,
        currentRows: payload.currentRows,
        archives: payload.archives,
        updatedAt: new Date().toISOString(),
      };
      await fs.writeFile(STATE_FILE, `${JSON.stringify(saved, null, 2)}\n`, "utf8");
      return sendJson(response, 200, { ok: true, updatedAt: saved.updatedAt });
    } catch (error) {
      console.error(error);
      return sendJson(response, 500, { error: "write_failed" });
    }
  }

  response.writeHead(405, { Allow: "GET, POST" });
  response.end();
}

async function serveStatic(request, response, pathname) {
  const requestedPath = pathname === "/" ? "/index.html" : pathname;
  const filePath = path.normalize(path.join(ROOT_DIR, fileURLToPath(`file://${requestedPath}`)));
  if (!filePath.startsWith(ROOT_DIR)) return sendJson(response, 403, { error: "forbidden" });
  try {
    const body = await fs.readFile(filePath);
    response.writeHead(200, { "Content-Type": MIME_TYPES[path.extname(filePath)] || "application/octet-stream" });
    response.end(body);
  } catch (error) {
    if (error.code === "ENOENT") return sendJson(response, 404, { error: "not_found" });
    console.error(error);
    return sendJson(response, 500, { error: "static_read_failed" });
  }
}

const server = http.createServer(async (request, response) => {
  const requestUrl = new URL(request.url, `http://${request.headers.host || `127.0.0.1:${PORT}`}`);
  if (requestUrl.pathname === "/api/health") {
    sendNoContent(response);
    return;
  }
  if (requestUrl.pathname === "/api/roadside-status-state") {
    await handleRoadsideState(request, response);
    return;
  }
  await serveStatic(request, response, decodeURIComponent(requestUrl.pathname));
});

server.listen(PORT, "127.0.0.1", () => {
  console.log(`Prototype server running at http://127.0.0.1:${PORT}`);
  console.log(`Roadside status state file: ${STATE_FILE}`);
});
