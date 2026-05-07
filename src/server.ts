import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import { readFile } from "node:fs/promises";
import { extname, join, normalize } from "node:path";
import { evaluateDiagnostics } from "./rules.js";
import { renderMarkdownReport } from "./report.js";
import type { OutlookDiagnostics } from "./types.js";

const port = Number(process.env.PORT ?? 3000);
const publicDir = join(process.cwd(), "web");
const samplesDir = join(process.cwd(), "samples");

const mimeTypes = new Map<string, string>([
  [".html", "text/html; charset=utf-8"],
  [".css", "text/css; charset=utf-8"],
  [".js", "text/javascript; charset=utf-8"],
  [".json", "application/json; charset=utf-8"],
  [".svg", "image/svg+xml"]
]);

const server = createServer(async (request, response) => {
  try {
    if (!request.url) {
      sendText(response, 400, "Bad request");
      return;
    }

    const url = new URL(request.url, `http://${request.headers.host ?? "localhost"}`);

    if (request.method === "POST" && url.pathname === "/api/diagnose") {
      await handleDiagnose(request, response);
      return;
    }

    if (request.method !== "GET") {
      sendText(response, 405, "Method not allowed");
      return;
    }

    if (url.pathname.startsWith("/samples/")) {
      await serveStaticFrom(samplesDir, url.pathname.replace(/^\/samples/, ""), response);
      return;
    }

    await serveStaticFrom(publicDir, url.pathname, response);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    sendJson(response, 500, { error: message });
  }
});

server.listen(port, "127.0.0.1", () => {
  console.log(`Outlook-TS web app running at http://127.0.0.1:${port}`);
});

async function handleDiagnose(request: IncomingMessage, response: ServerResponse): Promise<void> {
  const body = await readRequestBody(request);
  const diagnostics = JSON.parse(stripByteOrderMark(body)) as OutlookDiagnostics;
  const findings = evaluateDiagnostics(diagnostics);
  sendJson(response, 200, {
    diagnostics,
    findings,
    markdown: renderMarkdownReport(diagnostics, findings)
  });
}

async function serveStaticFrom(baseDir: string, pathname: string, response: ServerResponse): Promise<void> {
  const requestedPath = pathname === "/" ? "/index.html" : decodeURIComponent(pathname);
  const safePath = normalize(requestedPath).replace(/^(\.\.[/\\])+/, "");
  const filePath = join(baseDir, safePath);

  if (!filePath.startsWith(baseDir)) {
    sendText(response, 403, "Forbidden");
    return;
  }

  try {
    const data = await readFile(filePath);
    const contentType = mimeTypes.get(extname(filePath)) ?? "application/octet-stream";
    response.writeHead(200, {
      "Content-Type": contentType,
      "Cache-Control": "no-store"
    });
    response.end(data);
  } catch {
    sendText(response, 404, "Not found");
  }
}

function readRequestBody(request: IncomingMessage): Promise<string> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];

    request.on("data", (chunk: Buffer) => {
      chunks.push(chunk);
    });

    request.on("end", () => {
      resolve(Buffer.concat(chunks).toString("utf8"));
    });

    request.on("error", reject);
  });
}

function sendJson(response: ServerResponse, status: number, payload: unknown): void {
  response.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store"
  });
  response.end(JSON.stringify(payload));
}

function sendText(response: ServerResponse, status: number, message: string): void {
  response.writeHead(status, {
    "Content-Type": "text/plain; charset=utf-8",
    "Cache-Control": "no-store"
  });
  response.end(message);
}

function stripByteOrderMark(value: string): string {
  return value.charCodeAt(0) === 0xfeff ? value.slice(1) : value;
}
