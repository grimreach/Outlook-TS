import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import { spawn } from "node:child_process";
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

    if (request.method === "POST" && url.pathname === "/api/tools/run") {
      await handleToolRun(request, response);
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

async function handleToolRun(request: IncomingMessage, response: ServerResponse): Promise<void> {
  const body = await readRequestBody(request);
  const payload = JSON.parse(stripByteOrderMark(body)) as {
    tool?: string;
    mailbox?: string;
    olderThanYears?: number;
    purgeType?: string;
    confirmText?: string;
    connect?: boolean;
  };

  const mailbox = String(payload.mailbox ?? "").trim();
  if (!mailbox || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(mailbox)) {
    sendJson(response, 400, { error: "Enter a valid mailbox UPN." });
    return;
  }

  const olderThanYears = Number(payload.olderThanYears ?? 2);
  if (!Number.isInteger(olderThanYears) || olderThanYears < 1 || olderThanYears > 25) {
    sendJson(response, 400, { error: "Older-than years must be a whole number from 1 to 25." });
    return;
  }

  const connect = payload.connect !== false;
  const baseArgs = connect ? ["-Connect"] : [];
  const scriptDir = join(process.cwd(), "scripts");
  let args: string[];

  switch (payload.tool) {
    case "purview-preview":
      args = [
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        join(scriptDir, "Invoke-PurviewCalendarPurge.ps1"),
        "-Mailbox",
        mailbox,
        "-OlderThanYears",
        String(olderThanYears),
        ...baseArgs
      ];
      break;
    case "purview-purge":
      if (payload.confirmText !== mailbox) {
        sendJson(response, 400, { error: "Type the mailbox address exactly before purging." });
        return;
      }
      args = [
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        join(scriptDir, "Invoke-PurviewCalendarPurge.ps1"),
        "-Mailbox",
        mailbox,
        "-OlderThanYears",
        String(olderThanYears),
        "-Purge",
        "-PurgeType",
        payload.purgeType === "SoftDelete" ? "SoftDelete" : "HardDelete",
        ...baseArgs
      ];
      break;
    case "graph-preview":
      args = [
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        join(scriptDir, "Remove-OldCalendarItems.ps1"),
        "-UserId",
        mailbox,
        "-OlderThanYears",
        String(olderThanYears),
        "-OutputPath",
        ".\\old-calendar-items.json",
        "-CsvPath",
        ".\\old-calendar-items.csv",
        ...baseArgs
      ];
      break;
    case "retention-policy":
      if (payload.confirmText !== "APPLY") {
        sendJson(response, 400, { error: "Type APPLY before assigning a retention policy." });
        return;
      }
      args = [
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        join(scriptDir, "New-CalendarRetentionPolicy.ps1"),
        "-Identity",
        mailbox,
        "-OlderThanDays",
        String(olderThanYears * 365),
        "-RetentionAction",
        "DeleteAndAllowRecovery",
        "-Assign",
        "-StartManagedFolderAssistant",
        ...baseArgs
      ];
      break;
    default:
      sendJson(response, 400, { error: "Unknown tool." });
      return;
  }

  try {
    const result = await runPowerShell(args, 30 * 60 * 1000);
    sendJson(response, result.exitCode === 0 ? 200 : 500, result);
  } catch (error) {
    sendJson(response, 500, { error: error instanceof Error ? error.message : String(error) });
  }
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

function runPowerShell(args: string[], timeoutMs: number): Promise<{ exitCode: number | null; stdout: string; stderr: string }> {
  return new Promise((resolve, reject) => {
    const command = process.platform === "win32" ? "pwsh.exe" : "pwsh";
    const child = spawn(command, args, {
      cwd: process.cwd(),
      windowsHide: false
    });

    let stdout = "";
    let stderr = "";
    let settled = false;
    const timer = setTimeout(() => {
      settled = true;
      child.kill();
      reject(new Error("PowerShell command timed out."));
    }, timeoutMs);

    child.stdout.on("data", (chunk: Buffer) => {
      stdout += chunk.toString("utf8");
    });

    child.stderr.on("data", (chunk: Buffer) => {
      stderr += chunk.toString("utf8");
    });

    child.on("error", (error) => {
      if (!settled) {
        settled = true;
        clearTimeout(timer);
        reject(error);
      }
    });

    child.on("close", (exitCode) => {
      if (!settled) {
        settled = true;
        clearTimeout(timer);
        resolve({ exitCode, stdout, stderr });
      }
    });
  });
}
