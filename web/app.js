const fileInput = document.querySelector("#fileInput");
const dropzone = document.querySelector("#dropzone");
const jsonInput = document.querySelector("#jsonInput");
const analyzeButton = document.querySelector("#analyzeButton");
const exportMarkdown = document.querySelector("#exportMarkdown");
const loadClassicSample = document.querySelector("#loadClassicSample");
const loadCloudSample = document.querySelector("#loadCloudSample");
const statusText = document.querySelector("#status");
const findingsEl = document.querySelector("#findings");
const toolMailbox = document.querySelector("#toolMailbox");
const toolYears = document.querySelector("#toolYears");
const toolPurgeType = document.querySelector("#toolPurgeType");
const toolFolder = document.querySelector("#toolFolder");
const toolConfirm = document.querySelector("#toolConfirm");
const toolConnect = document.querySelector("#toolConnect");
const toolOutput = document.querySelector("#toolOutput");
const purviewPreviewButton = document.querySelector("#purviewPreviewButton");
const emailPreviewButton = document.querySelector("#emailPreviewButton");
const graphPreviewButton = document.querySelector("#graphPreviewButton");
const retentionButton = document.querySelector("#retentionButton");
const purviewPurgeButton = document.querySelector("#purviewPurgeButton");
const emailPurgeButton = document.querySelector("#emailPurgeButton");
const graphTenantId = document.querySelector("#graphTenantId");
const graphClientId = document.querySelector("#graphClientId");
const graphUserId = document.querySelector("#graphUserId");
const graphDaysBack = document.querySelector("#graphDaysBack");
const graphDaysForward = document.querySelector("#graphDaysForward");
const graphSignInButton = document.querySelector("#graphSignInButton");
const graphSignOutButton = document.querySelector("#graphSignOutButton");
const graphCollectButton = document.querySelector("#graphCollectButton");
const graphOutput = document.querySelector("#graphOutput");

const metaCollected = document.querySelector("#metaCollected");
const metaComputer = document.querySelector("#metaComputer");
const metaUser = document.querySelector("#metaUser");
const metaFindings = document.querySelector("#metaFindings");
const criticalCount = document.querySelector("#criticalCount");
const warningCount = document.querySelector("#warningCount");
const infoCount = document.querySelector("#infoCount");

let currentMarkdown = "";
const graphScopes = ["User.Read", "MailboxSettings.Read", "Mail.ReadBasic", "Mail.Read", "Calendars.Read"];

fileInput.addEventListener("change", async () => {
  const file = fileInput.files?.[0];
  if (file) {
    await loadFile(file);
  }
});

dropzone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropzone.classList.add("dragging");
});

dropzone.addEventListener("dragleave", () => {
  dropzone.classList.remove("dragging");
});

dropzone.addEventListener("drop", async (event) => {
  event.preventDefault();
  dropzone.classList.remove("dragging");
  const file = event.dataTransfer?.files?.[0];
  if (file) {
    await loadFile(file);
  }
});

analyzeButton.addEventListener("click", analyzeCurrentJson);
purviewPreviewButton.addEventListener("click", () => runTool("purview-preview"));
emailPreviewButton.addEventListener("click", () => runTool("email-preview"));
graphPreviewButton.addEventListener("click", () => runTool("graph-preview"));
retentionButton.addEventListener("click", () => runTool("retention-policy"));
purviewPurgeButton.addEventListener("click", () => runTool("purview-purge"));
emailPurgeButton.addEventListener("click", () => runTool("email-purge"));
graphSignInButton.addEventListener("click", signInGraph);
graphSignOutButton.addEventListener("click", signOutGraph);
graphCollectButton.addEventListener("click", collectGraphDiagnostics);

loadClassicSample.addEventListener("click", async () => {
  await loadSample("/samples/classic-modern-toggle.json");
});

loadCloudSample.addEventListener("click", async () => {
  await loadSample("/samples/cloud-expanded.json");
});

exportMarkdown.addEventListener("click", () => {
  if (!currentMarkdown) {
    return;
  }

  const blob = new Blob([currentMarkdown], { type: "text/markdown;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "outlook-ts-report.md";
  link.click();
  URL.revokeObjectURL(url);
});

void finishGraphRedirect();
loadGraphConfig();
renderGraphSession();

async function loadFile(file) {
  const text = await file.text();
  jsonInput.value = text;
  setStatus(`Loaded ${file.name}.`);
  await analyzeCurrentJson();
}

async function loadSample(path) {
  const response = await fetch(path, { cache: "no-store" });
  if (!response.ok) {
    setStatus(`Could not load sample: ${response.statusText}`, true);
    return;
  }

  jsonInput.value = await response.text();
  setStatus("Sample loaded.");
  await analyzeCurrentJson();
}

async function analyzeCurrentJson() {
  const payload = jsonInput.value.trim();
  if (!payload) {
    setStatus("Paste or upload diagnostics JSON first.", true);
    return;
  }

  setStatus("Analyzing...");
  analyzeButton.disabled = true;

  try {
    const response = await fetch("/api/diagnose", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: payload
    });

    const result = await response.json();
    if (!response.ok) {
      throw new Error(result.error ?? "Analysis failed.");
    }

    currentMarkdown = result.markdown ?? "";
    exportMarkdown.disabled = !currentMarkdown;
    renderReport(result.diagnostics, result.findings ?? []);
    setStatus("Report ready.");
  } catch (error) {
    currentMarkdown = "";
    exportMarkdown.disabled = true;
    setStatus(error instanceof Error ? error.message : String(error), true);
  } finally {
    analyzeButton.disabled = false;
  }
}

async function runTool(tool) {
  const mailbox = toolMailbox.value.trim();
  const olderThanYears = Number(toolYears.value || 2);
  const isDangerous = tool === "purview-purge" || tool === "email-purge" || tool === "retention-policy";

  if (!mailbox) {
    setToolOutput("Enter a mailbox first.", true);
    return;
  }

  if (isDangerous) {
    const required = tool === "retention-policy" ? "APPLY" : mailbox;
    if (toolConfirm.value.trim() !== required) {
      setToolOutput(`Confirmation required. Type ${required} in the Confirm box.`, true);
      return;
    }
  }

  setToolBusy(true);
  setToolOutput("Running PowerShell tool. Complete any sign-in prompts that open...");

  try {
    const response = await fetch("/api/tools/run", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        tool,
        mailbox,
        olderThanYears,
        purgeType: toolPurgeType.value,
        folderName: toolFolder.value.trim(),
        confirmText: toolConfirm.value.trim(),
        connect: toolConnect.checked
      })
    });

    const result = await response.json();
    if (!response.ok) {
      throw new Error(result.error ?? result.stderr ?? "Tool failed.");
    }

    const output = [
      result.stdout?.trim(),
      result.stderr?.trim() ? `STDERR:\n${result.stderr.trim()}` : "",
      `Exit code: ${result.exitCode}`
    ].filter(Boolean).join("\n\n");
    setToolOutput(output || "Tool completed.");
  } catch (error) {
    setToolOutput(error instanceof Error ? error.message : String(error), true);
  } finally {
    setToolBusy(false);
  }
}

function renderReport(diagnostics, findings) {
  const counts = countBySeverity(findings);
  criticalCount.textContent = String(counts.critical);
  warningCount.textContent = String(counts.warning);
  infoCount.textContent = String(counts.info);

  metaCollected.textContent = diagnostics?.collectedAt ?? "unknown";
  metaComputer.textContent = diagnostics?.computerName ?? "unknown";
  metaUser.textContent = diagnostics?.userName ?? diagnostics?.targetUserPrincipalName ?? "unknown";
  metaFindings.textContent = String(findings.length);

  findingsEl.classList.remove("empty");
  findingsEl.replaceChildren(...findings.map(renderFinding));
}

function renderFinding(finding) {
  const article = document.createElement("article");
  article.className = "finding";
  article.dataset.severity = finding.severity;

  const header = document.createElement("div");
  header.className = "findingHeader";

  const titleWrap = document.createElement("div");
  const title = document.createElement("h2");
  title.textContent = finding.title;
  const rule = document.createElement("p");
  rule.innerHTML = `Rule: <code>${escapeHtml(finding.id)}</code>`;
  titleWrap.append(title, rule);

  const badge = document.createElement("span");
  badge.className = `badge ${finding.severity}`;
  badge.textContent = finding.severity;
  header.append(titleWrap, badge);

  const evidence = document.createElement("ul");
  for (const item of finding.evidence ?? []) {
    const li = document.createElement("li");
    li.textContent = item;
    evidence.append(li);
  }

  const recommendation = document.createElement("p");
  recommendation.className = "recommendation";
  recommendation.textContent = finding.recommendation;

  article.append(header, evidence, recommendation);

  if (finding.repair?.command) {
    const repair = document.createElement("div");
    repair.className = "repair";

    const command = document.createElement("code");
    command.textContent = finding.repair.command;

    const copy = document.createElement("button");
    copy.type = "button";
    copy.className = "copyButton";
    copy.textContent = "Copy";
    copy.addEventListener("click", async () => {
      await navigator.clipboard.writeText(finding.repair.command);
      copy.textContent = "Copied";
      setTimeout(() => {
        copy.textContent = "Copy";
      }, 1200);
    });

    repair.append(command, copy);
    article.append(repair);
  }

  return article;
}

function countBySeverity(findings) {
  return findings.reduce(
    (counts, finding) => {
      counts[finding.severity] = (counts[finding.severity] ?? 0) + 1;
      return counts;
    },
    { critical: 0, warning: 0, info: 0 }
  );
}

function setStatus(message, isError = false) {
  statusText.textContent = message;
  statusText.style.color = isError ? "#c2410c" : "";
}

function setToolOutput(message, isError = false) {
  toolOutput.textContent = message;
  toolOutput.dataset.error = isError ? "true" : "false";
}

function setToolBusy(isBusy) {
  for (const button of [purviewPreviewButton, emailPreviewButton, graphPreviewButton, retentionButton, purviewPurgeButton, emailPurgeButton]) {
    button.disabled = isBusy;
  }
}

async function signInGraph() {
  const tenantId = graphTenantId.value.trim() || "common";
  const clientId = graphClientId.value.trim();

  if (!clientId) {
    setGraphOutput("Enter the Entra app client ID first.", true);
    return;
  }

  saveGraphConfig();
  const verifier = createRandomString(64);
  const challenge = await createPkceChallenge(verifier);
  const state = createRandomString(32);
  const redirectUri = getGraphRedirectUri();
  sessionStorage.setItem("outlookTsGraphAuth", JSON.stringify({ tenantId, clientId, verifier, state, redirectUri }));

  const params = new URLSearchParams({
    client_id: clientId,
    response_type: "code",
    redirect_uri: redirectUri,
    response_mode: "query",
    scope: graphScopes.join(" "),
    state,
    code_challenge: challenge,
    code_challenge_method: "S256",
    prompt: "select_account"
  });

  window.location.assign(`https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/authorize?${params}`);
}

async function finishGraphRedirect() {
  const params = new URLSearchParams(window.location.search);
  const code = params.get("code");
  if (!code) {
    return;
  }

  const pending = readJsonSession("outlookTsGraphAuth");
  if (!pending || params.get("state") !== pending.state) {
    setGraphOutput("Graph sign-in state did not match. Try signing in again.", true);
    return;
  }

  try {
    const body = new URLSearchParams({
      client_id: pending.clientId,
      grant_type: "authorization_code",
      code,
      redirect_uri: pending.redirectUri,
      code_verifier: pending.verifier,
      scope: graphScopes.join(" ")
    });

    const response = await fetch(`https://login.microsoftonline.com/${encodeURIComponent(pending.tenantId)}/oauth2/v2.0/token`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body
    });
    const token = await response.json();
    if (!response.ok) {
      throw new Error(token.error_description ?? token.error ?? "Token exchange failed.");
    }

    sessionStorage.setItem("outlookTsGraphToken", JSON.stringify({
      accessToken: token.access_token,
      expiresAt: Date.now() + Number(token.expires_in ?? 3600) * 1000,
      tenantId: pending.tenantId,
      clientId: pending.clientId,
      scopes: graphScopes
    }));
    sessionStorage.removeItem("outlookTsGraphAuth");
    window.history.replaceState({}, document.title, window.location.pathname);
    setGraphOutput("Signed in to Microsoft Graph.");
  } catch (error) {
    setGraphOutput(error instanceof Error ? error.message : String(error), true);
  } finally {
    renderGraphSession();
  }
}

function signOutGraph() {
  sessionStorage.removeItem("outlookTsGraphAuth");
  sessionStorage.removeItem("outlookTsGraphToken");
  renderGraphSession();
  setGraphOutput("Signed out.");
}

async function collectGraphDiagnostics() {
  const targetUser = graphUserId.value.trim();
  const token = getGraphToken();
  if (!token) {
    setGraphOutput("Sign in to Microsoft Graph first.", true);
    return;
  }

  if (!targetUser) {
    setGraphOutput("Enter a target user first.", true);
    return;
  }

  saveGraphConfig();
  setGraphBusy(true);
  setGraphOutput("Collecting Microsoft Graph diagnostics...");

  try {
    const daysBack = clampWholeNumber(Number(graphDaysBack.value || 14), 1, 365);
    const daysForward = clampWholeNumber(Number(graphDaysForward.value || 120), 1, 365);
    const diagnostics = await buildGraphDiagnostics(token.accessToken, targetUser, daysBack, daysForward);
    jsonInput.value = JSON.stringify(diagnostics, null, 2);
    setGraphOutput(`Collected Graph diagnostics for ${targetUser}.`);
    await analyzeCurrentJson();
  } catch (error) {
    setGraphOutput(error instanceof Error ? error.message : String(error), true);
  } finally {
    setGraphBusy(false);
  }
}

async function buildGraphDiagnostics(accessToken, targetUser, daysBack, daysForward) {
  const errors = [];
  const calendarErrors = [];
  const target = encodeURIComponent(targetUser);
  const collectedAt = new Date().toISOString();
  const folderResult = await graphGetPaged(accessToken, `/users/${target}/mailFolders?$top=100&$select=displayName,totalItemCount,unreadItemCount,childFolderCount`, errors, "List mail folders");
  const ruleResult = await graphGetPaged(accessToken, `/users/${target}/mailFolders/inbox/messageRules?$top=100`, errors, "List inbox rules");
  const calendarResult = await graphGetPaged(accessToken, `/users/${target}/calendars?$top=100`, calendarErrors, "List calendars");
  const settings = await graphGet(accessToken, `/users/${target}/mailboxSettings`, errors, "Read mailbox settings");
  const signedIn = await graphGet(accessToken, `/me?$select=userPrincipalName,mail,displayName`, [], "Read signed-in user");

  const syncStart = new Date(Date.now() - daysBack * 24 * 60 * 60 * 1000);
  const syncEnd = new Date(Date.now() + daysForward * 24 * 60 * 60 * 1000);
  const calendars = [];
  const allEvents = [];

  for (const calendar of calendarResult) {
    const calendarId = calendar.id;
    let events = [];
    if (calendarId) {
      events = await graphGetPaged(
        accessToken,
        `/users/${target}/calendars/${encodeURIComponent(calendarId)}/calendarView?startDateTime=${encodeURIComponent(syncStart.toISOString())}&endDateTime=${encodeURIComponent(syncEnd.toISOString())}&$top=100&$select=id,subject,start,end,type,isCancelled,isOrganizer,organizer,recurrence,lastModifiedDateTime`,
        calendarErrors,
        `Read calendar view for ${calendar.name ?? calendarId}`
      );
    }
    allEvents.push(...events);
    calendars.push({
      id: calendar.id,
      name: calendar.name,
      canEdit: calendar.canEdit,
      canShare: calendar.canShare,
      canViewPrivateItems: calendar.canViewPrivateItems,
      isDefaultCalendar: calendar.isDefaultCalendar,
      ownerAddress: calendar.owner?.address,
      eventCount: events.length,
      recurringEventCount: events.filter((event) => event.type && event.type !== "singleInstance").length,
      cancelledEventCount: events.filter((event) => event.isCancelled).length
    });
  }

  const largestFolders = folderResult
    .slice()
    .sort((a, b) => (b.totalItemCount ?? 0) - (a.totalItemCount ?? 0))
    .slice(0, 10)
    .map((folder) => ({
      displayName: folder.displayName,
      totalItemCount: folder.totalItemCount,
      unreadItemCount: folder.unreadItemCount,
      childFolderCount: folder.childFolderCount,
      isHidden: false
    }));

  const ruleSummaries = ruleResult.map(summarizeGraphRule);
  const mailboxTimeZone = settings?.timeZone;
  const defaultCalendar = calendars.find((calendar) => calendar.isDefaultCalendar) ?? calendars.find((calendar) => calendar.name === "Calendar") ?? null;

  return {
    collectedAt,
    targetUserPrincipalName: targetUser,
    graph: {
      collectedAt,
      collectorVersion: "0.2.0-browser",
      errors,
      signedInUserPrincipalName: signedIn?.userPrincipalName ?? signedIn?.mail,
      targetUserPrincipalName: targetUser,
      mailboxSettingsAvailable: Boolean(settings),
      mailboxSettings: {
        timeZone: settings?.timeZone ?? null,
        dateFormat: settings?.dateFormat ?? null,
        timeFormat: settings?.timeFormat ?? null,
        workingHoursTimeZone: settings?.workingHours?.timeZone?.name ?? null,
        automaticRepliesStatus: settings?.automaticRepliesSetting?.status ?? null
      },
      folderCount: folderResult.length,
      hiddenFolderCount: 0,
      largestFolders,
      inboxRuleCount: ruleResult.length,
      enabledInboxRuleCount: ruleResult.filter((rule) => rule.isEnabled).length,
      forwardingRuleCount: ruleSummaries.filter((rule) => rule.hasForwardingAction).length,
      suspiciousRules: ruleSummaries.filter((rule) => rule.hasForwardingAction || rule.hasDeleteOrMoveAction),
      calendar: {
        errors: calendarErrors,
        calendarCount: calendarResult.length,
        calendars,
        defaultCalendar,
        syncWindowStart: syncStart.toISOString(),
        syncWindowEnd: syncEnd.toISOString(),
        eventCount: allEvents.length,
        recurringEventCount: allEvents.filter((event) => event.type && event.type !== "singleInstance").length,
        cancelledEventCount: allEvents.filter((event) => event.isCancelled).length,
        exceptionEventCount: allEvents.filter((event) => event.type === "exception").length,
        eventTimeZoneMismatchCount: mailboxTimeZone ? allEvents.filter((event) => event.start?.timeZone && event.start.timeZone !== mailboxTimeZone).length : 0
      }
    }
  };
}

function summarizeGraphRule(rule) {
  const actions = rule.actions ?? {};
  const forwardTo = (actions.forwardTo ?? []).map((recipient) => recipient.emailAddress?.address ?? recipient.emailAddress?.name).filter(Boolean);
  const redirectTo = (actions.redirectTo ?? []).map((recipient) => recipient.emailAddress?.address ?? recipient.emailAddress?.name).filter(Boolean);
  return {
    displayName: rule.displayName,
    isEnabled: rule.isEnabled,
    sequence: rule.sequence,
    hasForwardingAction: forwardTo.length > 0 || redirectTo.length > 0,
    hasDeleteOrMoveAction: Boolean(actions.delete || actions.moveToFolder),
    forwardTo,
    redirectTo,
    moveToFolder: actions.moveToFolder ?? null
  };
}

async function graphGet(accessToken, path, errors, label) {
  try {
    const response = await fetch(`https://graph.microsoft.com${path}`, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });
    const result = await response.json();
    if (!response.ok) {
      throw new Error(result.error?.message ?? response.statusText);
    }
    return result;
  } catch (error) {
    errors.push(`${label} failed: ${error instanceof Error ? error.message : String(error)}`);
    return null;
  }
}

async function graphGetPaged(accessToken, path, errors, label) {
  const items = [];
  let next = `https://graph.microsoft.com${path}`;
  try {
    while (next) {
      const response = await fetch(next, {
        headers: { Authorization: `Bearer ${accessToken}` }
      });
      const result = await response.json();
      if (!response.ok) {
        throw new Error(result.error?.message ?? response.statusText);
      }
      items.push(...(result.value ?? []));
      next = result["@odata.nextLink"] ?? "";
    }
  } catch (error) {
    errors.push(`${label} failed: ${error instanceof Error ? error.message : String(error)}`);
  }
  return items;
}

function getGraphToken() {
  const token = readJsonSession("outlookTsGraphToken");
  if (!token?.accessToken || Number(token.expiresAt ?? 0) <= Date.now() + 60_000) {
    return null;
  }
  return token;
}

function loadGraphConfig() {
  const config = readJsonLocal("outlookTsGraphConfig");
  if (!config) {
    return;
  }
  graphTenantId.value = config.tenantId ?? "common";
  graphClientId.value = config.clientId ?? "";
  graphUserId.value = config.userId ?? "";
  graphDaysBack.value = config.daysBack ?? "14";
  graphDaysForward.value = config.daysForward ?? "120";
}

function saveGraphConfig() {
  localStorage.setItem("outlookTsGraphConfig", JSON.stringify({
    tenantId: graphTenantId.value.trim() || "common",
    clientId: graphClientId.value.trim(),
    userId: graphUserId.value.trim(),
    daysBack: graphDaysBack.value,
    daysForward: graphDaysForward.value
  }));
}

function renderGraphSession() {
  const token = getGraphToken();
  graphSignInButton.disabled = Boolean(token);
  graphSignOutButton.disabled = !token;
  graphCollectButton.disabled = !token;
  if (token) {
    setGraphOutput(`Signed in. Token expires at ${new Date(token.expiresAt).toLocaleString()}.`);
  }
}

function setGraphOutput(message, isError = false) {
  graphOutput.textContent = message;
  graphOutput.dataset.error = isError ? "true" : "false";
}

function setGraphBusy(isBusy) {
  graphCollectButton.disabled = isBusy || !getGraphToken();
  graphSignInButton.disabled = isBusy || Boolean(getGraphToken());
  graphSignOutButton.disabled = isBusy || !getGraphToken();
}

function getGraphRedirectUri() {
  return `${window.location.origin}${window.location.pathname}`;
}

function clampWholeNumber(value, min, max) {
  if (!Number.isFinite(value)) {
    return min;
  }
  return Math.min(max, Math.max(min, Math.trunc(value)));
}

function readJsonSession(key) {
  try {
    return JSON.parse(sessionStorage.getItem(key) ?? "null");
  } catch {
    return null;
  }
}

function readJsonLocal(key) {
  try {
    return JSON.parse(localStorage.getItem(key) ?? "null");
  } catch {
    return null;
  }
}

function createRandomString(length) {
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~";
  const bytes = new Uint8Array(length);
  crypto.getRandomValues(bytes);
  return Array.from(bytes, (byte) => alphabet[byte % alphabet.length]).join("");
}

async function createPkceChallenge(verifier) {
  const data = new TextEncoder().encode(verifier);
  const digest = await crypto.subtle.digest("SHA-256", data);
  return base64UrlEncode(new Uint8Array(digest));
}

function base64UrlEncode(bytes) {
  let binary = "";
  for (const byte of bytes) {
    binary += String.fromCharCode(byte);
  }
  return btoa(binary).replaceAll("+", "-").replaceAll("/", "_").replaceAll("=", "");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
