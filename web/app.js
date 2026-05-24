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

const metaCollected = document.querySelector("#metaCollected");
const metaComputer = document.querySelector("#metaComputer");
const metaUser = document.querySelector("#metaUser");
const metaFindings = document.querySelector("#metaFindings");
const criticalCount = document.querySelector("#criticalCount");
const warningCount = document.querySelector("#warningCount");
const infoCount = document.querySelector("#infoCount");

let currentMarkdown = "";

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
    const required = tool === "purview-purge" ? mailbox : "APPLY";
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

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
