import type { Finding, OutlookDiagnostics } from "./types.js";

const severityRank = {
  critical: 0,
  warning: 1,
  info: 2
} as const;

export function sortFindings(findings: Finding[]): Finding[] {
  return [...findings].sort((left, right) => {
    return severityRank[left.severity] - severityRank[right.severity] || left.title.localeCompare(right.title);
  });
}

export function renderMarkdownReport(input: OutlookDiagnostics, findings: Finding[]): string {
  const lines: string[] = [];
  lines.push("# Outlook-TS Diagnostic Report");
  lines.push("");
  lines.push(`Collected: ${input.collectedAt || "unknown"}`);
  lines.push(`Computer: ${input.computerName || "unknown"}`);
  lines.push(`User: ${input.userName || "unknown"}`);
  lines.push("");
  lines.push("## Findings");
  lines.push("");

  for (const finding of sortFindings(findings)) {
    lines.push(`### [${finding.severity.toUpperCase()}] ${finding.title}`);
    lines.push("");
    lines.push(`Rule: \`${finding.id}\``);
    lines.push("");
    lines.push("Evidence:");
    for (const item of finding.evidence) {
      lines.push(`- ${item}`);
    }
    lines.push("");
    lines.push(`Recommendation: ${finding.recommendation}`);
    if (finding.repair) {
      lines.push("");
      lines.push(`Repair: \`${finding.repair.command}\``);
    }
    lines.push("");
  }

  return lines.join("\n");
}

export function renderConsoleReport(input: OutlookDiagnostics, findings: Finding[]): string {
  const lines: string[] = [];
  lines.push("Outlook-TS Diagnostic Report");
  lines.push(`Collected: ${input.collectedAt || "unknown"}`);
  lines.push(`Computer: ${input.computerName || "unknown"}`);
  lines.push(`User: ${input.userName || "unknown"}`);
  lines.push("");

  for (const finding of sortFindings(findings)) {
    lines.push(`[${finding.severity.toUpperCase()}] ${finding.title}`);
    lines.push(`Rule: ${finding.id}`);
    for (const item of finding.evidence) {
      lines.push(`  - ${item}`);
    }
    lines.push(`Recommendation: ${finding.recommendation}`);
    if (finding.repair) {
      lines.push(`Repair: ${finding.repair.command}`);
    }
    lines.push("");
  }

  return lines.join("\n");
}
