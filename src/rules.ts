import type { Finding, OutlookDiagnostics } from "./types.js";

export function evaluateDiagnostics(input: OutlookDiagnostics): Finding[] {
  const findings: Finding[] = [];

  if (input.newOutlook?.useNewOutlook === 1) {
    findings.push({
      id: "new-outlook-toggle-enabled",
      severity: "critical",
      title: "Classic Outlook launch may be redirected to new Outlook",
      evidence: [
        `${input.newOutlook.useNewOutlookRegistryPath ?? "UseNewOutlook"} is set to 1.`,
        "Microsoft documents this toggle state as a reason classic Outlook launches redirect to new Outlook."
      ],
      recommendation:
        "Set UseNewOutlook to 0 for the affected Windows user, then launch Outlook (classic) from the Start menu.",
      repair: {
        script: "scripts/Repair-NewOutlookToggle.ps1",
        command: ".\\scripts\\Repair-NewOutlookToggle.ps1 -ForceClassicLaunch"
      }
    });
  }

  if (input.newOutlook?.installed && input.newOutlook.useNewOutlook === null) {
    findings.push({
      id: "new-outlook-installed-toggle-missing",
      severity: "warning",
      title: "New Outlook is installed but the classic/new toggle registry value was not found",
      evidence: [
        `New Outlook package: ${input.newOutlook.packageFullName ?? "installed"}.`,
        "The UseNewOutlook registry value was missing for this user."
      ],
      recommendation:
        "If classic Outlook still redirects, create the Preferences key and set UseNewOutlook to 0 for the affected user.",
      repair: {
        script: "scripts/Repair-NewOutlookToggle.ps1",
        command: ".\\scripts\\Repair-NewOutlookToggle.ps1 -ForceClassicLaunch"
      }
    });
  }

  const profileCount = input.classicOutlook?.profileCount ?? input.classicOutlook?.profiles?.length ?? 0;
  if (profileCount === 0) {
    findings.push({
      id: "classic-profile-missing",
      severity: "critical",
      title: "No classic Outlook profile was detected",
      evidence: ["The classic Outlook profile registry path did not contain any profiles."],
      recommendation:
        "Create a fresh classic Outlook profile after confirming Microsoft 365 sign-in and Autodiscover are healthy."
    });
  } else if (!input.classicOutlook?.defaultProfile) {
    findings.push({
      id: "classic-default-profile-missing",
      severity: "warning",
      title: "Classic Outlook profiles exist but no default profile was detected",
      evidence: [`Detected ${profileCount} classic Outlook profile(s).`, "DefaultProfile was empty or missing."],
      recommendation:
        "Set a default profile in Mail control panel or create a new clean profile and make it the default."
    });
  }

  for (const addin of input.addins ?? []) {
    if (addin.loadBehavior !== null && addin.loadBehavior !== undefined && ![0, 2, 3].includes(addin.loadBehavior)) {
      findings.push({
        id: `addin-unusual-load-behavior-${slug(addin.name)}`,
        severity: "warning",
        title: `Outlook add-in has unusual LoadBehavior: ${addin.name}`,
        evidence: [
          `${addin.hive}\\...\\${addin.name}`,
          `LoadBehavior=${addin.loadBehavior}`,
          addin.friendlyName ? `FriendlyName=${addin.friendlyName}` : "FriendlyName was not set."
        ],
        recommendation:
          "Disable this add-in during startup testing, then re-enable it only after classic Outlook opens reliably."
      });
    }
  }

  const recentCrashes = (input.eventLog ?? []).filter((entry) => {
    const text = `${entry.providerName ?? ""} ${entry.message ?? ""}`.toLowerCase();
    return text.includes("outlook") && (text.includes("fault") || text.includes("crash") || text.includes("hang"));
  });

  if (recentCrashes.length > 0) {
    findings.push({
      id: "recent-outlook-crashes",
      severity: "warning",
      title: "Recent Outlook crash or hang events were found",
      evidence: recentCrashes.slice(0, 3).map((entry) => {
        return `${entry.timeCreated ?? "unknown time"} ${entry.providerName ?? "unknown provider"} ${entry.id ?? ""}`.trim();
      }),
      recommendation:
        "Start classic Outlook in safe mode, test with COM add-ins disabled, then compare against a fresh profile."
    });
  }

  if (input.exchangeOnline?.cas?.owaEnabled === false) {
    findings.push({
      id: "owa-disabled",
      severity: "critical",
      title: "OWA is disabled for the mailbox",
      evidence: ["Exchange Online CAS setting OwaEnabled is false."],
      recommendation:
        "Enable OWA for this mailbox or confirm policy intent. New Outlook depends heavily on web-backed Outlook services."
    });
  }

  if (findings.length === 0) {
    findings.push({
      id: "no-blockers-detected",
      severity: "info",
      title: "No high-confidence blocker was detected",
      evidence: ["The supplied diagnostics did not match the current rule set."],
      recommendation:
        "Add Exchange Online and Graph collection for this user, then retest classic Outlook safe mode and a fresh profile."
    });
  }

  return findings;
}

function slug(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
}
