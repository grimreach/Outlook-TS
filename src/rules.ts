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

  if (input.classicOutlook) {
    const profileCount = input.classicOutlook.profileCount ?? input.classicOutlook.profiles?.length ?? 0;
    if (profileCount === 0) {
      findings.push({
        id: "classic-profile-missing",
        severity: "critical",
        title: "No classic Outlook profile was detected",
        evidence: ["The classic Outlook profile registry path did not contain any profiles."],
        recommendation:
          "Create a fresh classic Outlook profile after confirming Microsoft 365 sign-in and Autodiscover are healthy."
      });
    } else if (!input.classicOutlook.defaultProfile) {
      findings.push({
        id: "classic-default-profile-missing",
        severity: "warning",
        title: "Classic Outlook profiles exist but no default profile was detected",
        evidence: [`Detected ${profileCount} classic Outlook profile(s).`, "DefaultProfile was empty or missing."],
        recommendation:
          "Set a default profile in Mail control panel or create a new clean profile and make it the default."
      });
    }
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

  if (input.exchangeOnline) {
    findings.push(...evaluateExchangeOnline(input));
  }

  if (input.graph) {
    findings.push(...evaluateGraph(input));
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

function evaluateExchangeOnline(input: OutlookDiagnostics): Finding[] {
  const findings: Finding[] = [];
  const exchange = input.exchangeOnline;
  if (!exchange) {
    return findings;
  }

  for (const error of exchange.errors ?? []) {
    findings.push({
      id: "exchange-collector-error",
      severity: "warning",
      title: "Exchange Online collector reported an error",
      evidence: [error],
      recommendation:
        "Review the Exchange Online collector permissions, module installation, and target mailbox identity."
    });
  }

  if (exchange.mailboxFound === false) {
    findings.push({
      id: "exchange-mailbox-not-found",
      severity: "critical",
      title: "Exchange Online mailbox was not found",
      evidence: [`Target: ${input.targetUserPrincipalName ?? exchange.primarySmtpAddress ?? "unknown"}`],
      recommendation:
        "Confirm the user has an Exchange Online mailbox and that the collector was run against the correct tenant."
    });
    return findings;
  }

  if (exchange.cas?.mapiEnabled === false) {
    findings.push({
      id: "mapi-disabled",
      severity: "critical",
      title: "MAPI is disabled for the mailbox",
      evidence: ["Exchange Online CAS setting MapiEnabled is false."],
      recommendation:
        "Enable MAPI for the mailbox if classic Outlook for Windows is expected to connect."
    });
  }

  if (exchange.cas?.ewsEnabled === false) {
    findings.push({
      id: "ews-disabled",
      severity: "warning",
      title: "EWS is disabled for the mailbox",
      evidence: ["Exchange Online CAS setting EwsEnabled is false."],
      recommendation:
        "Confirm this is intentional. Some Outlook features and add-ins still depend on EWS-backed service access."
    });
  }

  if (exchange.quotaUsedPercent !== null && exchange.quotaUsedPercent !== undefined && exchange.quotaUsedPercent >= 95) {
    findings.push({
      id: "mailbox-quota-critical",
      severity: "critical",
      title: "Mailbox is at or above 95% of send/receive quota",
      evidence: [
        `Quota used: ${exchange.quotaUsedPercent.toFixed(1)}%`,
        `Total item size: ${exchange.totalItemSize ?? "unknown"}`,
        `Prohibit send/receive quota: ${exchange.prohibitSendReceiveQuota ?? exchange.prohibitSendQuota ?? "unknown"}`
      ],
      recommendation:
        "Reduce mailbox size, enable or expand archive, or adjust quota before troubleshooting client sync behavior."
    });
  } else if (exchange.quotaUsedPercent !== null && exchange.quotaUsedPercent !== undefined && exchange.quotaUsedPercent >= 85) {
    findings.push({
      id: "mailbox-quota-warning",
      severity: "warning",
      title: "Mailbox is approaching quota",
      evidence: [
        `Quota used: ${exchange.quotaUsedPercent.toFixed(1)}%`,
        `Total item size: ${exchange.totalItemSize ?? "unknown"}`
      ],
      recommendation:
        "Treat mailbox size as a possible contributor to sync and performance symptoms."
    });
  }

  if (exchange.forwardingSmtpAddress || exchange.forwardingAddress) {
    findings.push({
      id: "mailbox-forwarding-enabled",
      severity: "warning",
      title: "Mailbox-level forwarding is enabled",
      evidence: [
        `ForwardingSmtpAddress=${exchange.forwardingSmtpAddress ?? "not set"}`,
        `ForwardingAddress=${exchange.forwardingAddress ?? "not set"}`,
        `DeliverToMailboxAndForward=${exchange.deliverToMailboxAndForward ?? "unknown"}`
      ],
      recommendation:
        "Verify forwarding is expected. Unexpected forwarding can look like missing mail or inconsistent mailbox behavior."
    });
  }

  const archiveStatus = exchange.archiveStatus?.toLowerCase();
  if (archiveStatus && archiveStatus !== "active") {
    findings.push({
      id: "archive-mailbox-not-active",
      severity: "warning",
      title: "Archive mailbox is not active",
      evidence: [
        `ArchiveStatus=${exchange.archiveStatus}`,
        `ArchiveState=${exchange.archiveState ?? "unknown"}`
      ],
      recommendation:
        "Enable the archive mailbox for users with large mailbox or calendar workloads, especially when Outlook/OWA calendar edit behavior is failing."
    });
  }

  if (exchange.autoExpandingArchiveEnabled === false) {
    findings.push({
      id: "auto-expanding-archive-disabled",
      severity: "warning",
      title: "Auto-expanding archive is disabled",
      evidence: ["AutoExpandingArchiveEnabled=false"],
      recommendation:
        "Consider enabling auto-expanding archive when Microsoft identifies calendar or mailbox capacity pressure."
    });
  }

  const delegateCount = exchange.mailboxPermissionSummary?.nonInheritedPermissionCount ?? 0;
  if (delegateCount > 10) {
    findings.push({
      id: "many-full-access-delegates",
      severity: "warning",
      title: "Mailbox has many explicit FullAccess permissions",
      evidence: [
        `Explicit delegate count: ${delegateCount}`,
        `Sample: ${(exchange.mailboxPermissionSummary?.fullAccessDelegates ?? []).slice(0, 5).join(", ") || "none"}`
      ],
      recommendation:
        "Review mailbox delegation. Large or stale delegate sets can complicate automapping and shared mailbox behavior."
    });
  }

  findings.push(...evaluateExchangeCalendar(input));

  return findings;
}

function evaluateExchangeCalendar(input: OutlookDiagnostics): Finding[] {
  const findings: Finding[] = [];
  const exchange = input.exchangeOnline;
  if (!exchange) {
    return findings;
  }

  const calendarFolders = exchange.calendarFolders ?? [];
  const totalCalendarItems = calendarFolders.reduce((sum, folder) => {
    return sum + (folder.itemsInFolderAndSubfolders ?? folder.itemsInFolder ?? 0);
  }, 0);
  const largestCalendar = [...calendarFolders].sort((left, right) => {
    return (right.itemsInFolderAndSubfolders ?? right.itemsInFolder ?? 0) - (left.itemsInFolderAndSubfolders ?? left.itemsInFolder ?? 0);
  })[0];

  if (totalCalendarItems >= 7000) {
    findings.push({
      id: "calendar-item-count-capacity-risk",
      severity: "critical",
      title: "Calendar folders contain a high number of items",
      evidence: [
        `Total calendar items across calendar folders: ${totalCalendarItems}`,
        `Largest calendar folder: ${largestCalendar?.folderPath ?? largestCalendar?.name ?? "unknown"} (${largestCalendar ? largestCalendar.itemsInFolderAndSubfolders ?? largestCalendar.itemsInFolder ?? 0 : "unknown"} item(s))`,
        `ArchiveStatus=${exchange.archiveStatus ?? "unknown"}`,
        `AutoExpandingArchiveEnabled=${exchange.autoExpandingArchiveEnabled ?? "unknown"}`
      ],
      recommendation:
        "Treat this as calendar capacity risk. If OWA/new Outlook can create but not edit/save calendar items, confirm archive and auto-expanding archive are enabled with Enable-Mailbox -Archive and Enable-Mailbox -AutoExpandingArchive."
    });
  } else if (totalCalendarItems >= 4000) {
    findings.push({
      id: "calendar-item-count-warning",
      severity: "warning",
      title: "Calendar folders contain many items",
      evidence: [
        `Total calendar items across calendar folders: ${totalCalendarItems}`,
        `Largest calendar folder: ${largestCalendar?.folderPath ?? largestCalendar?.name ?? "unknown"} (${largestCalendar ? largestCalendar.itemsInFolderAndSubfolders ?? largestCalendar.itemsInFolder ?? 0 : "unknown"} item(s))`
      ],
      recommendation:
        "Monitor for calendar edit/sync symptoms and consider archive/retention review before this becomes a service-impacting calendar workload."
    });
  }

  if (calendarFolders.length >= 25) {
    findings.push({
      id: "exchange-many-calendar-folders",
      severity: "warning",
      title: "Exchange reports many calendar folders",
      evidence: [`Calendar folder count: ${calendarFolders.length}`],
      recommendation:
        "Review duplicate, shared, or stale calendar folders. If OWA is empty, confirm the user is viewing the folder that actually contains calendar items."
    });
  }

  if (exchange.defaultCalendar && (exchange.defaultCalendar.itemsInFolderAndSubfolders ?? exchange.defaultCalendar.itemsInFolder ?? 0) === 0) {
    findings.push({
      id: "exchange-default-calendar-empty",
      severity: "warning",
      title: "Exchange reports the default calendar is empty",
      evidence: [
        `Folder: ${exchange.defaultCalendar.folderPath ?? exchange.defaultCalendar.name ?? "Calendar"}`,
        `Items: ${exchange.defaultCalendar.itemsInFolder ?? 0}`
      ],
      recommendation:
        "If the user expects meetings in the default calendar, compare against Graph calendar view and check whether items are in another calendar folder."
    });
  }

  const populatedNonDefault = calendarFolders
    .filter((folder) => folder.folderPath !== exchange.defaultCalendar?.folderPath)
    .filter((folder) => (folder.itemsInFolderAndSubfolders ?? folder.itemsInFolder ?? 0) > 0)
    .slice(0, 5);
  if ((exchange.defaultCalendar?.itemsInFolder ?? 0) === 0 && populatedNonDefault.length > 0) {
    findings.push({
      id: "calendar-items-in-non-default-folders",
      severity: "warning",
      title: "Calendar items appear to be in non-default calendar folders",
      evidence: populatedNonDefault.map((folder) => {
        return `${folder.folderPath ?? folder.name ?? "unknown"} has ${folder.itemsInFolderAndSubfolders ?? folder.itemsInFolder ?? 0} item(s)`;
      }),
      recommendation:
        "Ask the user which calendar should contain the meetings. A migrated or recreated calendar can leave events outside the default calendar that new Outlook expects."
    });
  }

  const defaultPermission = (exchange.calendarFolderPermissions ?? []).find((permission) => {
    return permission.user?.toLowerCase() === "default";
  });
  if (defaultPermission?.accessRights?.some((right) => ["Editor", "PublishingEditor", "Owner"].includes(right))) {
    findings.push({
      id: "calendar-default-permission-elevated",
      severity: "warning",
      title: "Default calendar permission is elevated",
      evidence: [`Default permission: ${(defaultPermission.accessRights ?? []).join(", ")}`],
      recommendation:
        "Confirm this is intentional. Overly broad default permissions can confuse support triage and expose calendar data."
    });
  }

  return findings;
}

function evaluateGraph(input: OutlookDiagnostics): Finding[] {
  const findings: Finding[] = [];
  const graph = input.graph;
  if (!graph) {
    return findings;
  }

  for (const error of graph.errors ?? []) {
    findings.push({
      id: "graph-collector-error",
      severity: "warning",
      title: "Microsoft Graph collector reported an error",
      evidence: [error],
      recommendation:
        "Review Graph PowerShell installation, delegated scopes, admin consent, and target mailbox access."
    });
  }

  if (graph.mailboxSettingsAvailable === false) {
    findings.push({
      id: "graph-mailbox-settings-unavailable",
      severity: "warning",
      title: "Graph could not read mailbox settings",
      evidence: [`Target: ${graph.targetUserPrincipalName ?? input.targetUserPrincipalName ?? "unknown"}`],
      recommendation:
        "Confirm the signed-in account has access and Graph has MailboxSettings.Read or an equivalent approved permission."
    });
  }

  if ((graph.hiddenFolderCount ?? 0) >= 20) {
    findings.push({
      id: "many-hidden-mail-folders",
      severity: "warning",
      title: "Mailbox has many hidden folders",
      evidence: [
        `Hidden folders: ${graph.hiddenFolderCount}`,
        `Total folders sampled: ${graph.folderCount ?? "unknown"}`
      ],
      recommendation:
        "Review hidden/system folder growth. Large hidden folder counts can point to stale client, rule, or sync state."
    });
  }

  if ((graph.forwardingRuleCount ?? 0) > 0) {
    findings.push({
      id: "inbox-forwarding-rules",
      severity: "warning",
      title: "Inbox rules include forwarding or redirect actions",
      evidence: (graph.suspiciousRules ?? [])
        .filter((rule) => rule.hasForwardingAction)
        .slice(0, 5)
        .map((rule) => `${rule.displayName ?? "unnamed rule"} forwards or redirects mail`),
      recommendation:
        "Verify every forwarding or redirect rule with the user. Unexpected rules are a common cause of missing mail."
    });
  }

  if ((graph.inboxRuleCount ?? 0) >= 100) {
    findings.push({
      id: "many-inbox-rules",
      severity: "warning",
      title: "Mailbox has a large number of inbox rules",
      evidence: [
        `Inbox rules: ${graph.inboxRuleCount}`,
        `Enabled rules: ${graph.enabledInboxRuleCount ?? "unknown"}`
      ],
      recommendation:
        "Review and simplify inbox rules before treating the issue as purely client-side."
    });
  }

  findings.push(...evaluateGraphCalendar(input));

  return findings;
}

function evaluateGraphCalendar(input: OutlookDiagnostics): Finding[] {
  const findings: Finding[] = [];
  const calendar = input.graph?.calendar;
  if (!calendar) {
    return findings;
  }

  for (const error of calendar.errors ?? []) {
    findings.push({
      id: "graph-calendar-collector-error",
      severity: "warning",
      title: "Microsoft Graph calendar collector reported an error",
      evidence: [error],
      recommendation:
        "Confirm Graph has Calendars.Read consent and that the signed-in technician can read the target user's calendar."
    });
  }

  if (calendar.calendarCount === 0) {
    findings.push({
      id: "graph-calendar-none-found",
      severity: "critical",
      title: "Graph did not return any calendars for the mailbox",
      evidence: [`Target: ${input.graph?.targetUserPrincipalName ?? input.targetUserPrincipalName ?? "unknown"}`],
      recommendation:
        "Confirm the mailbox is healthy and that Graph calendar permissions are granted before treating this as a client reinstall issue."
    });
  }

  if ((calendar.calendarCount ?? 0) >= 25) {
    findings.push({
      id: "many-calendars",
      severity: "warning",
      title: "Mailbox has many calendars",
      evidence: [`Calendars returned by Graph: ${calendar.calendarCount}`],
      recommendation:
        "Review shared/duplicate calendars. A large calendar set can make new Outlook calendar sync symptoms harder to isolate."
    });
  }

  if ((calendar.defaultCalendar?.eventCount ?? 0) === 0 && (calendar.calendarCount ?? 0) > 0) {
    findings.push({
      id: "default-calendar-empty-in-window",
      severity: "warning",
      title: "Default calendar has no events in the sampled sync window",
      evidence: [
        `Window: ${calendar.syncWindowStart ?? "unknown"} to ${calendar.syncWindowEnd ?? "unknown"}`,
        `Default calendar: ${calendar.defaultCalendar?.name ?? "unknown"}`
      ],
      recommendation:
        "Compare against Outlook on the web. If OWA shows events but Graph does not, focus on mailbox/service access. If Graph shows none because events are in another calendar, check which calendar the user expects to sync."
    });
  }

  if ((calendar.recurringEventCount ?? 0) >= 100) {
    findings.push({
      id: "many-recurring-calendar-events",
      severity: "warning",
      title: "Calendar sync window contains many recurring events",
      evidence: [
        `Recurring events in sampled window: ${calendar.recurringEventCount}`,
        `Total events in sampled window: ${calendar.eventCount ?? "unknown"}`
      ],
      recommendation:
        "Ask whether the issue involves recurring meetings. Recurring series and exceptions are common places to compare new Outlook, OWA, and classic Outlook behavior."
    });
  }

  if ((calendar.cancelledEventCount ?? 0) >= 25) {
    findings.push({
      id: "many-cancelled-calendar-events",
      severity: "warning",
      title: "Calendar sync window contains many cancelled events",
      evidence: [
        `Cancelled events in sampled window: ${calendar.cancelledEventCount}`,
        `Window: ${calendar.syncWindowStart ?? "unknown"} to ${calendar.syncWindowEnd ?? "unknown"}`
      ],
      recommendation:
        "Review whether cancelled or stale meetings are the items failing to sync. Compare item visibility in OWA before reinstalling the client again."
    });
  }

  if ((calendar.eventTimeZoneMismatchCount ?? 0) > 0) {
    findings.push({
      id: "calendar-event-time-zone-mismatches",
      severity: "warning",
      title: "Some calendar events use a different time zone than mailbox settings",
      evidence: [
        `Mismatched events in sampled window: ${calendar.eventTimeZoneMismatchCount}`,
        `Mailbox time zone: ${input.graph?.mailboxSettings?.timeZone ?? "unknown"}`
      ],
      recommendation:
        "If the symptom is wrong meeting times, compare mailbox time zone, Windows time zone, OWA time zone, and the event time zones."
    });
  }

  return findings;
}

function slug(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
}
