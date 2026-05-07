export type Severity = "info" | "warning" | "critical";

export interface OutlookDiagnostics {
  collectedAt: string;
  computerName?: string;
  userName?: string;
  os?: {
    caption?: string;
    version?: string;
    buildNumber?: string;
  };
  office?: {
    platform?: string;
    clientVersionToReport?: string;
    updateChannel?: string;
    productReleaseIds?: string;
  };
  newOutlook?: {
    installed?: boolean;
    packageFullName?: string;
    version?: string;
    installLocation?: string;
    useNewOutlook?: number | null;
    useNewOutlookRegistryPath?: string;
  };
  classicOutlook?: {
    defaultProfile?: string | null;
    profiles?: string[];
    profileCount?: number;
    executablePaths?: string[];
  };
  addins?: AddinDiagnostic[];
  eventLog?: EventLogEntry[];
  exchangeOnline?: ExchangeOnlineDiagnostics;
  graph?: GraphDiagnostics;
}

export interface AddinDiagnostic {
  hive: string;
  name: string;
  friendlyName?: string;
  description?: string;
  loadBehavior?: number | null;
  manifest?: string;
}

export interface EventLogEntry {
  timeCreated?: string;
  providerName?: string;
  id?: number;
  levelDisplayName?: string;
  message?: string;
}

export interface ExchangeOnlineDiagnostics {
  mailboxFound?: boolean;
  primarySmtpAddress?: string;
  recipientTypeDetails?: string;
  litigationHoldEnabled?: boolean;
  archiveStatus?: string;
  issueWarningQuota?: string;
  prohibitSendQuota?: string;
  totalItemSize?: string;
  itemCount?: number;
  lastLogonTime?: string;
  cas?: {
    owaEnabled?: boolean;
    mapiEnabled?: boolean;
    imapEnabled?: boolean;
    popEnabled?: boolean;
    activeSyncEnabled?: boolean;
  };
}

export interface GraphDiagnostics {
  signedInUserPrincipalName?: string;
  mailboxSettingsAvailable?: boolean;
  folderCount?: number;
  hiddenFolderCount?: number;
  inboxRuleCount?: number;
}

export interface Finding {
  id: string;
  severity: Severity;
  title: string;
  evidence: string[];
  recommendation: string;
  repair?: {
    script: string;
    command: string;
  };
}
