export type Severity = "info" | "warning" | "critical";

export interface OutlookDiagnostics {
  collectedAt: string;
  schemaVersion?: number;
  computerName?: string;
  userName?: string;
  targetUserPrincipalName?: string;
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
  collectedAt?: string;
  collectorVersion?: string;
  errors?: string[];
  mailboxFound?: boolean;
  primarySmtpAddress?: string;
  displayName?: string;
  recipientTypeDetails?: string;
  accountDisabled?: boolean;
  litigationHoldEnabled?: boolean;
  archiveStatus?: string;
  forwardingSmtpAddress?: string | null;
  forwardingAddress?: string | null;
  deliverToMailboxAndForward?: boolean;
  hiddenFromAddressListsEnabled?: boolean;
  retentionPolicy?: string | null;
  roleAssignmentPolicy?: string | null;
  owaMailboxPolicy?: string | null;
  issueWarningQuota?: string;
  prohibitSendQuota?: string;
  prohibitSendReceiveQuota?: string;
  totalItemSize?: string;
  totalItemSizeBytes?: number | null;
  quotaUsedPercent?: number | null;
  itemCount?: number;
  deletedItemCount?: number;
  totalDeletedItemSize?: string;
  lastLogonTime?: string;
  mailboxPermissionSummary?: {
    fullAccessDelegates?: string[];
    nonInheritedPermissionCount?: number;
  };
  calendarFolders?: ExchangeCalendarFolderSummary[];
  defaultCalendar?: ExchangeCalendarFolderSummary;
  calendarFolderSettings?: {
    identity?: string;
    publishEnabled?: boolean;
    detailLevel?: string | null;
    searchableUrlEnabled?: boolean;
  };
  calendarFolderPermissions?: ExchangeCalendarPermissionSummary[];
  cas?: {
    owaEnabled?: boolean;
    mapiEnabled?: boolean;
    imapEnabled?: boolean;
    popEnabled?: boolean;
    activeSyncEnabled?: boolean;
    smtpClientAuthenticationDisabled?: boolean;
    ewsEnabled?: boolean;
  };
}

export interface ExchangeCalendarFolderSummary {
  name?: string;
  folderPath?: string;
  folderType?: string;
  itemsInFolder?: number;
  itemsInFolderAndSubfolders?: number;
  folderSize?: string;
  folderAndSubfolderSize?: string;
  oldestItemReceivedDate?: string | null;
  newestItemReceivedDate?: string | null;
}

export interface ExchangeCalendarPermissionSummary {
  user?: string;
  accessRights?: string[];
  sharingPermissionFlags?: string[];
}

export interface GraphDiagnostics {
  collectedAt?: string;
  collectorVersion?: string;
  errors?: string[];
  signedInUserPrincipalName?: string;
  targetUserPrincipalName?: string;
  mailboxSettingsAvailable?: boolean;
  mailboxSettings?: {
    timeZone?: string;
    dateFormat?: string;
    timeFormat?: string;
    workingHoursTimeZone?: string;
    automaticRepliesStatus?: string;
  };
  folderCount?: number;
  hiddenFolderCount?: number;
  largestFolders?: MailFolderSummary[];
  inboxRuleCount?: number;
  enabledInboxRuleCount?: number;
  forwardingRuleCount?: number;
  suspiciousRules?: MessageRuleSummary[];
  calendar?: CalendarDiagnostics;
}

export interface CalendarDiagnostics {
  errors?: string[];
  calendarCount?: number;
  calendars?: CalendarSummary[];
  defaultCalendar?: CalendarSummary;
  syncWindowStart?: string;
  syncWindowEnd?: string;
  eventCount?: number;
  recurringEventCount?: number;
  cancelledEventCount?: number;
  exceptionEventCount?: number;
  eventTimeZoneMismatchCount?: number;
}

export interface CalendarSummary {
  id?: string;
  name?: string;
  canEdit?: boolean;
  canShare?: boolean;
  canViewPrivateItems?: boolean;
  isDefaultCalendar?: boolean;
  ownerAddress?: string;
  eventCount?: number;
  recurringEventCount?: number;
  cancelledEventCount?: number;
}

export interface MailFolderSummary {
  displayName?: string;
  totalItemCount?: number;
  unreadItemCount?: number;
  childFolderCount?: number;
  isHidden?: boolean;
}

export interface MessageRuleSummary {
  displayName?: string;
  isEnabled?: boolean;
  sequence?: number;
  hasForwardingAction?: boolean;
  hasDeleteOrMoveAction?: boolean;
  forwardTo?: string[];
  redirectTo?: string[];
  moveToFolder?: string | null;
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
