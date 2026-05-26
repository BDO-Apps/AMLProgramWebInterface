/**
 * BDO Audit Software Dev Site — SharePoint list & library IDs
 * https://bdozw.sharepoint.com/sites/AuditSoftwareDevSite
 */
export const sharePointConfig = {
  hostname: import.meta.env.VITE_SP_HOSTNAME ?? 'bdozw.sharepoint.com',
  sitePath: import.meta.env.VITE_SP_SITE_PATH ?? '/sites/AuditSoftwareDevSite',
  siteUrl:
    import.meta.env.VITE_SHAREPOINT_SITE_URL ??
    'https://bdozw.sharepoint.com/sites/AuditSoftwareDevSite',

  /** Optional GUID overrides — leave unset to auto-resolve by list name from the site */
  lists: {
    clientOnboarding: import.meta.env.VITE_SP_LIST_CLIENT_ONBOARDING ?? '',
    directors: import.meta.env.VITE_SP_LIST_DIRECTORS ?? '',
    beneficialOwners: import.meta.env.VITE_SP_LIST_BENEFICIAL_OWNERS ?? '',
    auditLogs: import.meta.env.VITE_SP_LIST_AUDIT_LOGS ?? '',
  },

  /** Display / internal names used when resolving lists (case-insensitive) */
  listNames: {
    clientOnboarding: ['ClientOnboarding'],
    directors: ['Directors'],
    beneficialOwners: ['BeneficialOwners'],
    auditLogs: ['AuditLogs', 'Auditlogs'],
    documentLibrary: ['AMLOnboardingDocuments'],
  },

  documentLibrary: {
    name: 'AMLOnboardingDocuments',
    listId: import.meta.env.VITE_SP_LIBRARY_AML_DOCUMENTS ?? '',
  },

  /** Lookup column on child lists → ClientOnboarding */
  caseLookupField: 'CaseID',

  /** Set VITE_USE_SHAREPOINT_MOCK=true to use localStorage instead of SharePoint */
  useMock: import.meta.env.VITE_USE_SHAREPOINT_MOCK === 'true',
} as const;

export const graphScopes = [
  'User.Read',
  'User.Read.All',
  'Sites.ReadWrite.All',
  'Files.ReadWrite.All',
] as const;
