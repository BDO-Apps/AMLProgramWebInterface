/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_AZURE_CLIENT_ID?: string;
  readonly VITE_AZURE_TENANT_ID?: string;
  readonly VITE_REDIRECT_URI?: string;
  readonly VITE_SP_HOSTNAME?: string;
  readonly VITE_SP_SITE_PATH?: string;
  readonly VITE_SHAREPOINT_SITE_URL?: string;
  readonly VITE_SP_LIST_CLIENT_ONBOARDING?: string;
  readonly VITE_SP_LIST_DIRECTORS?: string;
  readonly VITE_SP_LIST_BENEFICIAL_OWNERS?: string;
  readonly VITE_SP_LIST_AUDIT_LOGS?: string;
  readonly VITE_SP_LIBRARY_AML_DOCUMENTS?: string;
  readonly VITE_USE_SHAREPOINT_MOCK?: string;
  readonly VITE_POWER_AUTOMATE_WORKFLOW_URL?: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
