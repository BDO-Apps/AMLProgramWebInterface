import { LogLevel, type Configuration } from '@azure/msal-browser';

export type AppRole = 'Preparer' | 'Compliance' | 'EngagementPartner' | 'RiskPartner';
export type AppOffice = 'Zimbabwe' | 'Malawi';

export interface TestUserProfile {
  email: string;
  role: AppRole;
  office: AppOffice;
  displayName: string;
  /** Allows Compliance Walkthrough role/office switching during testing */
  canOverrideRole: boolean;
}

const tenantId = import.meta.env.VITE_AZURE_TENANT_ID ?? 'be3d2a5f-945c-469b-bb8a-2c50395c4601';
const clientId = import.meta.env.VITE_AZURE_CLIENT_ID ?? '232439e8-2b9d-46cc-9ef5-b8cce8521b3e';
/** Must match a redirect URI registered under Authentication → SPA in Azure (client 232439e8-…) */
const redirectUri =
  import.meta.env.VITE_REDIRECT_URI ??
  (typeof window !== 'undefined' ? window.location.origin : 'http://localhost:3000');

/** Test personas mapped to implementation-plan roles (Section 3) */
export const testUsers: TestUserProfile[] = [
  {
    email: 't.moyo@bdo.co.zw',
    role: 'Preparer',
    office: 'Zimbabwe',
    displayName: 'Tendai Moyo',
    canOverrideRole: true,
  },
  {
    email: 'c.phiri@bdo.co.mw',
    role: 'Preparer',
    office: 'Malawi',
    displayName: 'Chikondi Phiri',
    canOverrideRole: true,
  },
  {
    email: 'tmutasa@bdo.co.zw',
    role: 'Compliance',
    office: 'Zimbabwe',
    displayName: 'Tendai Mutasa',
    canOverrideRole: true,
  },
  {
    email: 'itservices@bdo.co.zw',
    role: 'Compliance',
    office: 'Zimbabwe',
    displayName: 'IT Services',
    canOverrideRole: true,
  },
  {
    email: 'cmariro@bdo.co.zw',
    role: 'EngagementPartner',
    office: 'Zimbabwe',
    displayName: 'Clement Mariro',
    canOverrideRole: true,
  },
  {
    email: 'amugumwa@bdo.co.zw',
    role: 'RiskPartner',
    office: 'Zimbabwe',
    displayName: 'Amugumwa',
    canOverrideRole: true,
  },
  {
    email: 'test.user@bdo.co.zw',
    role: 'Preparer',
    office: 'Zimbabwe',
    displayName: 'Test User',
    canOverrideRole: true,
  },
];

export const roleConfig = {
  /** Emails that may use the Compliance Walkthrough role switcher */
  authorizedEmails: testUsers.filter((u) => u.canOverrideRole).map((u) => u.email.toLowerCase()),
  emailToProfile: Object.fromEntries(
    testUsers.map((u) => [u.email.toLowerCase(), u])
  ) as Record<string, TestUserProfile>,
};

export const msalConfig: Configuration = {
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri,
    postLogoutRedirectUri: redirectUri,
    navigateToLoginRequestUrl: true,
  },
  cache: {
    cacheLocation: 'localStorage',
    storeAuthStateInCookie: true,
  },
  system: {
    loggerOptions: {
      loggerCallback: (level: LogLevel, message: string, containsPii: boolean) => {
        if (containsPii) return;
        switch (level) {
          case LogLevel.Error:
            console.error(message);
            return;
          case LogLevel.Info:
            console.info(message);
            return;
          case LogLevel.Verbose:
            console.debug(message);
            return;
          case LogLevel.Warning:
            console.warn(message);
            return;
        }
      },
    },
  },
};

export const loginRequest = {
  scopes: [
    'User.Read',
    'User.Read.All',
    'Sites.ReadWrite.All',
    'Files.ReadWrite.All',
  ],
};

export const graphConfig = {
  graphMeEndpoint: 'https://graph.microsoft.com/v1.0/me',
  graphMePhotoEndpoint: 'https://graph.microsoft.com/v1.0/me/photo/$value',
  graphUsersEndpoint: 'https://graph.microsoft.com/v1.0/users',
  graphGroupsEndpoint: 'https://graph.microsoft.com/v1.0/me/memberOf',
};
