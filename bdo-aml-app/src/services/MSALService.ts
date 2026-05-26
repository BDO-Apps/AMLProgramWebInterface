import { InteractionRequiredAuthError } from '@azure/msal-browser';
import type { AccountInfo, AuthenticationResult } from '@azure/msal-browser';
import { msalInstance } from '../auth/msalInstance';
import {
  loginRequest,
  graphConfig,
  roleConfig,
  type AppOffice,
  type AppRole,
} from '../config/msalConfig';

export type { AppRole, AppOffice };

export interface M365User {
  name: string;
  email: string;
  username: string;
  role: AppRole;
  office: AppOffice;
  avatar?: string;
  accessToken?: string;
  /** Test accounts may switch role via Compliance Walkthrough bar */
  canOverrideRole?: boolean;
}

const ROLE_OVERRIDE_KEY = 'bdo_aml_role_override';

class MSALService {
  private currentUser: M365User | null = null;
  private listeners: ((user: M365User | null) => void)[] = [];
  private initialized = false;
  private initPromise: Promise<void> | null = null;

  public subscribe(callback: (user: M365User | null) => void) {
    this.listeners.push(callback);
    callback(this.currentUser);
    return () => {
      this.listeners = this.listeners.filter((l) => l !== callback);
    };
  }

  private notify() {
    this.listeners.forEach((callback) => callback(this.currentUser));
  }

  /** Call once at app startup — restores SSO session from MSAL cache */
  public initialize(): Promise<void> {
    if (!this.initPromise) {
      this.initPromise = this.doInitialize();
    }
    return this.initPromise;
  }

  private async doInitialize(): Promise<void> {
    if (this.initialized) return;

    await msalInstance.initialize();

    const redirectResult = await msalInstance.handleRedirectPromise();
    if (redirectResult?.account) {
      msalInstance.setActiveAccount(redirectResult.account);
      await this.syncUserFromAccount(redirectResult.account, redirectResult.accessToken);
    } else {
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        msalInstance.setActiveAccount(accounts[0]);
        await this.syncUserFromAccount(accounts[0]);
      }
    }

    this.initialized = true;
    this.notify();
  }

  private resolveAccountEmail(account: AccountInfo): string {
    const claims = account.idTokenClaims;
    const claimEmail =
      claims?.preferred_username ??
      (typeof claims?.email === 'string' ? claims.email : undefined) ??
      claims?.upn ??
      (Array.isArray(claims?.emails) ? claims.emails[0] : undefined);

    return (account.username || claimEmail || account.name || '').toLowerCase();
  }

  private mapRoleFromClaims(
    email: string,
    claims: AccountInfo['idTokenClaims']
  ): AppRole {
    const profile = roleConfig.emailToProfile[email];
    if (profile) {
      return profile.role;
    }

    const roleClaim = claims?.roles ?? claims?.role;
    if (Array.isArray(roleClaim)) {
      if (roleClaim.includes('RiskPartner')) return 'RiskPartner';
      if (roleClaim.includes('EngagementPartner')) return 'EngagementPartner';
      if (roleClaim.includes('Compliance')) return 'Compliance';
      if (roleClaim.includes('Preparer')) return 'Preparer';
    } else if (typeof roleClaim === 'string') {
      if (roleClaim.includes('RiskPartner')) return 'RiskPartner';
      if (roleClaim.includes('EngagementPartner')) return 'EngagementPartner';
      if (roleClaim.includes('Compliance')) return 'Compliance';
      if (roleClaim.includes('Preparer')) return 'Preparer';
    }

    return 'Preparer';
  }

  private resolveOffice(email: string): AppOffice {
    const profile = roleConfig.emailToProfile[email];
    if (profile) return profile.office;
    return email.endsWith('.mw') ? 'Malawi' : 'Zimbabwe';
  }

  private applyWalkthroughOverride(user: M365User): M365User {
    if (!user.canOverrideRole) return user;

    const saved = sessionStorage.getItem(ROLE_OVERRIDE_KEY);
    if (!saved) return user;

    try {
      const { role, office } = JSON.parse(saved) as { role: AppRole; office: AppOffice };
      return { ...user, role, office };
    } catch {
      return user;
    }
  }

  private mapAccountToUser(account: AccountInfo, accessToken?: string): M365User {
    const email = this.resolveAccountEmail(account);
    const profile = roleConfig.emailToProfile[email];
    const role = this.mapRoleFromClaims(email, account.idTokenClaims);
    const office = this.resolveOffice(email);
    const canOverrideRole = roleConfig.authorizedEmails.includes(email);

    const user: M365User = {
      name: profile?.displayName ?? account.name ?? email,
      email,
      username: email,
      role,
      office,
      accessToken,
      canOverrideRole,
    };

    return this.applyWalkthroughOverride(user);
  }

  private async fetchProfilePhoto(accessToken: string): Promise<string | undefined> {
    try {
      const response = await fetch(graphConfig.graphMePhotoEndpoint, {
        headers: { Authorization: `Bearer ${accessToken}` },
      });
      if (!response.ok) return undefined;
      const blob = await response.blob();
      return URL.createObjectURL(blob);
    } catch {
      return undefined;
    }
  }

  private async syncUserFromAccount(
    account: AccountInfo,
    accessToken?: string
  ): Promise<M365User> {
    let token = accessToken;
    if (!token) {
      try {
        token = await this.acquireTokenForAccount(account);
      } catch {
        token = undefined;
      }
    }

    const user = this.mapAccountToUser(account, token);
    if (token) {
      user.avatar = await this.fetchProfilePhoto(token);
    }

    this.currentUser = user;
    return user;
  }

  private async acquireTokenForAccount(account: AccountInfo): Promise<string> {
    const authResult = await msalInstance.acquireTokenSilent({
      ...loginRequest,
      account,
    });
    return authResult.accessToken;
  }

  public async login(): Promise<M365User> {
    await this.initialize();

    const loginResponse: AuthenticationResult = await msalInstance.loginPopup({
      scopes: loginRequest.scopes,
      prompt: 'select_account',
    });

    const account = loginResponse.account ?? msalInstance.getActiveAccount();
    if (!account) {
      throw new Error('Unable to determine signed-in Microsoft account.');
    }

    msalInstance.setActiveAccount(account);
    sessionStorage.removeItem(ROLE_OVERRIDE_KEY);

    const user = await this.syncUserFromAccount(account, loginResponse.accessToken);
    this.notify();
    return user;
  }

  public async loginRedirect(): Promise<void> {
    await this.initialize();
    await msalInstance.loginRedirect({ scopes: loginRequest.scopes, prompt: 'select_account' });
  }

  public async logout(): Promise<void> {
    sessionStorage.removeItem(ROLE_OVERRIDE_KEY);
    this.currentUser = null;
    this.notify();

    const account = msalInstance.getActiveAccount();
    if (account) {
      await msalInstance.logoutPopup({
        account,
        postLogoutRedirectUri: msalInstance.getConfiguration().auth.postLogoutRedirectUri,
      });
    }
  }

  public changeRole(role: AppRole, office: AppOffice = 'Zimbabwe'): M365User | null {
    if (!this.currentUser?.canOverrideRole) return this.currentUser;

    sessionStorage.setItem(ROLE_OVERRIDE_KEY, JSON.stringify({ role, office }));
    this.currentUser = { ...this.currentUser, role, office };
    this.notify();
    return this.currentUser;
  }

  public async getAccessToken(): Promise<string> {
    await this.initialize();

    const account = msalInstance.getActiveAccount();
    if (!account) {
      throw new Error('No active Microsoft account. Sign in again.');
    }

    try {
      const authResult = await msalInstance.acquireTokenSilent({ ...loginRequest, account });
      if (this.currentUser) {
        this.currentUser = { ...this.currentUser, accessToken: authResult.accessToken };
      }
      return authResult.accessToken;
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        const result = await msalInstance.acquireTokenPopup({ ...loginRequest });
        if (this.currentUser) {
          this.currentUser = { ...this.currentUser, accessToken: result.accessToken };
        }
        return result.accessToken;
      }
      throw error;
    }
  }

  public getCurrentUser(): M365User | null {
    return this.currentUser;
  }

  public isReady(): boolean {
    return this.initialized;
  }
}

export const msalService = new MSALService();
