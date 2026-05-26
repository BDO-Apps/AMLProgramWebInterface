import { PublicClientApplication } from '@azure/msal-browser';
import { msalConfig } from '../config/msalConfig';

/** Single MSAL client instance shared by the service and MsalProvider */
export const msalInstance = new PublicClientApplication(msalConfig);
