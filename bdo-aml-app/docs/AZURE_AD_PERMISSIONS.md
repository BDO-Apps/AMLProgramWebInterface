# Azure AD permissions for BDO AML app (admin consent)

The SPA uses **delegated** Microsoft Graph permissions so each signed-in user accesses SharePoint as themselves.

| Permission | Purpose |
| :--- | :--- |
| `User.Read` | Profile + sign-in |
| `User.Read.All` | Read directory users (optional, already in app) |
| `Sites.ReadWrite.All` | Read/write SharePoint lists (`ClientOnboarding`, etc.) |
| `Files.ReadWrite.All` | Upload files to `AMLOnboardingDocuments` |

**App registration (defaults in this project)**

- **Application (client) ID:** `232439e8-2b9d-46cc-9ef5-b8cce8521b3e`
- **Directory (tenant) ID:** `be3d2a5f-945c-469b-bb8a-2c50395c4601`

---

## Step 1 — Open the app registration

1. Sign in to [Azure Portal](https://portal.azure.com) with an account that can manage app registrations (Global Admin, Application Admin, or Cloud Application Admin).
2. Go to **Microsoft Entra ID** (formerly Azure Active Directory).
3. Open **App registrations**.
4. Search for your app by name or paste the **Client ID** `232439e8-2b9d-46cc-9ef5-b8cce8521b3e`.
5. Open the registration.

---

## Step 2 — Register redirect URI (fixes `AADSTS500113`)

**Error `AADSTS500113: No reply address is registered`** means Azure has no redirect URI matching what the browser sends (e.g. `http://localhost:3000`).

1. Open app **`232439e8-2b9d-46cc-9ef5-b8cce8521b3e`** → **Authentication**.
2. Click **+ Add a platform** → **Single-page application** (not “Web”).
3. Under **Redirect URIs**, add every URL you use to open the app:
   - `http://localhost:3000` (default Vite port in this project)
   - If you use the machine name or IP, also add e.g. `http://127.0.0.1:3000`
4. Click **Configure** / **Save**.
5. Remove duplicate **Web** platform entries that use the same URL (SPA + Web with same redirect can cause issues).

The value must match **exactly** (scheme, host, port, no trailing path) — check DevTools → Network on login for `redirect_uri=` in the authorize request.

Optional: set `VITE_REDIRECT_URI` in `.env` to force a fixed URI if you always use one origin.

---

## Step 3 — Add API permissions

1. Go to **API permissions**.
2. Click **+ Add a permission**.
3. Choose **Microsoft Graph**.
4. Choose **Delegated permissions** (not Application).
5. Search and tick:
   - `User.Read`
   - `User.Read.All` (if not already added)
   - `Sites.ReadWrite.All`
   - `Files.ReadWrite.All`
6. Click **Add permissions**.

You should see four delegated Graph permissions on the **API permissions** blade.

---

## Step 4 — Grant admin consent

Delegated permissions that require admin approval show status **Not granted for …**.

1. On **API permissions**, click **Grant admin consent for &lt;your tenant name&gt;**.
2. Confirm **Yes**.
3. After success, the **Status** column shows a green tick: **Granted for …** on each permission.

If **Grant admin consent** is greyed out, your account lacks rights — ask a Global Administrator or Privileged Role Administrator to perform this step.

---

## Step 5 — Users sign out and sign in again

Tokens are cached in the browser. Old tokens do not include new scopes.

Each user (including you):

1. In the AML app, click **Logout**.
2. Close the browser tab (optional but helps).
3. Open the app again → **Sign in with Microsoft**.
4. On the consent screen (first time after new permissions), accept the permissions.

**IT tip:** Clear site data for `localhost:3000` if login still fails:

- Edge/Chrome → DevTools (F12) → Application → Storage → Clear site data  
  or remove MSAL keys from **Local Storage**.

---

## Step 6 — Verify permissions in the token (optional)

1. Sign in to the app.
2. Open browser **DevTools** → **Network**.
3. Find a request to `graph.microsoft.com`.
4. Decode the access token at [jwt.ms](https://jwt.ms) (paste token only in a trusted environment).
5. Check claim `scp` (scopes) includes:
   - `Sites.ReadWrite.All`
   - `Files.ReadWrite.All`
   - `User.Read`

If `scp` is missing SharePoint scopes, consent or sign-in was not completed correctly.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
| :--- | :--- | :--- |
| `AADSTS65001` / consent required | Admin consent not granted | Repeat Step 4 |
| `403` / `Authorization_RequestDenied` on Graph | Missing scope in token | Sign out/in; verify API permissions |
| `Grant admin consent` disabled | Insufficient admin role | Use Global Admin |
| Works for admin, not other users | User consent blocked by policy | Tenant: **Enterprise applications** → **User settings** → consent policies; or admin consent for all |
| `invalid_client` on login | Wrong client ID in `.env` | Match `VITE_AZURE_CLIENT_ID` to portal |
| `AADSTS500113` no reply address | No SPA redirect URI in Azure | Step 2 — add `http://localhost:3000` under **Single-page application** |

---

## Least-privilege alternative (advanced)

For production, some organisations replace `Sites.ReadWrite.All` with:

- `Sites.Selected` + granting the app access to one site via PowerShell

That requires extra setup beyond this dev guide. `Sites.ReadWrite.All` is the straightforward option for the Audit Software Dev site.

---

## Related app configuration

Scopes requested at login are defined in `src/config/msalConfig.ts`:

```typescript
export const loginRequest = {
  scopes: [
    'User.Read',
    'User.Read.All',
    'Sites.ReadWrite.All',
    'Files.ReadWrite.All',
  ],
};
```

No code change is needed after Azure Portal setup — only admin consent and a fresh sign-in.
