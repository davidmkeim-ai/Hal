# HAL Entra App Registration

This guide walks through the Microsoft Entra setup HAL needs for local development and future Teams use.

## 1. Create the app registration

In the Microsoft Entra admin center:

1. Go to `Microsoft Entra ID` -> `App registrations` -> `New registration`
2. Name: `HAL`
3. Supported account types:
   Use `Accounts in this organizational directory only` unless you specifically need cross-tenant access.
4. Redirect URI:
   Platform: `Web`
   URI: `http://localhost:4000/auth/microsoft/callback`

Save the registration.

Record these values from the Overview page:

- `Application (client) ID`
- `Directory (tenant) ID`

## 2. Create a client secret

1. Open `Certificates & secrets`
2. Create `New client secret`
3. Copy the `Value` immediately after it is created

Use that copied secret value in HAL. Do not use the secret ID.

## 3. Add Microsoft Graph delegated permissions

In `API permissions`, add delegated Microsoft Graph permissions:

- `User.Read`
- `Calendars.ReadWrite`
- `offline_access`
- `openid`
- `profile`

For the current HAL scaffold, delegated permissions are the right fit because HAL is acting on behalf of the signed-in user.

## 4. Optional Teams SSO preparation

If you want to move HAL from web redirect auth to full Teams tab SSO later:

1. Open `Expose an API`
2. Set the Application ID URI to something like:
   `api://localhost/your-client-id`
   or for hosted environments:
   `api://hal.yourdomain.com/your-client-id`
3. Add a scope for Teams tab SSO later

This step is mainly to support the future `webApplicationInfo.resource` value in the Teams manifest.

## 5. Values to copy into HAL

### `.env`

Copy [.env.example](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\.env.example) to `.env`, then fill in:

```env
PORT=4000
SESSION_SECRET=use-a-long-random-string-here
HAL_BASE_URL=http://localhost:4000
HAL_TAB_PATH=/index.html
ENTRA_CLIENT_ID=your-application-client-id
ENTRA_CLIENT_SECRET=your-client-secret-value
ENTRA_TENANT_ID=your-directory-tenant-id
GRAPH_SCOPES=User.Read Calendars.ReadWrite offline_access openid profile
```

### `hal.config.example.js`

Update [hal.config.example.js](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\hal.config.example.js):

```js
window.HAL_CONFIG = {
  appName: "HAL",
  teamsAppId: "your-teams-app-id",
  entraClientId: "same-as-application-client-id",
  entraTenantId: "same-as-directory-tenant-id",
  tabBaseUrl: "http://localhost:4000",
  apiBaseUrl: "http://localhost:4000",
  graphScopes: [
    "User.Read",
    "Calendars.ReadWrite",
  ],
};
```

### `appPackage/manifest.json`

Update [appPackage/manifest.json](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\appPackage\manifest.json):

- `id`
  Use a dedicated Teams app ID GUID. This can be different from the Entra client ID.
- `developer.websiteUrl`
  `http://localhost:4000` for local development is not valid for Teams sideloading. For Teams, use a real HTTPS URL.
- `staticTabs[0].contentUrl`
  `https://your-hosted-domain/index.html`
- `staticTabs[0].websiteUrl`
  `https://your-hosted-domain/index.html`
- `validDomains`
  Your hosted domain only, no protocol
- `webApplicationInfo.id`
  Your Entra `Application (client) ID`
- `webApplicationInfo.resource`
  The Application ID URI from `Expose an API`

## 6. Important local-vs-Teams distinction

Local web testing:

- `http://localhost:4000/auth/microsoft/callback` is valid for the current web auth flow.

Actual Teams tab testing:

- Teams tabs must load from HTTPS.
- Your manifest URLs and valid domains must point to the hosted HTTPS version of HAL, not localhost.

## 7. Suggested order

1. Finish the Entra app registration
2. Fill in `.env`
3. Run HAL locally and test Microsoft sign-in in the browser
4. Host HAL on HTTPS
5. Update the Teams manifest with hosted values
6. Sideload into Teams
