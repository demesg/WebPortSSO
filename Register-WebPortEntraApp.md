# Register or update a WebPort app in Entra

`Register-WebPortEntraApp.ps1` looks for apps with the same plant name in the signed-in tenant. If there is no match, it creates a new app and enterprise application. A name match is reused automatically if the app's client ID matches the SQLite value `KIClientId`; the object ID and client ID are kept. If the client ID does not match, `-Force` is required. The name comparison ignores case. The script also enables Microsoft Entra Application Proxy through an existing active connector group.

With several name matches, the app whose client ID unambiguously matches SQLite is chosen. If none matches, the error message shows the `DisplayName` and this WebPort instance's `KIClientId`. If the client ID is missing or invalid, the message says so explicitly.

## The database controls the server settings

The script reads `C:\ProgramData\WebPort\db\webport.sqlite` with the SQLite provider from `C:\Program Files\WebPort`. The two root folders can be set with `-WebPortDataPath` and `-WebPortProgPath`, and a separate database with `-DatabasePath`.

- **SettingsDeviceGuid is read from the database.** The namespace is identified by `SSLCN` and, if needed, `Plant`. There is no hard-coded GUID and no parameter for guessing it. Several equally likely namespaces cause an error before registration.
- **SSLCN is required.** A missing, empty or invalid SSLCN stops the script before the Graph sign-in or any stop of WebPort, with a prompt to run LetsEncryptWebport first. No other URL is used as a fallback.
- **Plant is the plant name**, the value under **General settings → Plant** in WebPort. It is used exactly as the name of the app registration and the enterprise application. If Plant is empty, the script asks for the name. `-Plant` can be given for unattended runs; `-FacilityName` and `-DisplayName` are aliases for the same parameter.
- **ServerUrl** becomes `https://<SSLCN>/` and is Application Proxy's internal address. **RedirectUri** is on the external `msappproxy.net` address, see [Application Proxy](#application-proxy). Old server URLs or custom metadata keys are not used to build the addresses.
- **KIExternalTenantId** is used if it holds a valid value. Otherwise the tenant ID is taken from the Graph sign-in. `-TenantId` sets a tenant explicitly. An existing tenant must match the sign-in context.
- The app name is searched for with **Plant**. The existing **KIClientId** is then used to bind the name match safely to this WebPort instance. The client ID of the created or updated app is returned and saved after the registration.

No `WebPortIAM.*` settings are read, created or changed. Old rows of that kind are left as they are. Apart from `EntraRedirectUri` and the encrypted `EntraClientSecretProtected`, the script introduces no new database settings: none for scopes, object IDs or plaintext secrets.

## Running

Run in **PowerShell 7 (`pwsh`)**. The script requires it (`#Requires -Version 7.0`), because WebPort's SQLite provider `Microsoft.Data.Sqlite` can only be loaded there. The Graph module is required for registration:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

Inspect the database values without Graph, without stopping WebPort and without writing:

```powershell
.\Register-WebPortEntraApp.ps1 -ReadDatabaseOnly
# Show the roles in the new manifest:
.\Register-WebPortEntraApp.ps1 -ReadDatabaseOnly | Select-Object -ExpandProperty AppRoles
```

Preview the action:

```powershell
.\Register-WebPortEntraApp.ps1 -WhatIf
```

Register and write the result back (the default):

```powershell
$registration = .\Register-WebPortEntraApp.ps1
$registration | Select-Object Plant,ClientId,TenantId,ServerUrl,RedirectUri,ApplicationProxyEnabled,ApplicationProxyConnectorGroupName,AdminConsentGranted,DatabaseUpdated,DatabaseBackupPath,DatabaseUpdateError
```

Run in an elevated window if WebPort is to be stopped and the local database updated. Sign in with an Entra role that may manage apps, Application Proxy and delegated admin consent, for example Application Administrator. Cloud Application Administrator is enough only with `-SkipApplicationProxy`, because that role cannot manage Application Proxy. `-UseDeviceCode` uses device code sign-in. `-UseExistingConnection` reuses an already connected delegated Graph session, after checking its tenant and permissions.

Register a new Entra app for real, but without writing to the database:

```powershell
$registration = .\Register-WebPortEntraApp.ps1 -UpdateDatabase:$false
```

Update an existing app with the same plant name, or create it if it is missing:

```powershell
$registration = .\Register-WebPortEntraApp.ps1 -Force
# Entra only, without writing to the local database or stopping WebPort:
$registration = .\Register-WebPortEntraApp.ps1 -Force -UpdateDatabase:$false
# Update roles and consent without another secret:
$registration = .\Register-WebPortEntraApp.ps1 -Force -SkipClientSecret -UpdateDatabase:$false
```

If the name match's client ID already matches SQLite, `-Force` is not needed. The run is then idempotent and creates no new client secret, and the result has `SecretStatus=ExistingSecretNotRetrievable`. `-Force` is needed when the operator explicitly wants to take over an unambiguous name match with another client ID, and then normally creates a new secret.

An existing app is reused only if it is a single-tenant Web app without public client fallback or implicit grant, with a known WebPort callback: this run's callback, the internal `https://<SSLCN>/access/loginsso`, or an `msappproxy.net` address ending in `/access/loginsso`. Incompatible apps are rejected before any change, with or without `-Force`. Existing redirect URIs are kept, and missing ones are added. A home page that someone changed manually is kept. The enterprise application and the tenant-wide Graph consent are reused if they exist, otherwise they are created. Existing API permissions, consent scopes and optional claims are kept, and whatever is missing is added.

`-ReadDatabaseOnly` and `-WhatIf` do not ask for a missing Plant value. `-ReadDatabaseOnly` then returns `PlantRequired=true`, and the plan shows the app name as `(saknas)`. They make no Graph calls and therefore cannot tell whether the app already exists; the plan shows that the action creates or updates.

### Application Proxy

Application Proxy is published with the internal URL `https://<SSLCN>/` and an external URL on the tenant's `msappproxy.net` domain: `https://<plant>-<tenant>.msappproxy.net/`.

- **`<plant>`** is the Plant value in lowercase. Å, ä and ö become a, a and o, and characters other than letters and digits become hyphens. Example: `Anläggning Väster's hus` becomes `anlaggning-vaster-s-hus`. The name is shortened if `<plant>-<tenant>` would exceed 63 characters.
- **`<tenant>`** is the prefix of the tenant's initial `onmicrosoft.com` domain. The script reads it from Graph after the sign-in.
- **Custom address:** `-ApplicationProxyExternalUrl 'https://webport-contoso.msappproxy.net/'` sets the address explicitly. It must be an HTTPS root address.

Other settings:

- Entra pre-authentication (`aadPreAuthentication`), secure HTTP-only cookies and validation of the backend certificate.
- The host header is **not** translated. WebPort therefore receives the external address, which lets the add-on tell the two addresses apart.
- The add-on's callback `https://<external address>/access/loginsso` is registered as a redirect URI and saved in SQLite as `EntraRedirectUri`. The home page is set to `https://<external address>/entra/login`, so a click on the app's tile in My Apps signs the user straight into WebPort.
- The external root address `https://<external address>/` is also registered as a redirect URI before the app is published, because Application Proxy's own pre-authentication returns there.
- Entra accepts Application Proxy publishing only on an app that has an application ID URI. If the app has none, `api://<client ID>` is set first, as in Microsoft's guide. An existing value is kept.
- If Microsoft Graph rejects a call, the script shows Graph's error code and message, not just which step failed.
- The internal callback `https://<SSLCN>/access/loginsso` is registered as well, so that users on the local network can sign in directly against the server.
- On an existing app, the external callback, the root address and the internal callback are added if they are missing, and earlier redirect URIs are kept. The home page is changed to `/entra/login` only if it is empty or still an address the script set earlier (the external root address or `WebPortServer Url`). A home page that someone changed manually is kept. An older app with the internal SSLCN callback is accepted.

**Without Application Proxy.** If Entra has no connector at all, the app is registered without Application Proxy. The script shows this as a white information line and continues. The callback is then `https://<SSLCN>/access/loginsso` and the home page `https://<SSLCN>/entra/login`, and `EntraRedirectUri` is cleared, so that the add-on derives the callback from `WebPortServer Url`. If `-ApplicationProxyExternalUrl` or `-ApplicationProxyConnectorGroup` is given, the run stops instead.

If connectors exist but none is active, the run stops instead, before anything changes. Otherwise a connector that is temporarily down would silently move the callback to the internal address, and external sign-in would stop working.

`-SkipApplicationProxy` skips Application Proxy entirely. The script then does not look for connectors and does not request `Directory.ReadWrite.All`. As without a connector, the callback is the internal one and `EntraRedirectUri` is cleared. An existing publication in Entra is not changed. The switch cannot be combined with `-ApplicationProxyExternalUrl` or `-ApplicationProxyConnectorGroup`.

```powershell
$registration = .\Register-WebPortEntraApp.ps1 -SkipApplicationProxy
```

With Application Proxy, users can sign in on both addresses. Each sign-in stays on the address where it started: the callback, the binding cookie and the session cookie belong to the same host, and a sign-in started on one address cannot be completed on the other. A sign-in started on any other address, such as an IP address, is redirected to the external address. `/entra/status` answers on every address.

The script reads the publication back and verifies it before consent or a client secret is created: that it is active, that the addresses, pre-authentication and host header setting are correct, and that the right connector group is assigned.

Without the parameter, the script first chooses an active connector group that contains a connector on the local computer, then the active default group or the only active group. If several groups remain, the run stops before the app is created. Then choose the group explicitly by name or object ID:

```powershell
$registration = .\Register-WebPortEntraApp.ps1 -ApplicationProxyConnectorGroup 'WebPort-proxy'
```

The group must already exist and contain at least one connector with the status `active`. The connector must be able to resolve and reach `https://<SSLCN>/` internally, for example through internal DNS. The script installs or moves no connectors. It uses Microsoft Graph's beta endpoint, because the Application Proxy configuration is only available there so far.

## Entra configuration

The enterprise application is configured with **Assignment required = Yes** (`appRoleAssignmentRequired=true`), both on a new registration and when an existing app is reused. The script verifies the setting by reading it back before consent and before creating the secret. Assign users or groups in the enterprise application so that they can sign in. The setting is managed on the service principal, as described in [Microsoft Graph's documentation](https://learn.microsoft.com/en-us/graph/api/serviceprincipal-update?view=graph-rest-1.0).

The app is single-tenant (`AzureADMyOrg`), platform **Web**, with public client fallback and implicit grant turned off. It gets the delegated Graph permissions `openid profile email User.Read` and tenant-wide admin consent (`AllPrincipals`). No application permissions are assigned. The ID token is configured with the optional claim `email`.

The manifest's `appRoles` automatically contains the following roles, with the same values as the mapping built into the add-on:

| App role value in the JWT | Global permission in WebPort |
|---|---|
| `WebPort.Deny` | `NOACCESS` (0) |
| `WebPort.View` | `VIEW` (1) |
| `WebPort.Basic` | `BASIC` (2) |
| `WebPort.Advanced` | `ADVANCED` (3) |
| `WebPort.System` | `SYSTEM` (4) |
| `WebPort.Admin` | `ADMIN` (5) |

Each new role gets a unique GUID, `isEnabled=true` and `allowedMemberTypes=["User"]`, which allows assignment to both users and groups. On an existing app, WebPort roles keep their GUIDs so that assignments stay valid; names, descriptions and enablement are updated. Missing WebPort roles are added. Other roles are kept, including their enablement and member types. Conflicts, such as a WebPort role for application permissions or a role value that differs only in letter case, are rejected before any change. The script then reads the roles back and verifies the definitions before any client secret is created.

The script does not assign roles to users or groups. Make the assignments under **Enterprise applications → the app → Users and groups**. Roles are used by default, and the mapping from app role to permission level is built into the add-on. The defined roles are included with `-WithoutGraphGroups` as well; app roles do not require `User.Read`.

`-WithoutGraphGroups` leaves out `User.Read`. `-Scopes` may specify `openid profile email`, optionally with `User.Read`. These scopes apply to the app registration and consent; the add-on always requests `openid profile email` at sign-in.

The administration session requests `Application.ReadWrite.All`, `DelegatedPermissionGrant.ReadWrite.All` and `Directory.ReadWrite.All`. The last one is required for Application Proxy, connector groups and the tenant's `onmicrosoft.com` name, and is not requested with `-SkipApplicationProxy`. These administration permissions are not assigned to the new WebPort app. The Graph sign-in uses a process-scoped context.

A new client secret, named `WebPort Entra integration`, is created last, after the app and the consent have been verified. It gets the longest allowed lifetime, 24 months. `-SecretLifetimeMonths` sets a shorter lifetime, 1–24 months, and the tenant's policy may limit it further. The expiry date is shown in yellow at the end, together with the command for renewing the secret (`.\Register-WebPortEntraApp.ps1 -Force`). Renew it before it expires, or Entra sign-in stops working. `-SkipClientSecret` leaves out the secret if client authentication is arranged separately, for example with a certificate. It does not turn the app into a public client.

A secret is created for a new app, and with `-Force` also for an existing one, unless `-SkipClientSecret` is used. Old secrets are kept in Entra. The script replaces the secret in SQLite and restarts WebPort, see below. Verify sign-in before removing old secrets in Entra.

## Writing back to SQLite

With the default behaviour (`-UpdateDatabase`), an elevated window is required. Unless `-SkipClientSecret` is given, WebPort must also either be installed as the `WebPortServer` service under `LocalSystem` (running or stopped), or run as a program (`WebPortServer.exe`) started by the user who runs the script. Both are checked before anything changes in Entra.

After a successful Entra registration, the script stops WebPort, the service or the program, from the selected installation folder. The SQLite backup API creates a consistent backup, including committed WAL data, next to the database as `webport.sqlite.EntraIAM-backup-<timestamp>.sqlite`. The following WebPort keys are written in one transaction:

| Database key | Value |
|---|---|
| `Plant` | The plant name (**General settings → Plant**), also the app name in Entra |
| `WebPortServer Url` | `https://<SSLCN>/` |
| `EntraRedirectUri` | `https://<external address>/access/loginsso`, or empty without Application Proxy |
| `AuthenticationTypeKI` | `1`, turns on WebPort's SSO sign-in (through `/access/loginsso`, which the add-on handles) |
| `KIClientId` | The client ID of the created or updated app |
| `KIExternalTenantId` | Tenant GUID |
| `KIProviderUri` | `https://login.microsoftonline.com/<tenant GUID>` |
| `KICustomPolicy`, `KISignInPolicyName` | Empty for Entra |

Each UPDATE or INSERT matches both the DeviceGuid that was read and the key. Other devices are not changed. SSLCN is checked again before writing; if the certificate name, the namespace or a value to be written has changed since it was read, the transaction is aborted. Users and groups are not changed, and neither is the SSLCN value itself.

If the database was updated and a new secret was created, the secret is then saved in `Settings.EntraClientSecretProtected`. The add-on decrypts it with DPAPI `CurrentUser` as the account WebPort runs as, bound to the tenant and client ID. An administrator cannot encrypt for LocalSystem directly, so the script works like this:

1. It creates a folder under `C:\ProgramData\WebPort\Entra\secret-staging`. Only SYSTEM and administrators have access to it.
2. The folder receives a transfer copy of the secret, encrypted with DPAPI `LocalMachine`, and a small worker script without the secret.
3. The worker runs as the account WebPort runs as: for the service, through a temporary scheduled task as SYSTEM; for a program, directly as the current user. It checks that it runs as that account, re-encrypts the secret bound to the tenant and client ID, checks that it can be decrypted and writes the value in a transaction.
4. The scheduled task, the transfer copy and the folder are removed immediately afterwards, even on failure.

The secret never appears in plaintext on disk, in the task's arguments or in the console.

If WebPort was running before the run and the add-on is already installed, WebPort is started again, as a service or a program as before. Otherwise WebPort stays stopped until `Install-EntraModule.ps1` is run.

If the database update fails after the Entra changes, the result and the secret are returned anyway. `DatabaseUpdated=false`, `DatabaseUpdateError` and a warning show the error. Save the result and inspect the app before running again. `-Force` reuses the app but creates another secret unless `-SkipClientSecret` is given. A failure during the Graph steps can leave partial changes; the script does not delete the app as a recovery measure.

## Returned values and the secret

The result contains `Plant`, `DisplayName`, `ServerUrl`, `RedirectUri`, `TenantId`, `ClientId`, `ApplicationObjectId`, `ServicePrincipalObjectId`, `Authority`, `Scopes`, `ConsentedScopes`, `AppRoles`, `ConsentGrantId`, `AdminConsentGranted`, `ApplicationProxyEnabled`, `ApplicationProxyInternalUrl`, `ApplicationProxyExternalUrl`, `ApplicationProxyConnectorGroupId`, `ApplicationProxyConnectorGroupName`, `ClientSecret`, `SecretKeyId`, `SecretExpiresUtc`, `SecretStatus`, `ApplicationCreated`, `ApplicationUpdated`, `DatabasePath`, `SettingsDeviceGuid`, `ProposedDatabaseSettings`, `DatabaseUpdated`, `DatabaseBackupPath`, `DatabaseUpdateError`, `SecretStored` and `SecretStoreError`.

`ClientSecret` is a **SecureString** and is never shown. In SQLite it is stored only encrypted for the account WebPort runs as (`SecretStored=true`). If storing fails, the secret remains in `ClientSecret` and the error in `SecretStoreError`. Then run the script again with `-Force` to create and store a new secret.

The secret cannot be retrieved from Entra later. If a separate administrative recovery copy is needed, the result can be saved on Windows with DPAPI for the same Windows user on the same computer:

```powershell
$registration | Export-Clixml -LiteralPath '.\webport-entra.registration.clixml'
```

If another tool explicitly needs plaintext (avoid logging and transcripts):

```powershell
$plainSecret = [System.Net.NetworkCredential]::new('', $registration.ClientSecret).Password
```

## Verification

The registration tests run with simulated Graph calls, a simulated service and program, and real, separate SQLite test databases. The secret step runs the real worker script, but as the current user instead of SYSTEM. The tests cover, among other things:

- encryption and binding of the secret, and removal of the task and transfer file;
- that elevation and a supported WebPort runtime are required before Entra is changed;
- safe reuse when the SQLite client ID matches;
- enabling and reading back Application Proxy and its connector group, including the application ID URI;
- roles, name conflicts, `-Force`, preserved role IDs and custom roles, and added permissions and claims;
- new secrets without removal of old ones;
- rejection of ambiguous or incompatible apps before any change.

The tests stop no real WebPort.

The registration script prepares Entra and the existing WebPort settings. Creating roles in the manifest does not give users access by itself and does not change their local permissions. The assignment is made in Entra. The WebPort account is created by the add-on at the first sign-in, with the user's UPN as its name.

Sources: [create an app](https://learn.microsoft.com/en-us/graph/api/application-post-applications?view=graph-rest-1.0), [configure Application Proxy](https://learn.microsoft.com/en-us/graph/application-proxy-configure-api), [connector groups](https://learn.microsoft.com/en-us/graph/api/resources/connectorgroup?view=graph-rest-beta), [appRole fields](https://learn.microsoft.com/en-us/graph/api/resources/approle?view=graph-rest-1.0), [delegated consent](https://learn.microsoft.com/en-us/graph/api/oauth2permissiongrant-post?view=graph-rest-1.0), [client secret](https://learn.microsoft.com/en-us/graph/api/application-addpassword?view=graph-rest-1.0).
