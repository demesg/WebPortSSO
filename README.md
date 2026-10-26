# WebPort: Entra-SSO

Users sign in to WebPort with their Microsoft 365 work account. The add-on makes WebPort's **Login with Microsoft 365** button sign users in directly against your own Entra tenant, without the need for Kiona Identity.

- **Single sign-on** with the account users already have, and no separate WebPort passwords.
- **Who can sign in** is managed in Entra, by assigning users and groups to the app.
- **Permissions** come from app roles in Entra.
- **Remote access** through Microsoft Entra Application Proxy, set up by the registration script.

## How sign-in works

### How the add-on attaches to WebPort

```
 WebPortServer starts
   │
   ├─► WebPort loads every DLL in its folder and finds ExtEntraPkce.dll (implements IAddon)
   │     └─► EntraAddon.Initialize()
   │           1. Check WebPort's internals and the SQL statements (else: not started, event 3000)
   │           2. Read the settings from SQLite
   │           3. Put EntraHttpModule FIRST in WebPort's HTTP module list
   │
   └─► Every HTTPS request
         │
         ▼
   ┌──────────────────┐  /entra/login, /access/loginsso, /entra/status
   │ EntraHttpModule  │─────────────────────────────────────────────────► handled by the add-on
   └──────────────────┘
         │ all other paths
         ▼
   WebPort's own modules and controllers (unchanged)
```

### Authorization code flow with PKCE

```
   Browser                    WebPort + Entra add-on            Microsoft Entra ID
     │                                    │                                      │
     │ 1. GET /entra/login?return=/page   │                                      │
     │───────────────────────────────────►│                                      │
     │                                    │ create in memory, valid 5 min:       │
     │                                    │   state, nonce, browser binding,     │
     │                                    │   code_verifier (random secret)      │
     │                                    │   code_challenge=SHA256(verifier)    │
     │ 2. 303 redirect to Entra           │                                      │
     │    + cookie __Host-wpentra-…       │                                      │
     │      (binding; Secure, HttpOnly)   │                                      │
     │◄───────────────────────────────────│                                      │
     │                                    │                                      │
     │ 3. GET /<tenant>/oauth2/v2.0/authorize                                    │
     │      client_id, redirect_uri=https://<server>/access/loginsso,            │
     │      scope=openid profile email, response_mode=form_post,                 │
     │      state, nonce, code_challenge, code_challenge_method=S256             │
     │──────────────────────────────────────────────────────────────────────────►│
     │                                    │      sign-in, MFA, Conditional Access│
     │                                    │      "Assignment required" check     │
     │ 4. page that auto-submits code + state to redirect_uri                    │
     │◄──────────────────────────────────────────────────────────────────────────│
     │                                    │                                      │
     │ 5. POST /access/loginsso           │                                      │
     │      code, state + binding cookie  │                                      │
     │───────────────────────────────────►│                                      │
     │                                    │ state known, unused, not expired?    │
     │                                    │ binding cookie matches this state?   │
     │                                    │ client secret ◄── SQLite (DPAPI)     │
     │                                    │                                      │
     │                                    │ 6. POST /oauth2/v2.0/token           │
     │                                    │    code, code_verifier,              │
     │                                    │    client_id, client_secret          │
     │                                    │─────────────────────────────────────►│
     │                                    │      SHA256(code_verifier) must equal│
     │                                    │      code_challenge from step 3      │
     │                                    │ 7. id_token (signed JWT)             │
     │                                    │◄─────────────────────────────────────│
     │                                    │                                      │
     │                                    │ validate: signature (Entra keys),    │
     │                                    │   issuer, audience, lifetime,        │
     │                                    │   nonce, tenant (tid)                │
     │                                    │ roles ──► built-in role mapping      │
     │                                    │   ──► global level (no role: denied) │
     │                                    │ account named by UPN (lowercase)?    │
     │                                    │   yes: use it, sync its role level   │
     │                                    │   no:  JIT, create it with the       │
     │                                    │        role level                    │
     │                                    │ local checks: locked, expired,       │
     │                                    │   throttled, obligatory group        │
     │                                    │ new WebPort session + token,         │
     │                                    │   old session removed                │
     │ 8. 303 redirect to /page           │                                      │
     │    + new WebPort session cookie    │                                      │
     │◄───────────────────────────────────│                                      │
```

- **PKCE:** the `code_verifier` never leaves the server. An intercepted authorization code cannot be redeemed without it, and redeeming also requires the client secret.
- **state, nonce and the binding cookie** tie the callback to the browser that started the login. Each value can be used once and expires after five minutes.
- **The add-on does not use the access token.** Identity comes only from the ID token, validated with Microsoft's IdentityModel libraries: RS256 signature, issuer, audience, lifetime, nonce, token version and tenant.
- **The WebPort account is the one named by the user's UPN** (`preferred_username`, lowercase). If it exists, the add-on uses it. If not, the add-on creates it just in time (always on). The new account gets the global level of the user's Entra role, plus the `name` and `email` claims. If WebPort has an `ObligatoryGroup`, the account is added to it. Password and PIN are random and never disclosed, so Entra is the only way to sign in. The built-in `admin` account is never created or used. Without a mapped Entra role, no account is created.
- **Role sync:** in role mode, `Users.global` is set to the current Entra role on every sign-in, for every account that signs in through Entra.
- `GET /access/loginsso` without parameters starts the same flow as `/entra/login`.

### What the add-on stores

The add-on has **no tables of its own**. It uses WebPort's `Users` table: the account whose `username` is the user's UPN. The schema of `Users` is not changed, and `Users.kid` is not used, because it belongs to WebPort's Kiona sign-in and holds Kiona's identity, not the Entra object ID. Settings and the encrypted client secret are in WebPort's `Settings` table. The mapping from app role to global permission is built into the add-on, see [Assign users in Entra](#4-assign-users-in-entra). Sign-in events go to the Windows Application log, see [Verify](#6-verify).

Because the account is found by name, two things are the administrator's job:

- **A user's UPN changes** (for example a new surname or domain): rename the WebPort account to the new UPN before the user signs in again. Otherwise the user gets a new, empty account.
- **A user leaves:** lock or delete the WebPort account. Entra may give the same UPN to a new employee later, who would otherwise sign in to the old account and inherit its groups, settings and history. A locked account is denied; a deleted account is created anew on the next sign-in.

## Contents

The repository is the server package.

| File | Purpose |
|---|---|
| `Register-WebPortEntraApp.ps1` | **Everything Entra:** creates or updates the app registration and writes its settings to SQLite. Also stores the client secret in SQLite, encrypted for the account WebPort runs as. Self-contained |
| `Install-EntraModule.ps1` | **Everything add-on:** installs or upgrades `ExtEntraPkce.dll` and registers the event log source. Also performs rollback and uninstallation. Self-contained |
| `ExtEntraPkce.dll` | The compiled and tested add-on |

Detailed documentation: [Entra app registration](Register-WebPortEntraApp.md).

## Requirements

**Server**

- WebPort installed in `C:\Program Files\WebPort` with data in `C:\ProgramData\WebPort`. The add-on's tests have passed against WebPort 2.25.10 and 2.25.4. See [WebPort versions](#webport-versions) for how one DLL works with several versions.
- WebPort running either as the `WebPortServer` Windows service under **LocalSystem**, or as a program (`WebPortServer.exe` started by a user). See [Service or program](#service-or-program).
- A valid HTTPS certificate configured with LetsEncryptWebport, so that the `SSLCN` setting holds the server's DNS name. This address is the internal URL that Application Proxy publishes.
- An Entra Application Proxy connector group with at least one active connector that can reach `https://<SSLCN>/`.
- **PowerShell 7**.
- Local administrator rights for installation and maintenance.

**Entra ID**

- The `Microsoft.Graph.Authentication` PowerShell module:

  ```powershell
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
  ```

- An account with a role that can register applications, manage Application Proxy and grant tenant-wide delegated consent, for example Application Administrator. Cloud Application Administrator is enough only with `-SkipApplicationProxy`.

## Installation

A WebPort server needs only three files, in one folder:

```
Register-WebPortEntraApp.ps1
Install-EntraModule.ps1
ExtEntraPkce.dll
```

Run every step in an **elevated PowerShell 7** session.

Both scripts start by printing what they are going to change:

- the settings of the app in Entra (`Register-WebPortEntraApp.ps1`);
- each SQLite setting, with its current and new value;
- the files and the `WebPortServer.exe.config` entries (`Install-EntraModule.ps1`);
- where the backup is saved;
- whether WebPort is stopped and started again, as a service or as a program.

The scripts do not wait for approval: the changes follow directly. To review first, run with `-WhatIf`, which prints the list and stops without changing anything. Or add `-Confirm`, and the script asks before it makes any change.

### 1. Get the server package

Download the three files to one folder on the server, for example `C:\script`:

```powershell
New-Item -ItemType Directory -Path C:\script -Force | Out-Null
cd C:\script
$base = 'https://raw.githubusercontent.com/demesg/WebPortSSO/main'
foreach ($file in 'Register-WebPortEntraApp.ps1','Install-EntraModule.ps1','ExtEntraPkce.dll') {
    Invoke-WebRequest -Uri "$base/$file" -OutFile $file
}
Unblock-File -Path .\*.ps1, .\ExtEntraPkce.dll
```

`Unblock-File` removes the "downloaded from the internet" mark, so that PowerShell's execution policy does not stop the scripts. Run the remaining steps from the same folder.

### 2. Register the Entra application

Preview the settings read from the WebPort database first:

```powershell
.\Register-WebPortEntraApp.ps1 -ReadDatabaseOnly
```

Then register:

```powershell
$registration = .\Register-WebPortEntraApp.ps1
$registration | Select-Object Plant,ClientId,TenantId,RedirectUri,SecretExpiresUtc,DatabaseUpdated,SecretStored
```

The script signs in to Microsoft Graph and creates a single-tenant Web app named after the plant name in WebPort, **General settings → Plant**. It also creates the six `WebPort.*` app roles, sets *Assignment required* on the enterprise application, and enables Microsoft Entra Application Proxy through an existing active connector group. It grants admin consent for `openid profile email User.Read` and creates a client secret with the maximum lifetime of 24 months. The expiry date is shown in yellow at the end, with the command for renewing the secret. Use `-ApplicationProxyConnectorGroup <name-or-id>` when the tenant has several active connector groups and automatic selection is ambiguous.

**Addresses.** Users reach WebPort through Application Proxy on `https://<plant>-<tenant>.msappproxy.net/`:

- `<plant>` is **General settings → Plant** in lowercase ASCII.
- `<tenant>` is the prefix of the tenant's initial `onmicrosoft.com` domain, read from Graph after sign-in.
- `-ApplicationProxyExternalUrl` sets another address.

The add-on's callback is `<external address>/access/loginsso`. It is added to the app and saved in SQLite as `EntraRedirectUri`. The external root address is also registered as a redirect URI before the app is published, because Application Proxy's own pre-authentication returns there. Entra accepts Application Proxy publishing only on an app with an application ID URI, so the script sets `api://<client ID>` first if the app has none, as in Microsoft's guide. An existing application ID URI is kept. The internal callback `https://<SSLCN>/access/loginsso` is registered as well.

The app's home page is `<address>/entra/login`, so the app's tile in My Apps signs the user straight into WebPort. On an existing app the script replaces the home page only if it is empty or still an address the script set earlier (the external root or `WebPortServer Url`). A home page that someone changed manually stays as it is.

Users can sign in on both addresses: externally through Application Proxy, or directly on `https://<SSLCN>/` from the local network. Each sign-in stays on the address where it started. The callback, the binding cookie and the session cookie all belong to that host, and a sign-in started on one address cannot be completed on the other. Application Proxy passes the external host name through to WebPort unchanged, so that the add-on can tell the two addresses apart. A sign-in started on any other address, such as an IP address, is redirected to the external address.

**Without Application Proxy.** If the tenant has no Application Proxy connector at all, the script prints a white information line and registers the app without it. Users then sign in on `https://<SSLCN>/`. `-SkipApplicationProxy` skips Application Proxy explicitly. If connectors exist but none is active, the script stops instead, so that a temporary outage never moves the callback.

Before anything changes in Entra, the script checks that it runs as administrator and that WebPort runs as the LocalSystem service or as a program of the current user, see [Service or program](#service-or-program). After a successful registration it:

1. Stops WebPort, backs up the database and writes `Plant`, `KIClientId`, `KIExternalTenantId`, `KIProviderUri`, `WebPortServer Url`, `EntraRedirectUri` and `AuthenticationTypeKI=1`. The last one turns on WebPort's single sign-on: the SSO option on WebPort's sign-in page goes to `/access/loginsso`, which the add-on handles. `Plant` is the value of **General settings → Plant** in WebPort; if it is empty, the script asks for it and saves the answer there. The same name is used for the app in Entra.
2. Stores the new client secret in `Settings.EntraClientSecretProtected`. It is encrypted with DPAPI for the account WebPort runs as, bound to the tenant and client ID. For the service, a short-lived scheduled task running as SYSTEM does the encryption, and the task and its temporary transfer file are removed afterwards. The secret never appears in plaintext on disk, in task arguments or in the console.
3. Starts WebPort again if it was running and the add-on is already installed. Otherwise WebPort stays stopped until step 3.

An existing app with the same name is reused automatically when its client ID matches SQLite `KIClientId`. Use `-Force` only to update an otherwise unambiguous name match with a different client ID. See the [registration documentation](Register-WebPortEntraApp.md) for all options.

### 3. Install the add-on

```powershell
.\Install-EntraModule.ps1
```

The installer reads everything it needs from SQLite. It stops if the Entra settings or the encrypted secret are missing. It then:

1. Backs up `WebPortServer.exe.config`, the current DLL and the database to `C:\ProgramData\WebPort\Entra\install-backups\<timestamp>`.
2. Copies `ExtEntraPkce.dll` to `C:\Program Files\WebPort\addons\EntraPkce\`.
3. Creates no tables and writes no settings; the registration has already written them. Entra app roles need no setting. Only local permissions do, see [Assign users in Entra](#4-assign-users-in-entra). If the row `EntraRolesEnabled=1` exists, the installer removes it, because roles are already the default.
4. Registers the event log source `WebPort Entra` in the Windows Application log.
5. Compares the WebPort libraries the DLL was built against with the installed ones, and adds a `bindingRedirect` to `WebPortServer.exe.config` for each one that differs. See [WebPort versions](#webport-versions).
6. Removes legacy keys from `WebPortServer.exe.config`, if an older installation left any.
7. Starts WebPort and waits for `https://<SSLCN>/entra/status` (the internal address) to report the add-on as active, with the version of the DLL it just copied. The add-on only becomes active if the WebPort compatibility check passes.

If any step fails, the DLL and `WebPortServer.exe.config` are restored automatically. The same command upgrades an installed add-on.

### 4. Assign users in Entra

In the Entra admin center, open **Enterprise applications → \<Plant\> → Users and groups**, where \<Plant\> is the name under **General settings → Plant** in WebPort. Assign users or groups to one of the app roles. The mapping is built into the add-on and cannot be changed per server:

| App role | WebPort global permission |
|---|---|
| `WebPort.Deny` | `NOACCESS`; denies access even if other roles grant it |
| `WebPort.View` | `VIEW` |
| `WebPort.Basic` | `BASIC` |
| `WebPort.Advanced` | `ADVANCED` |
| `WebPort.System` | `SYSTEM` |
| `WebPort.Admin` | `ADMIN` |

Users without an assigned role cannot sign in. The rules:

- Only app roles in the ID token's `roles` claim count, not tenant administrator roles or the `wids` claim. Roles are read only after the token has been fully validated.
- Role values must match exactly, including case.
- With several roles, the highest permission applies, but `WebPort.Deny` always denies.
- Unknown roles give no permission, even if someone added them manually in Entra. A user with only unknown roles is denied.
- Local group, page and object permissions still apply; the role sets only the global level. Locks, validity period, sign-in restrictions and the obligatory group can still deny.
- Roles are evaluated at sign-in. A changed assignment in Entra takes effect at the user's next sign-in, not in a session that is already open.

Roles are the default and need no setting. Only if a server must use local WebPort permissions instead, add the row `EntraRolesEnabled` with the value `0` to WebPort's `Settings` table, under the same `DeviceGuid` as the add-on's other settings, and restart WebPort. Removing the row turns roles on again. In local mode, new accounts are created with `NOACCESS`, and an administrator grants permissions in WebPort.

### 5. First sign-in creates the WebPort account

Users sign in at `https://<plant>-<tenant>.msappproxy.net/entra/login`. No preparation in WebPort is needed.

- **First sign-in:** if no local account is named by the user's UPN (lowercase), the add-on creates one with the global permission level of the assigned app role. The event log shows `user_created` (event 1003).
- **Every sign-in:** the account named by the UPN is used, and in role mode its stored global level is set to the current role.

An existing local account is used only when its name equals the user's UPN. An account with a different name, such as `anna` instead of `anna.svensson@example.com`, is not used; the user gets a new account. To keep the old account, rename it to the UPN before the user's first sign-in. The same applies when a user's UPN changes later. When a user leaves, lock or delete the account, see [What the add-on stores](#what-the-add-on-stores).

### 6. Verify

Sign in at `https://<plant>-<tenant>.msappproxy.net/entra/login` with an account that has an app role. The registration summary prints the exact address.

Every sign-in is logged in the Windows Application log under the source `WebPort Entra`:

```powershell
Get-WinEvent -FilterHashtable @{LogName='Application';ProviderName='WebPort Entra'} -MaxEvents 20
```

| Event ID | Level | Meaning |
|---|---|---|
| 1000 | Information | Add-on started (`module_initialized`, with version) |
| 1001 | Information | Sign-in started (`authorize_started`) |
| 1002 | Information | Sign-in succeeded (`login_succeeded`, with user, object ID and permission level) |
| 1003 | Information | Account created (`user_created`) |
| 2000 | Warning | Sign-in failed; the first line is the reason, for example `no_mapped_role` or `token_http_401` with an `AADSTS` code |
| 3000 | Error | The add-on could not start (`module_initialization_failed`) |

The first line of each message is the result code. No secrets, tokens or authorization codes are logged. `AADSTS7000215` means that Entra rejected the client secret; check that the client ID and the secret belong together, and renew the secret if needed, see [Client secret renewal](#client-secret-renewal).

## Upgrading the add-on

Copy the new `ExtEntraPkce.dll` next to `Install-EntraModule.ps1` on the server and run:

```powershell
.\Install-EntraModule.ps1
```

## Service or program

WebPort can run as the Windows service `WebPortServer` or as a program, `WebPortServer.exe` started by a user. Both scripts detect which, and act accordingly:

| | Service | Program |
|---|---|---|
| Detected by | The service `WebPortServer` | A running `WebPortServer.exe` from the program folder, and its owner |
| Stopped by | Stopping the service | Closing the program; if it does not exit within 15 seconds, it is ended |
| Started again by | Starting the service | Starting `WebPortServer.exe` with the same arguments, if it belongs to the account running the script. Otherwise the script asks you to start it and waits up to five minutes |
| Client secret encrypted for | LocalSystem, through a short-lived scheduled task running as SYSTEM | The account running the program; run `Register-WebPortEntraApp.ps1` as that account |

**Why the account matters.** The client secret is the same in Entra in both cases. The copy in SQLite (`EntraClientSecretProtected`) is encrypted with Windows DPAPI for one account, and only that account can decrypt it. The add-on runs inside the WebPort process, so it can only read the secret if it was encrypted for the account WebPort runs as.

- `Register-WebPortEntraApp.ps1` stops before any change in Entra if it cannot encrypt for that account: a service running as another account than LocalSystem, a program owned by another user, or WebPort neither installed as a service nor running.
- `Install-EntraModule.ps1` tries to decrypt the secret when WebPort runs as a program of the account running the installer, and stops if it cannot.

**Switching between service and program** means a new account, so the stored secret no longer works. Run `.\Register-WebPortEntraApp.ps1 -Force` once in the new mode. It creates a new secret in Entra, because Entra never returns an existing one, and stores it for the new account. Then remove the old secret in Entra.

If WebPort is not running and has no service, the installer installs the add-on but cannot start WebPort or run the health check. Start WebPort yourself and check `/entra/status`.

## WebPort versions

### Why a DLL depends on the WebPort version

At startup, WebPort loads every DLL in its program folder, subfolders included, and starts each one that implements its `IAddon` interface. No setting is involved: placing `ExtEntraPkce.dll` in `addons\EntraPkce\` is enough.

The add-on uses WebPort's own libraries, `WebPortCommon.dll`, `HttpServer.dll` and `HttpServer.MVC.dll`. They are *strong-named*, so the DLL records the exact version it was compiled against, and .NET refuses to load it if another version is installed. This happens before any add-on code runs, so nothing is logged and `/entra/status` does not answer.

### How the installer handles it

The installer adds a **`bindingRedirect`** to `WebPortServer.exe.config`, because `ExtEntraPkce.dll` is compiled against the WebPort version it was built with, not the one that is installed. The redirect tells .NET "when the add-on asks for the WebPort version it was built with, use the installed one". WebPort already uses the same mechanism for many of its own libraries.

The `ExtEntraPkce.dll` in this package is built against an older WebPort. On a server with **the same WebPort**, every version matches, and the installer adds no redirect.

On a server with **a newer WebPort**, the installer adds a redirect for `WebPortCommon`, `HttpServer` and `HttpServer.MVC`, from the older version the DLL refers to, to the newer version that is installed. The IdentityModel libraries usually need nothing, because WebPort's own redirects already send every older version to the installed one.

**WebPort and the add-on both use the DLLs of the installed WebPort.** The program folder holds one copy of each library, the one that came with the installed WebPort, and the add-on ships none of its own. The redirect does not bring in or replace any library. It only lets the add-on use the installed one, although the add-on was compiled against an older one. The installer's plan lists the exact versions before anything changes, and `-Uninstall` removes exactly the lines it added.

### How the add-on checks compatibility

A redirect only lets .NET load the add-on. Whether the add-on works with the installed WebPort is decided at startup, by two checks:

- every internal WebPort type, method, property and field it uses exists with the expected signature;
- each of its four SQL statements against WebPort's `Users` and `Settings` tables still fits the database. It runs them with `EXPLAIN`, which makes SQLite check every table and column without reading or changing any row.

If either check fails, the add-on does not start, writes event 3000 with what is missing, and the installer rolls back. Passing the checks does not prove that WebPort's internals still behave the same, so test sign-in after every WebPort upgrade.

### After a WebPort upgrade

A WebPort upgrade or reinstall replaces `WebPortServer.exe.config`, and the redirects are lost. **Run `Install-EntraModule.ps1` again after every WebPort upgrade.** It recreates the redirects that the new version needs, or none if the DLL matches it.

## Client secret renewal

The secret created by the registration script expires after 24 months, the maximum (`-SecretLifetimeMonths` sets 1–24; a tenant policy can limit it further). Entra sign-in stops working when it expires, so renew it before then, for example after a calendar reminder a month ahead:

```powershell
.\Register-WebPortEntraApp.ps1 -Force
```

This adds a new secret in Entra and keeps the old ones. It then replaces `EntraClientSecretProtected` in SQLite and restarts WebPort. When sign-in works, remove the old secret under **App registrations → \<Plant\> → Certificates & secrets**.

## Rollback

```powershell
.\Install-EntraModule.ps1 -Rollback -BackupPath 'C:\ProgramData\WebPort\Entra\install-backups\<timestamp>'
```

This restores the previous DLL and `WebPortServer.exe.config`. The settings stay in SQLite. They do not affect WebPort without the add-on. The backup also holds a copy of the database, but do not restore it just to roll back the add-on: it would undo every change made in WebPort since the backup. Backups must stay outside `C:\Program Files\WebPort`, because WebPort loads every DLL it finds there recursively.

Rollback replaces the whole configuration file with the copy in the backup. To remove the add-on, use `-Uninstall` instead, which changes only what the installer added.

## Uninstall

```powershell
.\Install-EntraModule.ps1 -Uninstall -WhatIf          # show what would be removed
.\Install-EntraModule.ps1 -Uninstall                  # remove the add-on, keep its settings
.\Install-EntraModule.ps1 -Uninstall -RemoveSettings  # also remove its settings and the 1.2 tables
```

| What | `-Uninstall` | with `-RemoveSettings` |
|---|---|---|
| `addons\EntraPkce\` | Removed | Removed |
| `WebPortServer.exe.config` | Only the `bindingRedirect` entries the installer added are removed; the rest of the file stays as it is | Same |
| Event log source `WebPort Entra` | Removed; events already logged stay in the log | Same |
| `Settings`: `EntraClientSecretProtected`, `EntraRedirectUri`, `EntraRolesEnabled` | Kept, so that a later installation works at once | Removed |
| The `WebPortEntra*` tables of version 1.2 | Kept | Removed |
| `Settings`: `KIClientId`, `KIExternalTenantId`, `KIProviderUri`, `WebPortServer Url`, `AuthenticationTypeKI` | Never touched; they belong to WebPort's own sign-in. Set `AuthenticationTypeKI` to `0` in WebPort if its SSO option should be turned off | Never touched |
| Accounts in `Users` created by the add-on | Not touched | Not touched |
| The app registration in Entra | Not touched | Not touched |

Like an installation, the uninstallation shows its plan first, backs up the configuration file, the DLL and the database, and stops WebPort. It then starts WebPort again if it was running, and checks that `/entra/status` no longer answers as the add-on. If a step fails, the DLL and the configuration file are restored automatically. Settings removed with `-RemoveSettings` are not restored automatically, but they are in the backup's `webport.sqlite`.

Afterwards, lock or delete WebPort accounts that were only used through Entra, and delete the app under **App registrations** if it is no longer needed.

## License and trademarks

Released under the [MIT License](LICENSE). The software is provided as is, without warranty of any kind.

WebPort and Kiona are trademarks of their respective owners. Microsoft, Microsoft Entra and Windows are trademarks of the Microsoft group of companies. This project is independent and is not affiliated with, endorsed by or supported by any of them.
