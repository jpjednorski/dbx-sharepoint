# Authentication and secrets

`dbx-sharepoint` authenticates to Microsoft Graph using an Azure AD app registration with a client secret. All SharePoint calls are made as the application (app-only auth), not as a user.

## Minimum permissions (TL;DR)

For least-privilege setup, the library needs only:

1. **`Sites.Selected`** (Graph application permission) on the app registration, with tenant admin consent.
2. A **per-site grant** giving the app `read` or `write` role on each target site.

No `Sites.*.All` or `Files.*` permissions are required. See [Option A](#option-a-sitesselected-recommended) for the exact grant call.

## Who does what

Setup touches three systems, usually owned by different teams. Most engagements stall on the handoff, not the technical steps.

| Step | System | Who you need |
|---|---|---|
| Register the app, add Graph **application** permissions, grant tenant admin consent | Microsoft Entra ID (Azure AD) | Your Entra tenant admin team (Global Administrator or Privileged Role Administrator) |
| Authorize the app on specific SharePoint sites (Option A) | SharePoint / Graph | The SharePoint site owner decides which sites; the grant call itself currently requires a tenant admin to execute |
| Store the tenant ID, client ID, and client secret | Databricks | Workspace admin or whoever owns the target secret scope |

Line up all three owners before starting. The Entra admin and SharePoint site owner are often different people in different orgs, and the grant sequence only works top-down (tenant-level consent first, then per-site authorization).

---

## 1. Register an Azure AD application

In the Azure portal (or the Gov portal for Gov tenants), go to **Entra ID → App registrations → New registration**:

- Name: something descriptive, e.g. `databricks-sharepoint-reader`
- Supported account types: single tenant
- Redirect URI: leave blank (app-only flow)

From the app's overview page, record the **Application (client) ID** and **Directory (tenant) ID**.

Under **Certificates & secrets**, create a client secret and copy the **Value** immediately. It is shown only once. The Secret ID is a different field; do not confuse the two.

## 2. Grant Microsoft Graph permissions

The app needs application permissions (not delegated). Three options, in order of preference:

### Option A: `Sites.Selected` (recommended)

`Sites.Selected` alone grants access to no sites. An admin then grants the app a role on each site it should access.

1. API permissions → Add → Microsoft Graph → Application permissions → `Sites.Selected`.
2. Grant admin consent.
3. For each target site, a tenant admin runs:

   ```http
   POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions
   Content-Type: application/json

   {
     "roles": ["read"],              // or ["write"] or ["owner"]
     "grantedToIdentities": [{
       "application": {
         "id": "<client-id>",
         "displayName": "databricks-sharepoint-reader"
       }
     }]
   }
   ```

   Use `https://graph.microsoft.us` on Gov. Resolve `{site-id}` with:
   ```http
   GET https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site-name}
   ```

### Option B: `Sites.Read.All` or `Sites.ReadWrite.All`

Tenant-wide SharePoint access. Simpler operationally, but rarely approved in strict security reviews.

### Option C: `Files.Read.All` or `Files.ReadWrite.All`

Tenant-wide file and drive-item access across SharePoint and OneDrive. Useful when an existing app registration already holds them, or when tenant policy allows `Files.*` but not `Sites.*`.

`Files.*` application permissions cannot be scoped more narrowly in this flow:

- The `.All` suffix is tenant-wide by design.
- `Files.Read.Selected` and `Files.ReadWrite.Selected` are delegated-only (user-context); they do not apply to app-only auth.
- `Files.SelectedOperations.Selected` is app-only but applies to SharePoint Embedded containers, not classic team sites and document libraries.

For per-site scoping, use `Sites.Selected` (Option A).

### Operations by permission

| Library call | `Sites.Selected` role | Also works with |
|---|---|---|
| `list_files`, `download`, `read_excel`, `open_template` | `read` | `Sites.Read.All`, `Sites.ReadWrite.All`, `Files.Read.All`, `Files.ReadWrite.All` |
| `upload`, `write_excel`, `save` | `write` | `Sites.ReadWrite.All`, `Files.ReadWrite.All` |

## 3. Store credentials in a Databricks secret scope

```bash
databricks secrets create-scope sharepoint
databricks secrets put-secret sharepoint tenant-id       # Azure AD tenant GUID
databricks secrets put-secret sharepoint client-id       # Application (client) ID
databricks secrets put-secret sharepoint client-secret   # Client secret VALUE
databricks secrets put-secret sharepoint site-url        # https://myorg.sharepoint.us/sites/TeamSite
```

`site-url` is optional. Pass it to the factory instead when one scope serves multiple sites.

### Multiple environments

When a single scope serves several environments, use a prefix:

```bash
databricks secrets put-secret sharepoint prod-tenant-id      ...
databricks secrets put-secret sharepoint prod-client-id      ...
databricks secrets put-secret sharepoint prod-client-secret  ...
databricks secrets put-secret sharepoint prod-site-url       ...
```

```python
sp_prod = SharePointClient.from_databricks_secrets(dbutils=dbutils, scope="sharepoint", prefix="prod")
sp_dev  = SharePointClient.from_databricks_secrets(dbutils=dbutils, scope="sharepoint", prefix="dev")
```

## 4. Connect

```python
from dbx_sharepoint import SharePointClient

sp = SharePointClient.from_databricks_secrets(dbutils=dbutils, scope="sharepoint")
```

The factory reads the secrets, detects Gov versus Commercial from the site URL, and passes the correct authority to `ClientSecretCredential` (`login.microsoftonline.us` for Gov, `login.microsoftonline.com` for Commercial). Gov tenants fail to authenticate without this, so callers building a credential directly must set `authority` explicitly.

## 5. Alternatives

### Bring your own credential

Any `azure.identity` `TokenCredential` is accepted. Useful when credentials are managed outside Databricks, or with a managed identity from a non-Databricks process.

```python
from azure.identity import ClientSecretCredential
from dbx_sharepoint import SharePointClient

cred = ClientSecretCredential(
    tenant_id="...",
    client_id="...",
    client_secret="...",
    authority="https://login.microsoftonline.us",  # Gov. Omit for Commercial.
)
sp = SharePointClient(
    credential=cred,
    site_url="https://myorg.sharepoint.us/sites/TeamSite",
)
```

### Custom Graph endpoint

For tenants on a non-standard Graph endpoint, pass it explicitly:

```python
sp = SharePointClient(
    credential=cred,
    site_url="https://myorg.sharepoint.us/sites/TeamSite",
    graph_endpoint="https://graph.microsoft.us",
)
```

## Credential rotation

Client secrets expire. To rotate:

- Create a new secret in Azure AD before the current one expires.
- Update `client-secret` in the Databricks scope.
- Delete the old Azure AD secret once all workloads have picked up the new one.

The library holds no state that needs clearing. A fresh token is acquired on each request, so rotation takes effect on the next call.

## Security notes

- Never write the client secret to notebook output. `dbutils.secrets.get` returns a redacted value when printed in Databricks.
- Use `Sites.Selected` whenever possible.
- Restrict the secret scope ACL to only the users and service principals that need it.
- The library does not log request bodies or tokens. Exception messages from 4xx responses include the Graph error body verbatim; wrap calls in `try`/`except` and log a sanitized message when needed.
