# Azure Government and Commercial

`dbx-sharepoint` supports both Azure Commercial and Azure Government clouds. The correct endpoints are detected automatically from the SharePoint site URL, and no explicit configuration is typically required.

## Auto-detection

When a `SharePointClient` is created, the library parses the site URL's hostname:

| Hostname suffix | Environment | Login authority | Graph endpoint |
|---|---|---|---|
| `*.sharepoint.com` | Commercial | `https://login.microsoftonline.com` | `https://graph.microsoft.com` |
| `*.sharepoint.us` | Government | `https://login.microsoftonline.us` | `https://graph.microsoft.us` |

The matching Graph scope (`<endpoint>/.default`) is used when acquiring tokens.

If the hostname does not match either suffix, the library raises `ValueError`. Pass `graph_endpoint` explicitly to bypass detection.

## Why this matters

Azure Government is a separate cloud from Commercial, with a different identity provider, a different Graph service, and different certificates. Authenticating against `login.microsoftonline.com` with Gov tenant credentials fails with an authority mismatch. Tokens issued by the Commercial authority are not accepted by `graph.microsoft.us`.

The `from_databricks_secrets` factory handles this transparently. It resolves the site URL first, detects the environment, and constructs the `ClientSecretCredential` with the correct `authority` parameter. When constructing the credential directly, pass the authority explicitly:

```python
from azure.identity import ClientSecretCredential

cred = ClientSecretCredential(
    tenant_id="...",
    client_id="...",
    client_secret="...",
    authority="https://login.microsoftonline.us",  # Gov. Required.
)
```

## Manual override

For tenants on a non-standard endpoint (for example, a sovereign cloud not listed above), pass `graph_endpoint` explicitly. The login authority is inferred from the matching environment. For fully custom setups, construct the credential directly and pass it in.

```python
sp = SharePointClient(
    credential=cred,
    site_url="https://myorg.example-cloud/sites/TeamSite",
    graph_endpoint="https://graph.example-cloud",
)
```

## Things that work the same on Gov and Commercial

- Site resolution (`/v1.0/sites/{hostname}:/sites/{site-name}`)
- Drive item access (`/drives/root:/{path}:/content`)
- Paging, throttling, error codes
- The template/openpyxl workflow (no Graph involved)
- Shared-link downloads (`read_excel_from_shared_link`) — the shared link URL itself points to the correct cloud

## Gov-specific considerations

- App registration must live in a Gov tenant. A Commercial app registration cannot authenticate against Gov Graph. A separate app registration in the Gov Entra ID tenant is required.
- Admin consent is tenant-scoped. Consent granted in a Commercial tenant does not apply to Gov.
- Some Graph features lag Commercial. When a Graph endpoint works on `graph.microsoft.com` but not `graph.microsoft.us`, the Gov rollout has not reached it yet. The endpoints this library uses (`/sites`, `/drives`, `/drive/root:/...`) are long-standing and stable on both clouds.
- Network egress. The Databricks workspace must be able to reach `login.microsoftonline.us` and `graph.microsoft.us`. Behind a firewall with allow-lists, add both hostnames.
