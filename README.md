# acumatica-sharepoint-project-creation

A .NET 10 (isolated worker) **Azure Function** that polls Acumatica every 15 minutes for
newly-created projects and, for each one, creates a **SharePoint Document Set** in the
library mapped to the project's *practice*, stamped with metadata:

- **Project Id**
- **Customer Name**
- **Project Name**
- **Project Manager**

## How it works

```
Timer (every 15 min)
   │
   ▼
Read last-run watermark  ──►  Azure Blob (last-run.txt)
   │
   ▼
Acumatica: OAuth2 token ──►  query Generic Inquiry (OData feed)
                             $filter = CreatedDateTime gt <watermark>
   │
   ▼
For each new project:
   • resolve practice → target library + folder
   • create Document Set (PnP.Framework / CSOM, app-only cert)
   • set Project Id / Customer / Name / Manager metadata
   • (idempotent — skips if a set with that name already exists)
   │
   ▼
Advance watermark to newest processed CreatedDateTime
```

Key design points:

- **Watermark + overlap.** State is the created-datetime of the newest processed project.
  Each poll re-queries with a small overlap (`State:OverlapMinutes`) to avoid missing
  records from commit latency; duplicate work is prevented by the idempotency check.
- **Fail-safe ordering.** Projects are processed oldest-first. If one fails, the cycle
  halts and the watermark stays just before it, so it retries next run — nothing is skipped.
- **Everything is config.** GI name, field names, SharePoint columns, and practice→library
  mappings are all app settings — no code changes to adapt to your instance.

## Project layout

| Path | Purpose |
|------|---------|
| `Functions/ProjectSyncTimerFunction.cs` | Timer trigger (`%ProjectSyncSchedule%`) |
| `Functions/ProjectSyncHttpFunction.cs` | Manual **run-now** HTTP endpoint (`POST /api/sync/run`) |
| `ProjectSyncProcessor.cs` | Orchestration: watermark → query → create → advance |
| `Acumatica/` | OAuth2 token provider + GI OData client |
| `SharePoint/` | App-only cert auth + Document Set creation (CSOM/PnP) |
| `State/` | Blob-backed last-run store |
| `Options/` | Strongly-typed settings |

Tests live in `tests/ProjectSync.Functions.Tests` (xUnit + Moq).

## Prerequisites you must set up

### 1. Acumatica

- **Connected application** (System → Integration → Connected Applications) using the
  **Client Credentials** flow. Copy the **Client ID** and **Client Secret** into
  `Acumatica:ClientId` / `Acumatica:ClientSecret`.
- A **Generic Inquiry** that returns one row per project with columns for project id,
  name, customer, project manager, **practice**, and a **created date/time**. Expose it
  to OData (GI editor → *Make Visible on the UI* / OData checkbox). Put its name in
  `Acumatica:GenericInquiryName` and map the column names in the `Acumatica:*Field` settings.
  - The client reads `{BaseUrl}/t/{Tenant}/api/odata/gi/{GenericInquiryName}`.

### 2. SharePoint (Azure AD app-only, certificate)

- Register an **Entra ID app**. Add **SharePoint → Application permission**
  `Sites.Selected` (recommended) or `Sites.FullControl.All`, and **grant admin consent**.
- With `Sites.Selected`, grant the app write access to the specific site(s) via PnP
  PowerShell: `Grant-PnPAzureADAppSitePermission`.
- Create a **certificate**, upload the public key to the app registration, and provide the
  private key to the function via **either**:
  - `SharePoint:CertificateBase64` (+ `SharePoint:CertificatePassword`) — base64 of the `.pfx`, or
  - `SharePoint:CertificateThumbprint` — if the cert is installed in the host cert store
    (e.g. App Service `WEBSITE_LOAD_CERTIFICATES`).
- In each target library, enable the **Document Set** content type and create the metadata
  **columns**; set their internal names in `SharePoint:*Column`.

### 3. Practice → destination mapping

Map each Acumatica practice value to a library (and optional parent folder). Use `*` as a
catch-all. Example (flattened app-setting form):

```
SharePoint:PracticeMappings:Advisory:Library     = Project Documents
SharePoint:PracticeMappings:Advisory:ParentFolder = Advisory
SharePoint:PracticeMappings:*:Library            = Project Documents
SharePoint:PracticeMappings:*:ParentFolder       = Unmapped
```

A mapping entry may also set `SiteUrl` to target a different site collection per practice.

### 4. Storage

An Azure Storage account (the Functions `AzureWebJobsStorage` is reused by default for the
watermark blob; override with `State:ConnectionString`).

## Local development

1. Fill in `src/ProjectSync.Functions/local.settings.json` (it's git-ignored).
2. Start Azurite (or point `AzureWebJobsStorage` at a real account).
3. Install [Azure Functions Core Tools], then:

   ```bash
   cd src/ProjectSync.Functions
   func start
   ```

Build and test:

```bash
dotnet build
dotnet test
```

### Trigger a run manually

Instead of waiting for the timer, POST to the run-now endpoint (same cycle, returns a JSON
summary). Locally it needs no key; when deployed, pass the function key.

```bash
curl -X POST "http://localhost:7071/api/sync/run"
```

Response:

```json
{ "found": 2, "created": 2, "updated": 0, "hadFailure": false, "watermark": "2026-08-05T11:00:00.0000000+00:00" }
```

Deployed, include the key: `POST https://<app>.azurewebsites.net/api/sync/run?code=<function-key>`.

## Deployment

Deploy as a **.NET 10 isolated** Function App. Set every value from `local.settings.json`
as **Application settings** (use Key Vault references for `ClientSecret`,
`CertificatePassword`, and `CertificateBase64`). The schedule is controlled by
`ProjectSyncSchedule` (NCRONTAB, default `0 */15 * * * *`).

[Azure Functions Core Tools]: https://learn.microsoft.com/azure/azure-functions/functions-run-local
