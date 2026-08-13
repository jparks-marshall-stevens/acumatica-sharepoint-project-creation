# acumatica-sharepoint-project-creation

A **.NET 10 (isolated worker) Azure Function** that polls Acumatica every 15 minutes for
newly-created projects and, for each one in a configured practice, creates a **SharePoint
Document Set** — named from the project description, stamped with metadata, and permissioned
to the project manager and practice leader.

Currently scoped to the **Estate & Gift** practice → the **GiftEstate** SharePoint site.

## How it works

```
Timer (every 15 min)  ─►  read last-run watermark (Azure Blob)
   │
   ▼
Acumatica: OAuth2 token (ROPC/password grant)
   │        query Generic Inquiry over OData:  $filter = CreatedOn gt <watermark>
   ▼
For each new project (filtered to IncludedPractices, minus ExcludedProjectIds):
   • resolve practice → site + library + parent folder
   • create Document Set (content type "Project") named from first 40 chars of Description
   • set metadata: Project Id, Customer Name, Project Name, Project Manager (People field)
   • set permissions: break inheritance; Owners = Full Control;
                      Project Manager + Practice Leader = Edit
   • idempotent — dedup by the Project Id column; re-runs update instead of duplicating
   ▼
Advance + persist watermark (newest processed CreatedOn)
```

### Key design points
- **Moving-forward only.** `State:FirstRunLookbackHours = 0`, so the first run stamps the
  watermark at "now" and only new projects are processed — no historical backfill.
- **Folder naming.** First 40 chars of the Description (`SharePoint:DocumentSetNameMaxLength`),
  sanitized. Descriptions aren't unique, so **dedup is keyed on the Project Id column**, and a
  colliding folder name gets the project id appended. Blank descriptions fall back to the id.
- **People field.** The GI's `ProjectManager` column returns the PM's **email**; the function
  resolves it to a SharePoint user (`EnsureUser`). Emails outside the tenant are left blank (fail-soft).
- **Fail-safe ordering.** Oldest-first; if a project fails the cycle halts and the watermark holds
  just before it, so nothing is skipped. A permission-set failure is logged (trips the alert) but
  doesn't roll back the created set.
- **Everything is config** — GI name, field names, SharePoint columns, practice mappings, and
  permission behavior are all app settings.

## Project layout

| Path | Purpose |
|------|---------|
| `Functions/ProjectSyncTimerFunction.cs` | Timer trigger (`%ProjectSyncSchedule%`) |
| `Functions/ProjectSyncHttpFunction.cs` | Manual **run-now** endpoint (`POST /api/sync/run`) |
| `ProjectSyncProcessor.cs` | Orchestration: watermark → query → filter → create → advance |
| `Acumatica/` | OAuth2 (ROPC) token provider + GI OData client |
| `SharePoint/` | App-only cert auth, Document Set create, metadata, permissions (CSOM/PnP) |
| `State/` | Blob-backed last-run store |
| `Options/` | Strongly-typed settings |
| `tools/` | Standalone diagnostic consoles (see below) |

Tests: `tests/ProjectSync.Functions.Tests` (xUnit + Moq).

### Diagnostic tools (`tools/`)
Each reads config from `local.settings.json` and needs no Functions runtime:
- **`AcumaticaConnectivityTest`** — verifies Acumatica auth + GI field mapping.
- **`ProjectSyncDryRun -- <days>`** — previews what would be created (+ PM email-domain breakdown).
- **`SharePointConnectivityTest`** — verifies SharePoint auth, library, content type, column names (read-only).
- **`CreateOneDocumentSet -- <ProjectId>`** — creates/updates one set and reads back metadata + permissions.

## Configuration

### Acumatica (`Acumatica:*`)
Auth is **Resource Owner Password Credentials** (`GrantType = password`) — this Acumatica version's
connected-app UI doesn't offer client-credentials, and there's no static API key. Needs a service
user (`Username`/`Password`) plus the connected app's `ClientId`/`ClientSecret`.
- Instance is under a virtual directory: `BaseUrl = https://erp.marshall-stevens.com/acu`.
- GI `CUST-AzureFunction-SharePoint-FolderCreate` read at `{BaseUrl}/t/{Tenant}/api/odata/gi/{GI}`.
- Field names map GI OData properties: `ProjectId`, `CustomerName`, `Description` (= project name),
  `ProjectManager` (= email), `CreatedOn`. `ProjectManagerEmailField` points at the same column so
  the People field resolves.
- **`IncludedPractices`** allow-list (indexed): only these practices are synced; others are skipped
  but still advance the watermark. **`ExcludedProjectIds`**: hard-ignore list (e.g. `X`, the
  Non-Project Code) checked before the practice filter.

### SharePoint (`SharePoint:*`)
App-only **certificate** auth (Entra app `SharePoint:ClientId`, tenant `SharePoint:AzureAdTenant`).
Cert via `CertificateBase64` (+ `CertificatePassword`) — works cross-OS — or `CertificateThumbprint`
for local dev (cert in `CurrentUser\My`).
- `SiteUrl`, `DocumentSetContentType = Project` (a custom Document Set content type).
- Column **internal** names are space-encoded: `Project_x0020_Id`, `Customer_x0020_Name`,
  `Project_x0020_Name`, `Project_x0020_Manager` (a **People** field, `ProjectManagerIsPersonColumn = true`).
- **Permissions**: `SetProjectPermissions`, `PermissionLevel = Edit`, `RestrictPermissions = true`
  (break inheritance). Requires the app to have **Sites.Selected `fullControl`** on the site.

### Practice mappings (indexed list)
An indexed list (not a dictionary) so it binds from env vars — practice values contain spaces/`&`
which are invalid in config **key** names. Each entry maps a practice to a destination + leader:

```
SharePoint:PracticeMappings:0:Practice           = Estate & Gift
SharePoint:PracticeMappings:0:PracticeLeaderEmail = bjohnson@marshall-stevens.com
SharePoint:PracticeMappings:0:SiteUrl            = https://marshallstevens.sharepoint.com/sites/GiftEstate
SharePoint:PracticeMappings:0:Library            = Documents
SharePoint:PracticeMappings:0:ParentFolder       = Projects/Active
```

To add a practice: add an entry (`:1:…`) with its own site/library, add the practice to
`Acumatica:IncludedPractices`, grant the app `fullControl` on that site, and provision the library
(the "Project" content type + the metadata columns). `SharePointConnectivityTest` verifies a site's setup.

## Local development

1. Fill in `src/ProjectSync.Functions/local.settings.json` (git-ignored).
2. `dotnet build` / `dotnet test`.
3. With [Azure Functions Core Tools] + Azurite: `cd src/ProjectSync.Functions && func start`.

### Run-now endpoint
```bash
# normal run (deployed: add ?code=<host key>)
curl -X POST ".../api/sync/run"
# dry run — previews a `plan`, creates nothing, doesn't move the watermark
curl -X POST ".../api/sync/run?dryRun=true&days=30"
# targeted reprocess of one project — create/update + re-apply permissions, watermark untouched
curl -X POST ".../api/sync/run?projectId=10-31-21-74663&days=30"
```

## Deployment (Azure)

Deployed to **Windows Consumption** (resource group `rg-projectsync`, East US):

- Function App `func-projectsync-2276c5` — **Windows** Consumption, .NET 10 isolated, Functions v4.
- Storage `stprojsync2276c5`; App Insights; **Key Vault `kv-projsync-2276c5`**.
- Secrets (`acumatica-client-secret`, `acumatica-password`, `sharepoint-cert-base64`,
  `sharepoint-cert-password`) live in Key Vault; the Function App's **system-assigned managed
  identity** reads them via `@Microsoft.KeyVault(SecretUri=…)` app-setting references.
- App settings use **`__`** separators (env-var form). Deploy with `az functionapp deployment
  source config-zip`.
- **Failure alert**: an App Insights log alert (`ProjectSync-Errors`) emails jparks@marshall-stevens.com
  when the function logs an Error/exception in a 15-minute window.

> **Why Windows, not Linux Consumption:** Azure CLI 2.83 has a bug validating .NET 10 on Linux
> (`DOTNET-ISOLATED|10.0` vs its list of `10`), which breaks `config-zip`/`function list`, and Linux
> Consumption's SCM is unreachable when idle. Windows Consumption avoids both; the cert uses the
> base64/Key-Vault path so nothing depends on the OS cert store.

[Azure Functions Core Tools]: https://learn.microsoft.com/azure/azure-functions/functions-run-local
