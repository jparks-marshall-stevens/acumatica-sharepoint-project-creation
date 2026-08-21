# acumatica-sharepoint-project-creation

A **.NET 10 (isolated worker) Azure Function** app that keeps **SharePoint Document Set**
workspaces in sync with two source systems, across an engagement's lifecycle:

- **HubSpot** (pre-ERP **scoping** phase) — polls deals in scope and creates a workspace
  (`Status = Scoping`) keyed on the HubSpot deal id.
- **Acumatica** (ERP **execution** phase) — polls newly-created projects and creates a workspace
  (`Status = Execution`), then **reconciles** existing ones as the project team / PM / description change.

Each workspace is stamped with metadata, permissioned to the right people, and (optionally) gets a
`Client Uploads` subfolder with an anonymous **Request-files** link. Currently scoped to the
**Estate & Gift** practice → the **GiftEstate** SharePoint site.

The two sources run on **independent timers** (they hit different systems, and each writes its own
idempotent doc sets). One engagement gets **one workspace for its whole life**: it is born at scoping from a
HubSpot deal and **promoted in place** when it reaches the ERP — renamed, restamped, re-permissioned, never
moved and never duplicated. See [`docs/hubspot-scoping-integration.md`](docs/hubspot-scoping-integration.md).

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
   • create Document Set (content type "Project") named "{customer} ({project id})"
   • set metadata: Project Id, Customer Name, Project Name, Project Manager (People field)
   • set permissions: break inheritance; Owners = Full Control;
                      Project Manager + Practice Leader = Edit
   • idempotent — dedup by the Project Id column; re-runs update instead of duplicating
   ▼
Advance + persist watermark (newest processed CreatedOn)
```

### Reconcile (keep tracked sets in sync)
Beyond create-on-new, two timers keep **already-tracked** Acumatica doc sets current — signature-gated
(a SHA-256 of the desired state in a hidden `ProjectSyncSig` column), so unchanged sets cost zero writes:
- **Incremental** (`%ProjectSyncReconcileSchedule%`) — short-circuits on the team GI's `ModifiedOn`
  watermark, touching only projects whose team changed.
- **Full** (`%ProjectSyncFullReconcileSchedule%`, daily) — every tracked set, catching PM/description
  changes and team removals.

Permissions include the **project team** (from the `EPEmployeeContract` team GI) alongside PM + leader.

> **Schedule gap (intentional):** the two Acumatica timers (`ProjectSyncSchedule`,
> `ProjectSyncReconcileSchedule`) run `0 0,15,30,45 0-8,10-23 * * *` — every 15 min **except the 09:00
> UTC hour**. The Acumatica OAuth token endpoint hangs to a 100 s timeout every day at exactly
> 09:15:00 UTC (04:15 ET), failing those runs and firing the `ProjectSync-Errors` alert. Skipping the
> hour avoids it without swallowing errors; watermark-gating means nothing is lost. Don't "fix" this
> back to `*/15` until whatever runs on the ERP at 04:15 ET is identified.

### Promotion (scoping → execution)

When a HubSpot deal becomes an Acumatica project, the existing scoping workspace is **converted in place**
rather than duplicated. The link is the HubSpot **opportunity number** (`quote_number`, labelled *Opportunity
#*, e.g. `PQ005871`), recorded on the Acumatica project and exposed by the GI as **`PQCode`**
(`Acumatica:HubSpotLinkField`).

Lookup order on each Acumatica poll:

1. **`Project Id`** column — already tracked, just refresh metadata + permissions.
2. Only if `PQCode` is non-blank: the **`OpportunityId`** column, then the **`HubSpotDealId`** column. A hit on
   either is a **promotion** — so it links whether a person recorded the opportunity number or pasted the raw
   HubSpot record id.
3. No match → a fresh document set (a project that skipped scoping).

A promotion renames the folder to the project-id form (`FileLeafRef` — a rename, not a move: contents, item
id, and the client-upload link all follow it), stamps `Project Id` and `Status = Execution`, and re-applies
permissions for the delivery team. It does **not** re-run the client-uploads step — the scoping phase already
created that folder and minted its link. Reported separately as `Promoted=` in the cycle log and in
`ProjectSyncResult`.

Failure modes, both logged as warnings rather than silently swallowed:
- **`PQCode` blank** → the engagement gets a second folder; merge by hand. `ProjectSyncDryRun` reports how many
  projects in the window carry a link at all.
- **Two projects, one `PQCode`** → the second does not hijack the first one's workspace; it gets its own.

> Note: `PQCode` is currently **null on every GI row**, so promotion does not fire yet. Populating it at
> conversion is a process step, not a code change.

### Client uploads (optional)
When `SharePoint:CreateClientUploadLink = true`, each new workspace gets a `Client Uploads` subfolder and
an anonymous, upload-only ("Request files") Microsoft **Graph** link (30-day expiry) stamped as a plain
text value in the `ClientUploadLink` column (copy-and-send to the client). Requires a Graph
`Sites.Selected` grant on the site and tenant "Anyone" links enabled. No one is notified on upload
(the built-in notification goes to the app identity, which has no mailbox).

### Key design points
- **Moving-forward only.** `State:FirstRunLookbackHours = 0`, so the first run stamps the
  watermark at "now" and only new projects are processed — no historical backfill.
- **Folder naming.** `{first N chars of Customer Name} ({id})` (`N` =
  `SharePoint:DocumentSetNameMaxLength`, default 40), sanitized for SharePoint. The id is the **Project Id**
  for Acumatica sets (`Robert Palmer (10-31-21-74663)`) and the **opportunity number / PQCode** for scoping
  sets (`Blackstone Dilworth (PQ007180)`). Because the id is unique the names are effectively unique;
  **dedup is still keyed on the metadata columns, never the name**. Blank customer names fall back to the id.
  - Note the name is written **once at creation and never revisited**, while `Customer Name` is refreshed on
    every poll — so a customer that resolves differently later (e.g. the deal's client contact gets filled
    in) leaves the folder name stale. `BackfillOpportunityIds` re-syncs scoping folder names.
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
| `Functions/ProjectSyncTimerFunction.cs` | Acumatica create timer (`%ProjectSyncSchedule%`) |
| `Functions/ProjectSyncReconcileFunctions.cs` | Reconcile timers (incremental + full) |
| `Functions/HubSpotScopingFunctions.cs` | HubSpot scoping poll timer (`%HubSpotScopingSchedule%`) |
| `Functions/ProjectSyncHttpFunction.cs` | Manual **run-now** endpoint (`POST /api/sync/run`) |
| `ProjectSyncProcessor.cs` | Acumatica orchestration: watermark → query → filter → create/reconcile |
| `Acumatica/` | OAuth2 (ROPC) token provider + GI OData client (projects + team) |
| `HubSpot/` | OAuth/token provider, CRM v3 deals client, scoping poll processor |
| `SharePoint/` | Cert auth, Document Set create, metadata, permissions, Graph upload links (CSOM/PnP) |
| `State/` | Blob-backed last-run + named watermark store |
| `Options/` | Strongly-typed settings |
| `tools/` | Standalone diagnostic consoles (see below) |

Tests: `tests/ProjectSync.Functions.Tests` (xUnit + Moq).

### Diagnostic tools (`tools/`)
Each reads config from `local.settings.json` and needs no Functions runtime:
- **`ProvisionPracticeSite`** — provisions a NEW practice site so it mirrors Gift & Estate: reads G&E's
  "Project" Document Set content type + base columns and reproduces them on the target site, along with
  `Projects/Current` (locked) and versioning. Dry-run by default; `--apply` to write. Requires the
  one-time `bootstrap-practice-site.ps1` (site creation + cert grant) first. See "add a practice" above.
- **`AcumaticaConnectivityTest`** — verifies Acumatica auth + GI field mapping.
- **`ProjectSyncDryRun -- <days>`** — previews what would be created (+ PM email-domain breakdown).
- **`SharePointConnectivityTest`** — verifies SharePoint auth, library, content type, column names (read-only).
- **`SharePointHardening [--lock]`** — reports/enables versioning + recycle bin; locks the `Current` folder to code-only creation.
- **`CreateOneDocumentSet -- <ProjectId> [--delete]`** — creates/updates one Acumatica set; reads back metadata + permissions.
- **`ReconcileOnce [full|incremental]`** — runs a single reconcile pass against the real systems.
- **`BackfillOpportunityIds [--apply]`** — stamps `OpportunityId` on scoping workspaces created before that
  column existed, and renames folders still named after the raw deal id to the `{customer} ({PQCode})` form
  (dry-run by default; creates the column on `--apply`). Batch-reads deals by id, so deals that have since
  closed are still resolved. Touches that column and the folder leaf name only.
- **`HubSpotOAuthSetup`** — one-time: captures an OAuth refresh token via a local redirect.
- **`HubSpotConnectivityTest`** — verifies the HubSpot token; lists pipelines/stages + candidate properties.
- **`HubSpotPollOnce -- [lookbackHours]`** — dry-run plan of scoping workspaces that would be created.
- **`CreateOneScopingWorkspace -- <dealId> [--delete]`** — creates/updates one scoping set from a HubSpot deal.

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
- **`HubSpotLinkField = PQCode`** — the GI column carrying the HubSpot opportunity number, which drives
  promotion (see above). Blank disables the promotion path entirely.

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
SharePoint:PracticeMappings:0:ParentFolder       = Projects/Current
```

To add a practice, use the two-step provisioning flow (see `tools/ProvisionPracticeSite`):

1. **One-time, admin-only bootstrap** (`bootstrap-practice-site.ps1`, run interactively with your own
   SharePoint-admin identity — no app registration): creates the empty site collection matching the
   Gift & Estate template and grants the cert app `Sites.Selected fullControl` on it.
2. **Repeatable provisioning** (`ProvisionPracticeSite`, cert auth like every other tool): mirrors the
   G&E library structure onto the new site — the "Project" Document Set content type, the four base
   metadata columns (read off G&E so the *global* internal names match exactly), `Projects/Current`
   (locked to code-only creation), and versioning. It then prints the ready-to-paste `PracticeMappings`
   entry.

The `PracticeMappings` entry and the site provisioning are **inert on their own**: the sync only touches
a practice once it is added to `Acumatica:IncludedPractices` **and** `HubSpot:IncludedPractices` — that
allow-list addition is the single go-live switch. `SharePointConnectivityTest` verifies a site's setup.

### HubSpot scoping (`HubSpot:*`)
A second source: the **`HubSpotScopingPoll`** timer (`%HubSpotScopingSchedule%`) polls HubSpot deals
modified since a persisted watermark (`hubspot-deals`) and creates/updates a scoping workspace for each
in-scope deal, keyed on `HubSpotDealId`, with `Status = Scoping` and access for the deal **owner** +
practice leader. A workspace that has already been promoted (its `Project Id` column is set) is skipped —
otherwise a later edit to the deal would demote it and wipe the delivery team's access.

The sync is **authoritative** over an un-promoted scoping room: on each poll it re-stamps metadata,
re-derives the folder name from the current customer + PQCode and renames in place, and resets permissions
(break inheritance; Owners + deal owner + practice leader). Manual changes to a scoping room's name or
permissions are therefore reverted. Note the cadence: a room is only re-processed when **its deal is
modified** in HubSpot (a customer-name change counts), so drift introduced without a subsequent deal edit
persists until the deal next changes. `BackfillOpportunityIds` re-syncs names on demand in the meantime.

- **Auth**: OAuth refresh-token (`ClientId`/`ClientSecret`/`RefreshToken`) preferred, or a static
  private-app token (`AccessToken`) as a fallback. Token is a Key Vault secret in production.
- **In scope**: `PracticeProperty = practices` (a multi-select) contains an `IncludedPractices` value
  (e.g. `Estate & Gift`), **and** the deal stage is not in `TerminalStageIds` (Won/Lost/Closed).
- **`OpportunityIdProperty = quote_number`** — the *Opportunity #* property, stamped onto the workspace's
  `SharePoint:OpportunityIdColumn` on every poll and matched against `PQCode` at promotion. Deliberately
  separate from the deal id: the deal id is immutable identity (so it stays the idempotency key), while an
  opportunity number can be assigned or corrected later.
- **Customer name**: resolved from the deal's **client contact** — the contact's `company` text, then its
  associated company name, then the deal name (order via `CustomerCompanyTextFirst`).
- **Watermark + floors**: `FirstRunLookbackHours` (default 0 = moving-forward); `CreatedAfter` — an
  optional created-date floor so pre-existing open deals aren't backfilled when HubSpot bumps their
  modified date. Search paging respects HubSpot's 10,000-result window and retries on HTTP 429.

### Notifications (`Notifications:*`)
Workspace emails (created / access-added / client-upload) are sent via Microsoft Graph `sendMail` as
`FromAddress` (a no-reply mailbox; the app's `Mail.Send` is scoped to it by an Exchange Application
Access Policy). `Enabled = false` composes/sends nothing. Two rollout safety valves redirect email to
`TestRecipient` (intended recipients shown in the subject, BCC dropped):
- **`TestMode`** (bool) — **global**: every practice's email is redirected. `[TEST]` subject tag.
- **`SilentPractices`** (indexed list) — **per-practice**: only the named practices are redirected;
  every other practice emails its real recipients normally. `[SILENT]` subject tag. Use it to switch a
  **new** practice on for verification without diverting a **live** practice's (e.g. Estate & Gift) mail.
  Matched case-insensitively against the workspace's practice, token-by-token for multi-select values.
  Example: `Notifications:SilentPractices:0 = Marital Dissolution` (env-var form
  `Notifications__SilentPractices__0`). `TestMode`, when on, still overrides and redirects everything.

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
- **Functions**: `ProjectSyncTimer`, `ProjectSyncReconcileIncremental`, `ProjectSyncReconcileFull`,
  `HubSpotScopingPoll`, `ProjectSyncRunNow` (HTTP).
- Secrets (`acumatica-client-secret`, `acumatica-password`, `sharepoint-cert-base64`,
  `sharepoint-cert-password`, `hubspot-access-token`) live in Key Vault; the Function App's
  **system-assigned managed identity** reads them via `@Microsoft.KeyVault(…)` app-setting references.
- App settings use **`__`** separators (env-var form). Deploy with `az functionapp deployment
  source config-zip` (publish → zip the output → deploy; the `az` `function list` view can lag, so
  verify with `functionapp function show`).
- Beyond the base settings: reconcile needs the **team GI** (`Acumatica__TeamGenericInquiryName` + team
  fields) and `ProjectSyncReconcileSchedule` / `ProjectSyncFullReconcileSchedule`; HubSpot needs
  `HubSpot__AccessToken` (KV ref), `HubSpotScopingSchedule`, `HubSpot__PracticeProperty`,
  `HubSpot__IncludedPractices__0`, `HubSpot__TerminalStageIds__*`, and `HubSpot__CreatedAfter`
  (go-live floor); uploads need `SharePoint__CreateClientUploadLink` / `ClientUploadLinkScope`.
- **Failure alert**: an App Insights log alert (`ProjectSync-Errors`) emails jparks@marshall-stevens.com
  when the function logs an Error/exception in a 15-minute window.

> **Why Windows, not Linux Consumption:** Azure CLI 2.83 has a bug validating .NET 10 on Linux
> (`DOTNET-ISOLATED|10.0` vs its list of `10`), which breaks `config-zip`/`function list`, and Linux
> Consumption's SCM is unreachable when idle. Windows Consumption avoids both; the cert uses the
> base64/Key-Vault path so nothing depends on the OS cert store.

[Azure Functions Core Tools]: https://learn.microsoft.com/azure/azure-functions/functions-run-local
