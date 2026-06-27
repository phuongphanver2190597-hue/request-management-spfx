# Request Management SPFx Web Part

## Project Overview
SharePoint Framework (SPFx) web part for managing requests in SharePoint Online.
- Dashboard with summary cards (Total/Pending/Approved/Rejected)
- Request list with search, filter, refresh
- New Request form (modal, centered)
- View Detail modal (centered)
- Items filtered by current user's AssignedTo email (server-side, no admin bypass)

## Tech Stack
- SPFx: 1.23.1
- React: 17.0.1
- TypeScript: 5.3.3 (via `@microsoft/rush-stack-compiler-5.3`)
- Fluent UI: `@fluentui/react` v8 (^8.125.6) — NOT v9
- Node.js: **must use v22** (Node 24 is incompatible with SPFx 1.23.1)

## Node.js Setup (REQUIRED before every build)
System has Node 24. Must switch to Node 22 via fnm every session:
```powershell
$env:PATH = [System.Environment]::GetEnvironmentVariable("PATH","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH","User")
fnm env --shell powershell | Out-String | Invoke-Expression
fnm use 22
node --version  # must show v22.x.x
```

Install fnm if not present: `winget install Schniz.fnm`, then `fnm install 22`.

## Build Commands
```powershell
# Always switch to Node 22 first (see above), then:
gulp bundle --ship
gulp package-solution --ship
```
Output: `sharepoint/solution/request-management.sppkg`

## Deploy to SharePoint
1. Upload `sharepoint/solution/request-management.sppkg` to App Catalog
2. Tick **"Make this solution available to all sites"** (tenant-wide deployment)
3. Click Deploy
4. On the SharePoint page: Ctrl+F5 to refresh

## SharePoint List
- **List name**: `Requests` (configurable in web part property pane)
- **Required columns**:
  | Column | Type | Notes |
  |--------|------|-------|
  | Title | Single line text | Default SP column |
  | RequestCode | Single line text | |
  | Status | Choice | Options: Pending, Approved, Rejected |
  | AssignedTo | Person or Group | Users only, single select |
- Created / Author are auto-populated by SharePoint

## Key Implementation Details

### Data Fetching
- Uses `SPHttpClient` (no backend, no Graph API needed)
- Always filters by current user: `$filter=AssignedTo/EMail eq '${currentUserEmail}'`
- Expands `Author` and `AssignedTo` fields in the same query

### People Picker
- Custom `NormalPeoplePicker` (Fluent UI v8)
- Search endpoint: `/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`
- Resolve user before POST: `/_api/web/ensureuser` → returns SharePoint user ID
- POST uses `AssignedToId: userId` (not email string)

### Known Issues & Fixes
- `skipFeatureDeployment: true` requires a **non-empty** `features` array in `package-solution.json` — empty array causes `Reduce of empty array` build error
- `@pnp/spfx-controls-react` v3 is incompatible (PostCSS 8 conflict) — do NOT add it
- `node_modules` is gitignored — run `npm install` after cloning

## Project Structure
```
src/webparts/requestManagement/
  RequestManagementWebPart.ts       # Web part entry point, property pane
  components/
    RequestManagement.tsx           # Main component (all UI logic)
    IRequestManagementProps.ts      # Props interface
    IRequestManagementState.ts      # State + IRequest interfaces
config/
  package-solution.json             # Must have non-empty features[] array
sharepoint/solution/
  request-management.sppkg         # Deploy this file to App Catalog
```

## Git Remote
```
https://github.com/phuongphanver2190597-hue/request-management-spfx.git
```
