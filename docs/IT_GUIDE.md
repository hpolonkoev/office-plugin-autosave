# AutoSave Office Add-in — IT & Developer Guide

A Microsoft Office Web Add-in that automatically saves Word, Excel, and PowerPoint documents at a configurable interval. Once installed, autosave starts silently in the background every time a document is opened — no user interaction required.

---

## How It Works

- On document open, the add-in runtime loads automatically via the `OnDocumentOpen` event.
- A repeating timer calls the native Office save API for the host application (Word, Excel, or PowerPoint).
- Before saving, the add-in checks that the document has been saved to disk at least once. If it has not, a warning is shown in the task pane.
- After each successful save, the "Last saved" timestamp in the task pane is updated.
- The task pane (accessible via the "Open AutoSave" button in the Home ribbon) is only needed to change settings. Autosave runs whether the pane is open or closed.

### Architecture

| Component | Description |
|---|---|
| `manifest.xml` | Office add-in manifest — declares hosts, ribbon button, and the `OnDocumentOpen` event |
| `src/taskpane/taskpane.ts` | Main logic: timer, save calls, UI, settings |
| `src/i18n/i18n.ts` | Lightweight i18n module with locale fallback chain |
| `config/autosave-config.json` | IT admin configuration — intervals and user override policy |
| `locales/en.json` / `fr.json` / `nl.json` | Locale strings — new languages can be added without rebuilding |

### Shared Runtime

The manifest uses a **shared runtime** (`<Runtime lifetime="long">`). This keeps the JavaScript engine alive even when the task pane is closed, so the autosave timer continues running in the background.

### Locale Detection

On startup, the add-in reads `Office.context.displayLanguage` (e.g. `fr-BE`) and loads the matching locale file. Fallback chain: exact locale → base language → English.

---

## Local Development

### Prerequisites

- Node.js 18 or later
- A Microsoft 365 subscription (for testing in Office desktop)

### Setup

```bash
npm install
npm start
```

`npm start` installs a trusted development certificate and starts the webpack dev server at `https://localhost:3000`.

### Sideload via Trusted Catalog (Windows)

1. Create a folder, e.g. `C:\OfficeAddins\Autosave`, and share it (right-click → Properties → Sharing → Advanced Sharing). Note the share path, e.g. `\\YOURPC\Autosave`.
2. Copy `manifest.xml` into that folder.
3. In Word/Excel/PowerPoint: **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**.
4. Add the share path (`\\YOURPC\Autosave`) and tick **Show in Menu**.
5. Restart Office, then go to **Insert → Add-ins → My Add-ins → SHARED FOLDER → Refresh** and add the add-in.

> **Note:** The "This add-in could not be started" popup on document open is a local development artifact caused by the self-signed dev certificate not being trusted in background WebView2 contexts. Click **Cancel** to dismiss it and open the task pane manually the first time. This error does not occur on a production HTTPS deployment.

---

## Production Deployment

### 1. Build

```bash
npm run build
```

Output is in `dist/`. It contains `taskpane.html`, the JS bundle, `config/autosave-config.json`, `locales/`, and all assets.

### 2. Host the Static Files

The `dist/` folder must be served over **HTTPS**. Any static hosting provider works.

#### Option A — GitHub Pages (free) — current setup

This project uses GitHub Actions to build and deploy automatically. **`dist/` is intentionally excluded from git** (it is in `.gitignore`) because the build happens on GitHub's servers, not your machine.

How it works:
1. You push source code (`src/`, `config/`, `locales/`, etc.) to the `main` branch.
2. GitHub Actions runs `npm run build`, producing `dist/` on GitHub's servers.
3. The workflow uploads `dist/` directly to GitHub Pages storage — it never touches the git repository.
4. The live site is served from `https://hpolonkoev.github.io/office-plugin-autosave/`.

> **Never commit `dist/` to git.** It is a build artifact. If you need to check what was deployed, look at the GitHub Actions run log, not the repository.

#### Option B — Azure Static Web Apps (free tier)

1. In the Azure Portal, create a **Static Web App** and connect it to your GitHub repository.
2. Set the build output folder to `dist`.
3. Azure assigns a URL like `https://<name>.azurestaticapps.net`.

#### Option C — Internal Web Server (on-premises / VDI environments)

Host the `dist/` folder on any IIS, Nginx, or Apache server accessible to your users over HTTPS. A wildcard or SAN certificate covering the server hostname is required.

### 3. Update the Manifest

Replace every occurrence of `https://localhost:3000` in `manifest.xml` with your production base URL.

```xml
<!-- Example -->
<IconUrl DefaultValue="https://your-domain.com/assets/icon-32.png" />
<bt:Url id="Taskpane.Url" DefaultValue="https://your-domain.com/taskpane.html" />
```

There are approximately eight URLs to update. Search for `localhost:3000` to find them all.

---

## Deploying to Users

### Option 1 — Microsoft 365 Centralized Deployment (Recommended)

Centralized Deployment pushes the add-in to all selected users automatically. Users get the ribbon button without doing anything, and autosave starts on every document open from day one.

1. Sign in to the **Microsoft 365 Admin Center** (`admin.microsoft.com`).
2. Go to **Settings → Integrated apps → Upload custom apps**.
3. Upload your updated `manifest.xml`.
4. Under **Assign Users**, choose specific users, groups, or the entire organisation.
5. Click **Deploy**. The add-in appears in Office for targeted users within 24 hours (usually faster).

**Advantages:** Zero user interaction, works on all machines and devices where the user signs in with their M365 account, survives Office reinstalls and machine changes.

### Option 2 — Trusted Catalog via Group Policy (Domain / VDI)

Use this when you cannot use Centralized Deployment (e.g. on-premises Exchange, no M365 Admin access, or non-M365 licences).

#### Step A — Host the manifest

Place `manifest.xml` on an internal HTTPS server or a network share accessible from all machines. Example network share: `\\fileserver\OfficeAddins\Autosave\manifest.xml`.

#### Step B — Push the trusted catalog via Group Policy

1. Open **Group Policy Management** on a domain controller.
2. Create or edit a GPO linked to the OU containing your users or VDI machines.
3. Navigate to:
   `User Configuration → Preferences → Windows Settings → Registry`
4. Create the following registry values (repeat for each value):

| Key path | Value name | Type | Data |
|---|---|---|---|
| `HKCU\SOFTWARE\Microsoft\Office\16.0\WEF\TrustedCatalogs\{YOUR-GUID}` | `Url` | REG_SZ | `\\fileserver\OfficeAddins\Autosave` or `https://your-domain.com` |
| same key | `Flags` | REG_DWORD | `1` |

Replace `{YOUR-GUID}` with any new GUID (generate one at `[Guid]::NewGuid()` in PowerShell).

5. Apply the GPO. On next login, users will see the add-in in the Shared Folder catalog. They add it once; after that it persists.

#### Alternative: Registry push via PowerShell login script

```powershell
$guid = "{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}"
$path = "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\TrustedCatalogs\$guid"
if (-not (Test-Path $path)) {
    New-Item -Path $path -Force | Out-Null
    Set-ItemProperty -Path $path -Name "Url"   -Value "\\fileserver\OfficeAddins\Autosave"
    Set-ItemProperty -Path $path -Name "Flags" -Type DWord -Value 1
}
```

Run this script at user login via GPO or VDI session startup script.

### VDI-Specific Notes

| Scenario | Recommendation |
|---|---|
| **Citrix / VMware Horizon / Azure Virtual Desktop with persistent desktops** | Use GPO registry push (Option 2) or Centralized Deployment (Option 1). Both work identically to physical machines. |
| **Non-persistent (pooled) VDI images** | Use **Centralized Deployment** (Option 1) — it is tied to the user's M365 account, not the machine, so it survives image resets. |
| **Non-persistent VDI without M365 Centralized Deployment** | Bake the registry key into the golden image (see PowerShell script above) and host the manifest on an always-available internal HTTPS server. Users still need to add the add-in from the Shared Folder once per profile, but on non-persistent VDI this means once per session. Consider combining with a login script that auto-accepts the catalog. |
| **Static files accessible from VDI** | Ensure the `dist/` web server is reachable from VDI machines. If VDI machines have no internet access, host the files on an internal server (Option C above). |
| **Office version on VDI** | The `OnDocumentOpen` event and shared runtime require Office 2019 / Microsoft 365 Apps version 2009 or later. Verify with `File → Account → About Word/Excel/PowerPoint`. |

---

## IT Configuration Guide

### Changing the Autosave Interval

Edit `config/autosave-config.json` on the web server (no rebuild required):

```json
{
  "defaultIntervalMinutes": 5,
  "minIntervalMinutes": 2,
  "maxIntervalMinutes": 60,
  "allowUserOverride": true
}
```

| Field | Description |
|---|---|
| `defaultIntervalMinutes` | Interval used when no user preference has been saved |
| `minIntervalMinutes` | Minimum interval users can select |
| `maxIntervalMinutes` | Maximum interval users can select |
| `allowUserOverride` | `true` shows the settings panel; `false` hides it and shows a "managed by IT" message |

### Disabling User Overrides

Set `"allowUserOverride": false`. The interval controls are hidden; the add-in always uses `defaultIntervalMinutes`.

### Adding a New Language

1. Create `locales/xx.json` (where `xx` is the BCP 47 language tag, e.g. `de`, `es`, `pt-BR`) on the web server.
2. Copy all keys from `locales/en.json` and translate the values.
3. No rebuild or manifest change is needed. The add-in will automatically pick up the new file for users whose `Office.context.displayLanguage` matches.

Required keys (all must be present):

```
addin_name, activated_msg, toggle_on, toggle_off, last_saved, never_saved,
interval_label, interval_unit, save_settings, settings_saved,
unsaved_document, save_error, managed_by_it, protected_view
```

### After Any Config or Locale Change

Simply save the updated file on the web server. The add-in fetches `config/autosave-config.json` and the locale file fresh on every document open — no rebuild, no manifest update, no redeployment needed.

---

## Versioning & Release Checklist

The version follows **Semantic Versioning (SemVer)** and lives in two files that must always stay in sync:

| File | Field |
|---|---|
| `package.json` | `"version": "1.0.0"` |
| `manifest.xml` | `<Version>1.0.0</Version>` |

Microsoft 365 reads the manifest version to decide when to push an update to users — so keeping them in sync matters.

### When to bump the version

| What changed | Version bump | Example |
|---|---|---|
| Bug fix, typo in a locale string, small tweak | **Patch** | `1.0.0 → 1.0.1` |
| New feature, new config option, new language added | **Minor** | `1.0.0 → 1.1.0` |
| Manifest structure changed, breaking IT configuration | **Major** | `1.0.0 → 2.0.0` |

> **Not every commit needs a version bump.** Only bump when the change is meaningful enough to be called a release.

### How to release a new version

```bash
# 1. Bump package.json automatically (choose one)
npm version patch   # bug fix  → 1.0.0 to 1.0.1
npm version minor   # new feature → 1.0.0 to 1.1.0
npm version major   # breaking change → 1.0.0 to 2.0.0
```

`npm version` updates `package.json` and creates a git commit automatically.

```bash
# 2. Open manifest.xml and update <Version> to match package.json
#    Search for:  <Version>
#    Change to:   <Version>1.0.1</Version>  (whatever npm version set)

# 3. Commit the manifest change
git add manifest.xml
git commit -m "chore: sync manifest version to 1.0.1"

# 4. Push — GitHub Actions builds and deploys automatically
git push
```

The version badge in the README updates automatically once the push lands on GitHub.

### Quick reference — what triggers what

| Action | Triggers |
|---|---|
| `git push` to `main` (any change) | GitHub Actions rebuilds and redeploys the site |
| `npm version patch/minor/major` + manifest sync + push | All of the above + version badge updates |
| Edit `config/autosave-config.json` on the server directly | Takes effect on next document open — no push needed |
| Drop a new `locales/xx.json` on the server | Takes effect on next document open — no push needed |
