# Ehealth Innovations – central e-mail signature (Outlook add-in)

The add-in inserts the company signature into every new email, reply and forward in Outlook (web, new and classic Outlook for Windows, Mac, iOS, Android). There's no backend. The template is in `src/template.js`, and the user's details come from their Entra ID profile through Microsoft Graph (`/me`, delegated `User.Read`).

```
config.json             ← host, clientId, tenantId (fill these in)
manifest.template.xml   → dist/manifest.xml
src/launchevent.js      event handlers (NAA → Graph → cache → setSignatureAsync)
src/template.js         signature HTML (table-based, inline styles)
src/commands.html       runtime page for web / new Outlook / Mac / mobile
assets/                 logo, colour band, spacer, add-in icons
build.mjs               bundles everything into dist/
```

What the signature shows: **name**, **job title**, **T** (mobile, or the business number if there's no mobile), **E** (the address the mail is sent from) and a LinkedIn link. Lines for empty fields are left out. Details are cached per user for 24 hours (`cacheHours`). The **Handtekening invoegen** button on the compose ribbon clears the cache and inserts the signature again.

---

## Test with your own account first (about 1 hour)

### 0. Fill in your own profile
Go to the M365 admin center → **Users → Active users → Jelle** and fill in **Job title** and **Mobile phone**. At the moment your job title is empty, so that line would be left out.

### 1. Create the Azure Static Web App
1. In the Azure portal, go to **Create a resource → Static Web App**.
2. Choose Plan **Free**, Region **West Europe** and Deployment source **Other**.
3. When it's created, copy the **URL** (e.g. `https://gentle-sea-0a1b2c.azurestaticapps.net`).
4. Go to **Overview → Manage deployment token** and copy the token.

### 2. Register the Entra app
1. Go to **Entra admin center → App registrations → New registration**.
   - Name: `Ehealth Signature Add-in`
   - Supported account types: **Single tenant**
   - Redirect URI: platform **Single-page application (SPA)**, value `brk-multihub://<your-swa-host>`
     (origin only, e.g. `brk-multihub://gentle-sea-0a1b2c.azurestaticapps.net`, with no `https://` and no path)
2. Go to **API permissions**. `Microsoft Graph → User.Read (Delegated)` is already there. Click **Grant admin consent for Ehealth Innovations**.
   The event handler runs without a UI, so it can't ask for consent itself.
3. From **Overview**, copy the **Application (client) ID** and the **Directory (tenant) ID**.

### 3. Configure, build and deploy (on your Mac)
```bash
cd ~/Sites/ehealth-signature-addin
# edit config.json: host, clientId, tenantId
npm install
npm run build          # creates dist/
npm run preview        # opens dist/preview.html – check the layout
npm run validate       # Microsoft's manifest validator
SWA_CLI_DEPLOYMENT_TOKEN=<token> npm run deploy
```
Check that these URLs load in a browser:
- `<host>/commands.html`
- `<host>/launchevent.js`
- `<host>/assets/logo.png`
- `<host>/.well-known/microsoft-officeaddins-allowed.json` (classic Outlook needs this one)

### 4. Install it for your account only
**Option A: sideload (quickest, a few minutes).**
Open Outlook on the web → https://aka.ms/olksideload → **My add-ins → Custom add-ins → Add a custom add-in → Add from file** → choose `dist/manifest.xml`.
It becomes available in Outlook on the web and new Outlook straight away. Restart classic Outlook, Mac and mobile to pick it up.

**Option B: admin deployment, for yourself only (closer to how the rollout will work).**
Go to M365 admin center → **Settings → Integrated apps → Upload custom apps → Office Add-in**, upload `dist/manifest.xml` and under **Assign users** choose **Specific users** → yourself.
It can take a few hours, and at most 24 hours, before it shows up.
> Remove the sideloaded copy (option A) before you use option B. Both use the same add-in ID.

### 5. Test checklist
| | New | Reply | Forward | Send as / shared mailbox |
|---|---|---|---|---|
| Outlook on the web | | | | |
| New Outlook (Windows) | | | | |
| Classic Outlook (Windows) | | | | |
| Outlook for Mac | | | | |
| iOS | | | | |
| Android | | | | |

For each cell, check:
- the signature appears on its own
- there's no duplicate from your own Outlook signature
- the job title and phone are filled in (if you only see name + email, Graph failed; see below)

Then send a test mail to Gmail and Apple Mail, and check it in dark mode too.

### Troubleshooting
- **Only name and email appear.** The Graph lookup failed and the add-in used its fallback. Check that the redirect URI matches the host exactly, that admin consent was granted, and that clientId and tenantId in `config.json` are right. Then rebuild, redeploy and click **Handtekening invoegen** (which clears the cache).
- **Changed profile details don't show up.** They're cached for up to 24 hours. Click **Handtekening invoegen**.
- **Nothing appears in classic Outlook.** Check the `.well-known` URL above. Then turn on [runtime logging](https://learn.microsoft.com/office/dev/add-ins/testing/runtime-logging); lines are prefixed `[ehi-signature]`.
- **The logo shows as text.** The recipient's mail client blocks remote images. The alt text "Ehealth Innovations" appears instead. This is normal for hosted images.
- **Mobile replies.** The signature is added, but you only see it once you expand the compose window to full screen. This is a Microsoft limitation.

---

## Before the org-wide rollout
1. Optional: add a custom domain, e.g. `signature.ehealthinnovations.nl` (SWA → Custom domains, CNAME). Then change `host`, add a second redirect URI `brk-multihub://signature.ehealthinnovations.nl`, bump `version` in `config.json`, rebuild and deploy.
2. Move `deploy.github-workflow.yml` to `.github/workflows/deploy.yml` (`mkdir -p .github/workflows && mv deploy.github-workflow.yml .github/workflows/deploy.yml`), push the project to GitHub and add the secret `AZURE_STATIC_WEB_APPS_API_TOKEN`. It then deploys on every merge to `main`.
3. Integrated apps: assign the add-in to a pilot group, and after 1–2 weeks to the whole organisation.
4. Let users know. They don't need to do anything, and their own Outlook signature won't be inserted any more. They can still edit the signature by hand in a message.

## Maintenance
- **Changing the template** (`src/template.js`): build and deploy. You don't need to upload a new manifest.
- **Changing the manifest**: bump `version`, build, deploy, then update the add-in in Integrated apps.
- **Changing someone's details**: edit their Entra ID profile. The change shows up within 24 hours.

## Security / ISO notes
- The add-in only has delegated `User.Read`, so it reads the signed-in user's own profile and nothing else.
- There are no servers or databases. Email content never leaves Microsoft 365. The static host serves only HTML, JS and images.
- Record the Static Web App and the app registration in the asset register, and put template changes through the change process.
