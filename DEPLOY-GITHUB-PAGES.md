# Host the "Follow Up" add-in on GitHub Pages

Outlook only loads an add-in from **trusted HTTPS**. `localhost` fails to connect
(self-signed cert), which is why hosting on **GitHub Pages** is the reliable path.
These files are already flattened for Pages — every URL sits at the repo root.

## Steps

### 1. Put your Flow URL in place (already done)
`config.js` already contains the live `FLOW_URL` for **PP FollowUp - Create**.
Nothing to change unless you redeploy the flow.

### 2. Set the manifest URL to your repo
Decide a repo name (e.g. `followup-addin`). Then, inside this folder:

```bash
./set-url.sh <github-username> <repo-name>
# example: ./set-url.sh shubich followup-addin
```

This rewrites every `__BASEURL__` in `manifest.xml` to
`https://<github-username>.github.io/<repo-name>`.
(If you're on Windows without bash, just open `manifest.xml` and replace every
`__BASEURL__` with that same URL by hand.)

### 3. Push to GitHub
```bash
git init
git add .
git commit -m "Follow Up Outlook add-in"
git branch -M main
git remote add origin https://github.com/<github-username>/<repo-name>.git
git push -u origin main
```

### 4. Enable GitHub Pages
On GitHub: **repo → Settings → Pages →** Source **Deploy from a branch**,
Branch **main**, Folder **/(root)** → **Save**.
Wait ~1 minute, then open `https://<github-username>.github.io/<repo-name>/` —
you should see the "Follow Up — Outlook Add-in" landing page. That confirms HTTPS
is serving.

### 5. Sideload the manifest into Outlook
- **Outlook (web & new/desktop):** **Get Add-ins → My add-ins →
  Custom Addins → Add a custom add-in → Add from file** → pick `manifest.xml`.
- **Org-wide:** Microsoft 365 admin center → **Settings → Integrated apps →
  Upload custom apps** → upload `manifest.xml` → assign users.

Open a **sent** message → the **Follow Up** button appears on the ribbon → click it.

## Updating later
Edit any file, `git commit`, `git push`. Pages redeploys automatically.
If you changed a URL/icon path in `manifest.xml`, re-sideload the manifest;
changes to `taskpane.*` / `config.js` take effect on next open (you may need to
clear the Office cache or bump the manifest `<Version>` to force a refresh).

## Notes
- `.nojekyll` is included so GitHub serves the files verbatim (no Jekyll build).
- Generate a fresh `<Id>` GUID in `manifest.xml` if you ever publish a second,
  separate copy — two add-ins must not share the same Id.
- `config.js` holds the flow URL **with its signature** — keep the repo
  **private** if you don't want that endpoint public, or add the shared-secret
  check described in the main SETUP-GUIDE.md. (A private repo can still serve
  Pages on paid GitHub plans; otherwise use a private host.)
