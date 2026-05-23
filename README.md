# Excel Bullet Splitter

Static web app that opens `.xlsx` files in the browser, finds merged
"Nhận xét" cells in the configured sheets (default `DKQHT`, `NLPC`),
and splits the bullet points into one row each. Files are processed
entirely in the user's browser via JSZip; nothing is uploaded.

## Run locally

Open `index.html` in any modern browser, or serve the folder:

```powershell
npx serve .
```

## Deploy to GitHub Pages

1. Create a new GitHub repo and push this folder as the repo root:

   ```powershell
   git remote add origin https://github.com/<your-user>/<your-repo>.git
   git branch -M main
   git push -u origin main
   ```

2. In the GitHub repo's **Settings → Pages**:
   - **Source**: *Deploy from a branch*
   - **Branch**: `main` / `/ (root)`
   - Save.

3. Wait ~30 seconds; the site is live at
   `https://<your-user>.github.io/<your-repo>/`.

No build step is required — GitHub Pages serves the files as-is.

## Deploy to Vercel (alternative)

```powershell
npm i -g vercel
vercel
```

Answer the prompts: framework `Other`, no build command, output dir `.`.
Subsequent deploys: `vercel --prod`.

## Files

- `index.html` — UI: drop zone, file list, run button, ZIP download.
- `splitter.js` — browser splitter (surgical xlsx editor).
- `public/images/` — before/after preview screenshots shown in the UI.
- `vercel.json`, `package.json` — only used by Vercel / local Node testing.
  GitHub Pages ignores them.
