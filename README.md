# monthlyactivityishraf

Static Cloudflare-ready app for the attached inspection monthly activity workbook.

Files:
- `index.html`
- `styles.css`
- `app.js`
- `assets/template.xlsx`
- `worker.js`
- `wrangler.toml`
- `public/` for the static assets deployed by Wrangler

The app edits the bundled Excel template in the browser and downloads a finished workbook with the original print layout preserved.

## Shared phone/PC saving on Cloudflare

The browser keeps a local copy, but shared saving between devices needs the Cloudflare Worker in `worker.js` plus a KV namespace binding named `ACTIVITY_STORE`.

Wrangler deployment:
1. Create a KV namespace for the app.
2. Put that namespace id in `wrangler.toml` where it says `REPLACE_WITH_YOUR_KV_NAMESPACE_ID`.
3. Deploy the Worker.

Cloudflare dashboard deployment:
1. Upload/deploy the app with `worker.js`.
2. Add a KV namespace binding named `ACTIVITY_STORE`.
3. Redeploy. The app will use `/api/data` to sync the same saved data on phone and PC.
