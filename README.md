# Natural Joy Dashboard — Vercel Deployment

This folder is ready to deploy to Vercel. It contains:

- `index.html`, `styles.css`, `app.js`, `logo.jpeg`, `vendor/plotly.min.js`: dashboard website files.
- `api/data.py`: a Vercel Python Function that downloads the current OneDrive Excel workbook and converts it to dashboard data.
- `vercel.json`: project configuration.
- `.env.example`: the required OneDrive environment-variable name.

## Important privacy note

Do **not** upload the Excel workbook to GitHub or to the website folder. The workbook stays in the author's OneDrive. Do **not** paste the OneDrive link in source code if your GitHub repository will be public; add it as a Vercel Environment Variable instead.

## Deploy using GitHub and Vercel

1. Create a GitHub repository, for example `natural-joy-dashboard`.
2. Upload the **contents of this folder** to the repository root. The repository root must contain `index.html`, `app.js`, `styles.css`, `logo.jpeg`, `vercel.json`, and the `api` and `vendor` folders.
3. On Vercel, choose **Add New > Project** and import the GitHub repository.
4. Before deploying, open **Environment Variables** and add:
   - Name: `ONEDRIVE_SHARE_URL`
   - Value: the author's OneDrive Excel shared link.
   - Apply to: Production, Preview, and Development.
5. Click **Deploy**.
6. Open the generated Vercel URL. The website reads the latest saved OneDrive workbook through `/api/data`.

## Data updates

- The author edits and saves the master Excel workbook in OneDrive.
- The dashboard checks the OneDrive workbook every minute while it is open; the user can also press `تحديث البيانات`.
- No Excel file is permanently downloaded into the website storage.

## OneDrive permission required for link-only access

This no-login connection works only when the OneDrive shared link allows viewing/downloading without signing in, such as **Anyone with the link can view**. Do not grant edit access to dashboard viewers. For confidential investor information, use Microsoft Graph authentication instead of a publicly readable share link.
