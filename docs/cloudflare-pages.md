# Cloudflare Pages Deployment

**English** | [简体中文](cloudflare-pages.zh-CN.md)

Connect the GitHub repository to Cloudflare Pages, and every push to the production branch runs `npm run build` and publishes `dist/` as an HTTPS site. After users sideload the generated `manifest.xml`, Word loads the task pane from that domain.

## 1. Fork the repository

1. Open [siyuhong/Mermaid-Word-AddIn](https://github.com/siyuhong/Mermaid-Word-AddIn).
2. Click **Fork** in the upper-right corner to create a copy.

## 2. Create a Cloudflare Pages project

1. Sign in to the [Cloudflare dashboard](https://dash.cloudflare.com/).
2. Go to **Workers & Pages → Create application → Pages → Connect to Git**.
3. Authorize Cloudflare to access GitHub and select the fork you created in step 1.
4. Choose the production branch, for example `main`.

## 3. Configure the build

Fill in the build settings as follows:

| Field                  | Value                                            |
| ---------------------- | ------------------------------------------------ |
| Framework preset       | `None`                                           |
| Build command          | `npm run build`                                  |
| Build output directory | `dist`                                           |
| Root directory         | Leave empty to use the repository root directory |

Build environment variables:

| Variable         | Value                                                                                       |
| ---------------- | ------------------------------------------------------------------------------------------- |
| `ADDIN_BASE_URL` | The project's fixed production URL, for example `https://your-mermaid-word-demo.pages.dev/` |
| `NODE_VERSION`   | `22` (or another Node LTS version verified with this project)                               |

Notes:

- The production build replaces every `https://localhost:3000` and the corresponding `<AppDomain>` in `manifest.xml` with `ADDIN_BASE_URL`. If the variable is missing or the URL is invalid, `npm run build` fails immediately.
- `ADDIN_BASE_URL` must be an HTTPS URL and must not contain a username, password, query string, or fragment. Do not include a path such as `/taskpane.html`; a trailing slash is optional.
- Variables are scoped by environment. If you configure it only for Production, Preview (PR) deployments fail because `ADDIN_BASE_URL` is missing. Configure the Preview environment separately if you need PR previews.

## 4. Deploy

1. Click **Save and Deploy** and wait for dependency installation, the build, and the upload to finish.
2. Check the production domain Cloudflare assigned, for example `https://your-mermaid-word-demo.pages.dev`.
3. If the domain differs from `ADDIN_BASE_URL`, update the Production environment variable under **Settings → Environment variables** and redeploy.

Changing an environment variable does not regenerate an existing manifest; you must rebuild. The same applies after binding a custom domain: update `ADDIN_BASE_URL`, rebuild, and then ask users to download the manifest again.

## 5. Verify the deployment

### 5.1 Browser checks

Visit the following paths on the actual production domain:

| Path                  | Expected result                                                                        |
| --------------------- | -------------------------------------------------------------------------------------- |
| `/`                   | Shows installation instructions and the `manifest.xml` download link                   |
| `/manifest.xml`       | Downloads an XML file whose URLs point to the production site rather than `localhost`  |
| `/taskpane.html`      | The page and scripts load. Office.js functionality still needs to be verified in Word. |
| `/commands.html`      | The page loads; a blank page is expected.                                              |
| `/assets/icon-32.png` | Displays the icon.                                                                     |
| `/does-not-exist`     | Returns `404.html`.                                                                    |

If the last check returns the home page or a 200 response, verify that `public/404.html` is copied to `dist/` and that no SPA fallback or rewrite rule is intercepting the request.

### 5.2 Verify in Word

1. Download the manifest from `https://your-mermaid-word-demo.pages.dev/` (or `/manifest.xml`).
2. **Word on the web**: open an editable document, go to **Home → Add-ins → More Settings → Upload My Add-in**, and select the manifest.
3. **Windows desktop**: put the manifest in a shared folder, add the folder's UNC path under **Trust Center → Trusted Add-in Catalogs**, restart Word, and install the add-in from **Shared Folder**.
4. Open **Mermaid Editor** from the **Insert** tab in Word.
5. Verify the four main paths:
   - Modify the Mermaid code and confirm the editor works.
   - Confirm the live preview updates as you type.
   - Click **Insert | Replace** to insert the diagram into the document.
   - Select the inserted image, click **Edit Selected** to restore the source, modify it, and replace the image again.

## 6. Troubleshooting

- **Build error `Set ADDIN_BASE_URL ...`**: The Production environment variable is not configured, or it was configured only for Preview.
- **Build error `ADDIN_BASE_URL must be a deployed HTTPS URL ...`**: The URL is not HTTPS, contains a query string, fragment, username, or password, or still points to `localhost`.
- **Blank task pane or certificate error in Word**: Confirm that you are using the production domain rather than `localhost:3000`, and check that the URLs in the manifest match the current domain.
- **Word still shows an old version after redeployment**: `public/_headers` sets `no-cache` for all resources, but Word may still cache the manifest. Upload the manifest again or restart Word; if necessary, clear the WebView2 cache.
- **The task pane cannot be embedded in Word on the web**: Do not add `X-Frame-Options` to the site, and do not restrict `frame-ancestors` with a CSP header.

## Official references

- [Cloudflare Pages Git integration](https://developers.cloudflare.com/pages/get-started/git-integration/)
- [Cloudflare Pages build configuration](https://developers.cloudflare.com/pages/configuration/build-configuration/)
- [Cloudflare Pages headers](https://developers.cloudflare.com/pages/configuration/headers/)
- [Sideload Office Add-ins in Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)
- [Sideload add-ins on Windows from a network share](https://learn.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)
