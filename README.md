# Translation Assistant

Microsoft Word task-pane add-in for propagating tracked changes from a selected
English, French, or German reference column into the other two language cells.

## Deploy to Vercel

Vercel replaces both the production Docker containers and Nginx. The root
[`vercel.json`](vercel.json) deploys the Vite frontend and FastAPI backend as
two Vercel services, routes `/api/*` and `/manifest.xml` to FastAPI, and routes
`/taskpane/` explicitly to the Word UI before routing the remaining URLs to
Vite. Keep the Vercel project's Root Directory set to
the repository root. Under **Build and Deployment**, set **Framework Preset**
to **Services**; Vercel will otherwise auto-detect only FastAPI and `/` will
return its platform `404: NOT_FOUND`. Do not configure a custom build command
or output folder.

In **Vercel project settings → Environment Variables**, define every variable
from `backend/.env.example` for the relevant Production and Preview
environments. In production, set:

```text
PUBLIC_BASE_URL=https://your-stable-production-domain.example
```

`PUBLIC_BASE_URL` must be the stable HTTPS origin from which users will install
and run the add-in—not a short-lived deployment preview URL. Redeploy after
changing an environment variable. The production manifest is then available
at `https://your-stable-production-domain.example/manifest.xml`.

The source-column and language assignment are stored in the user's task-pane
browser storage and are included in each translation request. No mutable
configuration is kept in a Vercel function instance.

With Fluid Compute enabled, Vercel currently defaults Python functions to a
300-second duration. The task pane stops waiting after 180 seconds, so
especially large rows can still time out and should eventually be moved to a
queued job architecture. Function duration can be changed under the project's
**Settings → Functions** when the selected Vercel plan supports it.

## Local production-shaped environment

Prerequisites:

- Docker with Compose;
- trusted Office development certificates in
  `/Users/octavian/.office-addin-dev-certs` or a custom directory supplied as
  `OFFICE_CERTS_DIR`;
- `backend/.env.dev` containing every variable listed in
  `backend/.env.example`. Copy the example when creating another environment;
  the backend refuses to start if any entry is missing or invalid.

Start the application:

```sh
docker compose up -d --build
```

Local routes:

- installation page: `https://localhost:3000/`
- Word task pane: `https://localhost:3000/taskpane/`
- downloadable manifest: `https://localhost:3000/manifest.xml`
- proxied health check: `https://localhost:3000/api/health`

The first time after switching from the old single-page development setup,
remove the old add-in registration and upload the new `manifest.xml`. Its source
location is now `/taskpane/` rather than `/index.html`.

The local manifest intentionally has a different Office add-in ID and the name
**Translation Assistant (Local)**. This prevents Word from confusing it with
the production Vercel manifest, whose ID must remain stable for deployed-user
updates.

The frontend image is a multi-stage Vite/Nginx build. Nginx is the only public
container and proxies `/api/` to the private FastAPI service. FastAPI runs
without development reload and isn't exposed on a host port.

## Verification

```sh
cd frontend
npm run typecheck
npm run build

cd ../
docker compose exec -T backend python -m unittest discover -s tests -v
```

Before exposing the translation API beyond a controlled pilot, add
authentication and request/rate limits; possession of the manifest alone is
not an authorization mechanism.
