# Translation Assistant

Microsoft Word task-pane add-in for propagating English tracked changes into
French and German table cells.

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

For deployment, replace the localhost URLs and certificate with the production
domain, move secrets out of `.env.dev`, and add authentication before exposing
the translation API publicly.
