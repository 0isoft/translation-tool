# Local Word integration

The add-in has two separate pieces:

- `manifest.xml` is installed into Word. It identifies the add-in, limits it to
  Word, grants document access, and tells Word to load the task pane from
  `https://localhost:3000/index.html`.
- The Vite server provides that task pane page. Office.js is loaded by the page
  from Microsoft's CDN and gives the page access to the current Word document.

## Run the local services

From the project root:

```bash
docker compose up
```

Keep this terminal running while using the add-in. Before opening Word, verify
that `https://localhost:3000` opens without a certificate warning.

## Sideload into Word on macOS

1. Close Word.
2. In Finder, press **Command+Shift+G**.
3. Open this folder (create the final `wef` folder if needed):

   ```text
   ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
   ```

4. Copy `manifest.xml` from this project into that `wef` folder.
5. Reopen Word and open a document.
6. Choose **Home > Add-ins**, then select **Translation Assistant**.

Word reads the local copy of the manifest, but the task pane itself is served
live by Vite. Frontend edits therefore update through Vite without reinstalling
the manifest. Recopy the manifest and restart Word when manifest values change.

## Check the manifest

From `frontend`:

```bash
npm run validate:manifest
```
