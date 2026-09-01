# Translation Tool Architecture

## What this application is

This is a Microsoft Word task-pane add-in. It inspects tracked changes in the
configured English table cell, asks Claude to propagate semantic changes into
the German and French cells, and writes the result back through Office.js with
Track Changes enabled.

There is no database. The project does not use SQLite, SQLAlchemy, an ORM,
migrations, Redis, or browser local storage.

## Boundaries

```text
Word document
  tracked revisions + add-in baseline setting
          |
          v
Office.js task pane (frontend)
  inspect -> classify -> call backend -> validate stale state -> write edits
          |
          | HTTPS /api/* (Vite proxy)
          v
FastAPI transport (backend/app/main.py)
          |
          v
Translation application service
  deterministic numeric rules + Claude planning port
          |
          v
Anthropic adapter -> Claude Messages API
```

## Backend layout

- `app/main.py`: FastAPI routes, CORS, dependency construction, and mapping
  application errors to HTTP errors. It contains no translation algorithm.
- `app/application/translation_service.py`: the `propagate_cell_changes` use
  case. It classifies changes, combines deterministic and Claude results, and
  returns partial-failure information.
- `app/application/configuration.py`: process-local language configuration.
- `app/domain/models.py`: request, response, and Claude structured-output
  contracts.
- `app/domain/translation_rules.py`: pure paragraph classification, numeric
  propagation, plan normalization, safety checks, and numeric warnings.
- `app/domain/prompts.py`: the complete Claude system prompts.
- `app/ports/claude.py`: outbound Claude-planning interface.
- `app/adapters/anthropic.py`: Anthropic SDK, environment configuration,
  structured parsing, exponential retry, and jitter.

## Frontend layout

- `src/taskpane.ts`: UI event wiring and the top-level Word propagation use
  case. It coordinates adapters but does not contain HTTP or mutation-writing
  implementations.
- `src/domain/models.ts`: frontend contracts matching the backend API.
- `src/domain/textRules.ts`: blank-paragraph, whitespace, whole-addition, and
  minimal-span-diff rules.
- `src/adapters/translationApi.ts`: every browser `fetch` call.
- `src/adapters/wordRevisionStore.ts`: revision fingerprints and the embedded
  document baseline.
- `src/adapters/wordMutationWriter.ts`: minimal replacements, paragraph
  insertions, and Track Changes mode management.

## State and persistence

| State | Owner | Lifetime |
| --- | --- | --- |
| Document text and tracked revisions | Word DOCX | Saved with the document |
| Revision baseline | Word document setting `translationTool.revisionBaseline.v1` | Saved with the document and copied with it |
| Source column and three language assignments | Backend in-memory configuration | Lost/reset on backend restart |
| Claude API key, model, timeout, retry settings | `backend/.env.dev` | External configuration |
| Inspected rows, planned mutations, progress counters | Task-pane JavaScript memory | One button invocation |

The baseline is a JSON object containing counts of SHA-256 fingerprints. A
fingerprint hashes tracked-change author, date, type, and text. It is not a copy
of document text. Completed revisions enter the baseline; failed revisions are
subtracted so the next run sees them again.

The **Retry all tracked changes** button ignores the embedded baseline for that
execution, then replaces it with a new success/failure-aware baseline. It does
not accept/reject revisions or directly modify source text.

## HTTP calls

The frontend uses the native browser `fetch` API. Axios is not installed or
used. All calls are centralized in `src/adapters/translationApi.ts`.

1. `GET /api/config`
   - Vite proxies this to backend `GET /config`.
   - Returns source column and English/French/German assignments.
2. `PUT /api/config`
   - Replaces the in-memory language configuration.
   - This is not persisted across backend restarts.
3. `POST /api/translate-cell-changes`
   - One request per table row containing newly detected source revisions.
   - Timeout: 180 seconds in the task pane.
   - Header: `X-Request-ID` for log correlation.

`PUT /config/source-column` remains as a backwards-compatible backend route;
the current frontend does not call it.

## Translation request contract

Each request contains:

- `source_column`: logical language column 1, 2, or 3.
- `source_cell`: every meaningful current paragraph in the English cell.
- `changed_source_paragraphs`: only paragraphs containing revisions newer than
  the saved baseline, with original text, current text, and revision fragments.
- `targets`: the complete meaningful-paragraph lists for the two target cells,
  each with a fixed column number and expected language.

Blank Word paragraphs and cell markers are omitted from logical indices. The
whole document and unrelated table columns are never sent to Claude.

## Claude calls

There are two structured-output operations:

1. New-paragraph planner
   - Called once per wholly added English paragraph.
   - Receives the full three cells as context and the one new paragraph.
   - Returns `insert` or `none` independently for German and French, an index,
     and translated text.
   - `none` prevents duplicates during baseline recovery.
2. Existing-paragraph planner
   - Called once for all remaining changed paragraphs in that cell.
   - Returns `replace`, `insert`, or `none` edits for each target.
   - Every source paragraph must be accounted for; incomplete plans are
     rejected and retried.

The exact prompts are constants in `backend/app/domain/prompts.py`; there are no
hidden prompt fragments elsewhere.

Claude output is parsed directly into the Pydantic schemas in
`backend/app/domain/models.py`. Default retry behavior is three application
attempts with exponential backoff and jitter. SDK-level retry is separately
controlled by `ANTHROPIC_MAX_RETRIES`.

## Deterministic path

Numeric-only changes bypass Claude when both target cells can be matched safely.
Matching uses old numeric signatures plus unchanged numeric context. It never
replaces the "next number" by ordinal position. If either target is ambiguous,
both fall back to Claude.

## Word propagation sequence

1. Load backend configuration.
2. Read table row cell counts.
3. Choose physical language columns:
   - seven-or-more-cell source tables: physical columns 1/2/3;
   - compact six-cell test tables: physical columns 0/1/2.
4. Load paragraphs, reviewed original/current text, and tracked changes in
   batches of ten rows. Read revisions at paragraph and cell scope. If desktop
   Word omits a wholly inserted paragraph from those collections, inspect the
   cell OOXML for complete paragraphs contained in `w:ins` and map them only by
   exact normalized paragraph text.
5. Compare change fingerprints to the embedded baseline.
6. Ignore whitespace-only revisions and fully deleted source paragraphs.
7. Send one backend request for each affected row.
8. Verify returned target indices against the still-current Word text.
9. Enable Track Changes if necessary.
10. Apply the smallest unique changed span when possible; otherwise replace the
    paragraph. Insert wholly new paragraphs before/after semantic anchors.
11. Restore the user's original Track Changes mode.
12. Save a new baseline containing successful revisions while leaving failures
    pending.

## Network topology

- Word opens `https://localhost:3000/index.html` from `manifest.xml`.
- Vite serves the task pane over the Office development certificate.
- Vite proxies `/api` to `http://backend:8000` inside Docker Compose.
- FastAPI calls Anthropic over HTTPS using `ANTHROPIC_API_KEY`.
