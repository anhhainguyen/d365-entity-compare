# D365 Entity Compare — Agent Instructions

A Chrome MV3 extension (plus local proxy fallback) that compares D365 Finance & Operations entity configuration between two environments.

---

## Project Structure

| File | Role |
|---|---|
| `manifest.json` | MV3 extension descriptor — permissions: `tabs`, `storage`, `scripting` |
| `popup.html` | Entire UI, inline CSS; all DOM IDs listed below are coupled to `popup.js` |
| `popup.js` | All logic (~2 200 lines), single IIFE; **no build step, no framework** |
| `content.js` | Fetch proxy running inside D365 tabs (uses session cookies) |
| `background.js` | MV3 service worker — handles `OPEN_FULL_TAB` message only |
| `server.ps1` | PowerShell HTTP server (`localhost:8888`) + `/proxy` endpoint |
| `Launch-Tool.bat` | Starts `server.ps1` then opens Chrome at `http://localhost:8888/popup.html` |

---

## Two Runtime Modes

`popup.js` detects the mode at startup:

```js
var IS_EXT  = typeof chrome !== 'undefined' && !!chrome.runtime && !!chrome.runtime.id;
var IS_FILE = window.location.protocol === 'file:';
```

| Mode | When | Auth | D365 calls go through |
|---|---|---|---|
| **Extension** (`IS_EXT=true`) | Loaded from `chrome-extension://` | Session cookies (no token needed) | `content.js` via `chrome.tabs.sendMessage` |
| **Proxy** (`IS_EXT=false`) | Served by `server.ps1` at `localhost:8888` | Bearer token entered by user | `http://localhost:8888/proxy?url=...&token=...` |

Every D365 HTTP path branches on `IS_EXT`. Always maintain both branches.

---

## Critical Coupling: DOM IDs

`popup.js` accesses these IDs directly via `getElementById`. **Never rename or remove them from `popup.html` without updating every reference in `popup.js`.**

`pickerA`, `pickerB`, `stA`, `stB`, `modeNote`, `btnVal`, `spinVal`, `btnLoad`, `spinLoad`, `btnDiag`, `diagPanel`, `progressPanel`, `progressLabel`, `progressText`, `progressPct`, `progressFill`, `pfName`, `pfUrl`, `pfEditId`, `btnAddProfile`, `pfList`, `profileSettings`, `modSel`, `company`, `companyCustom`, `entitySearch`, `btnModLoad`, `btnReport`, `diffOnly`, `moduleProgressPanel`, `moduleProgressLabel`, `moduleProgressText`, `moduleProgressPct`, `moduleProgressFill`, `modDetailPanel`, `modDetailTbody`, `modDetailSummary`, `modDetailFilter`, `entityDiffPanel`, `btnFullPage`, `toast`, `tokenModal`, `modalCloseBtn`, `modalGotItBtn`, `tokenCmd`

---

## Entity Fetching — Candidate Waterfall

`getCandidateEndpoints(origin)` returns 11 URLs tried in order:

1. `/Metadata/DataEntities` (3 casing variants) — returns plain array, richest metadata (`AppModule`, `Tags`, `EntityCategory`)
2. `/data/DataEntities?$top=10000...` (5 query/casing variants) — returns `{value:[...]}`, partial metadata
3. `/data`, `/data/` — OData service document; returns only `{kind:'EntitySet', name, url}` — no module/category

`fetchEntities` accepts the **first candidate that returns HTTP 200 AND normalises to ≥1 usable entity**. An endpoint that returns 200 with zero usable entities is skipped.

The order is intentional — **do not reorder** without testing on both sandbox and production D365 environments.

---

## Entity Normalisation Rules

`normaliseEntities(raw)` handles both array and `{value:[]}` shapes. An entity is **dropped** unless:
- `e.category` is in `{master, reference, parameter, parameters}` **OR** `e.serviceDoc === true`

Module/group assignment is runtime-driven:
- exact module fields (`AppModule`, `ApplicationModule`, `Module`, `ModuleName`) are authoritative
- `Tags` can provide a metadata-backed module/group when module fields are absent
- bare service-document rows (`/data`, `/data/`) must not infer business modules from entity-name prefixes; they are grouped as `Raw Entity Sets`
- non-service rows with no module metadata are grouped as `Unclassified`

This intentionally excludes transactional data. The service-root fallback (`/data`) bypasses category filtering via `serviceDoc`, so all public entity sets pass through — some transactional entities may appear when only that fallback works.

---

## Storage: Dual-Write Pattern

Profiles and picker state are written to **both** `localStorage` and `chrome.storage.local`. On popup open, `syncProfilesFromStorage()` pulls from `chrome.storage.local` first (popup `localStorage` is cleared when the popup closes).

```
putProfiles(list)  → localStorage['d365_profiles'] + chrome.storage.local['d365_profiles']
persist()          → localStorage['d365_pick']     + chrome.storage.local['d365_pick']
restore()          → reads localStorage (already synced by syncProfilesFromStorage)
```

In proxy mode `_store()` returns `null`; only `localStorage` is used.

---

## In-Memory State

```js
STATE = {
  allRows: [],               // merged entity rows: {name, module, inA, inB, status}
  entitiesA/B: [],           // normalised entity objects
  entityMapA/B: {},          // {entityName → entity} indexes
  lblA, lblB: '',            // display names
  activeModule: '',          // current module filter value
  moduleDetailRows: [],      // compareEntityRecords results
  visibleModuleDetailRows: []// after applyModuleDetailFilters()
}
```

`STATE` is **not persisted** — reloading the popup or reloading entity data clears it.

---

## Module Compare Flow

1. `loadModuleEntities()` — iterates `STATE.allRows` filtered by `modSel`/`entitySearch`
2. For each entity: `fetchCollectionRows(url, slot, entityMeta)` → `/data/<collection>?$top=50&$count=true&cross-company=true` (+ optional `$filter=dataAreaId eq '...'`)
3. `findAllDifferentRowPairs(rowsA, rowsB)` — auto-detects business key fields by suffix heuristic (`Id`, `Code`, `Num`, `Key`, `Name`, `No`, `Ref`) then builds keyed maps and returns all differing/missing pairs
4. Results populate `modDetailTbody`; clicking a non-`No OData` row calls `showEntityDiff(row)` for field-level diff
5. Rate-limit retries: `fetchWithRetry` retries HTTP 429 up to 3× with exponential back-off (800 ms base)
6. Concurrency: `mapLimit(items, 4, worker)` limits simultaneous D365 requests

---

## Key Constraints & Pitfalls

- **`content.js` must `return true`** from its message listener to keep the async `sendResponse` channel open. Never change this.
- **Do not add `$select` to candidate endpoints** — restrictive `$select` clauses cause 404 on some D365 environments.
- **`probeEndpoint` ≠ `askContentScript`** — `askContentScript` was an old duplicate and has been removed; use `probeEndpoint` everywhere.
- **Validate only uses `$top=1`** in the candidate list (for a quick ping); `fetchEntities` uses the full `$top=10000` list.
- **The `collection` field** on a normalised entity is the OData collection name used in `/data/<collection>` queries. If it's empty, the entity is classified `No OData` and record comparison is skipped.
- **Service-doc entities** get their `collection` from `e.url || e.name` from the service document entry. They stay in the dataset, but their grouping is `Raw Entity Sets` unless richer module metadata came from another endpoint.
- **`buildRows`** merges source and target entity lists; the `status` field is `Match` | `Only in Source` | `Only in Target` (not a count comparison — that's `compareEntityRecords`).
- **HTML report** is a self-contained Blob URL opened in a new tab — no external dependencies. The inline `<script>` at the end of `buildReportHtml` wires all tab/filter interactivity.

---

## No Build Step

Edit files directly. Reload the extension in Chrome (`chrome://extensions` → reload) to test changes. There is no compiler, bundler, or test suite.
