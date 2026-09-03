# Velixo FACTORIALROW — Excel Add-in

An Excel add-in built with [Office.js](https://learn.microsoft.com/office/dev/add-ins/) that adds a
custom function to the **`TESTVELIXO`** namespace:

```
=TESTVELIXO.FACTORIALROW(N)
```

It spills the factorials `[0!, 1!, 2!, ..., N!]` into a **dynamic array (spill range)** — a real range
of `N + 1` cells, not a comma-separated string in one cell.

## Highlights

| Requirement | How it's met |
| --- | --- |
| Office.js + shared runtime + TypeScript | XML manifest declares a single `lifetime="long"` runtime; all source is TypeScript. |
| `TESTVELIXO.FACTORIALROW(N)` spilling `[0!..N!]` | `src/functions/functions.ts`; JSDoc `@returns {any[][]}` → `dimensionality: "matrix"`. |
| Task pane, row/column toggle | React + Fluent UI v9 (`src/taskpane/App.tsx`); toggling triggers a full recalculation so results re-orient live. |
| Setting persists across reloads | `OfficeRuntime.storage` (`src/shared/orientation.ts`). |
| Each `N!` computed only once | Persistent, ever-growing cache inside the web worker (`src/workers/factorial.worker.ts`). |
| Precision up to `N = 500` | All math in `BigInt`; values beyond `Number.MAX_SAFE_INTEGER` are returned as exact strings. |
| Computation offloaded to a web worker | `functions.ts` delegates to a long-lived worker via `worker-client.ts`. |

## Architecture

The shared runtime loads both the task-pane bundle and the custom-functions bundle into the **same
`window`**. Because separate webpack chunks don't share module-level state, orientation lives on a
`window.__velixo` global that both sides read/write; `OfficeRuntime.storage` persists it across
reloads.

```
Task pane (React/Fluent)                 Custom function (functions.ts)
  radio → persistOrientation()             reads getOrientation()  ─┐
    │  writes window.__velixo               validates 0..500         │
    │  writes OfficeRuntime.storage         │                        │
    └─ Excel full recalculation ────────────┘                        │
                                             postMessage(n) ──► Web worker
                                                                 persistent bigint[] cache
                                             ◄── string[] ──────  factorialSeries()
                                             toCellValue(): number if safe, else string
                                             toSpillRange(): [row] or [[col]]
```

Pure, host-independent logic lives in `src/shared/factorial.ts` and is reused by both the worker and
the unit tests (DRY) — no Excel host is needed to test the math or the spill shaping.

```
src/
  shared/      factorial.ts, orientation.ts, worker-protocol.ts   (pure, testable)
  functions/   functions.ts (custom function), worker-client.ts   (worker plumbing)
  workers/     factorial.worker.ts                                (off-thread compute + cache)
  taskpane/    index.tsx, App.tsx, App.css, taskpane.html         (React + Fluent UI)
test/          factorial.test.ts, orientation.test.ts,
               functions.test.ts                                  (Jest, no Excel host)
```

## Prerequisites

- Node.js 18+ and npm
- Microsoft Excel on the web (free) or Excel desktop

## Getting started

```bash
npm install
npm test          # run the unit tests
npm run build     # production build to dist/ (generates dist/functions.json)
npm run validate  # validate manifest.xml
npm run dev-server # serve the add-in over https://localhost:3000
```

The first `dev-server` run installs a local development HTTPS certificate
(`office-addin-dev-certs`); accept the prompt.

### Sideloading in Excel on the web

1. Run `npm run dev-server`.
2. Open a workbook at <https://office.com> → **Excel**.
3. **Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in**, and choose `manifest.xml`.
4. In any cell: `=TESTVELIXO.FACTORIALROW(5)` → spills `1  1  2  6  24  120`.
5. Open the task pane (**Home → Velixo → FACTORIALROW Settings**) and switch **Row / Column**; the
   result re-orients and the workbook recalculates. Reload the page — the choice is remembered.

On Excel desktop you can instead run `npm start`, which sideloads automatically.

### Try it

- `=TESTVELIXO.FACTORIALROW(25)` — later values spill as exact strings (BigInt, no precision loss).
- `=TESTVELIXO.FACTORIALROW(500)` — 501 lossless values.
- After computing `FACTORIALROW(500)`, `FACTORIALROW(200)` returns instantly (served from the cache).

## Notes & trade-offs

- Values `<= 18!` are numbers (Excel treats them numerically); `>= 19!` are exact strings, since
  `19!` already exceeds `Number.MAX_SAFE_INTEGER`. This is the intended lossless behaviour.
- `N` is capped at 500 per the brief; out-of-range input returns a `#VALUE!` custom-function error.
- The icons in `assets/` are simple generated placeholders.
- For deployment, replace the `https://localhost:3000/` URLs in `manifest.xml` and `webpack.config.js`.

## AI / LLM assistance disclosure

This project was built with the assistance of an AI coding agent (Anthropic's Claude, via Claude
Code). The collaboration worked as follows:

- **Research.** I used the agent to gather the current Microsoft documentation on the shared-runtime
  manifest configuration, how webpack bundles share state under the shared runtime, the web-worker
  pattern for Office Add-ins, the `functions.json` matrix metadata format, and `OfficeRuntime.storage`
  semantics. These findings shaped the architecture (notably the `window.__velixo` state-sharing
  approach, which is necessary because separate webpack chunks don't share module scope).
- **Implementation.** The agent scaffolded the project and wrote the source, tests, manifest, and
  build configuration to an architecture I reviewed and approved up front.
- **Verification.** All code was compiled, linted, and unit-tested locally; the add-in was
  sideloaded and exercised in Excel to confirm spilling, orientation switching, persistence, large-N
  precision, and caching behaviour.

Deficiencies I had to correct during the process are recorded in the git history. Every file was
reviewed for correctness and clarity before submission; I take full responsibility for the result.
