# Reconciliation: live `main` and the `bhb-platform` source (packet RHM, 2026-09-24)

This branch moves the Import Builder source from `bhb-platform` into this repo. The two sides had diverged:

- **Live `main` only:** `274628b` "Correct bilingual interface copy" (2026-09-09) and `9046969` "Increment 13 — repair 360 import workbook compatibility" (2026-09-18). Both were edited directly in the built files at the root. Neither existed in the `bhb-platform` source.
- **`bhb-platform` only:** `1028b7a` "Increment 360 collector mass relations" (2026-09-23). It added a single upload that auto-detects collector, relation template, legacy, allocation export and existing import, plus optional criteria, cohort warnings, the traceability-report rename and project-name autofill. It was never deployed.

**Method.** The published bundle `7599b4d` is byte-identical to `bhb-platform` `afadad4:40-standalone/360-import-builder/deploy/` (all 11 files). That makes `afadad4` the common base, and the two live commits are clean deltas on the built files. Each hunk was mapped back to `src/` and applied to the current source (`1028b7a`). The diffs were then re-run:

- The live `9046969` `core.js`, with `export` stripped from ours, differs from the new `src/core.js` **only** in mass-relations lines.
- The live `index.html` differs from the new `deploy/index.html` **only** in the mass-relations upload card, the autofill note and the traceability-report copy.
- The i18n dictionaries of the live `app.js` and the new `src/app.js` differ **only** in keys that mass-relations added or renamed: `files`, `filesHelp`, `audit`, `auditPrivacy`, `massTemplate*`, `respondentGuidance`, `cohortWarning`, `projectAutoFilled` and `criteriaConflict`.

Rule applied (RHM DoD 6): nothing is lost from either side. Where the two sides touch the same structure, the mass-relations structure wins and the live logic and copy are carried into it.

## Hunk dispositions

| # | Live hunk (commit, built file) | Lands in | Disposition | Proving check |
|---|---|---|---|---|
| C1 | `274628b` `app.js`: RO `privacyLine`, `reviewLead`, `name`, `email`, `conflict`, `recommended`, `recoveryText`, `sourceRejectedIssue` | `src/app.js` `text.ro` | Carried. Values edited in place in the dictionary instead of through the live `Object.assign` override, with the same resulting strings | `tests/increment13.test.mjs` "copy fix survives in the built deploy bundle"; the dictionary diff against live shows no difference on these keys |
| C2 | `274628b` `app.js`: EN `privacyLine`, `name`, `email`, `conflict`, `recommended`, `sourceRejectedIssue` (and `recoveryText`, which was already equal) | `src/app.js` `text.en` | Carried, as C1 | as C1 |
| C3 | `274628b` `app.js`: RO/EN `allocationHelp` ("live" → "activ"/"active") | `src/app.js` | Carried into the dictionary. **Not rendered**: mass-relations removed the separate allocation-file card, so no element uses this key. Winning behavior: single upload (mass-relations). Preserved: the corrected copy | dictionary diff |
| C4 | `274628b` `index.html`: `privacyLine` fallback text | `src/index.html` | Carried | increment13 test (html phrases) |
| C5 | `274628b` `index.html`: `reviewLead` fallback ("pentru fiecare participant și sursă") | `src/index.html` | Carried. The source still had the even older "izolate pe participant." | increment13 test |
| C6 | `274628b` `index.html`: `recoveryText` fallback | `src/index.html` | Carried | increment13 test |
| C7 | `274628b` `index.html`: `allocationHelp` fallback on the `#allocation-file` card | — | **Structural conflict, resolved.** Mass-relations replaced both inputs with one `#source-files` picker, so the card no longer exists. Winning: the mass-relations single upload. Preserved: the copy, in the dictionary (C3) | browser test "one project-workbook picker auto-fills the name from an existing import" |
| I1 | `9046969` `core.js`: new `roleKey()` (fold, then drop spaces, `_` and `-`) and `isSelfRole()` | `src/core.js` after `fold()` | Carried | increment13 "role aliases match ignoring spaces, hyphens and underscores" |
| I2 | `9046969` `core.js` `mapRole`: key by `roleKey`; `functionalmanager`; the `partener extern` alias is subsumed | `src/core.js` `mapRole` | Carried. Also applies to the mass-relations parser, which calls `mapRole` | increment13 role test and "mass-relations rows also benefit from tolerant role matching" |
| I3 | `9046969` `core.js` `parseCollectorRows`: `isSelf` uses `isSelfRole(sourceRole)` instead of `=== "Autoevaluare"` | `src/core.js` (the mass-relations criteria-aware version of the function) | Carried | increment13 "collector rows recognise any Autoevaluare spelling" |
| I4 | `9046969` `core.js` `parseLegacyRows`: evaluator name is `Nume Prenume` (last name first), including the normalization record | `src/core.js` | Carried | increment13 "secondary Rol/Nume/Prenume/Email template…" |
| I5 | `9046969` `core.js` `parseLegacyRows`: `isSelf` uses `isSelfRole` | `src/core.js` | Carried | same |
| I6 | `9046969` `core.js` `createProductionWorkbook`: N formula `G&" - "&B&" - "&D` instead of `CONCAT(...)` (avoids Excel's implicit-intersection `@` rewrite) | `src/core.js` (the criteria-aware version) | Carried. The existing contract test was updated from `CONCAT` to the `&` form, and `docs/production-contract.md` was updated to match | `core.test.mjs` "production workbook contract…"; increment13 "NumeInregistrare is an & formula…" (cell formula and raw XML: no `CONCAT`, `_xlfn` or `@`) |
| I7 | `9046969` `core.js`: `export` keywords stripped from the served `core.js` | — | **Not carried; no behavioral effect.** `index.html` loads only `app.js` (the classic bundle, which already has no exports). `core.js` is copied verbatim from `src/` by the unchanged `tools/build.mjs`, and the tests import the ES-module source. Carrying this would need a build change. The served `core.js` therefore goes back to having `export` lines, as in `7599b4d` | `package.test.mjs` classic-bundle tests; browser test "initializes when deploy index is opened directly from disk" |
| I8 | `9046969` `app.js`: one blank line added | — | Dropped (whitespace) | — |
| I9 | `9046969` `app.js`: bundled copies of I1–I6 | `deploy/app.js` | Generated by `npm run build` from `src/core.js` | increment13 "survive in the built deploy bundle" (`roleKey` and the `&` formula present in `deploy/app.js`, no `CONCAT(`) |

## Mass-relations behaviour the deploy brings live (from `bhb-platform`, not from this reconciliation)

These are visible changes for Vlad's click-through:

- One "Încarcă workbook-urile proiectului" picker replaces the separate respondent-files and allocation-export inputs, with format auto-detection. Collector, mass-relations template, legacy, allocation export and existing A:N import are all accepted.
- `Criteriu1`–`Criteriu5` are carried into the A:N output. Conflicting criteria for one participant block.
- More than 15 respondents is now a **warning**, where live blocks it, with a recommendation above 10.
- "Dovada de audit" / "audit receipt" is renamed to "raport de trasabilitate" / "traceability report".
- The project name is autofilled from an existing A:N import with one unique `NumeCampanie`.

## Open observations and Vlad's rulings

1. **Mass template self detection — kept per production contract, confirmed by Vlad (2026-09-24).** `parseMassRows` marks a row as self only when the evaluator email equals the participant email, as before. Increment 13's `isSelfRole` is not applied there, because the contract says the mass template omits the self row and the importer synthesises one. A mass-template row labelled `Autoevaluare` with a different email is therefore mapped to a `Manager` relation, not to self. This matches `docs/production-contract.md`: self is the `Manager` row whose evaluator and participant emails match, and the mass template omits self. No change.
2. **Legacy name order.** Increment 13 makes legacy evaluator names `Nume Prenume`. Any consultant who loaded legacy files on live before 2026-09-18 saw `Prenume Nume`. That is live behavior today, not a change made here.

3. **Internal docs are private (Vlad, 2026-09-24).** `docs/platform-bridge-proposal.md` is removed here and the collector's `docs/review-note.md` is removed in its own PR. Both are kept in `bhb-platform` `10-project/standalone-history/`. The two README links now read "internal note, kept in the private workspace".
4. **Generic project placeholder (Vlad, 2026-09-24).** The former placeholder was an internal project name. It becomes `ex. Proiect 360 2026` (RO) and `e.g. 360 Project 2026` (EN). The placeholder had no EN form before, so it is now an i18n key (`projectPlaceholder`, applied in `setLang` through `data-i18n-placeholder`). Check: `increment13.test.mjs` "project placeholder is generic in RO and EN".

## Served root, before and after (SHA-256)

| File | Live `main` `9046969` | This branch |
|---|---|---|
| `app.js` | `535e3fad361b87d49482d906a22f97059fd638e05809f26eefbeecb8a9af0614` | `b29eb7ebdeb3f6c147dbae0cda4471de6482d5b18072b00220031e79b88930f6` |
| `core.js` | `53cc2dc17049dbb27912aa2400b9a7f9e3dc937968f751f1f4dcbd0e58782c86` | `9b5970b9e327b13f2a5378e31122676807cbbd9e510423c8e0ca2231357f870b` |
| `index.html` | `9a4c950aaefcd205e652fb7d700f75a2281d4324d8e9f6ebdc5e80cbec7038ff` | `6ecf4e13f9615d8f4043736255b35e8c311fa0dc302bc83c73d5c3e952fdaeca` |
| `styles.css` | `5c61bdeafe3af47b9bb70a95786dc9d3f6e7b5d981b1daeaf8d6842b2ae25ee4` | `0976413e82b93f09625e0bf59c43472f70f8e8e2c1ffd1edb91b35a827c19e54` |
| `assets/` (7 files: 3 fonts, 2 images, SheetJS + licence) | unchanged | unchanged |

## Privacy review of everything this branch makes public

| Path | Content | Personal / client data | Disposition |
|---|---|---|---|
| `src/` | Canonical source of the served app (ES-module form), same assets | None. The former placeholder (an internal project name, live since `7599b4d`) is replaced; see ruling 4 | Publish |
| `deploy/` | Build output, byte-identical to the served root | As above | Publish |
| `tests/*.test.mjs`, `tests/browser.browser.mjs`, `tests/increment13.test.mjs` | Synthetic workbooks | Only `D210 …` names and `@example.invalid` addresses | Publish |
| `tests/fixtures/d210-project.json` | Marked "D210 SYNTHETIC — not real client data" | Synthetic only | Publish |
| `tools/` | Build, browser-test runner, root check | None | Publish |
| `docs/production-contract.md` | The A:N contract of the old production importer | None. Column names are already visible in the public app | Publish |
| `docs/platform-bridge-proposal.md` | Internal product-direction note | None personal | **Not published** (Vlad, 2026-09-24): removed from this branch. The copy is kept in `bhb-platform` `10-project/standalone-history/` (`074f3d8`). The README now calls it an internal note |
| `README.md` | Usage, formats, asset authorization, verification | Names Vlad as the asset authorizer and "BHB/Trend consultants"; no client data. The cross-folder link to the collector review note now points to the collector repo on GitHub | Publish |
| `THIRD_PARTY_NOTICES.md` | SheetJS and Poppins notices | None | Publish |
| `.github/` | Issue and PR templates (from the closed CI PR #2) and the `verify` workflow | None | Publish |
| `RECONCILIATION.md` | This file | None | Publish |

Pages note: the repo root keeps serving the app. The added Markdown files render as extra Jekyll pages (for example `/README.html`), and none contains Liquid tokens. `.github/` is not served.
