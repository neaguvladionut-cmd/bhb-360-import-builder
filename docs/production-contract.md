# Current production 360 import contract

This document is the implementation contract for the standalone internal import builder. It describes the old production 360 application as it must be served now; it is not the target architecture for the future platform.

## Inputs

The consultant supplies:

- a project/campaign name, unless it can be recovered from an existing production import;
- one or more standardized collector workbooks, completed mass-relations templates, supported legacy participant collection workbooks, allocation exports and/or existing production imports;
- optionally, the current project's questionnaire-allocation export or an existing A:N production import when adding participants to a live project.

The optional allocation export uses the observed `Accounts` sheet headers:

`Identifier`, `Assessor Name`, `Assessor Email`, `QuestionaireId`, `Assessed Email`, `Assessor Role`

The consultant explicitly supplies this export as the existing allocation state for the named project. Its identifiers are never mixed with another project's state.

An existing production import is recognized by the exact A:N headers. Its evaluator identifier, name and email columns are reused as allocation evidence. When all populated `NumeCampanie` cells contain one unique non-empty value, that value pre-fills the project field; multiple or missing campaign values leave the field for manual entry.

## Normalization and blocking rules

- Trim leading/trailing ordinary whitespace, non-breaking spaces and line breaks from names, emails, roles and language values.
- Compare evaluator emails case-insensitively after trimming. Do not guess or repair the address itself.
- The normalized evaluator email is the project-wide identity key. Reuse its identifier everywhere it appears, even when its relationship role differs between participants.
- Reject only a duplicate normalized evaluated-email + evaluator-email allocation; cross-participant evaluator reuse is valid.
- Block one email mapped to multiple identifiers and one identifier mapped to multiple emails.
- If one normalized email has different names, show all names and source files and require an explicit consultant choice. Preselect the allocation-export name when present.
- Unknown roles, missing participant identity, malformed emails and missing Manager rows block the affected participant. Ten non-self respondents in total is the general recommendation; more than 15 produces a warning but does not block export. Valid participants from other files remain reviewable.
- Participant criteria are optional and limited to five columns. When present, all rows for that participant must carry one consistent set; conflicting values in the same criterion column block that participant until the source is corrected.
- Keep user values beginning with `=`, `+`, `-` or `@` as literal text. Column N is the only formula-bearing column.

Known role conversions are deterministic:

| Source concept/value | Production `EvaluatorRol` |
|---|---|
| Autoevaluare / identified self | `Manager` |
| Manager / Functional Manager | `Manager` |
| Peer / Coleg | `Peer` |
| Subordonat | `Subordonat` |
| Stakeholder / Partener / PartenerExtern | `PartenerExtern` |

Labels are matched ignoring case, diacritics, spaces, hyphens and underscores, so `Auto-evaluare`, `Partener extern` or `Functional_Manager` map the same way (Increment 13). Any `Autoevaluare` spelling marks the self row in collector and legacy files.

The mass-relations template uses the same participant/respondent fields as the collector output, but omits the self-evaluation row: the importer synthesizes one self row for each participant. Each participant must have at least one distinct `Manager` relation. `Criteriu1`–`Criteriu5` are optional and are copied to every production row for that participant.

The self allocation is identified by matching normalized evaluated and evaluator email, not merely by the word `Manager`.

## Identifier allocation

- New project: sort normalized evaluator identities deterministically and allocate text identifiers `1`, `2`, `3` and onward.
- Existing project: reuse each valid identifier from the allocation export. Allocate new evaluator identities consecutively from the highest existing numeric identifier plus one, using a deterministic normalized-email order.
- The same normalized evaluator email always receives one identifier throughout the generated workbook.
- Existing identifiers are text in the output even when the source workbook stored them numerically.

## Exact output workbook

Produce one worksheet with these exact headers in this exact order:

| Column | Header | Value/type |
|---|---|---|
| A | `Marca` | evaluator identifier, text |
| B | `PersoanaEvaluata` | evaluated participant name, text |
| C | `PersoanaEvaluataEmail` | evaluated participant email, text, no hyperlink |
| D | `Evaluator` | evaluator name, text |
| E | `EvaluatorEmail` | evaluator email, text, no hyperlink |
| F | `EvaluatorRol` | exact production role, text |
| G | `NumeCampanie` | consultant-entered project name, text |
| H | `Limba` | uppercase two-letter questionnaire language, text |
| I | `Criteriu1` | optional participant criterion, text |
| J | `Criteriu2` | optional participant criterion, text |
| K | `Criteriu3` | optional participant criterion, text |
| L | `Criteriu4` | optional participant criterion, text |
| M | `Criteriu5` | optional participant criterion, text |
| N | `NumeInregistrare` | row-relative formula only |

Additional invariants:

- Every cell in A:M is authored as text; blanks in I:M remain text-compatible blanks.
- F is exactly one of `Manager`, `Peer`, `Subordonat`, `PartenerExtern`.
- H is uppercase and exactly two letters. `EN` is valid and creates a visible non-blocking consultant notice.
- Every data row `n` in N contains `=Gn&" - "&Bn&" - "&Dn`. It uses the `&` operator, not `CONCAT`, so Excel does not rewrite the formula with an implicit-intersection `@` (Increment 13, live 2026-09-18). The header is the only non-formula cell in N.
- Email cells have no hyperlink relationship and no hyperlink styling.
- No sample/example row remains.
- Rows are grouped deterministically by evaluated participant. Within each group, self is first and therefore the first `Manager` row; at least one distinct actual Manager follows. Remaining rows use a documented deterministic role/email order.
- `NumeInregistrare` is produced only by the formula, never separately authored as text.

## Consultant review surface

Before download, show:

- totals by participant, production role and questionnaire language;
- the self and actual Manager status for every participant;
- every `EN` allocation as a notice, not an error;
- blocking errors with participant/person, source file, reason and corrective action;
- every same-email/different-name conflict and the selected canonical name;
- identifier reuse/new-allocation counts and the next allocated identifier.

No production workbook is downloadable while a blocking error or unresolved canonical-name choice remains.

## In-memory source ledger

Every attempted participant or allocation file remains in the current browser-memory session until explicitly removed or replaced. The ledger records:

- original filename, byte size and full SHA-256 fingerprint of the original bytes;
- current adapter/state: transitional `pending`, then `collector`, `legacy`, `allocation`, `existing-production` or settled `rejected`;
- parsed, exportable, blocked and intentionally ignored row counts;
- field-normalization count plus warning/error counts.

An attempted source becomes blocking before its bytes are read or fingerprinted. While local checking is in progress, the UI and audit receipt classify it as `pending`, show `unavailable (pending)` for its not-yet-completed fingerprint, and do not claim a read failure or prescribe rejection recovery. An unreadable, unhashable or unsupported attempt becomes `rejected` only after failure settles and continues to block production export; its filename and byte size remain visible, while its fingerprint is explicitly `unavailable` if SHA-256 cannot be completed. The audit receipt records the failed stage and recovery evidence honestly. The source card's explicit **Replace** control replaces only that exact attempt; the unified project-workbook picker always adds a source, so unrelated files with identical filenames are preserved. Loading an allocation-bearing file replaces the previous allocation attempt and clears its old identifier state before file reading begins. Removing a source does not discard unrelated valid files. Closing or resetting clears the complete in-memory ledger; already downloaded files remain on the consultant's device.

For every changed value, normalization provenance retains source filename, source row, field, original literal value and normalized literal value. This includes whitespace/case/language cleanup and deterministic role translation.

## Separate audit receipt

After at least one source attempt, the consultant may explicitly download `audit-360-<project>.xlsx` using the Descarcă raportul de trasabilitate button, whether the current run is ready or blocked. It is never automatic and never added to, linked from or bundled with the production import workbook.

The receipt has exactly five sheets:

- `Run` — local generation time, project, readiness/totals, privacy warning and validation boundary;
- `Sources` — complete source ledger and dispositions;
- `Issues` — blockers/warnings, person/source/row evidence and recovery action;
- `Identifiers` — evaluator email, reused/new identifier, canonical name and allocation evidence;
- `Normalizations` — field-level original and normalized literal values.

Every audit cell is authored as literal text, with no formulas or hyperlinks, including values beginning with `=`, `+`, `-` or `@`. The receipt is client-confidential and local; the consultant controls its retention/deletion. It is diagnostic evidence, not a production-app input.

## Verification boundary

Automated verification must inspect workbook semantics and raw XML: exact headers/order, text cell types, formulas on every data row, absence of hyperlinks, escaped literal input and deterministic identities/rows. Reopen the file in Excel or LibreOffice as a separate fidelity check.

These checks establish **structural validation**. Only a controlled trial in the old production application can establish **production compatibility**; it must verify import acceptance, self-sheet assignment, notification creation and identifier reuse.
