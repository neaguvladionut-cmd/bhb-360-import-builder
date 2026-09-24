# Proposal: respondent enrollment in the future platform

**Status: proposal only.** This document records learning from the standalone bridge for later platform design. It does not amend the project decisions, specifications, `TASKS.md`, `NOW.md` or the current build structure, and it is not authorization to build platform functionality on this branch.

## Problem learned from the current process

The old workflow asks participants to fill spreadsheets, consultants to normalize and merge them, and a second spreadsheet importer to create project allocations. That makes formatting, role translation, identity reuse and the unusual self-row ordering the consultant's responsibility. A small error can surface only after notifications fail.

The standalone collector and importer reduce that risk for the current production app, but the downloaded-file and email handoff still exists.

## Recommended target

Respondent collection should be a native step inside each future 360 project:

1. The consultant creates/opens the project and invites its participants.
2. Each authenticated participant enters respondents in a guided form tied to that project and their own participant record.
3. The platform validates relationships, email identity, respondent limits and questionnaire language at entry time.
4. Submission creates or reuses the respondent identity and project allocation directly—there is no intermediate workbook and no production-role translation exposed to the participant.
5. The communication engine sends invitations, records delivery state and supports reminders and corrections from the same project.
6. The consultant sees project-wide completion, conflicts and language distribution before launching or extending the project.

In that target, self-evaluation is a first-class allocation type. `Stakeholder` is the participant-facing relationship label, and any legacy `PartenerExtern` mapping lives only in a compatibility adapter.

## Importer as fallback only

Retain an importer inside the future project's respondent area for migration, bulk correction and exceptional consultant-led onboarding. It should use the platform's domain model and APIs, not recreate the old A:N workbook internally. The old-production export should remain isolated as a legacy adapter and be removable when the old application is retired.

Fallback import should preserve the safeguards learned here:

- normalized email as an identity candidate, with explicit conflict handling rather than silent merges;
- project-scoped identifier/allocation reconciliation;
- visible name conflicts and source provenance;
- per-respondent questionnaire language;
- participant-level error isolation and an auditable review before committing changes;
- privacy, retention and deletion rules appropriate to personal contact data.

## Decisions still required in the platform increment

Before this becomes a platform build pack, Vlad/Architect must decide and record:

- who may add, edit, submit, reopen and lock respondent nominations;
- whether respondents require accounts before invitation acceptance;
- consent, privacy notice, retention and deletion behavior for nominated-but-never-invited people;
- duplicate-person resolution across organizations and projects;
- launch gates, late additions, withdrawal and replacement behavior;
- notification timing, bounce handling, reminders and consultant overrides;
- whether participant submissions are visible verbatim to consultants and when;
- audit-log requirements and the fallback import's rollback behavior.

The existing project decisions around 360 invitation chains and the communication engine remain the starting point. This proposal should be reconciled against the then-current decisions log and build pack in a future Architect/Vlad decision gate.
