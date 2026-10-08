# 360-IB-ADD — Definition of done

Date: 2026-10-08

1. Keep production-import and Accounts state with row/file provenance; load, remove, and replace each source independently.
2. Require one normalized campaign across state and export; allocation-only evidence requires an entered campaign.
3. Treat an evaluated email in state as an existing participant, preserving state name and criteria.
4. Compare normalized participant/respondent pairs and surface app-only changes.
5. Keep missing resend respondents and explain that removals happen in the 360 app.
6. Preserve the present new-participant rules.
7. Export new rows only, reusing identifiers and allocating from the observed maximum.
8. Show new, skipped, and warning splits and separate cohort size from export rows.
9. Block a new non-self Manager for an existing participant behind one named, reversible rule.
10. Use plain Romanian and English guidance and document the live-project workflow.
11. Cover the agreed synthetic cases and their near neighbours.
12. Build and run the specified checks; leave independent rendered inspection to the non-builder.
