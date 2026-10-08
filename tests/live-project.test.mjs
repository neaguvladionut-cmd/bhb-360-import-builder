import assert from "node:assert/strict";
import test from "node:test";
import { analyzeAllocation, analyzeProject } from "../src/core.js";

const state = [{ identifier: "71", participantName: "D210 Ada", participantEmail: "ada@example.invalid", name: "D210 Existing Manager", email: "manager@example.invalid", sourceRole: "Manager", campaign: "D210 Live", language: "RO", criteria: ["Leadership", "Integrity", "", "", ""], source: "d210-live.xlsx", rowNumber: 2, origin: "existing-production" }];
function incoming(rows) { return { id: "d210-resend", source: "d210-resend.xlsx", participantName: "D210 Ada", participantEmail: " ADA@example.invalid ", rows }; }
function row(name, email, role, isSelf = false, language = "RO", criteria = ["Leadership", "Integrity", "", "", ""]) { return { evaluatorName: name, evaluatorEmail: email, sourceRole: role, language, criteria, isSelf, source: "d210-resend.xlsx", rowNumber: 2 }; }

test("live project normalizes a resent pair, skips its self row, and blocks a new Manager", () => {
  const analysis = analyzeProject({ participants: [incoming([row("D210 Ada", "ada@example.invalid", "Autoevaluare", true), row("D210 Existing Manager", " MANAGER@example.invalid", "Manager"), row("D210 New Manager", "new.manager@example.invalid", "Manager")])], stateRecords: state, allocation: analyzeAllocation(state), projectName: "D210 Live" });
  assert.equal(analysis.outputRows.length, 0);
  assert.equal(analysis.skippedRows.length, 2);
  assert.equal(analysis.skippedRows[1].reason, "already-in-project");
  assert(analysis.blockers.some((item) => item.code === "new-manager-existing-participant"));
});

test("live project exports a different participant's respondent and blocks campaign and criteria changes", () => {
  const newPerson = { id: "d210-new", source: "d210-new.xlsx", participantName: "D210 Bogdan", participantEmail: "bogdan@example.invalid", rows: [row("D210 Bogdan", "bogdan@example.invalid", "Autoevaluare", true), row("D210 Existing Manager", "manager@example.invalid", "Manager")] };
  const exported = analyzeProject({ participants: [newPerson], stateRecords: state, allocation: analyzeAllocation(state), projectName: "D210 Live" });
  assert.equal(exported.ready, true);
  assert.equal(exported.outputRows.length, 2);
  assert.equal(exported.identifiers["manager@example.invalid"], "71");
  const changed = analyzeProject({ participants: [incoming([row("D210 Existing Manager", "manager@example.invalid", "Manager", false, "RO", ["Different", "Integrity", "", "", ""])])], stateRecords: state, allocation: analyzeAllocation(state), projectName: "Other campaign" });
  assert(changed.blockers.some((item) => item.code === "campaign-mismatch"));
  assert(changed.blockers.some((item) => item.code === "criteria-conflict"));
});
