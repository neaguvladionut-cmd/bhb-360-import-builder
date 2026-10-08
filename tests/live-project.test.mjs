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

test("near neighbours: same respondent for another participant exports, email case and spaces remain one pair", () => {
  const other = { id: "d210-other", source: "d210-other.xlsx", participantName: "D210 Bogdan", participantEmail: "bogdan@example.invalid", rows: [row("D210 Bogdan", "bogdan@example.invalid", "Autoevaluare", true), row("D210 Existing Manager", " MANAGER@example.invalid ", "Manager")] };
  const exported = analyzeProject({ participants: [other], stateRecords: state, allocation: analyzeAllocation(state), projectName: "D210 Live" });
  assert.equal(exported.ready, true);
  assert.equal(exported.outputRows.filter((item) => item.evaluatorEmail === "manager@example.invalid").length, 1);
  const resent = analyzeProject({ participants: [incoming([row("D210 Ada", " ADA@example.invalid ", "Autoevaluare", true), row("D210 Existing Manager", " MANAGER@example.invalid ", "Manager")])], stateRecords: state, allocation: analyzeAllocation(state), projectName: "D210 Live" });
  assert.equal(resent.outputRows.length, 0);
  assert.deepEqual(resent.skippedRows.map((item) => item.reason), ["already-in-project", "already-in-project"]);
  assert.equal(resent.nothingNew, true);
});

test("near neighbours: state-only participant and state-only respondent keep their distinct meanings", () => {
  const stateOnly = analyzeProject({ participants: [], stateRecords: state, allocation: analyzeAllocation(state), projectName: "D210 Live" });
  assert.equal(stateOnly.nothingNew, true);
  assert.equal(stateOnly.ready, false);
  const arrivingParticipant = { id: "d210-manager-now-participant", source: "d210-manager-now-participant.xlsx", participantName: "D210 Existing Manager", participantEmail: "manager@example.invalid", rows: [row("D210 Existing Manager", "manager@example.invalid", "Autoevaluare", true), row("D210 New Manager", "new.manager@example.invalid", "Manager")] };
  const newParticipant = analyzeProject({ participants: [arrivingParticipant], stateRecords: state, allocation: analyzeAllocation(state), projectName: "D210 Live" });
  assert.equal(newParticipant.ready, true);
  assert.deepEqual(newParticipant.outputRows.map((item) => item.isSelf), [true, false]);
  assert(newParticipant.outputRows.every((item) => item.participantEmail === "manager@example.invalid"));
});

test("near neighbours: duplicate state pair across files stays one known pair and a removal-looking resend offers nothing", () => {
  const duplicateState = state.map((item) => ({ ...item, source: "d210-live-copy.xlsx" }));
  const removalLooking = analyzeProject({ participants: [incoming([row("D210 Ada", "ada@example.invalid", "Autoevaluare", true)])], stateRecords: [...state, ...duplicateState], allocation: analyzeAllocation([...state, ...duplicateState]), projectName: "D210 Live" });
  assert.equal(removalLooking.ready, false);
  assert.equal(removalLooking.nothingNew, true);
  assert.equal(removalLooking.skippedRows.length, 1);
  assert.equal(removalLooking.skippedRows[0].reason, "already-in-project");
  assert.equal(removalLooking.blockers.filter((item) => item.code === "duplicate-allocation").length, 0);
});

test("near neighbours: state criteria inherit for a new Peer and a criteria or campaign change blocks", () => {
  const peer = incoming([row("D210 Ada", "ada@example.invalid", "Autoevaluare", true), row("D210 Existing Manager", "manager@example.invalid", "Manager"), row("D210 New Peer", "new.peer@example.invalid", "Peer")]);
  const inherited = analyzeProject({ participants: [peer], stateRecords: state, allocation: analyzeAllocation(state), projectName: "D210 Live" });
  assert.equal(inherited.ready, true);
  assert.deepEqual(inherited.outputRows[0].criteria, state[0].criteria);
  const conflict = analyzeProject({ participants: [incoming([row("D210 Ada", "ada@example.invalid", "Autoevaluare", true), row("D210 Changed", "changed@example.invalid", "Peer", false, "RO", ["Different", "Integrity", "", "", ""])])], stateRecords: state, allocation: analyzeAllocation(state), projectName: "Other campaign" });
  assert(conflict.blockers.some((item) => item.code === "criteria-conflict"));
  assert(conflict.blockers.some((item) => item.code === "campaign-mismatch"));
});

test("near neighbours: a participant-name conflict blocks until the settled state name is chosen", () => {
  const mismatched = incoming([row("D210 Ada Revised", "ada@example.invalid", "Autoevaluare", true), row("D210 New Peer", "new.peer@example.invalid", "Peer")]);
  mismatched.participantName = "D210 Ada Revised";
  const blocked = analyzeProject({ participants: [mismatched], stateRecords: state, allocation: analyzeAllocation(state), projectName: "D210 Live" });
  assert(blocked.blockers.some((item) => item.code === "participant-name-conflict"));
  const settled = structuredClone(mismatched);
  settled.participantName = "D210 Ada";
  settled.rows.find((item) => item.isSelf).evaluatorName = "D210 Ada";
  const chosen = analyzeProject({ participants: [settled], stateRecords: state, allocation: analyzeAllocation(state), projectName: "D210 Live" });
  assert.equal(chosen.ready, true);
  assert.equal(chosen.outputRows[0].participantName, "D210 Ada");
});

test("near neighbours: allocation-only evidence needs a campaign, new participants stay unchanged, and F2 blocks only the new Manager", () => {
  const fresh = { id: "d210-fresh", source: "d210-fresh.xlsx", participantName: "D210 Fresh", participantEmail: "fresh@example.invalid", rows: [row("D210 Fresh", "fresh@example.invalid", "Autoevaluare", true), row("D210 Fresh Manager", "fresh.manager@example.invalid", "Manager")] };
  const allocationOnly = analyzeAllocation([{ identifier: "99", name: "D210 Fresh Manager", email: "fresh.manager@example.invalid", participantEmail: "fresh@example.invalid", sourceRole: "Manager", source: "d210-accounts.xlsx", rowNumber: 2 }]);
  const missingCampaign = analyzeProject({ participants: [fresh], allocation: allocationOnly, projectName: "" });
  assert(missingCampaign.blockers.some((item) => item.code === "project-required"));
  const withCampaign = analyzeProject({ participants: [fresh], allocation: allocationOnly, projectName: "D210 Entered Campaign" });
  assert.equal(withCampaign.ready, true);
  assert.deepEqual(withCampaign.outputRows.map((item) => item.participantName), ["D210 Fresh", "D210 Fresh"]);
  const managerAndPeer = analyzeProject({ participants: [incoming([row("D210 Ada", "ada@example.invalid", "Autoevaluare", true), row("D210 New Manager", "new.manager@example.invalid", "Manager"), row("D210 New Peer", "new.peer@example.invalid", "Peer")])], stateRecords: state, allocation: analyzeAllocation(state), projectName: "D210 Live" });
  assert(managerAndPeer.blockers.some((item) => item.code === "new-manager-existing-participant"));
  assert.deepEqual(managerAndPeer.outputRows.map((item) => item.sourceRole), ["Peer"]);
});
