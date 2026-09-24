// Increment 13 (live 9046969, 2026-09-18): workbook-compatibility fixes that must survive the rehome.
// 1. NumeInregistrare uses the & operator, not CONCAT, so Excel does not rewrite it with implicit-intersection @.
// 2. Role labels are matched ignoring spaces, hyphens and underscores (roleKey), so the secondary
//    Rol/Nume/Prenume/Email collection template and its spelling variants are accepted.
// 3. Self rows are recognised from any Autoevaluare spelling (isSelfRole), in collector and legacy files.
// 4. Legacy evaluator names are "Nume Prenume" (last name first), as the Romanian template lists them.
// All data is synthetic and D210-marked.
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import { MASS_TEMPLATE_HEADERS, analyzeAllocation, analyzeProject, createProductionWorkbook, mapRole, parseParticipantWorkbook } from "../src/core.js";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);
const sandbox = { exports: {}, module: { exports: {} }, require, Buffer, process, setTimeout, clearTimeout };
vm.runInNewContext(await readFile(resolve(root, "src/assets/vendor/xlsx.full.min.js"), "utf8"), sandbox, { filename: "xlsx.full.min.js" });
const XLSX = sandbox.exports;

function workbookBytes(rows, sheetName = "Sheet1") {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), sheetName);
  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

test("Inc13: role aliases match ignoring spaces, hyphens and underscores", () => {
  for (const value of ["Autoevaluare", "Auto-evaluare", "auto evaluare", "AUTO_EVALUARE"]) assert.equal(mapRole(value), "Manager", value);
  for (const value of ["Functional Manager", "Functional-Manager", "functional_manager", "FunctionalManager"]) assert.equal(mapRole(value), "Manager", value);
  for (const value of ["Partener Extern", "Partener-Extern", "partener_extern", "PartenerExtern", "Partener", "Stakeholder"]) assert.equal(mapRole(value), "PartenerExtern", value);
  assert.equal(mapRole("Coleg"), "Peer");
  assert.equal(mapRole("Sub ordonat"), "Subordonat");
  assert.equal(mapRole("invented"), null);
});

test("Inc13: secondary Rol/Nume/Prenume/Email template parses with last-name-first evaluator names and spelled-apart self row", () => {
  const rows = [
    ["Rol", "Nume", "Prenume", "Email"],
    ["Auto-evaluare", "Exemplu", "D210 Ada", "ada@example.invalid"],
    ["Functional-Manager", "Manager", "D210 Mihai", "manager@example.invalid"],
    ["Partener extern", "Partner", "D210 Erin", "partner@example.invalid"],
  ];
  const parsed = parseParticipantWorkbook(XLSX, workbookBytes(rows), "d210-secondary.xlsx");
  assert.equal(parsed.type, "legacy");
  const [participant] = parsed.participants;
  const [self, manager, partner] = participant.rows;
  assert.equal(self.isSelf, true);
  assert.equal(self.evaluatorName, "Exemplu D210 Ada");
  assert.equal(participant.participantName, "Exemplu D210 Ada");
  assert.equal(manager.isSelf, false);
  assert.equal(manager.role, "Manager");
  assert.equal(manager.evaluatorName, "Manager D210 Mihai");
  assert.equal(partner.role, "PartenerExtern");
  const analysis = analyzeProject({ participants: parsed.participants, allocation: analyzeAllocation([]), projectName: "D210 Secondary" });
  assert.equal(analysis.ready, true, JSON.stringify(analysis.blockers));
  assert.equal(analysis.outputRows[0].isSelf, true);
  assert.equal(analysis.outputRows[0].evaluatorName, "Exemplu D210 Ada");
});

test("Inc13: collector rows recognise any Autoevaluare spelling as the self row", () => {
  const rows = [
    ["ParticipantName", "ParticipantEmail", "EvaluatorName", "EvaluatorEmail", "Relationship", "Language", "IsSelf"],
    ["D210 Ada", "ada@example.invalid", "D210 Ada", "ada.self@example.invalid", "Auto evaluare", "RO", ""],
    ["D210 Ada", "ada@example.invalid", "D210 Manager", "manager@example.invalid", "Manager", "RO", "NU"],
  ];
  const parsed = parseParticipantWorkbook(XLSX, workbookBytes(rows, "Respondenti"), "d210-collector-variant.xlsx");
  assert.equal(parsed.type, "collector");
  assert.equal(parsed.participants[0].rows[0].isSelf, true);
  assert.equal(parsed.participants[0].rows[1].isSelf, false);
});

test("Inc13: mass-relations rows also benefit from tolerant role matching", () => {
  const rows = [MASS_TEMPLATE_HEADERS,
    ["D210 Ada", "ada@example.invalid", "D210 Manager", "manager@example.invalid", "Functional-Manager", "RO", "", "", "", "", ""],
    ["D210 Ada", "ada@example.invalid", "D210 Partner", "partner@example.invalid", "Partener_Extern", "RO", "", "", "", "", ""]];
  const parsed = parseParticipantWorkbook(XLSX, workbookBytes(rows), "d210-mass-variant.xlsx");
  assert.equal(parsed.type, "mass");
  const roles = parsed.participants[0].rows.filter((row) => !row.isSelf).map((row) => row.role);
  assert.deepEqual(roles, ["Manager", "PartenerExtern"]);
});

test("Inc13: NumeInregistrare is an & formula with no function call Excel could rewrite with @", async () => {
  const rows = [
    ["ParticipantName", "ParticipantEmail", "EvaluatorName", "EvaluatorEmail", "Relationship", "Language", "IsSelf"],
    ["D210 Ada", "ada@example.invalid", "D210 Ada", "ada@example.invalid", "Autoevaluare", "RO", "DA"],
    ["D210 Ada", "ada@example.invalid", "D210 Manager", "manager@example.invalid", "Manager", "RO", "NU"],
    ["D210 Ada", "ada@example.invalid", "D210 Peer", "peer@example.invalid", "Peer", "RO", "NU"],
  ];
  const parsed = parseParticipantWorkbook(XLSX, workbookBytes(rows, "Respondenti"), "d210-formula.xlsx");
  const analysis = analyzeProject({ participants: parsed.participants, allocation: analyzeAllocation([]), projectName: "D210 Formula" });
  assert.equal(analysis.ready, true);
  const dir = await mkdtemp(join(tmpdir(), "d210-inc13-"));
  const file = join(dir, "import.xlsx");
  try {
    await writeFile(file, XLSX.write(createProductionWorkbook(XLSX, analysis, "D210 Formula"), { type: "buffer", bookType: "xlsx", compression: true, bookSST: true }));
    const sheet = XLSX.read(await readFile(file), { type: "buffer", cellFormula: true }).Sheets.Sheet1;
    for (let r = 2; r <= analysis.outputRows.length + 1; r++) {
      assert.equal(sheet[`N${r}`].f, `G${r}&" - "&B${r}&" - "&D${r}`);
      assert(!/CONCAT|_xlfn|@|\(/u.test(sheet[`N${r}`].f), sheet[`N${r}`].f);
    }
    const xml = execFileSync("unzip", ["-p", file, "xl/worksheets/sheet1.xml"], { encoding: "utf8" });
    const formulas = xml.match(/<f>[^<]*<\/f>/gu) || [];
    assert.equal(formulas.length, analysis.outputRows.length);
    assert(!formulas.some((f) => /CONCAT|_xlfn|@/u.test(f)), formulas.join(""));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("Inc13 and the copy fix survive in the built deploy bundle", async () => {
  const [app, html] = await Promise.all([readFile(resolve(root, "deploy/app.js"), "utf8"), readFile(resolve(root, "deploy/index.html"), "utf8")]);
  assert(app.includes('function roleKey(value) { return fold(value).replace(/[\\s_-]+/gu, ""); }'));
  assert(app.includes('f: `G${excelRow}&" - "&B${excelRow}&" - "&D${excelRow}`'));
  assert(!app.includes("CONCAT("));
  for (const phrase of ["Nicio încărcare online, fără analize de utilizare", "Aceeași adresă de email apare cu nume diferite", "Participant's email address", "Correct the file, then remove or replace the source"]) assert(app.includes(phrase), phrase);
  for (const phrase of ["Erorile sunt izolate pentru fiecare participant și sursă.", "Corectează fișierul indicat, apoi folosește opțiunea Înlocuiește", "Nicio încărcare online"]) assert(html.includes(phrase), phrase);
});

test("project placeholder is generic in RO and EN (Vlad, 2026-09-24)", async () => {
  const [app, html, deployHtml] = await Promise.all(["src/app.js", "src/index.html", "deploy/index.html"].map((f) => readFile(resolve(root, f), "utf8")));
  assert(deployHtml.includes('placeholder="ex. Proiect 360 2026"'));
  assert(html.includes('placeholder="ex. Proiect 360 2026" data-i18n-placeholder="projectPlaceholder"'));
  assert(app.includes('projectPlaceholder: "ex. Proiect 360 2026"') && app.includes('projectPlaceholder: "e.g. 360 Project 2026"'));
  assert(app.includes("node.placeholder = tr(node.dataset.i18nPlaceholder)"));
});
