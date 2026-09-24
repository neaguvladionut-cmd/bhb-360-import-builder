import assert from "node:assert/strict";
import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import { extname, resolve } from "node:path";
import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";
import test from "node:test";

const require = createRequire(import.meta.url);
const { chromium } = require(process.env.BHB_PLAYWRIGHT_MODULE);
const root = resolve("deploy");
const prefix = "/pages/d210-importer/";
const mime = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8", ".png": "image/png", ".webp": "image/webp", ".ttf": "font/ttf", ".txt": "text/plain; charset=utf-8" };

function startServer() {
  const server = createServer(async (request, response) => {
    const url = new URL(request.url, "http://127.0.0.1");
    if (!url.pathname.startsWith(prefix)) { response.writeHead(404).end(); return; }
    const relative = url.pathname.slice(prefix.length) || "index.html";
    if (relative.includes("..")) { response.writeHead(400).end(); return; }
    try { const bytes = await readFile(resolve(root, relative)); response.writeHead(200, { "content-type": mime[extname(relative)] || "application/octet-stream" }); response.end(bytes); }
    catch { response.writeHead(404).end(); }
  });
  return new Promise((resolveReady) => server.listen(0, "127.0.0.1", () => resolveReady(server)));
}

async function openImporter(browser, origin, pageErrors) {
  const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
  page.on("pageerror", (error) => pageErrors.push(error.message));
  await page.goto(`${origin}${prefix}`, { waitUntil: "networkidle" });
  await page.getByRole("button", { name: "EN", exact: true }).click();
  await page.locator("#project-name").fill("D210 Failure Boundary");
  return page;
}

async function uploadParticipant(page, name = "d210-participant.xlsx") {
  await page.evaluate((filename) => {
    const rows = [["ParticipantName","ParticipantEmail","EvaluatorName","EvaluatorEmail","Relationship","Language","IsSelf"],["D210 Boundary Participant","d210.boundary@example.invalid","D210 Boundary Participant","d210.boundary@example.invalid","Autoevaluare","RO","DA"],["D210 Boundary Participant","d210.boundary@example.invalid","D210 Boundary Manager","d210.boundary.manager@example.invalid","Manager","RO","NU"]];
    const workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), "Respondenti");
    const file = new File([XLSX.write(workbook, { type: "array", bookType: "xlsx" })], filename, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const transfer = new DataTransfer(); transfer.items.add(file); const input = document.querySelector("#source-files"); input.files = transfer.files; input.dispatchEvent(new Event("change", { bubbles: true }));
  }, name);
}

async function uploadAllocation(page, name = "d210-allocation.xlsx") {
  await page.evaluate((filename) => {
    const rows = [["Identifier","Assessor Name","Assessor Email","QuestionaireId","Assessed Email","Assessor Role"],["77","D210 Boundary Manager","d210.boundary.manager@example.invalid","D210-Q","d210.boundary@example.invalid","Manager"]];
    const workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), "Accounts");
    const file = new File([XLSX.write(workbook, { type: "array", bookType: "xlsx" })], filename, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const transfer = new DataTransfer(); transfer.items.add(file); const input = document.querySelector("#source-files"); input.files = transfer.files; input.dispatchEvent(new Event("change", { bubbles: true }));
  }, name);
}

async function auditSheets(page) {
  const downloadEvent = page.waitForEvent("download");
  await page.locator("#download-audit").click();
  const download = await downloadEvent;
  const bytes = await readFile(await download.path());
  return page.evaluate((values) => {
    const workbook = XLSX.read(new Uint8Array(values), { type: "array", cellFormula: false });
    return Object.fromEntries(workbook.SheetNames.map((name) => [name, XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: "" })]));
  }, [...bytes]);
}

async function dispatchUnreadable(page, inputSelector, filename, cardName = "", delayMs = 0) {
  await page.evaluate(({ inputSelector: selector, filename: name, cardName: target, delay }) => {
    if (!globalThis.__d210OriginalArrayBuffer) globalThis.__d210OriginalArrayBuffer = File.prototype.arrayBuffer;
    File.prototype.arrayBuffer = function arrayBuffer() { if (this.name === name) return new Promise((resolve, reject) => { if (delay < 0) globalThis.__d210RejectRead = reject; else setTimeout(() => reject(new DOMException("D210 synthetic read failure", "NotReadableError")), delay); }); return globalThis.__d210OriginalArrayBuffer.call(this); };
    const file = new File(["D210 unreadable fixture"], name, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const transfer = new DataTransfer(); transfer.items.add(file);
    const input = target ? [...document.querySelectorAll(".source-card")].find((card) => card.querySelector("strong")?.textContent === target).querySelector(selector) : document.querySelector(selector);
    input.files = transfer.files; input.dispatchEvent(new Event("change", { bubbles: true }));
  }, { inputSelector, filename, cardName, delay: delayMs });
}

test("nested importer runtime reviews D210 workbook locally and reset clears memory", async () => {
  const server = await startServer();
  const address = server.address();
  const origin = `http://127.0.0.1:${address.port}`;
  const requests = []; const consoleMessages = []; const pageErrors = [];
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
    page.on("request", (request) => requests.push(request.url()));
    page.on("console", (message) => consoleMessages.push(message.text()));
    page.on("pageerror", (error) => pageErrors.push(error.message));
    await page.goto(`${origin}${prefix}`, { waitUntil: "networkidle" });
    await page.getByRole("button", { name: "EN", exact: true }).click();
    await page.locator("#project-name").fill("D210 Runtime Project");
    await page.evaluate(() => {
      const rows = [
        ["ParticipantName","ParticipantEmail","EvaluatorName","EvaluatorEmail","Relationship","Language","IsSelf"],
        ["D210 Runtime Participant","d210.runtime@example.invalid","D210 Runtime Participant","d210.runtime@example.invalid","Autoevaluare","RO","DA"],
        ["D210 Runtime Participant","d210.runtime@example.invalid","D210 Runtime Manager","d210.manager@example.invalid","Manager","EN","NU"],
        ["D210 Runtime Participant","d210.runtime@example.invalid","D210 Runtime Peer","d210.peer@example.invalid","Peer","RO","NU"],
      ];
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), "Respondenti");
      const file = new File([XLSX.write(workbook, { type: "array", bookType: "xlsx" })], "d210-runtime.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const transfer = new DataTransfer(); transfer.items.add(file);
      const input = document.querySelector("#source-files"); input.files = transfer.files; input.dispatchEvent(new Event("change", { bubbles: true }));
    });
    await page.locator("#readiness-pill").getByText("Ready to export", { exact: true }).waitFor();
    const acceptedSource = page.locator(".source-card").filter({ hasText: "d210-runtime.xlsx" });
    assert.equal((await acceptedSource.locator(".fingerprint").textContent()).length, 64);
    assert((await acceptedSource.textContent()).includes("accepted"));
    const review = await page.locator("#review").textContent();
    for (const expected of ["Roles · Manager", "Roles · Peer", "Languages · EN", "D210 Runtime Manager", "d210.manager@example.invalid", "d210-runtime.xlsx"]) assert(review.includes(expected), expected);
    const secrets = ["D210 Runtime Project", "D210 Runtime Participant", "d210.runtime@example.invalid", "d210.manager@example.invalid"];
    assert.equal(pageErrors.length, 0, pageErrors.join("\n"));
    assert(requests.every((value) => { const url = new URL(value); return url.origin === origin && url.pathname.startsWith(prefix); }), requests.join("\n"));
    assert(secrets.every((secret) => !requests.some((value) => decodeURIComponent(value).includes(secret))));
    assert(secrets.every((secret) => !consoleMessages.some((value) => value.includes(secret))));
    assert(secrets.every((secret) => !page.url().includes(secret)));
    const persistence = await page.evaluate(async () => ({ local: localStorage.length, session: sessionStorage.length, indexed: indexedDB.databases ? (await indexedDB.databases()).length : 0, caches: globalThis.caches ? (await caches.keys()).length : 0 }));
    assert.deepEqual(persistence, { local: 0, session: 0, indexed: 0, caches: 0 });
    await page.evaluate(() => {
      const file = new File(["not an xlsx"], "d210-bad.xlsx", { type: "application/octet-stream" });
      const transfer = new DataTransfer(); transfer.items.add(file);
      const input = document.querySelector("#source-files"); input.files = transfer.files; input.dispatchEvent(new Event("change", { bubbles: true }));
    });
    await page.locator(".source-card").filter({ hasText: "d210-bad.xlsx" }).getByText("rejected", { exact: true }).waitFor();
    assert.equal(await page.locator("#download").isDisabled(), true);
    assert.equal(await page.locator("#download-audit").isEnabled(), true);
    const auditDownload = page.waitForEvent("download");
    await page.locator("#download-audit").click();
    assert.match((await auditDownload).suggestedFilename(), /^audit-360-D210-Runtime-Project\.xlsx$/u);
    await page.locator(".source-card").filter({ hasText: "d210-bad.xlsx" }).getByRole("button", { name: "Remove" }).click();
    await page.locator("#readiness-pill").getByText("Ready to export", { exact: true }).waitFor();
    await page.evaluate(() => {
      const file = new File(["broken replacement"], "d210-runtime.xlsx", { type: "application/octet-stream" });
      const transfer = new DataTransfer(); transfer.items.add(file);
      const card = [...document.querySelectorAll(".source-card")].find((item) => item.querySelector("strong")?.textContent === "d210-runtime.xlsx");
      const input = card.querySelector(".source-replace-input"); input.files = transfer.files; input.dispatchEvent(new Event("change", { bubbles: true }));
    });
    await page.locator(".source-card").filter({ hasText: "d210-runtime.xlsx" }).getByText("rejected", { exact: true }).waitFor();
    assert.equal(await page.locator(".source-card").count(), 1, "replacement must remove stale accepted source state");
    await page.evaluate(() => {
      const rows = [["ParticipantName","ParticipantEmail","EvaluatorName","EvaluatorEmail","Relationship","Language","IsSelf"],["D210 Runtime Participant","d210.runtime@example.invalid","D210 Runtime Participant","d210.runtime@example.invalid","Autoevaluare","RO","DA"],["D210 Runtime Participant","d210.runtime@example.invalid","D210 Runtime Manager","d210.manager@example.invalid","Manager","EN","NU"]];
      const workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), "Respondenti");
      const file = new File([XLSX.write(workbook, { type: "array", bookType: "xlsx" })], "d210-runtime.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const transfer = new DataTransfer(); transfer.items.add(file);
      const card = [...document.querySelectorAll(".source-card")].find((item) => item.querySelector("strong")?.textContent === "d210-runtime.xlsx");
      const input = card.querySelector(".source-replace-input"); input.files = transfer.files; input.dispatchEvent(new Event("change", { bubbles: true }));
    });
    await page.locator("#readiness-pill").getByText("Ready to export", { exact: true }).waitFor();
    assert.equal(await page.locator(".source-card").count(), 1, "corrected source must replace the rejected attempt");
    page.once("dialog", (dialog) => dialog.accept());
    await page.locator("#reset").click();
    assert.equal(await page.locator("#project-name").inputValue(), "");
    const body = await page.locator("body").textContent();
    assert(secrets.every((secret) => !body.includes(secret)));
  } finally { await browser.close(); await new Promise((resolveClose) => server.close(resolveClose)); }
});

test("first unreadable participant attempt stays visible, blocks A:N and enables an honest audit", async () => {
  const server = await startServer(); const address = server.address(); const origin = `http://127.0.0.1:${address.port}`;
  const pageErrors = []; const browser = await chromium.launch({ headless: true });
  try {
    const page = await openImporter(browser, origin, pageErrors);
    await dispatchUnreadable(page, "#source-files", "d210-first-unreadable.xlsx");
    const card = page.locator(".source-card").filter({ hasText: "d210-first-unreadable.xlsx" });
    await card.getByText("rejected", { exact: true }).waitFor();
    assert.equal(await card.locator(".fingerprint").textContent(), "unavailable");
    assert.equal(await page.locator("#download").isDisabled(), true);
    assert.equal(await page.locator("#download-audit").isEnabled(), true);
    assert.equal(pageErrors.length, 0, pageErrors.join("\n"));
    const audit = await auditSheets(page);
    assert.deepEqual(Object.keys(audit), ["Run","Sources","Issues","Identifiers","Normalizations"]);
    assert.equal(audit.Sources[1][0], "d210-first-unreadable.xlsx");
    assert.equal(audit.Sources[1][2], "unavailable");
    assert(audit.Issues[1][5].includes('"failureStage":"read"'), audit.Issues[1][5]);
    assert.equal(audit.Identifiers.length, 1, "failed first attempt must not mint identifiers");
  } finally { await browser.close(); await new Promise((resolveClose) => server.close(resolveClose)); }
});

test("unreadable allocation replacement invalidates stale allocation identifiers before reading", async () => {
  const server = await startServer(); const address = server.address(); const origin = `http://127.0.0.1:${address.port}`;
  const pageErrors = []; const browser = await chromium.launch({ headless: true });
  try {
    const page = await openImporter(browser, origin, pageErrors);
    await uploadParticipant(page); await page.locator(".source-card").filter({ hasText: "d210-participant.xlsx" }).getByText("accepted", { exact: true }).waitFor();
    await uploadAllocation(page); await page.locator(".source-card").filter({ hasText: "d210-allocation.xlsx" }).getByText("accepted", { exact: true }).waitFor();
    assert.equal(await page.locator(".summary-card").filter({ hasText: "Reused identifiers" }).locator("b").textContent(), "1");
    await dispatchUnreadable(page, ".source-replace-input", "d210-allocation-unreadable.xlsx", "d210-allocation.xlsx", -1);
    const card = page.locator(".source-card").filter({ hasText: "d210-allocation-unreadable.xlsx" });
    await card.getByText("checking", { exact: true }).waitFor();
    assert.equal(await page.locator(".source-card").filter({ hasText: "d210-allocation.xlsx" }).count(), 0, "old allocation must disappear before file read settles");
    assert.equal(await page.locator(".summary-card").filter({ hasText: "Reused identifiers" }).locator("b").textContent(), "0", "old identifiers must clear before file read settles");
    assert.equal(await page.locator("#download").isDisabled(), true);
    assert.equal(await card.locator(".fingerprint").textContent(), "unavailable (pending)");
    const pendingCardText = await card.textContent();
    assert(pendingCardText.includes("Local checking is in progress; production export remains blocked."));
    assert(!pendingCardText.includes("could not be read"));
    assert(!pendingCardText.includes("rejected"));
    const pendingAudit = await auditSheets(page);
    const pendingSource = pendingAudit.Sources.find((row) => row[0] === "d210-allocation-unreadable.xlsx");
    assert.equal(pendingSource[2], "unavailable (pending)");
    assert.equal(pendingSource[3], "pending");
    assert.equal(pendingSource[4], "pending");
    assert.equal(pendingSource[11], "0");
    const pendingIssue = pendingAudit.Issues.find((row) => row[2] === "d210-allocation-unreadable.xlsx");
    assert.equal(pendingIssue[1], "source-pending");
    assert.equal(pendingIssue[5], "{}");
    assert(pendingIssue[6].includes("checking is in progress"));
    assert(!pendingAudit.Issues.flat().join(" ").includes("source-rejected"));
    assert(!pendingAudit.Issues.flat().join(" ").includes("source-read-failed"));
    assert(!pendingAudit.Issues.flat().join(" ").includes("could not be read"));
    assert(!pendingAudit.Identifiers.flat().includes("77"), "pending audit must not retain the replaced allocation identifier");
    await page.evaluate(() => globalThis.__d210RejectRead(new DOMException("D210 synthetic read failure", "NotReadableError")));
    await card.getByText("rejected", { exact: true }).waitFor();
    assert.equal(await page.locator(".source-card").filter({ hasText: "d210-allocation.xlsx" }).count(), 0);
    assert.equal(await card.locator(".fingerprint").textContent(), "unavailable");
    assert.equal(await page.locator(".summary-card").filter({ hasText: "Reused identifiers" }).locator("b").textContent(), "0");
    assert.equal(await page.locator("#download").isDisabled(), true);
    assert.equal(await page.locator("#download-audit").isEnabled(), true);
    assert.equal(pageErrors.length, 0, pageErrors.join("\n"));
    const audit = await auditSheets(page);
    assert(audit.Sources.some((row) => row[0] === "d210-allocation-unreadable.xlsx" && row[2] === "unavailable" && row[3] === "rejected" && row[4] === "blocked"));
    assert(!audit.Identifiers.flat().includes("77"), "replaced allocation identifier must not survive in audit state");
    assert(audit.Issues.some((row) => row[1] === "source-rejected"));
    assert(audit.Issues.some((row) => String(row[5]).includes('"failureStage":"read"')));
  } finally { await browser.close(); await new Promise((resolveClose) => server.close(resolveClose)); }
});

test("SHA-256 digest rejection is contained as a blocked source attempt", async () => {
  const server = await startServer(); const address = server.address(); const origin = `http://127.0.0.1:${address.port}`;
  const pageErrors = []; const browser = await chromium.launch({ headless: true });
  try {
    const page = await openImporter(browser, origin, pageErrors);
    await page.evaluate(() => { Object.defineProperty(crypto.subtle, "digest", { configurable: true, value: async () => { throw new DOMException("D210 synthetic digest failure", "OperationError"); } }); });
    await uploadParticipant(page, "d210-digest-rejection.xlsx");
    const card = page.locator(".source-card").filter({ hasText: "d210-digest-rejection.xlsx" });
    await card.getByText("rejected", { exact: true }).waitFor();
    assert.equal(await card.locator(".fingerprint").textContent(), "unavailable");
    assert.equal(await page.locator("#download").isDisabled(), true);
    assert.equal(await page.locator("#download-audit").isEnabled(), true);
    assert.equal(pageErrors.length, 0, pageErrors.join("\n"));
    const audit = await auditSheets(page);
    assert.equal(audit.Sources[1][2], "unavailable");
    assert(audit.Issues[1][5].includes('"failureStage":"digest"'), audit.Issues[1][5]);
    assert.equal(audit.Identifiers.length, 1, "digest failure must not mint identifiers");
  } finally { await browser.close(); await new Promise((resolveClose) => server.close(resolveClose)); }
});

test("one project-workbook picker auto-fills the name from an existing import", async () => {
  const server = await startServer(); const address = server.address(); const origin = `http://127.0.0.1:${address.port}`;
  const pageErrors = []; const browser = await chromium.launch({ headless: true });
  try {
    const page = await openImporter(browser, origin, pageErrors);
    await page.locator("#project-name").fill("");
    await page.evaluate(() => {
      const rows = [
        ["Marca","PersoanaEvaluata","PersoanaEvaluataEmail","Evaluator","EvaluatorEmail","EvaluatorRol","NumeCampanie","Limba","Criteriu1","Criteriu2","Criteriu3","Criteriu4","Criteriu5","NumeInregistrare"],
        ["77","D210 Existing Participant","d210.existing@example.invalid","D210 Existing Manager","d210.existing.manager@example.invalid","Manager","D210 Existing Project","RO","","","","","",""]
      ];
      const workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), "Sheet1");
      const file = new File([XLSX.write(workbook, { type: "array", bookType: "xlsx" })], "d210-existing-import.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const transfer = new DataTransfer(); transfer.items.add(file); const input = document.querySelector("#source-files"); input.files = transfer.files; input.dispatchEvent(new Event("change", { bubbles: true }));
    });
    await page.locator(".source-card").filter({ hasText: "d210-existing-import.xlsx" }).getByText("accepted", { exact: true }).waitFor();
    assert.equal(await page.locator("#project-name").inputValue(), "D210 Existing Project");
    assert.equal(await page.locator("#project-auto-note").isVisible(), true);
    assert.equal(await page.locator("#participant-files").count(), 0);
    assert.equal(await page.locator("#allocation-file").count(), 0);
    assert.equal(await page.locator("#mass-template-file").count(), 0);
    assert.equal(pageErrors.length, 0, pageErrors.join("\n"));
  } finally { await browser.close(); await new Promise((resolveClose) => server.close(resolveClose)); }
});

test("import builder initializes when deploy index is opened directly from disk", async () => {
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage(); const errors = [];
    page.on("pageerror", (error) => errors.push(error.message));
    await page.goto(pathToFileURL(resolve(root, "index.html")).href, { waitUntil: "load" });
    await page.locator("#project-name").fill("D210 Direct File");
    assert.equal(await page.locator("#project-name").inputValue(), "D210 Direct File");
    assert.equal(errors.length, 0, errors.join("\n"));
  } finally { await browser.close(); }
});
