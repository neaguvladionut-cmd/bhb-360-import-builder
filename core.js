export const PRODUCTION_HEADERS = ["Marca","PersoanaEvaluata","PersoanaEvaluataEmail","Evaluator","EvaluatorEmail","EvaluatorRol","NumeCampanie","Limba","Criteriu1","Criteriu2","Criteriu3","Criteriu4","Criteriu5","NumeInregistrare"];
export const COLLECTOR_HEADERS = ["ParticipantName","ParticipantEmail","EvaluatorName","EvaluatorEmail","Relationship","Language","IsSelf"];
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/u;
const LANGUAGE_RE = /^[A-Z]{2}$/u;
const ROLE_ORDER = { Manager: 0, Peer: 1, Subordonat: 2, PartenerExtern: 3 };

export function normalizeText(value) {
  return String(value ?? "").replace(/[\u00a0\u2007\u202f]/gu, " ").replace(/\s+/gu, " ").trim();
}
export function normalizeEmail(value) { return normalizeText(value).toLocaleLowerCase("en-US"); }
export function normalizeLanguage(value) { return normalizeText(value).toUpperCase(); }
export function isValidEmail(value) { return EMAIL_RE.test(normalizeEmail(value)); }
function fold(value) { return normalizeText(value).normalize("NFD").replace(/[\u0300-\u036f]/gu, "").toLocaleLowerCase("en-US"); }

export function mapRole(value) {
  const key = fold(value);
  if (key === "autoevaluare") return "Manager";
  if (key === "manager" || key === "functional manager") return "Manager";
  if (key === "peer" || key === "coleg") return "Peer";
  if (key === "subordonat") return "Subordonat";
  if (key === "stakeholder" || key === "partener" || key === "partenerextern" || key === "partener extern") return "PartenerExtern";
  return null;
}

function rowsFromSheet(XLSX, sheet) { return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false, blankrows: false }); }
function exactHeaderIndex(rows, headers) {
  return rows.findIndex((row) => headers.every((header, index) => normalizeText(row[index]) === header));
}

function recordNormalization(normalizations, source, rowNumber, field, original, normalized) {
  const before = String(original ?? "");
  const after = String(normalized ?? "");
  if (before !== after) normalizations.push({ source, rowNumber, field, original: before, normalized: after });
}

function parseCollectorRows(rows, headerIndex, sourceName, normalizations) {
  const sourceRows = rows.slice(headerIndex + 1);
  const contentRows = sourceRows.map((row, index) => ({ row, rowNumber: headerIndex + index + 2 })).filter(({ row }) => row.some((value) => normalizeText(value)));
  const objects = contentRows.map(({ row, rowNumber }) => {
    const sourceRole = normalizeText(row[4]);
    const participantEmail = normalizeEmail(row[1]);
    const evaluatorEmail = normalizeEmail(row[3]);
    const values = [
      ["ParticipantName", row[0], normalizeText(row[0])], ["ParticipantEmail", row[1], participantEmail],
      ["EvaluatorName", row[2], normalizeText(row[2])], ["EvaluatorEmail", row[3], evaluatorEmail],
      ["Relationship", row[4], sourceRole], ["Language", row[5], normalizeLanguage(row[5])],
      ["ProductionRole", sourceRole, mapRole(sourceRole) || sourceRole],
    ];
    values.forEach(([field, original, normalized]) => recordNormalization(normalizations, sourceName, rowNumber, field, original, normalized));
    return {
      participantName: normalizeText(row[0]), participantEmail,
      evaluatorName: normalizeText(row[2]), evaluatorEmail,
      sourceRole, role: mapRole(sourceRole), language: normalizeLanguage(row[5]),
      isSelf: fold(row[6]) === "da" || fold(row[6]) === "yes" || sourceRole === "Autoevaluare" || (participantEmail && participantEmail === evaluatorEmail),
      source: sourceName, rowNumber, legacy: false,
    };
  });
  const grouped = new Map();
  for (const row of objects) {
    const key = row.participantEmail || `missing:${row.participantName}`;
    if (!grouped.has(key)) grouped.set(key, { id: `${sourceName}:${key}`, source: sourceName, adapter: "collector", participantName: row.participantName, participantEmail: row.participantEmail, rows: [] });
    grouped.get(key).rows.push(row);
  }
  return { participants: [...grouped.values()], ignoredRows: sourceRows.length - contentRows.length };
}

function parseLegacyRows(rows, headerIndex, sourceName, normalizations) {
  const header = rows[headerIndex].map(fold);
  const lastNameIndex = header.indexOf("nume");
  const firstNameIndex = header.indexOf("prenume");
  const emailIndex = header.indexOf("email");
  const roleIndex = header.indexOf("rol");
  const sourceRows = rows.slice(headerIndex + 1);
  const parsedRows = sourceRows.map((row, index) => {
    const sourceRole = normalizeText(row[roleIndex]);
    if (!sourceRole && !normalizeText(row[emailIndex]) && !normalizeText(row[lastNameIndex]) && !normalizeText(row[firstNameIndex])) return null;
    const rowNumber = headerIndex + index + 2;
    const evaluatorName = normalizeText([row[firstNameIndex], row[lastNameIndex]].filter(Boolean).join(" "));
    const evaluatorEmail = normalizeEmail(row[emailIndex]);
    recordNormalization(normalizations, sourceName, rowNumber, "EvaluatorName", [row[firstNameIndex], row[lastNameIndex]].filter(Boolean).join(" "), evaluatorName);
    recordNormalization(normalizations, sourceName, rowNumber, "EvaluatorEmail", row[emailIndex], evaluatorEmail);
    recordNormalization(normalizations, sourceName, rowNumber, "Relationship", row[roleIndex], sourceRole);
    recordNormalization(normalizations, sourceName, rowNumber, "ProductionRole", sourceRole, mapRole(sourceRole) || sourceRole);
    return {
      evaluatorName, evaluatorEmail, sourceRole, role: mapRole(sourceRole), language: "RO",
      isSelf: fold(sourceRole) === "autoevaluare", source: sourceName, rowNumber, legacy: true,
    };
  }).filter(Boolean);
  const self = parsedRows.find((row) => row.isSelf);
  const participantName = self?.evaluatorName || "";
  const participantEmail = self?.evaluatorEmail || "";
  parsedRows.forEach((row) => { row.participantName = participantName; row.participantEmail = participantEmail; });
  return { participants: [{ id: `${sourceName}:legacy`, source: sourceName, adapter: "legacy", participantName, participantEmail, rows: parsedRows, legacyLanguageAssumed: true }], ignoredRows: sourceRows.length - parsedRows.length };
}

export function parseParticipantWorkbook(XLSX, data, sourceName) {
  const workbook = XLSX.read(data, { type: data instanceof ArrayBuffer ? "array" : "buffer", cellFormula: false, cellHTML: false });
  for (const sheetName of workbook.SheetNames) {
    const rows = rowsFromSheet(XLSX, workbook.Sheets[sheetName]);
    const collectorHeader = exactHeaderIndex(rows, COLLECTOR_HEADERS);
    if (collectorHeader >= 0) {
      const normalizations = [];
      const parsed = parseCollectorRows(rows, collectorHeader, sourceName, normalizations);
      return { type: "collector", participants: parsed.participants, warnings: [], ignoredRows: parsed.ignoredRows, normalizations };
    }
  }
  for (const sheetName of workbook.SheetNames) {
    const rows = rowsFromSheet(XLSX, workbook.Sheets[sheetName]);
    const headerIndex = rows.findIndex((row) => {
      const values = row.map(fold);
      return values.includes("nume") && values.includes("prenume") && values.includes("email") && values.includes("rol");
    });
    if (headerIndex >= 0) {
      const normalizations = [];
      const parsed = parseLegacyRows(rows, headerIndex, sourceName, normalizations);
      return { type: "legacy", participants: parsed.participants, warnings: [{ code: "legacy-language", source: sourceName }], ignoredRows: parsed.ignoredRows, normalizations };
    }
  }
  throw new Error("unsupported-participant-workbook");
}

export function parseAllocationWorkbook(XLSX, data, sourceName = "allocation.xlsx") {
  const workbook = XLSX.read(data, { type: data instanceof ArrayBuffer ? "array" : "buffer", cellFormula: false, cellHTML: false });
  const sheet = workbook.Sheets.Accounts;
  if (!sheet) throw new Error("missing-accounts-sheet");
  const rows = rowsFromSheet(XLSX, sheet);
  const expected = ["Identifier","Assessor Name","Assessor Email","QuestionaireId","Assessed Email","Assessor Role"];
  const headerIndex = exactHeaderIndex(rows, expected);
  if (headerIndex < 0) throw new Error("invalid-allocation-headers");
  const sourceRows = rows.slice(headerIndex + 1);
  const contentRows = sourceRows.map((row, index) => ({ row, rowNumber: headerIndex + index + 2 })).filter(({ row }) => row.some((value) => normalizeText(value)));
  const normalizations = [];
  const records = contentRows.map(({ row, rowNumber }) => {
    const record = { identifier: normalizeText(row[0]), name: normalizeText(row[1]), email: normalizeEmail(row[2]), source: sourceName, rowNumber };
    recordNormalization(normalizations, sourceName, rowNumber, "Identifier", row[0], record.identifier);
    recordNormalization(normalizations, sourceName, rowNumber, "Assessor Name", row[1], record.name);
    recordNormalization(normalizations, sourceName, rowNumber, "Assessor Email", row[2], record.email);
    return record;
  });
  return { ...analyzeAllocation(records), ignoredRows: sourceRows.length - contentRows.length, normalizations };
}

export function analyzeAllocation(records = []) {
  const errors = [];
  const byEmail = new Map();
  const byId = new Map();
  let maxId = 0;
  for (const record of records) {
    if (!record.identifier || !isValidEmail(record.email)) {
      errors.push({ code: "allocation-invalid", source: record.source, rowNumber: record.rowNumber });
      continue;
    }
    maxId = Math.max(maxId, /^\d+$/u.test(record.identifier) ? Number(record.identifier) : 0);
    if (!byEmail.has(record.email)) byEmail.set(record.email, { identifiers: new Set(), names: new Set(), records: [] });
    byEmail.get(record.email).identifiers.add(record.identifier);
    if (record.name) byEmail.get(record.email).names.add(record.name);
    byEmail.get(record.email).records.push(record);
    if (!byId.has(record.identifier)) byId.set(record.identifier, new Set());
    byId.get(record.identifier).add(record.email);
  }
  for (const [email, info] of byEmail) if (info.identifiers.size > 1) errors.push({
    code: "email-multiple-identifiers", email, identifiers: [...info.identifiers].sort(),
    records: info.records.map(({ identifier, email: recordEmail, name, source, rowNumber }) => ({ identifier, email: recordEmail, name, source, rowNumber })),
    correctiveAction: "Corectează exportul astfel încât acest email să aibă un singur identificator.",
  });
  for (const [identifier, emails] of byId) if (emails.size > 1) errors.push({
    code: "identifier-multiple-emails", identifier, emails: [...emails].sort(),
    records: records.filter((record) => record.identifier === identifier).map(({ identifier: recordIdentifier, email, name, source, rowNumber }) => ({ identifier: recordIdentifier, email, name, source, rowNumber })),
    correctiveAction: "Corectează exportul astfel încât identificatorul să aparțină unui singur email.",
  });
  return { records, byEmail, byId, maxId, errors };
}

function normalizeParticipant(participant) {
  const participantName = normalizeText(participant.participantName);
  const participantEmail = normalizeEmail(participant.participantEmail);
  return {
    ...participant, participantName, participantEmail,
    rows: (participant.rows || []).map((row) => {
      const isSelf = Boolean(row.isSelf) || (participantEmail && normalizeEmail(row.evaluatorEmail) === participantEmail);
      return {
        ...row, participantName, participantEmail,
        evaluatorName: isSelf ? participantName : normalizeText(row.evaluatorName),
        evaluatorEmail: isSelf ? participantEmail : normalizeEmail(row.evaluatorEmail),
        sourceRole: normalizeText(row.sourceRole), role: mapRole(row.sourceRole),
        language: normalizeLanguage(row.language || "RO"), isSelf,
      };
    }),
  };
}

function mergeParticipants(participants) {
  const merged = new Map();
  participants.map(normalizeParticipant).forEach((participant) => {
    const key = participant.participantEmail || `missing:${participant.id}`;
    if (!merged.has(key)) merged.set(key, { ...participant, sources: [participant.source], rows: [...participant.rows] });
    else { const target = merged.get(key); target.rows.push(...participant.rows); target.sources.push(participant.source); if (!target.participantName) target.participantName = participant.participantName; }
  });
  return [...merged.values()];
}

function issue(code, participant, row, extra = {}) { return { code, participant: participant?.participantName || participant?.source || "—", source: row?.source || participant?.source || "—", rowNumber: row?.rowNumber, ...extra }; }

function addNameVariant(store, email, name, source, rowNumber, origin) {
  const normalized = normalizeText(name);
  if (!normalized) return;
  if (!store.has(email)) store.set(email, new Map());
  const variants = store.get(email);
  if (!variants.has(normalized)) variants.set(normalized, { name: normalized, provenance: [] });
  const provenance = variants.get(normalized).provenance;
  const item = { source: source || "—", rowNumber: rowNumber || null, origin };
  if (!provenance.some((entry) => entry.source === item.source && entry.rowNumber === item.rowNumber && entry.origin === item.origin)) provenance.push(item);
}

export function analyzeProject({ participants = [], allocation = analyzeAllocation([]), nameChoices = {}, projectName = "", sourceBlockers = [] }) {
  const blockers = [...sourceBlockers, ...(allocation?.errors || [])];
  const warnings = [];
  const normalizedParticipants = mergeParticipants(participants);
  const validRows = [];
  const participantSummaries = [];

  if (!normalizeText(projectName)) blockers.push({ code: "project-required" });
  if (!normalizedParticipants.length) blockers.push({ code: "files-required" });

  for (const participant of normalizedParticipants) {
    const local = [];
    if (!participant.participantName) local.push(issue("participant-name-missing", participant));
    if (!isValidEmail(participant.participantEmail)) local.push(issue("participant-email-invalid", participant));
    const selfRows = participant.rows.filter((row) => row.isSelf);
    const nonSelf = participant.rows.filter((row) => !row.isSelf);
    if (selfRows.length !== 1) local.push(issue("self-count", participant, null, { count: selfRows.length }));
    if (nonSelf.length > 15) local.push(issue("too-many", participant, null, { count: nonSelf.length }));
    const pairs = new Set();
    let managerCount = 0;
    for (const row of participant.rows) {
      if (!row.evaluatorName) local.push(issue("evaluator-name-missing", participant, row));
      if (!isValidEmail(row.evaluatorEmail)) local.push(issue("evaluator-email-invalid", participant, row));
      if (!row.role) local.push(issue("unknown-role", participant, row, { role: row.sourceRole }));
      if (!LANGUAGE_RE.test(row.language)) local.push(issue("language-invalid", participant, row, { language: row.language }));
      const pair = `${participant.participantEmail}|${row.evaluatorEmail}`;
      if (pairs.has(pair)) local.push(issue("duplicate-allocation", participant, row, { email: row.evaluatorEmail }));
      pairs.add(pair);
      if (!row.isSelf && row.role === "Manager" && row.evaluatorEmail !== participant.participantEmail) managerCount += 1;
      if (row.language === "EN") warnings.push(issue("english", participant, row, { evaluator: row.evaluatorName, email: row.evaluatorEmail }));
    }
    if (managerCount === 0) local.push(issue("manager-required", participant));
    blockers.push(...local);
    if (!local.length) validRows.push(...participant.rows);
    participantSummaries.push({ key: participant.participantEmail || participant.id, participant, blockers: local, selfCount: selfRows.length, managerCount, rowCount: participant.rows.length });
  }

  const namesByEmail = new Map();
  for (const row of validRows) {
    addNameVariant(namesByEmail, row.evaluatorEmail, row.evaluatorName, row.source, row.rowNumber, "participant");
  }
  const usedEmails = new Set(validRows.map((row) => row.evaluatorEmail));
  for (const [email, info] of allocation.byEmail || []) {
    if (!usedEmails.has(email)) continue;
    for (const record of info.records) addNameVariant(namesByEmail, email, record.name, record.source, record.rowNumber, "allocation");
  }
  const conflicts = [];
  const canonicalNames = {};
  for (const [email, names] of namesByEmail) {
    const variants = [...names.values()].map((variant) => ({
      ...variant,
      provenance: [...variant.provenance].sort((a, b) => a.source.localeCompare(b.source, "ro") || (a.rowNumber || 0) - (b.rowNumber || 0) || a.origin.localeCompare(b.origin)),
    })).sort((a, b) => a.name.localeCompare(b.name, "ro", { sensitivity: "variant" }));
    const allocationVariantNames = new Set(variants.filter((variant) => variant.provenance.some((entry) => entry.origin === "allocation")).map((variant) => variant.name));
    variants.forEach((variant) => { variant.recommended = allocationVariantNames.has(variant.name); });
    const defaultName = [...allocationVariantNames].sort((a,b) => a.localeCompare(b,"ro", { sensitivity: "variant" }))[0] || variants[0]?.name || "";
    if (variants.length > 1) {
      const selected = normalizeText(nameChoices[email]);
      const resolved = variants.some((variant) => variant.name === selected);
      conflicts.push({ email, names: variants.map((variant) => variant.name), variants, defaultName, resolved, selected: resolved ? selected : "" });
      if (!resolved) blockers.push({ code: "name-conflict", email, names: variants.map((variant) => variant.name) });
      else canonicalNames[email] = selected;
    } else canonicalNames[email] = variants[0]?.name || "";
  }

  const allEmails = [...new Set(validRows.map((row) => row.evaluatorEmail))].filter(isValidEmail).sort();
  const identifiers = {};
  const identifierEvidence = {};
  let nextIdentifier = allocation.maxId + 1;
  let reusedCount = 0;
  for (const email of allEmails) {
    const existing = allocation.byEmail?.get(email);
    if (existing && existing.identifiers.size === 1) {
      identifiers[email] = [...existing.identifiers][0]; reusedCount += 1;
      identifierEvidence[email] = { identifier: identifiers[email], disposition: "reused", records: existing.records.map(({ source, rowNumber }) => ({ source, rowNumber })) };
    }
  }
  for (const email of allEmails) if (!identifiers[email]) {
    identifiers[email] = String(nextIdentifier++);
    identifierEvidence[email] = { identifier: identifiers[email], disposition: "new", records: [] };
  }

  const outputRows = validRows.map((row) => ({ ...row, evaluatorName: canonicalNames[row.evaluatorEmail] || row.evaluatorName, identifier: identifiers[row.evaluatorEmail] })).sort((a,b) => {
    const participant = a.participantEmail.localeCompare(b.participantEmail);
    if (participant) return participant;
    if (a.isSelf !== b.isSelf) return a.isSelf ? -1 : 1;
    const role = (ROLE_ORDER[a.role] ?? 99) - (ROLE_ORDER[b.role] ?? 99);
    return role || a.evaluatorEmail.localeCompare(b.evaluatorEmail);
  });
  const byRole = {}; const byLanguage = {};
  outputRows.forEach((row) => { byRole[row.role] = (byRole[row.role] || 0) + 1; byLanguage[row.language] = (byLanguage[row.language] || 0) + 1; });
  return { blockers, warnings, conflicts, participantSummaries, outputRows, identifiers, identifierEvidence, canonicalNames, summary: { participants: normalizedParticipants.length, rows: outputRows.length, byRole, byLanguage, reusedCount, newCount: allEmails.length - reusedCount, nextIdentifier }, ready: blockers.length === 0 };
}

function textCell(value) { return { t: "s", v: String(value ?? "") }; }
export function createProductionWorkbook(XLSX, analysis, projectName) {
  if (!analysis.ready) throw new Error("Project has blocking errors");
  const sheet = {};
  PRODUCTION_HEADERS.forEach((header, column) => { sheet[XLSX.utils.encode_cell({ r: 0, c: column })] = textCell(header); });
  analysis.outputRows.forEach((row, index) => {
    const excelRow = index + 2;
    const values = [row.identifier,row.participantName,row.participantEmail,row.evaluatorName,row.evaluatorEmail,row.role,normalizeText(projectName),row.language,"","","","",""];
    values.forEach((value, column) => { sheet[XLSX.utils.encode_cell({ r: index + 1, c: column })] = textCell(value); });
    sheet[XLSX.utils.encode_cell({ r: index + 1, c: 13 })] = { t: "str", f: `CONCAT(G${excelRow}," - ",B${excelRow}," - ",D${excelRow})`, v: "" };
  });
  sheet["!ref"] = `A1:N${analysis.outputRows.length + 1}`;
  sheet["!cols"] = [{wch:10},{wch:26},{wch:32},{wch:26},{wch:32},{wch:18},{wch:28},{wch:10},{wch:12},{wch:12},{wch:12},{wch:12},{wch:12},{wch:55}];
  sheet["!autofilter"] = { ref: `A1:N${analysis.outputRows.length + 1}` };
  const workbook = XLSX.utils.book_new();
  workbook.Props = { Title: `360 import - ${normalizeText(projectName)}`, Author: "Business Health Bar", Comments: "Generated locally; structurally validated only." };
  XLSX.utils.book_append_sheet(workbook, sheet, "Sheet1");
  return workbook;
}

function literalSheet(XLSX, rows, widths = []) {
  const sheet = {};
  rows.forEach((row, rowIndex) => row.forEach((value, columnIndex) => {
    sheet[XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })] = textCell(value);
  }));
  const maxColumns = Math.max(1, ...rows.map((row) => row.length));
  sheet["!ref"] = `A1:${XLSX.utils.encode_col(maxColumns - 1)}${Math.max(1, rows.length)}`;
  sheet["!cols"] = widths.map((wch) => ({ wch }));
  if (rows.length) sheet["!autofilter"] = { ref: `A1:${XLSX.utils.encode_col(maxColumns - 1)}${rows.length}` };
  return sheet;
}

export function createAuditWorkbook(XLSX, { projectName = "", analysis, sources = [], generatedAt = "", normalizations = [] }) {
  const runRows = [
    ["Field", "Value"], ["Project", normalizeText(projectName) || "—"], ["Generated locally at", generatedAt || "—"],
    ["State", analysis?.ready ? "ready" : "blocked"], ["Participants", String(analysis?.summary?.participants ?? 0)],
    ["Exportable rows", String(analysis?.summary?.rows ?? 0)], ["Blockers", String(analysis?.blockers?.length ?? 0)],
    ["Warnings", String(analysis?.warnings?.length ?? 0)], ["Validation boundary", "Structurally validated only; production compatibility requires Vlad's controlled import in the current 360 application."],
    ["Privacy", "Client-confidential local receipt. The consultant controls storage and deletion; reset cannot delete a downloaded file."],
  ];
  const sourceRows = [["Filename","Bytes","SHA-256","Adapter","Disposition","Parsed rows","Exportable rows","Blocked rows","Intentionally ignored rows","Normalizations","Warnings","Errors"]];
  sources.forEach((source) => sourceRows.push([
    source.name, String(source.bytes ?? 0), source.sha256 || "", source.adapter || "rejected", source.disposition || "blocked",
    String(source.parsedRows ?? 0), String(source.exportableRows ?? 0), String(source.blockedRows ?? 0), String(source.ignoredRows ?? 0),
    String(source.normalizationCount ?? 0), String(source.warningCount ?? 0), String(source.errorCount ?? 0),
  ]));
  const issueRows = [["Severity","Code","Person / source","Source","Row","Details","Recovery"]];
  const allIssues = [...(analysis?.blockers || []).map((issue) => ({ ...issue, severity: "blocker" })), ...(analysis?.warnings || []).map((issue) => ({ ...issue, severity: "warning" }))];
  allIssues.forEach((issue) => issueRows.push([
    issue.severity, issue.code || "", issue.participant || issue.email || issue.identifier || issue.source || "—", issue.source || "—", String(issue.rowNumber ?? ""),
    JSON.stringify({ email: issue.email, identifier: issue.identifier, identifiers: issue.identifiers, emails: issue.emails, evaluator: issue.evaluator, names: issue.names, failureStage: issue.failureStage, errorCode: issue.errorCode, errorDetail: issue.errorDetail }),
    issue.correctiveAction || "Correct the named source, remove or replace it, then run review again.",
  ]));
  const identifierRows = [["Evaluator email","Identifier","Disposition","Canonical name","Evidence sources"]];
  Object.keys(analysis?.identifiers || {}).sort().forEach((email) => {
    const evidence = analysis.identifierEvidence?.[email] || {};
    identifierRows.push([email, analysis.identifiers[email], evidence.disposition || "new", analysis.canonicalNames?.[email] || "", (evidence.records || []).map((record) => `${record.source}:${record.rowNumber || "—"}`).join("; ")]);
  });
  const normalizationRows = [["Source","Row","Field","Original literal value","Normalized literal value"]];
  normalizations.forEach((item) => normalizationRows.push([item.source || "—", String(item.rowNumber ?? ""), item.field || "", item.original ?? "", item.normalized ?? ""]));
  const workbook = XLSX.utils.book_new();
  workbook.Props = { Title: `360 audit - ${normalizeText(projectName) || "project"}`, Author: "Business Health Bar", Comments: "Generated locally; client-confidential; structurally validated only." };
  XLSX.utils.book_append_sheet(workbook, literalSheet(XLSX, runRows, [28, 95]), "Run");
  XLSX.utils.book_append_sheet(workbook, literalSheet(XLSX, sourceRows, [30,12,68,16,16,14,16,14,22,16,12,12]), "Sources");
  XLSX.utils.book_append_sheet(workbook, literalSheet(XLSX, issueRows, [12,28,35,30,10,70,70]), "Issues");
  XLSX.utils.book_append_sheet(workbook, literalSheet(XLSX, identifierRows, [38,14,16,32,55]), "Identifiers");
  XLSX.utils.book_append_sheet(workbook, literalSheet(XLSX, normalizationRows, [30,10,24,55,55]), "Normalizations");
  workbook.SheetNames = Array.from(workbook.SheetNames);
  return workbook;
}

export async function sha256Hex(data) {
  const bytes = data instanceof ArrayBuffer ? new Uint8Array(data) : new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  const digest = await globalThis.crypto.subtle.digest("SHA-256", bytes);
  return [...new Uint8Array(digest)].map((value) => value.toString(16).padStart(2, "0")).join("");
}

export function safeFilePart(value) { return normalizeText(value).normalize("NFD").replace(/[\u0300-\u036f]/gu, "").replace(/[^a-z0-9_-]+/giu,"-").replace(/^-+|-+$/gu,"").slice(0,60) || "proiect-360"; }
