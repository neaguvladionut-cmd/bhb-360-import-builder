const PRODUCTION_HEADERS = ["Marca","PersoanaEvaluata","PersoanaEvaluataEmail","Evaluator","EvaluatorEmail","EvaluatorRol","NumeCampanie","Limba","Criteriu1","Criteriu2","Criteriu3","Criteriu4","Criteriu5","NumeInregistrare"];
const COLLECTOR_HEADERS = ["ParticipantName","ParticipantEmail","EvaluatorName","EvaluatorEmail","Relationship","Language","IsSelf"];
const COLLECTOR_HEADERS_WITH_CRITERIA = [...COLLECTOR_HEADERS,"Criteriu1","Criteriu2","Criteriu3","Criteriu4","Criteriu5"];
const MASS_TEMPLATE_HEADERS = ["ParticipantName","ParticipantEmail","EvaluatorName","EvaluatorEmail","Relationship","Language","Criteriu1","Criteriu2","Criteriu3","Criteriu4","Criteriu5"];
const RECOMMENDED_RESPONDENTS = 10;
const WARNING_RESPONDENTS = 15;
const CRITERIA_COUNT = 5;
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/u;
const LANGUAGE_RE = /^[A-Z]{2}$/u;
const ROLE_ORDER = { Manager: 0, Peer: 1, Subordonat: 2, PartenerExtern: 3 };

function normalizeText(value) {
  return String(value ?? "").replace(/[\u00a0\u2007\u202f]/gu, " ").replace(/\s+/gu, " ").trim();
}
function normalizeEmail(value) { return normalizeText(value).toLocaleLowerCase("en-US"); }
function normalizeLanguage(value) { return normalizeText(value).toUpperCase(); }
function isValidEmail(value) { return EMAIL_RE.test(normalizeEmail(value)); }
function fold(value) { return normalizeText(value).normalize("NFD").replace(/[\u0300-\u036f]/gu, "").toLocaleLowerCase("en-US"); }
function roleKey(value) { return fold(value).replace(/[\s_-]+/gu, ""); }
function isSelfRole(value) { return roleKey(value) === "autoevaluare"; }

function mapRole(value) {
  const key = roleKey(value);
  if (key === "autoevaluare") return "Manager";
  if (key === "manager" || key === "functionalmanager") return "Manager";
  if (key === "peer" || key === "coleg") return "Peer";
  if (key === "subordonat") return "Subordonat";
  if (key === "stakeholder" || key === "partener" || key === "partenerextern") return "PartenerExtern";
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

function parseCollectorRows(rows, headerIndex, sourceName, normalizations, includeCriteria = false) {
  const sourceRows = rows.slice(headerIndex + 1);
  const contentRows = sourceRows.map((row, index) => ({ row, rowNumber: headerIndex + index + 2 })).filter(({ row }) => row.some((value) => normalizeText(value)));
  const objects = contentRows.map(({ row, rowNumber }) => {
    const sourceRole = normalizeText(row[4]);
    const participantEmail = normalizeEmail(row[1]);
    const evaluatorEmail = normalizeEmail(row[3]);
    const criteria = Array.from({ length: CRITERIA_COUNT }, (_, criterionIndex) => normalizeText(row[7 + criterionIndex]));
    const values = [
      ["ParticipantName", row[0], normalizeText(row[0])], ["ParticipantEmail", row[1], participantEmail],
      ["EvaluatorName", row[2], normalizeText(row[2])], ["EvaluatorEmail", row[3], evaluatorEmail],
      ["Relationship", row[4], sourceRole], ["Language", row[5], normalizeLanguage(row[5])],
      ["ProductionRole", sourceRole, mapRole(sourceRole) || sourceRole],
      ...criteria.map((value, criterionIndex) => [`Criteriu${criterionIndex + 1}`, row[7 + criterionIndex], value]),
    ];
    values.forEach(([field, original, normalized]) => recordNormalization(normalizations, sourceName, rowNumber, field, original, normalized));
    return {
      participantName: normalizeText(row[0]), participantEmail,
      evaluatorName: normalizeText(row[2]), evaluatorEmail,
      sourceRole, role: mapRole(sourceRole), language: normalizeLanguage(row[5]), criteria,
      isSelf: fold(row[6]) === "da" || fold(row[6]) === "yes" || isSelfRole(sourceRole) || (participantEmail && participantEmail === evaluatorEmail),
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

function parseMassRows(rows, headerIndex, sourceName, normalizations) {
  const sourceRows = rows.slice(headerIndex + 1);
  const contentRows = sourceRows.map((row, index) => ({ row, rowNumber: headerIndex + index + 2 })).filter(({ row }) => row.some((value) => normalizeText(value)));
  const grouped = new Map();
  contentRows.forEach(({ row, rowNumber }) => {
    const sourceRole = normalizeText(row[4]);
    const participantEmail = normalizeEmail(row[1]);
    const evaluatorEmail = normalizeEmail(row[3]);
    const participantName = normalizeText(row[0]);
    const evaluatorName = normalizeText(row[2]);
    const criteria = Array.from({ length: CRITERIA_COUNT }, (_, criterionIndex) => normalizeText(row[6 + criterionIndex]));
    const values = [
      ["ParticipantName", row[0], participantName], ["ParticipantEmail", row[1], participantEmail],
      ["EvaluatorName", row[2], evaluatorName], ["EvaluatorEmail", row[3], evaluatorEmail],
      ["Relationship", row[4], sourceRole], ["Language", row[5], normalizeLanguage(row[5])],
      ["ProductionRole", sourceRole, mapRole(sourceRole) || sourceRole],
      ...criteria.map((value, criterionIndex) => [`Criteriu${criterionIndex + 1}`, row[6 + criterionIndex], value]),
    ];
    values.forEach(([field, original, normalized]) => recordNormalization(normalizations, sourceName, rowNumber, field, original, normalized));
    const rowValue = {
      participantName, participantEmail, evaluatorName, evaluatorEmail, sourceRole,
      role: mapRole(sourceRole), language: normalizeLanguage(row[5] || "RO"), criteria,
      isSelf: participantEmail && participantEmail === evaluatorEmail,
      source: sourceName, rowNumber, legacy: false, mass: true,
    };
    const key = participantEmail || `missing:${participantName}`;
    if (!grouped.has(key)) grouped.set(key, { id: `${sourceName}:${key}`, source: sourceName, adapter: "mass", participantName, participantEmail, rows: [] });
    grouped.get(key).rows.push(rowValue);
  });
  const participants = [...grouped.values()];
  participants.forEach((participant) => {
    if (participant.rows.some((row) => row.isSelf)) return;
    const first = participant.rows[0] || {};
    participant.rows.unshift({
      participantName: participant.participantName, participantEmail: participant.participantEmail,
      evaluatorName: participant.participantName, evaluatorEmail: participant.participantEmail,
      sourceRole: "Autoevaluare", role: "Manager", language: first.language || "RO", criteria: first.criteria || [],
      isSelf: true, source: sourceName, rowNumber: null, legacy: false, mass: true, synthetic: true,
    });
  });
  return { participants, ignoredRows: sourceRows.length - contentRows.length };
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
    const evaluatorName = normalizeText([row[lastNameIndex], row[firstNameIndex]].filter(Boolean).join(" "));
    const evaluatorEmail = normalizeEmail(row[emailIndex]);
    recordNormalization(normalizations, sourceName, rowNumber, "EvaluatorName", [row[lastNameIndex], row[firstNameIndex]].filter(Boolean).join(" "), evaluatorName);
    recordNormalization(normalizations, sourceName, rowNumber, "EvaluatorEmail", row[emailIndex], evaluatorEmail);
    recordNormalization(normalizations, sourceName, rowNumber, "Relationship", row[roleIndex], sourceRole);
    recordNormalization(normalizations, sourceName, rowNumber, "ProductionRole", sourceRole, mapRole(sourceRole) || sourceRole);
    return {
      evaluatorName, evaluatorEmail, sourceRole, role: mapRole(sourceRole), language: "RO",
      isSelf: isSelfRole(sourceRole), source: sourceName, rowNumber, legacy: true,
    };
  }).filter(Boolean);
  const self = parsedRows.find((row) => row.isSelf);
  const participantName = self?.evaluatorName || "";
  const participantEmail = self?.evaluatorEmail || "";
  parsedRows.forEach((row) => { row.participantName = participantName; row.participantEmail = participantEmail; });
  return { participants: [{ id: `${sourceName}:legacy`, source: sourceName, adapter: "legacy", participantName, participantEmail, rows: parsedRows, legacyLanguageAssumed: true }], ignoredRows: sourceRows.length - parsedRows.length };
}

function parseParticipantWorkbook(XLSX, data, sourceName) {
  const workbook = XLSX.read(data, { type: data instanceof ArrayBuffer ? "array" : "buffer", cellFormula: false, cellHTML: false });
  for (const sheetName of workbook.SheetNames) {
    const rows = rowsFromSheet(XLSX, workbook.Sheets[sheetName]);
    const massHeader = exactHeaderIndex(rows, MASS_TEMPLATE_HEADERS);
    if (massHeader >= 0) {
      const normalizations = [];
      const parsed = parseMassRows(rows, massHeader, sourceName, normalizations);
      return { type: "mass", participants: parsed.participants, warnings: [], ignoredRows: parsed.ignoredRows, normalizations };
    }
    const collectorHeaderWithCriteria = exactHeaderIndex(rows, COLLECTOR_HEADERS_WITH_CRITERIA);
    const collectorHeader = collectorHeaderWithCriteria >= 0 ? collectorHeaderWithCriteria : exactHeaderIndex(rows, COLLECTOR_HEADERS);
    if (collectorHeader >= 0) {
      const normalizations = [];
      const parsed = parseCollectorRows(rows, collectorHeader, sourceName, normalizations, collectorHeaderWithCriteria >= 0);
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

function parseExistingProjectWorkbook(XLSX, workbook, sourceName) {
  for (const sheetName of workbook.SheetNames) {
    const rows = rowsFromSheet(XLSX, workbook.Sheets[sheetName]);
    const headerIndex = exactHeaderIndex(rows, PRODUCTION_HEADERS);
    if (headerIndex < 0) continue;
    const sourceRows = rows.slice(headerIndex + 1);
    const contentRows = sourceRows.map((row, index) => ({ row, rowNumber: headerIndex + index + 2 })).filter(({ row }) => row.some((value) => normalizeText(value)));
    const normalizations = [];
    const records = contentRows.map(({ row, rowNumber }) => {
      const record = { identifier: normalizeText(row[0]), name: normalizeText(row[3]), email: normalizeEmail(row[4]), source: sourceName, rowNumber };
      recordNormalization(normalizations, sourceName, rowNumber, "Identifier", row[0], record.identifier);
      recordNormalization(normalizations, sourceName, rowNumber, "Assessor Name", row[3], record.name);
      recordNormalization(normalizations, sourceName, rowNumber, "Assessor Email", row[4], record.email);
      return record;
    });
    const projectNames = [...new Set(contentRows.map(({ row }) => normalizeText(row[6])).filter(Boolean))];
    return { ...analyzeAllocation(records), type: "allocation", origin: "existing-production", projectName: projectNames.length === 1 ? projectNames[0] : "", projectNames, ignoredRows: sourceRows.length - contentRows.length, normalizations };
  }
  return null;
}

function parseSourceWorkbook(XLSX, data, sourceName) {
  const workbook = XLSX.read(data, { type: data instanceof ArrayBuffer ? "array" : "buffer", cellFormula: false, cellHTML: false });
  const existing = parseExistingProjectWorkbook(XLSX, workbook, sourceName);
  if (existing) return existing;
  for (const sheetName of workbook.SheetNames) {
    const rows = rowsFromSheet(XLSX, workbook.Sheets[sheetName]);
    if (exactHeaderIndex(rows, ["Identifier","Assessor Name","Assessor Email","QuestionaireId","Assessed Email","Assessor Role"]) >= 0) return { ...parseAllocationWorkbook(XLSX, data, sourceName), type: "allocation", origin: "allocation-export" };
  }
  return parseParticipantWorkbook(XLSX, data, sourceName);
}

function parseAllocationWorkbook(XLSX, data, sourceName = "allocation.xlsx") {
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

function analyzeAllocation(records = []) {
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
        language: normalizeLanguage(row.language || "RO"),
        criteria: Array.from({ length: CRITERIA_COUNT }, (_, criterionIndex) => normalizeText(row.criteria?.[criterionIndex])),
        isSelf,
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

function analyzeProject({ participants = [], allocation = analyzeAllocation([]), nameChoices = {}, projectName = "", sourceBlockers = [] }) {
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
    if (nonSelf.length > RECOMMENDED_RESPONDENTS) warnings.push(issue("respondent-recommendation", participant, null, { count: nonSelf.length }));
    if (nonSelf.length > WARNING_RESPONDENTS) warnings.push(issue("respondent-cohort-warning", participant, null, { count: nonSelf.length }));
    const criteriaByIndex = Array.from({ length: CRITERIA_COUNT }, (_, criterionIndex) => [...new Set(participant.rows.map((row) => normalizeText(row.criteria?.[criterionIndex])).filter(Boolean))]);
    criteriaByIndex.forEach((values, criterionIndex) => {
      if (values.length > 1) local.push(issue("criteria-conflict", participant, null, { criterion: criterionIndex + 1, values }));
    });
    const participantCriteria = criteriaByIndex.map((values) => values[0] || "");
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
    if (!local.length) validRows.push(...participant.rows.map((row) => ({ ...row, criteria: participantCriteria })));
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
function createProductionWorkbook(XLSX, analysis, projectName) {
  if (!analysis.ready) throw new Error("Project has blocking errors");
  const sheet = {};
  PRODUCTION_HEADERS.forEach((header, column) => { sheet[XLSX.utils.encode_cell({ r: 0, c: column })] = textCell(header); });
  analysis.outputRows.forEach((row, index) => {
    const excelRow = index + 2;
    const criteria = Array.from({ length: CRITERIA_COUNT }, (_, criterionIndex) => row.criteria?.[criterionIndex] || "");
    const values = [row.identifier,row.participantName,row.participantEmail,row.evaluatorName,row.evaluatorEmail,row.role,normalizeText(projectName),row.language,...criteria];
    values.forEach((value, column) => { sheet[XLSX.utils.encode_cell({ r: index + 1, c: column })] = textCell(value); });
    sheet[XLSX.utils.encode_cell({ r: index + 1, c: 13 })] = { t: "str", f: `G${excelRow}&" - "&B${excelRow}&" - "&D${excelRow}`, v: "" };
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

function createAuditWorkbook(XLSX, { projectName = "", analysis, sources = [], generatedAt = "", normalizations = [] }) {
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

async function sha256Hex(data) {
  const bytes = data instanceof ArrayBuffer ? new Uint8Array(data) : new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  const digest = await globalThis.crypto.subtle.digest("SHA-256", bytes);
  return [...new Uint8Array(digest)].map((value) => value.toString(16).padStart(2, "0")).join("");
}

function safeFilePart(value) { return normalizeText(value).normalize("NFD").replace(/[\u0300-\u036f]/gu, "").replace(/[^a-z0-9_-]+/giu,"-").replace(/^-+|-+$/gu,"").slice(0,60) || "proiect-360"; }

const text = {
  ro: {
    skip: "Sari la spațiul de lucru", eyebrow: "Instrument intern · 360", title: "Din liste primite, într-un import curat.", lead: "Adună fișierele participanților, rezolvă excepțiile și descarcă un singur workbook verificat structural pentru aplicația 360 actuală.", localTitle: "Procesare exclusiv locală", localText: "Fișierele și datele personale nu părăsesc acest browser.", step1: "Încarcă listele", step2: "Corectează și confirmă", step3: "Descarcă importul", workbench: "Spațiu de lucru consultant", setup: "Pregătește proiectul", project: "Numele proiectului / campaniei", files: "Încarcă workbook-urile proiectului", filesHelp: "Selectează unul sau mai multe fișiere .xlsx: collector, șablon relații, legacy, export alocări sau import existent. Formatul este detectat automat.", massTemplate: "Șablon relații în masă", massTemplateHelp: "Opțional. Încarcă aici șablonul completat de HR sau de client.", allocation: "Export alocări existent", allocationHelp: "Opțional, doar pentru un proiect deja activ", privacyLine: "Nicio încărcare online, fără analize de utilizare și fără salvare în browser. Resetarea șterge datele din memorie; fișierele deja descărcate rămân pe dispozitiv.", reviewTitle: "Revizuire și corecturi", reviewLead: "Erorile sunt izolate pentru fiecare participant și sursă. Fișierul final devine disponibil numai după rezolvarea tuturor blocajelor.", waiting: "În așteptare", ready: "Gata pentru export", blocked: "Necesită corecturi", reset: "Șterge sesiunea", download: "Descarcă importul A:N", audit: "Descarcă raportul de trasabilitate", auditPrivacy: "Raportul de trasabilitate poate conține informații confidențiale despre client. Tu controlezi unde este păstrat și când este șters.", boundary: "Rezultatul este verificat structural. Compatibilitatea de producție se confirmă numai printr-un import controlat în aplicația actuală.", footer: "Instrument intern pentru pregătirea importului 360.", participants: "Participanți", allocations: "Rânduri valide", reused: "Identificatori reutilizați", newIds: "Identificatori noi", roles: "Roluri", languages: "Limbi", rowsLabel: "rânduri", rowLabel: "Rând", legacyNotice: "Fișierele legacy nu conțin limba chestionarului. RO este precompletat; verifică fiecare rând înainte de export.", englishNotice: "Chestionare EN — valide, verifică fiecare alocare înainte de lansare.", respondentGuidance: "Recomandare privind dimensiunea cohortelor", cohortWarning: "Cohorta depășește 15 respondenți. Verifică dacă toți sunt necesari; importul rămâne permis.", name: "Numele participantului", email: "Adresa de email a participantului", language: "Limbă", source: "Sursă", confirmName: "Confirmă numele ales", conflict: "Aceeași adresă de email apare cu nume diferite", recommended: "Recomandat pe baza exportului", correctAllocation: "Corectează sau înlocuiește exportul de alocări înainte de a continua.", clearConfirm: "Ștergi toate fișierele și corecturile din sesiune?", unsupported: "Fișierul nu are un format recunoscut", allocationError: "Exportul de alocări nu poate fi citit", projectAutoFilled: "Numele proiectului a fost preluat automat din fișierul existent.", projectPlaceholder: "ex. Proiect 360 2026", noData: "Încarcă fișiere pentru a începe revizuirea.", sourcesTitle: "Surse și trasabilitate", sourcesLead: "Fiecare încercare rămâne vizibilă, cu amprenta SHA-256 și rezultatul prelucrării.", removeSource: "Elimină", replaceSource: "Înlocuiește", sourceAccepted: "acceptată", sourceBlocked: "blocată", sourceRejected: "respinsă", sourcePending: "se verifică", adapter: "format", parsed: "citite", exportable: "exportabile", ignored: "ignorate intenționat", normalizations: "normalizări", recoveryTitle: "Cum revii după o eroare", recoveryText: "Corectează fișierul indicat, apoi folosește opțiunea Înlocuiește pentru sursa lui sau elimină-l și încarcă-l din nou. Revizuiește și regenerează ieșirile. Dacă închizi sau resetezi pagina, memoria se golește și sursele trebuie reîncărcate.", sourceRejectedIssue: "Sursa nu a putut fi citită. Corectează fișierul, apoi elimină sursa sau înlocuiește-o; nu reutilizăm în tăcere o stare mai veche.", sourcePendingIssue: "Verificarea locală este în curs; exportul de producție rămâne blocat.", criteriaConflict: "Criteriile diferă pentru același participant. Păstrează un singur set de criterii în sursă înainte de export.",
  },
  en: {
    skip: "Skip to workspace", eyebrow: "Internal tool · 360", title: "From received lists to a clean import.", lead: "Combine participant files, resolve exceptions and download one structurally validated workbook for the current 360 application.", localTitle: "Local processing only", localText: "Files and personal data never leave this browser.", step1: "Load lists", step2: "Correct and confirm", step3: "Download import", workbench: "Consultant workspace", setup: "Prepare the project", project: "Project / campaign name", files: "Load the project workbooks", filesHelp: "Select one or more .xlsx files: collector, relations template, legacy, allocation export or existing import. The format is detected automatically.", massTemplate: "Mass relations template", massTemplateHelp: "Optional. Load the template completed by HR or the client here.", allocation: "Existing allocation export", allocationHelp: "Optional, only for an already active project", privacyLine: "No online uploads, analytics, or browser storage. Reset clears in-memory data; downloaded files remain on the device.", reviewTitle: "Review and corrections", reviewLead: "Errors are isolated by participant and source. The final file is available only after every blocker is resolved.", waiting: "Waiting", ready: "Ready to export", blocked: "Corrections needed", reset: "Clear session", download: "Download A:N import", audit: "Download traceability report", auditPrivacy: "The traceability report may contain client-confidential information. You control where it is stored and when it is deleted.", boundary: "The result is structurally validated. Production compatibility requires a controlled import into the current application.", footer: "Internal tool for preparing the 360 import.", participants: "Participants", allocations: "Valid rows", reused: "Reused identifiers", newIds: "New identifiers", roles: "Roles", languages: "Languages", rowsLabel: "rows", rowLabel: "Row", legacyNotice: "Legacy files contain no questionnaire language. RO is prefilled; check every row before export.", englishNotice: "EN questionnaires — valid; review every allocation before launch.", respondentGuidance: "Cohort size guidance", cohortWarning: "This cohort exceeds 15 respondents. Check that everyone is necessary; import remains allowed.", name: "Participant's name", email: "Participant's email address", language: "Language", source: "Source", confirmName: "Confirm selected name", conflict: "The same email address appears with different names", recommended: "Recommended based on the allocation export", correctAllocation: "Correct or replace the allocation export before continuing.", clearConfirm: "Clear every file and correction in this session?", unsupported: "The workbook format is not recognised", allocationError: "The allocation export cannot be read", projectAutoFilled: "The project name was taken automatically from the existing workbook.", projectPlaceholder: "e.g. 360 Project 2026", noData: "Load files to begin review.", sourcesTitle: "Sources and traceability", sourcesLead: "Every attempt stays visible with its SHA-256 fingerprint and processing outcome.", removeSource: "Remove", replaceSource: "Replace", sourceAccepted: "accepted", sourceBlocked: "blocked", sourceRejected: "rejected", sourcePending: "checking", adapter: "format", parsed: "parsed", exportable: "exportable", ignored: "intentionally ignored", normalizations: "normalizations", recoveryTitle: "How to recover from an error", recoveryText: "Correct the named file, then use Replace on that source, or remove and load it again. Review and regenerate the outputs. Closing or resetting clears memory, so sources must be reloaded.", sourceRejectedIssue: "The source could not be read. Correct the file, then remove or replace the source; an older state is never reused silently.", sourcePendingIssue: "Local checking is in progress; production export remains blocked.", criteriaConflict: "Criteria differ for the same participant. Keep one criteria set in the source before export.",
  },
};

let lang = "ro";
let sourceLedger = [];
let nameChoices = {};
let lastAnalysis = null;
let lastSourceSnapshots = [];
let sourceSequence = 1;
const q = (selector) => document.querySelector(selector);
const tr = (key) => text[lang][key] || key;
const readFile = (file) => file.arrayBuffer();
function el(tag, className, textContent) { const node = document.createElement(tag); if (className) node.className = className; if (textContent !== undefined) node.textContent = textContent; return node; }

function issueMessage(issue) {
  if (issue.code === "source-pending") return tr("sourcePendingIssue");
  const ro = { "project-required": "Completează numele proiectului.", "files-required": "Încarcă cel puțin un fișier de respondenți.", "source-rejected": tr("sourceRejectedIssue"), "participant-name-missing": "Completează numele participantului.", "participant-email-invalid": "Completează un email valid pentru participant.", "self-count": "Este necesar un singur rând de autoevaluare.", "evaluator-name-missing": "Lipsește numele unui respondent.", "evaluator-email-invalid": "Email de respondent invalid.", "unknown-role": "Rol necunoscut; corectează fișierul sursă.", "language-invalid": "Limba trebuie să aibă două litere.", "duplicate-allocation": "Același respondent apare de două ori pentru participant.", "manager-required": "Lipsește un Manager distinct de participant.", "name-conflict": "Confirmă numele canonic pentru email.", "criteria-conflict": tr("criteriaConflict"), "allocation-invalid": "Rând invalid în exportul de alocări.", "email-multiple-identifiers": "Un email are mai mulți identificatori.", "identifier-multiple-emails": "Un identificator aparține mai multor emailuri." };
  const en = { "project-required": "Enter the project name.", "files-required": "Load at least one respondent file.", "source-rejected": tr("sourceRejectedIssue"), "participant-name-missing": "Enter the participant name.", "participant-email-invalid": "Enter a valid participant email.", "self-count": "Exactly one self-evaluation row is required.", "evaluator-name-missing": "A respondent name is missing.", "evaluator-email-invalid": "A respondent email is invalid.", "unknown-role": "Unknown role; correct the source workbook.", "language-invalid": "Language must use two letters.", "duplicate-allocation": "The same respondent appears twice for this participant.", "manager-required": "A Manager distinct from the participant is missing.", "name-conflict": "Confirm the canonical name for this email.", "criteria-conflict": tr("criteriaConflict"), "allocation-invalid": "Invalid row in the allocation export.", "email-multiple-identifiers": "One email has multiple identifiers.", "identifier-multiple-emails": "One identifier belongs to multiple emails." };
  return (lang === "ro" ? ro : en)[issue.code] || issue.code;
}
function allocationIssueTitle(issue) { if (issue.code === "email-multiple-identifiers") return `${issue.email}: ${lang === "ro" ? "identificatori" : "identifiers"} ${issue.identifiers.join(", ")}`; if (issue.code === "identifier-multiple-emails") return `${lang === "ro" ? "Identificator" : "Identifier"} ${issue.identifier}: ${issue.emails.join(", ")}`; return `${issue.source || "export"} · ${tr("rowLabel")} ${issue.rowNumber || "—"}: ${issueMessage(issue)}`; }
function allocationIssueDetails(issue) { return (issue.records || []).map((record) => `${record.source || "export"} · ${tr("rowLabel")} ${record.rowNumber || "—"}: ${record.email || "—"} ↔ ${record.identifier || "—"}`); }

function setLang(next) { lang = next; document.documentElement.lang = next; document.querySelectorAll("[data-language]").forEach((button) => { const active = button.dataset.language === next; button.classList.toggle("active", active); button.setAttribute("aria-pressed", String(active)); }); document.querySelectorAll("[data-i18n]").forEach((node) => { node.textContent = tr(node.dataset.i18n); }); document.querySelectorAll("[data-i18n-placeholder]").forEach((node) => { node.placeholder = tr(node.dataset.i18nPlaceholder); }); render(); }
function removeSource(id) { sourceLedger = sourceLedger.filter((source) => source.id !== id); nameChoices = {}; render(); }
function beginSourceAttempt(file, replaceId = null) { const attemptToken = `attempt-${sourceSequence++}`; const index = replaceId ? sourceLedger.findIndex((source) => source.id === replaceId) : -1; const pending = { id: index >= 0 ? replaceId : `source-${sourceSequence++}`, attemptToken, attemptState: "pending", kind: "source", name: file.name, bytes: file.size, sha256: "unavailable (pending)", adapter: "pending" }; if (index >= 0) sourceLedger.splice(index, 1, pending); else sourceLedger.push(pending); nameChoices = {}; render(); return pending; }
function settleSourceAttempt(pending, next) { const index = sourceLedger.findIndex((source) => source.id === pending.id && source.attemptToken === pending.attemptToken); if (index >= 0) sourceLedger.splice(index, 1, { ...next, id: pending.id, attemptToken: pending.attemptToken }); }
async function attemptSource(file, replaceId = null) { q("#global-message").hidden = true; const pending = beginSourceAttempt(file, replaceId); let stage = "read"; let data; let sha256 = "unavailable"; try { data = await readFile(file); stage = "digest"; sha256 = await sha256Hex(data); stage = "parse"; const parsed = parseSourceWorkbook(window.XLSX, data, file.name); const sourceKind = parsed.type === "allocation" ? "allocation" : "participant"; if (sourceKind === "allocation") sourceLedger = sourceLedger.filter((source) => source.id === pending.id || source.kind !== "allocation"); settleSourceAttempt(pending, { kind: sourceKind, name: file.name, bytes: file.size, sha256, adapter: parsed.type, parsed }); if (parsed.projectName && !q("#project-name").value.trim()) { q("#project-name").value = parsed.projectName; q("#project-auto-note").hidden = false; } } catch (error) { const failureStage = stage; const errorCode = failureStage === "read" ? "source-read-failed" : failureStage === "digest" ? "source-digest-failed" : "source-parse-failed"; settleSourceAttempt(pending, { kind: "source", name: file.name, bytes: file.size, sha256, adapter: "rejected", attemptState: "rejected", failureStage, errorCode, errorDetail: error?.message || "source-error" }); q("#global-message").textContent = `${file.name}: ${tr("unsupported")}. ${tr("sourceRejectedIssue")}`; q("#global-message").hidden = false; } nameChoices = {}; render(); }

q("#source-files").addEventListener("change", async (event) => { for (const file of event.target.files) await attemptSource(file); event.target.value = ""; });
q("#project-name").addEventListener("input", () => { q("#project-auto-note").hidden = true; render(); });
q("#reset").addEventListener("click", () => { if (!window.confirm(tr("clearConfirm"))) return; sourceLedger = []; nameChoices = {}; q("#project-name").value = ""; q("#project-auto-note").hidden = true; q("#global-message").hidden = true; render(); });
q("#download").addEventListener("click", () => { if (!lastAnalysis?.ready) return; const name = q("#project-name").value; const workbook = createProductionWorkbook(window.XLSX, lastAnalysis, name); window.XLSX.writeFile(workbook, `import-360-${safeFilePart(name)}.xlsx`, { compression: true, bookType: "xlsx" }); });
q("#download-audit").addEventListener("click", () => { if (!sourceLedger.length) return; const name = q("#project-name").value; const normalizations = sourceLedger.flatMap((source) => source.parsed?.normalizations || []); const workbook = createAuditWorkbook(window.XLSX, { projectName: name, analysis: lastAnalysis, sources: lastSourceSnapshots, normalizations, generatedAt: new Date().toLocaleString() }); window.XLSX.writeFile(workbook, `audit-360-${safeFilePart(name)}.xlsx`, { compression: true, bookType: "xlsx" }); });
document.querySelectorAll("[data-language]").forEach((button) => button.addEventListener("click", () => setLang(button.dataset.language)));

function currentInputs() { const participantSources = sourceLedger.filter((source) => source.kind === "participant" && ["collector", "legacy", "mass"].includes(source.adapter)); const allocationSource = sourceLedger.find((source) => source.kind === "allocation" && source.adapter === "allocation"); return { participants: participantSources.flatMap((source) => source.parsed.participants), allocation: allocationSource?.parsed || analyzeAllocation([]), sourceBlockers: sourceLedger.filter((source) => ["pending", "rejected"].includes(source.adapter)).map((source) => source.adapter === "pending" ? { code: "source-pending", source: source.name, sourceId: source.id, correctiveAction: tr("sourcePendingIssue") } : { code: "source-rejected", source: source.name, sourceId: source.id, failureStage: source.failureStage, errorCode: source.errorCode, errorDetail: source.errorDetail, correctiveAction: tr("sourceRejectedIssue") }) }; }
function sourceSnapshot(source, analysis) { if (source.adapter === "pending") return { name: source.name, bytes: source.bytes, sha256: source.sha256 || "unavailable (pending)", adapter: "pending", disposition: "pending", parsedRows: 0, exportableRows: 0, blockedRows: 0, ignoredRows: 0, normalizationCount: 0, warningCount: 0, errorCount: 0 }; if (source.adapter === "rejected") return { name: source.name, bytes: source.bytes, sha256: source.sha256 || "unavailable", adapter: "rejected", disposition: "blocked", parsedRows: 0, exportableRows: 0, blockedRows: 0, ignoredRows: 0, normalizationCount: 0, warningCount: 0, errorCount: 1 }; if (source.kind === "allocation") { const errorRecords = new Set((source.parsed.errors || []).flatMap((issue) => issue.records?.map((record) => `${record.source}:${record.rowNumber}`) || [`${issue.source}:${issue.rowNumber}`])); return { name: source.name, bytes: source.bytes, sha256: source.sha256, adapter: "allocation", disposition: source.parsed.errors.length ? "blocked" : "accepted", parsedRows: source.parsed.records.length, exportableRows: source.parsed.records.length - errorRecords.size, blockedRows: errorRecords.size, ignoredRows: source.parsed.ignoredRows || 0, normalizationCount: source.parsed.normalizations?.length || 0, warningCount: 0, errorCount: source.parsed.errors.length }; } const parsedRows = source.parsed.participants.reduce((sum, participant) => sum + participant.rows.length, 0); const blockedRows = source.parsed.participants.reduce((sum, participant) => { const summary = analysis.participantSummaries.find((item) => item.participant.participantEmail === participant.participantEmail || item.participant.id === participant.id); return sum + (summary?.blockers.length ? participant.rows.length : 0); }, 0); return { name: source.name, bytes: source.bytes, sha256: source.sha256, adapter: source.adapter, disposition: blockedRows ? "blocked" : "accepted", parsedRows, exportableRows: parsedRows - blockedRows, blockedRows, ignoredRows: source.parsed.ignoredRows || 0, normalizationCount: source.parsed.normalizations?.length || 0, warningCount: source.parsed.warnings?.length || 0, errorCount: blockedRows ? 1 : 0 }; }

function render() { const projectName = q("#project-name").value; const inputs = currentInputs(); lastAnalysis = analyzeProject({ ...inputs, nameChoices, projectName }); lastSourceSnapshots = sourceLedger.map((source) => sourceSnapshot(source, lastAnalysis)); renderSources(); const pill = q("#readiness-pill"); pill.className = `pill ${lastAnalysis.ready ? "ready" : inputs.participants.length || sourceLedger.length ? "blocked" : ""}`; pill.textContent = tr(lastAnalysis.ready ? "ready" : inputs.participants.length || sourceLedger.length ? "blocked" : "waiting"); q("#download").disabled = !lastAnalysis.ready; q("#download-audit").disabled = !sourceLedger.length; renderSummary(); renderNotices(); renderParticipants(inputs.participants); renderConflicts(); }
function renderSources() { const host = q("#source-ledger"); host.replaceChildren(); lastSourceSnapshots.forEach((snapshot, index) => { const source = sourceLedger[index]; const card = el("article", `source-card ${snapshot.disposition}`); const top = el("div", "source-card-head"); const title = el("div", ""); title.append(el("strong", "", snapshot.name), el("span", "source-state", tr(snapshot.adapter === "pending" ? "sourcePending" : snapshot.adapter === "rejected" ? "sourceRejected" : snapshot.disposition === "blocked" ? "sourceBlocked" : "sourceAccepted"))); const controls = el("div", "source-controls"); const replace = el("label", "source-replace", tr("replaceSource")); const replaceInput = document.createElement("input"); replaceInput.type = "file"; replaceInput.accept = ".xlsx"; replaceInput.className = "source-replace-input"; replaceInput.addEventListener("change", async () => { const file = replaceInput.files[0]; if (file) await attemptSource(file, source.id); }); replace.append(replaceInput); const remove = el("button", "source-remove", tr("removeSource")); remove.type = "button"; remove.addEventListener("click", () => removeSource(source.id)); controls.append(replace, remove); top.append(title, controls); card.append(top, el("code", "fingerprint", snapshot.sha256)); card.append(el("p", "source-stats", `${snapshot.bytes} bytes · ${tr("adapter")}: ${snapshot.adapter} · ${tr("parsed")}: ${snapshot.parsedRows} · ${tr("exportable")}: ${snapshot.exportableRows} · ${tr("ignored")}: ${snapshot.ignoredRows} · ${tr("normalizations")}: ${snapshot.normalizationCount}`)); if (snapshot.disposition !== "accepted") card.append(el("p", "source-recovery", tr(snapshot.adapter === "pending" ? "sourcePendingIssue" : "sourceRejectedIssue"))); host.append(card); }); }
function renderSummary() { const items = [[lastAnalysis.summary.participants, tr("participants")], [lastAnalysis.summary.rows, tr("allocations")], [lastAnalysis.summary.reusedCount, tr("reused")], [lastAnalysis.summary.newCount, tr("newIds")]]; Object.entries(lastAnalysis.summary.byRole).sort().forEach(([role, count]) => items.push([count, `${tr("roles")} · ${role}`])); Object.entries(lastAnalysis.summary.byLanguage).sort().forEach(([language, count]) => items.push([count, `${tr("languages")} · ${language}`])); q("#summary-cards").replaceChildren(...items.map(([value, label]) => { const card = el("div", "summary-card"); card.append(el("b", "", String(value)), el("span", "", label)); return card; })); }
function renderNotices() { const box = q("#notice-panel"); box.replaceChildren(); if (sourceLedger.some((source) => source.adapter === "legacy")) box.append(el("p", "", tr("legacyNotice"))); const english = lastAnalysis.warnings.filter((warning) => warning.code === "english"); if (english.length) { const card = el("section", "notice-card"); card.append(el("strong", "", tr("englishNotice"))); const list = el("ul", "issues neutral"); english.forEach((item) => list.append(el("li", "", `${item.participant} → ${item.evaluator} (${item.email}) · ${item.source}${item.rowNumber ? ` · ${tr("rowLabel")} ${item.rowNumber}` : ""}`))); card.append(list); box.append(card); } const cohortWarnings = lastAnalysis.warnings.filter((warning) => ["respondent-recommendation", "respondent-cohort-warning"].includes(warning.code)); if (cohortWarnings.length) { const card = el("section", "notice-card cohort-warning"); card.append(el("strong", "", tr("respondentGuidance"))); const list = el("ul", "issues neutral"); cohortWarnings.forEach((warning) => list.append(el("li", "", warning.code === "respondent-cohort-warning" ? tr("cohortWarning") : `${lang === "ro" ? "Peste recomandarea generală de 10 pentru" : "Above the general recommendation of 10 for"} ${warning.participant} (${warning.count})`))); card.append(list); box.append(card); } const allocationCodes = new Set(["allocation-invalid", "email-multiple-identifiers", "identifier-multiple-emails"]); const allocationIssues = lastAnalysis.blockers.filter((issue) => allocationCodes.has(issue.code)); if (allocationIssues.length) { const group = el("section", "allocation-errors"); group.append(el("strong", "", tr("correctAllocation"))); allocationIssues.forEach((issue) => { const card = el("article", "allocation-error"); card.append(el("b", "", allocationIssueTitle(issue))); const details = allocationIssueDetails(issue); if (details.length) { const list = el("ul", "issues"); details.forEach((detail) => list.append(el("li", "", detail))); card.append(list); } card.append(el("p", "", tr("correctAllocation"))); group.append(card); }); box.append(group); } const global = lastAnalysis.blockers.filter((blocker) => !blocker.participant && blocker.code !== "name-conflict" && !allocationCodes.has(blocker.code)); if (global.length) { const list = el("ul", "issues"); global.forEach((issue) => list.append(el("li", "", `${issue.source ? `${issue.source}: ` : ""}${issueMessage(issue)}`))); box.append(list); } }
function renderParticipants(participants) { const host = q("#participant-panels"); host.replaceChildren(); if (!participants.length) { host.append(el("p", "", tr("noData"))); return; } participants.forEach((participant) => { const summary = lastAnalysis.participantSummaries.find((item) => item.participant.participantEmail === participant.participantEmail) || lastAnalysis.participantSummaries.find((item) => item.participant.source === participant.source); const details = el("details", `participant ${summary?.blockers.length ? "has-errors" : ""}`); details.open = Boolean(summary?.blockers.length || ["legacy", "mass"].includes(participant.adapter)); const head = document.createElement("summary"); head.append(document.createTextNode(participant.participantName || participant.source), el("span", "", `${participant.source} · ${participant.rows.length} ${tr("rowsLabel")}`)); details.append(head); const body = el("div", "participant-body"); const correction = el("div", "correction-grid"); const nameLabel = el("label", "", tr("name")); const nameInput = document.createElement("input"); nameInput.value = participant.participantName; nameInput.addEventListener("change", () => { participant.participantName = nameInput.value; participant.rows.filter((row) => row.isSelf).forEach((row) => { row.evaluatorName = nameInput.value; }); render(); }); nameLabel.append(nameInput); const emailLabel = el("label", "", tr("email")); const emailInput = document.createElement("input"); emailInput.type = "email"; emailInput.value = participant.participantEmail; emailInput.addEventListener("change", () => { participant.participantEmail = emailInput.value; participant.rows.filter((row) => row.isSelf).forEach((row) => { row.evaluatorEmail = emailInput.value; }); render(); }); emailLabel.append(emailInput); correction.append(nameLabel, emailLabel); body.append(correction); if (summary?.blockers.length) { const list = el("ul", "issues"); summary.blockers.forEach((issue) => list.append(el("li", "", `${issue.source ? `${issue.source}: ` : ""}${issueMessage(issue)}`))); body.append(list); } const table = el("div", "language-table"); participant.rows.forEach((row) => { const line = el("div", "language-row"); line.append(el("span", "", row.evaluatorName || `${tr("rowLabel")} ${row.rowNumber || "—"}`), el("span", "", row.sourceRole || "—")); const label = el("label", "", tr("language")); const input = document.createElement("input"); input.maxLength = 2; input.value = row.language || "RO"; input.addEventListener("change", () => { row.language = input.value.toUpperCase(); render(); }); label.append(input); line.append(label); table.append(line); }); body.append(table); details.append(body); host.append(details); }); }
function renderConflicts() { const host = q("#conflict-panels"); host.replaceChildren(); lastAnalysis.conflicts.forEach((conflict) => { const card = el("section", "conflict"); card.append(el("h3", "", tr("conflict")), el("p", "", conflict.email)); const options = el("div", "conflict-options"); (conflict.variants || conflict.names.map((name) => ({ name, provenance: [] }))).forEach((variant) => { const label = document.createElement("label"); const radio = document.createElement("input"); radio.type = "radio"; radio.name = `conflict-${conflict.email}`; radio.value = variant.name; radio.checked = (conflict.selected || conflict.defaultName) === variant.name; const copy = el("span", "conflict-copy"); copy.append(document.createTextNode(variant.name)); if (variant.recommended) copy.append(el("em", "recommended", `★ ${tr("recommended")}`)); variant.provenance.forEach((entry) => copy.append(el("small", "", `${entry.source} · ${tr("rowLabel")} ${entry.rowNumber || "—"}`))); label.append(radio, copy); options.append(label); }); card.append(options); const button = el("button", "button small confirm-conflict", tr("confirmName")); button.type = "button"; button.addEventListener("click", () => { const selected = card.querySelector("input:checked"); if (selected) { nameChoices[conflict.email] = selected.value; render(); } }); card.append(button); host.append(card); }); }

setLang("ro");
