import { findSensitiveMatches, type SensitiveMatch } from "./sensitivePatterns";

declare const Office: any;
declare const Word: any;

export type RedactionResult = {
  trackChangesEnabled: boolean;
  headerUpdated: boolean;
  emails: number;
  phones: number;
  ssns: number;
  cards: number;
  bankAccounts: number;
  insurancePolicies: number;
  employeeIds: number;
  medicalRecordNumbers: number;
  redactionsTotal: number;
};

export async function runRedactionWorkflow(log: (line: string) => void): Promise<RedactionResult> {
  if (typeof Word === "undefined" || !Word?.run) {
    throw new Error("Word JavaScript API not available. Open this add-in inside Word.");
  }

  const trackingSupported =
    typeof Office !== "undefined" &&
    Office?.context?.requirements?.isSetSupported?.("WordApi", "1.5") === true;

  return await Word.run(async (context: any) => {
    const result: RedactionResult = {
      trackChangesEnabled: false,
      headerUpdated: false,
      emails: 0,
      phones: 0,
      ssns: 0,
      cards: 0,
      bankAccounts: 0,
      insurancePolicies: 0,
      employeeIds: 0,
      medicalRecordNumbers: 0,
      redactionsTotal: 0,
    };

    if (trackingSupported && context?.document && "changeTrackingMode" in context.document) {
      try {
        context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        result.trackChangesEnabled = true;
        log("Track Changes: enabled (WordApi 1.5).");
      } catch {
        log("Track Changes: supported but could not be enabled (continuing).");
      }
    } else {
      log("Track Changes: not available (WordApi 1.5 not supported).");
    }

    // 1) Add confidentiality header (tracked if tracking is enabled).
    log("Updating header: CONFIDENTIAL DOCUMENT");
    try {
      result.headerUpdated = await addConfidentialHeader(context);
      await context.sync();
      log(result.headerUpdated ? "Header updated." : "Header already present (no change).");
    } catch (e: any) {
      // Continue with redaction even if header update fails.
      result.headerUpdated = false;
      log("Header update failed (continuing with redaction).");
      log(formatOfficeError(e));
    }

    // 2) Load full document text.
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const text: string = body.text ?? "";
    log(`Scanning document text (${text.length.toLocaleString()} chars)…`);

    const matches = findSensitiveMatches(text);
    const emails = matches.filter((m) => m.type === "email");
    const phones = matches.filter((m) => m.type === "phone");
    const ssns = matches.filter((m) => m.type === "ssn");
    const cards = matches.filter((m) => m.type === "card");
    const bankAccounts = matches.filter((m) => m.type === "bank");
    const insurancePolicies = matches.filter((m) => m.type === "insurancePolicy");
    const employeeIds = matches.filter((m) => m.type === "employeeId");
    const medicalRecordNumbers = matches.filter((m) => m.type === "medicalRecordNumber");

    result.emails = emails.length;
    result.phones = phones.length;
    result.ssns = ssns.length;
    result.cards = cards.length;
    result.bankAccounts = bankAccounts.length;
    result.insurancePolicies = insurancePolicies.length;
    result.employeeIds = employeeIds.length;
    result.medicalRecordNumbers = medicalRecordNumbers.length;
    log(
      `Found: ${result.emails} emails, ${result.phones} phones, ${result.ssns} SSNs, ${result.cards} cards, ${result.bankAccounts} bank identifiers, ${result.insurancePolicies} insurance policies, ${result.employeeIds} employee IDs, ${result.medicalRecordNumbers} MRNs.`,
    );

    // 3) Redact by searching exact matches and replacing ranges.
    let redacted = 0;

    redacted += await redactGroup(context, body, emails, "[REDACTED EMAIL]", log);
    redacted += await redactGroup(context, body, phones, "[REDACTED PHONE]", log);
    redacted += await redactGroup(context, body, ssns, "[REDACTED SSN]", log);
    redacted += await redactGroup(context, body, cards, "[REDACTED CARD]", log);
    redacted += await redactGroup(context, body, bankAccounts, "[REDACTED BANK]", log);
    redacted += await redactGroup(context, body, insurancePolicies, "INS-[REDACTED]", log);
    redacted += await redactGroup(context, body, employeeIds, "EMP-[REDACTED]", log);
    redacted += await redactGroup(context, body, medicalRecordNumbers, "MRN-[REDACTED]", log);

    result.redactionsTotal = redacted;
    return result;
  });
}

async function redactGroup(
  context: any,
  body: any,
  items: SensitiveMatch[],
  replacement: string,
  log: (line: string) => void,
): Promise<number> {
  if (items.length === 0) return 0;
  let count = 0;

  // De-dupe already happens in matcher, but keep stable order for logs.
  for (const { value } of items) {
    const ranges = body.search(value, {
      matchCase: false,
      matchWholeWord: false,
      ignorePunct: false,
      ignoreSpace: false,
    });
    ranges.load("items");
    await context.sync();

    if (!ranges.items || ranges.items.length === 0) continue;

    for (const r of ranges.items) {
      r.insertText(replacement, Word.InsertLocation.replace);
      count += 1;
    }
    await context.sync();
    log(`Redacted ${ranges.items.length} occurrence(s) of: ${value}`);
  }

  return count;
}

async function addConfidentialHeader(context: any): Promise<boolean> {
  const sections = context.document.sections;
  sections.load("items");
  await context.sync();

  let changed = false;
  for (const section of sections.items) {
    const updated = await tryUpdateAnyHeaderInSection(context, section);
    if (updated) changed = true;
  }

  // Some Word hosts/docs can throw on Section.getHeader (GeneralException). In that case,
  // fall back to a top-of-document banner so the requirement is still met for the user.
  if (!changed) {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const bodyText: string = (body.text ?? "").trimStart();
    if (!bodyText.toUpperCase().startsWith("CONFIDENTIAL DOCUMENT")) {
      // Body.insertParagraph supports InsertLocation.start across hosts more reliably than Range.insertParagraph.
      const p = body.insertParagraph("CONFIDENTIAL DOCUMENT", Word.InsertLocation.start);
      p.font.bold = true;
      p.font.color = "#B00020";
      p.alignment = Word.Alignment.centered;
      changed = true;
    }
  }

  return changed;
}

async function tryUpdateAnyHeaderInSection(context: any, section: any): Promise<boolean> {
  const headerTypes = [
    Word.HeaderFooterType.primary,
    Word.HeaderFooterType.firstPage,
    Word.HeaderFooterType.evenPages,
  ];

  for (const type of headerTypes) {
    try {
      const header = section.getHeader(type);
      const headerRange = header.getRange();
      headerRange.load("text");
      await context.sync();

      const current: string = (headerRange.text ?? "").trim();
      if (current.toUpperCase().includes("CONFIDENTIAL DOCUMENT")) return false;

      // Range.insertParagraph does not support InsertLocation.start (it expects before/after in some hosts),
      // so use insertText at the start of the header range instead.
      const inserted = headerRange.insertText("CONFIDENTIAL DOCUMENT\r", Word.InsertLocation.start);
      inserted.font.bold = true;
      inserted.font.color = "#B00020";
      inserted.paragraphFormat.alignment = Word.Alignment.centered;
      return true;
    } catch {
      // Try the next header type (some docs/hosts reject certain header types or header access entirely).
    }
  }

  return false;
}

function formatOfficeError(e: any): string {
  const msg = e?.message ?? String(e);
  const code = e?.code ? ` code=${e.code}` : "";
  const debugInfo = e?.debugInfo ? ` debugInfo=${safeJson(e.debugInfo)}` : "";
  return `OfficeExtension.Error:${code} ${msg}${debugInfo}`.trim();
}

function safeJson(v: unknown): string {
  try {
    return JSON.stringify(v);
  } catch {
    return "[unserializable]";
  }
}


