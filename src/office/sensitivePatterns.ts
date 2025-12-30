export type SensitiveType =
  | "email"
  | "phone"
  | "ssn"
  | "card"
  | "bank"
  | "insurancePolicy"
  | "employeeId"
  | "medicalRecordNumber";

export type SensitiveMatch = {
  type: SensitiveType;
  value: string;
};

// Email: basic RFC-ish pattern, anchored with word boundaries to avoid trailing punctuation.
const EMAIL_RE = /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi;

// SSN: common US format.
const SSN_RE = /\b\d{3}-\d{2}-\d{4}\b/g;

// SSN "last 4" commonly appears like "... social security number ... 8234".
const SSN_LAST4_CONTEXT_RE = /\b(?:social security number|ssn)\b[^0-9]{0,80}(\d{4})\b/gi;

// Phone: broad US-style formats (+1 optional), allows (), spaces, dots, dashes.
// Example matches: 2125551212, 212-555-1212, (212) 555-1212, +1 212 555 1212, 212.555.1212
// Notes:
// - Uses lookarounds instead of \b so leading "(" is included.
// - Avoids matching as part of longer digit sequences.
const PHONE_RE =
  /(?<!\d)(?:\+?\d{1,3}[-.\s]*)?(?:\(\s*\d{3}\s*\)|\d{3})[-.\s]*\d{3}[-.\s]*\d{4}(?!\d)/g;

// Credit/debit card candidates (13–19 digits) with spaces/dashes allowed.
// Prefer Luhn-valid numbers, but also allow non-Luhn numbers when clearly labeled
// (e.g., "credit card number 4532-...") to handle redaction test docs.
const CARD_CANDIDATE_RE = /(?<!\d)(?:\d[ -]*?){13,19}(?!\d)/g;

// IBAN candidates (allow spaces) e.g. "GB82 WEST 1234 5698 7654 32"
const IBAN_CANDIDATE_RE = /(?<![A-Z0-9])[A-Z]{2}\d{2}(?:[ ]?[A-Z0-9]){11,30}(?![A-Z0-9])/gi;

// Bank-ish keywords followed by a number.
const ROUTING_RE = /\brouting(?:\s*number)?\s*[:\-]?\s*(\d{9})\b/gi;
const ACCOUNT_RE = /\b(?:account|acct)(?:\s*(?:number|no\.?|#))?\s*[:\-]?\s*(\d{6,17})\b/gi;
const SORT_CODE_RE = /\bsort\s*code\s*[:\-]?\s*(\d{2}[- ]?\d{2}[- ]?\d{2})\b/gi;

// Common document identifiers (prefix-based). We redact the value but preserve prefix via replacement string.
// Example from sample doc: "INS-44556677", "EMP-2024-5567", "MRN- 998877"
const INS_POLICY_RE = /\bINS[-\s]*\d{6,14}\b/gi;
const EMPLOYEE_ID_RE = /\bEMP[-\s]*\d{2,4}(?:[-\s]*\d{2,6})+\b/gi;
const MRN_RE = /\bMRN[-\s]*\d{4,14}\b/gi;

export function findSensitiveMatches(text: string): SensitiveMatch[] {
  const out: SensitiveMatch[] = [];

  for (const value of matchAllUnique(text, EMAIL_RE)) out.push({ type: "email", value });
  for (const value of matchAllUnique(text, SSN_RE)) out.push({ type: "ssn", value });
  for (const value of findSsnLast4(text)) out.push({ type: "ssn", value });
  for (const value of matchAllUnique(text, PHONE_RE)) out.push({ type: "phone", value });
  for (const value of findCreditCards(text)) out.push({ type: "card", value });
  for (const value of findBankAccounts(text)) out.push({ type: "bank", value });
  for (const value of matchAllUnique(text, INS_POLICY_RE)) out.push({ type: "insurancePolicy", value });
  for (const value of matchAllUnique(text, EMPLOYEE_ID_RE)) out.push({ type: "employeeId", value });
  for (const value of matchAllUnique(text, MRN_RE)) out.push({ type: "medicalRecordNumber", value });

  return out;
}

function matchAllUnique(text: string, re: RegExp): string[] {
  const unique = new Map<string, string>();
  re.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    const raw = (m[0] ?? "").trim();
    if (!raw) continue;
    const key = raw.toLowerCase();
    if (!unique.has(key)) unique.set(key, raw);
  }
  return [...unique.values()];
}

function findSsnLast4(text: string): string[] {
  const unique = new Map<string, string>();
  SSN_LAST4_CONTEXT_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = SSN_LAST4_CONTEXT_RE.exec(text)) !== null) {
    const last4 = (m[1] ?? "").trim();
    if (!/^\d{4}$/.test(last4)) continue;
    // Avoid redacting common years if they appear in SSN context by accident.
    const asNum = Number(last4);
    if (asNum >= 1900 && asNum <= 2099) continue;
    if (!unique.has(last4)) unique.set(last4, last4);
  }
  return [...unique.values()];
}

function findCreditCards(text: string): string[] {
  const unique = new Map<string, string>(); // key: digits only
  CARD_CANDIDATE_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = CARD_CANDIDATE_RE.exec(text)) !== null) {
    const raw = (m[0] ?? "").trim();
    const digits = raw.replace(/[^\d]/g, "");
    if (digits.length < 13 || digits.length > 19) continue;

    const luhnOk = luhnCheck(digits);
    const idx = typeof m.index === "number" ? m.index : -1;
    const ctx =
      idx >= 0
        ? text
            .slice(Math.max(0, idx - 50), Math.min(text.length, idx + raw.length + 50))
            .toLowerCase()
        : "";
    const keywordOk =
      /\b(?:credit\s*card|debit\s*card|card\s*number|visa|mastercard|amex|american\s*express)\b/.test(ctx);

    if (!luhnOk && !keywordOk) continue;
    if (!unique.has(digits)) unique.set(digits, raw);
  }
  return [...unique.values()];
}

function findBankAccounts(text: string): string[] {
  const unique = new Map<string, string>();

  // IBAN (validated).
  IBAN_CANDIDATE_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = IBAN_CANDIDATE_RE.exec(text)) !== null) {
    const raw = (m[0] ?? "").trim();
    const normalized = raw.replace(/\s+/g, "").toUpperCase();
    if (!isValidIban(normalized)) continue;
    if (!unique.has(normalized)) unique.set(normalized, raw);
  }

  // Routing number (US).
  ROUTING_RE.lastIndex = 0;
  while ((m = ROUTING_RE.exec(text)) !== null) {
    const raw = (m[1] ?? "").trim();
    if (!/^\d{9}$/.test(raw)) continue;
    const key = `routing:${raw}`;
    if (!unique.has(key)) unique.set(key, raw);
  }

  // Account number (US-ish).
  ACCOUNT_RE.lastIndex = 0;
  while ((m = ACCOUNT_RE.exec(text)) !== null) {
    const raw = (m[1] ?? "").trim();
    if (!/^\d{6,17}$/.test(raw)) continue;
    const key = `account:${raw}`;
    if (!unique.has(key)) unique.set(key, raw);
  }

  // Sort code (UK).
  SORT_CODE_RE.lastIndex = 0;
  while ((m = SORT_CODE_RE.exec(text)) !== null) {
    const raw = (m[1] ?? "").trim();
    const digits = raw.replace(/[^\d]/g, "");
    if (digits.length !== 6) continue;
    const key = `sort:${digits}`;
    if (!unique.has(key)) unique.set(key, raw);
  }

  return [...unique.values()];
}

function luhnCheck(digits: string): boolean {
  let sum = 0;
  let shouldDouble = false;

  for (let i = digits.length - 1; i >= 0; i--) {
    let d = digits.charCodeAt(i) - 48;
    if (d < 0 || d > 9) return false;
    if (shouldDouble) {
      d *= 2;
      if (d > 9) d -= 9;
    }
    sum += d;
    shouldDouble = !shouldDouble;
  }

  return sum % 10 === 0;
}

function isValidIban(iban: string): boolean {
  // Basic shape check.
  if (!/^[A-Z]{2}\d{2}[A-Z0-9]{11,30}$/.test(iban)) return false;

  // Rearrange: move first 4 chars to end.
  const rearranged = iban.slice(4) + iban.slice(0, 4);

  // Convert letters to numbers (A=10..Z=35) and compute mod 97 iteratively.
  let mod = 0;
  for (let i = 0; i < rearranged.length; i++) {
    const ch = rearranged.charCodeAt(i);
    if (ch >= 48 && ch <= 57) {
      mod = (mod * 10 + (ch - 48)) % 97;
    } else if (ch >= 65 && ch <= 90) {
      const val = ch - 55; // 'A'->10
      mod = (mod * 10 + Math.floor(val / 10)) % 97;
      mod = (mod * 10 + (val % 10)) % 97;
    } else {
      return false;
    }
  }

  return mod === 1;
}


