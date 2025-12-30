# Document Redaction Word Add-in (TypeScript)

Word taskpane add-in that:
- Enables **Track Changes** (when `WordApi 1.5` is available)
- Inserts a **CONFIDENTIAL DOCUMENT** header
- Redacts common sensitive data across the full document

## What gets redacted

- Emails
- Phone numbers
- SSNs (including SSN last-4 when mentioned in SSN context)
- Credit/debit card numbers
- Common bank identifiers (IBAN + routing/account/sort code patterns)
- Insurance policy numbers (e.g. `INS-44556677` → `INS-[REDACTED]`)
- Employee IDs (e.g. `EMP-2024-5567` → `EMP-[REDACTED]`)
- MRNs / medical record numbers (e.g. `MRN- 998877` → `MRN-[REDACTED]`)

## Run locally

1. Install dependencies:

   `npm install`

2. Start dev server + sideload attempt:

   `npm start`

3. If sideloading prompts about localhost loopback on Windows, you may need to run once in an elevated terminal.

## Manual sideload (recommended if auto-sideload prompts)

1. Run the dev server:
   - `npm run dev` (serves `https://localhost:3000`)
2. In Word (web or desktop), sideload `manifest.xml`.
   - Microsoft steps: `https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing`

## Notes
- This repo intentionally does **not** commit `node_modules`. Install with `npm install`.

