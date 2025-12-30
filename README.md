# Document Redaction Word Add-in (TypeScript)

This is a **Microsoft Word taskpane add-in** built with **TypeScript + Vite**.

Clicking the single primary button will:
- **Enable Track Changes** (only when `WordApi 1.5` is available)
- **Insert a "CONFIDENTIAL DOCUMENT" header**
- **Redact sensitive identifiers** across the *entire* document (search + replace)

All edits are performed inside `Word.run(...)`, so when Track Changes is available, the changes are visible in Word’s **Review → Track Changes** history.

## What gets redacted

- Emails
- Phone numbers
- SSNs (full SSN pattern + SSN “last 4” when it appears in SSN context)
- Credit/debit card numbers (Luhn-validated; also redacts when clearly labeled like “credit card number …”)
- Bank identifiers (IBAN validated + routing/account/sort code when keyword-labeled)
- Insurance policy numbers (e.g. `INS-44556677` → `INS-[REDACTED]`)
- Employee IDs (e.g. `EMP-2024-5567` → `EMP-[REDACTED]`)
- MRNs / medical record numbers (e.g. `MRN- 998877` → `MRN-[REDACTED]`)

## Quickstart

### 1) Install dependencies

```bash
npm install
```

### 2) Start the dev server + attempt sideload

```bash
npm start
```

- Serves the taskpane on **`https://localhost:3000`**
- Attempts to sideload `manifest.xml` into Word

If sideloading prompts about **localhost loopback** on Windows, you may need to run once in an **elevated (Administrator)** terminal.

## Manual sideload (recommended)

### 1) Run the dev server

```bash
npm run dev
```

### 2) Sideload the manifest in Word
- Use `manifest.xml` at the project root.
- Follow Microsoft’s sideload instructions: `https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing`

Once sideloaded, open the taskpane and click **“Redact & Mark Confidential”**.

## How it works (short)
- Loads the full document body text (`document.body.text`)
- Detects sensitive tokens using regex + validation (e.g., IBAN mod-97, card Luhn/keyword context)
- Uses Word search (`body.search(...)`) and replaces matches with redaction markers
- Inserts a “CONFIDENTIAL DOCUMENT” header and enables Track Changes when supported

## Notes
- This repo intentionally does **not** commit `node_modules`. Install with `npm install`.
- `dist/` is not committed; build output is generated with `npm run build`.

## Build

```bash
npm run build
```

## Repo structure
- `src/main.ts`: taskpane bootstrap + UI state/logging
- `src/ui/appShell.ts`: taskpane UI rendering
- `src/office/runRedactionWorkflow.ts`: Word API workflow (Track Changes, header, replacement)
- `src/office/sensitivePatterns.ts`: sensitive token detection + validation

