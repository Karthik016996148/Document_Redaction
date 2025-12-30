export type AppState = {
  officeReady: boolean;
  running: boolean;
  logs: string[];
  lastResult:
    | null
    | {
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
};

type AppHandlers = {
  onRun: () => void;
};

let root: HTMLElement | null = null;
let handlers: AppHandlers | null = null;

export function renderAppShell(state: AppState, h: AppHandlers) {
  handlers = h;
  root = document.getElementById("app");
  if (!root) throw new Error("Missing #app root element");
  root.innerHTML = "";
  root.appendChild(buildTree(state));
}

export function setAppState(state: AppState) {
  if (!root || !handlers) return;
  root.innerHTML = "";
  root.appendChild(buildTree(state));
}

function buildTree(state: AppState): HTMLElement {
  const wrap = el("div", "wrap");
  const card = el("div", "card");
  wrap.appendChild(card);

  const header = el("div", "header");
  card.appendChild(header);

  const titleRow = el("div", "titleRow");
  header.appendChild(titleRow);

  const title = el("h1", "title");
  title.textContent = "Document Redaction";
  titleRow.appendChild(title);

  const badge = el("div", "badge");
  badge.textContent = state.officeReady ? "Connected to Word" : "Not connected";
  titleRow.appendChild(badge);

  const subtitle = el("p", "subtitle");
  subtitle.textContent =
    "One click will enable Track Changes (if supported), add a CONFIDENTIAL header, and redact emails, phone numbers, SSNs, credit/debit cards, common bank identifiers, insurance policy numbers, employee IDs, and MRNs across the entire document.";
  header.appendChild(subtitle);

  const content = el("div", "content");
  card.appendChild(content);

  const btn = document.createElement("button");
  btn.className = "primaryBtn";
  btn.textContent = state.running ? "Running…" : "Redact & Mark Confidential";
  btn.disabled = !state.officeReady || state.running;
  btn.addEventListener("click", () => handlers?.onRun());
  content.appendChild(btn);

  const grid2 = el("div", "grid2");
  content.appendChild(grid2);

  grid2.appendChild(metric("Track Changes", state.officeReady ? "Auto" : "—", state.officeReady ? "ok" : "bad"));
  grid2.appendChild(
    metric(
      "Last run",
      state.lastResult
        ? `${state.lastResult.redactionsTotal} redactions`
        : state.running
          ? "In progress"
          : "Not run yet",
      state.lastResult ? "ok" : state.running ? "ok" : "bad",
    ),
  );

  const log = el("div", "log");
  content.appendChild(log);
  for (const line of state.logs.slice(-120)) {
    const ln = el("div", "logLine");
    ln.textContent = line;
    log.appendChild(ln);
  }

  return wrap;
}

function metric(label: string, value: string, tone: "ok" | "bad") {
  const pill = el("div", "pill");
  const l = el("p", "pillLabel");
  l.textContent = label;
  const v = el("p", `pillValue ${tone}`);
  v.textContent = value;
  pill.appendChild(l);
  pill.appendChild(v);
  return pill;
}

function el(tag: string, className?: string) {
  const e = document.createElement(tag);
  if (className) e.className = className;
  return e;
}


