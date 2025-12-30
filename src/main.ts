import "./styles.css";
import { runRedactionWorkflow } from "./office/runRedactionWorkflow";
import { renderAppShell, setAppState, type AppState } from "./ui/appShell";

declare const Office: any;
declare const Word: any;

const initialState: AppState = {
  officeReady: false,
  running: false,
  lastResult: null,
  logs: ["Loading…"],
};

renderAppShell(initialState, {
  onRun: async () => {
    if (!initialState.officeReady || initialState.running) return;
    if (typeof Word === "undefined" || !Word?.run) {
      initialState.logs = [
        "Word JavaScript API not available yet.",
        "If you're in Word, wait 1–2 seconds and try again.",
        "If you're not in Word, sideload the manifest into Microsoft Word (web or desktop) and open the taskpane there.",
      ];
      setAppState(initialState);
      return;
    }
    initialState.running = true;
    initialState.lastResult = null;
    initialState.logs = [];
    setAppState(initialState);

    try {
      const result = await runRedactionWorkflow((line) => {
        initialState.logs = [...initialState.logs, line];
        setAppState(initialState);
      });
      initialState.lastResult = result;
      initialState.logs = [
        ...initialState.logs,
        `Done. Redacted: ${result.redactionsTotal} (emails ${result.emails}, phones ${result.phones}, ssns ${result.ssns}, cards ${result.cards}, bank ${result.bankAccounts}, insurance ${result.insurancePolicies}, employeeIds ${result.employeeIds}, mrns ${result.medicalRecordNumbers}). Header updated: ${result.headerUpdated ? "yes" : "no"}. Track Changes: ${result.trackChangesEnabled ? "enabled" : "not available"}.`,
      ];
    } catch (e: any) {
      initialState.logs = [
        ...initialState.logs,
        `Error: ${e?.message ?? String(e)}`,
      ];
    } finally {
      initialState.running = false;
      setAppState(initialState);
    }
  },
});

function bootInOffice(info?: { host?: string; platform?: string }) {
  const host = info?.host ?? Office?.context?.host;
  const platform = info?.platform ?? Office?.context?.platform;

  const isWordHost =
    host === Office?.HostType?.Word ||
    host?.toString?.().toLowerCase?.() === "word";

  // Don't require Word.run at boot time—Word can populate its globals slightly after onReady in some hosts.
  initialState.officeReady = isWordHost;
  initialState.logs = isWordHost
    ? [
        `Ready in Word${platform ? ` (${platform})` : ""}.`,
        "Tip: If the button says Word API isn't available yet, wait a moment and retry.",
      ]
    : [
        `Loaded in Office host: ${host ?? "unknown"}.`,
        "This challenge must run inside Microsoft Word.",
        "Open Word (web or desktop), sideload the manifest, and open the taskpane there.",
      ];
  setAppState(initialState);
}

if (typeof Office !== "undefined" && Office?.onReady) {
  Office.onReady((info: any) => bootInOffice(info));
} else {
  initialState.logs = [
    "Office.js not detected. This taskpane must be loaded inside Word (desktop or web).",
  ];
  setAppState(initialState);
}


