import { initI18n, t } from "../i18n/i18n";
import "./taskpane.css";

interface AdminConfig {
  defaultIntervalMinutes: number;
  minIntervalMinutes: number;
  maxIntervalMinutes: number;
  allowUserOverride: boolean;
}

const SETTINGS_KEY = "autosave_interval_minutes";
const VERSION = "1.0.0";

let config: AdminConfig = {
  defaultIntervalMinutes: 1,
  minIntervalMinutes: 1,
  maxIntervalMinutes: 60,
  allowUserOverride: true,
};
let intervalMinutes = config.defaultIntervalMinutes;
let timerHandle: ReturnType<typeof setInterval> | null = null;
let isEnabled = true;
let isReadOnly = false;
let confirmedFilePath = false;

// ── Logging ───────────────────────────────────────────────────────────────────

function log(msg: string): void {
  const ts = new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" });
  console.log(`[AutoSave ${ts}] ${msg}`);
}

// ── DOM helpers ───────────────────────────────────────────────────────────────

function el<T extends HTMLElement>(id: string): T {
  return document.getElementById(id) as T;
}

function setText(id: string, text: string): void {
  const node = document.getElementById(id);
  if (node) node.textContent = text;
}

function show(id: string): void {
  const node = document.getElementById(id);
  if (node) node.style.display = "";
}

function hide(id: string): void {
  const node = document.getElementById(id);
  if (node) node.style.display = "none";
}

function setNotification(message: string, type: "info" | "warning" | "error" | "success"): void {
  const banner = el<HTMLDivElement>("notification-banner");
  if (!banner) return;
  banner.textContent = message;
  banner.className = `notification notification--${type}`;
  show("notification-banner");
}

function clearNotification(): void {
  hide("notification-banner");
}

// ── Config loading ────────────────────────────────────────────────────────────

async function loadConfig(): Promise<void> {
  try {
    const response = await fetch("config/autosave-config.json");
    if (response.ok) {
      config = (await response.json()) as AdminConfig;
    }
  } catch (err) {
    console.warn("Could not load autosave-config.json, using defaults.", err);
  }
}

// ── User settings ─────────────────────────────────────────────────────────────

function loadUserInterval(): number {
  try {
    const stored = Office.context.document.settings.get(SETTINGS_KEY) as number | null;
    if (stored !== null && stored !== undefined) {
      return Math.min(Math.max(stored, config.minIntervalMinutes), config.maxIntervalMinutes);
    }
  } catch {
    // settings not available
  }
  return config.defaultIntervalMinutes;
}

function saveUserInterval(minutes: number): void {
  try {
    Office.context.document.settings.set(SETTINGS_KEY, minutes);
    Office.context.document.settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        setNotification(t("settings_saved"), "success");
        setTimeout(clearNotification, 3000);
      }
    });
  } catch (err) {
    console.error("Could not save user settings.", err);
  }
}

// ── Save logic ────────────────────────────────────────────────────────────────

async function hasFilePath(): Promise<boolean> {
  if (confirmedFilePath) return true;
  return new Promise<boolean>((resolve) => {
    try {
      Office.context.document.getFilePropertiesAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const url = result.value.url;
          const hasPath = !!url && url.trim() !== "";
          if (hasPath) confirmedFilePath = true;
          resolve(hasPath);
        } else {
          resolve(false);
        }
      });
    } catch {
      resolve(false);
    }
  });
}

async function performSave(): Promise<void> {
  log(`Timer fired — isEnabled: ${isEnabled} isReadOnly: ${isReadOnly}`);
  if (!isEnabled || isReadOnly) return;

  const hasSavedPath = await hasFilePath();
  log(`hasFilePath: ${hasSavedPath}`);
  if (!hasSavedPath) {
    setNotification(t("unsaved_document"), "warning");
    return;
  }

  const host = Office.context.host;
  log(`Saving — host: ${host}`);

  try {
    if (host === Office.HostType.Excel) {
      await Excel.run(async (context) => {
        context.workbook.save(Excel.SaveBehavior.save);
        await context.sync();
      });
    } else if (host === Office.HostType.Word) {
      await Word.run(async (context) => {
        context.document.save();
        await context.sync();
      });
    } else {
      // PowerPoint — saveAsync exists at runtime but is absent from @types/office-js
      await new Promise<void>((resolve, reject) => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (Office.context.document as any).saveAsync((result: Office.AsyncResult<void>) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            const errMsg = result.error?.message ?? "Unknown error";
            if (errMsg.toLowerCase().includes("read-only") || errMsg.toLowerCase().includes("protected")) {
              isReadOnly = true;
              setNotification(t("protected_view"), "warning");
              stopTimer();
              resolve();
            } else {
              reject(new Error(errMsg));
            }
          }
        });
      });
    }

    log("Save succeeded");
    clearNotification();
    setText("last-saved-value", new Date().toLocaleTimeString());
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    log(`Save failed: ${message}`);
    setNotification(t("save_error"), "error");
  }
}

// ── Timer ─────────────────────────────────────────────────────────────────────

function startTimer(): void {
  stopTimer();
  if (isReadOnly) return;
  log(`Timer started — interval: ${intervalMinutes} min`);
  timerHandle = setInterval(() => {
    performSave().catch((err) => log(`Unhandled save error: ${err}`));
  }, intervalMinutes * 60 * 1000);
}

function stopTimer(): void {
  if (timerHandle !== null) {
    clearInterval(timerHandle);
    timerHandle = null;
  }
}

// ── UI ────────────────────────────────────────────────────────────────────────

function applyStrings(): void {
  setText("addin-title", t("addin_name"));
  setText("toggle-label", isEnabled ? t("toggle_on") : t("toggle_off"));
  setText("last-saved-label", t("last_saved"));
  setText("last-saved-value", t("never_saved"));
  setText("interval-label-text", t("interval_label"));
  setText("interval-unit-text", t("interval_unit"));
  setText("save-settings-btn", t("save_settings"));
  setText("version-text", VERSION);
  el<HTMLInputElement>("toggle-checkbox").setAttribute("aria-label", t("toggle_on"));
}

function renderSettingsPanel(): void {
  if (config.allowUserOverride) {
    show("settings-panel");
    hide("managed-by-it-msg");
    const slider = el<HTMLInputElement>("interval-slider");
    const numberInput = el<HTMLInputElement>("interval-number");
    slider.min = String(config.minIntervalMinutes);
    slider.max = String(config.maxIntervalMinutes);
    slider.value = String(intervalMinutes);
    numberInput.min = String(config.minIntervalMinutes);
    numberInput.max = String(config.maxIntervalMinutes);
    numberInput.value = String(intervalMinutes);
  } else {
    hide("settings-panel");
    show("managed-by-it-msg");
    setText("managed-by-it-text", t("managed_by_it"));
  }
}

function updateToggleUI(): void {
  const checkbox = el<HTMLInputElement>("toggle-checkbox");
  checkbox.checked = isEnabled;
  setText("toggle-label", isEnabled ? t("toggle_on") : t("toggle_off"));
  el<HTMLDivElement>("toggle-row").classList.toggle("toggle-row--off", !isEnabled);
}

function wireEvents(): void {
  const checkbox = el<HTMLInputElement>("toggle-checkbox");
  checkbox.addEventListener("change", () => {
    isEnabled = checkbox.checked;
    updateToggleUI();
    if (isEnabled) startTimer(); else stopTimer();
  });

  const slider = el<HTMLInputElement>("interval-slider");
  const numberInput = el<HTMLInputElement>("interval-number");

  slider.addEventListener("input", () => { numberInput.value = slider.value; });

  numberInput.addEventListener("input", () => {
    const val = parseInt(numberInput.value, 10);
    if (!isNaN(val)) {
      slider.value = String(Math.min(Math.max(val, config.minIntervalMinutes), config.maxIntervalMinutes));
    }
  });

  el<HTMLButtonElement>("save-settings-btn").addEventListener("click", () => {
    const raw = parseInt(numberInput.value, 10);
    if (isNaN(raw)) return;
    intervalMinutes = Math.min(Math.max(raw, config.minIntervalMinutes), config.maxIntervalMinutes);
    slider.value = String(intervalMinutes);
    numberInput.value = String(intervalMinutes);
    saveUserInterval(intervalMinutes);
    if (isEnabled) startTimer();
  });
}

// ── Event-based activation ────────────────────────────────────────────────────
// Registered at module load time so Office can find the function when the
// OnDocumentOpen event fires — before Office.onReady has even run.

// eslint-disable-next-line @typescript-eslint/no-explicit-any
(Office as any).actions?.associate("onDocumentOpen", (event: any) => {
  log("OnDocumentOpen event received");
  event.completed();
});

// ── Initialization ────────────────────────────────────────────────────────────

Office.onReady(async () => {
  const locale = Office.context.displayLanguage ?? "en";

  await loadConfig();
  await initI18n(locale);

  isReadOnly = await (async () => {
    try {
      return await new Promise<boolean>((resolve) => {
        Office.context.document.getFilePropertiesAsync((result) => {
          resolve(result.status !== Office.AsyncResultStatus.Succeeded);
        });
      });
    } catch {
      return false;
    }
  })();

  intervalMinutes = loadUserInterval();

  // Render UI — visible when user opens the task pane for settings.
  applyStrings();
  renderSettingsPanel();
  wireEvents();
  updateToggleUI();

  if (isReadOnly) {
    setNotification(t("protected_view"), "warning");
    isEnabled = false;
    updateToggleUI();
  } else {
    startTimer();
    // Persist auto-start so future document opens load the runtime automatically
    try {
      await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    } catch {
      // Not available in all hosts/versions — safe to ignore
    }
  }

  log("Runtime ready — autosave running in background");
});
