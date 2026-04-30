import { initI18n, t } from "../i18n/i18n";
import "./taskpane.css";

interface AdminConfig {
  defaultIntervalMinutes: number;
  minIntervalMinutes: number;
  maxIntervalMinutes: number;
  allowUserOverride: boolean;
}

const SETTINGS_KEY = "autosave_interval_minutes";
const VERSION = "1.0.3";

// State lives on `window` so it survives script re-evaluation when the task
// pane HTML is reloaded inside the same shared runtime. Module-level `let`
// variables reset on every re-eval; `window` properties do not.
declare global {
  interface Window {
    __as_initialized: boolean;
    __as_timerHandle: ReturnType<typeof setInterval> | null;
    __as_enabled: boolean;
    __as_intervalMinutes: number;
    __as_lastSaved: string;
    __as_isReadOnly: boolean;
    __as_confirmedFilePath: boolean;
    __as_config: AdminConfig | null;
  }
}

const state = {
  get initialized(): boolean { return window.__as_initialized ?? false; },
  set initialized(v: boolean) { window.__as_initialized = v; },

  get timerHandle(): ReturnType<typeof setInterval> | null { return window.__as_timerHandle ?? null; },
  set timerHandle(v: ReturnType<typeof setInterval> | null) { window.__as_timerHandle = v; },

  get enabled(): boolean { return window.__as_enabled ?? true; },
  set enabled(v: boolean) { window.__as_enabled = v; },

  get intervalMinutes(): number { return window.__as_intervalMinutes ?? 1; },
  set intervalMinutes(v: number) { window.__as_intervalMinutes = v; },

  get lastSaved(): string { return window.__as_lastSaved ?? ""; },
  set lastSaved(v: string) { window.__as_lastSaved = v; },

  get isReadOnly(): boolean { return window.__as_isReadOnly ?? false; },
  set isReadOnly(v: boolean) { window.__as_isReadOnly = v; },

  get confirmedFilePath(): boolean { return window.__as_confirmedFilePath ?? false; },
  set confirmedFilePath(v: boolean) { window.__as_confirmedFilePath = v; },

  get config(): AdminConfig | null { return window.__as_config ?? null; },
  set config(v: AdminConfig | null) { window.__as_config = v; },
};

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

const CONFIG_DEFAULTS: AdminConfig = {
  defaultIntervalMinutes: 1,
  minIntervalMinutes: 1,
  maxIntervalMinutes: 60,
  allowUserOverride: true,
};

async function loadConfig(): Promise<AdminConfig> {
  try {
    const response = await fetch("config/autosave-config.json");
    if (response.ok) return (await response.json()) as AdminConfig;
  } catch (err) {
    console.warn("Could not load autosave-config.json, using defaults.", err);
  }
  return { ...CONFIG_DEFAULTS };
}

// ── User settings ─────────────────────────────────────────────────────────────

function loadUserInterval(cfg: AdminConfig): number {
  try {
    const stored = Office.context.document.settings.get(SETTINGS_KEY) as number | null;
    if (stored !== null && stored !== undefined) {
      return Math.min(Math.max(stored, cfg.minIntervalMinutes), cfg.maxIntervalMinutes);
    }
  } catch {
    // settings not available yet
  }
  return cfg.defaultIntervalMinutes;
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
  if (state.confirmedFilePath) return true;
  return new Promise<boolean>((resolve) => {
    try {
      Office.context.document.getFilePropertiesAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const url = result.value.url;
          const hasPath = !!url && url.trim() !== "";
          if (hasPath) state.confirmedFilePath = true;
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
  log(`Timer fired — enabled: ${state.enabled}, readOnly: ${state.isReadOnly}`);
  if (!state.enabled || state.isReadOnly) return;

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
              state.isReadOnly = true;
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
    state.lastSaved = new Date().toLocaleTimeString();
    // Persist to document settings so the task pane shows the correct time
    // even after it has been closed and reopened (fresh window context).
    try {
      Office.context.document.settings.set("__as_lastSaved", state.lastSaved);
      Office.context.document.settings.set("__as_lastError", "");
      Office.context.document.settings.saveAsync(() => {});
    } catch { /* non-critical */ }
    clearNotification();
    setText("last-saved-value", state.lastSaved);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    log(`Save failed: ${message}`);
    // Persist error so it surfaces in the task pane next time it opens,
    // making silent background failures visible.
    try {
      Office.context.document.settings.set("__as_lastError", message);
      Office.context.document.settings.saveAsync(() => {});
    } catch { /* non-critical */ }
    setNotification(t("save_error"), "error");
  }
}

// ── Timer ─────────────────────────────────────────────────────────────────────

function startTimer(): void {
  stopTimer();
  if (state.isReadOnly) return;
  log(`Timer started — interval: ${state.intervalMinutes} min`);
  state.timerHandle = setInterval(() => {
    performSave().catch((err) => log(`Unhandled save error: ${err}`));
  }, state.intervalMinutes * 60 * 1000);
}

function stopTimer(): void {
  if (state.timerHandle !== null) {
    clearInterval(state.timerHandle);
    state.timerHandle = null;
  }
}

// ── Background initialization (idempotent) ────────────────────────────────────
// Safe to call from both onDocumentOpen and Office.onReady. The initialized
// guard means the timer is only ever started once per runtime lifetime,
// regardless of how many times the task pane HTML is reloaded.

async function initBackground(): Promise<void> {
  if (state.initialized) {
    log("Runtime already initialized — skipping re-init");
    return;
  }

  const cfg = await loadConfig();
  state.config = cfg;

  await initI18n(Office.context.displayLanguage ?? "en");

  // Use document.mode for read-only detection. getFilePropertiesAsync returns
  // a failed result for any unsaved document (no URL), which is NOT the same
  // as protected/read-only — that bug caused the add-in to disable itself on
  // brand-new documents.
  try {
    state.isReadOnly = Office.context.document.mode === Office.DocumentMode.ReadOnly;
  } catch {
    state.isReadOnly = false;
  }

  state.intervalMinutes = loadUserInterval(cfg);
  state.enabled = true;

  // Restore last-saved time from settings — survives task pane reloads and
  // makes background saves visible when the task pane is opened later.
  try {
    const persisted = Office.context.document.settings.get("__as_lastSaved") as string | null;
    if (persisted) state.lastSaved = persisted;
    const bgError = Office.context.document.settings.get("__as_lastError") as string | null;
    if (bgError) log(`Background error from previous session: ${bgError}`);
  } catch { /* non-critical */ }

  state.initialized = true;

  if (!state.isReadOnly) {
    startTimer();
  }

  log(`Background ready — interval: ${state.intervalMinutes}m, readOnly: ${state.isReadOnly}`);
}

// ── UI ────────────────────────────────────────────────────────────────────────

function applyStrings(): void {
  setText("addin-title", t("addin_name"));
  setText("toggle-label", state.enabled ? t("toggle_on") : t("toggle_off"));
  setText("last-saved-label", t("last_saved"));
  setText("last-saved-value", state.lastSaved || t("never_saved"));
  setText("interval-label-text", t("interval_label"));
  setText("interval-unit-text", t("interval_unit"));
  setText("save-settings-btn", t("save_settings"));
  setText("version-text", VERSION);
  el<HTMLInputElement>("toggle-checkbox").setAttribute("aria-label", t("toggle_on"));
}

function renderSettingsPanel(): void {
  const cfg = state.config!;
  if (cfg.allowUserOverride) {
    show("settings-panel");
    hide("managed-by-it-msg");
    const slider = el<HTMLInputElement>("interval-slider");
    const numberInput = el<HTMLInputElement>("interval-number");
    slider.min = String(cfg.minIntervalMinutes);
    slider.max = String(cfg.maxIntervalMinutes);
    slider.value = String(state.intervalMinutes);
    numberInput.min = String(cfg.minIntervalMinutes);
    numberInput.max = String(cfg.maxIntervalMinutes);
    numberInput.value = String(state.intervalMinutes);
  } else {
    hide("settings-panel");
    show("managed-by-it-msg");
    setText("managed-by-it-text", t("managed_by_it"));
  }
}

function updateToggleUI(): void {
  const checkbox = el<HTMLInputElement>("toggle-checkbox");
  checkbox.checked = state.enabled;
  setText("toggle-label", state.enabled ? t("toggle_on") : t("toggle_off"));
  el<HTMLDivElement>("toggle-row").classList.toggle("toggle-row--off", !state.enabled);
}

function wireEvents(): void {
  const cfg = state.config!;

  const checkbox = el<HTMLInputElement>("toggle-checkbox");
  checkbox.addEventListener("change", () => {
    state.enabled = checkbox.checked;
    updateToggleUI();
    if (state.enabled) startTimer(); else stopTimer();
  });

  const slider = el<HTMLInputElement>("interval-slider");
  const numberInput = el<HTMLInputElement>("interval-number");

  slider.addEventListener("input", () => { numberInput.value = slider.value; });

  numberInput.addEventListener("input", () => {
    const val = parseInt(numberInput.value, 10);
    if (!isNaN(val)) {
      slider.value = String(Math.min(Math.max(val, cfg.minIntervalMinutes), cfg.maxIntervalMinutes));
    }
  });

  el<HTMLButtonElement>("save-settings-btn").addEventListener("click", () => {
    const raw = parseInt(numberInput.value, 10);
    if (isNaN(raw)) return;
    state.intervalMinutes = Math.min(Math.max(raw, cfg.minIntervalMinutes), cfg.maxIntervalMinutes);
    slider.value = String(state.intervalMinutes);
    numberInput.value = String(state.intervalMinutes);
    saveUserInterval(state.intervalMinutes);
    if (state.enabled) startTimer();
  });
}

// ── Event-based activation ────────────────────────────────────────────────────
// Registered synchronously at module load so Office can find the function
// as soon as the shared runtime loads for OnDocumentOpen.

// eslint-disable-next-line @typescript-eslint/no-explicit-any
(Office as any).actions?.associate("onDocumentOpen", (event: any) => {
  // Complete the event immediately — with lifetime="long" the runtime stays
  // alive after event.completed(), so we can do async work afterwards.
  // Awaiting initBackground() before completing caused handler timeouts.
  event.completed();
  initBackground().catch((err) => log(`onDocumentOpen init error: ${err}`));
});

// ── Initialization ────────────────────────────────────────────────────────────

Office.onReady(async () => {
  // initBackground is idempotent — if onDocumentOpen already ran it, this
  // call returns immediately. If the task pane is opened directly (no prior
  // background load), this performs the full init.
  await initBackground();

  // Re-render UI from current state each time the task pane is shown.
  applyStrings();
  renderSettingsPanel();
  wireEvents();
  updateToggleUI();

  if (state.isReadOnly) {
    setNotification(t("protected_view"), "warning");
  }

  try {
    await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
  } catch {
    // Not available in all hosts/versions — safe to ignore
  }

  log(`Task pane ready — timer running: ${state.timerHandle !== null}`);
});
