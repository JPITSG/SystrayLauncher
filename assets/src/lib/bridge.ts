export interface ConfigData {
  url: string;
  windowTitle: string;
  onHideJs: string;
  onShowJs: string;
  sleepWhenInactive: boolean;
  openNewWindowsExternally: boolean;
  allowRunningInsecureContent: boolean;
  insecureContentOrigins: string;
  useStaticHostMappings: boolean;
  staticHostMappings: string;
  lockdownHeader: boolean;
  lockdownSecret: string;
  autoCheckForUpdates: boolean;
  debugLog: boolean;
}

export interface InitData {
  config: ConfigData;
  webView2Version: string;
  updateCompletedVersion: string;
}

export interface UpdateResult {
  status:
    | "newer"
    | "same"
    | "older"
    | "cancelled"
    | "error"
    | "completed";
  title: string;
  message: string;
  currentVersion: string;
  remoteVersion: string;
}

export interface UpdateProgress {
  kilobytesPerSecond: number;
}

type InitCallback = (data: InitData) => void;

let initCallback: InitCallback | null = null;
let updateResultCallback: ((result: UpdateResult) => void) | null = null;
let updateProgressCallback: ((progress: UpdateProgress) => void) | null = null;

export function onInit(cb: InitCallback) {
  initCallback = cb;
}

// Called by C via ExecuteScript
(window as unknown as Record<string, unknown>).onInit = (data: InitData) => {
  if (initCallback) initCallback(data);
};

(window as unknown as Record<string, unknown>).onUpdateResult = (
  result: UpdateResult
) => {
  if (updateResultCallback) updateResultCallback(result);
};

(window as unknown as Record<string, unknown>).onUpdateProgress = (
  progress: UpdateProgress
) => {
  if (updateProgressCallback) updateProgressCallback(progress);
};

export function onUpdateResult(cb: (result: UpdateResult) => void) {
  updateResultCallback = cb;
  return () => {
    if (updateResultCallback === cb) updateResultCallback = null;
  };
}

export function onUpdateProgress(cb: (progress: UpdateProgress) => void) {
  updateProgressCallback = cb;
  return () => {
    if (updateProgressCallback === cb) updateProgressCallback = null;
  };
}

export function getInit() {
  window.chrome.webview.postMessage(JSON.stringify({ action: "getInit" }));
}

export function saveSettings(config: ConfigData) {
  window.chrome.webview.postMessage(
    JSON.stringify({
      action: "saveSettings",
      url: config.url,
      windowTitle: config.windowTitle,
      onHideJs: config.onHideJs,
      onShowJs: config.onShowJs,
      sleepWhenInactive: config.sleepWhenInactive,
      openNewWindowsExternally: config.openNewWindowsExternally,
      allowRunningInsecureContent: config.allowRunningInsecureContent,
      insecureContentOrigins: config.insecureContentOrigins,
      useStaticHostMappings: config.useStaticHostMappings,
      staticHostMappings: config.staticHostMappings,
      lockdownHeader: config.lockdownHeader,
      lockdownSecret: config.lockdownSecret,
      autoCheckForUpdates: config.autoCheckForUpdates,
      debugLog: config.debugLog,
    })
  );
}

export function closeDialog() {
  window.chrome.webview.postMessage(JSON.stringify({ action: "close" }));
}

export function checkForUpdate() {
  window.chrome.webview.postMessage(JSON.stringify({ action: "checkUpdate" }));
}

export function cancelUpdateCheck() {
  window.chrome.webview.postMessage(
    JSON.stringify({ action: "cancelUpdateCheck" })
  );
}

export function installUpdate() {
  window.chrome.webview.postMessage(
    JSON.stringify({ action: "installUpdate" })
  );
}

export function dismissUpdate() {
  window.chrome.webview.postMessage(
    JSON.stringify({ action: "dismissUpdate" })
  );
}

export function dismissUpdateConfirmation() {
  window.chrome.webview.postMessage(
    JSON.stringify({ action: "dismissUpdateConfirmation" })
  );
}

export function reportHeight(height: number) {
  window.chrome.webview.postMessage(
    JSON.stringify({ action: "resize", height })
  );
}

declare global {
  interface Window {
    chrome: {
      webview: {
        postMessage(message: string): void;
      };
    };
  }
}
