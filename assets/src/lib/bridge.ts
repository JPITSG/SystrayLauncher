export interface ConfigData {
  url: string;
  windowTitle: string;
  onHideJs: string;
  onShowJs: string;
}

export interface InitData {
  config: ConfigData;
}

type InitCallback = (data: InitData) => void;

let initCallback: InitCallback | null = null;

export function onInit(cb: InitCallback) {
  initCallback = cb;
}

// Called by C via ExecuteScript
(window as unknown as Record<string, unknown>).onInit = (data: InitData) => {
  if (initCallback) initCallback(data);
};

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
    })
  );
}

export function closeDialog() {
  window.chrome.webview.postMessage(JSON.stringify({ action: "close" }));
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
