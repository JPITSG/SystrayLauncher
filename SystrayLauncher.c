// WebView2 needs Windows 10; declare that so Vista+ APIs (GetTickCount64,
// DWM attributes) are visible in the headers.
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0A00
#endif
#ifndef WINVER
#define WINVER 0x0A00
#endif

#include <windows.h>
#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include <wchar.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <dwmapi.h>
#include <math.h>
#include <assert.h>

#ifndef DWMWA_CLOAKED
#define DWMWA_CLOAKED 14
#endif

// Not present in older MinGW headers.
#ifndef WM_DPICHANGED
#define WM_DPICHANGED 0x02E0
#endif

// WebView2 headers required from SDK
#include "WebView2.h"
#include "resource.h"

#define WINDOW_SIZE_PERCENTAGE 0.9
#define RESOLUTION_CHANGE_DEBOUNCE_MS 1000
#define CONFIG_FILENAME L"config.ini"
#define APP_NAME L"SystrayLauncher"
#define MUTEX_NAME L"SystrayLauncher_SingleInstance_Mutex_9F8A7B6C"
#define TRAY_ICON_ID 100
#define WM_TRAYICON (WM_APP + 1)
#define ID_TRAY_MENU_REFRESH 1
#define ID_TRAY_MENU_CLEAR_CACHE 2
#define ID_TRAY_MENU_OPEN 3
#define ID_TRAY_MENU_CONFIGURE 5
#define ID_TRAY_MENU_EXIT 4
#define ID_TRAY_MENU_RESTART 6

// Registry settings
#define REG_COMPANY L"JPIT"
#define REG_APPNAME L"SystrayLauncher"
#define REG_KEY_PATH L"SOFTWARE\\JPIT\\SystrayLauncher"
#define REG_VALUE_URL L"URL"
#define REG_VALUE_TITLE L"WindowTitle"
#define REG_VALUE_ONHIDEJS L"OnHideJS"
#define REG_VALUE_ONSHOWJS L"OnShowJS"
#define REG_VALUE_SLEEP L"SleepWhenInactive"
#define REG_VALUE_SPELLCHECK L"SpellcheckLanguages"
#define REG_VALUE_NEWWINDOW L"OpenNewWindowsExternally"
#define REG_VALUE_CONFIGURED L"Configured"

#define ID_TIMER_INITIAL_HIDE_JS 2
#define INITIAL_HIDE_JS_DELAY_MS 2000
#define ID_TIMER_VISIBILITY_CHECK 3
#define VISIBILITY_CHECK_INTERVAL_MS 250
#define ID_TIMER_CFG_SHOW_FALLBACK 4
#define CFG_SHOW_FALLBACK_DELAY_MS 350
#define ID_TIMER_WEBVIEW_PREWARM 5
#define WEBVIEW_PREWARM_MS 60000
#define ID_TIMER_WEBVIEW_PRELOAD 6
#define WEBVIEW_PRELOAD_SETTLE_MS 1500
#define ID_TIMER_WEBVIEW_RECREATE 7
#define WEBVIEW_RECREATE_FALLBACK_MS 10000
#define ID_TIMER_POWER_RESUME 8
#define POWER_RESUME_KICK_DELAY_MS 2000
#define ID_TIMER_WEBVIEW_LIVENESS 9
// Each composition kick after a power resume is verified with a script ping;
// if the runtime does not answer within this window the kick is retried (the
// graphics stack can lag badly after hibernate), and after
// POWER_RESUME_MAX_KICKS failed attempts the WebView is rebuilt instead.
#define POWER_RESUME_LIVENESS_MS 3000
#define POWER_RESUME_KICK_RETRY_MS 4000
#define POWER_RESUME_MAX_KICKS 3

// After a suspend-resume failure the resume is retried on every activation
// tick (4x/s); if it keeps failing this long the runtime is torn down and
// rebuilt instead.
#define RESUME_FAILURE_RECREATE_THRESHOLD 12

// Rate limit for automatic WebView rebuilds after unexpected browser-process
// deaths, so a crash-looping runtime cannot spin rebuilds forever. A manual
// tray Refresh/Open resets the limiter.
#define REBUILD_BURST_WINDOW_MS (5 * 60 * 1000)
#define REBUILD_BURST_MAX 5

// The config dialog is normally shown by its first resize message; the
// fallback timer keeps waiting while WebView2 is still initializing and
// gives up (with an error) after this many 350 ms ticks.
#define CFG_SHOW_FALLBACK_MAX_TRIES 20

// Posted to the main window after the config dialog saves changed spell-check
// languages (asks the user to restart the WebView), and when the browser
// process has exited and the WebView can be rebuilt with the new languages.
#define WM_APP_SPELLCHECK_CHANGED (WM_APP + 2)
#define WM_APP_WEBVIEW_RECREATE (WM_APP + 3)

typedef struct {
    wchar_t url[2048];
    wchar_t windowTitle[256];
    wchar_t onHideJs[4096];
    wchar_t onShowJs[4096];
    BOOL sleepWhenInactive;
    BOOL openNewWindowsExternally;
    // Comma-separated BCP-47 tags (e.g. L"en-US,pl"); empty = don't manage
    // the WebView2 profile's spell-check settings at all.
    wchar_t spellcheckLanguages[512];
} Configuration;

typedef enum {
    JS_VISIBILITY_UNKNOWN = -1,
    JS_VISIBILITY_HIDDEN = 0,
    JS_VISIBILITY_SHOWN = 1
} JsVisibility;

// Globals
static Configuration g_config;
static HWND g_hwnd = NULL;
static HWND g_hwndOwner = NULL;  // Invisible owner window to prevent taskbar appearance
static ICoreWebView2Controller* g_webViewController = NULL;
static ICoreWebView2* g_webView = NULL;
static ICoreWebView2Environment* g_webViewEnv = NULL;
static NOTIFYICONDATAW g_nid = {0};
static UINT g_WM_TASKBARCREATED = 0;
static HANDLE g_hMutex = NULL;
static wchar_t g_iniPath[MAX_PATH];
static wchar_t g_initialUrl[2048];
static UINT_PTR g_timerId = 0;
static int g_lastScreenWidth = 0;
static int g_lastScreenHeight = 0;
static float g_lastDpiX = 0.0f;
static float g_lastDpiY = 0.0f;
static volatile LONG g_isInitialized = FALSE;
static volatile LONG g_webViewDesiredActive = FALSE;
static volatile LONG g_webViewDesiredVisible = FALSE;
static volatile LONG g_webViewPrewarmActive = FALSE;
static volatile LONG g_webViewSuspendPending = FALSE;
static volatile LONG g_webViewSuspended = FALSE;
static volatile LONG g_resetUrlOnNextShow = FALSE;
static volatile LONG g_sleepWhenInactive = FALSE;
static volatile LONG g_openNewWindowsExternally = FALSE;
static volatile LONG g_initialPreloadComplete = FALSE;
static volatile LONG g_webViewRecreatePending = FALSE;
static volatile LONG g_webViewCreatePending = FALSE;
static volatile LONG g_resumeFailureCount = 0;
// Post-power-resume recovery: set when the machine goes down or comes back up
// and cleared once the WebView has answered a liveness ping (or been rebuilt).
// While set, the page is kept warm so we never snapshot an unverified page.
static volatile LONG g_powerResumePending = FALSE;
static volatile LONG g_webViewPingOutstanding = FALSE;
static int g_powerKickCount = 0;
static ULONGLONG g_rebuildBurstStartTick = 0;
static LONG g_rebuildBurstCount = 0;
static EventRegistrationToken g_browserExitedToken;
static BOOL g_browserExitedRegistered = FALSE;
static HINSTANCE g_hInstance;
static JsVisibility g_jsVisibility = JS_VISIBILITY_UNKNOWN;
static wchar_t g_webView2Version[128] = L"Unknown";

// Config dialog WebView2 globals
static HWND g_cfgHwnd = NULL;
static ICoreWebView2Environment* g_cfgEnv = NULL;
static ICoreWebView2Controller* g_cfgController = NULL;
static ICoreWebView2* g_cfgWebView = NULL;
static BOOL g_cfgSaved = FALSE;
static BOOL g_cfgWindowShown = FALSE;
static int g_cfgShowFallbackTries = 0;

// Dynamic WebView2 loading
static WCHAR g_extractedDllPath[MAX_PATH] = {0};
typedef HRESULT (STDAPICALLTYPE *PFN_CreateCoreWebView2EnvironmentWithOptions)(
    LPCWSTR, LPCWSTR, void*,
    ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*);
static PFN_CreateCoreWebView2EnvironmentWithOptions fnCreateEnvironment = NULL;

// Forward declarations
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
void LoadConfiguration(const wchar_t* iniPath, Configuration* config);
void CreateDefaultIni(const wchar_t* iniPath);
void ParseConfigLine(wchar_t* line, Configuration* config);
void ShowMainWindow(void);
void HideMainWindow(void);
void CreateTrayIcon(HWND hwnd);
void ShowContextMenu(HWND hwnd);
void RefreshTrayIcon(void);
void CaptureDisplaySettings(void);
BOOL HasDisplaySettingsChanged(void);
void DebugPrint(const wchar_t* format, ...);
static void NormalizeSpellcheckLanguages(const wchar_t* in, wchar_t* out, size_t outLen);
static void PatchSpellcheckPreferences(void);
static void GetMainUserDataFolder(wchar_t path[MAX_PATH]);
static void CreateMainWebViewEnvironment(HWND hwnd);
static void BeginMainWebViewRecreate(void);
static void FinishMainWebViewRecreate(HWND hwnd);
static void HandleUnexpectedBrowserExit(HWND hwnd);
static void RebuildMainWebViewIfDead(void);
static void KickWebViewAfterPowerResume(HWND hwnd);
static void SendMainWebViewLivenessPing(void);
static void CheckMainWebViewLiveness(HWND hwnd);
static void RestartApplication(void);
static void RegisterBrowserExitedOnCurrentEnv(void);
static void UnregisterBrowserExitedFromCurrentEnv(void);
static void RegisterMainProcessFailedHandler(ICoreWebView2* webview2);
void ReloadTargetPage(void);
void ClearWebViewCacheAndReload(void);
void ExecuteJavaScript(const wchar_t* js);
static BOOL IsWebViewReady(void);
static BOOL IsWindowActuallyVisible(HWND hwnd);
static void UpdateJsVisibilityState(HWND hwnd);
static void StartVisibilityTimer(HWND hwnd);
static void StopVisibilityTimer(HWND hwnd);
static void ActivateMainWebView(void);
static void DeactivateMainWebView(void);
static void ResumeMainWebViewRuntime(void);
static void SetMainWebViewControllerVisible(BOOL visible);
static void PrewarmMainWebView(void);
static void ResetTargetPageIfNeeded(void);
static void OnMainNavigationCompleted(void);
static void RegisterMainNavigationCompletedHandler(ICoreWebView2* webview2);
static void RegisterMainNewWindowRequestedHandler(ICoreWebView2* webview2);
static void GetTargetWindowRect(int* x, int* y, int* w, int* h);

// Registry and config dialog functions
static BOOL LoadConfigFromRegistry(Configuration* config);
static BOOL SaveConfigToRegistry(const Configuration* config);
static BOOL IsFirstLaunch(void);
static void MarkAsConfigured(void);
static void ApplyConfiguration(void);
static BOOL load_webview2_loader(void);
static void ShowConfigWebViewDialog(void);

// WebView2 Callbacks
HRESULT STDMETHODCALLTYPE EnvCompletedHandler_QueryInterface(
    ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler* This,
    REFIID riid, void** ppvObject);
ULONG STDMETHODCALLTYPE EnvCompletedHandler_AddRef(
    ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler* This);
ULONG STDMETHODCALLTYPE EnvCompletedHandler_Release(
    ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler* This);
HRESULT STDMETHODCALLTYPE EnvCompletedHandler_Invoke(
    ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler* This,
    HRESULT result, ICoreWebView2Environment* environment);

typedef struct {
    ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandlerVtbl* lpVtbl;
    LONG refCount;
    HWND hwnd;
    wchar_t* userDataPath;
} EnvCompletedHandler;

HRESULT STDMETHODCALLTYPE ControllerCompletedHandler_QueryInterface(
    ICoreWebView2CreateCoreWebView2ControllerCompletedHandler* This,
    REFIID riid, void** ppvObject);
ULONG STDMETHODCALLTYPE ControllerCompletedHandler_AddRef(
    ICoreWebView2CreateCoreWebView2ControllerCompletedHandler* This);
ULONG STDMETHODCALLTYPE ControllerCompletedHandler_Release(
    ICoreWebView2CreateCoreWebView2ControllerCompletedHandler* This);
HRESULT STDMETHODCALLTYPE ControllerCompletedHandler_Invoke(
    ICoreWebView2CreateCoreWebView2ControllerCompletedHandler* This,
    HRESULT result, ICoreWebView2Controller* controller);

typedef struct {
    ICoreWebView2CreateCoreWebView2ControllerCompletedHandlerVtbl* lpVtbl;
    LONG refCount;
    HWND hwnd;
} ControllerCompletedHandler;

HRESULT STDMETHODCALLTYPE ClearBrowsingDataCompletedHandler_QueryInterface(
    ICoreWebView2ClearBrowsingDataCompletedHandler* This,
    REFIID riid, void** ppvObject);
ULONG STDMETHODCALLTYPE ClearBrowsingDataCompletedHandler_AddRef(
    ICoreWebView2ClearBrowsingDataCompletedHandler* This);
ULONG STDMETHODCALLTYPE ClearBrowsingDataCompletedHandler_Release(
    ICoreWebView2ClearBrowsingDataCompletedHandler* This);
HRESULT STDMETHODCALLTYPE ClearBrowsingDataCompletedHandler_Invoke(
    ICoreWebView2ClearBrowsingDataCompletedHandler* This,
    HRESULT errorCode);

typedef struct {
    ICoreWebView2ClearBrowsingDataCompletedHandlerVtbl* lpVtbl;
    LONG refCount;
} ClearBrowsingDataCompletedHandler;

// ExecuteScript completion handler (fire-and-forget)
HRESULT STDMETHODCALLTYPE ExecuteScriptCompletedHandler_QueryInterface(
    ICoreWebView2ExecuteScriptCompletedHandler* This,
    REFIID riid, void** ppvObject);
ULONG STDMETHODCALLTYPE ExecuteScriptCompletedHandler_AddRef(
    ICoreWebView2ExecuteScriptCompletedHandler* This);
ULONG STDMETHODCALLTYPE ExecuteScriptCompletedHandler_Release(
    ICoreWebView2ExecuteScriptCompletedHandler* This);
HRESULT STDMETHODCALLTYPE ExecuteScriptCompletedHandler_Invoke(
    ICoreWebView2ExecuteScriptCompletedHandler* This,
    HRESULT errorCode, LPCWSTR resultObjectAsJson);

typedef struct {
    ICoreWebView2ExecuteScriptCompletedHandlerVtbl* lpVtbl;
    LONG refCount;
} ExecuteScriptCompletedHandler;

typedef struct {
    ICoreWebView2ExecuteScriptCompletedHandlerVtbl* lpVtbl;
    LONG refCount;
} LivenessPingHandler;

// WebView suspend completion handler
HRESULT STDMETHODCALLTYPE TrySuspendCompletedHandler_QueryInterface(
    ICoreWebView2TrySuspendCompletedHandler* This,
    REFIID riid, void** ppvObject);
ULONG STDMETHODCALLTYPE TrySuspendCompletedHandler_AddRef(
    ICoreWebView2TrySuspendCompletedHandler* This);
ULONG STDMETHODCALLTYPE TrySuspendCompletedHandler_Release(
    ICoreWebView2TrySuspendCompletedHandler* This);
HRESULT STDMETHODCALLTYPE TrySuspendCompletedHandler_Invoke(
    ICoreWebView2TrySuspendCompletedHandler* This,
    HRESULT errorCode, BOOL result);

typedef struct {
    ICoreWebView2TrySuspendCompletedHandlerVtbl* lpVtbl;
    LONG refCount;
} TrySuspendCompletedHandler;

// Navigation completed handler (settles the initial preload / sleep state)
HRESULT STDMETHODCALLTYPE NavCompletedHandler_QueryInterface(
    ICoreWebView2NavigationCompletedEventHandler* This,
    REFIID riid, void** ppvObject);
ULONG STDMETHODCALLTYPE NavCompletedHandler_AddRef(
    ICoreWebView2NavigationCompletedEventHandler* This);
ULONG STDMETHODCALLTYPE NavCompletedHandler_Release(
    ICoreWebView2NavigationCompletedEventHandler* This);
HRESULT STDMETHODCALLTYPE NavCompletedHandler_Invoke(
    ICoreWebView2NavigationCompletedEventHandler* This,
    ICoreWebView2* sender, ICoreWebView2NavigationCompletedEventArgs* args);

typedef struct {
    ICoreWebView2NavigationCompletedEventHandlerVtbl* lpVtbl;
    LONG refCount;
} NavCompletedHandler;

// Configuration functions
void LoadConfiguration(const wchar_t* iniPath, Configuration* config) {
    wcscpy_s(config->url, 2048, L"https://www.google.com/");
    wcscpy_s(config->windowTitle, 256, L"Systray Launcher");
    config->onHideJs[0] = L'\0';
    config->onShowJs[0] = L'\0';
    config->sleepWhenInactive = FALSE;
    config->openNewWindowsExternally = FALSE;
    config->spellcheckLanguages[0] = L'\0';

    if (!PathFileExistsW(iniPath)) {
        CreateDefaultIni(iniPath);
        return;
    }

    FILE* file = NULL;
    if (_wfopen_s(&file, iniPath, L"r, ccs=UTF-8") != 0 || !file) {
        // Fallback: try ANSI
        char pathA[MAX_PATH];
        wcstombs(pathA, iniPath, MAX_PATH);
        FILE* fileA = NULL;
        if (fopen_s(&fileA, pathA, "r") == 0 && fileA) {
            char lineA[4096];
            while (fgets(lineA, sizeof(lineA), fileA)) {
                wchar_t lineW[4096];
                MultiByteToWideChar(CP_ACP, 0, lineA, -1, lineW, 4096);
                ParseConfigLine(lineW, config);
            }
            fclose(fileA);
        }
        return;
    }

    wchar_t line[4096];
    while (fgetws(line, sizeof(line)/sizeof(wchar_t), file)) {
        ParseConfigLine(line, config);
    }
    fclose(file);
}

void ParseConfigLine(wchar_t* line, Configuration* config) {
    while (iswspace(*line)) line++;
    if (line[0] == L'#' || line[0] == L';' || line[0] == L'\0' || line[0] == L'\n') return;

    wchar_t* equals = wcschr(line, L'=');
    if (!equals) return;
    *equals = L'\0';

    wchar_t* key = line;
    wchar_t* value = equals + 1;

    // Trim key
    while (iswspace(*key)) key++;
    size_t keyLen = wcslen(key);
    while (keyLen > 0 && iswspace(key[keyLen - 1])) key[--keyLen] = L'\0';
    for (wchar_t* p = key; *p; ++p) *p = towlower(*p);

    // Trim value
    while (iswspace(*value)) value++;
    size_t valLen = wcslen(value);
    while (valLen > 0 && iswspace(value[valLen - 1])) value[--valLen] = L'\0';
    if (valLen > 0 && value[valLen-1] == L'\n') value[--valLen] = L'\0';
    if (valLen > 0 && value[valLen-1] == L'\r') value[--valLen] = L'\0';

    if (wcscmp(key, L"url") == 0) {
        wcscpy_s(config->url, 2048, value);
    } else if (wcscmp(key, L"windowtitle") == 0) {
        wcscpy_s(config->windowTitle, 256, value);
    } else if (wcscmp(key, L"onhidejs") == 0) {
        wcscpy_s(config->onHideJs, 4096, value);
    } else if (wcscmp(key, L"onshowjs") == 0) {
        wcscpy_s(config->onShowJs, 4096, value);
    } else if (wcscmp(key, L"sleepwheninactive") == 0) {
        wchar_t c = towlower(value[0]);
        config->sleepWhenInactive = (c == L'1' || c == L't' || c == L'y');
    } else if (wcscmp(key, L"spellchecklanguages") == 0) {
        wcscpy_s(config->spellcheckLanguages, 512, value);
    } else if (wcscmp(key, L"opennewwindowsexternally") == 0) {
        wchar_t c = towlower(value[0]);
        config->openNewWindowsExternally = (c == L'1' || c == L't' || c == L'y');
    }
}

void CreateDefaultIni(const wchar_t* iniPath) {
    const wchar_t* content = L"# SystrayLauncher Configuration File\n"
                             L"# Lines starting with # or ; are comments\n\n"
                             L"url=https://www.google.com/\n\n"
                             L"windowtitle=Systray Launcher\n";
    FILE* file = NULL;
    _wfopen_s(&file, iniPath, L"w, ccs=UTF-8");
    if (file) {
        fputws(content, file);
        fclose(file);
    }
}

// Registry functions
static BOOL LoadConfigFromRegistry(Configuration* config) {
    HKEY hKey;
    LONG result = RegOpenKeyExW(HKEY_CURRENT_USER, REG_KEY_PATH, 0, KEY_READ, &hKey);
    if (result != ERROR_SUCCESS) {
        return FALSE;
    }

    DWORD dataSize;
    DWORD dataType;

    // Load URL
    dataSize = sizeof(config->url);
    if (RegQueryValueExW(hKey, REG_VALUE_URL, NULL, &dataType, (LPBYTE)config->url, &dataSize) != ERROR_SUCCESS) {
        wcscpy_s(config->url, 2048, L"https://www.google.com/");
    }

    // Load Window Title
    dataSize = sizeof(config->windowTitle);
    if (RegQueryValueExW(hKey, REG_VALUE_TITLE, NULL, &dataType, (LPBYTE)config->windowTitle, &dataSize) != ERROR_SUCCESS) {
        wcscpy_s(config->windowTitle, 256, L"Systray Launcher");
    }

    // Load OnHideJS
    dataSize = sizeof(config->onHideJs);
    if (RegQueryValueExW(hKey, REG_VALUE_ONHIDEJS, NULL, &dataType, (LPBYTE)config->onHideJs, &dataSize) != ERROR_SUCCESS) {
        config->onHideJs[0] = L'\0';
    }

    // Load OnShowJS
    dataSize = sizeof(config->onShowJs);
    if (RegQueryValueExW(hKey, REG_VALUE_ONSHOWJS, NULL, &dataType, (LPBYTE)config->onShowJs, &dataSize) != ERROR_SUCCESS) {
        config->onShowJs[0] = L'\0';
    }

    // Load SleepWhenInactive (default disabled)
    DWORD sleepVal = 0;
    dataSize = sizeof(sleepVal);
    if (RegQueryValueExW(hKey, REG_VALUE_SLEEP, NULL, &dataType, (LPBYTE)&sleepVal, &dataSize) == ERROR_SUCCESS) {
        config->sleepWhenInactive = (sleepVal != 0);
    } else {
        config->sleepWhenInactive = FALSE;
    }

    // Load OpenNewWindowsExternally (default disabled)
    DWORD newWinVal = 0;
    dataSize = sizeof(newWinVal);
    if (RegQueryValueExW(hKey, REG_VALUE_NEWWINDOW, NULL, &dataType, (LPBYTE)&newWinVal, &dataSize) == ERROR_SUCCESS) {
        config->openNewWindowsExternally = (newWinVal != 0);
    } else {
        config->openNewWindowsExternally = FALSE;
    }

    // Load SpellcheckLanguages (empty = spell-check settings not managed)
    dataSize = sizeof(config->spellcheckLanguages);
    if (RegQueryValueExW(hKey, REG_VALUE_SPELLCHECK, NULL, &dataType, (LPBYTE)config->spellcheckLanguages, &dataSize) != ERROR_SUCCESS) {
        config->spellcheckLanguages[0] = L'\0';
    } else {
        // REG_SZ data is not guaranteed to be null-terminated
        config->spellcheckLanguages[511] = L'\0';
    }

    RegCloseKey(hKey);
    return TRUE;
}

static BOOL SaveConfigToRegistry(const Configuration* config) {
    HKEY hKey;
    DWORD disposition;
    LONG result = RegCreateKeyExW(HKEY_CURRENT_USER, REG_KEY_PATH, 0, NULL,
                                   REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, &disposition);
    if (result != ERROR_SUCCESS) {
        return FALSE;
    }

    // Save URL
    RegSetValueExW(hKey, REG_VALUE_URL, 0, REG_SZ,
                   (const BYTE*)config->url, (DWORD)((wcslen(config->url) + 1) * sizeof(wchar_t)));

    // Save Window Title
    RegSetValueExW(hKey, REG_VALUE_TITLE, 0, REG_SZ,
                   (const BYTE*)config->windowTitle, (DWORD)((wcslen(config->windowTitle) + 1) * sizeof(wchar_t)));

    // Save OnHideJS
    RegSetValueExW(hKey, REG_VALUE_ONHIDEJS, 0, REG_SZ,
                   (const BYTE*)config->onHideJs, (DWORD)((wcslen(config->onHideJs) + 1) * sizeof(wchar_t)));

    // Save OnShowJS
    RegSetValueExW(hKey, REG_VALUE_ONSHOWJS, 0, REG_SZ,
                   (const BYTE*)config->onShowJs, (DWORD)((wcslen(config->onShowJs) + 1) * sizeof(wchar_t)));

    // Save SleepWhenInactive
    DWORD sleepVal = config->sleepWhenInactive ? 1 : 0;
    RegSetValueExW(hKey, REG_VALUE_SLEEP, 0, REG_DWORD,
                   (const BYTE*)&sleepVal, sizeof(sleepVal));

    // Save OpenNewWindowsExternally
    DWORD newWinVal = config->openNewWindowsExternally ? 1 : 0;
    RegSetValueExW(hKey, REG_VALUE_NEWWINDOW, 0, REG_DWORD,
                   (const BYTE*)&newWinVal, sizeof(newWinVal));

    // Save SpellcheckLanguages
    RegSetValueExW(hKey, REG_VALUE_SPELLCHECK, 0, REG_SZ,
                   (const BYTE*)config->spellcheckLanguages,
                   (DWORD)((wcslen(config->spellcheckLanguages) + 1) * sizeof(wchar_t)));

    RegCloseKey(hKey);
    return TRUE;
}

static BOOL IsFirstLaunch(void) {
    HKEY hKey;
    LONG result = RegOpenKeyExW(HKEY_CURRENT_USER, REG_KEY_PATH, 0, KEY_READ, &hKey);
    if (result != ERROR_SUCCESS) {
        return TRUE;  // Key doesn't exist = first launch
    }

    DWORD configured = 0;
    DWORD dataSize = sizeof(configured);
    result = RegQueryValueExW(hKey, REG_VALUE_CONFIGURED, NULL, NULL, (LPBYTE)&configured, &dataSize);
    RegCloseKey(hKey);

    return (result != ERROR_SUCCESS || configured == 0);
}

static void MarkAsConfigured(void) {
    HKEY hKey;
    DWORD disposition;
    LONG result = RegCreateKeyExW(HKEY_CURRENT_USER, REG_KEY_PATH, 0, NULL,
                                   REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, &disposition);
    if (result == ERROR_SUCCESS) {
        DWORD configured = 1;
        RegSetValueExW(hKey, REG_VALUE_CONFIGURED, 0, REG_DWORD, (const BYTE*)&configured, sizeof(configured));
        RegCloseKey(hKey);
    }
}

static void ApplyConfiguration(void) {
    // Update initial URL
    wcscpy_s(g_initialUrl, 2048, g_config.url);

    // Sync the "sleep when inactive" setting
    InterlockedExchange(&g_sleepWhenInactive, g_config.sleepWhenInactive ? TRUE : FALSE);

    // Sync the new-window handling setting (read live by the handler, so a
    // toggle applies without restarting the WebView)
    InterlockedExchange(&g_openNewWindowsExternally, g_config.openNewWindowsExternally ? TRUE : FALSE);

    // Update window title
    if (g_hwnd) {
        SetWindowTextW(g_hwnd, g_config.windowTitle);
    }

    // Update tray icon tooltip
    if (g_nid.hWnd) {
        wcscpy_s(g_nid.szTip, sizeof(g_nid.szTip)/sizeof(wchar_t), g_config.windowTitle);
        Shell_NotifyIconW(NIM_MODIFY, &g_nid);
    }

    // Navigate to the new URL only if the main window is visible
    if (g_webView && g_hwnd && IsWindowVisible(g_hwnd)) {
        g_webView->lpVtbl->Navigate(g_webView, g_config.url);
    }

    // Re-apply the correct active/sleep state for the (possibly changed)
    // setting: disabling sleep while hidden wakes the runtime back up;
    // enabling it suspends the already-loaded page.
    if (g_hwnd && IsWebViewReady()) {
        if (IsWindowActuallyVisible(g_hwnd)) {
            ActivateMainWebView();
        } else {
            DeactivateMainWebView();
        }
    }
}

// Dynamic WebView2 loader extraction
static BOOL load_webview2_loader(void) {
    HRSRC hRes = FindResource(NULL, MAKEINTRESOURCE(IDR_WEBVIEW2_DLL), RT_RCDATA);
    if (hRes) {
        HGLOBAL hData = LoadResource(NULL, hRes);
        DWORD dllSize = SizeofResource(NULL, hRes);
        const void *dllBytes = LockResource(hData);
        if (dllBytes && dllSize > 0) {
            WCHAR tempDir[MAX_PATH];
            DWORD tempLen = GetTempPathW(MAX_PATH, tempDir);
            if (tempLen > 0 && tempLen < MAX_PATH - 30) {
                swprintf(g_extractedDllPath, MAX_PATH, L"%sWebView2Loader.dll", tempDir);
                HANDLE hFile = CreateFileW(g_extractedDllPath, GENERIC_WRITE, 0, NULL,
                    CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
                if (hFile != INVALID_HANDLE_VALUE) {
                    DWORD written = 0;
                    WriteFile(hFile, dllBytes, dllSize, &written, NULL);
                    CloseHandle(hFile);
                    if (written == dllSize) {
                        HMODULE hMod = LoadLibraryW(g_extractedDllPath);
                        if (hMod) {
                            fnCreateEnvironment = (PFN_CreateCoreWebView2EnvironmentWithOptions)
                                GetProcAddress(hMod, "CreateCoreWebView2EnvironmentWithOptions");
                            if (fnCreateEnvironment) return TRUE;
                        }
                    }
                }
            }
        }
    }
    return FALSE;
}

// JSON helpers
static BOOL json_get_string(const char *json, const char *key, char *out, size_t outLen) {
    char search[128];
    snprintf(search, sizeof(search), "\"%s\"", key);
    const char *p = strstr(json, search);
    if (!p) return FALSE;
    p += strlen(search);
    while (*p == ' ' || *p == ':') p++;
    if (*p != '"') return FALSE;
    p++;
    size_t i = 0;
    while (*p && i < outLen - 1) {
        if (*p == '"') break;
        if (*p == '\\' && *(p + 1)) {
            p++;
            switch (*p) {
                case '"':  out[i++] = '"';  break;
                case '\\': out[i++] = '\\'; break;
                case 'n':  out[i++] = '\n'; break;
                case 'r':  out[i++] = '\r'; break;
                case 't':  out[i++] = '\t'; break;
                default:   out[i++] = *p;   break;
            }
        } else {
            out[i++] = *p;
        }
        p++;
    }
    out[i] = '\0';
    return TRUE;
}

static BOOL json_get_bool(const char *json, const char *key, BOOL defVal) {
    char search[128];
    snprintf(search, sizeof(search), "\"%s\"", key);
    const char *p = strstr(json, search);
    if (!p) return defVal;
    p += strlen(search);
    while (*p == ' ' || *p == ':') p++;
    if (strncmp(p, "true", 4) == 0) return TRUE;
    if (strncmp(p, "false", 5) == 0) return FALSE;
    if (*p == '1') return TRUE;
    if (*p == '0') return FALSE;
    return defVal;
}

static void json_escape_wstring(const wchar_t *in, wchar_t *out, size_t outLen) {
    size_t j = 0;
    for (size_t i = 0; in[i] && j < outLen - 2; i++) {
        wchar_t c = in[i];
        if (c == L'"' || c == L'\\') {
            if (j + 2 >= outLen) break;
            out[j++] = L'\\';
            out[j++] = c;
        } else if (c == L'\n') {
            if (j + 2 >= outLen) break;
            out[j++] = L'\\';
            out[j++] = L'n';
        } else if (c == L'\r') {
            if (j + 2 >= outLen) break;
            out[j++] = L'\\';
            out[j++] = L'r';
        } else if (c == L'\t') {
            if (j + 2 >= outLen) break;
            out[j++] = L'\\';
            out[j++] = L't';
        } else {
            out[j++] = c;
        }
    }
    out[j] = L'\0';
}

// The process is per-monitor DPI aware (see SystrayLauncher.manifest), so
// window pixels are physical pixels and anything sized from CSS pixels has
// to be scaled by the window's DPI. GetDpiForWindow is resolved dynamically
// (Windows 10 1607+); the GDI metric is the fallback and, for an aware
// process, also returns the real DPI.
typedef UINT (WINAPI *PFN_GetDpiForWindow)(HWND);
static UINT GetWindowDpi(HWND hwnd) {
    static PFN_GetDpiForWindow fnGetDpiForWindow = NULL;
    static BOOL resolved = FALSE;
    if (!resolved) {
        fnGetDpiForWindow = (PFN_GetDpiForWindow)GetProcAddress(
            GetModuleHandleW(L"user32.dll"), "GetDpiForWindow");
        resolved = TRUE;
    }
    if (fnGetDpiForWindow && hwnd) {
        UINT dpi = fnGetDpiForWindow(hwnd);
        if (dpi) return dpi;
    }
    HDC hdc = GetDC(hwnd);
    UINT dpi = (UINT)GetDeviceCaps(hdc, LOGPIXELSX);
    ReleaseDC(hwnd, hdc);
    return dpi ? dpi : 96;
}

// Config dialog WebView2 helpers
static void webview_cfg_execute_script(const wchar_t* script) {
    if (!g_cfgWebView || !script) return;

    ExecuteScriptCompletedHandler* handler =
        (ExecuteScriptCompletedHandler*)calloc(1, sizeof(ExecuteScriptCompletedHandler));
    if (!handler) return;

    static ICoreWebView2ExecuteScriptCompletedHandlerVtbl vtbl = {
        ExecuteScriptCompletedHandler_QueryInterface,
        ExecuteScriptCompletedHandler_AddRef,
        ExecuteScriptCompletedHandler_Release,
        ExecuteScriptCompletedHandler_Invoke
    };
    handler->lpVtbl = &vtbl;
    handler->refCount = 1;

    g_cfgWebView->lpVtbl->ExecuteScript(g_cfgWebView, script,
        (ICoreWebView2ExecuteScriptCompletedHandler*)handler);
    handler->lpVtbl->Release((ICoreWebView2ExecuteScriptCompletedHandler*)handler);
}

static void cfg_sync_controller_bounds(void) {
    if (!g_cfgController || !g_cfgHwnd) return;
    RECT bounds;
    GetClientRect(g_cfgHwnd, &bounds);
    g_cfgController->lpVtbl->put_Bounds(g_cfgController, bounds);
    g_cfgController->lpVtbl->put_IsVisible(g_cfgController, TRUE);
}

static void webview_push_init_config(void) {
    wchar_t eUrl[4096], eTitle[512], eHide[8192], eShow[8192], eSpell[1024];
    json_escape_wstring(g_config.url, eUrl, 4096);
    json_escape_wstring(g_config.windowTitle, eTitle, 512);
    json_escape_wstring(g_config.onHideJs, eHide, 8192);
    json_escape_wstring(g_config.onShowJs, eShow, 8192);
    json_escape_wstring(g_config.spellcheckLanguages, eSpell, 1024);

    wchar_t script[16384];
    swprintf(script, 16384,
        L"window.onInit({\"config\":{\"url\":\"%s\",\"windowTitle\":\"%s\",\"onHideJs\":\"%s\",\"onShowJs\":\"%s\",\"sleepWhenInactive\":%s,\"spellcheckLanguages\":\"%s\",\"openNewWindowsExternally\":%s}})",
        eUrl, eTitle, eHide, eShow, g_config.sleepWhenInactive ? L"true" : L"false", eSpell,
        g_config.openNewWindowsExternally ? L"true" : L"false");
    webview_cfg_execute_script(script);
}

// Minimal COM handler struct for config dialog (shared by all cfg handlers)
typedef struct {
    void* lpVtbl;
    LONG refCount;
} CfgHandler;

// Config dialog COM handlers — simplified pattern
static HRESULT STDMETHODCALLTYPE CfgHandler_QueryInterface(
    ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler *This, REFIID riid, void **ppv) {
    (void)riid;
    *ppv = This;
    ((CfgHandler*)This)->refCount++;
    return S_OK;
}
static ULONG STDMETHODCALLTYPE CfgHandler_AddRef(
    ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler *This) {
    return ++((CfgHandler*)This)->refCount;
}
static ULONG STDMETHODCALLTYPE CfgHandler_Release(
    ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler *This) {
    ULONG rc = --((CfgHandler*)This)->refCount;
    if (rc == 0) free(This);
    return rc;
}

static HRESULT STDMETHODCALLTYPE CfgCtrlCompleted_Invoke(
    ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*, HRESULT, ICoreWebView2Controller*);
static HRESULT STDMETHODCALLTYPE CfgMsgReceived_Invoke(
    ICoreWebView2WebMessageReceivedEventHandler*, ICoreWebView2*, ICoreWebView2WebMessageReceivedEventArgs*);

// A silent environment/controller failure used to leave the fallback timer
// to show a window with nothing inside it - the "blank config modal". Fail
// loudly and take the window down instead.
static void CfgReportInitFailureAndClose(HRESULT hr) {
    if (g_cfgHwnd) {
        KillTimer(g_cfgHwnd, ID_TIMER_CFG_SHOW_FALLBACK);
        PostMessage(g_cfgHwnd, WM_CLOSE, 0, 0);
    }
    wchar_t msg[256];
    swprintf_s(msg, 256,
        L"The configuration window could not initialize WebView2 (0x%08X).\n\n"
        L"Please check the Microsoft Edge WebView2 Runtime installation.",
        (unsigned)hr);
    MessageBoxW(NULL, msg, APP_NAME, MB_ICONERROR | MB_OK);
}

static HRESULT STDMETHODCALLTYPE CfgEnvCompleted_Invoke(
    ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler *This,
    HRESULT result, ICoreWebView2Environment *env) {
    (void)This;
    if (FAILED(result) || !env) {
        CfgReportInitFailureAndClose(FAILED(result) ? result : E_POINTER);
        return S_OK;
    }
    g_cfgEnv = env;
    env->lpVtbl->AddRef(env);

    static ICoreWebView2CreateCoreWebView2ControllerCompletedHandlerVtbl ctrlVtbl = {0};
    static BOOL init = FALSE;
    if (!init) {
        ctrlVtbl.QueryInterface = (HRESULT (STDMETHODCALLTYPE*)(ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*, REFIID, void**))CfgHandler_QueryInterface;
        ctrlVtbl.AddRef = (ULONG (STDMETHODCALLTYPE*)(ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*))CfgHandler_AddRef;
        ctrlVtbl.Release = (ULONG (STDMETHODCALLTYPE*)(ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*))CfgHandler_Release;
        ctrlVtbl.Invoke = CfgCtrlCompleted_Invoke;
        init = TRUE;
    }

    CfgHandler *handler = (CfgHandler*)malloc(sizeof(CfgHandler));
    handler->lpVtbl = &ctrlVtbl;
    handler->refCount = 1;

    env->lpVtbl->CreateCoreWebView2Controller(env, g_cfgHwnd,
        (ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*)handler);
    ((ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*)handler)->lpVtbl->Release(
        (ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*)handler);
    return S_OK;
}

static ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandlerVtbl g_cfgEnvVtbl = {
    CfgHandler_QueryInterface,
    CfgHandler_AddRef,
    CfgHandler_Release,
    CfgEnvCompleted_Invoke
};

static HRESULT STDMETHODCALLTYPE CfgCtrlCompleted_Invoke(
    ICoreWebView2CreateCoreWebView2ControllerCompletedHandler *This,
    HRESULT result, ICoreWebView2Controller *controller) {
    (void)This;
    if (FAILED(result) || !controller) {
        CfgReportInitFailureAndClose(FAILED(result) ? result : E_POINTER);
        return S_OK;
    }

    g_cfgController = controller;
    controller->lpVtbl->AddRef(controller);

    RECT bounds;
    GetClientRect(g_cfgHwnd, &bounds);
    controller->lpVtbl->put_Bounds(controller, bounds);
    controller->lpVtbl->put_IsVisible(controller, TRUE);

    ICoreWebView2 *webview = NULL;
    controller->lpVtbl->get_CoreWebView2(controller, &webview);
    if (!webview) {
        CfgReportInitFailureAndClose(E_FAIL);
        return E_FAIL;
    }
    g_cfgWebView = webview;

    ICoreWebView2Settings *settings = NULL;
    webview->lpVtbl->get_Settings(webview, &settings);
    if (settings) {
        settings->lpVtbl->put_AreDefaultContextMenusEnabled(settings, FALSE);
        settings->lpVtbl->put_AreDevToolsEnabled(settings, FALSE);
        settings->lpVtbl->put_IsStatusBarEnabled(settings, FALSE);
        settings->lpVtbl->put_IsZoomControlEnabled(settings, FALSE);
        settings->lpVtbl->Release(settings);
    }

    // Register message handler
    static ICoreWebView2WebMessageReceivedEventHandlerVtbl msgVtbl = {0};
    static BOOL msgInit = FALSE;
    if (!msgInit) {
        msgVtbl.QueryInterface = (HRESULT (STDMETHODCALLTYPE*)(ICoreWebView2WebMessageReceivedEventHandler*, REFIID, void**))CfgHandler_QueryInterface;
        msgVtbl.AddRef = (ULONG (STDMETHODCALLTYPE*)(ICoreWebView2WebMessageReceivedEventHandler*))CfgHandler_AddRef;
        msgVtbl.Release = (ULONG (STDMETHODCALLTYPE*)(ICoreWebView2WebMessageReceivedEventHandler*))CfgHandler_Release;
        msgVtbl.Invoke = CfgMsgReceived_Invoke;
        msgInit = TRUE;
    }

    CfgHandler *msgHandler = (CfgHandler*)malloc(sizeof(CfgHandler));
    msgHandler->lpVtbl = &msgVtbl;
    msgHandler->refCount = 1;

    EventRegistrationToken token;
    webview->lpVtbl->add_WebMessageReceived(webview, (ICoreWebView2WebMessageReceivedEventHandler*)msgHandler, &token);
    ((ICoreWebView2WebMessageReceivedEventHandler*)msgHandler)->lpVtbl->Release(
        (ICoreWebView2WebMessageReceivedEventHandler*)msgHandler);

    // Load embedded HTML from resources
    HRSRC hRes = FindResource(NULL, MAKEINTRESOURCE(IDR_HTML_UI), RT_RCDATA);
    if (hRes) {
        HGLOBAL hData = LoadResource(NULL, hRes);
        if (hData) {
            DWORD htmlSize = SizeofResource(NULL, hRes);
            const char *htmlUtf8 = (const char *)LockResource(hData);
            if (htmlUtf8 && htmlSize > 0) {
                int wLen = MultiByteToWideChar(CP_UTF8, 0, htmlUtf8, (int)htmlSize, NULL, 0);
                wchar_t *wHtml = malloc((wLen + 1) * sizeof(wchar_t));
                MultiByteToWideChar(CP_UTF8, 0, htmlUtf8, (int)htmlSize, wHtml, wLen);
                wHtml[wLen] = L'\0';
                webview->lpVtbl->NavigateToString(webview, wHtml);
                free(wHtml);
            }
        }
    }

    return S_OK;
}

static HRESULT STDMETHODCALLTYPE CfgMsgReceived_Invoke(
    ICoreWebView2WebMessageReceivedEventHandler *This,
    ICoreWebView2 *sender,
    ICoreWebView2WebMessageReceivedEventArgs *args) {
    (void)This; (void)sender;

    LPWSTR wMsg = NULL;
    args->lpVtbl->TryGetWebMessageAsString(args, &wMsg);
    if (!wMsg) return S_OK;

    int len = WideCharToMultiByte(CP_UTF8, 0, wMsg, -1, NULL, 0, NULL, NULL);
    char *msg = malloc(len);
    WideCharToMultiByte(CP_UTF8, 0, wMsg, -1, msg, len, NULL, NULL);
    CoTaskMemFree(wMsg);

    char action[64] = {0};
    json_get_string(msg, "action", action, sizeof(action));

    if (strcmp(action, "getInit") == 0) {
        webview_push_init_config();
    } else if (strcmp(action, "saveSettings") == 0) {
        char url[4096] = {0}, title[512] = {0}, hideJs[8192] = {0}, showJs[8192] = {0};
        char spellLangs[512] = {0};
        json_get_string(msg, "url", url, sizeof(url));
        json_get_string(msg, "windowTitle", title, sizeof(title));
        json_get_string(msg, "onHideJs", hideJs, sizeof(hideJs));
        json_get_string(msg, "onShowJs", showJs, sizeof(showJs));
        json_get_string(msg, "spellcheckLanguages", spellLangs, sizeof(spellLangs));

        wchar_t prevSpellLangs[512];
        wcscpy_s(prevSpellLangs, 512, g_config.spellcheckLanguages);

        MultiByteToWideChar(CP_UTF8, 0, url, -1, g_config.url, 2048);
        MultiByteToWideChar(CP_UTF8, 0, title, -1, g_config.windowTitle, 256);
        MultiByteToWideChar(CP_UTF8, 0, hideJs, -1, g_config.onHideJs, 4096);
        MultiByteToWideChar(CP_UTF8, 0, showJs, -1, g_config.onShowJs, 4096);
        wchar_t rawSpellLangs[512] = {0};
        MultiByteToWideChar(CP_UTF8, 0, spellLangs, -1, rawSpellLangs, 512);
        NormalizeSpellcheckLanguages(rawSpellLangs, g_config.spellcheckLanguages, 512);
        g_config.sleepWhenInactive = json_get_bool(msg, "sleepWhenInactive", FALSE);
        g_config.openNewWindowsExternally = json_get_bool(msg, "openNewWindowsExternally", FALSE);

        SaveConfigToRegistry(&g_config);
        MarkAsConfigured();
        ApplyConfiguration();

        // Spell-check languages are only read when the browser process
        // starts, so a change needs the main WebView rebuilt. Prompt on the
        // main window's thread once this dialog has closed itself.
        if (wcscmp(prevSpellLangs, g_config.spellcheckLanguages) != 0 &&
            g_hwnd && g_webViewController) {
            PostMessageW(g_hwnd, WM_APP_SPELLCHECK_CHANGED, 0, 0);
        }

        g_cfgSaved = TRUE;
        PostMessage(g_cfgHwnd, WM_CLOSE, 0, 0);
    } else if (strcmp(action, "close") == 0) {
        PostMessage(g_cfgHwnd, WM_CLOSE, 0, 0);
    } else if (strcmp(action, "resize") == 0) {
        char hStr[32] = {0};
        json_get_string(msg, "height", hStr, sizeof(hStr));
        int contentHeight = atoi(hStr);
        if (contentHeight <= 0) {
            // Try parsing as bare number (not quoted)
            const char *hp = strstr(msg, "\"height\"");
            if (hp) {
                hp += 8;
                while (*hp == ' ' || *hp == ':') hp++;
                contentHeight = atoi(hp);
            }
        }
        if (contentHeight > 0 && g_cfgHwnd) {
            // The page reports its height in CSS pixels; convert to the
            // physical pixels window sizes use.
            int physHeight = MulDiv(contentHeight, (int)GetWindowDpi(g_cfgHwnd), 96);
            RECT clientRect = {0}, windowRect = {0};
            GetClientRect(g_cfgHwnd, &clientRect);
            GetWindowRect(g_cfgHwnd, &windowRect);
            int chromeH = (windowRect.bottom - windowRect.top) - (clientRect.bottom - clientRect.top);
            int newWindowH = physHeight + chromeH;
            int windowW = windowRect.right - windowRect.left;

            // Keep the dialog inside the work area of its monitor. Sizing
            // with SWP_NOMOVE kept the top edge where a 380px-tall window
            // had been centered, so tall content grew past the bottom of
            // the screen; clamp the size (the page scrolls when it cannot
            // fit) and position the window explicitly.
            MONITORINFO mi = { sizeof(mi) };
            RECT work;
            HMONITOR mon = MonitorFromWindow(g_cfgHwnd, MONITOR_DEFAULTTONEAREST);
            if (!mon || !GetMonitorInfoW(mon, &mi)) {
                SystemParametersInfoW(SPI_GETWORKAREA, 0, &work, 0);
            } else {
                work = mi.rcWork;
            }
            int workW = work.right - work.left;
            int workH = work.bottom - work.top;
            if (newWindowH > workH) newWindowH = workH;
            if (windowW > workW) windowW = workW;

            // Always center on the measured height, clamped into the work
            // area. The page reports its height through a ResizeObserver
            // that fires more than once (a short first measurement, then the
            // real height): re-centering every time keeps the dialog
            // centered instead of anchoring its top edge and letting later
            // growth push it to the bottom of the screen.
            int posX = work.left + (workW - windowW) / 2;
            int posY = work.top + (workH - newWindowH) / 2;
            UINT flags = SWP_NOZORDER;
            if (g_cfgWindowShown) {
                flags |= SWP_NOACTIVATE;
            } else {
                flags |= SWP_SHOWWINDOW;
                KillTimer(g_cfgHwnd, ID_TIMER_CFG_SHOW_FALLBACK);
            }
            SetWindowPos(g_cfgHwnd, NULL, posX, posY, windowW, newWindowH, flags);
            g_cfgWindowShown = TRUE;
            cfg_sync_controller_bounds();
        }
    }

    free(msg);
    return S_OK;
}

// Config dialog window procedure
static LRESULT CALLBACK CfgWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_SIZE:
            cfg_sync_controller_bounds();
            return 0;

        case WM_DPICHANGED: {
            const RECT* suggested = (const RECT*)lParam;
            SetWindowPos(hwnd, NULL, suggested->left, suggested->top,
                         suggested->right - suggested->left,
                         suggested->bottom - suggested->top,
                         SWP_NOZORDER | SWP_NOACTIVATE);
            return 0;
        }

        case WM_TIMER:
            if (wParam == ID_TIMER_CFG_SHOW_FALLBACK) {
                if (g_cfgWindowShown) {
                    KillTimer(hwnd, ID_TIMER_CFG_SHOW_FALLBACK);
                    return 0;
                }
                if (g_cfgController) {
                    // Content exists but the resize message never came:
                    // show the window at its default size.
                    KillTimer(hwnd, ID_TIMER_CFG_SHOW_FALLBACK);
                    ShowWindow(hwnd, SW_SHOWNOACTIVATE);
                    UpdateWindow(hwnd);
                    g_cfgWindowShown = TRUE;
                    cfg_sync_controller_bounds();
                } else if (++g_cfgShowFallbackTries >= CFG_SHOW_FALLBACK_MAX_TRIES) {
                    // WebView2 creation neither completed nor reported
                    // failure; never present an empty window.
                    KillTimer(hwnd, ID_TIMER_CFG_SHOW_FALLBACK);
                    CfgReportInitFailureAndClose(HRESULT_FROM_WIN32(ERROR_TIMEOUT));
                }
                // Otherwise keep waiting: the periodic timer fires again.
                return 0;
            }
            break;

        case WM_CLOSE:
            g_cfgWindowShown = FALSE;
            KillTimer(hwnd, ID_TIMER_CFG_SHOW_FALLBACK);
            if (g_cfgController) {
                g_cfgController->lpVtbl->Close(g_cfgController);
                g_cfgController->lpVtbl->Release(g_cfgController);
                g_cfgController = NULL;
            }
            if (g_cfgWebView) {
                g_cfgWebView->lpVtbl->Release(g_cfgWebView);
                g_cfgWebView = NULL;
            }
            if (g_cfgEnv) {
                g_cfgEnv->lpVtbl->Release(g_cfgEnv);
                g_cfgEnv = NULL;
            }
            DestroyWindow(hwnd);
            return 0;

        case WM_DESTROY:
            g_cfgHwnd = NULL;
            g_cfgWindowShown = FALSE;
            KillTimer(hwnd, ID_TIMER_CFG_SHOW_FALLBACK);
            return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static void ShowConfigWebViewDialog(void) {
    if (g_cfgHwnd != NULL) {
        SetForegroundWindow(g_cfgHwnd);
        return;
    }

    if (!fnCreateEnvironment && !load_webview2_loader()) {
        MessageBoxW(NULL,
            L"Failed to load WebView2.\n\n"
            L"Please ensure the Microsoft Edge WebView2 Runtime is installed.\n"
            L"Download from: https://developer.microsoft.com/en-us/microsoft-edge/webview2/",
            L"SystrayLauncher", MB_ICONERROR | MB_OK);
        return;
    }

    // Register window class (once)
    static BOOL classRegistered = FALSE;
    if (!classRegistered) {
        WNDCLASSEXW wc = {0};
        wc.cbSize = sizeof(wc);
        wc.lpfnWndProc = CfgWndProc;
        wc.hInstance = g_hInstance;
        wc.hIcon = LoadIconW(g_hInstance, MAKEINTRESOURCEW(IDI_TRAYICON));
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
        wc.lpszClassName = L"SystrayLauncherCfgWnd";
        wc.hIconSm = LoadIconW(g_hInstance, MAKEINTRESOURCEW(IDI_TRAYICON));
        RegisterClassExW(&wc);
        classRegistered = TRUE;
    }

    // Center the initial (hidden) window in the primary work area; the
    // resize message from the page re-centers it at its real height before
    // it is shown, and clamps it to the work area of whatever monitor it
    // lands on.
    RECT workArea;
    SystemParametersInfoW(SPI_GETWORKAREA, 0, &workArea, 0);
    int dpi = (int)GetWindowDpi(NULL);
    int width = MulDiv(480, dpi, 96), height = MulDiv(380, dpi, 96);
    if (height > workArea.bottom - workArea.top) height = workArea.bottom - workArea.top;
    int posX = workArea.left + ((workArea.right - workArea.left) - width) / 2;
    int posY = workArea.top + ((workArea.bottom - workArea.top) - height) / 2;

    g_cfgHwnd = CreateWindowExW(0, L"SystrayLauncherCfgWnd", L"Configuration",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
        posX, posY, width, height,
        NULL, NULL, g_hInstance, NULL);

    if (!g_cfgHwnd) return;
    g_cfgWindowShown = FALSE;
    g_cfgShowFallbackTries = 0;
    SetTimer(g_cfgHwnd, ID_TIMER_CFG_SHOW_FALLBACK, CFG_SHOW_FALLBACK_DELAY_MS, NULL);

    // Build user data folder path
    WCHAR userDataFolder[MAX_PATH];
    DWORD tempLen = GetTempPathW(MAX_PATH, userDataFolder);
    if (tempLen > 0 && tempLen < MAX_PATH - 30) {
        wcscat(userDataFolder, L"SystrayLauncher.WebView2");
    } else {
        wcscpy(userDataFolder, L"");
    }

    CfgHandler *envHandler = (CfgHandler*)malloc(sizeof(CfgHandler));
    envHandler->lpVtbl = &g_cfgEnvVtbl;
    envHandler->refCount = 1;

    HRESULT hr = fnCreateEnvironment(NULL, userDataFolder[0] ? userDataFolder : NULL, NULL,
        (ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*)envHandler);
    ((ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*)envHandler)->lpVtbl->Release(
        (ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*)envHandler);

    if (FAILED(hr)) {
        MessageBoxW(NULL,
            L"Failed to initialize WebView2.\n\n"
            L"Please ensure the Microsoft Edge WebView2 Runtime is installed.",
            L"SystrayLauncher", MB_ICONERROR | MB_OK);
        DestroyWindow(g_cfgHwnd);
        g_cfgHwnd = NULL;
    }
}

// ---------------------------------------------------------------------------
// Spell checking
//
// WebView2 has no public API to pick spell-check languages (tracked in
// MicrosoftEdge/WebView2Feedback#2040). The engine is fully present though:
// Chromium reads the dictionary list from the profile's Preferences JSON at
// browser-process startup. This app owns its user data folder exclusively and
// is single-instance, so we can safely rewrite those keys while the browser
// process is not running. The dictionary list must stay a subset of
// intl.accept_languages or Chromium prunes it on startup, which is why both
// prefs are written together.
// ---------------------------------------------------------------------------

// Normalize a comma-separated list of spell-check language tags: trim
// whitespace, drop tokens with characters other than letters/digits/hyphens
// (they would be invalid tags and must not reach the JSON patch), dedupe, and
// fix the usual BCP-47 casing (en-us -> en-US, pt-br -> pt-BR) so the tags
// match Chromium's dictionary names.
static void NormalizeSpellcheckLanguages(const wchar_t* in, wchar_t* out, size_t outLen) {
    size_t o = 0;
    if (outLen == 0) return;
    out[0] = L'\0';
    if (!in) return;

    const wchar_t* p = in;
    while (*p) {
        while (*p == L',' || *p == L';' || iswspace(*p)) p++;
        if (!*p) break;

        wchar_t token[32];
        size_t t = 0;
        BOOL valid = TRUE;
        while (*p && *p != L',' && *p != L';' && !iswspace(*p)) {
            wchar_t c = (*p == L'_') ? L'-' : *p;
            if (!iswalnum(c) && c != L'-') valid = FALSE;
            if (t < 31) token[t++] = c; else valid = FALSE;
            p++;
        }
        token[t] = L'\0';
        if (!valid || t == 0) continue;

        // Casing per subtag: primary lowercase, 2-letter region UPPER,
        // 4-letter script Titlecase.
        size_t start = 0;
        for (size_t i = 0; i <= t && valid; i++) {
            if (token[i] == L'-' || token[i] == L'\0') {
                size_t len = i - start;
                if (len == 0) { valid = FALSE; break; }
                if (start == 0) {
                    for (size_t j = start; j < i; j++) token[j] = towlower(token[j]);
                } else if (len == 2) {
                    for (size_t j = start; j < i; j++) token[j] = towupper(token[j]);
                } else if (len == 4) {
                    token[start] = towupper(token[start]);
                    for (size_t j = start + 1; j < i; j++) token[j] = towlower(token[j]);
                } else {
                    for (size_t j = start; j < i; j++) token[j] = towlower(token[j]);
                }
                start = i + 1;
            }
        }
        if (!valid) continue;

        // Skip duplicates
        BOOL dup = FALSE;
        const wchar_t* q = out;
        while (*q) {
            const wchar_t* e = wcschr(q, L',');
            size_t len = e ? (size_t)(e - q) : wcslen(q);
            if (len == t && wcsncmp(q, token, t) == 0) { dup = TRUE; break; }
            q = e ? e + 1 : q + len;
        }
        if (dup) continue;

        size_t need = t + (o > 0 ? 1 : 0);
        if (o + need + 1 > outLen) break;
        if (o > 0) out[o++] = L',';
        wmemcpy(out + o, token, t);
        o += t;
        out[o] = L'\0';
    }
}

static void GetMainUserDataFolder(wchar_t path[MAX_PATH]) {
    path[0] = L'\0';
    SHGetFolderPathW(NULL, CSIDL_LOCAL_APPDATA, NULL, 0, path);
    PathAppendW(path, APP_NAME L"\\WebView2Data");
}

// --- Minimal JSON surgery (string- and nesting-aware, never guesses) -------

static const char* json_skip_ws(const char* p, const char* end) {
    while (p < end && (*p == ' ' || *p == '\t' || *p == '\n' || *p == '\r')) p++;
    return p;
}

// p points at the opening quote; returns one past the closing quote.
static const char* json_skip_string(const char* p, const char* end) {
    p++;
    while (p < end) {
        if (*p == '\\') { p += 2; continue; }
        if (*p == '"') return p + 1;
        p++;
    }
    return NULL;
}

// Returns one past the end of the value starting at p (string, object,
// array, number, bool or null), or NULL if the input is malformed/truncated.
static const char* json_skip_value(const char* p, const char* end) {
    p = json_skip_ws(p, end);
    if (p >= end) return NULL;
    if (*p == '"') return json_skip_string(p, end);
    if (*p == '{' || *p == '[') {
        int depth = 0;
        while (p < end) {
            if (*p == '"') {
                p = json_skip_string(p, end);
                if (!p) return NULL;
                continue;
            }
            if (*p == '{' || *p == '[') {
                depth++;
            } else if (*p == '}' || *p == ']') {
                depth--;
                if (depth == 0) return p + 1;
            }
            p++;
        }
        return NULL;
    }
    while (p < end && *p != ',' && *p != '}' && *p != ']' &&
           *p != ' ' && *p != '\t' && *p != '\n' && *p != '\r') {
        p++;
    }
    return p;
}

// obj points at '{'. Looks up "key" among this object's own members (not in
// nested objects). Returns the value start and sets *valEnd, or NULL.
static const char* json_object_find(const char* obj, const char* end,
                                    const char* key, const char** valEnd) {
    if (obj >= end || *obj != '{') return NULL;
    size_t klen = strlen(key);
    const char* p = obj + 1;
    for (;;) {
        p = json_skip_ws(p, end);
        if (p >= end || *p == '}') return NULL;
        if (*p != '"') return NULL;
        const char* ks = p + 1;
        const char* kq = json_skip_string(p, end);
        if (!kq) return NULL;
        const char* ke = kq - 1;
        p = json_skip_ws(kq, end);
        if (p >= end || *p != ':') return NULL;
        const char* vs = json_skip_ws(p + 1, end);
        const char* ve = json_skip_value(vs, end);
        if (!ve) return NULL;
        if ((size_t)(ke - ks) == klen && strncmp(ks, key, klen) == 0) {
            *valEnd = ve;
            return vs;
        }
        p = json_skip_ws(ve, end);
        if (p >= end) return NULL;
        if (*p == ',') { p++; continue; }
        return NULL;
    }
}

// Replace [from,to) in buf with insert; returns a new heap buffer.
static char* str_splice(const char* buf, size_t len, const char* from,
                        const char* to, const char* insert, size_t* newLen) {
    size_t pre = (size_t)(from - buf);
    size_t post = len - (size_t)(to - buf);
    size_t ins = strlen(insert);
    char* result = (char*)malloc(pre + ins + post + 1);
    if (!result) return NULL;
    memcpy(result, buf, pre);
    memcpy(result + pre, insert, ins);
    memcpy(result + pre + ins, to, post);
    result[pre + ins + post] = '\0';
    *newLen = pre + ins + post;
    return result;
}

// Ensure root.<parentKey>.<childKey> == valueJson. Returns a new heap buffer
// (and updates *newLen), or NULL when the document isn't a JSON object we can
// understand - the caller then leaves the file untouched. New keys are
// inserted at the end of their object so they win under Chromium's
// last-key-wins JSON parsing even in pathological duplicate-key files.
static char* json_set_nested(const char* json, size_t len,
                             const char* parentKey, const char* childKey,
                             const char* valueJson, size_t* newLen) {
    const char* end = json + len;
    const char* root = json_skip_ws(json, end);
    if (root >= end || *root != '{') return NULL;
    const char* rootEnd = json_skip_value(root, end);
    if (!rootEnd) return NULL;

    char insert[2048];
    const char* pve = NULL;
    const char* pv = json_object_find(root, rootEnd, parentKey, &pve);
    if (pv && *pv == '{') {
        const char* cve = NULL;
        const char* cv = json_object_find(pv, pve, childKey, &cve);
        if (cv) {
            return str_splice(json, len, cv, cve, valueJson, newLen);
        }
        const char* inner = json_skip_ws(pv + 1, pve);
        BOOL emptyObj = (inner < pve && *inner == '}');
        snprintf(insert, sizeof(insert), "%s\"%s\":%s",
                 emptyObj ? "" : ",", childKey, valueJson);
        return str_splice(json, len, pve - 1, pve - 1, insert, newLen);
    }
    if (pv) {
        // Parent exists but is not an object: replace it wholesale.
        snprintf(insert, sizeof(insert), "{\"%s\":%s}", childKey, valueJson);
        return str_splice(json, len, pv, pve, insert, newLen);
    }
    const char* innerRoot = json_skip_ws(root + 1, rootEnd);
    BOOL rootEmpty = (innerRoot < rootEnd && *innerRoot == '}');
    snprintf(insert, sizeof(insert), "%s\"%s\":{\"%s\":%s}",
             rootEmpty ? "" : ",", parentKey, childKey, valueJson);
    return str_splice(json, len, rootEnd - 1, rootEnd - 1, insert, newLen);
}

// Write the configured spell-check languages into the WebView2 profile's
// Preferences file. Must only run while the browser process for the main
// user data folder is not running (app startup, or after BrowserProcessExited).
static void PatchSpellcheckPreferences(void) {
    if (g_config.spellcheckLanguages[0] == L'\0') return;

    char langs[512];
    if (WideCharToMultiByte(CP_UTF8, 0, g_config.spellcheckLanguages, -1,
                            langs, sizeof(langs), NULL, NULL) <= 0) {
        return;
    }

    // "en-US,pl" -> ["en-US","pl"] for spellcheck.dictionaries, and a quoted
    // string for intl.accept_languages / intl.selected_languages.
    char dictJson[1200];
    char acceptJson[600];
    size_t d = 0;
    dictJson[d++] = '[';
    const char* tok = langs;
    BOOL first = TRUE;
    while (*tok) {
        const char* comma = strchr(tok, ',');
        size_t tlen = comma ? (size_t)(comma - tok) : strlen(tok);
        if (d + tlen + 4 >= sizeof(dictJson)) break;
        if (!first) dictJson[d++] = ',';
        dictJson[d++] = '"';
        memcpy(dictJson + d, tok, tlen);
        d += tlen;
        dictJson[d++] = '"';
        first = FALSE;
        tok = comma ? comma + 1 : tok + tlen;
    }
    dictJson[d++] = ']';
    dictJson[d] = '\0';
    snprintf(acceptJson, sizeof(acceptJson), "\"%s\"", langs);

    wchar_t prefsDir[MAX_PATH];
    GetMainUserDataFolder(prefsDir);
    // The runtime nests the actual browser profile in an "EBWebView"
    // subfolder of the user data folder, so the Preferences file lives at
    // <UDF>\EBWebView\Default\Preferences - not directly under the UDF.
    PathAppendW(prefsDir, L"EBWebView");
    PathAppendW(prefsDir, L"Default");
    wchar_t prefsPath[MAX_PATH];
    wcscpy_s(prefsPath, MAX_PATH, prefsDir);
    PathAppendW(prefsPath, L"Preferences");

    FILE* f = NULL;
    if (_wfopen_s(&f, prefsPath, L"rb") != 0 || !f) {
        // Profile doesn't exist yet (first run): seed a minimal Preferences
        // file; Chromium fills in defaults for everything else.
        SHCreateDirectoryExW(NULL, prefsDir, NULL);
        FILE* nf = NULL;
        if (_wfopen_s(&nf, prefsPath, L"wb") == 0 && nf) {
            fprintf(nf,
                "{\"intl\":{\"accept_languages\":%s,\"selected_languages\":%s},"
                "\"spellcheck\":{\"dictionaries\":%s}}",
                acceptJson, acceptJson, dictJson);
            fclose(nf);
            DebugPrint(L"[INFO] Seeded WebView2 Preferences with spell-check languages: %s\n",
                       g_config.spellcheckLanguages);
        }
        return;
    }

    fseek(f, 0, SEEK_END);
    long fsize = ftell(f);
    fseek(f, 0, SEEK_SET);
    if (fsize <= 0 || fsize > 32 * 1024 * 1024) {
        fclose(f);
        return;
    }
    char* buf = (char*)malloc((size_t)fsize + 1);
    if (!buf) {
        fclose(f);
        return;
    }
    size_t got = fread(buf, 1, (size_t)fsize, f);
    fclose(f);
    buf[got] = '\0';

    struct {
        const char* parent;
        const char* child;
        const char* value;
    } patches[] = {
        { "spellcheck", "dictionaries", dictJson },
        { "intl", "accept_languages", acceptJson },
        { "intl", "selected_languages", acceptJson },
    };

    char* cur = buf;
    size_t curLen = got;
    BOOL ok = TRUE;
    for (int i = 0; i < 3; i++) {
        size_t nextLen = 0;
        char* next = json_set_nested(cur, curLen, patches[i].parent,
                                     patches[i].child, patches[i].value, &nextLen);
        if (!next) {
            ok = FALSE;
            break;
        }
        if (cur != buf) free(cur);
        cur = next;
        curLen = nextLen;
    }

    if (!ok) {
        DebugPrint(L"[WARNING] Could not parse WebView2 Preferences; spell-check patch skipped\n");
        if (cur != buf) free(cur);
        free(buf);
        return;
    }

    if (curLen == got && memcmp(cur, buf, got) == 0) {
        DebugPrint(L"[INFO] Spell-check preferences already up to date\n");
        free(cur);
        free(buf);
        return;
    }

    // Write to a temp file and swap it in so a crash can't corrupt the profile.
    wchar_t tmpPath[MAX_PATH];
    if (swprintf_s(tmpPath, MAX_PATH, L"%s.stl_tmp", prefsPath) > 0) {
        FILE* wf = NULL;
        if (_wfopen_s(&wf, tmpPath, L"wb") == 0 && wf) {
            size_t written = fwrite(cur, 1, curLen, wf);
            fclose(wf);
            if (written == curLen &&
                MoveFileExW(tmpPath, prefsPath, MOVEFILE_REPLACE_EXISTING)) {
                DebugPrint(L"[INFO] Applied spell-check languages to WebView2 profile: %s\n",
                           g_config.spellcheckLanguages);
            } else {
                DeleteFileW(tmpPath);
                DebugPrint(L"[WARNING] Failed to update WebView2 Preferences file\n");
            }
        }
    }
    free(cur);
    free(buf);
}

// --- WebView rebuild (applies new spell-check languages without an app
// restart: Chromium only reads the prefs at browser-process startup) --------

// Create the WebView2 environment for the main window. The controller and
// WebView are then built by the completion handlers (EnvCompletedHandler et
// al). Used both from WM_CREATE and when rebuilding after a language change.
static void CreateMainWebViewEnvironment(HWND hwnd) {
    wchar_t userDataPath[MAX_PATH];
    GetMainUserDataFolder(userDataPath);
    SHCreateDirectoryExW(NULL, userDataPath, NULL);

    InterlockedExchange(&g_webViewCreatePending, TRUE);

    EnvCompletedHandler* envHandler = (EnvCompletedHandler*)calloc(1, sizeof(EnvCompletedHandler));
    if (!envHandler) {
        InterlockedExchange(&g_webViewCreatePending, FALSE);
        return;
    }

    static ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandlerVtbl envVtbl = {
        EnvCompletedHandler_QueryInterface,
        EnvCompletedHandler_AddRef,
        EnvCompletedHandler_Release,
        EnvCompletedHandler_Invoke
    };
    envHandler->lpVtbl = &envVtbl;
    envHandler->refCount = 1;
    envHandler->hwnd = hwnd;
    envHandler->userDataPath = _wcsdup(userDataPath);

    HRESULT hr = fnCreateEnvironment(NULL, userDataPath, NULL,
        (ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*)envHandler);
    envHandler->lpVtbl->Release((ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*)envHandler);
    if (FAILED(hr)) {
        InterlockedExchange(&g_webViewCreatePending, FALSE);
        DebugPrint(L"[WARNING] CreateCoreWebView2Environment call failed. HRESULT: 0x%08X\n", hr);
    }
}

typedef struct {
    ICoreWebView2BrowserProcessExitedEventHandlerVtbl* lpVtbl;
    LONG refCount;
} BrowserExitedHandler;

static HRESULT STDMETHODCALLTYPE BrowserExitedHandler_QueryInterface(
    ICoreWebView2BrowserProcessExitedEventHandler* This,
    REFIID riid, void** ppvObject) {
    if (IsEqualIID(riid, &IID_IUnknown) ||
        IsEqualIID(riid, &IID_ICoreWebView2BrowserProcessExitedEventHandler)) {
        *ppvObject = This;
        This->lpVtbl->AddRef(This);
        return S_OK;
    }
    *ppvObject = NULL;
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE BrowserExitedHandler_AddRef(
    ICoreWebView2BrowserProcessExitedEventHandler* This) {
    return InterlockedIncrement(&((BrowserExitedHandler*)This)->refCount);
}

static ULONG STDMETHODCALLTYPE BrowserExitedHandler_Release(
    ICoreWebView2BrowserProcessExitedEventHandler* This) {
    ULONG refCount = InterlockedDecrement(&((BrowserExitedHandler*)This)->refCount);
    if (refCount == 0) free(This);
    return refCount;
}

static HRESULT STDMETHODCALLTYPE BrowserExitedHandler_Invoke(
    ICoreWebView2BrowserProcessExitedEventHandler* This,
    ICoreWebView2Environment* sender,
    ICoreWebView2BrowserProcessExitedEventArgs* args) {
    (void)This; (void)sender; (void)args;
    DebugPrint(L"[INFO] WebView2 browser process exited\n");
    if (g_hwnd) PostMessageW(g_hwnd, WM_APP_WEBVIEW_RECREATE, 0, 0);
    return S_OK;
}

// The BrowserProcessExited event stays registered for the whole lifetime of
// each environment: it drives both the deliberate rebuild (spell-check
// change) and recovery from a browser process that died on its own - the
// latter used to leave a permanently blank container.
static void RegisterBrowserExitedOnCurrentEnv(void) {
    if (!g_webViewEnv || g_browserExitedRegistered) return;

    ICoreWebView2Environment5* env5 = NULL;
    if (FAILED(g_webViewEnv->lpVtbl->QueryInterface(g_webViewEnv,
            &IID_ICoreWebView2Environment5, (void**)&env5)) || !env5) {
        DebugPrint(L"[WARNING] BrowserProcessExited event not supported by this runtime\n");
        return;
    }

    BrowserExitedHandler* handler =
        (BrowserExitedHandler*)calloc(1, sizeof(BrowserExitedHandler));
    if (handler) {
        static ICoreWebView2BrowserProcessExitedEventHandlerVtbl exitVtbl = {
            BrowserExitedHandler_QueryInterface,
            BrowserExitedHandler_AddRef,
            BrowserExitedHandler_Release,
            BrowserExitedHandler_Invoke
        };
        handler->lpVtbl = &exitVtbl;
        handler->refCount = 1;
        if (SUCCEEDED(env5->lpVtbl->add_BrowserProcessExited(env5,
                (ICoreWebView2BrowserProcessExitedEventHandler*)handler,
                &g_browserExitedToken))) {
            g_browserExitedRegistered = TRUE;
        }
        handler->lpVtbl->Release((ICoreWebView2BrowserProcessExitedEventHandler*)handler);
    }
    env5->lpVtbl->Release(env5);
}

static void UnregisterBrowserExitedFromCurrentEnv(void) {
    if (!g_webViewEnv || !g_browserExitedRegistered) return;

    ICoreWebView2Environment5* env5 = NULL;
    if (SUCCEEDED(g_webViewEnv->lpVtbl->QueryInterface(g_webViewEnv,
            &IID_ICoreWebView2Environment5, (void**)&env5)) && env5) {
        env5->lpVtbl->remove_BrowserProcessExited(env5, g_browserExitedToken);
        env5->lpVtbl->Release(env5);
    }
    g_browserExitedRegistered = FALSE;
}

// Tear down the main WebView so its browser process exits; the Preferences
// patch and the rebuild happen in FinishMainWebViewRecreate once the
// BrowserProcessExited event fires (the file is only flushed - and unlocked
// for our purposes - when that process is gone). A fallback timer covers
// runtimes where the event can't be observed.
static void BeginMainWebViewRecreate(void) {
    if (!g_hwnd) return;
    if (InterlockedCompareExchange(&g_webViewRecreatePending, TRUE, TRUE) == TRUE) return;

    if (!g_webViewController) {
        // Nothing is running; the startup path will patch and create as usual.
        PatchSpellcheckPreferences();
        return;
    }

    InterlockedExchange(&g_webViewRecreatePending, TRUE);

    // BrowserProcessExited is already registered on the environment (done
    // when the environment was created); its handler will post the rebuild.
    SetTimer(g_hwnd, ID_TIMER_WEBVIEW_RECREATE, WEBVIEW_RECREATE_FALLBACK_MS, NULL);

    KillTimer(g_hwnd, ID_TIMER_INITIAL_HIDE_JS);
    KillTimer(g_hwnd, ID_TIMER_WEBVIEW_PREWARM);
    KillTimer(g_hwnd, ID_TIMER_WEBVIEW_PRELOAD);
    KillTimer(g_hwnd, ID_TIMER_POWER_RESUME);
    KillTimer(g_hwnd, ID_TIMER_WEBVIEW_LIVENESS);

    InterlockedExchange(&g_isInitialized, FALSE);
    InterlockedExchange(&g_initialPreloadComplete, FALSE);
    InterlockedExchange(&g_webViewSuspendPending, FALSE);
    InterlockedExchange(&g_webViewSuspended, FALSE);
    InterlockedExchange(&g_webViewPrewarmActive, FALSE);
    InterlockedExchange(&g_powerResumePending, FALSE);
    InterlockedExchange(&g_webViewPingOutstanding, FALSE);
    g_powerKickCount = 0;
    g_jsVisibility = JS_VISIBILITY_UNKNOWN;

    if (g_webView) {
        g_webView->lpVtbl->Release(g_webView);
        g_webView = NULL;
    }
    if (g_webViewController) {
        g_webViewController->lpVtbl->Close(g_webViewController);
        g_webViewController->lpVtbl->Release(g_webViewController);
        g_webViewController = NULL;
    }
    DebugPrint(L"[INFO] Main WebView closed; waiting for browser process exit\n");
}

static void FinishMainWebViewRecreate(HWND hwnd) {
    if (InterlockedExchange(&g_webViewRecreatePending, FALSE) != TRUE) return;
    KillTimer(hwnd, ID_TIMER_WEBVIEW_RECREATE);

    if (g_webViewEnv) {
        UnregisterBrowserExitedFromCurrentEnv();
        g_webViewEnv->lpVtbl->Release(g_webViewEnv);
        g_webViewEnv = NULL;
    }

    DebugPrint(L"[INFO] Rebuilding main WebView with updated spell-check languages\n");
    PatchSpellcheckPreferences();
    CreateMainWebViewEnvironment(hwnd);
}

// The browser process died without the app asking for it (crash, kill, out
// of memory, runtime servicing) or stopped honoring resume requests. Drop
// every stale COM object and build a fresh WebView.
static void HandleUnexpectedBrowserExit(HWND hwnd) {
    if (InterlockedCompareExchange(&g_webViewCreatePending, TRUE, TRUE) == TRUE) return;

    DebugPrint(L"[WARNING] WebView2 browser gone or unresponsive; rebuilding\n");

    KillTimer(hwnd, ID_TIMER_INITIAL_HIDE_JS);
    KillTimer(hwnd, ID_TIMER_WEBVIEW_PREWARM);
    KillTimer(hwnd, ID_TIMER_WEBVIEW_PRELOAD);
    KillTimer(hwnd, ID_TIMER_WEBVIEW_RECREATE);
    KillTimer(hwnd, ID_TIMER_POWER_RESUME);
    KillTimer(hwnd, ID_TIMER_WEBVIEW_LIVENESS);

    InterlockedExchange(&g_isInitialized, FALSE);
    InterlockedExchange(&g_initialPreloadComplete, FALSE);
    InterlockedExchange(&g_webViewSuspendPending, FALSE);
    InterlockedExchange(&g_webViewSuspended, FALSE);
    InterlockedExchange(&g_webViewPrewarmActive, FALSE);
    InterlockedExchange(&g_resumeFailureCount, 0);
    InterlockedExchange(&g_powerResumePending, FALSE);
    InterlockedExchange(&g_webViewPingOutstanding, FALSE);
    g_powerKickCount = 0;
    g_jsVisibility = JS_VISIBILITY_UNKNOWN;

    if (g_webView) {
        g_webView->lpVtbl->Release(g_webView);
        g_webView = NULL;
    }
    if (g_webViewController) {
        g_webViewController->lpVtbl->Close(g_webViewController);
        g_webViewController->lpVtbl->Release(g_webViewController);
        g_webViewController = NULL;
    }
    if (g_webViewEnv) {
        UnregisterBrowserExitedFromCurrentEnv();
        g_webViewEnv->lpVtbl->Release(g_webViewEnv);
        g_webViewEnv = NULL;
    }

    ULONGLONG now = GetTickCount64();
    if (g_rebuildBurstStartTick == 0 ||
        now - g_rebuildBurstStartTick > REBUILD_BURST_WINDOW_MS) {
        g_rebuildBurstStartTick = now;
        g_rebuildBurstCount = 0;
    }
    if (++g_rebuildBurstCount > REBUILD_BURST_MAX) {
        DebugPrint(L"[WARNING] Too many WebView rebuilds; waiting for a manual Refresh/Open\n");
        return;
    }

    PatchSpellcheckPreferences();
    CreateMainWebViewEnvironment(hwnd);
}

// Recovery entry point for the tray actions: if the WebView is gone (rebuild
// limiter tripped, or creation failed earlier) a Refresh/Open builds it anew.
static void RebuildMainWebViewIfDead(void) {
    if (g_webView || g_webViewController || g_webViewEnv) return;
    if (!g_hwnd) return;
    if (InterlockedCompareExchange(&g_webViewCreatePending, TRUE, TRUE) == TRUE) return;
    if (InterlockedCompareExchange(&g_webViewRecreatePending, TRUE, TRUE) == TRUE) return;

    g_rebuildBurstCount = 0;
    DebugPrint(L"[INFO] Rebuilding missing WebView from tray action\n");
    PatchSpellcheckPreferences();
    CreateMainWebViewEnvironment(g_hwnd);
}

// WebView2 Handler Implementations
HRESULT STDMETHODCALLTYPE EnvCompletedHandler_QueryInterface(
    ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler* This,
    REFIID riid, void** ppvObject) {
    if (IsEqualIID(riid, &IID_IUnknown) ||
        IsEqualIID(riid, &IID_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler)) {
        *ppvObject = This;
        This->lpVtbl->AddRef(This);
        return S_OK;
    }
    *ppvObject = NULL;
    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE EnvCompletedHandler_AddRef(
    ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler* This) {
    EnvCompletedHandler* handler = (EnvCompletedHandler*)This;
    return InterlockedIncrement(&handler->refCount);
}

ULONG STDMETHODCALLTYPE EnvCompletedHandler_Release(
    ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler* This) {
    EnvCompletedHandler* handler = (EnvCompletedHandler*)This;
    ULONG refCount = InterlockedDecrement(&handler->refCount);
    if (refCount == 0) {
        free(handler->userDataPath);
        free(handler);
    }
    return refCount;
}

HRESULT STDMETHODCALLTYPE EnvCompletedHandler_Invoke(
    ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler* This,
    HRESULT result, ICoreWebView2Environment* environment) {

    if (FAILED(result)) {
        InterlockedExchange(&g_webViewCreatePending, FALSE);
        MessageBoxW(NULL, L"WebView2 environment creation failed", L"Error", MB_OK | MB_ICONERROR);
        return result;
    }

    // Get WebView2 browser version string
    LPWSTR versionString = NULL;
    if (SUCCEEDED(environment->lpVtbl->get_BrowserVersionString(environment, &versionString)) && versionString) {
        wcscpy_s(g_webView2Version, 128, versionString);
        CoTaskMemFree(versionString);
    }

    // Keep the environment for the BrowserProcessExited event (deliberate
    // rebuilds and crash recovery both depend on it).
    if (g_webViewEnv) {
        UnregisterBrowserExitedFromCurrentEnv();
        g_webViewEnv->lpVtbl->Release(g_webViewEnv);
    }
    g_webViewEnv = environment;
    environment->lpVtbl->AddRef(environment);
    RegisterBrowserExitedOnCurrentEnv();

    EnvCompletedHandler* handler = (EnvCompletedHandler*)This;
    HWND hwnd = handler->hwnd;

    ControllerCompletedHandler* controllerHandler = (ControllerCompletedHandler*)calloc(1, sizeof(ControllerCompletedHandler));
    if (!controllerHandler) return E_OUTOFMEMORY;

    static ICoreWebView2CreateCoreWebView2ControllerCompletedHandlerVtbl controllerVtbl = {
        ControllerCompletedHandler_QueryInterface,
        ControllerCompletedHandler_AddRef,
        ControllerCompletedHandler_Release,
        ControllerCompletedHandler_Invoke
    };

    controllerHandler->lpVtbl = &controllerVtbl;
    controllerHandler->refCount = 1;
    controllerHandler->hwnd = hwnd;

    environment->lpVtbl->CreateCoreWebView2Controller(environment, hwnd,
        (ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*)controllerHandler);
    
    controllerHandler->lpVtbl->Release((ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*)controllerHandler);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ControllerCompletedHandler_QueryInterface(
    ICoreWebView2CreateCoreWebView2ControllerCompletedHandler* This,
    REFIID riid, void** ppvObject) {
    if (IsEqualIID(riid, &IID_IUnknown) ||
        IsEqualIID(riid, &IID_ICoreWebView2CreateCoreWebView2ControllerCompletedHandler)) {
        *ppvObject = This;
        This->lpVtbl->AddRef(This);
        return S_OK;
    }
    *ppvObject = NULL;
    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE ControllerCompletedHandler_AddRef(
    ICoreWebView2CreateCoreWebView2ControllerCompletedHandler* This) {
    ControllerCompletedHandler* handler = (ControllerCompletedHandler*)This;
    return InterlockedIncrement(&handler->refCount);
}

ULONG STDMETHODCALLTYPE ControllerCompletedHandler_Release(
    ICoreWebView2CreateCoreWebView2ControllerCompletedHandler* This) {
    ControllerCompletedHandler* handler = (ControllerCompletedHandler*)This;
    ULONG refCount = InterlockedDecrement(&handler->refCount);
    if (refCount == 0) {
        free(handler);
    }
    return refCount;
}

HRESULT STDMETHODCALLTYPE ControllerCompletedHandler_Invoke(
    ICoreWebView2CreateCoreWebView2ControllerCompletedHandler* This,
    HRESULT result, ICoreWebView2Controller* controller) {

    InterlockedExchange(&g_webViewCreatePending, FALSE);

    if (FAILED(result)) {
        MessageBoxW(NULL, L"WebView2 controller creation failed", L"Error", MB_OK | MB_ICONERROR);
        return result;
    }

    ControllerCompletedHandler* handler = (ControllerCompletedHandler*)This;
    HWND hwnd = handler->hwnd;

    g_webViewController = controller;
    g_webViewController->lpVtbl->AddRef(g_webViewController);

    ICoreWebView2* webview2 = NULL;
    controller->lpVtbl->get_CoreWebView2(controller, &webview2);
    if (webview2) {
        g_webView = webview2;
        
        BOOL initiallyVisible = IsWindowActuallyVisible(hwnd);

        // Pre-size the (still hidden) window to the size it will have when
        // shown, so the preload lays out and renders at the final dimensions
        // and the first open needs no reflow.
        if (!initiallyVisible) {
            int wx, wy, ww, wh;
            GetTargetWindowRect(&wx, &wy, &ww, &wh);
            SetWindowPos(hwnd, NULL, wx, wy, ww, wh, SWP_NOZORDER | SWP_NOACTIVATE);
        }

        RECT bounds;
        GetClientRect(hwnd, &bounds);
        controller->lpVtbl->put_Bounds(controller, bounds);

        // Requirement (c): keep the WebView rendering (IsVisible = TRUE) during
        // the preload even though the host window stays hidden. This is the
        // documented way to keep a WebView "warm" — the page loads and renders
        // off-screen so it is ready to display instantly. We only turn
        // rendering off (IsVisible = FALSE) at the moment we suspend for sleep.
        controller->lpVtbl->put_IsVisible(controller, TRUE);

        // Settle the sleep state only once the initial navigation finishes so
        // we never suspend a half-loaded page (see OnMainNavigationCompleted).
        RegisterMainNavigationCompletedHandler(webview2);
        RegisterMainNewWindowRequestedHandler(webview2);
        RegisterMainProcessFailedHandler(webview2);

        webview2->lpVtbl->Navigate(webview2, g_initialUrl);

        InterlockedExchange(&g_resumeFailureCount, 0);
        InterlockedExchange(&g_isInitialized, TRUE);
        InterlockedExchange(&g_webViewDesiredVisible, TRUE);
        ResumeMainWebViewRuntime();

        if (initiallyVisible && InterlockedExchange(&g_resetUrlOnNextShow, FALSE) == TRUE) {
            ResetTargetPageIfNeeded();
        }

        // Schedule initial JS sync after WebView is ready.
        if (g_config.onHideJs[0] != L'\0' || g_config.onShowJs[0] != L'\0') {
            SetTimer(hwnd, ID_TIMER_INITIAL_HIDE_JS, INITIAL_HIDE_JS_DELAY_MS, NULL);
        }
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE ClearBrowsingDataCompletedHandler_QueryInterface(
    ICoreWebView2ClearBrowsingDataCompletedHandler* This,
    REFIID riid, void** ppvObject) {
    if (IsEqualIID(riid, &IID_IUnknown) ||
        IsEqualIID(riid, &IID_ICoreWebView2ClearBrowsingDataCompletedHandler)) {
        *ppvObject = This;
        This->lpVtbl->AddRef(This);
        return S_OK;
    }
    *ppvObject = NULL;
    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE ClearBrowsingDataCompletedHandler_AddRef(
    ICoreWebView2ClearBrowsingDataCompletedHandler* This) {
    ClearBrowsingDataCompletedHandler* handler = (ClearBrowsingDataCompletedHandler*)This;
    return InterlockedIncrement(&handler->refCount);
}

ULONG STDMETHODCALLTYPE ClearBrowsingDataCompletedHandler_Release(
    ICoreWebView2ClearBrowsingDataCompletedHandler* This) {
    ClearBrowsingDataCompletedHandler* handler = (ClearBrowsingDataCompletedHandler*)This;
    ULONG refCount = InterlockedDecrement(&handler->refCount);
    if (refCount == 0) {
        free(handler);
    }
    return refCount;
}

HRESULT STDMETHODCALLTYPE ClearBrowsingDataCompletedHandler_Invoke(
    ICoreWebView2ClearBrowsingDataCompletedHandler* This,
    HRESULT errorCode) {
    if (FAILED(errorCode)) {
        DebugPrint(L"[WARNING] Clear browsing data failed. HRESULT: 0x%08X\n", errorCode);
    } else {
        DebugPrint(L"[INFO] Cleared WebView2 cache data\n");
    }
    ReloadTargetPage();
    return S_OK;
}

// ExecuteScript handler implementation (fire-and-forget)
HRESULT STDMETHODCALLTYPE ExecuteScriptCompletedHandler_QueryInterface(
    ICoreWebView2ExecuteScriptCompletedHandler* This,
    REFIID riid, void** ppvObject) {
    if (IsEqualIID(riid, &IID_IUnknown) ||
        IsEqualIID(riid, &IID_ICoreWebView2ExecuteScriptCompletedHandler)) {
        *ppvObject = This;
        This->lpVtbl->AddRef(This);
        return S_OK;
    }
    *ppvObject = NULL;
    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE ExecuteScriptCompletedHandler_AddRef(
    ICoreWebView2ExecuteScriptCompletedHandler* This) {
    ExecuteScriptCompletedHandler* handler = (ExecuteScriptCompletedHandler*)This;
    return InterlockedIncrement(&handler->refCount);
}

ULONG STDMETHODCALLTYPE ExecuteScriptCompletedHandler_Release(
    ICoreWebView2ExecuteScriptCompletedHandler* This) {
    ExecuteScriptCompletedHandler* handler = (ExecuteScriptCompletedHandler*)This;
    ULONG refCount = InterlockedDecrement(&handler->refCount);
    if (refCount == 0) {
        free(handler);
    }
    return refCount;
}

HRESULT STDMETHODCALLTYPE ExecuteScriptCompletedHandler_Invoke(
    ICoreWebView2ExecuteScriptCompletedHandler* This,
    HRESULT errorCode, LPCWSTR resultObjectAsJson) {
    (void)resultObjectAsJson;  // Unused
    if (FAILED(errorCode)) {
        DebugPrint(L"[WARNING] ExecuteScript failed. HRESULT: 0x%08X\n", errorCode);
    }
    return S_OK;
}

HRESULT STDMETHODCALLTYPE TrySuspendCompletedHandler_QueryInterface(
    ICoreWebView2TrySuspendCompletedHandler* This,
    REFIID riid, void** ppvObject) {
    if (IsEqualIID(riid, &IID_IUnknown) ||
        IsEqualIID(riid, &IID_ICoreWebView2TrySuspendCompletedHandler)) {
        *ppvObject = This;
        This->lpVtbl->AddRef(This);
        return S_OK;
    }
    *ppvObject = NULL;
    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE TrySuspendCompletedHandler_AddRef(
    ICoreWebView2TrySuspendCompletedHandler* This) {
    TrySuspendCompletedHandler* handler = (TrySuspendCompletedHandler*)This;
    return InterlockedIncrement(&handler->refCount);
}

ULONG STDMETHODCALLTYPE TrySuspendCompletedHandler_Release(
    ICoreWebView2TrySuspendCompletedHandler* This) {
    TrySuspendCompletedHandler* handler = (TrySuspendCompletedHandler*)This;
    ULONG refCount = InterlockedDecrement(&handler->refCount);
    if (refCount == 0) {
        free(handler);
    }
    return refCount;
}

HRESULT STDMETHODCALLTYPE TrySuspendCompletedHandler_Invoke(
    ICoreWebView2TrySuspendCompletedHandler* This,
    HRESULT errorCode, BOOL result) {
    (void)This;
    InterlockedExchange(&g_webViewSuspendPending, FALSE);

    if (SUCCEEDED(errorCode) && result) {
        InterlockedExchange(&g_webViewSuspended, TRUE);
        DebugPrint(L"[INFO] WebView2 suspend request completed\n");
    } else {
        InterlockedExchange(&g_webViewSuspended, FALSE);
        DebugPrint(L"[WARNING] WebView2 suspend request failed. HRESULT: 0x%08X, result: %d\n",
                   errorCode, result);
    }

    if (InterlockedCompareExchange(&g_webViewDesiredActive, TRUE, TRUE) == TRUE) {
        BOOL desiredVisible =
            InterlockedCompareExchange(&g_webViewDesiredVisible, TRUE, TRUE) == TRUE;
        ResumeMainWebViewRuntime();
        SetMainWebViewControllerVisible(desiredVisible);
    }

    return S_OK;
}

// Liveness ping handler: any answer at all (even an error code) proves the
// runtime is still talking to us. No answer within POWER_RESUME_LIVENESS_MS
// means it is wedged; CheckMainWebViewLiveness handles that case.
HRESULT STDMETHODCALLTYPE LivenessPingHandler_QueryInterface(
    ICoreWebView2ExecuteScriptCompletedHandler* This,
    REFIID riid, void** ppvObject) {
    if (IsEqualIID(riid, &IID_IUnknown) ||
        IsEqualIID(riid, &IID_ICoreWebView2ExecuteScriptCompletedHandler)) {
        *ppvObject = This;
        This->lpVtbl->AddRef(This);
        return S_OK;
    }
    *ppvObject = NULL;
    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE LivenessPingHandler_AddRef(
    ICoreWebView2ExecuteScriptCompletedHandler* This) {
    return InterlockedIncrement(&((LivenessPingHandler*)This)->refCount);
}

ULONG STDMETHODCALLTYPE LivenessPingHandler_Release(
    ICoreWebView2ExecuteScriptCompletedHandler* This) {
    ULONG refCount = InterlockedDecrement(&((LivenessPingHandler*)This)->refCount);
    if (refCount == 0) {
        free(This);
    }
    return refCount;
}

HRESULT STDMETHODCALLTYPE LivenessPingHandler_Invoke(
    ICoreWebView2ExecuteScriptCompletedHandler* This,
    HRESULT errorCode, LPCWSTR resultObjectAsJson) {
    (void)This; (void)errorCode; (void)resultObjectAsJson;
    InterlockedExchange(&g_webViewPingOutstanding, FALSE);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE NavCompletedHandler_QueryInterface(
    ICoreWebView2NavigationCompletedEventHandler* This,
    REFIID riid, void** ppvObject) {
    if (IsEqualIID(riid, &IID_IUnknown) ||
        IsEqualIID(riid, &IID_ICoreWebView2NavigationCompletedEventHandler)) {
        *ppvObject = This;
        This->lpVtbl->AddRef(This);
        return S_OK;
    }
    *ppvObject = NULL;
    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE NavCompletedHandler_AddRef(
    ICoreWebView2NavigationCompletedEventHandler* This) {
    NavCompletedHandler* handler = (NavCompletedHandler*)This;
    return InterlockedIncrement(&handler->refCount);
}

ULONG STDMETHODCALLTYPE NavCompletedHandler_Release(
    ICoreWebView2NavigationCompletedEventHandler* This) {
    NavCompletedHandler* handler = (NavCompletedHandler*)This;
    ULONG refCount = InterlockedDecrement(&handler->refCount);
    if (refCount == 0) {
        free(handler);
    }
    return refCount;
}

HRESULT STDMETHODCALLTYPE NavCompletedHandler_Invoke(
    ICoreWebView2NavigationCompletedEventHandler* This,
    ICoreWebView2* sender, ICoreWebView2NavigationCompletedEventArgs* args) {
    (void)This;
    (void)sender;
    (void)args;
    OnMainNavigationCompleted();
    return S_OK;
}

static void RegisterMainNavigationCompletedHandler(ICoreWebView2* webview2) {
    if (!webview2) return;

    NavCompletedHandler* handler =
        (NavCompletedHandler*)calloc(1, sizeof(NavCompletedHandler));
    if (!handler) return;

    static ICoreWebView2NavigationCompletedEventHandlerVtbl navVtbl = {
        NavCompletedHandler_QueryInterface,
        NavCompletedHandler_AddRef,
        NavCompletedHandler_Release,
        NavCompletedHandler_Invoke
    };
    handler->lpVtbl = &navVtbl;
    handler->refCount = 1;

    EventRegistrationToken token;
    HRESULT hr = webview2->lpVtbl->add_NavigationCompleted(
        webview2, (ICoreWebView2NavigationCompletedEventHandler*)handler, &token);
    if (FAILED(hr)) {
        DebugPrint(L"[WARNING] add_NavigationCompleted failed. HRESULT: 0x%08X\n", hr);
    }

    handler->lpVtbl->Release((ICoreWebView2NavigationCompletedEventHandler*)handler);
}

// New-window requests (target="_blank", window.open, "open in new tab"):
// when the setting is enabled, suppress the default WebView2 popup and hand
// the URL to the system default browser instead. Only http(s) URLs are passed
// to the shell so a page cannot make the app launch other schemes. Requests
// without a usable URL (e.g. about:blank popups that get scripted afterwards)
// fall through to the default popup, where such flows still work.
typedef struct {
    ICoreWebView2NewWindowRequestedEventHandlerVtbl* lpVtbl;
    LONG refCount;
} NewWindowHandler;

static HRESULT STDMETHODCALLTYPE NewWindowHandler_QueryInterface(
    ICoreWebView2NewWindowRequestedEventHandler* This,
    REFIID riid, void** ppvObject) {
    if (IsEqualIID(riid, &IID_IUnknown) ||
        IsEqualIID(riid, &IID_ICoreWebView2NewWindowRequestedEventHandler)) {
        *ppvObject = This;
        This->lpVtbl->AddRef(This);
        return S_OK;
    }
    *ppvObject = NULL;
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE NewWindowHandler_AddRef(
    ICoreWebView2NewWindowRequestedEventHandler* This) {
    return InterlockedIncrement(&((NewWindowHandler*)This)->refCount);
}

static ULONG STDMETHODCALLTYPE NewWindowHandler_Release(
    ICoreWebView2NewWindowRequestedEventHandler* This) {
    ULONG refCount = InterlockedDecrement(&((NewWindowHandler*)This)->refCount);
    if (refCount == 0) free(This);
    return refCount;
}

static HRESULT STDMETHODCALLTYPE NewWindowHandler_Invoke(
    ICoreWebView2NewWindowRequestedEventHandler* This,
    ICoreWebView2* sender, ICoreWebView2NewWindowRequestedEventArgs* args) {
    (void)This; (void)sender;

    if (InterlockedCompareExchange(&g_openNewWindowsExternally, TRUE, TRUE) != TRUE) {
        return S_OK;
    }

    LPWSTR uri = NULL;
    if (FAILED(args->lpVtbl->get_Uri(args, &uri)) || !uri) {
        return S_OK;
    }

    if (_wcsnicmp(uri, L"https://", 8) == 0 || _wcsnicmp(uri, L"http://", 7) == 0) {
        args->lpVtbl->put_Handled(args, TRUE);
        ShellExecuteW(NULL, L"open", uri, NULL, NULL, SW_SHOWNORMAL);
        DebugPrint(L"[INFO] Opened new-window link in default browser: %s\n", uri);
    }

    CoTaskMemFree(uri);
    return S_OK;
}

static void RegisterMainNewWindowRequestedHandler(ICoreWebView2* webview2) {
    if (!webview2) return;

    NewWindowHandler* handler = (NewWindowHandler*)calloc(1, sizeof(NewWindowHandler));
    if (!handler) return;

    static ICoreWebView2NewWindowRequestedEventHandlerVtbl newWindowVtbl = {
        NewWindowHandler_QueryInterface,
        NewWindowHandler_AddRef,
        NewWindowHandler_Release,
        NewWindowHandler_Invoke
    };
    handler->lpVtbl = &newWindowVtbl;
    handler->refCount = 1;

    EventRegistrationToken token;
    HRESULT hr = webview2->lpVtbl->add_NewWindowRequested(
        webview2, (ICoreWebView2NewWindowRequestedEventHandler*)handler, &token);
    if (FAILED(hr)) {
        DebugPrint(L"[WARNING] add_NewWindowRequested failed. HRESULT: 0x%08X\n", hr);
    }

    handler->lpVtbl->Release((ICoreWebView2NewWindowRequestedEventHandler*)handler);
}

// Process failures (crashed/killed renderer, dead browser process, hung
// page) previously went unnoticed, leaving the container permanently blank.
// Renderer-level failures are repaired in place with a reload; a dead
// browser process triggers a full rebuild of the WebView.
typedef struct {
    ICoreWebView2ProcessFailedEventHandlerVtbl* lpVtbl;
    LONG refCount;
} ProcessFailedHandler;

static HRESULT STDMETHODCALLTYPE ProcessFailedHandler_QueryInterface(
    ICoreWebView2ProcessFailedEventHandler* This,
    REFIID riid, void** ppvObject) {
    if (IsEqualIID(riid, &IID_IUnknown) ||
        IsEqualIID(riid, &IID_ICoreWebView2ProcessFailedEventHandler)) {
        *ppvObject = This;
        This->lpVtbl->AddRef(This);
        return S_OK;
    }
    *ppvObject = NULL;
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE ProcessFailedHandler_AddRef(
    ICoreWebView2ProcessFailedEventHandler* This) {
    return InterlockedIncrement(&((ProcessFailedHandler*)This)->refCount);
}

static ULONG STDMETHODCALLTYPE ProcessFailedHandler_Release(
    ICoreWebView2ProcessFailedEventHandler* This) {
    ULONG refCount = InterlockedDecrement(&((ProcessFailedHandler*)This)->refCount);
    if (refCount == 0) free(This);
    return refCount;
}

static HRESULT STDMETHODCALLTYPE ProcessFailedHandler_Invoke(
    ICoreWebView2ProcessFailedEventHandler* This,
    ICoreWebView2* sender, ICoreWebView2ProcessFailedEventArgs* args) {
    (void)This; (void)sender;

    COREWEBVIEW2_PROCESS_FAILED_KIND kind =
        COREWEBVIEW2_PROCESS_FAILED_KIND_BROWSER_PROCESS_EXITED;
    if (args) args->lpVtbl->get_ProcessFailedKind(args, &kind);
    DebugPrint(L"[WARNING] WebView2 process failure, kind=%d\n", (int)kind);

    switch (kind) {
        case COREWEBVIEW2_PROCESS_FAILED_KIND_BROWSER_PROCESS_EXITED:
            // Everything behind the controller is gone; rebuild from scratch
            // (the BrowserProcessExited event posts the same message, the
            // handler dedupes).
            if (g_hwnd) PostMessageW(g_hwnd, WM_APP_WEBVIEW_RECREATE, 0, 0);
            break;

        case COREWEBVIEW2_PROCESS_FAILED_KIND_RENDER_PROCESS_EXITED:
        case COREWEBVIEW2_PROCESS_FAILED_KIND_RENDER_PROCESS_UNRESPONSIVE:
            // The browser process is fine; only the page died. Reload it in
            // place, falling back to a fresh navigation.
            if (g_webView) {
                ResumeMainWebViewRuntime();
                if (FAILED(g_webView->lpVtbl->Reload(g_webView))) {
                    ReloadTargetPage();
                }
            }
            break;

        default:
            // GPU/utility/frame processes are restarted by the runtime.
            break;
    }

    return S_OK;
}

static void RegisterMainProcessFailedHandler(ICoreWebView2* webview2) {
    if (!webview2) return;

    ProcessFailedHandler* handler =
        (ProcessFailedHandler*)calloc(1, sizeof(ProcessFailedHandler));
    if (!handler) return;

    static ICoreWebView2ProcessFailedEventHandlerVtbl failVtbl = {
        ProcessFailedHandler_QueryInterface,
        ProcessFailedHandler_AddRef,
        ProcessFailedHandler_Release,
        ProcessFailedHandler_Invoke
    };
    handler->lpVtbl = &failVtbl;
    handler->refCount = 1;

    EventRegistrationToken token;
    HRESULT hr = webview2->lpVtbl->add_ProcessFailed(
        webview2, (ICoreWebView2ProcessFailedEventHandler*)handler, &token);
    if (FAILED(hr)) {
        DebugPrint(L"[WARNING] add_ProcessFailed failed. HRESULT: 0x%08X\n", hr);
    }

    handler->lpVtbl->Release((ICoreWebView2ProcessFailedEventHandler*)handler);
}

// Helper to execute JavaScript in WebView2
void ExecuteJavaScript(const wchar_t* js) {
    if (!g_webView || !js || js[0] == L'\0') return;

    ExecuteScriptCompletedHandler* handler =
        (ExecuteScriptCompletedHandler*)calloc(1, sizeof(ExecuteScriptCompletedHandler));
    if (!handler) return;

    static ICoreWebView2ExecuteScriptCompletedHandlerVtbl executeScriptVtbl = {
        ExecuteScriptCompletedHandler_QueryInterface,
        ExecuteScriptCompletedHandler_AddRef,
        ExecuteScriptCompletedHandler_Release,
        ExecuteScriptCompletedHandler_Invoke
    };

    handler->lpVtbl = &executeScriptVtbl;
    handler->refCount = 1;

    HRESULT hr = g_webView->lpVtbl->ExecuteScript(
        g_webView, js, (ICoreWebView2ExecuteScriptCompletedHandler*)handler);
    if (FAILED(hr)) {
        DebugPrint(L"[WARNING] ExecuteScript call failed. HRESULT: 0x%08X\n", hr);
    }

    handler->lpVtbl->Release((ICoreWebView2ExecuteScriptCompletedHandler*)handler);
}

static BOOL IsWebViewReady(void) {
    return InterlockedCompareExchange(&g_isInitialized, TRUE, TRUE) == TRUE && g_webView != NULL;
}

static ICoreWebView2_3* QueryMainWebView3(void) {
    if (!g_webView) return NULL;

    ICoreWebView2_3* webView3 = NULL;
    HRESULT hr = g_webView->lpVtbl->QueryInterface(
        g_webView, &IID_ICoreWebView2_3, (void**)&webView3);
    if (FAILED(hr) || !webView3) {
        DebugPrint(L"[WARNING] WebView2 suspend/resume API is not available. HRESULT: 0x%08X\n", hr);
        return NULL;
    }

    return webView3;
}

static void SyncMainWebViewBounds(void) {
    if (!g_webViewController || !g_hwnd) return;

    RECT bounds;
    GetClientRect(g_hwnd, &bounds);
    g_webViewController->lpVtbl->put_Bounds(g_webViewController, bounds);
}

static void SetMainWebViewControllerVisible(BOOL visible) {
    if (!g_webViewController) return;

    if (visible) {
        SyncMainWebViewBounds();
    }
    g_webViewController->lpVtbl->put_IsVisible(g_webViewController, visible);
}

static void ResumeMainWebViewRuntime(void) {
    LONG previousDesired = InterlockedExchange(&g_webViewDesiredActive, TRUE);
    if (!IsWebViewReady() || !g_webViewController) return;

    BOOL needsResume =
        previousDesired == FALSE ||
        InterlockedCompareExchange(&g_webViewSuspendPending, FALSE, FALSE) == TRUE ||
        InterlockedCompareExchange(&g_webViewSuspended, FALSE, FALSE) == TRUE;

    if (!needsResume) return;

    ICoreWebView2_3* webView3 = QueryMainWebView3();
    if (!webView3) return;

    HRESULT hr = webView3->lpVtbl->Resume(webView3);
    webView3->lpVtbl->Release(webView3);

    if (SUCCEEDED(hr)) {
        InterlockedExchange(&g_webViewSuspendPending, FALSE);
        InterlockedExchange(&g_webViewSuspended, FALSE);
        InterlockedExchange(&g_resumeFailureCount, 0);
        return;
    }

    // Keep the suspended flags set so every later activation retries the
    // resume. A runtime that stays unresumable (seen after the machine comes
    // back from hibernation) gets torn down and rebuilt instead of leaving a
    // frozen, white page on screen.
    DebugPrint(L"[WARNING] WebView2 resume failed. HRESULT: 0x%08X\n", hr);
    LONG failures = InterlockedIncrement(&g_resumeFailureCount);
    if (failures >= RESUME_FAILURE_RECREATE_THRESHOLD && g_hwnd) {
        InterlockedExchange(&g_resumeFailureCount, 0);
        PostMessageW(g_hwnd, WM_APP_WEBVIEW_RECREATE, 0, 0);
    }
}

static void SuspendMainWebViewRuntime(void) {
    InterlockedExchange(&g_webViewDesiredActive, FALSE);
    if (!IsWebViewReady() || !g_webViewController) return;

    if (InterlockedCompareExchange(&g_webViewSuspendPending, FALSE, FALSE) == TRUE ||
        InterlockedCompareExchange(&g_webViewSuspended, FALSE, FALSE) == TRUE) {
        return;
    }

    ICoreWebView2_3* webView3 = QueryMainWebView3();
    if (!webView3) return;

    TrySuspendCompletedHandler* handler =
        (TrySuspendCompletedHandler*)calloc(1, sizeof(TrySuspendCompletedHandler));
    if (!handler) {
        webView3->lpVtbl->Release(webView3);
        return;
    }

    static ICoreWebView2TrySuspendCompletedHandlerVtbl suspendVtbl = {
        TrySuspendCompletedHandler_QueryInterface,
        TrySuspendCompletedHandler_AddRef,
        TrySuspendCompletedHandler_Release,
        TrySuspendCompletedHandler_Invoke
    };

    handler->lpVtbl = &suspendVtbl;
    handler->refCount = 1;

    InterlockedExchange(&g_webViewSuspendPending, TRUE);
    HRESULT hr = webView3->lpVtbl->TrySuspend(
        webView3, (ICoreWebView2TrySuspendCompletedHandler*)handler);
    if (FAILED(hr)) {
        InterlockedExchange(&g_webViewSuspendPending, FALSE);
        DebugPrint(L"[WARNING] WebView2 TrySuspend call failed. HRESULT: 0x%08X\n", hr);
    }

    handler->lpVtbl->Release((ICoreWebView2TrySuspendCompletedHandler*)handler);
    webView3->lpVtbl->Release(webView3);
}

// Bring the WebView to the foreground state: runtime resumed + rendered, and
// the controller sized to the now-visible host window.
static void ActivateMainWebView(void) {
    InterlockedExchange(&g_webViewPrewarmActive, FALSE);
    if (g_hwnd) {
        KillTimer(g_hwnd, ID_TIMER_WEBVIEW_PREWARM);
        KillTimer(g_hwnd, ID_TIMER_WEBVIEW_PRELOAD);
    }

    InterlockedExchange(&g_webViewDesiredVisible, TRUE);
    ResumeMainWebViewRuntime();
    SetMainWebViewControllerVisible(TRUE);
}

// Move the WebView to the background (host window hidden).
//
// When the sleep setting is enabled and the initial preload has finished we
// stop rendering (IsVisible = FALSE) and suspend the runtime to save CPU.
// Otherwise — sleeping disabled (the default), or the preload is still in
// progress — we keep the page warm: rendering stays on so the off-screen page
// is ready to display instantly. The host window is hidden either way, so
// nothing is shown to the user.
static void DeactivateMainWebView(void) {
    BOOL sleepEnabled = InterlockedCompareExchange(&g_sleepWhenInactive, TRUE, TRUE) == TRUE;
    BOOL preloaded = InterlockedCompareExchange(&g_initialPreloadComplete, TRUE, TRUE) == TRUE;
    BOOL recovery = InterlockedCompareExchange(&g_powerResumePending, TRUE, TRUE) == TRUE;
    BOOL prewarming = InterlockedCompareExchange(&g_webViewPrewarmActive, TRUE, TRUE) == TRUE;

    // Never put the page to sleep while post-resume recovery is unverified
    // (a possibly-broken page must not be frozen into a suspend snapshot) or
    // while a tray-hover prewarm is keeping it warm. The steady-state hidden
    // ticks used to cancel both within 250 ms.
    if (sleepEnabled && preloaded && !recovery && !prewarming) {
        InterlockedExchange(&g_webViewPrewarmActive, FALSE);
        if (g_hwnd) {
            KillTimer(g_hwnd, ID_TIMER_WEBVIEW_PREWARM);
        }
        InterlockedExchange(&g_webViewDesiredVisible, FALSE);
        SetMainWebViewControllerVisible(FALSE);
        SuspendMainWebViewRuntime();
    } else {
        InterlockedExchange(&g_webViewDesiredVisible, TRUE);
        ResumeMainWebViewRuntime();
        SetMainWebViewControllerVisible(TRUE);
    }
}

// Pre-emptively wake a suspended WebView when the user hovers the tray icon, so
// a live, rendered copy of the page is ready before they open the window. Only
// relevant while the sleep setting is enabled — otherwise nothing is suspended.
static void PrewarmMainWebView(void) {
    if (!g_hwnd) return;
    if (InterlockedCompareExchange(&g_sleepWhenInactive, TRUE, TRUE) != TRUE) return;
    if (!IsWebViewReady()) return;
    if (IsWindowActuallyVisible(g_hwnd)) return;

    InterlockedExchange(&g_webViewPrewarmActive, TRUE);
    InterlockedExchange(&g_webViewDesiredVisible, TRUE);
    ResumeMainWebViewRuntime();
    SetMainWebViewControllerVisible(TRUE);  // render warm while the host stays hidden

    SetTimer(g_hwnd, ID_TIMER_WEBVIEW_PREWARM, WEBVIEW_PREWARM_MS, NULL);
    DebugPrint(L"[INFO] WebView2 prewarmed (warm render) from tray hover for %d ms\n", WEBVIEW_PREWARM_MS);
}

// Ask the runtime to run a trivial script; the completion handler clearing
// g_webViewPingOutstanding is the "pong". A synchronous call failure is
// treated as no answer (the flag stays set) so the liveness check escalates.
static void SendMainWebViewLivenessPing(void) {
    if (!IsWebViewReady()) return;
    if (InterlockedCompareExchange(&g_webViewPingOutstanding, TRUE, TRUE) == TRUE) return;

    LivenessPingHandler* handler =
        (LivenessPingHandler*)calloc(1, sizeof(LivenessPingHandler));
    if (!handler) return;

    static ICoreWebView2ExecuteScriptCompletedHandlerVtbl pingVtbl = {
        LivenessPingHandler_QueryInterface,
        LivenessPingHandler_AddRef,
        LivenessPingHandler_Release,
        LivenessPingHandler_Invoke
    };
    handler->lpVtbl = &pingVtbl;
    handler->refCount = 1;

    InterlockedExchange(&g_webViewPingOutstanding, TRUE);
    HRESULT hr = g_webView->lpVtbl->ExecuteScript(g_webView, L"1",
        (ICoreWebView2ExecuteScriptCompletedHandler*)handler);
    if (FAILED(hr)) {
        DebugPrint(L"[WARNING] Liveness ping could not be sent. HRESULT: 0x%08X\n", hr);
    }

    handler->lpVtbl->Release((ICoreWebView2ExecuteScriptCompletedHandler*)handler);
}

// After the machine resumes from sleep/hibernate the GPU-side composition
// surfaces backing the WebView can be gone and the runtime may stop answering
// altogether; a page in that state presents as a permanently white container.
// The graphics stack can take many seconds to come back after hibernate, so
// this runs as a short sequence rather than a single attempt: wake the page,
// re-assert the bounds, drop and re-add the visual tree, then verify with a
// script ping. CheckMainWebViewLiveness re-arms the kick on silence, or
// rebuilds the WebView when the runtime keeps ignoring us. Until recovery
// completes, DeactivateMainWebView keeps the page warm so a possibly-broken
// page is never frozen into a suspend snapshot.
static void KickWebViewAfterPowerResume(HWND hwnd) {
    if (!IsWebViewReady() || !g_webViewController) {
        // Nothing to kick; force the liveness check down the rebuild path.
        InterlockedExchange(&g_webViewPingOutstanding, TRUE);
        SetTimer(hwnd, ID_TIMER_WEBVIEW_LIVENESS, POWER_RESUME_LIVENESS_MS, NULL);
        return;
    }

    DebugPrint(L"[INFO] System resumed; refreshing WebView2 composition (attempt %d)\n",
               g_powerKickCount);
    ResumeMainWebViewRuntime();
    SyncMainWebViewBounds();
    g_webViewController->lpVtbl->put_IsVisible(g_webViewController, FALSE);
    g_webViewController->lpVtbl->put_IsVisible(g_webViewController, TRUE);
    g_webViewController->lpVtbl->NotifyParentWindowPositionChanged(g_webViewController);

    SendMainWebViewLivenessPing();
    SetTimer(hwnd, ID_TIMER_WEBVIEW_LIVENESS, POWER_RESUME_LIVENESS_MS, NULL);
}

// Runs POWER_RESUME_LIVENESS_MS after each kick. A cleared ping flag means
// the runtime answered: recovery is done and the normal visibility/sleep
// state can settle. Silence means the runtime is wedged: retry the kick a
// few times, then tear the WebView down and rebuild it.
static void CheckMainWebViewLiveness(HWND hwnd) {
    if (InterlockedCompareExchange(&g_powerResumePending, TRUE, TRUE) != TRUE) return;

    if (InterlockedCompareExchange(&g_webViewPingOutstanding, TRUE, TRUE) != TRUE) {
        InterlockedExchange(&g_powerResumePending, FALSE);
        g_powerKickCount = 0;
        DebugPrint(L"[INFO] WebView2 responsive after power resume\n");
        if (IsWindowActuallyVisible(hwnd)) {
            ActivateMainWebView();
        } else {
            DeactivateMainWebView();
        }
        return;
    }

    // Allow the next kick to ping again (a late pong is harmless).
    InterlockedExchange(&g_webViewPingOutstanding, FALSE);

    if (IsWebViewReady() && g_powerKickCount < POWER_RESUME_MAX_KICKS) {
        DebugPrint(L"[WARNING] WebView2 not answering after power resume; retrying\n");
        SetTimer(hwnd, ID_TIMER_POWER_RESUME, POWER_RESUME_KICK_RETRY_MS, NULL);
        return;
    }

    DebugPrint(L"[WARNING] WebView2 unresponsive after power resume; forcing rebuild\n");
    InterlockedExchange(&g_powerResumePending, FALSE);
    g_powerKickCount = 0;
    PostMessageW(hwnd, WM_APP_WEBVIEW_RECREATE, 0, 0);
}

// Called when a navigation completes. The first completion marks the initial
// preload as done, after which it is safe to suspend on hide without cutting a
// page load short.
static void OnMainNavigationCompleted(void) {
    InterlockedExchange(&g_initialPreloadComplete, TRUE);

    if (!g_hwnd) return;
    if (IsWindowActuallyVisible(g_hwnd)) return;  // shown: stay active
    if (InterlockedCompareExchange(&g_webViewPrewarmActive, TRUE, TRUE) == TRUE) return;  // hover prewarm in progress

    if (InterlockedCompareExchange(&g_sleepWhenInactive, TRUE, TRUE) == TRUE) {
        // Let the freshly-loaded page render for a short moment before
        // suspending, so the suspended snapshot is complete and resumes
        // instantly when the user opens or hovers.
        SetTimer(g_hwnd, ID_TIMER_WEBVIEW_PRELOAD, WEBVIEW_PRELOAD_SETTLE_MS, NULL);
    } else {
        // Sleep disabled: keep the page warm and running for instant opens.
        DeactivateMainWebView();
    }
}

// Data structure for occlusion check enumeration
typedef struct {
    HWND targetHwnd;
    HRGN visibleRgn;
} OcclusionCheckData;

// Callback for EnumWindows - subtracts each window above target from visible region
static BOOL CALLBACK OcclusionEnumProc(HWND hwnd, LPARAM lParam) {
    OcclusionCheckData* data = (OcclusionCheckData*)lParam;

    // Stop when we reach our own window (windows below us don't occlude us)
    if (hwnd == data->targetHwnd) {
        return FALSE;
    }

    // Skip invisible or minimized windows
    if (!IsWindowVisible(hwnd) || IsIconic(hwnd)) {
        return TRUE;
    }

    // Skip DWM-cloaked windows: suspended UWP apps, the lock-screen host and
    // ghost ApplicationFrameHost shells report IsWindowVisible=TRUE while
    // drawing nothing. Counting them as occluders makes the app believe the
    // window is covered and suspend a WebView the user is looking at.
    DWORD cloaked = 0;
    if (SUCCEEDED(DwmGetWindowAttribute(hwnd, DWMWA_CLOAKED, &cloaked,
                                        sizeof(cloaked))) && cloaked != 0) {
        return TRUE;
    }

    // Skip windows with no area
    RECT windowRect;
    if (!GetWindowRect(hwnd, &windowRect)) {
        return TRUE;
    }
    if (windowRect.right <= windowRect.left || windowRect.bottom <= windowRect.top) {
        return TRUE;
    }

    // Subtract this window's rect from our visible region
    HRGN windowRgn = CreateRectRgnIndirect(&windowRect);
    if (windowRgn) {
        CombineRgn(data->visibleRgn, data->visibleRgn, windowRgn, RGN_DIFF);
        DeleteObject(windowRgn);
    }

    return TRUE;
}

// Check if ANY part of the window is visible (not fully covered by other windows)
static BOOL IsWindowActuallyVisible(HWND hwnd) {
    if (!hwnd) return FALSE;
    if (!IsWindowVisible(hwnd)) return FALSE;
    if (IsIconic(hwnd)) return FALSE;

    RECT ourRect;
    if (!GetWindowRect(hwnd, &ourRect)) return FALSE;
    if (!MonitorFromRect(&ourRect, MONITOR_DEFAULTTONULL)) {
        // Monitors are still being re-enumerated (common right after resume
        // from hibernate, or during docking changes). Assume visible rather
        // than suspending a WebView the user may be looking at.
        return TRUE;
    }

    // Create a region representing our window
    HRGN visibleRgn = CreateRectRgnIndirect(&ourRect);
    if (!visibleRgn) return FALSE;

    OcclusionCheckData data;
    data.targetHwnd = hwnd;
    data.visibleRgn = visibleRgn;

    // EnumWindows enumerates top-level windows in z-order (top to bottom)
    // We subtract each window above us until we reach our own window
    EnumWindows(OcclusionEnumProc, (LPARAM)&data);

    // Check if any part of our window is still visible
    RECT boundingBox;
    int rgnType = GetRgnBox(visibleRgn, &boundingBox);
    DeleteObject(visibleRgn);

    // NULLREGION means our window is completely covered
    return (rgnType != NULLREGION);
}

static void StartVisibilityTimer(HWND hwnd) {
    SetTimer(hwnd, ID_TIMER_VISIBILITY_CHECK, VISIBILITY_CHECK_INTERVAL_MS, NULL);
    DebugPrint(L"[INFO] Started visibility check timer\n");
}

static void StopVisibilityTimer(HWND hwnd) {
    KillTimer(hwnd, ID_TIMER_VISIBILITY_CHECK);
    DebugPrint(L"[INFO] Stopped visibility check timer\n");
}

static void UpdateJsVisibilityState(HWND hwnd) {
    if (!IsWebViewReady()) return;

    JsVisibility newState = IsWindowActuallyVisible(hwnd) ? JS_VISIBILITY_SHOWN : JS_VISIBILITY_HIDDEN;
    if (newState == g_jsVisibility) {
        if (newState == JS_VISIBILITY_SHOWN) {
            ActivateMainWebView();
        } else {
            DeactivateMainWebView();
        }
        return;
    }

    if (newState == JS_VISIBILITY_SHOWN) {
        ActivateMainWebView();
    }

    g_jsVisibility = newState;
    if (newState == JS_VISIBILITY_SHOWN) {
        if (g_config.onShowJs[0] != L'\0') {
            ExecuteJavaScript(g_config.onShowJs);
            DebugPrint(L"[INFO] Executed onShowJs (window visible)\n");
        }
    } else {
        if (g_config.onHideJs[0] != L'\0') {
            ExecuteJavaScript(g_config.onHideJs);
            DebugPrint(L"[INFO] Executed onHideJs (window fully covered/hidden)\n");
        }
        DeactivateMainWebView();
    }
}

// Display helpers
void CaptureDisplaySettings(void) {
    HDC hdcScreen = GetDC(NULL);
    g_lastScreenWidth = GetSystemMetrics(SM_CXSCREEN);
    g_lastScreenHeight = GetSystemMetrics(SM_CYSCREEN);
    g_lastDpiX = (float)GetDeviceCaps(hdcScreen, LOGPIXELSX);
    g_lastDpiY = (float)GetDeviceCaps(hdcScreen, LOGPIXELSY);
    ReleaseDC(NULL, hdcScreen);
    DebugPrint(L"[INFO] Captured display: %dx%d @ %.1fx%.1f DPI\n",
               g_lastScreenWidth, g_lastScreenHeight, g_lastDpiX, g_lastDpiY);
}

BOOL HasDisplaySettingsChanged(void) {
    HDC hdcScreen = GetDC(NULL);
    int currentWidth = GetSystemMetrics(SM_CXSCREEN);
    int currentHeight = GetSystemMetrics(SM_CYSCREEN);
    float currentDpiX = (float)GetDeviceCaps(hdcScreen, LOGPIXELSX);
    float currentDpiY = (float)GetDeviceCaps(hdcScreen, LOGPIXELSY);
    ReleaseDC(NULL, hdcScreen);

    BOOL changed = (currentWidth != g_lastScreenWidth) ||
                   (currentHeight != g_lastScreenHeight) ||
                   (fabs(currentDpiX - g_lastDpiX) > 0.1f) ||
                   (fabs(currentDpiY - g_lastDpiY) > 0.1f);

    if (changed) {
        DebugPrint(L"[INFO] Display settings changed: %dx%d @ %.1fx%.1f DPI -> %dx%d @ %.1fx%.1f DPI\n",
                   g_lastScreenWidth, g_lastScreenHeight, g_lastDpiX, g_lastDpiY,
                   currentWidth, currentHeight, currentDpiX, currentDpiY);
    }
    return changed;
}

static void ResetTargetPageIfNeeded(void) {
    if (!g_webView || !g_initialUrl[0]) return;

    LPWSTR currentUrl = NULL;
    HRESULT hr = g_webView->lpVtbl->get_Source(g_webView, &currentUrl);
    if (SUCCEEDED(hr) && currentUrl) {
        if (wcscmp(currentUrl, g_initialUrl) != 0) {
            g_webView->lpVtbl->Navigate(g_webView, g_initialUrl);
            DebugPrint(L"[INFO] Reset URL to initial on show: %s (was: %s)\n", g_initialUrl, currentUrl);
        } else {
            DebugPrint(L"[INFO] URL unchanged, skipping navigation on show\n");
        }
        CoTaskMemFree(currentUrl);
        return;
    }

    g_webView->lpVtbl->Navigate(g_webView, g_initialUrl);
    DebugPrint(L"[INFO] Reset URL to initial on show (couldn't check current): %s\n", g_initialUrl);
}

// Compute the centered, 90%-of-work-area rectangle used for the main window.
static void GetTargetWindowRect(int* x, int* y, int* w, int* h) {
    RECT workArea;
    SystemParametersInfoW(SPI_GETWORKAREA, 0, &workArea, 0);

    int workWidth = workArea.right - workArea.left;
    int workHeight = workArea.bottom - workArea.top;

    int windowWidth = (int)(workWidth * WINDOW_SIZE_PERCENTAGE);
    int windowHeight = (int)(workHeight * WINDOW_SIZE_PERCENTAGE);

    *w = windowWidth;
    *h = windowHeight;
    *x = workArea.left + (workWidth - windowWidth) / 2;
    *y = workArea.top + (workHeight - windowHeight) / 2;
}

// Window management
void ShowMainWindow(void) {
    if (!g_hwnd) return;

    RebuildMainWebViewIfDead();

    int x, y, windowWidth, windowHeight;
    GetTargetWindowRect(&x, &y, &windowWidth, &windowHeight);

    SetWindowPos(g_hwnd, HWND_TOP, x, y, windowWidth, windowHeight,
                 SWP_SHOWWINDOW | SWP_FRAMECHANGED);
    ShowWindow(g_hwnd, SW_RESTORE);
    SetForegroundWindow(g_hwnd);

    ActivateMainWebView();

    if (InterlockedExchange(&g_resetUrlOnNextShow, FALSE) == TRUE) {
        if (IsWebViewReady()) {
            ResetTargetPageIfNeeded();
        } else {
            InterlockedExchange(&g_resetUrlOnNextShow, TRUE);
        }
    }

    // Start polling for visibility changes while window is shown
    StartVisibilityTimer(g_hwnd);
    UpdateJsVisibilityState(g_hwnd);

    DebugPrint(L"[INFO] Main window shown at %dx%d, size %dx%d\n", x, y, windowWidth, windowHeight);
}

void HideMainWindow(void) {
	if (!g_hwnd) return;

    // Stop visibility polling when window is hidden
    StopVisibilityTimer(g_hwnd);

    ShowWindow(g_hwnd, SW_HIDE);
    UpdateJsVisibilityState(g_hwnd);
    InterlockedExchange(&g_resetUrlOnNextShow, TRUE);

    DebugPrint(L"[INFO] Main window hidden\n");
}

// Tray icon functions
void CreateTrayIcon(HWND hwnd) {
    ZeroMemory(&g_nid, sizeof(g_nid));
    g_nid.cbSize = sizeof(NOTIFYICONDATAW);
    g_nid.hWnd = hwnd;
    g_nid.uID = TRAY_ICON_ID;
    g_nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    g_nid.uCallbackMessage = WM_TRAYICON;
    
    // Load icon from resources with proper DPI scaling
    HDC hdcScreen = GetDC(NULL);
    int dpiX = GetDeviceCaps(hdcScreen, LOGPIXELSX);
    ReleaseDC(NULL, hdcScreen);
    
    int iconSize = (dpiX >= 120) ? 32 : 16;
    g_nid.hIcon = (HICON)LoadImageW(g_hInstance, MAKEINTRESOURCEW(IDI_TRAYICON),
                                     IMAGE_ICON, iconSize, iconSize, LR_DEFAULTCOLOR);

    if (!g_nid.hIcon) {
        g_nid.hIcon = LoadIconW(NULL, (LPCWSTR)IDI_APPLICATION);
    }
    
    wcscpy_s(g_nid.szTip, sizeof(g_nid.szTip)/sizeof(wchar_t), g_config.windowTitle);
    Shell_NotifyIconW(NIM_ADD, &g_nid);
    DebugPrint(L"[INFO] Tray icon created with size %dx%d for DPI %d\n", iconSize, iconSize, dpiX);
}

void RefreshTrayIcon(void) {
    if (!g_nid.hWnd) return;
    
    // Delete old icon
    if (g_nid.hIcon) {
        DestroyIcon(g_nid.hIcon);
        g_nid.hIcon = NULL;
    }
    Shell_NotifyIconW(NIM_DELETE, &g_nid);
    
    // Recreate with new DPI settings
    CreateTrayIcon(g_nid.hWnd);
    DebugPrint(L"[INFO] Tray icon refreshed\n");
}

void ReloadTargetPage(void) {
    if (!g_webView) return;

    if (g_initialUrl[0]) {
        g_webView->lpVtbl->Navigate(g_webView, g_initialUrl);
        DebugPrint(L"[INFO] Reloaded target URL: %s\n", g_initialUrl);
    } else {
        g_webView->lpVtbl->Reload(g_webView);
        DebugPrint(L"[INFO] Reloaded current page\n");
    }
}

void ClearWebViewCacheAndReload(void) {
    if (!g_webView) return;

    ICoreWebView2_13* webview13 = NULL;
    HRESULT hr = g_webView->lpVtbl->QueryInterface(
        g_webView, &IID_ICoreWebView2_13, (void**)&webview13);
    if (FAILED(hr) || !webview13) {
        DebugPrint(L"[WARNING] WebView2 profile interface not available. HRESULT: 0x%08X\n", hr);
        ReloadTargetPage();
        return;
    }

    ICoreWebView2Profile* profile = NULL;
    hr = webview13->lpVtbl->get_Profile(webview13, &profile);
    webview13->lpVtbl->Release(webview13);
    if (FAILED(hr) || !profile) {
        DebugPrint(L"[WARNING] Failed to get WebView2 profile. HRESULT: 0x%08X\n", hr);
        ReloadTargetPage();
        return;
    }

    ICoreWebView2Profile2* profile2 = NULL;
    hr = profile->lpVtbl->QueryInterface(profile, &IID_ICoreWebView2Profile2, (void**)&profile2);
    profile->lpVtbl->Release(profile);
    if (FAILED(hr) || !profile2) {
        DebugPrint(L"[WARNING] WebView2 profile2 interface not available. HRESULT: 0x%08X\n", hr);
        ReloadTargetPage();
        return;
    }

    ClearBrowsingDataCompletedHandler* handler =
        (ClearBrowsingDataCompletedHandler*)calloc(1, sizeof(ClearBrowsingDataCompletedHandler));
    if (!handler) {
        profile2->lpVtbl->Release(profile2);
        ReloadTargetPage();
        return;
    }

    static ICoreWebView2ClearBrowsingDataCompletedHandlerVtbl clearBrowsingDataVtbl = {
        ClearBrowsingDataCompletedHandler_QueryInterface,
        ClearBrowsingDataCompletedHandler_AddRef,
        ClearBrowsingDataCompletedHandler_Release,
        ClearBrowsingDataCompletedHandler_Invoke
    };

    handler->lpVtbl = &clearBrowsingDataVtbl;
    handler->refCount = 1;

    COREWEBVIEW2_BROWSING_DATA_KINDS dataKinds =
        COREWEBVIEW2_BROWSING_DATA_KINDS_DISK_CACHE |
        COREWEBVIEW2_BROWSING_DATA_KINDS_CACHE_STORAGE |
        COREWEBVIEW2_BROWSING_DATA_KINDS_SERVICE_WORKERS;

    hr = profile2->lpVtbl->ClearBrowsingData(
        profile2, dataKinds, (ICoreWebView2ClearBrowsingDataCompletedHandler*)handler);
    if (FAILED(hr)) {
        DebugPrint(L"[WARNING] Failed to clear WebView2 cache. HRESULT: 0x%08X\n", hr);
        ReloadTargetPage();
    }

    handler->lpVtbl->Release((ICoreWebView2ClearBrowsingDataCompletedHandler*)handler);
    profile2->lpVtbl->Release(profile2);
}

// Restart the executable: launch a fresh instance and shut this one down.
// The new instance retries the single-instance mutex (see WinMain) until this
// process has released it during exit.
static void RestartApplication(void) {
    wchar_t exePath[MAX_PATH];
    if (GetModuleFileNameW(NULL, exePath, MAX_PATH) == 0) {
        return;
    }

    STARTUPINFOW si = { sizeof(si) };
    PROCESS_INFORMATION pi = {0};
    if (!CreateProcessW(exePath, NULL, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi)) {
        MessageBoxW(NULL, L"Failed to restart the application.", APP_NAME,
                    MB_OK | MB_ICONERROR);
        return;
    }
    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);

    // Same shutdown path as tray Exit.
    if (g_nid.hIcon) {
        DestroyIcon(g_nid.hIcon);
        g_nid.hIcon = NULL;
    }
    Shell_NotifyIconW(NIM_DELETE, &g_nid);
    PostQuitMessage(0);
}

void ShowContextMenu(HWND hwnd) {
    POINT pt;
    GetCursorPos(&pt);

    // Build version string for menu
    wchar_t versionLabel[160];
    swprintf_s(versionLabel, 160, L"WebView2: %s", g_webView2Version);

    HMENU hMenu = CreatePopupMenu();
    AppendMenuW(hMenu, MF_STRING | MF_GRAYED, 0, versionLabel);
    AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenuW(hMenu, MF_STRING, ID_TRAY_MENU_REFRESH, L"Refresh");
    AppendMenuW(hMenu, MF_STRING, ID_TRAY_MENU_CLEAR_CACHE, L"Refresh + Clear Cache");
    AppendMenuW(hMenu, MF_STRING, ID_TRAY_MENU_RESTART, L"Restart");
    AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenuW(hMenu, MF_STRING, ID_TRAY_MENU_OPEN, L"Open");
    AppendMenuW(hMenu, MF_STRING, ID_TRAY_MENU_CONFIGURE, L"Configure");
    AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenuW(hMenu, MF_STRING, ID_TRAY_MENU_EXIT, L"Exit");

    SetForegroundWindow(hwnd);
    TrackPopupMenu(hMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, pt.x, pt.y, 0, hwnd, NULL);
    DestroyMenu(hMenu);
}

// Window procedure
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
        case WM_CREATE:
            CaptureDisplaySettings();
            CreateMainWebViewEnvironment(hwnd);
            return 0;
            
        case WM_SIZE:
            if (wParam == SIZE_MINIMIZED) {
                DeactivateMainWebView();
            } else if (IsWindowVisible(hwnd)) {
                ActivateMainWebView();
            }
            return 0;
            
        case WM_DISPLAYCHANGE:
            DebugPrint(L"[INFO] Display change event received...\n");
            if (g_timerId) KillTimer(hwnd, g_timerId);
            g_timerId = SetTimer(hwnd, 1, RESOLUTION_CHANGE_DEBOUNCE_MS, NULL);
            return 0;

        case WM_DPICHANGED: {
            // Per-monitor DPI aware: adopt the size Windows suggests for the
            // new monitor's scale; WM_SIZE re-syncs the WebView bounds.
            const RECT* suggested = (const RECT*)lParam;
            SetWindowPos(hwnd, NULL, suggested->left, suggested->top,
                         suggested->right - suggested->left,
                         suggested->bottom - suggested->top,
                         SWP_NOZORDER | SWP_NOACTIVATE);
            return 0;
        }
            
        case WM_TIMER:
            if (wParam == 1) {
                KillTimer(hwnd, g_timerId);
                g_timerId = 0;
                if (HasDisplaySettingsChanged()) {
                    RefreshTrayIcon();
                    CaptureDisplaySettings();
                }
            } else if (wParam == ID_TIMER_INITIAL_HIDE_JS) {
                KillTimer(hwnd, ID_TIMER_INITIAL_HIDE_JS);
                // Start visibility polling after initial JS sync
                StartVisibilityTimer(hwnd);
                UpdateJsVisibilityState(hwnd);
            } else if (wParam == ID_TIMER_VISIBILITY_CHECK) {
                // Periodic check for window occlusion
                UpdateJsVisibilityState(hwnd);
            } else if (wParam == ID_TIMER_WEBVIEW_PREWARM) {
                KillTimer(hwnd, ID_TIMER_WEBVIEW_PREWARM);
                InterlockedExchange(&g_webViewPrewarmActive, FALSE);
                if (IsWindowActuallyVisible(hwnd)) {
                    ActivateMainWebView();
                } else {
                    DeactivateMainWebView();
                }
            } else if (wParam == ID_TIMER_WEBVIEW_PRELOAD) {
                KillTimer(hwnd, ID_TIMER_WEBVIEW_PRELOAD);
                // Initial preload has settled: suspend now if still hidden and
                // not being kept warm by a tray-hover prewarm.
                if (!IsWindowActuallyVisible(hwnd) &&
                    InterlockedCompareExchange(&g_webViewPrewarmActive, TRUE, TRUE) != TRUE) {
                    DeactivateMainWebView();
                }
            } else if (wParam == ID_TIMER_WEBVIEW_RECREATE) {
                // BrowserProcessExited never arrived; rebuild anyway.
                DebugPrint(L"[WARNING] Browser exit event not received; rebuilding WebView after timeout\n");
                FinishMainWebViewRecreate(hwnd);
            } else if (wParam == ID_TIMER_POWER_RESUME) {
                KillTimer(hwnd, ID_TIMER_POWER_RESUME);
                g_powerKickCount++;
                KickWebViewAfterPowerResume(hwnd);
            } else if (wParam == ID_TIMER_WEBVIEW_LIVENESS) {
                KillTimer(hwnd, ID_TIMER_WEBVIEW_LIVENESS);
                CheckMainWebViewLiveness(hwnd);
            }
            return 0;

        case WM_POWERBROADCAST:
            if (wParam == PBT_APMSUSPEND) {
                // Going down: from here on, nothing we believe about the
                // runtime's suspend/resume state can be trusted. The recovery
                // sequence starts when a resume broadcast arrives.
                InterlockedExchange(&g_powerResumePending, TRUE);
                return TRUE;
            }
            if (wParam == PBT_APMQUERYSUSPENDFAILED) {
                // The suspend was vetoed; there is nothing to recover from.
                InterlockedExchange(&g_powerResumePending, FALSE);
                return TRUE;
            }
            if (wParam == PBT_APMRESUMEAUTOMATIC || wParam == PBT_APMRESUMESUSPEND ||
                wParam == PBT_APMRESUMECRITICAL) {
                // (Re)start the recovery sequence. Several resume broadcasts
                // can arrive for a single resume; restarting is idempotent.
                // Give the graphics stack a moment to come back up before the
                // first kick (see KickWebViewAfterPowerResume).
                KillTimer(hwnd, ID_TIMER_POWER_RESUME);
                KillTimer(hwnd, ID_TIMER_WEBVIEW_LIVENESS);
                InterlockedExchange(&g_webViewPingOutstanding, FALSE);
                InterlockedExchange(&g_powerResumePending, TRUE);
                g_powerKickCount = 0;
                SetTimer(hwnd, ID_TIMER_POWER_RESUME, POWER_RESUME_KICK_DELAY_MS, NULL);
            }
            return TRUE;
            
        case WM_SYSCOMMAND:
            if ((wParam & 0xFFF0) == SC_MINIMIZE) {
                HideMainWindow();
                return 0;
            }
            break;
            
        case WM_CLOSE:
            HideMainWindow();
            return 0;
            
        case WM_DESTROY:
            KillTimer(hwnd, ID_TIMER_WEBVIEW_PREWARM);
            KillTimer(hwnd, ID_TIMER_WEBVIEW_PRELOAD);
            KillTimer(hwnd, ID_TIMER_POWER_RESUME);
            KillTimer(hwnd, ID_TIMER_WEBVIEW_LIVENESS);
            PostQuitMessage(0);
            return 0;
            
        case WM_APP_SPELLCHECK_CHANGED:
            if (MessageBoxW(NULL,
                    L"Spell-check languages were changed.\n\n"
                    L"The embedded web view has to restart for them to take effect, "
                    L"which reloads the page - unsaved work in the page will be lost.\n\n"
                    L"Restart the web view now?\n"
                    L"(Choosing No applies the change the next time the app starts.)",
                    APP_NAME, MB_YESNO | MB_ICONQUESTION | MB_SETFOREGROUND) == IDYES) {
                BeginMainWebViewRecreate();
            }
            return 0;

        case WM_APP_WEBVIEW_RECREATE:
            if (InterlockedCompareExchange(&g_webViewRecreatePending, TRUE, TRUE) == TRUE) {
                // Deliberate rebuild (spell-check change): the browser was
                // asked to exit and now has.
                FinishMainWebViewRecreate(hwnd);
            } else {
                // Unsolicited: the browser process died or stopped resuming.
                HandleUnexpectedBrowserExit(hwnd);
            }
            return 0;

        case WM_TRAYICON:
            switch (lParam) {
                case WM_MOUSEMOVE: PrewarmMainWebView(); break;
                case WM_LBUTTONDBLCLK: ShowMainWindow(); break;
                case WM_RBUTTONUP: ShowContextMenu(hwnd); break;
            }
            return 0;
            
        case WM_COMMAND:
            switch (wParam) {
                case ID_TRAY_MENU_REFRESH:
                    RebuildMainWebViewIfDead();
                    // If window is already restored and visible, don't reposition it
                    if (IsWindowVisible(g_hwnd) && !IsIconic(g_hwnd) && IsWindowActuallyVisible(g_hwnd)) {
                        SetForegroundWindow(g_hwnd);
                    } else {
                        ShowMainWindow();
                    }
                    ReloadTargetPage();
                    return 0;
                case ID_TRAY_MENU_CLEAR_CACHE:
                    RebuildMainWebViewIfDead();
                    // If window is already restored and visible, don't reposition it
                    if (IsWindowVisible(g_hwnd) && !IsIconic(g_hwnd) && IsWindowActuallyVisible(g_hwnd)) {
                        SetForegroundWindow(g_hwnd);
                    } else {
                        ShowMainWindow();
                    }
                    ClearWebViewCacheAndReload();
                    return 0;
                case ID_TRAY_MENU_OPEN:
                    ShowMainWindow();
                    return 0;
                case ID_TRAY_MENU_RESTART:
                    RestartApplication();
                    return 0;
                case ID_TRAY_MENU_CONFIGURE:
                    g_cfgSaved = FALSE;
                    ShowConfigWebViewDialog();
                    return 0;
                case ID_TRAY_MENU_EXIT:
                    // Clean up tray icon resources before exit
                    if (g_nid.hIcon) {
                        DestroyIcon(g_nid.hIcon);
                        g_nid.hIcon = NULL;
                    }
                    Shell_NotifyIconW(NIM_DELETE, &g_nid);
                    PostQuitMessage(0);
                    return 0;
            }
            break;
            
        default:
            if (uMsg == g_WM_TASKBARCREATED) {
                CreateTrayIcon(hwnd);
                return 0;
            }
    }
    return DefWindowProcW(hwnd, uMsg, wParam, lParam);
}

// Debug output (no-op in release builds)
void DebugPrint(const wchar_t* format, ...) {
#ifdef _DEBUG
    va_list args;
    va_start(args, format);
    wchar_t buffer[4096];
    vswprintf_s(buffer, sizeof(buffer)/sizeof(wchar_t), format, args);
    OutputDebugStringW(buffer);
    va_end(args);
#endif
}

// Entry point
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    g_hInstance = hInstance;
    
    // Single instance check. A restarted instance (tray Restart) can arrive
    // while the previous process is still shutting down, so retry briefly
    // before declaring another instance is running.
    g_hMutex = CreateMutexW(NULL, TRUE, MUTEX_NAME);
    for (int attempt = 0;
         g_hMutex && GetLastError() == ERROR_ALREADY_EXISTS && attempt < 10;
         attempt++) {
        CloseHandle(g_hMutex);
        Sleep(250);
        g_hMutex = CreateMutexW(NULL, TRUE, MUTEX_NAME);
    }
    if (g_hMutex && GetLastError() == ERROR_ALREADY_EXISTS) {
        CloseHandle(g_hMutex);
        g_hMutex = NULL;
        MessageBoxW(NULL, L"SystrayLauncher is already running.\n\nCheck your system tray for the application icon.",
                    L"Already Running", MB_OK | MB_ICONINFORMATION);
        return 0;
    }
    
    // Initialize COM
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) {
        MessageBoxW(NULL, L"COM initialization failed", L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }

    if (!load_webview2_loader()) {
        MessageBoxW(NULL,
            L"Failed to load WebView2.\n\n"
            L"Please ensure the Microsoft Edge WebView2 Runtime is installed.\n"
            L"Download from: https://developer.microsoft.com/en-us/microsoft-edge/webview2/",
            L"Error", MB_ICONERROR | MB_OK);
        CoUninitialize();
        if (g_hMutex) { ReleaseMutex(g_hMutex); CloseHandle(g_hMutex); }
        return 1;
    }

    // Get exe directory
    wchar_t exePath[MAX_PATH];
    GetModuleFileNameW(NULL, exePath, MAX_PATH);
    PathRemoveFileSpecW(exePath);
    wcscpy_s(g_iniPath, MAX_PATH, exePath);
    PathAppendW(g_iniPath, CONFIG_FILENAME);

    // Check if this is the first launch
    BOOL isFirstLaunch = IsFirstLaunch();

    // Try to load config from registry first
    if (!LoadConfigFromRegistry(&g_config)) {
        // Fallback to INI file (for migration or first launch)
        LoadConfiguration(g_iniPath, &g_config);
    }
    {
        // Sanitize hand-edited registry/INI values so the JSON patch and
        // Chromium both see well-formed language tags.
        wchar_t rawLangs[512];
        wcscpy_s(rawLangs, 512, g_config.spellcheckLanguages);
        NormalizeSpellcheckLanguages(rawLangs, g_config.spellcheckLanguages, 512);
    }
    wcscpy_s(g_initialUrl, 2048, g_config.url);
    InterlockedExchange(&g_sleepWhenInactive, g_config.sleepWhenInactive ? TRUE : FALSE);
    InterlockedExchange(&g_openNewWindowsExternally, g_config.openNewWindowsExternally ? TRUE : FALSE);

    // On first launch, show configuration dialog
    if (isFirstLaunch) {
        g_cfgSaved = FALSE;
        ShowConfigWebViewDialog();
        // Nested message loop — runs until config dialog is closed
        MSG cfgMsg;
        while (g_cfgHwnd && GetMessage(&cfgMsg, NULL, 0, 0)) {
            TranslateMessage(&cfgMsg);
            DispatchMessage(&cfgMsg);
        }
        if (!g_cfgSaved) {
            // User cancelled on first launch - exit
            CoUninitialize();
            if (g_hMutex) {
                ReleaseMutex(g_hMutex);
                CloseHandle(g_hMutex);
            }
            return 0;
        }
        wcscpy_s(g_initialUrl, 2048, g_config.url);
    }

    // Apply spell-check languages to the WebView2 profile before the browser
    // process launches (the Preferences file can only be edited while the
    // profile is not in use). The config dialog above uses a separate user
    // data folder, so it does not conflict with this.
    PatchSpellcheckPreferences();

    // Register invisible owner window class (prevents taskbar appearance)
    WNDCLASSEXW ownerWc = {0};
    ownerWc.cbSize = sizeof(ownerWc);
    ownerWc.lpfnWndProc = DefWindowProcW;
    ownerWc.hInstance = hInstance;
    ownerWc.lpszClassName = L"SystrayLauncherOwner";
    RegisterClassExW(&ownerWc);

    // Create invisible owner window
    g_hwndOwner = CreateWindowExW(0, L"SystrayLauncherOwner", L"",
                                  WS_POPUP, 0, 0, 0, 0,
                                  NULL, NULL, hInstance, NULL);
    
    // Register window class
    WNDCLASSEXW wc = {0};
    wc.cbSize = sizeof(wc);
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"SystrayLauncherClass";
    
    // Load embedded icon for the application window
    wc.hIcon = LoadIconW(g_hInstance, MAKEINTRESOURCEW(IDI_TRAYICON));
    wc.hIconSm = (HICON)LoadImageW(g_hInstance, MAKEINTRESOURCEW(IDI_TRAYICON),
                                   IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR);

    wc.hCursor = LoadCursorW(NULL, (LPCWSTR)IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    
    if (!RegisterClassExW(&wc)) {
        MessageBoxW(NULL, L"Failed to register window class", L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }
    
    // Create main window with invisible owner (prevents taskbar appearance)
    g_hwnd = CreateWindowExW(0, L"SystrayLauncherClass", g_config.windowTitle,
                            WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT,
                            1024, 768, g_hwndOwner, NULL, hInstance, NULL);
    if (!g_hwnd) {
        MessageBoxW(NULL, L"Failed to create window", L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }
    
    // Register for taskbar restart notifications
    g_WM_TASKBARCREATED = RegisterWindowMessageW(L"TaskbarCreated");
    
    // Create tray icon (loads embedded icon)
    CreateTrayIcon(g_hwnd);
    
    // Message loop
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    
    // Cleanup
    if (g_webView) g_webView->lpVtbl->Release(g_webView);
    if (g_webViewController) g_webViewController->lpVtbl->Release(g_webViewController);
    if (g_webViewEnv) {
        UnregisterBrowserExitedFromCurrentEnv();
        g_webViewEnv->lpVtbl->Release(g_webViewEnv);
    }
    
    // Clean up tray icon and its resources
    if (g_nid.hIcon) {
        DestroyIcon(g_nid.hIcon);
        g_nid.hIcon = NULL;
    }
    Shell_NotifyIconW(NIM_DELETE, &g_nid);
    
    CoUninitialize();
    if (g_hwndOwner) {
        DestroyWindow(g_hwndOwner);
    }
    if (g_hMutex) {
        ReleaseMutex(g_hMutex);
        CloseHandle(g_hMutex);
    }
    
    return (int)msg.wParam;
}
