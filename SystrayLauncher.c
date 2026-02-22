#include <windows.h>
#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include <wchar.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <math.h>
#include <assert.h>

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

// Registry settings
#define REG_COMPANY L"JPIT"
#define REG_APPNAME L"SystrayLauncher"
#define REG_KEY_PATH L"SOFTWARE\\JPIT\\SystrayLauncher"
#define REG_VALUE_URL L"URL"
#define REG_VALUE_TITLE L"WindowTitle"
#define REG_VALUE_ONHIDEJS L"OnHideJS"
#define REG_VALUE_ONSHOWJS L"OnShowJS"
#define REG_VALUE_CONFIGURED L"Configured"

#define ID_TIMER_INITIAL_HIDE_JS 2
#define INITIAL_HIDE_JS_DELAY_MS 2000
#define ID_TIMER_VISIBILITY_CHECK 3
#define VISIBILITY_CHECK_INTERVAL_MS 250

typedef struct {
    wchar_t url[2048];
    wchar_t windowTitle[256];
    wchar_t onHideJs[4096];
    wchar_t onShowJs[4096];
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
void SetSpellCheckLanguages(ICoreWebView2* webview, const wchar_t* languages);
void ReloadTargetPage(void);
void ClearWebViewCacheAndReload(void);
void ExecuteJavaScript(const wchar_t* js);
static BOOL IsWebViewReady(void);
static BOOL IsWindowActuallyVisible(HWND hwnd);
static void UpdateJsVisibilityState(HWND hwnd);
static void StartVisibilityTimer(HWND hwnd);
static void StopVisibilityTimer(HWND hwnd);

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

// Configuration functions
void LoadConfiguration(const wchar_t* iniPath, Configuration* config) {
    wcscpy_s(config->url, 2048, L"https://www.google.com/");
    wcscpy_s(config->windowTitle, 256, L"Systray Launcher");
    config->onHideJs[0] = L'\0';
    config->onShowJs[0] = L'\0';

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

static void webview_push_init_config(void) {
    wchar_t eUrl[4096], eTitle[512], eHide[8192], eShow[8192];
    json_escape_wstring(g_config.url, eUrl, 4096);
    json_escape_wstring(g_config.windowTitle, eTitle, 512);
    json_escape_wstring(g_config.onHideJs, eHide, 8192);
    json_escape_wstring(g_config.onShowJs, eShow, 8192);

    wchar_t script[16384];
    swprintf(script, 16384,
        L"window.onInit({\"config\":{\"url\":\"%s\",\"windowTitle\":\"%s\",\"onHideJs\":\"%s\",\"onShowJs\":\"%s\"}})",
        eUrl, eTitle, eHide, eShow);
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

static HRESULT STDMETHODCALLTYPE CfgEnvCompleted_Invoke(
    ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler *This,
    HRESULT result, ICoreWebView2Environment *env) {
    (void)This;
    if (FAILED(result) || !env) return result;
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
    if (FAILED(result) || !controller) return result;

    g_cfgController = controller;
    controller->lpVtbl->AddRef(controller);

    RECT bounds;
    GetClientRect(g_cfgHwnd, &bounds);
    controller->lpVtbl->put_Bounds(controller, bounds);

    ICoreWebView2 *webview = NULL;
    controller->lpVtbl->get_CoreWebView2(controller, &webview);
    if (!webview) return E_FAIL;
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
        json_get_string(msg, "url", url, sizeof(url));
        json_get_string(msg, "windowTitle", title, sizeof(title));
        json_get_string(msg, "onHideJs", hideJs, sizeof(hideJs));
        json_get_string(msg, "onShowJs", showJs, sizeof(showJs));

        MultiByteToWideChar(CP_UTF8, 0, url, -1, g_config.url, 2048);
        MultiByteToWideChar(CP_UTF8, 0, title, -1, g_config.windowTitle, 256);
        MultiByteToWideChar(CP_UTF8, 0, hideJs, -1, g_config.onHideJs, 4096);
        MultiByteToWideChar(CP_UTF8, 0, showJs, -1, g_config.onShowJs, 4096);

        SaveConfigToRegistry(&g_config);
        MarkAsConfigured();
        ApplyConfiguration();
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
            RECT clientRect = {0}, windowRect = {0};
            GetClientRect(g_cfgHwnd, &clientRect);
            GetWindowRect(g_cfgHwnd, &windowRect);
            int chromeH = (windowRect.bottom - windowRect.top) - (clientRect.bottom - clientRect.top);
            int newWindowH = contentHeight + chromeH;
            int windowW = windowRect.right - windowRect.left;
            UINT flags = SWP_NOMOVE | SWP_NOZORDER;
            if (g_cfgWindowShown) {
                flags |= SWP_NOACTIVATE;
            } else {
                flags |= SWP_SHOWWINDOW;
            }
            SetWindowPos(g_cfgHwnd, NULL, 0, 0, windowW, newWindowH, flags);
            g_cfgWindowShown = TRUE;
        }
    }

    free(msg);
    return S_OK;
}

// Config dialog window procedure
static LRESULT CALLBACK CfgWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_SIZE:
            if (g_cfgController) {
                RECT bounds;
                GetClientRect(hwnd, &bounds);
                g_cfgController->lpVtbl->put_Bounds(g_cfgController, bounds);
            }
            return 0;

        case WM_CLOSE:
            g_cfgWindowShown = FALSE;
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

    int screenW = GetSystemMetrics(SM_CXSCREEN);
    int screenH = GetSystemMetrics(SM_CYSCREEN);
    int width = 480, height = 380;
    int posX = (screenW - width) / 2;
    int posY = (screenH - height) / 2;

    g_cfgHwnd = CreateWindowExW(0, L"SystrayLauncherCfgWnd", L"Configuration",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
        posX, posY, width, height,
        NULL, NULL, g_hInstance, NULL);

    if (!g_cfgHwnd) return;
    g_cfgWindowShown = FALSE;

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

// Function to enable spell checking via settings
void SetSpellCheckLanguages(ICoreWebView2* webview, const wchar_t* languages) {
    if (!webview || !languages) return;
    
    // Get settings
    ICoreWebView2Settings* settings = NULL;
    HRESULT hr = webview->lpVtbl->get_Settings(webview, &settings);
    
    if (SUCCEEDED(hr) && settings) {
        // Enable context menus (required for spell check suggestions to appear)
        hr = settings->lpVtbl->put_AreDefaultContextMenusEnabled(settings, TRUE);
        if (SUCCEEDED(hr)) {
            DebugPrint(L"[INFO] Context menus enabled for spell checking\n");
        } else {
            DebugPrint(L"[WARNING] Failed to enable context menus. HRESULT: 0x%08X\n", hr);
        }
        
        DebugPrint(L"[INFO] Spell check will use languages configured in Edge (edge://settings/languages)\n");
        DebugPrint(L"[INFO] Configured languages: %s\n", languages);
        
        settings->lpVtbl->Release(settings);
    } else {
        DebugPrint(L"[WARNING] Failed to get WebView2 settings. HRESULT: 0x%08X\n", hr);
    }
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
        MessageBoxW(NULL, L"WebView2 environment creation failed", L"Error", MB_OK | MB_ICONERROR);
        return result;
    }

    // Get WebView2 browser version string
    LPWSTR versionString = NULL;
    if (SUCCEEDED(environment->lpVtbl->get_BrowserVersionString(environment, &versionString)) && versionString) {
        wcscpy_s(g_webView2Version, 128, versionString);
        CoTaskMemFree(versionString);
    }

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
        
        RECT bounds;
        GetClientRect(hwnd, &bounds);
        controller->lpVtbl->put_Bounds(controller, bounds);
        controller->lpVtbl->put_IsVisible(controller, TRUE);
        
        webview2->lpVtbl->Navigate(webview2, g_initialUrl);
        
        // Set spell check languages to en-US and pl-PL
        SetSpellCheckLanguages(webview2, L"en-US,pl-PL");

        InterlockedExchange(&g_isInitialized, TRUE);

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
    if (g_config.onHideJs[0] != L'\0' || g_config.onShowJs[0] != L'\0') {
        SetTimer(hwnd, ID_TIMER_VISIBILITY_CHECK, VISIBILITY_CHECK_INTERVAL_MS, NULL);
        DebugPrint(L"[INFO] Started visibility check timer\n");
    }
}

static void StopVisibilityTimer(HWND hwnd) {
    KillTimer(hwnd, ID_TIMER_VISIBILITY_CHECK);
    DebugPrint(L"[INFO] Stopped visibility check timer\n");
}

static void UpdateJsVisibilityState(HWND hwnd) {
    if (!IsWebViewReady()) return;

    JsVisibility newState = IsWindowActuallyVisible(hwnd) ? JS_VISIBILITY_SHOWN : JS_VISIBILITY_HIDDEN;
    if (newState == g_jsVisibility) return;

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

// Window management
void ShowMainWindow(void) {
    if (!g_hwnd) return;

    RECT workArea;
    SystemParametersInfoW(SPI_GETWORKAREA, 0, &workArea, 0);

    int workWidth = workArea.right - workArea.left;
    int workHeight = workArea.bottom - workArea.top;

    int windowWidth = (int)(workWidth * WINDOW_SIZE_PERCENTAGE);
    int windowHeight = (int)(workHeight * WINDOW_SIZE_PERCENTAGE);
    int x = workArea.left + (workWidth - windowWidth) / 2;
    int y = workArea.top + (workHeight - windowHeight) / 2;

    SetWindowPos(g_hwnd, HWND_TOP, x, y, windowWidth, windowHeight,
                 SWP_SHOWWINDOW | SWP_FRAMECHANGED);
    ShowWindow(g_hwnd, SW_RESTORE);
    SetForegroundWindow(g_hwnd);

    if (g_webViewController) {
        RECT bounds;
        GetClientRect(g_hwnd, &bounds);
        g_webViewController->lpVtbl->put_Bounds(g_webViewController, bounds);
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

	if (g_webView && g_initialUrl[0]) {
		LPWSTR currentUrl = NULL;
		HRESULT hr = g_webView->lpVtbl->get_Source(g_webView, &currentUrl);

		if (SUCCEEDED(hr) && currentUrl) {
			if (wcscmp(currentUrl, g_initialUrl) != 0) {
				g_webView->lpVtbl->Navigate(g_webView, g_initialUrl);
				DebugPrint(L"[INFO] Reset URL to initial on hide: %s (was: %s)\n", g_initialUrl, currentUrl);
			} else {
				DebugPrint(L"[INFO] URL unchanged, skipping navigation on hide\n");
			}
			CoTaskMemFree(currentUrl);
		} else {
			// Fallback: navigate anyway if we couldn't get current URL
			g_webView->lpVtbl->Navigate(g_webView, g_initialUrl);
			DebugPrint(L"[INFO] Reset URL to initial on hide (couldn't check current): %s\n", g_initialUrl);
		}
	}

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
            
            wchar_t userDataPath[MAX_PATH];
            SHGetFolderPathW(NULL, CSIDL_LOCAL_APPDATA, NULL, 0, userDataPath);
            PathAppendW(userDataPath, APP_NAME L"\\WebView2Data");
            CreateDirectoryW(userDataPath, NULL);
            
            EnvCompletedHandler* envHandler = (EnvCompletedHandler*)calloc(1, sizeof(EnvCompletedHandler));
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
            
            fnCreateEnvironment(NULL, userDataPath, NULL,
                (ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*)envHandler);
            envHandler->lpVtbl->Release((ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*)envHandler);
            return 0;
            
        case WM_SIZE:
            if (g_webViewController && wParam != SIZE_MINIMIZED && IsWindowVisible(hwnd)) {
                RECT bounds;
                GetClientRect(hwnd, &bounds);
                g_webViewController->lpVtbl->put_Bounds(g_webViewController, bounds);
            }
            return 0;
            
        case WM_DISPLAYCHANGE:
            DebugPrint(L"[INFO] Display change event received...\n");
            if (g_timerId) KillTimer(hwnd, g_timerId);
            g_timerId = SetTimer(hwnd, 1, RESOLUTION_CHANGE_DEBOUNCE_MS, NULL);
            return 0;
            
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
            }
            return 0;
            
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
            PostQuitMessage(0);
            return 0;
            
        case WM_TRAYICON:
            switch (lParam) {
                case WM_LBUTTONDBLCLK: ShowMainWindow(); break;
                case WM_RBUTTONUP: ShowContextMenu(hwnd); break;
            }
            return 0;
            
        case WM_COMMAND:
            switch (wParam) {
                case ID_TRAY_MENU_REFRESH:
                    // If window is already restored and visible, don't reposition it
                    if (IsWindowVisible(g_hwnd) && !IsIconic(g_hwnd) && IsWindowActuallyVisible(g_hwnd)) {
                        SetForegroundWindow(g_hwnd);
                    } else {
                        ShowMainWindow();
                    }
                    ReloadTargetPage();
                    return 0;
                case ID_TRAY_MENU_CLEAR_CACHE:
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
    
    // Single instance check
    g_hMutex = CreateMutexW(NULL, TRUE, MUTEX_NAME);
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
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
    wcscpy_s(g_initialUrl, 2048, g_config.url);

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
