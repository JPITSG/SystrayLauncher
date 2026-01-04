#include <windows.h>
#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <math.h>
#include <assert.h>

// WebView2 headers required from SDK
#include "WebView2.h"

#define WINDOW_SIZE_PERCENTAGE 0.9
#define RESOLUTION_CHANGE_DEBOUNCE_MS 1000
#define CONFIG_FILENAME L"config.ini"
#define APP_NAME L"NextcloudLauncher"
#define MUTEX_NAME L"NextcloudLauncher_SingleInstance_Mutex_9F8A7B6C"
#define TRAY_ICON_ID 100
#define WM_TRAYICON (WM_APP + 1)
#define IDI_TRAYICON 101  // Resource ID for embedded icon
#define ID_TRAY_MENU_REFRESH 1
#define ID_TRAY_MENU_CLEAR_CACHE 2
#define ID_TRAY_MENU_OPEN 3
#define ID_TRAY_MENU_EXIT 4
#define ID_TIMER_INITIAL_HIDE_JS 2
#define INITIAL_HIDE_JS_DELAY_MS 2000

typedef struct {
    wchar_t url[2048];
    wchar_t windowTitle[256];
    wchar_t onHideJs[4096];
    wchar_t onShowJs[4096];
} Configuration;

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
    wcscpy_s(config->windowTitle, 256, L"Nextcloud Launcher");
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
    const wchar_t* content = L"# NextcloudLauncher Configuration File\n"
                             L"# Lines starting with # or ; are comments\n\n"
                             L"url=https://www.google.com/\n\n"
                             L"windowtitle=Nextcloud Launcher\n";
    FILE* file = NULL;
    _wfopen_s(&file, iniPath, L"w, ccs=UTF-8");
    if (file) {
        fputws(content, file);
        fclose(file);
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

        // Schedule initial onHideJs execution (app starts hidden)
        if (g_config.onHideJs[0] != L'\0') {
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

    // Execute optional onShowJs from config
    if (g_config.onShowJs[0] != L'\0') {
        ExecuteJavaScript(g_config.onShowJs);
    }

    DebugPrint(L"[INFO] Main window shown at %dx%d, size %dx%d\n", x, y, windowWidth, windowHeight);
}

void HideMainWindow(void) {
	if (!g_hwnd) return;

    // Execute optional onHideJs from config before hiding
    if (g_config.onHideJs[0] != L'\0') {
        ExecuteJavaScript(g_config.onHideJs);
    }

    ShowWindow(g_hwnd, SW_HIDE);

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
    
    HMENU hMenu = CreatePopupMenu();
    AppendMenuW(hMenu, MF_STRING, ID_TRAY_MENU_REFRESH, L"Refresh");
    AppendMenuW(hMenu, MF_STRING, ID_TRAY_MENU_CLEAR_CACHE, L"Refresh + Clear Cache");
    AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenuW(hMenu, MF_STRING, ID_TRAY_MENU_OPEN, L"Open");
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
            
            CreateCoreWebView2EnvironmentWithOptions(NULL, userDataPath, NULL,
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
                // Execute onHideJs on startup (app starts hidden)
                if (g_config.onHideJs[0] != L'\0') {
                    ExecuteJavaScript(g_config.onHideJs);
                    DebugPrint(L"[INFO] Executed initial onHideJs\n");
                }
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
                    ReloadTargetPage();
                    return 0;
                case ID_TRAY_MENU_CLEAR_CACHE:
                    ClearWebViewCacheAndReload();
                    return 0;
                case ID_TRAY_MENU_OPEN:
                    ShowMainWindow();
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
        MessageBoxW(NULL, L"NextcloudLauncher is already running.\n\nCheck your system tray for the application icon.",
                    L"Already Running", MB_OK | MB_ICONINFORMATION);
        return 0;
    }
    
    // Initialize COM
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) {
        MessageBoxW(NULL, L"COM initialization failed", L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }
    
    // Get exe directory
    wchar_t exePath[MAX_PATH];
    GetModuleFileNameW(NULL, exePath, MAX_PATH);
    PathRemoveFileSpecW(exePath);
    wcscpy_s(g_iniPath, MAX_PATH, exePath);
    PathAppendW(g_iniPath, CONFIG_FILENAME);
    
    LoadConfiguration(g_iniPath, &g_config);
    wcscpy_s(g_initialUrl, 2048, g_config.url);
    
    // Register invisible owner window class (prevents taskbar appearance)
    WNDCLASSEXW ownerWc = {0};
    ownerWc.cbSize = sizeof(ownerWc);
    ownerWc.lpfnWndProc = DefWindowProcW;
    ownerWc.hInstance = hInstance;
    ownerWc.lpszClassName = L"NextcloudLauncherOwner";
    RegisterClassExW(&ownerWc);
    
    // Create invisible owner window
    g_hwndOwner = CreateWindowExW(0, L"NextcloudLauncherOwner", L"",
                                  WS_POPUP, 0, 0, 0, 0,
                                  NULL, NULL, hInstance, NULL);
    
    // Register window class
    WNDCLASSEXW wc = {0};
    wc.cbSize = sizeof(wc);
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"NextcloudLauncherClass";
    
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
    g_hwnd = CreateWindowExW(0, L"NextcloudLauncherClass", g_config.windowTitle,
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
