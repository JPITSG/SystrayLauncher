// WebView2 needs Windows 10; declare that so Vista+ APIs (GetTickCount64,
// DWM attributes) are visible in the headers.
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0A00
#endif
#ifndef WINVER
#define WINVER 0x0A00
#endif

#include <winsock2.h>
#include <ws2tcpip.h>
#include <windows.h>
#include <userenv.h>
#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include <wchar.h>
#include <wctype.h>
#include <shellapi.h>
#include <sddl.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <dwmapi.h>
#include <bcrypt.h>
#include <winhttp.h>
#include <winver.h>
#include <math.h>

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
#include "version.h"

// The C++ helper in the bundled 1.0.3650.58 SDK initializes the required
// target-version property to CORE_WEBVIEW_TARGET_PRODUCT_VERSION. Keep its
// value as the fallback for this project's plain-C options object.
#define WEBVIEW2_TARGET_COMPATIBLE_BROWSER_VERSION L"143.0.3650.58"

#define WINDOW_SIZE_PERCENTAGE 0.9
#define RESOLUTION_CHANGE_DEBOUNCE_MS 1000
#define CONFIG_FILENAME L"config.ini"
#define APP_NAME L"SystrayLauncher"
#define MUTEX_NAME L"SystrayLauncher_SingleInstance_Mutex_9F8A7B6C"
#define MAILTO_ACTIVATE_MESSAGE_NAME \
    L"SystrayLauncher_MailtoActivation_43DDF20A_891B_4C96_A7E2_8E4F46C0A7A1"
#define TRAY_ICON_ID 100
#define WM_TRAYICON (WM_APP + 1)
#define WM_APP_UPDATE_RESULT (WM_APP + 3)
#define WM_APP_UPDATE_PROGRESS (WM_APP + 4)
#define ID_TRAY_MENU_REFRESH 1
#define ID_TRAY_MENU_CLEAR_CACHE 2
#define ID_TRAY_MENU_OPEN 3
#define ID_TRAY_MENU_CONFIGURE 5
#define ID_TRAY_MENU_EXIT 4
#define ID_TRAY_MENU_RESTART 6

#define UPDATE_URL L"https://github.com/JPITSG/SystrayLauncher/raw/refs/heads/main/release/SystrayLauncher.exe"
#define UPDATE_MAX_BYTES (100ULL * 1024ULL * 1024ULL)
#define UPDATE_PROGRESS_INTERVAL_MS 250
#define UPDATE_HELPER_READY_MS 10000
#define UPDATE_HELPER_WAIT_MS 120000

// Registry settings
#define REG_COMPANY L"JPIT"
#define REG_APPNAME L"SystrayLauncher"
#define REG_KEY_PATH L"SOFTWARE\\JPIT\\SystrayLauncher"
#define REG_VALUE_URL L"URL"
#define REG_VALUE_TITLE L"WindowTitle"
#define REG_VALUE_START_MAXIMIZED L"StartMaximized"
#define REG_VALUE_RETURN_TO_TARGET_ON_DOUBLE_CLICK L"ReturnToTargetOnDoubleClick"
#define REG_VALUE_SHOW_IN_TASKBAR L"ShowInTaskbar"
#define REG_VALUE_HANDLE_MAILTO_LINKS L"HandleMailtoLinks"
#define REG_VALUE_MAILTO_TARGET_URL L"MailtoTargetURL"
#define REG_VALUE_ONHIDEJS L"OnHideJS"
#define REG_VALUE_ONSHOWJS L"OnShowJS"
#define REG_VALUE_SLEEP L"SleepWhenInactive"
#define REG_VALUE_NEWWINDOW L"OpenNewWindowsExternally"
#define REG_VALUE_INSECURE_CONTENT L"AllowRunningInsecureContent"
#define REG_VALUE_INSECURE_CONTENT_ORIGINS L"InsecureContentOrigins"
#define REG_VALUE_STATIC_HOSTS L"UseStaticHostMappings"
#define REG_VALUE_STATIC_HOST_MAPPINGS L"StaticHostMappings"
#define REG_VALUE_STATIC_HOST_DNS_FALLBACK L"StaticHostDnsFallback"
#define REG_VALUE_LOCKDOWN L"LockdownHeader"
#define REG_VALUE_LOCKDOWN_SECRET L"LockdownSecret"
#define REG_VALUE_AUTO_UPDATE L"AutoCheckForUpdates"
#define REG_VALUE_IGNORED_UPDATE_VERSION L"IgnoredUpdateVersion"
#define REG_VALUE_DEBUGLOG L"DebugLog"
#define REG_VALUE_CONFIGURED L"Configured"

// Per-user Default Apps registration. Registering makes SystrayLauncher an
// available MAILTO handler; Windows still requires the user to choose it as
// the default application.
#define REG_REGISTERED_APPLICATIONS_PATH L"Software\\RegisteredApplications"
#define REG_CAPABILITIES_PATH REG_KEY_PATH L"\\Capabilities"
#define REG_CAPABILITIES_REFERENCE L"Software\\JPIT\\SystrayLauncher\\Capabilities"
#define REG_MAILTO_PROGID L"SystrayLauncher.mailto"
#define REG_MAILTO_PROGID_PATH L"Software\\Classes\\" REG_MAILTO_PROGID
#define REG_APPLICATION_PATH \
    L"Software\\Classes\\Applications\\SystrayLauncher.exe"

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
#define ID_TIMER_URL_RESET 7
// How long a hidden window keeps its page state: re-opening within this
// window returns the user exactly where they were. Once it expires the page
// is reset to the configured URL in the background - still hidden - so the
// next open starts fresh without a visible navigation.
#define URL_RESET_AFTER_HIDE_MS 60000
// On-open health verification (see ArmMainHealthCheck): every 100 ms the
// page is asked for a frame heartbeat while the window stays up; a verdict
// falls after 1 s. The lifetime cap only bounds the wait for an in-flight
// rebuild to come up.
#define ID_TIMER_HEALTH_CHECK 10
#define HEALTH_CHECK_INTERVAL_MS 100
#define HEALTH_CHECK_VERDICT_TICKS 10
#define HEALTH_CHECK_LIFETIME_TICKS 600
#define MAIN_WINDOW_TITLE_CCH 768
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

// Backstop for the title's "Loading..." suffix. A navigation's completion
// event can be lost outright (runtime wedged after hibernate, renderer death
// mid-flight), which would otherwise pin the suffix forever; when no
// completion has arrived for this long the suffix is dropped.
#define ID_TIMER_NAV_TITLE_WATCHDOG 11
#define ID_TIMER_AUTO_UPDATE 12
#define AUTO_UPDATE_INTERVAL_MS (60u * 60u * 1000u)
#define NAV_TITLE_WATCHDOG_MS 60000

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

// Posted to the main window when the browser process has died or stopped
// responding and the WebView has to be rebuilt.
#define WM_APP_WEBVIEW_RECREATE (WM_APP + 2)

// Posted by the fallback proxy's worker threads when the address actually
// used for a mapped hostname changes, so the title can show the new route.
#define WM_APP_HOST_ROUTE_CHANGED (WM_APP + 5)

// Static host DNS fallback (see StartStaticHostProxy). A small loopback
// forward proxy inside the launcher process; the browser reaches it through
// --proxy-pac-url and the generated PAC script routes only the configured
// static hostnames to it, so all other traffic stays direct. Per host the
// proxy prefers the mapped address and falls back to standard DNS while the
// mapped address is unreachable, re-trying it at most once per interval.
#define HOST_PROXY_PAC_PATH "/proxy.pac"
#define HOST_PROXY_MAX_MAPPINGS 256
#define HOST_PROXY_MAX_TUNNELS 64
#define HOST_PROXY_LISTEN_BACKLOG 16
#define HOST_PROXY_HEAD_MAX_BYTES (16 * 1024)
#define HOST_PROXY_IO_BUFFER_BYTES (16 * 1024)
// Connect budget for the mapped address, shared by the in-band attempt and
// the side-car probe so both agree on what "reachable" means. Short, so a
// page load that races a dead mapped address stays under a second before
// the DNS fallback takes over.
#define HOST_PROXY_MAPPED_CONNECT_TIMEOUT_MS 800
#define HOST_PROXY_DNS_CONNECT_TIMEOUT_MS 10000
#define HOST_PROXY_PROBE_INTERVAL_MS 60000
#define HOST_PROXY_HEAD_READ_TIMEOUT_MS 15000
#define HOST_PROXY_POLL_TICK_MS 1000
#define HOST_PROXY_SHUTDOWN_WAIT_MS 5000

typedef struct {
    wchar_t url[2048];
    wchar_t windowTitle[256];
    BOOL startMaximized;
    BOOL returnToTargetOnDoubleClick;
    BOOL showInTaskbar;
    BOOL handleMailtoLinks;
    wchar_t mailtoTargetUrl[2048];
    wchar_t onHideJs[4096];
    wchar_t onShowJs[4096];
    BOOL sleepWhenInactive;
    BOOL openNewWindowsExternally;
    BOOL allowRunningInsecureContent;
    wchar_t insecureContentOrigins[2048];
    BOOL useStaticHostMappings;
    wchar_t staticHostMappings[2048];
    BOOL staticHostDnsFallback;
    BOOL lockdownHeader;
    wchar_t lockdownSecret[256];
    BOOL autoCheckForUpdates;
    BOOL debugLogEnabled;
} Configuration;

typedef enum {
    MAILTO_DEFAULT_APPS_NOT_OPENED = 0,
    MAILTO_DEFAULT_APPS_APP_PAGE,
    MAILTO_DEFAULT_APPS_GENERAL_PAGE
} MailtoDefaultAppsOpenResult;

typedef enum {
    JS_VISIBILITY_UNKNOWN = -1,
    JS_VISIBILITY_HIDDEN = 0,
    JS_VISIBILITY_SHOWN = 1
} JsVisibility;

// Globals
static Configuration g_config;
static HWND g_hwnd = NULL;
// Invisible owner used when the main window should stay out of the taskbar.
static HWND g_hwndOwner = NULL;
static ICoreWebView2Controller* g_webViewController = NULL;
static ICoreWebView2* g_webView = NULL;
static ICoreWebView2Environment* g_webViewEnv = NULL;
static NOTIFYICONDATAW g_nid = {0};
static UINT g_WM_TASKBARCREATED = 0;
static UINT g_WM_MAILTO_ACTIVATE = 0;
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
static volatile LONG g_mailtoActivationPending = FALSE;
static volatile LONG g_sleepWhenInactive = FALSE;
static volatile LONG g_openNewWindowsExternally = FALSE;
static volatile LONG g_lockdownHeader = FALSE;
// Whether the match-everything WebResourceRequested filter is registered on
// the current main WebView. Per-instance state: reset on every rebuild.
static BOOL g_lockdownFilterActive = FALSE;
static volatile LONG g_debugLogEnabled = FALSE;
static volatile LONG g_initialPreloadComplete = FALSE;
static volatile LONG g_webViewCreatePending = FALSE;
static volatile LONG g_resumeFailureCount = 0;
// Post-power-resume recovery: set when the machine goes down or comes back up
// and cleared once the WebView has answered a liveness ping (or been rebuilt).
// While set, the page is kept warm so we never snapshot an unverified page.
static volatile LONG g_powerResumePending = FALSE;
static volatile LONG g_webViewPingOutstanding = FALSE;
static int g_powerKickCount = 0;
// On-open health verification. g_presentationUnverified is set on every
// power transition and cleared only once a SHOWN window passes the on-screen
// check: the hidden-time recovery above can prove the runtime answers
// scripts, but never that pixels actually reach the screen - after hibernate
// the two can differ (the "white container"). g_framePongSeen is flipped by
// the requestAnimationFrame heartbeat the page posts back while the check
// runs (see ArmMainHealthCheck).
static volatile LONG g_presentationUnverified = FALSE;
static volatile LONG g_framePongSeen = FALSE;
static int g_healthTicks = 0;       // probing ticks with a live WebView
static int g_healthTotalTicks = 0;  // lifetime of this poll (safety cap)
static BOOL g_healthHealed = FALSE; // one rebuild per open
// UI-thread state used to distinguish a taskbar restore from ordinary resizes.
static BOOL g_mainWindowMinimized = FALSE;
static ULONGLONG g_rebuildBurstStartTick = 0;
static LONG g_rebuildBurstCount = 0;
static BOOL g_mainNavigationLoading = TRUE;
static UINT64 g_mainNavigationId = 0;
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
static BOOL g_configViewReady = FALSE;
static BOOL g_updateConfirmationPending = FALSE;
static int g_cfgShowFallbackTries = 0;
static volatile LONG g_updateCheckPending = FALSE;
static volatile LONG g_updateCheckAutomatic = FALSE;
static BOOL g_updateInstallReady = FALSE;
static volatile LONG g_updateRequestSequence = 0;
static HANDLE g_updateCancelEvent = NULL;
static volatile LONG g_updateSpeedKbps = 0;
static volatile LONG g_updateProgressPosted = FALSE;

// Static host DNS fallback proxy state.
typedef enum {
    HOST_PROXY_UNTESTED = 0,   // mapped address not yet tried this run
    HOST_PROXY_MAPPED_ACTIVE,  // mapped address answered; keep using it
    HOST_PROXY_FALLBACK        // mapped address unreachable; use standard DNS
} HostProxyBreakerState;

typedef struct {
    char host[254];        // lowercase ASCII hostname
    char address[64];      // numeric address literal, IPv6 brackets stripped
    int addressFamily;     // AF_INET or AF_INET6
    HostProxyBreakerState state;  // guarded by g_hostProxyLock
    ULONGLONG lastProbeTick;      // tick of the last failed mapped attempt
    BOOL probeInFlight;           // single-flight guard for side-car probes
    // Address most recently used to reach the host, shown in the window
    // title; starts as the mapped address. Guarded by g_hostProxyLock.
    char currentAddress[64];
} HostProxyMapping;

// Tunnel nodes are owned by their connection thread: the accept thread links
// a node into the list (under g_hostProxyLock) before starting the thread,
// and only the owning thread unlinks, closes and frees it. Everyone else -
// CloseHostTunnelsForMapping, StopStaticHostProxy - may only shutdown() the
// sockets of nodes found on the list while holding the lock, which is safe
// because a linked node cannot be freed concurrently.
typedef struct HostProxyTunnel {
    SOCKET client;
    SOCKET upstream;       // INVALID_SOCKET until connected
    int mappingIndex;      // index into g_hostProxyMappings, -1 before parse
    volatile LONG abortRequested;
    struct HostProxyTunnel* next;
    struct HostProxyTunnel* prev;
} HostProxyTunnel;

static BOOL g_winsockInitialized = FALSE;
// Guards breaker state, probe flags and the tunnel list. Never held across
// a blocking call (connect/getaddrinfo/send/recv/logging).
static CRITICAL_SECTION g_hostProxyLock;
static HostProxyMapping* g_hostProxyMappings = NULL;
static size_t g_hostProxyMappingCount = 0;
static SOCKET g_hostProxyListenSocket = INVALID_SOCKET;
static unsigned short g_hostProxyPort = 0;  // 0 = proxy not running
static HANDLE g_hostProxyAcceptThread = NULL;
static volatile LONG g_hostProxyStopping = FALSE;
static volatile LONG g_hostProxyWorkerCount = 0;  // connection + probe threads
static HostProxyTunnel* g_hostProxyTunnelList = NULL;
static char* g_hostProxyPacScript = NULL;

typedef struct {
    WORD major;
    WORD minor;
    WORD patch;
    WORD build;
} ExecutableVersion;

typedef enum {
    UPDATE_CHECK_SAME = 1,
    UPDATE_CHECK_NEWER,
    UPDATE_CHECK_OLDER,
    UPDATE_CHECK_CANCELLED,
    UPDATE_CHECK_ERROR
} UpdateCheckKind;

typedef struct {
    HWND targetWindow;
    BOOL automatic;
    UpdateCheckKind kind;
    ULONGLONG cacheBuster;
    ExecutableVersion runningVersion;
    ExecutableVersion availableVersion;
    wchar_t message[512];
    wchar_t targetPath[MAX_PATH];
    wchar_t stagedPath[MAX_PATH];
} UpdateCheckTask;

static UpdateCheckTask* volatile g_updatePostedResult = NULL;
static UpdateCheckTask* g_updateNoticeTask = NULL;
static UpdateCheckTask* g_updateReadyTask = NULL;
static wchar_t g_ignoredUpdateVersion[32] = L"";

// Dynamic WebView2 loading
static WCHAR g_extractedDllPath[MAX_PATH] = {0};
typedef HRESULT (STDAPICALLTYPE *PFN_CreateCoreWebView2EnvironmentWithOptions)(
    LPCWSTR, LPCWSTR, ICoreWebView2EnvironmentOptions*,
    ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*);
typedef HRESULT (STDAPICALLTYPE *PFN_GetAvailableCoreWebView2BrowserVersionString)(
    LPCWSTR, LPWSTR*);
static PFN_CreateCoreWebView2EnvironmentWithOptions fnCreateEnvironment = NULL;
static PFN_GetAvailableCoreWebView2BrowserVersionString fnGetAvailableBrowserVersion = NULL;

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
static void GetMainUserDataFolder(wchar_t path[MAX_PATH]);
static void CreateMainWebViewEnvironment(HWND hwnd);
static void HandleUnexpectedBrowserExit(HWND hwnd);
static void RebuildMainWebViewIfDead(void);
static void KickWebViewAfterPowerResume(HWND hwnd);
static void SendMainWebViewLivenessPing(void);
static void CheckMainWebViewLiveness(HWND hwnd);
static void RestartApplication(void);
static void StartUpdateCheck(BOOL automatic);
static BOOL IsOriginListSeparator(wchar_t c);
static BOOL IsValidStaticHostMapping(const wchar_t* mapping,
                                     const wchar_t** separatorOut);
static void CancelUpdateCheck(void);
static void InstallPreparedUpdate(BOOL reopenSettings);
static void DiscardPreparedUpdate(void);
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
static void ResetTargetPageInBackground(void);
static void ArmMainHealthCheck(void);
static void KickMainWebViewComposition(void);
static void OnMainNavigationCompleted(void);
static void BeginMainNavigationTitle(HWND hwnd, UINT64 navigationId);
static BOOL FinishMainNavigationTitle(HWND hwnd, UINT64 navigationId,
                                      BOOL navigationIdKnown);
static void RegisterMainNavigationStartingHandler(ICoreWebView2* webview2);
static void RegisterMainNavigationCompletedHandler(ICoreWebView2* webview2);
static void RegisterMainNewWindowRequestedHandler(ICoreWebView2* webview2);
static void RegisterMainWebMessageHandler(ICoreWebView2* webview2);
static void RegisterMainWebResourceRequestedHandler(ICoreWebView2* webview2);
static void ApplyLockdownRequestFilter(void);
static void GetTargetWindowRect(int* x, int* y, int* w, int* h);

// Registry and config dialog functions
static BOOL LoadConfigFromRegistry(Configuration* config);
static BOOL SaveConfigToRegistry(const Configuration* config);
static BOOL IsFirstLaunch(void);
static void MarkAsConfigured(void);
static void ApplyConfiguration(void);
static BOOL IsValidHttpNavigationUrl(const wchar_t* url);
static BOOL SetMailtoHandlerRegistration(BOOL enabled);
static MailtoDefaultAppsOpenResult OpenMailtoDefaultAppsSettings(void);
static BOOL IsMailtoProtocolInvocation(void);
static BOOL ForwardMailtoActivationToRunningInstance(void);
static BOOL ActivateMailtoDestination(void);
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

// Navigation starting handler (drives the animated native window title)
HRESULT STDMETHODCALLTYPE NavStartingHandler_QueryInterface(
    ICoreWebView2NavigationStartingEventHandler* This,
    REFIID riid, void** ppvObject);
ULONG STDMETHODCALLTYPE NavStartingHandler_AddRef(
    ICoreWebView2NavigationStartingEventHandler* This);
ULONG STDMETHODCALLTYPE NavStartingHandler_Release(
    ICoreWebView2NavigationStartingEventHandler* This);
HRESULT STDMETHODCALLTYPE NavStartingHandler_Invoke(
    ICoreWebView2NavigationStartingEventHandler* This,
    ICoreWebView2* sender, ICoreWebView2NavigationStartingEventArgs* args);

typedef struct {
    ICoreWebView2NavigationStartingEventHandlerVtbl* lpVtbl;
    LONG refCount;
} NavStartingHandler;

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
    config->startMaximized = FALSE;
    config->returnToTargetOnDoubleClick = TRUE;
    config->showInTaskbar = FALSE;
    config->handleMailtoLinks = FALSE;
    config->mailtoTargetUrl[0] = L'\0';
    config->onHideJs[0] = L'\0';
    config->onShowJs[0] = L'\0';
    config->sleepWhenInactive = FALSE;
    config->openNewWindowsExternally = FALSE;
    config->allowRunningInsecureContent = FALSE;
    config->insecureContentOrigins[0] = L'\0';
    config->useStaticHostMappings = FALSE;
    config->staticHostMappings[0] = L'\0';
    config->staticHostDnsFallback = FALSE;
    config->lockdownHeader = FALSE;
    config->lockdownSecret[0] = L'\0';
    config->autoCheckForUpdates = TRUE;
    config->debugLogEnabled = FALSE;

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
    } else if (wcscmp(key, L"startmaximized") == 0) {
        wchar_t c = towlower(value[0]);
        config->startMaximized = (c == L'1' || c == L't' || c == L'y');
    } else if (wcscmp(key, L"returntotargetondoubleclick") == 0) {
        wchar_t c = towlower(value[0]);
        config->returnToTargetOnDoubleClick =
            (c == L'1' || c == L't' || c == L'y');
    } else if (wcscmp(key, L"showintaskbar") == 0) {
        wchar_t c = towlower(value[0]);
        config->showInTaskbar = (c == L'1' || c == L't' || c == L'y');
    } else if (wcscmp(key, L"handlemailtolinks") == 0) {
        wchar_t c = towlower(value[0]);
        config->handleMailtoLinks = (c == L'1' || c == L't' || c == L'y');
    } else if (wcscmp(key, L"mailtotargeturl") == 0) {
        wcscpy_s(config->mailtoTargetUrl, 2048, value);
    } else if (wcscmp(key, L"onhidejs") == 0) {
        wcscpy_s(config->onHideJs, 4096, value);
    } else if (wcscmp(key, L"onshowjs") == 0) {
        wcscpy_s(config->onShowJs, 4096, value);
    } else if (wcscmp(key, L"sleepwheninactive") == 0) {
        wchar_t c = towlower(value[0]);
        config->sleepWhenInactive = (c == L'1' || c == L't' || c == L'y');
    } else if (wcscmp(key, L"opennewwindowsexternally") == 0) {
        wchar_t c = towlower(value[0]);
        config->openNewWindowsExternally = (c == L'1' || c == L't' || c == L'y');
    } else if (wcscmp(key, L"allowrunninginsecurecontent") == 0) {
        wchar_t c = towlower(value[0]);
        config->allowRunningInsecureContent = (c == L'1' || c == L't' || c == L'y');
    } else if (wcscmp(key, L"insecurecontentorigins") == 0) {
        wcscpy_s(config->insecureContentOrigins, 2048, value);
    } else if (wcscmp(key, L"usestatichostmappings") == 0) {
        wchar_t c = towlower(value[0]);
        config->useStaticHostMappings = (c == L'1' || c == L't' || c == L'y');
    } else if (wcscmp(key, L"statichostmappings") == 0) {
        wcscpy_s(config->staticHostMappings, 2048, value);
    } else if (wcscmp(key, L"statichostdnsfallback") == 0) {
        wchar_t c = towlower(value[0]);
        config->staticHostDnsFallback = (c == L'1' || c == L't' || c == L'y');
    } else if (wcscmp(key, L"lockdownheader") == 0) {
        wchar_t c = towlower(value[0]);
        config->lockdownHeader = (c == L'1' || c == L't' || c == L'y');
    } else if (wcscmp(key, L"lockdownsecret") == 0) {
        wcscpy_s(config->lockdownSecret, 256, value);
    } else if (wcscmp(key, L"autocheckforupdates") == 0) {
        wchar_t c = towlower(value[0]);
        config->autoCheckForUpdates = (c == L'1' || c == L't' || c == L'y');
    } else if (wcscmp(key, L"debuglog") == 0) {
        wchar_t c = towlower(value[0]);
        config->debugLogEnabled = (c == L'1' || c == L't' || c == L'y');
    }
}

void CreateDefaultIni(const wchar_t* iniPath) {
    const wchar_t* content = L"# SystrayLauncher Configuration File\n"
                             L"# Lines starting with # or ; are comments\n\n"
                             L"url=https://www.google.com/\n\n"
                             L"windowtitle=Systray Launcher\n"
                             L"startmaximized=false\n"
                             L"returntotargetondoubleclick=true\n"
                             L"showintaskbar=false\n"
                             L"handlemailtolinks=false\n"
                             L"mailtotargeturl=\n";
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

    // Load maximized-window preference (default disabled, preserving the
    // centered 90%-of-work-area behavior used by earlier versions).
    DWORD startMaximizedVal = 0;
    dataSize = sizeof(startMaximizedVal);
    if (RegQueryValueExW(hKey, REG_VALUE_START_MAXIMIZED, NULL, &dataType,
                         (LPBYTE)&startMaximizedVal, &dataSize) == ERROR_SUCCESS) {
        config->startMaximized = (startMaximizedVal != 0);
    } else {
        config->startMaximized = FALSE;
    }

    // The preceding release always returned to the configured target on a tray
    // double-click, so preserve that behavior when the value is absent.
    DWORD returnToTargetVal = 1;
    dataSize = sizeof(returnToTargetVal);
    if (RegQueryValueExW(hKey, REG_VALUE_RETURN_TO_TARGET_ON_DOUBLE_CLICK,
                         NULL, &dataType, (LPBYTE)&returnToTargetVal,
                         &dataSize) == ERROR_SUCCESS) {
        config->returnToTargetOnDoubleClick = (returnToTargetVal != 0);
    } else {
        config->returnToTargetOnDoubleClick = TRUE;
    }

    // Existing installations use an invisible owner to suppress the taskbar
    // button, so the new preference remains disabled when the value is absent.
    DWORD showInTaskbarVal = 0;
    dataSize = sizeof(showInTaskbarVal);
    if (RegQueryValueExW(hKey, REG_VALUE_SHOW_IN_TASKBAR, NULL, &dataType,
                         (LPBYTE)&showInTaskbarVal, &dataSize) == ERROR_SUCCESS) {
        config->showInTaskbar = (showInTaskbarVal != 0);
    } else {
        config->showInTaskbar = FALSE;
    }

    // Email-link handling is opt-in. The destination is kept when disabled so
    // a user can turn the handler off temporarily without losing it.
    DWORD handleMailtoVal = 0;
    dataSize = sizeof(handleMailtoVal);
    if (RegQueryValueExW(hKey, REG_VALUE_HANDLE_MAILTO_LINKS, NULL, &dataType,
                         (LPBYTE)&handleMailtoVal, &dataSize) == ERROR_SUCCESS) {
        config->handleMailtoLinks = (handleMailtoVal != 0);
    } else {
        config->handleMailtoLinks = FALSE;
    }

    config->mailtoTargetUrl[0] = L'\0';
    dataSize = sizeof(config->mailtoTargetUrl);
    if (RegQueryValueExW(hKey, REG_VALUE_MAILTO_TARGET_URL, NULL, &dataType,
                         (LPBYTE)config->mailtoTargetUrl, &dataSize) != ERROR_SUCCESS ||
        dataType != REG_SZ || dataSize < sizeof(wchar_t)) {
        config->mailtoTargetUrl[0] = L'\0';
    } else {
        config->mailtoTargetUrl[2047] = L'\0';
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

    // Load AllowRunningInsecureContent (default disabled)
    DWORD insecureContentVal = 0;
    dataSize = sizeof(insecureContentVal);
    if (RegQueryValueExW(hKey, REG_VALUE_INSECURE_CONTENT, NULL, &dataType,
                         (LPBYTE)&insecureContentVal, &dataSize) == ERROR_SUCCESS) {
        config->allowRunningInsecureContent = (insecureContentVal != 0);
    } else {
        config->allowRunningInsecureContent = FALSE;
    }

    // Load the exact HTTP origins that WebView2 may treat as trustworthy.
    config->insecureContentOrigins[0] = L'\0';
    dataSize = sizeof(config->insecureContentOrigins);
    if (RegQueryValueExW(hKey, REG_VALUE_INSECURE_CONTENT_ORIGINS, NULL, &dataType,
                         (LPBYTE)config->insecureContentOrigins, &dataSize) != ERROR_SUCCESS ||
        dataType != REG_SZ || dataSize < sizeof(wchar_t)) {
        config->insecureContentOrigins[0] = L'\0';
    } else {
        config->insecureContentOrigins[2047] = L'\0';
    }

    // Load per-container static hostname resolution (default disabled).
    DWORD staticHostsVal = 0;
    dataSize = sizeof(staticHostsVal);
    if (RegQueryValueExW(hKey, REG_VALUE_STATIC_HOSTS, NULL, &dataType,
                         (LPBYTE)&staticHostsVal, &dataSize) == ERROR_SUCCESS) {
        config->useStaticHostMappings = (staticHostsVal != 0);
    } else {
        config->useStaticHostMappings = FALSE;
    }

    config->staticHostMappings[0] = L'\0';
    dataSize = sizeof(config->staticHostMappings);
    if (RegQueryValueExW(hKey, REG_VALUE_STATIC_HOST_MAPPINGS, NULL, &dataType,
                         (LPBYTE)config->staticHostMappings, &dataSize) != ERROR_SUCCESS ||
        dataType != REG_SZ || dataSize < sizeof(wchar_t)) {
        config->staticHostMappings[0] = L'\0';
    } else {
        config->staticHostMappings[2047] = L'\0';
    }

    // Load the DNS-fallback mode for static host mappings (default disabled).
    DWORD staticHostFallbackVal = 0;
    dataSize = sizeof(staticHostFallbackVal);
    if (RegQueryValueExW(hKey, REG_VALUE_STATIC_HOST_DNS_FALLBACK, NULL, &dataType,
                         (LPBYTE)&staticHostFallbackVal, &dataSize) == ERROR_SUCCESS) {
        config->staticHostDnsFallback = (staticHostFallbackVal != 0);
    } else {
        config->staticHostDnsFallback = FALSE;
    }

    // Load LockdownHeader (default disabled)
    DWORD lockdownVal = 0;
    dataSize = sizeof(lockdownVal);
    if (RegQueryValueExW(hKey, REG_VALUE_LOCKDOWN, NULL, &dataType,
                         (LPBYTE)&lockdownVal, &dataSize) == ERROR_SUCCESS) {
        config->lockdownHeader = (lockdownVal != 0);
    } else {
        config->lockdownHeader = FALSE;
    }

    // Load the optional shared secret mixed into the lockdown key.
    config->lockdownSecret[0] = L'\0';
    dataSize = sizeof(config->lockdownSecret);
    if (RegQueryValueExW(hKey, REG_VALUE_LOCKDOWN_SECRET, NULL, &dataType,
                         (LPBYTE)config->lockdownSecret, &dataSize) != ERROR_SUCCESS ||
        dataType != REG_SZ || dataSize < sizeof(wchar_t)) {
        config->lockdownSecret[0] = L'\0';
    } else {
        config->lockdownSecret[255] = L'\0';
    }

    // Load automatic update checks (default enabled).
    DWORD autoUpdateVal = 1;
    dataSize = sizeof(autoUpdateVal);
    if (RegQueryValueExW(hKey, REG_VALUE_AUTO_UPDATE, NULL, &dataType,
                         (LPBYTE)&autoUpdateVal, &dataSize) == ERROR_SUCCESS) {
        config->autoCheckForUpdates = (autoUpdateVal != 0);
    } else {
        config->autoCheckForUpdates = TRUE;
    }

    dataSize = sizeof(g_ignoredUpdateVersion);
    if (RegQueryValueExW(hKey, REG_VALUE_IGNORED_UPDATE_VERSION, NULL,
                         &dataType, (LPBYTE)g_ignoredUpdateVersion,
                         &dataSize) != ERROR_SUCCESS ||
        dataType != REG_SZ || dataSize < sizeof(wchar_t)) {
        g_ignoredUpdateVersion[0] = L'\0';
    }
    g_ignoredUpdateVersion[
        (sizeof(g_ignoredUpdateVersion) / sizeof(wchar_t)) - 1] = L'\0';

    // Load DebugLog (default disabled)
    DWORD dbgVal = 0;
    dataSize = sizeof(dbgVal);
    if (RegQueryValueExW(hKey, REG_VALUE_DEBUGLOG, NULL, &dataType, (LPBYTE)&dbgVal, &dataSize) == ERROR_SUCCESS) {
        config->debugLogEnabled = (dbgVal != 0);
    } else {
        config->debugLogEnabled = FALSE;
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
    BOOL success = TRUE;

    // Save URL
    RegSetValueExW(hKey, REG_VALUE_URL, 0, REG_SZ,
                   (const BYTE*)config->url, (DWORD)((wcslen(config->url) + 1) * sizeof(wchar_t)));

    // Save Window Title
    RegSetValueExW(hKey, REG_VALUE_TITLE, 0, REG_SZ,
                   (const BYTE*)config->windowTitle, (DWORD)((wcslen(config->windowTitle) + 1) * sizeof(wchar_t)));

    // Save maximized-window preference.
    DWORD startMaximizedVal = config->startMaximized ? 1 : 0;
    RegSetValueExW(hKey, REG_VALUE_START_MAXIMIZED, 0, REG_DWORD,
                   (const BYTE*)&startMaximizedVal, sizeof(startMaximizedVal));

    DWORD returnToTargetVal = config->returnToTargetOnDoubleClick ? 1 : 0;
    RegSetValueExW(hKey, REG_VALUE_RETURN_TO_TARGET_ON_DOUBLE_CLICK, 0,
                   REG_DWORD, (const BYTE*)&returnToTargetVal,
                   sizeof(returnToTargetVal));

    DWORD showInTaskbarVal = config->showInTaskbar ? 1 : 0;
    RegSetValueExW(hKey, REG_VALUE_SHOW_IN_TASKBAR, 0, REG_DWORD,
                   (const BYTE*)&showInTaskbarVal, sizeof(showInTaskbarVal));

    DWORD handleMailtoVal = config->handleMailtoLinks ? 1 : 0;
    if (RegSetValueExW(hKey, REG_VALUE_HANDLE_MAILTO_LINKS, 0, REG_DWORD,
                       (const BYTE*)&handleMailtoVal,
                       sizeof(handleMailtoVal)) != ERROR_SUCCESS) {
        success = FALSE;
    }
    if (RegSetValueExW(
            hKey, REG_VALUE_MAILTO_TARGET_URL, 0, REG_SZ,
            (const BYTE*)config->mailtoTargetUrl,
            (DWORD)((wcslen(config->mailtoTargetUrl) + 1) *
                    sizeof(wchar_t))) != ERROR_SUCCESS) {
        success = FALSE;
    }

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

    // Save AllowRunningInsecureContent
    DWORD insecureContentVal = config->allowRunningInsecureContent ? 1 : 0;
    RegSetValueExW(hKey, REG_VALUE_INSECURE_CONTENT, 0, REG_DWORD,
                   (const BYTE*)&insecureContentVal, sizeof(insecureContentVal));

    // Save the origin-scoped trust allowlist used by the browser process.
    RegSetValueExW(hKey, REG_VALUE_INSECURE_CONTENT_ORIGINS, 0, REG_SZ,
                   (const BYTE*)config->insecureContentOrigins,
                   (DWORD)((wcslen(config->insecureContentOrigins) + 1) * sizeof(wchar_t)));

    // Save the per-container static hostname resolution rules.
    DWORD staticHostsVal = config->useStaticHostMappings ? 1 : 0;
    RegSetValueExW(hKey, REG_VALUE_STATIC_HOSTS, 0, REG_DWORD,
                   (const BYTE*)&staticHostsVal, sizeof(staticHostsVal));
    RegSetValueExW(hKey, REG_VALUE_STATIC_HOST_MAPPINGS, 0, REG_SZ,
                   (const BYTE*)config->staticHostMappings,
                   (DWORD)((wcslen(config->staticHostMappings) + 1) * sizeof(wchar_t)));
    DWORD staticHostFallbackVal = config->staticHostDnsFallback ? 1 : 0;
    RegSetValueExW(hKey, REG_VALUE_STATIC_HOST_DNS_FALLBACK, 0, REG_DWORD,
                   (const BYTE*)&staticHostFallbackVal, sizeof(staticHostFallbackVal));

    // Save LockdownHeader
    DWORD lockdownVal = config->lockdownHeader ? 1 : 0;
    RegSetValueExW(hKey, REG_VALUE_LOCKDOWN, 0, REG_DWORD,
                   (const BYTE*)&lockdownVal, sizeof(lockdownVal));

    // Save LockdownSecret
    RegSetValueExW(hKey, REG_VALUE_LOCKDOWN_SECRET, 0, REG_SZ,
                   (const BYTE*)config->lockdownSecret,
                   (DWORD)((wcslen(config->lockdownSecret) + 1) * sizeof(wchar_t)));

    // Save automatic update checks.
    DWORD autoUpdateVal = config->autoCheckForUpdates ? 1 : 0;
    RegSetValueExW(hKey, REG_VALUE_AUTO_UPDATE, 0, REG_DWORD,
                   (const BYTE*)&autoUpdateVal, sizeof(autoUpdateVal));
    RegSetValueExW(hKey, REG_VALUE_IGNORED_UPDATE_VERSION, 0, REG_SZ,
                   (const BYTE*)g_ignoredUpdateVersion,
                   (DWORD)((wcslen(g_ignoredUpdateVersion) + 1) *
                           sizeof(wchar_t)));

    // Save DebugLog
    DWORD dbgVal = config->debugLogEnabled ? 1 : 0;
    RegSetValueExW(hKey, REG_VALUE_DEBUGLOG, 0, REG_DWORD,
                   (const BYTE*)&dbgVal, sizeof(dbgVal));

    // Drop the value left behind by versions that had the (never functional)
    // spell-check option.
    RegDeleteValueW(hKey, L"SpellcheckLanguages");

    RegCloseKey(hKey);
    return success;
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

static BOOL SetRegistryString(HKEY root, const wchar_t* subkey,
                              const wchar_t* valueName,
                              const wchar_t* value) {
    HKEY key = NULL;
    DWORD disposition = 0;
    LONG result = RegCreateKeyExW(root, subkey, 0, NULL,
                                  REG_OPTION_NON_VOLATILE, KEY_SET_VALUE,
                                  NULL, &key, &disposition);
    if (result != ERROR_SUCCESS) return FALSE;

    result = RegSetValueExW(key, valueName, 0, REG_SZ, (const BYTE*)value,
                            (DWORD)((wcslen(value) + 1) * sizeof(wchar_t)));
    RegCloseKey(key);
    return result == ERROR_SUCCESS;
}

static BOOL DeleteRegistryTreeIfPresent(HKEY root, const wchar_t* subkey) {
    LONG result = RegDeleteTreeW(root, subkey);
    return result == ERROR_SUCCESS || result == ERROR_FILE_NOT_FOUND ||
           result == ERROR_PATH_NOT_FOUND;
}

static BOOL RemoveMailtoHandlerRegistration(void) {
    BOOL success = TRUE;
    HKEY registeredApps = NULL;
    LONG result = RegOpenKeyExW(HKEY_CURRENT_USER,
                                REG_REGISTERED_APPLICATIONS_PATH, 0,
                                KEY_SET_VALUE, &registeredApps);
    if (result == ERROR_SUCCESS) {
        // Remove both the current friendly registered name and the canonical
        // name used by 1.0.20 so upgrading does not leave a duplicate entry.
        const wchar_t* registeredNames[] = {
            APP_DISPLAY_NAME_WSTRING,
            APP_NAME
        };
        for (size_t i = 0;
             i < sizeof(registeredNames) / sizeof(registeredNames[0]); i++) {
            result = RegDeleteValueW(registeredApps, registeredNames[i]);
            if (result != ERROR_SUCCESS && result != ERROR_FILE_NOT_FOUND) {
                success = FALSE;
            }
        }
        RegCloseKey(registeredApps);
    } else if (result != ERROR_FILE_NOT_FOUND && result != ERROR_PATH_NOT_FOUND) {
        success = FALSE;
    }

    if (!DeleteRegistryTreeIfPresent(HKEY_CURRENT_USER,
                                     REG_CAPABILITIES_PATH)) {
        success = FALSE;
    }
    if (!DeleteRegistryTreeIfPresent(HKEY_CURRENT_USER,
                                     REG_MAILTO_PROGID_PATH)) {
        success = FALSE;
    }
    if (!DeleteRegistryTreeIfPresent(HKEY_CURRENT_USER,
                                     REG_APPLICATION_PATH)) {
        success = FALSE;
    }
    return success;
}

static BOOL IsValidHttpNavigationUrl(const wchar_t* url) {
    if (!url || !url[0]) return FALSE;
    for (const wchar_t* current = url; *current; current++) {
        if (*current <= L' ' || *current == L'\"') return FALSE;
    }

    URL_COMPONENTS components = {0};
    wchar_t hostname[256];
    components.dwStructSize = sizeof(components);
    components.lpszHostName = hostname;
    components.dwHostNameLength =
        (DWORD)(sizeof(hostname) / sizeof(hostname[0]));
    if (!WinHttpCrackUrl(url, 0, 0, &components) ||
        (components.nScheme != INTERNET_SCHEME_HTTP &&
         components.nScheme != INTERNET_SCHEME_HTTPS) ||
        components.dwHostNameLength == 0 ||
        components.dwHostNameLength >=
            sizeof(hostname) / sizeof(hostname[0])) {
        return FALSE;
    }
    return TRUE;
}

static BOOL SetMailtoHandlerRegistration(BOOL enabled) {
    if (!enabled) {
        BOOL removed = RemoveMailtoHandlerRegistration();
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
        return removed;
    }

    if (!IsValidHttpNavigationUrl(g_config.mailtoTargetUrl)) {
        RemoveMailtoHandlerRegistration();
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
        return FALSE;
    }

    wchar_t executablePath[MAX_PATH];
    DWORD pathLength = GetModuleFileNameW(NULL, executablePath, MAX_PATH);
    if (pathLength == 0 || pathLength >= MAX_PATH) return FALSE;

    wchar_t command[MAX_PATH + 32];
    wchar_t icon[MAX_PATH + 8];
    if (swprintf_s(command, sizeof(command) / sizeof(command[0]),
                   L"\"%s\" --mailto \"%%1\"", executablePath) < 0 ||
        swprintf_s(icon, sizeof(icon) / sizeof(icon[0]),
                   L"\"%s\",0", executablePath) < 0) {
        return FALSE;
    }

    // Migrate the canonical RegisteredApplications value written by 1.0.20.
    // The capability's ApplicationName must match its registered value name,
    // otherwise Windows can retain a second filename-based entry.
    BOOL legacyNameRemoved = TRUE;
    HKEY registeredApps = NULL;
    LONG result = RegOpenKeyExW(HKEY_CURRENT_USER,
                                REG_REGISTERED_APPLICATIONS_PATH, 0,
                                KEY_SET_VALUE, &registeredApps);
    if (result == ERROR_SUCCESS) {
        result = RegDeleteValueW(registeredApps, APP_NAME);
        if (result != ERROR_SUCCESS && result != ERROR_FILE_NOT_FOUND) {
            legacyNameRemoved = FALSE;
        }
        RegCloseKey(registeredApps);
    } else if (result != ERROR_FILE_NOT_FOUND && result != ERROR_PATH_NOT_FOUND) {
        legacyNameRemoved = FALSE;
    }

    BOOL success = legacyNameRemoved &&
        SetRegistryString(HKEY_CURRENT_USER, REG_MAILTO_PROGID_PATH, NULL,
                          L"System Tray Launcher email link") &&
        SetRegistryString(HKEY_CURRENT_USER, REG_MAILTO_PROGID_PATH,
                          L"URL Protocol", L"") &&
        SetRegistryString(HKEY_CURRENT_USER, REG_MAILTO_PROGID_PATH,
                          L"FriendlyTypeName",
                          L"System Tray Launcher email link") &&
        SetRegistryString(HKEY_CURRENT_USER,
                          REG_MAILTO_PROGID_PATH L"\\DefaultIcon", NULL,
                          icon) &&
        SetRegistryString(HKEY_CURRENT_USER,
                          REG_MAILTO_PROGID_PATH L"\\shell\\open\\command",
                          NULL, command) &&
        SetRegistryString(HKEY_CURRENT_USER, REG_APPLICATION_PATH,
                          L"FriendlyAppName", APP_DISPLAY_NAME_WSTRING) &&
        SetRegistryString(HKEY_CURRENT_USER,
                          REG_APPLICATION_PATH L"\\DefaultIcon", NULL,
                          icon) &&
        SetRegistryString(HKEY_CURRENT_USER, REG_CAPABILITIES_PATH,
                          L"ApplicationName", APP_DISPLAY_NAME_WSTRING) &&
        SetRegistryString(
            HKEY_CURRENT_USER, REG_CAPABILITIES_PATH,
            L"ApplicationDescription",
            L"Opens email links at the configured SystrayLauncher destination.") &&
        SetRegistryString(HKEY_CURRENT_USER, REG_CAPABILITIES_PATH,
                          L"ApplicationIcon", icon) &&
        SetRegistryString(HKEY_CURRENT_USER,
                          REG_CAPABILITIES_PATH L"\\UrlAssociations",
                          L"mailto", REG_MAILTO_PROGID) &&
        SetRegistryString(HKEY_CURRENT_USER,
                          REG_REGISTERED_APPLICATIONS_PATH,
                          APP_DISPLAY_NAME_WSTRING,
                          REG_CAPABILITIES_REFERENCE);

    if (!success) {
        RemoveMailtoHandlerRegistration();
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
        return FALSE;
    }

    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
    return TRUE;
}

static MailtoDefaultAppsOpenResult OpenMailtoDefaultAppsSettings(void) {
    HWND owner = g_cfgHwnd ? g_cfgHwnd : g_hwnd;
    HINSTANCE result = ShellExecuteW(
        owner, L"open",
        L"ms-settings:defaultapps?registeredAppUser=System%20Tray%20Launcher",
        NULL, NULL, SW_SHOWNORMAL);
    if ((INT_PTR)result > 32) return MAILTO_DEFAULT_APPS_APP_PAGE;

    // Older Windows 10 builds may not understand the per-application query;
    // fall back to the general Default Apps page before asking the user to
    // navigate there manually.
    result = ShellExecuteW(owner, L"open", L"ms-settings:defaultapps",
                           NULL, NULL, SW_SHOWNORMAL);
    return (INT_PTR)result > 32 ? MAILTO_DEFAULT_APPS_GENERAL_PAGE
                               : MAILTO_DEFAULT_APPS_NOT_OPENED;
}

static BOOL IsMailtoProtocolInvocation(void) {
    int argumentCount = 0;
    LPWSTR* arguments = CommandLineToArgvW(GetCommandLineW(), &argumentCount);
    if (!arguments) return FALSE;

    BOOL isMailto =
        argumentCount == 3 && wcscmp(arguments[1], L"--mailto") == 0 &&
        _wcsnicmp(arguments[2], L"mailto:", 7) == 0;
    LocalFree(arguments);
    return isMailto;
}

static BOOL ForwardMailtoActivationToRunningInstance(void) {
    if (g_WM_MAILTO_ACTIVATE == 0) return FALSE;

    // The mutex is created before the main window, so allow a process that is
    // still starting or restarting a short window in which to publish it.
    for (int attempt = 0; attempt < 30; attempt++) {
        HWND target = FindWindowW(L"SystrayLauncherClass", NULL);
        if (target) {
            DWORD processId = 0;
            GetWindowThreadProcessId(target, &processId);
            if (processId != 0) AllowSetForegroundWindow(processId);

            DWORD_PTR activationResult = 0;
            if (SendMessageTimeoutW(target, g_WM_MAILTO_ACTIVATE, 0, 0,
                                    SMTO_ABORTIFHUNG | SMTO_BLOCK, 2000,
                                    &activationResult) != 0 &&
                activationResult != 0) {
                return TRUE;
            }
        }
        Sleep(100);
    }
    return FALSE;
}

static BOOL GetConfiguredUrlHostname(wchar_t* hostname, size_t hostnameCount) {
    if (!hostname || hostnameCount < 2 || !g_config.url[0]) return FALSE;

    hostname[0] = L'\0';
    URL_COMPONENTS components = {0};
    components.dwStructSize = sizeof(components);
    components.lpszHostName = hostname;
    components.dwHostNameLength = (DWORD)hostnameCount;

    if (!WinHttpCrackUrl(g_config.url, 0, 0, &components) ||
        (components.nScheme != INTERNET_SCHEME_HTTP &&
         components.nScheme != INTERNET_SCHEME_HTTPS) ||
        components.dwHostNameLength == 0 ||
        components.dwHostNameLength >= hostnameCount) {
        hostname[0] = L'\0';
        return FALSE;
    }

    hostname[components.dwHostNameLength] = L'\0';
    for (const wchar_t* current = hostname; *current; current++) {
        if (iswspace(*current) || *current < L' ' ||
            wcschr(L"/\\?#@", *current) != NULL) {
            hostname[0] = L'\0';
            return FALSE;
        }
    }
    return TRUE;
}

// True when the configured host is an IP literal rather than a domain name;
// the title's route suffix only applies to domains. Colons and brackets can
// only appear in IPv6 literals, never in hostnames.
static BOOL IsHostnameIpLiteral(const wchar_t* hostname) {
    IN_ADDR v4;
    if (wcschr(hostname, L':') || wcschr(hostname, L'[')) return TRUE;
    return InetPtonW(AF_INET, hostname, &v4) == 1;
}

// Fetches the address to show next to the configured hostname in the title:
// the address the fallback proxy most recently used for it, or the mapped
// address when the strict resolver rules are active. FALSE when the static
// mappings do not apply to this hostname.
static BOOL GetStaticHostDisplayAddress(const wchar_t* hostname,
                                        wchar_t* address, size_t addressCch) {
    if (!g_config.useStaticHostMappings || !hostname[0] || addressCch < 2) {
        return FALSE;
    }
    if (IsHostnameIpLiteral(hostname)) return FALSE;

    if (g_hostProxyPort != 0) {
        // Fallback proxy active: report the route actually in use. Narrow
        // and lowercase for comparison against the ASCII mapping table.
        char narrowHost[254];
        size_t hostLength = wcslen(hostname);
        BOOL found = FALSE;
        if (hostLength >= sizeof(narrowHost)) return FALSE;
        for (size_t i = 0; i < hostLength; i++) {
            wchar_t c = hostname[i];
            if (c >= 0x80) return FALSE;
            if (c >= L'A' && c <= L'Z') c = c - L'A' + L'a';
            narrowHost[i] = (char)c;
        }
        narrowHost[hostLength] = '\0';
        EnterCriticalSection(&g_hostProxyLock);
        for (size_t i = 0; i < g_hostProxyMappingCount; i++) {
            if (strcmp(g_hostProxyMappings[i].host, narrowHost) != 0) continue;
            const char* current = g_hostProxyMappings[i].currentAddress;
            size_t length = strlen(current);
            if (length > 0 && length < addressCch) {
                for (size_t j = 0; j <= length; j++) {
                    address[j] = (wchar_t)(unsigned char)current[j];
                }
                found = TRUE;
            }
            break;
        }
        LeaveCriticalSection(&g_hostProxyLock);
        return found;
    }

    // Strict resolver rules: the mapped address is always the one in use.
    // Mirror the all-or-nothing validation of the argument builder - when
    // any entry is invalid, no rules were applied at all.
    const wchar_t* cursor = g_config.staticHostMappings;
    size_t hostLength = wcslen(hostname);
    while (*cursor) {
        while (*cursor && IsOriginListSeparator(*cursor)) cursor++;
        if (!*cursor) break;

        const wchar_t* begin = cursor;
        while (*cursor && !IsOriginListSeparator(*cursor)) cursor++;
        size_t tokenLength = (size_t)(cursor - begin);
        wchar_t token[2048];
        if (tokenLength == 0 || tokenLength >= sizeof(token) / sizeof(token[0])) {
            return FALSE;
        }
        wmemcpy(token, begin, tokenLength);
        token[tokenLength] = L'\0';

        const wchar_t* separator = NULL;
        if (!IsValidStaticHostMapping(token, &separator)) return FALSE;
        if ((size_t)(separator - token) != hostLength ||
            _wcsnicmp(token, hostname, hostLength) != 0) {
            continue;
        }

        const wchar_t* mappedValue = separator + 1;
        size_t valueLength = wcslen(mappedValue);
        if (valueLength >= 2 && mappedValue[0] == L'[' &&
            mappedValue[valueLength - 1] == L']') {
            mappedValue++;
            valueLength -= 2;
        }
        if (valueLength == 0 || valueLength >= addressCch) return FALSE;
        wmemcpy(address, mappedValue, valueLength);
        address[valueLength] = L'\0';
        return TRUE;
    }
    return FALSE;
}

static void AppendMainWindowTitlePart(wchar_t* title, size_t titleCount,
                                      const wchar_t* part) {
    if (!part || !part[0]) return;
    if (title[0]) wcscat_s(title, titleCount, L" \x2014 ");
    wcscat_s(title, titleCount, part);
}

static void UpdateMainWindowTitle(HWND hwnd) {
    if (!hwnd) return;

    wchar_t title[MAIN_WINDOW_TITLE_CCH] = L"";
    wchar_t hostname[256];
    const wchar_t* configuredTitle =
        g_config.windowTitle[0] ? g_config.windowTitle : APP_NAME;

    AppendMainWindowTitlePart(title, MAIN_WINDOW_TITLE_CCH, configuredTitle);
    if (GetConfiguredUrlHostname(hostname, sizeof(hostname) / sizeof(hostname[0]))) {
        wchar_t mappedAddress[64];
        if (GetStaticHostDisplayAddress(hostname, mappedAddress,
                                        sizeof(mappedAddress) /
                                            sizeof(mappedAddress[0]))) {
            wchar_t hostPart[336];
            swprintf_s(hostPart, sizeof(hostPart) / sizeof(hostPart[0]),
                       L"%s (%s)", hostname, mappedAddress);
            AppendMainWindowTitlePart(title, MAIN_WINDOW_TITLE_CCH, hostPart);
        } else {
            AppendMainWindowTitlePart(title, MAIN_WINDOW_TITLE_CCH, hostname);
        }
    }

    if (g_mainNavigationLoading) {
        AppendMainWindowTitlePart(title, MAIN_WINDOW_TITLE_CCH, L"Loading...");
    }

    SetWindowTextW(hwnd, title);
}

static void BeginMainNavigationTitle(HWND hwnd, UINT64 navigationId) {
    if (!hwnd) return;

    // Every sign of navigation progress re-arms the lost-completion backstop
    // (see ID_TIMER_NAV_TITLE_WATCHDOG).
    SetTimer(hwnd, ID_TIMER_NAV_TITLE_WATCHDOG, NAV_TITLE_WATCHDOG_MS, NULL);

    if (g_mainNavigationLoading && navigationId != 0 &&
        navigationId == g_mainNavigationId) {
        return;  // A redirect keeps the same navigation and animation cadence.
    }

    g_mainNavigationLoading = TRUE;
    g_mainNavigationId = navigationId;
    UpdateMainWindowTitle(hwnd);
}

static BOOL FinishMainNavigationTitle(HWND hwnd, UINT64 navigationId,
                                      BOOL navigationIdKnown) {
    if (!hwnd) return TRUE;

    // A superseded navigation may complete after its replacement has begun.
    // Ignore only completions that are provably older than the navigation
    // being animated (WebView2 assigns rising IDs). A failed navigation can
    // complete under an ID this app never saw start — the error-page commit
    // raises no NavigationStarting — so treating merely unfamiliar IDs as
    // stale would swallow that completion and pin "Loading..." over the
    // error page forever.
    if (g_mainNavigationLoading && navigationIdKnown &&
        g_mainNavigationId != 0 &&
        navigationId < g_mainNavigationId) {
        return FALSE;
    }

    if (g_mainNavigationLoading) {
        g_mainNavigationLoading = FALSE;
        g_mainNavigationId = 0;
        KillTimer(hwnd, ID_TIMER_NAV_TITLE_WATCHDOG);
        UpdateMainWindowTitle(hwnd);
    }
    return TRUE;
}

static void ApplyConfiguration(void) {
    // Update initial URL
    wcscpy_s(g_initialUrl, 2048, g_config.url);

    // Sync the "sleep when inactive" setting
    InterlockedExchange(&g_sleepWhenInactive, g_config.sleepWhenInactive ? TRUE : FALSE);

    // Sync the new-window handling setting (read live by the handler, so a
    // toggle applies without restarting the WebView)
    InterlockedExchange(&g_openNewWindowsExternally, g_config.openNewWindowsExternally ? TRUE : FALSE);

    // Sync the lockdown-header setting (read live by the request handler)
    // and align the request filter with it; a toggle applies without any
    // restart because only the filter, not the handler, changes.
    InterlockedExchange(&g_lockdownHeader, g_config.lockdownHeader ? TRUE : FALSE);
    ApplyLockdownRequestFilter();

    // Sync debug logging (read live by DebugPrint)
    InterlockedExchange(&g_debugLogEnabled, g_config.debugLogEnabled ? TRUE : FALSE);

    // Update the stable title parts immediately. If a navigation is already
    // in progress, its loading suffix remains in place.
    if (g_hwnd) {
        UpdateMainWindowTitle(g_hwnd);
    }

    // Update tray icon tooltip
    if (g_nid.hWnd) {
        wcscpy_s(g_nid.szTip, sizeof(g_nid.szTip)/sizeof(wchar_t), g_config.windowTitle);
        Shell_NotifyIconW(NIM_MODIFY, &g_nid);
    }

    // Navigate to the new URL right away when the window is visible. When it
    // is hidden, apply it in the background instead: the warm preload still
    // holds the previous page, and navigating now means the next open already
    // presents the new page without a visible navigation.
    if (g_webView && g_hwnd) {
        if (IsWindowVisible(g_hwnd)) {
            g_webView->lpVtbl->Navigate(g_webView, g_config.url);
        } else {
            ResetTargetPageInBackground();
        }
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
                            fnGetAvailableBrowserVersion =
                                (PFN_GetAvailableCoreWebView2BrowserVersionString)GetProcAddress(
                                    hMod, "GetAvailableCoreWebView2BrowserVersionString");
                            if (fnCreateEnvironment) return TRUE;
                        }
                    }
                }
            }
        }
    }
    return FALSE;
}

static void RefreshWebView2VersionString(void) {
    if (!fnGetAvailableBrowserVersion) return;

    LPWSTR versionString = NULL;
    if (SUCCEEDED(fnGetAvailableBrowserVersion(NULL, &versionString)) &&
        versionString && versionString[0] != L'\0') {
        wcscpy_s(g_webView2Version,
                 sizeof(g_webView2Version) / sizeof(wchar_t), versionString);
    }
    CoTaskMemFree(versionString);
}

// --- Self update -----------------------------------------------------------

typedef struct {
    HINTERNET session;
    HINTERNET connection;
    HINTERNET request;
} UpdateHttpRequest;

static void CloseUpdateHttpRequest(UpdateHttpRequest* http) {
    if (!http) return;
    if (http->request) WinHttpCloseHandle(http->request);
    if (http->connection) WinHttpCloseHandle(http->connection);
    if (http->session) WinHttpCloseHandle(http->session);
    ZeroMemory(http, sizeof(*http));
}

static void SetUpdateTaskError(UpdateCheckTask* task, LPCWSTR message,
                               DWORD errorCode) {
    if (!task) return;
    task->kind = UPDATE_CHECK_ERROR;
    if (errorCode) {
        swprintf_s(task->message, sizeof(task->message) / sizeof(wchar_t),
                   L"%s (Windows error %lu).", message, (unsigned long)errorCode);
    } else {
        wcscpy_s(task->message, sizeof(task->message) / sizeof(wchar_t), message);
    }
}

static BOOL CancelUpdateTaskIfRequested(UpdateCheckTask* task) {
    if (!task || !g_updateCancelEvent ||
        WaitForSingleObject(g_updateCancelEvent, 0) != WAIT_OBJECT_0) {
        return FALSE;
    }
    task->kind = UPDATE_CHECK_CANCELLED;
    task->message[0] = L'\0';
    return TRUE;
}

// Sample received bytes over monotonic milliseconds; round half up to whole
// kilobytes/second. Download sizes are bounded by UPDATE_MAX_BYTES.
static DWORD CalculateUpdateSpeedKbps(ULONGLONG receivedBytes,
                                      ULONGLONG elapsedMs) {
    if (!elapsedMs) return 0;
    ULONGLONG divisor = elapsedMs * 1024ULL;
    ULONGLONG speed = (receivedBytes * 1000ULL + divisor / 2) / divisor;
    return speed > MAXLONG ? MAXLONG : (DWORD)speed;
}

static void PublishUpdateProgress(UpdateCheckTask* task, DWORD speedKbps) {
    if (!task || CancelUpdateTaskIfRequested(task) ||
        !IsWindow(task->targetWindow)) {
        return;
    }

    InterlockedExchange(&g_updateSpeedKbps, (LONG)speedKbps);
    if (InterlockedCompareExchange(&g_updateProgressPosted, TRUE, FALSE) == FALSE &&
        !PostMessageW(task->targetWindow, WM_APP_UPDATE_PROGRESS, 0, 0)) {
        InterlockedExchange(&g_updateProgressPosted, FALSE);
    }
}

static BOOL OpenUpdateHttpRequest(LPCWSTR verb, ULONGLONG cacheBuster,
                                  UpdateHttpRequest* http, DWORD* statusCode) {
    if (!verb || !http) return FALSE;
    ZeroMemory(http, sizeof(*http));
    if (statusCode) *statusCode = 0;

    wchar_t hostName[256] = L"";
    wchar_t urlPath[2048] = L"";
    wchar_t extraInfo[512] = L"";
    URL_COMPONENTS components = {0};
    components.dwStructSize = sizeof(components);
    components.lpszHostName = hostName;
    components.dwHostNameLength = sizeof(hostName) / sizeof(wchar_t);
    components.lpszUrlPath = urlPath;
    components.dwUrlPathLength = sizeof(urlPath) / sizeof(wchar_t);
    components.lpszExtraInfo = extraInfo;
    components.dwExtraInfoLength = sizeof(extraInfo) / sizeof(wchar_t);
    if (!WinHttpCrackUrl(UPDATE_URL, 0, 0, &components)) return FALSE;

    wchar_t objectName[2560];
    if (wcscpy_s(objectName, sizeof(objectName) / sizeof(wchar_t), urlPath) != 0 ||
        wcscat_s(objectName, sizeof(objectName) / sizeof(wchar_t), extraInfo) != 0) {
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return FALSE;
    }

    // A unique query value makes each button click reach the current branch
    // artifact even when an HTTP proxy or GitHub edge cache retains the
    // previous response. HEAD and GET share the same value within one check.
    wchar_t cacheSuffix[64];
    wchar_t separator = wcschr(objectName, L'?') ? L'&' : L'?';
    int cacheSuffixLength = swprintf_s(cacheSuffix,
        sizeof(cacheSuffix) / sizeof(wchar_t), L"%lcslUpdate=%016llx",
        separator, (unsigned long long)cacheBuster);
    if (cacheSuffixLength <= 0 ||
        wcscat_s(objectName, sizeof(objectName) / sizeof(wchar_t), cacheSuffix) != 0) {
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return FALSE;
    }

    http->session = WinHttpOpen(L"SystrayLauncher Update",
        WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, WINHTTP_NO_PROXY_NAME,
        WINHTTP_NO_PROXY_BYPASS, 0);
    if (!http->session) goto fail;
    WinHttpSetTimeouts(http->session, 10000, 10000, 15000, 30000);

    http->connection = WinHttpConnect(http->session, hostName,
                                      components.nPort, 0);
    if (!http->connection) goto fail;

    DWORD flags = WINHTTP_FLAG_REFRESH;
    if (components.nScheme == INTERNET_SCHEME_HTTPS) flags |= WINHTTP_FLAG_SECURE;
    http->request = WinHttpOpenRequest(http->connection, verb, objectName,
        NULL, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, flags);
    if (!http->request) goto fail;

    static const wchar_t noCacheHeaders[] =
        L"Cache-Control: no-cache, no-store, max-age=0\r\nPragma: no-cache\r\n";
    WinHttpAddRequestHeaders(http->request, noCacheHeaders, (DWORD)-1L,
                            WINHTTP_ADDREQ_FLAG_ADD | WINHTTP_ADDREQ_FLAG_REPLACE);
    if (!WinHttpSendRequest(http->request, WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                            WINHTTP_NO_REQUEST_DATA, 0, 0, 0) ||
        !WinHttpReceiveResponse(http->request, NULL)) {
        goto fail;
    }

    DWORD status = 0;
    DWORD statusSize = sizeof(status);
    if (!WinHttpQueryHeaders(http->request,
            WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
            WINHTTP_HEADER_NAME_BY_INDEX, &status, &statusSize,
            WINHTTP_NO_HEADER_INDEX)) {
        goto fail;
    }
    if (statusCode) *statusCode = status;
    if (status != 200) {
        CloseUpdateHttpRequest(http);
        SetLastError(ERROR_WINHTTP_INVALID_SERVER_RESPONSE);
        return FALSE;
    }
    return TRUE;

fail: {
        DWORD errorCode = GetLastError();
        CloseUpdateHttpRequest(http);
        SetLastError(errorCode);
        return FALSE;
    }
}

static BOOL QueryUpdateContentLength(HINTERNET request, ULONGLONG* size) {
    if (!request || !size) return FALSE;
    wchar_t lengthText[64] = L"";
    DWORD lengthBytes = sizeof(lengthText);
    if (!WinHttpQueryHeaders(request, WINHTTP_QUERY_CONTENT_LENGTH,
            WINHTTP_HEADER_NAME_BY_INDEX, lengthText, &lengthBytes,
            WINHTTP_NO_HEADER_INDEX)) {
        return FALSE;
    }
    lengthText[(sizeof(lengthText) / sizeof(wchar_t)) - 1] = L'\0';

    wchar_t* end = NULL;
    unsigned long long parsed = _wcstoui64(lengthText, &end, 10);
    if (end == lengthText || !end || *end != L'\0' || parsed == 0) {
        SetLastError(ERROR_INVALID_DATA);
        return FALSE;
    }
    *size = (ULONGLONG)parsed;
    return TRUE;
}

static BOOL GetExecutableVersion(LPCWSTR path, ExecutableVersion* version) {
    if (!path || !*path || !version) {
        SetLastError(ERROR_INVALID_PARAMETER);
        return FALSE;
    }

    DWORD ignored = 0;
    DWORD infoSize = GetFileVersionInfoSizeW(path, &ignored);
    if (infoSize == 0) {
        DWORD errorCode = GetLastError();
        SetLastError(errorCode ? errorCode : ERROR_RESOURCE_DATA_NOT_FOUND);
        return FALSE;
    }

    BYTE* infoData = (BYTE*)malloc(infoSize);
    if (!infoData) {
        SetLastError(ERROR_NOT_ENOUGH_MEMORY);
        return FALSE;
    }

    if (!GetFileVersionInfoW(path, 0, infoSize, infoData)) {
        DWORD errorCode = GetLastError();
        free(infoData);
        SetLastError(errorCode ? errorCode : ERROR_INVALID_DATA);
        return FALSE;
    }

    VS_FIXEDFILEINFO* fixedInfo = NULL;
    UINT fixedInfoSize = 0;
    if (!VerQueryValueW(infoData, L"\\", (LPVOID*)&fixedInfo, &fixedInfoSize) ||
        !fixedInfo || fixedInfoSize < sizeof(*fixedInfo) ||
        fixedInfo->dwSignature != VS_FFI_SIGNATURE) {
        free(infoData);
        SetLastError(ERROR_INVALID_DATA);
        return FALSE;
    }

    version->major = HIWORD(fixedInfo->dwFileVersionMS);
    version->minor = LOWORD(fixedInfo->dwFileVersionMS);
    version->patch = HIWORD(fixedInfo->dwFileVersionLS);
    version->build = LOWORD(fixedInfo->dwFileVersionLS);
    free(infoData);
    return TRUE;
}

static int CompareExecutableVersions(const ExecutableVersion* left,
                                     const ExecutableVersion* right) {
    const WORD leftParts[] = {
        left->major, left->minor, left->patch, left->build
    };
    const WORD rightParts[] = {
        right->major, right->minor, right->patch, right->build
    };
    for (size_t index = 0; index < sizeof(leftParts) / sizeof(leftParts[0]); index++) {
        if (leftParts[index] < rightParts[index]) return -1;
        if (leftParts[index] > rightParts[index]) return 1;
    }
    return 0;
}

static void FormatExecutableVersion(const ExecutableVersion* version,
                                    wchar_t* text, size_t textCch) {
    if (!version || !text || textCch == 0) return;
    if (swprintf_s(text, textCch, L"%u.%u.%u.%u",
                   (unsigned int)version->major,
                   (unsigned int)version->minor,
                   (unsigned int)version->patch,
                   (unsigned int)version->build) <= 0) {
        text[0] = L'\0';
    }
}

static void FormatExecutableVersionForDisplay(const ExecutableVersion* version,
                                              wchar_t* text, size_t textCch) {
    if (!version || !text || textCch == 0) return;
    if (version->build == 0) {
        if (swprintf_s(text, textCch, L"%u.%u.%u",
                       (unsigned int)version->major,
                       (unsigned int)version->minor,
                       (unsigned int)version->patch) <= 0) {
            text[0] = L'\0';
        }
        return;
    }
    FormatExecutableVersion(version, text, textCch);
}

static BOOL BuildUpdateTempPath(wchar_t path[MAX_PATH], LPCWSTR role,
                                DWORD processId) {
    if (!path || !role || !*role || processId == 0) {
        SetLastError(ERROR_INVALID_PARAMETER);
        return FALSE;
    }

    wchar_t tempDirectory[MAX_PATH];
    DWORD tempLength = GetTempPathW(MAX_PATH, tempDirectory);
    if (tempLength == 0) return FALSE;
    if (tempLength >= MAX_PATH) {
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return FALSE;
    }

    int length = swprintf_s(path, MAX_PATH, L"%s%s-%s-%lu.exe",
                            tempDirectory, APP_NAME, role,
                            (unsigned long)processId);
    if (length <= 0 || length >= MAX_PATH) {
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return FALSE;
    }
    return TRUE;
}

static BOOL DeleteUpdateTempFile(LPCWSTR path) {
    if (!path || !*path) return FALSE;
    SetFileAttributesW(path, FILE_ATTRIBUTE_NORMAL);
    if (DeleteFileW(path)) return TRUE;
    DWORD errorCode = GetLastError();
    return errorCode == ERROR_FILE_NOT_FOUND || errorCode == ERROR_PATH_NOT_FOUND;
}

static BOOL QueryRemoteUpdateSize(UpdateCheckTask* task, ULONGLONG* size) {
    if (CancelUpdateTaskIfRequested(task)) return FALSE;

    UpdateHttpRequest http;
    DWORD status = 0;
    if (!OpenUpdateHttpRequest(L"HEAD", task->cacheBuster, &http, &status)) {
        DWORD errorCode = GetLastError();
        if (CancelUpdateTaskIfRequested(task)) return FALSE;
        if (status) {
            swprintf_s(task->message, sizeof(task->message) / sizeof(wchar_t),
                       L"The update server returned HTTP status %lu.",
                       (unsigned long)status);
            task->kind = UPDATE_CHECK_ERROR;
        } else {
            SetUpdateTaskError(task, L"Could not contact the update server", errorCode);
        }
        return FALSE;
    }
    if (CancelUpdateTaskIfRequested(task)) {
        CloseUpdateHttpRequest(&http);
        return FALSE;
    }

    BOOL ok = QueryUpdateContentLength(http.request, size);
    DWORD errorCode = ok ? ERROR_SUCCESS : GetLastError();
    CloseUpdateHttpRequest(&http);
    if (CancelUpdateTaskIfRequested(task)) return FALSE;
    if (!ok) {
        SetUpdateTaskError(task, L"The update server did not report a valid file size",
                           errorCode);
        return FALSE;
    }
    if (*size > UPDATE_MAX_BYTES) {
        SetUpdateTaskError(task, L"The available update is unexpectedly large", 0);
        return FALSE;
    }
    return TRUE;
}

static BOOL DownloadUpdateFile(UpdateCheckTask* task, ULONGLONG expectedSize) {
    if (CancelUpdateTaskIfRequested(task)) return FALSE;

    UpdateHttpRequest http;
    DWORD status = 0;
    if (!OpenUpdateHttpRequest(L"GET", task->cacheBuster, &http, &status)) {
        DWORD errorCode = GetLastError();
        if (CancelUpdateTaskIfRequested(task)) return FALSE;
        if (status) {
            swprintf_s(task->message, sizeof(task->message) / sizeof(wchar_t),
                       L"The update download returned HTTP status %lu.",
                       (unsigned long)status);
            task->kind = UPDATE_CHECK_ERROR;
        } else {
            SetUpdateTaskError(task, L"Could not download the update", errorCode);
        }
        return FALSE;
    }
    if (CancelUpdateTaskIfRequested(task)) {
        CloseUpdateHttpRequest(&http);
        return FALSE;
    }

    ULONGLONG downloadSize = 0;
    if (QueryUpdateContentLength(http.request, &downloadSize) &&
        downloadSize != expectedSize) {
        CloseUpdateHttpRequest(&http);
        SetUpdateTaskError(task,
            L"The available update changed while it was being downloaded. Try again", 0);
        return FALSE;
    }
    if (CancelUpdateTaskIfRequested(task)) {
        CloseUpdateHttpRequest(&http);
        return FALSE;
    }

    DeleteUpdateTempFile(task->stagedPath);
    HANDLE file = CreateFileW(task->stagedPath, GENERIC_WRITE, 0, NULL,
                              CREATE_NEW, FILE_ATTRIBUTE_NORMAL, NULL);
    if (file == INVALID_HANDLE_VALUE) {
        DWORD errorCode = GetLastError();
        CloseUpdateHttpRequest(&http);
        SetUpdateTaskError(task, L"Could not create the staged update file", errorCode);
        return FALSE;
    }

    BOOL ok = TRUE;
    ULONGLONG totalWritten = 0;
    ULONGLONG speedWindowBytes = 0;
    ULONGLONG speedWindowStarted = GetTickCount64();
    BYTE buffer[64 * 1024];
    while (ok) {
        if (CancelUpdateTaskIfRequested(task)) {
            ok = FALSE;
            break;
        }

        DWORD bytesRead = 0;
        if (!WinHttpReadData(http.request, buffer, sizeof(buffer), &bytesRead)) {
            DWORD errorCode = GetLastError();
            if (!CancelUpdateTaskIfRequested(task)) {
                SetUpdateTaskError(task, L"The update download was interrupted", errorCode);
            }
            ok = FALSE;
            break;
        }
        if (CancelUpdateTaskIfRequested(task)) {
            ok = FALSE;
            break;
        }
        if (bytesRead == 0) break;
        if (totalWritten + bytesRead > expectedSize) {
            SetUpdateTaskError(task, L"The downloaded update has an invalid size", 0);
            ok = FALSE;
            break;
        }

        DWORD bytesWritten = 0;
        if (!WriteFile(file, buffer, bytesRead, &bytesWritten, NULL)) {
            SetUpdateTaskError(task, L"Could not write the staged update", GetLastError());
            ok = FALSE;
            break;
        }
        if (bytesWritten != bytesRead) {
            SetUpdateTaskError(task, L"Could not write the staged update",
                               ERROR_WRITE_FAULT);
            ok = FALSE;
            break;
        }
        totalWritten += bytesWritten;
        speedWindowBytes += bytesRead;

        ULONGLONG now = GetTickCount64();
        ULONGLONG elapsed = now - speedWindowStarted;
        if (elapsed >= UPDATE_PROGRESS_INTERVAL_MS) {
            PublishUpdateProgress(task,
                CalculateUpdateSpeedKbps(speedWindowBytes, elapsed));
            speedWindowBytes = 0;
            speedWindowStarted = now;
        }
    }

    if (ok && CancelUpdateTaskIfRequested(task)) ok = FALSE;
    if (ok && totalWritten != expectedSize) {
        SetUpdateTaskError(task, L"The downloaded update is incomplete", 0);
        ok = FALSE;
    }
    if (ok && !FlushFileBuffers(file)) {
        SetUpdateTaskError(task, L"Could not finish writing the staged update", GetLastError());
        ok = FALSE;
    }
    CloseHandle(file);
    CloseUpdateHttpRequest(&http);

    if (ok && CancelUpdateTaskIfRequested(task)) ok = FALSE;
    DWORD binaryType = 0;
    if (ok && (!GetBinaryTypeW(task->stagedPath, &binaryType) ||
               binaryType != SCS_64BIT_BINARY)) {
        SetUpdateTaskError(task, L"The downloaded file is not a valid 64-bit application", 0);
        ok = FALSE;
    }
    if (!ok) DeleteUpdateTempFile(task->stagedPath);
    return ok;
}

static void DiscardUpdateTask(UpdateCheckTask* task) {
    if (!task) return;
    if (task->stagedPath[0]) {
        DeleteUpdateTempFile(task->stagedPath);
    }
    free(task);
}

static void PublishUpdateTask(UpdateCheckTask* task) {
    CancelUpdateTaskIfRequested(task);
    InterlockedExchange(&g_updateCheckPending, FALSE);
    InterlockedExchange(&g_updateCheckAutomatic, FALSE);
    if (!task || !IsWindow(task->targetWindow)) {
        DiscardUpdateTask(task);
        return;
    }

    UpdateCheckTask* previous = (UpdateCheckTask*)InterlockedExchangePointer(
        (PVOID volatile*)&g_updatePostedResult, task);
    DiscardUpdateTask(previous);
    if (!PostMessageW(task->targetWindow, WM_APP_UPDATE_RESULT, 0, 0)) {
        UpdateCheckTask* unclaimed = (UpdateCheckTask*)InterlockedExchangePointer(
            (PVOID volatile*)&g_updatePostedResult, NULL);
        DiscardUpdateTask(unclaimed);
    }
}

static DWORD WINAPI UpdateCheckThread(LPVOID parameter) {
    UpdateCheckTask* task = (UpdateCheckTask*)parameter;
    if (CancelUpdateTaskIfRequested(task)) {
        PublishUpdateTask(task);
        return 0;
    }

    DWORD pathLength = GetModuleFileNameW(NULL, task->targetPath,
                                         sizeof(task->targetPath) / sizeof(wchar_t));
    if (pathLength == 0 || pathLength >= sizeof(task->targetPath) / sizeof(wchar_t)) {
        SetUpdateTaskError(task, L"Could not determine the running executable path",
                           GetLastError());
        PublishUpdateTask(task);
        return 0;
    }

    if (!GetExecutableVersion(task->targetPath, &task->runningVersion)) {
        SetUpdateTaskError(task, L"Could not read the running application version",
                           GetLastError());
        PublishUpdateTask(task);
        return 0;
    }
    if (CancelUpdateTaskIfRequested(task)) {
        PublishUpdateTask(task);
        return 0;
    }

    ULONGLONG remoteSize = 0;
    if (!QueryRemoteUpdateSize(task, &remoteSize)) {
        PublishUpdateTask(task);
        return 0;
    }
    if (!BuildUpdateTempPath(task->stagedPath, L"download",
                             GetCurrentProcessId())) {
        SetUpdateTaskError(task, L"Could not create the temporary update path",
                           GetLastError());
        PublishUpdateTask(task);
        return 0;
    }
    if (CancelUpdateTaskIfRequested(task)) {
        PublishUpdateTask(task);
        return 0;
    }

    if (!DownloadUpdateFile(task, remoteSize)) {
        PublishUpdateTask(task);
        return 0;
    }
    if (CancelUpdateTaskIfRequested(task)) {
        PublishUpdateTask(task);
        return 0;
    }

    if (!GetExecutableVersion(task->stagedPath, &task->availableVersion)) {
        SetUpdateTaskError(task,
            L"The downloaded application does not contain valid version information",
            GetLastError());
        PublishUpdateTask(task);
        return 0;
    }
    if (CancelUpdateTaskIfRequested(task)) {
        PublishUpdateTask(task);
        return 0;
    }

    DebugPrint(L"[INFO] Update versions: running %u.%u.%u.%u, available %u.%u.%u.%u\n",
               (unsigned int)task->runningVersion.major,
               (unsigned int)task->runningVersion.minor,
               (unsigned int)task->runningVersion.patch,
               (unsigned int)task->runningVersion.build,
               (unsigned int)task->availableVersion.major,
               (unsigned int)task->availableVersion.minor,
               (unsigned int)task->availableVersion.patch,
               (unsigned int)task->availableVersion.build);
    int comparison = CompareExecutableVersions(&task->availableVersion,
                                               &task->runningVersion);
    task->kind = comparison > 0 ? UPDATE_CHECK_NEWER
               : comparison < 0 ? UPDATE_CHECK_OLDER
                                : UPDATE_CHECK_SAME;
    PublishUpdateTask(task);
    return 0;
}

static BOOL ParseUpdateProcessId(LPCWSTR text, DWORD* processId) {
    if (!text || !processId || !*text) return FALSE;
    wchar_t* end = NULL;
    unsigned long value = wcstoul(text, &end, 10);
    if (!end || *end != L'\0' || value == 0) return FALSE;
    *processId = (DWORD)value;
    return TRUE;
}

static BOOL ValidateUpdateTempFilePair(LPCWSTR helperPath,
                                       LPCWSTR stagedPath,
                                       DWORD processId) {
    if (!helperPath || !stagedPath || !*helperPath || !*stagedPath) return FALSE;

    wchar_t expectedHelperName[96], expectedStagedName[96];
    int helperNameLength = swprintf_s(expectedHelperName,
        sizeof(expectedHelperName) / sizeof(wchar_t),
        APP_NAME L"-updater-%lu.exe", (unsigned long)processId);
    int stagedNameLength = swprintf_s(expectedStagedName,
        sizeof(expectedStagedName) / sizeof(wchar_t),
        APP_NAME L"-download-%lu.exe", (unsigned long)processId);
    if (helperNameLength <= 0 || stagedNameLength <= 0 ||
        _wcsicmp(PathFindFileNameW(helperPath), expectedHelperName) != 0 ||
        _wcsicmp(PathFindFileNameW(stagedPath), expectedStagedName) != 0) {
        return FALSE;
    }

    wchar_t helperDirectory[MAX_PATH], stagedDirectory[MAX_PATH];
    if (wcscpy_s(helperDirectory, MAX_PATH, helperPath) != 0 ||
        wcscpy_s(stagedDirectory, MAX_PATH, stagedPath) != 0 ||
        !PathRemoveFileSpecW(helperDirectory) ||
        !PathRemoveFileSpecW(stagedDirectory)) {
        return FALSE;
    }
    return _wcsicmp(helperDirectory, stagedDirectory) == 0;
}

static HANDLE DuplicateUpdateLaunchToken(HANDLE process) {
    HANDLE processToken = NULL;
    HANDLE launchToken = NULL;
    if (!process ||
        !OpenProcessToken(process, TOKEN_QUERY | TOKEN_DUPLICATE,
                          &processToken)) {
        return NULL;
    }
    DuplicateTokenEx(processToken, MAXIMUM_ALLOWED, NULL,
                     SecurityImpersonation, TokenPrimary, &launchToken);
    CloseHandle(processToken);
    return launchToken;
}

static BOOL LaunchUpdateTarget(LPCWSTR targetPath, LPCWSTR stagedPath,
                               LPCWSTR helperPath, DWORD helperProcessId,
                               DWORD oldProcessId, HANDLE launchToken,
                               BOOL successfulUpdate, BOOL reopenSettings) {
    wchar_t commandLine[MAX_PATH * 3 + 256];
    LPCWSTR finishAction = successfulUpdate
        ? L"--finish-update"
        : L"--finish-update-cleanup";
    int commandLength = swprintf_s(commandLine,
        sizeof(commandLine) / sizeof(wchar_t),
        L"\"%s\" %s %lu %lu \"%s\" \"%s\"%s", targetPath, finishAction,
        (unsigned long)helperProcessId, (unsigned long)oldProcessId,
        stagedPath, helperPath,
        successfulUpdate && reopenSettings ? L" --reopen-settings" : L"");
    if (commandLength <= 0 ||
        commandLength >= (int)(sizeof(commandLine) / sizeof(wchar_t))) {
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return FALSE;
    }

    STARTUPINFOW startupInfo = {0};
    startupInfo.cb = sizeof(startupInfo);
    PROCESS_INFORMATION processInfo = {0};
    BOOL launched = FALSE;
    if (launchToken) {
        wchar_t tokenCommandLine[MAX_PATH * 3 + 256];
        wcscpy_s(tokenCommandLine,
                 sizeof(tokenCommandLine) / sizeof(wchar_t), commandLine);
        LPVOID environment = NULL;
        BOOL hasEnvironment = CreateEnvironmentBlock(&environment, launchToken,
                                                     FALSE);
        launched = CreateProcessWithTokenW(launchToken, 0, targetPath,
                                           tokenCommandLine,
                                           hasEnvironment
                                               ? CREATE_UNICODE_ENVIRONMENT : 0,
                                           environment, NULL,
                                           &startupInfo, &processInfo);
        if (environment) DestroyEnvironmentBlock(environment);
    }
    if (!launched) {
        ZeroMemory(&processInfo, sizeof(processInfo));
        launched = CreateProcessW(targetPath, commandLine, NULL, NULL, FALSE,
                                  0, NULL, NULL, &startupInfo, &processInfo);
    }
    if (launched) {
        CloseHandle(processInfo.hProcess);
        CloseHandle(processInfo.hThread);
    }
    return launched;
}

static int RestartAfterUpdateFailure(LPCWSTR targetPath, LPCWSTR stagedPath,
                                     LPCWSTR helperPath, DWORD oldProcessId,
                                     HANDLE launchToken, LPCWSTR message) {
    MessageBoxW(NULL, message, APP_NAME L" Update", MB_OK | MB_ICONERROR);
    LaunchUpdateTarget(targetPath, stagedPath, helperPath,
                       GetCurrentProcessId(), oldProcessId, launchToken, FALSE, FALSE);
    if (launchToken) CloseHandle(launchToken);
    DeleteUpdateTempFile(stagedPath);
    SetFileAttributesW(helperPath, FILE_ATTRIBUTE_NORMAL);
    MoveFileExW(helperPath, NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
    return 1;
}

static int RunUpdateApplyHelper(DWORD oldProcessId, LPCWSTR readyEventName,
                                LPCWSTR targetPath, LPCWSTR stagedPath,
                                BOOL reopenSettings) {
    wchar_t expectedEventPrefix[96];
    int prefixLength = swprintf_s(expectedEventPrefix,
        sizeof(expectedEventPrefix) / sizeof(wchar_t),
        L"Local\\SystrayLauncher_UpdateReady_%lu_",
        (unsigned long)oldProcessId);
    if (prefixLength <= 0 || !readyEventName ||
        _wcsnicmp(readyEventName, expectedEventPrefix,
                  (size_t)prefixLength) != 0) {
        return ERROR_INVALID_DATA;
    }

    HANDLE readyEvent = OpenEventW(EVENT_MODIFY_STATE, FALSE, readyEventName);
    if (!readyEvent) return (int)GetLastError();

    HANDLE oldProcess = OpenProcess(SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION,
                                    FALSE, oldProcessId);
    if (!oldProcess) {
        DWORD errorCode = GetLastError();
        CloseHandle(readyEvent);
        return (int)errorCode;
    }

    wchar_t oldProcessPath[MAX_PATH];
    DWORD oldProcessPathLength = sizeof(oldProcessPath) / sizeof(wchar_t);
    if (!QueryFullProcessImageNameW(oldProcess, 0, oldProcessPath,
                                    &oldProcessPathLength)) {
        DWORD errorCode = GetLastError();
        CloseHandle(oldProcess);
        CloseHandle(readyEvent);
        return (int)errorCode;
    }
    if (_wcsicmp(oldProcessPath, targetPath) != 0) {
        CloseHandle(oldProcess);
        CloseHandle(readyEvent);
        return ERROR_INVALID_DATA;
    }

    wchar_t helperPath[MAX_PATH];
    DWORD helperPathLength = GetModuleFileNameW(NULL, helperPath,
                                               sizeof(helperPath) / sizeof(wchar_t));
    DWORD binaryType = 0;
    if (helperPathLength == 0 || helperPathLength >= MAX_PATH ||
        !ValidateUpdateTempFilePair(helperPath, stagedPath, oldProcessId)) {
        CloseHandle(oldProcess);
        CloseHandle(readyEvent);
        return ERROR_INVALID_DATA;
    }
    if (!GetBinaryTypeW(stagedPath, &binaryType)) {
        DWORD errorCode = GetLastError();
        CloseHandle(oldProcess);
        CloseHandle(readyEvent);
        return (int)errorCode;
    }
    if (binaryType != SCS_64BIT_BINARY) {
        CloseHandle(oldProcess);
        CloseHandle(readyEvent);
        return ERROR_BAD_EXE_FORMAT;
    }

    // The helper is elevated only for file replacement. Preserve a primary
    // token from the original process so the restarted launcher normally
    // returns to the user's non-elevated session.
    HANDLE launchToken = DuplicateUpdateLaunchToken(oldProcess);

    // Only let the parent exit once this helper has verified every path and
    // owns the process handle it must wait on.
    if (!SetEvent(readyEvent)) {
        DWORD errorCode = GetLastError();
        CloseHandle(oldProcess);
        CloseHandle(readyEvent);
        if (launchToken) CloseHandle(launchToken);
        return (int)errorCode;
    }
    CloseHandle(readyEvent);

    DWORD waitResult = WaitForSingleObject(oldProcess, UPDATE_HELPER_WAIT_MS);
    CloseHandle(oldProcess);
    if (waitResult != WAIT_OBJECT_0) {
        if (launchToken) CloseHandle(launchToken);
        DeleteUpdateTempFile(stagedPath);
        SetFileAttributesW(helperPath, FILE_ATTRIBUTE_NORMAL);
        MoveFileExW(helperPath, NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
        MessageBoxW(NULL, L"The running application did not close in time.",
                    APP_NAME L" Update", MB_OK | MB_ICONERROR);
        return ERROR_TIMEOUT;
    }

    wchar_t replacementPath[MAX_PATH], backupPath[MAX_PATH];
    DWORD helperProcessId = GetCurrentProcessId();
    int replacementLength = swprintf_s(replacementPath,
        sizeof(replacementPath) / sizeof(wchar_t), L"%s.new.%lu.exe",
        targetPath, (unsigned long)helperProcessId);
    int backupLength = swprintf_s(backupPath,
        sizeof(backupPath) / sizeof(wchar_t), L"%s.backup.%lu.exe",
        targetPath, (unsigned long)helperProcessId);
    if (replacementLength <= 0 || replacementLength >= MAX_PATH ||
        backupLength <= 0 || backupLength >= MAX_PATH) {
        return RestartAfterUpdateFailure(targetPath, stagedPath, helperPath,
            oldProcessId, launchToken,
            L"The update paths were too long. The previous version will restart.");
    }

    SetFileAttributesW(replacementPath, FILE_ATTRIBUTE_NORMAL);
    DeleteFileW(replacementPath);
    SetFileAttributesW(backupPath, FILE_ATTRIBUTE_NORMAL);
    DeleteFileW(backupPath);
    if (!CopyFileW(stagedPath, replacementPath, FALSE)) {
        return RestartAfterUpdateFailure(targetPath, stagedPath, helperPath,
            oldProcessId, launchToken,
            L"The update could not be prepared. The previous version will restart.");
    }

    DWORD targetAttributes = GetFileAttributesW(targetPath);
    BOOL clearedReadOnly = FALSE;
    if (targetAttributes != INVALID_FILE_ATTRIBUTES &&
        (targetAttributes & FILE_ATTRIBUTE_READONLY)) {
        clearedReadOnly = SetFileAttributesW(
            targetPath, targetAttributes & ~FILE_ATTRIBUTE_READONLY);
    }
    if (!ReplaceFileW(targetPath, replacementPath, backupPath,
                      REPLACEFILE_WRITE_THROUGH, NULL, NULL)) {
        if (clearedReadOnly) SetFileAttributesW(targetPath, targetAttributes);
        DeleteFileW(replacementPath);
        return RestartAfterUpdateFailure(targetPath, stagedPath, helperPath,
            oldProcessId, launchToken,
            L"The executable could not be replaced. The previous version will restart.");
    }

    if (!LaunchUpdateTarget(targetPath, stagedPath, helperPath,
                            helperProcessId, oldProcessId, launchToken, TRUE,
                            reopenSettings)) {
        DeleteFileW(targetPath);
        if (!MoveFileExW(backupPath, targetPath,
                         MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
            if (launchToken) CloseHandle(launchToken);
            DeleteUpdateTempFile(stagedPath);
            SetFileAttributesW(helperPath, FILE_ATTRIBUTE_NORMAL);
            MoveFileExW(helperPath, NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
            MessageBoxW(NULL,
                L"The updated application could not start and the previous executable "
                L"could not be restored. A backup remains beside the application.",
                APP_NAME L" Update", MB_OK | MB_ICONERROR);
            return 1;
        }
        return RestartAfterUpdateFailure(targetPath, stagedPath, helperPath,
            oldProcessId, launchToken,
            L"The updated application could not start. The previous version was restored.");
    }
    if (launchToken) CloseHandle(launchToken);

    if (!DeleteFileW(backupPath)) {
        MoveFileExW(backupPath, NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
    }
    DeleteUpdateTempFile(stagedPath);
    SetFileAttributesW(helperPath, FILE_ATTRIBUTE_NORMAL);
    MoveFileExW(helperPath, NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
    return 0;
}

static BOOL FinishUpdateCleanup(DWORD helperProcessId, DWORD oldProcessId,
                                LPCWSTR stagedPath, LPCWSTR helperPath) {
    wchar_t targetPath[MAX_PATH], expectedStagedPath[MAX_PATH];
    wchar_t expectedHelperPath[MAX_PATH];
    DWORD targetLength = GetModuleFileNameW(NULL, targetPath,
                                           sizeof(targetPath) / sizeof(wchar_t));
    if (targetLength == 0 || targetLength >= MAX_PATH ||
        !BuildUpdateTempPath(expectedStagedPath, L"download", oldProcessId) ||
        !BuildUpdateTempPath(expectedHelperPath, L"updater", oldProcessId) ||
        _wcsicmp(stagedPath, expectedStagedPath) != 0 ||
        _wcsicmp(helperPath, expectedHelperPath) != 0) {
        return FALSE;
    }

    HANDLE helperProcess = OpenProcess(SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION,
                                       FALSE, helperProcessId);
    if (helperProcess) {
        wchar_t runningHelperPath[MAX_PATH];
        DWORD runningHelperPathLength = MAX_PATH;
        if (QueryFullProcessImageNameW(helperProcess, 0, runningHelperPath,
                                      &runningHelperPathLength) &&
            _wcsicmp(runningHelperPath, helperPath) == 0) {
            WaitForSingleObject(helperProcess, UPDATE_HELPER_WAIT_MS);
        }
        CloseHandle(helperProcess);
    }
    for (int attempt = 0;
         attempt < 20 && !DeleteUpdateTempFile(stagedPath);
         ++attempt) {
        Sleep(100);
    }
    for (int attempt = 0;
         attempt < 20 && !DeleteUpdateTempFile(helperPath);
         ++attempt) {
        Sleep(100);
    }
    return TRUE;
}

// Returns an exit code and sets handled for the temporary updater process.
// Both finish modes perform cleanup and continue normal application startup;
// updateCompleted identifies only a successful executable replacement.
static int HandleUpdateCommandLine(BOOL* handled, BOOL* updateCompleted,
                                    BOOL* reopenSettings) {
    if (handled) *handled = FALSE;
    if (updateCompleted) *updateCompleted = FALSE;
    if (reopenSettings) *reopenSettings = FALSE;
    int argumentCount = 0;
    LPWSTR* arguments = CommandLineToArgvW(GetCommandLineW(), &argumentCount);
    if (!arguments) return 0;

    // The optional choice belongs only to a validated updater handoff. Never
    // persist it, and accept legacy handoffs with the default (closed).
    BOOL wantsSettings = argumentCount == 7 &&
        wcscmp(arguments[6], L"--reopen-settings") == 0;
    BOOL validArguments = argumentCount == 6 || wantsSettings;
    int result = 0;
    if (validArguments && wcscmp(arguments[1], L"--apply-update") == 0) {
        DWORD oldProcessId = 0;
        if (handled) *handled = TRUE;
        if (!ParseUpdateProcessId(arguments[2], &oldProcessId)) {
            result = ERROR_INVALID_PARAMETER;
        } else {
            result = RunUpdateApplyHelper(oldProcessId, arguments[3],
                                          arguments[4], arguments[5], wantsSettings);
        }
    } else if (validArguments &&
               (wcscmp(arguments[1], L"--finish-update") == 0 ||
                wcscmp(arguments[1], L"--finish-update-cleanup") == 0)) {
        DWORD helperProcessId = 0, oldProcessId = 0;
        if (ParseUpdateProcessId(arguments[2], &helperProcessId) &&
            ParseUpdateProcessId(arguments[3], &oldProcessId)) {
            BOOL recognizedHandoff = FinishUpdateCleanup(
                helperProcessId, oldProcessId, arguments[4], arguments[5]);
            if (recognizedHandoff && updateCompleted &&
                wcscmp(arguments[1], L"--finish-update") == 0) {
                *updateCompleted = TRUE;
                if (reopenSettings) *reopenSettings = wantsSettings;
            }
        }
    }
    LocalFree(arguments);
    return result;
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

static void CfgSendUpdateResultWithVersions(LPCWSTR status, LPCWSTR title,
                                            LPCWSTR message,
                                            LPCWSTR currentVersion,
                                            LPCWSTR remoteVersion,
                                            BOOL automatic) {
    if (!g_cfgWebView || !status || !title || !message ||
        !currentVersion || !remoteVersion) return;
    wchar_t escapedStatus[64], escapedTitle[256], escapedMessage[1024];
    wchar_t escapedCurrentVersion[64], escapedRemoteVersion[64];
    json_escape_wstring(status, escapedStatus,
                        sizeof(escapedStatus) / sizeof(wchar_t));
    json_escape_wstring(title, escapedTitle,
                        sizeof(escapedTitle) / sizeof(wchar_t));
    json_escape_wstring(message, escapedMessage,
                        sizeof(escapedMessage) / sizeof(wchar_t));
    json_escape_wstring(currentVersion, escapedCurrentVersion,
                        sizeof(escapedCurrentVersion) / sizeof(wchar_t));
    json_escape_wstring(remoteVersion, escapedRemoteVersion,
                        sizeof(escapedRemoteVersion) / sizeof(wchar_t));

    wchar_t script[1792];
    int written = swprintf_s(script, sizeof(script) / sizeof(wchar_t),
        L"window.onUpdateResult({\"status\":\"%s\",\"title\":\"%s\","
        L"\"message\":\"%s\",\"currentVersion\":\"%s\","
        L"\"remoteVersion\":\"%s\",\"automatic\":%s})",
        escapedStatus, escapedTitle, escapedMessage,
        escapedCurrentVersion, escapedRemoteVersion,
        automatic ? L"true" : L"false");
    if (written > 0) webview_cfg_execute_script(script);
}

static void CfgSendUpdateResult(LPCWSTR status, LPCWSTR title, LPCWSTR message) {
    CfgSendUpdateResultWithVersions(status, title, message, L"", L"", FALSE);
}

static void CfgSendUpdateProgress(DWORD speedKbps) {
    wchar_t script[160];
    int written = swprintf_s(script, sizeof(script) / sizeof(wchar_t),
        L"window.onUpdateProgress({\"kilobytesPerSecond\":%lu})",
        (unsigned long)speedKbps);
    if (written > 0) webview_cfg_execute_script(script);
}

static void DiscardPendingUpdateNotice(void) {
    UpdateCheckTask* task = g_updateNoticeTask;
    g_updateNoticeTask = NULL;
    DiscardUpdateTask(task);
}

static void StartUpdateCheck(BOOL automatic) {
    HWND targetWindow = g_hwnd ? g_hwnd : g_cfgHwnd;
    if (!targetWindow) return;
    if (automatic && (g_updateNoticeTask || g_updateReadyTask)) return;
    if (InterlockedCompareExchangePointer(
            (PVOID volatile*)&g_updatePostedResult, NULL, NULL) != NULL) {
        if (!automatic) {
            CfgSendUpdateResult(L"error", L"Update check in progress",
                L"Another update check is still finishing. Try again shortly.");
        }
        return;
    }
    if (InterlockedCompareExchange(&g_updateCheckPending, TRUE, FALSE) != FALSE) {
        if (!automatic) {
            CfgSendUpdateResult(L"error", L"Update check in progress",
                L"Another update check is still finishing. Try again shortly.");
        }
        return;
    }
    InterlockedExchange(&g_updateCheckAutomatic, automatic ? TRUE : FALSE);

    // Every accepted request starts from scratch. Automatic requests are
    // skipped above while a result is awaiting user action, avoiding an
    // hourly re-download of the same prepared executable.
    DiscardPendingUpdateNotice();
    DiscardPreparedUpdate();

    if (!g_updateCancelEvent) {
        g_updateCancelEvent = CreateEventW(NULL, TRUE, FALSE, NULL);
        if (!g_updateCancelEvent) {
            DWORD errorCode = GetLastError();
            InterlockedExchange(&g_updateCheckPending, FALSE);
            InterlockedExchange(&g_updateCheckAutomatic, FALSE);
            wchar_t message[256];
            swprintf_s(message, sizeof(message) / sizeof(wchar_t),
                L"Could not initialize update cancellation (Windows error %lu).",
                (unsigned long)errorCode);
            CfgSendUpdateResultWithVersions(L"error", L"Update failed", message,
                                            L"", L"", automatic);
            return;
        }
    }
    ResetEvent(g_updateCancelEvent);
    InterlockedExchange(&g_updateSpeedKbps, 0);
    InterlockedExchange(&g_updateProgressPosted, FALSE);

    UpdateCheckTask* task = (UpdateCheckTask*)calloc(1, sizeof(UpdateCheckTask));
    if (!task) {
        InterlockedExchange(&g_updateCheckPending, FALSE);
        InterlockedExchange(&g_updateCheckAutomatic, FALSE);
        CfgSendUpdateResultWithVersions(L"error", L"Update failed",
            L"There was not enough memory to check for updates.",
            L"", L"", automatic);
        return;
    }
    task->targetWindow = targetWindow;
    task->automatic = automatic;
    LONG sequence = InterlockedIncrement(&g_updateRequestSequence);
    task->cacheBuster =
        ((GetTickCount64() ^ GetCurrentProcessId()) << 32) | (DWORD)sequence;
    if (task->cacheBuster == 0) task->cacheBuster = 1;

    HANDLE thread = CreateThread(NULL, 0, UpdateCheckThread, task, 0, NULL);
    if (!thread) {
        DWORD errorCode = GetLastError();
        free(task);
        InterlockedExchange(&g_updateCheckPending, FALSE);
        InterlockedExchange(&g_updateCheckAutomatic, FALSE);
        wchar_t message[256];
        swprintf_s(message, sizeof(message) / sizeof(wchar_t),
                   L"Could not start the update check (Windows error %lu).",
                   (unsigned long)errorCode);
        CfgSendUpdateResultWithVersions(L"error", L"Update failed", message,
                                        L"", L"", automatic);
        return;
    }
    CloseHandle(thread);
}

static BOOL IsIgnoredUpdateVersion(const ExecutableVersion* version) {
    wchar_t formatted[32];
    if (!version || !g_ignoredUpdateVersion[0]) return FALSE;
    FormatExecutableVersion(version, formatted,
                            sizeof(formatted) / sizeof(wchar_t));
    return wcscmp(formatted, g_ignoredUpdateVersion) == 0;
}

static void SaveIgnoredUpdateVersion(LPCWSTR version) {
    HKEY key;
    DWORD disposition;
    if (!version) return;
    wcsncpy_s(g_ignoredUpdateVersion,
              sizeof(g_ignoredUpdateVersion) / sizeof(wchar_t), version,
              _TRUNCATE);
    if (RegCreateKeyExW(HKEY_CURRENT_USER, REG_KEY_PATH, 0, NULL,
                        REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &key,
                        &disposition) == ERROR_SUCCESS) {
        RegSetValueExW(key, REG_VALUE_IGNORED_UPDATE_VERSION, 0, REG_SZ,
                       (const BYTE*)g_ignoredUpdateVersion,
                       (DWORD)((wcslen(g_ignoredUpdateVersion) + 1) *
                               sizeof(wchar_t)));
        RegCloseKey(key);
    }
}

static void PresentPendingUpdateNotice(void) {
    if (!g_configViewReady || !g_cfgWebView || !g_updateNoticeTask) return;

    UpdateCheckTask* task = g_updateNoticeTask;
    g_updateNoticeTask = NULL;
    LPCWSTR status = NULL;
    LPCWSTR title = NULL;
    LPCWSTR message = NULL;
    wchar_t currentVersion[32] = L"";
    wchar_t remoteVersion[32] = L"";
    BOOL installable = FALSE;

    if (task->kind == UPDATE_CHECK_CANCELLED) {
        status = L"cancelled";
        title = L"";
        message = L"";
    } else if (task->kind == UPDATE_CHECK_ERROR) {
        status = L"error";
        title = L"Update failed";
        message = task->message;
    } else {
        FormatExecutableVersionForDisplay(
            &task->runningVersion, currentVersion,
            sizeof(currentVersion) / sizeof(wchar_t));
        FormatExecutableVersionForDisplay(
            &task->availableVersion, remoteVersion,
            sizeof(remoteVersion) / sizeof(wchar_t));
        if (task->kind == UPDATE_CHECK_NEWER) {
            status = L"newer";
            title = L"Update available";
            message = L"A newer version is ready to install.";
            installable = TRUE;
        } else if (task->kind == UPDATE_CHECK_SAME) {
            status = L"same";
            title = L"You're up to date";
            message = L"The remote build matches your current version. "
                      L"You can force a reinstall if needed.";
            installable = !task->automatic;
        } else if (task->kind == UPDATE_CHECK_OLDER) {
            status = L"older";
            title = L"No update available";
            message = L"The remote build is older than your current version.";
        }
    }

    if (!status) {
        DiscardUpdateTask(task);
        return;
    }
    if (installable) {
        DiscardPreparedUpdate();
        g_updateReadyTask = task;
    }
    CfgSendUpdateResultWithVersions(status, title, message,
                                    currentVersion, remoteVersion,
                                    task->automatic);
    if (!installable) DiscardUpdateTask(task);
}

static void QueueUpdateNotice(UpdateCheckTask* task) {
    DiscardPendingUpdateNotice();
    g_updateNoticeTask = task;
    PresentPendingUpdateNotice();
}

static void HandleCompletedUpdateCheck(UpdateCheckTask* task) {
    if (!task) return;
    InterlockedExchange(&g_updateProgressPosted, FALSE);
    InterlockedExchange(&g_updateSpeedKbps, 0);

    if (task->kind == UPDATE_CHECK_CANCELLED) {
        DebugPrint(L"[INFO] Update check cancelled\n");
    } else if (task->kind == UPDATE_CHECK_ERROR) {
        DebugPrint(L"[WARNING] Update check failed: %s\n", task->message);
    }

    if (task->automatic && !g_config.autoCheckForUpdates) {
        DiscardUpdateTask(task);
        return;
    }

    if (task->automatic && task->kind == UPDATE_CHECK_NEWER &&
        IsIgnoredUpdateVersion(&task->availableVersion)) {
        DebugPrint(L"[INFO] Automatic update prompt suppressed for ignored version\n");
        task->kind = UPDATE_CHECK_CANCELLED;
        if (g_cfgHwnd) {
            QueueUpdateNotice(task);
        } else {
            DiscardUpdateTask(task);
        }
        return;
    }

    if (task->automatic && task->kind == UPDATE_CHECK_NEWER) {
        QueueUpdateNotice(task);
        ShowConfigWebViewDialog();
        if (!g_cfgHwnd) DiscardPendingUpdateNotice();
        return;
    }

    if (g_cfgHwnd) {
        QueueUpdateNotice(task);
    } else {
        DiscardUpdateTask(task);
    }
}

static void IgnorePreparedUpdateVersion(const char* requestedVersion) {
    UpdateCheckTask* task = g_updateReadyTask;
    wchar_t preparedVersion[32];
    wchar_t requestedVersionW[32] = L"";
    if (!task || !task->automatic || task->kind != UPDATE_CHECK_NEWER ||
        !requestedVersion ||
        !MultiByteToWideChar(CP_UTF8, 0, requestedVersion, -1,
                             requestedVersionW,
                             sizeof(requestedVersionW) / sizeof(wchar_t))) {
        CfgSendUpdateResult(L"error", L"Update unavailable",
            L"The update version could not be ignored. Check for updates again.");
        return;
    }
    FormatExecutableVersion(&task->availableVersion, preparedVersion,
                            sizeof(preparedVersion) / sizeof(wchar_t));
    if (wcscmp(preparedVersion, requestedVersionW) != 0) {
        CfgSendUpdateResult(L"error", L"Update unavailable",
            L"The update version changed. Check for updates again.");
        return;
    }
    SaveIgnoredUpdateVersion(preparedVersion);
    DebugPrint(L"[INFO] Automatic update version added to the ignore list\n");
    DiscardPreparedUpdate();
}

static void CancelUpdateCheck(void) {
    if (InterlockedCompareExchange(&g_updateCheckPending, FALSE, FALSE) == TRUE &&
        g_updateCancelEvent) {
        DebugPrint(L"[INFO] Update check cancellation requested\n");
        SetEvent(g_updateCancelEvent);
    }
}

static HANDLE CreateUpdateReadyEvent(DWORD processId, wchar_t* eventName,
                                     size_t eventNameCch) {
    if (!processId || !eventName || eventNameCch < 96) {
        SetLastError(ERROR_INVALID_PARAMETER);
        return NULL;
    }

    ULONGLONG nonce = GetTickCount64() ^
                      ((ULONGLONG)GetCurrentThreadId() << 32);
    BCryptGenRandom(NULL, (PUCHAR)&nonce, sizeof(nonce),
                    BCRYPT_USE_SYSTEM_PREFERRED_RNG);
    int nameLength = swprintf_s(eventName, eventNameCch,
        L"Local\\SystrayLauncher_UpdateReady_%lu_%016llx",
        (unsigned long)processId, (unsigned long long)nonce);
    if (nameLength <= 0 || nameLength >= (int)eventNameCch) {
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return NULL;
    }

    // The elevated helper can run with a split administrator token (or with
    // alternate administrator credentials). Grant interactive users access
    // to this random, session-local event so either UAC path can acknowledge
    // readiness without exposing any file or process permissions.
    PSECURITY_DESCRIPTOR descriptor = NULL;
    if (!ConvertStringSecurityDescriptorToSecurityDescriptorW(
            L"D:P(A;;GA;;;IU)(A;;GA;;;BA)(A;;GA;;;SY)",
            SDDL_REVISION_1, &descriptor, NULL)) {
        return NULL;
    }
    SECURITY_ATTRIBUTES securityAttributes = {0};
    securityAttributes.nLength = sizeof(securityAttributes);
    securityAttributes.lpSecurityDescriptor = descriptor;

    HANDLE readyEvent = CreateEventW(&securityAttributes, TRUE, FALSE,
                                     eventName);
    DWORD errorCode = readyEvent ? ERROR_SUCCESS : GetLastError();
    LocalFree(descriptor);
    if (!readyEvent) SetLastError(errorCode);
    return readyEvent;
}

static BOOL LaunchStagedUpdate(LPCWSTR stagedPath, LPCWSTR targetPath,
                               BOOL reopenSettings) {
    if (!stagedPath || !targetPath || !*stagedPath || !*targetPath) {
        SetLastError(ERROR_INVALID_PARAMETER);
        return FALSE;
    }

    DWORD oldProcessId = GetCurrentProcessId();
    wchar_t helperPath[MAX_PATH];
    if (!BuildUpdateTempPath(helperPath, L"updater", oldProcessId)) {
        return FALSE;
    }
    DeleteUpdateTempFile(helperPath);
    if (!CopyFileW(targetPath, helperPath, TRUE)) return FALSE;
    SetFileAttributesW(helperPath, FILE_ATTRIBUTE_NORMAL);

    // CopyFile preserves alternate data streams. The source is already the
    // running, user-approved executable, so do not carry its download-zone
    // marker onto the short-lived updater copy and trigger a second warning.
    wchar_t zonePath[MAX_PATH + 32];
    if (swprintf_s(zonePath, sizeof(zonePath) / sizeof(wchar_t),
                   L"%s:Zone.Identifier", helperPath) > 0) {
        DeleteFileW(zonePath);
    }

    wchar_t readyEventName[160];
    HANDLE readyEvent = CreateUpdateReadyEvent(oldProcessId, readyEventName,
        sizeof(readyEventName) / sizeof(wchar_t));
    if (!readyEvent) {
        DWORD errorCode = GetLastError();
        DeleteUpdateTempFile(helperPath);
        SetLastError(errorCode);
        return FALSE;
    }

    wchar_t parameters[MAX_PATH * 2 + 512];
    int parameterLength = swprintf_s(parameters,
        sizeof(parameters) / sizeof(wchar_t),
        L"--apply-update %lu \"%s\" \"%s\" \"%s\"%s",
        (unsigned long)oldProcessId, readyEventName, targetPath, stagedPath,
        reopenSettings ? L" --reopen-settings" : L"");
    if (parameterLength <= 0 ||
        parameterLength >= (int)(sizeof(parameters) / sizeof(wchar_t))) {
        CloseHandle(readyEvent);
        DeleteUpdateTempFile(helperPath);
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return FALSE;
    }

    SHELLEXECUTEINFOW executeInfo = {0};
    executeInfo.cbSize = sizeof(executeInfo);
    executeInfo.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_NOASYNC;
    executeInfo.hwnd = g_cfgHwnd;
    executeInfo.lpVerb = L"runas";
    executeInfo.lpFile = helperPath;
    executeInfo.lpParameters = parameters;
    executeInfo.nShow = SW_HIDE;
    BOOL elevated = ShellExecuteExW(&executeInfo);
    if (!elevated || !executeInfo.hProcess) {
        DWORD errorCode = elevated ? ERROR_INVALID_HANDLE : GetLastError();
        if (!errorCode) errorCode = ERROR_ACCESS_DENIED;
        CloseHandle(readyEvent);
        DeleteUpdateTempFile(helperPath);
        SetLastError(errorCode);
        return FALSE;
    }

    HANDLE waitHandles[2] = { readyEvent, executeInfo.hProcess };
    DWORD waitResult = WaitForMultipleObjects(2, waitHandles, FALSE,
                                              UPDATE_HELPER_READY_MS);
    DWORD errorCode = ERROR_SUCCESS;
    if (waitResult != WAIT_OBJECT_0) {
        if (waitResult == WAIT_OBJECT_0 + 1) {
            DWORD exitCode = ERROR_INSTALL_FAILURE;
            if (!GetExitCodeProcess(executeInfo.hProcess, &exitCode) ||
                exitCode == ERROR_SUCCESS || exitCode == STILL_ACTIVE) {
                exitCode = ERROR_INSTALL_FAILURE;
            }
            errorCode = exitCode;
        } else {
            errorCode = waitResult == WAIT_TIMEOUT ? ERROR_TIMEOUT
                                                   : GetLastError();
            if (!errorCode) errorCode = ERROR_INSTALL_FAILURE;
        }
    }
    CloseHandle(executeInfo.hProcess);
    CloseHandle(readyEvent);

    if (waitResult != WAIT_OBJECT_0) {
        DeleteUpdateTempFile(helperPath);
        SetLastError(errorCode);
        return FALSE;
    }
    return TRUE;
}

static void DiscardPreparedUpdate(void) {
    UpdateCheckTask* task = g_updateReadyTask;
    g_updateReadyTask = NULL;
    DiscardUpdateTask(task);
}

static void InstallPreparedUpdate(BOOL reopenSettings) {
    UpdateCheckTask* task = g_updateReadyTask;
    g_updateReadyTask = NULL;
    if (!task || (task->kind != UPDATE_CHECK_NEWER &&
                  task->kind != UPDATE_CHECK_SAME)) {
        DiscardUpdateTask(task);
        CfgSendUpdateResult(L"error", L"Update unavailable",
            L"The prepared update is no longer available. Check for updates again.");
        return;
    }

    if (LaunchStagedUpdate(task->stagedPath, task->targetPath, reopenSettings)) {
        DebugPrint(L"[INFO] Update accepted; exiting for replacement\n");
        g_updateInstallReady = TRUE;
        free(task);  // The updater process now owns the staged file.
        if (g_cfgHwnd) PostMessageW(g_cfgHwnd, WM_CLOSE, 0, 0);
        return;
    }

    DWORD errorCode = GetLastError();
    wchar_t message[384];
    LPCWSTR title = L"Update failed";
    if (errorCode == ERROR_CANCELLED) {
        title = L"Update cancelled";
        wcscpy_s(message, sizeof(message) / sizeof(wchar_t),
            L"Administrator approval was cancelled. Your current version is still running.");
    } else {
        swprintf_s(message, sizeof(message) / sizeof(wchar_t),
            L"The elevated update process could not be started (Windows error %lu).",
            (unsigned long)errorCode);
    }
    DebugPrint(L"[WARNING] %s\n", message);
    CfgSendUpdateResult(L"error", title, message);
    DiscardUpdateTask(task);
}

static void cfg_sync_controller_bounds(void) {
    if (!g_cfgController || !g_cfgHwnd) return;
    RECT bounds;
    GetClientRect(g_cfgHwnd, &bounds);
    g_cfgController->lpVtbl->put_Bounds(g_cfgController, bounds);
    g_cfgController->lpVtbl->put_IsVisible(g_cfgController, TRUE);
}

static void webview_push_init_config(void) {
    wchar_t eUrl[4096], eTitle[512], eMailtoTargetUrl[4096];
    wchar_t eHide[8192], eShow[8192], eInsecureOrigins[4096];
    wchar_t eStaticHosts[4096], eLockdownSecret[512], eWebView2Version[256];
    wchar_t eUpdateCompletedVersion[64];
    json_escape_wstring(g_config.url, eUrl, 4096);
    json_escape_wstring(g_config.windowTitle, eTitle, 512);
    json_escape_wstring(g_config.mailtoTargetUrl, eMailtoTargetUrl, 4096);
    json_escape_wstring(g_config.onHideJs, eHide, 8192);
    json_escape_wstring(g_config.onShowJs, eShow, 8192);
    json_escape_wstring(g_config.insecureContentOrigins, eInsecureOrigins, 4096);
    json_escape_wstring(g_config.staticHostMappings, eStaticHosts, 4096);
    json_escape_wstring(g_config.lockdownSecret, eLockdownSecret, 512);
    json_escape_wstring(g_webView2Version, eWebView2Version, 256);
    json_escape_wstring(g_updateConfirmationPending ? APP_VERSION_WSTRING : L"",
                        eUpdateCompletedVersion, 64);

    // Sized for every field at maximum, fully escaped, plus the JSON scaffold.
    // A fixed 16K buffer used to silently truncate the script for large JS
    // hooks, which broke onInit and left the dialog blank.
    const size_t scriptCch =
        4096 + 512 + 4096 + 8192 + 8192 + 4096 + 4096 + 512 + 256 + 64 + 1152;
    wchar_t* script = (wchar_t*)malloc(scriptCch * sizeof(wchar_t));
    if (!script) return;
    int written = swprintf(script, scriptCch,
        L"window.onInit({\"config\":{\"url\":\"%s\",\"windowTitle\":\"%s\",\"startMaximized\":%s,\"returnToTargetOnDoubleClick\":%s,\"showInTaskbar\":%s,\"handleMailtoLinks\":%s,\"mailtoTargetUrl\":\"%s\",\"onHideJs\":\"%s\",\"onShowJs\":\"%s\",\"sleepWhenInactive\":%s,\"openNewWindowsExternally\":%s,\"allowRunningInsecureContent\":%s,\"insecureContentOrigins\":\"%s\",\"useStaticHostMappings\":%s,\"staticHostMappings\":\"%s\",\"staticHostDnsFallback\":%s,\"lockdownHeader\":%s,\"lockdownSecret\":\"%s\",\"autoCheckForUpdates\":%s,\"updateCheckPending\":%s,\"updatePromptPending\":%s,\"debugLog\":%s},\"webView2Version\":\"%s\",\"updateCompletedVersion\":\"%s\"})",
        eUrl, eTitle, g_config.startMaximized ? L"true" : L"false",
        g_config.returnToTargetOnDoubleClick ? L"true" : L"false",
        g_config.showInTaskbar ? L"true" : L"false",
        g_config.handleMailtoLinks ? L"true" : L"false",
        eMailtoTargetUrl,
        eHide, eShow, g_config.sleepWhenInactive ? L"true" : L"false",
        g_config.openNewWindowsExternally ? L"true" : L"false",
        g_config.allowRunningInsecureContent ? L"true" : L"false",
        eInsecureOrigins,
        g_config.useStaticHostMappings ? L"true" : L"false",
        eStaticHosts,
        g_config.staticHostDnsFallback ? L"true" : L"false",
        g_config.lockdownHeader ? L"true" : L"false",
        eLockdownSecret,
        g_config.autoCheckForUpdates ? L"true" : L"false",
        (InterlockedCompareExchange(&g_updateCheckPending,
                                    FALSE, FALSE) == TRUE ||
         InterlockedCompareExchangePointer(
             (PVOID volatile*)&g_updatePostedResult, NULL, NULL) != NULL)
            ? L"true" : L"false",
        g_updateNoticeTask ? L"true" : L"false",
        g_config.debugLogEnabled ? L"true" : L"false",
        eWebView2Version,
        eUpdateCompletedVersion);
    if (written > 0) {
        webview_cfg_execute_script(script);
    }
    free(script);
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
    } else if (strcmp(action, "configReady") == 0) {
        if (g_cfgHwnd) {
            BOOL updateWorkAlreadyActive =
                InterlockedCompareExchange(&g_updateCheckPending,
                                           FALSE, FALSE) == TRUE ||
                InterlockedCompareExchangePointer(
                    (PVOID volatile*)&g_updatePostedResult, NULL, NULL) != NULL ||
                g_updateNoticeTask || g_updateReadyTask;
            BOOL checkAutomatically =
                json_get_bool(msg, "checkAutomatically", FALSE);
            g_configViewReady = TRUE;
            PresentPendingUpdateNotice();
            if (checkAutomatically && g_config.autoCheckForUpdates &&
                !g_updateConfirmationPending && !updateWorkAlreadyActive) {
                StartUpdateCheck(TRUE);
            }
        }
    } else if (strcmp(action, "checkUpdate") == 0) {
        StartUpdateCheck(json_get_bool(msg, "automatic", FALSE));
    } else if (strcmp(action, "cancelUpdateCheck") == 0) {
        CancelUpdateCheck();
    } else if (strcmp(action, "installUpdate") == 0) {
        InstallPreparedUpdate(json_get_bool(msg, "reopenSettings", FALSE));
    } else if (strcmp(action, "dismissUpdate") == 0) {
        DiscardPreparedUpdate();
    } else if (strcmp(action, "ignoreUpdateVersion") == 0) {
        char version[32] = {0};
        json_get_string(msg, "version", version, sizeof(version));
        IgnorePreparedUpdateVersion(version);
    } else if (strcmp(action, "dismissUpdateConfirmation") == 0) {
        g_updateConfirmationPending = FALSE;
    } else if (strcmp(action, "saveSettings") == 0) {
        char url[4096] = {0}, title[512] = {0}, hideJs[8192] = {0}, showJs[8192] = {0};
        char mailtoTargetUrl[8192] = {0};
        char insecureOrigins[8192] = {0}, staticHosts[8192] = {0};
        json_get_string(msg, "url", url, sizeof(url));
        json_get_string(msg, "windowTitle", title, sizeof(title));
        json_get_string(msg, "mailtoTargetUrl", mailtoTargetUrl,
                        sizeof(mailtoTargetUrl));
        json_get_string(msg, "onHideJs", hideJs, sizeof(hideJs));
        json_get_string(msg, "onShowJs", showJs, sizeof(showJs));
        json_get_string(msg, "insecureContentOrigins", insecureOrigins,
                        sizeof(insecureOrigins));
        json_get_string(msg, "staticHostMappings", staticHosts,
                        sizeof(staticHosts));

        Configuration previousConfig = g_config;
        BOOL oldShowInTaskbar = previousConfig.showInTaskbar;
        BOOL oldHandleMailtoLinks = previousConfig.handleMailtoLinks;
        MultiByteToWideChar(CP_UTF8, 0, url, -1, g_config.url, 2048);
        MultiByteToWideChar(CP_UTF8, 0, title, -1, g_config.windowTitle, 256);
        if (MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS,
                                mailtoTargetUrl, -1,
                                g_config.mailtoTargetUrl, 2048) == 0) {
            g_config.mailtoTargetUrl[0] = L'\0';
        }
        MultiByteToWideChar(CP_UTF8, 0, hideJs, -1, g_config.onHideJs, 4096);
        MultiByteToWideChar(CP_UTF8, 0, showJs, -1, g_config.onShowJs, 4096);
        g_config.startMaximized = json_get_bool(msg, "startMaximized", FALSE);
        g_config.returnToTargetOnDoubleClick =
            json_get_bool(msg, "returnToTargetOnDoubleClick", TRUE);
        g_config.showInTaskbar = json_get_bool(msg, "showInTaskbar", FALSE);
        g_config.handleMailtoLinks =
            json_get_bool(msg, "handleMailtoLinks", FALSE);
        g_config.sleepWhenInactive = json_get_bool(msg, "sleepWhenInactive", FALSE);
        g_config.openNewWindowsExternally = json_get_bool(msg, "openNewWindowsExternally", FALSE);
        BOOL oldAllowRunningInsecureContent = g_config.allowRunningInsecureContent;
        wchar_t oldInsecureContentOrigins[2048];
        wcscpy_s(oldInsecureContentOrigins, 2048, g_config.insecureContentOrigins);
        g_config.allowRunningInsecureContent =
            json_get_bool(msg, "allowRunningInsecureContent", FALSE);
        if (MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, insecureOrigins, -1,
                                g_config.insecureContentOrigins, 2048) == 0) {
            g_config.insecureContentOrigins[0] = L'\0';
        }
        BOOL oldUseStaticHostMappings = g_config.useStaticHostMappings;
        BOOL oldStaticHostDnsFallback = g_config.staticHostDnsFallback;
        wchar_t oldStaticHostMappings[2048];
        wcscpy_s(oldStaticHostMappings, 2048, g_config.staticHostMappings);
        g_config.useStaticHostMappings =
            json_get_bool(msg, "useStaticHostMappings", FALSE);
        if (MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, staticHosts, -1,
                                g_config.staticHostMappings, 2048) == 0) {
            g_config.staticHostMappings[0] = L'\0';
        }
        g_config.staticHostDnsFallback =
            json_get_bool(msg, "staticHostDnsFallback", FALSE);
        char lockdownSecret[1024] = {0};
        json_get_string(msg, "lockdownSecret", lockdownSecret, sizeof(lockdownSecret));
        g_config.lockdownHeader = json_get_bool(msg, "lockdownHeader", FALSE);
        if (MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, lockdownSecret, -1,
                                g_config.lockdownSecret, 256) == 0) {
            g_config.lockdownSecret[0] = L'\0';
        }
        g_config.autoCheckForUpdates =
            json_get_bool(msg, "autoCheckForUpdates", TRUE);
        g_config.debugLogEnabled = json_get_bool(msg, "debugLog", FALSE);

        if (g_config.handleMailtoLinks &&
            !IsValidHttpNavigationUrl(g_config.mailtoTargetUrl)) {
            g_config = previousConfig;
            MessageBoxW(g_cfgHwnd,
                        L"Enter a valid http:// or https:// destination URL "
                        L"before enabling email-link handling.",
                        APP_NAME, MB_OK | MB_ICONWARNING);
            free(msg);
            return S_OK;
        }

        if (!SaveConfigToRegistry(&g_config)) {
            g_config = previousConfig;
            MessageBoxW(g_cfgHwnd,
                        L"The settings could not be saved. Please try again.",
                        APP_NAME, MB_OK | MB_ICONERROR);
            free(msg);
            return S_OK;
        }
        BOOL mailtoRegistrationUpdated =
            SetMailtoHandlerRegistration(g_config.handleMailtoLinks);
        MarkAsConfigured();
        ApplyConfiguration();

        g_cfgSaved = TRUE;
        PostMessage(g_cfgHwnd, WM_CLOSE, 0, 0);

        if (!mailtoRegistrationUpdated) {
            MessageBoxW(
                g_cfgHwnd,
                g_config.handleMailtoLinks
                    ? L"The settings were saved, but SystrayLauncher could not "
                      L"be registered with Windows as an email-link handler. "
                      L"It will try again the next time it starts."
                    : L"The settings were saved, but SystrayLauncher's email-link "
                      L"registration could not be removed. Please try saving again.",
                APP_NAME, MB_OK | MB_ICONWARNING);
        } else if (!oldHandleMailtoLinks && g_config.handleMailtoLinks) {
            MailtoDefaultAppsOpenResult settingsResult =
                OpenMailtoDefaultAppsSettings();
            if (settingsResult != MAILTO_DEFAULT_APPS_APP_PAGE) {
                MessageBoxW(
                    g_cfgHwnd,
                    settingsResult == MAILTO_DEFAULT_APPS_GENERAL_PAGE
                        ? L"Windows Default Apps is open. Select "
                          L"System Tray Launcher and assign it to MAILTO links to "
                          L"finish enabling email-link handling."
                        : L"SystrayLauncher is registered for email links, but "
                          L"Windows Settings could not be opened. Open Settings "
                          L"> Apps > Default apps, select System Tray Launcher, and "
                          L"assign it to MAILTO links.",
                    APP_NAME, MB_OK | MB_ICONINFORMATION);
            }
        }

        if (g_hwnd &&
            (oldAllowRunningInsecureContent != g_config.allowRunningInsecureContent ||
             wcscmp(oldInsecureContentOrigins, g_config.insecureContentOrigins) != 0 ||
             oldUseStaticHostMappings != g_config.useStaticHostMappings ||
             oldStaticHostDnsFallback != g_config.staticHostDnsFallback ||
             wcscmp(oldStaticHostMappings, g_config.staticHostMappings) != 0 ||
             oldShowInTaskbar != g_config.showInTaskbar)) {
            // Browser arguments and main-window ownership are established at
            // startup. Close the dialog first, then restart so the new process
            // creates them with the updated settings.
            DebugPrint(L"[INFO] Startup-only setting changed; restarting launcher\n");
            PostMessage(g_hwnd, WM_COMMAND, ID_TRAY_MENU_RESTART, 0);
        }
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
        // Optional desired width: sent when the page reflows the settings
        // into two columns; absent (0) keeps the current window width.
        char wStr[32] = {0};
        json_get_string(msg, "width", wStr, sizeof(wStr));
        int contentWidth = atoi(wStr);
        if (contentWidth <= 0) {
            const char *wp = strstr(msg, "\"width\"");
            if (wp) {
                wp += 7;
                while (*wp == ' ' || *wp == ':') wp++;
                contentWidth = atoi(wp);
            }
        }
        // Content-driven sizing must not fight a maximized (or minimized)
        // window; WM_SIZE keeps the WebView bounds in sync there.
        if (contentHeight > 0 && g_cfgHwnd &&
            !IsZoomed(g_cfgHwnd) && !IsIconic(g_cfgHwnd)) {
            // The page reports its height in CSS pixels; convert to the
            // physical pixels window sizes use.
            int physHeight = MulDiv(contentHeight, (int)GetWindowDpi(g_cfgHwnd), 96);
            RECT clientRect = {0}, windowRect = {0};
            GetClientRect(g_cfgHwnd, &clientRect);
            GetWindowRect(g_cfgHwnd, &windowRect);
            int chromeH = (windowRect.bottom - windowRect.top) - (clientRect.bottom - clientRect.top);
            int newWindowH = physHeight + chromeH;
            int windowW = windowRect.right - windowRect.left;
            if (contentWidth > 0) {
                int chromeW = (windowRect.right - windowRect.left) -
                              (clientRect.right - clientRect.left);
                windowW = MulDiv(contentWidth, (int)GetWindowDpi(g_cfgHwnd), 96) +
                          chromeW;
            }

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
        case WM_APP_UPDATE_PROGRESS:
            InterlockedExchange(&g_updateProgressPosted, FALSE);
            if (InterlockedCompareExchange(&g_updateCheckPending,
                                           FALSE, FALSE) == TRUE) {
                DWORD speedKbps = (DWORD)InterlockedCompareExchange(
                    &g_updateSpeedKbps, 0, 0);
                CfgSendUpdateProgress(speedKbps);
            }
            return 0;

        case WM_APP_UPDATE_RESULT: {
            UpdateCheckTask* task = (UpdateCheckTask*)InterlockedExchangePointer(
                (PVOID volatile*)&g_updatePostedResult, NULL);
            HandleCompletedUpdateCheck(task);
            return 0;
        }

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
            if (InterlockedCompareExchange(&g_updateCheckPending,
                                           FALSE, FALSE) == TRUE &&
                (InterlockedCompareExchange(&g_updateCheckAutomatic,
                                            FALSE, FALSE) == FALSE ||
                 !g_hwnd) &&
                g_updateCancelEvent) {
                SetEvent(g_updateCancelEvent);
            }
            if (!g_hwnd) {
                DiscardUpdateTask((UpdateCheckTask*)InterlockedExchangePointer(
                    (PVOID volatile*)&g_updatePostedResult, NULL));
            }
            DiscardPendingUpdateNotice();
            DiscardPreparedUpdate();
            g_cfgHwnd = NULL;
            g_cfgWindowShown = FALSE;
            g_configViewReady = FALSE;
            KillTimer(hwnd, ID_TIMER_CFG_SHOW_FALLBACK);
            if (g_updateInstallReady) PostQuitMessage(0);
            return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static void ShowConfigWebViewDialog(void) {
    if (g_cfgHwnd != NULL) {
        if (IsIconic(g_cfgHwnd)) ShowWindow(g_cfgHwnd, SW_RESTORE);
        else ShowWindow(g_cfgHwnd, SW_SHOW);
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

    // Same frame styles as the main window so both get identical caption
    // rendering; the fixed dialog frame used before drew a more compact
    // title bar that looked out of place next to the main window.
    g_cfgHwnd = CreateWindowExW(0, L"SystrayLauncherCfgWnd", L"Configuration",
        WS_OVERLAPPEDWINDOW,
        posX, posY, width, height,
        NULL, NULL, g_hInstance, NULL);

    if (!g_cfgHwnd) return;
    g_cfgWindowShown = FALSE;
    g_configViewReady = FALSE;
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

static void GetMainUserDataFolder(wchar_t path[MAX_PATH]) {
    path[0] = L'\0';
    SHGetFolderPathW(NULL, CSIDL_LOCAL_APPDATA, NULL, 0, path);
    PathAppendW(path, APP_NAME L"\\WebView2Data");
}

static BOOL IsOriginListSeparator(wchar_t c) {
    return c == L',' || c == L';' || iswspace(c);
}

static BOOL IsValidOriginPort(const wchar_t* begin, const wchar_t* end) {
    if (begin >= end) return FALSE;

    unsigned long port = 0;
    for (const wchar_t* p = begin; p < end; ++p) {
        if (*p < L'0' || *p > L'9') return FALSE;
        port = port * 10 + (unsigned long)(*p - L'0');
        if (port > 65535) return FALSE;
    }
    return port > 0;
}

// AdditionalBrowserArguments is a command line, so only accept exact ASCII
// HTTP origins here. Besides matching Chromium's documented origin format,
// this prevents a registry or INI value from introducing another switch.
static BOOL IsValidInsecureContentOrigin(const wchar_t* origin) {
    static const wchar_t httpPrefix[] = L"http://";
    const size_t prefixLen = (sizeof(httpPrefix) / sizeof(httpPrefix[0])) - 1;
    size_t length = wcslen(origin);
    if (length <= prefixLen || _wcsnicmp(origin, httpPrefix, prefixLen) != 0) {
        return FALSE;
    }

    const wchar_t* authority = origin + prefixLen;
    const wchar_t* end = origin + length;

    if (*authority == L'[') {
        const wchar_t* closeBracket = wcschr(authority + 1, L']');
        if (!closeBracket || closeBracket == authority + 1) return FALSE;

        int colonCount = 0;
        for (const wchar_t* p = authority + 1; p < closeBracket; ++p) {
            wchar_t c = *p;
            if (c == L':') {
                colonCount++;
            } else if (!((c >= L'0' && c <= L'9') ||
                         (c >= L'a' && c <= L'f') ||
                         (c >= L'A' && c <= L'F') || c == L'.')) {
                return FALSE;
            }
        }
        if (colonCount < 2) return FALSE;

        if (closeBracket + 1 == end) return TRUE;
        return closeBracket + 1 < end && closeBracket[1] == L':' &&
               IsValidOriginPort(closeBracket + 2, end);
    }

    const wchar_t* portSeparator = NULL;
    for (const wchar_t* p = authority; p < end; ++p) {
        if (*p == L':') {
            if (portSeparator) return FALSE;
            portSeparator = p;
        }
    }

    const wchar_t* hostEnd = portSeparator ? portSeparator : end;
    if (authority == hostEnd) return FALSE;
    for (const wchar_t* p = authority; p < hostEnd; ++p) {
        wchar_t c = *p;
        if (!((c >= L'0' && c <= L'9') ||
              (c >= L'a' && c <= L'z') ||
              (c >= L'A' && c <= L'Z') ||
              c == L'.' || c == L'-' || c == L'_')) {
            return FALSE;
        }
    }

    return !portSeparator || IsValidOriginPort(portSeparator + 1, end);
}

static wchar_t* BuildInsecureContentBrowserArguments(size_t* originCount) {
    if (originCount) *originCount = 0;
    if (!g_config.allowRunningInsecureContent) return NULL;

    const wchar_t* configured = g_config.insecureContentOrigins;
    size_t configuredLength = wcslen(configured);
    wchar_t* origins = (wchar_t*)calloc(configuredLength + 1, sizeof(wchar_t));
    if (!origins) return NULL;

    size_t originsLength = 0;
    size_t count = 0;
    const wchar_t* cursor = configured;
    while (*cursor) {
        while (*cursor && IsOriginListSeparator(*cursor)) cursor++;
        if (!*cursor) break;

        const wchar_t* begin = cursor;
        while (*cursor && !IsOriginListSeparator(*cursor)) cursor++;
        size_t tokenLength = (size_t)(cursor - begin);
        wchar_t token[2048];
        if (tokenLength == 0 || tokenLength >= sizeof(token) / sizeof(token[0])) {
            free(origins);
            DebugPrint(L"[WARNING] Insecure-content origin list is too long or malformed\n");
            return NULL;
        }
        wmemcpy(token, begin, tokenLength);
        token[tokenLength] = L'\0';

        if (!IsValidInsecureContentOrigin(token)) {
            free(origins);
            DebugPrint(L"[WARNING] Insecure-content setting requires exact ASCII http:// origins\n");
            return NULL;
        }

        if (count > 0) origins[originsLength++] = L',';
        wmemcpy(origins + originsLength, token, tokenLength);
        originsLength += tokenLength;
        origins[originsLength] = L'\0';
        count++;
    }

    if (count == 0) {
        free(origins);
        DebugPrint(L"[WARNING] Insecure-content setting is enabled but no HTTP origins are configured\n");
        return NULL;
    }

    static const wchar_t argumentPrefix[] =
        L"--unsafely-treat-insecure-origin-as-secure=";
    size_t argumentLength = wcslen(argumentPrefix) + originsLength;
    wchar_t* arguments = (wchar_t*)malloc((argumentLength + 1) * sizeof(wchar_t));
    if (!arguments) {
        free(origins);
        return NULL;
    }
    wcscpy_s(arguments, argumentLength + 1, argumentPrefix);
    wcscat_s(arguments, argumentLength + 1, origins);
    free(origins);
    if (originCount) *originCount = count;
    return arguments;
}

static BOOL IsValidStaticHostName(const wchar_t* begin, const wchar_t* end) {
    size_t length = (size_t)(end - begin);
    if (length == 0 || length > 253) return FALSE;

    const wchar_t* labelStart = begin;
    for (const wchar_t* p = begin; p <= end; ++p) {
        if (p == end || *p == L'.') {
            size_t labelLength = (size_t)(p - labelStart);
            if (labelLength == 0 || labelLength > 63 ||
                *labelStart == L'-' || p[-1] == L'-') {
                return FALSE;
            }
            labelStart = p + 1;
            continue;
        }

        wchar_t c = *p;
        if (!((c >= L'0' && c <= L'9') ||
              (c >= L'a' && c <= L'z') ||
              (c >= L'A' && c <= L'Z') ||
              c == L'-' || c == L'_')) {
            return FALSE;
        }
    }
    return TRUE;
}

static BOOL IsValidStaticIpAddress(const wchar_t* begin, const wchar_t* end) {
    if (begin >= end) return FALSE;

    int addressFamily = AF_INET;
    if (*begin == L'[') {
        if (end - begin < 4 || end[-1] != L']') return FALSE;
        begin++;
        end--;
        addressFamily = AF_INET6;
    }

    size_t length = (size_t)(end - begin);
    if (length == 0 || length >= 64) return FALSE;

    wchar_t address[64];
    wmemcpy(address, begin, length);
    address[length] = L'\0';

    union {
        IN_ADDR ipv4;
        IN6_ADDR ipv6;
    } parsed;
    return InetPtonW(addressFamily, address, &parsed) == 1;
}

static BOOL IsValidStaticHostMapping(const wchar_t* mapping,
                                     const wchar_t** separatorOut) {
    const wchar_t* separator = wcschr(mapping, L':');
    if (!separator || separator == mapping || separator[1] == L'\0') {
        return FALSE;
    }

    const wchar_t* end = mapping + wcslen(mapping);
    if (!IsValidStaticHostName(mapping, separator) ||
        !IsValidStaticIpAddress(separator + 1, end)) {
        return FALSE;
    }

    if (separatorOut) *separatorOut = separator;
    return TRUE;
}

// Build Chromium resolver rules only from exact ASCII hostnames and literal
// IP addresses. The strict grammar prevents a registry or INI value from
// escaping the quoted switch value or introducing another browser argument.
static wchar_t* BuildStaticHostBrowserArguments(size_t* mappingCount) {
    if (mappingCount) *mappingCount = 0;
    if (!g_config.useStaticHostMappings) return NULL;

    // DNS-fallback mode: the mappings are enforced by the local fallback
    // proxy instead of resolver rules; point the browser at its PAC script.
    // Port 0 means the proxy failed to start - fall through and emit the
    // strict resolver rules so the mappings still apply.
    if (g_config.staticHostDnsFallback && g_hostProxyPort != 0) {
        const size_t argumentCch = 64;
        wchar_t* arguments = (wchar_t*)malloc(argumentCch * sizeof(wchar_t));
        if (!arguments) return NULL;
        swprintf_s(arguments, argumentCch,
                   L"--proxy-pac-url=http://127.0.0.1:%u" HOST_PROXY_PAC_PATH,
                   (unsigned)g_hostProxyPort);
        if (mappingCount) *mappingCount = g_hostProxyMappingCount;
        return arguments;
    }

    const wchar_t* configured = g_config.staticHostMappings;
    size_t configuredLength = wcslen(configured);
    size_t rulesCapacity = configuredLength * 2 + 1;
    wchar_t* rules = (wchar_t*)calloc(rulesCapacity, sizeof(wchar_t));
    if (!rules) return NULL;

    size_t rulesLength = 0;
    size_t count = 0;
    const wchar_t* cursor = configured;
    while (*cursor) {
        while (*cursor && IsOriginListSeparator(*cursor)) cursor++;
        if (!*cursor) break;

        const wchar_t* begin = cursor;
        while (*cursor && !IsOriginListSeparator(*cursor)) cursor++;
        size_t tokenLength = (size_t)(cursor - begin);
        wchar_t token[2048];
        if (tokenLength == 0 || tokenLength >= sizeof(token) / sizeof(token[0])) {
            free(rules);
            DebugPrint(L"[WARNING] Static host mapping list is too long or malformed\n");
            return NULL;
        }
        wmemcpy(token, begin, tokenLength);
        token[tokenLength] = L'\0';

        const wchar_t* separator = NULL;
        if (!IsValidStaticHostMapping(token, &separator)) {
            free(rules);
            DebugPrint(L"[WARNING] Static hosts require exact hostname:IP mappings\n");
            return NULL;
        }

        size_t hostLength = (size_t)(separator - token);
        size_t addressLength = tokenLength - hostLength - 1;
        size_t needed = (count > 0 ? 1 : 0) + 4 + hostLength + 1 + addressLength;
        if (rulesLength + needed + 1 > rulesCapacity) {
            free(rules);
            return NULL;
        }

        if (count > 0) rules[rulesLength++] = L',';
        wmemcpy(rules + rulesLength, L"MAP ", 4);
        rulesLength += 4;
        wmemcpy(rules + rulesLength, token, hostLength);
        rulesLength += hostLength;
        rules[rulesLength++] = L' ';
        wmemcpy(rules + rulesLength, separator + 1, addressLength);
        rulesLength += addressLength;
        rules[rulesLength] = L'\0';
        count++;
    }

    if (count == 0) {
        free(rules);
        DebugPrint(L"[WARNING] Static host mapping is enabled but no mappings are configured\n");
        return NULL;
    }

    static const wchar_t argumentPrefix[] = L"--host-resolver-rules=\"";
    size_t argumentLength = wcslen(argumentPrefix) + rulesLength + 1;
    wchar_t* arguments = (wchar_t*)malloc((argumentLength + 1) * sizeof(wchar_t));
    if (!arguments) {
        free(rules);
        return NULL;
    }
    wcscpy_s(arguments, argumentLength + 1, argumentPrefix);
    wcscat_s(arguments, argumentLength + 1, rules);
    wcscat_s(arguments, argumentLength + 1, L"\"");
    free(rules);
    if (mappingCount) *mappingCount = count;
    return arguments;
}

static wchar_t* JoinBrowserArguments(LPCWSTR first, LPCWSTR second) {
    size_t firstLength = first ? wcslen(first) : 0;
    size_t secondLength = second ? wcslen(second) : 0;
    if (firstLength == 0 && secondLength == 0) return NULL;

    size_t totalLength = firstLength + secondLength +
                         (firstLength > 0 && secondLength > 0 ? 1 : 0);
    wchar_t* joined = (wchar_t*)malloc((totalLength + 1) * sizeof(wchar_t));
    if (!joined) return NULL;
    joined[0] = L'\0';
    if (firstLength > 0) wcscat_s(joined, totalLength + 1, first);
    if (firstLength > 0 && secondLength > 0) {
        wcscat_s(joined, totalLength + 1, L" ");
    }
    if (secondLength > 0) wcscat_s(joined, totalLength + 1, second);
    return joined;
}

// --- Static host DNS fallback proxy ---
//
// When "fall back to standard DNS" is enabled for the static host mappings,
// the mappings are enforced here instead of through --host-resolver-rules.
// The browser is pointed at a generated PAC script (served by this listener)
// that routes only the mapped hostnames through the proxy with a DIRECT
// fallback, so a dead proxy degrades to plain direct connections. For each
// mapped hostname the proxy keeps a small circuit breaker: prefer the mapped
// address, fall back to standard DNS resolution within the same connection
// when it does not answer, and while fallen back re-try the mapped address
// at most once per HOST_PROXY_PROBE_INTERVAL_MS via a side-car probe that
// never delays the request that triggered it.

static BOOL HostProxySendAll(SOCKET s, const char* data, int length) {
    int sent = 0;
    while (sent < length) {
        int chunk = send(s, data + sent, length - sent, 0);
        if (chunk <= 0) return FALSE;
        sent += chunk;
    }
    return TRUE;
}

static void HostProxySendSimpleResponse(SOCKET s, const char* status) {
    char response[128];
    int length = snprintf(response, sizeof(response),
                          "HTTP/1.1 %s\r\nContent-Length: 0\r\nConnection: close\r\n\r\n",
                          status);
    if (length > 0) HostProxySendAll(s, response, length);
}

static void HostProxyServePac(SOCKET client) {
    if (!g_hostProxyPacScript) {
        HostProxySendSimpleResponse(client, "404 Not Found");
        return;
    }
    size_t bodyLength = strlen(g_hostProxyPacScript);
    char header[160];
    int headerLength = snprintf(header, sizeof(header),
                                "HTTP/1.1 200 OK\r\n"
                                "Content-Type: application/x-ns-proxy-autoconfig\r\n"
                                "Content-Length: %u\r\n"
                                "Connection: close\r\n\r\n",
                                (unsigned)bodyLength);
    if (headerLength > 0 && HostProxySendAll(client, header, headerLength)) {
        HostProxySendAll(client, g_hostProxyPacScript, (int)bodyLength);
    }
}

// Non-blocking connect with a timeout. Completion is detected with select()
// plus SO_ERROR rather than WSAPoll: WSAPoll on Windows builds before
// 10 2004 never reports failed connect attempts, which would turn every
// unreachable mapped address - the exact case this feature handles - into a
// hang. WSAPoll is only used for established-socket readiness elsewhere.
static SOCKET HostProxyConnectWithTimeout(const struct addrinfo* address, DWORD timeoutMs) {
    SOCKET s = socket(address->ai_family, address->ai_socktype, address->ai_protocol);
    if (s == INVALID_SOCKET) return INVALID_SOCKET;
    u_long nonBlocking = 1;
    if (ioctlsocket(s, FIONBIO, &nonBlocking) != 0) {
        closesocket(s);
        return INVALID_SOCKET;
    }
    if (connect(s, address->ai_addr, (int)address->ai_addrlen) != 0) {
        if (WSAGetLastError() != WSAEWOULDBLOCK) {
            closesocket(s);
            return INVALID_SOCKET;
        }
        fd_set writeSet, exceptSet;
        FD_ZERO(&writeSet);
        FD_ZERO(&exceptSet);
        FD_SET(s, &writeSet);
        FD_SET(s, &exceptSet);
        struct timeval timeout;
        timeout.tv_sec = (long)(timeoutMs / 1000);
        timeout.tv_usec = (long)((timeoutMs % 1000) * 1000);
        int selected = select(0, NULL, &writeSet, &exceptSet, &timeout);
        int soError = 0;
        int soErrorLength = sizeof(soError);
        if (selected <= 0 ||
            getsockopt(s, SOL_SOCKET, SO_ERROR, (char*)&soError, &soErrorLength) != 0 ||
            soError != 0) {
            closesocket(s);
            return INVALID_SOCKET;
        }
    }
    u_long blocking = 0;
    if (ioctlsocket(s, FIONBIO, &blocking) != 0) {
        closesocket(s);
        return INVALID_SOCKET;
    }
    return s;
}

// Connect to the configured mapped address. Shared by the in-band attempt
// and the side-car probe so both agree on what "reachable" means.
static SOCKET HostProxyConnectMapped(const HostProxyMapping* mapping, unsigned short port) {
    char portString[8];
    snprintf(portString, sizeof(portString), "%u", (unsigned)port);
    struct addrinfo hints;
    memset(&hints, 0, sizeof(hints));
    hints.ai_family = mapping->addressFamily;
    hints.ai_socktype = SOCK_STREAM;
    hints.ai_protocol = IPPROTO_TCP;
    hints.ai_flags = AI_NUMERICHOST | AI_NUMERICSERV;
    struct addrinfo* result = NULL;
    if (getaddrinfo(mapping->address, portString, &hints, &result) != 0 || !result) {
        return INVALID_SOCKET;
    }
    SOCKET s = HostProxyConnectWithTimeout(result, HOST_PROXY_MAPPED_CONNECT_TIMEOUT_MS);
    freeaddrinfo(result);
    return s;
}

// Standard system resolution - hosts file, configured DNS and caches all
// behave exactly as they would for a direct connection. On success the
// numeric address that answered is reported for the title's route display.
static SOCKET HostProxyConnectViaDns(const char* host, unsigned short port,
                                     char* usedAddress, size_t usedAddressSize) {
    char portString[8];
    snprintf(portString, sizeof(portString), "%u", (unsigned)port);
    if (usedAddress && usedAddressSize > 0) usedAddress[0] = '\0';
    struct addrinfo hints;
    memset(&hints, 0, sizeof(hints));
    hints.ai_family = AF_UNSPEC;
    hints.ai_socktype = SOCK_STREAM;
    hints.ai_protocol = IPPROTO_TCP;
    struct addrinfo* results = NULL;
    if (getaddrinfo(host, portString, &hints, &results) != 0 || !results) {
        return INVALID_SOCKET;
    }
    SOCKET s = INVALID_SOCKET;
    for (const struct addrinfo* address = results; address; address = address->ai_next) {
        if (InterlockedCompareExchange(&g_hostProxyStopping, FALSE, FALSE)) break;
        s = HostProxyConnectWithTimeout(address, HOST_PROXY_DNS_CONNECT_TIMEOUT_MS);
        if (s != INVALID_SOCKET) {
            if (usedAddress && usedAddressSize > 0 &&
                getnameinfo(address->ai_addr, (socklen_t)address->ai_addrlen,
                            usedAddress, (DWORD)usedAddressSize, NULL, 0,
                            NI_NUMERICHOST) != 0) {
                usedAddress[0] = '\0';
            }
            break;
        }
    }
    freeaddrinfo(results);
    return s;
}

// Force the live tunnels of a mapping off their current path so traffic
// migrates when the breaker flips. Only shuts the sockets down; the owning
// connection threads notice, exit and clean up (see HostProxyTunnel).
static void CloseHostTunnelsForMapping(int mappingIndex, HostProxyTunnel* except) {
    EnterCriticalSection(&g_hostProxyLock);
    for (HostProxyTunnel* node = g_hostProxyTunnelList; node; node = node->next) {
        if (node == except || node->mappingIndex != mappingIndex) continue;
        InterlockedExchange(&node->abortRequested, TRUE);
        if (node->client != INVALID_SOCKET) shutdown(node->client, SD_BOTH);
        if (node->upstream != INVALID_SOCKET) shutdown(node->upstream, SD_BOTH);
    }
    LeaveCriticalSection(&g_hostProxyLock);
}

// Records the address most recently used to reach a mapped host and nudges
// the UI thread to refresh the window title when the route actually changed.
static void HostProxySetCurrentAddress(HostProxyMapping* mapping,
                                       const char* address) {
    BOOL changed = FALSE;
    size_t length;
    if (!address || !address[0]) return;
    length = strlen(address);
    if (length >= sizeof(mapping->currentAddress)) return;
    EnterCriticalSection(&g_hostProxyLock);
    if (strcmp(mapping->currentAddress, address) != 0) {
        memcpy(mapping->currentAddress, address, length + 1);
        changed = TRUE;
    }
    LeaveCriticalSection(&g_hostProxyLock);
    if (changed && g_hwnd) {
        PostMessageW(g_hwnd, WM_APP_HOST_ROUTE_CHANGED, 0, 0);
    }
}

typedef struct {
    int mappingIndex;
    unsigned short port;
} HostProxyProbeTask;

static DWORD WINAPI HostProxyProbeThread(LPVOID param) {
    HostProxyProbeTask* task = (HostProxyProbeTask*)param;
    HostProxyMapping* mapping = &g_hostProxyMappings[task->mappingIndex];
    SOCKET probe = HostProxyConnectMapped(mapping, task->port);
    BOOL recovered = FALSE;
    EnterCriticalSection(&g_hostProxyLock);
    mapping->probeInFlight = FALSE;
    if (probe != INVALID_SOCKET) {
        if (!InterlockedCompareExchange(&g_hostProxyStopping, FALSE, FALSE) &&
            mapping->state == HOST_PROXY_FALLBACK) {
            mapping->state = HOST_PROXY_MAPPED_ACTIVE;
            recovered = TRUE;
        }
    } else {
        mapping->lastProbeTick = GetTickCount64();
    }
    LeaveCriticalSection(&g_hostProxyLock);
    if (probe != INVALID_SOCKET) closesocket(probe);
    if (recovered) {
        DebugPrint(L"[INFO] Mapped address for '%S' answered; resuming mapped routing\n",
                   mapping->host);
        HostProxySetCurrentAddress(mapping, mapping->address);
        CloseHostTunnelsForMapping(task->mappingIndex, NULL);
    }
    free(task);
    InterlockedDecrement(&g_hostProxyWorkerCount);
    return 0;
}

// Side-car probe: fired by a request that arrives while a mapping is fallen
// back and its cooldown has elapsed. The triggering request proceeds via DNS
// immediately; only the NEXT connections benefit from a successful probe.
static void HostProxyStartProbeIfDue(int mappingIndex, unsigned short port) {
    HostProxyMapping* mapping = &g_hostProxyMappings[mappingIndex];
    BOOL launch = FALSE;
    EnterCriticalSection(&g_hostProxyLock);
    if (mapping->state == HOST_PROXY_FALLBACK && !mapping->probeInFlight &&
        GetTickCount64() - mapping->lastProbeTick >= HOST_PROXY_PROBE_INTERVAL_MS) {
        mapping->probeInFlight = TRUE;
        launch = TRUE;
    }
    LeaveCriticalSection(&g_hostProxyLock);
    if (!launch) return;

    HostProxyProbeTask* task = (HostProxyProbeTask*)malloc(sizeof(*task));
    HANDLE thread = NULL;
    if (task) {
        task->mappingIndex = mappingIndex;
        task->port = port;
        InterlockedIncrement(&g_hostProxyWorkerCount);
        thread = CreateThread(NULL, 0, HostProxyProbeThread, task, 0, NULL);
        if (!thread) {
            InterlockedDecrement(&g_hostProxyWorkerCount);
            free(task);
        }
    }
    if (thread) {
        CloseHandle(thread);
    } else {
        // Roll back the single-flight claim or probing would wedge forever.
        EnterCriticalSection(&g_hostProxyLock);
        mapping->probeInFlight = FALSE;
        LeaveCriticalSection(&g_hostProxyLock);
    }
}

// The circuit breaker. Establishes the upstream connection for one request,
// never dropping it: when the mapped address fails the same connection is
// retried through standard DNS before giving up. Only connect-phase results
// move the breaker - mid-stream closes are normal (keep-alive teardown) and
// must not trip it.
static SOCKET HostProxyEstablishUpstream(HostProxyTunnel* tunnel, int mappingIndex,
                                         unsigned short port) {
    HostProxyMapping* mapping = &g_hostProxyMappings[mappingIndex];
    EnterCriticalSection(&g_hostProxyLock);
    tunnel->mappingIndex = mappingIndex;
    HostProxyBreakerState state = mapping->state;
    LeaveCriticalSection(&g_hostProxyLock);

    SOCKET upstream = INVALID_SOCKET;
    if (state == HOST_PROXY_FALLBACK) {
        char dnsAddress[64];
        HostProxyStartProbeIfDue(mappingIndex, port);
        upstream = HostProxyConnectViaDns(mapping->host, port,
                                          dnsAddress, sizeof(dnsAddress));
        if (upstream != INVALID_SOCKET) {
            HostProxySetCurrentAddress(mapping, dnsAddress);
        }
    } else {
        upstream = HostProxyConnectMapped(mapping, port);
        if (upstream != INVALID_SOCKET) {
            BOOL announced = FALSE;
            EnterCriticalSection(&g_hostProxyLock);
            if (mapping->state != HOST_PROXY_MAPPED_ACTIVE) {
                mapping->state = HOST_PROXY_MAPPED_ACTIVE;
                announced = TRUE;
            }
            LeaveCriticalSection(&g_hostProxyLock);
            if (announced) {
                DebugPrint(L"[INFO] Static host '%S' using mapped address\n", mapping->host);
            }
            HostProxySetCurrentAddress(mapping, mapping->address);
        } else {
            BOOL wasActive;
            EnterCriticalSection(&g_hostProxyLock);
            wasActive = (mapping->state == HOST_PROXY_MAPPED_ACTIVE);
            mapping->state = HOST_PROXY_FALLBACK;
            mapping->lastProbeTick = GetTickCount64();
            LeaveCriticalSection(&g_hostProxyLock);
            DebugPrint(L"[WARNING] Mapped address for '%S' unreachable; using DNS resolution\n",
                       mapping->host);
            if (wasActive) CloseHostTunnelsForMapping(mappingIndex, tunnel);
            char dnsAddress[64];
            upstream = HostProxyConnectViaDns(mapping->host, port,
                                              dnsAddress, sizeof(dnsAddress));
            if (upstream != INVALID_SOCKET) {
                HostProxySetCurrentAddress(mapping, dnsAddress);
            }
        }
    }
    if (upstream != INVALID_SOCKET) {
        EnterCriticalSection(&g_hostProxyLock);
        tunnel->upstream = upstream;
        LeaveCriticalSection(&g_hostProxyLock);
        // An eviction that raced the connect above may have missed the new
        // socket; make sure it observes the abort immediately.
        if (InterlockedCompareExchange(&tunnel->abortRequested, FALSE, FALSE)) {
            shutdown(upstream, SD_BOTH);
        }
    }
    return upstream;
}

static int FindHostProxyMapping(const char* host) {
    for (size_t i = 0; i < g_hostProxyMappingCount; i++) {
        if (strcmp(g_hostProxyMappings[i].host, host) == 0) return (int)i;
    }
    return -1;
}

// Reads until the blank line ending the request head. Bytes that arrive
// beyond the head (early tunnel data, a request body prefix) are kept in the
// buffer - *totalRead past *headLength - and must be forwarded, not dropped.
// Returns 0 on success, 1 on an oversized head, 2 on timeout/close/abort.
static int HostProxyReadRequestHead(HostProxyTunnel* tunnel, char* buffer, int capacity,
                                    int* headLength, int* totalRead) {
    int received = 0;
    int scanned = 0;
    ULONGLONG deadline = GetTickCount64() + HOST_PROXY_HEAD_READ_TIMEOUT_MS;
    *headLength = 0;
    *totalRead = 0;
    for (;;) {
        for (int i = scanned; i < received; i++) {
            if (buffer[i] != '\n' || i < 1) continue;
            // A blank line ends the head: either a bare LF or a CRLF, no
            // matter how the preceding line was terminated.
            if (buffer[i - 1] == '\n' ||
                (i >= 2 && buffer[i - 1] == '\r' && buffer[i - 2] == '\n')) {
                *headLength = i + 1;
                *totalRead = received;
                return 0;
            }
        }
        scanned = (received > 3) ? received - 3 : 0;
        if (received >= capacity - 1) return 1;
        if (InterlockedCompareExchange(&tunnel->abortRequested, FALSE, FALSE) ||
            InterlockedCompareExchange(&g_hostProxyStopping, FALSE, FALSE)) {
            return 2;
        }
        ULONGLONG now = GetTickCount64();
        if (now >= deadline) return 2;
        ULONGLONG remaining = deadline - now;
        INT wait = (remaining < HOST_PROXY_POLL_TICK_MS)
                       ? (INT)remaining : HOST_PROXY_POLL_TICK_MS;
        WSAPOLLFD pollFd;
        pollFd.fd = tunnel->client;
        pollFd.events = POLLRDNORM;
        pollFd.revents = 0;
        int pollResult = WSAPoll(&pollFd, 1, wait);
        if (pollResult < 0) return 2;
        if (pollResult == 0) continue;
        int chunk = recv(tunnel->client, buffer + received, capacity - 1 - received, 0);
        if (chunk <= 0) return 2;
        received += chunk;
    }
}

typedef struct {
    char method[16];
    char host[254];
    unsigned short port;
    BOOL isConnect;
    BOOL isOriginForm;
    const char* path;     // absolute-form: path+query start within the head
    int pathLength;
    const char* headers;  // first byte after the request line
} HostProxyRequest;

static BOOL HostProxyParseAuthority(const char* authority, int length, BOOL requirePort,
                                    unsigned short defaultPort, char* host,
                                    size_t hostSize, unsigned short* port) {
    if (length <= 0) return FALSE;
    // Reject userinfo outright; browsers never send it to a proxy.
    for (int i = 0; i < length; i++) {
        if (authority[i] == '@') return FALSE;
    }
    const char* hostBegin = authority;
    const char* hostEnd = NULL;
    const char* portBegin = NULL;
    if (authority[0] == '[') {
        const char* closeBracket = (const char*)memchr(authority, ']', (size_t)length);
        if (!closeBracket || closeBracket == authority + 1) return FALSE;
        hostBegin = authority + 1;
        hostEnd = closeBracket;
        if (closeBracket + 1 < authority + length) {
            if (closeBracket[1] != ':') return FALSE;
            portBegin = closeBracket + 2;
        }
    } else {
        const char* colon = (const char*)memchr(authority, ':', (size_t)length);
        if (colon) {
            hostEnd = colon;
            portBegin = colon + 1;
        } else {
            hostEnd = authority + length;
        }
    }
    size_t hostLength = (size_t)(hostEnd - hostBegin);
    if (hostLength == 0 || hostLength >= hostSize) return FALSE;
    for (size_t i = 0; i < hostLength; i++) {
        char c = hostBegin[i];
        if ((unsigned char)c >= 0x80) return FALSE;
        if (c >= 'A' && c <= 'Z') c = (char)(c - 'A' + 'a');
        host[i] = c;
    }
    host[hostLength] = '\0';
    if (portBegin) {
        const char* end = authority + length;
        if (portBegin >= end) return FALSE;
        unsigned long value = 0;
        for (const char* p = portBegin; p < end; p++) {
            if (*p < '0' || *p > '9') return FALSE;
            value = value * 10 + (unsigned long)(*p - '0');
            if (value > 65535) return FALSE;
        }
        if (value == 0) return FALSE;
        *port = (unsigned short)value;
    } else {
        if (requirePort) return FALSE;
        *port = defaultPort;
    }
    return TRUE;
}

// Parses a NUL-terminated request head into its routing-relevant pieces:
// CONNECT authority-form, plain-HTTP absolute-form, or origin-form (used
// only for the PAC endpoint). Anything else is rejected.
static BOOL HostProxyParseRequestHead(const char* head, HostProxyRequest* request) {
    memset(request, 0, sizeof(*request));
    const char* lineEnd = strchr(head, '\n');
    if (!lineEnd) return FALSE;
    request->headers = lineEnd + 1;
    const char* requestLineEnd = lineEnd;
    if (requestLineEnd > head && requestLineEnd[-1] == '\r') requestLineEnd--;
    const char* methodEnd = (const char*)memchr(head, ' ', (size_t)(requestLineEnd - head));
    if (!methodEnd) return FALSE;
    size_t methodLength = (size_t)(methodEnd - head);
    if (methodLength == 0 || methodLength >= sizeof(request->method)) return FALSE;
    memcpy(request->method, head, methodLength);
    request->method[methodLength] = '\0';
    const char* target = methodEnd + 1;
    if (target >= requestLineEnd) return FALSE;
    const char* targetEnd = (const char*)memchr(target, ' ', (size_t)(requestLineEnd - target));
    if (!targetEnd || targetEnd == target) return FALSE;
    int targetLength = (int)(targetEnd - target);

    if (strcmp(request->method, "CONNECT") == 0) {
        request->isConnect = TRUE;
        // Authority-form; browsers always include the port here.
        return HostProxyParseAuthority(target, targetLength, TRUE, 0, request->host,
                                       sizeof(request->host), &request->port);
    }
    if (target[0] == '/') {
        request->isOriginForm = TRUE;
        request->path = target;
        request->pathLength = targetLength;
        return TRUE;
    }
    if (targetLength > 7 && _strnicmp(target, "http://", 7) == 0) {
        const char* authority = target + 7;
        int authorityLength = targetLength - 7;
        int authorityEnd = 0;
        while (authorityEnd < authorityLength && authority[authorityEnd] != '/' &&
               authority[authorityEnd] != '?') {
            authorityEnd++;
        }
        if (!HostProxyParseAuthority(authority, authorityEnd, FALSE, 80, request->host,
                                     sizeof(request->host), &request->port)) {
            return FALSE;
        }
        request->path = authority + authorityEnd;
        request->pathLength = authorityLength - authorityEnd;
        return TRUE;
    }
    return FALSE;
}

// Rebuilds an absolute-form plain-HTTP head as an origin-form request with
// close-delimited framing: the request line is rewritten to the path, the
// hop-by-hop headers are dropped, and "Connection: close" is forced so the
// origin server ends the exchange (the proxy does not parse responses).
static char* HostProxyRewriteAbsoluteHead(const HostProxyRequest* request,
                                          int* rewrittenLength) {
    size_t capacity = HOST_PROXY_HEAD_MAX_BYTES + 128;
    char* rewritten = (char*)malloc(capacity);
    if (!rewritten) return NULL;
    int written = snprintf(rewritten, capacity, "%s %s%.*s HTTP/1.1\r\n",
                           request->method,
                           (request->pathLength == 0 || request->path[0] == '?') ? "/" : "",
                           request->pathLength, request->path);
    if (written < 0) {
        free(rewritten);
        return NULL;
    }
    const char* cursor = request->headers;
    while (*cursor) {
        const char* nextLine = strchr(cursor, '\n');
        const char* lineEnd = nextLine ? nextLine : cursor + strlen(cursor);
        const char* trimmedEnd = lineEnd;
        if (trimmedEnd > cursor && trimmedEnd[-1] == '\r') trimmedEnd--;
        if (trimmedEnd == cursor) break;  // blank line: end of the headers
        size_t lineLength = (size_t)(trimmedEnd - cursor);
        if (!(lineLength >= 6 && _strnicmp(cursor, "Proxy-", 6) == 0) &&
            !(lineLength >= 11 && _strnicmp(cursor, "Connection:", 11) == 0) &&
            !(lineLength >= 10 && _strnicmp(cursor, "Keep-Alive", 10) == 0)) {
            if ((size_t)written + lineLength + 2 >= capacity) {
                free(rewritten);
                return NULL;
            }
            memcpy(rewritten + written, cursor, lineLength);
            written += (int)lineLength;
            rewritten[written++] = '\r';
            rewritten[written++] = '\n';
        }
        if (!nextLine) break;
        cursor = nextLine + 1;
    }
    static const char terminator[] = "Connection: close\r\n\r\n";
    size_t terminatorLength = sizeof(terminator) - 1;
    if ((size_t)written + terminatorLength >= capacity) {
        free(rewritten);
        return NULL;
    }
    memcpy(rewritten + written, terminator, terminatorLength);
    written += (int)terminatorLength;
    *rewrittenLength = written;
    return rewritten;
}

// Blind bidirectional byte pump. Deliberately touches no breaker state: a
// dying tunnel is normal (keep-alive teardown, page navigation) and must
// not be mistaken for an unreachable mapped address.
static void HostProxyPumpTunnel(HostProxyTunnel* tunnel, const char* leftover,
                                int leftoverLength) {
    char* buffer = (char*)malloc(HOST_PROXY_IO_BUFFER_BYTES);
    if (!buffer) return;
    if (leftoverLength > 0 &&
        !HostProxySendAll(tunnel->upstream, leftover, leftoverLength)) {
        free(buffer);
        return;
    }
    BOOL clientOpen = TRUE;
    BOOL upstreamOpen = TRUE;
    while (clientOpen || upstreamOpen) {
        if (InterlockedCompareExchange(&tunnel->abortRequested, FALSE, FALSE) ||
            InterlockedCompareExchange(&g_hostProxyStopping, FALSE, FALSE)) {
            break;
        }
        WSAPOLLFD fds[2];
        int fdCount = 0;
        int clientIndex = -1;
        int upstreamIndex = -1;
        if (clientOpen) {
            fds[fdCount].fd = tunnel->client;
            fds[fdCount].events = POLLRDNORM;
            fds[fdCount].revents = 0;
            clientIndex = fdCount++;
        }
        if (upstreamOpen) {
            fds[fdCount].fd = tunnel->upstream;
            fds[fdCount].events = POLLRDNORM;
            fds[fdCount].revents = 0;
            upstreamIndex = fdCount++;
        }
        int pollResult = WSAPoll(fds, (ULONG)fdCount, HOST_PROXY_POLL_TICK_MS);
        if (pollResult < 0) break;
        if (pollResult == 0) continue;
        if (clientIndex >= 0 && fds[clientIndex].revents) {
            if (fds[clientIndex].revents & POLLRDNORM) {
                int got = recv(tunnel->client, buffer, HOST_PROXY_IO_BUFFER_BYTES, 0);
                if (got <= 0) {
                    clientOpen = FALSE;
                    shutdown(tunnel->upstream, SD_SEND);
                } else if (!HostProxySendAll(tunnel->upstream, buffer, got)) {
                    break;
                }
            } else {
                clientOpen = FALSE;
                shutdown(tunnel->upstream, SD_SEND);
            }
        }
        if (upstreamIndex >= 0 && fds[upstreamIndex].revents) {
            if (fds[upstreamIndex].revents & POLLRDNORM) {
                int got = recv(tunnel->upstream, buffer, HOST_PROXY_IO_BUFFER_BYTES, 0);
                if (got <= 0) {
                    upstreamOpen = FALSE;
                    shutdown(tunnel->client, SD_SEND);
                } else if (!HostProxySendAll(tunnel->client, buffer, got)) {
                    break;
                }
            } else {
                upstreamOpen = FALSE;
                shutdown(tunnel->client, SD_SEND);
            }
        }
    }
    free(buffer);
}

static DWORD WINAPI HostProxyConnectionThread(LPVOID param) {
    HostProxyTunnel* tunnel = (HostProxyTunnel*)param;
    char* head = (char*)malloc(HOST_PROXY_HEAD_MAX_BYTES);
    char* rewrittenHead = NULL;
    int headLength = 0;
    int totalRead = 0;
    int rewrittenLength = 0;

    if (!head) goto cleanup;
    {
        int readResult = HostProxyReadRequestHead(tunnel, head, HOST_PROXY_HEAD_MAX_BYTES,
                                                  &headLength, &totalRead);
        if (readResult == 1) {
            HostProxySendSimpleResponse(tunnel->client, "400 Bad Request");
            goto cleanup;
        }
        if (readResult != 0) goto cleanup;
    }
    {
        // NUL-terminate the head for parsing; the byte at headLength is the
        // start of any early tunnel data and is restored before forwarding.
        char savedByte = head[headLength];
        head[headLength] = '\0';
        HostProxyRequest request;
        BOOL parsed = HostProxyParseRequestHead(head, &request);

        if (parsed && request.isOriginForm) {
            if (strcmp(request.method, "GET") == 0 &&
                request.pathLength == (int)(sizeof(HOST_PROXY_PAC_PATH) - 1) &&
                strncmp(request.path, HOST_PROXY_PAC_PATH,
                        (size_t)request.pathLength) == 0) {
                HostProxyServePac(tunnel->client);
            } else {
                HostProxySendSimpleResponse(tunnel->client, "404 Not Found");
            }
            goto cleanup;
        }
        if (!parsed) {
            HostProxySendSimpleResponse(tunnel->client, "400 Bad Request");
            goto cleanup;
        }

        int mappingIndex = FindHostProxyMapping(request.host);
        if (mappingIndex < 0) {
            // Only the configured static hostnames are proxied; refusing
            // everything else keeps the listener from being an open proxy.
            HostProxySendSimpleResponse(tunnel->client, "403 Forbidden");
            goto cleanup;
        }

        if (!request.isConnect) {
            rewrittenHead = HostProxyRewriteAbsoluteHead(&request, &rewrittenLength);
            if (!rewrittenHead) {
                HostProxySendSimpleResponse(tunnel->client, "400 Bad Request");
                goto cleanup;
            }
        }

        head[headLength] = savedByte;

        SOCKET upstream = HostProxyEstablishUpstream(tunnel, mappingIndex, request.port);
        if (upstream == INVALID_SOCKET) {
            HostProxySendSimpleResponse(tunnel->client, "502 Bad Gateway");
            goto cleanup;
        }

        if (request.isConnect) {
            static const char established[] = "HTTP/1.1 200 Connection Established\r\n\r\n";
            if (!HostProxySendAll(tunnel->client, established,
                                  (int)(sizeof(established) - 1))) {
                goto cleanup;
            }
        } else {
            if (!HostProxySendAll(upstream, rewrittenHead, rewrittenLength)) goto cleanup;
        }
        HostProxyPumpTunnel(tunnel, head + headLength, totalRead - headLength);
    }

cleanup:
    EnterCriticalSection(&g_hostProxyLock);
    if (tunnel->prev) tunnel->prev->next = tunnel->next;
    else g_hostProxyTunnelList = tunnel->next;
    if (tunnel->next) tunnel->next->prev = tunnel->prev;
    LeaveCriticalSection(&g_hostProxyLock);
    if (tunnel->client != INVALID_SOCKET) closesocket(tunnel->client);
    if (tunnel->upstream != INVALID_SOCKET) closesocket(tunnel->upstream);
    free(tunnel);
    free(head);
    free(rewrittenHead);
    InterlockedDecrement(&g_hostProxyWorkerCount);
    return 0;
}

static DWORD WINAPI HostProxyAcceptThread(LPVOID param) {
    (void)param;
    for (;;) {
        struct sockaddr_storage peer;
        int peerLength = sizeof(peer);
        SOCKET client = accept(g_hostProxyListenSocket, (struct sockaddr*)&peer,
                               &peerLength);
        if (client == INVALID_SOCKET) {
            if (InterlockedCompareExchange(&g_hostProxyStopping, FALSE, FALSE)) break;
            int error = WSAGetLastError();
            if (error == WSAECONNRESET || error == WSAEINTR) continue;
            if (error == WSAENOBUFS || error == WSAEMFILE) {
                // Transient resource exhaustion (typical shortly after a
                // resume): keep the listener alive rather than abandoning
                // the mapped path until the next launcher restart.
                Sleep(100);
                continue;
            }
            DebugPrint(L"[ERROR] Static host proxy accept failed (%d); listener stopped\n",
                       error);
            break;
        }
        // The listener is bound to 127.0.0.1 so remote peers cannot reach
        // it; drop anything unexpected anyway.
        BOOL loopback = FALSE;
        if (peer.ss_family == AF_INET) {
            loopback = (((struct sockaddr_in*)&peer)->sin_addr.s_addr ==
                        htonl(INADDR_LOOPBACK));
        }
        if (!loopback ||
            InterlockedCompareExchange(&g_hostProxyWorkerCount, 0, 0) >=
                HOST_PROXY_MAX_TUNNELS) {
            // Over capacity: refuse; the PAC's DIRECT fallback keeps loads
            // working while the browser backs off.
            closesocket(client);
            continue;
        }
        HostProxyTunnel* tunnel = (HostProxyTunnel*)calloc(1, sizeof(*tunnel));
        if (!tunnel) {
            closesocket(client);
            continue;
        }
        tunnel->client = client;
        tunnel->upstream = INVALID_SOCKET;
        tunnel->mappingIndex = -1;
        EnterCriticalSection(&g_hostProxyLock);
        tunnel->next = g_hostProxyTunnelList;
        if (g_hostProxyTunnelList) g_hostProxyTunnelList->prev = tunnel;
        g_hostProxyTunnelList = tunnel;
        LeaveCriticalSection(&g_hostProxyLock);
        InterlockedIncrement(&g_hostProxyWorkerCount);
        HANDLE thread = CreateThread(NULL, 0, HostProxyConnectionThread, tunnel, 0, NULL);
        if (!thread) {
            EnterCriticalSection(&g_hostProxyLock);
            if (tunnel->prev) tunnel->prev->next = tunnel->next;
            else g_hostProxyTunnelList = tunnel->next;
            if (tunnel->next) tunnel->next->prev = tunnel->prev;
            LeaveCriticalSection(&g_hostProxyLock);
            InterlockedDecrement(&g_hostProxyWorkerCount);
            closesocket(client);
            free(tunnel);
            continue;
        }
        CloseHandle(thread);
    }
    return 0;
}

// Builds the runtime mapping table from the same configuration string (and
// with the same all-or-nothing validation) as the strict resolver rules.
static BOOL ParseStaticHostProxyMappings(void) {
    HostProxyMapping* mappings =
        (HostProxyMapping*)calloc(HOST_PROXY_MAX_MAPPINGS, sizeof(HostProxyMapping));
    if (!mappings) return FALSE;

    size_t count = 0;
    const wchar_t* cursor = g_config.staticHostMappings;
    while (*cursor) {
        while (*cursor && IsOriginListSeparator(*cursor)) cursor++;
        if (!*cursor) break;

        const wchar_t* begin = cursor;
        while (*cursor && !IsOriginListSeparator(*cursor)) cursor++;
        size_t tokenLength = (size_t)(cursor - begin);
        wchar_t token[2048];
        if (tokenLength == 0 || tokenLength >= sizeof(token) / sizeof(token[0])) {
            free(mappings);
            return FALSE;
        }
        wmemcpy(token, begin, tokenLength);
        token[tokenLength] = L'\0';

        const wchar_t* separator = NULL;
        if (!IsValidStaticHostMapping(token, &separator)) {
            free(mappings);
            return FALSE;
        }
        if (count >= HOST_PROXY_MAX_MAPPINGS) break;

        HostProxyMapping* mapping = &mappings[count];
        size_t hostLength = (size_t)(separator - token);
        if (hostLength >= sizeof(mapping->host)) {
            free(mappings);
            return FALSE;
        }
        for (size_t i = 0; i < hostLength; i++) {
            wchar_t c = token[i];
            if (c >= L'A' && c <= L'Z') c = c - L'A' + L'a';
            mapping->host[i] = (char)c;
        }
        mapping->host[hostLength] = '\0';

        const wchar_t* address = separator + 1;
        size_t addressLength = wcslen(address);
        BOOL ipv6 = (address[0] == L'[');
        if (ipv6) {
            address++;
            addressLength -= 2;
        }
        if (addressLength == 0 || addressLength >= sizeof(mapping->address)) {
            free(mappings);
            return FALSE;
        }
        for (size_t i = 0; i < addressLength; i++) {
            mapping->address[i] = (char)address[i];
        }
        mapping->address[addressLength] = '\0';
        mapping->addressFamily = ipv6 ? AF_INET6 : AF_INET;
        mapping->state = HOST_PROXY_UNTESTED;
        mapping->lastProbeTick = 0;
        mapping->probeInFlight = FALSE;
        memcpy(mapping->currentAddress, mapping->address,
               sizeof(mapping->currentAddress));

        // Duplicate hostname: first entry wins, matching resolver rules.
        BOOL duplicate = FALSE;
        for (size_t i = 0; i < count; i++) {
            if (strcmp(mappings[i].host, mapping->host) == 0) {
                duplicate = TRUE;
                break;
            }
        }
        if (duplicate) memset(mapping, 0, sizeof(*mapping));
        else count++;
    }

    if (count == 0) {
        free(mappings);
        return FALSE;
    }
    g_hostProxyMappings = mappings;
    g_hostProxyMappingCount = count;
    return TRUE;
}

static char* BuildHostProxyPacScript(void) {
    size_t capacity = 192;
    for (size_t i = 0; i < g_hostProxyMappingCount; i++) {
        capacity += strlen(g_hostProxyMappings[i].host) + 32;
    }
    char* script = (char*)malloc(capacity);
    if (!script) return NULL;
    int written = snprintf(script, capacity,
                           "function FindProxyForURL(url, host) {\n"
                           "  host = host.toLowerCase();\n"
                           "  if (");
    if (written < 0) {
        free(script);
        return NULL;
    }
    for (size_t i = 0; i < g_hostProxyMappingCount; i++) {
        int chunk = snprintf(script + written, capacity - (size_t)written,
                             "%shost == \"%s\"", i > 0 ? " ||\n      " : "",
                             g_hostProxyMappings[i].host);
        if (chunk < 0 || (size_t)written + (size_t)chunk >= capacity) {
            free(script);
            return NULL;
        }
        written += chunk;
    }
    int chunk = snprintf(script + written, capacity - (size_t)written,
                         ")\n    return \"PROXY 127.0.0.1:%u; DIRECT\";\n"
                         "  return \"DIRECT\";\n}\n",
                         (unsigned)g_hostProxyPort);
    if (chunk < 0 || (size_t)written + (size_t)chunk >= capacity) {
        free(script);
        return NULL;
    }
    return script;
}

// Starts the fallback proxy: parse the mapping table, bind an ephemeral
// loopback port (queried back so the PAC URL can embed it), build the PAC
// script and spawn the accept loop. Fails soft - the caller falls back to
// the strict resolver rules and the feature degrades to today's behavior.
static BOOL StartStaticHostProxy(void) {
    if (!g_winsockInitialized) {
        DebugPrint(L"[WARNING] Static host fallback proxy unavailable: Winsock init failed\n");
        return FALSE;
    }
    if (!ParseStaticHostProxyMappings()) {
        DebugPrint(L"[WARNING] Static host fallback proxy disabled: no valid mappings\n");
        return FALSE;
    }
    InitializeCriticalSection(&g_hostProxyLock);
    g_hostProxyListenSocket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
    if (g_hostProxyListenSocket == INVALID_SOCKET) goto fail;
    {
        struct sockaddr_in bindAddress;
        memset(&bindAddress, 0, sizeof(bindAddress));
        bindAddress.sin_family = AF_INET;
        bindAddress.sin_addr.s_addr = htonl(INADDR_LOOPBACK);
        bindAddress.sin_port = 0;
        if (bind(g_hostProxyListenSocket, (struct sockaddr*)&bindAddress,
                 sizeof(bindAddress)) != 0) {
            goto fail;
        }
        int addressLength = sizeof(bindAddress);
        if (getsockname(g_hostProxyListenSocket, (struct sockaddr*)&bindAddress,
                        &addressLength) != 0) {
            goto fail;
        }
        g_hostProxyPort = ntohs(bindAddress.sin_port);
    }
    if (g_hostProxyPort == 0) goto fail;
    if (listen(g_hostProxyListenSocket, HOST_PROXY_LISTEN_BACKLOG) != 0) goto fail;
    g_hostProxyPacScript = BuildHostProxyPacScript();
    if (!g_hostProxyPacScript) goto fail;
    g_hostProxyAcceptThread = CreateThread(NULL, 0, HostProxyAcceptThread, NULL, 0, NULL);
    if (!g_hostProxyAcceptThread) goto fail;
    DebugPrint(L"[INFO] Static host fallback proxy listening on 127.0.0.1:%u (%u mapping(s))\n",
               (unsigned)g_hostProxyPort, (unsigned)g_hostProxyMappingCount);
    return TRUE;

fail:
    DebugPrint(L"[WARNING] Static host fallback proxy failed to start (%d)\n",
               WSAGetLastError());
    if (g_hostProxyListenSocket != INVALID_SOCKET) {
        closesocket(g_hostProxyListenSocket);
        g_hostProxyListenSocket = INVALID_SOCKET;
    }
    free(g_hostProxyPacScript);
    g_hostProxyPacScript = NULL;
    free(g_hostProxyMappings);
    g_hostProxyMappings = NULL;
    g_hostProxyMappingCount = 0;
    DeleteCriticalSection(&g_hostProxyLock);
    g_hostProxyPort = 0;
    return FALSE;
}

// Bounded shutdown: unblock the accept loop by closing its socket, shut
// down every live tunnel, then wait for the worker count to drain. If a
// worker somehow fails to exit in time the shared state is deliberately
// leaked instead of freed under a live thread - the process is exiting.
static void StopStaticHostProxy(void) {
    if (g_hostProxyPort == 0 && g_hostProxyListenSocket == INVALID_SOCKET) return;
    InterlockedExchange(&g_hostProxyStopping, TRUE);
    if (g_hostProxyListenSocket != INVALID_SOCKET) {
        closesocket(g_hostProxyListenSocket);
        g_hostProxyListenSocket = INVALID_SOCKET;
    }
    BOOL acceptThreadExited = TRUE;
    if (g_hostProxyAcceptThread) {
        acceptThreadExited =
            (WaitForSingleObject(g_hostProxyAcceptThread,
                                 HOST_PROXY_SHUTDOWN_WAIT_MS) == WAIT_OBJECT_0);
        CloseHandle(g_hostProxyAcceptThread);
        g_hostProxyAcceptThread = NULL;
    }
    EnterCriticalSection(&g_hostProxyLock);
    for (HostProxyTunnel* node = g_hostProxyTunnelList; node; node = node->next) {
        InterlockedExchange(&node->abortRequested, TRUE);
        if (node->client != INVALID_SOCKET) shutdown(node->client, SD_BOTH);
        if (node->upstream != INVALID_SOCKET) shutdown(node->upstream, SD_BOTH);
    }
    LeaveCriticalSection(&g_hostProxyLock);
    DWORD waited = 0;
    while (InterlockedCompareExchange(&g_hostProxyWorkerCount, 0, 0) > 0 &&
           waited < HOST_PROXY_SHUTDOWN_WAIT_MS) {
        Sleep(50);
        waited += 50;
    }
    if (acceptThreadExited &&
        InterlockedCompareExchange(&g_hostProxyWorkerCount, 0, 0) == 0) {
        DeleteCriticalSection(&g_hostProxyLock);
        free(g_hostProxyMappings);
        g_hostProxyMappings = NULL;
        g_hostProxyMappingCount = 0;
        free(g_hostProxyPacScript);
        g_hostProxyPacScript = NULL;
    } else {
        DebugPrint(L"[WARNING] Static host proxy thread did not exit in time\n");
    }
    g_hostProxyPort = 0;
}

// Power-resume hook: whatever the breaker believed before a suspend is
// stale, so let the first request after resume re-probe immediately instead
// of waiting out a cooldown started before the machine went down. Mappings
// still on the mapped address self-correct on their next in-band connect.
static void HostProxyExpireFallbackCooldowns(void) {
    if (g_hostProxyPort == 0) return;
    ULONGLONG now = GetTickCount64();
    EnterCriticalSection(&g_hostProxyLock);
    for (size_t i = 0; i < g_hostProxyMappingCount; i++) {
        if (g_hostProxyMappings[i].state == HOST_PROXY_FALLBACK) {
            g_hostProxyMappings[i].lastProbeTick =
                (now > HOST_PROXY_PROBE_INTERVAL_MS)
                    ? now - HOST_PROXY_PROBE_INTERVAL_MS : 0;
        }
    }
    LeaveCriticalSection(&g_hostProxyLock);
}

// Plain-C implementation of the base WebView2 environment-options COM
// interface. The SDK's convenience implementation requires C++/WRL, while
// this application deliberately remains a single C translation unit.
typedef struct {
    ICoreWebView2EnvironmentOptionsVtbl* lpVtbl;
    LONG refCount;
    LPWSTR additionalBrowserArguments;
    LPWSTR language;
    LPWSTR targetCompatibleBrowserVersion;
    BOOL allowSingleSignOnUsingOSPrimaryAccount;
} MainEnvironmentOptions;

static HRESULT CopyEnvironmentOptionString(LPCWSTR source, LPWSTR* value) {
    if (!value) return E_POINTER;
    *value = NULL;
    if (!source) return S_OK;

    size_t bytes = (wcslen(source) + 1) * sizeof(wchar_t);
    LPWSTR copy = (LPWSTR)CoTaskMemAlloc(bytes);
    if (!copy) return E_OUTOFMEMORY;
    memcpy(copy, source, bytes);
    *value = copy;
    return S_OK;
}

static HRESULT SetEnvironmentOptionString(LPWSTR* destination, LPCWSTR value) {
    LPWSTR copy = NULL;
    if (value) {
        size_t bytes = (wcslen(value) + 1) * sizeof(wchar_t);
        copy = (LPWSTR)CoTaskMemAlloc(bytes);
        if (!copy) return E_OUTOFMEMORY;
        memcpy(copy, value, bytes);
    }
    CoTaskMemFree(*destination);
    *destination = copy;
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE MainEnvironmentOptions_QueryInterface(
    ICoreWebView2EnvironmentOptions* This, REFIID riid, void** ppvObject) {
    if (!ppvObject) return E_POINTER;
    *ppvObject = NULL;
    if (IsEqualIID(riid, &IID_IUnknown) ||
        IsEqualIID(riid, &IID_ICoreWebView2EnvironmentOptions)) {
        *ppvObject = This;
        This->lpVtbl->AddRef(This);
        return S_OK;
    }
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE MainEnvironmentOptions_AddRef(
    ICoreWebView2EnvironmentOptions* This) {
    return (ULONG)InterlockedIncrement(&((MainEnvironmentOptions*)This)->refCount);
}

static ULONG STDMETHODCALLTYPE MainEnvironmentOptions_Release(
    ICoreWebView2EnvironmentOptions* This) {
    MainEnvironmentOptions* options = (MainEnvironmentOptions*)This;
    ULONG refCount = (ULONG)InterlockedDecrement(&options->refCount);
    if (refCount == 0) {
        CoTaskMemFree(options->additionalBrowserArguments);
        CoTaskMemFree(options->language);
        CoTaskMemFree(options->targetCompatibleBrowserVersion);
        free(options);
    }
    return refCount;
}

static HRESULT STDMETHODCALLTYPE MainEnvironmentOptions_get_AdditionalBrowserArguments(
    ICoreWebView2EnvironmentOptions* This, LPWSTR* value) {
    return CopyEnvironmentOptionString(
        ((MainEnvironmentOptions*)This)->additionalBrowserArguments, value);
}

static HRESULT STDMETHODCALLTYPE MainEnvironmentOptions_put_AdditionalBrowserArguments(
    ICoreWebView2EnvironmentOptions* This, LPCWSTR value) {
    return SetEnvironmentOptionString(
        &((MainEnvironmentOptions*)This)->additionalBrowserArguments, value);
}

static HRESULT STDMETHODCALLTYPE MainEnvironmentOptions_get_Language(
    ICoreWebView2EnvironmentOptions* This, LPWSTR* value) {
    return CopyEnvironmentOptionString(((MainEnvironmentOptions*)This)->language, value);
}

static HRESULT STDMETHODCALLTYPE MainEnvironmentOptions_put_Language(
    ICoreWebView2EnvironmentOptions* This, LPCWSTR value) {
    return SetEnvironmentOptionString(&((MainEnvironmentOptions*)This)->language, value);
}

static HRESULT STDMETHODCALLTYPE MainEnvironmentOptions_get_TargetCompatibleBrowserVersion(
    ICoreWebView2EnvironmentOptions* This, LPWSTR* value) {
    return CopyEnvironmentOptionString(
        ((MainEnvironmentOptions*)This)->targetCompatibleBrowserVersion, value);
}

static HRESULT STDMETHODCALLTYPE MainEnvironmentOptions_put_TargetCompatibleBrowserVersion(
    ICoreWebView2EnvironmentOptions* This, LPCWSTR value) {
    return SetEnvironmentOptionString(
        &((MainEnvironmentOptions*)This)->targetCompatibleBrowserVersion, value);
}

static HRESULT STDMETHODCALLTYPE MainEnvironmentOptions_get_AllowSingleSignOnUsingOSPrimaryAccount(
    ICoreWebView2EnvironmentOptions* This, BOOL* allow) {
    if (!allow) return E_POINTER;
    *allow = ((MainEnvironmentOptions*)This)->allowSingleSignOnUsingOSPrimaryAccount;
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE MainEnvironmentOptions_put_AllowSingleSignOnUsingOSPrimaryAccount(
    ICoreWebView2EnvironmentOptions* This, BOOL allow) {
    ((MainEnvironmentOptions*)This)->allowSingleSignOnUsingOSPrimaryAccount = allow;
    return S_OK;
}

static BOOL ValidateMainEnvironmentOptions(ICoreWebView2EnvironmentOptions* options) {
    LPWSTR arguments = NULL;
    LPWSTR language = NULL;
    LPWSTR targetVersion = NULL;
    BOOL allowSingleSignOn = TRUE;

    HRESULT argumentsHr = options->lpVtbl->get_AdditionalBrowserArguments(
        options, &arguments);
    HRESULT languageHr = options->lpVtbl->get_Language(options, &language);
    HRESULT targetHr = options->lpVtbl->get_TargetCompatibleBrowserVersion(
        options, &targetVersion);
    HRESULT singleSignOnHr =
        options->lpVtbl->get_AllowSingleSignOnUsingOSPrimaryAccount(
            options, &allowSingleSignOn);

    BOOL valid = SUCCEEDED(argumentsHr) && arguments && arguments[0] != L'\0' &&
                 SUCCEEDED(languageHr) &&
                 SUCCEEDED(targetHr) && targetVersion && targetVersion[0] != L'\0' &&
                 SUCCEEDED(singleSignOnHr) && !allowSingleSignOn;

    CoTaskMemFree(arguments);
    CoTaskMemFree(language);
    CoTaskMemFree(targetVersion);
    return valid;
}

static ICoreWebView2EnvironmentOptions* CreateMainEnvironmentOptions(
    LPCWSTR additionalBrowserArguments) {
    static ICoreWebView2EnvironmentOptionsVtbl vtbl = {
        MainEnvironmentOptions_QueryInterface,
        MainEnvironmentOptions_AddRef,
        MainEnvironmentOptions_Release,
        MainEnvironmentOptions_get_AdditionalBrowserArguments,
        MainEnvironmentOptions_put_AdditionalBrowserArguments,
        MainEnvironmentOptions_get_Language,
        MainEnvironmentOptions_put_Language,
        MainEnvironmentOptions_get_TargetCompatibleBrowserVersion,
        MainEnvironmentOptions_put_TargetCompatibleBrowserVersion,
        MainEnvironmentOptions_get_AllowSingleSignOnUsingOSPrimaryAccount,
        MainEnvironmentOptions_put_AllowSingleSignOnUsingOSPrimaryAccount
    };

    MainEnvironmentOptions* options =
        (MainEnvironmentOptions*)calloc(1, sizeof(MainEnvironmentOptions));
    if (!options) return NULL;
    options->lpVtbl = &vtbl;
    options->refCount = 1;
    ICoreWebView2EnvironmentOptions* interfaceOptions =
        (ICoreWebView2EnvironmentOptions*)options;

    // A non-null target version is mandatory for a custom options object.
    // Prefer the installed runtime version so this remains compatible with
    // machines that have not yet updated to the bundled SDK's corresponding
    // runtime; fall back to the SDK default if version discovery is unavailable.
    LPWSTR installedBrowserVersion = NULL;
    LPCWSTR targetVersion = WEBVIEW2_TARGET_COMPATIBLE_BROWSER_VERSION;
    if (fnGetAvailableBrowserVersion &&
        SUCCEEDED(fnGetAvailableBrowserVersion(NULL, &installedBrowserVersion)) &&
        installedBrowserVersion && installedBrowserVersion[0] != L'\0') {
        targetVersion = installedBrowserVersion;
    }

    if (FAILED(MainEnvironmentOptions_put_AdditionalBrowserArguments(
            interfaceOptions, additionalBrowserArguments)) ||
        FAILED(MainEnvironmentOptions_put_TargetCompatibleBrowserVersion(
            interfaceOptions, targetVersion)) ||
        !ValidateMainEnvironmentOptions(interfaceOptions)) {
        CoTaskMemFree(installedBrowserVersion);
        MainEnvironmentOptions_Release((ICoreWebView2EnvironmentOptions*)options);
        return NULL;
    }
    CoTaskMemFree(installedBrowserVersion);
    return interfaceOptions;
}

// Create the WebView2 environment for the main window. The controller and
// WebView are then built by the completion handlers (EnvCompletedHandler et
// al). Used from WM_CREATE and when rebuilding after a browser-process death.
static void CreateMainWebViewEnvironment(HWND hwnd) {
    BeginMainNavigationTitle(hwnd, 0);

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

    ICoreWebView2EnvironmentOptions* environmentOptions = NULL;
    size_t insecureOriginCount = 0;
    size_t staticHostCount = 0;
    wchar_t* insecureContentArguments =
        BuildInsecureContentBrowserArguments(&insecureOriginCount);
    wchar_t* staticHostArguments =
        BuildStaticHostBrowserArguments(&staticHostCount);
    wchar_t* browserArguments =
        JoinBrowserArguments(insecureContentArguments, staticHostArguments);
    free(insecureContentArguments);
    free(staticHostArguments);
    if (browserArguments) {
        environmentOptions = CreateMainEnvironmentOptions(browserArguments);
        free(browserArguments);
        if (!environmentOptions) {
            DebugPrint(L"[WARNING] Could not construct valid WebView2 environment options\n");
        } else if (insecureOriginCount > 0) {
            DebugPrint(L"[WARNING] Treating %lu configured HTTP origin(s) as trustworthy\n",
                       (unsigned long)insecureOriginCount);
        }
        if (environmentOptions && staticHostCount > 0) {
            DebugPrint(L"[INFO] Applying %lu static host mapping(s) to the web container\n",
                       (unsigned long)staticHostCount);
        }
    }

    HRESULT hr = fnCreateEnvironment(NULL, userDataPath, environmentOptions,
        (ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*)envHandler);
    if (hr == E_INVALIDARG && environmentOptions) {
        // Keep the launcher usable if a future runtime changes its options
        // contract. The retry uses WebView2's defaults because it omits all
        // configured browser arguments.
        DebugPrint(L"[WARNING] WebView2 rejected environment options; retrying without them\n");
        hr = fnCreateEnvironment(NULL, userDataPath, NULL,
            (ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*)envHandler);
    }
    if (environmentOptions) {
        environmentOptions->lpVtbl->Release(environmentOptions);
    }
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
// each environment: it drives recovery from a browser process that died on
// its own, which used to leave a permanently blank container.
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

// The browser process died without the app asking for it (crash, kill, out
// of memory, runtime servicing) or stopped honoring resume requests. Drop
// every stale COM object and build a fresh WebView.
static void HandleUnexpectedBrowserExit(HWND hwnd) {
    if (InterlockedCompareExchange(&g_webViewCreatePending, TRUE, TRUE) == TRUE) return;

    DebugPrint(L"[WARNING] WebView2 browser gone or unresponsive; rebuilding\n");

    KillTimer(hwnd, ID_TIMER_INITIAL_HIDE_JS);
    KillTimer(hwnd, ID_TIMER_WEBVIEW_PREWARM);
    KillTimer(hwnd, ID_TIMER_WEBVIEW_PRELOAD);
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
    g_lockdownFilterActive = FALSE;  // filters die with the WebView instance

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

    CreateMainWebViewEnvironment(hwnd);
}

// Recovery entry point for the tray actions: if the WebView is gone (rebuild
// limiter tripped, or creation failed earlier) a Refresh/Open builds it anew.
static void RebuildMainWebViewIfDead(void) {
    if (g_webView || g_webViewController || g_webViewEnv) return;
    if (!g_hwnd) return;
    if (InterlockedCompareExchange(&g_webViewCreatePending, TRUE, TRUE) == TRUE) return;

    g_rebuildBurstCount = 0;
    DebugPrint(L"[INFO] Rebuilding missing WebView from tray action\n");
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

    // Keep the environment for the BrowserProcessExited event (crash
    // recovery depends on it).
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

        // Settle the sleep state only once navigations finish so we never
        // suspend a half-loaded page (see OnMainNavigationCompleted and the
        // in-flight guard in DeactivateMainWebView).
        RegisterMainNavigationStartingHandler(webview2);
        RegisterMainNavigationCompletedHandler(webview2);
        RegisterMainNewWindowRequestedHandler(webview2);
        RegisterMainProcessFailedHandler(webview2);
        RegisterMainWebMessageHandler(webview2);

        // Attach the lockdown machinery before the first navigation so the
        // initial page load already carries the header when enabled.
        RegisterMainWebResourceRequestedHandler(webview2);
        g_lockdownFilterActive = FALSE;  // fresh WebView has no filters yet
        ApplyLockdownRequestFilter();

        BOOL mailtoPending =
            InterlockedExchange(&g_mailtoActivationPending, FALSE) == TRUE;
        const wchar_t* initialNavigationUrl = g_initialUrl;
        if (mailtoPending && g_config.handleMailtoLinks &&
            IsValidHttpNavigationUrl(g_config.mailtoTargetUrl)) {
            initialNavigationUrl = g_config.mailtoTargetUrl;
            DebugPrint(L"[INFO] Opening configured email-link destination during WebView startup\n");
        }
        webview2->lpVtbl->Navigate(webview2, initialNavigationUrl);

        InterlockedExchange(&g_resumeFailureCount, 0);
        InterlockedExchange(&g_isInitialized, TRUE);
        InterlockedExchange(&g_webViewDesiredVisible, TRUE);
        ResumeMainWebViewRuntime();

        if (initiallyVisible && InterlockedExchange(&g_resetUrlOnNextShow, FALSE) == TRUE) {
            ResetTargetPageIfNeeded();
        }

        if (initiallyVisible) {
            // A (re)built container under a visible window gets verified like
            // any open. Re-arming resets the counters but not the one-heal-
            // per-open budget, so a rebuild that stays broken cannot loop.
            ArmMainHealthCheck();
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

HRESULT STDMETHODCALLTYPE NavStartingHandler_QueryInterface(
    ICoreWebView2NavigationStartingEventHandler* This,
    REFIID riid, void** ppvObject) {
    if (IsEqualIID(riid, &IID_IUnknown) ||
        IsEqualIID(riid, &IID_ICoreWebView2NavigationStartingEventHandler)) {
        *ppvObject = This;
        This->lpVtbl->AddRef(This);
        return S_OK;
    }
    *ppvObject = NULL;
    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE NavStartingHandler_AddRef(
    ICoreWebView2NavigationStartingEventHandler* This) {
    NavStartingHandler* handler = (NavStartingHandler*)This;
    return InterlockedIncrement(&handler->refCount);
}

ULONG STDMETHODCALLTYPE NavStartingHandler_Release(
    ICoreWebView2NavigationStartingEventHandler* This) {
    NavStartingHandler* handler = (NavStartingHandler*)This;
    ULONG refCount = InterlockedDecrement(&handler->refCount);
    if (refCount == 0) free(handler);
    return refCount;
}

HRESULT STDMETHODCALLTYPE NavStartingHandler_Invoke(
    ICoreWebView2NavigationStartingEventHandler* This,
    ICoreWebView2* sender, ICoreWebView2NavigationStartingEventArgs* args) {
    (void)This;
    (void)sender;

    UINT64 navigationId = 0;
    if (args) args->lpVtbl->get_NavigationId(args, &navigationId);
    DebugPrint(L"[INFO] Main navigation %I64u starting\n", navigationId);
    BeginMainNavigationTitle(g_hwnd, navigationId);
    return S_OK;
}

static void RegisterMainNavigationStartingHandler(ICoreWebView2* webview2) {
    if (!webview2) return;

    NavStartingHandler* handler =
        (NavStartingHandler*)calloc(1, sizeof(NavStartingHandler));
    if (!handler) return;

    static ICoreWebView2NavigationStartingEventHandlerVtbl navVtbl = {
        NavStartingHandler_QueryInterface,
        NavStartingHandler_AddRef,
        NavStartingHandler_Release,
        NavStartingHandler_Invoke
    };
    handler->lpVtbl = &navVtbl;
    handler->refCount = 1;

    EventRegistrationToken token;
    HRESULT hr = webview2->lpVtbl->add_NavigationStarting(
        webview2, (ICoreWebView2NavigationStartingEventHandler*)handler, &token);
    if (FAILED(hr)) {
        DebugPrint(L"[WARNING] add_NavigationStarting failed. HRESULT: 0x%08X\n", hr);
    }

    handler->lpVtbl->Release((ICoreWebView2NavigationStartingEventHandler*)handler);
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

    UINT64 navigationId = 0;
    BOOL navigationIdKnown =
        args && SUCCEEDED(args->lpVtbl->get_NavigationId(args, &navigationId));

    // Read the outcome for the log. Failures land on an error page (or a
    // server error body); the completion still ends the loading title either
    // way, so a settled error is never presented as still loading.
    BOOL isSuccess = TRUE;
    COREWEBVIEW2_WEB_ERROR_STATUS webErrorStatus =
        COREWEBVIEW2_WEB_ERROR_STATUS_UNKNOWN;
    if (args) {
        args->lpVtbl->get_IsSuccess(args, &isSuccess);
        args->lpVtbl->get_WebErrorStatus(args, &webErrorStatus);
    }
    if (isSuccess) {
        DebugPrint(L"[INFO] Main navigation %I64u completed\n", navigationId);
    } else {
        DebugPrint(L"[WARNING] Main navigation %I64u failed (WebErrorStatus %d)\n",
                   navigationId, (int)webErrorStatus);
    }

    if (FinishMainNavigationTitle(g_hwnd, navigationId, navigationIdKnown)) {
        OnMainNavigationCompleted();
    }
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

// Web messages from the main page: the only one the app understands is the
// frame-heartbeat pong posted by the health probe (see SendMainFrameProbe);
// anything else a page happens to post is ignored.
typedef struct {
    ICoreWebView2WebMessageReceivedEventHandlerVtbl* lpVtbl;
    LONG refCount;
} MainMsgHandler;

static HRESULT STDMETHODCALLTYPE MainMsgHandler_QueryInterface(
    ICoreWebView2WebMessageReceivedEventHandler* This,
    REFIID riid, void** ppvObject) {
    if (IsEqualIID(riid, &IID_IUnknown) ||
        IsEqualIID(riid, &IID_ICoreWebView2WebMessageReceivedEventHandler)) {
        *ppvObject = This;
        This->lpVtbl->AddRef(This);
        return S_OK;
    }
    *ppvObject = NULL;
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE MainMsgHandler_AddRef(
    ICoreWebView2WebMessageReceivedEventHandler* This) {
    return InterlockedIncrement(&((MainMsgHandler*)This)->refCount);
}

static ULONG STDMETHODCALLTYPE MainMsgHandler_Release(
    ICoreWebView2WebMessageReceivedEventHandler* This) {
    ULONG refCount = InterlockedDecrement(&((MainMsgHandler*)This)->refCount);
    if (refCount == 0) free(This);
    return refCount;
}

static HRESULT STDMETHODCALLTYPE MainMsgHandler_Invoke(
    ICoreWebView2WebMessageReceivedEventHandler* This,
    ICoreWebView2* sender, ICoreWebView2WebMessageReceivedEventArgs* args) {
    (void)This; (void)sender;

    LPWSTR msg = NULL;
    if (SUCCEEDED(args->lpVtbl->TryGetWebMessageAsString(args, &msg)) && msg) {
        if (wcscmp(msg, L"SystrayLauncher.framePong") == 0) {
            InterlockedExchange(&g_framePongSeen, TRUE);
        }
        CoTaskMemFree(msg);
    }
    return S_OK;
}

static void RegisterMainWebMessageHandler(ICoreWebView2* webview2) {
    if (!webview2) return;

    MainMsgHandler* handler = (MainMsgHandler*)calloc(1, sizeof(MainMsgHandler));
    if (!handler) return;

    static ICoreWebView2WebMessageReceivedEventHandlerVtbl msgVtbl = {
        MainMsgHandler_QueryInterface,
        MainMsgHandler_AddRef,
        MainMsgHandler_Release,
        MainMsgHandler_Invoke
    };
    handler->lpVtbl = &msgVtbl;
    handler->refCount = 1;

    EventRegistrationToken token;
    HRESULT hr = webview2->lpVtbl->add_WebMessageReceived(
        webview2, (ICoreWebView2WebMessageReceivedEventHandler*)handler, &token);
    if (FAILED(hr)) {
        DebugPrint(L"[WARNING] add_WebMessageReceived failed. HRESULT: 0x%08X\n", hr);
    }

    handler->lpVtbl->Release((ICoreWebView2WebMessageReceivedEventHandler*)handler);
}

// --- X-Lockdown request header ---------------------------------------------
//
// Optional gateway token: when enabled, every HTTP(S) request the container
// issues carries an X-Lockdown header holding the request's own User-Agent
// value, AES-256-CBC encrypted under a key derived from the current UTC hour
// and an optional shared secret, then base64 encoded. A gateway derives the
// keys for the previous, current and next hour, tries each, and passes the
// request when a decryption matches the request's User-Agent header — a
// rolling access token that needs no state or clock precision on the client.
// The PHP counterpart is in the README; the layout must match it exactly:
//
//   key          = SHA-256(secret_utf8 + "|" + "YYYY-MM-DD HH:00:00")  (UTC)
//   header value = base64(IV[16] || ciphertext)          (PKCS#7 padding)

#define LOCKDOWN_HEADER_NAME L"X-Lockdown"
#define LOCKDOWN_UA_MAX 1024        // UTF-8 User-Agent bytes incl. NUL
#define LOCKDOWN_VALUE_CCH 1600     // base64(16 + padded UA) + NUL

static BOOL LockdownSha256(const BYTE* data, ULONG dataLen, BYTE hash[32]) {
    BCRYPT_ALG_HANDLE alg = NULL;
    if (!BCRYPT_SUCCESS(BCryptOpenAlgorithmProvider(&alg, BCRYPT_SHA256_ALGORITHM,
                                                    NULL, 0))) {
        return FALSE;
    }

    BOOL ok = FALSE;
    BCRYPT_HASH_HANDLE hashHandle = NULL;
    if (BCRYPT_SUCCESS(BCryptCreateHash(alg, &hashHandle, NULL, 0, NULL, 0, 0))) {
        ok = BCRYPT_SUCCESS(BCryptHashData(hashHandle, (PUCHAR)data, dataLen, 0)) &&
             BCRYPT_SUCCESS(BCryptFinishHash(hashHandle, hash, 32, 0));
        BCryptDestroyHash(hashHandle);
    }
    BCryptCloseAlgorithmProvider(alg, 0);
    return ok;
}

static BOOL LockdownAesCbcEncrypt(const BYTE key[32], const BYTE iv[16],
                                  const BYTE* plain, ULONG plainLen,
                                  BYTE* cipher, ULONG cipherCap, ULONG* cipherLen) {
    BCRYPT_ALG_HANDLE alg = NULL;
    if (!BCRYPT_SUCCESS(BCryptOpenAlgorithmProvider(&alg, BCRYPT_AES_ALGORITHM,
                                                    NULL, 0))) {
        return FALSE;
    }

    BOOL ok = FALSE;
    if (BCRYPT_SUCCESS(BCryptSetProperty(alg, BCRYPT_CHAINING_MODE,
            (PUCHAR)BCRYPT_CHAIN_MODE_CBC, sizeof(BCRYPT_CHAIN_MODE_CBC), 0))) {
        BCRYPT_KEY_HANDLE keyHandle = NULL;
        if (BCRYPT_SUCCESS(BCryptGenerateSymmetricKey(alg, &keyHandle, NULL, 0,
                (PUCHAR)key, 32, 0))) {
            BYTE ivCopy[16];  // BCryptEncrypt advances the IV in place
            memcpy(ivCopy, iv, sizeof(ivCopy));
            ok = BCRYPT_SUCCESS(BCryptEncrypt(keyHandle, (PUCHAR)plain, plainLen,
                    NULL, ivCopy, sizeof(ivCopy), cipher, cipherCap, cipherLen,
                    BCRYPT_BLOCK_PADDING));
            BCryptDestroyKey(keyHandle);
        }
    }
    BCryptCloseAlgorithmProvider(alg, 0);
    return ok;
}

// Build the header value for the given request User-Agent. The result is
// cached per (UTC hour, User-Agent): a page load fires hundreds of
// subresource requests but the value only changes on the hour. Single-thread
// use only — WebView2 raises its events on the UI thread that created it.
static BOOL BuildLockdownHeaderValue(const char* uaUtf8, wchar_t* out, size_t outCch) {
    static char cachedUa[LOCKDOWN_UA_MAX];
    static wchar_t cachedValue[LOCKDOWN_VALUE_CCH];
    static BOOL cacheValid = FALSE;
    static WORD cachedYear, cachedMonth, cachedDay, cachedHour;

    SYSTEMTIME st;
    GetSystemTime(&st);  // UTC by definition

    if (cacheValid &&
        st.wYear == cachedYear && st.wMonth == cachedMonth &&
        st.wDay == cachedDay && st.wHour == cachedHour &&
        strcmp(uaUtf8, cachedUa) == 0) {
        wcscpy_s(out, outCch, cachedValue);
        return TRUE;
    }

    char secretUtf8[768] = "";
    WideCharToMultiByte(CP_UTF8, 0, g_config.lockdownSecret, -1,
                        secretUtf8, sizeof(secretUtf8), NULL, NULL);
    char keyMaterial[832];
    int keyMaterialLen = snprintf(keyMaterial, sizeof(keyMaterial),
        "%s|%04u-%02u-%02u %02u:00:00",
        secretUtf8, st.wYear, st.wMonth, st.wDay, st.wHour);
    if (keyMaterialLen <= 0 || keyMaterialLen >= (int)sizeof(keyMaterial)) {
        return FALSE;
    }

    BYTE key[32];
    if (!LockdownSha256((const BYTE*)keyMaterial, (ULONG)keyMaterialLen, key)) {
        return FALSE;
    }

    BYTE payload[16 + LOCKDOWN_UA_MAX + 16];  // IV || ciphertext (padded)
    if (!BCRYPT_SUCCESS(BCryptGenRandom(NULL, payload, 16,
            BCRYPT_USE_SYSTEM_PREFERRED_RNG))) {
        return FALSE;
    }

    ULONG cipherLen = 0;
    if (!LockdownAesCbcEncrypt(key, payload, (const BYTE*)uaUtf8,
            (ULONG)strlen(uaUtf8), payload + 16,
            (ULONG)(sizeof(payload) - 16), &cipherLen)) {
        return FALSE;
    }

    char base64[LOCKDOWN_VALUE_CCH];
    DWORD base64Len = (DWORD)sizeof(base64);
    if (!CryptBinaryToStringA(payload, 16 + cipherLen,
            CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, base64, &base64Len)) {
        return FALSE;
    }

    // Base64 is plain ASCII; widen it for SetHeader.
    if (MultiByteToWideChar(CP_ACP, 0, base64, -1, out, (int)outCch) == 0) {
        return FALSE;
    }

    cachedYear = st.wYear; cachedMonth = st.wMonth;
    cachedDay = st.wDay; cachedHour = st.wHour;
    strcpy_s(cachedUa, sizeof(cachedUa), uaUtf8);
    wcscpy_s(cachedValue, LOCKDOWN_VALUE_CCH, out);
    cacheValid = TRUE;
    return TRUE;
}

typedef struct {
    ICoreWebView2WebResourceRequestedEventHandlerVtbl* lpVtbl;
    LONG refCount;
} LockdownRequestHandler;

static HRESULT STDMETHODCALLTYPE LockdownRequestHandler_QueryInterface(
    ICoreWebView2WebResourceRequestedEventHandler* This,
    REFIID riid, void** ppvObject) {
    if (IsEqualIID(riid, &IID_IUnknown) ||
        IsEqualIID(riid, &IID_ICoreWebView2WebResourceRequestedEventHandler)) {
        *ppvObject = This;
        This->lpVtbl->AddRef(This);
        return S_OK;
    }
    *ppvObject = NULL;
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE LockdownRequestHandler_AddRef(
    ICoreWebView2WebResourceRequestedEventHandler* This) {
    return InterlockedIncrement(&((LockdownRequestHandler*)This)->refCount);
}

static ULONG STDMETHODCALLTYPE LockdownRequestHandler_Release(
    ICoreWebView2WebResourceRequestedEventHandler* This) {
    ULONG refCount = InterlockedDecrement(&((LockdownRequestHandler*)This)->refCount);
    if (refCount == 0) free(This);
    return refCount;
}

static HRESULT STDMETHODCALLTYPE LockdownRequestHandler_Invoke(
    ICoreWebView2WebResourceRequestedEventHandler* This,
    ICoreWebView2* sender, ICoreWebView2WebResourceRequestedEventArgs* args) {
    (void)This; (void)sender;

    if (InterlockedCompareExchange(&g_lockdownHeader, TRUE, TRUE) != TRUE) {
        return S_OK;
    }

    ICoreWebView2WebResourceRequest* request = NULL;
    if (FAILED(args->lpVtbl->get_Request(args, &request)) || !request) {
        return S_OK;
    }

    ICoreWebView2HttpRequestHeaders* headers = NULL;
    if (SUCCEEDED(request->lpVtbl->get_Headers(request, &headers)) && headers) {
        // Encrypt the exact User-Agent value this request carries; a request
        // without one (rare) gets the empty string, which the verifier
        // compares against its equally absent User-Agent header.
        LPWSTR uaWide = NULL;
        headers->lpVtbl->GetHeader(headers, L"User-Agent", &uaWide);

        char uaUtf8[LOCKDOWN_UA_MAX] = "";
        if (uaWide) {
            if (WideCharToMultiByte(CP_UTF8, 0, uaWide, -1, uaUtf8,
                                    sizeof(uaUtf8), NULL, NULL) == 0) {
                uaUtf8[0] = '\0';
            }
            CoTaskMemFree(uaWide);
        }

        wchar_t value[LOCKDOWN_VALUE_CCH];
        if (BuildLockdownHeaderValue(uaUtf8, value, LOCKDOWN_VALUE_CCH)) {
            headers->lpVtbl->SetHeader(headers, LOCKDOWN_HEADER_NAME, value);
        } else {
            DebugPrint(L"[WARNING] Could not build X-Lockdown header value\n");
        }
        headers->lpVtbl->Release(headers);
    }
    request->lpVtbl->Release(request);
    return S_OK;
}

static void RegisterMainWebResourceRequestedHandler(ICoreWebView2* webview2) {
    if (!webview2) return;

    LockdownRequestHandler* handler =
        (LockdownRequestHandler*)calloc(1, sizeof(LockdownRequestHandler));
    if (!handler) return;

    static ICoreWebView2WebResourceRequestedEventHandlerVtbl requestVtbl = {
        LockdownRequestHandler_QueryInterface,
        LockdownRequestHandler_AddRef,
        LockdownRequestHandler_Release,
        LockdownRequestHandler_Invoke
    };
    handler->lpVtbl = &requestVtbl;
    handler->refCount = 1;

    EventRegistrationToken token;
    HRESULT hr = webview2->lpVtbl->add_WebResourceRequested(
        webview2, (ICoreWebView2WebResourceRequestedEventHandler*)handler, &token);
    if (FAILED(hr)) {
        DebugPrint(L"[WARNING] add_WebResourceRequested failed. HRESULT: 0x%08X\n", hr);
    }

    handler->lpVtbl->Release((ICoreWebView2WebResourceRequestedEventHandler*)handler);
}

// Add or remove the match-everything request filter so it agrees with the
// lockdown setting. The handler itself stays registered either way: only an
// active filter makes requests round-trip through this process, so a
// disabled setting costs nothing, and a config-dialog toggle applies live.
static void ApplyLockdownRequestFilter(void) {
    if (!g_webView) return;

    BOOL want = g_config.lockdownHeader ? TRUE : FALSE;
    if (want == g_lockdownFilterActive) return;

    // Prefer the filter variant that also covers service-worker and shared-
    // worker initiated requests; runtimes without it intercept requests from
    // documents only.
    HRESULT hr = E_NOINTERFACE;
    ICoreWebView2_22* webview22 = NULL;
    if (SUCCEEDED(g_webView->lpVtbl->QueryInterface(g_webView,
            &IID_ICoreWebView2_22, (void**)&webview22)) && webview22) {
        hr = want
            ? webview22->lpVtbl->AddWebResourceRequestedFilterWithRequestSourceKinds(
                  webview22, L"*", COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL,
                  COREWEBVIEW2_WEB_RESOURCE_REQUEST_SOURCE_KINDS_ALL)
            : webview22->lpVtbl->RemoveWebResourceRequestedFilterWithRequestSourceKinds(
                  webview22, L"*", COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL,
                  COREWEBVIEW2_WEB_RESOURCE_REQUEST_SOURCE_KINDS_ALL);
        webview22->lpVtbl->Release(webview22);
    }
    if (FAILED(hr)) {
        hr = want
            ? g_webView->lpVtbl->AddWebResourceRequestedFilter(
                  g_webView, L"*", COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL)
            : g_webView->lpVtbl->RemoveWebResourceRequestedFilter(
                  g_webView, L"*", COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL);
    }

    if (SUCCEEDED(hr)) {
        g_lockdownFilterActive = want;
        DebugPrint(want ? L"[INFO] X-Lockdown header enabled for all requests\n"
                        : L"[INFO] X-Lockdown header disabled\n");
    } else {
        DebugPrint(L"[WARNING] Could not update X-Lockdown request filter. HRESULT: 0x%08X\n", hr);
    }
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
    // (a possibly-broken page must not be frozen into a suspend snapshot),
    // while a tray-hover prewarm is keeping it warm, or while a navigation is
    // still in flight — suspending mid-navigation freezes the load half-done
    // and its completion event may never arrive (the eventual completion
    // re-runs the settle-then-suspend path). The steady-state hidden ticks
    // used to cancel both within 250 ms.
    if (sleepEnabled && preloaded && !recovery && !prewarming &&
        !g_mainNavigationLoading) {
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

    // Hovering delivers a continuous WM_MOUSEMOVE stream. While a prewarm is
    // already active the page is warm; just re-arm the timeout instead of
    // redoing the occlusion scan and cross-process resume/show calls for
    // every mouse move.
    if (InterlockedCompareExchange(&g_webViewPrewarmActive, TRUE, TRUE) == TRUE) {
        SetTimer(g_hwnd, ID_TIMER_WEBVIEW_PREWARM, WEBVIEW_PREWARM_MS, NULL);
        return;
    }

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

// Rebuild the WebView's presentation without touching the page: re-assert
// the bounds (with a one-pixel jiggle that forces the compositor to
// reallocate its surfaces), drop and re-add the visual tree, and refresh the
// parent-position bookkeeping. This is the cheap repair for composition
// surfaces lost across sleep/hibernate.
static void KickMainWebViewComposition(void) {
    if (!g_webViewController || !g_hwnd) return;

    ResumeMainWebViewRuntime();

    RECT bounds;
    GetClientRect(g_hwnd, &bounds);
    if (bounds.bottom - bounds.top > 1) {
        RECT shrunk = bounds;
        shrunk.bottom -= 1;
        g_webViewController->lpVtbl->put_Bounds(g_webViewController, shrunk);
    }
    g_webViewController->lpVtbl->put_Bounds(g_webViewController, bounds);
    g_webViewController->lpVtbl->put_IsVisible(g_webViewController, FALSE);
    g_webViewController->lpVtbl->put_IsVisible(g_webViewController, TRUE);
    g_webViewController->lpVtbl->NotifyParentWindowPositionChanged(g_webViewController);
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
    KickMainWebViewComposition();

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
            DebugPrint(L"[INFO] Reset URL to configured target: %s (was: %s)\n",
                       g_initialUrl, currentUrl);
        } else {
            DebugPrint(L"[INFO] URL already at configured target; skipping navigation\n");
        }
        CoTaskMemFree(currentUrl);
        return;
    }

    g_webView->lpVtbl->Navigate(g_webView, g_initialUrl);
    DebugPrint(L"[INFO] Reset URL to configured target (couldn't check current): %s\n",
               g_initialUrl);
}

// Reset the hidden page to the configured URL without showing anything: wake
// the runtime in case the sleep setting suspended it, navigate, and let
// OnMainNavigationCompleted settle it back to its idle state. The page keeps
// rendering off-screen, so the next open presents the result with no flash.
// When the WebView is not available, defer the reset to the next show instead.
static void ResetTargetPageInBackground(void) {
    if (IsWebViewReady()) {
        ResumeMainWebViewRuntime();
        ResetTargetPageIfNeeded();
    } else {
        InterlockedExchange(&g_resetUrlOnNextShow, TRUE);
    }
}

// --- On-open health verification ------------------------------------------
//
// The hidden-time power-resume recovery above can only prove the runtime
// answers scripts; whether pixels actually reach the screen is unknowable
// until the window is shown. So every show runs this check: for up to a
// second (100 ms ticks, each first confirming the window is still up) the
// page is asked for a frame heartbeat, and a container that cannot produce
// one is torn down and rebuilt. Page CONTENT is deliberately irrelevant -
// a 404 or a blank document heartbeats just as well as the real page.

// Ask the page's compositor for proof of life. requestAnimationFrame only
// fires when the renderer is producing frames for a visible page, so the
// pong (posted back as a web message, see MainMsgHandler_Invoke) covers the
// whole path from script execution to frame production.
static void SendMainFrameProbe(void) {
    ExecuteJavaScript(
        L"requestAnimationFrame(function(){"
        L"try{window.chrome.webview.postMessage('SystrayLauncher.framePong');}catch(e){}"
        L"});");
}

// Begin (or restart) the health poll. Called on every show, and again when a
// rebuilt WebView comes up under a visible window; the caller manages
// g_healthHealed so a rebuild that stays broken cannot loop.
static void ArmMainHealthCheck(void) {
    if (!g_hwnd) return;
    g_healthTicks = 0;
    g_healthTotalTicks = 0;
    InterlockedExchange(&g_framePongSeen, FALSE);
    SetTimer(g_hwnd, ID_TIMER_HEALTH_CHECK, HEALTH_CHECK_INTERVAL_MS, NULL);
}

// Detached-surface detector: after hibernate the renderer can keep producing
// frames (so the heartbeat passes) into a composition surface that is no
// longer attached to the window - the screen just shows a uniform white
// rectangle. Sample a 4x4 interior grid of the client area straight from the
// screen; a fully uniform color is treated as "not actually presenting".
// Only called while this window is foreground, and only until one on-screen
// verification after a power transition has passed, so a legitimately
// uniform page can trigger at most one needless rebuild per resume.
static BOOL IsClientAreaUniformColor(HWND hwnd) {
    RECT rc;
    if (!GetClientRect(hwnd, &rc)) return FALSE;
    int w = rc.right - rc.left, h = rc.bottom - rc.top;
    if (w < 32 || h < 32) return FALSE;

    POINT origin = {0, 0};
    ClientToScreen(hwnd, &origin);

    HDC screen = GetDC(NULL);
    if (!screen) return FALSE;

    COLORREF first = CLR_INVALID;
    BOOL uniform = TRUE;
    for (int iy = 0; iy < 4 && uniform; iy++) {
        for (int ix = 0; ix < 4; ix++) {
            int x = origin.x + w * (2 * ix + 1) / 8;
            int y = origin.y + h * (2 * iy + 1) / 8;
            COLORREF c = GetPixel(screen, x, y);
            if (c == CLR_INVALID) { uniform = FALSE; break; }
            if (first == CLR_INVALID) {
                first = c;
            } else if (c != first) {
                uniform = FALSE;
                break;
            }
        }
    }
    ReleaseDC(NULL, screen);
    return uniform;
}

// Compute the main window's configured rectangle. Maximized windows are
// pre-sized to the work area so the hidden WebView lays out close to its final
// dimensions; Windows applies the exact maximized frame when the window opens.
static void GetTargetWindowRect(int* x, int* y, int* w, int* h) {
    RECT workArea;
    SystemParametersInfoW(SPI_GETWORKAREA, 0, &workArea, 0);

    int workWidth = workArea.right - workArea.left;
    int workHeight = workArea.bottom - workArea.top;

    double scale = g_config.startMaximized ? 1.0 : WINDOW_SIZE_PERCENTAGE;
    int windowWidth = (int)(workWidth * scale);
    int windowHeight = (int)(workHeight * scale);

    *w = windowWidth;
    *h = windowHeight;
    *x = workArea.left + (workWidth - windowWidth) / 2;
    *y = workArea.top + (workHeight - windowHeight) / 2;
}

// Window management
void ShowMainWindow(void) {
    if (!g_hwnd) return;

    // Re-opened before the post-hide reset fired: keep the page exactly as
    // the user left it.
    KillTimer(g_hwnd, ID_TIMER_URL_RESET);

    RebuildMainWebViewIfDead();

    int x, y, windowWidth, windowHeight;
    GetTargetWindowRect(&x, &y, &windowWidth, &windowHeight);

    if (g_config.startMaximized) {
        ShowWindow(g_hwnd, SW_MAXIMIZE);
        SetWindowPos(g_hwnd, HWND_TOP, 0, 0, 0, 0,
                     SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW | SWP_FRAMECHANGED);
    } else {
        // Restore first in case the previous opening was maximized, then
        // enforce the configured centered size before Windows paints again.
        ShowWindow(g_hwnd, SW_RESTORE);
        SetWindowPos(g_hwnd, HWND_TOP, x, y, windowWidth, windowHeight,
                     SWP_SHOWWINDOW | SWP_FRAMECHANGED);
    }
    SetForegroundWindow(g_hwnd);

    ActivateMainWebView();

    // A power transition happened since the container was last verified on
    // screen: refresh the composition up front so a surface lost across
    // sleep/hibernate never gets a chance to present as a white window.
    if (InterlockedCompareExchange(&g_presentationUnverified, TRUE, TRUE) == TRUE) {
        KickMainWebViewComposition();
    }

    if (InterlockedExchange(&g_resetUrlOnNextShow, FALSE) == TRUE) {
        if (IsWebViewReady()) {
            ResetTargetPageIfNeeded();
        } else {
            InterlockedExchange(&g_resetUrlOnNextShow, TRUE);
        }
    }

    // Verify the container actually renders now that it is on screen.
    g_healthHealed = FALSE;
    ArmMainHealthCheck();

    // Start polling for visibility changes while window is shown
    StartVisibilityTimer(g_hwnd);
    UpdateJsVisibilityState(g_hwnd);

    RECT shownRect;
    if (GetWindowRect(g_hwnd, &shownRect)) {
        DebugPrint(L"[INFO] Main window shown at %ldx%ld, size %ldx%ld%s\n",
                   shownRect.left, shownRect.top,
                   shownRect.right - shownRect.left,
                   shownRect.bottom - shownRect.top,
                   g_config.startMaximized ? L" (maximized)" : L"");
    }
}

void HideMainWindow(void) {
    if (!g_hwnd) return;

    // Stop visibility polling when window is hidden
    StopVisibilityTimer(g_hwnd);
    KillTimer(g_hwnd, ID_TIMER_HEALTH_CHECK);

    ShowWindow(g_hwnd, SW_HIDE);
    UpdateJsVisibilityState(g_hwnd);

    // Don't reset the page yet: a quick re-open should land the user exactly
    // where they were. The reset happens in the background once the window
    // has stayed hidden for the grace period (re-armed on every hide).
    SetTimer(g_hwnd, ID_TIMER_URL_RESET, URL_RESET_AFTER_HIDE_MS, NULL);

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

static BOOL ActivateMailtoDestination(void) {
    if (!g_config.handleMailtoLinks ||
        !IsValidHttpNavigationUrl(g_config.mailtoTargetUrl)) {
        DebugPrint(L"[WARNING] Ignored MAILTO activation because email-link handling is not configured\n");
        return FALSE;
    }

    // Preserve the request while the asynchronous WebView creation path is
    // still running. Its controller callback consumes this flag and uses the
    // email destination for the first navigation instead of briefly loading
    // the primary URL first.
    InterlockedExchange(&g_mailtoActivationPending, TRUE);
    ShowMainWindow();

    if (IsWebViewReady()) {
        HRESULT hr = g_webView->lpVtbl->Navigate(
            g_webView, g_config.mailtoTargetUrl);
        if (SUCCEEDED(hr)) {
            InterlockedExchange(&g_mailtoActivationPending, FALSE);
            DebugPrint(L"[INFO] Opened configured email-link destination\n");
        } else {
            DebugPrint(L"[WARNING] Email-link navigation failed. HRESULT: 0x%08X\n",
                       hr);
        }
    }
    return TRUE;
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

    HMENU hMenu = CreatePopupMenu();
    AppendMenuW(hMenu, MF_STRING, ID_TRAY_MENU_OPEN, L"Open");
    AppendMenuW(hMenu, MF_STRING, ID_TRAY_MENU_REFRESH, L"Refresh");
    AppendMenuW(hMenu, MF_STRING, ID_TRAY_MENU_CLEAR_CACHE, L"Refresh + Clear Cache");
    AppendMenuW(hMenu, MF_STRING, ID_TRAY_MENU_RESTART, L"Restart");
    AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenuW(hMenu, MF_STRING, ID_TRAY_MENU_CONFIGURE, L"Configure");
    AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenuW(hMenu, MF_STRING, ID_TRAY_MENU_EXIT, L"Exit");

    SetForegroundWindow(hwnd);
    TrackPopupMenu(hMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, pt.x, pt.y, 0, hwnd, NULL);
    // Documented requirement for tray menus (KB135788): force a task switch
    // after TrackPopupMenu, or the menu will not dismiss when the user
    // clicks outside it.
    PostMessageW(hwnd, WM_NULL, 0, 0);
    DestroyMenu(hMenu);
}

// Window procedure
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
        case WM_CREATE:
            CaptureDisplaySettings();
            CreateMainWebViewEnvironment(hwnd);
            return 0;

        case WM_APP_UPDATE_PROGRESS:
            InterlockedExchange(&g_updateProgressPosted, FALSE);
            if (InterlockedCompareExchange(&g_updateCheckPending,
                                           FALSE, FALSE) == TRUE &&
                g_configViewReady) {
                DWORD speedKbps = (DWORD)InterlockedCompareExchange(
                    &g_updateSpeedKbps, 0, 0);
                CfgSendUpdateProgress(speedKbps);
            }
            return 0;

        case WM_APP_UPDATE_RESULT:
            {
                UpdateCheckTask* task = (UpdateCheckTask*)InterlockedExchangePointer(
                    (PVOID volatile*)&g_updatePostedResult, NULL);
                HandleCompletedUpdateCheck(task);
            }
            return 0;
            
        case WM_SIZE: {
            BOOL restoredFromTaskbar = g_mainWindowMinimized;
            if (wParam == SIZE_MINIMIZED) {
                g_mainWindowMinimized = TRUE;
                StopVisibilityTimer(hwnd);
                KillTimer(hwnd, ID_TIMER_HEALTH_CHECK);
                UpdateJsVisibilityState(hwnd);
                DeactivateMainWebView();
            } else if (IsWindowVisible(hwnd)) {
                g_mainWindowMinimized = FALSE;
                ActivateMainWebView();
                if (restoredFromTaskbar) {
                    // A restore through the taskbar bypasses ShowMainWindow,
                    // so resume the same visibility and health lifecycle here.
                    StartVisibilityTimer(hwnd);
                    UpdateJsVisibilityState(hwnd);
                    g_healthHealed = FALSE;
                    ArmMainHealthCheck();
                    DebugPrint(L"[INFO] Main window restored from taskbar\n");
                }
            } else {
                g_mainWindowMinimized = FALSE;
            }
            return 0;
        }
            
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
                // Initial JS sync. Occlusion polling is only useful while the
                // window is shown (ShowMainWindow starts it): a hidden window
                // cannot become visible on its own, so polling it would burn
                // EnumWindows/DWM/WebView calls 4x per second forever.
                UpdateJsVisibilityState(hwnd);
                if (IsWindowVisible(hwnd)) {
                    StartVisibilityTimer(hwnd);
                }
            } else if (wParam == ID_TIMER_VISIBILITY_CHECK) {
                // Periodic check for window occlusion
                UpdateJsVisibilityState(hwnd);
                // Once the window is withdrawn only ShowMainWindow can bring
                // it back, and that restarts the timer; stop polling until then.
                if (!IsWindowVisible(hwnd)) {
                    StopVisibilityTimer(hwnd);
                }
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
            } else if (wParam == ID_TIMER_URL_RESET) {
                KillTimer(hwnd, ID_TIMER_URL_RESET);
                // The window has stayed hidden past the grace period: put the
                // configured URL back now, while nothing is on screen, so the
                // next open starts fresh with no visible navigation.
                if (!IsWindowVisible(hwnd)) {
                    ResetTargetPageInBackground();
                }
            } else if (wParam == ID_TIMER_HEALTH_CHECK) {
                // On-open health verification. Each tick first confirms the
                // window is still up - closed/minimized/fully-covered windows
                // end the check (a covered page legitimately stops producing
                // frames, so there is nothing to measure).
                if (!IsWindowVisible(hwnd) || IsIconic(hwnd) ||
                    !IsWindowActuallyVisible(hwnd) ||
                    ++g_healthTotalTicks > HEALTH_CHECK_LIFETIME_TICKS) {
                    KillTimer(hwnd, ID_TIMER_HEALTH_CHECK);
                } else if (!IsWebViewReady() || !g_webViewController) {
                    // A rebuild is in flight; absence is not ill health. The
                    // measurement restarts once the new WebView is up.
                    g_healthTicks = 0;
                    InterlockedExchange(&g_framePongSeen, FALSE);
                } else {
                    SendMainFrameProbe();
                    g_healthTicks++;
                    BOOL framesFlowing =
                        InterlockedCompareExchange(&g_framePongSeen, TRUE, TRUE) == TRUE;
                    BOOL unverified =
                        InterlockedCompareExchange(&g_presentationUnverified, TRUE, TRUE) == TRUE;
                    if (framesFlowing && !unverified) {
                        // Frames are flowing and no power transition is in
                        // question - verified, nothing further to prove.
                        KillTimer(hwnd, ID_TIMER_HEALTH_CHECK);
                    } else if (g_healthTicks >= HEALTH_CHECK_VERDICT_TICKS) {
                        BOOL healthy = framesFlowing;
                        if (healthy && unverified) {
                            // Heartbeat passed, but this is the first look
                            // since a power transition: also check that the
                            // frames reach the screen. Only meaningful while
                            // frontmost; otherwise accept the heartbeat and
                            // keep the flag for the next open.
                            if (GetForegroundWindow() == hwnd) {
                                if (IsClientAreaUniformColor(hwnd)) {
                                    healthy = FALSE;
                                    DebugPrint(L"[WARNING] Container heartbeat OK but screen uniform after power resume\n");
                                } else {
                                    InterlockedExchange(&g_presentationUnverified, FALSE);
                                }
                            }
                        }
                        if (healthy) {
                            KillTimer(hwnd, ID_TIMER_HEALTH_CHECK);
                            DebugPrint(L"[INFO] Container verified healthy after open\n");
                        } else if (!g_healthHealed) {
                            // One heal per open: a full rebuild, the only
                            // repair that covers every failure mode seen
                            // after hibernate. The user just opened the
                            // window, so bypass the burst limiter like any
                            // manual tray action.
                            g_healthHealed = TRUE;
                            DebugPrint(L"[WARNING] Container unhealthy %d ms after open; rebuilding\n",
                                       g_healthTicks * HEALTH_CHECK_INTERVAL_MS);
                            g_rebuildBurstStartTick = 0;
                            g_rebuildBurstCount = 0;
                            HandleUnexpectedBrowserExit(hwnd);
                            g_healthTicks = 0;
                            InterlockedExchange(&g_framePongSeen, FALSE);
                        } else {
                            KillTimer(hwnd, ID_TIMER_HEALTH_CHECK);
                            DebugPrint(L"[WARNING] Container still unhealthy after rebuild; waiting for next open\n");
                        }
                    }
                }
            } else if (wParam == ID_TIMER_POWER_RESUME) {
                KillTimer(hwnd, ID_TIMER_POWER_RESUME);
                g_powerKickCount++;
                KickWebViewAfterPowerResume(hwnd);
            } else if (wParam == ID_TIMER_WEBVIEW_LIVENESS) {
                KillTimer(hwnd, ID_TIMER_WEBVIEW_LIVENESS);
                CheckMainWebViewLiveness(hwnd);
            } else if (wParam == ID_TIMER_NAV_TITLE_WATCHDOG) {
                KillTimer(hwnd, ID_TIMER_NAV_TITLE_WATCHDOG);
                if (g_mainNavigationLoading) {
                    DebugPrint(L"[WARNING] Main navigation %I64u never completed; dropping loading title\n",
                               g_mainNavigationId);
                    g_mainNavigationLoading = FALSE;
                    g_mainNavigationId = 0;
                    UpdateMainWindowTitle(hwnd);
                    // The in-flight navigation was also holding the page out
                    // of suspension; let the hidden-state policy settle now.
                    if (!IsWindowActuallyVisible(hwnd)) {
                        DeactivateMainWebView();
                    }
                }
            } else if (wParam == ID_TIMER_AUTO_UPDATE) {
                if (g_config.autoCheckForUpdates) StartUpdateCheck(TRUE);
            }
            return 0;

        case WM_POWERBROADCAST:
            if (wParam == PBT_APMSUSPEND) {
                // Going down: from here on, nothing we believe about the
                // runtime's suspend/resume state can be trusted. The recovery
                // sequence starts when a resume broadcast arrives, and the
                // presentation stays unverified until a SHOWN window passes
                // the on-open health check.
                InterlockedExchange(&g_powerResumePending, TRUE);
                InterlockedExchange(&g_presentationUnverified, TRUE);
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
                InterlockedExchange(&g_presentationUnverified, TRUE);
                g_powerKickCount = 0;
                // Whatever the fallback proxy believed about mapped-address
                // reachability is stale across a suspend; re-probe on the
                // first request instead of waiting out the cooldown.
                HostProxyExpireFallbackCooldowns();
                SetTimer(hwnd, ID_TIMER_POWER_RESUME, POWER_RESUME_KICK_DELAY_MS, NULL);
            }
            return TRUE;
            
        case WM_SYSCOMMAND:
            if ((wParam & 0xFFF0) == SC_MINIMIZE) {
                if (!g_config.showInTaskbar) {
                    HideMainWindow();
                    return 0;
                }
                // Taskbar mode uses normal Windows minimization. DefWindowProc
                // sends WM_SIZE/SIZE_MINIMIZED, which settles the WebView state.
            }
            break;
            
        case WM_CLOSE:
            HideMainWindow();
            return 0;
            
        case WM_DESTROY:
            KillTimer(hwnd, ID_TIMER_WEBVIEW_PREWARM);
            KillTimer(hwnd, ID_TIMER_WEBVIEW_PRELOAD);
            KillTimer(hwnd, ID_TIMER_URL_RESET);
            KillTimer(hwnd, ID_TIMER_HEALTH_CHECK);
            KillTimer(hwnd, ID_TIMER_POWER_RESUME);
            KillTimer(hwnd, ID_TIMER_WEBVIEW_LIVENESS);
            KillTimer(hwnd, ID_TIMER_NAV_TITLE_WATCHDOG);
            KillTimer(hwnd, ID_TIMER_AUTO_UPDATE);
            PostQuitMessage(0);
            return 0;
            
        case WM_APP_WEBVIEW_RECREATE:
            // The browser process died or stopped resuming; rebuild.
            HandleUnexpectedBrowserExit(hwnd);
            return 0;

        case WM_APP_HOST_ROUTE_CHANGED:
            // The fallback proxy switched between the mapped address and
            // DNS resolution for a mapped hostname; refresh the title.
            UpdateMainWindowTitle(hwnd);
            return 0;

        case WM_TRAYICON:
            switch (lParam) {
                case WM_MOUSEMOVE: PrewarmMainWebView(); break;
                case WM_LBUTTONDBLCLK:
                    if (g_config.returnToTargetOnDoubleClick) {
                        // In Home mode, request the configured target even when
                        // the window is already up. ShowMainWindow consumes the
                        // request now or preserves it across a WebView rebuild.
                        InterlockedExchange(&g_resetUrlOnNextShow, TRUE);
                    }
                    ShowMainWindow();
                    break;
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
            if (g_WM_MAILTO_ACTIVATE != 0 &&
                uMsg == g_WM_MAILTO_ACTIVATE) {
                return ActivateMailtoDestination() ? 1 : 0;
            }
            if (uMsg == g_WM_TASKBARCREATED) {
                // Explorer restarted: re-add the icon. RefreshTrayIcon also
                // destroys the old HICON, which a bare CreateTrayIcon leaks.
                RefreshTrayIcon();
                return 0;
            }
    }
    return DefWindowProcW(hwnd, uMsg, wParam, lParam);
}

// Diagnostic output. Debug builds always emit to the debugger. Release
// builds are silent unless the user enables the debug log in the config
// dialog, which appends timestamped lines to
// %LOCALAPPDATA%\SystrayLauncher\debug.log so field incidents (blank
// containers, rebuild storms) can be diagnosed after the fact.
static void GetDebugLogPath(wchar_t path[MAX_PATH]) {
    path[0] = L'\0';
    if (FAILED(SHGetFolderPathW(NULL, CSIDL_LOCAL_APPDATA, NULL, 0, path))) {
        path[0] = L'\0';
        return;
    }
    PathAppendW(path, APP_NAME);
    PathAppendW(path, L"debug.log");
}

static void AppendDebugLogLine(const wchar_t* line) {
    wchar_t path[MAX_PATH];
    GetDebugLogPath(path);
    if (!path[0]) return;

    wchar_t dir[MAX_PATH];
    wcscpy_s(dir, MAX_PATH, path);
    PathRemoveFileSpecW(dir);
    SHCreateDirectoryExW(NULL, dir, NULL);

    // Cap growth: once per process, if the log has passed ~1 MB shift it to
    // debug.old.log (keeping one previous generation) before appending.
    static BOOL rotationChecked = FALSE;
    if (!rotationChecked) {
        rotationChecked = TRUE;
        WIN32_FILE_ATTRIBUTE_DATA fad;
        if (GetFileAttributesExW(path, GetFileExInfoStandard, &fad) &&
            fad.nFileSizeHigh == 0 && fad.nFileSizeLow > 1024 * 1024) {
            wchar_t oldPath[MAX_PATH];
            wcscpy_s(oldPath, MAX_PATH, dir);
            PathAppendW(oldPath, L"debug.old.log");
            MoveFileExW(path, oldPath, MOVEFILE_REPLACE_EXISTING);
        }
    }

    FILE* f = NULL;
    if (_wfopen_s(&f, path, L"a, ccs=UTF-8") != 0 || !f) return;
    SYSTEMTIME st;
    GetLocalTime(&st);
    fwprintf(f, L"[%04u-%02u-%02u %02u:%02u:%02u.%03u] %s",
             st.wYear, st.wMonth, st.wDay,
             st.wHour, st.wMinute, st.wSecond, st.wMilliseconds, line);
    size_t len = wcslen(line);
    if (len == 0 || line[len - 1] != L'\n') {
        fputwc(L'\n', f);
    }
    fclose(f);
}

void DebugPrint(const wchar_t* format, ...) {
    BOOL logEnabled = InterlockedCompareExchange(&g_debugLogEnabled, TRUE, TRUE) == TRUE;
#ifndef _DEBUG
    if (!logEnabled) return;
#endif
    va_list args;
    va_start(args, format);
    wchar_t buffer[4096];
    vswprintf_s(buffer, sizeof(buffer)/sizeof(wchar_t), format, args);
    va_end(args);
#ifdef _DEBUG
    OutputDebugStringW(buffer);
#endif
    if (logEnabled) {
        AppendDebugLogLine(buffer);
    }
}

// Entry point
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    BOOL updateHelperHandled = FALSE;
    BOOL updateCompleted = FALSE;
    BOOL reopenSettings = FALSE;
    int updateHelperResult = HandleUpdateCommandLine(
        &updateHelperHandled, &updateCompleted, &reopenSettings);
    if (updateHelperHandled) return updateHelperResult;

    BOOL mailtoInvocation = IsMailtoProtocolInvocation();
    g_hInstance = hInstance;
    g_WM_MAILTO_ACTIVATE = RegisterWindowMessageW(
        MAILTO_ACTIVATE_MESSAGE_NAME);
    
    // Single instance check. A restarted instance (tray Restart) can arrive
    // while the previous process is still shutting down, so retry briefly
    // before declaring another instance is running.
    g_hMutex = CreateMutexW(NULL, TRUE, MUTEX_NAME);
    DWORD mutexStatus = g_hMutex ? GetLastError() : ERROR_SUCCESS;
    if (mailtoInvocation && g_hMutex &&
        mutexStatus == ERROR_ALREADY_EXISTS) {
        CloseHandle(g_hMutex);
        g_hMutex = NULL;
        if (ForwardMailtoActivationToRunningInstance()) return 0;

        MessageBoxW(
            NULL,
            L"SystrayLauncher is already running, but it could not accept "
            L"this email link. Try the link again or open the launcher from "
            L"the system tray.",
            APP_NAME, MB_OK | MB_ICONWARNING);
        return 1;
    }

    for (int attempt = 0;
         g_hMutex && mutexStatus == ERROR_ALREADY_EXISTS && attempt < 10;
         attempt++) {
        CloseHandle(g_hMutex);
        Sleep(250);
        g_hMutex = CreateMutexW(NULL, TRUE, MUTEX_NAME);
        mutexStatus = g_hMutex ? GetLastError() : ERROR_SUCCESS;
    }
    if (g_hMutex && mutexStatus == ERROR_ALREADY_EXISTS) {
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

    // Winsock is only exercised by the static host fallback proxy, but the
    // init is cheap and unconditional so mapping validation (InetPtonW) and
    // the proxy share one lifetime.
    WSADATA wsaData;
    g_winsockInitialized = (WSAStartup(MAKEWORD(2, 2), &wsaData) == 0);

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
    RefreshWebView2VersionString();

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
    InterlockedExchange(&g_sleepWhenInactive, g_config.sleepWhenInactive ? TRUE : FALSE);
    InterlockedExchange(&g_openNewWindowsExternally, g_config.openNewWindowsExternally ? TRUE : FALSE);
    InterlockedExchange(&g_lockdownHeader, g_config.lockdownHeader ? TRUE : FALSE);
    InterlockedExchange(&g_debugLogEnabled, g_config.debugLogEnabled ? TRUE : FALSE);
    DebugPrint(L"[INFO] SystrayLauncher starting\n");
    if (!SetMailtoHandlerRegistration(g_config.handleMailtoLinks)) {
        DebugPrint(L"[WARNING] Could not synchronize MAILTO handler registration\n");
    }

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

    BOOL activateMailtoOnStartup =
        mailtoInvocation && g_config.handleMailtoLinks &&
        IsValidHttpNavigationUrl(g_config.mailtoTargetUrl);
    if (mailtoInvocation && !activateMailtoOnStartup) {
        MessageBoxW(
            NULL,
            L"Email-link handling is not enabled or does not have a valid "
            L"destination URL. Open SystrayLauncher's configuration to set it up.",
            APP_NAME, MB_OK | MB_ICONINFORMATION);
    }

    // Start the static host fallback proxy before the main window exists:
    // WM_CREATE builds the WebView2 environment, and the PAC URL embedded in
    // its browser arguments needs the proxy's port. The proxy then persists
    // untouched across WebView rebuilds. If it cannot start, the strict
    // resolver rules are emitted instead (see BuildStaticHostBrowserArguments).
    if (g_config.useStaticHostMappings && g_config.staticHostDnsFallback) {
        if (!StartStaticHostProxy()) {
            DebugPrint(L"[WARNING] Falling back to strict static host mappings\n");
        }
    }

    // Register the invisible owner class used when the taskbar button is off.
    WNDCLASSEXW ownerWc = {0};
    ownerWc.cbSize = sizeof(ownerWc);
    ownerWc.lpfnWndProc = DefWindowProcW;
    ownerWc.hInstance = hInstance;
    ownerWc.lpszClassName = L"SystrayLauncherOwner";
    RegisterClassExW(&ownerWc);

    // Create the invisible owner; it remains unused in taskbar-button mode.
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
    
    // An unowned WS_EX_APPWINDOW receives a taskbar button. The default keeps
    // the invisible owner used by earlier releases, which remains tray-only.
    DWORD mainWindowExStyle = g_config.showInTaskbar ? WS_EX_APPWINDOW : 0;
    HWND mainWindowOwner = g_config.showInTaskbar ? NULL : g_hwndOwner;
    g_hwnd = CreateWindowExW(mainWindowExStyle, L"SystrayLauncherClass",
                            g_config.windowTitle,
                            WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT,
                            1024, 768, mainWindowOwner, NULL, hInstance, NULL);
    if (!g_hwnd) {
        MessageBoxW(NULL, L"Failed to create window", L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }
    if (g_WM_MAILTO_ACTIVATE != 0) {
        typedef BOOL (WINAPI *ChangeWindowMessageFilterExFn)(
            HWND, UINT, DWORD, PCHANGEFILTERSTRUCT);
        HMODULE user32 = GetModuleHandleW(L"user32.dll");
        ChangeWindowMessageFilterExFn changeWindowMessageFilterEx =
            user32 ? (ChangeWindowMessageFilterExFn)GetProcAddress(
                         user32, "ChangeWindowMessageFilterEx")
                   : NULL;
        if (changeWindowMessageFilterEx) {
            changeWindowMessageFilterEx(g_hwnd, g_WM_MAILTO_ACTIVATE,
                                        MSGFLT_ALLOW, NULL);
        }
    }
    
    // Register for taskbar restart notifications
    g_WM_TASKBARCREATED = RegisterWindowMessageW(L"TaskbarCreated");
    
    // Create tray icon (loads embedded icon)
    CreateTrayIcon(g_hwnd);

    if (activateMailtoOnStartup) {
        ActivateMailtoDestination();
    }

    // A successful replacement starts exactly once with --finish-update.
    // Reopen settings and show the confirmation only when that update's
    // confirmation requested it; recovery and ordinary launches stay silent.
    if (updateCompleted && reopenSettings) {
        g_updateConfirmationPending = TRUE;
        ShowConfigWebViewDialog();
    }

    SetTimer(g_hwnd, ID_TIMER_AUTO_UPDATE, AUTO_UPDATE_INTERVAL_MS, NULL);
    if (g_config.autoCheckForUpdates) StartUpdateCheck(TRUE);
    
    // Message loop
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    
    // Cleanup. Close() the controller (as the rebuild paths do) so the
    // browser process shuts down and flushes its profile promptly instead of
    // waiting to notice the host process disappear.
    if (g_hwnd) KillTimer(g_hwnd, ID_TIMER_AUTO_UPDATE);
    if (g_updateCancelEvent) SetEvent(g_updateCancelEvent);
    DiscardUpdateTask((UpdateCheckTask*)InterlockedExchangePointer(
        (PVOID volatile*)&g_updatePostedResult, NULL));
    DiscardPendingUpdateNotice();
    DiscardPreparedUpdate();
    if (g_webView) g_webView->lpVtbl->Release(g_webView);
    if (g_webViewController) {
        g_webViewController->lpVtbl->Close(g_webViewController);
        g_webViewController->lpVtbl->Release(g_webViewController);
    }
    if (g_webViewEnv) {
        UnregisterBrowserExitedFromCurrentEnv();
        g_webViewEnv->lpVtbl->Release(g_webViewEnv);
    }

    // The browser process has been told to close, so no new proxy
    // connections are coming; drain the fallback proxy and Winsock last.
    StopStaticHostProxy();
    if (g_winsockInitialized) WSACleanup();

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
