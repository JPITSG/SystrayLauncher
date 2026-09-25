/* Deterministic platform boundary for the extracted production C functions. */
#include <assert.h>
#include <stdarg.h>
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <wchar.h>

typedef int BOOL, LONG, HRESULT;
typedef unsigned ULONG, UINT, DWORD;
typedef uint64_t UINT64, ULONGLONG;
typedef intptr_t LPARAM, LONG_PTR;
typedef uintptr_t WPARAM;
typedef wchar_t *LPWSTR;
typedef const wchar_t *LPCWSTR;
typedef struct { int left, top, right, bottom; } RECT;
typedef struct Window {
    BOOL visible, minimized, child, cloaked, dwmFails, destroyed, shaped;
    DWORD pid;
    LONG_PTR style;
    RECT rect, frame;
} Window, *HWND;
typedef struct { unsigned char pixels[32][32]; } Region, *HRGN;
typedef struct Hook *HWINEVENTHOOK;
typedef void (*WinEventProc)(HWINEVENTHOOK, DWORD, HWND, LONG, LONG, DWORD, DWORD);
typedef struct Hook { BOOL active; DWORD first, last, pid; WinEventProc fn; } Hook;
#define TRUE 1
#define FALSE 0
#define NULLREGION 1
#define SIMPLEREGION 2
#define COMPLEXREGION 3
#define ERROR 0
#define S_OK 0
#define SUCCEEDED(hr) ((hr) >= 0)
#define FAILED(hr) ((hr) < 0)
#define CALLBACK
#define STDMETHODCALLTYPE
#define WINAPI
typedef void VOID, *PVOID;
typedef struct MIB_IPFORWARD_ROW2 *PMIB_IPFORWARD_ROW2;
typedef enum { MibParameterNotification, MibAddInstance, MibDeleteInstance,
               MibInitialNotification } MIB_NOTIFICATION_TYPE;
#define WM_APP 0x8000
#define OBJID_WINDOW 0
#define CHILDID_SELF 0
#define GA_ROOT 2
#define GWL_EXSTYLE (-20)
#define WS_EX_LAYERED 1
#define WS_EX_TRANSPARENT 2
#define WS_EX_NOREDIRECTIONBITMAP 4
#define DWMWA_CLOAKED 14
#define DWMWA_EXTENDED_FRAME_BOUNDS 9
#define MONITOR_DEFAULTTONULL 0
#define RGN_DIFF 4
#define WINEVENT_OUTOFCONTEXT 0
#define EVENT_SYSTEM_FOREGROUND 3
#define EVENT_SYSTEM_CAPTUREEND 9
#define EVENT_SYSTEM_MOVESIZESTART 10
#define EVENT_SYSTEM_MOVESIZEEND 11
#define EVENT_SYSTEM_MINIMIZESTART 22
#define EVENT_SYSTEM_MINIMIZEEND 23
#define EVENT_SYSTEM_DESKTOPSWITCH 32
#define EVENT_OBJECT_DESTROY 0x8001
#define EVENT_OBJECT_REORDER 0x8004
#define EVENT_OBJECT_LOCATIONCHANGE 0x800b
#define EVENT_OBJECT_STATECHANGE 0x800a
#define EVENT_OBJECT_CLOAKED 0x8017
#define EVENT_OBJECT_UNCLOAKED 0x8018
#define DebugPrint(...) ((void)0)

static Window windows[80];
static HWND zorder[80], g_hwnd = &windows[0], foreground;
static size_t windowCount;
static BOOL enumerationFails, regionFails, combineFails, monitorMissing, timerFails;
static unsigned scans, liveRegions, timerSets, hookSets, hookRemoves;
static unsigned timers[32], posts, lastPost;
static Hook hooks[512];
static int failHookAt;
static ULONGLONG ticks = 1000;
static struct { wchar_t onShowJs[64], onHideJs[64]; } g_config;
static wchar_t g_initialUrl[2048] = L"https://example.com/";
typedef enum { JS_VISIBILITY_UNKNOWN = -1, JS_VISIBILITY_HIDDEN,
               JS_VISIBILITY_SHOWN } JsVisibility;
static JsVisibility g_jsVisibility = JS_VISIBILITY_UNKNOWN;
static HWINEVENTHOOK g_visibilityHooks[8];
static struct { DWORD processId; HWINEVENTHOOK hook; } g_visibilityLocationHooks[64];
static size_t g_visibilityLocationHookCount;

static LONG InterlockedExchange(volatile LONG *p, LONG v) { LONG old = *p; *p = v; return old; }
static LONG InterlockedCompareExchange(volatile LONG *p, LONG v, LONG expected) {
    LONG old = *p; if (old == expected) *p = v; return old;
}
static LONG InterlockedIncrement(volatile LONG *p) { return ++*p; }
static LONG InterlockedDecrement(volatile LONG *p) { return --*p; }
static BOOL IsWindowVisible(HWND w) { return w && w->visible && !w->destroyed; }
static BOOL IsIconic(HWND w) { return w && w->minimized; }
static HWND GetForegroundWindow(void) { return foreground; }
static HWND GetAncestor(HWND w, UINT mode) { return w->child || w->destroyed ? NULL : w; }
static LONG_PTR GetWindowLongPtrW(HWND w, int index) { return w->style; }
static DWORD GetWindowThreadProcessId(HWND w, DWORD *pid) { *pid = w->pid; return !w->destroyed; }
static HRESULT DwmGetWindowAttribute(HWND w, DWORD attr, void *out, size_t size) {
    if (w->dwmFails) return -1;
    if (attr == DWMWA_CLOAKED) *(DWORD *)out = w->cloaked;
    else *(RECT *)out = w->frame;
    return S_OK;
}
static BOOL GetWindowRect(HWND w, RECT *r) { *r = w->rect; return !w->destroyed; }
static BOOL GetClientRect(HWND w, RECT *r) { *r = (RECT){0,0,20,20}; return TRUE; }
static BOOL EqualRect(const RECT *a, const RECT *b) { return memcmp(a,b,sizeof(*a)) == 0; }
static BOOL IsRectEmpty(const RECT *r) { return r->right <= r->left || r->bottom <= r->top; }
static void *MonitorFromRect(const RECT *r, DWORD flags) { return monitorMissing ? NULL : (void *)1; }
static BOOL IntersectRect(RECT *r, const RECT *a, const RECT *b) {
    *r = (RECT){a->left > b->left ? a->left : b->left,
                a->top > b->top ? a->top : b->top,
                a->right < b->right ? a->right : b->right,
                a->bottom < b->bottom ? a->bottom : b->bottom};
    return !IsRectEmpty(r);
}
static BOOL SetRectRgn(HRGN r, int left, int top, int right, int bottom) {
    for (int y=0;y<32;++y) for (int x=0;x<32;++x)
        r->pixels[y][x] = x >= left && x < right && y >= top && y < bottom;
    return TRUE;
}
static HRGN CreateRectRgn(int l, int t, int r, int b) {
    if (regionFails) return NULL;
    HRGN region = calloc(1, sizeof(*region)); assert(region); ++liveRegions;
    SetRectRgn(region,l,t,r,b); return region;
}
static HRGN CreateRectRgnIndirect(const RECT *r) { return CreateRectRgn(r->left,r->top,r->right,r->bottom); }
static void DeleteObject(HRGN r) { assert(r && liveRegions); --liveRegions; free(r); }
static int GetRgnBox(HRGN r, RECT *out) {
    for (int y=0;y<32;++y) for (int x=0;x<32;++x)
        if (r->pixels[y][x]) return COMPLEXREGION;
    return NULLREGION;
}
static int CombineRgn(HRGN out, HRGN a, HRGN b, int op) {
    if (combineFails) { memset(out,0,sizeof(*out)); return ERROR; }
    assert(op == RGN_DIFF);
    for (int y=0;y<32;++y) for (int x=0;x<32;++x) out->pixels[y][x] = a->pixels[y][x] && !b->pixels[y][x];
    RECT dummy; return GetRgnBox(out,&dummy);
}
static int GetWindowRgn(HWND w, HRGN r) { return w->shaped ? COMPLEXREGION : ERROR; }
static BOOL EnumWindows(BOOL (*fn)(HWND, LPARAM), LPARAM arg) {
    ++scans; if (enumerationFails) return FALSE;
    for (size_t i=0;i<windowCount;++i) if (!fn(zorder[i],arg)) return FALSE;
    return TRUE;
}
static HWINEVENTHOOK SetWinEventHook(DWORD first, DWORD last, void *module,
        WinEventProc fn, DWORD pid, DWORD thread, DWORD flags) {
    assert(first != EVENT_OBJECT_LOCATIONCHANGE || pid != 0); // No global movement stream.
    ++hookSets;
    if (failHookAt && (int)hookSets == failHookAt) return NULL;
    for (size_t i=0;i<512;++i) if (!hooks[i].active) {
        hooks[i] = (Hook){TRUE,first,last,pid,fn}; return &hooks[i];
    }
    assert(0); return NULL;
}
static BOOL UnhookWinEvent(HWINEVENTHOOK hook) { assert(hook->active); hook->active=FALSE; ++hookRemoves; return TRUE; }
static UINT SetTimer(HWND hwnd, UINT id, UINT delay, void *proc) {
    ++timerSets; if (timerFails) return 0; timers[id]=delay; return id;
}
static BOOL KillTimer(HWND hwnd, UINT id) { timers[id]=0; return TRUE; }
static BOOL PostMessageW(HWND hwnd, UINT message, WPARAM w, LPARAM l) { ++posts; lastPost=message; return TRUE; }
static ULONGLONG GetTickCount64(void) { return ticks; }
static void CoTaskMemFree(void *p) { free(p); }
static int swprintf_s(wchar_t *out, size_t n, const wchar_t *format, ...) {
    // Production's Windows format has a different uint64 length modifier.
    wchar_t adjusted[512]; size_t j=0;
    for (size_t i=0;format[i];++i) {
        if (wcsncmp(format+i,L"%I64u",5)==0) {
            wmemcpy(adjusted+j,L"%llu",4); j+=4; i+=4;
        } else adjusted[j++]=format[i];
    }
    adjusted[j]=0; va_list args; va_start(args,format);
    int result=vswprintf(out,n,adjusted,args); va_end(args); return result;
}

typedef struct ICoreWebView2TrySuspendCompletedHandler ICoreWebView2TrySuspendCompletedHandler;
typedef struct {
    HRESULT (*QueryInterface)(ICoreWebView2TrySuspendCompletedHandler *, const void *, void **);
    ULONG (*AddRef)(ICoreWebView2TrySuspendCompletedHandler *);
    ULONG (*Release)(ICoreWebView2TrySuspendCompletedHandler *);
    HRESULT (*Invoke)(ICoreWebView2TrySuspendCompletedHandler *, HRESULT, BOOL);
} ICoreWebView2TrySuspendCompletedHandlerVtbl;
struct ICoreWebView2TrySuspendCompletedHandler { ICoreWebView2TrySuspendCompletedHandlerVtbl *lpVtbl; };
typedef struct ICoreWebView2ExecuteScriptCompletedHandlerVtbl { void *unused; } ICoreWebView2ExecuteScriptCompletedHandlerVtbl;
typedef struct { ICoreWebView2ExecuteScriptCompletedHandlerVtbl *lpVtbl; } ICoreWebView2ExecuteScriptCompletedHandler;
typedef struct ICoreWebView2_3 ICoreWebView2_3;
typedef struct {
    HRESULT (*Resume)(ICoreWebView2_3 *);
    ULONG (*Release)(ICoreWebView2_3 *);
    HRESULT (*TrySuspend)(ICoreWebView2_3 *, ICoreWebView2TrySuspendCompletedHandler *);
} ICoreWebView2_3Vtbl;
struct ICoreWebView2_3 { ICoreWebView2_3Vtbl *lpVtbl; };
typedef struct ICoreWebView2Controller ICoreWebView2Controller;
typedef struct {
    HRESULT (*put_Bounds)(ICoreWebView2Controller *, RECT);
    HRESULT (*put_IsVisible)(ICoreWebView2Controller *, BOOL);
} ControllerVtbl;
struct ICoreWebView2Controller { ControllerVtbl *lpVtbl; };
typedef struct ICoreWebView2 ICoreWebView2;
typedef struct {
    HRESULT (*QueryInterface)(ICoreWebView2 *, const void *, void **);
    HRESULT (*get_Source)(ICoreWebView2 *, LPWSTR *);
    HRESULT (*Navigate)(ICoreWebView2 *, LPCWSTR);
} WebViewVtbl;
struct ICoreWebView2 { WebViewVtbl *lpVtbl; };
static const int IID_ICoreWebView2_3;
static ICoreWebView2 *g_webView;
static ICoreWebView2Controller *g_webViewController;
static unsigned resumes, suspends, shows, boundsSets, navigations, scripts, kicks, rebuilds;
static BOOL runtimeSuspended, rendered=TRUE, resumeFails, visibilityFails, sourceFails, uniformPage;
static wchar_t currentUrl[2048]=L"https://example.com/", lastScript[512];
static ICoreWebView2TrySuspendCompletedHandler *pending[16];
static size_t pendingCount;
static HRESULT TrySuspendCompletedHandler_QueryInterface(ICoreWebView2TrySuspendCompletedHandler *h,
        const void *iid, void **out) { *out=h; return S_OK; }
static HRESULT mock_resume(ICoreWebView2_3 *v) { ++resumes; if (resumeFails) return -1; runtimeSuspended=FALSE; return S_OK; }
static ULONG mock_release(ICoreWebView2_3 *v) { return 1; }
static HRESULT mock_suspend(ICoreWebView2_3 *v, ICoreWebView2TrySuspendCompletedHandler *h) {
    assert(!rendered); assert(pendingCount<16); ++suspends;
    h->lpVtbl->AddRef(h); pending[pendingCount++]=h; return S_OK;
}
static ICoreWebView2_3Vtbl runtimeVtbl={mock_resume,mock_release,mock_suspend};
static ICoreWebView2_3 runtime={&runtimeVtbl};
static HRESULT mock_query(ICoreWebView2 *v, const void *iid, void **out) { *out=&runtime; return S_OK; }
static HRESULT mock_source(ICoreWebView2 *v, LPWSTR *out) {
    if (sourceFails) { *out=NULL; return -1; }
    *out=malloc((wcslen(currentUrl)+1)*sizeof(wchar_t)); wcscpy(*out,currentUrl); return S_OK;
}
static HRESULT mock_navigate(ICoreWebView2 *v, LPCWSTR url) {
    assert(rendered && !runtimeSuspended); ++navigations; wcscpy(currentUrl,url); return S_OK;
}
static HRESULT mock_bounds(ICoreWebView2Controller *v, RECT rect) { ++boundsSets; return S_OK; }
static HRESULT mock_visible(ICoreWebView2Controller *v, BOOL value) {
    ++shows; if (visibilityFails) return -1; rendered=value;
    if (value) runtimeSuspended=FALSE;
    return S_OK;
}
static WebViewVtbl webVtbl={mock_query,mock_source,mock_navigate};
static unsigned targetReloads, cooldownExpiries;
static void ReloadTargetPage(void) { ++targetReloads; }
static void HostProxyExpireFallbackCooldowns(void) { ++cooldownExpiries; }
typedef enum {
    COREWEBVIEW2_WEB_ERROR_STATUS_UNKNOWN,
    COREWEBVIEW2_WEB_ERROR_STATUS_CERTIFICATE_COMMON_NAME_IS_INCORRECT,
    COREWEBVIEW2_WEB_ERROR_STATUS_CERTIFICATE_EXPIRED,
    COREWEBVIEW2_WEB_ERROR_STATUS_CLIENT_CERTIFICATE_CONTAINS_ERRORS,
    COREWEBVIEW2_WEB_ERROR_STATUS_CERTIFICATE_REVOKED,
    COREWEBVIEW2_WEB_ERROR_STATUS_CERTIFICATE_IS_INVALID,
    COREWEBVIEW2_WEB_ERROR_STATUS_SERVER_UNREACHABLE,
    COREWEBVIEW2_WEB_ERROR_STATUS_TIMEOUT,
    COREWEBVIEW2_WEB_ERROR_STATUS_ERROR_HTTP_INVALID_SERVER_RESPONSE,
    COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_ABORTED,
    COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_RESET,
    COREWEBVIEW2_WEB_ERROR_STATUS_DISCONNECTED,
    COREWEBVIEW2_WEB_ERROR_STATUS_CANNOT_CONNECT,
    COREWEBVIEW2_WEB_ERROR_STATUS_HOST_NAME_NOT_RESOLVED,
    COREWEBVIEW2_WEB_ERROR_STATUS_OPERATION_CANCELED,
    COREWEBVIEW2_WEB_ERROR_STATUS_REDIRECT_FAILED,
    COREWEBVIEW2_WEB_ERROR_STATUS_UNEXPECTED_ERROR,
    COREWEBVIEW2_WEB_ERROR_STATUS_VALID_AUTHENTICATION_CREDENTIALS_REQUIRED,
    COREWEBVIEW2_WEB_ERROR_STATUS_VALID_PROXY_AUTHENTICATION_REQUIRED
} COREWEBVIEW2_WEB_ERROR_STATUS;
static ICoreWebView2 web={&webVtbl};
static ControllerVtbl controllerVtbl={mock_bounds,mock_visible};
static ICoreWebView2Controller controller={&controllerVtbl};
static void ExecuteJavaScript(LPCWSTR script) { ++scripts; wcscpy(lastScript,script); }
static void BeginMainNavigationTitle(HWND hwnd, UINT64 navigationId);
static void KickMainWebViewComposition(void);
static BOOL IsClientAreaUniformColor(HWND hwnd) { return uniformPage; }
static void HandleUnexpectedBrowserExit(HWND hwnd) { ++rebuilds; }
