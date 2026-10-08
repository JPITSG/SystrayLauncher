"""Run the production new-window handler with mocked WebView2 and shell calls.

Run: python3 -m unittest discover -s tests -p test_new_windows.py -v
Requires cc. Windows/WebView2 are not run by this test.
"""
from pathlib import Path
import re
import subprocess
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
SOURCE = (ROOT / 'SystrayLauncher.c').read_text()

HARNESS = r'''
#define _POSIX_C_SOURCE 200809L
#include <assert.h>
#include <stdint.h>
#include <stdlib.h>
#include <string.h>
#include <wchar.h>
typedef int BOOL;
typedef int32_t HRESULT, LONG;
typedef wchar_t* LPWSTR;
typedef void ICoreWebView2, ICoreWebView2NewWindowRequestedEventHandler;
#define TRUE 1
#define FALSE 0
#define S_OK 0
#define E_FAIL ((HRESULT)-1)
#define FAILED(hr) ((hr) < 0)
#define STDMETHODCALLTYPE
#define SW_SHOWNORMAL 1
#define DebugPrint(...) ((void)0)
#define _wcsnicmp wcsncasecmp

typedef struct ICoreWebView2WindowFeatures ICoreWebView2WindowFeatures;
typedef struct {
    unsigned long (*Release)(ICoreWebView2WindowFeatures*);
    HRESULT (*get_ShouldDisplayToolbar)(ICoreWebView2WindowFeatures*, BOOL*);
} FeaturesVtbl;
struct ICoreWebView2WindowFeatures { FeaturesVtbl* lpVtbl; };
typedef struct ICoreWebView2NewWindowRequestedEventArgs ICoreWebView2NewWindowRequestedEventArgs;
typedef struct {
    HRESULT (*get_WindowFeatures)(ICoreWebView2NewWindowRequestedEventArgs*, ICoreWebView2WindowFeatures**);
    HRESULT (*get_Uri)(ICoreWebView2NewWindowRequestedEventArgs*, LPWSTR*);
    HRESULT (*put_Handled)(ICoreWebView2NewWindowRequestedEventArgs*, BOOL);
} ArgsVtbl;
struct ICoreWebView2NewWindowRequestedEventArgs { ArgsVtbl* lpVtbl; };

static volatile LONG g_openNewWindowsExternally;
static BOOL toolbar, missingFeatures, handled;
static HRESULT featuresResult, toolbarResult, uriResult;
static const wchar_t* uri;
static wchar_t launchedUri[512];
static LPWSTR allocatedUri;
static int launches, featureRefs, uriAllocations, handledCalls;

static LONG InterlockedCompareExchange(volatile LONG* target, LONG value, LONG expected) {
    LONG old = *target;
    if (old == expected) *target = value;
    return old;
}
static unsigned long ReleaseFeatures(ICoreWebView2WindowFeatures* This) {
    (void)This;
    assert(featureRefs == 1);
    return --featureRefs;
}
static HRESULT GetToolbar(ICoreWebView2WindowFeatures* This, BOOL* value) {
    (void)This;
    assert(featureRefs == 1);
    *value = toolbar;
    return toolbarResult;
}
static FeaturesVtbl featuresVtbl = {ReleaseFeatures, GetToolbar};
static ICoreWebView2WindowFeatures features = {&featuresVtbl};

static HRESULT GetFeatures(ICoreWebView2NewWindowRequestedEventArgs* This,
                           ICoreWebView2WindowFeatures** value) {
    (void)This;
    assert(featureRefs == 0);
    *value = missingFeatures ? NULL : &features;
    if (*value) ++featureRefs;
    return featuresResult;
}
static HRESULT GetUri(ICoreWebView2NewWindowRequestedEventArgs* This, LPWSTR* value) {
    (void)This;
    *value = NULL;
    if (FAILED(uriResult) || !uri) return uriResult;
    assert(uriAllocations == 0);
    size_t bytes = (wcslen(uri) + 1) * sizeof(wchar_t);
    allocatedUri = malloc(bytes);
    assert(allocatedUri);
    memcpy(allocatedUri, uri, bytes);
    *value = allocatedUri;
    ++uriAllocations;
    return S_OK;
}
static HRESULT PutHandled(ICoreWebView2NewWindowRequestedEventArgs* This, BOOL value) {
    (void)This;
    handled = value;
    ++handledCalls;
    return S_OK;
}
static ArgsVtbl argsVtbl = {GetFeatures, GetUri, PutHandled};
static ICoreWebView2NewWindowRequestedEventArgs args = {&argsVtbl};

static void CoTaskMemFree(void* memory) {
    assert(uriAllocations == 1 && memory == allocatedUri);
    free(memory);
    allocatedUri = NULL;
    --uriAllocations;
}
static void* ShellExecuteW(void* owner, const wchar_t* action, const wchar_t* file,
                          const wchar_t* parameters, const wchar_t* directory, int show) {
    assert(!owner && !parameters && !directory && show == SW_SHOWNORMAL);
    assert(wcscmp(action, L"open") == 0 && file && wcslen(file) < 512);
    wcscpy(launchedUri, file);
    ++launches;
    return (void*)33;
}

/* PRODUCTION */

static void reset(BOOL external, BOOL popup, const wchar_t* url) {
    assert(featureRefs == 0 && uriAllocations == 0);
    g_openNewWindowsExternally = external;
    toolbar = !popup;
    missingFeatures = FALSE;
    featuresResult = toolbarResult = uriResult = S_OK;
    uri = url;
    launches = handledCalls = 0;
    handled = FALSE;
    launchedUri[0] = 0;
}
static void expect_route(BOOL external) {
    assert(NewWindowHandler_Invoke(NULL, NULL, &args) == S_OK);
    assert(launches == external);
    assert(handled == external && handledCalls == external);
    if (external) assert(wcscmp(launchedUri, uri) == 0);
    assert(featureRefs == 0 && uriAllocations == 0);
}

int main(int argc, char** argv) {
    assert(argc == 2);
    if (strcmp(argv[1], "popups") == 0) {
        // WebView2 reports the same popup classification for explicit popup
        // features and conventional sized popups, including target _blank.
        const wchar_t* urls[] = {L"https://example.com/login", L"http://example.com/help", L"about:blank"};
        for (int enabled = 0; enabled <= 1; ++enabled) {
            for (size_t i = 0; i < sizeof(urls) / sizeof(urls[0]); ++i) {
                reset(enabled, TRUE, urls[i]);
                expect_route(FALSE);
            }
        }
    } else if (strcmp(argv[1], "links") == 0) {
        const wchar_t* urls[] = {L"https://example.com/page?q=1#section",
            L"http://example.com/page", L"HTTPS://example.com/page", L"HTTP://example.com/page"};
        // A live toggle takes effect on the next request without rebuilding.
        for (int enabled = 0; enabled <= 1; ++enabled) {
            for (size_t i = 0; i < sizeof(urls) / sizeof(urls[0]); ++i) {
                reset(enabled, FALSE, urls[i]);
                expect_route(enabled);
            }
        }
        reset(FALSE, FALSE, urls[0]);
        expect_route(FALSE);
    } else if (strcmp(argv[1], "schemes") == 0) {
        const wchar_t* urls[] = {L"about:blank", L"", NULL,
            L"mailto:person@example.com", L"file:///C:/example.txt",
            L"javascript:alert(1)", L"data:text/html,hello", L"custom-app:launch",
            L"httpx://example.com", L"https:example.com"};
        for (size_t i = 0; i < sizeof(urls) / sizeof(urls[0]); ++i) {
            reset(TRUE, FALSE, urls[i]);
            expect_route(FALSE);
        }
    } else if (strcmp(argv[1], "errors") == 0) {
        reset(TRUE, FALSE, L"https://example.com/login");
        featuresResult = E_FAIL;
        missingFeatures = TRUE;
        expect_route(FALSE);

        reset(TRUE, FALSE, L"https://example.com/login");
        missingFeatures = TRUE; // Successful getter with no object.
        expect_route(FALSE);

        reset(TRUE, FALSE, L"https://example.com/login");
        featuresResult = E_FAIL; // Any returned reference must be released.
        expect_route(FALSE);

        reset(TRUE, FALSE, L"https://example.com/login");
        toolbarResult = E_FAIL; // Ignore the output value on failure.
        expect_route(FALSE);

        reset(TRUE, FALSE, L"https://example.com/login");
        uriResult = E_FAIL;
        expect_route(FALSE);
    } else {
        return 1;
    }
    return 0;
}
'''


class NewWindowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        match = re.search(
            r'^static HRESULT STDMETHODCALLTYPE NewWindowHandler_Invoke\([^;]*?\)\s*\{.*?^}',
            SOURCE, re.M | re.S)
        if not match:
            raise AssertionError('Missing production NewWindowHandler_Invoke')
        cls.directory = tempfile.TemporaryDirectory(prefix='systraylauncher-windows-')
        cls.addClassCleanup(cls.directory.cleanup)
        path = Path(cls.directory.name)
        source = path / 'windows.c'
        cls.exe = path / 'windows'
        source.write_text(HARNESS.replace('/* PRODUCTION */', match.group()))
        subprocess.run(['cc', '-std=c11', '-Wall', '-Wextra', '-Werror',
                        str(source), '-o', str(cls.exe)], check=True)

    def test_popups_stay_in_webview_with_either_checkbox_state(self):
        subprocess.run([str(self.exe), 'popups'], check=True, timeout=10)

    def test_ordinary_links_follow_the_checkbox(self):
        subprocess.run([str(self.exe), 'links'], check=True, timeout=10)

    def test_only_http_and_https_can_open_externally(self):
        subprocess.run([str(self.exe), 'schemes'], check=True, timeout=10)

    def test_missing_metadata_stays_in_webview(self):
        subprocess.run([str(self.exe), 'errors'], check=True, timeout=10)


if __name__ == '__main__':
    unittest.main()
