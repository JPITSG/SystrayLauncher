"""Compile the native close gate; the browser suite checks edit decisions.

Run: python3 -m unittest discover -s tests -p test_config_close.py -v
Requires cc. Windows/WebView2 are not run by this test.
"""
from pathlib import Path
import re
import subprocess
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
SOURCE = (ROOT / 'SystrayLauncher.c').read_text()


def function(name):
    match = re.search(r'^static [^;{]*\b' + name + r'\([^;{]*\) \{', SOURCE, re.M)
    assert match, name
    return SOURCE[match.start():SOURCE.index('\n}', match.end()) + 2]


HARNESS = r'''
#include <assert.h>
#include <wchar.h>
typedef int BOOL;
#define TRUE 1
#define FALSE 0
static BOOL g_configViewReady, g_configCloseApproved, g_updateInstallReady;
static void *g_cfgWebView;
static int requests;
static void webview_cfg_execute_script(const wchar_t *script) {
    assert(wcscmp(script, L"window.onCloseRequested()") == 0);
    requests++;
}
/* PRODUCTION */
int main(void) {
    assert(!RequestConfigClose());  // Loading or failed WebView can still close.
    g_cfgWebView = (void *)1;
    assert(!RequestConfigClose());
    g_configViewReady = TRUE;
    assert(RequestConfigClose());
    assert(RequestConfigClose());  // Repeated X does not bypass the prompt.
    assert(requests == 2);
    g_configCloseApproved = TRUE;
    assert(!RequestConfigClose());  // Save, discard, or unchanged configuration.
    g_configCloseApproved = FALSE;
    g_updateInstallReady = TRUE;
    assert(!RequestConfigClose());  // Updater must finish its handoff.
    g_updateInstallReady = FALSE;
    g_cfgWebView = NULL;
    assert(!RequestConfigClose());
    assert(requests == 2);
    return 0;
}
'''


class ConfigCloseTests(unittest.TestCase):
    def test_native_close_gate(self):
        with tempfile.TemporaryDirectory(prefix='systray-config-close-') as directory:
            path = Path(directory)
            (path / 'test.c').write_text(HARNESS.replace(
                '/* PRODUCTION */', function('RequestConfigClose')))
            subprocess.run(['cc', '-std=c11', '-Wall', '-Wextra', '-Werror',
                            str(path / 'test.c'), '-o', str(path / 'test')], check=True)
            subprocess.run([str(path / 'test')], check=True, timeout=10)

    def test_save_discard_and_window_lifecycle_wiring(self):
        proc = function('CfgWndProc')
        close = proc.split('case WM_CLOSE:', 1)[1].split('case WM_DESTROY:', 1)[0]
        gate = close.index('if (RequestConfigClose()) return 0;')
        self.assertLess(gate, close.index('->Close('))
        self.assertLess(gate, close.index('DestroyWindow('))
        self.assertIn('g_configCloseApproved = FALSE;', proc.split('case WM_DESTROY:', 1)[1])
        self.assertIn('g_configCloseApproved = FALSE;', function('ShowConfigWebViewDialog'))
        handler = function('CfgMsgReceived_Invoke')
        save = handler.split('strcmp(action, "saveSettings")', 1)[1].split('strcmp(action, "close")', 1)[0]
        failed_save = save.split('if (!SaveConfigToRegistry(&g_config)) {', 1)[1].split('\n        }', 1)[0]
        self.assertIn('return S_OK;', failed_save)
        self.assertNotIn('g_configCloseApproved = TRUE;', failed_save)
        self.assertLess(save.index('g_cfgSaved = TRUE;'), save.index('g_configCloseApproved = TRUE;'))
        self.assertLess(save.index('g_configCloseApproved = TRUE;'), save.index('PostMessage(g_cfgHwnd, WM_CLOSE'))
        close = handler.split('strcmp(action, "close")', 1)[1].split('strcmp(action, "resize")', 1)[0]
        self.assertLess(close.index('g_configCloseApproved = TRUE;'), close.index('PostMessage(g_cfgHwnd, WM_CLOSE'))
        self.assertNotIn('g_cfgSaved = TRUE;', close)


if __name__ == '__main__':
    unittest.main()
