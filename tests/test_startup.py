"""Compile the production Start with Windows logic on Linux with a mocked registry.

Run: python3 -m unittest discover -s tests -p test_startup.py -v
Requires cc. Windows is not run by this test.
"""
import pathlib
import re
import subprocess
import tempfile
import unittest

ROOT = pathlib.Path(__file__).resolve().parents[1]
SOURCE = (ROOT / 'SystrayLauncher.c').read_text()


def function(name):
    match = re.search(r'static [^;{}]+\b' + name + r'\([^;{}]*\)\s*\{', SOURCE)
    end, depth = match.end(), 1
    while depth:
        depth += (SOURCE[end] == '{') - (SOURCE[end] == '}')
        end += 1
    return SOURCE[match.start():end]


def define(name):
    # Include backslash-continued lines.
    return re.search(r'^#define ' + name + r'\b(?:.*\\\n)*.*$', SOURCE, re.M).group(0)


class StartupTests(unittest.TestCase):
    def test_run_entry_state_and_changes(self):
        # The mock asserts the documented Windows key paths, so the production
        # defines are checked rather than copied.
        prelude = r'''
#include <assert.h>
#include <stdarg.h>
#include <stdint.h>
#include <stdio.h>
#include <string.h>
#include <wchar.h>
#include <wctype.h>
typedef int BOOL;
typedef int32_t LONG;
typedef uint32_t DWORD;
typedef unsigned char BYTE;
typedef void *HKEY, *HMODULE;
#define TRUE 1
#define FALSE 0
#define MAX_PATH 260
#define ERROR_SUCCESS 0
#define ERROR_FILE_NOT_FOUND 2
#define ERROR_PATH_NOT_FOUND 3
#define ERROR_ACCESS_DENIED 5
#define ERROR_MORE_DATA 234
#define REG_SZ 1
#define REG_OPTION_NON_VOLATILE 0
#define KEY_SET_VALUE 2
#define RRF_RT_REG_SZ 2
#define RRF_RT_REG_BINARY 8
#define HKEY_CURRENT_USER ((HKEY)1)
#define RUN_KEY ((HKEY)2)
static const wchar_t *RUN = L"Software\\Microsoft\\Windows\\CurrentVersion\\Run";
static const wchar_t *APPROVED =
    L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\StartupApproved\\Run";
static wchar_t modulePath[400] = L"C:\\Apps\\SystrayLauncher.exe";
static wchar_t runValue[400];
static BYTE marker[12];
static DWORD markerSize;
static BOOL runExists, markerExists, approvedKeyExists = TRUE;
static LONG setError, deleteError;
// Windows wide printf uses %s where the host libc uses %ls.
static int swprintf_s(wchar_t *out, size_t size, const wchar_t *fmt, ...) {
    wchar_t converted[64]; size_t j = 0;
    for (size_t i = 0; fmt[i]; ++i) {
        converted[j++] = fmt[i];
        if (fmt[i] == L'%' && fmt[i+1] == L's') converted[j++] = L'l';
    }
    converted[j] = 0;
    va_list ap; va_start(ap, fmt);
    int result = vswprintf(out, size, converted, ap);
    va_end(ap); return result;
}
static int _wcsicmp(const wchar_t *a, const wchar_t *b) {
    while (*a && towlower(*a) == towlower(*b)) { ++a; ++b; }
    return (int)towlower(*a) - (int)towlower(*b);
}
static DWORD GetModuleFileNameW(HMODULE module, wchar_t *path, DWORD size) {
    assert(!module);
    DWORD length = (DWORD)wcslen(modulePath);
    wcsncpy(path, modulePath, size);
    if (length >= size) { path[size - 1] = 0; return size; }  // Truncated.
    return length;
}
static LONG copyValue(const void *value, DWORD needed, void *data, DWORD *size) {
    if (*size < needed) { *size = needed; return ERROR_MORE_DATA; }
    memcpy(data, value, needed); *size = needed; return ERROR_SUCCESS;
}
static LONG RegGetValueW(HKEY root, const wchar_t *subkey, const wchar_t *name,
                         DWORD flags, DWORD *type, void *data, DWORD *size) {
    assert(root == HKEY_CURRENT_USER && !wcscmp(name, L"SystrayLauncher") && !type);
    if (!wcscmp(subkey, RUN)) {
        assert(flags == RRF_RT_REG_SZ);
        if (!runExists) return ERROR_FILE_NOT_FOUND;
        return copyValue(runValue, (DWORD)((wcslen(runValue) + 1) * sizeof(wchar_t)),
                         data, size);
    }
    assert(!wcscmp(subkey, APPROVED) && flags == RRF_RT_REG_BINARY);
    if (!markerExists) return ERROR_FILE_NOT_FOUND;
    return copyValue(marker, markerSize, data, size);
}
static LONG RegCreateKeyExW(HKEY root, const wchar_t *subkey, DWORD reserved,
                            wchar_t *cls, DWORD options, DWORD access,
                            void *security, HKEY *key, DWORD *disposition) {
    (void)reserved; (void)cls; (void)security; (void)disposition;
    assert(root == HKEY_CURRENT_USER && !wcscmp(subkey, RUN));
    assert(options == REG_OPTION_NON_VOLATILE && access == KEY_SET_VALUE);
    *key = RUN_KEY; return ERROR_SUCCESS;
}
static LONG RegSetValueExW(HKEY key, const wchar_t *name, DWORD reserved,
                           DWORD type, const BYTE *data, DWORD size) {
    (void)reserved;
    assert(key == RUN_KEY && !wcscmp(name, L"SystrayLauncher") && type == REG_SZ);
    assert(size == (wcslen((const wchar_t *)data) + 1) * sizeof(wchar_t));
    if (setError) return setError;
    wcscpy(runValue, (const wchar_t *)data); runExists = TRUE;
    return ERROR_SUCCESS;
}
static LONG RegCloseKey(HKEY key) { assert(key == RUN_KEY); return ERROR_SUCCESS; }
static LONG RegDeleteKeyValueW(HKEY root, const wchar_t *subkey, const wchar_t *name) {
    assert(root == HKEY_CURRENT_USER && !wcscmp(name, L"SystrayLauncher"));
    if (deleteError) return deleteError;
    BOOL run = !wcscmp(subkey, RUN);
    assert(run || !wcscmp(subkey, APPROVED));
    if (!run && !approvedKeyExists) return ERROR_PATH_NOT_FOUND;
    BOOL *exists = run ? &runExists : &markerExists;
    if (!*exists) return ERROR_FILE_NOT_FOUND;
    *exists = FALSE; return ERROR_SUCCESS;
}
static void setMarker(BYTE first) {
    memset(marker, 0, sizeof(marker)); marker[0] = first;
    markerSize = sizeof(marker); markerExists = TRUE;
}
'''
        main = r'''
int main(void) {
    const wchar_t *quoted = L"\"C:\\Apps\\SystrayLauncher.exe\"";
    // Off without a Run value, or when it launches something else.
    assert(!IsStartWithWindowsEnabled());
    runExists = TRUE;
    wcscpy(runValue, L"\"C:\\Other\\SystrayLauncher.exe\"");
    assert(!IsStartWithWindowsEnabled());
    wcscpy(runValue, L"C:\\Apps\\SystrayLauncher.exe");
    assert(!IsStartWithWindowsEnabled());
    wmemset(runValue, L'x', 300); runValue[300] = 0;
    assert(!IsStartWithWindowsEnabled());

    // Enabling writes the quoted executable path and clears a disabled marker.
    setMarker(3);
    assert(SetStartWithWindows(TRUE));
    assert(runExists && !wcscmp(runValue, quoted) && !markerExists);
    assert(IsStartWithWindowsEnabled());

    // Paths match case-insensitively; an odd StartupApproved byte disables.
    wcscpy(runValue, L"\"c:\\apps\\SYSTRAYLAUNCHER.exe\"");
    assert(IsStartWithWindowsEnabled());
    for (int first = 0; first < 8; ++first) {
        setMarker((BYTE)first);
        assert(IsStartWithWindowsEnabled() == !(first & 1));
    }
    markerSize = 0;
    assert(IsStartWithWindowsEnabled());

    // Disabling removes both values; nothing left to remove still succeeds.
    setMarker(2);
    assert(SetStartWithWindows(FALSE) && !runExists && !markerExists);
    assert(SetStartWithWindows(FALSE) && !IsStartWithWindowsEnabled());
    approvedKeyExists = FALSE;
    assert(SetStartWithWindows(TRUE) && IsStartWithWindowsEnabled());
    assert(SetStartWithWindows(FALSE) && !runExists);
    approvedKeyExists = TRUE;

    // Registry failures are reported, as is a path too long to register.
    setError = ERROR_ACCESS_DENIED;
    assert(!SetStartWithWindows(TRUE) && !runExists);
    setError = 0;
    assert(SetStartWithWindows(TRUE));
    deleteError = ERROR_ACCESS_DENIED;
    assert(!SetStartWithWindows(FALSE) && runExists);
    deleteError = 0;
    wmemset(modulePath, L'x', MAX_PATH); modulePath[MAX_PATH] = 0;
    assert(!IsStartWithWindowsEnabled());
    runExists = FALSE;
    assert(!SetStartWithWindows(TRUE) && !runExists);
    puts("Start with Windows state, change, and failure checks passed");
}
'''
        defines = ['APP_NAME', 'REG_STARTUP_RUN_PATH', 'REG_STARTUP_APPROVED_RUN_PATH']
        names = ['SetRegistryString', 'DeleteRegistryValueIfPresent',
                 'GetStartupCommand', 'IsStartWithWindowsEnabled', 'SetStartWithWindows']
        with tempfile.TemporaryDirectory(prefix='systray-startup-tests-') as tmp:
            source = pathlib.Path(tmp) / 'test.c'
            source.write_text(prelude + '\n'.join(map(define, defines)) + '\n' +
                              '\n'.join(map(function, names)) + main)
            binary = pathlib.Path(tmp) / 'test'
            subprocess.run(['cc', '-std=c11', '-Wall', '-Wextra', '-Werror',
                            str(source), '-o', str(binary)], check=True)
            subprocess.run([str(binary)], check=True)

    def test_dialog_and_save_wiring(self):
        init = function('webview_push_init_config')
        self.assertIn(r'\"startWithWindows\":%s,\"autoCheckForUpdates\":%s', init)
        self.assertIn('IsStartWithWindowsEnabled() ? L"true" : L"false",\n'
                      '        g_config.autoCheckForUpdates', init)
        save = function('CfgMsgReceived_Invoke')
        # A missing field or an unchanged toggle never touches the Run entry.
        self.assertIn('json_get_bool(msg, "startWithWindows", startWithWindowsEnabled)', save)
        self.assertIn('if (startWithWindows != startWithWindowsEnabled &&\n'
                      '            !SetStartWithWindows(startWithWindows))', save)
        # The failure warning is shown before the dialog is told to close.
        self.assertLess(save.index('could not "\n                      L"be set to start with Windows'),
                        save.index('g_cfgSaved = TRUE;'))


if __name__ == '__main__':
    unittest.main()
