"""Linux regression checks of production updater logic; no Windows app is run.

Run: python3 -m unittest discover -s tests -p test_update.py -v
Requires a host C compiler (cc).
"""
import pathlib
import subprocess
import tempfile
import unittest

ROOT = pathlib.Path(__file__).resolve().parents[1]
SOURCE = (ROOT / 'SystrayLauncher.c').read_text()


def function(name):
    # Find the definition, skipping forward declarations.
    import re
    match = re.search(r'static [^;{}]+\b' + name + r'\([^;{}]*\)\s*\{', SOURCE)
    start = match.start()
    end = match.end()
    depth = 1
    while depth:
        depth += (SOURCE[end] == '{') - (SOURCE[end] == '}')
        end += 1
    return SOURCE[start:end]


class UpdateTests(unittest.TestCase):
    def test_native_speed_and_relaunch(self):
        # Compile the actual functions with process/cleanup APIs mocked. Windows
        # wide printf uses %s where the host libc uses %ls; adapt only that API.
        prelude = r'''
#include <assert.h>
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <wchar.h>
#include <stdarg.h>
typedef int BOOL;
typedef uint32_t DWORD;
typedef uint64_t ULONGLONG;
typedef const wchar_t *LPCWSTR;
typedef wchar_t *LPWSTR;
typedef void *HANDLE;
typedef void *LPVOID;
#define TRUE 1
#define FALSE 0
#define MAXLONG INT32_MAX
#define MAX_PATH 260
#define ERROR_INVALID_PARAMETER 87
#define ERROR_INSUFFICIENT_BUFFER 122
#define CREATE_UNICODE_ENVIRONMENT 1024
#define ZeroMemory(p,n) memset(p,0,n)
#define wcscpy_s(p,n,s) wcscpy(p,s)
typedef struct { int cb; } STARTUPINFOW;
typedef struct { HANDLE hProcess, hThread; } PROCESS_INFORMATION;
static wchar_t captured[2048], argvStorage[2048];
static LPWSTR args[16];
static int argcMock, applied, cleaned, tokenAttempts, plainAttempts;
static BOOL appliedReopen, cleanupValid = TRUE, tokenSucceeds = TRUE;
static BOOL processSucceeds = TRUE;
static int swprintf_s(wchar_t *out, size_t size, const wchar_t *fmt, ...) {
    wchar_t converted[512]; size_t j = 0;
    for (size_t i = 0; fmt[i]; ++i) {
        if (fmt[i] == L'%' && fmt[i+1] == L's') {
            converted[j++] = L'%'; converted[j++] = L'l';
            converted[j++] = L's'; ++i;
        } else converted[j++] = fmt[i];
    }
    converted[j] = 0;
    va_list ap; va_start(ap, fmt);
    int result = vswprintf(out, size, converted, ap);
    va_end(ap); return result;
}
static void SetLastError(DWORD error) { (void)error; }
static void CloseHandle(HANDLE h) { (void)h; }
static BOOL CreateEnvironmentBlock(LPVOID *env, HANDLE token, BOOL inherit) {
    (void)token; (void)inherit; *env = NULL; return FALSE;
}
static void DestroyEnvironmentBlock(LPVOID env) { (void)env; }
static BOOL CreateProcessWithTokenW(HANDLE token, DWORD flags, LPCWSTR target,
    LPWSTR command, DWORD creation, LPVOID env, LPCWSTR cwd,
    STARTUPINFOW *si, PROCESS_INFORMATION *pi) {
    (void)token; (void)flags; (void)target; (void)creation; (void)env;
    (void)cwd; (void)si; (void)pi;
    ++tokenAttempts; wcscpy(captured, command); return tokenSucceeds;
}
static BOOL CreateProcessW(LPCWSTR target, LPWSTR command, void *pa, void *ta,
    BOOL inherit, DWORD flags, LPVOID env, LPCWSTR cwd,
    STARTUPINFOW *si, PROCESS_INFORMATION *pi) {
    (void)target; (void)pa; (void)ta; (void)inherit; (void)flags;
    (void)env; (void)cwd; (void)si; (void)pi;
    ++plainAttempts; wcscpy(captured, command); return processSucceeds;
}
static LPCWSTR GetCommandLineW(void) { return L"mock"; }
static LPWSTR *CommandLineToArgvW(LPCWSTR line, int *count) {
    (void)line; *count = argcMock; return args;
}
static void LocalFree(void *p) { (void)p; }
static int RunUpdateApplyHelper(DWORD pid, LPCWSTR event, LPCWSTR target,
    LPCWSTR staged, BOOL reopen) {
    assert(pid == 10); (void)event; (void)target; (void)staged;
    ++applied; appliedReopen = reopen; return 0;
}
static BOOL FinishUpdateCleanup(DWORD helper, DWORD old, LPCWSTR staged,
    LPCWSTR path) {
    assert(helper == 20 && old == 10);
    assert(wcscmp(staged, L"C:\\Temp Dir\\download.exe") == 0);
    assert(wcscmp(path, L"C:\\Temp Dir\\helper.exe") == 0);
    ++cleaned; return cleanupValid;
}
// Split the generated test commands, including quoted paths with spaces.
static void setCommand(LPCWSTR line) {
    wcscpy(argvStorage, line); argcMock = 0; wchar_t *p = argvStorage;
    while (*p) {
        while (*p == L' ') ++p;
        if (!*p) break;
        if (*p == L'"') {
            args[argcMock++] = ++p;
            while (*p && *p != L'"') ++p;
        } else {
            args[argcMock++] = p;
            while (*p && *p != L' ') ++p;
        }
        if (*p) *p++ = 0;
    }
}
'''
        main = r'''
int main(void) {
    assert(CalculateUpdateSpeedKbps(25600,250) == 100);
    assert(CalculateUpdateSpeedKbps(25727,250) == 100);
    assert(CalculateUpdateSpeedKbps(25728,250) == 101);
    assert(CalculateUpdateSpeedKbps(1,1000) == 0);
    assert(CalculateUpdateSpeedKbps(0,250) == 0);
    assert(CalculateUpdateSpeedKbps(100,0) == 0);
    assert(CalculateUpdateSpeedKbps(102400,1000) == 100);
    assert(CalculateUpdateSpeedKbps(102400,2000) == 50);
    assert(CalculateUpdateSpeedKbps(100ULL*1024*1024,250) == 409600);
    assert(CalculateUpdateSpeedKbps(1ULL<<40,1) == MAXLONG);
    for (int success = 0; success <= 1; ++success) {
        for (int choice = 0; choice <= 1; ++choice) {
            for (int token = 0; token <= 1; ++token) {
                tokenSucceeds = token;
                tokenAttempts = plainAttempts = 0;
                assert(LaunchUpdateTarget(L"C:\\App Dir\\launcher.exe",
                    L"C:\\Temp Dir\\download.exe", L"C:\\Temp Dir\\helper.exe",
                    20, 10, (HANDLE)1, success, choice));
                assert(tokenAttempts == 1 && plainAttempts == !token);
                assert(!!wcsstr(captured,L"--reopen-settings") == (success && choice));
                setCommand(captured);
                BOOL handled = TRUE, completed = TRUE, reopen = TRUE;
                int before = cleaned;
                HandleUpdateCommandLine(&handled,&completed,&reopen);
                assert(!handled && completed == success);
                assert(reopen == (success && choice) && cleaned == before+1);
                cleanupValid = FALSE;
                HandleUpdateCommandLine(&handled,&completed,&reopen);
                assert(!completed && !reopen);
                cleanupValid = TRUE;
            }
        }
    }
    for (int choice = 0; choice <= 1; ++choice) {
        setCommand(choice
            ? L"helper --apply-update 10 event target staged --reopen-settings"
            : L"helper --apply-update 10 event target staged");
        BOOL handled, completed, reopen;
        HandleUpdateCommandLine(&handled,&completed,&reopen);
        assert(handled && !completed && !reopen && appliedReopen == choice);
    }
    assert(applied == 2);
    LPCWSTR unrelated[] = {
        L"app", L"app --reopen-settings", L"app --restart",
        L"app --finish-update 20 10 staged helper --unknown",
        L"app --finish-update 0 10 staged helper --reopen-settings",
        L"app --finish-update 20 bad staged helper --reopen-settings",
        L"app --finish-update 20 10 staged helper --reopen-settings extra"
    };
    for (unsigned i = 0; i < sizeof(unrelated)/sizeof(*unrelated); ++i) {
        BOOL handled = TRUE, completed = TRUE, reopen = TRUE;
        setCommand(unrelated[i]);
        HandleUpdateCommandLine(&handled,&completed,&reopen);
        assert(!handled && !completed && !reopen);
    }
    tokenSucceeds = processSucceeds = FALSE;
    assert(!LaunchUpdateTarget(L"target",L"staged",L"helper",20,10,NULL,TRUE,TRUE));
    puts("Native speed, handoff, token fallback, failure and unrelated-launch checks passed");
}
'''
        names = ['CalculateUpdateSpeedKbps', 'ParseUpdateProcessId',
                 'LaunchUpdateTarget', 'HandleUpdateCommandLine']
        with tempfile.TemporaryDirectory(prefix='systray-update-tests-') as tmp:
            source = pathlib.Path(tmp) / 'test.c'
            source.write_text(prelude + '\n'.join(map(function, names)) + main)
            binary = pathlib.Path(tmp) / 'test'
            subprocess.run(['cc', '-std=c11', '-Wall', '-Wextra', '-Werror',
                            str(source), '-o', str(binary)], check=True)
            subprocess.run([str(binary)], check=True)

    def test_handoff_and_progress_integration(self):
        self.assertIn('InstallPreparedUpdate(json_get_bool(msg, "reopenSettings", FALSE))', SOURCE)
        self.assertIn('LaunchStagedUpdate(task->stagedPath, task->targetPath, reopenSettings)', SOURCE)
        self.assertIn('reopenSettings ? L" --reopen-settings" : L""', function('LaunchStagedUpdate'))
        self.assertIn('launchToken, TRUE,\n                            reopenSettings)', function('RunUpdateApplyHelper'))
        self.assertIn('launchToken, FALSE, FALSE)', function('RestartAfterUpdateFailure'))
        self.assertIn('if (updateCompleted && reopenSettings)', SOURCE)
        download = function('DownloadUpdateFile')
        self.assertIn('GetTickCount64()', download)
        self.assertIn('speedWindowBytes += bytesRead', download)
        self.assertIn('elapsed >= UPDATE_PROGRESS_INTERVAL_MS', download)
        self.assertIn('InterlockedCompareExchange(&g_updateProgressPosted, TRUE, FALSE)', function('PublishUpdateProgress'))


if __name__ == '__main__':
    unittest.main()
