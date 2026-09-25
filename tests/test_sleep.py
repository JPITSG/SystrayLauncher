"""Run production visibility/sleep functions with deterministic Win32/COM fakes.

These cover state transitions and resource use, not Windows compositor behavior.
Run: python3 -m unittest discover -s tests -p test_sleep.py -v
"""
import pathlib
import re
import subprocess
import tempfile
import unittest

ROOT = pathlib.Path(__file__).resolve().parents[1]
SOURCE = (ROOT / 'SystrayLauncher.c').read_text()


def block(start):
    end = SOURCE.index('{', start) + 1
    depth = 1
    while depth:
        depth += (SOURCE[end] == '{') - (SOURCE[end] == '}')
        end += 1
    return SOURCE[start:end]


def function(name):
    match = re.search(r'^(?:static )?[\w *]+\b' + name +
                      r'\([^;{}]*\)\s*\{', SOURCE, re.M)
    if not match:
        raise ValueError(name)
    return block(match.start())


class SleepTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.directory = tempfile.TemporaryDirectory()
        cls.addClassCleanup(cls.directory.cleanup)
        path = pathlib.Path(cls.directory.name)
        cls.executable = path / 'sleep'
        names = [
            'IsWebViewReady', 'QueryMainWebView3', 'SyncMainWebViewBounds',
            'SetMainWebViewControllerVisible', 'ResumeMainWebViewRuntime',
            'SuspendMainWebViewRuntime', 'TrySuspendCompletedHandler_AddRef',
            'TrySuspendCompletedHandler_Release', 'TrySuspendCompletedHandler_Invoke',
            'LivenessPingHandler_Invoke', 'ActivateMainWebView', 'DeactivateMainWebView',
            'PrewarmMainWebView', 'PrepareMainWebViewNavigation',
            'OnMainNavigationCompleted', 'VisibilityWinEventProc',
            'ClearVisibilityLocationHooks', 'SyncVisibilityLocationHooks',
            'OcclusionEnumProc', 'IsWindowActuallyVisible', 'QueueVisibilityCheck',
            'StartVisibilityTracking', 'StopVisibilityTracking',
            'UpdateJsVisibilityState', 'ResetTargetPageIfNeeded',
            'ResetTargetPageInBackground', 'SendMainFrameProbe',
            'StopMainHealthCheck', 'ArmMainHealthCheck', 'CheckMainWebViewHealth',
            'RetryFailedMainLoad', 'IsConnectivityLoadFailure', 'OnNetworkRouteChange',
        ]
        definitions = [function(name) for name in names]
        declarations = '\n'.join(f[:f.index('{')].strip() + ';' for f in definitions)
        types = '\n'.join(re.search(
            r'typedef struct \{[^{}]*\} ' + name + ';', SOURCE).group()
            for name in ['TrySuspendCompletedHandler', 'LivenessPingHandler', 'OcclusionCheckData'])
        constants = '\n'.join(line for line in SOURCE.splitlines()
                              if re.match(r'#define (ID_TIMER_|VISIBILITY_|WEBVIEW_PRE|'
                                          r'WEBVIEW_RESUME|RESUME_FAILURE|MAX_VISIBILITY|'
                                          r'HEALTH_CHECK_TIMEOUT|WM_APP_WEBVIEW_RECREATE|'
                                          r'WM_APP_VISIBILITY_WAKE|WM_APP_NETWORK_CHANGED|'
                                          r'NETWORK_CHANGE_SETTLE)', line))
        globals_ = '\n'.join(line for line in SOURCE.splitlines()
                             if re.match(r'static (?:volatile LONG|UINT64|ULONGLONG|BOOL|int|UINT|RECT) '
                                         r'g_(?:webView(?:Desired|Suspend|Generation|Settle)|'
                                         r'isInitialized|sleepWhenInactive|initialPreloadComplete|'
                                         r'resumeFailureCount|powerResumePending|presentationUnverified|'
                                         r'framePongSeen|frameProbeId|health|controller|'
                                         r'visibility(?:Tracking|HooksAvailable|CheckDelay)|'
                                         r'suspendRequestId|livenessRequestId|prewarmLastHoverTick|preloadSettleDeadline|'
                                         r'mainNavigationLoading|mainLoadFailed|resetUrlOnNextShow|'
                                         r'webViewPrewarmActive|webViewPingOutstanding|'
                                         r'networkChangePosted)', line))
        timer_ids = ['ID_TIMER_VISIBILITY_CHECK', 'ID_TIMER_WEBVIEW_PREWARM',
                     'ID_TIMER_WEBVIEW_PRELOAD', 'ID_TIMER_WEBVIEW_RESUME_RETRY',
                     'ID_TIMER_HEALTH_CHECK', 'ID_TIMER_NETWORK_SETTLE']
        branches = []
        for timer in timer_ids:
            start = SOURCE.index('if (wParam == ' + timer + ')', SOURCE.index('LRESULT CALLBACK WindowProc'))
            branches.append(block(start))
        timers = 'static int fire_timer(UINT wParam) { HWND hwnd = g_hwnd;\n' + \
            ' else '.join(branches) + '\nreturn 0; }\n'
        label = 'case WM_APP_NETWORK_CHANGED:'
        start = SOURCE.index(label, SOURCE.index('LRESULT CALLBACK WindowProc')) + len(label)
        timers += ('static int network_message(void) { HWND hwnd = g_hwnd;\n' +
                   SOURCE[start:SOURCE.index('return 0;', start)] + 'return 0; }\n')
        harness = '\n'.join([
            (ROOT / 'tests/sleep_stubs.h').read_text(), constants, types, globals_,
            declarations, '\n'.join(definitions), timers,
            (ROOT / 'tests/sleep_cases.c').read_text(),
        ])
        source = path / 'sleep.c'
        source.write_text(harness)
        subprocess.run(['cc', '-std=c11', '-Wall', '-Wextra',
                        '-Werror', '-Wno-unused-function', '-Wno-unused-parameter',
                        '-Wno-unused-variable', str(source), '-o', str(cls.executable)], check=True)

    def run_case(self, name):
        subprocess.run([str(self.executable), name], check=True)

    def test_partial_coverage_and_uncover_events(self):
        self.run_case('coverage')

    def test_transparency_shapes_and_uncertain_geometry(self):
        self.run_case('geometry')

    def test_no_idle_polling_and_coalesced_filtered_events(self):
        self.run_case('events')

    def test_hooks_disabled_hidden_minimized_and_feature_off(self):
        self.run_case('lifecycle')

    def test_stale_suspend_and_liveness_completions(self):
        self.run_case('callbacks')

    def test_navigation_settle_and_background_reset(self):
        self.run_case('navigation')

    def test_prewarm_and_resume_failure_bounds(self):
        self.run_case('recovery')

    def test_health_checks_wait_for_load_and_preserve_plain_pages(self):
        self.run_case('health')

    def test_failed_loads_retry_on_network_change_and_return_only(self):
        self.run_case('retry')

    def test_route_notifications_are_released_and_load_outcomes_recorded(self):
        main = SOURCE[SOURCE.index('int WINAPI WinMain('):]
        registered = main.index('NotifyRouteChange2(AF_UNSPEC, OnNetworkRouteChange')
        self.assertLess(registered, main.index('while (GetMessage(&msg'))
        self.assertGreater(main.index('CancelMibChangeNotify2(g_networkChangeHandle)'),
                           main.index('while (GetMessage(&msg'))
        window_proc = SOURCE.index('LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, '
                                   'WPARAM wParam, LPARAM lParam) {')
        destroy = SOURCE.index('case WM_DESTROY:', window_proc)
        self.assertIn('KillTimer(hwnd, ID_TIMER_NETWORK_SETTLE);',
                      SOURCE[destroy:SOURCE.index('return 0;', destroy)])
        completed = function('NavCompletedHandler_Invoke')
        self.assertRegex(completed, r'webErrorStatus != COREWEBVIEW2_WEB_ERROR_STATUS_OPERATION_CANCELED\)'
                                    r'\s*\{\s*g_mainLoadFailed = !isSuccess &&\s*'
                                    r'IsConnectivityLoadFailure\(webErrorStatus, httpStatusCode\);')
        self.assertIn('g_mainLoadFailed = FALSE;', function('HandleUnexpectedBrowserExit'))
