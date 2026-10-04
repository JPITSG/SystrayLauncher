"""Run the modal-loop pointer fix against a simulated ShowCursor counter.

Run: python3 -m unittest discover -s tests -p test_modal_cursor.py -v
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
    # Match the definition, not a forward declaration of the same name.
    match = re.search(r'^(?:static )?[^\n;]*\b' + name + r'\([^;]*?\)\s*\{',
                      SOURCE, re.MULTILINE)
    assert match, name
    depth, end = 1, match.end()
    while depth:
        depth += (SOURCE[end] == '{') - (SOURCE[end] == '}')
        end += 1
    return SOURCE[match.start():end]


HARNESS = r'''
#include <assert.h>
typedef int BOOL;
#define TRUE 1
#define FALSE 0
#define SM_MOUSEPRESENT 19
static int counter, calls, mousePresent = 1;
static int ShowCursor(BOOL show) { ++calls; return show ? ++counter : --counter; }
static int GetSystemMetrics(int index) { assert(index == SM_MOUSEPRESENT); return mousePresent; }
/* PRODUCTION */
static void run(int start) { counter = start; calls = 0; }
int main(void) {
    // Chromium hid the pointer (-1): the loop shows it, the end gives it back.
    run(-1); ModalLoopCursorBegin(); assert(counter == 0 && calls == 1);
    ModalLoopCursorBegin(); assert(counter == 0 && calls == 1);  // No double raise.
    ModalLoopCursorEnd(); assert(counter == -1 && calls == 2);
    ModalLoopCursorEnd(); assert(counter == -1 && calls == 2);  // Nothing owed.

    // Several outstanding hides are all lifted, then all restored.
    run(-3); ModalLoopCursorBegin(); assert(counter == 0 && calls == 3);
    ModalLoopCursorEnd(); assert(counter == -3);

    // Already visible: the probe is undone at once and nothing is owed.
    run(0); ModalLoopCursorBegin(); assert(counter == 0 && calls == 2);
    ModalLoopCursorEnd(); assert(counter == 0 && calls == 2);

    // Chromium shows the pointer itself during the loop: still balanced.
    run(-1); ModalLoopCursorBegin(); ++counter; ModalLoopCursorEnd();
    assert(counter == 0);

    // A runaway counter is raised a bounded number of times.
    run(-100); ModalLoopCursorBegin(); assert(counter == -84 && calls == 16);
    ModalLoopCursorEnd(); assert(counter == -100);

    // No mouse: the counter rests at -1 by design and is left alone.
    mousePresent = 0; run(-1); ModalLoopCursorBegin(); ModalLoopCursorEnd();
    assert(counter == -1 && calls == 0);
    return 0;
}
'''


class ModalCursorTests(unittest.TestCase):
    def test_counter_is_raised_for_the_loop_and_given_back(self):
        production = '\n'.join([
            re.search(r'^#define MODAL_CURSOR_MAX_RAISES .*$', SOURCE, re.M).group(0),
            re.search(r'^static int g_modalCursorRaises\b.*$', SOURCE, re.M).group(0),
            function('ModalLoopCursorBegin'), function('ModalLoopCursorEnd')])
        with tempfile.TemporaryDirectory(prefix='systraylauncher-cursor-') as directory:
            path = Path(directory)
            (path / 'cursor.c').write_text(HARNESS.replace('/* PRODUCTION */', production))
            subprocess.run(['cc', '-std=c11', '-Wall', '-Wextra', '-Werror',
                            str(path / 'cursor.c'), '-o', str(path / 'cursor')], check=True)
            subprocess.run([str(path / 'cursor')], check=True, timeout=10)

    def test_both_windows_wrap_their_modal_loops(self):
        # Entering and leaving pass on to DefWindowProcW (break, not return).
        wiring = (r'case WM_ENTERMENULOOP:\s*case WM_ENTERSIZEMOVE:\s*'
                  r'ModalLoopCursorBegin\(\);\s*break;\s*'
                  r'case WM_EXITMENULOOP:\s*case WM_EXITSIZEMOVE:\s*'
                  r'ModalLoopCursorEnd\(\);\s*break;')
        for name in ('WindowProc', 'CfgWndProc'):
            proc = function(name)
            self.assertRegex(proc, wiring, name)
            self.assertRegex(proc, r'return DefWindowProcW\([^;]*\);\s*\}$', name)
            # Before the switch only the existing interceptions run.
            self.assertNotIn('ModalLoopCursor', proc.split('switch (', 1)[0], name)


if __name__ == '__main__':
    unittest.main()
