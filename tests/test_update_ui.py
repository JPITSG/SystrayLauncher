"""Exercise the built configuration UI with a mocked WebView2 message bridge.

Run after make: python3 -m unittest discover -s tests -p test_update_ui.py -v
Requires Python Playwright and Chromium. Set CHROMIUM_EXECUTABLE to override
Playwright's browser, and PLAYWRIGHT_NODEJS_PATH to override its Node runtime.
"""
import os
import re
from pathlib import Path
import unittest

from playwright.sync_api import sync_playwright, expect

ROOT = Path(__file__).resolve().parents[1]

FULL_CONFIG = {
    'url': 'https://example.com', 'windowTitle': 'Test',
    'startMaximized': False, 'returnToTargetOnDoubleClick': True,
    'showInTaskbar': False, 'handleMailtoLinks': True,
    'mailtoTargetUrl': 'https://example.com/mail',
    'onHideJs': 'pause();', 'onShowJs': 'resume();',
    'sleepWhenInactive': False, 'openNewWindowsExternally': False,
    'allowRunningInsecureContent': True,
    'insecureContentOrigins': 'http://example.com,http://other.test',
    'useStaticHostMappings': True,
    'staticHostMappings': 'example.com:127.0.0.1,example.com:127.0.0.2',
    'staticHostDnsFallback': False, 'lockdownHeader': True,
    'lockdownSecret': 'test-secret', 'startWithWindows': False,
    'autoCheckForUpdates': False, 'debugLog': False,
    'updateCheckPending': False, 'updatePromptPending': False,
}


class UpdateUiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.playwright = sync_playwright().start()
        options = {'headless': True}
        if os.environ.get('CHROMIUM_EXECUTABLE'):
            options['executable_path'] = os.environ['CHROMIUM_EXECUTABLE']
        cls.browser = cls.playwright.chromium.launch(**options)

    @classmethod
    def tearDownClass(cls):
        cls.browser.close()
        cls.playwright.stop()

    def setUp(self):
        self.page = self.browser.new_page(viewport={'width': 900, 'height': 1000})
        self.errors = []
        self.page.on('pageerror', lambda error: self.errors.append(str(error)))
        self.page.add_init_script('''
            window.messages = [];
            window.chrome = { webview: { postMessage(message) {
                const data = JSON.parse(message);
                window.messages.push(data);
                if (data.action === 'getInit') setTimeout(() => window.onInit({
                    config: { url: 'https://example.com', windowTitle: 'Test',
                        onHideJs: '', onShowJs: '', autoCheckForUpdates: false },
                    webView2Version: 'test', updateCompletedVersion: ''
                }), 0);
            } } };
        ''')
        self.page.goto((ROOT / 'assets/dist/index.html').as_uri())
        expect(self.page.get_by_role('button', name='Update', exact=True)).to_be_visible()

    def tearDown(self):
        self.page.close()
        self.assertEqual(self.errors, [])

    def reset_config(self, config=None):
        if config is None:
            self.page.reload()
        else:
            self.page.evaluate('window.onInit(null)')
            expect(self.page.locator('#windowTitle')).to_have_count(0)
            self.page.evaluate('''config => {
                window.messages = [];
                window.onInit({config, webView2Version: 'test', updateCompletedVersion: ''});
            }''', config)
        self.page.wait_for_function("window.messages.some(m => m.action === 'configReady')")

    def close_actions(self):
        return self.page.evaluate(
            "window.messages.filter(m => ['close', 'saveSettings'].includes(m.action))")

    def request_native_close(self):
        self.page.evaluate('window.onCloseRequested()')

    def expect_close_prompt(self):
        dialog = self.page.get_by_role('alertdialog', name='Unsaved changes')
        expect(dialog).to_be_visible()
        expect(dialog.get_by_text('Save changes before closing?', exact=True)).to_be_visible()
        expect(self.page.get_by_role('alertdialog')).to_have_count(1)
        self.assertTrue(self.page.locator('#windowTitle').evaluate("el => !!el.closest('[inert]')"))
        self.assertEqual(self.close_actions(), [])
        return dialog

    def result(self, status='newer', automatic=False):
        self.page.evaluate('result => window.onUpdateResult(result)', {
            'status': status, 'automatic': automatic, 'title': 'Update result',
            'message': 'Result of checking for an update.',
            'currentVersion': '1.0.23', 'remoteVersion': '1.0.24',
        })

    def last_message(self, action):
        return self.page.evaluate(
            'action => window.messages.filter(m => m.action === action).at(-1)', action)

    def test_progress_style_and_cancellation(self):
        self.page.get_by_role('button', name='Update', exact=True).click()
        button = self.page.get_by_role('button', name='Stop update check and download')
        expect(button).to_have_text('Checking...')
        expect(button).to_have_class(re.compile('bg-red-600'))
        for percent, expected in [(0, 0), (7, 7), (42.9, 42), (100, 100),
                                  (150, 100), (-5, 0)]:
            self.page.evaluate('percent => window.onUpdateProgress({percent})', percent)
            expect(button).to_have_text(f'Checking ({expected}%)...')
            expect(button).to_be_enabled()  # Clicking again stops the download.
        button.click()
        expect(button).to_have_text('Stopping...')
        expect(button).to_be_disabled()
        self.assertIsNotNone(self.last_message('cancelUpdateCheck'))
        self.result('cancelled')
        expect(self.page.get_by_role('alertdialog')).to_have_count(0)
        self.page.get_by_role('button', name='Update', exact=True).click()
        expect(button).to_have_text('Checking...')

    def test_per_confirmation_choice(self):
        for status in ['newer', 'same']:
            for choice in [False, True]:
                self.result(status)
                dialog = self.page.get_by_role('alertdialog')
                checkbox = dialog.get_by_label('Reopen settings after update', exact=True)
                expect(checkbox).not_to_be_checked()
                expect(dialog.get_by_text('1.0.23', exact=True)).to_be_visible()
                expect(dialog.get_by_text('1.0.24', exact=True)).to_be_visible()
                if choice:
                    checkbox.check()
                dialog.get_by_role('button', name='Update' if status == 'newer' else 'Force update', exact=True).click()
                self.assertEqual(self.last_message('installUpdate'), {
                    'action': 'installUpdate', 'reopenSettings': choice})
                expect(checkbox).to_be_disabled()
                expect(dialog.get_by_role('button', name='Starting...')).to_be_disabled()
                self.result('error')  # UAC cancellation or helper failure.
                expect(dialog.get_by_label('Reopen settings after update')).to_have_count(0)
                dialog.get_by_role('button', name='OK', exact=True).click()

    def test_dismiss_ignore_and_downgrade(self):
        for automatic in [False, True]:
            self.result(automatic=automatic)
            dialog = self.page.get_by_role('alertdialog')
            checkbox = dialog.get_by_label('Reopen settings after update')
            expect(checkbox).not_to_be_checked()
            checkbox.check()
            dialog.get_by_role('button', name='Ignore this version' if automatic else 'Cancel', exact=True).click()
            expect(dialog).to_have_count(0)
            self.assertIsNone(self.last_message('installUpdate'))
        self.result()
        expect(self.page.get_by_label('Reopen settings after update')).not_to_be_checked()
        self.result('older')
        dialog = self.page.get_by_role('alertdialog')
        expect(dialog.get_by_label('Reopen settings after update')).to_have_count(0)
        expect(dialog.get_by_role('button', name='Update', exact=True)).to_have_count(0)
        expect(dialog.get_by_role('button', name='Force update', exact=True)).to_have_count(0)

    def test_static_hosts_order_and_save_reload(self):
        self.page.get_by_label('Resolve listed hostnames to static IP addresses', exact=True).check()
        mappings = self.page.get_by_label('Static host mappings', exact=True)
        mappings.fill('DOMAIN.com:1.2.3.4\nother.test:127.0.0.1\n'
                      'domain.com:3.4.5.6,domain.com:1.2.3.4\n'
                      'domain.com:[2001:0DB8:0:0::1]\ndomain.com:[2001:db8::1]')
        fallback = self.page.get_by_label(
            'Fall back to standard DNS when all mapped addresses are unreachable', exact=True)
        expect(fallback).not_to_be_checked()
        expected = ('domain.com:1.2.3.4,other.test:127.0.0.1,'
                    'domain.com:3.4.5.6,domain.com:[2001:db8::1]')
        for checked in [True, False]:
            fallback.set_checked(checked)
            self.page.get_by_role('button', name='Save', exact=True).click()
            saved = self.last_message('saveSettings')
            self.assertIsNotNone(saved)
            self.assertEqual(saved['staticHostMappings'], expected)
            self.assertEqual(saved['staticHostDnsFallback'], checked)
            self.assertTrue(saved['useStaticHostMappings'])
        # Recreate the dialog with the saved config, as after restarting.
        self.page.evaluate('window.onInit(null)')
        expect(mappings).to_have_count(0)
        self.page.evaluate('config => window.onInit({config, webView2Version: "test"})', saved)
        expect(self.page.get_by_label('Static host mappings', exact=True)).to_have_value(
            expected.replace(',', '\n'))
        expect(self.page.get_by_label(
            'Fall back to standard DNS when all mapped addresses are unreachable', exact=True)
        ).not_to_be_checked()

    def test_start_with_windows_placement_and_save(self):
        toggle = self.page.get_by_label('Start with Windows', exact=True)
        expect(toggle).not_to_be_checked()
        expect(self.page.get_by_text(
            'Launches in the tray when you sign in to Windows.', exact=True)).to_be_visible()
        ids = self.page.evaluate(
            "() => [...document.querySelectorAll('input[type=checkbox]')].map(e => e.id)")
        self.assertEqual(ids[ids.index('startWithWindows') + 1], 'autoCheckForUpdates')
        for checked in [True, False]:
            toggle.set_checked(checked)
            self.page.get_by_role('button', name='Save', exact=True).click()
            self.assertIs(self.last_message('saveSettings')['startWithWindows'], checked)
        # Reopening reflects the Windows state reported by the host.
        saved = dict(self.last_message('saveSettings'), startWithWindows=True)
        self.page.evaluate('window.onInit(null)')
        expect(toggle).to_have_count(0)
        self.page.evaluate('config => window.onInit({config, webView2Version: "test"})', saved)
        expect(toggle).to_be_checked()

    def test_static_hosts_still_reject_invalid_entries(self):
        self.page.get_by_label('Resolve listed hostnames to static IP addresses', exact=True).check()
        mappings = self.page.get_by_label('Static host mappings', exact=True)
        for invalid in ['domain.com:not-an-ip', 'domain.com:999.1.1.1',
                        'domain.com:2001:db8::1', 'bad/host:1.2.3.4']:
            mappings.fill(f'domain.com:1.2.3.4\n{invalid}')
            self.page.get_by_role('button', name='Save', exact=True).click()
            self.assertIsNone(self.last_message('saveSettings'))
            expect(mappings).to_have_class(re.compile('border-red-500'))

    def test_unchanged_config_closes_without_prompt(self):
        # Defaulted fields and comma-separated lists must not look edited on load.
        for config in [None, FULL_CONFIG]:
            for route in ['Cancel', 'native', 'Escape']:
                with self.subTest(config=config, route=route):
                    self.reset_config(config)
                    if route == 'native':
                        self.request_native_close()
                    elif route == 'Escape':
                        self.page.keyboard.press('Escape')
                    else:
                        self.page.get_by_role('button', name='Cancel', exact=True).click()
                    self.assertEqual(self.close_actions(), [{'action': 'close'}])
                    expect(self.page.get_by_role('alertdialog')).to_have_count(0)

    def test_every_setting_guards_close_and_reverting_clears_changes(self):
        for field, value in FULL_CONFIG.items():
            if field in ['updateCheckPending', 'updatePromptPending']:
                continue
            with self.subTest(field=field):
                self.reset_config(FULL_CONFIG)
                control = self.page.locator('#' + field)
                if isinstance(value, bool):
                    control.set_checked(not value)
                else:
                    control.fill('changed')
                self.request_native_close()
                self.expect_close_prompt()
                self.request_native_close()  # Repeated X must not force a close.
                dialog = self.expect_close_prompt()
                dialog.get_by_role('button', name='Keep editing', exact=True).click()
                expect(self.page.get_by_role('alertdialog')).to_have_count(0)
                self.assertEqual(self.close_actions(), [])
                if isinstance(value, bool):
                    expect(control).to_be_checked(checked=not value)
                    control.set_checked(value)
                else:
                    expect(control).to_have_value('changed')
                    control.fill(value.replace(',', '\n') if field in [
                        'insecureContentOrigins', 'staticHostMappings'] else value)
                self.page.get_by_role('button', name='Cancel', exact=True).click()
                self.assertEqual(self.close_actions(), [{'action': 'close'}])
                expect(self.page.get_by_role('alertdialog')).to_have_count(0)

    def test_close_prompt_keyboard_overlay_and_discard(self):
        for width in [900, 480]:
            with self.subTest(width=width):
                self.page.set_viewport_size({'width': width, 'height': 700})
                self.reset_config()
                self.page.locator('#startWithWindows').check()
                cancel = self.page.get_by_role('button', name='Cancel', exact=True)
                cancel.click()
                dialog = self.expect_close_prompt()
                keep = dialog.get_by_role('button', name='Keep editing', exact=True)
                save = dialog.get_by_role('button', name='Save', exact=True)
                expect(keep).to_be_focused()
                self.page.keyboard.press('Tab')
                self.page.keyboard.press('Tab')
                expect(save).to_be_focused()
                self.page.keyboard.press('Tab')
                expect(keep).to_be_focused()
                self.page.keyboard.press('Shift+Tab')
                expect(save).to_be_focused()
                self.assertFalse(self.page.locator('#url').evaluate(
                    'el => { el.focus(); return document.activeElement === el; }'))
                self.assertEqual(dialog.evaluate(
                    'el => getComputedStyle(el.parentElement).backgroundColor'),
                    'rgba(0, 0, 0, 0.35)')
                box = dialog.bounding_box()
                self.assertAlmostEqual(box['x'] + box['width'] / 2, width / 2, delta=1)
                self.assertAlmostEqual(box['y'] + box['height'] / 2, 350, delta=1)
                self.page.mouse.click(8, 8)
                self.expect_close_prompt()
                expect(save).to_be_focused()
                self.page.keyboard.press('Escape')
                expect(dialog).to_have_count(0)
                expect(cancel).to_be_focused()
                expect(self.page.locator('#startWithWindows')).to_be_checked()
                self.page.keyboard.press('Escape')
                dialog = self.expect_close_prompt()
                self.page.keyboard.press('Enter')  # Keep editing is the safe default.
                expect(dialog).to_have_count(0)
                self.request_native_close()
                dialog = self.expect_close_prompt()
                dialog.get_by_role('button', name='Discard', exact=True).click()
                self.assertEqual(self.close_actions(), [{'action': 'close'}])

    def test_prompt_save_matches_normal_save(self):
        saved = []
        for through_prompt in [False, True]:
            self.reset_config(FULL_CONFIG)
            self.page.locator('#startWithWindows').check()
            self.page.locator('#windowTitle').fill('Changed title')
            self.page.locator('#staticHostMappings').fill(
                'EXAMPLE.com:127.0.0.2\nexample.com:127.0.0.1\nexample.com:127.0.0.2')
            scope = self.page
            if through_prompt:
                self.request_native_close()
                scope = self.expect_close_prompt()
            scope.get_by_role('button', name='Save', exact=True).click()
            expected = {key: value for key, value in FULL_CONFIG.items()
                        if key not in ['updateCheckPending', 'updatePromptPending']}
            expected.update(action='saveSettings', startWithWindows=True,
                            windowTitle='Changed title',
                            staticHostMappings='example.com:127.0.0.2,example.com:127.0.0.1')
            self.assertEqual(self.close_actions(), [expected])
            saved.append(self.last_message('saveSettings'))
        self.assertEqual(saved[0], saved[1])
        # A newly opened configuration uses the saved values as its baseline.
        self.reset_config(saved[1])
        self.request_native_close()
        self.assertEqual(self.close_actions(), [{'action': 'close'}])

    def test_invalid_prompt_save_returns_to_the_invalid_field(self):
        for field, value, message in [
            ('url', '', 'URL cannot be empty.'),
            ('mailtoTargetUrl', 'invalid', 'Enter a valid destination URL.'),
            ('insecureContentOrigins', 'https://example.com', 'Only http:// origins are allowed: https://example.com'),
            ('insecureContentOrigins', '', 'Add at least one HTTP origin to allow.'),
            ('staticHostMappings', 'invalid', 'Use hostname:IP format: invalid'),
            ('staticHostMappings', '', 'Add at least one hostname and IP address.'),
        ]:
            with self.subTest(field=field, value=value):
                self.reset_config(FULL_CONFIG)
                self.page.locator('#' + field).fill(value)
                self.request_native_close()
                dialog = self.expect_close_prompt()
                dialog.get_by_role('button', name='Save', exact=True).click()
                expect(dialog).to_have_count(0)
                expect(self.page.get_by_text(message, exact=True)).to_be_visible()
                expect(self.page.locator('#' + field)).to_be_focused()
                self.assertEqual(self.close_actions(), [])

    def test_unsaved_prompt_takes_priority_over_updates(self):
        self.page.locator('#debugLog').check()
        self.request_native_close()
        self.expect_close_prompt()
        self.result(automatic=True)
        dialog = self.expect_close_prompt()
        dialog.get_by_role('button', name='Keep editing', exact=True).click()
        update = self.page.get_by_role('alertdialog', name='Update result')
        expect(update).to_be_visible()
        self.assertEqual(update.evaluate(
            'el => getComputedStyle(el.parentElement).backgroundColor'),
            'rgba(0, 0, 0, 0.35)')
        self.request_native_close()
        self.expect_close_prompt()
        self.page.keyboard.press('Escape')
        expect(update).to_be_visible()
        self.page.keyboard.press('Escape')  # An update prompt consumes Escape.
        expect(update).to_be_visible()
        update.get_by_role('button', name='Cancel', exact=True).click()
        expect(self.page.get_by_role('alertdialog')).to_have_count(0)
        expect(self.page.locator('#debugLog')).to_be_checked()
        self.assertEqual(self.close_actions(), [])

    def test_update_state_is_not_an_unsaved_setting(self):
        self.result()
        self.page.get_by_label('Reopen settings after update').check()
        self.request_native_close()
        self.assertEqual(self.close_actions(), [{'action': 'close'}])
        expect(self.page.get_by_role('alertdialog', name='Unsaved changes')).to_have_count(0)


if __name__ == '__main__':
    unittest.main()
