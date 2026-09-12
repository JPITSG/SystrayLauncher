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

    def result(self, status='newer', automatic=False):
        self.page.evaluate('result => window.onUpdateResult(result)', {
            'status': status, 'automatic': automatic, 'title': 'Update result',
            'message': 'Result of checking for an update.',
            'currentVersion': '1.0.23', 'remoteVersion': '1.0.24',
        })

    def last_message(self, action):
        return self.page.evaluate(
            'action => window.messages.filter(m => m.action === action).at(-1)', action)

    def test_speed_style_and_cancellation(self):
        self.page.get_by_role('button', name='Update', exact=True).click()
        button = self.page.get_by_role('button', name='Stop update check and download')
        expect(button).to_have_text('Checking...')
        expect(button).to_have_class(re.compile('bg-red-600'))
        for speed, expected in [(100, 100), (100.49, 100), (100.5, 101),
                                (12345, 12345), (0, 0)]:
            self.page.evaluate('speed => window.onUpdateProgress({kilobytesPerSecond: speed})', speed)
            expect(button).to_have_text(f'Checking ({expected}kb/s)...')
            expect(button).to_be_enabled()  # Existing click-to-cancel behavior.
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

    def test_static_hosts_still_reject_invalid_entries(self):
        self.page.get_by_label('Resolve listed hostnames to static IP addresses', exact=True).check()
        mappings = self.page.get_by_label('Static host mappings', exact=True)
        for invalid in ['domain.com:not-an-ip', 'domain.com:999.1.1.1',
                        'domain.com:2001:db8::1', 'bad/host:1.2.3.4']:
            mappings.fill(f'domain.com:1.2.3.4\n{invalid}')
            self.page.get_by_role('button', name='Save', exact=True).click()
            self.assertIsNone(self.last_message('saveSettings'))
            expect(mappings).to_have_class(re.compile('border-red-500'))


if __name__ == '__main__':
    unittest.main()
