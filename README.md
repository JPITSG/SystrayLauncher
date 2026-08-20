# SystrayLauncher

A lightweight Windows system tray application that hosts a WebView2 browser window. Designed for web apps that you want quick access to without cluttering your taskbar.

## Features

- **System Tray Integration** - Runs in the system tray; double-click to open, close/minimize to hide
- **WebView2 Browser** - Uses Microsoft Edge WebView2 for modern web compatibility
- **Configurable** - Set custom URL, window title, and JavaScript hooks via GUI
- **JavaScript Hooks** - Execute custom JavaScript when the window is shown or hidden (useful for pausing/resuming web app activity)
- **External Link Handling** - Optionally open new windows/tabs (`target="_blank"`, `window.open`) in the system default browser instead of a WebView2 popup
- **Optional Mixed Content** - Explicitly allow an HTTPS page to load HTTP content from configured legacy origins
- **Lockdown Header** - Optionally stamp every request with a rolling, hour-keyed `X-Lockdown` token a gateway can require as an extra access layer
- **Preloaded on Startup** - The page is loaded into the WebView at launch so it is ready the moment you open the window
- **Hide Grace Period** - Re-opening within 60 seconds of hiding returns exactly where you left off; after that the page resets to the configured URL in the background, so the next open starts fresh with no visible reload
- **Self-Healing Container** - Every time the window opens, the app verifies the embedded browser is actually rendering (frame heartbeat, plus a screen check after resume from sleep/hibernate) and automatically rebuilds it if it is not — no more permanently blank windows after hibernation
- **Optional CPU Saving** - Opt-in "sleep when inactive" suspends the web container while hidden to save CPU on laptops, and pre-emptively wakes it when you hover the tray icon
- **Registry Storage** - Settings persist in Windows Registry (`HKCU\SOFTWARE\JPIT\SystrayLauncher`)
- **Single Instance** - Only one instance can run at a time
- **First-Launch Setup** - Configuration dialog appears automatically on first run

## Context Menu Options

Right-click the tray icon to access:

- **WebView2 Version** - Displays the current WebView2 runtime version
- **Refresh** - Reloads the page and brings window to foreground
- **Refresh + Clear Cache** - Clears browser cache and reloads
- **Open** - Shows the main window
- **Configure** - Opens the settings dialog
- **Exit** - Closes the application

## Configuration

Settings available in the Configure dialog:

| Setting | Description |
|---------|-------------|
| Window Title | Title shown in the window title bar |
| URL | The web page to load |
| JavaScript on Hide | JS executed when window is fully covered or hidden |
| JavaScript on Show | JS executed when window becomes visible |
| Allow listed HTTP origins on HTTPS pages | Lets an HTTPS page embed content from explicitly listed HTTP origins. Enter one origin per line (for example, `http://device.local:8080`). The option is disabled by default and restarts the launcher when changed. |
| Send X-Lockdown header | Adds an `X-Lockdown` header to every request the embedded browser makes: the request's own User-Agent encrypted with a key derived from the current UTC hour and an optional shared secret (see [Lockdown Header](#lockdown-header)). Toggling applies immediately. Disabled by default. |
| Open new windows in the default browser | When enabled, links that would open a new window or tab launch in the system default browser instead of a WebView2 popup. Only `http(s)` links are handed to the browser. Popups that must script back to the opening page (some login flows) may not work while enabled. Disabled by default. |
| Sleep web container when inactive | When enabled, suspends the WebView to save CPU while the window is hidden, and pre-emptively wakes it on tray-icon hover. The page is always preloaded at startup regardless of this setting. Disabled by default. |
| Enable debug logging | Appends timestamped diagnostic events (recovery attempts, web view rebuilds, power transitions) to `%LOCALAPPDATA%\SystrayLauncher\debug.log` (rotated at ~1 MB). Useful when reporting issues. Disabled by default. |

The mixed-content option starts WebView2 with the documented
[`--unsafely-treat-insecure-origin-as-secure`](https://learn.microsoft.com/en-us/microsoft-edge/webview2/concepts/webview-features-flags#available-webview2-browser-flags)
switch for the exact origins entered in the dialog. Paste a full HTTP URL if
convenient; the dialog stores only its scheme, host, and port. Every HTTP origin
reached through an iframe redirect or used for its HTTP subresources must be
listed separately.

Treating an HTTP origin as trustworthy weakens transport security and may also
enable browser APIs that normally require a secure context. It does not bypass
Content Security Policy, `X-Frame-Options`/`frame-ancestors`, certificate
errors, or cross-origin access controls. WebView2 documents browser flags for
development use rather than as a stable production API, so this legacy
compatibility option may change with future Edge runtime versions.

## Lockdown Header

When **Send X-Lockdown header** is enabled, every HTTP(S) request the embedded
browser issues (pages, subresources, XHR/fetch, iframes, service workers on
current runtimes) carries an additional header:

```
X-Lockdown: base64( IV[16 bytes] + AES-256-CBC( User-Agent ) )
```

- The plaintext is the exact `User-Agent` value sent with that same request.
- The AES key is `SHA-256( secret + "|" + "YYYY-MM-DD HH:00:00" )`, where the
  timestamp is the current UTC hour (minutes and seconds zeroed) and `secret`
  is the shared secret from the dialog (empty string if none is set).
- The IV is random and prepended; padding is PKCS#7.

A gateway verifies by trying the keys for the previous, current, and next hour
and comparing the decryption against the request's own `User-Agent` header.
The rolling window means client and server clocks only have to agree to
roughly an hour. In PHP:

```php
function lockdown_pass(string $secret): bool {
    $raw = base64_decode($_SERVER['HTTP_X_LOCKDOWN'] ?? '', true);
    if ($raw === false || strlen($raw) < 32) return false;
    $iv = substr($raw, 0, 16);
    $ciphertext = substr($raw, 16);
    $ua = $_SERVER['HTTP_USER_AGENT'] ?? '';
    foreach ([0, -3600, 3600] as $offset) {
        $key = hash('sha256', $secret . '|' . gmdate('Y-m-d H:00:00', time() + $offset), true);
        $plain = openssl_decrypt($ciphertext, 'aes-256-cbc', $key, OPENSSL_RAW_DATA, $iv);
        if ($plain !== false && hash_equals($plain, $ua)) return true;
    }
    return false;
}
```

Notes:

- **Set a shared secret.** This application is open source, so without a
  secret the scheme is only obscurity — anyone can derive the hourly key and
  mint a valid header. With a secret, forging requires knowing it. The secret
  is compared byte-for-byte (UTF-8, whitespace-trimmed by the dialog), so use
  the identical string on both sides.
- The token is replayable within its roughly two-hour validity window by
  anyone who observes it (including third-party hosts the page embeds), and
  it authenticates only the launcher, not a user. Treat it as one extra gate
  in front of your normal authentication, not a replacement for it.
- WebSocket connections are not intercepted by WebView2 and never carry the
  header, so enforce it on ordinary HTTP(S) endpoints only.

## Icon Customization

The application uses a single icon file (`icon.ico`) that appears in multiple locations:

| Location | Description |
|----------|-------------|
| System Tray | Small icon in the notification area (16x16 or 32x32 depending on DPI) |
| Window Title Bar | Icon shown in the top-left corner of the main window |
| Taskbar | Icon displayed when the window is visible |
| Alt-Tab Switcher | Icon shown when cycling through windows |

### Replacing the Icon

1. Replace `icon.svg` with your own SVG file
2. Run `make icon` to generate `icon.ico` (requires ImageMagick)
3. Rebuild with `make`

This generates a multi-resolution `.ico` containing 16x16, 24x24, 32x32, 48x48, and 256x256 sizes, covering all DPI scaling levels. The icon is embedded into the executable at compile time via `resource.rc`.

## Requirements

- Windows 10/11
- [WebView2 Runtime](https://developer.microsoft.com/en-us/microsoft-edge/webview2/) (usually pre-installed on Windows 10/11)

## Building from Source

Requires MinGW-w64 cross-compiler and the WebView2 SDK.

```bash
# Download WebView2 SDK (one-time)
make deps

# Build
make

# Output: SystrayLauncher.exe + WebView2Loader.dll in release/
```

## License

[MIT](LICENSE)
