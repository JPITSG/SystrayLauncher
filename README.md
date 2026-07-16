# SystrayLauncher

A lightweight Windows system tray application that hosts a WebView2 browser window. Designed for web apps that you want quick access to without cluttering your taskbar.

## Features

- **System Tray Integration** - Runs in the system tray; double-click to open, close/minimize to hide
- **WebView2 Browser** - Uses Microsoft Edge WebView2 for modern web compatibility
- **Configurable** - Set custom URL, window title, and JavaScript hooks via GUI
- **JavaScript Hooks** - Execute custom JavaScript when the window is shown or hidden (useful for pausing/resuming web app activity)
- **Spell Checking** - Enable the built-in Chromium spell checker for any set of languages (e.g. English + Polish at the same time), configurable from the GUI
- **External Link Handling** - Optionally open new windows/tabs (`target="_blank"`, `window.open`) in the system default browser instead of a WebView2 popup
- **Preloaded on Startup** - The page is loaded into the WebView at launch so it is ready the moment you open the window
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
| Spell-check languages | Comma-separated language tags to spell-check simultaneously, e.g. `en-US,pl`. Empty (the default) leaves the WebView2 default behavior untouched. See [Spell Checking](#spell-checking). |
| Open new windows in the default browser | When enabled, links that would open a new window or tab launch in the system default browser instead of a WebView2 popup. Only `http(s)` links are handed to the browser. Popups that must script back to the opening page (some login flows) may not work while enabled. Disabled by default. |
| Sleep web container when inactive | When enabled, suspends the WebView to save CPU while the window is hidden, and pre-emptively wakes it on tray-icon hover. The page is always preloaded at startup regardless of this setting. Disabled by default. |

## Spell Checking

WebView2 ships the full Chromium spell checker but (as of 2026) exposes no API to
choose its languages — a fresh WebView2 profile spell-checks US English only,
regardless of Windows or Edge language settings (tracked in
[WebView2Feedback #2040](https://github.com/MicrosoftEdge/WebView2Feedback/issues/2040)).

SystrayLauncher works around this: at startup — before the browser process
launches — it writes the configured languages into its private WebView2
profile (`spellcheck.dictionaries` and `intl.accept_languages` in
`%LOCALAPPDATA%\SystrayLauncher\WebView2Data\Default\Preferences`). The app is
single-instance and owns that profile exclusively, so the edit is safe. If you
change the languages while the app is running, it offers to restart the
embedded web view (the languages are only read at browser startup); answering
No applies them on the next launch instead.

Notes:

- Use the language tags Chromium's dictionaries use: `en-US`, `en-GB`, `pl`,
  `de`, `fr`, `es`, `it`, `nl`, `pt-BR`, `ru`, `sv`, ... Multiple languages are
  checked simultaneously.
- Each language needs spell-check data on the machine. Windows usually has
  this once the language is added under **Settings → Time & Language →
  Language & region** (the "Basic typing" feature). From PowerShell:
  `Add-WindowsCapability -Online -Name "Language.BasicTyping~~~pl-PL~0.0.1.0"`
- Right-click a misspelled word in the web view for suggestions and
  "Add to dictionary". The setting also becomes the browser's
  `Accept-Language` list.
- Leaving the field empty means SystrayLauncher never touches the profile's
  spell-check configuration.

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
