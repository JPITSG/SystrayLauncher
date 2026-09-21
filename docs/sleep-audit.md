# Sleep and visibility audit — 1.0.26

The reported symptom was a visible window displaying the white host background
after WebView2 had been hidden for suspension. The audit found several paths
capable of causing that symptom, along with separate battery drains. The report
does not claim to reproduce the original incident on a Windows desktop.

## Findings and changes

| Finding | Change |
| --- | --- |
| Uncovering waited for the next 10-second poll; focus had no direct wake handler. | Windows events queue one 50 ms exposure check. Activation and taskbar restoration wake directly. |
| Occlusion subtracted every window's outer rectangle, including transparency, shaped regions and invisible resize borders. | Use DWM frame bounds. Exclude layered, transparent, nonredirected, cloaked and shaped occluders. Any geometry, enumeration or observation uncertainty keeps a shown window awake. |
| Allocating the target region could fail and incorrectly report it invisible. | Fail awake, release temporary regions, and require a complete scan that actually encountered the target. |
| Increasing the poll interval also lengthened resume failure recovery. | Three bounded resume attempts, with one-second retries independent of visibility events. |
| Every unchanged visibility tick reasserted controller visibility and bounds. | Cache successful state changes and bounds. Invalidate the cache on replacement and power recovery. |
| A cancelled suspend or old WebView could complete after a newer operation. | Identify suspend requests and recovery pings, reject stale completions, detach old COM objects before closing, and tag queued rebuilds by WebView generation. |
| A failed/old recovery ping could count as healthy. | Accept only successful callbacks belonging to the current recovery attempt. |
| A URL reset woke the runtime before discovering that no navigation was necessary. No completion then arrived to put it back to sleep. | Compare the URL before waking. Actual navigation explicitly restores rendering and sets the loading guard before entering WebView2. |
| Visibility/recovery callbacks could suspend during the 1.5-second navigation settle interval. The first failed load could also stay awake forever. | A shared settle guard protects the full interval. The navigation watchdog also settles the sleep state. Deadlines reject queued timers from previous navigations. |
| The health check repeatedly enumerated desktop windows and injected scripts at 100 ms intervals; a slow load or plain page could cause an unnecessary rebuild. | One frame request and one three-second deadline per attempt, started after navigation. Accept only the current frame response. Repair composition before a bounded rebuild. Do not rebuild because of uniform color or a missing frame while behind another app. |
| A replacement under a covered but shown window could reposition that window. | Preposition only a genuinely hidden, nonminimized host. |
| Every tray mouse movement reset a Windows timer. | Update a timestamp; the single expiry timer accounts for later hovers. |

## Work and resource bounds

- No repeating visibility timer. Window-event bursts share a pending check;
  further events cannot postpone an exposure check.
- Awake coverage checks are coalesced over 10 seconds. Sleeping exposure checks
  use 50 ms. A check does not schedule another check by itself.
- No global location-change subscription. Subscribe to the processes owning
  visible potential occluders, reuse those hooks, and remove obsolete hooks.
  Ignore caret, accessibility-child and child-window events before scheduling.
- At most 64 process subscriptions plus eight narrow global event hooks. If
  that bound or hook installation prevents complete observation, stay awake.
- No visibility hooks while tray-hidden or minimized. No hooks when sleep and
  both custom visibility scripts are disabled.
- Two reusable GDI regions per scan. No region allocation per covering window.
- Navigation, hover and health timers are one-shot. Actual suspension and
  controller changes happen at transitions, rather than on every event.
- Existing power-recovery and automatic-rebuild limits remain bounded. Other
  application timers, including update checks, are independent of this feature.

## Automated verification

`python3 -m unittest discover -s tests -p test_sleep.py -v` compiles the actual C
functions with deterministic Win32 and WebView2 boundary fakes. Cases cover:

- Full cover, partial cover, a one-pixel exposed strip, multiple occluders, and
  full-to-partial transitions driven by an event.
- Transparent, shaped, cloaked and minimized occluders; invisible borders;
  missing monitors; failed allocations/enumeration; missing targets; hook
  failures and capacity exhaustion.
- Event bursts, ignored child/caret events, timer priority, stale timer messages,
  absence of idle rearming, hook cleanup, and redundant COM-call suppression.
- Cancelled and superseded suspension, stale recovery callbacks, pending
  navigation, protected settling, and unchanged/changed background URLs.
- Hover event storms, hover expiry, resume retry limits, slow initial loads,
  composition repair, plain pages, and bounded rebuild attempts.

These tests establish policy and callback behavior. They do not emulate the
Windows compositor, accessibility-event delivery, GPU drivers, or WebView2's
internal suspension implementation. The release is cross-built with MinGW;
Windows desktop validation remains outstanding.

## Windows validation matrix

With sleep enabled and debug logging on, verify that the page remains responsive
and does not reload during each exposure transition:

1. Cover the entire window with one opaque window, then move or resize the
   covering window to reveal a narrow strip without focusing the launcher.
2. Repeat using two covering windows; minimize, close, or terminate one of them.
3. Cover partially with a transparent overlay, rounded/shaped window, tray
   flyout, task switcher, or the launcher's own configuration dialog.
4. Hide to tray, reopen before and after 60 seconds, and repeat with the page
   already at the configured URL. Hover without opening and let prewarm expire.
5. Minimize/restore through the taskbar, maximize, change monitor/DPI, dock,
   switch virtual desktops, and lock/unlock the session.
6. Trigger a slow navigation while hidden, then show it during loading and
   during settling. Toggle sleep off and on while covered or hidden.
7. Sleep/hibernate Windows with the launcher visible and hidden. Repeat with a
   legitimately blank page. Confirm that blank content alone causes no reload.
8. Observe host CPU/wake activity after the desktop becomes idle: visibility
   scans must stop, and hidden/minimized windows must hold no visibility hooks.

## Platform references

- [WebView2 suspension and resumption](https://learn.microsoft.com/en-us/microsoft-edge/webview2/reference/win32/icorewebview2_3): suspension requires an invisible controller, is asynchronous and best effort; navigation and becoming visible can resume it.
- [Out-of-context Windows event hooks](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwineventhook) and [event constants](https://learn.microsoft.com/en-us/windows/win32/winauto/event-constants): the observation mechanism and the distinction between window and accessibility-object changes.
- [Window regions](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindowrgn): a bounding rectangle does not describe the pixels drawn by a shaped window.
