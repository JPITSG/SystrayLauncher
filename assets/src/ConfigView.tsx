import { useEffect, useLayoutEffect, useRef, useState } from "react";
import {
  type ConfigData,
  type UpdateResult,
  saveSettings,
  closeDialog,
  checkForUpdate,
  cancelUpdateCheck,
  configReady,
  installUpdate,
  dismissUpdate,
  ignoreUpdateVersion,
  dismissUpdateConfirmation,
  onUpdateResult,
  onUpdateProgress,
  setDesiredContentWidth,
} from "./lib/bridge";
import { Button } from "./components/ui/button";
import { Checkbox } from "./components/ui/checkbox";
import { Input } from "./components/ui/input";
import { Label } from "./components/ui/label";
import { Textarea } from "./components/ui/textarea";

interface Props {
  config: ConfigData;
  webView2Version: string;
  updateCompletedVersion: string;
}

function normalizeHttpOrigins(raw: string): string {
  const entries = raw
    .split(/[\r\n,]+/)
    .map((entry) => entry.trim())
    .filter(Boolean);
  const origins: string[] = [];

  for (const entry of entries) {
    let parsed: URL;
    try {
      parsed = new URL(entry);
    } catch {
      throw new Error(`Invalid HTTP URL: ${entry}`);
    }
    if (parsed.protocol !== "http:") {
      throw new Error(`Only http:// origins are allowed: ${entry}`);
    }
    if (parsed.username || parsed.password) {
      throw new Error("Origins cannot contain a username or password.");
    }
    if (!origins.includes(parsed.origin)) origins.push(parsed.origin);
  }

  return origins.join(",");
}

function normalizeStaticHostMappings(raw: string): string {
  const entries = raw
    .split(/[\r\n,]+/)
    .map((entry) => entry.trim())
    .filter(Boolean);
  const mappings: string[] = [];
  const mappedAddresses = new Map<string, string>();

  for (const entry of entries) {
    const separator = entry.indexOf(":");
    if (separator <= 0 || separator === entry.length - 1) {
      throw new Error(`Use hostname:IP format: ${entry}`);
    }

    const rawHostname = entry.slice(0, separator).trim();
    const rawAddress = entry.slice(separator + 1).trim();
    let parsedHostname: URL;
    try {
      parsedHostname = new URL(`http://${rawHostname}/`);
    } catch {
      throw new Error(`Invalid hostname: ${rawHostname}`);
    }

    const hostname = parsedHostname.hostname.toLowerCase();
    const labels = hostname.split(".");
    if (
      !hostname ||
      hostname.length > 253 ||
      parsedHostname.host !== parsedHostname.hostname ||
      parsedHostname.username ||
      parsedHostname.password ||
      parsedHostname.pathname !== "/" ||
      parsedHostname.search ||
      parsedHostname.hash ||
      hostname.startsWith("[") ||
      labels.some(
        (label) =>
          !label ||
          label.length > 63 ||
          !/^[a-z0-9_-]+$/i.test(label) ||
          label.startsWith("-") ||
          label.endsWith("-")
      )
    ) {
      throw new Error(`Invalid hostname: ${rawHostname}`);
    }

    let address: string;
    if (rawAddress.startsWith("[") && rawAddress.endsWith("]")) {
      try {
        const parsedAddress = new URL(`http://${rawAddress}/`);
        address = parsedAddress.hostname.toLowerCase();
        if (!address.startsWith("[") || !address.includes(":")) {
          throw new Error();
        }
      } catch {
        throw new Error(`Invalid IP address for ${hostname}: ${rawAddress}`);
      }
    } else {
      const octets = rawAddress.split(".");
      if (
        octets.length !== 4 ||
        octets.some(
          (octet) => !/^\d{1,3}$/.test(octet) || Number(octet) > 255
        )
      ) {
        throw new Error(`Invalid IP address for ${hostname}: ${rawAddress}`);
      }
      address = octets.map((octet) => String(Number(octet))).join(".");
    }

    const previousAddress = mappedAddresses.get(hostname);
    if (previousAddress && previousAddress !== address) {
      throw new Error(`Hostname is mapped more than once: ${hostname}`);
    }
    if (!previousAddress) {
      mappedAddresses.set(hostname, address);
      mappings.push(`${hostname}:${address}`);
    }
  }

  return mappings.join(",");
}

export default function ConfigView({
  config,
  webView2Version,
  updateCompletedVersion,
}: Props) {
  const [windowTitle, setWindowTitle] = useState(config.windowTitle);
  const [url, setUrl] = useState(config.url);
  const [startMaximized, setStartMaximized] = useState(
    config.startMaximized ?? false
  );
  const [returnToTargetOnDoubleClick, setReturnToTargetOnDoubleClick] = useState(
    config.returnToTargetOnDoubleClick ?? true
  );
  const [showInTaskbar, setShowInTaskbar] = useState(
    config.showInTaskbar ?? false
  );
  const [onHideJs, setOnHideJs] = useState(config.onHideJs);
  const [onShowJs, setOnShowJs] = useState(config.onShowJs);
  const [sleepWhenInactive, setSleepWhenInactive] = useState(
    config.sleepWhenInactive ?? false
  );
  const [openNewWindowsExternally, setOpenNewWindowsExternally] = useState(
    config.openNewWindowsExternally ?? false
  );
  const [allowRunningInsecureContent, setAllowRunningInsecureContent] = useState(
    config.allowRunningInsecureContent ?? false
  );
  const [insecureContentOrigins, setInsecureContentOrigins] = useState(
    (config.insecureContentOrigins ?? "").split(",").join("\n")
  );
  const [useStaticHostMappings, setUseStaticHostMappings] = useState(
    config.useStaticHostMappings ?? false
  );
  const [staticHostMappings, setStaticHostMappings] = useState(
    (config.staticHostMappings ?? "").split(",").join("\n")
  );
  const [staticHostDnsFallback, setStaticHostDnsFallback] = useState(
    config.staticHostDnsFallback ?? false
  );
  const [lockdownHeader, setLockdownHeader] = useState(
    config.lockdownHeader ?? false
  );
  const [lockdownSecret, setLockdownSecret] = useState(
    config.lockdownSecret ?? ""
  );
  const [autoCheckForUpdates, setAutoCheckForUpdates] = useState(
    config.autoCheckForUpdates ?? true
  );
  const [debugLog, setDebugLog] = useState(config.debugLog ?? false);
  // Two-column reflow: when a single column would overflow the monitor's
  // usable height (so the dialog would have to scroll), the settings latch
  // into two balanced CSS columns at roughly double width for the rest of
  // this dialog session; the host window follows the reported size.
  const fieldsRef = useRef<HTMLDivElement>(null);
  const [twoColumn, setTwoColumn] = useState(false);
  const [twoColumnWidth, setTwoColumnWidth] = useState(0);
  useLayoutEffect(() => {
    if (twoColumn) return;
    const fields = fieldsRef.current;
    if (!fields) return;
    const evaluate = () => {
      // Work-area height minus a generous allowance for the window frame.
      const usable = window.screen.availHeight - 96;
      const baseWidth = document.body.clientWidth;
      if (
        document.body.scrollHeight > usable &&
        window.screen.availWidth >= baseWidth * 2 + 64
      ) {
        const width = baseWidth * 2 + 32;
        setDesiredContentWidth(width);
        setTwoColumnWidth(width);
        setTwoColumn(true);
      }
    };
    evaluate();
    const observer = new ResizeObserver(evaluate);
    observer.observe(fields);
    return () => observer.disconnect();
  }, [twoColumn]);
  const [urlError, setUrlError] = useState("");
  const [insecureOriginsError, setInsecureOriginsError] = useState("");
  const [staticHostsError, setStaticHostsError] = useState("");
  const [updateChecking, setUpdateChecking] = useState(
    config.updateCheckPending ?? false
  );
  const [updateCancelling, setUpdateCancelling] = useState(false);
  const [updateSpeedKbps, setUpdateSpeedKbps] = useState<number | null>(null);
  const [updateAlert, setUpdateAlert] = useState<UpdateResult | null>(() =>
    updateCompletedVersion
      ? {
          status: "completed",
          title: "Update complete",
          message: `SystrayLauncher has been updated to version ${updateCompletedVersion}.`,
          currentVersion: "",
          remoteVersion: "",
          automatic: false,
        }
      : null
  );
  const automaticUpdateStarted = useRef(false);

  useEffect(() => {
    const removeResultListener = onUpdateResult((result) => {
      setUpdateChecking(false);
      setUpdateCancelling(false);
      setUpdateSpeedKbps(null);
      if (result.status === "cancelled") {
        setUpdateAlert((current) =>
          result.automatic && current?.status === "completed" ? current : null
        );
      } else if (result.automatic && result.status !== "newer") {
        setUpdateAlert((current) =>
          current?.status === "completed" ? current : null
        );
      } else {
        setUpdateAlert(result);
      }
    });
    const removeProgressListener = onUpdateProgress((progress) => {
      setUpdateSpeedKbps(Math.max(0, Math.round(progress.kilobytesPerSecond)));
    });

    const shouldCheckAutomatically =
      config.autoCheckForUpdates &&
      !updateCompletedVersion &&
      !config.updateCheckPending &&
      !config.updatePromptPending &&
      !automaticUpdateStarted.current;
    if (shouldCheckAutomatically) {
      automaticUpdateStarted.current = true;
      setUpdateChecking(true);
    }
    configReady(shouldCheckAutomatically);

    return () => {
      removeResultListener();
      removeProgressListener();
    };
  }, [
    config.autoCheckForUpdates,
    config.updateCheckPending,
    config.updatePromptPending,
    updateCompletedVersion,
  ]);

  function handleUpdate() {
    if (updateChecking) {
      setUpdateCancelling(true);
      cancelUpdateCheck();
      return;
    }
    setUpdateAlert(null);
    setUpdateChecking(true);
    setUpdateCancelling(false);
    setUpdateSpeedKbps(null);
    checkForUpdate(false);
  }

  function handleInstallUpdate() {
    setUpdateChecking(true);
    setUpdateCancelling(false);
    setUpdateSpeedKbps(null);
    installUpdate();
  }

  function handleDismissUpdate() {
    if (updateAlert?.status === "completed") {
      dismissUpdateConfirmation();
    } else {
      dismissUpdate();
    }
    setUpdateAlert(null);
  }

  function handleIgnoreUpdateVersion() {
    if (!updateAlert?.remoteVersion) return;
    ignoreUpdateVersion(updateAlert.remoteVersion);
    setUpdateAlert(null);
  }

  function handleSave() {
    const trimmedUrl = url.trim();
    if (!trimmedUrl) {
      setUrlError("URL cannot be empty.");
      return;
    }

    let normalizedInsecureOrigins = "";
    if (allowRunningInsecureContent) {
      try {
        normalizedInsecureOrigins = normalizeHttpOrigins(insecureContentOrigins);
      } catch (error) {
        setInsecureOriginsError(
          error instanceof Error ? error.message : "Invalid HTTP origin."
        );
        return;
      }
    } else if (insecureContentOrigins.trim()) {
      try {
        normalizedInsecureOrigins = normalizeHttpOrigins(insecureContentOrigins);
      } catch {
        // A disabled legacy/INI value must not prevent saving other settings.
        normalizedInsecureOrigins = "";
      }
    }
    if (allowRunningInsecureContent && !normalizedInsecureOrigins) {
      setInsecureOriginsError("Add at least one HTTP origin to allow.");
      return;
    }

    let normalizedStaticHosts = "";
    if (useStaticHostMappings) {
      try {
        normalizedStaticHosts = normalizeStaticHostMappings(staticHostMappings);
      } catch (error) {
        setStaticHostsError(
          error instanceof Error ? error.message : "Invalid static host mapping."
        );
        return;
      }
    } else if (staticHostMappings.trim()) {
      try {
        normalizedStaticHosts = normalizeStaticHostMappings(staticHostMappings);
      } catch {
        // A disabled legacy/INI value must not prevent saving other settings.
        normalizedStaticHosts = "";
      }
    }
    if (useStaticHostMappings && !normalizedStaticHosts) {
      setStaticHostsError("Add at least one hostname and IP address.");
      return;
    }

    setUrlError("");
    setInsecureOriginsError("");
    setStaticHostsError("");
    saveSettings({
      url: trimmedUrl,
      windowTitle,
      startMaximized,
      returnToTargetOnDoubleClick,
      showInTaskbar,
      onHideJs,
      onShowJs,
      sleepWhenInactive,
      openNewWindowsExternally,
      allowRunningInsecureContent,
      insecureContentOrigins: normalizedInsecureOrigins,
      useStaticHostMappings,
      staticHostMappings: normalizedStaticHosts,
      staticHostDnsFallback,
      lockdownHeader,
      lockdownSecret: lockdownSecret.trim(),
      autoCheckForUpdates,
      updateCheckPending: config.updateCheckPending,
      updatePromptPending: config.updatePromptPending,
      debugLog,
    });
  }

  return (
    <div
      className="p-4 space-y-3"
      style={
        // maxWidth absorbs sub-pixel DPI rounding and scrollbar-width
        // differences between the wanted and granted window size, so the
        // fixed-width layout can never spill into a horizontal scrollbar.
        twoColumn ? { width: twoColumnWidth, maxWidth: "100%" } : undefined
      }
    >
      {/* Layout-only wrapper: single column normally, two balanced columns
          once the settings outgrow the screen; each block stays intact. */}
      <div
        ref={fieldsRef}
        className={
          twoColumn
            ? "columns-2 gap-x-8 [&>*]:mb-3 [&>*]:break-inside-avoid"
            : "space-y-3"
        }
      >
      <div className="space-y-1">
        <Label htmlFor="windowTitle">Window Title</Label>
        <Input
          id="windowTitle"
          value={windowTitle}
          onChange={(e) => setWindowTitle(e.target.value)}
        />
      </div>

      <div className="space-y-1">
        <Label htmlFor="url">URL</Label>
        <Input
          id="url"
          value={url}
          onChange={(e) => {
            setUrl(e.target.value);
            if (urlError) setUrlError("");
          }}
          className={urlError ? "border-red-500" : ""}
        />
        {urlError && (
          <p className="text-red-600 text-[11px]">{urlError}</p>
        )}
      </div>

      <div className="flex items-start gap-2 pt-1">
        <Checkbox
          id="startMaximized"
          className="mt-0.5"
          checked={startMaximized}
          onChange={(e) => setStartMaximized(e.target.checked)}
        />
        <div className="space-y-0.5">
          <Label htmlFor="startMaximized" className="cursor-pointer">
            Open main window maximized
          </Label>
          <p className="text-neutral-500 text-[11px] leading-snug">
            Fills the available desktop whenever the window opens. Leave off to
            open it centered at 90% of the work area.
          </p>
        </div>
      </div>

      <div className="flex items-start gap-2 pt-1">
        <Checkbox
          id="showInTaskbar"
          className="mt-0.5"
          checked={showInTaskbar}
          onChange={(e) => setShowInTaskbar(e.target.checked)}
        />
        <div className="space-y-0.5">
          <Label htmlFor="showInTaskbar" className="cursor-pointer">
            Show main window in the taskbar
          </Label>
          <p className="text-neutral-500 text-[11px] leading-snug">
            Adds a taskbar button while the main window is open. Leave off to
            keep the launcher tray-only. Changing this setting restarts the
            launcher.
          </p>
        </div>
      </div>

      <div className="flex items-start gap-2 pt-1">
        <Checkbox
          id="returnToTargetOnDoubleClick"
          className="mt-0.5"
          checked={returnToTargetOnDoubleClick}
          onChange={(e) => setReturnToTargetOnDoubleClick(e.target.checked)}
        />
        <div className="space-y-0.5">
          <Label
            htmlFor="returnToTargetOnDoubleClick"
            className="cursor-pointer"
          >
            Return to configured URL on tray double-click
          </Label>
          <p className="text-neutral-500 text-[11px] leading-snug">
            Returns to the configured URL when you double-click the tray icon,
            including while the main window is already open. Leave off to only
            open or focus the current page.
          </p>
        </div>
      </div>

      <div className="space-y-1">
        <Label htmlFor="onHideJs">
          JavaScript on Hide (window fully covered)
        </Label>
        <Textarea
          id="onHideJs"
          rows={3}
          value={onHideJs}
          onChange={(e) => setOnHideJs(e.target.value)}
        />
      </div>

      <div className="space-y-1">
        <Label htmlFor="onShowJs">
          JavaScript on Show (window becomes visible)
        </Label>
        <Textarea
          id="onShowJs"
          rows={3}
          value={onShowJs}
          onChange={(e) => setOnShowJs(e.target.value)}
        />
      </div>

      <div className="space-y-1 pt-1">
        <div className="flex items-start gap-2">
          <Checkbox
            id="useStaticHostMappings"
            className="mt-0.5"
            checked={useStaticHostMappings}
            onChange={(e) => {
              setUseStaticHostMappings(e.target.checked);
              if (staticHostsError) setStaticHostsError("");
            }}
          />
          <div className="space-y-0.5">
            <Label htmlFor="useStaticHostMappings" className="cursor-pointer">
              Resolve listed hostnames to static IP addresses
            </Label>
            <p className="text-red-600 text-[11px] leading-snug">
              Routes the listed hostnames to their configured IP addresses
              inside this web container, bypassing normal DNS. Changing this
              setting restarts the launcher.
            </p>
          </div>
        </div>
        {useStaticHostMappings && (
          <div className="ml-6 space-y-1">
            <Label htmlFor="staticHostMappings">Static host mappings</Label>
            <Textarea
              id="staticHostMappings"
              rows={2}
              maxLength={1800}
              placeholder={"device.local:192.168.1.20\napi.example.com:10.0.0.8"}
              value={staticHostMappings}
              onChange={(e) => {
                setStaticHostMappings(e.target.value);
                if (staticHostsError) setStaticHostsError("");
              }}
              className={staticHostsError ? "border-red-500" : ""}
            />
            <p className="text-neutral-500 text-[11px] leading-snug">
              One hostname:IP mapping per line or comma-separated. HTTPS
              certificates are still checked against the hostname. List each
              subdomain separately; wrap IPv6 addresses in brackets.
            </p>
            {staticHostsError && (
              <p className="text-red-600 text-[11px]">{staticHostsError}</p>
            )}
            <div className="flex items-start gap-2 pt-1">
              <Checkbox
                id="staticHostDnsFallback"
                className="mt-0.5"
                checked={staticHostDnsFallback}
                onChange={(e) => setStaticHostDnsFallback(e.target.checked)}
              />
              <div className="space-y-0.5">
                <Label htmlFor="staticHostDnsFallback" className="cursor-pointer">
                  Fall back to standard DNS when a mapped address is unreachable
                </Label>
                <p className="text-neutral-500 text-[11px] leading-snug">
                  Routes only the listed hostnames through a small local helper
                  that connects to the mapped address when it responds and
                  quietly uses normal DNS resolution while it does not,
                  re-trying the mapped address about once a minute. Useful when
                  the mapped addresses are reachable only from certain
                  networks. Changing this setting restarts the launcher.
                </p>
              </div>
            </div>
          </div>
        )}
      </div>

      <div className="space-y-1 pt-1">
        <div className="flex items-start gap-2">
          <Checkbox
            id="allowRunningInsecureContent"
            className="mt-0.5"
            checked={allowRunningInsecureContent}
            onChange={(e) => {
              setAllowRunningInsecureContent(e.target.checked);
              if (insecureOriginsError) setInsecureOriginsError("");
            }}
          />
          <div className="space-y-0.5">
            <Label
              htmlFor="allowRunningInsecureContent"
              className="cursor-pointer"
            >
              Allow listed HTTP origins on HTTPS pages
            </Label>
            <p className="text-red-600 text-[11px] leading-snug">
              Treats the listed HTTP origins as trustworthy so they can load in
              HTTPS pages. This weakens browser security for those origins.
              Changing this setting restarts the launcher.
            </p>
          </div>
        </div>
        {allowRunningInsecureContent && (
          <div className="ml-6 space-y-1">
            <Label htmlFor="insecureContentOrigins">Allowed HTTP origins</Label>
            <Textarea
              id="insecureContentOrigins"
              rows={2}
              maxLength={1800}
              placeholder={"http://device.local:8080\nhttp://192.168.1.20"}
              value={insecureContentOrigins}
              onChange={(e) => {
                setInsecureContentOrigins(e.target.value);
                if (insecureOriginsError) setInsecureOriginsError("");
              }}
              className={insecureOriginsError ? "border-red-500" : ""}
            />
            <p className="text-neutral-500 text-[11px] leading-snug">
              One URL per line or comma-separated. You may paste a full URL;
              only its origin (scheme, host, and port) is saved. List every HTTP
              origin used by the iframe and its redirects.
            </p>
            {insecureOriginsError && (
              <p className="text-red-600 text-[11px]">{insecureOriginsError}</p>
            )}
          </div>
        )}
      </div>

      <div className="space-y-1 pt-1">
        <div className="flex items-start gap-2">
          <Checkbox
            id="lockdownHeader"
            className="mt-0.5"
            checked={lockdownHeader}
            onChange={(e) => setLockdownHeader(e.target.checked)}
          />
          <div className="space-y-0.5">
            <Label htmlFor="lockdownHeader" className="cursor-pointer">
              Send X-Lockdown header
            </Label>
            <p className="text-neutral-500 text-[11px] leading-snug">
              Stamps every request with an X-Lockdown header: the browser's
              User-Agent encrypted with a key derived from the current UTC hour
              and the secret below. A gateway that knows the secret can require
              the header as an extra access check (server recipe in the README).
            </p>
          </div>
        </div>
        {lockdownHeader && (
          <div className="ml-6 space-y-1">
            <Label htmlFor="lockdownSecret">Shared secret</Label>
            <Input
              id="lockdownSecret"
              maxLength={200}
              placeholder="use the same value on the server"
              value={lockdownSecret}
              onChange={(e) => setLockdownSecret(e.target.value)}
            />
            <p className="text-neutral-500 text-[11px] leading-snug">
              Mixed into the hourly encryption key. Optional, but without it
              anyone who knows the (public) scheme can forge the header.
              Leading and trailing spaces are ignored.
            </p>
          </div>
        )}
      </div>

      <div className="flex items-start gap-2 pt-1">
        <Checkbox
          id="openNewWindowsExternally"
          className="mt-0.5"
          checked={openNewWindowsExternally}
          onChange={(e) => setOpenNewWindowsExternally(e.target.checked)}
        />
        <div className="space-y-0.5">
          <Label htmlFor="openNewWindowsExternally" className="cursor-pointer">
            Open new windows in the default browser
          </Label>
          <p className="text-neutral-500 text-[11px] leading-snug">
            Links that would open a new window or tab launch in your system
            browser instead of a WebView2 popup. Popups that need to talk back
            to the page (some login flows) may not work while this is on.
          </p>
        </div>
      </div>

      <div className="flex items-start gap-2 pt-1">
        <Checkbox
          id="sleepWhenInactive"
          className="mt-0.5"
          checked={sleepWhenInactive}
          onChange={(e) => setSleepWhenInactive(e.target.checked)}
        />
        <div className="space-y-0.5">
          <Label htmlFor="sleepWhenInactive" className="cursor-pointer">
            Sleep web container when inactive
          </Label>
          <p className="text-neutral-500 text-[11px] leading-snug">
            Suspends the page to save CPU while the window is hidden. The page is
            still preloaded at startup and wakes when you hover the tray icon.
          </p>
        </div>
      </div>

      <div className="flex items-start gap-2 pt-1">
        <Checkbox
          id="autoCheckForUpdates"
          className="mt-0.5"
          checked={autoCheckForUpdates}
          onChange={(e) => setAutoCheckForUpdates(e.target.checked)}
        />
        <div className="space-y-0.5">
          <Label htmlFor="autoCheckForUpdates" className="cursor-pointer">
            Automatically check for updates
          </Label>
          <p className="text-neutral-500 text-[11px] leading-snug">
            Checks at startup, whenever this dialog opens, and every 60 minutes.
            Prompts only when a newer version is available.
          </p>
        </div>
      </div>

      <div className="flex items-start gap-2 pt-1">
        <Checkbox
          id="debugLog"
          className="mt-0.5"
          checked={debugLog}
          onChange={(e) => setDebugLog(e.target.checked)}
        />
        <div className="space-y-0.5">
          <Label htmlFor="debugLog" className="cursor-pointer">
            Enable debug logging
          </Label>
          <p className="text-neutral-500 text-[11px] leading-snug">
            Writes diagnostic events (recovery attempts, web view rebuilds,
            power transitions) to{" "}
            %LOCALAPPDATA%\SystrayLauncher\debug.log. Useful when reporting
            issues; leave off for normal use.
          </p>
        </div>
      </div>
      </div>

      <div className="flex items-center justify-between gap-3 pt-1">
        <span
          className="select-none whitespace-nowrap text-[11px] leading-none tabular-nums text-neutral-400"
          title="Application version / WebView2 version"
        >
          v{__APP_VERSION__} / {webView2Version}
        </span>
        <div className="flex items-center gap-2">
          <Button
            variant={updateChecking ? "destructive" : "outline"}
            size="sm"
            className="min-w-[5rem]"
            disabled={updateCancelling}
            aria-label={
              updateChecking ? "Stop update check and download" : undefined
            }
            title={
              updateChecking ? "Stop update check and download" : undefined
            }
            onClick={handleUpdate}
          >
            {updateCancelling
              ? "Stopping..."
              : updateChecking
                ? updateSpeedKbps === null
                  ? "Checking..."
                  : `Checking (${updateSpeedKbps.toLocaleString()} KB/s)...`
                : "Update"}
          </Button>
          <Button
            variant="outline"
            size="sm"
            className="min-w-[5rem]"
            onClick={closeDialog}
          >
            Cancel
          </Button>
          <Button size="sm" className="min-w-[5rem]" onClick={handleSave}>
            Save
          </Button>
        </div>
      </div>

      {updateAlert && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/35 p-4">
          <div
            role="alertdialog"
            aria-modal="true"
            aria-labelledby="update-alert-title"
            aria-describedby="update-alert-message"
            className="w-full max-w-sm space-y-3 rounded-lg border border-neutral-200 bg-white p-4 shadow-xl"
          >
            <div className="space-y-1">
              <h2 id="update-alert-title" className="text-sm font-semibold">
                {updateAlert.title}
              </h2>
              <p
                id="update-alert-message"
                className="text-xs leading-relaxed text-neutral-600"
              >
                {updateAlert.message}
              </p>
            </div>
            {updateAlert.currentVersion && updateAlert.remoteVersion && (
              <dl className="grid grid-cols-[1fr_auto] gap-x-4 gap-y-1 rounded-md border border-neutral-200 bg-neutral-50 px-3 py-2 text-xs">
                <dt className="text-neutral-500">Current version</dt>
                <dd className="font-medium tabular-nums text-neutral-900">
                  {updateAlert.currentVersion}
                </dd>
                <dt className="text-neutral-500">Remote version</dt>
                <dd className="font-medium tabular-nums text-neutral-900">
                  {updateAlert.remoteVersion}
                </dd>
              </dl>
            )}
            <div className="flex justify-end gap-2">
              {updateAlert.status === "newer" && updateAlert.automatic && (
                <Button
                  variant="outline"
                  size="sm"
                  disabled={updateChecking}
                  onClick={handleIgnoreUpdateVersion}
                >
                  Ignore this version
                </Button>
              )}
              {(updateAlert.status === "newer" ||
                updateAlert.status === "same") && (
                <Button
                  variant="outline"
                  size="sm"
                  autoFocus
                  disabled={updateChecking}
                  onClick={handleDismissUpdate}
                >
                  Cancel
                </Button>
              )}
              <Button
                size="sm"
                autoFocus={
                  updateAlert.status !== "newer" &&
                  updateAlert.status !== "same"
                }
                disabled={updateChecking}
                onClick={
                  updateAlert.status === "newer" ||
                  updateAlert.status === "same"
                    ? handleInstallUpdate
                    : handleDismissUpdate
                }
              >
                {updateChecking
                  ? "Starting..."
                  : updateAlert.status === "same"
                    ? "Force update"
                    : updateAlert.status === "newer"
                      ? "Update"
                      : "OK"}
              </Button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
