import { useState } from "react";
import { type ConfigData, saveSettings, closeDialog } from "./lib/bridge";
import { Button } from "./components/ui/button";
import { Checkbox } from "./components/ui/checkbox";
import { Input } from "./components/ui/input";
import { Label } from "./components/ui/label";
import { Textarea } from "./components/ui/textarea";

interface Props {
  config: ConfigData;
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

export default function ConfigView({ config }: Props) {
  const [windowTitle, setWindowTitle] = useState(config.windowTitle);
  const [url, setUrl] = useState(config.url);
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
  const [lockdownHeader, setLockdownHeader] = useState(
    config.lockdownHeader ?? false
  );
  const [lockdownSecret, setLockdownSecret] = useState(
    config.lockdownSecret ?? ""
  );
  const [debugLog, setDebugLog] = useState(config.debugLog ?? false);
  const [urlError, setUrlError] = useState("");
  const [insecureOriginsError, setInsecureOriginsError] = useState("");

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

    setUrlError("");
    setInsecureOriginsError("");
    saveSettings({
      url: trimmedUrl,
      windowTitle,
      onHideJs,
      onShowJs,
      sleepWhenInactive,
      openNewWindowsExternally,
      allowRunningInsecureContent,
      insecureContentOrigins: normalizedInsecureOrigins,
      lockdownHeader,
      lockdownSecret: lockdownSecret.trim(),
      debugLog,
    });
  }

  return (
    <div className="p-4 space-y-3">
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

      <div className="flex justify-end gap-2 pt-1">
        <Button variant="outline" size="sm" className="min-w-[5rem]" onClick={closeDialog}>
          Cancel
        </Button>
        <Button size="sm" className="min-w-[5rem]" onClick={handleSave}>
          Save
        </Button>
      </div>
    </div>
  );
}
