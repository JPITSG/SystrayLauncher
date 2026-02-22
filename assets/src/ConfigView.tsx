import { useState } from "react";
import { type ConfigData, saveSettings, closeDialog } from "./lib/bridge";
import { Button } from "./components/ui/button";
import { Input } from "./components/ui/input";
import { Label } from "./components/ui/label";
import { Textarea } from "./components/ui/textarea";

interface Props {
  config: ConfigData;
}

export default function ConfigView({ config }: Props) {
  const [windowTitle, setWindowTitle] = useState(config.windowTitle);
  const [url, setUrl] = useState(config.url);
  const [onHideJs, setOnHideJs] = useState(config.onHideJs);
  const [onShowJs, setOnShowJs] = useState(config.onShowJs);
  const [urlError, setUrlError] = useState("");

  function handleSave() {
    const trimmedUrl = url.trim();
    if (!trimmedUrl) {
      setUrlError("URL cannot be empty.");
      return;
    }
    setUrlError("");
    saveSettings({
      url: trimmedUrl,
      windowTitle,
      onHideJs,
      onShowJs,
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
