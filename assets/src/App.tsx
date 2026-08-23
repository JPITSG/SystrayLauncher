import { useEffect, useRef, useState } from "react";
import { type InitData, onInit, getInit, reportSize } from "./lib/bridge";
import ConfigView from "./ConfigView";

export default function App() {
  const [initData, setInitData] = useState<InitData | null>(null);
  const rootRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    onInit((data) => setInitData(data));
    getInit();
  }, []);

  useEffect(() => {
    const el = rootRef.current;
    if (!el || !initData) return;

    const report = () => {
      // Ask for more width only when the content genuinely wants it (the
      // two-column layout sets an explicit width wider than the window);
      // reporting the fluid width back would just echo the window size.
      const width =
        el.scrollWidth > window.innerWidth + 1 ? Math.ceil(el.scrollWidth) : 0;
      reportSize(Math.ceil(el.scrollHeight), width);
    };
    const rafId = requestAnimationFrame(report);

    const observer = new ResizeObserver(report);
    observer.observe(el);
    return () => {
      cancelAnimationFrame(rafId);
      observer.disconnect();
    };
  }, [initData]);

  if (!initData) return null;

  return (
    <div ref={rootRef}>
      <ConfigView
        config={initData.config}
        webView2Version={initData.webView2Version ?? "Unknown"}
        updateCompletedVersion={initData.updateCompletedVersion ?? ""}
      />
    </div>
  );
}
