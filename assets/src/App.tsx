import { useEffect, useRef, useState } from "react";
import { type InitData, onInit, getInit, reportHeight } from "./lib/bridge";
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
    if (!el) return;
    const observer = new ResizeObserver((entries) => {
      for (const entry of entries) {
        reportHeight(Math.ceil(entry.contentRect.height));
      }
    });
    observer.observe(el);
    return () => observer.disconnect();
  }, []);

  return (
    <div ref={rootRef}>
      {initData && <ConfigView config={initData.config} />}
    </div>
  );
}
