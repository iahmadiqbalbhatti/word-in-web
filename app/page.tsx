"use client";

import dynamic from "next/dynamic";

// WordInWeb touches the DOM/canvas at import time, so it must never run on the server.
const TemplateEditor = dynamic(() => import("@/components/TemplateEditor"), {
  ssr: false,
  loading: () => <p style={{ padding: 24 }}>Loading editor…</p>,
});

export default function Page() {
  return <TemplateEditor />;
}