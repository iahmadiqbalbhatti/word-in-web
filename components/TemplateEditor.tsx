"use client";

import { useCallback, useMemo, useRef, useState } from "react";
import { DocxView, DocxToolbar, type DocxViewApi } from "wordinweb";
import { matters, matterLabel, flattenKeys, safeFileName } from "@/lib/matters";
import { fillTemplate, downloadBytes, TemplateError } from "@/lib/fillTemplate";

const STARTER_TEMPLATE = "/templates/make-pleading-template.docx";

type Notice = { kind: "ok" | "warn" | "error"; text: string } | null;

export default function TemplateEditor() {
  const [api, setApi] = useState<DocxViewApi | null>(null);
  const [source, setSource] = useState<string | File>(STARTER_TEMPLATE);
  const [docKey, setDocKey] = useState(0); // bump to force a clean remount of the editor
  const [matterIndex, setMatterIndex] = useState(0);
  const [notice, setNotice] = useState<Notice>(null);
  const [missingFonts, setMissingFonts] = useState<string[]>([]);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const matter = matters[matterIndex];
  const variableKeys = useMemo(() => flattenKeys(matter), [matter]);

  const openSource = useCallback((next: string | File) => {
    setApi(null); // the old api is invalid once the editor unmounts
    setSource(next);
    setDocKey((k) => k + 1);
    setNotice(null);
  }, []);

  const handleExport = useCallback(() => {
    if (!api) return;
    try {
      const templateBytes = api.save(); // current editor content, {{vars}} still in place
      const { bytes, unresolved } = fillTemplate(templateBytes, matter);

      const name = `Answer_${safeFileName(matter.case.number)}.docx`;
      downloadBytes(bytes, name);

      setNotice(
        unresolved.length
          ? {
              kind: "warn",
              text: `Exported ${name}, but these tags had no data and were left as-is: ${unresolved
                .map((u) => `{{${u}}}`)
                .join(", ")}`,
            }
          : {
              kind: "ok",
              text: `Exported ${name} with ${matterLabel(matterIndex)} data.`,
            },
      );
    } catch (err) {
      if (err instanceof TemplateError) {
        setNotice({
          kind: "error",
          text: [err.message, ...err.details].join(" — "),
        });
      } else {
        setNotice({ kind: "error", text: (err as Error).message });
      }
    }
  }, [api, matter, matterIndex]);

  // Toolbar "Download" saves the raw template (variables intact), not the filled version.
  const saveTemplate = useCallback((bytes: Uint8Array) => {
    downloadBytes(bytes, "template.docx");
  }, []);

  const copyTag = (key: string) => {
    navigator.clipboard?.writeText(`{{${key}}}`);
    setNotice({
      kind: "ok",
      text: `Copied {{${key}}} — paste it into the document.`,
    });
  };

  const insertTag = (key: string) => {
    const tag = `{{${key}}}`;

    // Inserts at the caret (or replaces the current selection), like typing.
    if (api?.insertSymbol(tag)) {
      setNotice(null);
      return;
    }

    // false = the editor has no caret yet (user hasn't clicked into the document)
    navigator.clipboard?.writeText(tag);
    setNotice({
      kind: "warn",
      text: `Click inside the document to place the cursor first. ${tag} was copied to your clipboard instead.`,
    });
  };

  return (
    <div style={styles.page}>
      <div style={styles.bar}>
        <label style={styles.label}>
          Matter
          <select
            value={matterIndex}
            onChange={(e) => setMatterIndex(Number(e.target.value))}
            style={styles.select}
          >
            {matters.map((m, i) => (
              <option key={m.case.number} value={i}>
                {matterLabel(i)} — {m.case.number}
              </option>
            ))}
          </select>
        </label>

        <button onClick={handleExport} disabled={!api} style={styles.primary}>
          Export .docx
        </button>

        <span style={{ flex: 1 }} />

        <button
          onClick={() => fileInputRef.current?.click()}
          style={styles.secondary}
        >
          Open .docx…
        </button>
        <button
          onClick={() => openSource(STARTER_TEMPLATE)}
          style={styles.secondary}
        >
          Reset to starter
        </button>
        <input
          ref={fileInputRef}
          type="file"
          accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          hidden
          onChange={(e) => {
            const file = e.target.files?.[0];
            if (file) openSource(file);
            e.target.value = "";
          }}
        />
      </div>

      <div style={styles.vars}>
        <span style={styles.muted}>Tags (click to insert at cursor):</span>
        {variableKeys.map((k) => (
          <button
            key={k}
            title="Insert at cursor"
            disabled={!api}
            // Keeps focus and the caret in the editor. Without this the click steals it.
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => insertTag(k)}
            style={styles.chip}
          >
            {`{{${k}}}`}
          </button>
        ))}
      </div>

      {notice && (
        <div
          role="status"
          style={{ ...styles.notice, ...styles[`notice_${notice.kind}`] }}
        >
          {notice.text}
        </div>
      )}
      {missingFonts.length > 0 && (
        <div style={{ ...styles.notice, ...styles.notice_warn }}>
          Fonts not installed here (rendered with a substitute, layout may
          differ from Word): {missingFonts.join(", ")}
        </div>
      )}

      {api && <DocxToolbar api={api} mode="advanced" onSave={saveTemplate} />}

      <DocxView
        key={docKey}
        source={source}
        editable
        commentAuthor="Ahmad"
        onReady={setApi}
        onMissingFonts={(m) => setMissingFonts(m.map((f) => f.family))}
        onError={(e) =>
          setNotice({
            kind: "error",
            text: `Could not open document: ${e.message}`,
          })
        }
        style={{ flex: 1, minHeight: 0 }}
      />
    </div>
  );
}

const styles = {
  page: { display: "flex", flexDirection: "column", height: "100vh" },
  bar: {
    display: "flex",
    alignItems: "center",
    gap: 12,
    padding: "10px 16px",
    background: "#fff",
    borderBottom: "1px solid #e5e7eb",
  },
  label: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    fontSize: 14,
    fontWeight: 600,
    color: "#000",
  },
  select: {
    padding: "6px 10px",
    fontSize: 14,
    borderRadius: 6,
    border: "1px solid #d1d5db",
  },
  primary: {
    padding: "7px 16px",
    fontSize: 14,
    fontWeight: 600,
    color: "#fff",
    background: "#1a73e8",
    border: 0,
    borderRadius: 6,
    cursor: "pointer",
  },
  secondary: {
    padding: "6px 12px",
    fontSize: 13,
    background: "#fff",
    border: "1px solid #d1d5db",
    borderRadius: 6,
    cursor: "pointer",
  },
  vars: {
    display: "flex",
    flexWrap: "wrap",
    alignItems: "center",
    gap: 6,
    padding: "8px 16px",
    background: "#f9fafb",
    color: "#000",
    borderBottom: "1px solid #e5e7eb",
  },
  muted: { fontSize: 12, color: "#6b7280", marginRight: 4 },
  chip: {
    fontFamily: "ui-monospace, monospace",
    fontSize: 12,
    padding: "2px 8px",
    background: "#eef2ff",
    border: "1px solid #c7d2fe",
    borderRadius: 999,
    cursor: "pointer",
  },
  notice: { padding: "8px 16px", fontSize: 13 },
  notice_ok: { background: "#ecfdf5", color: "#065f46" },
  notice_warn: { background: "#fffbeb", color: "#92400e" },
  notice_error: { background: "#fef2f2", color: "#991b1b" },
} satisfies Record<string, React.CSSProperties>;
