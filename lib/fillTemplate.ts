import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";

export type TemplateData = Record<string, unknown>;

export interface FillResult {
  bytes: Uint8Array;
  /** Tags in the document with no value in `data` (left as-is in the output). */
  unresolved: string[];
}

/** Thrown for malformed templates, e.g. an unclosed "{{". */
export class TemplateError extends Error {
  details: string[];
  constructor(message: string, details: string[] = []) {
    super(message);
    this.name = "TemplateError";
    this.details = details;
  }
}

/** Resolve "a.b.c" against nested data. "." returns the current scope (used inside loops). */
const dotPathParser = (tag: string) => ({
  get(scope: unknown) {
    const path = tag.trim();
    if (path === ".") return scope;
    return path
      .split(".")
      .reduce<unknown>(
        (acc, key) =>
          acc == null ? undefined : (acc as Record<string, unknown>)[key],
        scope,
      );
  },
});

export function fillTemplate(
  templateBytes: Uint8Array,
  data: TemplateData,
): FillResult {
  const unresolved = new Set<string>();

  try {
    const zip = new PizZip(templateBytes);
    const doc = new Docxtemplater(zip, {
      delimiters: { start: "{{", end: "}}" },
      paragraphLoop: true,
      linebreaks: true,
      parser: dotPathParser,
      nullGetter(part: { value?: string }) {
        const name = String(part.value ?? "");
        unresolved.add(name);
        return `{{${name}}}`;
      },
    });

    doc.render(data);

    const bytes = doc.getZip().generate({
      type: "uint8array",
      compression: "DEFLATE",
    }) as Uint8Array;

    return { bytes, unresolved: [...unresolved] };
  } catch (err) {
    const e = err as {
      message?: string;
      properties?: {
        errors?: Array<{
          message?: string;
          properties?: { explanation?: string };
        }>;
        explanation?: string;
      };
    };
    const details =
      e.properties?.errors?.map(
        (x) =>
          x.properties?.explanation ?? x.message ?? "Unknown template error",
      ) ?? (e.properties?.explanation ? [e.properties.explanation] : []);
    throw new TemplateError(e.message ?? "Failed to fill template", details);
  }
}

export function downloadBytes(bytes: Uint8Array, fileName: string) {
  const blob = new Blob([bytes as BlobPart], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
