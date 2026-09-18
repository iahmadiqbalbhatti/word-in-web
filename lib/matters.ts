import raw from "@/data/matters.json";

export type Matter = (typeof raw)[number];
export const matters: Matter[] = raw;

const WORDS = [
  "One",
  "Two",
  "Three",
  "Four",
  "Five",
  "Six",
  "Seven",
  "Eight",
  "Nine",
  "Ten",
];

export const matterLabel = (index: number) =>
  `Matter ${WORDS[index] ?? index + 1}`;

/** { court: { county } } -> ["court.county"] — the tags you can type into the doc. */
export function flattenKeys(obj: unknown, prefix = ""): string[] {
  if (obj === null || typeof obj !== "object" || Array.isArray(obj)) return [];
  return Object.entries(obj as Record<string, unknown>).flatMap(([k, v]) => {
    const path = prefix ? `${prefix}.${k}` : k;
    return v !== null && typeof v === "object" && !Array.isArray(v)
      ? flattenKeys(v, path)
      : [path];
  });
}

export const safeFileName = (s: string) =>
  s.replace(/[^\w.\-]+/g, "_").replace(/^_+|_+$/g, "");
