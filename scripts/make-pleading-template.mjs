// node apps/web/scripts/make-pleading-template.mjs # writes public/templates/make-pleading-template.docx
// Now apps/web/public/templates/make-pleading-template.docx is generated
// we can use it as a template for the pleading

import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  Header,
  Footer,
  AlignmentType,
  BorderStyle,
  WidthType,
  HeightRule,
  LineRuleType,
  PageBorderDisplay,
  PageBorderOffsetFrom,
  PageBorderZOrder,
  PageNumber,
  WpsShapeRun,
  HorizontalPositionRelativeFrom,
  VerticalPositionRelativeFrom,
  TextWrappingType,
} from "docx";
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

/* ---------- California pleading geometry (CRC 2.100-2.119) ----------
   Page: Letter, 8.5in x 11in = 12240 x 15840 twips (612pt x 792pt).

   Line pitch: 28 numbered lines at exactly 24pt each (12pt font,
   double-spaced) = 672pt of body text height, fixed and non-negotiable
   (this is what makes text line N land beside printed number N).

   MARGIN MATH - why bottom is NOT a literal 1.0in
   ------------------------------------------------
   672pt (body) + top margin + bottom margin must total <= 792pt (page).
   That leaves only 792 - 672 = 120pt (1.667in) to split between the
   top and bottom margins combined.

   Top margin here is fixed at the full CRC-standard 72pt (1.0in), which
   leaves at most 120 - 72 = 48pt (0.667in) for the bottom margin - not
   72pt (1.0in). A literal 1.0in bottom margin, combined with a full
   1.0in top margin, would require 672 + 72 + 72 = 816pt of page height,
   24pt (one whole numbered line) more than an 11in page actually has.

   So: bottom margin is set to 47pt (940 twips) here - the largest value
   that still fits all 28 lines, minus 1pt of slack so renderers that add
   a hair of extra height (e.g. LibreOffice's table borders add ~0.75pt)
   don't push line 28 onto a second page. This is standard practice for
   28-line/24pt numbered pleading paper: top stays at a full 1in, and the
   bottom margin absorbs the shortfall. If a literal 1.0in bottom margin
   is a hard requirement, the only way to get it (given top = 1in and
   24pt/line) is to drop to 27 numbered lines instead of 28 - ask if you
   want that variant instead.

   ONE-LINE-PER-NUMBER RULE
   Every line of body text (attorney block, court title, caption, body,
   signature block) is set at EXACTLY 24pt, the same pitch as the line
   numbers, so text line N always sits beside printed number N.
   Never use a line height that is not a multiple of 24pt in the body. */
const TWIP = 20;
const LINE = 24 * TWIP; // 480 : one numbered line (exact 24pt)
const LINES = 28;
const MARGIN = { top: 1440, bottom: 940, left: 2160, right: 720 }; // 1in / 47pt (see math above) / 1.5in / 0.5in
const TEXT_W = 12240 - MARGIN.left - MARGIN.right; // 9360

const FONT = "Times New Roman";
const run = (text, o = {}) => new TextRun({ text, font: FONT, size: 24, ...o });
const exact = (line) => ({
  before: 0,
  after: 0,
  line,
  lineRule: LineRuleType.EXACT,
});

/** One numbered line of text (exact 24pt). The ONLY paragraph builder used in the body. */
const ln = (text, o = {}, p = {}) =>
  new Paragraph({ spacing: exact(LINE), children: [run(text, o)], ...p });
/** One empty numbered line. */
const blank = (p = {}) =>
  new Paragraph({ spacing: exact(LINE), children: [run("")], ...p });

const center = { alignment: AlignmentType.CENTER };
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const rule = { style: BorderStyle.SINGLE, size: 6, color: "000000" };

/* ---------- Line numbers 1-28: floating DrawingML text box in the header ----------
   Anchored to the page (not the text), so it repeats on every page and always
   sits on the same 28-line grid, whatever is typed in the body.
   Its top edge = the body's top margin, and it uses the same exact-24pt
   paragraphs as the body, so number N and body line N share one baseline.

   COMPATIBILITY NOTE: WpsShapeRun (floating text box) is the one element in
   this file most likely to render differently across apps. It is well
   supported in Word and LibreOffice. Google Docs' .docx importer has a
   narrower OOXML feature set and has historically been inconsistent with
   floating/behindDocument shapes - always spot-check the generated file by
   opening it directly in Google Docs before relying on it there. If the
   line numbers don't survive Google Docs import, that importer gap - not
   this code - is the cause; there is no way to guarantee it without an
   actual render check on the current Google Docs importer. */
const EMU = 12700; // 1pt in EMU
const PX = 0.75; // docx-js transformation units are px (1px = .75pt)
const lineNumbers = new Paragraph({
  spacing: exact(12 * TWIP),
  children: [
    new WpsShapeRun({
      type: "wps",
      transformation: { width: 54 / PX, height: (LINES * 24) / PX },
      floating: {
        horizontalPosition: {
          relative: HorizontalPositionRelativeFrom.PAGE,
          offset: 36 * EMU,
        },
        verticalPosition: {
          relative: VerticalPositionRelativeFrom.PAGE,
          offset: (MARGIN.top / TWIP) * EMU,
        },
        behindDocument: true,
        allowOverlap: true,
        wrap: { type: TextWrappingType.NONE },
      },
      bodyProperties: {
        margins: { top: 0, bottom: 0, left: 0, right: 0 },
        noAutoFit: true,
      },
      children: Array.from(
        { length: LINES },
        (_, i) =>
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            spacing: exact(LINE),
            children: [run(String(i + 1))],
          }),
      ),
    }),
  ],
});

/* ---------- Footer: page number, rule, document title (CRC 2.110) ---------- */
const small = (t) => run(t, { size: 20 });
const footer = new Footer({
  children: [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: exact(12 * TWIP),
      border: { bottom: rule },
      children: [
        new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 20 }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: exact(12 * TWIP),
      children: [
        small(
          "DEFENDANT {defendant.name_caps}'S ANSWER TO PLAINTIFF'S COMPLAINT",
        ),
      ],
    }),
  ],
});

/* ---------- Caption table (party block | case no. + title) ----------
   Every paragraph inside is a 24pt `ln` line. The row height is EXACT and a
   whole number of lines, so the table can never push later text off the grid
   (table borders otherwise add a fraction of a point and shift everything
   below it). If a very long party name / title wraps past CAPTION_LINES,
   raise the constant by 1.                                                  */
const CAPTION_LINES = 8;
const cell = (width, children, borders) =>
  new TableCell({
    width: { size: width, type: WidthType.DXA },
    margins: { top: 0, bottom: 0, left: 100, right: 100 },
    borders: { top: none, left: none, bottom: none, right: none, ...borders },
    children,
  });
const LEFT_W = 4680,
  RIGHT_W = TEXT_W - LEFT_W;
const caption = new Table({
  width: { size: TEXT_W, type: WidthType.DXA },
  columnWidths: [LEFT_W, RIGHT_W],
  borders: {
    top: none,
    bottom: none,
    left: none,
    right: none,
    insideHorizontal: none,
    insideVertical: none,
  },
  rows: [
    new TableRow({
      cantSplit: true,
      height: { value: CAPTION_LINES * LINE, rule: HeightRule.EXACT },
      children: [
        cell(
          LEFT_W,
          [
            ln("{plaintiff.name_caps},"),
            ln("Plaintiff,", {}, { indent: { left: 1440 } }),
            ln("v.", {}, { indent: { left: 720 } }),
            ln("{defendant.name_caps},"),
            ln("Defendant(s).", {}, { indent: { left: 1440 } }),
          ],
          { right: rule, bottom: rule },
        ),
        cell(
          RIGHT_W,
          [
            ln("Case No. {matter.case_number}"),
            ln(
              "DEFENDANT {defendant.name_caps}'S ANSWER TO PLAINTIFF {plaintiff.name_caps}'S COMPLAINT",
              { bold: true },
            ),
          ],
          {},
        ),
      ],
    }),
  ],
});

/* ---------- Document ---------- */
const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: FONT, size: 24 },
        paragraph: { spacing: exact(LINE) },
      },
    },
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 }, // Letter: 8.5in x 11in
          margin: { ...MARGIN, header: 720, footer: 360 },
          borders: {
            // double rule left of the text, single rule right of it, full body height
            // NOTE: page borders are a known soft spot for Google Docs' .docx
            // importer - verify visually there; if dropped, that's the
            // importer's OOXML support gap, not a config error here.
            pageBorders: {
              display: PageBorderDisplay.ALL_PAGES,
              offsetFrom: PageBorderOffsetFrom.TEXT,
              zOrder: PageBorderZOrder.BACK,
            },
            pageBorderLeft: {
              style: BorderStyle.DOUBLE,
              size: 6,
              color: "000000",
              space: 8,
            },
            pageBorderRight: {
              style: BorderStyle.SINGLE,
              size: 6,
              color: "000000",
              space: 6,
            },
          },
        },
      },
      headers: { default: new Header({ children: [lineNumbers] }) },
      footers: { default: footer },
      children: [
        // Lines 1-7: attorney block, one text line per numbered line
        ln("{attorney.name} (SBN {plaintiff_attorney.bar})"),
        ln("{plaintiff_attorney.email}"),
        ln("Telephone: {plaintiff_attorney.phone}"),
        ln("Email: {attorney.email}"),
        ln("Attorneys for Defendant {defendant.name}"),

        // Lines 8-9: court title, line 10 spacer
        ln("SUPERIOR COURT OF THE STATE OF CALIFORNIA", { bold: true }, center),
        ln("COUNTY OF {court.county}", { bold: true }, center),
        blank(),

        // Lines 11-18: caption
        caption,

        // Line 19 spacer, then the body starts on line 20
        blank(),
        ln(
          "Defendant {defendant.name_caps} answers Plaintiff {plaintiff.name_caps}'s Complaint as follows:",
          {},
          { indent: { firstLine: 0 } },
        ),
        blank(),

        // Signature block: every line is still one numbered line
        ln("Dated: {matter.created_at}", {}, { keepNext: true }),
        blank({ keepNext: true }),
        ln(
          "By: ______________________________",
          {},
          { indent: { left: 4680 }, keepNext: true },
        ),
        ln("{attorney.name}", {}, { indent: { left: 5040 }, keepNext: true }),
        ln(
          "Attorneys for Defendant ({defendant.name_caps})",
          {},
          { indent: { left: 5040 } },
        ),
      ],
    },
  ],
});

// Resolve paths from the script's own location, so it works from any directory
const scriptDir = dirname(fileURLToPath(import.meta.url)); // apps/web/scripts
const outDir = resolve(scriptDir, "../public/templates"); // apps/web/public/templates
const outFile = resolve(outDir, "make-pleading-template.docx");

mkdirSync(outDir, { recursive: true });
writeFileSync(outFile, await Packer.toBuffer(doc));
console.log("Wrote", outFile);
