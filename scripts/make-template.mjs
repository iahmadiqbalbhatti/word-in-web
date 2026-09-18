import { Document, Packer, Paragraph, TextRun, AlignmentType } from "docx";
import { mkdirSync, writeFileSync } from "node:fs";

const run = (text, opts = {}) =>
  new TextRun({ text, font: "Times New Roman", size: 24, ...opts });
const p = (child, opts = {}) =>
  new Paragraph({ spacing: { after: 200 }, children: [child], ...opts });
const center = { alignment: AlignmentType.CENTER };

const doc = new Document({
  sections: [
    {
      children: [
        p(run("SUPERIOR COURT OF CALIFORNIA", { bold: true }), { ...center, spacing: { after: 0 } }),
        p(run("COUNTY OF {{court.county}}", { bold: true }), center),
        p(run("{{plaintiff.name_caps}},")),
        p(run("Plaintiff,"), { indent: { left: 720 } }),
        p(run("v.")),
        p(run("{{defendant.name_caps}},")),
        p(run("Defendant(s)."), { indent: { left: 720 } }),
        p(run("Case No. {{case.number}}")),
        p(run("DEFENDANT {{defendant.name_caps}}'S ANSWER", { bold: true }), { ...center, spacing: { after: 0 } }),
        p(run("TO PLAINTIFF {{plaintiff.name_caps}}'S COMPLAINT", { bold: true }), center),
      ],
    },
  ],
});

mkdirSync("public/templates", { recursive: true });
writeFileSync("public/templates/answer-template.docx", await Packer.toBuffer(doc));
console.log("Wrote public/templates/answer-template.docx");