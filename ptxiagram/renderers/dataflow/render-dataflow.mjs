import pptxgen from "pptxgenjs";
import fs from "node:fs/promises";
import path from "node:path";
import { drawNode, drawCards, FONT, EDGE_VARIANTS } from "../../lib/shapes.mjs";
import { connector } from "../../lib/routes.mjs";
import { repairPptx } from "../../lib/repair-pptx.mjs";
import { validateSchema, validateDataflowLayout } from "../../lib/validate.mjs";
import { computeDataflowLayout, DATAFLOW_LAYOUT, sameRow } from "../../lib/layout.mjs";

export async function renderDataflow(inputPath, outputPath) {
  const raw = JSON.parse(await fs.readFile(inputPath, "utf8"));
  await validateSchema("dataflow", raw);

  const { nodeBoxes, stageCenterX } = computeDataflowLayout(raw);
  const problems = validateDataflowLayout(raw);
  if (problems.length) {
    throw new Error(`Dataflow layout validation failed:\n- ${problems.join("\n- ")}`);
  }

  const pptx = new pptxgen();
  pptx.defineLayout({ name: "WIDE", width: 13.33, height: 7.5 });
  pptx.layout = "WIDE";
  const slide = pptx.addSlide();

  slide.addText(raw.meta.title, {
    x: 0.4, y: 0.2, w: 12.5, h: 0.5,
    fontFace: FONT, fontSize: 20, bold: true, color: "0F172A",
  });
  if (raw.meta.subtitle) {
    slide.addText(raw.meta.subtitle, {
      x: 0.4, y: 0.65, w: 12.5, h: 0.3,
      fontFace: FONT, fontSize: 11, color: "64748B",
    });
  }

  raw.stages.forEach((s, i) => {
    slide.addText(s.label, {
      x: stageCenterX[i] - 1.0, y: DATAFLOW_LAYOUT.STAGE_LABEL_Y, w: 2.0, h: 0.3,
      align: "center", fontFace: FONT, fontSize: 11, bold: true, color: "64748B",
    });
  });

  for (const n of raw.nodes) {
    drawNode(slide, { ...nodeBoxes[n.id], label: n.label, sublabel: n.sublabel, type: n.type, tag: n.tag });
  }

  for (const f of raw.flows) {
    const variant = EDGE_VARIANTS[f.variant || "default"];
    const from = nodeBoxes[f.from], to = nodeBoxes[f.to];
    if (!from) throw new Error(`Flow references unknown node "${f.from}"`);
    if (!to) throw new Error(`Flow references unknown node "${f.to}"`);
    const label = f.classification ? `${f.label} (${f.classification})` : f.label;
    connector(slide, from, to, {
      color: variant.color, dashed: variant.dashed, label,
      route: sameRow(from, to) ? "straight" : "elbow-right",
    });
  }

  if (raw.cards && raw.cards.length) {
    drawCards(slide, raw.cards, 6.5);
  }

  slide.addText("Made with a Claude Skill · native PPTX shapes, editable in PowerPoint", {
    x: 0.4, y: 7.15, w: 12.5, h: 0.3,
    fontFace: FONT, fontSize: 8, color: "94A3B8", italic: true,
  });

  const outFile = outputPath || raw.meta.output || "output.pptx";
  await fs.mkdir(path.dirname(outFile), { recursive: true });
  await pptx.writeFile({ fileName: outFile });
  await repairPptx(outFile);
  return outFile;
}
