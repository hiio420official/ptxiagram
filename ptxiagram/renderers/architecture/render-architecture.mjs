import pptxgen from "pptxgenjs";
import fs from "node:fs/promises";
import { drawNode, drawBoundary, drawCards, FONT, EDGE_VARIANTS } from "../../lib/shapes.mjs";
import { connector } from "../../lib/routes.mjs";
import { repairPptx } from "../../lib/repair-pptx.mjs";
import { validateSchema, validateArchitectureLayout } from "../../lib/validate.mjs";
import { computeArchitectureLayout } from "../../lib/layout.mjs";

export async function renderArchitecture(inputPath, outputPath) {
  const raw = JSON.parse(await fs.readFile(inputPath, "utf8"));
  await validateSchema("architecture", raw);

  const { boxes, boundaryBoxes } = computeArchitectureLayout(raw);
  const problems = validateArchitectureLayout(raw);
  if (problems.length) {
    throw new Error(`Architecture layout validation failed:\n- ${problems.join("\n- ")}`);
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

  // boundaries first so they sit visually under the components they frame
  for (const b of boundaryBoxes) {
    drawBoundary(slide, { box: b.box, label: b.label, ...(b.color ? { color: b.color } : {}) });
  }

  for (const c of raw.components) {
    drawNode(slide, { ...boxes[c.id], label: c.label, sublabel: c.sublabel, type: c.type, tag: c.tag });
  }

  for (const conn of raw.connections) {
    const variant = EDGE_VARIANTS[conn.variant || "default"];
    const from = boxes[conn.from], to = boxes[conn.to];
    if (!from) throw new Error(`Connection references unknown component "${conn.from}"`);
    if (!to) throw new Error(`Connection references unknown component "${conn.to}"`);
    connector(slide, from, to, {
      color: variant.color, dashed: variant.dashed, label: conn.label,
      route: conn.route, bias: conn.bias, clearY: conn.clearY,
    });
  }

  if (raw.cards && raw.cards.length) {
    drawCards(slide, raw.cards, 6.0);
  }

  slide.addText("Made with a Claude Skill · native PPTX shapes, editable in PowerPoint", {
    x: 0.4, y: 7.15, w: 12.5, h: 0.3,
    fontFace: FONT, fontSize: 8, color: "94A3B8", italic: true,
  });

  const outFile = outputPath || raw.meta.output || "output.pptx";
  await pptx.writeFile({ fileName: outFile });
  await repairPptx(outFile);
  return outFile;
}
