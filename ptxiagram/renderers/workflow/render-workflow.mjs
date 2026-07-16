import pptxgen from "pptxgenjs";
import fs from "node:fs/promises";
import path from "node:path";
import { drawNode, drawLane, drawCards, FONT, EDGE_VARIANTS } from "../../lib/shapes.mjs";
import { connector } from "../../lib/routes.mjs";
import { repairPptx } from "../../lib/repair-pptx.mjs";
import { validateSchema, validateWorkflowLayout } from "../../lib/validate.mjs";
import { computeWorkflowLayout } from "../../lib/layout.mjs";

export async function renderWorkflow(inputPath, outputPath) {
  const raw = JSON.parse(await fs.readFile(inputPath, "utf8"));
  await validateSchema("workflow", raw);

  const { laneBoxes, nodeBoxes, cardsY } = computeWorkflowLayout(raw);
  const problems = validateWorkflowLayout(raw);
  if (problems.length) {
    throw new Error(`Workflow layout validation failed:\n- ${problems.join("\n- ")}`);
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

  for (const lane of raw.lanes) {
    drawLane(slide, { ...laneBoxes[lane.id], label: lane.label });
  }

  for (const n of raw.nodes) {
    drawNode(slide, { ...nodeBoxes[n.id], label: n.label, sublabel: n.sublabel, type: n.type, tag: n.tag });
  }

  for (const e of raw.edges) {
    const variant = EDGE_VARIANTS[e.variant || "default"];
    const from = nodeBoxes[e.from], to = nodeBoxes[e.to];
    if (!from) throw new Error(`Edge references unknown node "${e.from}"`);
    if (!to) throw new Error(`Edge references unknown node "${e.to}"`);
    connector(slide, from, to, {
      color: variant.color, dashed: variant.dashed,
      label: e.label, route: e.route || "straight", bias: e.bias,
    });
  }

  if (raw.cards && raw.cards.length) {
    drawCards(slide, raw.cards, cardsY);
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
