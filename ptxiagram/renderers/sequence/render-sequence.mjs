import pptxgen from "pptxgenjs";
import fs from "node:fs/promises";
import path from "node:path";
import { drawNode, drawLifeline, drawCards, FONT, EDGE_VARIANTS } from "../../lib/shapes.mjs";
import { drawArrow } from "../../lib/routes.mjs";
import { repairPptx } from "../../lib/repair-pptx.mjs";
import { validateSchema, validateSequenceLayout } from "../../lib/validate.mjs";
import { computeSequenceLayout } from "../../lib/layout.mjs";

export async function renderSequence(inputPath, outputPath) {
  const raw = JSON.parse(await fs.readFile(inputPath, "utf8"));
  await validateSchema("sequence", raw);

  const { participantBoxes, centerX, messageYs, lifelineBottom, cardsY } = computeSequenceLayout(raw);
  const problems = validateSequenceLayout(raw);
  if (problems.length) {
    throw new Error(`Sequence layout validation failed:\n- ${problems.join("\n- ")}`);
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

  for (const p of raw.participants) {
    drawNode(slide, { ...participantBoxes[p.id], label: p.label, sublabel: p.sublabel, type: p.type });
    drawLifeline(slide, {
      x: centerX[p.id],
      fromY: participantBoxes[p.id].y + participantBoxes[p.id].h,
      toY: lifelineBottom,
    });
  }

  raw.messages.forEach((m, i) => {
    const variant = EDGE_VARIANTS[m.variant || "default"];
    drawArrow(slide, {
      x1: centerX[m.from], x2: centerX[m.to], y: messageYs[i],
      color: variant.color, dashed: variant.dashed, label: m.label,
    });
  });

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
