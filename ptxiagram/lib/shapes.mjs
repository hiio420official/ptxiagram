// Semantic node types -> native PPTX shape preset + color pair.
// Shape preset strings are raw OOXML preset geometry names (pptxgenjs
// accepts these directly, no need to import the ShapeType enum).
export const NODE_TYPES = {
  client: { shape: "ellipse", fill: "F1F5F9", stroke: "94A3B8" },
  frontend: { shape: "rect", fill: "E2E8F0", stroke: "94A3B8" },
  neutral: { shape: "roundRect", fill: "E2E8F0", stroke: "94A3B8" },
  backend: { shape: "roundRect", fill: "D1FAE5", stroke: "10B981" },
  database: { shape: "can", fill: "DBEAFE", stroke: "3B82F6" },
  cache: { shape: "hexagon", fill: "DBEAFE", stroke: "3B82F6" },
  queue: { shape: "cube", fill: "DBEAFE", stroke: "3B82F6" },
  external: { shape: "cloud", fill: "FEF3C7", stroke: "F59E0B" },
  security: { shape: "roundRect", fill: "FFE4E6", stroke: "F43F5E" },
  decision: { shape: "diamond", fill: "FFE4E6", stroke: "F43F5E" },
};

export const FONT = "맑은 고딕";

// Edge/connection variant -> color + dash pair, so JSON authors pick a
// semantic variant instead of guessing hex codes.
export const EDGE_VARIANTS = {
  default: { color: "94A3B8", dashed: false },
  emphasis: { color: "10B981", dashed: false },
  security: { color: "F43F5E", dashed: true },
  dashed: { color: "64748B", dashed: true },
};

// Draws a labeled node and returns its box {x,y,w,h} for connector anchoring.
export function drawNode(slide, { x, y, w = 1.8, h = 0.95, label, sublabel, type = "neutral", tag }) {
  const style = NODE_TYPES[type] || NODE_TYPES.neutral;
  slide.addShape(style.shape, {
    x, y, w, h,
    fill: { color: style.fill },
    line: { color: style.stroke, width: 1.5 },
    shadow: { type: "outer", color: "1E293B", opacity: 0.18, blur: 4, offset: 1, angle: 90 },
  });

  const hasSub = Boolean(sublabel);
  slide.addText(label, {
    x, y: y + h / 2 - (hasSub ? 0.28 : 0.14), w, h: 0.32,
    align: "center", valign: "middle", fontFace: FONT, fontSize: 12.5, bold: true, color: "0F172A",
  });
  if (hasSub) {
    slide.addText(sublabel, {
      x, y: y + h / 2 + 0.08, w, h: 0.3,
      align: "center", fontFace: FONT, fontSize: 9, color: "64748B",
    });
  }
  if (tag) {
    slide.addText(tag, {
      x, y: y + h - 0.24, w, h: 0.22,
      align: "center", fontFace: FONT, fontSize: 7.5, color: style.stroke,
    });
  }
  return { x, y, w, h };
}

// Draws a dashed boundary frame (precomputed by lib/layout.mjs's
// computeArchitectureLayout, which owns the padding math) with a label in
// its top-left corner. Mirrors archify's `wraps` boundary concept.
export function drawBoundary(slide, { box, label, color = "F59E0B" }) {
  slide.addShape("rect", {
    ...box,
    fill: { type: "none" },
    line: { color, width: 1.25, dashType: "dash" },
  });
  slide.addText(label, {
    x: box.x + 0.15, y: box.y + 0.1, w: box.w - 0.3, h: 0.3,
    fontFace: FONT, fontSize: 10, color, bold: true,
  });
}

// Draws a row of summary cards (title + bullet items) spanning the slide
// width, starting at the given y. Used for footnotes/exception details that
// would otherwise need extra nodes and crossing arrows.
export function drawCards(slide, cards, y) {
  const totalW = 12.5, gap = 0.25;
  const w = (totalW - gap * (cards.length - 1)) / cards.length;
  cards.forEach((card, i) => {
    const x = 0.4 + i * (w + gap);
    const color = card.color ? `#${card.color}` : "3B82F6";
    slide.addText(card.title, {
      x, y, w, h: 0.3,
      fontFace: FONT, fontSize: 11, bold: true, color: "0F172A",
    });
    const bulletText = card.items.map((t) => `• ${t}`).join("\n");
    slide.addText(bulletText, {
      x, y: y + 0.35, w, h: 1.3,
      fontFace: FONT, fontSize: 9, color: "475569", valign: "top", lineSpacing: 14,
    });
    slide.addShape("rect", { x, y: y + 0.02, w: 0.05, h: 0.22, fill: { color: card.color || "3B82F6" }, line: { type: "none" } });
  });
}

// Draws a swimlane frame with a title strip. Returns the lane's box.
export function drawLane(slide, { x, y, w, h, label }) {
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: "F8FAFC" },
    line: { color: "CBD5E1", width: 1, dashType: "dash" },
  });
  slide.addText(label, {
    x: x + 0.15, y: y + 0.05, w: 3, h: 0.3,
    fontFace: FONT, fontSize: 11, color: "64748B", bold: true,
  });
  return { x, y, w, h };
}
