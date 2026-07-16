import { FONT } from "./shapes.mjs";

export function centerTop(n) { return { x: n.x + n.w / 2, y: n.y }; }
export function centerBottom(n) { return { x: n.x + n.w / 2, y: n.y + n.h }; }
export function centerLeft(n) { return { x: n.x, y: n.y + n.h / 2 }; }
export function centerRight(n) { return { x: n.x + n.w, y: n.y + n.h / 2 }; }

// Normalizes a point-to-point line into a {x,y,w,h} box (pptxgenjs "line"
// shapes are a bounding box + direction, so a purely vertical or horizontal
// segment always has w=0 or h=0 here).
function boxFromPoints(p1, p2) {
  return { x: p1.x, y: p1.y, w: p2.x - p1.x, h: p2.y - p1.y };
}

/**
 * Pure geometry: returns the ordered line segments (each an arrow-less
 * {x,y,w,h,arrow} box, `arrow: true` only on the final one) and the label
 * anchor a connector between `from` and `to` would use, without touching a
 * slide. The renderer draws these; the validator walks them to check for
 * crossings against unrelated nodes. See routes' JSDoc in `connector()`
 * below for what each route means.
 */
export function computeSegments(from, to, opts = {}) {
  const { route = "straight", bias = 0.5, clearY, label } = opts;

  if (route === "straight") {
    const a = centerRight(from), b = centerLeft(to);
    return {
      segments: [{ ...boxFromPoints(a, b), arrow: true }],
      labelAnchor: label ? { x: a.x, y: a.y, w: b.x - a.x } : null,
    };
  }

  if (route === "drop") {
    const a = centerBottom(from), b = centerTop(to);
    if (Math.abs(a.x - b.x) < 0.05) {
      return {
        segments: [{ ...boxFromPoints(a, { x: a.x, y: b.y }), arrow: true }],
        labelAnchor: label ? { x: a.x - 0.6, y: (a.y + b.y) / 2, w: 1.2 } : null,
      };
    }
    const midY = a.y + (b.y - a.y) * bias;
    return {
      segments: [
        { ...boxFromPoints(a, { x: a.x, y: midY }), arrow: false },
        { ...boxFromPoints({ x: Math.min(a.x, b.x), y: midY }, { x: Math.max(a.x, b.x), y: midY }), arrow: false },
        { ...boxFromPoints({ x: b.x, y: midY }, b), arrow: true },
      ],
      labelAnchor: label ? { x: Math.min(a.x, b.x) - 0.6, y: midY, w: Math.abs(b.x - a.x) + 1.2 } : null,
    };
  }

  if (route === "elbow-right") {
    const a = centerRight(from), b = centerLeft(to);
    if (Math.abs(a.y - b.y) < 0.05) {
      return {
        segments: [{ ...boxFromPoints(a, { x: b.x, y: a.y }), arrow: true }],
        labelAnchor: label ? { x: a.x, y: a.y, w: b.x - a.x } : null,
      };
    }
    const midX = a.x + (b.x - a.x) * bias;
    return {
      segments: [
        { ...boxFromPoints(a, { x: midX, y: a.y }), arrow: false },
        { ...boxFromPoints({ x: midX, y: Math.min(a.y, b.y) }, { x: midX, y: Math.max(a.y, b.y) }), arrow: false },
        { ...boxFromPoints({ x: midX, y: b.y }, b), arrow: true },
      ],
      labelAnchor: label ? { x: midX - 0.6, y: b.y, w: 1.2 } : null,
    };
  }

  if (route === "arc-over-top") {
    if (clearY === undefined) throw new Error("route 'arc-over-top' requires opts.clearY");
    const a = centerTop(from), b = centerTop(to);
    return {
      segments: [
        { ...boxFromPoints(a, { x: a.x, y: clearY }), arrow: false },
        { ...boxFromPoints({ x: Math.min(a.x, b.x), y: clearY }, { x: Math.max(a.x, b.x), y: clearY }), arrow: false },
        { ...boxFromPoints({ x: b.x, y: clearY }, b), arrow: true },
      ],
      labelAnchor: label ? { x: Math.min(a.x, b.x), y: clearY, w: Math.abs(b.x - a.x) } : null,
    };
  }

  throw new Error(`Unknown route: ${route}`);
}

/** Draws a connector between two node boxes {x,y,w,h}. See computeSegments
 * for what each `route` means; this is the thin drawing layer over it.
 * opts: { color, dashed, label, route, bias, clearY } */
export function connector(slide, from, to, opts = {}) {
  const { color = "94A3B8", dashed = false, label } = opts;
  const { segments, labelAnchor } = computeSegments(from, to, opts);

  for (const s of segments) {
    slide.addShape("line", {
      x: s.x, y: s.y, w: s.w, h: s.h,
      line: {
        color, width: 2,
        dashType: dashed ? "dash" : "solid",
        ...(s.arrow ? { endArrowType: "triangle" } : {}),
      },
    });
  }

  if (label && labelAnchor) {
    slide.addText(label, {
      x: labelAnchor.x, y: labelAnchor.y - 0.32, w: labelAnchor.w, h: 0.28,
      align: "center", fontFace: FONT, fontSize: 8.5, color: "475569",
    });
  }
}

/** Draws a single point-to-point arrow at an explicit y (sequence
 * messages, self-loops) instead of anchoring to node box sides. Direction
 * matters: the arrowhead lands at (x2,y). opts: { color, dashed, label } */
export function drawArrow(slide, { x1, x2, y, color = "94A3B8", dashed = false, label }) {
  slide.addShape("line", {
    x: Math.min(x1, x2), y, w: Math.abs(x2 - x1), h: 0,
    line: { color, width: 2, dashType: dashed ? "dash" : "solid", endArrowType: "triangle" },
  });
  if (label) {
    slide.addText(label, {
      x: Math.min(x1, x2), y: y - 0.26, w: Math.abs(x2 - x1), h: 0.24,
      align: "center", fontFace: FONT, fontSize: 8.5, color: "475569",
    });
  }
}
