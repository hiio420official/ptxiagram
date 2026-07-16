import pptxgen from "pptxgenjs";

const pptx = new pptxgen();
pptx.defineLayout({ name: "WIDE", width: 13.33, height: 7.5 });
pptx.layout = "WIDE";

const FONT = "맑은 고딕";
const LINE = pptx.ShapeType.line;

const slide = pptx.addSlide();

slide.addText("온라인 쇼핑몰 주문처리 아키텍처", {
  x: 0.4, y: 0.2, w: 12, h: 0.5,
  fontFace: FONT, fontSize: 20, bold: true, color: "0F172A",
});
slide.addText("클라이언트 → API 서버 → 내부 인프라 / 외부 결제 PG", {
  x: 0.4, y: 0.65, w: 12, h: 0.3,
  fontFace: FONT, fontSize: 11, color: "64748B",
});

// internal infra boundary (dashed region, like a VPC / trust boundary)
slide.addShape(pptx.ShapeType.rect, {
  x: 5.5, y: 1.3, w: 4.1, h: 4.6,
  fill: { type: "none" },
  line: { color: "F59E0B", width: 1.25, dashType: "dash" },
});
slide.addText("사내 인프라 (VPC)", {
  x: 5.65, y: 1.4, w: 3, h: 0.3,
  fontFace: FONT, fontSize: 10, color: "B45309", bold: true,
});

function shape(type, { x, y, w, h, label, sublabel, fill, stroke, fontSize = 12 }) {
  slide.addShape(type, {
    x, y, w, h,
    fill: { color: fill },
    line: { color: stroke, width: 1.5 },
    shadow: { type: "outer", color: "1E293B", opacity: 0.18, blur: 4, offset: 1, angle: 90 },
  });
  slide.addText(label, {
    x, y: y + h / 2 - (sublabel ? 0.28 : 0.14), w, h: 0.32,
    align: "center", valign: "middle", fontFace: FONT, fontSize, bold: true, color: "0F172A",
  });
  if (sublabel) {
    slide.addText(sublabel, {
      x, y: y + h / 2 + 0.08, w, h: 0.3,
      align: "center", fontFace: FONT, fontSize: 8.5, color: "64748B",
    });
  }
  return { x, y, w, h };
}

const client = shape(pptx.ShapeType.ellipse, {
  x: 0.4, y: 3.1, w: 1.5, h: 1.1, label: "사용자", sublabel: "브라우저", fill: "F1F5F9", stroke: "94A3B8",
});
const web = shape(pptx.ShapeType.rect, {
  x: 2.35, y: 3.1, w: 1.7, h: 1.1, label: "웹 / 앱", sublabel: "React", fill: "E2E8F0", stroke: "94A3B8",
});
const api = shape(pptx.ShapeType.roundRect, {
  x: 5.9, y: 3.1, w: 1.85, h: 1.1, label: "API 서버", sublabel: "주문 처리", fill: "D1FAE5", stroke: "10B981",
});
const db = shape(pptx.ShapeType.can, {
  x: 8.4, y: 1.55, w: 1.0, h: 1.2, label: "주문 DB", fill: "DBEAFE", stroke: "3B82F6", fontSize: 10.5,
});
const cache = shape(pptx.ShapeType.hexagon, {
  x: 8.3, y: 3.1, w: 1.15, h: 1.1, label: "캐시", sublabel: "Redis", fill: "DBEAFE", stroke: "3B82F6", fontSize: 11,
});
const queue = shape(pptx.ShapeType.cube, {
  x: 8.3, y: 4.55, w: 1.15, h: 1.05, label: "알림 큐", fill: "DBEAFE", stroke: "3B82F6", fontSize: 10.5,
});
const pg = shape(pptx.ShapeType.cloud, {
  x: 10.4, y: 3.0, w: 2.1, h: 1.3, label: "결제 PG사", sublabel: "외부 서비스", fill: "FEF3C7", stroke: "F59E0B",
});

function centerRight(n) { return { x: n.x + n.w, y: n.y + n.h / 2 }; }
function centerLeft(n) { return { x: n.x, y: n.y + n.h / 2 }; }
function centerTop(n) { return { x: n.x + n.w / 2, y: n.y }; }

// routes above a row of nodes entirely (e.g. an external service reached
// past a boundary box) instead of drawing a straight line that would cut
// through whatever sits between `from` and `to`
function arcOverTop(from, to, color, dashed, label, clearY) {
  const a = centerTop(from), b = centerTop(to);
  slide.addShape(LINE, { x: a.x, y: clearY, w: 0, h: a.y - clearY, line: { color, width: 2, dashType: dashed ? "dash" : "solid" } });
  slide.addShape(LINE, { x: Math.min(a.x, b.x), y: clearY, w: Math.abs(b.x - a.x), h: 0, line: { color, width: 2, dashType: dashed ? "dash" : "solid" } });
  slide.addShape(LINE, {
    x: b.x, y: clearY, w: 0, h: b.y - clearY,
    line: { color, width: 2, endArrowType: "triangle", dashType: dashed ? "dash" : "solid" },
  });
  if (label) {
    slide.addText(label, {
      x: Math.min(a.x, b.x), y: clearY - 0.32, w: Math.abs(b.x - a.x), h: 0.28,
      align: "center", fontFace: FONT, fontSize: 8.5, color: "475569",
    });
  }
}

function straight(from, to, color, dashed, label) {
  const a = centerRight(from), b = centerLeft(to);
  slide.addShape(LINE, {
    x: a.x, y: a.y, w: b.x - a.x, h: 0,
    line: { color, width: 2, endArrowType: "triangle", dashType: dashed ? "dash" : "solid" },
  });
  if (label) {
    slide.addText(label, {
      x: a.x, y: a.y - 0.3, w: b.x - a.x, h: 0.28,
      align: "center", fontFace: FONT, fontSize: 8.5, color: "475569",
    });
  }
}

// hub-and-spoke: right side of `from` fans out to several targets stacked
// vertically, jogging through a shared mid-x when the target isn't level
function elbowRight(from, to, color, label) {
  const a = centerRight(from), b = centerLeft(to);
  if (Math.abs(a.y - b.y) < 0.05) {
    straight(from, to, color, false, label);
    return;
  }
  const midX = a.x + (b.x - a.x) * 0.45;
  slide.addShape(LINE, { x: a.x, y: a.y, w: midX - a.x, h: 0, line: { color, width: 2 } });
  slide.addShape(LINE, { x: midX, y: Math.min(a.y, b.y), w: 0, h: Math.abs(b.y - a.y), line: { color, width: 2 } });
  slide.addShape(LINE, {
    x: midX, y: b.y, w: b.x - midX, h: 0,
    line: { color, width: 2, endArrowType: "triangle" },
  });
  if (label) {
    slide.addText(label, {
      x: midX - 0.6, y: b.y - 0.3, w: 1.2, h: 0.28,
      align: "center", fontFace: FONT, fontSize: 8, color: "475569",
    });
  }
}

straight(client, web, "94A3B8");
straight(web, api, "10B981");
elbowRight(api, db, "3B82F6", "주문 저장");
elbowRight(api, cache, "3B82F6");
elbowRight(api, queue, "3B82F6", "발송 알림");
arcOverTop(api, pg, "F59E0B", true, "결제 요청 (async)", 1.0);

slide.addText("Made with a Claude Skill prototype · native PPTX shapes, editable in PowerPoint", {
  x: 0.4, y: 7.1, w: 12.5, h: 0.3,
  fontFace: FONT, fontSize: 8, color: "94A3B8", italic: true,
});

await pptx.writeFile({ fileName: "prototype/output/architecture-demo.pptx" });
console.log("done");
