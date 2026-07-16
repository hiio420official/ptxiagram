import pptxgen from "pptxgenjs";

const pptx = new pptxgen();
pptx.defineLayout({ name: "WIDE", width: 13.33, height: 7.5 });
pptx.layout = "WIDE";

const FONT = "맑은 고딕";
const LINE = pptx.ShapeType.line;
const RRECT = pptx.ShapeType.rect;
const ROUND = pptx.ShapeType.roundRect;

const slide = pptx.addSlide();

slide.addText("신규 직원 온보딩 프로세스", {
  x: 0.4, y: 0.2, w: 12, h: 0.5,
  fontFace: FONT, fontSize: 20, bold: true, color: "0F172A",
});
slide.addText("HR팀 → IT팀 → 팀장", {
  x: 0.4, y: 0.65, w: 12, h: 0.3,
  fontFace: FONT, fontSize: 11, color: "64748B",
});

const lanes = [
  { label: "HR팀", y: 1.15, h: 1.75 },
  { label: "IT팀", y: 3.05, h: 1.75 },
  { label: "팀장", y: 4.95, h: 1.75 },
];
for (const lane of lanes) {
  slide.addShape(RRECT, {
    x: 0.4, y: lane.y, w: 12.5, h: lane.h,
    fill: { color: "F8FAFC" },
    line: { color: "CBD5E1", width: 1, dashType: "dash" },
  });
  slide.addText(lane.label, {
    x: 0.55, y: lane.y + 0.05, w: 2, h: 0.3,
    fontFace: FONT, fontSize: 11, color: "64748B", bold: true,
  });
}

function node({ x, y, label, sublabel, fill, stroke, w = 1.8, h = 0.95 }) {
  slide.addShape(ROUND, {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: fill },
    line: { color: stroke, width: 1.5 },
    shadow: { type: "outer", color: "1E293B", opacity: 0.2, blur: 4, offset: 1, angle: 90 },
  });
  slide.addText(label, {
    x, y: y + 0.1, w, h: 0.4,
    align: "center", fontFace: FONT, fontSize: 12.5, bold: true, color: "0F172A",
  });
  if (sublabel) {
    slide.addText(sublabel, {
      x, y: y + 0.52, w, h: 0.35,
      align: "center", fontFace: FONT, fontSize: 9, color: "64748B",
    });
  }
  return { x, y, w, h };
}

const n1 = node({ x: 0.6, y: 1.5, label: "입사 서류 접수", sublabel: "인사팀 접수 확인", fill: "E2E8F0", stroke: "94A3B8" });
const n2 = node({ x: 3.0, y: 1.5, label: "계정 발급 요청", sublabel: "IT팀에 요청 전달", fill: "E2E8F0", stroke: "94A3B8" });
const n3 = node({ x: 3.0, y: 3.4, label: "계정 생성", sublabel: "사내 계정 · 이메일", fill: "D1FAE5", stroke: "10B981" });
const n4 = node({ x: 5.4, y: 3.4, label: "장비 지급", sublabel: "노트북 · 출입카드", fill: "D1FAE5", stroke: "10B981" });
const n5 = node({ x: 7.8, y: 5.3, label: "킥오프 미팅", sublabel: "팀 소개 · 업무 안내", fill: "DBEAFE", stroke: "3B82F6" });
const n6 = node({ x: 10.2, y: 5.3, label: "온보딩 완료", sublabel: "1주차 체크인 예약", fill: "DBEAFE", stroke: "3B82F6" });

function centerBottom(n) { return { x: n.x + n.w / 2, y: n.y + n.h }; }
function centerTop(n) { return { x: n.x + n.w / 2, y: n.y }; }
function centerRight(n) { return { x: n.x + n.w, y: n.y + n.h / 2 }; }
function centerLeft(n) { return { x: n.x, y: n.y + n.h / 2 }; }

// same-row straight connector
function hArrow(from, to, color, label) {
  const a = centerRight(from), b = centerLeft(to);
  slide.addShape(LINE, {
    x: a.x, y: a.y, w: b.x - a.x, h: 0,
    line: { color, width: 2, endArrowType: "triangle" },
  });
  if (label) {
    slide.addText(label, {
      x: a.x, y: a.y - 0.32, w: b.x - a.x, h: 0.3,
      align: "center", fontFace: FONT, fontSize: 8.5, color: "475569",
    });
  }
}

// cross-lane connector: straight vertical if x aligns, otherwise a
// drop -> jog -> drop "Z" route through the vertical midpoint (mirrors
// archify's `route: drop` with bias 0.5)
function dropArrow(from, to, color, label) {
  const a = centerBottom(from), b = centerTop(to);
  const midY = a.y + (b.y - a.y) * 0.5;

  if (Math.abs(a.x - b.x) < 0.05) {
    slide.addShape(LINE, {
      x: a.x, y: a.y, w: 0, h: b.y - a.y,
      line: { color, width: 2, endArrowType: "triangle" },
    });
  } else {
    slide.addShape(LINE, { x: a.x, y: a.y, w: 0, h: midY - a.y, line: { color, width: 2 } });
    slide.addShape(LINE, {
      x: Math.min(a.x, b.x), y: midY, w: Math.abs(b.x - a.x), h: 0,
      line: { color, width: 2 },
    });
    slide.addShape(LINE, {
      x: b.x, y: midY, w: 0, h: b.y - midY,
      line: { color, width: 2, endArrowType: "triangle" },
    });
  }
  if (label) {
    slide.addText(label, {
      x: Math.min(a.x, b.x) - 0.9, y: midY - 0.32, w: Math.abs(b.x - a.x) + 1.8, h: 0.3,
      align: "center", fontFace: FONT, fontSize: 8.5, color: "475569",
    });
  }
}

hArrow(n1, n2, "94A3B8");
dropArrow(n2, n3, "10B981", "요청 전달");
hArrow(n3, n4, "10B981");
dropArrow(n4, n5, "3B82F6", "장비 수령 후");
hArrow(n5, n6, "3B82F6");

slide.addText("Made with a Claude Skill prototype · native PPTX shapes, editable in PowerPoint", {
  x: 0.4, y: 7.1, w: 12.5, h: 0.3,
  fontFace: FONT, fontSize: 8, color: "94A3B8", italic: true,
});

await pptx.writeFile({ fileName: "prototype/output/onboarding-process-demo.pptx" });
console.log("done");
