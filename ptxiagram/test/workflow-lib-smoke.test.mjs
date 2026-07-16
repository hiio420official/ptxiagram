import pptxgen from "pptxgenjs";
import { drawNode, drawLane, FONT } from "../lib/shapes.mjs";
import { connector } from "../lib/routes.mjs";
import { repairPptx } from "../lib/repair-pptx.mjs";

const pptx = new pptxgen();
pptx.defineLayout({ name: "WIDE", width: 13.33, height: 7.5 });
pptx.layout = "WIDE";
const slide = pptx.addSlide();

slide.addText("신규 직원 온보딩 프로세스 (lib 재구현)", {
  x: 0.4, y: 0.2, w: 12, h: 0.5,
  fontFace: FONT, fontSize: 20, bold: true, color: "0F172A",
});

drawLane(slide, { x: 0.4, y: 1.15, w: 12.5, h: 1.75, label: "HR팀" });
drawLane(slide, { x: 0.4, y: 3.05, w: 12.5, h: 1.75, label: "IT팀" });
drawLane(slide, { x: 0.4, y: 4.95, w: 12.5, h: 1.75, label: "팀장" });

const n1 = drawNode(slide, { x: 0.6, y: 1.5, label: "입사 서류 접수", sublabel: "인사팀 접수 확인", type: "neutral" });
const n2 = drawNode(slide, { x: 3.0, y: 1.5, label: "계정 발급 요청", sublabel: "IT팀에 요청 전달", type: "neutral" });
const n3 = drawNode(slide, { x: 3.0, y: 3.4, label: "계정 생성", sublabel: "사내 계정 · 이메일", type: "backend" });
const n4 = drawNode(slide, { x: 5.4, y: 3.4, label: "장비 지급", sublabel: "노트북 · 출입카드", type: "backend" });
const n5 = drawNode(slide, { x: 7.8, y: 5.3, label: "킥오프 미팅", sublabel: "팀 소개 · 업무 안내", type: "database" });
const n6 = drawNode(slide, { x: 10.2, y: 5.3, label: "온보딩 완료", sublabel: "1주차 체크인 예약", type: "database" });

connector(slide, n1, n2, { color: "94A3B8", route: "straight" });
connector(slide, n2, n3, { color: "10B981", route: "drop", label: "요청 전달" });
connector(slide, n3, n4, { color: "10B981", route: "straight" });
connector(slide, n4, n5, { color: "3B82F6", route: "drop", label: "장비 수령 후" });
connector(slide, n5, n6, { color: "3B82F6", route: "straight" });

const outPath = "test/output/lib-smoke-test.pptx";
await pptx.writeFile({ fileName: outPath });
await repairPptx(outPath);
console.log("done");
