import pptxgen from "pptxgenjs";

// Isolate what Hancom Office (한쇼) considers "corrupted" in a pptxgenjs
// file that both PowerPoint-compatible LibreOffice renders fine and
// python-pptx / raw XML inspection show as structurally valid OOXML.
// Each file changes exactly one variable from the previous baseline.

async function testA_onlyRects() {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9"; // standard preset, not a custom defineLayout
  const slide = pptx.addSlide();
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 1, y: 1, w: 2, h: 1, rectRadius: 0.06,
    fill: { color: "E2E8F0" },
    line: { color: "94A3B8", width: 1.5 },
  });
  slide.addText("사각형만", { x: 1, y: 1.1, w: 2, h: 0.4, align: "center", fontFace: "맑은 고딕", fontSize: 12 });
  await pptx.writeFile({ fileName: "prototype/output/diag-a-only-rects.pptx" });
}

async function testB_zeroHeightLine() {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";
  const slide = pptx.addSlide();
  slide.addShape(pptx.ShapeType.line, {
    x: 1, y: 1, w: 2, h: 0,
    line: { color: "94A3B8", width: 2, endArrowType: "triangle" },
  });
  await pptx.writeFile({ fileName: "prototype/output/diag-b-zero-height-line.pptx" });
}

async function testC_nonZeroLine() {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";
  const slide = pptx.addSlide();
  slide.addShape(pptx.ShapeType.line, {
    x: 1, y: 1, w: 2, h: 0.02,
    line: { color: "94A3B8", width: 2, endArrowType: "triangle" },
  });
  await pptx.writeFile({ fileName: "prototype/output/diag-c-nonzero-line.pptx" });
}

async function testD_customLayout() {
  const pptx = new pptxgen();
  pptx.defineLayout({ name: "WIDE", width: 13.33, height: 7.5 });
  pptx.layout = "WIDE";
  const slide = pptx.addSlide();
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 1, y: 1, w: 2, h: 1, rectRadius: 0.06,
    fill: { color: "E2E8F0" },
    line: { color: "94A3B8", width: 1.5 },
  });
  slide.addText("커스텀 레이아웃", { x: 1, y: 1.1, w: 2, h: 0.4, align: "center", fontFace: "맑은 고딕", fontSize: 12 });
  await pptx.writeFile({ fileName: "prototype/output/diag-d-custom-layout.pptx" });
}

async function testE_shadow() {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";
  const slide = pptx.addSlide();
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 1, y: 1, w: 2, h: 1, rectRadius: 0.06,
    fill: { color: "E2E8F0" },
    line: { color: "94A3B8", width: 1.5 },
    shadow: { type: "outer", color: "1E293B", opacity: 0.2, blur: 4, offset: 1, angle: 90 },
  });
  await pptx.writeFile({ fileName: "prototype/output/diag-e-shadow.pptx" });
}

await testA_onlyRects();
await testB_zeroHeightLine();
await testC_nonZeroLine();
await testD_customLayout();
await testE_shadow();
console.log("done");
