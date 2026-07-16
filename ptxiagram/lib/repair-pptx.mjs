import fs from "node:fs/promises";
import JSZip from "jszip";

/**
 * Fixes known PptxGenJS OOXML defects that PowerPoint and Hancom Office's
 * 한쇼 flag as a corrupted/damaged document (LibreOffice is lenient and
 * renders the unfixed output fine, which is why the defects go unnoticed
 * until a stricter parser opens the file).
 *
 * Source: https://github.com/gitbrent/PptxGenJS/issues/1449
 *   - Defect 5: notesMaster1.xml always ships 6 placeholder shapes
 *     (hdr/dt/sldImg/body/ftr/sldNum) that PowerPoint-family parsers
 *     consider malformed. PptxGenJS emits this notes scaffold on every
 *     file regardless of whether notes are used, so every generated
 *     pptx carries it.
 *   - Defect 6: empty ppt/charts/ and ppt/embeddings/ directories are
 *     always created even when no chart/embedding exists.
 *
 * Strategy: since our diagrams never use PPT speaker notes or embedded
 * charts, the simplest fix is to strip the notes infrastructure and the
 * empty scaffold directories entirely rather than trying to make them
 * spec-conformant.
 */
export async function repairPptx(filePath) {
  const buf = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(buf);

  // Defect 5: drop notes parts
  for (const name of Object.keys(zip.files)) {
    if (name.startsWith("ppt/notesMasters/") || name.startsWith("ppt/notesSlides/")) {
      zip.remove(name);
    }
  }

  // Defect 6: drop empty chart/embedding scaffolding
  for (const name of Object.keys(zip.files)) {
    if (name.startsWith("ppt/charts/") || name.startsWith("ppt/embeddings/")) {
      zip.remove(name);
    }
  }

  // [Content_Types].xml: remove notesMaster/notesSlide Overrides
  const ctPath = "[Content_Types].xml";
  let ct = await zip.file(ctPath).async("string");
  ct = ct.replace(/<Override[^>]*PartName="\/ppt\/notesMasters\/[^"]*"[^>]*\/>/g, "");
  ct = ct.replace(/<Override[^>]*PartName="\/ppt\/notesSlides\/[^"]*"[^>]*\/>/g, "");
  zip.file(ctPath, ct);

  // presentation.xml: remove <p:notesMasterIdLst>
  const presPath = "ppt/presentation.xml";
  let pres = await zip.file(presPath).async("string");
  pres = pres.replace(/<p:notesMasterIdLst>[\s\S]*?<\/p:notesMasterIdLst>/g, "");
  zip.file(presPath, pres);

  // presentation.xml.rels: remove notesMaster relationship
  const presRelsPath = "ppt/_rels/presentation.xml.rels";
  let presRels = await zip.file(presRelsPath).async("string");
  presRels = presRels.replace(/<Relationship[^>]*Type="[^"]*\/notesMaster"[^>]*\/>/g, "");
  zip.file(presRelsPath, presRels);

  // each slide's .rels: remove notesSlide relationship
  const slideRelsPattern = /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/;
  for (const name of Object.keys(zip.files)) {
    if (slideRelsPattern.test(name)) {
      let content = await zip.file(name).async("string");
      const cleaned = content.replace(/<Relationship[^>]*Type="[^"]*\/notesSlide"[^>]*\/>/g, "");
      if (cleaned !== content) zip.file(name, cleaned);
    }
  }

  const out = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  await fs.writeFile(filePath, out);
}
