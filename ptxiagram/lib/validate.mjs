import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import Ajv2020 from "ajv/dist/2020.js";
import { computeWorkflowLayout, computeArchitectureLayout, computeSequenceLayout, SLIDE_W, SLIDE_H } from "./layout.mjs";
import { computeSegments } from "./routes.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SCHEMAS_DIR = path.join(__dirname, "..", "schemas");

let ajvInstance;
async function getAjv() {
  if (ajvInstance) return ajvInstance;
  const ajv = new Ajv2020({ allErrors: true, strict: false });
  const common = JSON.parse(await fs.readFile(path.join(SCHEMAS_DIR, "common.schema.json"), "utf8"));
  ajv.addSchema(common, "common.schema.json");
  ajvInstance = ajv;
  return ajv;
}

/** Validates `data` against schemas/<type>.schema.json. Throws with a
 * readable message (path + reason per error) if invalid. */
export async function validateSchema(type, data) {
  const ajv = await getAjv();
  const schemaPath = path.join(SCHEMAS_DIR, `${type}.schema.json`);
  const schema = JSON.parse(await fs.readFile(schemaPath, "utf8"));
  const validateFn = ajv.getSchema(schema.$id) || ajv.compile(schema);
  const ok = validateFn(data);
  if (!ok) {
    const messages = validateFn.errors.map((e) => `${e.instancePath || "/"} ${e.message}`);
    throw new Error(`Schema validation failed:\n- ${messages.join("\n- ")}`);
  }
  return true;
}

function rectsOverlap(a, b) {
  return a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y;
}

const CJK_RE = /[ㄱ-힝一-鿿]/;

// Crude CJK-aware text width estimate (inches) for a given font size (pt).
// Not exact glyph metrics -- good enough to flag labels that will visibly
// overflow their shape, which is the failure mode this exists to catch.
function estimateTextWidth(text, fontSizePt) {
  const unit = (fontSizePt / 72) * 0.62;
  let units = 0;
  for (const ch of text) units += CJK_RE.test(ch) ? 1.0 : 0.55;
  return units * unit;
}

// All routes in lib/routes.mjs only ever emit axis-aligned segments, so
// treating each as a simple bounding box is exact, not an approximation.
function segmentIntersectsRect(seg, rect, epsilon = 0.02) {
  const x1 = Math.min(seg.x, seg.x + seg.w), x2 = Math.max(seg.x, seg.x + seg.w);
  const y1 = Math.min(seg.y, seg.y + seg.h), y2 = Math.max(seg.y, seg.y + seg.h);
  return (
    x1 < rect.x + rect.w - epsilon && x2 > rect.x + epsilon &&
    y1 < rect.y + rect.h - epsilon && y2 > rect.y + epsilon
  );
}

function offSlide(box) {
  return box.x < 0 || box.y < 0 || box.x + box.w > SLIDE_W || box.y + box.h > SLIDE_H;
}

function checkCommon({ problems, ids, boxes, labelOf, fontSizePt, edgesOrConns, resolveEndpoints, edgeLabel }) {
  for (const id of ids) {
    const box = boxes[id];
    const label = labelOf(id);
    const w = estimateTextWidth(label, fontSizePt);
    if (w > box.w - 0.1) {
      problems.push(`"${id}" label "${label}" (~${w.toFixed(2)}in) is wider than its shape (${box.w.toFixed(2)}in) — shorten the label or move detail to sublabel.`);
    }
    if (offSlide(box)) {
      problems.push(`"${id}" is off-slide (slide is ${SLIDE_W}x${SLIDE_H}in).`);
    }
  }

  for (let i = 0; i < ids.length; i++) {
    for (let j = i + 1; j < ids.length; j++) {
      if (rectsOverlap(boxes[ids[i]], boxes[ids[j]])) {
        problems.push(`"${ids[i]}" and "${ids[j]}" overlap.`);
      }
    }
  }

  for (const e of edgesOrConns) {
    const { from, to } = resolveEndpoints(e);
    if (!from || !to) continue; // unknown-id errors surface from the renderer/schema instead
    const { segments } = computeSegments(from, to, {
      route: e.route || "straight", bias: e.bias, clearY: e.clearY, label: e.label,
    });
    for (const id of ids) {
      if (id === e.from || id === e.to) continue;
      const rect = boxes[id];
      if (segments.some((seg) => segmentIntersectsRect(seg, rect))) {
        problems.push(`${edgeLabel(e)} crosses "${id}" — pick a different route/bias/clearY or reposition.`);
      }
    }
  }

  return problems;
}

export function validateWorkflowLayout(raw) {
  const { nodeBoxes } = computeWorkflowLayout(raw);
  const labelOf = (id) => raw.nodes.find((n) => n.id === id).label;
  return checkCommon({
    problems: [],
    ids: Object.keys(nodeBoxes),
    boxes: nodeBoxes,
    labelOf,
    fontSizePt: 12.5,
    edgesOrConns: raw.edges,
    resolveEndpoints: (e) => ({ from: nodeBoxes[e.from], to: nodeBoxes[e.to] }),
    edgeLabel: (e) => `Edge "${e.from}"->"${e.to}"`,
  });
}

export function validateArchitectureLayout(raw) {
  const { boxes } = computeArchitectureLayout(raw);
  const labelOf = (id) => raw.components.find((c) => c.id === id).label;
  return checkCommon({
    problems: [],
    ids: Object.keys(boxes),
    boxes,
    labelOf,
    fontSizePt: 12.5,
    edgesOrConns: raw.connections,
    resolveEndpoints: (e) => ({ from: boxes[e.from], to: boxes[e.to] }),
    edgeLabel: (e) => `Connection "${e.from}"->"${e.to}"`,
  });
}

export function validateSequenceLayout(raw) {
  const problems = [];
  const { participantBoxes, messageYs, lifelineBottom } = computeSequenceLayout(raw);
  const ids = Object.keys(participantBoxes);

  for (const p of raw.participants) {
    const box = participantBoxes[p.id];
    const w = estimateTextWidth(p.label, 12.5);
    if (w > box.w - 0.1) {
      problems.push(`Participant "${p.id}" label "${p.label}" (~${w.toFixed(2)}in) is wider than its shape (${box.w.toFixed(2)}in) — shorten the label or move detail to sublabel.`);
    }
    if (offSlide(box)) {
      problems.push(`Participant "${p.id}" is off-slide (slide is ${SLIDE_W}x${SLIDE_H}in).`);
    }
  }

  for (let i = 0; i < ids.length; i++) {
    for (let j = i + 1; j < ids.length; j++) {
      if (rectsOverlap(participantBoxes[ids[i]], participantBoxes[ids[j]])) {
        problems.push(`Participants "${ids[i]}" and "${ids[j]}" overlap.`);
      }
    }
  }

  raw.messages.forEach((m, i) => {
    if (m.from === m.to) {
      problems.push(`Message ${i} ("${m.label}") has the same participant "${m.from}" as from and to — self-messages aren't supported yet.`);
    }
    if (!participantBoxes[m.from]) problems.push(`Message ${i} references unknown participant "${m.from}".`);
    if (!participantBoxes[m.to]) problems.push(`Message ${i} references unknown participant "${m.to}".`);
    if (messageYs[i] > lifelineBottom - 0.3) {
      problems.push(`Message ${i} ("${m.label}") falls below the lifeline (too many messages for the available space) — split into fewer messages or drop the cards to free up room.`);
    }
  });

  return problems;
}

/** Runs schema validation, then layout collision checks. Returns the list
 * of layout problems (empty = clean); throws if schema validation fails. */
export async function validateAll(type, raw) {
  await validateSchema(type, raw);
  if (type === "workflow") return validateWorkflowLayout(raw);
  if (type === "architecture") return validateArchitectureLayout(raw);
  if (type === "sequence") return validateSequenceLayout(raw);
  throw new Error(`Unknown diagram type: ${type}`);
}
