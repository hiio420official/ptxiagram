#!/usr/bin/env node
import fs from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import JSZip from "jszip";
import { renderWorkflow } from "../renderers/workflow/render-workflow.mjs";
import { renderArchitecture } from "../renderers/architecture/render-architecture.mjs";
import { renderSequence } from "../renderers/sequence/render-sequence.mjs";
import { renderDataflow } from "../renderers/dataflow/render-dataflow.mjs";
import { renderLifecycle } from "../renderers/lifecycle/render-lifecycle.mjs";
import { validateAll } from "../lib/validate.mjs";

const execFileAsync = promisify(execFile);

const RENDERERS = {
  workflow: renderWorkflow,
  architecture: renderArchitecture,
  sequence: renderSequence,
  dataflow: renderDataflow,
  lifecycle: renderLifecycle,
};

function usage() {
  console.log(`Usage:
  cli render <type> <input.json> [output.pptx]
  cli validate <type> <input.json> [--json]
  cli check <output.pptx>
  cli doctor

Types: workflow, architecture, sequence, dataflow, lifecycle`);
}

async function cmdRender(type, inputPath, outputPath) {
  const renderFn = RENDERERS[type];
  if (!renderFn) throw new Error(`Unknown type "${type}". Expected one of: ${Object.keys(RENDERERS).join(", ")}`);
  const outFile = await renderFn(inputPath, outputPath);
  console.log(outFile);
}

async function cmdValidate(type, inputPath, asJson) {
  const raw = JSON.parse(await fs.readFile(inputPath, "utf8"));
  try {
    const problems = await validateAll(type, raw);
    const result = { ok: problems.length === 0, type, input: inputPath, problems };
    if (asJson) {
      console.log(JSON.stringify(result, null, 2));
    } else if (problems.length === 0) {
      console.log("OK: schema and layout are clean.");
    } else {
      console.log(`Layout problems (${problems.length}):`);
      for (const p of problems) console.log(`- ${p}`);
    }
    if (problems.length) process.exitCode = 1;
  } catch (err) {
    if (asJson) {
      console.log(JSON.stringify({ ok: false, type, input: inputPath, error: err.message }, null, 2));
    } else {
      console.error(err.message);
    }
    process.exitCode = 1;
  }
}

function findSoffice() {
  const candidates = [
    "C:/Program Files/LibreOffice/program/soffice.exe",
    "/usr/bin/soffice",
    "/opt/libreoffice/program/soffice",
  ];
  return candidates.find((p) => existsSync(p));
}

async function cmdCheck(pptxPath) {
  const buf = await fs.readFile(pptxPath);
  const zip = await JSZip.loadAsync(buf);
  const problems = [];

  const hasNotes = Object.keys(zip.files).some(
    (n) => n.startsWith("ppt/notesMasters/") || n.startsWith("ppt/notesSlides/")
  );
  if (hasNotes) problems.push("notesMaster/notesSlide parts present — repairPptx() was not applied (Hancom Office will flag this as corrupted).");

  const hasEmptyScaffold = Object.keys(zip.files).some(
    (n) => n.startsWith("ppt/charts/") || n.startsWith("ppt/embeddings/")
  );
  if (hasEmptyScaffold) problems.push("empty ppt/charts/ or ppt/embeddings/ scaffolding present.");

  const slideFiles = Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));
  const texts = [];
  for (const name of slideFiles) {
    const content = await zip.file(name).async("string");
    const matches = [...content.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]);
    texts.push(...matches);
  }

  console.log(JSON.stringify({ ok: problems.length === 0, file: pptxPath, problems, embeddedText: texts }, null, 2));

  const soffice = findSoffice();
  if (soffice) {
    const outDir = path.join(path.dirname(pptxPath), "render");
    await fs.mkdir(outDir, { recursive: true });
    await execFileAsync(soffice, ["--headless", "--convert-to", "png", "--outdir", outDir, pptxPath]);
    const pngName = path.basename(pptxPath).replace(/\.pptx$/, ".png");
    console.log(`PNG preview: ${path.join(outDir, pngName)}`);
  } else {
    console.log("LibreOffice not found — skipping PNG preview. Install it for visual QA (soffice --headless --convert-to png).");
  }

  if (problems.length) process.exitCode = 1;
}

async function cmdDoctor() {
  const checks = [];
  try {
    await import("pptxgenjs");
    checks.push(["pptxgenjs", true]);
  } catch {
    checks.push(["pptxgenjs", false]);
  }
  try {
    await import("jszip");
    checks.push(["jszip", true]);
  } catch {
    checks.push(["jszip", false]);
  }
  try {
    await import("ajv/dist/2020.js");
    checks.push(["ajv", true]);
  } catch {
    checks.push(["ajv", false]);
  }
  const soffice = findSoffice();
  checks.push(["LibreOffice (optional, enables `check` PNG preview + visual QA)", Boolean(soffice)]);

  for (const [name, ok] of checks) {
    console.log(`[${ok ? "ok" : "MISSING"}] ${name}`);
  }
  if (checks.some(([, ok]) => !ok)) process.exitCode = 1;
}

const [, , cmd, ...rest] = process.argv;

try {
  if (cmd === "render") {
    const [type, input, output] = rest;
    if (!type || !input) { usage(); process.exit(1); }
    await cmdRender(type, input, output);
  } else if (cmd === "validate") {
    const asJson = rest.includes("--json");
    const [type, input] = rest.filter((a) => a !== "--json");
    if (!type || !input) { usage(); process.exit(1); }
    await cmdValidate(type, input, asJson);
  } else if (cmd === "check") {
    const [pptxPath] = rest;
    if (!pptxPath) { usage(); process.exit(1); }
    await cmdCheck(pptxPath);
  } else if (cmd === "doctor") {
    await cmdDoctor();
  } else {
    usage();
    process.exit(cmd ? 1 : 0);
  }
} catch (err) {
  console.error(err.message);
  process.exit(1);
}
