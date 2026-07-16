---
name: ptxiagram
description: Create architecture and workflow diagrams as native, editable PowerPoint shapes (real rectangles, connectors, and text boxes a user can click and restyle in PowerPoint or Hancom 한쇼) instead of an embedded image. Accepts a plain-language description, turns it into a small JSON file, and renders it with pptxgenjs. Use when the user asks for a diagram, architecture drawing, or process flow that needs to end up inside a .pptx deck or a Korean-office-suite-compatible presentation - not when they just want a standalone image or an HTML/web diagram (use an SVG-based diagram tool for that instead).
license: MIT
metadata:
  version: "0.1.0"
  status: prototype
  based_on: "gitbrent/PptxGenJS (MIT)"
---

# ptxiagram

Turns a plain-language description of a system or process into a single
`.pptx` slide built from **native PowerPoint shapes** — real rectangles,
cylinders, connectors, and text boxes — not a picture pasted into a slide.
Anyone can then click a box in PowerPoint (or Hancom 한쇼) and change its
color, text, or position directly.

This is the thing to reach for when the deliverable is a slide deck. If the
deliverable is a standalone image, an HTML page, or a README diagram, an
SVG/HTML-based diagram tool is the better fit — this skill's whole reason to
exist is that the output has to remain editable inside PowerPoint.

## Install

```bash
npx skills add hiio420official/ptxiagram -g
```

or as an npm package (CLI + library):

```bash
npm install -g ptxiagram
ptxiagram doctor
```

## Setup

```bash
node bin/cli.mjs doctor
```

Confirms pptxgenjs, jszip, and ajv are present, and reports whether
LibreOffice is installed (optional — enables the `check` command's PNG
preview for visual QA; without it you still get full schema/layout
validation and the embedded-text dump, just no rendered preview image).

## Choosing a diagram type

| Type | Use for | Schema |
|------|---------|--------|
| `workflow` | Cross-team processes, approval chains, anything with swimlanes and a happy path | `schemas/workflow.schema.json` |
| `architecture` | System components, services, datastores, trust boundaries, hub-and-spoke integrations | `schemas/architecture.schema.json` |

## The loop

1. **Read the matching example** before writing JSON by hand — copy its
   shape, don't guess the field names:
   - `examples/onboarding.workflow.json`
   - `examples/order-processing.architecture.json`
2. Write `<name>.<type>.json`.
3. Render:
   ```bash
   node bin/cli.mjs render workflow <input>.json <output>.pptx
   ```
   This validates the JSON (schema + layout) and **always** repairs the
   output for Hancom Office compatibility before writing the final file —
   see "Why every file gets repaired" below. If layout validation fails, the
   error names exactly what's wrong (which nodes overlap, which label is
   wider than its shape, which edge cuts through an unrelated node) — fix
   the JSON and re-run; don't hand-edit the generated pptx.
4. Spot-check before handing it over:
   ```bash
   node bin/cli.mjs check <output>.pptx
   ```
   Confirms the Hancom repair actually took (no leftover notes/chart
   scaffolding), dumps every embedded text run so you can read the real
   strings back (LibreOffice's PNG rendering has shown font-substitution
   artifacts that don't reflect the actual file — trust this text dump over
   the preview image for spelling/content, and trust the PNG for layout).
   If LibreOffice is installed, this also writes a PNG preview next to the
   file so you can visually confirm layout before delivering it.

## Workflow JSON shape

```json
{
  "schema_version": 1,
  "diagram_type": "workflow",
  "meta": { "title": "...", "subtitle": "...", "output": "out.pptx" },
  "lanes": [{ "id": "hr", "label": "HR팀" }],
  "nodes": [
    { "id": "apply", "lane": "hr", "col": 0, "type": "neutral", "label": "...", "sublabel": "..." }
  ],
  "edges": [
    { "from": "apply", "to": "request", "variant": "default", "route": "straight" }
  ],
  "cards": []
}
```

**Layout budget**: lanes stack vertically and are auto-positioned (default
height 1.75in). Each lane has 6 column slots (`col` 0–5, each 2.0in wide);
a node claims exactly one slot, so default-width nodes can never overlap —
omit `col` and the renderer assigns the next free slot in that lane in JSON
order. Default node size is 1.8×0.95in.

## Architecture JSON shape

```json
{
  "schema_version": 1,
  "diagram_type": "architecture",
  "meta": { "title": "...", "output": "out.pptx" },
  "components": [
    { "id": "api", "type": "backend", "label": "API 서버", "pos": [5.9, 3.1], "size": [1.85, 1.1] }
  ],
  "boundaries": [{ "label": "사내 인프라 (VPC)", "wraps": ["api", "db"] }],
  "connections": [
    { "from": "api", "to": "db", "variant": "default", "route": "elbow-right", "label": "주문 저장" }
  ],
  "cards": []
}
```

**Layout budget**: free placement in inches on a 13.33×7.5in (16:9) slide.
`boundaries[].wraps` takes component ids, not coordinates — the renderer
computes a padded frame around them, so never hand-compute a boundary box.

## Node types

`client` `frontend` `neutral` `backend` `database` `cache` `queue`
`external` `security` `decision` — each maps to a shape preset (ellipse,
rect, roundRect, can/cylinder, hexagon, cube, cloud, diamond) and a color
pair. Pick by role, not by desired color.

## Route presets

Connectors never need hand-computed coordinates — pick the preset that
matches the story:

- `straight` — adjacent nodes, same row. One segment.
- `drop` — top-to-bottom flow (cross-lane, or stacked components). Straight
  if the two nodes share an x-center, otherwise a vertical→horizontal→vertical
  jog through the midpoint (`bias` picks where, default 0.5).
- `elbow-right` — hub-and-spoke fan-out where targets sit at different row
  heights. Straight if the two nodes share a y-center, otherwise a
  horizontal→vertical→horizontal jog.
- `arc-over-top` — routes above a shared clearance line (`clearY`, required)
  instead of cutting through whatever sits between two nodes. Use this for
  a connection that has to reach past a boundary box or a row of components
  a straight/elbow line would otherwise cross.

`node overlaps`, `label wider than its shape`, and `edge crosses an
unrelated node` are exactly the failure modes `render`/`validate` catch
before a file is written — if you hit one, change the route or `bias`
rather than fighting raw coordinates.

## Why every file gets repaired

pptxgenjs ships a notes-slide scaffold (6 placeholder shapes) on every
file it generates, regardless of whether notes are used. PowerPoint
silently repairs this on open; Hancom Office's 한쇼 does not — it shows a
"damaged document" dialog instead. `lib/repair-pptx.mjs` strips the notes
infrastructure and the always-empty `charts`/`embeddings` scaffolding
after generation and before the file is written. Every `render` call does
this automatically; `check` verifies it actually happened. Source:
[gitbrent/PptxGenJS#1449](https://github.com/gitbrent/PptxGenJS/issues/1449).

## What's not built yet

- `sequence`, `dataflow`, `lifecycle` diagram types (only `workflow` and
  `architecture` exist so far)
- No CJK exact-metrics font measurement — the label-width check is a
  conservative estimate, not real glyph widths
