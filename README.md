# ptxiagram

**Generate architecture and workflow diagrams as native, editable PowerPoint shapes — not an embedded image.**

ptxiagram is an agent skill for Claude (and a standalone CLI/library) that turns a plain-English description of a system or process into a `.pptx` slide built from real PowerPoint shapes: rectangles, cylinders, clouds, connectors, text boxes. Anyone can open the result in PowerPoint — or Hancom 한쇼 — and click a box to change its color, text, or position, because it *is* a box, not a picture of one.

- **Native shapes, not a picture** — every box and arrow is a real PPTX shape, editable after generation
- **Two diagram types today** — `workflow` (swimlanes, happy-path flows) and `architecture` (free-placed components, trust boundaries, hub-and-spoke)
- **A validator that catches real layout bugs before rendering** — node overlap, labels wider than their shape, and connectors that cut through an unrelated node all fail loudly with a specific fix, instead of shipping a broken-looking slide
- **Ships a Hancom Office compatibility fix** — every generated file is post-processed to strip a documented PptxGenJS defect that PowerPoint silently repairs but Hancom's 한쇼 flags as a corrupted document (see below)
- **Route presets instead of hand-computed coordinates** — `straight`, `drop`, `elbow-right`, `arc-over-top` cover swimlane transitions, hub-and-spoke fan-out, and routing around a boundary box

## Install

As a Claude Skill:

```bash
npx skills add hiio420official/ptxiagram -g
```

As a CLI / library:

```bash
npm install -g ptxiagram
ptxiagram doctor
```

## Quick start

```bash
ptxiagram render workflow examples/onboarding.workflow.json out.pptx
ptxiagram check out.pptx
```

`render` validates the input (schema + layout) before writing anything, and always repairs the output for Hancom Office compatibility. `check` confirms the repair took, dumps every embedded text run for a spelling/content spot-check, and — if LibreOffice is installed — writes a PNG preview for a quick visual check.

See [`ptxiagram/SKILL.md`](ptxiagram/SKILL.md) for the full JSON shape of both diagram types, the node-type → shape/color mapping, and the route presets.

## Why every file gets repaired

PptxGenJS ships a notes-slide scaffold (6 placeholder shapes) on every file it generates, whether or not notes are used. PowerPoint silently repairs this on open. Hancom Office's 한쇼 does not — it shows a "damaged document" dialog instead, on every file, every time. `lib/repair-pptx.mjs` strips the notes infrastructure and the always-empty `charts`/`embeddings` scaffolding before the file is written. Source: [gitbrent/PptxGenJS#1449](https://github.com/gitbrent/PptxGenJS/issues/1449).

## License

MIT
