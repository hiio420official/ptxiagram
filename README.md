# ptxiagram

**Generate architecture, workflow, and sequence diagrams as native, editable PowerPoint shapes — not an embedded image.**

> Built with [Claude Code](https://claude.com/claude-code) (Claude Sonnet 5) — from the initial prototype through the pptxgenjs/Hancom Office compatibility investigation, the JSON schema design, the layout validator, and this package's release automation. See the [commit history](https://github.com/hiio420official/ptxiagram/commits/main) for the build process.

ptxiagram is an agent skill for Claude (and a standalone CLI/library) that turns a plain-English description of a system or process into a `.pptx` slide built from real PowerPoint shapes: rectangles, cylinders, clouds, connectors, text boxes. Anyone can open the result in PowerPoint — or Hancom 한쇼 — and click a box to change its color, text, or position, because it *is* a box, not a picture of one.

- **Native shapes, not a picture** — every box and arrow is a real PPTX shape, editable after generation
- **Three diagram types today** — `workflow` (swimlanes, happy-path flows), `architecture` (free-placed components, trust boundaries, hub-and-spoke), and `sequence` (participants + lifelines, API call chains)
- **A validator that catches real layout bugs before rendering** — node overlap, labels wider than their shape, and connectors that cut through an unrelated node all fail loudly with a specific fix, instead of shipping a broken-looking slide
- **Ships a Hancom Office compatibility fix** — every generated file is post-processed to strip a documented PptxGenJS defect that PowerPoint silently repairs but Hancom's 한쇼 flags as a corrupted document (see below)
- **Route presets instead of hand-computed coordinates** — `straight`, `drop`, `elbow-right`, `arc-over-top` cover swimlane transitions, hub-and-spoke fan-out, and routing around a boundary box

## Use it in an agent (Claude Code, Codex, Cursor, ...)

Install with the [`skills` CLI](https://github.com/vercel-labs/skills) — it detects every supported agent on your machine and wires the skill into each one:

```bash
npx skills add hiio420official/ptxiagram -g
```

Then just ask in chat:

```text
Use ptxiagram to draw an architecture diagram for this repo's runtime,
as a PPTX I can drop into a slide deck.
```

```text
Use ptxiagram to make a workflow diagram of our incident response process,
with a lane per team, so I can paste it into the postmortem deck.
```

```text
Use ptxiagram to draw a sequence diagram of this API's cache-fallback path,
from request to response, for the engineering review deck.
```

The agent reads `SKILL.md`, writes a small JSON file describing your nodes and
edges, and runs the CLI to produce a `.pptx` with real, editable shapes.

**Per-agent notes:**
- **Claude Code**: works out of the box after `npx skills add`.
- **Codex CLI / opencode**: install with an explicit `--agent` flag if you don't want it installed everywhere, e.g. `npx skills use hiio420official/ptxiagram@ptxiagram --agent codex`.
- **Any other `skills`-supported agent** (Cursor, Windsurf, Continue, Antigravity, Gemini CLI, ...): the same `npx skills add` command installs to whichever of these are detected on your machine.

## Use it standalone (CLI / library)

```bash
npm install -g ptxiagram
ptxiagram doctor
ptxiagram render workflow examples/onboarding.workflow.json out.pptx
ptxiagram check out.pptx
```

Or as a library:

```js
import { renderWorkflow } from "ptxiagram";
await renderWorkflow("my-diagram.workflow.json", "out.pptx");
```

`render` validates the input (schema + layout) before writing anything, and always repairs the output for Hancom Office compatibility. `check` confirms the repair took, dumps every embedded text run for a spelling/content spot-check, and — if LibreOffice is installed — writes a PNG preview for a quick visual check.

See [`ptxiagram/SKILL.md`](ptxiagram/SKILL.md) for the full JSON shape of all three diagram types, the node-type → shape/color mapping, and the route presets.

## Why every file gets repaired

PptxGenJS ships a notes-slide scaffold (6 placeholder shapes) on every file it generates, whether or not notes are used. PowerPoint silently repairs this on open. Hancom Office's 한쇼 does not — it shows a "damaged document" dialog instead, on every file, every time. `lib/repair-pptx.mjs` strips the notes infrastructure and the always-empty `charts`/`embeddings` scaffolding before the file is written. Source: [gitbrent/PptxGenJS#1449](https://github.com/gitbrent/PptxGenJS/issues/1449).

## License

MIT
