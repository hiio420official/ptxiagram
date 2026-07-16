import { renderWorkflow } from "../workflow/render-workflow.mjs";

// lifecycle.schema.json is structurally identical to workflow.schema.json
// (lanes/nodes/edges, same layout budget) -- it only differs in its
// diagram_type const and typical usage (states + terminal outcomes instead
// of process steps), so it reuses the workflow renderer wholesale rather
// than duplicating the same ~70 lines under a different file name.
export async function renderLifecycle(inputPath, outputPath) {
  return renderWorkflow(inputPath, outputPath, "lifecycle");
}
