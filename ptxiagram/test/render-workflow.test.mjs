import { renderWorkflow } from "../renderers/workflow/render-workflow.mjs";

const out = await renderWorkflow("examples/onboarding.workflow.json");
console.log("rendered:", out);
