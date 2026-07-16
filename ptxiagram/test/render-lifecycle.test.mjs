import { renderLifecycle } from "../renderers/lifecycle/render-lifecycle.mjs";

const out = await renderLifecycle("examples/agent-run.lifecycle.json");
console.log("rendered:", out);
