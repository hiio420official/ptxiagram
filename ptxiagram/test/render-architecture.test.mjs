import { renderArchitecture } from "../renderers/architecture/render-architecture.mjs";

const out = await renderArchitecture("examples/order-processing.architecture.json");
console.log("rendered:", out);
