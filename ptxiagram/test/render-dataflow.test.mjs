import { renderDataflow } from "../renderers/dataflow/render-dataflow.mjs";

const out = await renderDataflow("examples/product-analytics.dataflow.json");
console.log("rendered:", out);
