import { renderSequence } from "../renderers/sequence/render-sequence.mjs";

const out = await renderSequence("examples/cache-miss.sequence.json");
console.log("rendered:", out);
