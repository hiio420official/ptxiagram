import fs from "node:fs/promises";
import { validateAll } from "../lib/validate.mjs";

const wf = JSON.parse(await fs.readFile("examples/onboarding.workflow.json", "utf8"));
const arch = JSON.parse(await fs.readFile("examples/order-processing.architecture.json", "utf8"));

console.log("--- workflow (should be clean) ---");
console.log(await validateAll("workflow", wf));

console.log("--- architecture (should be clean) ---");
console.log(await validateAll("architecture", arch));

console.log("--- deliberately broken workflow ---");
const broken = JSON.parse(JSON.stringify(wf));
// force an overlap: put two nodes in the same lane+col
broken.nodes[1].col = 0; // collides with node[0] (apply) which is also col 0
// force an off-slide label
broken.nodes[0].label = "이것은 아주 아주 아주 길고 긴 라벨입니다 넘칠 예정";
const problems = await validateAll("workflow", broken);
console.log(problems);
