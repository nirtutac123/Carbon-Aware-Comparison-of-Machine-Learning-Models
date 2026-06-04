import fs from "node:fs";
import path from "node:path";
import { PresentationFile } from "@oai/artifact-tool";

const [pptxPath, outputDir] = process.argv.slice(2);
if (!pptxPath || !outputDir) {
  console.error("Usage: node render_pptx_previews.mjs <deck.pptx> <output_dir>");
  process.exit(2);
}

fs.mkdirSync(outputDir, { recursive: true });
const bytes = fs.readFileSync(pptxPath);
const presentation = await PresentationFile.importPptx(bytes);
const count =
  typeof presentation.slides.count === "function"
    ? presentation.slides.count()
    : presentation.slides.count;

for (let i = 0; i < count; i += 1) {
  const slide = presentation.slides.getItem(i);
  const blob = await presentation.export({ slide, format: "png" });
  const buffer = Buffer.from(await blob.arrayBuffer());
  fs.writeFileSync(path.join(outputDir, `slide-${String(i + 1).padStart(2, "0")}.png`), buffer);
}

const inspect = await presentation.inspect();
fs.writeFileSync(path.join(outputDir, "inspect.ndjson"), inspect.ndjson, "utf8");
console.log(JSON.stringify({ pptxPath, outputDir, slides: count }, null, 2));
