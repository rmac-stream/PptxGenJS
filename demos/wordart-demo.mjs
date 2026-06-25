/**
 * Standalone WordArt demo — generates WordArt_Demo.pptx
 * Run: node demos/wordart-demo.mjs
 */
import pptxgen from "../src/bld/pptxgen.es.js";

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.title = "WordArt Demo";

// ---- Slide 1: warp transforms ----
let s1 = pptx.addSlide();
s1.background = { color: "FFFFFF" };
s1.addText("WordArt: Transforms (presetShape)", { x: 0.5, y: 0.2, w: 12.3, h: 0.5, fontSize: 20, bold: true, color: "666666" });

s1.addWordArt("Arch Up", { x: 0.5, y: 0.9, w: 4.0, h: 1.4, fontSize: 40, bold: true, color: "4472C4", presetShape: "textArchUp" });
s1.addWordArt("Wave", { x: 4.7, y: 0.9, w: 4.0, h: 1.4, fontSize: 40, bold: true, color: "ED7D31", presetShape: "textWave1" });
s1.addWordArt("Circle", { x: 9.0, y: 0.8, w: 3.5, h: 1.6, fontSize: 28, bold: true, color: "70AD47", presetShape: "textCircle" });
s1.addWordArt("Inflate", { x: 0.5, y: 2.6, w: 4.0, h: 1.4, fontSize: 40, bold: true, color: "A5A5A5", presetShape: "textInflate" });
s1.addWordArt("Chevron", { x: 4.7, y: 2.6, w: 4.0, h: 1.4, fontSize: 40, bold: true, color: "FFC000", presetShape: "textChevron" });
s1.addWordArt("Cascade", { x: 9.0, y: 2.6, w: 4.0, h: 1.4, fontSize: 40, bold: true, color: "5B9BD5", presetShape: "textCascadeDown" });

// ---- Slide 2: gradient fills ----
let s2 = pptx.addSlide();
s2.background = { color: "1A1A1A" };
s2.addText("WordArt: Gradient Fills", { x: 0.5, y: 0.2, w: 12.3, h: 0.5, fontSize: 20, bold: true, color: "CCCCCC" });

s2.addWordArt("Sunset", {
	x: 0.5, y: 1.0, w: 12.0, h: 1.6, fontSize: 88, bold: true,
	gradient: { angle: 45, stops: [{ color: "FF512F", position: 0 }, { color: "F09819", position: 100 }] },
});
s2.addWordArt("Ocean", {
	x: 0.5, y: 3.0, w: 12.0, h: 1.6, fontSize: 88, bold: true,
	gradient: { angle: 90, stops: [{ color: "2193B0", position: 0 }, { color: "6DD5ED", position: 50 }, { color: "FFFFFF", position: 100 }] },
});

// ---- Slide 3: the works — warp + gradient + outline ----
let s3 = pptx.addSlide();
s3.background = { color: "FFFFFF" };
s3.addText("WordArt: Warp + Gradient + Outline", { x: 0.5, y: 0.2, w: 12.3, h: 0.5, fontSize: 20, bold: true, color: "666666" });

s3.addWordArt("PptxGenJS", {
	x: 0.5, y: 1.4, w: 12.0, h: 2.6, fontSize: 110, bold: true,
	presetShape: "textInflate",
	gradient: { angle: 90, stops: [{ color: "FFC000", position: 0 }, { color: "ED7D31", position: 50 }, { color: "C00000", position: 100 }] },
	outline: { color: "404040", size: 1.25 },
});

await pptx.writeFile({ fileName: "WordArt_Demo.pptx" });
console.log("Wrote WordArt_Demo.pptx");
