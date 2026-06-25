/**
 * NAME: demo_wordart.mjs
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DESC: Common test/demo slides for all library features
 * DEPS: Used by various demos (./demos/browser, ./demos/node, etc.)
 * VER.: 4.0.1
 */

import { BASE_TABLE_OPTS, BASE_TEXT_OPTS_L, BASE_TEXT_OPTS_R } from "./enums.mjs";

export function genSlides_WordArt(pptx) {
	pptx.addSection({ title: "WordArt" });

	let slide = pptx.addSlide({ sectionTitle: "WordArt" });
	slide.addTable([[{ text: "WordArt Examples", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	// 1: Warp only
	slide.addWordArt("Arch Up", {
		x: 0.5, y: 0.9, w: 5.5, h: 1.2,
		fontSize: 48, bold: true, color: pptx.colors.ACCENT1,
		presetShape: "textArchUp",
	});

	// 2: Gradient text fill (no warp)
	slide.addWordArt("Gradient", {
		x: 7.0, y: 0.9, w: 5.5, h: 1.2,
		fontSize: 54, bold: true,
		gradient: { angle: 45, stops: [{ color: "FF0000", position: 0 }, { color: "0000FF", position: 100 }] },
	});

	// 3: Warp + gradient + outline (full WordArt)
	slide.addWordArt("PptxGenJS", {
		x: 0.5, y: 2.6, w: 12.0, h: 2.2,
		fontSize: 96, bold: true,
		presetShape: "textInflate",
		gradient: { angle: 90, stops: [{ color: "FFC000", position: 0 }, { color: "ED7D31", position: 50 }, { color: "C00000", position: 100 }] },
		outline: { color: "404040", size: 1 },
	});

	return slide;
}
