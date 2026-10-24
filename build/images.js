// Picture-only copy of the lecture: each slide is one full-bleed image (rendered from the PDF),
// so nothing depends on fonts installed on the viewing machine. Speaker notes are carried over.
// Usage: node images.js <png-dir> <notes.json> <out.pptx>
const fs = require("fs");
const path = require("path");
const sharp = require("sharp");
const pptxgen = require("pptxgenjs");

async function main() {
  const [dir, notesFile, out] = process.argv.slice(2);
  const notes = JSON.parse(fs.readFileSync(notesFile, "utf8"));
  const pngs = fs.readdirSync(dir).filter((f) => f.endsWith(".png")).sort();
  if (pngs.length !== notes.length) throw new Error(`${pngs.length} images but ${notes.length} slides`);
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.title = "الكسوف الكلي للشمس ٢ أغسطس ٢٠٢٧ — نسخة مصوّرة";
  for (let i = 0; i < pngs.length; i++) {
    const jpg = await sharp(path.join(dir, pngs[i])).jpeg({ quality: 90, mozjpeg: true }).toBuffer();
    const s = pres.addSlide();
    s.addImage({ data: "image/jpeg;base64," + jpg.toString("base64"), x: 0, y: 0, w: 13.333, h: 7.5,
      altText: "شريحة " + (i + 1) });
    if (notes[i]) s.addNotes(notes[i]);
  }
  await pres.writeFile({ fileName: out });
  console.log("wrote", out, pngs.length, "slides");
}
main().catch((e) => { console.error(e); process.exit(1); });
