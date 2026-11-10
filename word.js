/* Word export adapted from the supplied HTML: shapes as a high-resolution image, editable text boxes. */
function loadImg(svg){ return new Promise((res, rej) => { const im = new Image(); im.onload = () => res(im); im.onerror = () => rej(new Error("svg")); im.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg); }); }
async function toPng(p, withText){
  const sc = 2, o = withText ? { labels: false, bg: true } : { labels: false, bg: false, noDrawText: true };
  const { svg, m } = fullSvg(p, [], o), im = await loadImg(svg);
  const c = document.createElement("canvas"); c.width = W * sc; c.height = H * sc; const x = c.getContext("2d"); x.drawImage(im, 0, 0, c.width, c.height);
  if (withText) {
    try { await Promise.all(["12px Arial", "bold 12px Arial", "italic 12px Arial"].map(f => document.fonts.load(f))); } catch (e) {}
    x.scale(sc, sc); x.lineJoin = "round";
    for (const it of m.labels) { const d = [0, 0], n = it.lines.length, y0 = it.y - (n - 1) * it.size * 1.15 / 2 + it.size * 0.35;
      x.font = `${it.style} ${it.weight} ${it.size}px ${it.family}`; x.textAlign = it.anchor === "end" ? "right" : it.anchor === "middle" ? "center" : "left";
      it.lines.forEach((ln, i) => { const yy = y0 + i * it.size * 1.15 + d[1], xx = it.x + d[0];
        if (it.halo) { x.strokeStyle = "#FFFFFF"; x.lineWidth = 3.5; x.strokeText(ln, xx, yy); } x.fillStyle = it.fill; x.fillText(ln, xx, yy); }); }
  }
  const blob = await new Promise((res, rej) => c.toBlob(b => b ? res(b) : rej(new Error("png")), "image/png"));
  return { blob, m };
}
/* Word: grafica come immagine sullo sfondo, etichette come caselle di testo modificabili (stessa costruzione usata finora) */
const EMU = 6300, OFFX = 621000, OFFY = 440000, PT_PX = 0.496;
function docxParts(items){
  let uid = 100; const nid = () => ++uid;
  const xe = t => String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const colx = c => { c = c.replace("#", "").toUpperCase(); return c.length === 3 ? [...c].map(q => q + q).join("") : c; };
  const anchor = (g, x, y, cx, cy, behind, name) => { const i = nid();
    return `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${i}" behindDoc="${behind ? 1 : 0}" locked="${behind ? 1 : 0}" layoutInCell="1" allowOverlap="1">` +
      `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${Math.round(x)}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>${Math.round(y)}</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="${Math.round(cx)}" cy="${Math.round(cy)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${i}" name="${name} ${i}"/><wp:cNvGraphicFramePr/>${g}</wp:anchor></w:drawing></w:r>`; };
  const pic = (rid, cx, cy) => { const i = nid();
    return `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:nvPicPr><pic:cNvPr id="${i}" name="grafica${i}.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(cx)}" cy="${Math.round(cy)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`; };
  const textbox = it => {
    const size = it.size, lines = it.lines, n = lines.length;
    const wpx = Math.max(...lines.map(l => l.length)) * size * (it.family.includes("Serif") ? 0.68 : 0.58) + 14, hpx = n * size * 1.15 + size * 0.6 + 4;
    const y0b = it.y - (n - 1) * size * 1.15 / 2 + size * 0.35, top = y0b - size - 2;
    const left = it.anchor === "start" ? it.x : it.anchor === "end" ? it.x - wpx : it.x - wpx / 2, jc = { start: "left", end: "right", middle: "center" }[it.anchor];
    const font = it.family.includes("Serif") ? "Times New Roman" : "Calibri", sz = Math.max(8, Math.round(size * PT_PX * 2));
    const rpr = `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/>${it.weight === "bold" ? "<w:b/>" : ""}${it.style === "italic" ? "<w:i/>" : ""}<w:color w:val="${colx(it.fill)}"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr>`;
    const runs = lines.map((l, k) => (k ? `<w:r>${rpr}<w:br/></w:r>` : "") + `<w:r>${rpr}<w:t xml:space="preserve">${xe(l)}</w:t></w:r>`).join("");
    const para = `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${Math.round(size * 1.15 * PT_PX * 20)}" w:lineRule="exact"/><w:jc w:val="${jc}"/></w:pPr>${runs}</w:p>`;
    const g = `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>` +
      `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(wpx * EMU)}" cy="${Math.round(hpx * EMU)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>` +
      `<wps:txbx><w:txbxContent>${para}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic>`;
    let dml = anchor(g, OFFX + left * EMU, OFFY + top * EMU, wpx * EMU, hpx * EMU, false, "Etichetta");
    dml = dml.slice(5, -6); const i = nid(), ptv = v => (v / 12700).toFixed(2) + "pt";
    const vml = `<w:pict><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>` +
      `<v:shape id="Etichetta${i}" o:spid="_x0000_s${i}" type="#_x0000_t202" filled="f" stroked="f" style="position:absolute;margin-left:${ptv(OFFX + left * EMU)};margin-top:${ptv(OFFY + top * EMU)};width:${ptv(wpx * EMU)};height:${ptv(hpx * EMU)};z-index:${i};mso-position-horizontal-relative:page;mso-position-vertical-relative:page">` +
      `<v:textbox inset="0,0,0,0"><w:txbxContent>${para}</w:txbxContent></v:textbox></v:shape></w:pict>`;
    return `<w:r><mc:AlternateContent><mc:Choice Requires="wps">${dml}</mc:Choice><mc:Fallback>${vml}</mc:Fallback></mc:AlternateContent></w:r>`;
  };
  const body = [], rels = [];
  items.forEach((it, k) => { const rid = "rIdImg" + k;
    rels.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/tavola${k}.png"/>`);
    const img = anchor(pic(rid, W * EMU, H * EMU), OFFX, OFFY, W * EMU, H * EMU, true, "Grafica");
    body.push(`<w:p><w:pPr>${k ? "<w:pageBreakBefore/>" : ""}<w:spacing w:before="0" w:after="0"/></w:pPr>${img}${it.labels.map(textbox).join("")}</w:p>`); });
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"';
  const X = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  return {
    "[Content_Types].xml": X + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
    "_rels/.rels": X + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    "word/document.xml": X + `<w:document ${NS}><w:body>${body.join("")}<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:body></w:document>`,
    "word/styles.xml": X + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="12"/><w:lang w:val="it-IT"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>',
    "word/settings.xml": X + '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>',
    "word/_rels/document.xml.rels": X + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdSe" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>' + rels.join("") + "</Relationships>"
  };
}
async function buildDocx(plates){
  if (!window.JSZip) throw new Error("zip");
  const items = [];
  for (const p of plates) {
    const { blob, m } = await toPng(p, false), off = {};
    const labels = m.labels.map(it => { const d = off[it.key] || [0, 0]; return { ...it, x: it.x + d[0], y: it.y + d[1] }; });
    for (const d of []) if (d.k === "t") { const t = d.t || [0, 0]; labels.push({ x: d.x1 + t[0], y: d.y1 + t[1] - (d.s || 14) * 0.35, size: d.s || 14, anchor: "start", fill: d.c, lines: [d.txt], family: F_SANS, weight: "normal", style: "normal" }); }
    items.push({ blob, labels });
  }
  const zip = new window.JSZip(), parts = docxParts(items);
  for (const [k, v] of Object.entries(parts)) zip.file(k, v);
  items.forEach((it, k) => zip.file("word/media/tavola" + k + ".png", it.blob));
  return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", compression: "DEFLATE" });
}
