import JSZip from "jszip";

export const NS = {
  cp: "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
  dc: "http://purl.org/dc/elements/1.1/",
  dcterms: "http://purl.org/dc/terms/",
  xsi: "http://www.w3.org/2001/XMLSchema-instance",
  ep: "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
  cu: "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
  vt: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
};

// [prefix, tag, label, kind]  kind: "d" تاريخ، "t" نص عادي (RTL)، "l" رقم/قيمة LTR، "m" نص طويل
export const CORE = [
  ["dc", "title", "العنوان", "t"], ["dc", "subject", "الموضوع", "t"], ["dc", "creator", "المؤلف", "t"],
  ["cp", "keywords", "الكلمات المفتاحية", "t"], ["dc", "description", "التعليقات", "m"],
  ["cp", "category", "الفئة", "t"], ["cp", "contentStatus", "الحالة", "t"], ["dc", "language", "اللغة", "l"],
  ["cp", "version", "الإصدار", "l"], ["dc", "identifier", "المعرّف", "l"],
  ["cp", "lastModifiedBy", "آخر من عدّل", "t"], ["cp", "revision", "رقم المراجعة", "l"],
  ["dcterms", "created", "تاريخ الإنشاء (UTC)", "d"], ["dcterms", "modified", "تاريخ آخر تعديل (UTC)", "d"],
  ["dcterms", "lastPrinted", "آخر طباعة (UTC)", "d"],
];
export const APP = [
  ["ep", "Application", "التطبيق", "l"], ["ep", "AppVersion", "إصدار التطبيق", "l"],
  ["ep", "Company", "الشركة", "t"], ["ep", "Manager", "المدير", "t"], ["ep", "Template", "القالب", "l"],
  ["ep", "TotalTime", "وقت التحرير (دقائق)", "l"], ["ep", "Pages", "الصفحات", "l"], ["ep", "Words", "الكلمات", "l"],
  ["ep", "Characters", "الأحرف", "l"], ["ep", "CharactersWithSpaces", "الأحرف مع المسافات", "l"],
  ["ep", "Lines", "الأسطر", "l"], ["ep", "Paragraphs", "الفقرات", "l"],
  ["ep", "DocSecurity", "الحماية (DocSecurity)", "l"], ["ep", "HyperlinkBase", "أساس الروابط", "l"],
];
export const TYPES = ["lpwstr", "i4", "r8", "bool", "filetime"];

const PARTS = {
  core: { p: "docProps/core.xml", ct: "application/vnd.openxmlformats-package.core-properties+xml",
    rt: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
    root: `<cp:coreProperties xmlns:cp="${NS.cp}" xmlns:dc="${NS.dc}" xmlns:dcterms="${NS.dcterms}" xmlns:xsi="${NS.xsi}"/>` },
  app: { p: "docProps/app.xml", ct: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
    rt: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
    root: `<Properties xmlns="${NS.ep}"/>` },
  cust: { p: "docProps/custom.xml", ct: "application/vnd.openxmlformats-officedocument.custom-properties+xml",
    rt: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
    root: `<Properties xmlns="${NS.cu}" xmlns:vt="${NS.vt}"/>` },
};

const parse = (s) => new DOMParser().parseFromString(s, "application/xml");
const ser = (d) => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + new XMLSerializer().serializeToString(d.documentElement);
const getV = (d, ns, tag) => (d.getElementsByTagNameNS(NS[ns], tag)[0]?.textContent || "").trim();

export async function loadDocx(file) {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  if (!zip.file("[Content_Types].xml") || !zip.file("_rels/.rels")) throw new Error("not docx");
  const docs = {};
  for (const k of ["core", "app", "cust"]) {
    const e = zip.file(PARTS[k].p);
    docs[k] = parse(e ? await e.async("string") : PARTS[k].root);
  }
  const values = {};
  for (const [k, list] of [["core", CORE], ["app", APP]])
    for (const [ns, tag, , kind] of list) {
      const v = getV(docs[k], ns, tag);
      values[`${k}_${tag}`] = kind === "d" ? v.slice(0, 19) : v;
    }
  const custom = [...docs.cust.getElementsByTagNameNS(NS.cu, "property")].map((p) => {
    const c = p.firstElementChild;
    return { name: p.getAttribute("name"), tag: c ? c.localName : "lpwstr", val: c ? c.textContent : "" };
  });
  return { zip, docs, values, custom };
}

function setEl(d, ns, tag, val, isDate) {
  let e = d.getElementsByTagNameNS(NS[ns], tag)[0];
  if (!val) { e?.remove(); return; }
  if (!e) { e = d.createElementNS(NS[ns], `${ns}:${tag}`); d.documentElement.appendChild(e); }
  e.textContent = val;
  if (isDate) e.setAttributeNS(NS.xsi, "xsi:type", "dcterms:W3CDTF");
}

async function ensure(zip, k) {
  const P = PARTS[k];
  if (zip.file(P.p)) return;
  const ct = await zip.file("[Content_Types].xml").async("string");
  zip.file("[Content_Types].xml", ct.replace("</Types>", `<Override PartName="/${P.p}" ContentType="${P.ct}"/></Types>`));
  const rl = await zip.file("_rels/.rels").async("string");
  zip.file("_rels/.rels", rl.replace("</Relationships>", `<Relationship Id="rIdProp${k}" Type="${P.rt}" Target="${P.p}"/></Relationships>`));
}

const withSeconds = (v) => (v.length === 16 ? v + ":00" : v) + "Z";

// يعدّل ملفات docProps فقط؛ بقية محتويات الـ docx تبقى كما هي دون أي تغيير.
export async function buildDocx({ zip, docs, values, custom }) {
  for (const [k, list] of [["core", CORE], ["app", APP]])
    for (const [ns, tag, , kind] of list) {
      let v = (values[`${k}_${tag}`] || "").trim();
      if (kind === "d" && v) v = withSeconds(v);
      setEl(docs[k], ns, tag, v, kind === "d");
    }
  const cd = parse(PARTS.cust.root);
  custom.filter((c) => c.name.trim()).forEach((c, n) => {
    const p = cd.createElementNS(NS.cu, "property");
    p.setAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
    p.setAttribute("pid", n + 2);
    p.setAttribute("name", c.name.trim());
    const e = cd.createElementNS(NS.vt, "vt:" + c.tag);
    let v = c.val;
    if (c.tag === "bool") v = /^(true|1|yes|نعم)$/i.test(v) ? "true" : "false";
    if (c.tag === "filetime" && v) v = /Z$/.test(v) ? v : withSeconds(v);
    e.textContent = v;
    p.appendChild(e);
    cd.documentElement.appendChild(p);
  });
  for (const k of ["core", "app"]) { await ensure(zip, k); zip.file(PARTS[k].p, ser(docs[k])); }
  if (custom.some((c) => c.name.trim()) || zip.file(PARTS.cust.p)) {
    await ensure(zip, "cust");
    zip.file(PARTS.cust.p, ser(cd));
  }
  return zip.generateAsync({
    type: "blob", compression: "DEFLATE",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}
