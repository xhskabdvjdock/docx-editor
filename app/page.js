"use client";
import { useRef, useState } from "react";
import { APP, CORE, TYPES, batText, buildDocx, loadDocx, localInput, psLines, utcToLocalInput } from "../lib/docx";

function Field({ k, def, value, onChange }) {
  const [, tag, label, kind] = def;
  const id = `${k}_${tag}`;
  const common = { id, value, onChange: (e) => onChange(id, e.target.value) };
  return (
    <div>
      <label htmlFor={id}>{label} ({tag})</label>
      {kind === "m" ? <textarea rows={3} {...common} />
        : kind === "d" ? <input type="datetime-local" step="1" className="ltr" {...common} />
        : <input className={kind === "l" ? "ltr" : ""} {...common} />}
    </div>
  );
}

export default function Page() {
  const [doc, setDoc] = useState(null);
  const [file, setFile] = useState(null);
  const [values, setValues] = useState({});
  const [custom, setCustom] = useState([]);
  const [msg, setMsg] = useState({ t: "", c: "" });
  const [over, setOver] = useState(false);
  const [sys, setSys] = useState({ created: "", modified: "", accessed: "" });
  const input = useRef(null);

  async function open(f) {
    if (!f) return;
    try {
      const d = await loadDocx(f);
      const l = localInput(new Date(f.lastModified)); setSys({ created: l, modified: l, accessed: l });
      setDoc(d); setFile(f); setValues(d.values); setCustom(d.custom); setMsg({ t: "", c: "" });
    } catch {
      setMsg({ t: "تعذّرت قراءة الملف. تأكد أنه ملف Word بصيغة docx صحيحة.", c: "var(--er)" });
    }
  }

  async function exportDoc() {
    try {
      const blob = await buildDocx({ ...doc, values, custom });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = outName;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 4000);
      setMsg({ t: "تم تصدير الملف ✓", c: "var(--ok)" });
    } catch (e) {
      setMsg({ t: "حدث خطأ أثناء التصدير: " + e.message, c: "var(--er)" });
    }
  }

  const outName = file ? file.name.replace(/\.[^.]+$/, "") + "-edited.docx" : "";
  const sameAsInternal = () => {
    const c = utcToLocalInput(values.core_created), m = utcToLocalInput(values.core_modified);
    setSys({ created: c, modified: m || c, accessed: m || c });
  };
  const saveBat = () => {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([batText(outName, sys)], { type: "text/plain" }));
    a.download = "set-file-dates.bat";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  };
  const setVal = (id, v) => setValues((s) => ({ ...s, [id]: v }));
  const setC = (n, patch) => setCustom((s) => s.map((c, i) => (i === n ? { ...c, ...patch } : c)));

  return (
    <main>
      <h1>محرر خصائص ملفات Word</h1>
      <p className="s">عدّل كل خصائص ملف docx (Properties ← Details) ثم صدّره بدون المساس بالمحتوى. المعالجة تتم بالكامل على جهازك، ولا يُرفع الملف لأي مكان.</p>

      <div className={"drop" + (over ? " on" : "")} onClick={() => input.current.click()}
        onDragOver={(e) => { e.preventDefault(); setOver(true); }} onDragLeave={() => setOver(false)}
        onDrop={(e) => { e.preventDefault(); setOver(false); open(e.dataTransfer.files[0]); }}>
        <b>اضغط لاختيار ملف .docx</b> أو اسحبه هنا
        <input ref={input} type="file" accept=".docx,.dotx,.docm" hidden onChange={(e) => open(e.target.files[0])} />
      </div>

      {!doc && msg.t && <p style={{ color: msg.c }}>{msg.t}</p>}

      {doc && (
        <>
          <p className="s" style={{ marginTop: 14 }}>الملف: {file.name}</p>
          <div className="card"><h2>الوصف (Description)</h2>
            <div className="grid">{CORE.map((d) => <Field key={d[1]} k="core" def={d} value={values[`core_${d[1]}`] || ""} onChange={setVal} />)}</div></div>
          <div className="card"><h2>الأصل والإحصائيات (Origin / Content)</h2>
            <div className="grid">{APP.map((d) => <Field key={d[1]} k="app" def={d} value={values[`app_${d[1]}`] || ""} onChange={setVal} />)}</div>
            <p className="hint">TotalTime = إجمالي وقت التحرير بالدقائق.</p></div>
          <div className="card"><h2>خصائص مخصصة (Custom)</h2>
            {custom.map((c, n) => (
              <div className="row" key={n}>
                <input value={c.name} placeholder="الاسم" onChange={(e) => setC(n, { name: e.target.value })} />
                <select value={c.tag} onChange={(e) => setC(n, { tag: e.target.value })}>
                  {[...new Set([...TYPES, c.tag])].map((t) => <option key={t}>{t}</option>)}
                </select>
                <input className="ltr" value={c.val} placeholder="القيمة" onChange={(e) => setC(n, { val: e.target.value })} />
                <button className="d" onClick={() => setCustom((s) => s.filter((_, i) => i !== n))}>حذف</button>
              </div>
            ))}
            <button className="g" onClick={() => setCustom((s) => [...s, { name: "", tag: "lpwstr", val: "" }])}>+ إضافة خاصية</button></div>
          <div className="card"><h2>تواريخ الملف في النظام (Date created / modified / accessed)</h2>
            <p className="hint" style={{ marginTop: 0, marginBottom: 12 }}>هذه التواريخ هي التي يعرضها Windows في Properties ← Details ← File، وهي خارج الملف نفسه. المتصفح لا يستطيع تغييرها، لذلك يولّد الموقع أمراً تشغّله على جهازك بعد التصدير. التوقيت هنا محلي (توقيت جهازك).</p>
            <div className="grid">
              {[["created", "Date created"], ["modified", "Date modified"], ["accessed", "Date accessed"]].map(([k, lb]) => (
                <div key={k}><label htmlFor={"sys_" + k}>{lb}</label>
                  <input id={"sys_" + k} type="datetime-local" step="1" className="ltr" value={sys[k]} onChange={(e) => setSys((s) => ({ ...s, [k]: e.target.value }))} /></div>
              ))}
            </div>
            <div className="bar" style={{ margin: "12px 0" }}>
              <button className="g" onClick={sameAsInternal}>نفس التواريخ الداخلية أعلاه</button>
              <button onClick={saveBat}>تنزيل set-file-dates.bat</button>
              <button className="g" onClick={() => navigator.clipboard.writeText(psLines(outName, sys).join("\n"))}>نسخ أوامر PowerShell</button>
            </div>
            <pre className="ltr" style={{ margin: 0, padding: 10, background: "var(--bg)", border: "1px solid var(--bd)", borderRadius: 8, overflowX: "auto", fontSize: 12.5 }}>{psLines(outName, sys).join("\n")}</pre>
            <p className="hint">الخطوات: صدّر الملف ← ضع set-file-dates.bat في نفس مجلده ← انقر عليه مرتين. (أو الصق الأوامر في PowerShell داخل نفس المجلد.) إن غيّرت اسم الملف بعد التصدير فعدّل الاسم في الأمر.</p></div>
          <div className="bar">
            <button onClick={exportDoc}>تصدير الملف المعدّل</button>
            <button className="g" onClick={() => open(file)}>إعادة القيم الأصلية</button>
            <span style={{ color: msg.c || "var(--mut)" }}>{msg.t}</span>
          </div>
          <p className="hint">ملاحظة: نسخ الملف إلى فلاش/ضغطه/رفعه قد يغيّر تواريخ النظام، فاضبطها على النسخة النهائية.</p>
        </>
      )}
    </main>
  );
}
