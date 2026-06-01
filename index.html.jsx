import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

const INSPECTION_CODES = {
  ZRM: "ZRM – Zero Measurement",
  EXC: "EXC – Excavation",
  GEM: "GEM – Geomechanics-Work/Test",
  IN1: "IN1 – In situ rebar formwork – Top Slab",
  IN2: "IN2 – In situ concreting – Lean Concrete",
  IST: "IST – Installation – Geotextile",
  WP3: "WP3 – Pile Foundation Waterproofing",
  CON: "CON – Concreting Works",
  INS: "INS – Installation Works",
  GEO: "GEO – Geomechanics Works",
  OTHER: "Other (specify below)",
};

const DETAIL_CODES = { IR: "IR – Irrigation", WS: "WS – Water Supply" };

const INS_FULL = {
  ZRM: "ZRM – Zero Measurement",
  EXC: "EXC- Excavation",
  GEM: "GEM- Geomechanics-Work/Test",
  IN1: "IN1-In situ rebar formwork – Top Slab",
  IN2: "IN2-In situ concreting– Lean Concrete",
  IST: "IST - Installation - Geotextile",
  WP3: "WP3-Pile Foundation Waterproofing – bituminous emulsion",
  CON: "CON – Concreting Works",
  INS: "INS – Installation Works",
  GEO: "GEO – Geomechanics Works",
};

const emptyRSC = () => ({
  id: Date.now() + Math.random(),
  rscno: "",
  docdate: "",
  inscode: "ZRM",
  inscodeother: "",
  detcode: "IR",
  seqcode: "",
  section: "04C-Section 4C",
  object: "",
  startch: "",
  endch: "",
  side: "",
  estdt: "",
  desc: "",
  msc: "N/A",
  dwg: "",
  qcp: "",
  _autoSeq: null,
});

function padSeq(n) {
  return String(n).padStart(5, "0");
}

function buildRSCNo(location, prefix, type, seq) {
  const loc = location.replace("KM ", "");
  return `26437-RSC-04C-UT-${type}-${prefix}-${padSeq(seq)}`;
}

function buildRSCContent(r) {
  const insCodeFull =
    r.inscode === "OTHER"
      ? r.inscodeother
      : INS_FULL[r.inscode] || INSPECTION_CODES[r.inscode] || r.inscode;
  const detCodeFull = DETAIL_CODES[r.detcode] || r.detcode;
  const sep = "━".repeat(70);
  return `NORTH MACEDONIA CORRIDOR 8 AND CORRIDOR 10-D MOTORWAY PROJECT
REQUEST FOR INSPECTION, TESTING AND APPROVAL FOR SPECIALIST CONTRACTOR (RSC)
${sep}

Specialist Contractor : AK INVEST DOOEL TETOVO
Employer              : Public Enterprise for State Roads (PESR)
Engineer              : IRD Engineering SRL

${sep}
PROJECT CODE      : 26437
SECTION CODE      : ${r.section}
DISCIPLINE CODE   : UT1 – Specialist Contractor 1
INSPECTION CODE   : ${insCodeFull}
INSPECTION DETAIL : ${detCodeFull}
SEQ. CODE         : ${r.seqcode}

RSC DOCUMENT NUMBER : ${r.rscno} [rev:00]
DOC. DATE           : ${r.docdate}

NAME OF SC   : AK INVEST DOOEL TETOVO
CONTRACT NO. : MK102-UT1-SOSC00-00004

${sep}
LOCATION INFORMATION
${sep}
Section  : ${r.section}
Object   : ${r.object}
Element  : N/A
Chainage : Start: ${r.startch}    End: ${r.endch}    Side: ${r.side}

${sep}
DESCRIPTION OF WORK
${sep}
${r.desc}

${sep}
REFERENCES
${sep}
Material Approval Request (MSC) No : ${r.msc}
Drawing No                         : ${r.dwg}
Related QCP No                     : ${r.qcp}

${sep}
REQUEST FOR INSPECTION AND TESTING FOR SPECIALIST CONTRACTOR
${sep}
The Special Contractor hereby notifies the Engineer and the Contractor that the
activity described above of the works is ready for inspection and testing.

ISSUED BY (Name and Surname & Signature):
SPECIALIST CONTRACTOR  : AK INVEST DOOEL TETOVO
NAME AND SURNAME       : NENAD GAJIC
SIGNATURE              : ____________________

ESTIMATED DATE AND TIME FOR INSPECTION: ${r.estdt}

Note: The exact timing of inspection and testing is to be communicated by the
Special Contractor to the Engineer and Contractor during the day in a timely manner.

${sep}
REQUEST FOR APPROVAL FOR SPECIALIST CONTRACTOR
${sep}
Described works above have been completed in compliance with the technical
specifications, Quality Control Plans (QCP), and design requirements. the Contractor
and Contractor-Related Parties shall not be responsible for and shall be released from
any liability whatsoever arising out of and/or in connection with the services and/or
works performed by Specialist Contractors.

SUBMITTED BY THE SPECIALIST CONTRACTOR:
THE SPECIALIST CONTRACTOR'S SITE SUPERVISOR:
NAME AND SURNAME : NENAD GAJIC
SIGNATURE        : ____________________

ATTACHMENTS:
[  ]

${sep}
EVALUATION BY THE ENGINEER (Name and Surname & Signature):
${sep}
NAME AND SURNAME       SIGNATURE         DATE
SITE INSPECTOR    :    ______________    __________
SURVEY ENGINEER   :    ______________    __________
QC ENGINEER       :    ______________    __________

Comments (if any):
_______________________________________________

○ A – Approved        ○ B – Approved with Comment        ○ C – Not Approved

For the Engineer:
Name and Surname: ------------------------
Signature: --------------------
Date: ----------------

Level 3 - BEJV Internal | BEJV Data Sensitivity Level`;
}

function buildNotifText(rscs) {
  const line = "─".repeat(100);
  let out = `${line}\nDAILY INSPECTION NOTIFICATION LIST – UTILITIES\n${line}\n\n`;
  rscs.forEach((r, i) => {
    out += `${i + 1}  | RSC: ${r.rscno}\n`;
    out += `   | Section: Section 4  |  Object: ${r.startch}  |  From: ${r.startch}  |  To: ${r.endch}  |  Side/Element: ${r.side}\n`;
    out += `   | Estimated: ${r.estdt.replace(/\//g, ".")}\n`;
    out += `   | Item: ${r.desc}\n`;
    out += `${line}\n`;
  });
  return out;
}

function buildEmail(rscs, dateLabel, recipient, emailType) {
  const subjType = emailType === "resubmittal" ? "Resubmittal" : "Submittal";
  const intro =
    emailType === "resubmittal"
      ? `Please find attached the resubmitted RSCs and the Daily Notification Invitation List for ${dateLabel}. A mistake was identified in the documents, and a new RSC has also been submitted.`
      : `Please find attached the RSCs and the Daily Notification Invitation List for ${dateLabel}.`;
  const lines = rscs.map((r) => `${r.rscno} - ${r.desc};`).join("\n\n");
  return `Subject: ${subjType} of RSC's and Daily Notification list for ${dateLabel}\n\nDear ${recipient},\n\n${intro}\n\nRSCs:\n\n${lines}\n\nDaily Notification Invitation List ${dateLabel}, in PDF format.\n\nContact list From Ak Invest:\n\nSite Engineer: Darko Tanushevski 075-295-438\nProject Coordinator: Nenad Gajic 072-320-544\nSurvey Engineer: Martin Nasevski 072-240-604\n\nUtility Department Contact List: Section 4: Site Supervisor: Sefa Gucum — 072 921 037\n\nBest Regards,\nNenad`;
}

function CopyBtn({ getText }) {
  const [ok, setOk] = useState(false);
  return (
    <button
      onClick={() => navigator.clipboard.writeText(getText()).then(() => { setOk(true); setTimeout(() => setOk(false), 1800); })}
      style={{ padding: "4px 12px", fontSize: 12, borderRadius: 6, border: "1px solid #ccc", cursor: "pointer", background: ok ? "#e6f4ea" : "#fff", color: ok ? "#2e7d32" : "#555", fontFamily: "inherit" }}
    >{ok ? "✓ Copied" : "Copy"}</button>
  );
}

function OutBox({ label, text }) {
  return (
    <div style={{ marginBottom: 18 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 5 }}>
        <span style={{ fontWeight: 600, fontSize: 13 }}>{label}</span>
        <CopyBtn getText={() => text} />
      </div>
      <pre style={{ background: "#f7f7f7", border: "1px solid #e0e0e0", borderRadius: 8, padding: "10px 13px", fontSize: 11.5, fontFamily: "monospace", whiteSpace: "pre-wrap", maxHeight: 280, overflowY: "auto", lineHeight: 1.6, color: "#222" }}>{text}</pre>
    </div>
  );
}

const inp = { width: "100%", padding: "6px 8px", borderRadius: 6, border: "1px solid #ddd", fontSize: 12.5, fontFamily: "inherit", boxSizing: "border-box" };
const lbl = { display: "block", fontSize: 11, color: "#777", marginBottom: 3 };

function RSCForm({ rsc, idx, onChange, onDelete, excelData }) {
  const set = (field) => (e) => {
    const val = e.target.value;
    let updated = { ...rsc, [field]: val };

    if ((field === "startch" || field === "inscode") && excelData) {
      const loc = field === "startch" ? val : rsc.startch;
      const type = field === "inscode" ? val : rsc.inscode;
      if (loc && type && type !== "OTHER") {
        const prefix = excelData.lokacii[loc] || rsc.detcode;
        const counts = excelData.brojaci[type];
        let nextSeq = 1;
        if (counts) {
          nextSeq = (prefix === "IR" ? (counts.lastIR || 0) : (counts.lastWS || 0)) + 1;
        }
        const newRscNo = buildRSCNo(loc, prefix, type, nextSeq);
        updated = {
          ...updated,
          detcode: prefix,
          seqcode: padSeq(nextSeq),
          rscno: newRscNo,
          _autoSeq: nextSeq,
          _autoPrefix: prefix,
          _autoLoc: loc,
          _autoType: type,
        };
      }
    }
    onChange(updated);
  };

  const autoHint = rsc._autoSeq
    ? { background: "#eef7ee", border: "1px solid #b7ddb7" }
    : {};

  return (
    <div style={{ border: "1px solid #e0e0e0", borderRadius: 10, padding: "14px 16px", marginBottom: 12, background: "#fafafa", position: "relative" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
        <span style={{ fontWeight: 600, fontSize: 13, color: "#333" }}>RSC #{idx + 1}</span>
        {idx > 0 && <button onClick={onDelete} style={{ background: "none", border: "none", cursor: "pointer", color: "#bbb", fontSize: 16 }}>✕</button>}
      </div>

      {excelData && (
        <div style={{ marginBottom: 10 }}>
          <label style={lbl}>Локација (KM) — автоматски број</label>
          <select style={{ ...inp, ...autoHint }} value={rsc.startch} onChange={set("startch")}>
            <option value="">— Избери локација —</option>
            {Object.keys(excelData.lokacii).map(k => <option key={k} value={k}>{k} ({excelData.lokacii[k]})</option>)}
            <option value="manual">Рачно внесување</option>
          </select>
          {rsc._autoSeq && (
            <div style={{ fontSize: 11, color: "#2e7d32", marginTop: 3 }}>
              ✓ Автоматски: {rsc.rscno} (следен: #{rsc._autoSeq})
            </div>
          )}
        </div>
      )}

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        <div>
          <label style={lbl}>RSC Document Number</label>
          <input style={{ ...inp, ...autoHint }} value={rsc.rscno} onChange={set("rscno")} placeholder="26437-RSC-04C-UT-ZRM-IR-00003" />
        </div>
        <div>
          <label style={lbl}>Doc. Date (DD/MM/YYYY)</label>
          <input style={inp} value={rsc.docdate} onChange={set("docdate")} placeholder="01/06/2026" />
        </div>
        <div>
          <label style={lbl}>Inspection Code</label>
          <select style={inp} value={rsc.inscode} onChange={set("inscode")}>
            {Object.entries(INSPECTION_CODES).map(([k, v]) => <option key={k} value={k}>{v}</option>)}
          </select>
        </div>
        <div>
          <label style={lbl}>Inspection Code — ако е Other</label>
          <input style={inp} value={rsc.inscodeother} onChange={set("inscodeother")} placeholder="e.g. WP4-Custom" disabled={rsc.inscode !== "OTHER"} />
        </div>
        <div>
          <label style={lbl}>Inspection Detail (IR / WS)</label>
          <select style={inp} value={rsc.detcode} onChange={set("detcode")}>
            {Object.entries(DETAIL_CODES).map(([k, v]) => <option key={k} value={k}>{v}</option>)}
          </select>
        </div>
        <div>
          <label style={lbl}>Seq. Code</label>
          <input style={{ ...inp, ...autoHint }} value={rsc.seqcode} onChange={set("seqcode")} placeholder="00003" />
        </div>
        <div>
          <label style={lbl}>Section</label>
          <input style={inp} value={rsc.section} onChange={set("section")} />
        </div>
        <div>
          <label style={lbl}>Object</label>
          <input style={inp} value={rsc.object} onChange={set("object")} placeholder="Irrigation / Water Supply" />
        </div>
        <div>
          <label style={lbl}>Start Chainage</label>
          <input style={inp} value={rsc.startch === "manual" ? "" : rsc.startch} onChange={set("startch")} placeholder="KM 8+720" />
        </div>
        <div>
          <label style={lbl}>End Chainage</label>
          <input style={inp} value={rsc.endch} onChange={set("endch")} placeholder="KM 8+720" />
        </div>
        <div>
          <label style={lbl}>Side</label>
          <input style={inp} value={rsc.side} onChange={set("side")} placeholder="LHS" />
        </div>
        <div>
          <label style={lbl}>Estimated Date &amp; Time</label>
          <input style={inp} value={rsc.estdt} onChange={set("estdt")} placeholder="01/06/2026 09:00" />
        </div>
      </div>

      <div style={{ marginTop: 10 }}>
        <label style={lbl}>Description of Work</label>
        <textarea style={{ ...inp, minHeight: 54, resize: "vertical" }} value={rsc.desc} onChange={set("desc")} placeholder="Relocation of irrigation line at KM 8+720 – LHS – (from 0+022.70 to 0+038.42)" />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginTop: 10 }}>
        <div><label style={lbl}>Material Approval (MSC)</label><input style={inp} value={rsc.msc} onChange={set("msc")} /></div>
        <div><label style={lbl}>Drawing No</label><input style={inp} value={rsc.dwg} onChange={set("dwg")} placeholder="1; 2.1" /></div>
        <div><label style={lbl}>QCP No</label><input style={inp} value={rsc.qcp} onChange={set("qcp")} placeholder="26437-000-QC-QCP-00001" /></div>
      </div>
    </div>
  );
}

function HistoryTable({ vnes }) {
  return (
    <div style={{ overflowX: "auto" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
        <thead>
          <tr style={{ background: "#f0f0f0" }}>
            {["Дата", "Локација", "Тип", "Број", "RSC Код", "Префикс"].map(h => (
              <th key={h} style={{ padding: "5px 8px", textAlign: "left", border: "1px solid #e0e0e0", fontWeight: 600 }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {[...vnes].reverse().slice(0, 20).map((row, i) => (
            <tr key={i} style={{ background: i % 2 === 0 ? "#fff" : "#fafafa" }}>
              <td style={{ padding: "4px 8px", border: "1px solid #eee" }}>{row.date}</td>
              <td style={{ padding: "4px 8px", border: "1px solid #eee" }}>{row.loc}</td>
              <td style={{ padding: "4px 8px", border: "1px solid #eee" }}><span style={{ background: "#e8f0fe", color: "#1a73e8", padding: "1px 6px", borderRadius: 4, fontSize: 11 }}>{row.type}</span></td>
              <td style={{ padding: "4px 8px", border: "1px solid #eee", textAlign: "center" }}>{row.num}</td>
              <td style={{ padding: "4px 8px", border: "1px solid #eee", fontFamily: "monospace", fontSize: 11 }}>{row.code}</td>
              <td style={{ padding: "4px 8px", border: "1px solid #eee" }}>{row.prefix}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export default function App() {
  const [rscs, setRscs] = useState([emptyRSC()]);
  const [notifDate, setNotifDate] = useState(new Date().toISOString().split("T")[0]);
  const [recipient, setRecipient] = useState("Emil");
  const [emailType, setEmailType] = useState("submittal");
  const [outputs, setOutputs] = useState(null);
  const [tab, setTab] = useState("input");
  const [excelData, setExcelData] = useState(null);
  const [excelFileName, setExcelFileName] = useState(null);
  const [rawWorkbook, setRawWorkbook] = useState(null);
  const fileRef = useRef();

  const handleExcelUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const wb = XLSX.read(ev.target.result, { type: "array" });
      setRawWorkbook(wb);
      setExcelFileName(file.name);

      const bsheet = wb.Sheets["BROJACI"];
      const lsheet = wb.Sheets["LOKACII"];
      const vsheet = wb.Sheets["VNES"];

      const brows = XLSX.utils.sheet_to_json(bsheet, { header: 1 }).slice(1);
      const lrows = XLSX.utils.sheet_to_json(lsheet, { header: 1 }).slice(1);
      const vrows = XLSX.utils.sheet_to_json(vsheet, { header: 1 }).slice(1);

      const brojaci = {};
      brows.forEach(r => { if (r[0]) brojaci[r[0]] = { last: r[1] || 0, lastIR: r[2] || 0, lastWS: r[3] || 0 }; });

      const lokacii = {};
      lrows.forEach(r => { if (r[0]) lokacii[r[0]] = r[1]; });

      const vnes = vrows.filter(r => r[0]).map(r => ({
        date: r[0], loc: r[1], type: r[2], num: r[3], code: r[4], prefix: r[5]
      }));

      setExcelData({ brojaci, lokacii, vnes });
    };
    reader.readAsArrayBuffer(file);
  };

  const updateRSC = useCallback((idx, updated) => {
    setRscs(prev => prev.map((r, i) => i === idx ? updated : r));
  }, []);

  const generate = () => {
    const d = new Date(notifDate);
    const dateLabel = notifDate ? `${d.getDate()}.${d.getMonth() + 1}.${d.getFullYear()}` : "[DATE]";

    const rscTexts = rscs.map(r => ({ rscno: r.rscno, text: buildRSCContent(r) }));
    const notifText = buildNotifText(rscs);
    const emailText = buildEmail(rscs, dateLabel, recipient, emailType);
    setOutputs({ rscTexts, notifText, emailText });
    setTab("output");
  };

  const exportUpdatedExcel = () => {
    if (!rawWorkbook || !excelData) return;
    const wb = XLSX.utils.book_new();

    // Update BROJACI
    const newBro = { ...excelData.brojaci };
    rscs.forEach(r => {
      if (!r._autoSeq || !r._autoType || r.inscode === "OTHER") return;
      const type = r._autoType || r.inscode;
      const prefix = r._autoPrefix || r.detcode;
      if (!newBro[type]) newBro[type] = { last: 0, lastIR: 0, lastWS: 0 };
      newBro[type].last = Math.max(newBro[type].last, r._autoSeq);
      if (prefix === "IR") newBro[type].lastIR = Math.max(newBro[type].lastIR, r._autoSeq);
      else newBro[type].lastWS = Math.max(newBro[type].lastWS, r._autoSeq);
    });

    const broData = [["Тип", "Последен Број", "Последен Број IR", "Последен Број WS"]];
    Object.entries(newBro).forEach(([k, v]) => broData.push([k, v.last, v.lastIR, v.lastWS]));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(broData), "BROJACI");

    // Update VNES
    const today = new Date();
    const todayStr = `${String(today.getDate()).padStart(2,"0")}.${String(today.getMonth()+1).padStart(2,"0")}.${today.getFullYear()}`;
    const vnesData = [["Дата", "Локација (KM)", "Тип", "Број.", "RSC код", "Префикс"]];
    excelData.vnes.forEach(r => vnesData.push([r.date, r.loc, r.type, r.num, r.code, r.prefix]));
    rscs.forEach(r => {
      if (!r.rscno) return;
      const loc = r.startch && r.startch !== "manual" ? r.startch : r.startch;
      const type = r.inscode === "OTHER" ? (r.inscodeother || "OTHER") : r.inscode;
      const seq = r._autoSeq || parseInt(r.seqcode) || 0;
      const prefix = r._autoPrefix || r.detcode;
      const code = `${loc.replace("KM ", "")}-${prefix}-${type}-${padSeq(seq)}`;
      vnesData.push([todayStr, loc, type, seq, code, prefix]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(vnesData), "VNES");

    // Keep LOKACII
    const lokData = [["Локација (KM)", "Префикс"]];
    Object.entries(excelData.lokacii).forEach(([k, v]) => lokData.push([k, v]));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(lokData), "LOKACII");

    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = excelFileName || "RSC_Generator_updated.xlsx";
    a.click();
    URL.revokeObjectURL(url);
  };

  const btnTab = (active) => ({
    padding: "7px 16px", borderRadius: 7, border: active ? "none" : "1px solid #ddd",
    background: active ? "#1a1a1a" : "transparent", color: active ? "#fff" : "#555",
    cursor: "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: active ? 500 : 400,
  });

  const primaryBtn = { background: "#1a1a1a", color: "#fff", border: "none", padding: "9px 24px", borderRadius: 8, fontSize: 13, cursor: "pointer", fontFamily: "inherit", fontWeight: 500 };
  const secBtn = { background: "transparent", color: "#1a1a1a", border: "1px solid #ccc", padding: "9px 20px", borderRadius: 8, fontSize: 13, cursor: "pointer", fontFamily: "inherit" };

  return (
    <div style={{ fontFamily: "system-ui, sans-serif", maxWidth: 720, margin: "0 auto", padding: "16px 12px" }}>
      <h2 style={{ fontSize: 17, fontWeight: 600, marginBottom: 4, color: "#1a1a1a" }}>RSC Document Generator</h2>
      <p style={{ fontSize: 12, color: "#888", marginBottom: 14 }}>AK Invest — North Macedonia Corridor 8 &amp; 10-D</p>

      {/* Excel upload banner */}
      <div style={{ border: excelData ? "1px solid #b7ddb7" : "1.5px dashed #ccc", borderRadius: 10, padding: "10px 14px", marginBottom: 14, background: excelData ? "#f0faf0" : "#fafafa", display: "flex", alignItems: "center", gap: 12 }}>
        {excelData ? (
          <>
            <span style={{ fontSize: 20 }}>📊</span>
            <div style={{ flex: 1 }}>
              <div style={{ fontSize: 13, fontWeight: 500, color: "#2e7d32" }}>✓ {excelFileName} вчитан</div>
              <div style={{ fontSize: 11, color: "#555" }}>
                {Object.keys(excelData.brojaci).length} типови · {Object.keys(excelData.lokacii).length} локации · {excelData.vnes.length} записи во историја
              </div>
            </div>
            <button onClick={() => fileRef.current.click()} style={{ ...secBtn, padding: "5px 12px", fontSize: 12 }}>Смени</button>
          </>
        ) : (
          <>
            <span style={{ fontSize: 20 }}>📎</span>
            <div style={{ flex: 1 }}>
              <div style={{ fontSize: 13, fontWeight: 500 }}>Прикачи RSC_Generator Excel</div>
              <div style={{ fontSize: 11, color: "#888" }}>За автоматски сериски броеви и историја</div>
            </div>
            <button onClick={() => fileRef.current.click()} style={{ ...primaryBtn, padding: "7px 16px", fontSize: 12 }}>Прикачи .xlsx</button>
          </>
        )}
        <input ref={fileRef} type="file" accept=".xlsx" style={{ display: "none" }} onChange={handleExcelUpload} />
      </div>

      <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
        <button style={btnTab(tab === "input")} onClick={() => setTab("input")}>Внес</button>
        <button style={btnTab(tab === "output")} onClick={() => setTab("output")}>Излез</button>
        {excelData && <button style={btnTab(tab === "history")} onClick={() => setTab("history")}>Историја</button>}
      </div>

      {tab === "input" && (
        <div>
          {rscs.map((r, i) => (
            <RSCForm key={r.id} rsc={r} idx={i} onChange={(u) => updateRSC(i, u)} onDelete={() => setRscs(p => p.filter((_, j) => j !== i))} excelData={excelData} />
          ))}
          <button onClick={() => setRscs(p => [...p, emptyRSC()])} style={{ width: "100%", padding: 9, border: "1.5px dashed #ccc", borderRadius: 8, background: "transparent", cursor: "pointer", color: "#888", fontSize: 13, fontFamily: "inherit", marginBottom: 14 }}>
            + Додај уште еден RSC
          </button>

          <div style={{ border: "1px solid #e0e0e0", borderRadius: 10, padding: "13px 15px", background: "#fafafa", marginBottom: 14 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: "#555", marginBottom: 9 }}>Поставки за email &amp; нотификација</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
              <div>
                <label style={lbl}>Датум на нотификација</label>
                <input type="date" value={notifDate} onChange={e => setNotifDate(e.target.value)} style={{ ...inp }} />
              </div>
              <div>
                <label style={lbl}>Email примач</label>
                <input value={recipient} onChange={e => setRecipient(e.target.value)} style={inp} placeholder="Emil" />
              </div>
              <div>
                <label style={lbl}>Тип на email</label>
                <select value={emailType} onChange={e => setEmailType(e.target.value)} style={inp}>
                  <option value="submittal">Submittal</option>
                  <option value="resubmittal">Resubmittal</option>
                </select>
              </div>
            </div>
          </div>

          <div style={{ display: "flex", gap: 10 }}>
            <button onClick={generate} style={primaryBtn}>Генерирај ↗</button>
            {excelData && outputs && (
              <button onClick={exportUpdatedExcel} style={secBtn}>⬇ Превземи ажуриран Excel</button>
            )}
          </div>
        </div>
      )}

      {tab === "output" && outputs && (
        <div>
          {outputs.rscTexts.map((r, i) => <OutBox key={i} label={`📄 RSC ${i + 1}: ${r.rscno || "(без број)"}`} text={r.text} />)}
          <OutBox label="📋 Daily Notification Invitation List" text={outputs.notifText} />
          <OutBox label="✉️ Email текст" text={outputs.emailText} />
          {excelData && (
            <div style={{ marginTop: 10 }}>
              <button onClick={exportUpdatedExcel} style={primaryBtn}>⬇ Превземи ажуриран Excel (со новите записи)</button>
              <div style={{ fontSize: 11, color: "#888", marginTop: 5 }}>Ќе ги ажурира BROJACI и VNES табелите</div>
            </div>
          )}
        </div>
      )}

      {tab === "output" && !outputs && (
        <div style={{ color: "#888", fontSize: 13, padding: "20px 0" }}>Нема излез уште — пополни го формуларот и кликни Генерирај.</div>
      )}

      {tab === "history" && excelData && (
        <div>
          <div style={{ fontSize: 13, fontWeight: 500, marginBottom: 10 }}>Последни 20 RSC записи</div>
          <HistoryTable vnes={excelData.vnes} />
          <div style={{ marginTop: 14 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: "#555", marginBottom: 8 }}>Тековни бројачи</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 8 }}>
              {Object.entries(excelData.brojaci).map(([type, v]) => (
                <div key={type} style={{ border: "1px solid #e0e0e0", borderRadius: 8, padding: "8px 10px", background: "#fafafa" }}>
                  <div style={{ fontSize: 11, color: "#888" }}>Тип</div>
                  <div style={{ fontWeight: 600, fontSize: 14 }}>{type}</div>
                  <div style={{ fontSize: 11, marginTop: 4 }}>
                    <span style={{ color: "#555" }}>IR: </span><strong>{v.lastIR}</strong>
                    <span style={{ color: "#555", marginLeft: 8 }}>WS: </span><strong>{v.lastWS}</strong>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
