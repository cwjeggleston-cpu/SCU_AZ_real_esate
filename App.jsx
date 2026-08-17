import React, { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

/* ============================================================
   SCU PHX METRO SPACE PLANNER — v2
   Southern California University of Health Sciences
   Runs as a Claude artifact AND as a deployed Vite/Netlify app.
   ============================================================ */

const LOGO_WHITE = "data:image/png;base64,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";

const BRAND = {
  blue: "#004987",
  navy: "#00335F",
  ink: "#1C2B3A",
  slate: "#5A6B7C",
  mist: "#EEF3F8",
  line: "#D5DFE9",
  gold: "#F2A900",
  red: "#C43D3D",
  green: "#2E8B57",
};

const PROGRAM_COLORS = {
  dc: "#004987",
  dpt: "#1B8A8F",
  otd: "#7C6DAF",
  slp: "#C86B3C",
  pa: "#F2A900",
  gc: "#6BA84F",
  anat: "#8A93A0",
  shared: "#2F6DA4",
};

const ARC_PA_DEADLINE = new Date(2027, 9, 1); // October 1, 2027
const STORAGE_KEY = "scu-phx-planner-v3";

/* ------------- Storage adapter: artifact OR deployed site ------------- */
const store = {
  async get(key) {
    if (typeof window !== "undefined" && window.storage && window.storage.get) {
      try { return await window.storage.get(key); } catch (e) { return null; }
    }
    try {
      const v = window.localStorage.getItem(key);
      return v ? { key, value: v } : null;
    } catch (e) { return null; }
  },
  async set(key, value) {
    if (typeof window !== "undefined" && window.storage && window.storage.set) {
      try { return await window.storage.set(key, value); } catch (e) { return null; }
    }
    try { window.localStorage.setItem(key, value); return { key, value }; } catch (e) { return null; }
  },
  async delete(key) {
    if (typeof window !== "undefined" && window.storage && window.storage.delete) {
      try { return await window.storage.delete(key); } catch (e) { return null; }
    }
    try { window.localStorage.removeItem(key); return { key, deleted: true }; } catch (e) { return null; }
  },
};

/* ------------------ Default data model ------------------ */

const defaultData = () => ({
  multiplier: 1.45,
  overlapSF: 5500,        // DPT cohort-overlap lab need
  surgeRentAnnual: 0,     // est. $/yr when renting surge space
  programs: [
    {
      id: "dc", code: "DC", name: "Doctor of Chiropractic", usesMultiplier: true, sharePct: 0,
      rooms: [
        { name: "ADJ Lab 1", nsf: 1800 },
        { name: "ADJ Lab 2", nsf: 1800 },
        { name: "Diag. Lab 1", nsf: 1800 },
        { name: "Diag. Lab 2", nsf: 1800 },
        { name: "X-Ray", nsf: 600 },
        { name: "Substitute for Room 210", nsf: 1800 },
      ],
    },
    { id: "dpt", code: "DPT", name: "Doctor of Physical Therapy", usesMultiplier: true, sharePct: 100,
      rooms: [{ name: "Main Lab / Instructional", nsf: 5500 }] },
    { id: "otd", code: "OTD", name: "Doctor of Occupational Therapy", usesMultiplier: true, sharePct: 100,
      rooms: [{ name: "Main Lab / Instructional", nsf: 5500 }] },
    { id: "slp", code: "SLP", name: "Speech-Language Pathology", usesMultiplier: true, sharePct: 100,
      rooms: [{ name: "Main Lab / Instructional", nsf: 5500 }] },
    { id: "pa", code: "PA", name: "Physician Assistant (ARC-PA)", usesMultiplier: true, sharePct: 0,
      rooms: [{ name: "Instructional Suite", nsf: 9000 }] },
    { id: "gc", code: "GC", name: "Genetic Counseling", usesMultiplier: true, sharePct: 100,
      rooms: [{ name: "Instructional Suite", nsf: 2000 }] },
    { id: "anat", code: "ANAT", name: "Anatomy Lab (est. GSF)", usesMultiplier: false, sharePct: 0,
      rooms: [{ name: "Anatomy Lab", nsf: 2000 }] },
  ],
  sites: [
    { id: "sonoran", name: "Sonoran — Tempe", capacity: 10062, rate: 28.71, ti: 0,
      note: "Current lease: 10,062 RSF @ $28.71/RSF/yr. Raise capacity to model the expanded ask." },
    { id: "scottsdale", name: "Scottsdale — Candidate", capacity: 50000, rate: 45, ti: 0,
      note: "≈50,000 SF candidate site @ $45/SF/yr (whole-floor option)." },
    { id: "offsite", name: "Other Off-Site (TBD)", capacity: 0, rate: 0, ti: 0,
      note: "Placeholder for any additional location. Capacity 0 = no cap defined." },
  ],
  scenarios: [
    {
      id: "s1", name: "1 · All at Sonoran",
      note: "Every program at Sonoran; DPT overlap handled via rented surge space.",
      sharing: false, includeAnatomy: true,
      overlap: { mode: "surge", siteId: "sonoran" },
      assign: { dc: "sonoran", dpt: "sonoran", otd: "sonoran", slp: "sonoran", pa: "sonoran", gc: "sonoran", anat: "sonoran" },
    },
    {
      id: "s2", name: "2 · SPLIT Campus: PA/GC separate",
      note: "PA & GC at Scottsdale; everyone else at Sonoran. DPT overlap via surge rental.",
      sharing: false, includeAnatomy: true,
      overlap: { mode: "surge", siteId: "sonoran" },
      assign: { dc: "sonoran", dpt: "sonoran", otd: "sonoran", slp: "sonoran", pa: "scottsdale", gc: "scottsdale", anat: "sonoran" },
    },
    {
      id: "s3", name: "3 · SPLIT Campus: DC/PA off campus",
      note: "DC & PA at Scottsdale; everyone else at Sonoran. DPT overlap via surge rental.",
      sharing: false, includeAnatomy: true,
      overlap: { mode: "surge", siteId: "sonoran" },
      assign: { dc: "scottsdale", dpt: "sonoran", otd: "sonoran", slp: "sonoran", pa: "scottsdale", gc: "sonoran", anat: "sonoran" },
    },
    {
      id: "s4", name: "4 · All at Scottsdale",
      note: "Everyone at Scottsdale — no Sonoran, no surge. DPT overlap absorbed on-site (whole floor).",
      sharing: false, includeAnatomy: true,
      overlap: { mode: "absorb", siteId: "scottsdale" },
      assign: { dc: "scottsdale", dpt: "scottsdale", otd: "scottsdale", slp: "scottsdale", pa: "scottsdale", gc: "scottsdale", anat: "scottsdale" },
    },
  ],
  activeScenarioId: "s1",
});

/* ------------------ Helpers ------------------ */

const fmt = (n) => Math.round(n).toLocaleString("en-US");
const fmtMoney = (n) => "$" + Math.round(n).toLocaleString("en-US");

function programNSF(p) {
  return p.rooms.reduce((s, r) => s + (Number(r.nsf) || 0), 0);
}
function programGSF(p, multiplier) {
  const nsf = programNSF(p);
  return p.usesMultiplier ? nsf * multiplier : nsf;
}

/* Core scenario math.
   Shared-core model: for programs with sharePct > 0 at the same (non-surge)
   site when sharing is on, the shareable slice of base NSF is counted ONCE
   (largest need wins). The multiplier remainder — offices, storage,
   circulation — stays dedicated to each program. */
function computeScenario(scn, data) {
  const mult = data.multiplier;
  const overhead = Math.max(0, mult - 1);
  const overlapSF = Number(data.overlapSF) || 0;
  const overlap = scn.overlap || { mode: "unresolved", siteId: "" };

  const bySite = {};
  data.sites.forEach((s) => (bySite[s.id] = { site: s, blocks: [], used: 0 }));
  bySite.unassigned = { site: { id: "unassigned", name: "Unassigned", capacity: 0 }, blocks: [], used: 0 };

  const included = data.programs.filter((p) => p.id !== "anat" || scn.includeAnatomy);

  const groups = {};
  included.forEach((p) => {
    const sid = bySite[scn.assign[p.id]] ? scn.assign[p.id] : "unassigned";
    (groups[sid] = groups[sid] || []).push(p);
  });

  let shareSavings = 0;

  Object.entries(groups).forEach(([sid, progs]) => {
    const bucket = bySite[sid];

    const sharers = scn.sharing ? progs.filter((p) => (p.sharePct || 0) > 0 && p.id !== "anat") : [];
    const nonSharers = progs.filter((p) => !sharers.includes(p));

    let pool = 0, sumShared = 0;
    sharers.forEach((p) => {
      const N = programNSF(p);
      const S = N * ((p.sharePct || 0) / 100);
      const D = N - S;
      pool = Math.max(pool, S);
      sumShared += S;
      bucket.blocks.push({ program: p, gsf: S * overhead + D * mult, kind: "dedicated" });
    });
    if (sharers.length > 0) {
      bucket.blocks.push({
        program: { id: "shared", code: "SHARED", name: "Shared core — " + sharers.map((p) => p.code).join(" · ") },
        gsf: pool, kind: "pool", members: sharers.map((p) => p.code),
      });
      shareSavings += sumShared - pool;
    }
    nonSharers.forEach((p) => bucket.blocks.push({ program: p, gsf: programGSF(p, mult), kind: "full" }));
    bucket.used = bucket.blocks.reduce((s, b) => s + b.gsf, 0);
  });

  // ---- DPT cohort-overlap resolution ----
  // "absorb": lands as a flat block in a permanent site (rides the host allocation's overhead).
  // "surge": rented off-site 1wk × 3/yr — tracked separately from permanent GSF.
  let surgeSF = 0;
  if (overlap.mode === "absorb" && overlapSF > 0) {
    const host = bySite[overlap.siteId] || bySite.unassigned;
    host.blocks.push({
      program: { id: "overlap", code: "DPT+", name: "DPT cohort overlap (2nd simultaneous lab)" },
      gsf: overlapSF, kind: "overlap",
    });
    host.used += overlapSF;
  } else if (overlap.mode === "surge") {
    surgeSF = overlapSF;
  }
  const overlapUnresolved = overlap.mode !== "absorb" && overlap.mode !== "surge" && overlapSF > 0;

  const total = Object.values(bySite).reduce((s, b) => s + b.used, 0);
  const overs = Object.values(bySite).filter((b) => b.site.capacity > 0 && b.used > b.site.capacity);

  const paSiteId = scn.assign.pa;
  const paBucket = paSiteId ? bySite[paSiteId] : null;
  const paAtRisk =
    !paSiteId || paSiteId === "unassigned" ||
    (paBucket && paBucket.site.capacity > 0 && paBucket.used > paBucket.site.capacity);

  let annualCost = 0, tiCost = 0;
  Object.values(bySite).forEach((b) => {
    if (b.site.rate > 0) annualCost += b.used * b.site.rate;
    if (b.site.ti > 0) tiCost += b.used * b.site.ti;
  });
  if (overlap.mode === "surge") annualCost += Number(data.surgeRentAnnual) || 0;

  return { bySite, total, overs, paAtRisk, shareSavings, annualCost, tiCost, overlap, overlapSF, surgeSF, overlapUnresolved };
}


/* ------------------ Excel export ------------------ */
function exportExcel(data) {
  const wb = XLSX.utils.book_new();
  const siteRows = [...data.sites, { id: "unassigned", name: "Unassigned" }];
  const results = data.scenarios.map((s) => computeScenario(s, data));

  // Sheet 1: scenario comparison
  const cmp = [];
  cmp.push(["SCU PHX Metro Space Scenarios", "", "", "", ""]);
  cmp.push(["Exported", new Date().toLocaleDateString("en-US")]);
  cmp.push([]);
  cmp.push(["Metric", ...data.scenarios.map((s) => s.name)]);
  siteRows.forEach((site) => {
    cmp.push([site.name + " (GSF)", ...results.map((r) => { const b = r.bySite[site.id]; return b && b.used > 0 ? Math.round(b.used) : 0; })]);
  });
  cmp.push(["Surge rental (SF)", ...results.map((r) => Math.round(r.surgeSF || 0))]);
  cmp.push(["Total permanent GSF", ...results.map((r) => Math.round(r.total))]);
  cmp.push(["DPT overlap resolution", ...results.map((r) => {
    const m = (r.overlap || {}).mode;
    if (m === "surge") return "Rented surge space";
    if (m === "absorb") return "Absorbed at " + (((data.sites.find((x) => x.id === r.overlap.siteId) || {}).name) || "?");
    return "UNRESOLVED";
  })]);
  cmp.push(["Shared-core savings (GSF)", ...results.map((r) => Math.round(r.shareSavings || 0))]);
  cmp.push(["Anatomy Lab", ...data.scenarios.map((s) => (s.includeAnatomy ? "Included" : "Excluded"))]);
  cmp.push(["PA / ARC-PA status", ...results.map((r) => (r.paAtRisk ? "AT RISK" : "Housed"))]);
  cmp.push(["Est. annual lease ($)", ...results.map((r) => Math.round(r.annualCost || 0))]);
  cmp.push(["Est. one-time TI ($)", ...results.map((r) => Math.round(r.tiCost || 0))]);
  cmp.push([]);
  cmp.push(["GSF multiplier", data.multiplier]);
  cmp.push(["DPT cohort overlap (SF)", data.overlapSF]);
  const ws1 = XLSX.utils.aoa_to_sheet(cmp);
  ws1["!cols"] = [{ wch: 30 }, ...data.scenarios.map(() => ({ wch: 26 }))];
  XLSX.utils.book_append_sheet(wb, ws1, "Scenario Compare");

  // Sheet 2: program space model
  const pm = [["Program", "Room", "NSF", "Multiplier", "GSF"]];
  data.programs.forEach((p) => {
    p.rooms.forEach((r, i) => {
      pm.push([i === 0 ? p.code + " — " + p.name : "", r.name, Number(r.nsf) || 0, "", ""]);
    });
    const nsf = programNSF(p);
    pm.push(["", p.code + " subtotal", nsf, p.usesMultiplier ? data.multiplier : "flat", Math.round(programGSF(p, data.multiplier))]);
    pm.push([]);
  });
  pm.push(["Note: GSF multiplier covers restrooms, hallways, storage, and faculty offices."]);
  const ws2 = XLSX.utils.aoa_to_sheet(pm);
  ws2["!cols"] = [{ wch: 38 }, { wch: 28 }, { wch: 10 }, { wch: 10 }, { wch: 10 }];
  XLSX.utils.book_append_sheet(wb, ws2, "Program Space Model");

  // Sheet 3: sites & costs
  const st = [["Site", "Capacity (GSF)", "Lease rate ($/SF/yr)", "TI ($/SF one-time)", "Note"]];
  data.sites.forEach((sx) => st.push([sx.name, sx.capacity || 0, sx.rate || 0, sx.ti || 0, sx.note || ""]));
  st.push([]);
  st.push(["Surge rental est. ($/yr)", data.surgeRentAnnual || 0]);
  const ws3 = XLSX.utils.aoa_to_sheet(st);
  ws3["!cols"] = [{ wch: 28 }, { wch: 14 }, { wch: 18 }, { wch: 16 }, { wch: 60 }];
  XLSX.utils.book_append_sheet(wb, ws3, "Sites & Costs");

  const d = new Date();
  const stamp = d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
  XLSX.writeFile(wb, "SCU_PHX_Space_Scenarios_" + stamp + ".xlsx");
}

/* ------------------ Small UI atoms ------------------ */

function NumInput({ value, onChange, width = 92, step = 100, min = 0, max }) {
  return (
    <input
      type="number"
      value={value}
      step={step}
      min={min}
      max={max}
      onChange={(e) => onChange(e.target.value === "" ? 0 : Number(e.target.value))}
      style={{
        width, padding: "6px 8px", border: `1px solid ${BRAND.line}`,
        borderRadius: 6, fontFamily: "'Public Sans', sans-serif", fontSize: 13,
        color: BRAND.ink, background: "#fff", fontVariantNumeric: "tabular-nums",
      }}
    />
  );
}

function Tag({ children, color = BRAND.slate, bg = BRAND.mist }) {
  return (
    <span style={{
      fontSize: 10.5, fontWeight: 700, letterSpacing: "0.06em", textTransform: "uppercase",
      color, background: bg, padding: "3px 8px", borderRadius: 4, whiteSpace: "nowrap",
    }}>{children}</span>
  );
}

function Toggle({ checked, onChange, label, children }) {
  return (
    <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, cursor: "pointer" }}>
      <span
        role="switch"
        aria-checked={checked}
        tabIndex={0}
        onClick={() => onChange(!checked)}
        onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); onChange(!checked); } }}
        style={{
          width: 36, height: 20, borderRadius: 999, flexShrink: 0, position: "relative",
          background: checked ? BRAND.blue : BRAND.line, transition: "background 0.15s", cursor: "pointer",
        }}
      >
        <span style={{
          position: "absolute", top: 2, left: checked ? 18 : 2, width: 16, height: 16,
          borderRadius: "50%", background: "#fff", transition: "left 0.15s",
          boxShadow: "0 1px 2px rgba(0,0,0,0.25)",
        }} />
      </span>
      <span>{label}</span>
      {children}
    </label>
  );
}

function Info({ text }) {
  const [open, setOpen] = useState(false);
  return (
    <span style={{ position: "relative", display: "inline-flex" }}>
      <button
        onClick={(e) => { e.preventDefault(); e.stopPropagation(); setOpen((o) => !o); }}
        aria-label="More info"
        title={text}
        style={{
          width: 16, height: 16, borderRadius: "50%", border: `1.5px solid ${BRAND.blue}`,
          background: open ? BRAND.blue : "#fff", color: open ? "#fff" : BRAND.blue,
          fontSize: 10, fontWeight: 800, lineHeight: 1, cursor: "pointer",
          display: "inline-flex", alignItems: "center", justifyContent: "center", padding: 0,
          fontFamily: "'Public Sans', sans-serif",
        }}
      >?</button>
      {open && (
        <span
          onClick={(e) => { e.preventDefault(); e.stopPropagation(); setOpen(false); }}
          style={{
            position: "absolute", bottom: "calc(100% + 8px)", left: "50%", transform: "translateX(-50%)",
            width: 250, background: BRAND.ink, color: "#fff", fontSize: 11.5, lineHeight: 1.5,
            fontWeight: 400, padding: "10px 12px", borderRadius: 8, zIndex: 50,
            boxShadow: "0 4px 14px rgba(0,20,40,0.35)", cursor: "pointer", textTransform: "none",
            letterSpacing: "normal",
          }}
        >{text}</span>
      )}
    </span>
  );
}

const SURGE_INFO = "Driven by DPT running two cohorts' labs at the same time — one week, three times per year. If SCU stays at Sonoran, this 5,500 SF must be rented (e.g., a conference center). If a permanent site like Scottsdale has enough square footage, it can be absorbed there instead and no rental is needed.";

/* ------------------ Stacking diagram ------------------ */

function StackDiagram({ result, sites }) {
  const maxRef = Math.max(
    ...Object.values(result.bySite).map((b) => Math.max(b.used, b.site.capacity || 0)),
    result.surgeSF || 0,
    1
  );
  const H = 250;
  const order = [...sites.map((s) => s.id), "unassigned"];
  const hatch = (c1, c2) => `repeating-linear-gradient(45deg, ${c1}, ${c1} 6px, ${c2} 6px, ${c2} 12px)`;

  return (
    <div style={{ display: "flex", gap: 18, alignItems: "flex-end", flexWrap: "wrap" }}>
      {order.map((sid) => {
        const b = result.bySite[sid];
        if (!b) return null;
        if (sid === "unassigned" && b.blocks.length === 0) return null;
        const cap = b.site.capacity || 0;
        const over = cap > 0 && b.used > cap;
        const capY = cap > 0 ? (cap / maxRef) * H : null;
        const pct = cap > 0 ? Math.round((b.used / cap) * 100) : null;
        return (
          <div key={sid} style={{ width: 150, flexShrink: 0 }}>
            <div style={{ position: "relative", height: H, borderBottom: `3px solid ${BRAND.ink}` }}>
              {capY !== null && (
                <div style={{
                  position: "absolute", left: -6, right: -6, bottom: capY,
                  borderTop: `2px dashed ${over ? BRAND.red : BRAND.slate}`, zIndex: 2,
                }}>
                  <span style={{
                    position: "absolute", right: 0, top: -16, fontSize: 10, fontWeight: 700,
                    color: over ? BRAND.red : BRAND.slate, fontVariantNumeric: "tabular-nums",
                  }}>cap {fmt(cap)}</span>
                </div>
              )}
              <div style={{
                position: "absolute", left: 14, right: 14, bottom: 0,
                display: "flex", flexDirection: "column-reverse",
              }}>
                {b.blocks.filter((x) => x.gsf > 0).map((x, i) => {
                  const h = Math.max((x.gsf / maxRef) * H, 4);
                  const isPool = x.kind === "pool";
                  const isOverlap = x.kind === "overlap";
                  return (
                    <div
                      key={i}
                      title={isPool ? `${x.program.name}: ${fmt(x.gsf)} SF counted once`
                        : isOverlap ? `${x.program.name}: +${fmt(x.gsf)} SF absorbed here`
                        : `${x.program.code}: ${fmt(x.gsf)} GSF`}
                      style={{
                        height: h,
                        background: isPool ? hatch(BRAND.blue, PROGRAM_COLORS.shared)
                          : isOverlap ? hatch(PROGRAM_COLORS.dpt, "#3FA9AE")
                          : PROGRAM_COLORS[x.program.id],
                        borderTop: "2px solid rgba(255,255,255,0.85)",
                        display: "flex", alignItems: "center", justifyContent: "center",
                        color: "#fff", fontSize: 10.5, fontWeight: 700,
                        letterSpacing: "0.04em", overflow: "hidden", textAlign: "center",
                      }}
                    >
                      {h >= 15 ? (isPool ? `SHARED ${fmt(x.gsf)}` : isOverlap ? `DPT+ ${fmt(x.gsf)}` : `${x.program.code} ${fmt(x.gsf)}`) : ""}
                    </div>
                  );
                })}
              </div>
            </div>
            <div style={{ marginTop: 8 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: BRAND.ink, fontFamily: "'Montserrat', sans-serif" }}>
                {b.site.name}
              </div>
              <div style={{ fontSize: 12, color: over ? BRAND.red : BRAND.slate, fontWeight: over ? 700 : 500, fontVariantNumeric: "tabular-nums" }}>
                {fmt(b.used)} GSF{pct !== null ? ` · ${pct}% of cap` : ""}
                {over ? ` · over by ${fmt(b.used - cap)}` : ""}
              </div>
            </div>
          </div>
        );
      })}

      {/* Surge rental — off-books temporary space, drawn apart from permanent sites */}
      {result.overlap && result.overlap.mode === "surge" && result.surgeSF > 0 && (
        <div style={{ width: 150, flexShrink: 0, marginLeft: 8 }}>
          <div style={{ position: "relative", height: H, borderBottom: `3px dashed ${BRAND.slate}` }}>
            <div style={{ position: "absolute", left: 14, right: 14, bottom: 0 }}>
              <div
                title={`Rented surge space: ${fmt(result.surgeSF)} SF, one week × three times per year`}
                style={{
                  height: Math.max((result.surgeSF / maxRef) * H, 22),
                  background: hatch(PROGRAM_COLORS.dpt, "#3FA9AE"),
                  border: `2px dashed #fff`, boxSizing: "border-box",
                  display: "flex", alignItems: "center", justifyContent: "center",
                  color: "#fff", fontSize: 10.5, fontWeight: 700, letterSpacing: "0.04em", textAlign: "center",
                }}
              >DPT+ {fmt(result.surgeSF)}</div>
            </div>
          </div>
          <div style={{ marginTop: 8 }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: BRAND.ink, fontFamily: "'Montserrat', sans-serif", display: "flex", alignItems: "center", gap: 5 }}>
              Surge Rental <Info text={SURGE_INFO} />
            </div>
            <div style={{ fontSize: 12, color: BRAND.slate, fontVariantNumeric: "tabular-nums" }}>{fmt(result.surgeSF)} SF · rented</div>
            <div style={{ fontSize: 10, color: BRAND.slate, fontStyle: "italic" }}>conference center · 1 wk × 3/yr</div>
          </div>
        </div>
      )}

      {/* Unresolved overlap — floating block that still needs a landing spot */}
      {result.overlapUnresolved && (
        <div style={{ width: 150, flexShrink: 0, marginLeft: 8, alignSelf: "center" }}>
          <div
            title={SURGE_INFO}
            style={{
              height: Math.max((result.overlapSF / maxRef) * H, 34),
              background: hatch("#E3B341", BRAND.gold),
              border: `2px dashed ${BRAND.ink}`, borderRadius: 6,
              display: "flex", alignItems: "center", justifyContent: "center",
              color: BRAND.ink, fontSize: 10.5, fontWeight: 800, letterSpacing: "0.04em",
              textAlign: "center", padding: 4, boxSizing: "border-box",
              transform: "rotate(-2deg)",
            }}
          >DPT+ {fmt(result.overlapSF)} SF<br />no landing spot</div>
          <div style={{ marginTop: 8, fontSize: 11, color: BRAND.gold, fontWeight: 700 }}>⚠ Unresolved overlap</div>
        </div>
      )}
    </div>
  );
}

/* ------------------ Main App ------------------ */

export default function App() {
  const [data, setData] = useState(null);
  const [tab, setTab] = useState("planner");
  const [saveState, setSaveState] = useState("idle");
  const saveTimer = useRef(null);
  const firstLoad = useRef(true);

  useEffect(() => {
    (async () => {
      try {
        const res = await store.get(STORAGE_KEY);
        if (res && res.value) {
          const d = defaultData();
          setData({ ...d, ...JSON.parse(res.value) });
          return;
        }
      } catch (e) { /* fresh start */ }
      setData(defaultData());
    })();
  }, []);

  useEffect(() => {
    if (!data) return;
    if (firstLoad.current) { firstLoad.current = false; return; }
    setSaveState("saving");
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      const ok = await store.set(STORAGE_KEY, JSON.stringify(data));
      setSaveState(ok ? "saved" : "error");
      if (ok) setTimeout(() => setSaveState("idle"), 1600);
    }, 700);
    return () => clearTimeout(saveTimer.current);
  }, [data]);

  const daysLeft = Math.max(0, Math.ceil((ARC_PA_DEADLINE - new Date()) / 86400000));

  if (!data) {
    return (
      <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: BRAND.mist, fontFamily: "'Public Sans', sans-serif", color: BRAND.slate }}>
        Loading planner…
      </div>
    );
  }

  const scn = data.scenarios.find((s) => s.id === data.activeScenarioId) || data.scenarios[0];
  const result = computeScenario(scn, data);

  const update = (fn) => setData((d) => { const nd = JSON.parse(JSON.stringify(d)); fn(nd); return nd; });
  const updateScn = (fn) => update((d) => fn(d.scenarios.find((x) => x.id === d.activeScenarioId)));
  const setAssign = (pid, sid) => updateScn((s) => (s.assign[pid] = sid));

  const duplicateScenario = () => update((d) => {
    const src = d.scenarios.find((x) => x.id === d.activeScenarioId);
    const copy = JSON.parse(JSON.stringify(src));
    copy.id = "c" + Date.now();
    copy.name = src.name.replace(/^\d+ · /, "") + " (copy)";
    d.scenarios.push(copy);
    d.activeScenarioId = copy.id;
  });

  const deleteScenario = () => update((d) => {
    if (d.scenarios.length <= 1) return;
    d.scenarios = d.scenarios.filter((x) => x.id !== d.activeScenarioId);
    d.activeScenarioId = d.scenarios[0].id;
  });

  const resetAll = async () => {
    await store.delete(STORAGE_KEY);
    firstLoad.current = true;
    setData(defaultData());
  };

  const card = {
    background: "#fff", border: `1px solid ${BRAND.line}`, borderRadius: 10,
    padding: 20, boxShadow: "0 1px 2px rgba(0,40,80,0.05)",
  };
  const h2 = {
    fontFamily: "'Montserrat', sans-serif", fontWeight: 800, fontSize: 13,
    letterSpacing: "0.12em", textTransform: "uppercase", color: BRAND.blue,
    margin: "0 0 14px 0",
  };
  const btn = (primary) => ({
    padding: "8px 14px", borderRadius: 6, border: primary ? "none" : `1px solid ${BRAND.line}`,
    background: primary ? BRAND.blue : "#fff", color: primary ? "#fff" : BRAND.ink,
    fontFamily: "'Montserrat', sans-serif", fontWeight: 700, fontSize: 12,
    cursor: "pointer", letterSpacing: "0.03em",
  });

  const tabs = [
    ["planner", "Scenario Planner"],
    ["compare", "Compare"],
    ["programs", "Program Space Model"],
    ["sites", "Sites & Costs"],
  ];

  return (
    <div style={{ minHeight: "100vh", background: BRAND.mist, fontFamily: "'Public Sans', sans-serif", color: BRAND.ink }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;700;800&family=Public+Sans:wght@400;500;600;700&display=swap');
        input:focus, select:focus, button:focus-visible, [role="switch"]:focus-visible { outline: 2px solid ${BRAND.gold}; outline-offset: 1px; }
        @media (prefers-reduced-motion: reduce) { * { transition: none !important; } }
      `}</style>

      {/* ---------- Header ---------- */}
      <header style={{ background: `linear-gradient(135deg, ${BRAND.navy} 0%, ${BRAND.blue} 70%)`, padding: "18px 28px", display: "flex", alignItems: "center", gap: 20, flexWrap: "wrap" }}>
        <img src={LOGO_WHITE} alt="Southern California University of Health Sciences" style={{ height: 40, maxWidth: "60vw", objectFit: "contain" }} />
        <div style={{ borderLeft: "1px solid rgba(255,255,255,0.35)", paddingLeft: 20 }}>
          <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 800, fontSize: 17, color: "#fff", letterSpacing: "0.02em" }}>
            PHX Metro Space Planner
          </div>
          <div style={{ fontSize: 11.5, color: "rgba(255,255,255,0.75)" }}>
            Live scenario modeling · Tempe & Scottsdale
          </div>
        </div>
        <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 14 }}>
          <div style={{ textAlign: "right" }}>
            <div style={{ fontSize: 10, fontWeight: 700, letterSpacing: "0.1em", color: BRAND.gold, textTransform: "uppercase" }}>
              ARC-PA Site Walk · Oct 2027
            </div>
            <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 800, fontSize: 20, color: "#fff", fontVariantNumeric: "tabular-nums" }}>
              {daysLeft.toLocaleString()} days
            </div>
          </div>
          <div style={{ fontSize: 10.5, color: "rgba(255,255,255,0.65)", minWidth: 52, textAlign: "right" }}>
            {saveState === "saving" ? "Saving…" : saveState === "saved" ? "Saved ✓" : saveState === "error" ? "Save failed" : ""}
          </div>
        </div>
      </header>

      {/* ---------- Tabs ---------- */}
      <nav style={{ background: "#fff", borderBottom: `1px solid ${BRAND.line}`, padding: "0 28px", display: "flex", gap: 6, overflowX: "auto" }}>
        {tabs.map(([id, label]) => (
          <button key={id} onClick={() => setTab(id)} style={{
            padding: "13px 16px", background: "none", border: "none", cursor: "pointer",
            fontFamily: "'Montserrat', sans-serif", fontWeight: 700, fontSize: 12.5,
            letterSpacing: "0.04em", whiteSpace: "nowrap",
            color: tab === id ? BRAND.blue : BRAND.slate,
            borderBottom: tab === id ? `3px solid ${BRAND.gold}` : "3px solid transparent",
          }}>{label}</button>
        ))}
        <button onClick={() => exportExcel(data)} style={{
          marginLeft: "auto", alignSelf: "center", background: BRAND.blue, border: "none", borderRadius: 6,
          color: "#fff", fontSize: 11.5, fontWeight: 700, cursor: "pointer", whiteSpace: "nowrap", padding: "7px 12px",
          fontFamily: "'Montserrat', sans-serif",
        }}>⬇ Export Excel</button>
        <button onClick={resetAll} style={{
          alignSelf: "center", background: "none", border: "none",
          color: BRAND.slate, fontSize: 11.5, cursor: "pointer", whiteSpace: "nowrap", padding: "8px 4px", marginLeft: 8,
        }}>Reset to defaults</button>
      </nav>

      <main style={{ padding: "24px 28px 60px", maxWidth: 1240, margin: "0 auto" }}>

        {/* ================= PLANNER ================= */}
        {tab === "planner" && (
          <>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 18, alignItems: "center" }}>
              {data.scenarios.map((s) => (
                <button key={s.id} onClick={() => update((d) => (d.activeScenarioId = s.id))} style={{
                  padding: "9px 16px", borderRadius: 999, cursor: "pointer",
                  border: s.id === scn.id ? `2px solid ${BRAND.blue}` : `1px solid ${BRAND.line}`,
                  background: s.id === scn.id ? BRAND.blue : "#fff",
                  color: s.id === scn.id ? "#fff" : BRAND.ink,
                  fontFamily: "'Montserrat', sans-serif", fontWeight: 700, fontSize: 12,
                }}>{s.name}</button>
              ))}
              <button onClick={duplicateScenario} style={btn(false)}>+ Duplicate</button>
              {data.scenarios.length > 1 && (
                <button onClick={deleteScenario} style={{ ...btn(false), color: BRAND.red }}>Delete</button>
              )}
            </div>

            {result.paAtRisk && (
              <div style={{ ...card, borderLeft: `5px solid ${BRAND.red}`, marginBottom: 14, padding: "14px 20px", display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                <Tag color="#fff" bg={BRAND.red}>ARC-PA Risk</Tag>
                <span style={{ fontSize: 13 }}>
                  PA is {!scn.assign.pa || scn.assign.pa === "unassigned" ? "not assigned to a site" : "assigned to a site that is over capacity"}. A walkable, ready PA space is required for the October 2027 site visit.
                </span>
              </div>
            )}
            {result.overs.map((b) => (
              <div key={b.site.id} style={{ ...card, borderLeft: `5px solid ${BRAND.red}`, marginBottom: 14, padding: "14px 20px", display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                <Tag color="#fff" bg={BRAND.red}>Over Capacity</Tag>
                <span style={{ fontSize: 13 }}>
                  <b>{b.site.name}</b> needs {fmt(b.used)} GSF against {fmt(b.site.capacity)} available — short <b>{fmt(b.used - b.site.capacity)} GSF</b>. This is the ask if pursuing this scenario.
                </span>
              </div>
            ))}
            {result.overlapUnresolved && (
              <div style={{ ...card, borderLeft: `5px solid ${BRAND.gold}`, marginBottom: 14, padding: "14px 20px", display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                <Tag color="#fff" bg={BRAND.gold}>Hanging {fmt(result.overlapSF)} SF</Tag>
                <span style={{ fontSize: 13 }}>
                  DPT's cohort overlap has no landing spot in this scenario. Choose a resolution:
                </span>
                <button onClick={() => updateScn((sc) => (sc.overlap = { mode: "surge", siteId: "sonoran" }))} style={btn(false)}>Rent surge space</button>
                <button onClick={() => updateScn((sc) => (sc.overlap = { mode: "absorb", siteId: "scottsdale" }))} style={btn(false)}>Absorb at Scottsdale</button>
              </div>
            )}

            <div style={{ display: "grid", gridTemplateColumns: "minmax(340px, 5fr) minmax(380px, 7fr)", gap: 18, alignItems: "start" }}>
              {/* assignment board */}
              <section style={card}>
                <h2 style={h2}>Program Assignments</h2>
                <input
                  value={scn.name}
                  onChange={(e) => updateScn((s) => (s.name = e.target.value))}
                  style={{ width: "100%", boxSizing: "border-box", fontFamily: "'Montserrat', sans-serif", fontWeight: 700, fontSize: 15, color: BRAND.ink, border: `1px solid ${BRAND.line}`, borderRadius: 6, padding: "8px 10px", marginBottom: 6 }}
                />
                <input
                  value={scn.note || ""}
                  placeholder="Scenario note…"
                  onChange={(e) => updateScn((s) => (s.note = e.target.value))}
                  style={{ width: "100%", boxSizing: "border-box", fontSize: 12, color: BRAND.slate, border: `1px solid ${BRAND.line}`, borderRadius: 6, padding: "7px 10px", marginBottom: 12 }}
                />

                {/* scenario toggles */}
                <div style={{ display: "flex", flexDirection: "column", gap: 10, padding: "12px 12px", background: BRAND.mist, borderRadius: 8, marginBottom: 12 }}>
                  {/* DPT cohort overlap resolution */}
                  <div style={{ padding: "10px 12px", background: "#fff", border: `1px solid ${BRAND.line}`, borderRadius: 8 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, fontWeight: 700, color: BRAND.ink, marginBottom: 8 }}>
                      DPT cohort overlap · <NumInput value={data.overlapSF} step={500} width={70} onChange={(v) => update((d) => (d.overlapSF = v))} /> SF <Info text={SURGE_INFO} />
                    </div>
                    {[
                      ["surge", "Rent surge space (conference center · 1 wk × 3/yr)"],
                      ["absorb", "Absorb into a permanent site"],
                      ["unresolved", "Unresolved — flag it"],
                    ].map(([mode, label]) => (
                      <label key={mode} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, padding: "4px 0", cursor: "pointer" }}>
                        <input
                          type="radio"
                          name={`overlap-${scn.id}`}
                          checked={(scn.overlap || {}).mode === mode || (!scn.overlap && mode === "unresolved")}
                          onChange={() => updateScn((sc) => (sc.overlap = { mode, siteId: (sc.overlap && sc.overlap.siteId) || "scottsdale" }))}
                        />
                        <span>{label}</span>
                        {mode === "surge" && (scn.overlap || {}).mode === "surge" && (
                          <span style={{ display: "inline-flex", alignItems: "center", gap: 4, marginLeft: 4, color: BRAND.slate, fontSize: 11.5 }}>
                            est. <NumInput value={data.surgeRentAnnual} step={1000} width={82} onChange={(v) => update((d) => (d.surgeRentAnnual = v))} /> $/yr
                          </span>
                        )}
                        {mode === "absorb" && (scn.overlap || {}).mode === "absorb" && (
                          <select
                            value={(scn.overlap || {}).siteId || "scottsdale"}
                            onChange={(e) => updateScn((sc) => (sc.overlap = { mode: "absorb", siteId: e.target.value }))}
                            style={{ padding: "4px 6px", borderRadius: 6, border: `1px solid ${BRAND.line}`, fontSize: 11.5, background: "#fff" }}
                          >
                            {data.sites.map((st) => <option key={st.id} value={st.id}>{st.name}</option>)}
                          </select>
                        )}
                      </label>
                    ))}
                  </div>
                  <Toggle
                    checked={!!scn.sharing}
                    onChange={(v) => updateScn((s) => (s.sharing = v))}
                    label="Shared core space (co-located programs share their base lab)"
                  >
                    {result.shareSavings > 0 && <Tag color="#fff" bg={BRAND.green}>saves {fmt(result.shareSavings)} GSF</Tag>}
                  </Toggle>
                  <Toggle
                    checked={!!scn.includeAnatomy}
                    onChange={(v) => updateScn((s) => (s.includeAnatomy = v))}
                    label="Include Anatomy Lab"
                  >
                    {!scn.includeAnatomy && <Tag>excluded from totals</Tag>}
                  </Toggle>
                </div>

                {data.programs.map((p) => {
                  const isAnatOff = p.id === "anat" && !scn.includeAnatomy;
                  const gsf = programGSF(p, data.multiplier);
                  const sharesHere = scn.sharing && (p.sharePct || 0) > 0 && p.id !== "anat";
                  return (
                    <div key={p.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "9px 0", borderBottom: `1px solid ${BRAND.mist}`, opacity: isAnatOff ? 0.4 : 1 }}>
                      <span style={{ width: 12, height: 12, borderRadius: 3, background: PROGRAM_COLORS[p.id], flexShrink: 0 }} />
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontWeight: 700, fontSize: 13, display: "flex", gap: 6, alignItems: "center" }}>
                          {p.code}
                          {sharesHere && <span style={{ fontSize: 9.5, fontWeight: 700, color: BRAND.blue, background: "#DCE9F5", padding: "1px 6px", borderRadius: 3 }}>shares {p.sharePct}%</span>}
                        </div>
                        <div style={{ fontSize: 11, color: BRAND.slate, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{p.name}</div>
                      </div>
                      <div style={{ textAlign: "right", fontVariantNumeric: "tabular-nums", fontSize: 12.5, fontWeight: 600, minWidth: 92 }}>
                        {isAnatOff ? "—" : fmt(gsf)}{" "}
                        <span style={{ color: BRAND.slate, fontWeight: 400 }}>GSF</span>
                      </div>
                      <select
                        value={scn.assign[p.id] || "unassigned"}
                        disabled={isAnatOff}
                        onChange={(e) => setAssign(p.id, e.target.value)}
                        style={{ padding: "6px 8px", borderRadius: 6, border: `1px solid ${BRAND.line}`, fontSize: 12, fontFamily: "'Public Sans', sans-serif", background: "#fff", maxWidth: 170 }}
                      >
                        {data.sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                        <option value="unassigned">— Unassigned —</option>
                      </select>
                    </div>
                  );
                })}
                <p style={{ fontSize: 11, color: BRAND.slate, marginTop: 12, lineHeight: 1.5 }}>
                  Set each program's shareable % on the Program Space Model tab. The DPT overlap block counts at flat SF — as rented surge space it's temporary, and when absorbed it rides the host site's existing overhead.
                </p>
              </section>

              {/* stacking diagram + totals */}
              <section style={card}>
                <h2 style={h2}>Site Stacking Plan — Live</h2>
                <StackDiagram result={result} sites={data.sites} />
                <div style={{ marginTop: 22, borderTop: `2px solid ${BRAND.line}`, paddingTop: 14, display: "flex", flexWrap: "wrap", gap: 26 }}>
                  <div>
                    <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", color: BRAND.slate }}>Permanent footprint</div>
                    <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 800, fontSize: 24, fontVariantNumeric: "tabular-nums" }}>{fmt(result.total)} <span style={{ fontSize: 13, color: BRAND.slate }}>GSF</span></div>
                  </div>
                  {result.surgeSF > 0 && (
                    <div>
                      <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", color: BRAND.slate }}>+ Surge rental</div>
                      <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 800, fontSize: 24, fontVariantNumeric: "tabular-nums" }}>{fmt(result.surgeSF)} <span style={{ fontSize: 13, color: BRAND.slate }}>SF</span></div>
                    </div>
                  )}
                  {result.shareSavings > 0 && (
                    <div>
                      <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", color: BRAND.green }}>Saved via sharing</div>
                      <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 800, fontSize: 24, color: BRAND.green, fontVariantNumeric: "tabular-nums" }}>−{fmt(result.shareSavings)}</div>
                    </div>
                  )}
                  <div>
                    <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", color: BRAND.slate }}>GSF multiplier</div>
                    <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 800, fontSize: 24, fontVariantNumeric: "tabular-nums" }}>×{data.multiplier}</div>
                  </div>
                  {result.annualCost > 0 && (
                    <div>
                      <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", color: BRAND.slate }}>Est. annual lease</div>
                      <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 800, fontSize: 24, fontVariantNumeric: "tabular-nums" }}>{fmtMoney(result.annualCost)}</div>
                    </div>
                  )}
                  {result.tiCost > 0 && (
                    <div>
                      <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", color: BRAND.slate }}>Est. one-time TI</div>
                      <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 800, fontSize: 24, fontVariantNumeric: "tabular-nums" }}>{fmtMoney(result.tiCost)}</div>
                    </div>
                  )}
                </div>
                <p style={{ fontSize: 11, color: BRAND.slate, marginTop: 12, lineHeight: 1.5 }}>
                  The striped SHARED block is the pooled core lab, counted once per site (largest sharer's need). Each sharing program keeps its dedicated remainder — offices, storage, circulation from the multiplier — as its own block. The dashed Surge Rental plate is temporary rented space, kept separate from permanent GSF. Without sharing, treat totals as the not-to-exceed number.
                </p>
              </section>
            </div>
          </>
        )}

        {/* ================= COMPARE ================= */}
        {tab === "compare" && (
          <section style={card}>
            <h2 style={h2}>All Scenarios — Side by Side</h2>
            <div style={{ overflowX: "auto" }}>
              <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 12.5, fontVariantNumeric: "tabular-nums" }}>
                <thead>
                  <tr>
                    <th style={{ textAlign: "left", padding: "8px 10px", borderBottom: `2px solid ${BRAND.blue}`, fontFamily: "'Montserrat', sans-serif", fontSize: 11, letterSpacing: "0.06em", color: BRAND.blue }}>METRIC</th>
                    {data.scenarios.map((s) => (
                      <th key={s.id} style={{ textAlign: "right", padding: "8px 10px", borderBottom: `2px solid ${BRAND.blue}`, fontFamily: "'Montserrat', sans-serif", fontSize: 11, letterSpacing: "0.03em", color: BRAND.blue, minWidth: 130 }}>{s.name}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {[...data.sites, { id: "unassigned", name: "Unassigned" }].map((site) => (
                    <tr key={site.id}>
                      <td style={{ padding: "8px 10px", borderBottom: `1px solid ${BRAND.mist}`, fontWeight: 600 }}>{site.name}</td>
                      {data.scenarios.map((s) => {
                        const r = computeScenario(s, data);
                        const b = r.bySite[site.id];
                        const over = b && b.site.capacity > 0 && b.used > b.site.capacity;
                        return (
                          <td key={s.id} style={{ padding: "8px 10px", borderBottom: `1px solid ${BRAND.mist}`, textAlign: "right", color: over ? BRAND.red : BRAND.ink, fontWeight: over ? 700 : 400 }}>
                            {b && b.used > 0 ? fmt(b.used) : "—"}{over ? " ⚠" : ""}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                  <tr>
                    <td style={{ padding: "10px", fontWeight: 800, fontFamily: "'Montserrat', sans-serif" }}>Total permanent GSF</td>
                    {data.scenarios.map((s) => (
                      <td key={s.id} style={{ padding: "10px", textAlign: "right", fontWeight: 800, fontFamily: "'Montserrat', sans-serif" }}>{fmt(computeScenario(s, data).total)}</td>
                    ))}
                  </tr>
                  <tr>
                    <td style={{ padding: "8px 10px", color: BRAND.slate }}>Shared-core savings</td>
                    {data.scenarios.map((s) => {
                      const r = computeScenario(s, data);
                      return <td key={s.id} style={{ padding: "8px 10px", textAlign: "right", color: r.shareSavings > 0 ? BRAND.green : BRAND.ink }}>{r.shareSavings > 0 ? "−" + fmt(r.shareSavings) : "—"}</td>;
                    })}
                  </tr>
                  <tr>
                    <td style={{ padding: "8px 10px", color: BRAND.slate }}>Surge rental (SF)</td>
                    {data.scenarios.map((s) => {
                      const r = computeScenario(s, data);
                      return <td key={s.id} style={{ padding: "8px 10px", textAlign: "right", color: BRAND.slate }}>{r.surgeSF > 0 ? fmt(r.surgeSF) : "—"}</td>;
                    })}
                  </tr>
                  <tr>
                    <td style={{ padding: "8px 10px", color: BRAND.slate }}>DPT overlap resolution</td>
                    {data.scenarios.map((s) => {
                      const r = computeScenario(s, data);
                      const m = (r.overlap || {}).mode;
                      const siteName = m === "absorb" ? ((data.sites.find((x) => x.id === r.overlap.siteId) || {}).name || "?") : "";
                      return (
                        <td key={s.id} style={{ padding: "8px 10px", textAlign: "right" }}>
                          {m === "surge" ? <span style={{ color: BRAND.slate }}>Rented</span>
                            : m === "absorb" ? <span style={{ color: BRAND.slate }}>Absorbed · {siteName.split(" — ")[0]}</span>
                            : <Tag color="#fff" bg={BRAND.gold}>Unresolved</Tag>}
                        </td>
                      );
                    })}
                  </tr>
                  <tr>
                    <td style={{ padding: "8px 10px", color: BRAND.slate }}>Anatomy Lab</td>
                    {data.scenarios.map((s) => (
                      <td key={s.id} style={{ padding: "8px 10px", textAlign: "right", color: BRAND.slate }}>{s.includeAnatomy ? "Included" : "Excluded"}</td>
                    ))}
                  </tr>
                  <tr>
                    <td style={{ padding: "8px 10px", color: BRAND.slate }}>PA / ARC-PA status</td>
                    {data.scenarios.map((s) => {
                      const r = computeScenario(s, data);
                      return (
                        <td key={s.id} style={{ padding: "8px 10px", textAlign: "right" }}>
                          {r.paAtRisk ? <Tag color="#fff" bg={BRAND.red}>At risk</Tag> : <Tag color="#fff" bg={BRAND.green}>Housed</Tag>}
                        </td>
                      );
                    })}
                  </tr>
                  <tr>
                    <td style={{ padding: "8px 10px", color: BRAND.slate }}>Est. annual lease</td>
                    {data.scenarios.map((s) => {
                      const r = computeScenario(s, data);
                      return <td key={s.id} style={{ padding: "8px 10px", textAlign: "right" }}>{r.annualCost > 0 ? fmtMoney(r.annualCost) : "—"}</td>;
                    })}
                  </tr>
                  <tr>
                    <td style={{ padding: "8px 10px", color: BRAND.slate }}>Est. one-time TI</td>
                    {data.scenarios.map((s) => {
                      const r = computeScenario(s, data);
                      return <td key={s.id} style={{ padding: "8px 10px", textAlign: "right" }}>{r.tiCost > 0 ? fmtMoney(r.tiCost) : "—"}</td>;
                    })}
                  </tr>
                </tbody>
              </table>
            </div>
            <p style={{ fontSize: 11, color: BRAND.slate, marginTop: 14 }}>
              Lease cost = site GSF in use × that site's rate (Sonoran $28.71/RSF/yr · Scottsdale $45/SF/yr by default — editable on Sites & Costs). Red values flag a site whose scenario demand exceeds its set capacity.
            </p>
          </section>
        )}

        {/* ================= PROGRAMS ================= */}
        {tab === "programs" && (
          <>
            <section style={{ ...card, marginBottom: 18, display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
              <h2 style={{ ...h2, margin: 0 }}>GSF Multiplier</h2>
              <NumInput value={data.multiplier} step={0.05} onChange={(v) => update((d) => (d.multiplier = v))} width={80} />
              <span style={{ fontSize: 12, color: BRAND.slate }}>Applied to net program space to cover restrooms, hallways, storage, and faculty offices. Anatomy is entered as flat GSF and toggled per scenario.</span>
            </section>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(330px, 1fr))", gap: 16 }}>
              {data.programs.map((p, pi) => {
                const nsf = programNSF(p);
                const gsf = programGSF(p, data.multiplier);
                const shared = nsf * ((p.sharePct || 0) / 100);
                return (
                  <section key={p.id} style={{ ...card, borderTop: `4px solid ${PROGRAM_COLORS[p.id]}` }}>
                    <div style={{ display: "flex", alignItems: "baseline", gap: 8, marginBottom: 4 }}>
                      <span style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 800, fontSize: 16 }}>{p.code}</span>
                      <span style={{ fontSize: 11.5, color: BRAND.slate }}>{p.name}</span>
                    </div>
                    {p.rooms.map((r, ri) => (
                      <div key={ri} style={{ display: "flex", gap: 8, alignItems: "center", marginTop: 8 }}>
                        <input
                          value={r.name}
                          onChange={(e) => update((d) => (d.programs[pi].rooms[ri].name = e.target.value))}
                          style={{ flex: 1, minWidth: 0, padding: "6px 8px", border: `1px solid ${BRAND.line}`, borderRadius: 6, fontSize: 12.5 }}
                        />
                        <NumInput value={r.nsf} width={82} onChange={(v) => update((d) => (d.programs[pi].rooms[ri].nsf = v))} />
                        <button
                          onClick={() => update((d) => d.programs[pi].rooms.splice(ri, 1))}
                          aria-label={`Remove ${r.name}`}
                          style={{ border: "none", background: "none", color: BRAND.slate, cursor: "pointer", fontSize: 15, padding: 2 }}
                        >×</button>
                      </div>
                    ))}
                    <button onClick={() => update((d) => d.programs[pi].rooms.push({ name: "New room", nsf: 0 }))} style={{ ...btn(false), marginTop: 10, fontSize: 11 }}>+ Add room</button>

                    {p.id !== "anat" && (
                      <div style={{ marginTop: 12, padding: "10px 12px", background: BRAND.mist, borderRadius: 8, display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                        <span style={{ fontSize: 11.5, fontWeight: 700, color: BRAND.blue }}>Shareable core</span>
                        <NumInput value={p.sharePct || 0} step={5} min={0} max={100} width={60} onChange={(v) => update((d) => (d.programs[pi].sharePct = Math.max(0, Math.min(100, v))))} />
                        <span style={{ fontSize: 11.5, color: BRAND.slate }}>% of NSF {shared > 0 ? `= ${fmt(shared)} SF shareable` : "(never shares)"}</span>
                      </div>
                    )}
                    {p.id === "anat" && (
                      <div style={{ marginTop: 12, padding: "10px 12px", background: BRAND.mist, borderRadius: 8, fontSize: 11.5, color: BRAND.slate }}>
                        Toggled on/off per scenario on the planner — the question is whether the lab is needed at all.
                      </div>
                    )}

                    <div style={{ marginTop: 12, paddingTop: 10, borderTop: `1px solid ${BRAND.mist}`, display: "flex", justifyContent: "space-between", fontSize: 12.5, fontVariantNumeric: "tabular-nums" }}>
                      <span style={{ color: BRAND.slate }}>{fmt(nsf)} NSF {p.usesMultiplier ? `× ${data.multiplier}` : "(flat)"}</span>
                      <span style={{ fontWeight: 800, fontFamily: "'Montserrat', sans-serif" }}>{fmt(gsf)} GSF</span>
                    </div>
                  </section>
                );
              })}
            </div>
          </>
        )}

        {/* ================= SITES ================= */}
        {tab === "sites" && (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(400px, 1fr))", gap: 16 }}>
            {data.sites.map((s, si) => (
              <section key={s.id} style={card}>
                <input
                  value={s.name}
                  onChange={(e) => update((d) => (d.sites[si].name = e.target.value))}
                  style={{ width: "100%", boxSizing: "border-box", fontFamily: "'Montserrat', sans-serif", fontWeight: 800, fontSize: 15, border: `1px solid ${BRAND.line}`, borderRadius: 6, padding: "8px 10px", color: BRAND.ink }}
                />
                <div style={{ display: "flex", gap: 18, flexWrap: "wrap", marginTop: 14 }}>
                  <label style={{ fontSize: 11.5, color: BRAND.slate, display: "flex", flexDirection: "column", gap: 4 }}>
                    Capacity (GSF) — 0 = no cap
                    <NumInput value={s.capacity} step={500} onChange={(v) => update((d) => (d.sites[si].capacity = v))} />
                  </label>
                  <label style={{ fontSize: 11.5, color: BRAND.slate, display: "flex", flexDirection: "column", gap: 4 }}>
                    Lease rate ($/SF/yr)
                    <NumInput value={s.rate} step={0.5} onChange={(v) => update((d) => (d.sites[si].rate = v))} width={76} />
                  </label>
                  <label style={{ fontSize: 11.5, color: BRAND.slate, display: "flex", flexDirection: "column", gap: 4 }}>
                    TI ($/SF one-time)
                    <NumInput value={s.ti} step={5} onChange={(v) => update((d) => (d.sites[si].ti = v))} width={70} />
                  </label>
                </div>
                <input
                  value={s.note || ""}
                  placeholder="Note…"
                  onChange={(e) => update((d) => (d.sites[si].note = e.target.value))}
                  style={{ width: "100%", boxSizing: "border-box", marginTop: 12, fontSize: 12, color: BRAND.slate, border: `1px solid ${BRAND.line}`, borderRadius: 6, padding: "7px 10px" }}
                />
              </section>
            ))}
            <section style={{ ...card, display: "flex", alignItems: "center", justifyContent: "center" }}>
              <button
                onClick={() => update((d) => d.sites.push({ id: "site" + Date.now(), name: "New Site", capacity: 0, rate: 0, ti: 0, note: "" }))}
                style={btn(true)}
              >+ Add a site</button>
            </section>
          </div>
        )}
      </main>

      <footer style={{ padding: "16px 28px", fontSize: 11, color: BRAND.slate, borderTop: `1px solid ${BRAND.line}`, background: "#fff" }}>
        Southern California University of Health Sciences · Facilities scenario model · Figures update live and save automatically. NSF = net program space; GSF applies the ×{data.multiplier} multiplier; surge rental is temporary space tracked apart from permanent GSF.
      </footer>
    </div>
  );
}
