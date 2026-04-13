"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
// @ts-ignore
import Papa from "papaparse";

export default function Home() {
  const [summaryData, setSummaryData] = useState<any[]>([]);
  const [detailData, setDetailData] = useState<any[]>([]);
  const [excluded, setExcluded] = useState("");
  const [ranges, setRanges] = useState("");

  const [voicemails, setVoicemails] = useState(0);
  const [callbacks, setCallbacks] = useState(0);

  function parseFilters() {
    const excludedList = excluded.split(",").map(x => x.trim()).filter(Boolean).map(Number);

    const rangeList = ranges.split(",").map(r => {
      const [start, end] = r.split("-").map(Number);
      return { start, end };
    });

    return { excludedList, rangeList };
  }

  function isExcluded(ext: string, excludedList: number[], rangeList: any[]) {
    if (!ext) return true;
    const num = parseInt(ext);
    if (isNaN(num)) return true;

    if (excludedList.includes(num)) return true;

    for (let r of rangeList) {
      if (num >= r.start && num <= r.end) return true;
    }

    return false;
  }

  async function handleFile(file: File) {
    const text = await file.text();
    const parsed = Papa.parse<any>(text, { header: true }).data;

    const { excludedList, rangeList } = parseFilters();

    let summary: any = {};
    let detail: any[] = [];

    let vmCount = 0;
    let callMap: any = {};
    let callbackCount = 0;

    parsed.forEach((row: any) => {
      const ext = row["To User"]?.toString().trim();
      const disposition = row["Disposition"] || "";
      const caller = row["From"] || "";

      if (ext && !isExcluded(ext, excludedList, rangeList)) {
        const date = row["Call Date"];

        if (!summary[date]) summary[date] = { total: 0 };
        summary[date].total++;

        detail.push({
          Date: date,
          Extension: ext
        });

        if (disposition.toLowerCase().includes("vmail")) {
          vmCount++;
        }

        if (!callMap[caller]) {
          callMap[caller] = 1;
        } else {
          callbackCount++;
        }
      }
    });

    const formattedSummary = Object.entries(summary).map(([date, val]: any) => {
      const d = new Date(date);
      return {
        Day: d.toLocaleDateString("en-US", { weekday: "long" }),
        Date: d.toLocaleDateString("en-US"),
        Total: val.total
      };
    });

    setSummaryData(formattedSummary);
    setDetailData(detail);
    setVoicemails(vmCount);
    setCallbacks(callbackCount);
  }

  function downloadExcel() {
    const wb = XLSX.utils.book_new();

    const ws1 = XLSX.utils.json_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(wb, ws1, "Summary");

    const ws2 = XLSX.utils.json_to_sheet(detailData);
    XLSX.utils.book_append_sheet(wb, ws2, "Details");

    XLSX.writeFile(wb, "call_report.xlsx");
  }

  const totalCalls = summaryData.reduce((sum, r) => sum + r.Total, 0);

  return (
    <div style={outer}>
      <div style={card}>

        {/* HEADER */}
        <div style={header}>
          <img src="/logo.png" style={logo} />
          <h1 style={title}>Call Dashboard</h1>
        </div>

        {/* STATS */}
        {summaryData.length > 0 && (
          <div style={statsRow}>
            <Stat label="Total Calls" value={totalCalls} />
            <Stat label="Voicemails" value={voicemails} />
            <Stat label="Callbacks" value={callbacks} />
          </div>
        )}

        {/* INPUTS */}
        <div style={form}>
          <Label text="Exclude Extensions" />
          <input
            placeholder="300, 800"
            onChange={e => setExcluded(e.target.value)}
            style={input}
          />

          <Label text="Exclude Ranges" />
          <input
            placeholder="400-499, 700-799"
            onChange={e => setRanges(e.target.value)}
            style={input}
          />

          <Label text="Upload Call File" />
          <input
            type="file"
            onChange={e => handleFile(e.target.files?.[0] as File)}
            style={fileInput}
          />
        </div>

        {/* DOWNLOAD */}
        {summaryData.length > 0 && (
          <button onClick={downloadExcel} style={button}>
            Download Excel Report
          </button>
        )}

        {/* TABLE */}
        {summaryData.length > 0 && (
          <table style={table}>
            <thead>
              <tr>
                <th style={th}>Day</th>
                <th style={th}>Date</th>
                <th style={th}>Total Calls</th>
              </tr>
            </thead>
            <tbody>
              {summaryData.map((r, i) => (
                <tr key={i}>
                  <td style={td}>{r.Day}</td>
                  <td style={td}>{r.Date}</td>
                  <td style={td}>{r.Total}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}

      </div>
    </div>
  );
}

/* COMPONENTS */

function Stat({ label, value }: any) {
  return (
    <div style={statCard}>
      <div style={statLabel}>{label}</div>
      <div style={statValue}>{value}</div>
    </div>
  );
}

function Label({ text }: any) {
  return <div style={label}>{text}</div>;
}

/* STYLES */

const outer = {
  backgroundColor: "#f3f4f6",
  minHeight: "100vh",
  padding: 40,
};

const card = {
  maxWidth: 900,
  margin: "auto",
  background: "white",
  padding: 30,
  borderRadius: 12,
  boxShadow: "0 10px 30px rgba(0,0,0,0.08)",
};

const header = {
  display: "flex",
  alignItems: "center",
  gap: 15,
  borderBottom: "1px solid #e5e7eb",
  paddingBottom: 15,
};

const logo = { height: 45 };

const title = {
  margin: 0,
  color: "#111",
  fontSize: 24,
  fontWeight: 600,
};

const statsRow = {
  display: "flex",
  gap: 15,
  marginTop: 20,
};

const statCard = {
  flex: 1,
  background: "#f9fafb",
  border: "1px solid #e5e7eb",
  padding: 15,
  borderRadius: 10,
};

const statLabel = { fontSize: 13, color: "#666" };
const statValue = { fontSize: 22, fontWeight: "bold", color: "#111" };

const form = { marginTop: 25 };

const label = {
  fontSize: 13,
  color: "#555",
  marginBottom: 5,
};

const input = {
  width: "100%",
  padding: 12,
  marginBottom: 15,
  borderRadius: 6,
  border: "1px solid #d1d5db",
  color: "#111",
};

const fileInput = {
  marginTop: 5,
  marginBottom: 15,
  color: "#111",
};

const button = {
  marginTop: 10,
  padding: "12px 18px",
  background: "#2563eb",
  color: "white",
  border: "none",
  borderRadius: 6,
  cursor: "pointer",
};

const table = {
  width: "100%",
  marginTop: 20,
  borderCollapse: "collapse" as const,
};

const th = {
  textAlign: "left" as const,
  borderBottom: "2px solid #e5e7eb",
  padding: 10,
};

const td = {
  padding: 10,
  borderBottom: "1px solid #f1f1f1",
};
