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

    parsed.forEach((row: any) => {
      const ext = row["To User"]?.toString().trim();

      if (ext && !isExcluded(ext, excludedList, rangeList)) {
        const date = row["Call Date"];

        if (!summary[date]) summary[date] = { total: 0 };
        summary[date].total++;

        detail.push({
          Date: date,
          Extension: ext
        });
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
          <h1 style={{ margin: 0 }}>Call Dashboard</h1>
        </div>

        {/* SUMMARY */}
        {summaryData.length > 0 && (
          <div style={summaryBox}>
            Total Calls: <strong>{totalCalls}</strong>
          </div>
        )}

        {/* INPUTS */}
        <div style={{ marginTop: 20 }}>
          <input placeholder="Exclude extensions (300,800)" onChange={e => setExcluded(e.target.value)} style={input} />
          <input placeholder="Exclude ranges (400-499)" onChange={e => setRanges(e.target.value)} style={input} />
          <input type="file" onChange={e => handleFile(e.target.files?.[0] as File)} style={{ marginTop: 10 }} />
        </div>

        {/* DOWNLOAD */}
        {summaryData.length > 0 && (
          <button onClick={downloadExcel} style={button}>
            Download Excel
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

/* 🎨 STYLES */

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
  boxShadow: "0 6px 20px rgba(0,0,0,0.08)",
};

const header = {
  display: "flex",
  alignItems: "center",
  gap: 15,
  borderBottom: "1px solid #eee",
  paddingBottom: 10,
};

const logo = {
  height: 45,
};

const summaryBox = {
  marginTop: 20,
  padding: 15,
  background: "#f9fafb",
  border: "1px solid #e5e7eb",
  borderRadius: 8,
};

const input = {
  width: "100%",
  padding: 10,
  marginTop: 10,
  borderRadius: 6,
  border: "1px solid #ccc",
};

const button = {
  marginTop: 15,
  padding: "10px 15px",
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
  borderBottom: "2px solid #ddd",
  padding: 10,
  textAlign: "left" as const,
};

const td = {
  borderBottom: "1px solid #eee",
  padding: 10,
};
