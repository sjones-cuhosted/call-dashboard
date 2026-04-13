"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
// @ts-ignore
import Papa from "papaparse";

export default function Home() {
  const [results, setResults] = useState<any>(null);
  const [excluded, setExcluded] = useState("");
  const [ranges, setRanges] = useState("");

  function parseFilters() {
    const excludedList = excluded
      .split(",")
      .map((x) => x.trim())
      .filter(Boolean)
      .map(Number);

    const rangeList = ranges.split(",").map((r) => {
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

    parsed.forEach((row: any) => {
      const ext = row["To User"]?.toString().trim();

      if (ext && !isExcluded(ext, excludedList, rangeList)) {
        const date = row["Call Date"];

        if (!summary[date]) {
          summary[date] = { total: 0 };
        }

        summary[date].total++;
      }
    });

    const formatted = Object.entries(summary).map(([date, val]: any) => {
      const d = new Date(date);
      return {
        Day: d.toLocaleDateString("en-US", { weekday: "long" }),
        Date: d.toLocaleDateString("en-US"),
        Total: val.total,
      };
    });

    setResults(formatted);
  }

  function downloadExcel() {
    if (!results) return;

    const ws = XLSX.utils.json_to_sheet(results);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Summary");

    XLSX.writeFile(wb, "call_report.xlsx");
  }

  const totalCalls = results
    ? results.reduce((sum: number, r: any) => sum + r.Total, 0)
    : 0;

  return (
    <div
      style={{
        backgroundColor: "#f9fafb",
        minHeight: "100vh",
        padding: 40,
        fontFamily: "Arial, sans-serif",
        color: "#111",
      }}
    >
      <div
        style={{
          maxWidth: 900,
          margin: "auto",
          background: "white",
          padding: 30,
          borderRadius: 12,
          boxShadow: "0 4px 12px rgba(0,0,0,0.08)",
        }}
      >
        <h1 style={{ fontSize: 28 }}>📞 Call Dashboard</h1>

        {/* 🔥 SUMMARY CARDS */}
        {results && (
          <div style={{ display: "flex", gap: 20, marginTop: 20 }}>
            <div style={card}>
              <h3>Total Calls</h3>
              <p style={cardNumber}>{totalCalls}</p>
            </div>
          </div>
        )}

        <div style={{ marginTop: 20 }}>
          <label style={label}>Exclude Extensions</label>
          <input
            placeholder="300,800"
            onChange={(e) => setExcluded(e.target.value)}
            style={inputStyle}
          />

          <label style={label}>Exclude Ranges</label>
          <input
            placeholder="400-499,700-799"
            onChange={(e) => setRanges(e.target.value)}
            style={inputStyle}
          />

          <label style={label}>Upload File</label>
          <input
            type="file"
            onChange={(e) => handleFile(e.target.files?.[0] as File)}
          />
        </div>

        {/* 🔥 DOWNLOAD BUTTON */}
        {results && (
          <button onClick={downloadExcel} style={button}>
            Download Excel
          </button>
        )}

        {results && (
          <table
            style={{
              marginTop: 20,
              width: "100%",
              borderCollapse: "collapse",
            }}
          >
            <thead>
              <tr style={{ background: "#f3f4f6" }}>
                <th style={th}>Day</th>
                <th style={th}>Date</th>
                <th style={th}>Total Calls</th>
              </tr>
            </thead>
            <tbody>
              {results.map((r: any, i: number) => (
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

const card = {
  flex: 1,
  background: "#f3f4f6",
  padding: 20,
  borderRadius: 10,
};

const cardNumber = {
  fontSize: 28,
  fontWeight: "bold",
};

const label = {
  display: "block",
  marginTop: 10,
  marginBottom: 5,
  fontWeight: "bold",
};

const inputStyle = {
  width: "100%",
  padding: 10,
  marginBottom: 10,
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

const th = {
  padding: 10,
  textAlign: "left" as const,
};

const td = {
  padding: 10,
  borderTop: "1px solid #eee",
};
