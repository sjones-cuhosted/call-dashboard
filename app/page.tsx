"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
// @ts-ignore
import Papa from "papaparse";

export default function Home() {
  const [results, setResults] = useState<any>(null);
  const [rawData, setRawData] = useState<any[]>([]);
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

        // Summary
        if (!summary[date]) summary[date] = { total: 0 };
        summary[date].total++;

        // Detail
        detail.push({
          Date: date,
          Extension: ext
        });
      }
    });

    const formatted = Object.entries(summary).map(([date, val]: any) => {
      const d = new Date(date);
      return {
        Day: d.toLocaleDateString("en-US", { weekday: "long" }),
        Date: d.toLocaleDateString("en-US"),
        Total: val.total
      };
    });

    setResults(formatted);
    setRawData(detail);
  }

  function downloadExcel() {
    const wb = XLSX.utils.book_new();

    // Tab 1
    const ws1 = XLSX.utils.json_to_sheet(results);
    XLSX.utils.book_append_sheet(wb, ws1, "Summary");

    // Tab 2
    const ws2 = XLSX.utils.json_to_sheet(rawData);
    XLSX.utils.book_append_sheet(wb, ws2, "Details");

    XLSX.writeFile(wb, "call_report.xlsx");
  }

  const totalCalls = results
    ? results.reduce((sum: number, r: any) => sum + r.Total, 0)
    : 0;

  return (
    <div style={{ backgroundColor: "#f9fafb", minHeight: "100vh", padding: 40 }}>
      <div style={{ maxWidth: 900, margin: "auto", background: "white", padding: 30, borderRadius: 12 }}>
        <h1>📞 Call Dashboard</h1>

        {results && (
          <div style={{ marginTop: 20 }}>
            <strong>Total Calls: {totalCalls}</strong>
          </div>
        )}

        <input placeholder="Exclude extensions" onChange={e => setExcluded(e.target.value)} />
        <br /><br />
        <input placeholder="Exclude ranges" onChange={e => setRanges(e.target.value)} />
        <br /><br />
        <input type="file" onChange={e => handleFile(e.target.files?.[0] as File)} />

        {results && (
          <>
            <br /><br />
            <button onClick={downloadExcel}>Download Excel (2 Tabs)</button>
          </>
        )}
      </div>
    </div>
  );
}
