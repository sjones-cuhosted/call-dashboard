"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
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

    const rangeList = ranges
      .split(",")
      .map((r) => {
        const [start, end] = r.split("-").map(Number);
        return { start, end };
      });

    return { excludedList, rangeList };
  }

  function isExcluded(ext: string, excludedList: number[], rangeList: any[]) {
    if (!ext) return true;

    const cleaned = ext.toString().trim();
    const num = parseInt(cleaned);

    if (isNaN(num)) return true;

    if (excludedList.includes(num)) return true;

    for (let r of rangeList) {
      if (num >= r.start && num <= r.end) return true;
    }

    return false;
  }

  function clean(num: string) {
    if (!num) return null;
    return num.replace(/\D/g, "").slice(-10);
  }

  function getUniqueCalls(inbound: any[]) {
    const seen = new Set();
    return inbound.filter((call) => {
      const key = call.from + "_" + call.date;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  async function handleFile(file: File) {
    const text = await file.text();
    const parsed = Papa.parse<any>(text, { header: true }).data;

    const { excludedList, rangeList } = parseFilters();

    let inbound: any[] = [];
    let outbound: any[] = [];
    let voicemails: any = {};

    parsed.forEach((row: any) => {
      const toExt = row["To User"]?.toString().trim();
      const fromExt = row["From User"]?.toString().trim();

      // VOICEMAIL
      if (
        row["To"]?.toLowerCase().includes("vmail") ||
        row["Term Device"]?.toLowerCase().includes("vmail") ||
        row["To User"]?.toLowerCase().includes("vmail")
      ) {
        const date = row["Call Date"];
        if (!voicemails[date]) voicemails[date] = 0;
        voicemails[date]++;
      }

      // INBOUND (FILTERED)
      if (
        toExt &&
        toExt !== "nan" &&
        toExt !== "domain" &&
        !isExcluded(toExt, excludedList, rangeList)
      ) {
        inbound.push({
          date: row["Call Date"],
          datetime: new Date(row["Call Date"] + " " + row["Call Time"]),
          from: clean(row["From"]),
          ext: toExt,
        });
      }

      // OUTBOUND (FILTERED TOO)
      if (
        fromExt &&
        fromExt !== "nan" &&
        fromExt !== "domain" &&
        !isExcluded(fromExt, excludedList, rangeList)
      ) {
        outbound.push({
          datetime: new Date(row["Call Date"] + " " + row["Call Time"]),
          to: clean(row["To"]),
          ext: fromExt,
        });
      }
    });

    const uniqueInbound = getUniqueCalls(inbound);

    let summary: any = {};
    let extSummary: any = {};

    uniqueInbound.forEach((call) => {
      if (!summary[call.date]) {
        summary[call.date] = {
          total: 0,
          callbacks: 0,
          voicemails: voicemails[call.date] || 0,
        };
      }

      summary[call.date].total++;

      if (!extSummary[call.date]) extSummary[call.date] = {};
      if (!extSummary[call.date][call.ext]) {
        extSummary[call.date][call.ext] = {
          inbound: 0,
          callbacks: 0,
        };
      }

      extSummary[call.date][call.ext].inbound++;
    });

    uniqueInbound.forEach((call) => {
      const match = outbound.find(
        (o) => o.to === call.from && o.datetime > call.datetime
      );

      if (match) {
        summary[call.date].callbacks++;

        if (extSummary[call.date]?.[match.ext]) {
          extSummary[call.date][match.ext].callbacks++;
        }
      }
    });

    const formatted = Object.entries(summary)
      .map(([date, val]: any) => {
        const d = new Date(date);

        return {
          Day: d.toLocaleDateString("en-US", { weekday: "long" }),
          "Formatted Date": d.toLocaleDateString("en-US"),
          "Total Calls": val.total,
          Voicemails: val.voicemails,
          Callbacks: val.callbacks,
        };
      })
      .sort(
        (a: any, b: any) =>
          new Date(a["Formatted Date"]).getTime() -
          new Date(b["Formatted Date"]).getTime()
      );

    setResults({ summary: formatted, extSummary });
  }

  function downloadExcel() {
    const wb = XLSX.utils.book_new();

    const ws1 = XLSX.utils.json_to_sheet(results.summary);
    XLSX.utils.book_append_sheet(wb, ws1, "Summary");

    let extRows: any[] = [];

    const sortedDates = Object.keys(results.extSummary).sort(
      (a, b) => new Date(a).getTime() - new Date(b).getTime()
    );

    sortedDates.forEach((date, index) => {
      const exts = results.extSummary[date];

      Object.entries(exts).forEach(([ext, val]: any) => {
        extRows.push({
          "Formatted Date": date,
          Extension: ext,
          "Inbound Calls": val.inbound,
          Callbacks: val.callbacks,
        });
      });

      if (index < sortedDates.length - 1) {
        extRows.push({});
      }
    });

    const ws2 = XLSX.utils.json_to_sheet(extRows);
    XLSX.utils.book_append_sheet(wb, ws2, "Extension Detail");

    XLSX.writeFile(wb, "call_report.xlsx");
  }

  return (
    <div style={{ padding: 40 }}>
      <h1>Call Dashboard</h1>

      <input
        placeholder="Exclude extensions (300,450,800)"
        onChange={(e) => setExcluded(e.target.value)}
      />

      <br /><br />

      <input
        placeholder="Exclude ranges (400-499,700-799)"
        onChange={(e) => setRanges(e.target.value)}
      />

      <br /><br />

      <input
        type="file"
        onChange={(e) => handleFile(e.target.files?.[0] as File)}
      />

      {results?.summary && (
        <>
          <h2>Summary</h2>

          <table border={1}>
            <thead>
              <tr>
                <th>Day</th>
                <th>Date</th>
                <th>Total Calls</th>
                <th>Voicemails</th>
                <th>Callbacks</th>
              </tr>
            </thead>
            <tbody>
              {results.summary.map((r: any, i: number) => (
                <tr key={i}>
                  <td>{r.Day}</td>
                  <td>{r["Formatted Date"]}</td>
                  <td>{r["Total Calls"]}</td>
                  <td>{r.Voicemails}</td>
                  <td>{r.Callbacks}</td>
                </tr>
              ))}
            </tbody>
          </table>

          <br />

          <button onClick={downloadExcel}>
            Download Final Excel
          </button>
        </>
      )}
    </div>
  );
}