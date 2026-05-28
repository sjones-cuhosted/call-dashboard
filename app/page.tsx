"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
// @ts-ignore
import Papa from "papaparse";

export default function Home() {
  const [extensionStats, setExtensionStats] = useState<any[]>([]);
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
      .map((r) => r.trim())
      .filter(Boolean)
      .map((r) => {
        const [start, end] = r.split("-").map(Number);
        return { start, end };
      });

    return { excludedList, rangeList };
  }

  function isExcluded(
    ext: string,
    excludedList: number[],
    rangeList: any[]
  ) {
    const num = parseInt(ext);

    if (isNaN(num)) return true;

    if (excludedList.includes(num)) return true;

    for (const r of rangeList) {
      if (num >= r.start && num <= r.end) {
        return true;
      }
    }

    return false;
  }

  async function handleFile(file: File) {
    const text = await file.text();

    const parsed = Papa.parse<any>(text, {
      header: true,
      skipEmptyLines: true,
    }).data;

    const { excludedList, rangeList } = parseFilters();

    let extMap: any = {};

    parsed.forEach((row: any) => {
      const ext = row["To User"]?.toString().trim();

      if (!ext) return;

      if (
        isExcluded(
          ext,
          excludedList,
          rangeList
        )
      ) {
        return;
      }

      const toField = (row["To"] || "")
        .toString()
        .toLowerCase();

      if (!extMap[ext]) {
        extMap[ext] = {
          Extension: ext,
          TotalInbound: 0,
          Voicemails: 0,
        };
      }

      //
      // COUNT INBOUND
      //
      extMap[ext].TotalInbound++;

      //
      // COUNT VOICEMAILS
      //
      if (toField.includes("vmail")) {
        extMap[ext].Voicemails++;
      }
    });

    const stats = Object.values(extMap).sort(
      (a: any, b: any) =>
        b.TotalInbound - a.TotalInbound
    );

    setExtensionStats(stats);
  }

  function downloadExcel() {
    const wb = XLSX.utils.book_new();

    // CLONE DATA
    const exportData = [...extensionStats];

    // TOTALS ROW
    exportData.push({
      Extension: "TOTALS",

      TotalInbound: extensionStats.reduce(
        (sum, r) => sum + r.TotalInbound,
        0
      ),

      Voicemails: extensionStats.reduce(
        (sum, r) => sum + r.Voicemails,
        0
      ),
    });

    const ws1 = XLSX.utils.json_to_sheet(
      exportData
    );

    XLSX.utils.book_append_sheet(
      wb,
      ws1,
      "Extension Summary"
    );

    XLSX.writeFile(wb, "call_report.xlsx");
  }

  const totalInbound = extensionStats.reduce(
    (sum, r) => sum + r.TotalInbound,
    0
  );

  const totalVM = extensionStats.reduce(
    (sum, r) => sum + r.Voicemails,
    0
  );

  return (
    <div style={outer}>
      <div style={card}>

        {/* HEADER */}
        <div style={header}>
          <img src="/logo.png" style={logo} />

          <h1 style={title}>
            Call Dashboard
          </h1>
        </div>

        {/* STATS */}
        {extensionStats.length > 0 && (
          <div style={statsRow}>
            <Stat
              label="Total Inbound Calls"
              value={totalInbound}
            />

            <Stat
              label="Total Voicemails"
              value={totalVM}
            />
          </div>
        )}

        {/* FILTERS */}
        <div style={{ marginTop: 25 }}>
          <Label text="Exclude Extensions" />

          <input
            placeholder="300,800"
            onChange={(e) =>
              setExcluded(e.target.value)
            }
            style={input}
          />

          <Label text="Exclude Ranges" />

          <input
            placeholder="400-499,700-799"
            onChange={(e) =>
              setRanges(e.target.value)
            }
            style={input}
          />

          <Label text="Upload Call Records CSV" />

          <input
            type="file"
            onChange={(e) =>
              handleFile(
                e.target.files?.[0] as File
              )
            }
            style={fileInput}
          />
        </div>

        {/* DOWNLOAD */}
        {extensionStats.length > 0 && (
          <button
            onClick={downloadExcel}
            style={button}
          >
            Download Excel Report
          </button>
        )}

        {/* TABLE */}
        {extensionStats.length > 0 && (
          <table style={table}>
            <thead style={thead}>
              <tr>
                <th style={th}>
                  Extension
                </th>

                <th style={th}>
                  Total Inbound
                </th>

                <th style={th}>
                  Voicemails
                </th>
              </tr>
            </thead>

            <tbody>
              {extensionStats.map(
                (r: any, i: number) => (
                  <tr key={i}>
                    <td style={td}>
                      {r.Extension}
                    </td>

                    <td style={td}>
                      {r.TotalInbound}
                    </td>

                    <td style={td}>
                      {r.Voicemails}
                    </td>
                  </tr>
                )
              )}
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
      <div style={statLabel}>
        {label}
      </div>

      <div style={statValue}>
        {value}
      </div>
    </div>
  );
}

function Label({ text }: any) {
  return (
    <div style={label}>
      {text}
    </div>
  );
}

/* STYLES */

const outer = {
  backgroundColor: "#f3f4f6",
  minHeight: "100vh",
  padding: 40,
};

const card = {
  maxWidth: 1200,
  margin: "auto",
  background: "white",
  padding: 30,
  borderRadius: 12,
  boxShadow:
    "0 10px 30px rgba(0,0,0,0.08)",
};

const header = {
  display: "flex",
  alignItems: "center",
  gap: 15,
  borderBottom:
    "1px solid #e5e7eb",
  paddingBottom: 15,
};

const logo = {
  height: 45,
};

const title = {
  margin: 0,
  color: "#111",
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
  border:
    "1px solid #e5e7eb",
  padding: 15,
  borderRadius: 10,
};

const statLabel = {
  fontSize: 13,
  color: "#666",
};

const statValue = {
  fontSize: 24,
  fontWeight: "bold",
  color: "#111",
};

const label = {
  fontSize: 13,
  marginBottom: 5,
  color: "#555",
};

const input = {
  width: "100%",
  padding: 12,
  marginBottom: 15,
  borderRadius: 6,
  border:
    "1px solid #d1d5db",
  color: "#111",
};

const fileInput = {
  marginBottom: 20,
  color: "#111",
};

const button = {
  padding: "12px 18px",
  background: "#2563eb",
  color: "white",
  border: "none",
  borderRadius: 6,
  cursor: "pointer",
};

const table = {
  width: "100%",
  marginTop: 25,
  borderCollapse:
    "collapse" as const,
};

const thead = {
  background: "#f3f4f6",
};

const th = {
  textAlign: "left" as const,
  padding: 12,
  color: "#111",
  borderBottom:
    "2px solid #e5e7eb",
};

const td = {
  padding: 12,
  borderBottom:
    "1px solid #f1f1f1",
  color: "#111",
};
