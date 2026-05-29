"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
// @ts-ignore
import Papa from "papaparse";

export default function Home() {
  const [extensionStats, setExtensionStats] = useState<any[]>([]);
  const [totalCalls, setTotalCalls] = useState(0);

  async function handleFile(file: File) {
  const text = await file.text();

  const parsed = Papa.parse<any>(text, {
    header: true,
    skipEmptyLines: true,
  }).data;

  let extMap: any = {};

  let totalExtensionCalls = 0;
  let totalVoicemails = 0;

  parsed.forEach((row: any) => {
    const toField = (
      row["To"] || ""
    ).toString().trim();

    const lowerTo =
      toField.toLowerCase();

    //
    // IGNORE THESE
    //
    if (
      lowerTo.includes("system") ||
      lowerTo.includes("speakaccount") ||
      lowerTo.includes("call-queue")
    ) {
      return;
    }

    //
    // FIND EXTENSION
    // MATCHES:
    // 201
    // VMail (201)
    //
    const match =
      toField.match(/\d{3}/);

    if (!match) return;

    const ext = match[0];

    const num = parseInt(ext);

    //
    // EXCLUDE SYSTEM EXTENSIONS
    //
    if (
      num === 300 ||
      (num >= 400 &&
        num <= 402) ||
      (num >= 700 &&
        num <= 799)
    ) {
      return;
    }

    //
    // CREATE EXTENSION
    //
    if (!extMap[ext]) {
      extMap[ext] = {
        Extension: ext,
        Calls: 0,
        Voicemails: 0,
      };
    }

    //
    // EVERY VALID EXTENSION ROW
    // COUNTS AS A CALL
    //
    extMap[ext].Calls++;
    totalExtensionCalls++;

    //
    // COUNT VOICEMAILS
    //
    if (
      lowerTo.includes("vmail")
    ) {
      extMap[ext].Voicemails++;
      totalVoicemails++;
    }
  });

  const stats = Object.values(
    extMap
  ).sort(
    (a: any, b: any) =>
      b.Calls - a.Calls
  );

  setExtensionStats(stats);

//
// SUMMARY CARDS
//
setTotalCalls(
  totalExtensionCalls
);

}

function downloadExcel() {
    const wb = XLSX.utils.book_new();

    const exportData = [
      ...extensionStats,
    ];

    //
    // TOTALS ROW
    //
    exportData.push({
      Extension: "TOTALS",

      Calls: extensionStats.reduce(
        (sum, r) =>
          sum + r.Calls,
        0
      ),

      Voicemails:
        extensionStats.reduce(
          (sum, r) =>
            sum + r.Voicemails,
          0
        ),
    });

    const ws1 =
      XLSX.utils.json_to_sheet(
        exportData
      );

    XLSX.utils.book_append_sheet(
      wb,
      ws1,
      "Extension Summary"
    );

    XLSX.writeFile(
      wb,
      "call_report.xlsx"
    );
  }

  const totalVM =
    extensionStats.reduce(
      (sum, r) =>
        sum + r.Voicemails,
      0
    );

  return (
    <div style={outer}>
      <div style={card}>

        {/* HEADER */}
        <div style={header}>
          <img
            src="/logo.png"
            style={logo}
          />

          <h1 style={title}>
            Call Dashboard
          </h1>
        </div>

        {/* SUMMARY */}
        <div style={statsRow}>
          <Stat
  label="Total Extension Calls"
  value={totalCalls}
/>

          <Stat
            label="Total Voicemails"
            value={totalVM}
          />
        </div>

        {/* FILE UPLOAD */}
        <div
          style={{
            marginTop: 25,
          }}
        >
          <Label text="Upload PBX CSV" />

          <input
            type="file"
            onChange={(e) =>
              handleFile(
                e.target
                  .files?.[0] as File
              )
            }
            style={fileInput}
          />
        </div>

        {/* DOWNLOAD */}
        {extensionStats.length >
          0 && (
          <button
            onClick={
              downloadExcel
            }
            style={button}
          >
            Download Excel
            Report
          </button>
        )}

        {/* TABLE */}
        {extensionStats.length >
          0 && (
          <table style={table}>
            <thead style={thead}>
              <tr>
                <th style={th}>
                  Extension
                </th>

                <th style={th}>
                  Calls
                </th>

                <th style={th}>
                  Voicemails
                </th>
              </tr>
            </thead>

            <tbody>
              {extensionStats.map(
                (
                  r: any,
                  i: number
                ) => (
                  <tr key={i}>
                    <td style={td}>
                      {
                        r.Extension
                      }
                    </td>

                    <td style={td}>
                      {r.Calls}
                    </td>

                    <td style={td}>
                      {
                        r.Voicemails
                      }
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

function Stat({
  label,
  value,
}: any) {
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

function Label({
  text,
}: any) {
  return (
    <div style={label}>
      {text}
    </div>
  );
}

/* STYLES */

const outer = {
  backgroundColor:
    "#f3f4f6",
  minHeight: "100vh",
  padding: 40,
};

const card = {
  maxWidth: 1100,
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
  background:
    "#f9fafb",
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

const fileInput = {
  marginBottom: 20,
  color: "#111",
};

const button = {
  padding: "12px 18px",
  background:
    "#2563eb",
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
  background:
    "#f3f4f6",
};

const th = {
  textAlign:
    "left" as const,
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
