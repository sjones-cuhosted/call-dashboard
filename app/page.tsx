"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
// @ts-ignore
import Papa from "papaparse";

export default function Home() {
  const [extensionStats, setExtensionStats] = useState<any[]>([]);
  const [excluded, setExcluded] = useState("");
  const [ranges, setRanges] = useState("");

  const [totalCallbacks, setTotalCallbacks] = useState(0);

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

  function cleanNumber(value: string) {
    return value.replace(/\D/g, "");
  }

  function isInternalDial(dialed: string) {
    const d = dialed.toLowerCase();

    return (
      d.includes("x721") ||
      d.includes("x722") ||
      d.includes("x723")
    );
  }

  function isOutboundCall(
    from: string,
    fromName: string,
    dialed: string
  ) {
    const cleanFrom = cleanNumber(from);
    const cleanDialed = cleanNumber(dialed);

    //
    // RULE 1
    // Main office number dialing out
    //
    if (
      cleanFrom === "9723149330" &&
      !isInternalDial(dialed)
    ) {
      return true;
    }

    //
    // RULE 2
    // Extension/user dialing 10 digit number
    //
    const extensionPattern = /^\d{3,4}$/;

    if (
      extensionPattern.test(fromName.trim()) &&
      cleanDialed.length === 10
    ) {
      return true;
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

    //
    // TRACK MISSED CALLERS
    //
    let missedCallers = new Set<string>();

    //
    // TRACK UNIQUE CALLBACKS
    //
    let callbackTracker = new Set<string>();

    parsed.forEach((row: any) => {
      const toUser = row["To User"]?.toString().trim();
      const fromUser = row["From User"]?.toString().trim();

      const from = (row["From"] || "").toString();
      const fromName = (row["From Name"] || "").toString();
      const dialed = (row["Dialed"] || "").toString();

      const cleanFrom = cleanNumber(from);
      const cleanDialed = cleanNumber(dialed);

      const toField = (row["To"] || "")
        .toString()
        .toLowerCase();

      //
      // =====================================
      // OUTBOUND CALL DETECTION
      // =====================================
      //
      const outbound = isOutboundCall(
        from,
        fromName,
        dialed
      );

      //
      // =====================================
      // INBOUND PROCESSING
      // =====================================
      //
      if (
        !outbound &&
        toUser &&
        !isExcluded(
          toUser,
          excludedList,
          rangeList
        )
      ) {

        if (!extMap[toUser]) {
          extMap[toUser] = {
            Extension: toUser,
            Total: 0,
            Answered: 0,
            Voicemail: 0,
            HungUp: 0,
            Callbacks: 0,
          };
        }

        extMap[toUser].Total++;

        //
        // VOICEMAIL
        //
        if (toField.includes("vmail")) {
          extMap[toUser].Voicemail++;

          if (cleanFrom.length === 10) {
            missedCallers.add(cleanFrom);
          }
        }

        //
        // HUNG UP
        //
        else if (
          toField.includes("speakaccount") ||
          toField.includes("system")
        ) {
          extMap[toUser].HungUp++;

          if (cleanFrom.length === 10) {
            missedCallers.add(cleanFrom);
          }
        }

        //
        // ANSWERED
        //
        else {
          extMap[toUser].Answered++;
        }
      }

      //
      // =====================================
      // CALLBACK PROCESSING
      // =====================================
      //
      if (
        outbound &&
        fromUser &&
        !isExcluded(
          fromUser,
          excludedList,
          rangeList
        )
      ) {

        if (!extMap[fromUser]) {
          extMap[fromUser] = {
            Extension: fromUser,
            Total: 0,
            Answered: 0,
            Voicemail: 0,
            HungUp: 0,
            Callbacks: 0,
          };
        }

        //
        // CALLBACK MATCH
        //
        if (
          missedCallers.has(cleanDialed) &&
          !callbackTracker.has(cleanDialed)
        ) {
          extMap[fromUser].Callbacks++;

          callbackTracker.add(cleanDialed);
        }
      }
    });

    const stats = Object.values(extMap).sort(
      (a: any, b: any) => b.Total - a.Total
    );

    setExtensionStats(stats);
    setTotalCallbacks(callbackTracker.size);
  }

  function downloadExcel() {
    const wb = XLSX.utils.book_new();

    const ws1 = XLSX.utils.json_to_sheet(
      extensionStats
    );

    XLSX.utils.book_append_sheet(
      wb,
      ws1,
      "Extension Summary"
    );

    XLSX.writeFile(wb, "call_report.xlsx");
  }

  const totalCalls = extensionStats.reduce(
    (sum, r) => sum + r.Total,
    0
  );

  const totalAnswered = extensionStats.reduce(
    (sum, r) => sum + r.Answered,
    0
  );

  const totalVM = extensionStats.reduce(
    (sum, r) => sum + r.Voicemail,
    0
  );

  const totalHungUp = extensionStats.reduce(
    (sum, r) => sum + r.HungUp,
    0
  );

  return (
    <div style={outer}>
      <div style={card}>

        {/* HEADER */}
        <div style={header}>
          <img src="/logo.png" style={logo} />
          <h1 style={title}>Call Dashboard</h1>
        </div>

        {/* STATS */}
        {extensionStats.length > 0 && (
          <div style={statsRow}>
            <Stat label="Total Calls" value={totalCalls} />
            <Stat label="Answered" value={totalAnswered} />
            <Stat label="Voicemails" value={totalVM} />
            <Stat label="Hung Up" value={totalHungUp} />
            <Stat label="Callbacks" value={totalCallbacks} />
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
                <th style={th}>Extension</th>
                <th style={th}>Total</th>
                <th style={th}>Answered</th>
                <th style={th}>Voicemail</th>
                <th style={th}>Hung Up</th>
                <th style={th}>Callbacks</th>
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
                      {r.Total}
                    </td>

                    <td style={td}>
                      {r.Answered}
                    </td>

                    <td style={td}>
                      {r.Voicemail}
                    </td>

                    <td style={td}>
                      {r.HungUp}
                    </td>

                    <td style={td}>
                      {r.Callbacks}
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
  maxWidth: 1300,
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
  flexWrap: "wrap" as const,
};

const statCard = {
  flex: 1,
  minWidth: 160,
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
