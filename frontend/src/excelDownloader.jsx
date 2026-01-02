// src/ExcelDownloader.jsx
import React, { useState } from "react";

export default function ExcelDownloader() {
  const [backend, setBackend] = useState(
    localStorage.getItem("backend") ||
      "https://rejectedly-normative-emma.ngrok-free.dev"
  );

  const [includeInvoice, setIncludeInvoice] = useState(true);
  const [includeEstimate, setIncludeEstimate] = useState(true);
  const [includeCreditMemo, setIncludeCreditMemo] = useState(true);

  const connectQBO = () => {
    const cleanBase = backend.replace(/\/$/, "");
    if (!cleanBase) {
      alert("Please enter a valid Backend URL first.");
      return;
    }
    const url = `${cleanBase}/qbo/auth`;
    window.open(url, "_blank", "width=900,height=800");
    alert("Login to QuickBooks, then return here.");
  };

  const exportAll = async () => {
    const cleanBase = backend.replace(/\/$/, "");
    if (!cleanBase) {
      alert("Please enter a valid Backend URL first.");
      return;
    }

    if (!includeInvoice && !includeEstimate && !includeCreditMemo) {
      alert("Please select at least one document type.");
      return;
    }

    const payload = {
      includeInvoice,
      includeEstimate,
      includeCreditMemo,
    };

    const r = await fetch(`${cleanBase}/export/selected/excel`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!r.ok) {
      const msg = await r.text();
      alert("Export failed: " + msg);
      return;
    }

    const blob = await r.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "QBO-Selected-Docs.xlsx";
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="min-h-screen bg-slate-100 py-10 px-4">
      <div className="max-w-xl mx-auto bg-white shadow-md rounded-xl p-6">
        <h1 className="text-2xl font-bold mb-4 text-slate-900">
          QBO – Selected Docs Export
        </h1>
        <p className="text-xs text-slate-500 mb-4">
          Use this screen if you want to generate a single Excel file for only
          selected document types (Invoices / Estimates / Credit Memos).
        </p>

        {/* Backend URL */}
        <div className="mb-4">
          <label className="block text-sm font-medium text-slate-700 mb-1">
            Backend URL
          </label>
          <input
            className="w-full border border-slate-300 focus:border-indigo-500 focus:ring-indigo-500 rounded-lg px-3 py-2 text-sm outline-none"
            value={backend}
            onChange={(e) => {
              const value = e.target.value;
              setBackend(value);
              localStorage.setItem("backend", value);
            }}
            placeholder="https://your-ngrok-url.ngrok-free.app"
          />
        </div>

        {/* Connect QBO */}
        <button
          onClick={connectQBO}
          className="w-full bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium py-2.5 rounded-lg mb-5"
        >
          🔗 Connect QBO
        </button>

        {/* Checkboxes */}
        <div className="border border-slate-200 rounded-lg p-4 mb-4 bg-slate-50">
          <h2 className="text-sm font-semibold text-slate-800 mb-3">
            Select Document Types
          </h2>

          <label className="flex items-center gap-2 mb-2 text-sm text-slate-700">
            <input
              type="checkbox"
              checked={includeInvoice}
              onChange={() => setIncludeInvoice(!includeInvoice)}
            />
            <span>Invoice</span>
          </label>

          <label className="flex items-center gap-2 mb-2 text-sm text-slate-700">
            <input
              type="checkbox"
              checked={includeEstimate}
              onChange={() => setIncludeEstimate(!includeEstimate)}
            />
            <span>Estimate</span>
          </label>

          <label className="flex items-center gap-2 mb-1 text-sm text-slate-700">
            <input
              type="checkbox"
              checked={includeCreditMemo}
              onChange={() => setIncludeCreditMemo(!includeCreditMemo)}
            />
            <span>Credit Memo</span>
          </label>
        </div>

        <button
          onClick={exportAll}
          className="w-full bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-medium py-2.5 rounded-lg"
        >
          📥 Export Selected to Excel
        </button>
      </div>
    </div>
  );
}
