// src/App.jsx
import React, { useEffect, useMemo, useState } from "react";

export default function App() {
  const [backend, setBackend] = useState(
    localStorage.getItem("backend") || "https://destined-flannelly-kieran.ngrok-free.dev"
  );

  // ✅ Global date filters (ALL exports)
  const [fromDate, setFromDate] = useState(localStorage.getItem("fromDate") || "");
  const [toDate, setToDate] = useState(localStorage.getItem("toDate") || "");

  // ✅ Persist filterBy also
  const [filterBy, setFilterBy] = useState(localStorage.getItem("filterBy") || "invoice"); // invoice | payment

  const [activeTab, setActiveTab] = useState("line");

  // ✅ Status (MAIN / FROM / TO)
  const [qboStatus, setQboStatus] = useState(null);
  const [statusLoading, setStatusLoading] = useState(false);

  // ✅ clean base url (remove trailing slash)
  const cleanBase = useMemo(() => (backend || "").replace(/\/+$/, ""), [backend]);

  // Persist
  useEffect(() => localStorage.setItem("backend", backend), [backend]);
  useEffect(() => localStorage.setItem("fromDate", fromDate), [fromDate]);
  useEffect(() => localStorage.setItem("toDate", toDate), [toDate]);
  useEffect(() => localStorage.setItem("filterBy", filterBy), [filterBy]);

  // ✅ Fetch status
  const refreshStatus = async () => {
    try {
      if (!cleanBase) return;
      setStatusLoading(true);

      const r = await fetch(`${cleanBase}/qbo/status`);
      const txt = await r.text();

      if (!r.ok) {
        console.log("STATUS API ERROR:", r.status, txt);
        setQboStatus(null);
        return;
      }

      const data = JSON.parse(txt);
      setQboStatus(data);
    } catch (e) {
      console.log("refreshStatus error:", e.message);
      setQboStatus(null);
    } finally {
      setStatusLoading(false);
    }
  };

  // Load status on backend change
  useEffect(() => {
    refreshStatus();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [cleanBase]);

  // ✅ Listen popup success (postMessage from /data_access)
  useEffect(() => {
    const handler = (event) => {
      const data = event?.data || {};
      if (data.type !== "QBO_CONNECTED") return;

      // ✅ Backend already sends: which = "main/from/to"
      // But safety: if someone sends "FROM" etc.
      const rawWhich = String(data.which || data.label || "main").toLowerCase();
      const which =
        rawWhich.includes("from") ? "from" : rawWhich.includes("to") ? "to" : "main";

      const companyName = data.companyName || "";
      const realmId = data.realmId || null;

      // ✅ Instant UI update
      setQboStatus((prev) => {
        const safe = prev || { main: {}, from: {}, to: {} };

        if (which === "from") {
          return {
            ...safe,
            from: { connected: true, companyName, realmId },
          };
        }

        if (which === "to") {
          return {
            ...safe,
            to: { connected: true, companyName, realmId },
          };
        }

        return {
          ...safe,
          main: { connected: true, companyName, realmId },
        };
      });

      // ✅ Final truth refresh
      setTimeout(() => refreshStatus(), 250);
      setTimeout(() => refreshStatus(), 1200);
    };

    window.addEventListener("message", handler);
    return () => window.removeEventListener("message", handler);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [cleanBase]);

  // -------- CONNECT BUTTONS --------
  const connectQBO = () => {
    if (!cleanBase) return alert("Please enter a valid Backend URL first.");
    window.open(`${cleanBase}/qbo/auth`, "_blank", "width=900,height=800");
  };

  const connectQBOFrom = () => {
    if (!cleanBase) return alert("Please enter a valid Backend URL first.");
    window.open(`${cleanBase}/qbo/auth-from`, "_blank", "width=900,height=800");
  };

  const connectQBOTo = () => {
    if (!cleanBase) return alert("Please enter a valid Backend URL first.");
    window.open(`${cleanBase}/qbo/auth-to`, "_blank", "width=900,height=800");
  };

  // ✅ DISCONNECT BUTTONS
  const disconnectQBO = async (which = "main") => {
    try {
      if (!cleanBase) return alert("Please enter a valid Backend URL first.");

      const r = await fetch(`${cleanBase}/qbo/disconnect?which=${which}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
      });

      if (!r.ok) {
        const txt = await r.text();
        return alert("Disconnect failed: " + txt);
      }

      await refreshStatus();
    } catch (e) {
      alert("Disconnect error: " + e.message);
    }
  };

  // -------- EXPORT HELPERS --------
  const openExport = (path) => {
    if (!cleanBase) return alert("Please enter a valid Backend URL first.");

    const qs = new URLSearchParams();
    if (fromDate) qs.set("fromDate", fromDate);
    if (toDate) qs.set("toDate", toDate);

    const url = `${cleanBase}${path}${qs.toString() ? `?${qs.toString()}` : ""}`;
    window.open(url, "_blank");
  };

  const exportAllocation = async () => {
    try {
      if (!cleanBase) return alert("Please enter a valid Backend URL first.");

      const url = `${cleanBase}/export/allocation/excel`;
      const payload = {
        fromDate: fromDate || "",
        toDate: toDate || "",
        filterBy,
      };

      const r = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      if (!r.ok) {
        const txt = await r.text();
        alert("Allocation export failed: " + txt);
        return;
      }

      const blob = await r.blob();
      const dlUrl = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = dlUrl;
      a.download = "QBO-Allocation.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(dlUrl);
    } catch (e) {
      alert("Error: " + e.message);
    }
  };

  // Attachments tab state
  const [attachTypes, setAttachTypes] = useState({
    Invoice: true,
    CreditMemo: true,
    Bill: true,
    VendorCredit: true,
    SalesReceipt: false,
    Estimate: false,
    CreditCardCharge: false,
    Purchase: false,
    Check: false,
    DelayedCharge: false,
    JournalEntry: false,
    Payment: false,
    RefundReceipt: false,
  });

  const [attachProgress, setAttachProgress] = useState(0);
  const [attachStats, setAttachStats] = useState(null);
  const [attachErrors, setAttachErrors] = useState([]);
  const [isAttachRunning, setIsAttachRunning] = useState(false);

  const toggleAttachType = (key) => {
    setAttachTypes((prev) => ({ ...prev, [key]: !prev[key] }));
  };

  const getSelectedTypes = () =>
    Object.entries(attachTypes)
      .filter(([, val]) => val)
      .map(([key]) => key);

  const startAttachmentScan = async () => {
    if (!cleanBase) return alert("Please enter a valid Backend URL first.");
    const selectedTypes = getSelectedTypes();
    if (selectedTypes.length === 0) return alert("Please select at least one transaction type.");

    setIsAttachRunning(true);
    setAttachProgress(5);
    setAttachErrors([]);
    setAttachStats(null);

    try {
      const r = await fetch(`${cleanBase}/attachments/sync`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ docTypes: selectedTypes }),
      });

      if (!r.ok) {
        const txt = await r.text();
        setAttachErrors([txt]);
        setAttachProgress(0);
        return;
      }

      const data = await r.json();
      setAttachStats(data.stats || {});
      setAttachErrors(data.errors || []);
      setAttachProgress(100);
    } catch (e) {
      setAttachErrors([e.message]);
      setAttachProgress(0);
    } finally {
      setIsAttachRunning(false);
    }
  };

  const startAttachmentCopy = async () => {
    if (!cleanBase) return alert("Please enter a valid Backend URL first.");
    const selectedTypes = getSelectedTypes();
    if (selectedTypes.length === 0) return alert("Please select at least one transaction type.");

    setIsAttachRunning(true);
    setAttachProgress(5);
    setAttachErrors([]);
    setAttachStats(null);

    try {
      const r = await fetch(`${cleanBase}/attachments/copy`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ docTypes: selectedTypes }),
      });

      if (!r.ok) {
        const txt = await r.text();
        setAttachErrors([txt]);
        setAttachProgress(0);
        return;
      }

      const data = await r.json();
      setAttachStats(data.summary || {});
      setAttachErrors(data.errors || []);
      setAttachProgress(100);
    } catch (e) {
      setAttachErrors([e.message]);
      setAttachProgress(0);
    } finally {
      setIsAttachRunning(false);
    }
  };

  // -------- UI HELPERS --------
  const mainName =
    qboStatus?.main?.connected
      ? (qboStatus.main.companyName || qboStatus.main.realmId)
      : "Not connected";

  const fromName =
    qboStatus?.from?.connected
      ? (qboStatus.from.companyName || qboStatus.from.realmId)
      : "Not connected";

  const toName =
    qboStatus?.to?.connected
      ? (qboStatus.to.companyName || qboStatus.to.realmId)
      : "Not connected";

  const mainConnected = !!qboStatus?.main?.connected;
  const fromConnected = !!qboStatus?.from?.connected;
  const toConnected = !!qboStatus?.to?.connected;

  return (
    <div className="min-h-screen bg-slate-100 py-10 px-4">
      <div className="max-w-5xl mx-auto">
        {/* HEADER */}
        <header className="mb-8">
          <h1 className="text-3xl font-bold text-slate-900 text-center">
            QBO Excel Control Panel
          </h1>

          {/* ✅ Home page company name */}
        </header>

        {/* CONNECTION / BACKEND CARD */}
        <div className="bg-white shadow-md rounded-2xl p-6 mb-6">
          <div className="grid md:grid-cols-[2fr,1fr] gap-4 items-end">
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1">
                Backend URL
              </label>

              <input
                className="w-full border border-slate-300 focus:border-indigo-500 focus:ring-indigo-500 rounded-lg px-3 py-2 text-sm outline-none"
                value={backend}
                onChange={(e) => setBackend(e.target.value)}
                placeholder="https://xxxx.ngrok-free.dev"
              />

              <p className="text-xs text-slate-400 mt-1">
                Must point to your Node/Express backend (ngrok / deployed).
              </p>

              {/* Global date filters */}
              <div className="grid md:grid-cols-2 gap-3 mt-4">
                <div>
                  <label className="block text-xs font-medium text-slate-700 mb-1">
                    From Date (all exports)
                  </label>
                  <input
                    type="date"
                    value={fromDate}
                    onChange={(e) => setFromDate(e.target.value)}
                    className="w-full border border-slate-300 rounded-lg px-2 py-1.5 text-sm focus:border-indigo-500 focus:ring-indigo-500 outline-none"
                  />
                </div>

                <div>
                  <label className="block text-xs font-medium text-slate-700 mb-1">
                    To Date (all exports)
                  </label>
                  <input
                    type="date"
                    value={toDate}
                    onChange={(e) => setToDate(e.target.value)}
                    className="w-full border border-slate-300 rounded-lg px-2 py-1.5 text-sm focus:border-indigo-500 focus:ring-indigo-500 outline-none"
                  />
                </div>
              </div>

              <p className="text-[11px] text-slate-500 mt-2">
                Blank dates = export ALL. Filled dates = TxnDate filter.
              </p>
            </div>

            <div className="flex flex-col gap-2">
              <button
                onClick={connectQBO}
                className="w-full bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium px-4 py-2.5 rounded-lg transition-colors"
              >
                🔗 Connect QBO (Main)
              </button>

              {/* ✅ DISCONNECT MAIN */}
         

              <button
                onClick={refreshStatus}
                className="w-full bg-slate-800 hover:bg-slate-900 text-white text-xs font-medium px-4 py-2 rounded-lg"
              >
                🔄 Refresh Status
              </button>
            </div>
          </div>
        </div>

        {/* TABS */}
        <div className="bg-white shadow-md rounded-2xl">
          <div className="border-b border-slate-200 flex overflow-x-auto text-sm">
            <button
              onClick={() => setActiveTab("line")}
              className={`px-4 py-3 flex-1 min-w-[120px] text-center ${
                activeTab === "line"
                  ? "text-indigo-600 border-b-2 border-indigo-600 font-semibold bg-indigo-50"
                  : "text-slate-600 hover:bg-slate-50"
              }`}
            >
              🧾 Line-Level Exports
            </button>

            <button
              onClick={() => setActiveTab("allocation")}
              className={`px-4 py-3 flex-1 min-w-[120px] text-center ${
                activeTab === "allocation"
                  ? "text-indigo-600 border-b-2 border-indigo-600 font-semibold bg-indigo-50"
                  : "text-slate-600 hover:bg-slate-50"
              }`}
            >
              📊 Allocation Report
            </button>

            <button
              onClick={() => setActiveTab("overpayment")}
              className={`px-4 py-3 flex-1 min-w-[120px] text-center ${
                activeTab === "overpayment"
                  ? "text-indigo-600 border-b-2 border-indigo-600 font-semibold bg-indigo-50"
                  : "text-slate-600 hover:bg-slate-50"
              }`}
            >
              💰 Overpayments
            </button>

            <button
              onClick={() => setActiveTab("attachments")}
              className={`px-4 py-3 flex-1 min-w-[120px] text-center ${
                activeTab === "attachments"
                  ? "text-indigo-600 border-b-2 border-indigo-600 font-semibold bg-indigo-50"
                  : "text-slate-600 hover:bg-slate-50"
              }`}
            >
              📎 Attachments
            </button>
          </div>

          <div className="p-6">
            {activeTab === "line" && (
              <section>
                <h2 className="text-lg font-semibold text-slate-900 mb-1">
                  Line-Level Exports
                </h2>
                <p className="text-xs text-slate-500 mb-4">
                  Global date range applies automatically.
                </p>

                <div className="grid md:grid-cols-3 gap-3">
                  <button
                    onClick={() => openExport("/export/invoices")}
                    className="bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-medium py-2.5 rounded-lg"
                  >
                    📄 Export Invoices
                  </button>

                  <button
                    onClick={() => openExport("/export/estimates")}
                    className="bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium py-2.5 rounded-lg"
                  >
                    📝 Export Estimates
                  </button>

                  <button
                    onClick={() => openExport("/export/creditmemos")}
                    className="bg-purple-600 hover:bg-purple-700 text-white text-sm font-medium py-2.5 rounded-lg"
                  >
                    💳 Export Credit Memos
                  </button>
                </div>
              </section>
            )}

            {activeTab === "allocation" && (
              <section>
                <h2 className="text-lg font-semibold text-slate-900 mb-1">
                  Allocation Export
                </h2>

                <div className="mb-4">
                  <span className="block text-xs font-medium text-slate-700 mb-1">
                    Filter By
                  </span>
                  <div className="flex flex-col md:flex-row gap-3 text-xs text-slate-700">
                    <label className="inline-flex items-center gap-2">
                      <input
                        type="radio"
                        value="invoice"
                        checked={filterBy === "invoice"}
                        onChange={() => setFilterBy("invoice")}
                      />
                      <span>Invoice / Credit TxnDate</span>
                    </label>

                    <label className="inline-flex items-center gap-2">
                      <input
                        type="radio"
                        value="payment"
                        checked={filterBy === "payment"}
                        onChange={() => setFilterBy("payment")}
                      />
                      <span>Payment / BillPayment TxnDate</span>
                    </label>
                  </div>
                </div>

                <button
                  onClick={exportAllocation}
                  className="w-full md:w-auto bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium py-2.5 px-6 rounded-lg"
                >
                  📊 Export Allocation Excel
                </button>
              </section>
            )}

            {activeTab === "overpayment" && (
              <section>
                <h2 className="text-lg font-semibold text-slate-900 mb-1">
                  Overpayments
                </h2>

                <button
                  onClick={() => openExport("/export/overpayments")}
                  className="w-full md:w-auto bg-orange-600 hover:bg-orange-700 text-white text-sm font-medium py-2.5 px-6 rounded-lg"
                >
                  💰 Export Overpayments Excel
                </button>
              </section>
            )}

            {activeTab === "attachments" && (
              <section>
                <h2 className="text-lg font-semibold text-slate-900 mb-1">
                  Attachment Scan & Copy (FROM → TO)
                </h2>

                <div className="grid md:grid-cols-2 gap-3 mb-3">
                  <button
                    onClick={connectQBOFrom}
                    className="w-full bg-sky-600 hover:bg-sky-700 text-white text-sm font-medium py-2.5 rounded-lg"
                  >
                    🔗 Connect QBO FROM
                  </button>

                  <button
                    onClick={connectQBOTo}
                    className="w-full bg-teal-600 hover:bg-teal-700 text-white text-sm font-medium py-2.5 rounded-lg"
                  >
                    🔗 Connect QBO TO
                  </button>
                </div>

                {/* ✅ Disconnect FROM/TO */}
                <div className="grid md:grid-cols-2 gap-3 mb-5">
                  <button
                    onClick={() => disconnectQBO("from")}
                    disabled={!fromConnected}
                    className={`w-full text-white text-sm font-medium py-2.5 rounded-lg ${
                      fromConnected ? "bg-red-600 hover:bg-red-700" : "bg-red-300 cursor-not-allowed"
                    }`}
                  >
                    🔌 Disconnect FROM
                  </button>

                  <button
                    onClick={() => disconnectQBO("to")}
                    disabled={!toConnected}
                    className={`w-full text-white text-sm font-medium py-2.5 rounded-lg ${
                      toConnected ? "bg-red-600 hover:bg-red-700" : "bg-red-300 cursor-not-allowed"
                    }`}
                  >
                    🔌 Disconnect TO
                  </button>
                </div>

                <div className="border border-slate-200 rounded-xl p-4 bg-slate-50 mb-4">
                  <h3 className="text-sm font-semibold text-slate-800 mb-2">
                    Select QBO Transaction Types
                  </h3>
                  <div className="grid md:grid-cols-3 gap-2 text-xs text-slate-700">
                    {Object.keys(attachTypes).map((key) => (
                      <label key={key} className="inline-flex items-center gap-2">
                        <input
                          type="checkbox"
                          checked={attachTypes[key]}
                          onChange={() => toggleAttachType(key)}
                        />
                        <span>{key}</span>
                      </label>
                    ))}
                  </div>
                </div>

                <div className="mb-4">
                  <div className="flex items-center justify-between mb-1">
                    <span className="text-xs font-medium text-slate-700">
                      Progress
                    </span>
                    <span className="text-xs text-slate-500">{attachProgress}%</span>
                  </div>
                  <div className="w-full bg-slate-200 rounded-full h-2 overflow-hidden">
                    <div
                      className="h-2 rounded-full bg-indigo-600 transition-all"
                      style={{ width: `${attachProgress}%` }}
                    />
                  </div>
                </div>

                <div className="flex flex-col md:flex-row gap-3 mb-4">
                  <button
                    onClick={startAttachmentScan}
                    disabled={isAttachRunning}
                    className={`w-full md:w-auto ${
                      isAttachRunning
                        ? "bg-indigo-300 cursor-not-allowed"
                        : "bg-indigo-600 hover:bg-indigo-700"
                    } text-white text-sm font-medium py-2.5 px-6 rounded-lg`}
                  >
                    {isAttachRunning ? "Running..." : "🔍 Scan Attachments (FROM)"}
                  </button>

                  <button
                    onClick={startAttachmentCopy}
                    disabled={isAttachRunning}
                    className={`w-full md:w-auto ${
                      isAttachRunning
                        ? "bg-amber-300 cursor-not-allowed"
                        : "bg-amber-600 hover:bg-amber-700"
                    } text-white text-sm font-medium py-2.5 px-6 rounded-lg`}
                  >
                    {isAttachRunning ? "Running..." : "📂 Copy Attachments"}
                  </button>
                </div>

                <div className="mt-4 grid md:grid-cols-2 gap-4">
                  <div className="border border-slate-200 rounded-xl p-4">
                    <h3 className="text-sm font-semibold text-slate-800 mb-2">Stats</h3>
                    {!attachStats ? (
                      <p className="text-xs text-slate-500">Run scan/copy to see results.</p>
                    ) : (
                      <pre className="text-xs text-slate-700 whitespace-pre-wrap">
                        {JSON.stringify(attachStats, null, 2)}
                      </pre>
                    )}
                  </div>

                  <div className="border border-slate-200 rounded-xl p-4">
                    <h3 className="text-sm font-semibold text-slate-800 mb-2">Error Log</h3>
                    {attachErrors.length === 0 ? (
                      <p className="text-xs text-slate-500">No errors yet.</p>
                    ) : (
                      <ul className="text-xs text-red-600 space-y-1 max-h-60 overflow-auto">
                        {attachErrors.map((e, i) => (
                          <li key={i}>• {e}</li>
                        ))}
                      </ul>
                    )}
                  </div>
                </div>
              </section>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
