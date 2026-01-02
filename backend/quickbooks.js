// quickbooks.js
import express from "express";
import axios from "axios";
import dotenv from "dotenv";
import ExcelJS from "exceljs";
import { loadTaxMaster, extractLineTaxes } from "./extractTax.js";

dotenv.config();
const router = express.Router();

const { CLIENT_ID, CLIENT_SECRET } = process.env;

const AUTH_URL = "https://appcenter.intuit.com/connect/oauth2";
const TOKEN_URL = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer";
const MINOR = 75;

// ------------------------------
// MAIN / FROM / TO connections
// ------------------------------
global.QBO = global.QBO || {
  access_token: null,
  refresh_token: null,
  realmId: null,
  companyName: null,
};

global.QBO_FROM = global.QBO_FROM || {
  access_token: null,
  refresh_token: null,
  realmId: null,
  companyName: null,
};

global.QBO_TO = global.QBO_TO || {
  access_token: null,
  refresh_token: null,
  realmId: null,
  companyName: null,
};

// ======================================================================================
// Helpers
// ======================================================================================
function buildDateWhere(field, from, to) {
  if (!from && !to) return "";
  if (from && to) return `${field} >= '${from}' AND ${field} <= '${to}'`;
  if (from) return `${field} >= '${from}'`;
  return `${field} <= '${to}'`;
}

function getConn(which = "main") {
  const w = String(which || "main").toLowerCase();
  if (w === "from") return global.QBO_FROM;
  if (w === "to") return global.QBO_TO;
  return global.QBO;
}

function setConn(which, obj) {
  const w = String(which || "main").toLowerCase();
  if (w === "from") global.QBO_FROM = obj;
  else if (w === "to") global.QBO_TO = obj;
  else global.QBO = obj;
}

async function runQuery({ realmId, token, query }) {
  const url = `https://quickbooks.api.intuit.com/v3/company/${realmId}/query?query=${encodeURIComponent(
    query
  )}&minorversion=${MINOR}`;

  const r = await axios.get(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  });

  return r.data.QueryResponse || {};
}

async function fetchDocsPage({
  entityName,
  realmId,
  token,
  whereClause = "",
  startPos = 1,
  pageSize = 1000,
}) {
  let query = `SELECT * FROM ${entityName}`;
  if (whereClause && whereClause.trim()) query += ` WHERE ${whereClause}`;
  query += ` STARTPOSITION ${startPos} MAXRESULTS ${pageSize}`;
  const resp = await runQuery({ realmId, token, query });
  const list = resp?.[entityName] || [];
  const totalCount = resp?.totalCount;
  return { list, totalCount };
}

async function fetchAllDocs({ entityName, realmId, token, whereClause = "" }) {
  const pageSize = 1000;
  let startPos = 1;
  let all = [];

  while (true) {
    let query = `SELECT * FROM ${entityName}`;
    if (whereClause && whereClause.trim()) query += ` WHERE ${whereClause}`;
    query += ` STARTPOSITION ${startPos} MAXRESULTS ${pageSize}`;

    const resp = await runQuery({ realmId, token, query });
    const list = resp?.[entityName] || [];
    if (!list.length) break;

    all = all.concat(list);
    if (list.length < pageSize) break;
    startPos += pageSize;
  }
  return all;
}

async function fetchDoc(type, id, realmId, token) {
  const url = `https://quickbooks.api.intuit.com/v3/company/${realmId}/${type}/${id}?minorversion=${MINOR}`;
  return axios.get(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  });
}

async function getCompanyName(realmId, token) {
  const url = `https://quickbooks.api.intuit.com/v3/company/${realmId}/companyinfo/${realmId}?minorversion=${MINOR}`;
  const r = await axios.get(url, {
    headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
  });
  const ci = r.data?.CompanyInfo || {};
  return (
    ci.CompanyName ||
    ci.LegalName ||
    ci.CompanyNameFormatted ||
    ci.CompanyNameOnChecks ||
    ""
  );
}

// ======================================================================================
// STATUS ROUTE
// ======================================================================================
router.get("/qbo/status", (req, res) => {
  const main = global.QBO || {};
  const from = global.QBO_FROM || {};
  const to = global.QBO_TO || {};

  return res.json({
    main: {
      connected: !!(main.access_token && main.realmId),
      realmId: main.realmId || null,
      companyName: main.companyName || null,
    },
    from: {
      connected: !!(from.access_token && from.realmId),
      realmId: from.realmId || null,
      companyName: from.companyName || null,
    },
    to: {
      connected: !!(to.access_token && to.realmId),
      realmId: to.realmId || null,
      companyName: to.companyName || null,
    },
  });
});

// ======================================================================================
// OAuth Redirects
// ======================================================================================
function buildAuthUrl(state) {
  if (!global.PUBLIC_URL) return null;
  const REDIRECT_URI = `${global.PUBLIC_URL}/data_access`;
  const scope =
    "com.intuit.quickbooks.accounting openid profile email phone address";

  return `${AUTH_URL}?client_id=${CLIENT_ID}&redirect_uri=${encodeURIComponent(
    REDIRECT_URI
  )}&response_type=code&scope=${encodeURIComponent(scope)}&state=${state}`;
}

router.get("/qbo/auth", (req, res) => {
  const url = buildAuthUrl("qbo_main");
  if (!url) return res.send("⏳ PUBLIC_URL not set (ngrok/public domain missing).");
  return res.redirect(url);
});

router.get("/qbo/auth-from", (req, res) => {
  const url = buildAuthUrl("qbo_from");
  if (!url) return res.send("⏳ PUBLIC_URL not set (ngrok/public domain missing).");
  return res.redirect(url);
});

router.get("/qbo/auth-to", (req, res) => {
  const url = buildAuthUrl("qbo_to");
  if (!url) return res.send("⏳ PUBLIC_URL not set (ngrok/public domain missing).");
  return res.redirect(url);
});

// ======================================================================================
// OAuth Callback
// ======================================================================================
router.get("/data_access", async (req, res) => {
  const { code, realmId, state = "qbo_main" } = req.query;
  if (!code) return res.send("❌ Missing authorization code");

  try {
    const REDIRECT_URI = `${global.PUBLIC_URL}/data_access`;
    const basicAuth = Buffer.from(`${CLIENT_ID}:${CLIENT_SECRET}`).toString(
      "base64"
    );

    const body = new URLSearchParams({
      grant_type: "authorization_code",
      code,
      redirect_uri: REDIRECT_URI,
    }).toString();

    const r = await axios.post(TOKEN_URL, body, {
      headers: {
        Authorization: `Basic ${basicAuth}`,
        "Content-Type": "application/x-www-form-urlencoded",
      },
    });

    const tokenObj = {
      access_token: r.data.access_token,
      refresh_token: r.data.refresh_token,
      realmId,
      companyName: "",
    };

    try {
      tokenObj.companyName = await getCompanyName(realmId, tokenObj.access_token);
    } catch {
      tokenObj.companyName = "";
    }

    const sendPopupSuccess = (label) => {
      const name = tokenObj.companyName || realmId || "";
      const which = String(label || "").toLowerCase(); // main/from/to

      const safeCompany = String(tokenObj.companyName || "")
        .replace(/\\/g, "\\\\")
        .replace(/"/g, '\\"')
        .replace(/\n/g, " ");

      const safeRealm = String(realmId || "")
        .replace(/\\/g, "\\\\")
        .replace(/"/g, '\\"')
        .replace(/\n/g, " ");

      return res.send(`
        <html>
          <head><title>QBO Connected</title><meta charset="utf-8" /></head>
          <body style="font-family: Arial; padding: 18px;">
            <h3>✅ QBO ${label} Connected</h3>
            <p style="margin: 6px 0;"><b>${name}</b></p>
            <p style="color:#555; font-size: 13px;">This window will close automatically...</p>
            <script>
              try {
                if (window.opener) {
                  window.opener.postMessage(
                    {
                      type: "QBO_CONNECTED",
                      which: "${which}",
                      companyName: "${safeCompany}",
                      realmId: "${safeRealm}"
                    },
                    "*"
                  );
                }
              } catch(e) {}
              setTimeout(() => window.close(), 1500);
            </script>
          </body>
        </html>
      `);
    };

    if (state === "qbo_from") {
      setConn("from", tokenObj);
      console.log("✅ QBO FROM connected:", realmId, tokenObj.companyName);
      return sendPopupSuccess("FROM");
    } else if (state === "qbo_to") {
      setConn("to", tokenObj);
      console.log("✅ QBO TO connected:", realmId, tokenObj.companyName);
      return sendPopupSuccess("TO");
    } else {
      setConn("main", tokenObj);
      console.log("✅ QBO MAIN connected:", realmId, tokenObj.companyName);
      return sendPopupSuccess("MAIN");
    }
  } catch (err) {
    console.log("OAuth ERROR:", err.response?.data || err.message);
    return res.status(500).json(err.response?.data || { error: err.message });
  }
});

// ======================================================================================
// DISCONNECT
// ======================================================================================
router.post("/qbo/disconnect", (req, res) => {
  const which = (req.body?.which || req.query?.which || "main").toLowerCase();
  const blank = { access_token: null, refresh_token: null, realmId: null, companyName: null };
  setConn(which, { ...blank });
  return res.json({ success: true, disconnected: which });
});

// ======================================================================================
// Debug Extract Routes
// ======================================================================================
router.get("/extract/estimate/:id", async (req, res) => {
  try {
    const QBO = getConn("main");
    if (!QBO.access_token) return res.send("❌ Not connected to QBO.");

    const estimateId = req.params.id;
    const taxMaster = await loadTaxMaster(QBO.realmId, QBO.access_token);
    const doc = await fetchDoc("estimate", estimateId, QBO.realmId, QBO.access_token);
    const result = extractLineTaxes(doc.data.Estimate, taxMaster);
    res.json({ type: "Estimate", estimateId, taxLines: result });
  } catch (err) {
    res.status(500).json(err.response?.data || { error: err.message });
  }
});

router.get("/extract/creditmemo/:id", async (req, res) => {
  try {
    const QBO = getConn("main");
    if (!QBO.access_token) return res.send("❌ Not connected to QBO.");

    const cmId = req.params.id;
    const taxMaster = await loadTaxMaster(QBO.realmId, QBO.access_token);
    const doc = await fetchDoc("creditmemo", cmId, QBO.realmId, QBO.access_token);
    const result = extractLineTaxes(doc.data.CreditMemo, taxMaster);
    res.json({ type: "CreditMemo", creditMemoId: cmId, taxLines: result });
  } catch (err) {
    res.status(500).json(err.response?.data || { error: err.message });
  }
});

router.get("/raw/invoice/:id", async (req, res) => {
  try {
    const QBO = getConn("main");
    if (!QBO.access_token) return res.send("❌ Not connected to QBO.");

    const url = `https://quickbooks.api.intuit.com/v3/company/${QBO.realmId}/invoice/${req.params.id}?minorversion=${MINOR}`;
    const r = await axios.get(url, {
      headers: { Authorization: `Bearer ${QBO.access_token}`, Accept: "application/json" },
    });
    res.json(r.data);
  } catch (err) {
    res.status(500).json(err.response?.data || { error: err.message });
  }
});

// ======================================================================================
// ✅ EXPORT INVOICES (Streaming + Pagination)  ✅ OOM SAFE  ✅ ONLY ONE ROUTE
// ======================================================================================
router.get("/export/invoices", async (req, res) => {
  try {
    const QBO = getConn("main");
    if (!QBO.access_token || !QBO.realmId) {
      return res.status(400).send("❌ Not connected to QBO.");
    }

    const { fromDate = "", toDate = "" } = req.query;
    const where = buildDateWhere("TxnDate", fromDate, toDate);
    const taxMaster = await loadTaxMaster(QBO.realmId, QBO.access_token);

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", "attachment; filename=Invoices_All.xlsx");

    const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: res });
    const sheet = workbook.addWorksheet("Invoices");

    const header = [
      "Doc Type",
      "Invoice ID",
      "Invoice Number",
      "Txn Date",
      "Due Date",
      "Created Time",
      "Updated Time",

      "Customer ID",
      "Customer Name",
      "Customer Email",
      "Customer GSTIN",

      "Terms",
      "Tax Mode",
      "Currency",
      "Exchange Rate",
      "PO Number",
      "Reference Number",
      "Private Note",

      "Discount Total",
      "Tax Total",
      "Total Amount",
      "Balance",

      "Bill Addr Line1",
      "Bill Addr Line2",
      "Bill City",
      "Bill State",
      "Bill Postal Code",
      "Bill Country",

      "Ship Addr Line1",
      "Ship Addr Line2",
      "Ship City",
      "Ship State",
      "Ship Postal Code",
      "Ship Country",

      "Line No",
      "Item ID",
      "Item Name",
      "Class ID",
      "Class Name",
      "Service Date",
      "Description",
      "Qty",
      "Rate",

      "Line Amount Raw",
      "Line Amount Excl Tax",
      "Tax Code",
      "Tax Rate",
      "Line Tax Amount",
      "Line Amount Incl Tax",

      "CGST %",
      "SGST %",
      "IGST %",
      "CGST Amount",
      "SGST Amount",
      "IGST Amount",
    ];
    sheet.addRow(header).commit();

    const getInvoiceDiscountTotal = (doc) => {
      const invoiceLines = Array.isArray(doc.Line) ? doc.Line : [];
      let discountTotal = 0;

      for (const l of invoiceLines) {
        const amt = Number(l.Amount || 0);
        if (l.DetailType === "DiscountLineDetail") {
          discountTotal += Math.abs(amt);
          continue;
        }
        const itemName = l?.SalesItemLineDetail?.ItemRef?.name || "";
        if (String(itemName).toLowerCase().includes("discount")) {
          discountTotal += Math.abs(amt);
        }
      }
      return +discountTotal.toFixed(2);
    };

    const pageSize = 500;
    let startPos = 1;

    while (true) {
      const { list } = await fetchDocsPage({
        entityName: "Invoice",
        realmId: QBO.realmId,
        token: QBO.access_token,
        whereClause: where,
        startPos,
        pageSize,
      });

      if (!list.length) break;

      for (const doc of list) {
        const billTo = doc.BillAddr || {};
        const shipTo = doc.ShipAddr || {};

        const lines = extractLineTaxes(doc, taxMaster);
        const discountTotal = getInvoiceDiscountTotal(doc);

        const taxMode = doc.GlobalTaxCalculation || "TaxExcluded";
        const refNumber = doc.CustomerMemo?.value || doc.DocNumber || "";
        const privateNote =
          doc?.PrivateNote && String(doc.PrivateNote).trim()
            ? String(doc.PrivateNote).trim()
            : "";

        for (const ln of lines) {
          const cgst =
            ln.taxBreakup.find((t) =>
              (t.name || "").toUpperCase().includes("CGST")
            )?.rate || 0;

          const sgst =
            ln.taxBreakup.find((t) =>
              (t.name || "").toUpperCase().includes("SGST")
            )?.rate || 0;

          const igst =
            ln.taxBreakup.find((t) =>
              (t.name || "").toUpperCase().includes("IGST")
            )?.rate || 0;

          const lineAmountRaw = ln.amount || 0;
          const lineTaxAmount = ln.taxAmount || 0;

          let lineAmountExclTax = lineAmountRaw;
          let lineAmountInclTax = lineAmountRaw;

          if (taxMode === "TaxInclusive") {
            lineAmountExclTax = +(lineAmountRaw - lineTaxAmount).toFixed(2);
            lineAmountInclTax = lineAmountRaw;
          } else {
            lineAmountExclTax = lineAmountRaw;
            lineAmountInclTax = +(lineAmountRaw + lineTaxAmount).toFixed(2);
          }

          const cgstAmt = +((lineAmountExclTax * cgst) / 100).toFixed(2);
          const sgstAmt = +((lineAmountExclTax * sgst) / 100).toFixed(2);
          const igstAmt = +((lineAmountExclTax * igst) / 100).toFixed(2);

          sheet
            .addRow([
              "Invoice",
              doc.Id || "",
              doc.DocNumber || "",
              doc.TxnDate || "",
              doc.DueDate || "",
              doc.MetaData?.CreateTime || "",
              doc.MetaData?.LastUpdatedTime || "",

              doc.CustomerRef?.value || "",
              doc.CustomerRef?.name || "",
              doc.BillEmail?.Address || "",
              "",

              doc.SalesTermRef?.name || doc.SalesTermRef?.value || "",
              taxMode,
              doc.CurrencyRef?.value || "",
              doc.ExchangeRate || "",
              doc.PONumber || "",
              refNumber,
              privateNote,

              discountTotal,
              doc.TxnTaxDetail?.TotalTax || 0,
              doc.TotalAmt || 0,
              doc.Balance || 0,

              billTo.Line1 || "",
              billTo.Line2 || "",
              billTo.City || "",
              billTo.CountrySubDivisionCode || "",
              billTo.PostalCode || "",
              billTo.Country || "",

              shipTo.Line1 || "",
              shipTo.Line2 || "",
              shipTo.City || "",
              shipTo.CountrySubDivisionCode || "",
              shipTo.PostalCode || "",
              shipTo.Country || "",

              ln.lineNo,
              ln.itemId,
              ln.itemName,
              ln.classId,
              ln.className,
              ln.serviceDate,
              ln.description,
              ln.qty,
              ln.rate,

              lineAmountRaw,
              lineAmountExclTax,
              ln.taxCode,
              ln.taxRate,
              lineTaxAmount,
              lineAmountInclTax,

              cgst,
              sgst,
              igst,
              cgstAmt,
              sgstAmt,
              igstAmt,
            ])
            .commit();
        }
      }

      if (list.length < pageSize) break;
      startPos += pageSize;
    }

    sheet.commit();
    await workbook.commit();
  } catch (err) {
    console.error("EXPORT INVOICES ERROR:", err.response?.data || err.message || err);
    res.status(500).json(err.response?.data || { error: err.message });
  }
});

// ======================================================================================
// ✅ EXPORT ESTIMATES (same as yours, non-stream ok)
// ======================================================================================
router.get("/export/estimates", async (req, res) => {
  try {
    const QBO = getConn("main");
    if (!QBO.access_token) return res.status(400).send("❌ Not connected to QBO.");

    const { fromDate = "", toDate = "" } = req.query;
    const where = buildDateWhere("TxnDate", fromDate, toDate);

    const taxMaster = await loadTaxMaster(QBO.realmId, QBO.access_token);
    const estimates = await fetchAllDocs({
      entityName: "Estimate",
      realmId: QBO.realmId,
      token: QBO.access_token,
      whereClause: where,
    });

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Estimates");

    sheet.addRow([
      "Doc Type","Doc ID","Doc Number","Txn Date","Expiry Date","Created Date","Last Updated Time",
      "Customer ID","Customer Name","Customer Email","Terms","Currency","Exchange Rate","PO Number",
      "Ref Number","Private Note","Sales Tax Total","Txn Total Amount",
      "Bill To Line1","Bill To Line2","Bill To City","Bill To State","Bill To PostalCode","Bill To Country",
      "Ship To Line1","Ship To Line2","Ship To City","Ship To State","Ship To PostalCode","Ship To Country",
      "Line No","Item ID","Item Name","Class ID","Class Name","Service Date","Line Description",
      "Qty","Rate","Line Amount","Line Tax Code","Line Tax %","Line Tax Amount",
      "CGST %","SGST %","IGST %","CGST Amt","SGST Amt","IGST Amt",
    ]);

    for (const doc of estimates) {
      let full = doc;
      const hasLinesInQuery = Array.isArray(full.Line) && full.Line.length > 0;

      if (!hasLinesInQuery) {
        try {
          const fetched = await fetchDoc("estimate", doc.Id, QBO.realmId, QBO.access_token);
          full = fetched.data.Estimate;
        } catch {}
      }

      const taxMode = full?.GlobalTaxCalculation || "TaxExcluded";
      const lines = extractLineTaxes(full, taxMaster);

      const billTo = full?.BillAddr || {};
      const shipTo = full?.ShipAddr || {};

      if (!lines || lines.length === 0) {
        sheet.addRow([
          "Estimate", full?.Id || "", full?.DocNumber || "", full?.TxnDate || "", full?.ExpirationDate || "",
          full?.MetaData?.CreateTime || "", full?.MetaData?.LastUpdatedTime || "",
          full?.CustomerRef?.value || "", full?.CustomerRef?.name || "", full?.BillEmail?.Address || "",
          full?.SalesTermRef?.name || full?.SalesTermRef?.value || "",
          full?.CurrencyRef?.value || "", full?.ExchangeRate || "", full?.PONumber || "",
          full?.CustomerMemo?.value || "", String(full?.PrivateNote || "").trim(),
          full?.TxnTaxDetail?.TotalTax || 0, full?.TotalAmt || 0,
          billTo.Line1 || "", billTo.Line2 || "", billTo.City || "", billTo.CountrySubDivisionCode || "",
          billTo.PostalCode || "", billTo.Country || "",
          shipTo.Line1 || "", shipTo.Line2 || "", shipTo.City || "", shipTo.CountrySubDivisionCode || "",
          shipTo.PostalCode || "", shipTo.Country || "",
          "", "", "", "", "", "", "", "", 0, 0, 0, "", 0, 0, 0, 0, 0, 0
        ]);
        continue;
      }

      for (const ln of lines) {
        const cgst = ln.taxBreakup?.find((t) => (t.name || "").toUpperCase().includes("CGST"))?.rate || 0;
        const sgst = ln.taxBreakup?.find((t) => (t.name || "").toUpperCase().includes("SGST"))?.rate || 0;
        const igst = ln.taxBreakup?.find((t) => (t.name || "").toUpperCase().includes("IGST"))?.rate || 0;

        const lineAmountRaw = ln.amount || 0;
        const lineTaxAmount = ln.taxAmount || 0;

        let lineAmountExclTax = lineAmountRaw;
        if (taxMode === "TaxInclusive") lineAmountExclTax = +(lineAmountRaw - lineTaxAmount).toFixed(2);

        const cgstAmt = +((lineAmountExclTax * cgst) / 100).toFixed(2);
        const sgstAmt = +((lineAmountExclTax * sgst) / 100).toFixed(2);
        const igstAmt = +((lineAmountExclTax * igst) / 100).toFixed(2);

        sheet.addRow([
          "Estimate", full?.Id || "", full?.DocNumber || "", full?.TxnDate || "", full?.ExpirationDate || "",
          full?.MetaData?.CreateTime || "", full?.MetaData?.LastUpdatedTime || "",
          full?.CustomerRef?.value || "", full?.CustomerRef?.name || "", full?.BillEmail?.Address || "",
          full?.SalesTermRef?.name || full?.SalesTermRef?.value || "",
          full?.CurrencyRef?.value || "", full?.ExchangeRate || "", full?.PONumber || "",
          full?.CustomerMemo?.value || "", String(full?.PrivateNote || "").trim(),
          full?.TxnTaxDetail?.TotalTax || 0, full?.TotalAmt || 0,
          billTo.Line1 || "", billTo.Line2 || "", billTo.City || "", billTo.CountrySubDivisionCode || "",
          billTo.PostalCode || "", billTo.Country || "",
          shipTo.Line1 || "", shipTo.Line2 || "", shipTo.City || "", shipTo.CountrySubDivisionCode || "",
          shipTo.PostalCode || "", shipTo.Country || "",
          ln.lineNo, ln.itemId, ln.itemName, ln.classId, ln.className, ln.serviceDate, ln.description,
          ln.qty, ln.rate, lineAmountRaw, ln.taxCode, ln.taxRate, lineTaxAmount,
          cgst, sgst, igst, cgstAmt, sgstAmt, igstAmt
        ]);
      }
    }

    res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition","attachment; filename=Estimates_All.xlsx");
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.log("EXPORT ESTIMATES ERROR:", err.response?.data || err.message || err);
    res.status(500).json(err.response?.data || { error: err.message });
  }
});

// ======================================================================================
// ✅ CREDIT MEMOS (same logic as yours)
// ======================================================================================
router.get("/export/creditmemos", async (req, res) => {
  try {
    const QBO = getConn("main");
    if (!QBO.access_token) return res.status(400).send("❌ Not connected to QBO.");

    const { fromDate = "", toDate = "" } = req.query;
    const where = buildDateWhere("TxnDate", fromDate, toDate);

    const taxMaster = await loadTaxMaster(QBO.realmId, QBO.access_token);
    const cms = await fetchAllDocs({
      entityName: "CreditMemo",
      realmId: QBO.realmId,
      token: QBO.access_token,
      whereClause: where,
    });

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("CreditMemos");

    sheet.addRow([
      "Doc Type","Doc ID","Doc Number","Txn Date","Created Date","Last Updated Time",
      "Customer ID","Customer Name","Customer Email","Terms","Currency","Exchange Rate",
      "Ref Number","Private Note","Discount","Sales Tax Total","Txn Total Amount","Balance",
      "Bill To Line1","Bill To Line2","Bill To City","Bill To State","Bill To PostalCode","Bill To Country",
      "Ship To Line1","Ship To Line2","Ship To City","Ship To State","Ship To PostalCode","Ship To Country",
      "Line No","Item ID","Item Name","Class ID","Class Name","Service Date","Line Description",
      "Qty","Rate","Line Amount","Line Tax Code","Line Tax %","Line Tax Amount",
      "CGST %","SGST %","IGST %","CGST Amt","SGST Amt","IGST Amt",
    ]);

    const getDocDiscountTotal = (doc) => {
      const docLines = Array.isArray(doc.Line) ? doc.Line : [];
      let discountTotal = 0;
      for (const l of docLines) {
        const amt = Number(l.Amount || 0);
        if (l.DetailType === "DiscountLineDetail") {
          discountTotal += Math.abs(amt);
          continue;
        }
        const itemName = l?.SalesItemLineDetail?.ItemRef?.name || "";
        if (String(itemName).toLowerCase().includes("discount")) discountTotal += Math.abs(amt);
      }
      return +discountTotal.toFixed(2);
    };

    for (const doc of cms) {
      const billTo = doc.BillAddr || {};
      const shipTo = doc.ShipAddr || {};
      const lines = extractLineTaxes(doc, taxMaster);
      const discountTotal = getDocDiscountTotal(doc);

      for (const ln of lines) {
        const cgst = ln.taxBreakup.find((t) => (t.name || "").toUpperCase().includes("CGST"))?.rate || 0;
        const sgst = ln.taxBreakup.find((t) => (t.name || "").toUpperCase().includes("SGST"))?.rate || 0;
        const igst = ln.taxBreakup.find((t) => (t.name || "").toUpperCase().includes("IGST"))?.rate || 0;

        const cgstAmt = +((Number(ln.amount || 0) * cgst) / 100).toFixed(2);
        const sgstAmt = +((Number(ln.amount || 0) * sgst) / 100).toFixed(2);
        const igstAmt = +((Number(ln.amount || 0) * igst) / 100).toFixed(2);

        sheet.addRow([
          "CreditMemo",
          doc.Id || "",
          doc.DocNumber || "",
          doc.TxnDate || "",
          doc.MetaData?.CreateTime || "",
          doc.MetaData?.LastUpdatedTime || "",
          doc.CustomerRef?.value || "",
          doc.CustomerRef?.name || "",
          doc.BillEmail?.Address || "",
          doc.SalesTermRef?.name || doc.SalesTermRef?.value || "",
          doc.CurrencyRef?.value || "",
          doc.ExchangeRate || "",
          doc.CustomerMemo?.value || "",
          String(doc.PrivateNote || "").trim(),
          discountTotal,
          doc.TxnTaxDetail?.TotalTax || 0,
          doc.TotalAmt || 0,
          doc.Balance || 0,
          billTo.Line1 || "",
          billTo.Line2 || "",
          billTo.City || "",
          billTo.CountrySubDivisionCode || "",
          billTo.PostalCode || "",
          billTo.Country || "",
          shipTo.Line1 || "",
          shipTo.Line2 || "",
          shipTo.City || "",
          shipTo.CountrySubDivisionCode || "",
          shipTo.PostalCode || "",
          shipTo.Country || "",
          ln.lineNo,
          ln.itemId,
          ln.itemName,
          ln.classId,
          ln.className,
          ln.serviceDate,
          ln.description,
          ln.qty,
          ln.rate,
          ln.amount,
          ln.taxCode,
          ln.taxRate,
          ln.taxAmount,
          cgst,
          sgst,
          igst,
          cgstAmt,
          sgstAmt,
          igstAmt,
        ]);
      }
    }

    res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition","attachment; filename=CreditMemos_All.xlsx");
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.log("EXPORT CREDITMEMOS ERROR:", err.response?.data || err.message || err);
    res.status(500).json(err.response?.data || { error: err.message });
  }
});


router.post("/export/allocation/excel", async (req, res) => {
  try {
    if (!QBO.access_token || !QBO.realmId) {
      return res.status(400).send("❌ Not connected to QuickBooks");
    }

    const { fromDate, toDate, filterBy = "invoice" } = req.body || {};

    // If filterBy=invoice => filter CM/VC/Invoice by TxnDate
    // If filterBy=payment => filter Payment/BillPayment by TxnDate
    const whereTxn = buildDateWhere("TxnDate", fromDate, toDate);

    const cmWhere = filterBy === "invoice" ? whereTxn : "";
    const vcWhere = filterBy === "invoice" ? whereTxn : "";
    const invWhere = filterBy === "invoice" ? whereTxn : "";
    const payWhere = filterBy === "payment" ? whereTxn : "";
    const billPayWhere = filterBy === "payment" ? whereTxn : "";

    const [creditMemos, vendorCredits, payments, billPayments, invoices] =
      await Promise.all([
        fetchAllDocs({
          entityName: "CreditMemo",
          realmId: QBO.realmId,
          token: QBO.access_token,
          whereClause: cmWhere,
        }),
        fetchAllDocs({
          entityName: "VendorCredit",
          realmId: QBO.realmId,
          token: QBO.access_token,
          whereClause: vcWhere,
        }),
        fetchAllDocs({
          entityName: "Payment",
          realmId: QBO.realmId,
          token: QBO.access_token,
          whereClause: payWhere,
        }),
        fetchAllDocs({
          entityName: "BillPayment",
          realmId: QBO.realmId,
          token: QBO.access_token,
          whereClause: billPayWhere,
        }),
        fetchAllDocs({
          entityName: "Invoice",
          realmId: QBO.realmId,
          token: QBO.access_token,
          whereClause: invWhere,
        }),
      ]);

    // ✅ InvoiceId -> { number, total }  (Invoice Number + Invoice Amount)
    const invoiceMap = new Map();
    invoices.forEach((inv) =>
      invoiceMap.set(String(inv.Id), {
        number: inv.DocNumber || "",
        total: inv.TotalAmt || 0,
      })
    );

    // Buckets
    const cmAlloc = new Map();
    const vcAlloc = new Map();

    creditMemos.forEach((cm) => cmAlloc.set(String(cm.Id), { doc: cm, allocs: [] }));
    vendorCredits.forEach((vc) => vcAlloc.set(String(vc.Id), { doc: vc, allocs: [] }));

    const norm = (t) => (t || "").toLowerCase();

    // 🔹 Helper: comma / array invoice IDs ko alag-alag rows me todne ke liye
    const explodeIds = (val) => {
      if (!val) return [""];
      if (Array.isArray(val)) return val.length ? val.map(String) : [""];
      const s = String(val);
      if (s.includes(",")) {
        return s.split(",").map((x) => x.trim()).filter(Boolean);
      }
      return [s.trim()];
    };

    // ----------------------------------------------------------------------------------
    // 1) PAYMENT -> CREDIT MEMO (Invoice IDs collected at PAYMENT level)
    // ----------------------------------------------------------------------------------
    payments.forEach((p) => {
      const lines = p.Line || [];

      // ✅ Collect ALL invoice ids & ALL creditmemo ids from the whole Payment
      const allInvoiceIds = [];
      const allCreditMemoIds = [];

      lines.forEach((line) => {
        const linked = line.LinkedTxn || [];
        linked.forEach((ln) => {
          const t = norm(ln.TxnType);
          if (t === "invoice" && ln.TxnId) allInvoiceIds.push(String(ln.TxnId));
          if (t === "creditmemo" && ln.TxnId) allCreditMemoIds.push(String(ln.TxnId));
        });
      });

      if (allCreditMemoIds.length === 0) return;

      // unique
      const uniqInvoiceIds = [...new Set(allInvoiceIds)];
      const uniqCreditMemoIds = [...new Set(allCreditMemoIds)];

      const refNumber =
        p.PaymentRefNum ||
        p.DocNumber ||
        p.TxnNumber ||
        p.ReferenceNumber ||
        "";

      const invCount = uniqInvoiceIds.length || 1;

      // Payment total: QBO me TotalAmt hota hai
      const paymentTotal = Number(p.TotalAmt ?? 0);
      const amountPerInvoice = +(paymentTotal / invCount).toFixed(2);

      uniqCreditMemoIds.forEach((cmId) => {
        const bucket = cmAlloc.get(String(cmId));
        if (!bucket) return;

        bucket.allocs.push({
          type: "Payment",
          sourceId: p.Id,
          date: p.TxnDate || "",
          amount: uniqInvoiceIds.length ? amountPerInvoice : paymentTotal,
          refNumber,
          appliedInvoiceIds: uniqInvoiceIds.length ? uniqInvoiceIds : [""], // ✅ array
        });
      });
    });

    // ----------------------------------------------------------------------------------
    // 2) BILL PAYMENT -> VENDOR CREDIT (ADD Bill ID + Bill Number)
    // ----------------------------------------------------------------------------------
    billPayments.forEach((bp) => {
      const refNumber =
        bp.CheckPayment?.CheckNumber ||
        bp.PaymentRefNum ||
        bp.DocNumber ||
        bp.TxnNumber ||
        bp.ReferenceNumber ||
        "";

      const billId = bp.Id; // ✅ BillPayment ID
      const billNumber = bp.DocNumber || ""; // ✅ BillPayment DocNumber (Bill No in your requirement)

      (bp.Line || []).forEach((line) => {
        (line.LinkedTxn || []).forEach((ln) => {
          if (norm(ln.TxnType) === "vendorcredit") {
            const bucket = vcAlloc.get(String(ln.TxnId));
            if (!bucket) return;

            const amount = Number(ln.Amount ?? line.Amount ?? 0);

            bucket.allocs.push({
              type: "BillPayment",
              sourceId: bp.Id,
              date: bp.TxnDate || "",
              amount,
              refNumber,
              billId, // ✅ NEW
              billNumber, // ✅ NEW
            });
          }
        });
      });
    });

    // ----------------------------------------------------------------------------------
    // 3) Excel
    // ----------------------------------------------------------------------------------
    const workbook = new ExcelJS.Workbook();

    // ===========================
    // CREDIT MEMO SHEET
    // ===========================
    const cmSheet = workbook.addWorksheet("CreditMemoAllocation");
    cmSheet.addRow([
      "CreditMemo ID",
      "CreditMemo Number",
      "CreditMemo Date",
      "Customer ID",
      "Customer Name",
      "Currency",
      "CreditMemo Total",
      "Total Allocated",
      "Remaining Balance",
      "Applied Invoice ID",
      "Applied Invoice Number",
      "Applied Invoice Amount", // ✅ NEW
      "Alloc Type",
      "Alloc Source ID",
      "Alloc Date",
      "Alloc Amount",
      "Alloc Ref Number",
    ]);

    cmAlloc.forEach(({ doc, allocs }, id) => {
      const cmTotal = Number(doc.TotalAmt || 0);
      const customerId = doc.CustomerRef?.value || "";
      const customerName = doc.CustomerRef?.name || "";
      const currency = doc.CurrencyRef?.value || "";
      const docNumber = doc.DocNumber || "";
      const date = doc.TxnDate || "";

      // allocated = sum(alloc amount * count(invoiceIds)) (because we explode rows)
      const allocated = allocs.reduce((s, a) => {
        const invCount = explodeIds(a.appliedInvoiceIds || a.appliedInvoiceId).length;
        return s + Number(a.amount || 0) * invCount;
      }, 0);

      const remaining = +(cmTotal - allocated).toFixed(2);

      if (allocs.length === 0) {
        cmSheet.addRow([
          id,
          docNumber,
          date,
          customerId,
          customerName,
          currency,
          cmTotal,
          0,
          cmTotal,
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
        ]);
        return;
      }

      allocs.forEach((a) => {
        const invIds = explodeIds(a.appliedInvoiceIds || a.appliedInvoiceId);

        invIds.forEach((invId) => {
          const invData = invId ? invoiceMap.get(String(invId)) || {} : {};
          const appliedInvoiceNumber = invData.number || "";
          const appliedInvoiceAmount = invData.total ?? "";

          cmSheet.addRow([
            id,
            docNumber,
            date,
            customerId,
            customerName,
            currency,
            cmTotal,
            +allocated.toFixed(2),
            remaining,
            invId,
            appliedInvoiceNumber,
            appliedInvoiceAmount, // ✅ NEW
            a.type,
            a.sourceId,
            a.date,
            a.amount, // per invoice amount
            a.refNumber,
          ]);
        });
      });
    });

    // ===========================
    // VENDOR CREDIT SHEET
    // ===========================
    const vcSheet = workbook.addWorksheet("VendorCreditAllocation");
    vcSheet.addRow([
      "VendorCredit ID",
      "VendorCredit Number",
      "VendorCredit Date",
      "Vendor ID",
      "Vendor Name",
      "Currency",
      "VendorCredit Total",
      "Total Allocated",
      "Remaining Balance",
      "Applied Bill ID", // ✅ NEW
      "Applied Bill Number", // ✅ NEW
      "Alloc Type",
      "Alloc Source ID",
      "Alloc Date",
      "Alloc Amount",
      "Alloc Ref Number",
    ]);

    vcAlloc.forEach(({ doc, allocs }, id) => {
      const vcTotal = Number(doc.TotalAmt || 0);
      const vendorId = doc.VendorRef?.value || "";
      const vendorName = doc.VendorRef?.name || "";
      const currency = doc.CurrencyRef?.value || "";
      const docNumber = doc.DocNumber || "";
      const date = doc.TxnDate || "";

      const allocated = allocs.reduce((s, a) => s + Number(a.amount || 0), 0);
      const remaining = +(vcTotal - allocated).toFixed(2);

      if (allocs.length === 0) {
        vcSheet.addRow([
          id,
          docNumber,
          date,
          vendorId,
          vendorName,
          currency,
          vcTotal,
          0,
          vcTotal,
          "",
          "",
          "",
          "",
          "",
          "",
          "",
        ]);
      } else {
        allocs.forEach((a) => {
          vcSheet.addRow([
            id,
            docNumber,
            date,
            vendorId,
            vendorName,
            currency,
            vcTotal,
            +allocated.toFixed(2),
            remaining,
            a.billId || "", // ✅ Bill ID
            a.billNumber || "", // ✅ Bill No
            a.type,
            a.sourceId,
            a.date,
            a.amount,
            a.refNumber,
          ]);
        });
      }
    });

    // ----------------------------------------------------------------------------------
    // SEND FILE
    // ----------------------------------------------------------------------------------
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=QBO-CreditMemo-VendorCredit-Allocation.xlsx"
    );

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.log("ALLOCATION EXPORT ERROR:", err.response?.data || err.message);
    res.status(500).json(err.response?.data || { error: err.message });
  }
});



// ======================================================================================

// ======================================================================================
// ✅ FIX: Deposit To blank issue + Correct fields for QBO Payment
// Reason: Payment object usually has DepositToAccountRef (not DepositToAccountRef.name always),
// and sometimes only .value comes. So we must map AccountId -> AccountName
// ======================================================================================
router.get("/export/overpayments", async (req, res) => {
  try {
    if (!QBO.access_token || !QBO.realmId) {
      return res.status(400).send("❌ Not connected to QuickBooks");
    }

    const { fromDate = "", toDate = "" } = req.query;
    const whereTxn = buildDateWhere("TxnDate", fromDate, toDate);

    // ✅ Fetch Bank Accounts for DepositTo mapping
    const bankAccounts = await fetchAllDocs({
      entityName: "Account",
      realmId: QBO.realmId,
      token: QBO.access_token,
     whereClause: "AccountType IN ('Bank','Other Current Asset')",
    });

    const accountMap = new Map();
    bankAccounts.forEach((a) => {
      accountMap.set(String(a.Id), {
        name: a.Name || "",
        code: a.AcctNum || "",
        full: `${a.AcctNum ? a.AcctNum + " " : ""}${a.Name || ""}`.trim(),
      });
    });

    const [payments, billPayments] = await Promise.all([
      fetchAllDocs({
        entityName: "Payment",
        realmId: QBO.realmId,
        token: QBO.access_token,
        whereClause: whereTxn,
      }),
      fetchAllDocs({
        entityName: "BillPayment",
        realmId: QBO.realmId,
        token: QBO.access_token,
        whereClause: whereTxn,
      }),
    ]);

    const norm = (t) => String(t || "").toLowerCase();

    const chunk = (arr, size) => {
      const out = [];
      for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
      return out;
    };

    const fetchByIds = async (entityName, ids) => {
      const uniq = [...new Set(ids.filter(Boolean).map(String))];
      if (!uniq.length) return [];
      const parts = chunk(uniq, 30);

      const all = [];
      for (const idsChunk of parts) {
        const inClause = idsChunk.map((x) => `'${x}'`).join(", ");
        const whereClause = `Id IN (${inClause})`;
        const rows = await fetchAllDocs({
          entityName,
          realmId: QBO.realmId,
          token: QBO.access_token,
          whereClause,
        });
        all.push(...(rows || []));
      }
      return all;
    };

    const getLinkedTxnsFromPayment = (p) => {
      const out = [];
      (p.Line || []).forEach((line) => {
        (line.LinkedTxn || []).forEach((ln) => {
          out.push({
            txnType: ln.TxnType || "",
            txnId: ln.TxnId || "",
            amount: Number(ln.Amount ?? line.Amount ?? 0),
          });
        });
      });
      (p.LinkedTxn || []).forEach((ln) => {
        out.push({
          txnType: ln.TxnType || "",
          txnId: ln.TxnId || "",
          amount: Number(ln.Amount ?? 0),
        });
      });
      return out;
    };

    const getLinkedTxnsFromBillPayment = (bp) => {
      const out = [];
      (bp.Line || []).forEach((line) => {
        (line.LinkedTxn || []).forEach((ln) => {
          out.push({
            txnType: ln.TxnType || "",
            txnId: ln.TxnId || "",
            amount: Number(ln.Amount ?? line.Amount ?? 0),
          });
        });
      });
      (bp.LinkedTxn || []).forEach((ln) => {
        out.push({
          txnType: ln.TxnType || "",
          txnId: ln.TxnId || "",
          amount: Number(ln.Amount ?? 0),
        });
      });
      return out;
    };

    const paymentLinked = payments.map((p) => ({
      p,
      linked: getLinkedTxnsFromPayment(p),
    }));

    const billPaymentLinked = billPayments.map((bp) => ({
      bp,
      linked: getLinkedTxnsFromBillPayment(bp),
    }));

    // Collect linked IDs (customer side)
    const invoiceIds = [];
    const creditMemoIds = [];
    const depositIds = [];
    paymentLinked.forEach(({ linked }) => {
      linked.forEach((ln) => {
        const t = norm(ln.txnType);
        if (t === "invoice") invoiceIds.push(ln.txnId);
        else if (t === "creditmemo") creditMemoIds.push(ln.txnId);
        else if (t === "deposit") depositIds.push(ln.txnId);
      });
    });

    // Collect linked IDs (vendor side)
    const billIds = [];
    const vendorCreditIds = [];
    billPaymentLinked.forEach(({ linked }) => {
      linked.forEach((ln) => {
        const t = norm(ln.txnType);
        if (t === "bill") billIds.push(ln.txnId);
        else if (t === "vendorcredit") vendorCreditIds.push(ln.txnId);
      });
    });

    const [invoices, creditMemos, deposits, bills, vendorCredits] = await Promise.all([
      fetchByIds("Invoice", invoiceIds),
      fetchByIds("CreditMemo", creditMemoIds),
      (async () => {
        try {
          return await fetchByIds("Deposit", depositIds);
        } catch {
          return [];
        }
      })(),
      fetchByIds("Bill", billIds),
      fetchByIds("VendorCredit", vendorCreditIds),
    ]);

    const buildTxnMap = (docs, type) => {
      const m = new Map();
      (docs || []).forEach((d) => {
        const id = String(d.Id || "");
        m.set(id, {
          id,
          type,
          date: d.TxnDate || "",
          no: d.DocNumber || "",
          dueDate: d.DueDate || "",
          amount: Number(d.TotalAmt ?? 0),
          openBalance: Number(d.Balance ?? d.RemainingCredit ?? d.TotalAmt ?? 0),
        });
      });
      return m;
    };

    const invoiceMap = buildTxnMap(invoices, "Invoice");
    const creditMemoMap = buildTxnMap(creditMemos, "CreditMemo");
    const depositMap = buildTxnMap(deposits, "Deposit");
    const billMap = buildTxnMap(bills, "Bill");
    const vendorCreditMap = buildTxnMap(vendorCredits, "VendorCredit");

    const getTxnInfo = (txnType, txnId) => {
      const t = norm(txnType);
      const id = String(txnId || "");
      if (t === "invoice") return invoiceMap.get(id);
      if (t === "creditmemo") return creditMemoMap.get(id);
      if (t === "deposit") return depositMap.get(id);
      if (t === "bill") return billMap.get(id);
      if (t === "vendorcredit") return vendorCreditMap.get(id);
      return null;
    };

    // ✅ Extract Deposit To safely
    const getDepositToLabel = (p) => {
      const ref =
        p.DepositToAccountRef || // ✅ normal
        p.DepositToRef || // sometimes
        p.DepositTo || // rare
        null;

      const id = ref?.value ? String(ref.value) : "";
      const directName = ref?.name || "";

      if (directName) return directName;
      if (id && accountMap.has(id)) return accountMap.get(id).full || accountMap.get(id).name;
      return ""; // if still blank
    };

    const workbook = new ExcelJS.Workbook();

    // =========================================================
    // SHEET 1: CustomerPaymentApplyLines
    // =========================================================
    const s1 = workbook.addWorksheet("CustomerPaymentApplyLines");
    s1.addRow([
      "Payment ID",
      "Payment Date",
      "Customer",
      "Amount Received",
      "Payment Ref No",
      "Deposit To", // ✅ should be filled now
      "Linked Txn Type",
      "Linked Txn ID",
      "Txn Date",
      "No.",
      "Due Date",
      "Amount",
      "Open Balance",
      "Applied (Payment Column)",
    ]);

    const customerOverpayRows = [];

    paymentLinked.forEach(({ p, linked }) => {
      const paymentTotal = Number(p.TotalAmt ?? 0);
      const customerName = p.CustomerRef?.name || "";
      const paymentRef =
        p.PaymentRefNum || p.DocNumber || p.TxnNumber || p.ReferenceNumber || "";

      const depositTo = getDepositToLabel(p); // ✅ FIXED

      const appliedTotal = linked.reduce((s, ln) => s + Number(ln.amount || 0), 0);
      const overpay = +(paymentTotal - appliedTotal).toFixed(2);

      if (!linked.length) {
        s1.addRow([
          p.Id || "",
          p.TxnDate || "",
          customerName,
          paymentTotal,
          paymentRef,
          depositTo,
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
        ]);
      }

      linked.forEach((ln) => {
        const info = getTxnInfo(ln.txnType, ln.txnId) || {};
        s1.addRow([
          p.Id || "",
          p.TxnDate || "",
          customerName,
          paymentTotal,
          paymentRef,
          depositTo,
          ln.txnType || "",
          ln.txnId || "",
          info.date || "",
          info.no || "",
          info.dueDate || "",
          info.amount ?? "",
          info.openBalance ?? "",
          Number(ln.amount || 0),
        ]);
      });

      if (overpay > 0) {
        customerOverpayRows.push({
          type: "Payment",
          paymentId: p.Id || "",
          paymentDate: p.TxnDate || "",
          customer: customerName,
          amountReceived: paymentTotal,
          appliedTotal: +appliedTotal.toFixed(2),
          unapplied: overpay,
          paymentRef,
          depositTo, // ✅ should be filled now
          email:
            p.BillEmail?.Address ||
            p.PrimaryEmailAddr?.Address ||
            p.CustomerRef?.email ||
            "",
          currency: p.CurrencyRef?.value || "",
          exchangeRate: p.ExchangeRate || 1,
        });
      }
    });

    // =========================================================
    // SHEET 2: CustomerOverpaymentSummary (this is your screenshot sheet)
    // =========================================================
    const s2 = workbook.addWorksheet("CustomerOverpaymentSummary");
    s2.addRow([
      "Type",
      "Payment ID",
      "Payment Date",
      "Customer",
      "Amount Received",
      "Applied Total",
      "Unapplied",
      "Payment Ref",
      "Deposit To",     // ✅ FIXED
      "Email",
      "Currency",
      "Exchange Rate",
    ]);

    customerOverpayRows.forEach((r) => {
      s2.addRow([
        r.type,
        r.paymentId,
        r.paymentDate,
        r.customer,
        r.amountReceived,
        r.appliedTotal,
        r.unapplied,
        r.paymentRef,
        r.depositTo,     // ✅ should show like "120025 aaNAB ... - AUD"
        r.email,
        r.currency,
        r.exchangeRate,
      ]);
    });

    // =========================================================
    // SHEET 3: VendorBillPaymentApplyLines
    // =========================================================
    const s3 = workbook.addWorksheet("VendorBillPaymentApplyLines");
    s3.addRow([
      "BillPayment ID",
      "Payment Date",
      "Vendor",
      "Payment Total",
      "Ref No / Check No",
      "Bank Account",
      "Linked Txn Type",
      "Linked Txn ID",
      "Txn Date",
      "No.",
      "Due Date",
      "Amount",
      "Open Balance",
      "Applied",
    ]);

    const vendorOverpayRows = [];

    billPaymentLinked.forEach(({ bp, linked }) => {
      const paymentTotal = Number(bp.TotalAmt ?? 0);
      const vendorName = bp.VendorRef?.name || "";
      const refNo =
        bp.CheckPayment?.CheckNumber ||
        bp.PaymentRefNum ||
        bp.DocNumber ||
        bp.TxnNumber ||
        bp.ReferenceNumber ||
        "";

      const bankAccount =
        bp.CheckPayment?.BankAccountRef?.name ||
        bp.CreditCardPayment?.CCAccountRef?.name ||
        "";

      const appliedTotal = linked.reduce((s, ln) => s + Number(ln.amount || 0), 0);
      const overpay = +(paymentTotal - appliedTotal).toFixed(2);

      if (!linked.length) {
        s3.addRow([
          bp.Id || "",
          bp.TxnDate || "",
          vendorName,
          paymentTotal,
          refNo,
          bankAccount,
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
        ]);
      }

      linked.forEach((ln) => {
        const info = getTxnInfo(ln.txnType, ln.txnId) || {};
        s3.addRow([
          bp.Id || "",
          bp.TxnDate || "",
          vendorName,
          paymentTotal,
          refNo,
          bankAccount,
          ln.txnType || "",
          ln.txnId || "",
          info.date || "",
          info.no || "",
          info.dueDate || "",
          info.amount ?? "",
          info.openBalance ?? "",
          Number(ln.amount || 0),
        ]);
      });

      if (overpay > 0) {
        vendorOverpayRows.push([
          "BillPayment",
          bp.Id || "",
          bp.TxnDate || "",
          vendorName,
          paymentTotal,
          +appliedTotal.toFixed(2),
          overpay,
          refNo,
          bankAccount,
        ]);
      }
    });

    // =========================================================
    // SHEET 4: VendorOverpaymentSummary
    // =========================================================
    const s4 = workbook.addWorksheet("VendorOverpaymentSummary");
    s4.addRow([
      "Type",
      "BillPayment ID",
      "Payment Date",
      "Vendor",
      "Payment Total",
      "Applied Total",
      "Unapplied",
      "Ref No / Check No",
      "Bank Account",
    ]);
    vendorOverpayRows.forEach((r) => s4.addRow(r));

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=QBO-UI-Overpayments.xlsx"
    );

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.log("OVERPAYMENT EXPORT ERROR:", err.response?.data || err.message);
    res.status(500).json(err.response?.data || { error: err.message });
  }
});

export default router;
