// extractTax.js
import axios from "axios";

// -------------------------------
// Load Tax Codes + Tax Rates
// -------------------------------
export async function loadTaxMaster(realmId, token) {
  const q = (query) =>
    axios.get(
      `https://quickbooks.api.intuit.com/v3/company/${realmId}/query?query=${encodeURIComponent(
        query
      )}&minorversion=75`,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json",
        },
      }
    );

  const [codes, rates] = await Promise.all([
    q("SELECT * FROM TaxCode"),
    q("SELECT * FROM TaxRate"),
  ]);

  const taxCodes = codes.data?.QueryResponse?.TaxCode || [];
  const taxRates = rates.data?.QueryResponse?.TaxRate || [];

  return { taxCodes, taxRates };
}

// -------------------------------
// Extract Line-Level Taxes
// Returns array of:
// {
//   lineNo,
//   lineIndex,
//   description,
//   qty,
//   rate,
//   amount,
//   taxCode,
//   taxRate,
//   taxAmount,
//   taxBreakup: [{ rate, name }],
//   itemId,
//   itemName,
//   classId,
//   className,
//   serviceDate
// }
// -------------------------------
export function extractLineTaxes(doc, taxMaster) {
  if (!doc || !doc.Line || !Array.isArray(doc.Line)) return [];

  // Simple approach: doc level tax percent
  const taxPercent =
    doc?.TxnTaxDetail?.TaxLine?.[0]?.TaxLineDetail?.TaxPercent || 0;

  const result = [];
  let lineCounter = 1;

  doc.Line.forEach((line, idx) => {
    if (line.DetailType !== "SalesItemLineDetail") return;

    const detail = line.SalesItemLineDetail || {};
    const qty = detail.Qty ?? null;
    const rate = detail.UnitPrice ?? null;
    const amount = line.Amount || 0;

    const taxAmount = +(amount * (taxPercent / 100)).toFixed(2);
    const taxCodeRef = detail?.TaxCodeRef?.value || null;

    let taxBreakup = [];
    if (taxMaster?.taxCodes && taxMaster?.taxRates) {
      const code = taxMaster.taxCodes.find(
        (c) => c.Id == taxCodeRef || c.Name == taxCodeRef
      );

      if (code?.SalesTaxRateList?.TaxRateDetail) {
        taxBreakup = code.SalesTaxRateList.TaxRateDetail.map((d) => {
          const rateObj = taxMaster.taxRates.find(
            (r) => r.Id == d.TaxRateRef.value
          );
          return rateObj
            ? { rate: rateObj.RateValue, name: rateObj.Name }
            : null;
        }).filter(Boolean);
      }
    }

    const itemRef = detail.ItemRef || {};
    const classRef = detail.ClassRef || {};

    result.push({
      lineNo: lineCounter++,
      lineIndex: idx,
      description: line.Description || "",
      qty,
      rate,
      amount,
      taxCode: taxCodeRef,
      taxRate: taxPercent,
      taxAmount,
      taxBreakup:
        taxBreakup.length > 0
          ? taxBreakup
          : [{ rate: taxPercent, name: "GST" }],
      itemId: itemRef.value || "",
      itemName: itemRef.name || "",
      classId: classRef.value || "",
      className: classRef.name || "",
      serviceDate: detail.ServiceDate || "",
    });
  });

  return result;
}
