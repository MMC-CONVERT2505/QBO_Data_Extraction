// attachments.js
import express from "express";
import axios from "axios";
import FormData from "form-data";

const router = express.Router();

/* --------------------------------------------------
   FROM / TO helpers – tokens quickbooks.js ke global me store honge
-------------------------------------------------- */

function getFromConnection() {
  return global.QBO_FROM || {
    access_token: null,
    refresh_token: null,
    realmId: null,
  };
}

function getToConnection() {
  return global.QBO_TO || {
    access_token: null,
    refresh_token: null,
    realmId: null,
  };
}

/* --------------------------------------------------
   Generic helpers: rate-limit retry + QBO queries
-------------------------------------------------- */

// QBO query with retry (rate limit / server error)
async function qboQueryWithRetry({ realmId, token, query }, maxRetries = 3) {
  const url = `https://quickbooks.api.intuit.com/v3/company/${realmId}/query?query=${encodeURIComponent(
    query
  )}&minorversion=75`;

  let attempt = 0;
  let lastError = null;

  while (attempt <= maxRetries) {
    try {
      const r = await axios.get(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json",
        },
      });
      return r.data.QueryResponse || {};
    } catch (err) {
      lastError = err;
      const status = err.response?.status;

      if (status === 429 || status === 500 || status === 503) {
        const waitMs = 1000 * Math.pow(2, attempt); // 1s,2s,4s...
        console.log(
          `⏳ QBO retry (status=${status}, attempt=${
            attempt + 1
          }) after ${waitMs / 1000}s`
        );
        await new Promise((resolve) => setTimeout(resolve, waitMs));
        attempt++;
      } else {
        throw err;
      }
    }
  }

  throw lastError;
}

// All Attachable for given entity type (Invoice, CreditMemo, Bill, ...)
async function fetchAttachablesByType(connection, entityType) {
  const { realmId, access_token } = connection;
  const pageSize = 100;
  let startPos = 1;
  let all = [];

  while (true) {
    let query = `SELECT * FROM Attachable WHERE AttachableRef.EntityRef.type = '${entityType}'`;
    query += ` STARTPOSITION ${startPos} MAXRESULTS ${pageSize}`;

    const resp = await qboQueryWithRetry(
      {
        realmId,
        token: access_token,
        query,
      },
      3
    );

    const list = resp.Attachable || [];
    if (!list.length) break;

    all = all.concat(list);
    if (list.length < pageSize) break;
    startPos += pageSize;
  }

  return all;
}

/* --------------------------------------------------
   DOC FETCH + MATCHING HELPERS
   Abhi match DocNumber se, baad me yahi jagah pe
   extra conditions add kar sakte hain.
-------------------------------------------------- */

// QBO transaction endpoint name map
const ENTITY_ENDPOINT_MAP = {
  Invoice: "invoice",
  CreditMemo: "creditmemo",
  Bill: "bill",
  VendorCredit: "vendorcredit",
  SalesReceipt: "salesreceipt",
  Estimate: "estimate",
  CreditCardCharge: "creditcardcharge",
  Purchase: "purchase",
  Check: "check",
  DelayedCharge: "delayedcharge",
  JournalEntry: "journalentry",
  Payment: "payment",
  RefundReceipt: "refundreceipt",
};

// Source document fetch by ID
async function fetchSourceDocById(connection, entityType, id) {
  const endpoint = ENTITY_ENDPOINT_MAP[entityType];
  if (!endpoint) return null;

  const url = `https://quickbooks.api.intuit.com/v3/company/${connection.realmId}/${endpoint}/${id}?minorversion=75`;

  const r = await axios.get(url, {
    headers: {
      Authorization: `Bearer ${connection.access_token}`,
      Accept: "application/json",
    },
  });

  const rootKey = entityType; // e.g. Invoice, Bill
  return r.data[rootKey] || null;
}

// DocNumber निकालने का एक ही function
// 👉 future me yahi pe extra fields add kar sakte ho (e.g. date, customer name)
function getDocNumberForMatch(entityType, doc) {
  return (
    doc.DocNumber ||
    doc.TxnNumber ||
    doc.RefNumber ||
    doc.PaymentRefNum ||
    ""
  );
}

// Target company me DocNumber se matching transaction
async function findTargetByDocNumber(connection, entityType, docNumber) {
  if (!docNumber) return null;

  const escaped = docNumber.replace(/'/g, "''"); // SQL escape '
  const query = `SELECT * FROM ${entityType} WHERE DocNumber = '${escaped}'`;

  const resp = await qboQueryWithRetry(
    {
      realmId: connection.realmId,
      token: connection.access_token,
      query,
    },
    3
  );

  const list = resp[entityType] || [];
  if (!list.length) return null;

  // Assuming DocNumber unique – pehla hi le lo
  return list[0];
}

/* --------------------------------------------------
   FILE DOWNLOAD + UPLOAD HELPERS
-------------------------------------------------- */

// FROM company se attachment file download
async function downloadAttachmentFile(FROM, attachable) {
  const url = attachable.FileAccessUri || attachable.TempDownloadUri;
  if (!url) return null;

  const r = await axios.get(url, {
    responseType: "arraybuffer",
    headers: {
      Authorization: `Bearer ${FROM.access_token}`,
      Accept: "*/*",
    },
  });

  const buffer = Buffer.from(r.data);
  const fileName = attachable.FileName || "attachment.bin";

  return { buffer, fileName };
}

// TO company pe upload + attach
async function uploadAttachmentToTarget(TO, entityType, targetId, file, note) {
  const url = `https://quickbooks.api.intuit.com/v3/company/${TO.realmId}/upload?minorversion=75`;

  const metadata = {
    AttachableRef: [
      {
        EntityRef: {
          type: entityType,
          value: targetId,
        },
      },
    ],
    FileName: file.fileName,
    Note: note || "",
    Category: "Document",
  };

  const form = new FormData();
  form.append("file_metadata_01", JSON.stringify(metadata), {
    contentType: "application/json",
  });
  form.append("file_content_01", file.buffer, {
    filename: file.fileName,
  });

  const headers = {
    ...form.getHeaders(),
    Authorization: `Bearer ${TO.access_token}`,
    Accept: "application/json",
  };

  const r = await axios.post(url, form, { headers });
  return r.data;
}

/* --------------------------------------------------
   1. ATTACHMENT SCAN (already working)
-------------------------------------------------- */

// POST /attachments/sync
// Body: { docTypes: ["Invoice","CreditMemo", ...] }
router.post("/attachments/sync", async (req, res) => {
  try {
    const FROM = getFromConnection();
    const TO = getToConnection();

    if (!FROM?.access_token || !FROM?.realmId) {
      return res.status(400).json({
        error:
          "FROM company not connected. Please connect using /qbo/auth-from first.",
      });
    }
    if (!TO?.access_token || !TO?.realmId) {
      return res.status(400).json({
        error:
          "TO company not connected. Please connect using /qbo/auth-to first.",
      });
    }

    const {
      docTypes = [
        "Invoice",
        "CreditMemo",
        "Bill",
        "VendorCredit",
        "SalesReceipt",
        "Estimate",
        "CreditCardCharge",
        "Purchase",
        "Check",
        "DelayedCharge",
        "JournalEntry",
        "Payment",
        "RefundReceipt",
      ],
    } = req.body || {};

    const stats = {};
    const errors = [];

    for (const type of docTypes) {
      try {
        console.log(`📎 Scanning attachments for type: ${type}`);
        const attachables = await fetchAttachablesByType(FROM, type);

        const total = attachables.length;
        const withFileUri = attachables.filter(
          (a) => a.FileAccessUri || a.TempDownloadUri
        ).length;

        stats[type] = {
          totalAttachables: total,
          withFileUri,
        };
      } catch (err) {
        console.log(
          `Attachment scan error for ${type}:`,
          err.response?.data || err.message
        );
        errors.push(
          `Type ${type}: ${
            err.response?.data?.Fault?.Error?.[0]?.Message || err.message
          }`
        );
      }
    }

    return res.json({
      success: true,
      message:
        "Attachment scan completed for FROM company. Next step: download & apply to TO.",
      stats,
      errors,
    });
  } catch (err) {
    console.log("ATTACHMENTS SYNC ERROR:", err.response?.data || err.message);
    return res
      .status(500)
      .json(err.response?.data || { error: err.message });
  }
});

/* --------------------------------------------------
   2. ATTACHMENT COPY (DocNumber-based matching)
-------------------------------------------------- */

// POST /attachments/copy
// Body: { docTypes: ["Invoice","CreditMemo", ...] }
router.post("/attachments/copy", async (req, res) => {
  try {
    const FROM = getFromConnection();
    const TO = getToConnection();

    if (!FROM?.access_token || !FROM?.realmId) {
      return res.status(400).json({
        error:
          "FROM company not connected. Please connect using /qbo/auth-from first.",
      });
    }
    if (!TO?.access_token || !TO?.realmId) {
      return res.status(400).json({
        error:
          "TO company not connected. Please connect using /qbo/auth-to first.",
      });
    }

    const {
      docTypes = [
        "Invoice",
        "CreditMemo",
        "Bill",
        "VendorCredit",
        "SalesReceipt",
        "Estimate",
      ],
    } = req.body || {};

    const summary = {};
    const errors = [];

    // cache so same txn id ke liye baar-baar GET na ho
    const sourceDocCache = {};

    for (const type of docTypes) {
      console.log(`🚚 Copying attachments for type: ${type}`);

      const typeStats = {
        totalAttachables: 0,
        totalLinks: 0,
        copied: 0,
        skippedNoFile: 0,
        missingSourceDoc: 0,
        missingDocNumber: 0,
        missingTargetDoc: 0,
        uploadFailed: 0,
      };

      try {
        const attachables = await fetchAttachablesByType(FROM, type);
        typeStats.totalAttachables = attachables.length;

        for (const att of attachables) {
          const refs = att.AttachableRef || [];
          const fileUrl = att.FileAccessUri || att.TempDownloadUri;

          // Agar file hi nahi hai to skip
          if (!fileUrl) {
            typeStats.skippedNoFile++;
            continue;
          }

          for (const ref of refs) {
            if (!ref.EntityRef || ref.EntityRef.type !== type) continue;

            typeStats.totalLinks++;

            const sourceId = ref.EntityRef.value;
            if (!sourceId) {
              typeStats.missingSourceDoc++;
              continue;
            }

            // Source doc fetch (cached)
            const cacheKey = `${type}:${sourceId}`;
            let sourceDoc = sourceDocCache[cacheKey];

            if (!sourceDoc) {
              try {
                sourceDoc = await fetchSourceDocById(FROM, type, sourceId);
                sourceDocCache[cacheKey] = sourceDoc;
              } catch (err) {
                typeStats.missingSourceDoc++;
                errors.push(
                  `Source fetch failed (${type} #${sourceId}): ${
                    err.response?.data?.Fault?.Error?.[0]?.Message ||
                    err.message
                  }`
                );
                continue;
              }
            }

            if (!sourceDoc) {
              typeStats.missingSourceDoc++;
              continue;
            }

            const docNumber = getDocNumberForMatch(type, sourceDoc);
            if (!docNumber) {
              typeStats.missingDocNumber++;
              errors.push(
                `No DocNumber for source ${type} #${sourceId} (attachment ${att.Id})`
              );
              continue;
            }

            // Target doc find by DocNumber
            let targetDoc;
            try {
              targetDoc = await findTargetByDocNumber(TO, type, docNumber);
            } catch (err) {
              errors.push(
                `Target lookup failed (${type}, DocNumber=${docNumber}): ${
                  err.response?.data?.Fault?.Error?.[0]?.Message || err.message
                }`
              );
            }

            if (!targetDoc || !targetDoc.Id) {
              typeStats.missingTargetDoc++;
              errors.push(
                `No target ${type} found in TO for DocNumber=${docNumber} (source ${sourceId}, attachment ${att.Id})`
              );
              continue;
            }

            // File download
            let file;
            try {
              file = await downloadAttachmentFile(FROM, att);
              if (!file) {
                typeStats.skippedNoFile++;
                continue;
              }
            } catch (err) {
              typeStats.skippedNoFile++;
              errors.push(
                `Download failed (attachment ${att.Id}, ${type} DocNumber=${docNumber}): ${
                  err.response?.data?.Fault?.Error?.[0]?.Message || err.message
                }`
              );
              continue;
            }

            // Upload to TO
            try {
              await uploadAttachmentToTarget(
                TO,
                type,
                targetDoc.Id,
                file,
                att.Note || ""
              );
              typeStats.copied++;
            } catch (err) {
              typeStats.uploadFailed++;
              errors.push(
                `Upload failed (attachment ${att.Id} → TO ${type} #${targetDoc.Id}, DocNumber=${docNumber}): ${
                  err.response?.data?.Fault?.Error?.[0]?.Message || err.message
                }`
              );
            }
          }
        }
      } catch (err) {
        errors.push(
          `Type ${type} copy failed (top-level): ${
            err.response?.data?.Fault?.Error?.[0]?.Message || err.message
          }`
        );
      }

      summary[type] = typeStats;
    }

    return res.json({
      success: true,
      message:
        "Attachment copy finished (DocNumber-based match). For extra matching rules we can adjust getDocNumberForMatch / findTargetByDocNumber later.",
      summary,
      errors,
    });
  } catch (err) {
    console.log("ATTACHMENTS COPY ERROR:", err.response?.data || err.message);
    return res
      .status(500)
      .json(err.response?.data || { error: err.message });
  }
});

export default router;
