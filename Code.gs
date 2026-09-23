/**
 * GST 2-Way Reconciliation Engine - Google Apps Script Backend
 * Developed for V. Singhi & Associates / vsinghidelhi
 * Features:
 *   - Check 1: Vendor GSTIN Pivot Macro Reconciliation (IGST, CGST, SGST)
 *   - Check 2: Micro Reconciliation with independent 3-tax-head verification
 *   - Place of Supply / Tax Head Mismatch detection (IGST vs CGST/SGST)
 *   - Side-by-side invoice & tax comparison columns with transparent match reasons
 */

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('VSA GST 2-Way Reconciliation Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// -------------------------------------------------------------
// 1. DATA NORMALIZATION HELPERS
// -------------------------------------------------------------

function cleanGstin(val) {
  if (!val) return "";
  return String(val).trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function cleanInvoiceNo(val) {
  if (!val) return "";
  var s = String(val).trim().toUpperCase();
  // 1. FY normalization: 2025-26 -> 25-26
  s = s.replace(/20(\d{2})[-/]?(\d{2})/g, '$1$2');
  // 2. Single digit padding: /5/ -> /05/
  s = s.replace(/([/-])([0-9])([/-])/g, '$10$2$3');
  // 3. Remove non-alphanumeric
  s = s.replace(/[^A-Z0-9]/g, '');
  // 4. Remove leading zeros
  s = s.replace(/^0+/, '');
  // 5. De-duplicate double paste (e.g. ABCABC -> ABC)
  var half = Math.floor(s.length / 2);
  if (half >= 5 && s.slice(0, half) === s.slice(half)) {
    s = s.slice(0, half);
  }
  return s ? s : "0";
}

function safeFloat(val) {
  if (val === null || val === undefined || val === "") return 0.0;
  var num = parseFloat(String(val).replace(/,/g, ''));
  return isNaN(num) ? 0.0 : Math.round(num * 100) / 100;
}

function getLevenshteinDistance(s1, s2) {
  if (s1.length < s2.length) return getLevenshteinDistance(s2, s1);
  if (s2.length === 0) return s1.length;
  var prev = [];
  for (var i = 0; i <= s2.length; i++) prev[i] = i;
  for (var i = 0; i < s1.length; i++) {
    var curr = [i + 1];
    for (var j = 0; j < s2.length; j++) {
      var ins = prev[j + 1] + 1;
      var dels = curr[j] + 1;
      var subs = prev[j] + (s1[i] !== s2[j] ? 1 : 0);
      curr.push(Math.min(ins, dels, subs));
    }
    prev = curr;
  }
  return prev[prev.length - 1];
}

function diagnoseSmartMatch(bBill, pDoc, bClean, pClean, bDate, pDate) {
  if (!bClean || !pClean) return { match: false, reason: "" };
  if (bBill.indexOf(',') !== -1 || bBill.indexOf(';') !== -1) return { match: false, reason: "" };

  if (bClean === pClean) {
    return { match: true, reason: "Normalized: FY format (2025-26 -> 25-26) / leading zeros / separators" };
  }

  // 1-char typo (e.g. 0048 vs 004B)
  if (bClean.length === pClean.length && bClean.length >= 2) {
    var diffCount = 0;
    var c1 = '', c2 = '';
    for (var i = 0; i < bClean.length; i++) {
      if (bClean[i] !== pClean[i]) {
        diffCount++;
        c1 = bClean[i];
        c2 = pClean[i];
      }
    }
    if (diffCount === 1) {
      return { match: true, reason: "1-Char Typo: '" + bBill + "' vs '" + pDoc + "' (Char '" + c1 + "' vs '" + c2 + "')" };
    }
  }

  // Levenshtein edit distance = 1
  if (Math.abs(bClean.length - pClean.length) <= 1 && bClean.length >= 3 && pClean.length >= 3) {
    if (getLevenshteinDistance(bClean, pClean) === 1) {
      return { match: true, reason: "1-Char Edit: '" + bBill + "' vs '" + pDoc + "'" };
    }
  }

  // Prefix omission (e.g. 079 vs PSB-079)
  if (bClean.length >= 2 && pClean.length >= 4) {
    var bNum = bClean.replace(/^0+/, '');
    var pNum = pClean.replace(/^[A-Z]+/, '').replace(/^0+/, '');
    if (bNum && pNum && bNum === pNum) {
      var pfx = (pClean.match(/^[A-Z]+/) || [''])[0];
      return { match: true, reason: "Prefix Omitted in Books: '" + pfx + "' ('" + bBill + "' vs '" + pDoc + "')" };
    }
    if (pClean.slice(-bClean.length) === bClean) {
      var omitted = pClean.slice(0, pClean.length - bClean.length);
      return { match: true, reason: "Prefix Omitted in Books: '" + omitted + "' ('" + bBill + "' vs '" + pDoc + "')" };
    }
  }

  // Known Abbreviation (e.g. TechBOT vs TB)
  if (bClean.length >= 6 && pClean.length >= 6 && bClean.slice(0, 4) === pClean.slice(0, 4) && bClean.slice(-4) === pClean.slice(-4)) {
    return { match: true, reason: "Abbreviation Match: '" + bBill + "' vs '" + pDoc + "'" };
  }

  // Double-paste correction in Books
  if (bClean.length >= 8 && pClean.length >= 6 && bClean.indexOf(pClean) !== -1) {
    return { match: true, reason: "Double-paste in Books corrected: '" + bBill + "' vs '" + pDoc + "'" };
  }

  // Date entered in Doc No
  var cleanPNums = pDoc.replace(/[^0-9]/g, '');
  if (cleanPNums.length >= 6 && bDate) {
    var cleanBDate = bDate.replace(/[^0-9]/g, '');
    if (cleanPNums.indexOf(cleanBDate) !== -1 || cleanBDate.indexOf(cleanPNums) !== -1) {
      return { match: true, reason: "Date entered in Doc No: '" + pDoc + "'" };
    }
  }

  return { match: false, reason: "" };
}

// -------------------------------------------------------------
// 2. CORE RECONCILIATION ENGINE (CHECK 1 & CHECK 2)
// -------------------------------------------------------------

function runReconciliation(booksRows, portalRows, tolerance) {
  if (tolerance === undefined || tolerance === null || isNaN(tolerance)) {
    tolerance = 2.0;
  }

  // ---------------------------------------------------------
  // CHECK 1: VENDOR GSTIN PIVOT (MACRO CHECK WITH 3 HEADS)
  // ---------------------------------------------------------
  var booksVendorAgg = {};
  for (var i = 0; i < booksRows.length; i++) {
    var b = booksRows[i];
    var g = cleanGstin(b.gstin);
    if (!g) continue;
    if (!booksVendorAgg[g]) {
      booksVendorAgg[g] = {
        name: b.vendor_name || b.name || "",
        count: 0,
        taxable: 0.0,
        total_tax: 0.0,
        igst: 0.0,
        cgst: 0.0,
        sgst: 0.0
      };
    }
    booksVendorAgg[g].count += 1;
    booksVendorAgg[g].taxable += safeFloat(b.taxable);
    booksVendorAgg[g].total_tax += safeFloat(b.total_tax);
    booksVendorAgg[g].igst += safeFloat(b.igst);
    booksVendorAgg[g].cgst += safeFloat(b.cgst);
    booksVendorAgg[g].sgst += safeFloat(b.sgst);
  }

  var portalVendorAgg = {};
  for (var j = 0; j < portalRows.length; j++) {
    var p = portalRows[j];
    var g = cleanGstin(p.gstin);
    if (!g) continue;
    if (!portalVendorAgg[g]) {
      portalVendorAgg[g] = {
        name: p.supplier_name || p.name || "",
        count: 0,
        taxable: 0.0,
        total_tax: 0.0,
        igst: 0.0,
        cgst: 0.0,
        sgst: 0.0
      };
    }
    portalVendorAgg[g].count += 1;
    portalVendorAgg[g].taxable += safeFloat(p.taxable);
    portalVendorAgg[g].total_tax += safeFloat(p.total_tax);
    portalVendorAgg[g].igst += safeFloat(p.igst);
    portalVendorAgg[g].cgst += safeFloat(p.cgst);
    portalVendorAgg[g].sgst += safeFloat(p.sgst);
  }

  var allGstinsSet = {};
  for (var g in booksVendorAgg) allGstinsSet[g] = true;
  for (var g in portalVendorAgg) allGstinsSet[g] = true;
  var allGstins = Object.keys(allGstinsSet).sort();

  var vendorSummary = [];
  for (var k = 0; k < allGstins.length; k++) {
    var gst = allGstins[k];
    var bInfo = booksVendorAgg[gst] || { name: "", count: 0, taxable: 0, total_tax: 0, igst: 0, cgst: 0, sgst: 0 };
    var pInfo = portalVendorAgg[gst] || { name: "", count: 0, taxable: 0, total_tax: 0, igst: 0, cgst: 0, sgst: 0 };

    var bTax = Math.round(bInfo.total_tax * 100) / 100;
    var pTax = Math.round(pInfo.total_tax * 100) / 100;
    var taxDiff = Math.round((bTax - pTax) * 100) / 100;

    var bIgst = Math.round(bInfo.igst * 100) / 100, pIgst = Math.round(pInfo.igst * 100) / 100;
    var bCgst = Math.round(bInfo.cgst * 100) / 100, pCgst = Math.round(pInfo.cgst * 100) / 100;
    var bSgst = Math.round(bInfo.sgst * 100) / 100, pSgst = Math.round(pInfo.sgst * 100) / 100;

    var igstDiff = Math.round((bIgst - pIgst) * 100) / 100;
    var cgstDiff = Math.round((bCgst - pCgst) * 100) / 100;
    var sgstDiff = Math.round((bSgst - pSgst) * 100) / 100;

    var dynamicTol = (tolerance === 0) ? 0 : tolerance;

    var status = "";
    if (booksVendorAgg[gst] && portalVendorAgg[gst]) {
      var headsMatch = (Math.abs(igstDiff) <= dynamicTol && Math.abs(cgstDiff) <= dynamicTol && Math.abs(sgstDiff) <= dynamicTol);
      if (Math.abs(taxDiff) <= tolerance && headsMatch) {
        status = "100% Matched";
      } else if (Math.abs(taxDiff) <= dynamicTol && !headsMatch) {
        status = "Tax Head Mismatch (IGST vs CGST/SGST)";
      } else {
        status = "Tax Variance";
      }
    } else if (booksVendorAgg[gst]) {
      status = "Only in Books";
    } else {
      status = "Only in Portal";
    }

    vendorSummary.push({
      gstin: gst,
      vendor_name: bInfo.name || pInfo.name || "Vendor",
      status: status,
      books_tax: bTax,
      portal_tax: pTax,
      tax_variance: taxDiff,
      books_igst: bIgst,
      portal_igst: pIgst,
      igst_diff: igstDiff,
      books_cgst: bCgst,
      portal_cgst: pCgst,
      cgst_diff: cgstDiff,
      books_sgst: bSgst,
      portal_sgst: pSgst,
      sgst_diff: sgstDiff,
      books_taxable: Math.round(bInfo.taxable * 100) / 100,
      portal_taxable: Math.round(pInfo.taxable * 100) / 100,
      taxable_variance: Math.round((bInfo.taxable - pInfo.taxable) * 100) / 100,
      books_inv_count: bInfo.count,
      portal_inv_count: pInfo.count
    });
  }

  // ---------------------------------------------------------
  // CHECK 2: INVOICE & TAX HEAD MICRO RECONCILIATION
  // ---------------------------------------------------------
  var booksByKey = {};
  for (var i = 0; i < booksRows.length; i++) {
    var b = booksRows[i];
    var g = cleanGstin(b.gstin);
    var cInv = cleanInvoiceNo(b.bill_no);
    var key = g + "||" + cInv;
    if (!cInv || cInv === "0") {
      var docRef = b.sap_trans_no || b.trans_no || ("ROW_" + i);
      key = g + "||DOC_" + docRef;
    }

    if (!booksByKey[key]) {
      booksByKey[key] = {
        branch: b.branch || "",
        sap_trans_no: b.sap_trans_no || b.trans_no || "",
        inv_date: b.inv_date || "",
        bill_no: b.bill_no || "",
        clean_bill_no: cInv,
        vendor_name: b.vendor_name || "",
        gstin: g,
        line_count: 0,
        taxable: 0.0,
        total_tax: 0.0,
        igst: 0.0,
        cgst: 0.0,
        sgst: 0.0,
        row_indices: []
      };
    }
    booksByKey[key].line_count += 1;
    booksByKey[key].taxable = Math.round((booksByKey[key].taxable + safeFloat(b.taxable)) * 100) / 100;
    booksByKey[key].total_tax = Math.round((booksByKey[key].total_tax + safeFloat(b.total_tax)) * 100) / 100;
    booksByKey[key].igst = Math.round((booksByKey[key].igst + safeFloat(b.igst)) * 100) / 100;
    booksByKey[key].cgst = Math.round((booksByKey[key].cgst + safeFloat(b.cgst)) * 100) / 100;
    booksByKey[key].sgst = Math.round((booksByKey[key].sgst + safeFloat(b.sgst)) * 100) / 100;
    booksByKey[key].row_indices.push(i);
  }

  var portalByKey = {};
  var portalByGstin = {};
  for (var j = 0; j < portalRows.length; j++) {
    var p = portalRows[j];
    var g = cleanGstin(p.gstin);
    var cDoc = cleanInvoiceNo(p.doc_no);
    var key = g + "||" + cDoc;
    portalByKey[key] = j;

    if (!portalByGstin[g]) portalByGstin[g] = [];
    portalByGstin[g].push(j);
  }

  var matchedBooksKeys = {};
  var matchedPortalIndices = {};

  // Pass 1: Exact Key Match
  for (var key in booksByKey) {
    if (portalByKey.hasOwnProperty(key)) {
      var pIdx = portalByKey[key];
      var pRow = portalRows[pIdx];
      var bAgg = booksByKey[key];

      var diffTot = Math.round(Math.abs(bAgg.total_tax - safeFloat(pRow.total_tax)) * 100) / 100;
      var diffIgst = Math.round(Math.abs(bAgg.igst - safeFloat(pRow.igst)) * 100) / 100;
      var diffCgst = Math.round(Math.abs(bAgg.cgst - safeFloat(pRow.cgst)) * 100) / 100;
      var diffSgst = Math.round(Math.abs(bAgg.sgst - safeFloat(pRow.sgst)) * 100) / 100;

      var st = "", rsn = "";
      if (diffTot <= tolerance) {
        if (diffIgst <= tolerance && diffCgst <= tolerance && diffSgst <= tolerance) {
          st = "Matched (Exact)";
          rsn = "Exact Match: Invoice No & All Tax Heads (IGST/CGST/SGST)";
        } else {
          st = "Tax Head Mismatch (IGST vs CGST/SGST)";
          rsn = "Total Tax matches, but Head Mismatch: Books(I:" + bAgg.igst + ", C:" + bAgg.cgst + ", S:" + bAgg.sgst + ") vs Portal(I:" + pRow.igst + ", C:" + pRow.cgst + ", S:" + pRow.sgst + ")";
        }
      } else {
        st = "Value Mismatch";
        rsn = "Invoice matches, but Tax difference is ₹" + diffTot.toFixed(2);
      }

      matchedBooksKeys[key] = { pIdx: pIdx, status: st, reason: rsn };
      matchedPortalIndices[pIdx] = true;
    }
  }

  // Pass 2: Explainable Smart / Typo Match
  for (var key in booksByKey) {
    if (matchedBooksKeys[key]) continue;

    var bAgg = booksByKey[key];
    var gst = bAgg.gstin;
    var cBill = bAgg.clean_bill_no;
    var candidates = portalByGstin[gst] || [];

    var bestPIdx = -1;
    var bestStatus = "";
    var bestReason = "";

    for (var c = 0; c < candidates.length; c++) {
      var pIdx = candidates[c];
      if (matchedPortalIndices[pIdx]) continue;
      var pRow = portalRows[pIdx];

      var diffTot = Math.round(Math.abs(bAgg.total_tax - safeFloat(pRow.total_tax)) * 100) / 100;
      var diffIgst = Math.round(Math.abs(bAgg.igst - safeFloat(pRow.igst)) * 100) / 100;
      var diffCgst = Math.round(Math.abs(bAgg.cgst - safeFloat(pRow.cgst)) * 100) / 100;
      var diffSgst = Math.round(Math.abs(bAgg.sgst - safeFloat(pRow.sgst)) * 100) / 100;

      if (diffTot <= tolerance) {
        var diag = diagnoseSmartMatch(bAgg.bill_no, pRow.doc_no, cBill, cleanInvoiceNo(pRow.doc_no), bAgg.inv_date, pRow.doc_date);
        var isMatch = diag.match;
        var reason = diag.reason;

        var isBlankBill = (!cBill || cBill === "0" || (cBill + "").indexOf("DOC_") === 0);
        if (!isMatch && isBlankBill) {
          var sameTaxCount = 0;
          for (var ci = 0; ci < candidates.length; ci++) {
            if (Math.abs(safeFloat(portalRows[candidates[ci]].total_tax) - bAgg.total_tax) <= tolerance) sameTaxCount++;
          }
          if (sameTaxCount === 1) {
            isMatch = true;
            reason = "Blank Bill No in Books: Unique bill for vendor " + bAgg.vendor_name;
          }
        }

        if (isMatch) {
          if (diffIgst <= tolerance && diffCgst <= tolerance && diffSgst <= tolerance) {
            bestStatus = "Matched (Smart/Typo)";
          } else {
            bestStatus = "Tax Head Mismatch (IGST vs CGST/SGST)";
            reason += " | Head Mismatch: Books(I:" + bAgg.igst + ", C:" + bAgg.cgst + ", S:" + bAgg.sgst + ") vs Portal(I:" + pRow.igst + ", C:" + pRow.cgst + ", S:" + pRow.sgst + ")";
          }
          bestPIdx = pIdx;
          bestReason = reason;
          break;
        }
      }
    }

    if (bestPIdx !== -1) {
      matchedBooksKeys[key] = { pIdx: bestPIdx, status: bestStatus, reason: bestReason };
      matchedPortalIndices[bestPIdx] = true;
    }
  }

  // ---------------------------------------------------------
  // BUILD SIDE-BY-SIDE RECONCILIATION TABLES (1 Row per Invoice)
  // ---------------------------------------------------------
  var booksVsPortal = [];
  var actionableList = [];

  for (var key in booksByKey) {
    var b = booksByKey[key];
    var transDisplay = b.sap_trans_no || "";
    if (b.line_count > 1) {
      transDisplay = (transDisplay ? transDisplay + " " : "") + "(" + b.line_count + " items)";
    }

    var bRecon = {
      branch: b.branch || "",
      sap_trans_no: transDisplay,
      inv_date: b.inv_date || "",
      bill_no: b.bill_no || "",
      vendor_name: b.vendor_name || "",
      gstin: b.gstin,
      line_count: b.line_count,
      taxable: safeFloat(b.taxable),
      igst: safeFloat(b.igst),
      cgst: safeFloat(b.cgst),
      sgst: safeFloat(b.sgst),
      total_tax: safeFloat(b.total_tax),
      portal_doc_no: "-",
      portal_doc_date: "-",
      portal_taxable: 0.0,
      taxable_diff: safeFloat(b.taxable),
      portal_igst: 0.0,
      igst_diff: safeFloat(b.igst),
      portal_cgst: 0.0,
      cgst_diff: safeFloat(b.cgst),
      portal_sgst: 0.0,
      sgst_diff: safeFloat(b.sgst),
      portal_tax: 0.0,
      tax_diff: safeFloat(b.total_tax),
      match_status: "Only in Books (Missing in 2B)",
      match_reason: "Vendor has not filed invoice in GSTR-1 or GSTIN mismatch" + (b.line_count > 1 ? " (" + b.line_count + " items)" : "")
    };

    if (matchedBooksKeys[key]) {
      var matchInfo = matchedBooksKeys[key];
      var pRow = portalRows[matchInfo.pIdx];
      bRecon.portal_doc_no = pRow.doc_no || "";
      bRecon.portal_doc_date = pRow.doc_date || "";
      bRecon.portal_taxable = safeFloat(pRow.taxable);
      bRecon.taxable_diff = Math.round((safeFloat(b.taxable) - safeFloat(pRow.taxable)) * 100) / 100;
      bRecon.portal_igst = safeFloat(pRow.igst);
      bRecon.igst_diff = Math.round((safeFloat(b.igst) - safeFloat(pRow.igst)) * 100) / 100;
      bRecon.portal_cgst = safeFloat(pRow.cgst);
      bRecon.cgst_diff = Math.round((safeFloat(b.cgst) - safeFloat(pRow.cgst)) * 100) / 100;
      bRecon.portal_sgst = safeFloat(pRow.sgst);
      bRecon.sgst_diff = Math.round((safeFloat(b.sgst) - safeFloat(pRow.sgst)) * 100) / 100;
      bRecon.portal_tax = safeFloat(pRow.total_tax);
      bRecon.tax_diff = Math.round((b.total_tax - safeFloat(pRow.total_tax)) * 100) / 100;
      bRecon.match_status = matchInfo.status;
      if (b.line_count > 1) {
        bRecon.match_reason = matchInfo.reason + " (Consolidated " + b.line_count + " line items in Books)";
      } else {
        bRecon.match_reason = matchInfo.reason;
      }
    }

    booksVsPortal.push(bRecon);

    if (bRecon.match_status === "Only in Books (Missing in 2B)") {
      actionableList.push({
        action_type: "Vendor Follow-up (Missing in 2B)",
        gstin: bRecon.gstin,
        party_name: bRecon.vendor_name,
        books_bill_no: bRecon.bill_no,
        portal_doc_no: "-",
        invoice_date: bRecon.inv_date,
        taxable: bRecon.taxable,
        total_tax: bRecon.total_tax,
        heads_breakup: "Books(I:" + bRecon.igst + ", C:" + bRecon.cgst + ", S:" + bRecon.sgst + ") | Portal(-)",
        recommended_action: "Send follow-up email/reminder to Vendor with Bill details"
      });
    } else if (bRecon.match_status.indexOf("Head Mismatch") !== -1) {
      actionableList.push({
        action_type: "Tax Head POS Correction",
        gstin: bRecon.gstin,
        party_name: bRecon.vendor_name,
        books_bill_no: bRecon.bill_no,
        portal_doc_no: bRecon.portal_doc_no,
        invoice_date: bRecon.inv_date,
        taxable: bRecon.taxable,
        total_tax: bRecon.total_tax,
        heads_breakup: "Books(I:" + bRecon.igst + ", C:" + bRecon.cgst + ", S:" + bRecon.sgst + ") vs Portal(I:" + bRecon.portal_igst + ", C:" + bRecon.portal_cgst + ", S:" + bRecon.portal_sgst + ")",
        recommended_action: "Reclassify entry in SAP: Intra-state vs Inter-state tax heads"
      });
    } else if (bRecon.match_status === "Value Mismatch") {
      actionableList.push({
        action_type: "Value Discrepancy Investigation",
        gstin: bRecon.gstin,
        party_name: bRecon.vendor_name,
        books_bill_no: bRecon.bill_no,
        portal_doc_no: bRecon.portal_doc_no,
        invoice_date: bRecon.inv_date,
        taxable: bRecon.taxable,
        total_tax: bRecon.total_tax,
        heads_breakup: "Books Tax: ₹" + bRecon.total_tax + " vs Portal Tax: ₹" + bRecon.portal_tax,
        recommended_action: "Check tax rate/partial booking (Portal Tax: ₹" + bRecon.portal_tax + ")"
      });
    }
  }

  // View 2: Portal vs Books (Side-by-Side)
  var portalToBooks = {};
  for (var key in matchedBooksKeys) {
    portalToBooks[matchedBooksKeys[key].pIdx] = { key: key, status: matchedBooksKeys[key].status, reason: matchedBooksKeys[key].reason };
  }

  var portalVsBooks = [];
  for (var j = 0; j < portalRows.length; j++) {
    var p = portalRows[j];
    var pRecon = {
      gstin: cleanGstin(p.gstin),
      supplier_name: p.supplier_name || "",
      doc_no: p.doc_no || "",
      doc_date: p.doc_date || "",
      taxable: safeFloat(p.taxable),
      igst: safeFloat(p.igst),
      cgst: safeFloat(p.cgst),
      sgst: safeFloat(p.sgst),
      total_tax: safeFloat(p.total_tax),
      books_bill_no: "-",
      books_inv_date: "-",
      books_taxable: 0.0,
      taxable_diff: safeFloat(p.taxable),
      books_igst: 0.0,
      igst_diff: safeFloat(p.igst),
      books_cgst: 0.0,
      cgst_diff: safeFloat(p.cgst),
      books_sgst: 0.0,
      sgst_diff: safeFloat(p.sgst),
      books_tax: 0.0,
      tax_diff: safeFloat(p.total_tax),
      match_status: "Only in Portal (Unbooked)",
      match_reason: "Bill available in 2B but not recorded in SAP Books"
    };

    if (portalToBooks[j]) {
      var m = portalToBooks[j];
      var bAgg = booksByKey[m.key];
      pRecon.books_bill_no = bAgg.bill_no;
      pRecon.books_inv_date = bAgg.inv_date;
      pRecon.books_taxable = Math.round(bAgg.taxable * 100) / 100;
      pRecon.taxable_diff = Math.round((safeFloat(p.taxable) - bAgg.taxable) * 100) / 100;
      pRecon.books_igst = Math.round(bAgg.igst * 100) / 100;
      pRecon.igst_diff = Math.round((safeFloat(p.igst) - bAgg.igst) * 100) / 100;
      pRecon.books_cgst = Math.round(bAgg.cgst * 100) / 100;
      pRecon.cgst_diff = Math.round((safeFloat(p.cgst) - bAgg.cgst) * 100) / 100;
      pRecon.books_sgst = Math.round(bAgg.sgst * 100) / 100;
      pRecon.sgst_diff = Math.round((safeFloat(p.sgst) - bAgg.sgst) * 100) / 100;
      pRecon.books_tax = Math.round(bAgg.total_tax * 100) / 100;
      pRecon.tax_diff = Math.round((safeFloat(p.total_tax) - bAgg.total_tax) * 100) / 100;
      pRecon.match_status = m.status;
      pRecon.match_reason = m.reason;
    }

    portalVsBooks.push(pRecon);

    if (pRecon.match_status === "Only in Portal (Unbooked)") {
      actionableList.push({
        action_type: "Accounting Booking Pending",
        gstin: pRecon.gstin,
        party_name: pRecon.supplier_name,
        books_bill_no: "-",
        portal_doc_no: pRecon.doc_no,
        invoice_date: pRecon.doc_date,
        taxable: pRecon.taxable,
        total_tax: pRecon.total_tax,
        heads_breakup: "Portal(I:" + pRecon.igst + ", C:" + pRecon.cgst + ", S:" + pRecon.sgst + ") | Books(-)",
        recommended_action: "Check physical bill copy and book in SAP to claim ITC"
      });
    }
  }

  // ---------------------------------------------------------
  // KPI CALCULATIONS
  // ---------------------------------------------------------
  var totalBooksTax = 0;
  var matchedTax = 0;
  var missingBooksTax = 0;
  var matchedCount = 0;
  var missingBooksCount = 0;

  for (var i = 0; i < booksVsPortal.length; i++) {
    var r = booksVsPortal[i];
    totalBooksTax += r.total_tax;
    if (r.match_status.indexOf("Matched") !== -1) {
      matchedTax += r.total_tax;
      matchedCount += 1;
    } else if (r.match_status === "Only in Books (Missing in 2B)") {
      missingBooksTax += r.total_tax;
      missingBooksCount += 1;
    }
  }

  var totalPortalTax = 0;
  var unbookedPortalTax = 0;
  var unbookedPortalCount = 0;
  for (var j = 0; j < portalVsBooks.length; j++) {
    var r = portalVsBooks[j];
    totalPortalTax += r.total_tax;
    if (r.match_status === "Only in Portal (Unbooked)") {
      unbookedPortalTax += r.total_tax;
      unbookedPortalCount += 1;
    }
  }

  var kpis = {
    totalBooksTax: Math.round(totalBooksTax * 100) / 100,
    totalPortalTax: Math.round(totalPortalTax * 100) / 100,
    netDifference: Math.round((totalBooksTax - totalPortalTax) * 100) / 100,
    matchedTax: Math.round(matchedTax * 100) / 100,
    matchedCount: matchedCount,
    missingBooksTax: Math.round(missingBooksTax * 100) / 100,
    missingBooksCount: missingBooksCount,
    unbookedPortalTax: Math.round(unbookedPortalTax * 100) / 100,
    unbookedPortalCount: unbookedPortalCount,
    totalBooksRows: booksRows.length,
    totalPortalRows: portalRows.length,
    totalVendors: vendorSummary.length
  };

  return {
    kpis: kpis,
    vendorSummary: vendorSummary,
    booksVsPortal: booksVsPortal,
    portalVsBooks: portalVsBooks,
    actionableList: actionableList
  };
}

// -------------------------------------------------------------
// 3. EXPORT RECONCILIATION TO GOOGLE SHEET (SIDE-BY-SIDE)
// -------------------------------------------------------------

function exportToGoogleSheet(reconData) {
  var ss = SpreadsheetApp.create("VSA GST Reconciliation Report - " + Utilities.formatDate(new Date(), "GMT+5:30", "dd-MMM-yyyy HH:mm"));
  var navyColor = "#1F4E79";

  // TAB 1: KPI Dashboard
  var sheetKpi = ss.getActiveSheet();
  sheetKpi.setName("Dashboard & Summary");
  sheetKpi.appendRow(["GST 2-WAY RECONCILIATION SUMMARY (PORTAL vs SAP BOOKS)"]);
  sheetKpi.getRange("A1").setFontSize(14).setFontWeight("bold").setFontColor(navyColor);
  sheetKpi.appendRow(["Generated at: " + new Date().toLocaleString()]);
  sheetKpi.appendRow([]);
  sheetKpi.appendRow(["Key Metric", "Count", "Total Amount (₹)"]);
  sheetKpi.getRange("A4:C4").setBackground(navyColor).setFontColor("#FFFFFF").setFontWeight("bold");

  var k = reconData.kpis;
  sheetKpi.appendRow(["Total ITC Recorded in SAP Books", k.totalBooksRows, k.totalBooksTax]);
  sheetKpi.appendRow(["Total ITC Available in Portal (2B)", k.totalPortalRows, k.totalPortalTax]);
  sheetKpi.appendRow(["Net Variance (Books - Portal)", "-", k.netDifference]);
  sheetKpi.appendRow(["Reconciled ITC (Matched)", k.matchedCount, k.matchedTax]);
  sheetKpi.appendRow(["ITC at Risk (Missing in 2B)", k.missingBooksCount, k.missingBooksTax]);
  sheetKpi.appendRow(["Unclaimed ITC (Unbooked)", k.unbookedPortalCount, k.unbookedPortalTax]);
  sheetKpi.getRange("C5:C10").setNumberFormat("#,##0.00");

  // TAB 2: Check 1 - Vendor Summary with 3 Heads
  var sheetVendor = ss.insertSheet("1_Vendor_Pivot_Check");
  sheetVendor.appendRow([
    "Supplier GSTIN", "Vendor Name", "Status",
    "Books Total Tax", "Portal Total Tax", "Tax Variance",
    "Books IGST", "Portal IGST", "IGST Variance",
    "Books CGST", "Portal CGST", "CGST Variance",
    "Books SGST", "Portal SGST", "SGST Variance",
    "Books Taxable", "Portal Taxable", "Taxable Variance",
    "Books Count", "Portal Count"
  ]);
  sheetVendor.getRange("A1:T1").setBackground(navyColor).setFontColor("#FFFFFF").setFontWeight("bold");
  var vRows = [];
  for (var i = 0; i < reconData.vendorSummary.length; i++) {
    var v = reconData.vendorSummary[i];
    vRows.push([
      v.gstin, v.vendor_name, v.status,
      v.books_tax, v.portal_tax, v.tax_variance,
      v.books_igst, v.portal_igst, v.igst_diff,
      v.books_cgst, v.portal_cgst, v.cgst_diff,
      v.books_sgst, v.portal_sgst, v.sgst_diff,
      v.books_taxable, v.portal_taxable, v.taxable_variance,
      v.books_inv_count, v.portal_inv_count
    ]);
  }
  if (vRows.length > 0) {
    sheetVendor.getRange(2, 1, vRows.length, vRows[0].length).setValues(vRows);
    sheetVendor.getRange(2, 4, vRows.length, 15).setNumberFormat("#,##0.00");
  }

  // TAB 3: Check 2 - Books vs Portal (Side-by-Side)
  var sheetBvP = ss.insertSheet("2_Books_vs_Portal");
  sheetBvP.appendRow([
    "Supplier GSTIN", "Vendor Name",
    "Books Bill No", "Portal Doc No",
    "Match Status", "Match Reason / Typo Basis",
    "Books Inv Date", "Portal Doc Date",
    "Books Taxable", "Portal Taxable", "Taxable Diff",
    "Books IGST", "Portal IGST", "IGST Diff",
    "Books CGST", "Portal CGST", "CGST Diff",
    "Books SGST", "Portal SGST", "SGST Diff",
    "Books Total Tax", "Portal Total Tax", "Total Tax Diff",
    "Branch", "SAP Trans No"
  ]);
  sheetBvP.getRange("A1:Y1").setBackground(navyColor).setFontColor("#FFFFFF").setFontWeight("bold");
  var bRows = [];
  for (var i = 0; i < reconData.booksVsPortal.length; i++) {
    var b = reconData.booksVsPortal[i];
    bRows.push([
      b.gstin, b.vendor_name,
      b.bill_no, b.portal_doc_no,
      b.match_status, b.match_reason,
      b.inv_date, b.portal_doc_date,
      b.taxable, b.portal_taxable, b.taxable_diff,
      b.igst, b.portal_igst, b.igst_diff,
      b.cgst, b.portal_cgst, b.cgst_diff,
      b.sgst, b.portal_sgst, b.sgst_diff,
      b.total_tax, b.portal_tax, b.tax_diff,
      b.branch, b.sap_trans_no
    ]);
  }
  if (bRows.length > 0) {
    sheetBvP.getRange(2, 1, bRows.length, bRows[0].length).setValues(bRows);
    sheetBvP.getRange(2, 9, bRows.length, 15).setNumberFormat("#,##0.00");
  }

  // TAB 4: Check 2 - Portal vs Books (Side-by-Side)
  var sheetPvB = ss.insertSheet("3_Portal_vs_Books");
  sheetPvB.appendRow([
    "Supplier GSTIN", "Supplier Name",
    "Portal Doc No", "Books Bill No",
    "Match Status", "Match Reason / Typo Basis",
    "Portal Doc Date", "Books Inv Date",
    "Portal Taxable", "Books Taxable", "Taxable Diff",
    "Portal IGST", "Books IGST", "IGST Diff",
    "Portal CGST", "Books CGST", "CGST Diff",
    "Portal SGST", "Books SGST", "SGST Diff",
    "Portal Total Tax", "Books Total Tax", "Total Tax Diff"
  ]);
  sheetPvB.getRange("A1:V1").setBackground(navyColor).setFontColor("#FFFFFF").setFontWeight("bold");
  var pRows = [];
  for (var j = 0; j < reconData.portalVsBooks.length; j++) {
    var p = reconData.portalVsBooks[j];
    pRows.push([
      p.gstin, p.supplier_name,
      p.doc_no, p.books_bill_no,
      p.match_status, p.match_reason,
      p.doc_date, p.books_inv_date,
      p.taxable, p.books_taxable, p.taxable_diff,
      p.igst, p.books_igst, p.igst_diff,
      p.cgst, p.books_cgst, p.cgst_diff,
      p.sgst, p.books_sgst, p.sgst_diff,
      p.total_tax, p.books_tax, p.tax_diff
    ]);
  }
  if (pRows.length > 0) {
    sheetPvB.getRange(2, 1, pRows.length, pRows[0].length).setValues(pRows);
    sheetPvB.getRange(2, 9, pRows.length, 15).setNumberFormat("#,##0.00");
  }

  // TAB 5: Actionable Discrepancies
  var sheetAct = ss.insertSheet("4_Actionable_Discrepancies");
  sheetAct.appendRow([
    "Category", "Supplier GSTIN", "Party Name",
    "Books Bill No", "Portal Doc No", "Date",
    "Taxable Value", "Total Tax", "Heads Breakup (IGST/CGST/SGST)",
    "Recommended Action"
  ]);
  sheetAct.getRange("A1:J1").setBackground(navyColor).setFontColor("#FFFFFF").setFontWeight("bold");
  var aRows = [];
  for (var a = 0; a < reconData.actionableList.length; a++) {
    var act = reconData.actionableList[a];
    aRows.push([
      act.action_type, act.gstin, act.party_name,
      act.books_bill_no, act.portal_doc_no, act.invoice_date,
      act.taxable, act.total_tax, act.heads_breakup,
      act.recommended_action
    ]);
  }
  if (aRows.length > 0) {
    sheetAct.getRange(2, 1, aRows.length, aRows[0].length).setValues(aRows);
    sheetAct.getRange(2, 7, aRows.length, 2).setNumberFormat("#,##0.00");
  }

  return {
    sheetUrl: ss.getUrl(),
    sheetName: ss.getName()
  };
}
