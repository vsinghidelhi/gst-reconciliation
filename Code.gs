/**
 * GST 2-Way Reconciliation Engine - Google Apps Script Backend
 * Developed for V. Singhi & Associates / vsinghidelhi
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
  var s = String(val).trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
  s = s.replace(/^0+/, '');
  return s ? s : "0";
}

function safeFloat(val) {
  if (val === null || val === undefined || val === "") return 0.0;
  var num = parseFloat(String(val).replace(/,/g, ''));
  return isNaN(num) ? 0.0 : Math.round(num * 100) / 100;
}

function stringSimilarity(s1, s2) {
  if (!s1 || !s2) return 0.0;
  if (s1 === s2) return 1.0;
  if (s1.indexOf(s2) !== -1 || s2.indexOf(s1) !== -1) return 0.85;
  if (s1.length >= 4 && s2.length >= 4 && s1.slice(-4) === s2.slice(-4)) return 0.80;
  return 0.0;
}

// -------------------------------------------------------------
// 2. CORE RECONCILIATION ENGINE (CHECK 1 & CHECK 2)
// -------------------------------------------------------------

function runReconciliation(booksRows, portalRows, tolerance) {
  if (!tolerance) tolerance = 2.0;

  // ---------------------------------------------------------
  // CHECK 1: VENDOR GSTIN PIVOT (MACRO CHECK)
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
    var bInfo = booksVendorAgg[gst] || { name: "", count: 0, taxable: 0, total_tax: 0 };
    var pInfo = portalVendorAgg[gst] || { name: "", count: 0, taxable: 0, total_tax: 0 };

    var bTax = Math.round(bInfo.total_tax * 100) / 100;
    var pTax = Math.round(pInfo.total_tax * 100) / 100;
    var taxDiff = Math.round((bTax - pTax) * 100) / 100;

    var status = "";
    if (booksVendorAgg[gst] && portalVendorAgg[gst]) {
      status = Math.abs(taxDiff) <= tolerance ? "100% Matched" : "Tax Variance";
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
      books_taxable: Math.round(bInfo.taxable * 100) / 100,
      portal_taxable: Math.round(pInfo.taxable * 100) / 100,
      taxable_variance: Math.round((bInfo.taxable - pInfo.taxable) * 100) / 100,
      books_inv_count: bInfo.count,
      portal_inv_count: pInfo.count
    });
  }

  // ---------------------------------------------------------
  // CHECK 2: INVOICE & AMOUNT LEVEL MATCHING (MICRO CHECK)
  // ---------------------------------------------------------
  // Multi-line aggregation in Books: key = cleanGstin + "||" + cleanInvoiceNo
  var booksByKey = {};
  for (var i = 0; i < booksRows.length; i++) {
    var b = booksRows[i];
    var g = cleanGstin(b.gstin);
    var cInv = cleanInvoiceNo(b.bill_no);
    var key = g + "||" + cInv;

    if (!booksByKey[key]) {
      booksByKey[key] = {
        gstin: g,
        bill_no: b.bill_no || "",
        clean_bill_no: cInv,
        vendor_name: b.vendor_name || "",
        taxable: 0.0,
        total_tax: 0.0,
        row_indices: []
      };
    }
    booksByKey[key].taxable += safeFloat(b.taxable);
    booksByKey[key].total_tax += safeFloat(b.total_tax);
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
      var diff = Math.abs(bAgg.total_tax - safeFloat(pRow.total_tax));

      if (diff <= tolerance) {
        matchedBooksKeys[key] = { pIdx: pIdx, status: "Matched (Exact)" };
      } else {
        matchedBooksKeys[key] = { pIdx: pIdx, status: "Value Mismatch" };
      }
      matchedPortalIndices[pIdx] = true;
    }
  }

  // Pass 2: Smart / Typo Match
  for (var key in booksByKey) {
    if (matchedBooksKeys[key]) continue;

    var bAgg = booksByKey[key];
    var gst = bAgg.gstin;
    var cBill = bAgg.clean_bill_no;
    var candidates = portalByGstin[gst] || [];

    var bestPIdx = -1;
    for (var c = 0; c < candidates.length; c++) {
      var pIdx = candidates[c];
      if (matchedPortalIndices[pIdx]) continue;
      var pRow = portalRows[pIdx];
      var diff = Math.abs(bAgg.total_tax - safeFloat(pRow.total_tax));

      if (diff <= tolerance) {
        var pDoc = cleanInvoiceNo(pRow.doc_no);
        var sim = stringSimilarity(cBill, pDoc);
        if (sim >= 0.80 || pDoc.indexOf(cBill) !== -1 || cBill.indexOf(pDoc) !== -1) {
          bestPIdx = pIdx;
          break;
        }
      }
    }

    if (bestPIdx !== -1) {
      matchedBooksKeys[key] = { pIdx: bestPIdx, status: "Matched (Smart/Typo)" };
      matchedPortalIndices[bestPIdx] = true;
    }
  }

  // ---------------------------------------------------------
  // BUILD 2-WAY RECONCILIATION TABLES
  // ---------------------------------------------------------
  // View 1: Books vs Portal
  var booksVsPortal = [];
  var actionableList = [];

  for (var i = 0; i < booksRows.length; i++) {
    var b = booksRows[i];
    var key = cleanGstin(b.gstin) + "||" + cleanInvoiceNo(b.bill_no);

    var bRecon = {
      branch: b.branch || "",
      sap_trans_no: b.sap_trans_no || b.trans_no || "",
      inv_date: b.inv_date || "",
      bill_no: b.bill_no || "",
      vendor_name: b.vendor_name || "",
      gstin: cleanGstin(b.gstin),
      taxable: safeFloat(b.taxable),
      total_tax: safeFloat(b.total_tax),
      portal_doc_no: "-",
      portal_tax: 0.0,
      tax_diff: safeFloat(b.total_tax),
      match_status: "Only in Books (Missing in 2B)"
    };

    if (matchedBooksKeys[key]) {
      var matchInfo = matchedBooksKeys[key];
      var pRow = portalRows[matchInfo.pIdx];
      var bAgg = booksByKey[key];
      bRecon.portal_doc_no = pRow.doc_no || "";
      bRecon.portal_tax = safeFloat(pRow.total_tax);
      bRecon.tax_diff = Math.round((bAgg.total_tax - safeFloat(pRow.total_tax)) * 100) / 100;
      bRecon.match_status = matchInfo.status;
    }

    booksVsPortal.push(bRecon);

    if (bRecon.match_status === "Only in Books (Missing in 2B)") {
      actionableList.push({
        action_type: "Vendor Follow-up (Missing in 2B)",
        gstin: bRecon.gstin,
        party_name: bRecon.vendor_name,
        invoice_no: bRecon.bill_no,
        invoice_date: bRecon.inv_date,
        taxable: bRecon.taxable,
        total_tax: bRecon.total_tax,
        variance: bRecon.total_tax,
        recommended_action: "Send follow-up email/reminder to Vendor with Bill details"
      });
    } else if (bRecon.match_status === "Value Mismatch") {
      actionableList.push({
        action_type: "Value Discrepancy Investigation",
        gstin: bRecon.gstin,
        party_name: bRecon.vendor_name,
        invoice_no: bRecon.bill_no,
        invoice_date: bRecon.inv_date,
        taxable: bRecon.taxable,
        total_tax: bRecon.total_tax,
        variance: bRecon.tax_diff,
        recommended_action: "Check tax rate/partial booking (Portal Tax: ₹" + bRecon.portal_tax + ")"
      });
    }
  }

  // View 2: Portal vs Books
  var portalToBooks = {};
  for (var key in matchedBooksKeys) {
    portalToBooks[matchedBooksKeys[key].pIdx] = { key: key, status: matchedBooksKeys[key].status };
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
      total_tax: safeFloat(p.total_tax),
      books_bill_no: "-",
      books_tax: 0.0,
      tax_diff: safeFloat(p.total_tax),
      match_status: "Only in Portal (Unbooked)"
    };

    if (portalToBooks[j]) {
      var k = portalToBooks[j].key;
      var bAgg = booksByKey[k];
      pRecon.books_bill_no = bAgg.bill_no;
      pRecon.books_tax = Math.round(bAgg.total_tax * 100) / 100;
      pRecon.tax_diff = Math.round((safeFloat(p.total_tax) - bAgg.total_tax) * 100) / 100;
      pRecon.match_status = portalToBooks[j].status;
    }

    portalVsBooks.push(pRecon);

    if (pRecon.match_status === "Only in Portal (Unbooked)") {
      actionableList.push({
        action_type: "Accounting Booking Pending",
        gstin: pRecon.gstin,
        party_name: pRecon.supplier_name,
        invoice_no: pRecon.doc_no,
        invoice_date: pRecon.doc_date,
        taxable: pRecon.taxable,
        total_tax: pRecon.total_tax,
        variance: pRecon.total_tax,
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
// 3. EXPORT RECONCILIATION TO GOOGLE SHEET
// -------------------------------------------------------------

function exportToGoogleSheet(reconData) {
  var ss = SpreadsheetApp.create("VSA GST Reconciliation Report - " + Utilities.formatDate(new Date(), "GMT+5:30", "dd-MMM-yyyy HH:mm"));
  var navyColor = "#1F4E79";

  // TAB 1: KPI Dashboard
  var sheetKpi = ss.getActiveSheet();
  sheetKpi.setName("Dashboard & Summary");
  sheetKpi.appendRow(["GST RECONCILIATION SUMMARY (PORTAL vs SAP BOOKS)"]);
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

  // TAB 2: Check 1 - Vendor Summary
  var sheetVendor = ss.insertSheet("1_Vendor_Pivot_Check");
  sheetVendor.appendRow([
    "Supplier GSTIN", "Vendor Name", "Status", "Books Tax", "Portal Tax", "Tax Variance",
    "Books Taxable", "Portal Taxable", "Taxable Variance", "Books Count", "Portal Count"
  ]);
  sheetVendor.getRange("A1:K1").setBackground(navyColor).setFontColor("#FFFFFF").setFontWeight("bold");
  var vRows = [];
  for (var i = 0; i < reconData.vendorSummary.length; i++) {
    var v = reconData.vendorSummary[i];
    vRows.push([
      v.gstin, v.vendor_name, v.status, v.books_tax, v.portal_tax, v.tax_variance,
      v.books_taxable, v.portal_taxable, v.taxable_variance, v.books_inv_count, v.portal_inv_count
    ]);
  }
  if (vRows.length > 0) {
    sheetVendor.getRange(2, 1, vRows.length, vRows[0].length).setValues(vRows);
    sheetVendor.getRange(2, 4, vRows.length, 6).setNumberFormat("#,##0.00");
  }

  // TAB 3: Check 2 - Books vs Portal
  var sheetBvP = ss.insertSheet("2_Books_vs_Portal");
  sheetBvP.appendRow([
    "Branch", "SAP Trans No", "Invoice Date", "Vendor Bill No", "Vendor Name", "Vendor GSTIN",
    "Books Taxable", "Books Total Tax", "Portal Doc No", "Portal Tax", "Tax Diff", "Reconciliation Status"
  ]);
  sheetBvP.getRange("A1:L1").setBackground(navyColor).setFontColor("#FFFFFF").setFontWeight("bold");
  var bRows = [];
  for (var i = 0; i < reconData.booksVsPortal.length; i++) {
    var b = reconData.booksVsPortal[i];
    bRows.push([
      b.branch, b.sap_trans_no, b.inv_date, b.bill_no, b.vendor_name, b.gstin,
      b.taxable, b.total_tax, b.portal_doc_no, b.portal_tax, b.tax_diff, b.match_status
    ]);
  }
  if (bRows.length > 0) {
    sheetBvP.getRange(2, 1, bRows.length, bRows[0].length).setValues(bRows);
    sheetBvP.getRange(2, 7, bRows.length, 5).setNumberFormat("#,##0.00");
  }

  // TAB 4: Check 2 - Portal vs Books
  var sheetPvB = ss.insertSheet("3_Portal_vs_Books");
  sheetPvB.appendRow([
    "Supplier GSTIN", "Supplier Name", "Portal Doc No", "Portal Doc Date",
    "Portal Taxable", "Portal Total Tax", "Matched Books Bill", "Books Tax", "Tax Diff", "Booking Status"
  ]);
  sheetPvB.getRange("A1:J1").setBackground(navyColor).setFontColor("#FFFFFF").setFontWeight("bold");
  var pRows = [];
  for (var j = 0; j < reconData.portalVsBooks.length; j++) {
    var p = reconData.portalVsBooks[j];
    pRows.push([
      p.gstin, p.supplier_name, p.doc_no, p.doc_date,
      p.taxable, p.total_tax, p.books_bill_no, p.books_tax, p.tax_diff, p.match_status
    ]);
  }
  if (pRows.length > 0) {
    sheetPvB.getRange(2, 1, pRows.length, pRows[0].length).setValues(pRows);
    sheetPvB.getRange(2, 5, pRows.length, 5).setNumberFormat("#,##0.00");
  }

  // TAB 5: Actionable Discrepancies
  var sheetAct = ss.insertSheet("4_Actionable_Discrepancies");
  sheetAct.appendRow([
    "Category", "Supplier GSTIN", "Party Name", "Invoice No", "Date",
    "Taxable Value", "Total Tax", "Variance / Amount", "Recommended Action"
  ]);
  sheetAct.getRange("A1:I1").setBackground(navyColor).setFontColor("#FFFFFF").setFontWeight("bold");
  var aRows = [];
  for (var a = 0; a < reconData.actionableList.length; a++) {
    var act = reconData.actionableList[a];
    aRows.push([
      act.action_type, act.gstin, act.party_name, act.invoice_no, act.invoice_date,
      act.taxable, act.total_tax, act.variance, act.recommended_action
    ]);
  }
  if (aRows.length > 0) {
    sheetAct.getRange(2, 1, aRows.length, aRows[0].length).setValues(aRows);
    sheetAct.getRange(2, 6, aRows.length, 3).setNumberFormat("#,##0.00");
  }

  return {
    sheetUrl: ss.getUrl(),
    sheetName: ss.getName()
  };
}
