// ==UserScript==
// @name         QuickBooks Invoice Print + Pick Slip (New UI Table Fix)
// @namespace    http://tampermonkey.net/
// @version      5.7
// @description  Print / Pick Slip for QuickBooks invoice editor using exact field selectors and header-mapped column
// @author       Raj - Gorkhari
// @match        https://qbo.intuit.com/*
// @include      https://qbo.intuit.com/app/invoice?*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
  "use strict";

  const CONFIG = {
    buttonCheckIntervalMs: 2000,
    mutationDebounceMs: 500,
    initialBootDelayMs: 2200,
    stableReadAttempts: 6,
    stableReadDelayMs: 250,
    debug: true,
  };

  const STATE = {
    addButtonsInFlight: false,
    mutationTimer: null,
    currentInvoiceId: null,
  };

  function log(...args) {
    if (CONFIG.debug) console.log("[QBO Pick Slip]", ...args);
  }

  function warn(...args) {
    console.warn("[QBO Pick Slip]", ...args);
  }

  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  function normalizeText(value) {
    return String(value ?? "").replace(/\s+/g, " ").trim();
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function tryParseNumber(raw) {
    if (raw == null) return null;
    const text = String(raw)
      .replace(/,/g, "")
      .replace(/[A$]/g, "")
      .trim();
    if (!text) return null;
    const value = Number(text);
    return Number.isFinite(value) ? value : null;
  }

  function isInvoicePage() {
    const url = location.href;
    const title = document.title || "";
    return url.includes("/invoice") || url.includes("txnId=") || title.toLowerCase().includes("invoice");
  }

  function getInvoiceId() {
    const match = location.href.match(/[?&]txnId=([^&]+)/);
    return match ? decodeURIComponent(match[1]) : null;
  }

  function getInvoiceRoot() {
    const roots = [
      document.querySelector(".trowser-view .body"),
      document.querySelector('[data-automation-id="invoice-form"]'),
      document.querySelector('[data-automation-id="invoice-editor"]'),
      document.querySelector("#qbo-main"),
      document.querySelector("#app"),
      document.querySelector("main"),
      document.body,
    ].filter(Boolean);

    for (const root of roots) {
      const text = root.textContent || "";
      if (text.includes("Product/service") && text.includes("Qty")) {
        return root;
      }
    }

    return document.body;
  }

  function isInvoiceEditorOpen() {
    return isInvoicePage() && !!getInvoiceRoot();
  }

  function getInputValue(el) {
    if (!el) return "";

    if (typeof el.value === "string" && normalizeText(el.value)) {
      return normalizeText(el.value);
    }

    const attrValue = el.getAttribute?.("value");
    if (typeof attrValue === "string" && normalizeText(attrValue)) {
      return normalizeText(attrValue);
    }

    return "";
  }

  function getCellValue(cell) {
    if (!cell) return "";

    const directInput = cell.querySelector('input, textarea, select');
    const directInputValue = getInputValue(directInput);
    if (directInputValue) return directInputValue;

    const combo = cell.querySelector('[role="combobox"]');
    const comboValue = getInputValue(combo);
    if (comboValue) return comboValue;

    const text = normalizeText(cell.textContent || "");
    return text;
  }

  function findInvoiceTable(root) {
    const tables = Array.from(root.querySelectorAll("table"));

    for (const table of tables) {
      const headers = Array.from(table.querySelectorAll("thead th")).map((th) =>
        normalizeText(th.textContent)
      );

      const hasProduct = headers.some((h) => /product\s*\/?\s*service/i.test(h));
      const hasDescription = headers.some((h) => /^description$/i.test(h));
      const hasQty = headers.some((h) => /^qty$/i.test(h) || /^quantity$/i.test(h));

      if (hasProduct && hasDescription && hasQty) {
        return table;
      }
    }

    return null;
  }

  function buildHeaderIndexMap(table) {
    const headers = Array.from(table.querySelectorAll("thead th"));
    const map = {};

    headers.forEach((th, index) => {
      const label = normalizeText(th.textContent).toLowerCase();

      if (/product\s*\/?\s*service/.test(label)) map.product = index;
      else if (label === "sku") map.sku = index;
      else if (label === "description") map.description = index;
      else if (label === "qty" || label === "quantity") map.qty = index;
      else if (label === "rate") map.rate = index;
      else if (label === "amount") map.amount = index;
      else if (label === "gst") map.gst = index;
      else if (label === "service date") map.serviceDate = index;
    });

    return map;
  }

  function extractCustomFieldValue(root, labelText) {
    const fields = Array.from(root.querySelectorAll(".custom-form-field"));

    for (const field of fields) {
      const label = normalizeText(
        field.querySelector(".ReadAndWriteCustomFieldStyles__RethinkCFLabel-kivstf-6")?.textContent || ""
      );

      if (label.toLowerCase() === labelText.toLowerCase()) {
        const input = field.querySelector("input, textarea, select");
        return getInputValue(input);
      }
    }

    return "";
  }

  function extractHeaderData(root) {
    const billingAddress =
      getInputValue(root.querySelector('textarea[aria-label="billToTextAreaLabel"]')) || "N/A";

    const shippingAddress =
      getInputValue(root.querySelector('textarea[aria-label="shipToTextAreaLabel"]')) || "N/A";

    const invoiceNumber =
      normalizeText(
        root.querySelector('[data-automation-id="readonly_reference_number"] span')?.textContent ||
        root.querySelector('[data-automation-id="readonly_reference_number"]')?.textContent ||
        ""
      ) || "N/A";

    const invoiceDate =
      getInputValue(root.querySelector('input[data-testid="txn_date"]')) || "N/A";

    return {
      billingAddress,
      shippingAddress,
      invoiceNumber,
      invoiceDate,
      orderNumber: extractCustomFieldValue(root, "ORDER NUMBER"),
      jobName: extractCustomFieldValue(root, "JOB NAME"),
      phoneNumber: extractCustomFieldValue(root, "Phone"),
    };
  }

  function isCategoryRow(descriptionText) {
    return /^\*+\s*.+?\s*\*+$/.test((descriptionText || "").trim());
  }

  function extractRows(root) {
    const table = findInvoiceTable(root);
    if (!table) {
      warn("Invoice table not found");
      return [];
    }

    const headerMap = buildHeaderIndexMap(table);
    log("Header map:", headerMap);

    if (
      headerMap.product == null ||
      headerMap.sku == null ||
      headerMap.description == null ||
      headerMap.qty == null
    ) {
      warn("Required columns not found", headerMap);
      return [];
    }

    const tbody =
      table.querySelector('tbody[data-smart-table-body="true"]') ||
      table.querySelector("tbody");

    if (!tbody) {
      warn("Invoice tbody not found");
      return [];
    }

    const rowEls = Array.from(tbody.querySelectorAll('tr[data-automation-id^="line "]'));
    log("Found line rows:", rowEls.length);

    const rows = [];

    for (const tr of rowEls) {
      const cells = Array.from(tr.querySelectorAll("td"));
      if (!cells.length) continue;

      const productName = getCellValue(cells[headerMap.product]);
      const sku = getCellValue(cells[headerMap.sku]);
      const description = getCellValue(cells[headerMap.description]);
      const quantity = tryParseNumber(getCellValue(cells[headerMap.qty])) ?? 0;

      if (!productName && !sku && !description) continue;

      rows.push({
        key: `${productName}|${sku}|${description}|${quantity}`,
        productName,
        sku,
        description,
        quantity,
      });
    }

    log("Extracted rows:", rows);
    return rows;
  }

  function extractData() {
    const root = getInvoiceRoot();
    return {
      ...extractHeaderData(root),
      rows: extractRows(root),
    };
  }

  function stableRowsSignature(rows) {
    return JSON.stringify(
      rows.map((r) => ({
        productName: r.productName,
        sku: r.sku,
        description: r.description,
        quantity: r.quantity,
      }))
    );
  }

  async function getStableExtractedData() {
    let previousSig = null;
    let latest = extractData();

    for (let i = 0; i < CONFIG.stableReadAttempts; i++) {
      latest = extractData();
      const sig = stableRowsSignature(latest.rows);

      if (latest.rows.length > 0 && sig === previousSig) {
        return latest;
      }

      previousSig = sig;
      await sleep(CONFIG.stableReadDelayMs);
    }

    return latest;
  }

  function createButton(id, text, clickHandler, left) {
    const button = document.createElement("button");
    button.id = id;
    button.type = "button";
    button.textContent = text;
    button.style.cssText = `
      position: fixed;
      bottom: 4px;
      left: ${left};
      padding: 8px 22px;
      background-color: #2ca01c;
      color: white;
      border: none;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.08);
      cursor: pointer;
      font-size: 14px;
      font-weight: 600;
      z-index: 10000;
      transition: all 0.2s ease;
    `;

    button.addEventListener("mouseenter", () => {
      button.style.backgroundColor = "#248f17";
      button.style.transform = "translateY(-2px)";
      button.style.boxShadow = "0 4px 12px rgba(0,0,0,0.12)";
    });

    button.addEventListener("mouseleave", () => {
      button.style.backgroundColor = "#2ca01c";
      button.style.transform = "translateY(0)";
      button.style.boxShadow = "0 2px 8px rgba(0,0,0,0.08)";
    });

    button.addEventListener("click", clickHandler);
    return button;
  }

  function removeButtons() {
    document.getElementById("custom-print-button")?.remove();
    document.getElementById("custom-pick-slip-button")?.remove();
  }

  async function addButtons() {
    if (STATE.addButtonsInFlight) return;
    STATE.addButtonsInFlight = true;

    try {
      if (!isInvoicePage() || !isInvoiceEditorOpen()) {
        removeButtons();
        return;
      }

      const invoiceId = getInvoiceId();
      if (invoiceId && invoiceId !== STATE.currentInvoiceId) {
        STATE.currentInvoiceId = invoiceId;
        removeButtons();
      }

      if (!document.getElementById("custom-print-button")) {
        document.body.appendChild(
          createButton("custom-print-button", "🖨️ Print", async () => {
            await generateProductTable(false);
          }, "15%")
        );
      }

      if (!document.getElementById("custom-pick-slip-button")) {
        document.body.appendChild(
          createButton("custom-pick-slip-button", "📋 Pick Slip", async () => {
            await generateProductTable(true);
          }, "calc(15% + 140px)")
        );
      }
    } finally {
      STATE.addButtonsInFlight = false;
    }
  }

  function buildProductTable(rows, combineQuantities) {
    if (!rows.length) return "<p>No valid products found.</p>";

    let html = "";

    if (combineQuantities) {
      // Pick Slip:
      // - Product Name = Product/service only
      // - SKU = SKU column
      // - Quantity = Qty column
      // - Exclude category rows like ** Kitchen **
      const grouped = new Map();

      rows.forEach((row) => {
        const qty = Number(row.quantity || 0);

        if (isCategoryRow(row.description) && !row.productName && !row.sku) return;
        if (!row.productName && !row.sku) return;

        const groupKey = row.sku
          ? `SKU:${row.sku}`
          : `NAME:${row.productName}`;

        if (!grouped.has(groupKey)) {
          grouped.set(groupKey, {
            productName: row.productName || "",
            sku: row.sku || "",
            quantity: qty,
          });
        } else {
          grouped.get(groupKey).quantity += qty;
        }
      });

      grouped.forEach((value) => {
        html += `
          <tr style="height:30px;">
            <td>${escapeHtml(value.productName)}</td>
            <td>${escapeHtml(value.sku)}</td>
            <td style="text-align:right;">${escapeHtml(value.quantity)}</td>
          </tr>
        `;
      });
    } else {
      // Print:
      // - normally use Product/service
      // - only use Description when it is a category row like ** Kitchen **
      rows.forEach((row) => {
        let displayName = row.productName || "";

        if (!displayName && isCategoryRow(row.description)) {
          displayName = row.description;
        }

        if (!displayName && !row.sku) return;

        html += `
          <tr style="height:30px;">
            <td>${escapeHtml(displayName)}</td>
            <td>${escapeHtml(row.sku || "")}</td>
            <td style="text-align:right;">${escapeHtml(row.quantity || "")}</td>
          </tr>
        `;
      });
    }

    if (!html) return "<p>No valid products found.</p>";

    return `
      <table class="product-table">
        <thead style="background:lightgrey;">
          <tr>
            <th>Product Name</th>
            <th>SKU</th>
            <th>Quantity</th>
          </tr>
        </thead>
        <tbody>${html}</tbody>
      </table>
    `;
  }

  function generatePrintLayout(data, productTable) {
    const orderNumber = data.orderNumber || "";
    const phoneNumber = data.phoneNumber || "";

    return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Invoice ${escapeHtml(data.invoiceNumber)}</title>
  <style>
    * { box-sizing: border-box; }
    body { font-family: Arial, sans-serif; line-height: 1.6; margin: 0; padding: 0; }
    .maincontainer { width: 100%; padding: 20px; margin: 0 auto; line-height: 1.4; }
    .maincontainer h3 { margin-top: 0; margin-bottom: 0; }
    .ContactService textarea { line-height: 1.65; margin-bottom: -5px; }
    .TMheader { width: 100%; display: flex; justify-content: space-between; margin-bottom: 0px; }
    .TMheader-left { width: 70%; display: flex; font-size: 13px; gap: 20px; }
    .TMheader-right { width: 30%; font-size: 10px; }
    .deliveryNote { display: flex; justify-content: space-between; margin-bottom: 10px; font-size: 13px; width: 100%; }
    .deliveryNote p { margin-bottom: 0; white-space: pre-line; }
    .deliveryNote > div { width: 33.3%; white-space: pre-line; padding-right: 15px; }
    .orderProducts { margin-top: 20px; border-top: 1px solid #000; }
    .product-table { text-align: left; width: 100%; margin-top: 20px; font-size: 14px; border-collapse: collapse; }
    .product-table th { padding: 8px; text-align: left; }
    .product-table td { padding: 6px; vertical-align: top; }
    .product-table tbody tr td:last-child { text-align: right; }
    .orderNote { display: flex; justify-content: space-between; font-size: 14px; margin: 12px 0; }
    .input-group { display: flex; justify-content: space-between; margin-top: 20px; font-size: 12px; }
    hr { border: none; border-top: 1px solid #ccc; margin: 10px 0; }
  </style>
</head>
<body>
  <div class="maincontainer">
    <div class="TMheader">
      <div class="TMheader-left">
        <div class="CompanyInfo">
          <div><b>Adelaide Bathroom & Kitchen Supplies</b></div>
          <div>2/831 Lower North East Rd, Dernancourt</div>
          <div>(08) 7006 5181</div>
          <div>Sales@abksupplies.com.au</div>
          <div>ABN 13 695 032 804</div>
        </div>
      </div>
      <div class="TMheader-right">
        <div class="ContactService">
          <div><b>Received In Good Order & Condition</b></div>
          <div><textarea rows="1" style="height:30px;width:200px;" placeholder="Name:"></textarea></div>
          <div><textarea rows="1" style="height:40px;width:200px;" placeholder="Sign:"></textarea></div>
          <div><textarea rows="1" style="height:30px;width:200px;" placeholder="Date: __ / __ / ____"></textarea></div>
        </div>
      </div>
    </div>

    <h3 style="margin-bottom:10px">Delivery Note</h3>

    <div class="deliveryNote">
      <div><b>INVOICE TO</b><p>${escapeHtml(data.billingAddress)}</p></div>
      <div><b>SHIP TO</b><p>${escapeHtml(data.shippingAddress)}</p></div>
      <div><b>INVOICE NO.:</b><span>${escapeHtml(data.invoiceNumber)}</span><br/><b>DATE:</b><span>${escapeHtml(data.invoiceDate)}</span></div>
    </div>

    <hr>

    <div class="orderNote">
      <div><b>ORDER NUMBER</b><br/>${escapeHtml(orderNumber)}</div>
      <div><b>JOB NAME</b><br/>${escapeHtml(data.jobName || "")}</div>
      <div><b>PHONE</b><br/>${escapeHtml(phoneNumber)}</div>
    </div>

    <div class="orderProducts">${productTable}</div>

    <hr/>

    <div class="input-group">
      <span>Picked By: _______________</span>
      <span>Checked By: _______________</span>
    </div>
  </div>
</body>
</html>
    `;
  }

  async function generateProductTable(combineQuantities) {
    const data = await getStableExtractedData();

    if (!data.rows.length) {
      alert("No line items found in QuickBooks invoice table.");
      return;
    }

    const productTable = buildProductTable(data.rows, combineQuantities);
    const printLayout = generatePrintLayout(data, productTable);

    const newWindow = window.open("", "_blank", "width=900,height=700");
    if (!newWindow) {
      alert("Popup blocked. Please allow popups for QuickBooks.");
      return;
    }

    newWindow.document.write(printLayout);
    newWindow.document.close();

    newWindow.onload = () => {
      setTimeout(() => {
        newWindow.focus();
        newWindow.print();
      }, 300);
    };
  }

  function refreshButtons() {
    clearTimeout(STATE.mutationTimer);
    STATE.mutationTimer = setTimeout(() => {
      if (isInvoicePage() && isInvoiceEditorOpen()) addButtons();
      else removeButtons();
    }, CONFIG.mutationDebounceMs);
  }

  function setupObservers() {
    const originalPushState = history.pushState;
    history.pushState = function () {
      const result = originalPushState.apply(this, arguments);
      setTimeout(refreshButtons, 500);
      return result;
    };

    const originalReplaceState = history.replaceState;
    history.replaceState = function () {
      const result = originalReplaceState.apply(this, arguments);
      setTimeout(refreshButtons, 500);
      return result;
    };

    window.addEventListener("popstate", () => {
      setTimeout(refreshButtons, 500);
    });

    const observer = new MutationObserver(() => {
      refreshButtons();
    });

    observer.observe(document.body, { childList: true, subtree: true });
  }

  setInterval(() => {
    if (isInvoicePage() && isInvoiceEditorOpen()) addButtons();
    else removeButtons();
  }, CONFIG.buttonCheckIntervalMs);

  setupObservers();

  setTimeout(() => {
    if (isInvoicePage() && isInvoiceEditorOpen()) addButtons();
  }, CONFIG.initialBootDelayMs);
})();