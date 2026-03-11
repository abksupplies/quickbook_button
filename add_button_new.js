// ==UserScript==
// @name         QuickBooks Invoice Print + Pick Slip (Stable Version)
// @namespace    http://tampermonkey.net/
// @version      2.9
// @description  Adds reliable Print and Pick Slip buttons to QuickBooks invoices with stable row extraction and safer quantity handling
// @author       Raj - Gorkhari
// @match        https://qbo.intuit.com/*
// @include      https://qbo.intuit.com/app/invoice?*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
  "use strict";

  const CONFIG = {
    buttonCheckIntervalMs: 2500,
    mutationDebounceMs: 600,
    initialBootDelayMs: 2000,
    rootWaitTimeoutMs: 15000,
    stableReadAttempts: 8,
    stableReadDelayMs: 350,
    rowReadyTimeoutMs: 12000,
    rowReadyPollMs: 250,
    debug: true,
  };

  const STATE = {
    currentInvoiceId: null,
    addButtonsInFlight: false,
    buttonsMounted: false,
    mutationDebounceTimer: null,
    observer: null,
    lastUrl: location.href,
  };

  // ---------------------------
  // Logging
  // ---------------------------

  function log(...args) {
    if (CONFIG.debug) console.log("[QBO Pick Slip]", ...args);
  }

  function warn(...args) {
    console.warn("[QBO Pick Slip]", ...args);
  }

  function error(...args) {
    console.error("[QBO Pick Slip]", ...args);
  }

  // ---------------------------
  // Utilities
  // ---------------------------

  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function normalizeText(value) {
    return String(value ?? "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function isVisible(el) {
    if (!el || !document.contains(el)) return false;
    const style = window.getComputedStyle(el);
    if (style.display === "none" || style.visibility === "hidden" || style.opacity === "0") {
      return false;
    }
    const rect = el.getBoundingClientRect();
    return rect.width > 0 && rect.height > 0;
  }

  function tryParseNumber(raw) {
    if (raw == null) return null;
    const text = String(raw).replace(/,/g, "").trim();
    if (!text) return null;
    const value = Number(text);
    return Number.isFinite(value) ? value : null;
  }

  function getElementValue(el) {
    if (!el) return "";
    if (typeof el.value === "string" && el.value.trim()) return el.value.trim();
    const attrValue = el.getAttribute?.("value");
    if (typeof attrValue === "string" && attrValue.trim()) return attrValue.trim();
    const text = el.textContent;
    if (typeof text === "string" && text.trim()) return text.trim();
    return "";
  }

  function stableStringifyRows(rows) {
    return JSON.stringify(
      rows.map((row) => ({
        key: row.key,
        productName: row.productName,
        description: row.description,
        sku: row.sku,
        quantity: row.quantity,
      }))
    );
  }

  // ---------------------------
  // Page / Context Detection
  // ---------------------------

  function isInvoicePage() {
    const url = window.location.href;
    return (
      url.includes("qbo.intuit.com/app/invoice") ||
      url.includes("/app/invoice?") ||
      url.includes("/invoice?txnId") ||
      url.includes("/invoice")
    );
  }

  function getInvoiceId() {
    const match = window.location.href.match(/[?&]txnId=([^&]+)/);
    return match ? decodeURIComponent(match[1]) : null;
  }

  function getInvoiceRoot() {
    const overlayRoot = document.querySelector(".trowser-view .body");
    if (overlayRoot && overlayRoot.children.length > 0 && isVisible(overlayRoot)) {
      return overlayRoot;
    }

    const candidates = [
      document.querySelector('[data-automation-id="invoice-form"]'),
      document.querySelector('[data-automation-id="invoice-editor"]'),
      document.querySelector(".invoice-content"),
      document.querySelector("#qbo-main"),
      document.querySelector("#app"),
      document.body,
    ].filter(Boolean);

    for (const candidate of candidates) {
      if (
        candidate.querySelector(".dgrid-row") ||
        candidate.querySelector('[data-qbo-bind="text: referenceNumber"]') ||
        candidate.querySelector("textarea.topFieldInput.address") ||
        candidate.querySelector("#shippingAddress") ||
        candidate.querySelector(".custom-form")
      ) {
        return candidate;
      }
    }

    return null;
  }

  function isInvoiceEditorOpen() {
    const root = getInvoiceRoot();
    return !!(root && isVisible(root));
  }

  async function waitForInvoiceRoot(timeoutMs = CONFIG.rootWaitTimeoutMs) {
    const start = Date.now();

    while (Date.now() - start < timeoutMs) {
      const root = getInvoiceRoot();
      if (root && isVisible(root)) return root;
      await sleep(200);
    }

    return null;
  }

  // ---------------------------
  // Row Extraction
  // ---------------------------

  function getRowElements(root) {
    if (!root) return [];
    return Array.from(root.querySelectorAll(".dgrid-row")).filter((row) => document.contains(row));
  }

  function getSKU(row) {
    const selectors = [
      ".field-sku",
      '[data-automation-id="sku"]',
      ".sku-field",
      ".itemSKU",
      '[aria-label*="SKU"]',
      '[data-testid*="sku"]',
    ];

    for (const selector of selectors) {
      const el = row.querySelector(selector);
      const value = normalizeText(getElementValue(el));
      if (value) return value;
    }

    return "";
  }

  function getQuantity(row) {
    const selectors = [
      ".field-quantity-inner",
      '[data-automation-id="quantity"]',
      ".quantity-field",
      'input[aria-label*="Quantity"]',
      'input[name*="quantity"]',
      '[data-testid*="quantity"]',
      ".field-qty",
    ];

    for (const selector of selectors) {
      const el = row.querySelector(selector);
      if (!el) continue;

      const value = getElementValue(el);
      const parsed = tryParseNumber(value);
      if (parsed != null) return parsed;
    }

    // Fallback: search any input-like elements in the row that look like quantity
    const fallbackInputs = Array.from(row.querySelectorAll("input, [contenteditable='true'], span, div"));
    for (const el of fallbackInputs) {
      const aria = (el.getAttribute?.("aria-label") || "").toLowerCase();
      const name = (el.getAttribute?.("name") || "").toLowerCase();
      const cls = (el.className || "").toString().toLowerCase();

      const looksLikeQuantity =
        aria.includes("quantity") || name.includes("quantity") || cls.includes("quantity") || cls.includes("qty");

      if (!looksLikeQuantity) continue;

      const value = getElementValue(el);
      const parsed = tryParseNumber(value);
      if (parsed != null) return parsed;
    }

    return 0;
  }

  function getProductName(row) {
    const selectors = [
      ".itemColumn",
      '[data-automation-id="item-name"]',
      ".field-product-service .itemColumn",
      ".itemName",
    ];

    for (const selector of selectors) {
      const el = row.querySelector(selector);
      const value = normalizeText(getElementValue(el));
      if (value) return value;
    }

    return "";
  }

  function getDescription(row) {
    const selectors = [
      ".field-description div",
      ".field-description",
      '[data-automation-id="description"]',
      ".description-field",
    ];

    for (const selector of selectors) {
      const el = row.querySelector(selector);
      const value = normalizeText(getElementValue(el));
      if (value) return value;
    }

    return "";
  }

  function isLikelyDataRow(row) {
    const productName = getProductName(row);
    const description = getDescription(row);
    const sku = getSKU(row);
    const quantity = getQuantity(row);

    return Boolean(productName || description || sku || quantity);
  }

  function extractHeaderData(root) {
    const data = {
      billingAddress: getElementValue(root.querySelector("textarea.topFieldInput.address")) || "N/A",
      shippingAddress: getElementValue(root.querySelector("#shippingAddress")) || "N/A",
      invoiceNumber:
        normalizeText(getElementValue(root.querySelector('[data-qbo-bind="text: referenceNumber"]'))) || "N/A",
      invoiceDate: getElementValue(root.querySelector(".dijitDateTextBox input.dijitInputInner")) || "N/A",
      orderNumber: "",
      jobName: "",
      phoneNumber: "",
    };

    const formElement = root.querySelector(".custom-form");
    if (formElement) {
      const formFields = Array.from(formElement.querySelectorAll(".custom-form-field"));

      const findFieldValue = (labelText) => {
        const field = formFields.find((f) => {
          const label = normalizeText(f.querySelector("label")?.textContent);
          return label.toLowerCase() === labelText.toLowerCase();
        });
        return field ? getElementValue(field.querySelector("input, textarea, select")) : "";
      };

      data.orderNumber = findFieldValue("ORDER NUMBER");
      data.jobName = findFieldValue("JOB NAME");
      data.phoneNumber = findFieldValue("Phone");
    }

    return data;
  }

  function extractRows(root) {
    const rowElements = getRowElements(root);
    const rows = [];

    for (const rowEl of rowElements) {
      if (!isLikelyDataRow(rowEl)) continue;

      const productName = getProductName(rowEl);
      const description = getDescription(rowEl);
      const sku = getSKU(rowEl);
      const quantity = getQuantity(rowEl);

      // Skip obviously incomplete ghost rows
      if (!productName && !description && !sku) continue;

      const key = sku || `NAME:${productName}` || `DESC:${description}`;

      rows.push({
        key,
        productName,
        description,
        sku,
        quantity,
      });
    }

    return rows;
  }

  function extractData() {
    const root = getInvoiceRoot();

    if (!root) {
      error("Invoice root not found during extraction");
      return {
        billingAddress: "N/A",
        shippingAddress: "N/A",
        invoiceNumber: "N/A",
        invoiceDate: "N/A",
        orderNumber: "",
        jobName: "",
        phoneNumber: "",
        rows: [],
      };
    }

    const header = extractHeaderData(root);
    const rows = extractRows(root);

    return {
      ...header,
      rows,
    };
  }

  // ---------------------------
  // Stable Read / Readiness
  // ---------------------------

  function hasUsableRows(data) {
    if (!data || !Array.isArray(data.rows) || data.rows.length === 0) return false;

    return data.rows.some((row) => {
      const hasIdentity = Boolean(row.productName || row.description || row.sku);
      const hasQty = typeof row.quantity === "number" && row.quantity >= 0;
      return hasIdentity && hasQty;
    });
  }

  async function waitForRowsReady(timeoutMs = CONFIG.rowReadyTimeoutMs) {
    const start = Date.now();

    while (Date.now() - start < timeoutMs) {
      const data = extractData();

      if (hasUsableRows(data)) {
        // One extra stability delay to avoid mid-render scrape
        await sleep(250);
        const secondData = extractData();

        if (
          hasUsableRows(secondData) &&
          stableStringifyRows(data.rows) === stableStringifyRows(secondData.rows)
        ) {
          return true;
        }
      }

      await sleep(CONFIG.rowReadyPollMs);
    }

    return false;
  }

  async function getStableExtractedData() {
    let previousSignature = null;
    let previousData = null;

    for (let i = 0; i < CONFIG.stableReadAttempts; i++) {
      const data = extractData();
      const signature = stableStringifyRows(data.rows);

      log(`Stable read attempt ${i + 1}/${CONFIG.stableReadAttempts}`, data.rows);

      if (hasUsableRows(data) && signature && previousSignature === signature) {
        return data;
      }

      previousSignature = signature;
      previousData = data;
      await sleep(CONFIG.stableReadDelayMs);
    }

    warn("Returning best-effort data; row state did not fully stabilize.");
    return previousData || extractData();
  }

  // ---------------------------
  // Buttons
  // ---------------------------

  function createButton(id, text, clickHandler, left) {
    const button = document.createElement("button");
    button.id = id;
    button.type = "button";
    button.textContent = text;
    button.style.cssText = `
      position: fixed;
      bottom: 8px;
      left: ${left};
      padding: 12px 22px;
      background-color: #2ca01c;
      color: #fff;
      border: none;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.08);
      cursor: pointer;
      font-size: 14px;
      font-weight: 600;
      z-index: 100000;
      transition: all 0.2s ease;
    `;

    button.addEventListener("mouseenter", () => {
      button.style.backgroundColor = "#248f17";
      button.style.transform = "translateY(-1px)";
      button.style.boxShadow = "0 4px 12px rgba(0,0,0,0.14)";
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
    const ids = ["custom-print-button", "custom-pick-slip-button"];

    for (const id of ids) {
      const el = document.getElementById(id);
      if (el) el.remove();
    }

    STATE.buttonsMounted = false;
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

      if (document.getElementById("custom-print-button") && document.getElementById("custom-pick-slip-button")) {
        STATE.buttonsMounted = true;
        return;
      }

      const root = await waitForInvoiceRoot();
      if (!root) {
        warn("Invoice root not found in time.");
        return;
      }

      const ready = await waitForRowsReady();
      if (!ready) {
        warn("Invoice rows not ready yet.");
        return;
      }

      if (!document.getElementById("custom-print-button")) {
        const printButton = createButton(
          "custom-print-button",
          "🖨️ Print",
          async () => {
            await generateProductTable(false);
          },
          "15%"
        );
        document.body.appendChild(printButton);
      }

      if (!document.getElementById("custom-pick-slip-button")) {
        const pickSlipButton = createButton(
          "custom-pick-slip-button",
          "📋 Pick Slip",
          async () => {
            await generateProductTable(true);
          },
          "calc(15% + 150px)"
        );
        document.body.appendChild(pickSlipButton);
      }

      STATE.buttonsMounted = true;
      log("Buttons added successfully.");
    } catch (err) {
      error("Failed to add buttons:", err);
    } finally {
      STATE.addButtonsInFlight = false;
    }
  }

  // ---------------------------
  // Print / Pick Slip Generation
  // ---------------------------

  function buildProductTable(rows, combineQuantities) {
    if (!Array.isArray(rows) || rows.length === 0) {
      return "<p>No valid products found.</p>";
    }

    let tableRows = "";

    if (combineQuantities) {
      const grouped = new Map();

      for (const row of rows) {
        const identityKey = row.sku || row.productName || row.description;
        if (!identityKey) continue;

        const groupKey = row.sku ? `SKU:${row.sku}` : `NAME:${identityKey}`;

        if (!grouped.has(groupKey)) {
          grouped.set(groupKey, {
            productName: row.productName || row.description || "",
            sku: row.sku || "",
            quantity: Number(row.quantity || 0),
          });
        } else {
          const existing = grouped.get(groupKey);
          existing.quantity += Number(row.quantity || 0);
        }
      }

      for (const [, value] of grouped) {
        tableRows += `
          <tr style="height: 30px;">
            <td>${escapeHtml(value.productName)}</td>
            <td>${escapeHtml(value.sku)}</td>
            <td style="text-align:right;">${escapeHtml(value.quantity)}</td>
          </tr>
        `;
      }
    } else {
      for (const row of rows) {
        const displayName = row.productName || (row.sku ? "" : row.description);
        if (!displayName && !row.sku) continue;

        tableRows += `
          <tr style="height: 30px;">
            <td>${escapeHtml(displayName)}</td>
            <td>${escapeHtml(row.sku)}</td>
            <td style="text-align:right;">${escapeHtml(row.quantity || "")}</td>
          </tr>
        `;
      }
    }

    if (!tableRows) {
      return "<p>No valid products found.</p>";
    }

    return `
      <table class="product-table">
        <thead style="background: lightgrey;">
          <tr>
            <th>Product Name</th>
            <th>SKU</th>
            <th>Quantity</th>
          </tr>
        </thead>
        <tbody>
          ${tableRows}
        </tbody>
      </table>
    `;
  }

  async function generateProductTable(combineQuantities) {
    try {
      const root = await waitForInvoiceRoot();
      if (!root) {
        alert("Invoice was not found. Please wait for the invoice to load and try again.");
        return;
      }

      const ready = await waitForRowsReady();
      if (!ready) {
        alert("Invoice lines are still loading. Please wait a moment and try again.");
        return;
      }

      const data = await getStableExtractedData();

      if (!data || !Array.isArray(data.rows) || data.rows.length === 0) {
        alert("No product data found. Please ensure the invoice is fully loaded.");
        return;
      }

      log("Final extracted data:", data);

      const productTable = buildProductTable(data.rows, combineQuantities);
      const printLayout = generatePrintLayout(data, productTable);

      const printWindow = window.open("", "_blank", "width=1000,height=800,noopener,noreferrer");
      if (!printWindow) {
        alert("Popup blocked. Please allow popups for QuickBooks and try again.");
        return;
      }

      printWindow.document.open();
      printWindow.document.write(printLayout);
      printWindow.document.close();

      // Wait for images/layout before printing
      printWindow.onload = () => {
        setTimeout(() => {
          printWindow.focus();
          printWindow.print();
        }, 350);
      };
    } catch (err) {
      error("Printing failed:", err);
      alert("Printing failed. Check the browser console for details.");
    }
  }

  function generatePrintLayout(data, productTable) {
    return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Invoice ${escapeHtml(data.invoiceNumber)}</title>
  <style>
    * { box-sizing: border-box; }
    body {
      font-family: Arial, sans-serif;
      line-height: 1.45;
      margin: 0;
      padding: 0;
      color: #000;
    }
    .maincontainer {
      width: 100%;
      padding: 20px;
      margin: 0 auto;
    }
    .maincontainer h3 {
      margin-top: 0;
      margin-bottom: 8px;
    }
    .ContactService textarea {
      line-height: 1.5;
      margin-bottom: 4px;
      resize: none;
      font-family: Arial, sans-serif;
    }
    .TMheader {
      width: 100%;
      display: flex;
      justify-content: space-between;
      gap: 16px;
      margin-bottom: 8px;
    }
    .TMheader-left {
      width: 70%;
      display: flex;
      gap: 20px;
      font-size: 13px;
    }
    .TMheader-left img {
      max-width: 150px;
      height: auto;
    }
    .TMheader-right {
      width: 30%;
      font-size: 10px;
    }
    .deliveryNote {
      display: flex;
      justify-content: space-between;
      gap: 16px;
      margin-bottom: 10px;
      font-size: 13px;
      width: 100%;
    }
    .deliveryNote p {
      margin: 4px 0 0;
      white-space: pre-line;
    }
    .deliveryNote > div {
      width: 33.33%;
      white-space: pre-line;
      word-break: break-word;
    }
    .orderProducts {
      margin-top: 18px;
      border-top: 1px solid #000;
      padding-top: 6px;
    }
    .product-table {
      text-align: left;
      width: 100%;
      margin-top: 10px;
      font-size: 13px;
      border-collapse: collapse;
    }
    .product-table th {
      padding: 8px;
      text-align: left;
      border-bottom: 1px solid #bbb;
    }
    .product-table td {
      padding: 6px 8px;
      border-bottom: 1px solid #eee;
      vertical-align: top;
    }
    .product-table tbody tr td:last-child {
      text-align: right;
      white-space: nowrap;
    }
    .orderNote {
      display: flex;
      justify-content: space-between;
      gap: 16px;
      font-size: 13px;
      margin: 10px 0 12px;
    }
    .orderNote > div {
      width: 33.33%;
      word-break: break-word;
    }
    .input-group {
      display: flex;
      justify-content: space-between;
      font-size: 12px;
      margin-top: 16px;
    }
    hr {
      border: none;
      border-top: 1px solid #ccc;
      margin: 10px 0;
    }
    @media print {
      body { margin: 0; }
      .maincontainer { padding: 14px; }
    }
  </style>
</head>
<body>
  <div class="maincontainer">
    <div class="TMheader">
      <div class="TMheader-left">
        <div class="imageLogo">
          <img src="https://c17.qbo.intuit.com/qbo17/ext/Image/show/249862794341519/1?14857441280001" alt="Logo" />
        </div>
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

    <h3>Delivery Note</h3>

    <div class="deliveryNote">
      <div class="deliveryNote-1">
        <b>INVOICE TO</b>
        <p>${escapeHtml(data.billingAddress)}</p>
      </div>
      <div class="deliveryNote-2">
        <b>SHIP TO</b>
        <p>${escapeHtml(data.shippingAddress)}</p>
      </div>
      <div class="deliveryNote-3">
        <b>INVOICE NO.:</b>
        <span>${escapeHtml(data.invoiceNumber)}</span><br/>
        <b>DATE:</b>
        <span>${escapeHtml(data.invoiceDate)}</span>
      </div>
    </div>

    <hr>

    <div class="orderNote">
      <div><b>ORDER NUMBER</b><br/>${escapeHtml(data.orderNumber)}</div>
      <div><b>JOB NAME</b><br/>${escapeHtml(data.jobName)}</div>
      <div><b>PHONE</b><br/>${escapeHtml(data.phoneNumber)}</div>
    </div>

    <div class="orderProducts">
      ${productTable}
    </div>

    <hr>

    <div class="input-group">
      <span>Picked By: _______________</span>
      <span>Checked By: _______________</span>
    </div>
  </div>
</body>
</html>
    `;
  }

  // ---------------------------
  // SPA Navigation / Observers
  // ---------------------------

  function scheduleButtonRefresh() {
    clearTimeout(STATE.mutationDebounceTimer);
    STATE.mutationDebounceTimer = setTimeout(() => {
      if (isInvoicePage() && isInvoiceEditorOpen()) {
        addButtons();
      } else {
        removeButtons();
      }
    }, CONFIG.mutationDebounceMs);
  }

  function patchHistoryMethods() {
    const originalPushState = history.pushState;
    const originalReplaceState = history.replaceState;

    history.pushState = function () {
      const result = originalPushState.apply(this, arguments);
      setTimeout(handleUrlChange, 300);
      return result;
    };

    history.replaceState = function () {
      const result = originalReplaceState.apply(this, arguments);
      setTimeout(handleUrlChange, 300);
      return result;
    };
  }

  function handleUrlChange() {
    if (location.href !== STATE.lastUrl) {
      log("URL changed:", STATE.lastUrl, "=>", location.href);
      STATE.lastUrl = location.href;
      scheduleButtonRefresh();
    }
  }

  function setupObservers() {
    patchHistoryMethods();

    window.addEventListener("popstate", () => {
      setTimeout(handleUrlChange, 300);
    });

    STATE.observer = new MutationObserver(() => {
      scheduleButtonRefresh();
    });

    STATE.observer.observe(document.body, {
      childList: true,
      subtree: true,
    });
  }

  function startPeriodicCheck() {
    setInterval(() => {
      if (isInvoicePage() && isInvoiceEditorOpen()) {
        addButtons();
      } else {
        removeButtons();
      }
    }, CONFIG.buttonCheckIntervalMs);
  }

  // ---------------------------
  // Init
  // ---------------------------

  async function init() {
    setupObservers();
    startPeriodicCheck();

    await sleep(CONFIG.initialBootDelayMs);

    if (isInvoicePage() && isInvoiceEditorOpen()) {
      addButtons();
    }
  }

  init();
})();