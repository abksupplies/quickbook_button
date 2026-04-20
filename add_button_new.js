// ==UserScript==
// @name         QuickBooks Invoice Print + Pick Slip (Stable + Safer)
// @namespace    http://tampermonkey.net/
// @version      3.6
// @description  Adds Print and Pick Slip buttons to QuickBooks invoices with broader UI support
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
    mutationDebounceMs: 700,
    initialBootDelayMs: 2500,
    stableReadAttempts: 7,
    stableReadDelayMs: 350,
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
    return String(value ?? "").replace(/\s+/g, " ").trim();
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

  function tryParseNumber(raw) {
    if (raw == null) return null;
    const text = String(raw).replace(/,/g, "").trim();
    if (!text) return null;
    const value = Number(text);
    return Number.isFinite(value) ? value : null;
  }

  function isInvoicePage() {
    const url = location.href;
    return url.includes("/invoice") || url.includes("txnId=");
  }

  function getInvoiceId() {
    const match = location.href.match(/[?&]txnId=([^&]+)/);
    return match ? decodeURIComponent(match[1]) : null;
  }

  function getInvoiceRoot() {
    const selectors = [
      ".trowser-view .body",
      '[data-automation-id="invoice-form"]',
      '[data-automation-id="invoice-editor"]',
      '[data-testid*="invoice"]',
      '[class*="invoice"]',
      "#qbo-main",
      "#app",
      "main",
      "body",
    ];

    for (const selector of selectors) {
      const el = document.querySelector(selector);
      if (!el) continue;

      const text = el.textContent || "";
      if (
        text.includes("Receive payment") ||
        text.includes("BALANCE DUE") ||
        text.includes("Amounts are") ||
        text.includes("ORDER NUMBER") ||
        text.includes("JOB NAME") ||
        text.includes("Invoice")
      ) {
        return el;
      }
    }

    return document.body;
  }

  function isInvoiceEditorOpen() {
    return isInvoicePage() && !!getInvoiceRoot();
  }

  function getProductNameOld(row) {
    const selectors = [
      ".itemColumn",
      '[data-automation-id="item-name"]',
      ".itemName",
      '[data-testid*="item"]',
    ];
    for (const selector of selectors) {
      const el = row.querySelector(selector);
      const value = normalizeText(getElementValue(el));
      if (value) return value;
    }
    return "";
  }

  function getDescriptionOld(row) {
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

  function getSKUOld(row) {
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

  function getQuantityOld(row) {
    const selectors = [
      ".field-quantity-inner",
      '[data-automation-id="quantity"]',
      ".quantity-field",
      ".field-qty",
      'input[aria-label*="Quantity"]',
      'input[name*="quantity"]',
      '[data-testid*="quantity"]',
    ];

    for (const selector of selectors) {
      const el = row.querySelector(selector);
      if (!el) continue;
      const parsed = tryParseNumber(getElementValue(el));
      if (parsed != null) return parsed;
    }

    const allCandidates = Array.from(row.querySelectorAll("input, span, div"));
    for (const el of allCandidates) {
      const cls = String(el.className || "").toLowerCase();
      const aria = String(el.getAttribute?.("aria-label") || "").toLowerCase();
      const name = String(el.getAttribute?.("name") || "").toLowerCase();

      if (
        cls.includes("quantity") ||
        cls.includes("qty") ||
        aria.includes("quantity") ||
        name.includes("quantity")
      ) {
        const parsed = tryParseNumber(getElementValue(el));
        if (parsed != null) return parsed;
      }
    }

    return 0;
  }

  function isHeaderOnlyRow(productName, description, sku, quantity) {
    const text = (productName || description || "").trim();
    if (sku) return false;
    if (Number(quantity || 0) !== 0) return false;
    if (!text) return true;
    if (/^\*+\s*.+?\s*\*+$/.test(text)) return true;
    return false;
  }

  function extractHeaderData(root) {
    const text = root?.textContent || "";

    const invoiceNumber =
      normalizeText(getElementValue(root.querySelector('[data-qbo-bind="text: referenceNumber"]'))) ||
      (text.match(/Invoice no\.?\s*([A-Za-z0-9-]+)/i)?.[1] ?? "N/A");

    return {
      billingAddress: getElementValue(root.querySelector("textarea.topFieldInput.address")) || "N/A",
      shippingAddress: getElementValue(root.querySelector("#shippingAddress")) || "N/A",
      invoiceNumber,
      invoiceDate: getElementValue(root.querySelector(".dijitDateTextBox input.dijitInputInner")) || "N/A",
      orderNumber: "",
      jobName: "",
      phoneNumber: "",
    };
  }

  function extractRowsFromOldUI(root) {
    const rowElements = Array.from(root.querySelectorAll(".dgrid-row"));
    const rows = [];

    rowElements.forEach((row) => {
      const productName = getProductNameOld(row);
      const description = getDescriptionOld(row);
      const sku = getSKUOld(row);
      const quantity = getQuantityOld(row);

      if (!productName && !description && !sku) return;

      rows.push({
        key: sku || productName || description,
        productName,
        description,
        sku,
        quantity,
      });
    });

    return rows;
  }

  function scoreHeader(text) {
    const t = normalizeText(text).toLowerCase();

    if (!t) return null;
    if (t.includes("product") || t.includes("service") || t === "item") return "product";
    if (t === "sku" || t.includes("sku")) return "sku";
    if (t === "quantity" || t === "qty" || t.includes("quantity")) return "quantity";
    if (t.includes("description")) return "description";
    return null;
  }

  function extractRowsFromGenericGrid(root) {
    const tables = Array.from(root.querySelectorAll("table"));
    const results = [];

    for (const table of tables) {
      const headerCells = Array.from(table.querySelectorAll("thead th, tr th"));
      if (!headerCells.length) continue;

      const headerMap = {};
      headerCells.forEach((cell, idx) => {
        const role = scoreHeader(cell.textContent);
        if (role) headerMap[role] = idx;
      });

      if (headerMap.product == null && headerMap.sku == null && headerMap.quantity == null) {
        continue;
      }

      const bodyRows = Array.from(table.querySelectorAll("tbody tr, tr")).filter((tr) => tr.querySelectorAll("td").length);
      for (const tr of bodyRows) {
        const cells = Array.from(tr.querySelectorAll("td"));
        if (!cells.length) continue;

        const productName = headerMap.product != null ? normalizeText(cells[headerMap.product]?.textContent) : "";
        const description = headerMap.description != null ? normalizeText(cells[headerMap.description]?.textContent) : "";
        const sku = headerMap.sku != null ? normalizeText(cells[headerMap.sku]?.textContent) : "";
        const quantity = headerMap.quantity != null ? (tryParseNumber(cells[headerMap.quantity]?.textContent) ?? 0) : 0;

        if (!productName && !description && !sku) continue;

        results.push({
          key: sku || productName || description,
          productName,
          description,
          sku,
          quantity,
        });
      }

      if (results.length) return results;
    }

    return [];
  }

  function extractRowsFromColumnIdGrid(root) {
    const rows = Array.from(root.querySelectorAll('tr, [role="row"]'));
    const results = [];

    for (const row of rows) {
      const productCell =
        row.querySelector('[data-column-id="productName"]') ||
        row.querySelector('[data-column-id="productService"]') ||
        row.querySelector('[data-column-id="itemName"]') ||
        row.querySelector('[data-column-id="name"]');

      const skuCell =
        row.querySelector('[data-column-id="sku"]') ||
        row.querySelector('[data-column-id="itemSku"]');

      const qtyCell =
        row.querySelector('[data-column-id="quantity"]') ||
        row.querySelector('[data-column-id="qty"]');

      const descCell =
        row.querySelector('[data-column-id="description"]');

      const productName = normalizeText(getElementValue(productCell));
      const description = normalizeText(getElementValue(descCell));
      const sku = normalizeText(getElementValue(skuCell));
      const quantity = tryParseNumber(getElementValue(qtyCell)) ?? 0;

      if (!productName && !description && !sku) continue;

      results.push({
        key: sku || productName || description,
        productName,
        description,
        sku,
        quantity,
      });
    }

    return results;
  }

  function extractRows() {
    const root = getInvoiceRoot();

    let rows = extractRowsFromOldUI(root);
    if (rows.length) {
      log("Rows found using old UI selectors:", rows.length);
      return rows;
    }

    rows = extractRowsFromColumnIdGrid(root);
    if (rows.length) {
      log("Rows found using data-column-id grid:", rows.length);
      return rows;
    }

    rows = extractRowsFromGenericGrid(root);
    if (rows.length) {
      log("Rows found using generic table/grid:", rows.length);
      return rows;
    }

    log("No rows found with current selectors.");
    return [];
  }

  function extractData() {
    const root = getInvoiceRoot();

    return {
      ...extractHeaderData(root),
      rows: extractRows(),
    };
  }

  function stableRowsSignature(rows) {
    return JSON.stringify(
      rows.map((r) => ({
        key: r.key,
        sku: r.sku,
        productName: r.productName,
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
      bottom: 10px;
      left: ${left};
      padding: 10px 18px;
      background-color: #2ca01c;
      color: white;
      border: none;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.15);
      cursor: pointer;
      font-size: 14px;
      font-weight: 600;
      z-index: 999999;
    `;
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
      const grouped = new Map();

      rows.forEach((row) => {
        const qty = Number(row.quantity || 0);

        if (isHeaderOnlyRow(row.productName, row.description, row.sku, row.quantity)) return;
        if (!row.sku && qty === 0) return;

        const groupKey = row.sku ? `SKU:${row.sku}` : `NAME:${row.productName || row.description}`;

        if (!grouped.has(groupKey)) {
          grouped.set(groupKey, {
            productName: row.productName || row.description || "",
            sku: row.sku || "",
            quantity: qty,
          });
        } else {
          grouped.get(groupKey).quantity += qty;
        }
      });

      grouped.forEach((value) => {
        html += `
          <tr>
            <td>${escapeHtml(value.productName)}</td>
            <td>${escapeHtml(value.sku)}</td>
            <td style="text-align:right;">${escapeHtml(value.quantity)}</td>
          </tr>
        `;
      });
    } else {
      rows.forEach((row) => {
        const displayName = row.productName || (row.sku ? "" : row.description);
        if (displayName || row.sku) {
          html += `
            <tr>
              <td>${escapeHtml(displayName)}</td>
              <td>${escapeHtml(row.sku)}</td>
              <td style="text-align:right;">${escapeHtml(row.quantity || "")}</td>
            </tr>
          `;
        }
      });
    }

    if (!html) return "<p>No valid products found.</p>";

    return `
      <table class="product-table">
        <thead>
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

  async function generateProductTable(combineQuantities) {
    const data = await getStableExtractedData();

    if (!data.rows.length) {
      alert("No product rows found in the current QuickBooks UI. Open DevTools and check console logs starting with [QBO Pick Slip].");
      return;
    }

    const productTable = buildProductTable(data.rows, combineQuantities);

    const printLayout = `
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Invoice ${escapeHtml(data.invoiceNumber)}</title>
<style>
body { font-family: Arial, sans-serif; padding: 20px; }
.product-table { width: 100%; border-collapse: collapse; margin-top: 20px; }
.product-table th, .product-table td { padding: 8px; border-bottom: 1px solid #ddd; text-align: left; }
.product-table td:last-child { text-align: right; }
</style>
</head>
<body>
<h3>Delivery Note</h3>
<div><b>Invoice No:</b> ${escapeHtml(data.invoiceNumber)}</div>
<div><b>Date:</b> ${escapeHtml(data.invoiceDate)}</div>
<div style="margin-top:20px;">${productTable}</div>
</body>
</html>
    `;

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