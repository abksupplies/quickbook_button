// ==UserScript==
// @name         QuickBooks Invoice Print + Pick Slip (Stable + Safer)
// @namespace    http://tampermonkey.net/
// @version      3.3
// @description  Adds Print and Pick Slip buttons to QuickBooks invoices with safer quantity handling and stable extraction
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
    initialBootDelayMs: 1800,
    stableReadAttempts: 6,
    stableReadDelayMs: 250,
    debug: true,
  };

  const STATE = {
    currentInvoiceId: null,
    addButtonsInFlight: false,
    mutationTimer: null,
    lastUrl: location.href,
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
    const candidates = [
      document.querySelector(".trowser-view .body"),
      document.querySelector('[data-automation-id="invoice-form"]'),
      document.querySelector('[data-automation-id="invoice-editor"]'),
      document.querySelector(".invoice-content"),
      document.querySelector("#qbo-main"),
      document.querySelector("#app"),
      document.body,
    ].filter(Boolean);

    for (const c of candidates) {
      if (
        c.querySelector(".dgrid-row") ||
        c.querySelector('[data-qbo-bind="text: referenceNumber"]') ||
        c.querySelector("textarea.topFieldInput.address") ||
        c.querySelector("#shippingAddress") ||
        c.querySelector(".custom-form")
      ) {
        return c;
      }
    }

    return null;
  }

  function isInvoiceEditorOpen() {
    return !!getInvoiceRoot();
  }

  function getRowElements(root) {
    if (!root) return [];
    return Array.from(root.querySelectorAll(".dgrid-row"));
  }

  function getSKU(row) {
    const selectors = [
      ".field-sku",
      '[data-automation-id="sku"]',
      ".sku-field",
      ".itemSKU",
      '[aria-label*="SKU"]',
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
      ".field-qty",
      'input[aria-label*="Quantity"]',
      'input[name*="quantity"]',
    ];

    for (const selector of selectors) {
      const el = row.querySelector(selector);
      if (!el) continue;

      const value = getElementValue(el);
      const parsed = tryParseNumber(value);
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

  function getProductName(row) {
    const selectors = [
      ".itemColumn",
      '[data-automation-id="item-name"]',
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
        return field ? getElementValue(field.querySelector("input")) : "";
      };

      data.orderNumber = findFieldValue("ORDER NUMBER");
      data.jobName = findFieldValue("JOB NAME");
      data.phoneNumber = findFieldValue("Phone");
    }

    return data;
  }

  function isHeaderOnlyRow(productName, description, sku, quantity) {
    const text = (productName || description || "").trim();

    // No SKU + zero qty + looks like a section heading
    if (sku) return false;
    if (Number(quantity || 0) !== 0) return false;
    if (!text) return true;

    // Matches rows like ** Kitchen **, ** WC **, ** Ensuite **
    if (/^\*+\s*.+?\s*\*+$/.test(text)) return true;

    // Optional: also skip very short all-text labels with no SKU and no qty
    // Uncomment only if needed:
    // if (!sku && Number(quantity || 0) === 0 && text.length <= 40) return true;

    return false;
  }

  function extractRows(root) {
    const rowElements = getRowElements(root);
    const rows = [];

    rowElements.forEach((row) => {
      const productName = getProductName(row);
      const description = getDescription(row);
      const sku = getSKU(row);
      const quantity = getQuantity(row);

      const displayText = productName || description || "";

      // Skip truly empty rows
      if (!productName && !description && !sku) return;

      // Skip header / section rows
      if (isHeaderOnlyRow(productName, description, sku, quantity)) {
        return;
      }

      rows.push({
        key: sku || productName || description,
        productName,
        description,
        sku,
        quantity,
        displayText,
      });
    });

    return rows;
  }

  function extractData() {
    const root = getInvoiceRoot();

    if (!root) {
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

    return {
      ...extractHeaderData(root),
      rows: extractRows(root),
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
      bottom: 4px;
      left: ${left};
      padding: 2px 24px;
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
      if (!isInvoicePage()) {
        removeButtons();
        return;
      }

      if (!isInvoiceEditorOpen()) {
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
        const groupKey = row.sku ? `SKU:${row.sku}` : `NAME:${row.productName || row.description}`;
        if (!grouped.has(groupKey)) {
          grouped.set(groupKey, {
            productName: row.productName || row.description || "",
            sku: row.sku || "",
            quantity: Number(row.quantity || 0),
          });
        } else {
          grouped.get(groupKey).quantity += Number(row.quantity || 0);
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
      rows.forEach((row) => {
        const displayName = row.productName || (row.sku ? "" : row.description);

        if (displayName || row.sku) {
          html += `
            <tr style="height:30px;">
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
        <thead style="background:lightgrey;">
          <tr>
            <th>Product Name</th>
            <th>SKU</th>
            <th>Quantity</th>
          </tr>
        </thead>
        <tbody>
          ${html}
        </tbody>
      </table>
    `;
  }

  async function generateProductTable(combineQuantities) {
    const data = await getStableExtractedData();

    if (!data.rows.length) {
      alert("No product data found. Please ensure the invoice is fully loaded.");
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

  function generatePrintLayout(data, productTable) {
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
    .TMheader-left img { max-width: 150px; height: auto; }
    .TMheader-right { width: 30%; font-size: 10px; }
    .deliveryNote { display: flex; justify-content: space-between; margin-bottom: 10px; font-size: 13px; width: 100%; }
    .deliveryNote p { margin-bottom: 0; }
    .deliveryNote > div { width: 33.3%; white-space: pre-line; }
    .orderProducts { margin-top: 20px; border-top: 1px solid #000; }
    .product-table { text-align: left; width: 100%; margin-top: 20px; font-size: 14px; border-collapse: collapse; }
    .product-table th { padding: 8px; text-align: left; }
    .product-table td { padding: 6px; }
    .product-table tbody tr td:last-child { text-align: right; }
    .orderNote { display: flex; justify-content: space-between; font-size: 14px; margin: 12px 0; }
    .input-group { display: flex; justify-content: space-between; margin-top: 20px; font-size: 12px; }
    hr { border: none; border-top: 1px solid #ccc; margin: 10px 0; }
    @media print { body { margin: 0; } .maincontainer { padding: 15px; } }
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

    <h3 style="margin-bottom:10px">Delivery Note</h3>

    <div class="deliveryNote">
      <div>
        <b>INVOICE TO</b>
        <p>${escapeHtml(data.billingAddress)}</p>
      </div>
      <div>
        <b>SHIP TO</b>
        <p>${escapeHtml(data.shippingAddress)}</p>
      </div>
      <div>
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