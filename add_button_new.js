// ==UserScript==
// @name         Add Print and Pick Slip Buttons to QuickBooks Invoice (Stable Qty/SKU)
// @namespace    http://tampermonkey.net/
// @version      3.0
// @description  Adds "Print" and "Pick Slip" buttons to QuickBooks Invoice overlay and prints a stable pick slip (fixes random qty concatenation / row virtualization issues)
// @author       Raj - Gorkhari (Hardened)
// @match        https://qbo.intuit.com/*
// @include      https://qbo.intuit.com/app/invoice?*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
  "use strict";

  let currentInvoiceId = null;

  const BUTTON_CHECK_INTERVAL = 2000; // Check every 2 seconds
  const DATA_LOAD_TIMEOUT = 12000; // Wait up to 12 seconds for data
  const DATA_CHECK_INTERVAL = 250; // Check every 250ms

  const GRID_STABLE_MAX_WAIT = 5000; // up to 5s to wait for stable rows
  const GRID_STABLE_HOLD_MS = 650; // must stay same for this long to be "stable"
  const QTY_SANITY_MAX = 999; // skip insane qty values (prevents 12121 disasters)

  // ---------------------------
  // Helpers
  // ---------------------------

  function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms));
  }

  function isInvoiceOverlayOpen() {
    const trowserView = document.querySelector(".trowser-view");
    if (!trowserView) return false;

    const trowserBody = document.querySelector(".trowser-view .body");
    if (!trowserBody) return false;

    const hasContent = trowserBody.children.length > 0;
    const isVisible =
      window.getComputedStyle(trowserView).display !== "none" &&
      window.getComputedStyle(trowserBody).display !== "none";

    return hasContent && isVisible;
  }

  function isInvoicePage() {
    const url = window.location.href;
    return url.includes("qbo.intuit.com/app/invoice?") || url.includes("/invoice?txnId");
  }

  function getInvoiceId() {
    const match = window.location.href.match(/txnId=([^&]+)/);
    return match ? match[1] : null;
  }

  function waitForData(selector, timeout = DATA_LOAD_TIMEOUT) {
    return new Promise((resolve) => {
      const startTime = Date.now();

      const checkInterval = setInterval(() => {
        const invoiceContainer = document.querySelector(".trowser-view .body");
        if (!invoiceContainer) {
          if (Date.now() - startTime >= timeout) {
            clearInterval(checkInterval);
            resolve(false);
          }
          return;
        }

        const element = invoiceContainer.querySelector(selector);
        if (element) {
          const val = (element.value ?? element.textContent ?? "").toString().trim();
          if (val) {
            clearInterval(checkInterval);
            resolve(true);
            return;
          }
        }

        if (Date.now() - startTime >= timeout) {
          clearInterval(checkInterval);
          resolve(false);
        }
      }, DATA_CHECK_INTERVAL);
    });
  }

  async function isInvoiceDataLoaded() {
    const invoiceContainer = document.querySelector(".trowser-view .body");
    if (!invoiceContainer) return false;

    const checks = [
      waitForData('.dgrid-row[role="row"]', 3500), // rows exist
      waitForData('[data-qbo-bind="text: referenceNumber"]', 3500), // invoice no
      waitForData("textarea.topFieldInput.address", 3500), // billing address
    ];

    const results = await Promise.all(checks);
    return results.some(Boolean);
  }

  async function waitForStableGrid(invoiceContainer) {
    // Wait until row count stays the same for GRID_STABLE_HOLD_MS
    const start = Date.now();
    let lastCount = -1;

    while (Date.now() - start < GRID_STABLE_MAX_WAIT) {
      const rows = invoiceContainer.querySelectorAll('.dgrid-row[role="row"]');
      const count = rows.length;

      if (count === lastCount && count > 0) {
        await sleep(GRID_STABLE_HOLD_MS);
        const rows2 = invoiceContainer.querySelectorAll('.dgrid-row[role="row"]');
        if (rows2.length === count) return true;
      }

      lastCount = count;
      await sleep(150);
    }
    return false;
  }

  // Extract first numeric token only
  function firstNumberToken(str) {
    if (!str) return null;
    const m = String(str).replace(/\s+/g, " ").trim().match(/-?\d+(\.\d+)?/);
    return m ? Number(m[0]) : null;
  }

  function safeText(el) {
    if (!el) return "";
    return ((el.innerText || el.textContent || el.value || "") + "").trim();
  }

  // ---------------------------
  // Button UI
  // ---------------------------

  function createButton(id, text, clickHandler) {
    const button = document.createElement("button");
    button.id = id;
    button.type = "button";
    button.textContent = text;

    button.style.cssText = `
      position: fixed;
      bottom: 4px;
      left: ${id === "custom-print-button" ? "15%" : "calc(15% + 140px)"};
      padding: 12px 24px;
      background-color: #2ca01c;
      color: white;
      border: none;
      border-radius: 6px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.06);
      cursor: pointer;
      font-size: 14px;
      font-weight: 600;
      z-index: 10000;
      transition: all 0.2s ease;
    `;

    button.addEventListener("mouseenter", () => {
      button.style.backgroundColor = "#248f17";
      button.style.transform = "translateY(-2px)";
      button.style.boxShadow = "0 4px 12px rgba(0,0,0,0.10)";
    });

    button.addEventListener("mouseleave", () => {
      button.style.backgroundColor = "#2ca01c";
      button.style.transform = "translateY(0)";
      button.style.boxShadow = "0 2px 8px rgba(0,0,0,0.06)";
    });

    button.addEventListener("click", clickHandler);
    return button;
  }

  function removeButtons() {
    const printButton = document.getElementById("custom-print-button");
    const pickSlipButton = document.getElementById("custom-pick-slip-button");
    if (printButton) printButton.remove();
    if (pickSlipButton) pickSlipButton.remove();
  }

  async function addButtons() {
    if (!isInvoicePage()) {
      removeButtons();
      return;
    }

    if (!isInvoiceOverlayOpen()) {
      removeButtons();
      return;
    }

    const invoiceId = getInvoiceId();
    if (invoiceId !== currentInvoiceId) {
      currentInvoiceId = invoiceId;
      removeButtons();
    }

    if (document.getElementById("custom-print-button")) return;

    const dataLoaded = await isInvoiceDataLoaded();
    if (!dataLoaded) return;

    const printButton = createButton("custom-print-button", "🖨️ Print", async () => {
      await generateProductTable(false);
    });
    document.body.appendChild(printButton);

    const pickSlipButton = createButton("custom-pick-slip-button", "📋 Pick Slip", async () => {
      await generateProductTable(true);
    });
    document.body.appendChild(pickSlipButton);
  }

  // ---------------------------
  // Data extraction (Hardened)
  // ---------------------------

  function getSKU(row) {
    // Try to target the sku cell, then take first token only
    const cell =
      row.querySelector(".dgrid-cell.field-sku") ||
      row.querySelector(".field-sku") ||
      row.querySelector('[data-automation-id="sku"]') ||
      row.querySelector(".sku-field") ||
      row.querySelector(".itemSKU");

    const t = safeText(cell);
    return t ? t.split(/\s+/)[0] : "";
  }

  function getQuantity(row) {
    // 1) Prefer input value (most reliable if editor exists)
    const input =
      row.querySelector('input[data-automation-id="quantity"]') ||
      row.querySelector('input[aria-label*="Quantity"]') ||
      row.querySelector('input[name*="quantity"]');

    if (input && input.value != null) {
      const n = firstNumberToken(input.value);
      return n != null ? n : 0;
    }

    // 2) Otherwise read ONLY the quantity cell, then parse ONLY first number token
    const cell =
      row.querySelector(".dgrid-cell.field-quantity") ||
      row.querySelector(".field-quantity") ||
      row.querySelector('[data-automation-id="quantity"]') ||
      row.querySelector(".field-quantity-inner") ||
      row.querySelector(".quantity-field");

    const txt = safeText(cell);
    const n = firstNumberToken(txt);
    return n != null ? n : 0;
  }

  function getProductName(row) {
    // QBO often uses .itemColumn for item name
    const el = row.querySelector(".itemColumn") || row.querySelector('[data-automation-id="item"]');
    return safeText(el);
  }

  function getDescription(row) {
    // Description may be inside .field-description div
    const el =
      row.querySelector(".field-description div") ||
      row.querySelector(".field-description") ||
      row.querySelector('[data-automation-id="description"]');
    return safeText(el);
  }

  function extractCustomFields(invoiceContainer) {
    const out = { orderNumber: "", jobName: "", phoneNumber: "" };

    const formElement = invoiceContainer.querySelector(".custom-form");
    if (!formElement) return out;

    const formFields = Array.from(formElement.querySelectorAll(".custom-form-field"));
    const getFieldValue = (labelText) => {
      const f = formFields.find(
        (x) => (x.querySelector("label")?.textContent || "").trim() === labelText
      );
      return f?.querySelector("input")?.value || "";
    };

    out.orderNumber = getFieldValue("ORDER NUMBER");
    out.jobName = getFieldValue("JOB NAME");
    out.phoneNumber = getFieldValue("Phone");

    return out;
  }

  function getFilteredRows(invoiceContainer) {
    // Only "real" visible item rows
    return Array.from(invoiceContainer.querySelectorAll('.dgrid-row[role="row"]'))
      .filter((r) => !r.closest('[aria-hidden="true"]'))
      .filter((r) => r.offsetParent !== null) // visible
      .filter((r) => r.querySelector(".itemColumn") || r.querySelector('[data-automation-id="item"]'));
  }

  async function extractData() {
    const invoiceContainer = document.querySelector(".trowser-view .body");
    if (!invoiceContainer) {
      console.error("Invoice overlay container not found");
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

    // Wait for grid to stabilize (virtualized grid / repaint protection)
    await waitForStableGrid(invoiceContainer);

    const billingAddress =
      invoiceContainer.querySelector("textarea.topFieldInput.address")?.value || "N/A";

    const shippingAddress =
      invoiceContainer.querySelector("#shippingAddress")?.value ||
      invoiceContainer.querySelector('textarea[name*="shippingAddress"]')?.value ||
      "N/A";

    const invoiceNumber =
      safeText(invoiceContainer.querySelector('[data-qbo-bind="text: referenceNumber"]')) || "N/A";

    const invoiceDate =
      invoiceContainer.querySelector(".dijitDateTextBox input.dijitInputInner")?.value ||
      invoiceContainer.querySelector('input[aria-label*="Invoice date"]')?.value ||
      "N/A";

    const custom = extractCustomFields(invoiceContainer);

    const data = {
      billingAddress,
      shippingAddress,
      invoiceNumber,
      invoiceDate,
      orderNumber: custom.orderNumber,
      jobName: custom.jobName,
      phoneNumber: custom.phoneNumber,
      rows: [],
    };

    const rows = getFilteredRows(invoiceContainer);

    rows.forEach((row) => {
      const productName = getProductName(row);
      const description = getDescription(row);
      const sku = getSKU(row);
      const quantity = getQuantity(row);

      // Sanity & validity checks
      if (!productName && !sku && !description) return;
      if (!Number.isFinite(quantity) || quantity <= 0) return;
      if (quantity > QTY_SANITY_MAX) {
        console.warn("Suspicious quantity skipped:", { sku, productName, quantity });
        return;
      }

      data.rows.push({ productName, description, sku, quantity });
    });

    return data;
  }

  // ---------------------------
  // Printing
  // ---------------------------

  async function generateProductTable(combineQuantities) {
    const data = await extractData();

    if (!data.rows || data.rows.length === 0) {
      alert("No product data found. Please ensure the invoice is fully loaded.");
      return;
    }

    const skuMap = new Map();
    let productTable = "";

    data.rows.forEach((row) => {
      if (combineQuantities) {
        // Combine quantities for Pick Slip by SKU
        if (row.sku && skuMap.has(row.sku)) {
          skuMap.get(row.sku).quantity += row.quantity;
        } else if (row.sku) {
          skuMap.set(row.sku, { productName: row.productName, quantity: row.quantity });
        } else {
          // No SKU? fallback to productName key so it doesn't disappear
          const key = `NO-SKU:${row.productName || row.description || "UNKNOWN"}`;
          if (skuMap.has(key)) skuMap.get(key).quantity += row.quantity;
          else skuMap.set(key, { productName: row.productName || row.description || "UNKNOWN", quantity: row.quantity });
        }
      } else {
        // Print logic: Include description only if both product name and SKU are missing
        const displayName = row.productName || (row.sku ? "" : row.description);
        if (displayName || row.sku) {
          productTable += `
            <tr style="margin-bottom: 5px; height: 30px;">
              <td>${escapeHtml(displayName)}</td>
              <td>${escapeHtml(row.sku || "")}</td>
              <td style="text-align:right;">${row.quantity || ""}</td>
            </tr>
          `;
        }
      }
    });

    if (combineQuantities) {
      // Generate table for Pick Slip
      skuMap.forEach((value, sku) => {
        const shownSku = sku.startsWith("NO-SKU:") ? "" : sku;
        productTable += `
          <tr style="margin-bottom: 5px; height: 30px;">
            <td>${escapeHtml(value.productName || "")}</td>
            <td>${escapeHtml(shownSku)}</td>
            <td style="text-align:right;">${value.quantity}</td>
          </tr>
        `;
      });
    }

    const finalTable = productTable
      ? `
        <table class="product-table">
          <thead style="background:lightgrey; margin-bottom:5px;">
            <tr>
              <th>Product Name</th>
              <th>SKU</th>
              <th>Quantity</th>
            </tr>
          </thead>
          <tbody>
            ${productTable}
          </tbody>
        </table>
      `
      : "<p>No valid products found.</p>";

    const printLayout = generatePrintLayout(data, finalTable);

    const newWindow = window.open("", "_blank", "width=900,height=700");
    if (newWindow) {
      newWindow.document.write(printLayout);
      newWindow.document.close();
      newWindow.focus();
      newWindow.print();
    }
  }

  function escapeHtml(str) {
    return String(str ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function generatePrintLayout(data, productTable) {
    return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Invoice ${escapeHtml(data.invoiceNumber)}</title>
  <style>
    * { box-sizing: border-box; }
    body {
      font-family: Arial, sans-serif;
      line-height: 1.6;
      margin: 0;
      padding: 0;
    }
    .maincontainer {
      width: 100%;
      padding: 20px;
      margin: 0 auto;
      line-height: 1.4;
    }
    .maincontainer h3 {
      margin-top: 0;
      margin-bottom: 0;
    }
    .ContactService textarea {
      line-height: 1.65;
      margin-bottom: -5px;
    }
    .TMheader {
      width: 100%;
      display: flex;
      justify-content: space-between;
      margin-bottom: 0px;
    }
    .TMheader-left {
      width: 70%;
      display: flex;
      font-size: 13px;
      gap: 20px;
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
      margin-bottom: 10px;
      font-size: 13px;
      width: 100%;
    }
    .deliveryNote p { margin-bottom: 0; }
    .deliveryNote > div {
      width: 33.3%;
      white-space: pre-line;
    }
    .orderProducts {
      margin-top: 20px;
      border-top: 1px solid #000;
    }
    .product-table {
      text-align: left;
      width: 100%;
      margin-top: 20px;
      font-size: 14px;
      border-collapse: collapse;
    }
    .product-table th {
      padding: 8px;
      text-align: left;
    }
    .product-table td { padding: 6px; }
    .product-table tbody tr td:last-child { text-align: right; }
    .orderNote {
      display: flex;
      justify-content: space-between;
      font-size: 14px;
      margin: 12px 0;
    }
    .input-group {
      display: flex;
      justify-content: space-between;
      margin-top: 20px;
      font-size: 12px;
    }
    hr {
      border: none;
      border-top: 1px solid #ccc;
      margin: 10px 0;
    }
    @media print {
      body { margin: 0; }
      .maincontainer { padding: 15px; }
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
          <div><textarea class="form-control" rows="1" style="height:30px;width:200px;" placeholder="Name:"></textarea></div>
          <div><textarea class="form-control" rows="1" style="height:40px;width:200px;" placeholder="Sign:"></textarea></div>
          <div><textarea class="form-control" rows="1" style="height:30px;width:200px;" placeholder="Date: __ / __ / ____"></textarea></div>
        </div>
      </div>
    </div>

    <h3 style="margin-bottom:10px">Delivery Note</h3>

    <div class="deliveryNote">
      <div class="deliveryNote-1">
        <b>INVOICE TO</b>
        <p class="InvoiceInfo">${escapeHtml(data.billingAddress)}</p>
      </div>
      <div class="deliveryNote-2">
        <b>SHIP TO</b>
        <p class="ShipInfo">${escapeHtml(data.shippingAddress)}</p>
      </div>
      <div class="deliveryNote-3">
        <b>INVOICE NO.:</b>
        <span class="InvoiceNumber">${escapeHtml(data.invoiceNumber)}</span><br/>
        <b>DATE:</b>
        <span class="InvoiceDate">${escapeHtml(data.invoiceDate)}</span>
      </div>
    </div>

    <hr>

    <div class="orderNote">
      <div class="orderNote-1"><b>ORDER NUMBER</b><br/>${escapeHtml(data.orderNumber || "")}</div>
      <div class="orderNote-2"><b>JOB NAME</b><br/>${escapeHtml(data.jobName || "")}</div>
      <div class="orderNote-3"><b>PHONE</b><br/>${escapeHtml(data.phoneNumber || "")}</div>
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
</html>`;
  }

  // ---------------------------
  // Observers / page changes
  // ---------------------------

  function setupObservers() {
    const originalPushState = history.pushState;
    history.pushState = function () {
      const result = originalPushState.apply(this, arguments);
      setTimeout(() => {
        if (isInvoiceOverlayOpen()) addButtons();
        else removeButtons();
      }, 900);
      return result;
    };

    window.addEventListener("popstate", () => {
      setTimeout(() => {
        if (isInvoiceOverlayOpen()) addButtons();
        else removeButtons();
      }, 900);
    });

    const observer = new MutationObserver(() => {
      if (isInvoicePage() && isInvoiceOverlayOpen()) {
        if (!document.getElementById("custom-print-button")) addButtons();
      } else if (!isInvoiceOverlayOpen()) {
        removeButtons();
      }
    });

    observer.observe(document.body, { childList: true, subtree: true });

    // Specifically watch overlay changes
    const overlayObserver = new MutationObserver(() => {
      if (isInvoiceOverlayOpen()) addButtons();
      else removeButtons();
    });

    const checkForOverlay = setInterval(() => {
      const trowserView = document.querySelector(".trowser-view");
      if (trowserView) {
        overlayObserver.observe(trowserView, {
          childList: true,
          subtree: true,
          attributes: true,
          attributeFilter: ["style", "class"],
        });
        clearInterval(checkForOverlay);
      }
    }, 500);
  }

  // Periodic check
  setInterval(() => {
    if (isInvoicePage() && isInvoiceOverlayOpen()) addButtons();
    else removeButtons();
  }, BUTTON_CHECK_INTERVAL);

  // Initialize
  setupObservers();
  setTimeout(() => {
    if (isInvoiceOverlayOpen()) addButtons();
  }, 2000);
})();
