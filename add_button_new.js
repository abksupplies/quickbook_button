// ==UserScript==
// @name         QuickBooks Invoice Print + Pick Slip (Stable + Safer)
// @namespace    http://tampermonkey.net/
// @version      3.7
// @description  Adds Print and Pick Slip buttons to QuickBooks invoices for the new sales forms UI
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
    mutationDebounceMs: 600,
    initialBootDelayMs: 2200,
    stableReadAttempts: 8,
    stableReadDelayMs: 300,
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
    const title = document.title || "";
    return url.includes("/invoice") || url.includes("txnId=") || title.toLowerCase().includes("invoice");
  }

  function getInvoiceId() {
    const match = location.href.match(/[?&]txnId=([^&]+)/);
    return match ? decodeURIComponent(match[1]) : null;
  }

  function getInvoiceRoot() {
    const candidates = [
      document.querySelector(".trowser-view .body"),
      document.querySelector('[data-automation-id="invoice-form"]'),
      document.querySelector('[data-automation-id="invoice-editor"]'),
      document.querySelector('[data-testid*="invoice"]'),
      document.querySelector("#qbo-main"),
      document.querySelector("#app"),
      document.querySelector("main"),
      document.body,
    ].filter(Boolean);

    for (const c of candidates) {
      const text = c.textContent || "";
      if (
        text.includes("Product or service") &&
        text.includes("Qty") &&
        text.includes("Amount")
      ) {
        return c;
      }
    }

    return document.body;
  }

  function isInvoiceEditorOpen() {
    return isInvoicePage() && !!getInvoiceRoot();
  }

  function extractHeaderData(root) {
    const text = root?.textContent || "";

    const invoiceNumber =
      normalizeText(
        getElementValue(root.querySelector('[data-qbo-bind="text: referenceNumber"]')) ||
        getElementValue(root.querySelector('input[aria-label*="Invoice no"]')) ||
        getElementValue(root.querySelector('input[aria-label*="Invoice number"]'))
      ) ||
      (text.match(/\bInvoice\s+(\d{2,})\b/i)?.[1] ?? "N/A");

    const invoiceDate =
      getElementValue(root.querySelector(".dijitDateTextBox input.dijitInputInner")) ||
      getElementValue(root.querySelector('input[aria-label*="Invoice date"]')) ||
      "N/A";

    function findLabeledInput(labelText) {
      const labels = Array.from(root.querySelectorAll("label, div, span"));
      const match = labels.find((el) => normalizeText(el.textContent).toLowerCase() === labelText.toLowerCase());
      if (!match) return "";
      const container =
        match.closest("div") ||
        match.parentElement;
      if (!container) return "";
      return getElementValue(container.querySelector("input, textarea, select"));
    }

    return {
      billingAddress: getElementValue(root.querySelector("textarea.topFieldInput.address")) || "N/A",
      shippingAddress: getElementValue(root.querySelector("#shippingAddress")) || "N/A",
      invoiceNumber,
      invoiceDate,
      orderNumber: findLabeledInput("ORDER NUMBER"),
      jobName: findLabeledInput("JOB NAME"),
      phoneNumber: findLabeledInput("Phone"),
    };
  }

  function isHeaderOnlyRow(productName, description, sku, quantity) {
    const text = (productName || description || "").trim();
    if (sku) return false;
    if (Number(quantity || 0) !== 0) return false;
    if (!text) return true;
    if (/^\*+\s*.+?\s*\*+$/.test(text)) return true;
    return false;
  }

  function findLineItemsRegion(root) {
    const all = Array.from(root.querySelectorAll("div, section, table, [role='table']"));
    for (const el of all) {
      const text = normalizeText(el.textContent);
      if (
        text.includes("Product or service") &&
        text.includes("Description") &&
        text.includes("Qty") &&
        text.includes("Amount") &&
        text.includes("GST")
      ) {
        return el;
      }
    }
    return null;
  }

  function extractRowsFromVisibleGrid(root) {
    const region = findLineItemsRegion(root);
    if (!region) {
      warn("Line items region not found.");
      return [];
    }

    const rows = [];
    const candidates = Array.from(region.querySelectorAll("tr, [role='row'], .dgrid-row, div"));

    for (const row of candidates) {
      const text = normalizeText(row.textContent);
      if (!text) continue;

      // Skip header row
      if (
        text.includes("Product or service") &&
        text.includes("Description") &&
        text.includes("Qty")
      ) {
        continue;
      }

      // Skip toolbar/footer rows
      if (
        text.includes("Add product or service") ||
        text.includes("Clear all lines")
      ) {
        continue;
      }

      // Look for cells in this row
      let cells = Array.from(row.querySelectorAll("td, [role='cell']"));

      // Fallback for div-based grids
      if (!cells.length) {
        cells = Array.from(row.children).filter((el) => normalizeText(el.textContent));
      }

      if (cells.length < 3) continue;

      const cellTexts = cells.map((c) => normalizeText(c.textContent));

      // Expected new UI row shape from screenshot:
      // [drag, #, Product/service, Description, Qty, Rate, Amount, GST, delete]
      // Sometimes fewer wrapper cells may appear, so search flexibly.
      let productName = "";
      let description = "";
      let quantity = 0;

      // Prefer index-based mapping when row is wide enough
      if (cellTexts.length >= 7) {
        productName = cellTexts[2] || "";
        description = cellTexts[3] || "";
        quantity = tryParseNumber(cellTexts[4]) ?? 0;
      } else {
        // Fallback heuristic:
        // first long non-numeric cell = product/service
        // next long non-numeric cell = description
        const meaningful = cellTexts.filter(Boolean);
        const textCells = meaningful.filter((t) => !/^[-A$0-9.,()%]+$/.test(t));

        productName = textCells[0] || "";
        description = textCells[1] || "";
        const qtyCandidate = meaningful.find((t) => /^\d+(\.\d+)?$/.test(t));
        quantity = tryParseNumber(qtyCandidate) ?? 0;
      }

      // Description-only extraction bug guard:
      if (!productName && description) continue;

      // Skip empty add-new-line row
      if (!productName && !description && quantity === 0) continue;

      rows.push({
        key: productName || description,
        productName,
        description,
        sku: "",
        quantity,
      });
    }

    // De-duplicate obvious accidental captures
    const deduped = [];
    const seen = new Set();

    for (const row of rows) {
      const sig = JSON.stringify([row.productName, row.description, row.quantity]);
      if (seen.has(sig)) continue;
      seen.add(sig);
      deduped.push(row);
    }

    log("Visible grid rows extracted:", deduped);
    return deduped;
  }

  function extractRows() {
    const root = getInvoiceRoot();
    return extractRowsFromVisibleGrid(root);
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
        productName: r.productName,
        description: r.description,
        sku: r.sku,
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
      const grouped = new Map();

      rows.forEach((row) => {
        const qty = Number(row.quantity || 0);

        if (isHeaderOnlyRow(row.productName, row.description, row.sku, row.quantity)) return;
        if (!row.productName && qty === 0) return;

        const groupKey = row.sku
          ? `SKU:${row.sku}`
          : `NAME:${row.productName || row.description}`;

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

  async function generateProductTable(combineQuantities) {
    const data = await getStableExtractedData();

    if (!data.rows.length) {
      alert("No line items found in the new QuickBooks invoice grid.");
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