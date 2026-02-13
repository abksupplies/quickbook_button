// ==UserScript==
// @name         Add Print and Pick Slip Buttons to QuickBooks Invoice
// @namespace    http://tampermonkey.net/
// @version      2.8
// @description  Adds "Print" and "Pick Slip" buttons to QuickBooks Invoice (works for existing + newly created invoices)
// @author       Raj - Gorkhari (Improved)
// @match        https://qbo.intuit.com/*
// @include      https://qbo.intuit.com/app/invoice?*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
  "use strict";

  let currentInvoiceId = null;

  const BUTTON_CHECK_INTERVAL = 2000; // Check every 2 seconds
  const DATA_LOAD_TIMEOUT = 10000; // Wait up to 10 seconds for data
  const DATA_CHECK_INTERVAL = 500; // Check every 500ms

  // ---------------------------
  // Page / Context Detection
  // ---------------------------

  // Function to check if the current page is an invoice page
  // (QuickBooks can have multiple invoice URL shapes depending on navigation)
  function isInvoicePage() {
    const url = window.location.href;
    return (
      url.includes("qbo.intuit.com/app/invoice") ||
      url.includes("/app/invoice?") ||
      url.includes("/invoice?txnId") ||
      url.includes("/invoice") // keep broad because QBO changes routes
    );
  }

  // Function to extract invoice ID from URL (existing invoices)
  function getInvoiceId() {
    const match = window.location.href.match(/txnId=([^&]+)/);
    return match ? match[1] : null;
  }

  /**
   * Returns the root container that holds the invoice editor.
   * - Existing invoices opened from list often appear in the right-side overlay (trowser)
   * - New invoices created inside QBO often open as a full-page editor (no trowser)
   */
  function getInvoiceRoot() {
    // 1) Overlay (trowser) invoice
    const overlayRoot = document.querySelector(".trowser-view .body");
    if (overlayRoot && overlayRoot.children.length > 0) return overlayRoot;

    // 2) Full page invoice editor: try a few likely containers
    const candidates = [
      document.querySelector('[data-automation-id="invoice-form"]'),
      document.querySelector('[data-automation-id="invoice-editor"]'),
      document.querySelector(".invoice-content"),
      document.querySelector("#qbo-main"),
      document.querySelector("#app"),
      document.body,
    ].filter(Boolean);

    for (const c of candidates) {
      // Looks like an invoice editor if it contains any of these
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
    const root = getInvoiceRoot();
    if (!root) return false;

    // Visible-ish check
    const style = window.getComputedStyle(root);
    if (style.display === "none" || style.visibility === "hidden") return false;

    // If overlay exists but is hidden, treat as closed
    const trowserView = document.querySelector(".trowser-view");
    if (trowserView) {
      const tvStyle = window.getComputedStyle(trowserView);
      // If trowser exists and is visible, root is overlayRoot - OK
      // If trowser exists but hidden, full page might still exist - don't block.
      // So we don't early-return false here.
      void tvStyle;
    }

    return true;
  }

  // ---------------------------
  // Data Loading Helpers
  // ---------------------------

  function waitForData(root, selector, timeout = DATA_LOAD_TIMEOUT) {
    return new Promise((resolve) => {
      const startTime = Date.now();

      const checkInterval = setInterval(() => {
        const elapsed = Date.now() - startTime;

        if (!root || !document.contains(root)) {
          if (elapsed >= timeout) {
            clearInterval(checkInterval);
            resolve(false);
          }
          return;
        }

        const element = root.querySelector(selector);
        if (element) {
          const text = (element.value ?? element.textContent ?? "").toString().trim();
          if (text) {
            clearInterval(checkInterval);
            resolve(true);
            return;
          }
        }

        if (elapsed >= timeout) {
          clearInterval(checkInterval);
          resolve(false);
        }
      }, DATA_CHECK_INTERVAL);
    });
  }

  async function isInvoiceDataLoaded() {
    const root = getInvoiceRoot();
    if (!root) return false;

    // Check for multiple indicators that data is loaded within the editor
    const checks = [
      waitForData(root, ".dgrid-row", 3000), // Line rows
      waitForData(root, '[data-qbo-bind="text: referenceNumber"]', 3000), // Invoice number
      waitForData(root, "textarea.topFieldInput.address", 3000), // Billing address
    ];

    const results = await Promise.all(checks);
    return results.some(Boolean);
  }

  // ---------------------------
  // Buttons
  // ---------------------------

  function createButton(id, text, clickHandler) {
    const button = document.createElement("button");
    button.id = id;
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
            box-shadow: 0 2px 8px rgba(0,0,0,0.01);
            cursor: pointer;
            font-size: 14px;
            font-weight: 600;
            z-index: 10000;
            transition: all 0.3s ease;
        `;

    button.addEventListener("mouseenter", () => {
      button.style.backgroundColor = "#248f17";
      button.style.transform = "translateY(-2px)";
      button.style.boxShadow = "0 4px 12px rgba(0,0,0,0.1)";
    });

    button.addEventListener("mouseleave", () => {
      button.style.backgroundColor = "#2ca01c";
      button.style.transform = "translateY(0)";
      button.style.boxShadow = "0 2px 8px rgba(0,0,0,0.06)";
    });

    button.addEventListener("click", clickHandler);
    return button;
  }

  async function addButtons() {
    // Must be on invoice page
    if (!isInvoicePage()) {
      removeButtons();
      return;
    }

    // Invoice editor must be present (overlay OR full page)
    if (!isInvoiceEditorOpen()) {
      removeButtons();
      return;
    }

    // Track invoice id changes (existing invoices). New invoices may not have txnId yet.
    const invoiceId = getInvoiceId();
    if (invoiceId && invoiceId !== currentInvoiceId) {
      currentInvoiceId = invoiceId;
      removeButtons(); // Remove old buttons for new invoice
    }

    // Don't add if buttons already exist
    if (document.getElementById("custom-print-button")) return;

    // Wait for data to load before adding buttons
    const dataLoaded = await isInvoiceDataLoaded();
    if (!dataLoaded) {
      console.warn("QuickBooks invoice data not fully loaded yet");
      return; // Will retry on next interval
    }

    const printButton = createButton("custom-print-button", "🖨️ Print", () =>
      generateProductTable(false)
    );
    document.body.appendChild(printButton);

    const pickSlipButton = createButton("custom-pick-slip-button", "📋 Pick Slip", () =>
      generateProductTable(true)
    );
    document.body.appendChild(pickSlipButton);

    console.log("Buttons added successfully");
  }

  function removeButtons() {
    const printButton = document.getElementById("custom-print-button");
    const pickSlipButton = document.getElementById("custom-pick-slip-button");
    if (printButton) {
      printButton.remove();
      console.log("Buttons removed");
    }
    if (pickSlipButton) pickSlipButton.remove();
  }

  // ---------------------------
  // Data Extraction (scoped to invoice root)
  // ---------------------------

  function extractData() {
    const root = getInvoiceRoot();
    if (!root) {
      console.error("Invoice editor root not found");
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

    // Try multiple selectors for SKU as fallback
    function getSKU(row) {
      const selectors = [".field-sku", '[data-automation-id="sku"]', ".sku-field", ".itemSKU"];
      for (const selector of selectors) {
        const element = row.querySelector(selector);
        if (element && element.textContent.trim()) return element.textContent.trim();
      }
      return "";
    }

    // Try multiple selectors for quantity
    function getQuantity(row) {
      const selectors = [".field-quantity-inner", '[data-automation-id="quantity"]', ".quantity-field"];
      for (const selector of selectors) {
        const element = row.querySelector(selector);
        if (element) {
          const text = element.textContent.trim();
          const qty = parseInt(text, 10);
          if (!isNaN(qty)) return qty;
        }
      }
      return 0;
    }

    const data = {
      billingAddress: root.querySelector("textarea.topFieldInput.address")?.value || "N/A",
      shippingAddress: root.querySelector("#shippingAddress")?.value || "N/A",
      invoiceNumber:
        root.querySelector('[data-qbo-bind="text: referenceNumber"]')?.textContent?.trim() ||
        "N/A",
      invoiceDate: root.querySelector(".dijitDateTextBox input.dijitInputInner")?.value || "N/A",
      rows: [],
    };

    // Extract custom form fields - scoped to invoice root
    const formElement = root.querySelector(".custom-form");
    if (formElement) {
      const formFields = Array.from(formElement.querySelectorAll(".custom-form-field"));
      data.orderNumber =
        formFields
          .find((f) => f.querySelector("label")?.textContent.trim() === "ORDER NUMBER")
          ?.querySelector("input")?.value || "";
      data.jobName =
        formFields.find((f) => f.querySelector("label")?.textContent.trim() === "JOB NAME")
          ?.querySelector("input")?.value || "";
      data.phoneNumber =
        formFields.find((f) => f.querySelector("label")?.textContent.trim() === "Phone")
          ?.querySelector("input")?.value || "";
    } else {
      data.orderNumber = "";
      data.jobName = "";
      data.phoneNumber = "";
    }

    // Extract product rows - scoped to invoice root
    const rows = root.querySelectorAll(".dgrid-row");
    console.log(`Found ${rows.length} product rows`);

    rows.forEach((row) => {
      const productNameElement = row.querySelector(".itemColumn");
      const descriptionElement = row.querySelector(".field-description div");

      const productName = productNameElement?.textContent.trim() || "";
      const description = descriptionElement?.textContent.trim() || "";
      const sku = getSKU(row);
      const quantity = getQuantity(row);

      if (productName || sku || description) {
        data.rows.push({ productName, description, sku, quantity });
      }
    });

    return data;
  }

  // ---------------------------
  // Print / Pick Slip Generation
  // ---------------------------

  function generateProductTable(combineQuantities) {
    const data = extractData();

    if (data.rows.length === 0) {
      alert("No product data found. Please ensure the invoice is fully loaded.");
      return;
    }

    const skuMap = new Map();
    let productTable = "";

    data.rows.forEach((row) => {
      if (combineQuantities) {
        // Combine quantities for Pick Slip by SKU
        if (row.sku && skuMap.has(row.sku)) {
          const existing = skuMap.get(row.sku);
          existing.quantity += row.quantity;
        } else if (row.sku) {
          skuMap.set(row.sku, {
            productName: row.productName,
            quantity: row.quantity,
          });
        }
      } else {
        // Print logic: Include description only if both product name and SKU are missing
        const displayName = row.productName || (row.sku ? "" : row.description);

        if (displayName || row.sku) {
          productTable += `
                        <tr style="margin-bottom: 5px; height: 30px;">
                            <td>${escapeHtml(displayName)}</td>
                            <td>${escapeHtml(row.sku)}</td>
                            <td style="text-align:right;">${row.quantity || ""}</td>
                        </tr>
                    `;
        }
      }
    });

    if (combineQuantities) {
      skuMap.forEach((value, sku) => {
        productTable += `
                    <tr style="margin-bottom: 5px; height: 30px;">
                        <td>${escapeHtml(value.productName)}</td>
                        <td>${escapeHtml(sku)}</td>
                        <td style="text-align:right;">${value.quantity}</td>
                    </tr>
                `;
      });
    }

    if (productTable) {
      productTable = `
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
            `;
    } else {
      productTable = "<p>No valid products found.</p>";
    }

    const printLayout = generatePrintLayout(data, productTable);

    const newWindow = window.open("", "_blank", "width=800,height=600");
    if (newWindow) {
      newWindow.document.write(printLayout);
      newWindow.document.close();
      newWindow.print();
    }
  }

  // Prevent HTML injection from invoice fields
  function escapeHtml(str) {
    return (str ?? "")
      .toString()
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
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
            <div class="orderNote-1"><b>ORDER NUMBER</b><br/>${escapeHtml(data.orderNumber)}</div>
            <div class="orderNote-2"><b>JOB NAME</b><br/>${escapeHtml(data.jobName)}</div>
            <div class="orderNote-3"><b>PHONE</b><br/>${escapeHtml(data.phoneNumber)}</div>
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

  // ---------------------------
  // Observers / SPA navigation
  // ---------------------------

  function setupObservers() {
    // Override pushState to detect SPA navigation
    const originalPushState = history.pushState;
    history.pushState = function () {
      const result = originalPushState.apply(this, arguments);
      setTimeout(() => {
        if (isInvoicePage() && isInvoiceEditorOpen()) addButtons();
        else removeButtons();
      }, 600);
      return result;
    };

    // popstate for back/forward
    window.addEventListener("popstate", () => {
      setTimeout(() => {
        if (isInvoicePage() && isInvoiceEditorOpen()) addButtons();
        else removeButtons();
      }, 600);
    });

    // Mutation observer for dynamic content changes
    const observer = new MutationObserver(() => {
      if (isInvoicePage() && isInvoiceEditorOpen()) {
        if (!document.getElementById("custom-print-button")) addButtons();
      } else {
        removeButtons();
      }
    });

    observer.observe(document.body, { childList: true, subtree: true });
  }

  // Periodic check
  setInterval(() => {
    if (isInvoicePage() && isInvoiceEditorOpen()) addButtons();
    else removeButtons();
  }, BUTTON_CHECK_INTERVAL);

  // Initialize
  setupObservers();
  setTimeout(() => {
    if (isInvoicePage() && isInvoiceEditorOpen()) addButtons();
  }, 2000);
})();
