// ==UserScript==
// @name         QuickBooks Invoice Print + Pick Slip
// @namespace    http://tampermonkey.net/
// @version      6.3
// @description  Generates delivery notes and pick slips from the current QuickBooks invoice UI
// @author       Raj - Gorkhari
// @match        https://qbo.intuit.com/*
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
    debug: false,
  };

  const STATE = {
    addButtonsInFlight: false,
    mutationTimer: null,
    currentInvoiceId: null,
    printing: false,
  };

  // =========================================================
  // Logging and general helpers
  // =========================================================

  function log(...args) {
    if (CONFIG.debug) {
      console.log("[QBO Delivery Note]", ...args);
    }
  }

  function warn(...args) {
    console.warn("[QBO Delivery Note]", ...args);
  }

  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  function normalizeText(value) {
    return String(value ?? "")
      .replace(/\u00a0/g, " ")
      .replace(/[ \t]+/g, " ")
      .trim();
  }

  function normalizeMultiline(value) {
    return String(value ?? "")
      .replace(/\r\n?/g, "\n")
      .split("\n")
      .map((line) => normalizeText(line))
      .filter(Boolean)
      .join("\n");
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
      .replace(/[A$]/gi, "")
      .trim();

    if (!text) return null;

    const value = Number(text);
    return Number.isFinite(value) ? value : null;
  }

  function getInputValue(element) {
    if (!element) return "";

    if (typeof element.value === "string") {
      const value =
        element.tagName === "TEXTAREA"
          ? normalizeMultiline(element.value)
          : normalizeText(element.value);

      if (value) return value;
    }

    const attributeValue = element.getAttribute?.("value");

    if (
      typeof attributeValue === "string" &&
      normalizeText(attributeValue)
    ) {
      return normalizeText(attributeValue);
    }

    return "";
  }

  function getCellValue(cell) {
    if (!cell) return "";

    const preferredElements = Array.from(
      cell.querySelectorAll(
        [
          'input[aria-label^="Product or service line"]',
          'input[data-testid*="product line"]',
          'input[aria-label^="Quantity line"]',
          'input[data-testid^="quantity line"]',
          'textarea[data-testid="Description_field"]',
          "input",
          "textarea",
          "select",
          '[role="combobox"]',
        ].join(",")
      )
    );

    for (const element of preferredElements) {
      const value = getInputValue(element);

      if (value) {
        return value;
      }
    }

    return normalizeText(cell.textContent || "");
  }

  function queryFirst(searchRoot, selectors) {
    if (!searchRoot) return null;

    for (const selector of selectors) {
      const element = searchRoot.querySelector(selector);

      if (element) {
        return element;
      }
    }

    return null;
  }

  function getValueFromRootOrDocument(root, selectors) {
    const rootElement = queryFirst(root, selectors);
    const rootValue = getInputValue(rootElement);

    if (rootValue) {
      return rootValue;
    }

    if (root !== document) {
      const documentElement = queryFirst(document, selectors);
      return getInputValue(documentElement);
    }

    return "";
  }

  function joinCustomerAndAddress(customerName, billingAddress) {
    const customer = normalizeText(customerName);
    const address = normalizeMultiline(billingAddress);

    if (!customer) {
      return address || "N/A";
    }

    if (!address) {
      return customer;
    }

    const addressLines = address
      .split("\n")
      .map((line) => normalizeText(line))
      .filter(Boolean);

    const normalizedCustomer = customer.toLowerCase();

    const customerAlreadyPresent = addressLines.some(
      (line, index) => {
        if (index > 1) return false;

        const normalizedLine = line.toLowerCase();

        return (
          normalizedLine === normalizedCustomer ||
          normalizedLine.startsWith(
            `${normalizedCustomer} `
          ) ||
          normalizedCustomer.startsWith(
            `${normalizedLine} `
          )
        );
      }
    );

    if (customerAlreadyPresent) {
      return address;
    }

    return `${customer}\n${address}`;
  }

  // =========================================================
  // QuickBooks page detection
  // =========================================================

  function isInvoicePage() {
    const url = location.href;
    const title = document.title || "";

    return (
      url.includes("/invoice") ||
      url.includes("txnId=") ||
      title.toLowerCase().includes("invoice")
    );
  }

  function getInvoiceId() {
    const match = location.href.match(
      /[?&]txnId=([^&]+)/
    );

    return match
      ? decodeURIComponent(match[1])
      : null;
  }

  function getInvoiceRoot() {
    const candidates = [
      document.querySelector(
        '[id="sales-forms-ui/edit_or_preview_form"]'
      ),
      document.querySelector(
        '[data-id="editForm"]'
      ),
      document.querySelector(
        '[data-automation-id="RethinkLayout"]'
      ),
      document.querySelector(
        ".trowser-view .body"
      ),
      document.querySelector(
        '[data-automation-id="invoice-form"]'
      ),
      document.querySelector(
        '[data-automation-id="invoice-editor"]'
      ),
      document.querySelector("#qbo-main"),
      document.querySelector("#app"),
      document.querySelector("main"),
      document.body,
    ].filter(Boolean);

    for (const candidate of candidates) {
      const hasCustomer =
        candidate.querySelector(
          'input[aria-label="Customer"]'
        ) ||
        candidate.querySelector(
          '[data-cy="quickfill-contact"]'
        );

      const hasInvoiceNumber =
        candidate.querySelector(
          '[data-automation-id="readonly_reference_number"]'
        );

      const hasInvoiceTable =
        candidate.querySelector(
          'tbody[data-smart-table-body="true"]'
        ) ||
        candidate.querySelector(
          'input[aria-label^="Product or service line"]'
        );

      if (
        hasCustomer ||
        hasInvoiceNumber ||
        hasInvoiceTable
      ) {
        return candidate;
      }
    }

    return null;
  }

  function isInvoiceEditorOpen() {
    return (
      isInvoicePage() &&
      Boolean(getInvoiceRoot())
    );
  }

  // =========================================================
  // Invoice header extraction
  // =========================================================

  function extractCustomerName(root) {
    const customerSelectors = [
      'input[aria-label="Customer"]',
      '[data-cy="quickfill-contact"] input[role="combobox"]',
      '.qf-contact input[data-testid="__textField"]',
      '.qf-contact input',
    ];

    const customerInput =
      queryFirst(root, customerSelectors) ||
      queryFirst(document, customerSelectors);

    const customerInputValue =
      getInputValue(customerInput);

    if (customerInputValue) {
      return customerInputValue;
    }

    const quickfill =
      root?.querySelector(
        '[data-cy="quickfill-contact"] .QuickfillsContainer[data-props]'
      ) ||
      document.querySelector(
        '[data-cy="quickfill-contact"] .QuickfillsContainer[data-props]'
      );

    const rawProperties =
      quickfill?.getAttribute("data-props");

    if (rawProperties) {
      try {
        const parsedProperties =
          JSON.parse(rawProperties);

        const displayName = normalizeText(
          parsedProperties?.contact?.displayName ||
            ""
        );

        if (displayName) {
          return displayName;
        }
      } catch (error) {
        warn(
          "Could not parse customer QuickFill data.",
          error
        );
      }
    }

    return "";
  }

  function extractCustomFieldValue(
    root,
    targetLabel
  ) {
    const normalizedTarget =
      normalizeText(targetLabel).toLowerCase();

    /*
     * QuickBooks may render the custom-field form outside the
     * invoice root selected above. Search the invoice root first
     * and then the whole document.
     */
    const searchRoots = [];

    if (root) {
      searchRoots.push(root);
    }

    if (root !== document) {
      searchRoots.push(document);
    }

    for (const searchRoot of searchRoots) {
      const fields = Array.from(
        searchRoot.querySelectorAll(
          ".custom-form-field"
        )
      );

      for (const field of fields) {
        const labelCandidates = Array.from(
          field.querySelectorAll(
            [
              '[class*="RethinkCFLabel"]',
              ".custom-field-input > div > div:first-child",
              ".custom-field-input div",
            ].join(",")
          )
        );

        const matchingLabel =
          labelCandidates.find((element) => {
            return (
              normalizeText(
                element.textContent
              ).toLowerCase() ===
              normalizedTarget
            );
          });

        if (!matchingLabel) {
          continue;
        }

        /*
         * Read only the nested input. Do not use the wrapper's
         * textContent because that includes the label.
         */
        const input =
          field.querySelector(
            ".custom-fields-input-text-field input"
          ) ||
          field.querySelector(
            ".custom-field-input input"
          ) ||
          field.querySelector(
            "input, textarea, select"
          );

        if (!input) {
          return "";
        }

        return getInputValue(input);
      }
    }

    /*
     * Extra fallback using the stable custom-field definition IDs
     * present in the input names.
     */
    const knownFieldSuffixes = {
      "order number": "1700000000000127549",
      "job name": "1700000000000127550",
      phone: "1700000000000494174",
    };

    const fieldSuffix =
      knownFieldSuffixes[normalizedTarget];

    if (fieldSuffix) {
      const fallbackInput =
        document.querySelector(
          `.custom-form-field input[name$="${fieldSuffix}"]`
        ) ||
        document.querySelector(
          `input[name$="${fieldSuffix}"]`
        );

      return getInputValue(fallbackInput);
    }

    return "";
  }

  function extractHeaderData(root) {
    const customerName =
      extractCustomerName(root);

    const rawBillingAddress =
      getValueFromRootOrDocument(root, [
        'textarea[aria-label="billToTextAreaLabel"]',
        '[class*="billToAddress"] textarea',
      ]);

    const billingAddress =
      joinCustomerAndAddress(
        customerName,
        rawBillingAddress
      );

    const shippingAddress =
      getValueFromRootOrDocument(root, [
        'textarea[aria-label="Ship to"]',
        'textarea[aria-label="shipToTextAreaLabel"]',
        '[class*="shipToAddress"] textarea',
        '[class*="shipTo-"] textarea',
      ]) || "N/A";

    const invoiceNumberElement =
      queryFirst(root, [
        '[data-automation-id="readonly_reference_number"] span',
        '[data-automation-id="readonly_reference_number"]',
      ]) ||
      queryFirst(document, [
        '[data-automation-id="readonly_reference_number"] span',
        '[data-automation-id="readonly_reference_number"]',
      ]);

    const invoiceNumber =
      normalizeText(
        invoiceNumberElement?.textContent || ""
      ) || "N/A";

    const invoiceDate =
      getValueFromRootOrDocument(root, [
        'input[data-testid="txn_date"]',
        'input[aria-label="Invoice date"]',
      ]) || "N/A";

    const orderNumber =
      extractCustomFieldValue(
        root,
        "ORDER NUMBER"
      );

    const jobName =
      extractCustomFieldValue(
        root,
        "JOB NAME"
      );

    const phoneNumber =
      extractCustomFieldValue(
        root,
        "Phone"
      );

    log("Extracted header data:", {
      customerName,
      billingAddress,
      shippingAddress,
      invoiceNumber,
      invoiceDate,
      orderNumber,
      jobName,
      phoneNumber,
    });

    return {
      customerName,
      billingAddress,
      shippingAddress,
      invoiceNumber,
      invoiceDate,
      orderNumber,
      jobName,
      phoneNumber,
    };
  }

  // =========================================================
  // Invoice line extraction
  // =========================================================

  function findInvoiceTable(root) {
    const searchRoots = [root, document].filter(
      Boolean
    );

    for (const searchRoot of searchRoots) {
      const tables = Array.from(
        searchRoot.querySelectorAll("table")
      );

      for (const table of tables) {
        const headers = Array.from(
          table.querySelectorAll("thead th")
        ).map((header) =>
          normalizeText(
            header.textContent
          ).toLowerCase()
        );

        const hasProduct = headers.some(
          (header) =>
            /product\s*\/?\s*service/.test(
              header
            )
        );

        const hasSku = headers.some(
          (header) => header === "sku"
        );

        const hasDescription = headers.some(
          (header) =>
            header === "description"
        );

        const hasQuantity = headers.some(
          (header) =>
            header === "qty" ||
            header === "quantity"
        );

        if (
          hasProduct &&
          hasSku &&
          hasDescription &&
          hasQuantity
        ) {
          return table;
        }
      }
    }

    return null;
  }

  function buildHeaderIndexMap(table) {
    const map = {};

    const headers = Array.from(
      table.querySelectorAll("thead th")
    );

    headers.forEach((header, index) => {
      const label = normalizeText(
        header.textContent
      ).toLowerCase();

      if (
        /product\s*\/?\s*service/.test(label)
      ) {
        map.product = index;
      } else if (label === "sku") {
        map.sku = index;
      } else if (label === "description") {
        map.description = index;
      } else if (
        label === "qty" ||
        label === "quantity"
      ) {
        map.quantity = index;
      }
    });

    return map;
  }

  function extractRows(root) {
    const table = findInvoiceTable(root);

    if (!table) {
      warn(
        "QuickBooks invoice product table was not found."
      );

      return [];
    }

    const headerMap =
      buildHeaderIndexMap(table);

    if (
      headerMap.product == null ||
      headerMap.sku == null ||
      headerMap.description == null ||
      headerMap.quantity == null
    ) {
      warn(
        "Required invoice columns were not found.",
        headerMap
      );

      return [];
    }

    const tableBody =
      table.querySelector(
        'tbody[data-smart-table-body="true"]'
      ) ||
      table.querySelector("tbody");

    if (!tableBody) {
      warn(
        "QuickBooks invoice table body was not found."
      );

      return [];
    }

    const rowElements = Array.from(
      tableBody.querySelectorAll(
        'tr[data-automation-id^="line "]'
      )
    );

    const rows = [];

    for (const rowElement of rowElements) {
      const cells = Array.from(
        rowElement.querySelectorAll(
          ":scope > td"
        )
      );

      if (!cells.length) {
        continue;
      }

      const productName =
        getCellValue(
          cells[headerMap.product]
        );

      const sku =
        getCellValue(cells[headerMap.sku]);

      const description =
        getCellValue(
          cells[headerMap.description]
        );

      const rawQuantity =
        getCellValue(
          cells[headerMap.quantity]
        );

      const parsedQuantity =
        tryParseNumber(rawQuantity);

      const quantity =
        parsedQuantity == null
          ? 0
          : parsedQuantity;

      if (
        !productName &&
        !sku &&
        !description
      ) {
        continue;
      }

      rows.push({
        productName,
        sku,
        description,
        quantity,
      });
    }

    log("Extracted invoice rows:", rows);

    return rows;
  }

  function extractData() {
    const root = getInvoiceRoot();

    if (!root) {
      return {
        customerName: "",
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

  // =========================================================
  // Stable extraction
  // =========================================================

  function createDataSignature(data) {
    return JSON.stringify({
      customerName: data.customerName,
      billingAddress: data.billingAddress,
      shippingAddress: data.shippingAddress,
      invoiceNumber: data.invoiceNumber,
      invoiceDate: data.invoiceDate,
      orderNumber: data.orderNumber,
      jobName: data.jobName,
      phoneNumber: data.phoneNumber,
      rows: data.rows.map((row) => ({
        productName: row.productName,
        sku: row.sku,
        description: row.description,
        quantity: row.quantity,
      })),
    });
  }

  async function getStableExtractedData() {
    let previousSignature = "";
    let latestData = extractData();

    for (
      let attempt = 0;
      attempt < CONFIG.stableReadAttempts;
      attempt++
    ) {
      latestData = extractData();

      const currentSignature =
        createDataSignature(latestData);

      if (
        latestData.rows.length > 0 &&
        currentSignature ===
          previousSignature
      ) {
        return latestData;
      }

      previousSignature =
        currentSignature;

      await sleep(
        CONFIG.stableReadDelayMs
      );
    }

    return latestData;
  }

  // =========================================================
  // Product-table generation
  // =========================================================

  function isCategoryDescription(
    description
  ) {
    return /^\*+\s*.+?\s*\*+$/.test(
      normalizeText(description)
    );
  }

  function formatQuantity(quantity) {
    const number = Number(quantity);

    if (
      !Number.isFinite(number) ||
      number === 0
    ) {
      return "";
    }

    return Number.isInteger(number)
      ? String(number)
      : String(number);
  }

  function buildProductTable(
    rows,
    combineQuantities
  ) {
    let tableRows = "";

    if (combineQuantities) {
      /*
       * Pick Slip:
       * - exclude category/description-only rows
       * - combine identical SKUs
       * - Product Name comes from Product/service
       */
      const groupedProducts = new Map();

      for (const row of rows) {
        const categoryRow =
          !row.productName &&
          !row.sku &&
          isCategoryDescription(
            row.description
          );

        if (categoryRow) {
          continue;
        }

        if (
          !row.productName &&
          !row.sku
        ) {
          continue;
        }

        const quantity =
          Number(row.quantity);

        if (!Number.isFinite(quantity)) {
          continue;
        }

        const normalizedSku =
          normalizeText(row.sku);

        const normalizedName =
          normalizeText(row.productName);

        const key = normalizedSku
          ? `SKU:${normalizedSku.toUpperCase()}`
          : `NAME:${normalizedName.toLowerCase()}`;

        if (
          !groupedProducts.has(key)
        ) {
          groupedProducts.set(key, {
            productName: normalizedName,
            sku: normalizedSku,
            quantity,
          });
        } else {
          groupedProducts.get(
            key
          ).quantity += quantity;
        }
      }

      for (const product of groupedProducts.values()) {
        tableRows += `
          <tr>
            <td>${escapeHtml(
              product.productName
            )}</td>
            <td>${escapeHtml(
              product.sku
            )}</td>
            <td class="quantity-cell">${escapeHtml(
              formatQuantity(
                product.quantity
              )
            )}</td>
          </tr>
        `;
      }
    } else {
      /*
       * Print:
       * - Product/service for normal products
       * - Description only for category headers such as ** Kitchen **
       */
      for (const row of rows) {
        let displayName =
          normalizeText(row.productName);

        const categoryRow =
          !displayName &&
          !normalizeText(row.sku) &&
          isCategoryDescription(
            row.description
          );

        if (categoryRow) {
          displayName =
            normalizeText(
              row.description
            );
        }

        if (
          !displayName &&
          !normalizeText(row.sku)
        ) {
          continue;
        }

        tableRows += `
          <tr class="${
            categoryRow
              ? "category-row"
              : ""
          }">
            <td>${escapeHtml(
              displayName
            )}</td>
            <td>${escapeHtml(
              row.sku
            )}</td>
            <td class="quantity-cell">${escapeHtml(
              categoryRow
                ? ""
                : formatQuantity(
                    row.quantity
                  )
            )}</td>
          </tr>
        `;
      }
    }

    if (!tableRows) {
      return (
        "<p>No valid invoice products were found.</p>"
      );
    }

    return `
      <table class="product-table">
        <thead>
          <tr>
            <th>Product Name</th>
            <th>SKU</th>
            <th class="quantity-heading">Quantity</th>
          </tr>
        </thead>
        <tbody>
          ${tableRows}
        </tbody>
      </table>
    `;
  }

  // =========================================================
  // Print layout
  // =========================================================

  function generatePrintLayout(
    data,
    productTable
  ) {
    return `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Invoice ${escapeHtml(
    data.invoiceNumber
  )}</title>

  <style>
    * {
      box-sizing: border-box;
    }

    html,
    body {
      margin: 0;
      padding: 0;
    }

    body {
      font-family: Arial, sans-serif;
      font-size: 13px;
      line-height: 1.35;
      color: #000;
    }

    .maincontainer {
      width: 100%;
      margin: 0 auto;
      padding: 14px;
    }

    .top-header {
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 20px;
    }

    .company-info {
      width: 68%;
    }

    .company-info div:first-child {
      font-weight: 700;
    }

    .received-box {
      width: 230px;
      font-size: 10px;
    }

    .received-title {
      font-weight: 700;
      margin-bottom: 2px;
    }

    .received-box textarea {
      display: block;
      width: 100%;
      resize: none;
      padding: 3px;
      font-family: Arial, sans-serif;
      font-size: 11px;
      border: 1px solid #777;
      border-bottom: 0;
    }

    .received-box textarea:last-child {
      border-bottom: 1px solid #777;
    }

    .delivery-title {
      margin: 17px 0 10px;
      font-size: 20px;
      line-height: 1;
    }

    .address-grid {
      display: grid;
      grid-template-columns:
        1fr 1fr 0.75fr;
      gap: 28px;
      width: 100%;
    }

    .address-heading,
    .order-heading {
      font-weight: 700;
      margin-bottom: 7px;
    }

    .address-value {
      white-space: pre-line;
      min-height: 65px;
    }

    .invoice-summary {
      line-height: 1.4;
    }

    .separator {
      border: 0;
      border-top: 1px solid #bbb;
      margin: 10px 0;
    }

    .order-grid {
      display: grid;
      grid-template-columns:
        1fr 1fr 0.6fr;
      gap: 28px;
      padding: 2px 0 9px;
      align-items: start;
    }

    
.order-grid > div:nth-child(1) {
    text-align: left;
}


.order-grid > div:nth-child(2) {
    text-align: left;         
    justify-self: center;      
    width: 100%;
    max-width: 420px;
}


.order-grid > div:nth-child(3) {
    text-align: right;         
    justify-self: end;
}

    .order-value {
      min-height: 17px;
    }

    .products-section {
      margin-top: 8px;
      border-top: 1px solid #000;
      padding-top: 12px;
    }

    .product-table {
      width: 100%;
      border-collapse: collapse;
      table-layout: fixed;
    }

    .product-table th {
      background: #d2d2d2;
      padding: 8px;
      font-weight: 700;
      text-align: left;
    }

    .product-table th:nth-child(1),
    .product-table td:nth-child(1) {
      width: 77%;
    }

    .product-table th:nth-child(2),
    .product-table td:nth-child(2) {
      width: 15%;
    }

    .product-table th:nth-child(3),
    .product-table td:nth-child(3) {
      width: 8%;
    }

    .product-table td {
      padding: 6px 8px;
      vertical-align: top;
    }

    .quantity-heading,
    .quantity-cell {
      text-align: right !important;
    }

    .category-row td {
      font-weight: 400;
      padding-top: 8px;
      padding-bottom: 3px;
    }

    .signoff {
      display: flex;
      justify-content: space-between;
      border-top: 1px solid #ccc;
      margin-top: 10px;
      padding-top: 15px;
      font-size: 12px;
    }

    @media print {
      @page {
        size: portrait;
        margin: 8mm;
      }

      .maincontainer {
        padding: 0;
      }

      .product-table thead {
        display: table-header-group;
      }

      .product-table tr {
        break-inside: avoid;
      }
    }
  </style>
</head>

<body>
  <div class="maincontainer">
    <div class="top-header">
      <div class="company-info">
        <div>
          Adelaide Bathroom &amp;
          Kitchen Supplies
        </div>
        <div>
          2/831 Lower North East Rd,
          Dernancourt
        </div>
        <div>(08) 7006 5181</div>
        <div>
          Sales@abksupplies.com.au
        </div>
        <div>ABN 13 695 032 804</div>
      </div>

      <div class="received-box">
        <div class="received-title">
          Received In Good Order &amp;
          Condition
        </div>

        <textarea
          rows="1"
          placeholder="Name:"
        ></textarea>

        <textarea
          rows="2"
          placeholder="Sign:"
        ></textarea>

        <textarea
          rows="1"
          placeholder="Date: __ / __ / ____"
        ></textarea>
      </div>
    </div>

    <h2 class="delivery-title">
      Delivery Note
    </h2>

    <div class="address-grid">
      <div>
        <div class="address-heading">
          INVOICE TO
        </div>

        <div class="address-value">${escapeHtml(
          data.billingAddress
        )}</div>
      </div>

      <div>
        <div class="address-heading">
          SHIP TO
        </div>

        <div class="address-value">${escapeHtml(
          data.shippingAddress
        )}</div>
      </div>

      <div class="invoice-summary">
        <div>
          <strong>INVOICE NO.:</strong>${escapeHtml(
            data.invoiceNumber
          )}
        </div>

        <div>
          <strong>DATE:</strong>${escapeHtml(
            data.invoiceDate
          )}
        </div>
      </div>
    </div>

    <hr class="separator">

    <div class="order-grid">
      <div>
        <div class="order-heading">
          ORDER NUMBER
        </div>

        <div class="order-value">${escapeHtml(
          data.orderNumber || ""
        )}</div>
      </div>

      <div>
        <div class="order-heading">
          JOB NAME
        </div>

        <div class="order-value">${escapeHtml(
          data.jobName || ""
        )}</div>
      </div>

      <div>
        <div class="order-heading">
          PHONE
        </div>

        <div class="order-value">${escapeHtml(
          data.phoneNumber || ""
        )}</div>
      </div>
    </div>

    <div class="products-section">
      ${productTable}
    </div>

    <div class="signoff">
      <span>
        Picked By: _______________
      </span>

      <span>
        Checked By: _______________
      </span>
    </div>
  </div>
</body>
</html>
    `;
  }

  // =========================================================
  // Print execution
  // =========================================================

  async function generateProductTable(
    combineQuantities
  ) {
    if (STATE.printing) return;

    STATE.printing = true;

    try {
      const data =
        await getStableExtractedData();

      if (!data.rows.length) {
        alert(
          "No invoice line items were found. Please allow the invoice to finish loading and try again."
        );

        return;
      }

      const productTable =
        buildProductTable(
          data.rows,
          combineQuantities
        );

      const printLayout =
        generatePrintLayout(
          data,
          productTable
        );

      const printWindow = window.open(
        "",
        "_blank",
        "width=1000,height=760"
      );

      if (!printWindow) {
        alert(
          "The print window was blocked. Please allow popups for qbo.intuit.com."
        );

        return;
      }

      printWindow.document.open();
      printWindow.document.write(
        printLayout
      );
      printWindow.document.close();

      const triggerPrint = () => {
        setTimeout(() => {
          printWindow.focus();
          printWindow.print();
        }, 350);
      };

      if (
        printWindow.document.readyState ===
        "complete"
      ) {
        triggerPrint();
      } else {
        printWindow.addEventListener(
          "load",
          triggerPrint,
          { once: true }
        );
      }
    } finally {
      STATE.printing = false;
    }
  }

  // =========================================================
  // Custom buttons
  // =========================================================

  function createButton(
    id,
    text,
    clickHandler,
    left
  ) {
    const button =
      document.createElement("button");

    button.id = id;
    button.type = "button";
    button.textContent = text;

    button.style.cssText = `
      position: fixed;
      bottom: 4px;
      left: ${left};
      min-width: 112px;
      height: 32px;
      padding: 4px 16px;
      background: #2ca01c;
      color: #fff;
      border: 0;
      border-radius: 5px;
      box-shadow:
        0 2px 8px rgba(0,0,0,.12);
      cursor: pointer;
      font-family: Arial, sans-serif;
      font-size: 14px;
      font-weight: 600;
      line-height: 24px;
      z-index: 2147483647;
      transition:
        background-color .15s ease,
        transform .15s ease,
        box-shadow .15s ease;
    `;

    button.addEventListener(
      "mouseenter",
      () => {
        button.style.backgroundColor =
          "#248f17";

        button.style.transform =
          "translateY(-1px)";

        button.style.boxShadow =
          "0 4px 12px rgba(0,0,0,.16)";
      }
    );

    button.addEventListener(
      "mouseleave",
      () => {
        button.style.backgroundColor =
          "#2ca01c";

        button.style.transform =
          "translateY(0)";

        button.style.boxShadow =
          "0 2px 8px rgba(0,0,0,.12)";
      }
    );

    button.addEventListener(
      "click",
      clickHandler
    );

    return button;
  }

  function removeButtons() {
    document
      .getElementById(
        "custom-print-button"
      )
      ?.remove();

    document
      .getElementById(
        "custom-pick-slip-button"
      )
      ?.remove();
  }

  async function addButtons() {
    if (STATE.addButtonsInFlight) {
      return;
    }

    STATE.addButtonsInFlight = true;

    try {
      if (
        !isInvoicePage() ||
        !isInvoiceEditorOpen()
      ) {
        removeButtons();
        return;
      }

      const invoiceId =
        getInvoiceId();

      if (
        invoiceId &&
        invoiceId !==
          STATE.currentInvoiceId
      ) {
        STATE.currentInvoiceId =
          invoiceId;

        removeButtons();
      }

      if (
        !document.getElementById(
          "custom-print-button"
        )
      ) {
        document.body.appendChild(
          createButton(
            "custom-print-button",
            "🖨️ Print",
            () =>
              generateProductTable(false),
            "14%"
          )
        );
      }

      if (
        !document.getElementById(
          "custom-pick-slip-button"
        )
      ) {
        document.body.appendChild(
          createButton(
            "custom-pick-slip-button",
            "📋 Pick Slip",
            () =>
              generateProductTable(true),
            "calc(14% + 125px)"
          )
        );
      }
    } finally {
      STATE.addButtonsInFlight = false;
    }
  }

  // =========================================================
  // QuickBooks SPA navigation handling
  // =========================================================

  function refreshButtons() {
    clearTimeout(STATE.mutationTimer);

    STATE.mutationTimer =
      setTimeout(() => {
        if (
          isInvoicePage() &&
          isInvoiceEditorOpen()
        ) {
          addButtons();
        } else {
          removeButtons();
        }
      }, CONFIG.mutationDebounceMs);
  }

  function setupObservers() {
    const originalPushState =
      history.pushState;

    history.pushState = function (
      ...args
    ) {
      const result =
        originalPushState.apply(
          history,
          args
        );

      setTimeout(
        refreshButtons,
        500
      );

      return result;
    };

    const originalReplaceState =
      history.replaceState;

    history.replaceState = function (
      ...args
    ) {
      const result =
        originalReplaceState.apply(
          history,
          args
        );

      setTimeout(
        refreshButtons,
        500
      );

      return result;
    };

    window.addEventListener(
      "popstate",
      () => {
        setTimeout(
          refreshButtons,
          500
        );
      }
    );

    const observer =
      new MutationObserver(() => {
        refreshButtons();
      });

    observer.observe(document.body, {
      childList: true,
      subtree: true,
    });
  }

  // =========================================================
  // Initialise
  // =========================================================

  setupObservers();

  setInterval(() => {
    if (
      isInvoicePage() &&
      isInvoiceEditorOpen()
    ) {
      addButtons();
    } else {
      removeButtons();
    }
  }, CONFIG.buttonCheckIntervalMs);

  setTimeout(() => {
    if (
      isInvoicePage() &&
      isInvoiceEditorOpen()
    ) {
      addButtons();
    }
  }, CONFIG.initialBootDelayMs);
})();
