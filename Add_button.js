// ==UserScript==
// @name         Add Print and Pick Slip Buttons to QuickBooks Invoice
// @namespace    http://tampermonkey.net/
// @version      1.91
// @description  Adds "Print" and "Pick Slip" buttons to QuickBooks Invoice page, checks URL every minute to ensure buttons are removed if URL no longer matches
// @author       Raj - Gorkhari
// @match        https://qbo.intuit.com/*
// @include      https://qbo.intuit.com/app/invoice?*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // Function to check if the current page is an invoice page
    function isInvoicePage() {
        const url = window.location.href;
        const invoiceSelectors = [
            '.invoice-details-container',
            '[data-qbo-bind="text: referenceNumber"]',
            '.dgrid-row .itemColumn'
        ];
        return (
            url.includes('qbo.intuit.com/app/invoice?') ||
            url.includes('/invoice?txnId') ||
            invoiceSelectors.some(selector => document.querySelector(selector) !== null)
        );
    }

    // Function to add the "Print" and "Pick Slip" buttons
    function addButtons() {
        if (!isInvoicePage()) return;

        // Create buttons (example buttons, adjust as needed)
        const printButton = document.createElement('button');
        printButton.textContent = 'Print';
        printButton.id = 'printButton';
        printButton.onclick = function() { window.print(); };

        const pickSlipButton = document.createElement('button');
        pickSlipButton.textContent = 'Pick Slip';
        pickSlipButton.id = 'pickSlipButton';
        pickSlipButton.onclick = function() { alert('Pick Slip action'); };

        // Add buttons to the page (adjust selector as needed)
        const container = document.querySelector('.invoice-details-container');
        if (container) {
            container.appendChild(printButton);
            container.appendChild(pickSlipButton);
        }
    }

    // Function to remove the buttons
    function removeButtons() {
        const printButton = document.getElementById('printButton');
        const pickSlipButton = document.getElementById('pickSlipButton');
        if (printButton) printButton.remove();
        if (pickSlipButton) pickSlipButton.remove();
    }

    // Delay initialization to ensure page is fully loaded
    function safeInitialize() {
        if (isInvoicePage()) {
            // Add buttons if on invoice page
            setTimeout(addButtons, 2000);
        } else {
            // Remove buttons if not on invoice page
            removeButtons();
        }
    }

    // Setup observers for history changes, popstate, and mutations
    function setupObservers() {
        // History API listener
        (function(history){
            const pushState = history.pushState;
            history.pushState = function(state) {
                const result = pushState.apply(history, arguments);
                window.dispatchEvent(new Event('urlchange'));
                return result;
            };
        })(window.history);

        // Event listeners
        window.addEventListener('urlchange', safeInitialize);
        window.addEventListener('popstate', safeInitialize);

        // Mutation observer for detecting page changes
        const observer = new MutationObserver(safeInitialize);
        observer.observe(document.body, {
            childList: true,
            subtree: true
        });
    }

    // Periodic URL check every minute
    setInterval(function() {
        if (isInvoicePage()) {
            addButtons();
        } else {
            removeButtons();
        }
    }, 60000); // Check every 60,000 milliseconds (1 minute)

    // Initial setup
    setupObservers();
    safeInitialize();
    // Function to create a button
    function createButton(id, text, clickHandler) {
        const button = document.createElement('button');
        button.id = id;
        button.textContent = text;
        button.style.position = 'fixed';
        button.style.bottom = '8px';
        button.style.left = id === 'custom-print-button' ? '400px' : '500px';
        button.style.padding = '10px 20px';
        button.style.backgroundColor = '#00ab00';
        button.style.color = 'white';
        button.style.border = 'none';
        button.style.borderRadius = '5px';
        button.style.boxShadow = '0 2px 5px rgba(0,0,0,0.3)';
        button.style.cursor = 'pointer';
        button.style.zIndex = '1000';
        button.addEventListener('click', clickHandler);
        return button;
    }

    // Function to add the buttons to the page
    function addButtons() {
        if (document.getElementById('custom-print-button') || document.getElementById('custom-pick-slip-button')) return;
        const printButtonClickHandler = () => generateProductTable(false);
        const pickSlipButtonClickHandler = () => generateProductTable(true);
        const printButton = createButton('custom-print-button', 'Print', printButtonClickHandler);
        document.body.appendChild(printButton);
        const pickSlipButton = createButton('custom-pick-slip-button', 'Pick Slip', pickSlipButtonClickHandler);
        document.body.appendChild(pickSlipButton);
    }

    function generateProductTable(combineQuantities) {
        const rows = document.querySelectorAll('.dgrid-row');
        const billingAddress = document.querySelector('textarea.topFieldInput.address')?.value || '';
        const shippingAddress = document.getElementById('shippingAddress')?.value || '';
        const referenceNumberElement = document.querySelector('[data-qbo-bind="text: referenceNumber"]');
        const invoiceNumber = referenceNumberElement?.textContent?.trim() || '';
        const invoiceDateInput = document.querySelector('.dijitDateTextBox input.dijitInputInner');
        const invoiceDate = invoiceDateInput ? invoiceDateInput.value : '';

        // Get form details
        const formElement = document.querySelector('.custom-form');
        const formFields = Array.from(formElement.querySelectorAll('.custom-form-field'));
        const orderNumber = formFields.find(field => field.querySelector('label')?.textContent.trim() === 'ORDER NUMBER')?.querySelector('input')?.value || '';
        const jobName = formFields.find(field => field.querySelector('label')?.textContent.trim() === 'JOB NAME')?.querySelector('input')?.value || '';
        const phoneNumber = formFields.find(field => field.querySelector('label')?.textContent.trim() === 'Phone')?.querySelector('input')?.value || '';
        const skuMap = new Map();
        let productTable = '';
        rows.forEach(row => {
            const productNameElement = row.querySelector('.itemColumn');
            const descriptionElement = row.querySelector('.field-description div');
            const productName = productNameElement?.textContent.trim() || "";
            const description = descriptionElement?.textContent.trim() || "";
            const sku = row.querySelector('.field-sku')?.textContent.trim() || "";
            const quantityText = row.querySelector('.field-quantity-inner')?.textContent.trim() || "";
            const quantity = parseInt(quantityText, 10) || '';

            if (combineQuantities) {
                // Combine quantities for Pick Slip
                if (skuMap.has(sku)) {
                    const existing = skuMap.get(sku);
                    existing.quantity += quantity;
                } else {
                    skuMap.set(sku, { productName, quantity });
                }
            } else {
                // Print logic: Include description only if both product name and SKU are missing
                const displayName = productName || (sku ? '' : description);

                if (displayName || sku) {
                    productTable += `
                        <tr style="margin-bottom: 5px; height: 30px;">
                            <td>${displayName}</td>
                            <td>${sku}</td>
                            <td style="text-align:right;">${quantity}</td>
                        </tr>
                    `;
                }
            }
        });

        if (combineQuantities) {
            // Generate table for Pick Slip
            skuMap.forEach((value, key) => {
                productTable += `
                    <tr style="margin-bottom: 5px; height: 30px;">
                        <td>${value.productName}</td>
                        <td>${key}</td>
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
            productTable = '<p>No valid products found.</p>';
        }

const printLayout = `
           <html>
                    <head>
<style>
         .product-table tr {
                margin-bottom: 10px;
                height: 30px;
            }

        body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
        }

        .maincontainer {
            width: 100%;
            box-sizing: border-box;
            padding: 0 15px;
            margin: 20px auto;
        }

        .TMheader-left {
              width:70%;
              display: flex;
              font-size: 13px;
          }

        .TMheader {
            width:100%;
            display: flex;
            justify-content: space-between;
            margin-bottom: -25px;
        }

        .TMheader-left img {
            max-width: 150px;
        }

        .TMheader-right {
           width: 30%;
           font-size: 10px;
         }

        .deliveryNote {
         display: flex;
         margin-bottom: -30px;
         white-space: pre-line;
         flex-direction: row;
         justify-content: space-between;
         font-size: 13px;
         width: 100%;
        }

        .deliveryNote-1, .deliveryNote-2, .deliveryNote-3 {
            width: 33.3%;
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
        }
        .product-table tbody tr td:first-child {
            padding-right: 10px;
        }
        .product-table tbody tr td:last-child {
            text-align: right;
        }
        .orderNote {
          display: flex;
          justify-content: space-between;
          font-size: 14px;
        }
         .footer {
           position: relative;
           bottom: 0px;
           display: flex;
           flex-direction: column;
           align-items: center;
           font-size: 12px;
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
                <div class="CompnayInfo">
                    <div><b>Adelaide Bathroom & Kitchen Supplies</b></div>
                    <div>2/831 Lower North East Rd, Dernancourt</div>
                    <div>(08) 7006 5181</div>
                    <div>Sales@abksupplies.com.au</div>
                    <div>ABN 13 695 032 804</div>
                </div>
            </div>
            <div class="TMheader-right">
        <div class="QuoteHeaderLeft">
                                <div class="ContactSevice">
                                    <div><b>Received In Good Order & Condition</b></div>
                                    <div><textarea class="form-control" id="FormControlTextarea1" rows="1" style="height:30px;width:200px;" placeholder="Name:"></textarea></div>
                                    <div><textarea class="form-control" id="FormControlTextarea1" rows="1" style="height:40;width:200px;" placeholder="Sign:"></textarea></div>
                                    <div><textarea class="form-control" id="FormControlTextarea1" rows="1" style="height:30px;width:200px;"placeholder="Date: __ / __ / ____"></textarea></div>
                                </div>
                            </div>
            </div>
        </div>
        <h3 style="margin-bottom:0px">Delivery Note</h3>
        <div class="deliveryNote">
            <div class="deliveryNote-1">
                <b>INVOICE TO</b>
                <p class="InvoiceInfo">${billingAddress}</p>
            </div>
            <div class="deliveryNote-2">
                <b>SHIP TO</b>
                <p class="ShipInfo">${shippingAddress}</p>
            </div>
            <div class="deliveryNote-3">
                <b>INVOICE NO.:</b>
                <span class="InvoiceNumber">${invoiceNumber}</span>
                <b>DATE:</b>
                <span class="InvoiceDate">${invoiceDate}</span>
            </div>
        </div>
        <hr>
        <div class="orderNote">
            <div class="orderNote-1"><b>ORDER NUMBER</b><br/>${orderNumber}</div>
            <div class="orderNote-2"><b>JOB NAME<br/></b>${jobName}</div>
            <div class="orderNote-3"><b>PHONE<br/></b>${phoneNumber}</div>
        </div>
        <div class="orderProducts">
            ${productTable}
        </div>
        <hr/>
        <div class="input-group mb-3" style="display:flex; justify-content: space-between;">
  <span class="input-group-text">Picked By : _______________.</span>
  <span class="input-group-text">Checked By : _______________.</span>
</div>
    </div>
</body>
</html>
        `;

        const newWindow = window.open('', '_blank', 'width=800,height=600');
        newWindow.document.write(printLayout);
        newWindow.document.close();
        newWindow.print();
    }
})();
