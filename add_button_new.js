// ==UserScript==
// @name         Add Print and Pick Slip Buttons to QuickBooks Invoice
// @namespace    http://tampermonkey.net/
// @version      2.0
// @description  Adds "Print" and "Pick Slip" buttons to QuickBooks Invoice page with improved data loading
// @author       Raj - Gorkhari (Improved)
// @match        https://qbo.intuit.com/*
// @include      https://qbo.intuit.com/app/invoice?*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    let currentInvoiceId = null;
    const BUTTON_CHECK_INTERVAL = 5000; // Check every 5 seconds instead of 60
    const DATA_LOAD_TIMEOUT = 10000; // Wait up to 10 seconds for data
    const DATA_CHECK_INTERVAL = 500; // Check every 500ms

    // Function to check if the current page is an invoice page
    function isInvoicePage() {
        const url = window.location.href;
        return (
            url.includes('qbo.intuit.com/app/invoice?') ||
            url.includes('/invoice?txnId')
        );
    }

    // Function to extract invoice ID from URL
    function getInvoiceId() {
        const match = window.location.href.match(/txnId=([^&]+)/);
        return match ? match[1] : null;
    }

    // Function to wait for data to be loaded
    function waitForData(selector, timeout = DATA_LOAD_TIMEOUT) {
        return new Promise((resolve, reject) => {
            const startTime = Date.now();
            
            const checkInterval = setInterval(() => {
                const element = document.querySelector(selector);
                const elapsed = Date.now() - startTime;
                
                if (element && element.textContent.trim()) {
                    clearInterval(checkInterval);
                    resolve(true);
                } else if (elapsed >= timeout) {
                    clearInterval(checkInterval);
                    resolve(false); // Don't reject, just return false
                }
            }, DATA_CHECK_INTERVAL);
        });
    }

    // Function to check if invoice data is loaded
    async function isInvoiceDataLoaded() {
        // Check for multiple indicators that data is loaded
        const checks = [
            waitForData('.dgrid-row', 3000), // Product rows
            waitForData('[data-qbo-bind="text: referenceNumber"]', 3000), // Invoice number
            waitForData('textarea.topFieldInput.address', 3000) // Billing address
        ];

        const results = await Promise.all(checks);
        return results.some(result => result); // At least one should be loaded
    }

    // Function to create a button
    function createButton(id, text, clickHandler) {
        const button = document.createElement('button');
        button.id = id;
        button.textContent = text;
        button.style.cssText = `
            position: fixed;
            bottom: 20px;
            left: ${id === 'custom-print-button' ? '20px' : '150px'};
            padding: 12px 24px;
            background-color: #2ca01c;
            color: white;
            border: none;
            border-radius: 6px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.2);
            cursor: pointer;
            font-size: 14px;
            font-weight: 600;
            z-index: 10000;
            transition: all 0.3s ease;
        `;
        
        button.addEventListener('mouseenter', () => {
            button.style.backgroundColor = '#248f17';
            button.style.transform = 'translateY(-2px)';
            button.style.boxShadow = '0 4px 12px rgba(0,0,0,0.3)';
        });
        
        button.addEventListener('mouseleave', () => {
            button.style.backgroundColor = '#2ca01c';
            button.style.transform = 'translateY(0)';
            button.style.boxShadow = '0 2px 8px rgba(0,0,0,0.2)';
        });
        
        button.addEventListener('click', clickHandler);
        return button;
    }

    // Function to add the buttons to the page
    async function addButtons() {
        if (!isInvoicePage()) {
            removeButtons();
            return;
        }

        // Check if invoice ID has changed (new invoice)
        const invoiceId = getInvoiceId();
        if (invoiceId !== currentInvoiceId) {
            currentInvoiceId = invoiceId;
            removeButtons(); // Remove old buttons for new invoice
        }

        // Don't add if buttons already exist
        if (document.getElementById('custom-print-button')) {
            return;
        }

        // Wait for data to load before adding buttons
        const dataLoaded = await isInvoiceDataLoaded();
        if (!dataLoaded) {
            console.warn('QuickBooks data not fully loaded yet');
            return; // Will retry on next interval
        }

        const printButton = createButton(
            'custom-print-button', 
            '🖨️ Print', 
            () => generateProductTable(false)
        );
        document.body.appendChild(printButton);

        const pickSlipButton = createButton(
            'custom-pick-slip-button', 
            '📋 Pick Slip', 
            () => generateProductTable(true)
        );
        document.body.appendChild(pickSlipButton);
        
        console.log('Buttons added successfully');
    }

    // Function to remove the buttons
    function removeButtons() {
        const printButton = document.getElementById('custom-print-button');
        const pickSlipButton = document.getElementById('custom-pick-slip-button');
        if (printButton) printButton.remove();
        if (pickSlipButton) pickSlipButton.remove();
    }

    // Improved data extraction with fallbacks
    function extractData() {
        // Try multiple selectors for SKU as fallback
        function getSKU(row) {
            const selectors = [
                '.field-sku',
                '[data-automation-id="sku"]',
                '.sku-field',
                '.itemSKU'
            ];
            
            for (const selector of selectors) {
                const element = row.querySelector(selector);
                if (element && element.textContent.trim()) {
                    return element.textContent.trim();
                }
            }
            return "";
        }

        // Try multiple selectors for quantity
        function getQuantity(row) {
            const selectors = [
                '.field-quantity-inner',
                '[data-automation-id="quantity"]',
                '.quantity-field'
            ];
            
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
            billingAddress: document.querySelector('textarea.topFieldInput.address')?.value || 'N/A',
            shippingAddress: document.getElementById('shippingAddress')?.value || 'N/A',
            invoiceNumber: document.querySelector('[data-qbo-bind="text: referenceNumber"]')?.textContent?.trim() || 'N/A',
            invoiceDate: document.querySelector('.dijitDateTextBox input.dijitInputInner')?.value || 'N/A',
            rows: []
        };

        // Extract custom form fields
        const formElement = document.querySelector('.custom-form');
        if (formElement) {
            const formFields = Array.from(formElement.querySelectorAll('.custom-form-field'));
            data.orderNumber = formFields.find(f => f.querySelector('label')?.textContent.trim() === 'ORDER NUMBER')?.querySelector('input')?.value || '';
            data.jobName = formFields.find(f => f.querySelector('label')?.textContent.trim() === 'JOB NAME')?.querySelector('input')?.value || '';
            data.phoneNumber = formFields.find(f => f.querySelector('label')?.textContent.trim() === 'Phone')?.querySelector('input')?.value || '';
        } else {
            data.orderNumber = '';
            data.jobName = '';
            data.phoneNumber = '';
        }

        // Extract product rows
        const rows = document.querySelectorAll('.dgrid-row');
        rows.forEach(row => {
            const productNameElement = row.querySelector('.itemColumn');
            const descriptionElement = row.querySelector('.field-description div');
            
            const productName = productNameElement?.textContent.trim() || "";
            const description = descriptionElement?.textContent.trim() || "";
            const sku = getSKU(row);
            const quantity = getQuantity(row);

            // Only add if we have meaningful data
            if (productName || sku || description) {
                data.rows.push({
                    productName,
                    description,
                    sku,
                    quantity
                });
            }
        });

        return data;
    }

    function generateProductTable(combineQuantities) {
        const data = extractData();
        
        if (data.rows.length === 0) {
            alert('No product data found. Please ensure the invoice is fully loaded.');
            return;
        }

        const skuMap = new Map();
        let productTable = '';

        data.rows.forEach(row => {
            if (combineQuantities) {
                // Combine quantities for Pick Slip by SKU
                if (row.sku && skuMap.has(row.sku)) {
                    const existing = skuMap.get(row.sku);
                    existing.quantity += row.quantity;
                } else if (row.sku) {
                    skuMap.set(row.sku, { 
                        productName: row.productName, 
                        quantity: row.quantity 
                    });
                }
            } else {
                // Print logic: Include description only if both product name and SKU are missing
                const displayName = row.productName || (row.sku ? '' : row.description);

                if (displayName || row.sku) {
                    productTable += `
                        <tr style="margin-bottom: 5px; height: 30px;">
                            <td>${displayName}</td>
                            <td>${row.sku}</td>
                            <td style="text-align:right;">${row.quantity || ''}</td>
                        </tr>
                    `;
                }
            }
        });

        if (combineQuantities) {
            // Generate table for Pick Slip
            skuMap.forEach((value, sku) => {
                productTable += `
                    <tr style="margin-bottom: 5px; height: 30px;">
                        <td>${value.productName}</td>
                        <td>${sku}</td>
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

        const printLayout = generatePrintLayout(data, productTable);
        
        const newWindow = window.open('', '_blank', 'width=800,height=600');
        if (newWindow) {
            newWindow.document.write(printLayout);
            newWindow.document.close();
            newWindow.print();
        }
    }

    function generatePrintLayout(data, productTable) {
        return `
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Invoice ${data.invoiceNumber}</title>
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
        }

        .TMheader {
            width: 100%;
            display: flex;
            justify-content: space-between;
            margin-bottom: 20px;
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
            margin-bottom: 20px;
            font-size: 13px;
            width: 100%;
        }

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
        
        .product-table td {
            padding: 8px;
        }
        
        .product-table tbody tr td:last-child {
            text-align: right;
        }

        .orderNote {
            display: flex;
            justify-content: space-between;
            font-size: 14px;
            margin: 20px 0;
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
            margin: 15px 0;
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
                <p class="InvoiceInfo">${data.billingAddress}</p>
            </div>
            <div class="deliveryNote-2">
                <b>SHIP TO</b>
                <p class="ShipInfo">${data.shippingAddress}</p>
            </div>
            <div class="deliveryNote-3">
                <b>INVOICE NO.:</b>
                <span class="InvoiceNumber">${data.invoiceNumber}</span><br/>
                <b>DATE:</b>
                <span class="InvoiceDate">${data.invoiceDate}</span>
            </div>
        </div>
        
        <hr>
        
        <div class="orderNote">
            <div class="orderNote-1"><b>ORDER NUMBER</b><br/>${data.orderNumber}</div>
            <div class="orderNote-2"><b>JOB NAME</b><br/>${data.jobName}</div>
            <div class="orderNote-3"><b>PHONE</b><br/>${data.phoneNumber}</div>
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

    // Setup URL change detection
    function setupObservers() {
        // Override pushState
        const originalPushState = history.pushState;
        history.pushState = function() {
            const result = originalPushState.apply(this, arguments);
            setTimeout(addButtons, 1000);
            return result;
        };

        // Listen for popstate
        window.addEventListener('popstate', () => {
            setTimeout(addButtons, 1000);
        });

        // Mutation observer for dynamic content
        const observer = new MutationObserver(() => {
            if (isInvoicePage() && !document.getElementById('custom-print-button')) {
                addButtons();
            }
        });

        observer.observe(document.body, {
            childList: true,
            subtree: true
        });
    }

    // Periodic check
    setInterval(() => {
        if (isInvoicePage()) {
            addButtons();
        } else {
            removeButtons();
        }
    }, BUTTON_CHECK_INTERVAL);

    // Initialize
    setupObservers();
    setTimeout(addButtons, 2000);
})();