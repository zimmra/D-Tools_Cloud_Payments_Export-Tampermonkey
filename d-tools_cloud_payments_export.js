// ==UserScript==
// @name         D-Tools Cloud Billing Table to CSV & Excel Downloader
// @namespace    D-Tools
// @version      3.2
// @description  Add download CSV and Excel buttons for D-Tools Cloud billing table
// @author       Payton Zimmerer
// @match        https://d-tools.cloud/billing/home
// @grant        none
// @downloadURL  https://raw.githubusercontent.com/zimmra/D-Tools_Cloud_Payments_Export-Tampermonkey/refs/heads/main/d-tools_cloud_payments_export.js
// ==/UserScript==

(function() {
    'use strict';
    // Load ExcelJS library for Excel export
    function loadExcelJS(callback) {
        const script = document.createElement('script');
        script.src = "https://cdn.jsdelivr.net/npm/exceljs@4.3.0/dist/exceljs.min.js";
        script.onload = callback;
        document.head.appendChild(script);
    }

    // Function to clean text content
    function cleanText(text) {
        return text.replace(/\s+/g, ' ').trim();
    }

    // Function to extract cell content safely
    function extractCellContent(cell) {
        if (!cell) return '';
        const flexColumn = cell.querySelector('.flex-column');
        if (flexColumn) {
            return cleanText(flexColumn.textContent);
        }
        const link = cell.querySelector('a');
        if (link) {
            return cleanText(link.textContent);
        }
        return cleanText(cell.textContent);
    }

    // Function to format currency
    function formatCurrency(text) {
        return text.replace('$', '').replace(/,/g, '').trim() || '0.00';
    }

    async function gatherAllTableData(table) {
        const allData = [];
    
        // Get total number of pages
        const pageSpans = Array.from(document.querySelectorAll('span.page-link'))
            .filter(span => !span.classList.contains('disabled-link') && /^\d+$/.test(span.textContent.trim()));
        const totalPages = Math.max(...pageSpans.map(span => parseInt(span.textContent.trim())));
    
        // Get data from the first page
        let rows = Array.from(table.querySelectorAll('tbody tr'));
        allData.push(...rows);
    
        console.log(`Gathered ${rows.length} rows from page 1`);
    
        let currentPageNumber = 1;
    
        while (currentPageNumber < totalPages) {
            console.log(`Processing page ${currentPageNumber + 1} of ${totalPages}`);
    
            // Find the "Next" button
            const nextButton = document.querySelector('mat-icon[svgicon="keyboardArrowRight"]');
            if (!nextButton || nextButton.classList.contains('disabled-link')) {
                console.warn('Next button not found or disabled. Stopping navigation.');
                break;
            }
    
            nextButton.click();
    
            // Wait for the page to change
            let pageChanged = false;
            let attempts = 0;
            const maxAttempts = 20;
    
            while (!pageChanged && attempts < maxAttempts) {
                await new Promise(resolve => setTimeout(resolve, 1000)); // Wait 1 second
    
                const activePageSpan = document.querySelector('span.page-link.active');
                const activePageNum = parseInt(activePageSpan?.textContent.trim() || '0');
    
                if (activePageNum === currentPageNumber + 1) {
                    pageChanged = true;
                    currentPageNumber = activePageNum;
                    console.log(`Successfully navigated to page ${currentPageNumber}`);
                } else {
                    console.log(`Waiting for page to change... (attempt ${attempts + 1})`);
                }
    
                attempts++;
            }
    
            if (!pageChanged) {
                console.error('Failed to change to the next page after maximum attempts.');
                break;
            }
    
            // Get data from the current page
            rows = Array.from(document.querySelector('.table-container table').querySelectorAll('tbody tr'));
            allData.push(...rows);
    
            console.log(`Gathered ${rows.length} rows from page ${currentPageNumber}`);
        }
    
        // Ensure the last page is processed explicitly
        if (currentPageNumber === totalPages) {
            console.log('Processing the last page explicitly.');
            rows = Array.from(document.querySelector('.table-container table').querySelectorAll('tbody tr'));
            allData.push(...rows);
            console.log(`Gathered ${rows.length} rows from the last page (${currentPageNumber}).`);
        }
    
        // Return to page 1
        const firstPageButton = document.querySelector('span.page-link:not(.active):not(.disabled-link):first-child');
        if (firstPageButton) {
            firstPageButton.click();
            await new Promise(resolve => setTimeout(resolve, 2000));
        }
    
        console.log(`Total rows gathered: ${allData.length}`);
        return allData;
    }



    // Function to convert table data to CSV string
    async function tableToCSV(table) {
        const rows = await gatherAllTableData(table);
        const headers = [
            "Type", "Client", "Project/CO/Contract/Call", "Payment Term",
            "Billing Date", "Due Date", "Total Amount", "Requested", "Paid", "Status"
        ];

        let csvContent = headers.join(',') + '\n';

        rows.forEach(row => {
            const columns = row.querySelectorAll('td');
            if (columns.length === 10) {
                const rowData = [
                    extractCellContent(columns[0]),
                    extractCellContent(columns[1]),
                    extractCellContent(columns[2]),
                    extractCellContent(columns[3]),
                    extractCellContent(columns[4]),
                    extractCellContent(columns[5]),
                    formatCurrency(extractCellContent(columns[6])),
                    formatCurrency(extractCellContent(columns[7])),
                    formatCurrency(extractCellContent(columns[8])),
                    row.querySelector('.status-height-width span')?.textContent || ''
                ].map(text => {
                    text = text.replace(/"/g, '""');
                    return text.includes(',') || text.includes('"') || text.includes('\n')
                        ? `"${text}"`
                        : text;
                });

                csvContent += rowData.join(',') + '\n';
            }
        });

        return csvContent;
    }

    // Function to download CSV
    function downloadCSV(csvContent, filename) {
        const BOM = '\uFEFF';
        const blob = new Blob([BOM + csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);

        link.setAttribute('href', url);
        link.setAttribute('download', filename);
        link.style.visibility = 'hidden';

        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    }

    // Function to convert table data to Excel format based on provided Excel file's styles
    async function tableToExcelData(table) {
        const rows = await gatherAllTableData(table);
        const headers = [
            "Type", "Client", "Project/CO/Contract/Call", "Payment Term",
            "Billing Date", "Due Date", "Total Amount", "Requested", "Paid", "Status"
        ];

        const data = [headers];

        rows.forEach(row => {
            const columns = row.querySelectorAll('td');
            if (columns.length === 10) {
                const rowData = [
                    extractCellContent(columns[0]),
                    extractCellContent(columns[1]),
                    extractCellContent(columns[2]),
                    extractCellContent(columns[3]),
                    extractCellContent(columns[4]),
                    extractCellContent(columns[5]),
                    formatCurrency(extractCellContent(columns[6])),
                    formatCurrency(extractCellContent(columns[7])),
                    formatCurrency(extractCellContent(columns[8])),
                    row.querySelector('.status-height-width span')?.textContent || ''
                ];
                data.push(rowData);
            }
        });

        return data;
    }

    // Download Excel file with ExcelJS
    async function downloadExcel(data, filename) {
        try {
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Billing Data');

            // Ensure data is an array
            if (!Array.isArray(data)) {
                console.error('Data is not an array:', data);
                return;
            }

            // Add headers and data to worksheet
            worksheet.addRows(data);

            // Get the table range
            const lastRow = worksheet.rowCount;
            const lastCol = worksheet.columnCount;
            const tableRef = `A1:${String.fromCharCode(64 + lastCol)}${lastRow}`;

            // Create a table with explicit rows
            worksheet.addTable({
                name: 'BillingTable',
                ref: tableRef,
                headerRow: true,
                totalsRow: false,
                style: {
                    theme: 'TableStyleMedium9',
                    showFirstColumn: false,
                    showLastColumn: false,
                    showRowStripes: true,
                    showColumnStripes: false
                },
                columns: [
                    { name: "Type", filterButton: true },
                    { name: "Client", filterButton: true },
                    { name: "Project/CO/Contract/Call", filterButton: true },
                    { name: "Payment Term", filterButton: true },
                    { name: "Billing Date", filterButton: true },
                    { name: "Due Date", filterButton: true },
                    { name: "Total Amount", filterButton: true },
                    { name: "Requested", filterButton: true },
                    { name: "Paid", filterButton: true },
                    { name: "Status", filterButton: true }
                ],
                rows: data.slice(1) // Add all rows except header row
            });

            // Set column widths
            const minWidths = [120, 150, 200, 120, 100, 100, 120, 100, 100, 100];
            worksheet.columns.forEach((column, i) => {
                if (i < minWidths.length) {
                    const maxLength = Math.max(...data.map(row => String(row[i] || '').length));
                    column.width = Math.max(minWidths[i] / 7, maxLength + 2);
                }
            });

            // Format currency columns
            for(let rowNum = 2; rowNum <= data.length; rowNum++) {
                ['G', 'H', 'I'].forEach(col => {
                    const cell = worksheet.getCell(`${col}${rowNum}`);
                    cell.numFmt = '$#,##0.00';
                    if(cell.value) {
                        cell.value = Number(cell.value) || 0;
                    }
                });
            }

            // Generate buffer and create download
            const buffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            });
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.download = filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);

        } catch (error) {
            console.error('Error generating Excel file:', error);
        }
    }

    // Function to create and add the download buttons
    function addDownloadButtons(table) {
        const csvButton = document.createElement('button');
        const excelButton = document.createElement('button');
        const buttonContainer = document.createElement('div');

        csvButton.textContent = 'Download CSV';
        excelButton.textContent = 'Download Excel';
        buttonContainer.style.cssText = 'display: flex; gap: 10px;';

        const buttonStyle = `
            margin: 0;
            padding: 8px 16px;
            background-color: #00bc81;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 500;
            transition: background-color 0.2s;
        `;

        csvButton.style.cssText = buttonStyle;
        excelButton.style.cssText = buttonStyle;

        csvButton.addEventListener('mouseover', () => csvButton.style.backgroundColor = '#00895e');
        csvButton.addEventListener('mouseout', () => csvButton.style.backgroundColor = '#00bc81');
        excelButton.addEventListener('mouseover', () => excelButton.style.backgroundColor = '#00895e');
        excelButton.addEventListener('mouseout', () => excelButton.style.backgroundColor = '#00bc81');

        csvButton.setAttribute('data-dtools-export', 'csv');
        excelButton.setAttribute('data-dtools-export', 'excel');

        csvButton.addEventListener('click', async () => {
            csvButton.disabled = true;
            csvButton.textContent = 'Gathering Data...';
            try {
                const now = new Date();
                const timestamp = now.toISOString().slice(0,10);
                const csvContent = await tableToCSV(table);
                downloadCSV(csvContent, `d-tools-billing-${timestamp}.csv`);
            } finally {
                csvButton.disabled = false;
                csvButton.textContent = 'Download CSV';
            }
        });

        excelButton.addEventListener('click', async () => {
            excelButton.disabled = true;
            excelButton.textContent = 'Gathering Data...';
            try {
                const now = new Date();
                const timestamp = now.toISOString().slice(0,10);
                const data = await tableToExcelData(table);
                await downloadExcel(data, `d-tools-billing-${timestamp}.xlsx`);
            } finally {
                excelButton.disabled = false;
                excelButton.textContent = 'Download Excel';
            }
        });

        buttonContainer.appendChild(csvButton);
        buttonContainer.appendChild(excelButton);

        // Find the specific flex-row container that's a sibling of the table's container
        const tableContainer = table.closest('.table-container');
        const summaryContainer = tableContainer?.closest('.summary-container');
        const flexRowContainer = summaryContainer?.querySelector('.flex-row.space-between-center');
        
        if (flexRowContainer && tableContainer) {
            // Only insert if this flex-row is before the table
            const flexRowRect = flexRowContainer.getBoundingClientRect();
            const tableRect = tableContainer.getBoundingClientRect();
            
            if (flexRowRect.top < tableRect.top) {
                // Find the "New payment request" button container
                const newPaymentButton = flexRowContainer.querySelector('button[mat-button]')?.closest('div');
                if (newPaymentButton) {
                    flexRowContainer.insertBefore(buttonContainer, newPaymentButton);
                    return;
                }
                flexRowContainer.appendChild(buttonContainer);
                return;
            }
        }

        // Fallback to original placement if conditions aren't met
        console.warn('Target flex-row container not found or not preceding table, falling back to table container');
        if (tableContainer) {
            tableContainer.insertBefore(buttonContainer, table);
        } else {
            table.parentElement.insertBefore(buttonContainer, table);
        }
    }

    // Function to wait for table to load
    function waitForTable() {
        const maxAttempts = 20;
        let attempts = 0;

        const checkTable = setInterval(() => {
            const table = document.querySelector('.table-container table');
            attempts++;

            if (table && table.querySelector('tbody tr')) {
                clearInterval(checkTable);
                // Check if buttons already exist to prevent duplicates
                const existingButtons = document.querySelector('button[data-dtools-export]');
                if (!existingButtons) {
                    addDownloadButtons(table);
                }
            } else if (attempts >= maxAttempts) {
                clearInterval(checkTable);
                console.log('Table not found after maximum attempts');
            }
        }, 500);
    }

    // Initialize when page loads and on route changes
    function initialize() {
        // Watch for URL changes (SPA navigation)
        let lastUrl = location.href;
        new MutationObserver(() => {
            const url = location.href;
            if (url !== lastUrl) {
                lastUrl = url;
                if (url.includes('/billing/home')) {
                    waitForTable();
                }
            }
        }).observe(document, { subtree: true, childList: true });

        // Initial load
        if (location.href.includes('/billing/home')) {
            waitForTable();
        }
    }

    // Start initialization
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initialize);
    } else {
        initialize();
    }

    // Load ExcelJS instead of SheetJS
    loadExcelJS(() => console.log("ExcelJS loaded for Excel export"));
})();

