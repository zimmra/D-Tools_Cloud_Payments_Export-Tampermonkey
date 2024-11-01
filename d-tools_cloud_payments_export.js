// ==UserScript==
// @name         D-Tools Cloud Billing Table to CSV & Excel Downloader
// @namespace    D-Tools
// @version      2.2
// @description  Add download CSV and Excel buttons for D-Tools Cloud billing table
// @author       Payton Zimmerer
// @match        https://d-tools.cloud/billing/home
// @grant        none
// @downloadURL  https://raw.githubusercontent.com/zimmra/D-Tools_Cloud_Payments_Export-Tampermonkey/refs/heads/main/d-tools_cloud_payments_export.js
// ==/UserScript==

(function() {
    'use strict';

    // Load SheetJS library for Excel export
    function loadSheetJS(callback) {
        const script = document.createElement('script');
        script.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.16.9/xlsx.full.min.js";
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

    // Function to convert table data to CSV string
    function tableToCSV(table) {
        const rows = table.querySelectorAll('tbody tr');
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

    // Function to calculate string width in Excel units
    function getStringWidth(str) {
        // Rough approximation: each character is about 1 unit wide
        // Numbers and special chars are slightly narrower, CJK chars are wider
        return str.length * 1.2;
    }

    // Convert table data to Excel format based on provided Excel file's styles
    function tableToExcelData(table) {
        const rows = table.querySelectorAll('tbody tr');
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

    // Download Excel file with SheetJS
    function downloadExcel(data, filename) {
        const ws = XLSX.utils.aoa_to_sheet(data);
        const wb = XLSX.utils.book_new();
        
        // Convert to table and set table styling
        const range = XLSX.utils.decode_range(ws['!ref']);
        const tableRef = {
            s: { r: 0, c: 0 },
            e: { r: range.e.r, c: range.e.c }
        };
        
        // Add table definition
        ws['!table'] = {
            ref: XLSX.utils.encode_range(tableRef),
            name: 'BillingTable',
            style: {
                theme: 'TableStyleMedium9',
                showFirstColumn: false,
                showLastColumn: false,
                showRowStripes: true,
                showColumnStripes: false
            }
        };

        // Minimum column widths (in Excel units)
        const minWidths = [
            { wpx: 120 }, { wpx: 150 }, { wpx: 200 }, { wpx: 120 }, 
            { wpx: 100 }, { wpx: 100 }, { wpx: 120 }, { wpx: 100 },
            { wpx: 100 }, { wpx: 100 }
        ];

        // Calculate maximum content width for each column
        const maxWidths = new Array(10).fill(0);
        
        // Process all rows including header
        data.forEach(row => {
            row.forEach((cell, colIndex) => {
                const cellWidth = getStringWidth(String(cell));
                maxWidths[colIndex] = Math.max(maxWidths[colIndex], cellWidth);
            });
        });

        // Apply the larger of minimum width or content width
        ws['!cols'] = minWidths.map((col, index) => {
            // Convert minimum width from pixels to Excel units (approximately)
            const minExcelUnits = col.wpx / 7;
            // Add some padding to the content width
            const contentWidth = maxWidths[index] + 2;
            
            return {
                // Use the larger of the two widths
                width: Math.max(minExcelUnits, contentWidth),
                // Keep the pixel width as a minimum
                wpx: col.wpx
            };
        });

        // Apply cell styles
        for(let R = tableRef.s.r; R <= tableRef.e.r; R++) {
            for(let C = tableRef.s.c; C <= tableRef.e.c; C++) {
                const cell_address = XLSX.utils.encode_cell({r: R, c: C});
                if(!ws[cell_address]) continue;
                
                // Initialize cell style if not exists
                if(!ws[cell_address].s) ws[cell_address].s = {};
                
                // Header row styling
                if(R === 0) {
                    ws[cell_address].s = {
                        fill: {fgColor: {rgb: "4472C4"}},
                        font: {color: {rgb: "FFFFFF"}, bold: true},
                        alignment: {horizontal: "center", vertical: "center", wrapText: true}
                    };
                }
                // Data rows styling
                else {
                    ws[cell_address].s = {
                        font: {color: {rgb: "000000"}},
                        alignment: {horizontal: "left", vertical: "center", wrapText: true}
                    };
                    
                    // Apply number format to amount columns (6,7,8)
                    if(C >= 6 && C <= 8) {
                        ws[cell_address].z = '$#,##0.00';
                        ws[cell_address].s.alignment.horizontal = "right";
                    }
                }
            }
        }

        XLSX.utils.book_append_sheet(wb, ws, "Billing Data");
        XLSX.writeFile(wb, filename);
    }

    // Function to create and add the download buttons
    function addDownloadButtons(table) {
        const csvButton = document.createElement('button');
        const excelButton = document.createElement('button');
        
        csvButton.textContent = 'Download CSV';
        excelButton.textContent = 'Download Excel';

        const buttonStyle = `
            margin: 10px;
            padding: 8px 16px;
            background-color: #0072bc;
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
        
        csvButton.addEventListener('mouseover', () => csvButton.style.backgroundColor = '#005a96');
        csvButton.addEventListener('mouseout', () => csvButton.style.backgroundColor = '#0072bc');
        excelButton.addEventListener('mouseover', () => excelButton.style.backgroundColor = '#005a96');
        excelButton.addEventListener('mouseout', () => excelButton.style.backgroundColor = '#0072bc');
        
        csvButton.addEventListener('click', () => {
            const now = new Date();
            const timestamp = now.toISOString().slice(0,10);
            const csvContent = tableToCSV(table);
            downloadCSV(csvContent, `d-tools-billing-${timestamp}.csv`);
        });

        excelButton.addEventListener('click', () => {
            const now = new Date();
            const timestamp = now.toISOString().slice(0,10);
            const data = tableToExcelData(table);
            downloadExcel(data, `d-tools-billing-${timestamp}.xlsx`);
        });

        const tableContainer = table.closest('.table-container');
        if (tableContainer) {
            tableContainer.insertBefore(csvButton, table);
            tableContainer.insertBefore(excelButton, table);
        } else {
            table.parentElement.insertBefore(csvButton, table);
            table.parentElement.insertBefore(excelButton, table);
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
                addDownloadButtons(table);
            } else if (attempts >= maxAttempts) {
                clearInterval(checkTable);
                console.log('Table not found after maximum attempts');
            }
        }, 500);
    }

    // Initialize when page loads
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', waitForTable);
    } else {
        waitForTable();
    }

    // Load SheetJS for Excel functionality
    loadSheetJS(() => console.log("SheetJS loaded for Excel export"));
})();

