// ==UserScript==
// @name         D-Tools Cloud Billing Table to CSV Downloader
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  Add download CSV button for D-Tools Cloud billing table
// @author       Your name
// @match        https://d-tools.cloud/billing/home
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // Function to clean text content
    function cleanText(text) {
        return text.replace(/\s+/g, ' ').trim();
    }

    // Function to extract cell content safely
    function extractCellContent(cell) {
        // Handle different cell content structures
        if (!cell) return '';
        
        // For cells with nested elements (like the Type column)
        const flexColumn = cell.querySelector('.flex-column');
        if (flexColumn) {
            return cleanText(flexColumn.textContent);
        }
        
        // For cells with links
        const link = cell.querySelector('a');
        if (link) {
            return cleanText(link.textContent);
        }
        
        // For regular cells
        return cleanText(cell.textContent);
    }

    // Function to format currency
    function formatCurrency(text) {
        // Remove $ sign and handle empty values
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
                    extractCellContent(columns[0]), // Type
                    extractCellContent(columns[1]), // Client
                    extractCellContent(columns[2]), // Project/CO/Contract/Call
                    extractCellContent(columns[3]), // Payment Term
                    extractCellContent(columns[4]), // Billing Date
                    extractCellContent(columns[5]), // Due Date
                    formatCurrency(extractCellContent(columns[6])), // Total Amount
                    formatCurrency(extractCellContent(columns[7])), // Requested
                    formatCurrency(extractCellContent(columns[8])), // Paid
                    row.querySelector('.status-height-width span')?.textContent || '' // Status
                ].map(text => {
                    // Escape special characters and wrap in quotes if necessary
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

    // Function to create and add the download button
    function addDownloadButton(table) {
        const button = document.createElement('button');
        button.textContent = 'Download CSV';
        button.style.cssText = `
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
        
        button.addEventListener('mouseover', () => {
            button.style.backgroundColor = '#005a96';
        });
        
        button.addEventListener('mouseout', () => {
            button.style.backgroundColor = '#0072bc';
        });
        
        button.addEventListener('click', () => {
            const now = new Date();
            const timestamp = now.toISOString().slice(0,10);
            const csvContent = tableToCSV(table);
            downloadCSV(csvContent, `d-tools-billing-${timestamp}.csv`);
        });

        // Insert button before the table container
        const tableContainer = table.closest('.table-container');
        if (tableContainer) {
            tableContainer.insertBefore(button, table);
        } else {
            table.parentElement.insertBefore(button, table);
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
                addDownloadButton(table);
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
})();
