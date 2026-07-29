/*
=========================================
Excel AI
excel.js
Version 1.0
=========================================
*/

(function () {

    function tableToCSV(tableId) {

        const table = document.getElementById(tableId);

        if (!table) {
            alert("Table not found.");
            return "";
        }

        let csv = [];

        const rows = table.querySelectorAll("tr");

        rows.forEach(row => {

            let cols = row.querySelectorAll("th, td");

            let rowData = [];

            cols.forEach(col => {

                let value = col.innerText
                    .replace(/"/g, '""')
                    .replace(/\n/g, " ");

                rowData.push(`"${value}"`);

            });

            csv.push(rowData.join(","));

        });

        return csv.join("\n");

    }

    function downloadCSV(filename, csv) {

        const blob = new Blob([csv], {
            type: "text/csv;charset=utf-8;"
        });

        const link = document.createElement("a");

        const url = URL.createObjectURL(blob);

        link.href = url;

        link.download = filename;

        document.body.appendChild(link);

        link.click();

        document.body.removeChild(link);

        URL.revokeObjectURL(url);

    }

    function downloadExcel() {

        const csv = tableToCSV("dataTable");

        if (!csv) return;

        const fileName =
            "ExcelAI_" +
            new Date().toISOString().slice(0, 10) +
            ".csv";

        downloadCSV(fileName, csv);

    }

    // Attach button event after page loads
    window.addEventListener("DOMContentLoaded", () => {

        const btn = document.getElementById("downloadExcel");

        if (btn) {

            btn.addEventListener("click", downloadExcel);

        }

    });

    // Make function available globally
    window.downloadExcel = downloadExcel;

    console.log("✅ excel.js Loaded");

})();
