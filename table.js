/*
=========================================
 Excel AI
 table.js
 Version : 2.0
=========================================
*/

(function () {

    /**
     * Render Dataset Table
     * @param {string} datasetName
     */
    function renderTable(datasetName) {

        const container = document.getElementById("tableContainer");

        if (!container) {
            console.error("tableContainer not found.");
            return;
        }

        // ******** IMPORTANT ********
        // Remove everything before creating new table
        container.replaceChildren();

        const data = getSampleData(datasetName);

        if (!data || data.length === 0) {

            container.innerHTML = `
                <div style="padding:20px;text-align:center;color:#888;">
                    No sample data found.
                </div>
            `;

            return;
        }

        const table = document.createElement("table");
        table.id = "dataTable";
        table.className = "excel-table";

        //----------------------------------
        // Header
        //----------------------------------

        const thead = document.createElement("thead");
        const headerRow = document.createElement("tr");

        Object.keys(data[0]).forEach(header => {

            const th = document.createElement("th");
            th.textContent = header;

            headerRow.appendChild(th);

        });

        thead.appendChild(headerRow);

        table.appendChild(thead);

        //----------------------------------
        // Body
        //----------------------------------

        const tbody = document.createElement("tbody");

        data.forEach(item => {

            const tr = document.createElement("tr");

            Object.values(item).forEach(value => {

                const td = document.createElement("td");

                td.textContent = value;

                tr.appendChild(td);

            });

            tbody.appendChild(tr);

        });

        table.appendChild(tbody);

        container.appendChild(table);

    }

    /**
     * Clear Table
     */
    function clearTable() {

        const container = document.getElementById("tableContainer");

        if (container) {

            container.replaceChildren();

        }

    }

    /**
     * Get Current Table Data
     */
    function getCurrentTableData() {

        const table = document.getElementById("dataTable");

        if (!table) return [];

        const headers = [];

        table.querySelectorAll("thead th").forEach(th => {

            headers.push(th.innerText);

        });

        const rows = [];

        table.querySelectorAll("tbody tr").forEach(tr => {

            const obj = {};

            tr.querySelectorAll("td").forEach((td, index) => {

                obj[headers[index]] = td.innerText;

            });

            rows.push(obj);

        });

        return rows;

    }

    // Global Functions
    window.renderTable = renderTable;
    window.clearTable = clearTable;
    window.getCurrentTableData = getCurrentTableData;

    console.log("✅ table.js Loaded");

})();