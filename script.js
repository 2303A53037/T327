let originalData = [];

document.getElementById("excelFile").addEventListener("change", function(e) {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet);

        originalData = jsonData;
        renderTable(jsonData);
    };

    reader.readAsArrayBuffer(file);
});

function renderTable(data) {
    if (data.length === 0) {
        document.getElementById("tableContainer").innerHTML = "<p>No data available.</p>";
        return;
    }

    let html = `<table><thead><tr>`;
    const keys = Object.keys(data[0]);

    keys.forEach(key => html += `<th>${key}</th>`);
    html += `</tr></thead><tbody>`;

    data.forEach(row => {
        html += `<tr>`;
        keys.forEach(key => html += `<td>${row[key] || '-'}</td>`);
        html += `</tr>`;
    });

    html += `</tbody></table>`;
    document.getElementById("tableContainer").innerHTML = html;
}

function filterByYear() {
    const from = parseInt(document.getElementById("fromYear").value);
    const to = parseInt(document.getElementById("toYear").value);

    if (!from || !to || from > to) return alert("Please enter valid years.");

    const filtered = originalData.filter(record => {
        const year = parseInt(record.Year || record.year);
        return year >= from && year <= to;
    });

    renderTable(filtered);
}

function exportToExcel() {
    const worksheet = XLSX.utils.json_to_sheet(originalData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Publications");
    XLSX.writeFile(workbook, "publication_summary.xlsx");
}

function exportToWord() {
    const blob = new Blob([document.getElementById("tableContainer").innerHTML], {
        type: "application/msword"
    });
    saveAs(blob, "publication_summary.doc");
}
