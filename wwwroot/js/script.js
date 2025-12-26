// ===== Elements =====
const fileInput = document.getElementById("fileInput");
const uploadBtn = document.getElementById("uploadBtn");
const feedback = document.getElementById("feedback");
const tableBody = document.querySelector("#resultTable tbody");
const uploadArea = document.getElementById("uploadArea");
const uploadProgress = document.getElementById("uploadProgress");

const duplicatesCard = document.getElementById("duplicatesCard");
const duplicatesTableBody = document.querySelector("#duplicatesTable tbody");
const downloadDuplicatesBtn = document.getElementById("downloadDuplicatesBtn");

let duplicateRows = []; // store duplicates for download

// ===== Enable upload button =====
fileInput.addEventListener("change", () => {
    uploadBtn.disabled = !fileInput.files.length;
});

// ===== Clickable Upload Area =====
uploadArea.addEventListener("click", () => fileInput.click());

// ===== Drag & Drop =====
uploadArea.addEventListener("dragover", e => { e.preventDefault(); uploadArea.classList.add("drag-over"); });
uploadArea.addEventListener("dragleave", () => { uploadArea.classList.remove("drag-over"); });
uploadArea.addEventListener("drop", e => {
    e.preventDefault();
    uploadArea.classList.remove("drag-over");
    if (e.dataTransfer.files.length) {
        fileInput.files = e.dataTransfer.files;
        uploadBtn.disabled = false;
    }
});

// ===== Upload Handler =====
uploadBtn.addEventListener("click", async () => {
    if (!fileInput.files.length) return;

    const formData = new FormData();
    formData.append("file", fileInput.files[0]);

    // Reset UI
    tableBody.innerHTML = "";
    duplicatesTableBody.innerHTML = "";
    duplicateRows = [];
    duplicatesCard.style.display = "none";
    showAlert("Uploading...", "info");
    uploadProgress.style.width = "0%";

    try {
        const res = await fetch("/api/import/excel", { method: "POST", body: formData });
        const text = await res.text();

        if (!res.ok) return showAlert("Upload failed: " + text, "error");

        let data;
        try { data = JSON.parse(text); } catch { return showAlert("Server returned invalid JSON", "error"); }

        console.log("Server response:", data); // 🔹 debug

        // ✅ Support PascalCase and camelCase
        const previewData = data.Preview || data.preview || [];
        const importedRows = data.ImportedRows ?? data.importedRows ?? 0;
        const duplicateCount = data.DuplicateRows ?? data.duplicateRows ?? 0;

        if (!previewData.length) return showAlert("No preview data returned from server", "error");

        // Success alert
        showAlert(`Imported: ${importedRows}, Duplicates skipped: ${duplicateCount}`, "success");

        // Populate preview table
        previewData.forEach(r => {
            const tr = document.createElement("tr");
            let statusClass = "";
            switch (r.Status || r.status) {
                case "DUPLICATE": statusClass = "status-duplicate"; duplicateRows.push(r); break;
                case "PASS": statusClass = "status-pass"; break;
                case "FAIL": statusClass = "status-fail"; break;
            }

            tr.innerHTML = `
                <td>${r.Task ?? r.task}</td>
                <td>${r.TestPart ?? r.testPart}</td>
                <td>${new Date(r.TestDate ?? r.testDate).toLocaleString()}</td>
                <td>${r.Value ?? r.value}</td>
                <td class="${statusClass}">${r.Status ?? r.status}</td>
            `;
            tableBody.appendChild(tr);
        });

        // Show duplicates table if any
        if (duplicateRows.length) {
            duplicatesCard.style.display = "block";
            duplicateRows.forEach(r => {
                const tr = document.createElement("tr");
                tr.innerHTML = `
                    <td>${r.Task ?? r.task}</td>
                    <td>${r.TestPart ?? r.testPart}</td>
                    <td>${new Date(r.TestDate ?? r.testDate).toLocaleString()}</td>
                    <td>${r.Value ?? r.value}</td>
                `;
                duplicatesTableBody.appendChild(tr);
            });
        }

        // Animate progress bar
        uploadProgress.style.width = "100%";

    } catch (err) {
        showAlert("Unexpected error: " + err.message, "error");
        console.error(err);
    }
});

// ===== Download Duplicates as CSV =====
downloadDuplicatesBtn.addEventListener("click", () => {
    if (!duplicateRows.length) return;

    const headers = ["Task", "Test Part", "Test Date", "Value"];
    const csvRows = [headers.join(",")];

    duplicateRows.forEach(r => {
        const row = [
            r.Task ?? r.task,
            r.TestPart ?? r.testPart,
            new Date(r.TestDate ?? r.testDate).toISOString(),
            r.Value ?? r.value
        ];
        csvRows.push(row.join(","));
    });

    const csvContent = "data:text/csv;charset=utf-8," + csvRows.join("\n");
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", "duplicates.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
});

// ===== Alert helper =====
function showAlert(message, type = "info") {
    feedback.innerHTML = `<div class="alert ${type}">${message}</div>`;
    const alertDiv = feedback.querySelector(".alert");
    setTimeout(() => { if (alertDiv) alertDiv.style.opacity = "0"; }, 4000);
    setTimeout(() => { if (alertDiv) alertDiv.remove(); }, 4500);
}
