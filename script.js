let date = new Date().toLocaleDateString();

function processFile() {
    const file = document.getElementById("file-input").files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = function (e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });

            // Replace with the index of the sheet that contains the information needed on vouchers
            const sheetName = workbook.SheetNames[2];
            const sheet = workbook.Sheets[sheetName];

            // Convert sheet to JSON
            const jsonData = XLSX.utils.sheet_to_json(sheet);

            generateVouchers(jsonData);
        };
        reader.readAsArrayBuffer(file);
    } else {
        alert("Please select a file");
    }
}

function generateVouchers(data) {
    const container = document.getElementById("vouchers-container");
    container.innerHTML = ""; // Clear previous vouchers

    data.forEach((row, index) => {
        // Ensure row contains necessary fields
        if (!row.Name || !row.Amount || !row.RoomNumber) {
            console.error(`Missing data in row ${index + 1}`, row);
            return; // Skip this row if data is missing
        }

        // Create a new div for each voucher
        const voucherDiv = document.createElement("div");
        voucherDiv.classList.add("voucher");

        // Create voucher content
        voucherDiv.innerHTML = `
            <div class="sideTab">
                <p class="boxTab">${date}</p>
                <p class="boxTab" id="nameTab">${row.Name}</p>
                <p class="boxTab">${row.RoomNumber}</p>
                <p class="boxTab">${row.Amount}.00</p>
            </div>
            <div class="voucherMain">
                <p id="mainAmount" class="height1">${row.Amount}</p>
                <p id="harvester" class="height1">Harvester</p>
                <p id="mainName" class="height1">${row.Name}</p>
                <p id="mainDate" class="height1">${date}</p>            
        `;

        // Append the voucher to the container
        container.appendChild(voucherDiv);


    });
    window.print();
}
