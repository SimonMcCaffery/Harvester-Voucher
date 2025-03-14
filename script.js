

function processFile(){
    const file = document.getElementById("file-input").files[0];
    if (file){
        const reader = new FileReader();
        reader.onload = function(e){
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});

            //replace with index of sheet that contains the information needed on voucher
            const sheetName = workbook.SheetNames[2];
            const sheet = workbook.Sheets[sheetName];

            //convert to JSON
            const jsonData = XLSX.utils.sheet_to_json(sheet)

            generateVouchers(jsonData);
        };
        reader.readAsArrayBuffer(file);
    }else{
        alert("please select a file");
    }
}

function generateVouchers(data){
    const container = document.getElementById("vouchers-container");
    container.innerHTML = '';
}