// Excel data
let personnelData = [];

// File Upload
document.getElementById("excelFile").addEventListener("change", function(e){

    const file = e.target.files[0];

    if(!file){
        return;
    }

    const reader = new FileReader();

    reader.onload = function(event){

        const data = new Uint8Array(event.target.result);

        const workbook = XLSX.read(data,{type:"array"});

        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        personnelData = XLSX.utils.sheet_to_json(sheet);

        showTable(personnelData);

        updateDashboard();

    };

    reader.readAsArrayBuffer(file);

});

// Show Table
function showTable(data){

    let html="";

    data.forEach(function(item){

        html += `
        <tr>
            <td>${item.ID || ""}</td>
            <td>${item.Rank || ""}</td>
            <td>${item.Name || ""}</td>
            <td>${item.Trade || ""}</td>
            <td>${item.Section || ""}</td>
            <td>${item.Status || ""}</td>
        </tr>
        `;

    });

    document.getElementById("tableBody").innerHTML = html;

}

// Dashboard
function updateDashboard(){

    document.getElementById("total").innerHTML = personnelData.length;

}

// Search
document.getElementById("search").addEventListener("keyup",function(){

    let value = this.value.toLowerCase();

    let result = personnelData.filter(function(item){

        return JSON.stringify(item).toLowerCase().includes(value);

    });

    showTable(result);

});