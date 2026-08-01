// ===============================
// 35 ST BN Personnel Management
// Part-1 : Excel Import
// ===============================

let personnel = [];

const excelInput = document.getElementById("excelFile");
const searchInput = document.getElementById("search");
const tableBody = document.getElementById("tableBody");

excelInput.addEventListener("change", loadExcel);

function loadExcel(e){

    const file = e.target.files[0];

    if(!file){
        alert("Please select an Excel file.");
        return;
    }

    const reader = new FileReader();

    reader.onload = function(evt){

        const data = new Uint8Array(evt.target.result);

        const workbook = XLSX.read(data,{
            type:"array"
        });

        // আপনার Collection Sheet
        const sheetName = "Collection";

        const sheet = workbook.Sheets[sheetName];

        if(!sheet){
            alert("Collection Sheet Not Found");
            return;
        }

        personnel = XLSX.utils.sheet_to_json(sheet,{
            defval:""
        });

        console.log(personnel);

        renderTable(personnel);

        updateDashboard();

    }

    reader.readAsArrayBuffer(file);

}

function renderTable(data){

    tableBody.innerHTML="";

    data.forEach(person=>{

        tableBody.innerHTML+=`

<tr>

<td>${person["ID No"]}</td>

<td>${person["RANK"]}</td>

<td>${person["NAME"]}</td>

<td>${person["TRADE"]}</td>

<td>${person["COY"]}</td>

<td>${person["DISTRICT"]}</td>

</tr>

`;

    });

}

function updateDashboard(){

    document.getElementById("total").innerHTML =
    personnel.length;

}