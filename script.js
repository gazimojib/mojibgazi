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

        updateDashboard()updateSection(
    "HQ BN",
    "hqTotal",
    "hqCollected",
    "hqRemaining",
    "hqBar"
);

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
// ===============================
// Part-2 : Dashboard + Search
// ===============================

// Search
searchInput.addEventListener("keyup", function () {

    const keyword = this.value.toLowerCase();

    const filtered = personnel.filter(person => {

        return (
            String(person["ID No"]).toLowerCase().includes(keyword) ||
            String(person["NAME"]).toLowerCase().includes(keyword) ||
            String(person["RANK"]).toLowerCase().includes(keyword) ||
            String(person["TRADE"]).toLowerCase().includes(keyword) ||
            String(person["COY"]).toLowerCase().includes(keyword) ||
            String(person["DISTRICT"]).toLowerCase().includes(keyword)
        );

    });

    renderTable(filtered);

});

// Dashboard
function updateDashboard(){

    const total = personnel.length;

    const collected = personnel.filter(p =>
        String(p["Collected?"]).trim() === "ডাটা পাওয়া গেছে"
    ).length;

    const remaining = total - collected;

    document.getElementById("total").innerText = total;
    document.getElementById("collected").innerText = collected;
    document.getElementById("remaining").innerText = remaining;

    console.log("Total:", total);
    console.log("Collected:", collected);
    console.log("Remaining:", remaining);

}
// Section Count

document.getElementById("hqCount").innerText =
personnel.filter(x=>x["COY"]=="HQ BN").length;

document.getElementById("c59Count").innerText =
personnel.filter(x=>x["COY"]=="59 COY").length;

document.getElementById("c60Count").innerText =
personnel.filter(x=>x["COY"]=="60 COY").length;

document.getElementById("c61Count").innerText =
personnel.filter(x=>x["COY"]=="61 COY").length;

document.getElementById("mtplCount").innerText =
personnel.filter(x=>x["COY"]=="MTPL").length;

document.getElementById("emeCount").innerText =
personnel.filter(x=>x["COY"]=="EME").length;