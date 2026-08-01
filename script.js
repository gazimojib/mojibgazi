// Global data
let personnelData = [];

// Load Excel
document.getElementById("excelFile").addEventListener("change", function (e) {

    const file = e.target.files[0];

    const reader = new FileReader();

    reader.onload = function (event) {

        const data = new Uint8Array(event.target.result);

        const workbook = XLSX.read(data, { type: "array" });

        const sheetName = workbook.SheetNames[0];

        const sheet = workbook.Sheets[sheetName];

        personnelData = XLSX.utils.sheet_to_json(sheet);

        loadTable(personnelData);

        updateDashboard();

    };

    reader.readAsArrayBuffer(file);

});

function loadTable(data){

    const table=document.getElementById("tableBody");

    table.innerHTML="";

    data.forEach(person=>{

        table.innerHTML+=`

<tr>

<td>${person["ID"]||""}</td>

<td>${person["Rank"]||""}</td>

<td>${person["Name"]||""}</td>

<td>${person["Trade"]||""}</td>

<td>${person["Section"]||""}</td>

<td>${person["Collected"]||""}</td>

</tr>

`;

    });

}

function updateDashboard(){

document.getElementById("total").innerText=personnelData.length;

let collected=personnelData.filter(x=>x["Collected"]==="ডাটা পাওয়া গেছে").length;

document.getElementById("collected").innerText=collected;

document.getElementById("remaining").innerText=personnelData.length-collected;

}

// Search

document.getElementById("search").addEventListener("keyup",function(){

let value=this.value.toLowerCase();

let result=personnelData.filter(person=>{

return Object.values(person).join(" ").toLowerCase().includes(value);

});

loadTable(result);

});