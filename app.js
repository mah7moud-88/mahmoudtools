Office.onReady(() => {

const dataBox = document.getElementById("dataBox");
const repeatCountEl = document.getElementById("repeatCount");
const status = document.getElementById("status");
const progressBar = document.getElementById("progressBar");

let lastSourceValues = [];

/* STATUS */
function setStatus(t){ status.innerText = t; }

/* UPLOAD (safe) */
const uploadBtn = document.getElementById("uploadBtn");
const fileInput = document.createElement("input");
fileInput.type = "file";

uploadBtn.onclick = () => fileInput.click();

fileInput.onchange = e => {
const file = e.target.files[0];
if(!file) return;

const reader = new FileReader();

reader.onload = ev => {
const data = new Uint8Array(ev.target.result);
const wb = XLSX.read(data, {type:"array"});
const sheet = wb.Sheets[wb.SheetNames[0]];
const json = XLSX.utils.sheet_to_json(sheet,{header:1});

const cleaned = json.flat().map(v=>String(v).trim()).filter(v=>v);
dataBox.value = cleaned.join("\n");

setStatus("Loaded");
};

reader.readAsArrayBuffer(file);
};

/* CLEAR */
document.getElementById("clearBtn").onclick = () => {
dataBox.value="";
progressBar.style.width="0%";
setStatus("Cleared");
};

/* STOP */
let stop=false;
document.getElementById("stopBtn").onclick = () => {
stop=true;
setStatus("Stopped");
};

/* PASTE */
document.getElementById("run").onclick = async () => {

let values = dataBox.value.split("\n").filter(v=>v.trim());

lastSourceValues=[...values];
stop=false;

try{

await Excel.run(async context=>{

const sheet=context.workbook.worksheets.getActiveWorksheet();
const range=context.workbook.getSelectedRange();
range.load("rowIndex,columnIndex");

await context.sync();

const startRow=range.rowIndex;
const col=range.columnIndex;

const target=sheet.getRangeByIndexes(startRow,col,10000,1);

let i=0;

for(let r=0;r<10000;r++){

if(stop) break;
if(i>=values.length) break;

const cell=target.getCell(r,0);
cell.values=[[values[i++]]];

progressBar.style.width=Math.round((i/values.length)*100)+"%";
}

await context.sync();
setStatus("Done");

});

}catch(e){
setStatus("Error "+e.message);
}

};

/* CHECK */
document.getElementById("checkBtn").onclick = async () => {

await Excel.run(async context=>{

const r=context.workbook.getSelectedRange();
r.load("values");
await context.sync();

const count=r.values.flat().filter(v=>String(v).trim()).length;

setStatus("Pasted: "+count+" | Source: "+lastSourceValues.length);

});

};

/* SEARCH SAFE */
const searchBtn=document.getElementById("searchBtn");
const searchBox=document.getElementById("searchBox");
const searchInput=document.getElementById("searchInput");
const doSearchBtn=document.getElementById("doSearchBtn");

/* check existence before use */
if(searchBtn && searchBox){

searchBtn.onclick=()=>{
searchBox.style.display =
searchBox.style.display==="block"?"none":"block";
};

}

if(doSearchBtn){

doSearchBtn.onclick=()=>{
console.log("Search:",searchInput.value);
setStatus("Searching: "+searchInput.value);
};

}

});