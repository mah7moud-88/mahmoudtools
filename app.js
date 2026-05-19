Office.onReady(() => {

const uploadBtn = document.getElementById("uploadBtn");
const fileInput = document.getElementById("fileInput");
const dataBox = document.getElementById("dataBox");
const repeatCountEl = document.getElementById("repeatCount");
const status = document.getElementById("status");
const progressBar = document.getElementById("progressBar");

const checkBtn = document.getElementById("checkBtn");
const reportBtn = document.getElementById("reportBtn");

const loadedCount = document.getElementById("loadedCount");
const totalDash = document.getElementById("totalDash");
const progressDash = document.getElementById("progressDash");
const timeDash = document.getElementById("timeDash");

let stopRequested = false;
let startTime = 0;
let timerInterval = null;
let lastSourceValues = [];

/* TIMER */
function startTimer() {
startTime = Date.now();
timerInterval = setInterval(() => {
timeDash.innerText = Math.floor((Date.now() - startTime)/1000) + "s";
}, 500);
}

function stopTimer() {
clearInterval(timerInterval);
}

/* LIVE COUNT */
function updateLiveCount() {
const values = (dataBox.value || "")
.split("\n")
.map(v => v.trim())
.filter(v => v !== "");

loadedCount.innerText = values.length;
status.innerText = "Ready | " + values.length + " items";
}

dataBox.addEventListener("input", updateLiveCount);

/* UPLOAD */
uploadBtn.onclick = () => fileInput.click();

fileInput.onchange = (e) => {
const file = e.target.files[0];
if (!file) return;

const reader = new FileReader();

reader.onload = (event) => {

const data = new Uint8Array(event.target.result);
const workbook = XLSX.read(data, { type: "array" });

const sheet = workbook.Sheets[workbook.SheetNames[0]];

const json = XLSX.utils.sheet_to_json(sheet, {
header: 1,
defval: ""
});

const cleaned = (json || [])
.flat()
.map(v => String(v || "").trim())
.filter(v => v !== "");

dataBox.value = cleaned.join("\n");
loadedCount.innerText = cleaned.length;

status.innerText = "Loaded ✅";

};

reader.readAsArrayBuffer(file);
};

/* CLEAR */
document.getElementById("clearBtn").onclick = () => {

dataBox.value = "";
progressBar.style.width = "0%";

loadedCount.innerText = "0";
totalDash.innerText = "0";
progressDash.innerText = "0%";
timeDash.innerText = "0s";

status.innerText = "Cleared 🧹";

};

/* STOP */
document.getElementById("stopBtn").onclick = () => {
stopRequested = true;
status.innerText = "Stopped ⛔";
stopTimer();
};

/* PASTE */
document.getElementById("run").onclick = async () => {

const repeatTimes = parseInt(repeatCountEl.value || "0");

const baseValues = (dataBox.value || "")
.split("\n")
.map(v => v.trim())
.filter(v => v !== "");

let values = [];

if (repeatTimes <= 0) values = baseValues;
else {
for (const v of baseValues) {
for (let i = 0; i < repeatTimes; i++) {
values.push(v);
}
}
}

if (!values.length) return alert("اكتب أو ارفع بيانات");

lastSourceValues = [...values];

stopRequested = false;
progressBar.style.width = "0%";

status.innerText = "Processing ⚡";

startTimer();

try {

await Excel.run(async (context) => {

const sheet = context.workbook.worksheets.getActiveWorksheet();

const selected = context.workbook.getSelectedRange();
selected.load("rowIndex,columnIndex");

await context.sync();

const startRow = selected.rowIndex;
const col = selected.columnIndex;

const range = sheet.getRangeByIndexes(startRow, col, 100000, 1);

const visible = range.getSpecialCells(Excel.SpecialCellType.visible);

visible.load("areas/items/rowCount");

await context.sync();

let i = 0;

for (const area of visible.areas.items) {

for (let r = 0; r < area.rowCount; r++) {

if (stopRequested) break;
if (i >= values.length) break;

area.getCell(r,0).values = [[values[i++]]];

let percent = Math.round((i / values.length) * 100);
progressBar.style.width = percent + "%";

}

if (stopRequested) break;
}

await context.sync();

progressBar.style.width = "100%";
status.innerText = "Done 🚀";
stopTimer();

});

} catch (err) {
status.innerText = "Error ❌ " + err.message;
stopTimer();
}

};

/* CHECK */
checkBtn.onclick = async () => {

await Excel.run(async (context) => {

const range = context.workbook.getSelectedRange();
range.load("values");
await context.sync();

const excel = (range.values || []).flat().filter(v => v !== "");
const box = (dataBox.value || "").split("\n").filter(v => v.trim() !== "");

status.innerText = `📊 Excel: ${excel.length} | Box: ${box.length}`;

});

};

/* REPORT */
reportBtn.onclick = async () => {

await Excel.run(async (context) => {

const range = context.workbook.getSelectedRange();
range.load("values,rowCount,columnCount");

await context.sync();

if (range.columnCount < 2)
return status.innerText = "Select 2 columns";

let data = range.values;
let report = [["Account","Value"]];

for (let i=0;i<range.rowCount;i++){
report.push([data[i][0], data[i][1]]);
}

const wb = XLSX.utils.book_new();
const ws = XLSX.utils.aoa_to_sheet(report);

XLSX.utils.book_append_sheet(wb, ws, "Report");

XLSX.writeFile(wb, "report.xlsx");

status.innerText = "Report downloaded ✅";

});

};

});