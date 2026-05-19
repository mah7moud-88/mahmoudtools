Office.onReady(() => {

const uploadBtn = document.getElementById("uploadBtn");
const fileInput = document.getElementById("fileInput");
const dataBox = document.getElementById("dataBox");
const repeatCountEl = document.getElementById("repeatCount");
const status = document.getElementById("status");
const progressBar = document.getElementById("progressBar");

const checkBtn = document.getElementById("checkBtn");
const reportBtn = document.getElementById("reportBtn");

let stopRequested = false;
let lastSourceValues = [];

/* =========================
   LIVE COUNT
========================= */
function updateLiveCount() {

    const values = (dataBox.value || "")
        .split("\n")
        .map(v => v.trim())
        .filter(v => v !== "");

    status.innerText = `🟢 Ready | Items: ${values.length}`;
}

dataBox.addEventListener("input", updateLiveCount);

/* =========================
   UPLOAD EXCEL
========================= */
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

status.innerText = `Loaded ✅ (${cleaned.length})`;

};

reader.readAsArrayBuffer(file);
};

/* =========================
   CLEAR
========================= */
document.getElementById("clearBtn").onclick = () => {

dataBox.value = "";
progressBar.style.width = "0%";

status.innerText = "Cleared 🧹";
};

/* =========================
   STOP
========================= */
document.getElementById("stopBtn").onclick = () => {

stopRequested = true;
status.innerText = "Stopped ⛔";

};

/* =========================
   PASTE LOGIC
========================= */
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
}}}

if (!values.length) {
alert("اكتب أو ارفع بيانات");
return;
}

lastSourceValues = [...values];

stopRequested = false;
progressBar.style.width = "0%";

status.innerText = "Processing ⚡";

try {

await Excel.run(async (context) => {

const sheet = context.workbook.worksheets.getActiveWorksheet();

const selected = context.workbook.getSelectedRange();
selected.load("rowIndex,columnIndex");

await context.sync();

const startRow = selected.rowIndex;
const col = selected.columnIndex;

const scanRange = sheet.getRangeByIndexes(startRow, col, 100000, 1);

const visibleRange = scanRange.getSpecialCells(Excel.SpecialCellType.visible);

visibleRange.load("areas/items/rowCount");

await context.sync();

let valueIndex = 0;

for (const area of visibleRange.areas.items) {

for (let r = 0; r < area.rowCount; r++) {

if (stopRequested) break;

if (valueIndex >= values.length) break;

const cell = area.getCell(r, 0);
cell.values = [[values[valueIndex]]];

valueIndex++;

const percent = Math.round((valueIndex / values.length) * 100);
progressBar.style.width = percent + "%";

}

if (stopRequested) break;
}

await context.sync();

progressBar.style.width = "100%";
status.innerText = "Done 🚀";

});

} catch (err) {
console.log(err);
status.innerText = "Error ❌ " + err.message;
}

};

/* =========================
   CHECK BUTTON
========================= */
checkBtn.onclick = async () => {

status.innerText = "Checking... 🔍";

try {

await Excel.run(async (context) => {

const range = context.workbook.getSelectedRange();
range.load("values");

await context.sync();

const excelValues = (range.values || []).flat()
.map(v => String(v || "").trim())
.filter(v => v !== "");

const pastedCount = excelValues.length;
const sourceCount = (dataBox.value || "")
.split("\n")
.filter(v => v.trim() !== "").length;

status.innerText =
`📊 Excel: ${pastedCount} | Box: ${sourceCount}`;

});

} catch (err) {
console.log(err);
status.innerText = "Check Error ❌ " + err.message;
}

};

/* =========================
   REPORT BUTTON
========================= */
reportBtn.onclick = async () => {

status.innerText = "Generating Report... 📊";

try {

await Excel.run(async (context) => {

const range = context.workbook.getSelectedRange();
range.load("values, rowCount, columnCount");

await context.sync();

if (range.columnCount < 2) {
status.innerText = "⚠️ Select 2 columns (Account + Value)";
return;
}

const data = range.values;

const report = [["Account", "Value"]];

for (let i = 0; i < range.rowCount; i++) {

const acc = data[i][0];
const val = data[i][1];

if (acc !== "" || val !== "") {
report.push([acc, val]);
}

}

// download Excel file
const wb = XLSX.utils.book_new();
const ws = XLSX.utils.aoa_to_sheet(report);

XLSX.utils.book_append_sheet(wb, ws, "Report");

XLSX.writeFile(wb, "Excel_Report.xlsx");

status.innerText = "Report Downloaded ✅";

});

} catch (err) {
console.log(err);
status.innerText = "Report Error ❌ " + err.message;
}

};

});