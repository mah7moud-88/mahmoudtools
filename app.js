Office.onReady(() => {

const uploadBtn = document.getElementById("uploadBtn");
const fileInput = document.getElementById("fileInput");
const dataBox = document.getElementById("dataBox");
const repeatCountEl = document.getElementById("repeatCount");
const status = document.getElementById("status");
const progressBar = document.getElementById("progressBar");

const checkBtn = document.getElementById("checkBtn");

const loadedCount = document.getElementById("loadedCount");
const totalDash = document.getElementById("totalDash");
const progressDash = document.getElementById("progressDash");
const timeDash = document.getElementById("timeDash");

let stopRequested = false;
let startTime = 0;
let timerInterval = null;

/* TIMER */
function startTimer() {
startTime = Date.now();
timerInterval = setInterval(() => {
timeDash.innerText = Math.floor((Date.now() - startTime) / 1000) + "s";
}, 500);
}

function stopTimer() {
clearInterval(timerInterval);
}

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
totalDash.innerText = cleaned.length;

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

/* PASTE - FIXED LIVE PROGRESS */
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

stopRequested = false;

let total = values.length;
let done = 0;

progressBar.style.width = "0%";
loadedCount.innerText = total;
totalDash.innerText = total;
progressDash.innerText = "0%";

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

/* 🔥 LIVE LOOP (IMPORTANT FIX) */
for (let i = 0; i < total; i++) {

if (stopRequested) break;

const cell = sheet.getCell(startRow + i, col);
cell.values = [[values[i]]];

done++;

let percent = Math.round((done / total) * 100);

progressBar.style.width = percent + "%";
progressDash.innerText = percent + "%";
status.innerText = `Processing ⚡ ${done}/${total}`;

}

/* commit changes */
await context.sync();

status.innerText = "Done 🚀";
progressBar.style.width = "100%";
progressDash.innerText = "100%";

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

});