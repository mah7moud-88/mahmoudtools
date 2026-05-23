Office.onReady(() => {

const uploadBtn = document.getElementById("uploadBtn");
const fileInput = document.getElementById("fileInput");
const dataBox = document.getElementById("dataBox");
const repeatCountEl = document.getElementById("repeatCount");
const status = document.getElementById("status");
const progressBar = document.getElementById("progressBar");

const loadedCount = document.getElementById("loadedCount");
const totalDash = document.getElementById("totalDash");
const progressDash = document.getElementById("progressDash");
const timeDash = document.getElementById("timeDash");

const checkBtn = document.getElementById("checkBtn");
const filterModeEl = document.getElementById("filterMode");
const filterLabel = document.getElementById("filterLabel");

let stopRequested = false;
let startTime = 0;
let timerInterval = null;

let lastSourceValues = [];

/* =========================
   TIMER
========================= */
function startTimer() {

    startTime = Date.now();

    clearInterval(timerInterval);

    timerInterval = setInterval(() => {

        const sec =
            Math.floor((Date.now() - startTime) / 1000);

        timeDash.innerText = sec + "s";

    }, 500);
}

function stopTimer() {
    clearInterval(timerInterval);
}

/* =========================
   STATUS
========================= */
function setStatus(text) {
    status.innerText = text;
}

/* =========================
   DASHBOARD
========================= */
function updateDashboard(total, progress) {
    totalDash.innerText = total || 0;
    progressDash.innerText = (progress || 0) + "%";
}

/* =========================
   LIVE COUNT
========================= */
dataBox.addEventListener("input", () => {

    const values = dataBox.value
        .split("\n")
        .map(v => v.trim())
        .filter(v => v !== "");

    loadedCount.innerText = values.length;
});

/* =========================
   UPLOAD
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

const cleaned = json.flat()
.map(v => String(v).trim())
.filter(v => v !== "");

if (cleaned.length > 100000) {
    alert("الملف كبير جدًا");
    return;
}

dataBox.value = cleaned.join("\n");

loadedCount.innerText = cleaned.length;

setStatus("Loaded ✅");

fileInput.value = "";
};

reader.readAsArrayBuffer(file);
};

/* =========================
   CLEAR
========================= */
document.getElementById("clearBtn").onclick = () => {

dataBox.value = "";

progressBar.style.width = "0%";

loadedCount.innerText = "0";
totalDash.innerText = "0";
progressDash.innerText = "0%";
timeDash.innerText = "0s";

setStatus("Cleared 🧹");
};

/* =========================
   STOP
========================= */
document.getElementById("stopBtn").onclick = () => {

stopRequested = true;
setStatus("Stopped ⛔");
stopTimer();
};

/* =========================
   FILTER LABEL TOGGLE
========================= */
filterModeEl.addEventListener("change", function () {

filterLabel.innerText =
    this.checked
    ? "🔍 Filter Mode"
    : "🚀 Fast Mode";
});

/* =========================
   PASTE ENGINE
========================= */
document.getElementById("run").onclick = async () => {

const repeatTimes = parseInt(repeatCountEl.value || "0");

const baseValues = dataBox.value
.split("\n")
.map(v => v.trim())
.filter(v => v !== "");

/* REPEAT LOGIC (UNCHANGED) */
let values = [];

if (repeatTimes <= 0) {

    values = baseValues;

} else {

    for (const v of baseValues) {
        for (let i = 0; i < repeatTimes; i++) {
            values.push(v);
        }
    }
}

if (!values.length) {
    alert("اكتب أو ارفع بيانات");
    return;
}

lastSourceValues = [...values];

stopRequested = false;
progressBar.style.width = "0%";

setStatus("Processing ⚡");

startTimer();

updateDashboard(values.length, 0);

try {

await Excel.run(async (context) => {

    context.application.suspendScreenUpdatingUntilNextSync();

    const sheet =
        context.workbook.worksheets.getActiveWorksheet();

    const selected =
        context.workbook.getSelectedRange();

    selected.load("rowIndex,columnIndex,rowCount,columnCount");

    await context.sync();

    const scanRange =
        sheet.getRangeByIndexes(
            selected.rowIndex,
            selected.columnIndex,
            selected.rowCount,
            selected.columnCount
        );

    const useFilter = filterModeEl.checked;

    let valueIndex = 0;

    /* =========================
       🚀 FAST MODE
    ========================= */
    if (!useFilter) {

        const batch = values.map(v => [v]);

        const targetRange =
            sheet.getRangeByIndexes(
                selected.rowIndex,
                selected.columnIndex,
                batch.length,
                1
            );

        targetRange.values = batch;

        progressBar.style.width = "100%";
        updateDashboard(values.length, 100);
        setStatus("Done 🚀 Fast Mode");
    }

    /* =========================
       🔍 FILTER MODE
    ========================= */
    else {

        let visibleRange;

        try {

            visibleRange =
                scanRange.getSpecialCells(
                    Excel.SpecialCellType.visible
                );

        } catch {

            setStatus("No visible cells ❌");
            stopTimer();
            return;
        }

        visibleRange.load("values");

        await context.sync();

        const flatVisible = visibleRange.values.flat();

        let pastedCount = 0;
        let emptyCount = 0;

        for (const v of flatVisible) {

            const val = String(v).trim();

            if (val === "") {
                emptyCount++;
            } else {
                pastedCount++;
            }
        }

        const sourceCount = lastSourceValues.length || 0;

        const missing = sourceCount - pastedCount;

        if (missing > 0) {

            setStatus(
                `⚠️ Missing: ${missing} | Visible: ${pastedCount} | Empty: ${emptyCount}`
            );

        } else {

            setStatus(
                `✅ OK | Visible: ${pastedCount} | Empty: ${emptyCount}`
            );
        }

        progressBar.style.width = "100%";
        updateDashboard(values.length, 100);
    }

    stopTimer();
});

} catch (err) {

    console.log(err);

    setStatus("Error ❌ " + err.message);

    stopTimer();
}
};

/* =========================
   CHECK (FIXED - VISIBLE ONLY)
========================= */
checkBtn.onclick = async () => {

setStatus("Checking... 🔍");

try {

await Excel.run(async (context) => {

    const sheet =
        context.workbook.worksheets.getActiveWorksheet();

    const range =
        context.workbook.getSelectedRange();

    range.load("rowIndex,columnIndex,rowCount,columnCount");

    await context.sync();

    const scanRange =
        sheet.getRangeByIndexes(
            range.rowIndex,
            range.columnIndex,
            range.rowCount,
            range.columnCount
        );

    let visibleRange;

    try {

        visibleRange =
            scanRange.getSpecialCells(
                Excel.SpecialCellType.visible
            );

    } catch {

        setStatus("No visible cells ❌");
        return;
    }

    visibleRange.load("values");

    await context.sync();

    const flat = visibleRange.values.flat();

    let pastedCount = 0;
    let emptyCount = 0;

    for (const v of flat) {

        const val = String(v).trim();

        if (val === "") {
            emptyCount++;
        } else {
            pastedCount++;
        }
    }

    const sourceCount =
        lastSourceValues.length || 0;

    const missing =
        sourceCount - pastedCount;

    if (missing > 0) {

        setStatus(
            `⚠️ Missing: ${missing} | Visible: ${pastedCount} | Empty: ${emptyCount}`
        );

    } else {

        setStatus(
            `✅ OK | Visible: ${pastedCount} | Empty: ${emptyCount}`
        );
    }

});

} catch (err) {

    setStatus("Check Error ❌ " + err.message);
}

};

});