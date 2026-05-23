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

let stopRequested = false;
let startTime = 0;
let timerInterval = null;

let lastSourceValues = [];

/* TIMER */
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

/* STATUS */
function setStatus(text) {

    status.innerText = text;
}

/* DASHBOARD */
function updateDashboard(total, progress) {

    totalDash.innerText = total || 0;

    progressDash.innerText = (progress || 0) + "%";
}

/* LIVE COUNT */
function updateLiveCount() {

    const values = dataBox.value
        .split("\n")
        .map(v => v.trim())
        .filter(v => v !== "");

    loadedCount.innerText = values.length;
}

dataBox.addEventListener("input", updateLiveCount);

/* UPLOAD */
uploadBtn.onclick = () => fileInput.click();

fileInput.onchange = (e) => {

    const file = e.target.files[0];

    if (!file) return;

    const reader = new FileReader();

    reader.onload = (event) => {

        const data =
            new Uint8Array(event.target.result);

        const workbook =
            XLSX.read(data, { type: "array" });

        const sheet =
            workbook.Sheets[
                workbook.SheetNames[0]
            ];

        const json =
            XLSX.utils.sheet_to_json(sheet, {
                header: 1,
                defval: ""
            });

        const cleaned = json
            .flat()
            .map(v => String(v).trim())
            .filter(v => v !== "");

        /* LIMIT */
        if (cleaned.length > 100000) {

            alert("الملف ضخم جدًا");

            return;
        }

        dataBox.value = cleaned.join("\n");

        loadedCount.innerText = cleaned.length;

        setStatus("Loaded ✅");

        /* RESET FILE INPUT */
        fileInput.value = "";
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

    setStatus("Cleared 🧹");
};

/* STOP */
document.getElementById("stopBtn").onclick = () => {

    stopRequested = true;

    setStatus("Stopped ⛔");

    stopTimer();
};

/* PASTE */
document.getElementById("run").onclick = async () => {

    const repeatTimes =
        parseInt(repeatCountEl.value || "0");

    const baseValues = dataBox.value
        .split("\n")
        .map(v => v.trim())
        .filter(v => v !== "");

    let values = [];

    /* REPEAT (OLD LOGIC) */
    if (repeatTimes <= 0) {

        values = baseValues;

    } else {

        for (const v of baseValues) {

            for (let i = 0; i < repeatTimes; i++) {

                values.push(v);
            }
        }
    }

    /* EMPTY CHECK */
    if (!values.length) {

        alert("اكتب أو ارفع بيانات");

        return;
    }

    /* SAVE SOURCE */
    lastSourceValues = [...values];

    stopRequested = false;

    progressBar.style.width = "0%";

    setStatus("Processing ⚡");

    startTimer();

    updateDashboard(values.length, 0);

    try {

        await Excel.run(async (context) => {

            /* SPEED BOOST */
            context.application
                .suspendScreenUpdatingUntilNextSync();

            const sheet =
                context.workbook.worksheets
                .getActiveWorksheet();

            const selected =
                context.workbook.getSelectedRange();

            selected.load("rowIndex,columnIndex");

            await context.sync();

            const startRow = selected.rowIndex;

            const col = selected.columnIndex;

            const scanRange =
                sheet.getRangeByIndexes(
                    startRow,
                    col,
                    100000,
                    1
                );

            let visibleRange;

            /* FILTER SAFE */
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

            visibleRange.load("areas/items/rowCount");

            await context.sync();

            let valueIndex = 0;

            /* LOOP VISIBLE AREAS */
            for (const area of visibleRange.areas.items) {

                if (stopRequested) break;

                let batch = [];

                /* BUILD BATCH */
                for (let r = 0; r < area.rowCount; r++) {

                    if (stopRequested) break;

                    if (valueIndex >= values.length)
                        break;

                    batch.push([
                        values[valueIndex]
                    ]);

                    valueIndex++;
                }

                /* WRITE BATCH */
                if (batch.length > 0) {

                    const targetRange =
                        area.getCell(0, 0)
                        .getResizedRange(
                            batch.length - 1,
                            0
                        );

                    targetRange.values = batch;
                }

                /* PROGRESS */
                const percent = Math.round(
                    (valueIndex / values.length) * 100
                );

                progressBar.style.width =
                    percent + "%";

                updateDashboard(
                    values.length,
                    percent
                );
            }

            await context.sync();

            /* FINAL STATUS */
            if (!stopRequested) {

                progressBar.style.width = "100%";

                updateDashboard(
                    values.length,
                    100
                );

                setStatus("Done 🚀");

            } else {

                setStatus("Stopped ⛔");
            }

            stopTimer();
        });

    } catch (err) {

        console.log(err);

        setStatus("Error ❌ " + err.message);

        stopTimer();
    }
};

/* CHECK */
checkBtn.onclick = async () => {

    setStatus("Checking... 🔍");

    try {

        await Excel.run(async (context) => {

            const range =
                context.workbook.getSelectedRange();

            range.load("values");

            await context.sync();

            const excelValues =
                range.values.flat();

            const pastedCount = excelValues
                .map(v => String(v).trim())
                .filter(v => v !== "")
                .length;

            const sourceCount =
                lastSourceValues.length || 0;

            setStatus(
                `📊 Pasted: ${pastedCount} value(s) | Source: ${sourceCount}`
            );
        });

    } catch (err) {

        console.log(err);

        setStatus(
            "Check Error ❌ " + err.message
        );
    }
};

});