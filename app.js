Office.onReady(() => {

const uploadBtn = document.getElementById("uploadBtn");
const fileInput = document.getElementById("fileInput") || document.createElement("input");
const dataBox = document.getElementById("dataBox");
const repeatCountEl = document.getElementById("repeatCount");
const status = document.getElementById("status");
const progressBar = document.getElementById("progressBar");

const loadedCount = document.getElementById("loadedCount");
const totalDash = document.getElementById("totalDash");
const progressDash = document.getElementById("progressDash");
const timeDash = document.getElementById("timeDash");

const checkBtn = document.getElementById("checkBtn");
const searchBtn = document.getElementById("searchBtn");

let stopRequested = false;
let startTime = 0;
let timerInterval = null;

let lastSourceValues = [];

/* TIMER */
function startTimer() {
    startTime = Date.now();
    clearInterval(timerInterval);

    timerInterval = setInterval(() => {
        const sec = Math.floor((Date.now() - startTime) / 1000);
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

/* SEARCH FIXED */
searchBtn.onclick = () => {
    prompt("🔎 ادخل الأرقام (افصل بينهم بفاصلة):");
};

/* CHECK */
checkBtn.onclick = async () => {
    setStatus("Checking... 🔍");

    try {
        await Excel.run(async (context) => {

            const range = context.workbook.getSelectedRange();
            range.load("values");

            await context.sync();

            const excelValues = range.values.flat();

            const pastedCount = excelValues
                .map(v => String(v).trim())
                .filter(v => v !== "")
                .length;

            const sourceCount = lastSourceValues.length || 0;

            setStatus(`📊 Pasted: ${pastedCount} | Source: ${sourceCount}`);
        });

    } catch (err) {
        setStatus("Error ❌ " + err.message);
    }
};

});