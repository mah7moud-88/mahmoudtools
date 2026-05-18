Office.onReady(() => {

    const uploadBtn = document.getElementById("uploadBtn");
    const fileInput = document.getElementById("fileInput");
    const dataBox = document.getElementById("dataBox");
    const repeatCountEl = document.getElementById("repeatCount");
    const status = document.getElementById("status");
    const countEl = document.getElementById("count");
    const liveCountEl = document.getElementById("liveCount");
    const progressBar = document.getElementById("progressBar");

    let stopRequested = false;

    // ======================
    // Live Count
    // ======================
    function updateLiveCount() {

        const values = dataBox.value
            .split("\n")
            .map(v => v.trim())
            .filter(v => v !== "");

        liveCountEl.innerText = values.length;
    }

    dataBox.addEventListener("input", updateLiveCount);

    // ======================
    // Upload
    // ======================
    uploadBtn.onclick = () => {
        fileInput.value = "";
        fileInput.click();
    };

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

            const cleaned = json
                .flat()
                .map(v => String(v).trim())
                .filter(v => v !== "");

            dataBox.value = cleaned.join("\n");

            updateLiveCount();

            status.innerText = "Loaded ✅";
        };

        reader.readAsArrayBuffer(file);
    };

    // ======================
    // Clear
    // ======================
    document.getElementById("clearBtn").onclick = () => {

        dataBox.value = "";

        updateLiveCount();

        countEl.innerText = "0";

        progressBar.style.width = "0%";

        status.innerText = "Cleared 🧹";
    };

    // ======================
    // Stop
    // ======================
    document.getElementById("stopBtn").onclick = () => {

        stopRequested = true;

        status.innerText = "⛔ Stopping...";
    };

    // ======================
    // Paste
    // ======================
    document.getElementById("run").onclick = async () => {

        const repeatTimes = parseInt(repeatCountEl.value || "0");

        const baseValues = dataBox.value
            .split("\n")
            .map(v => v.trim())
            .filter(v => v !== "");

        updateLiveCount();

        let values = [];

        // ======================
        // Repeat Logic
        // ======================
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

        stopRequested = false;
        progressBar.style.width = "0%";
        status.innerText = "Processing... ⏳";

        try {

            await Excel.run(async (context) => {

                const sheet = context.workbook.worksheets.getActiveWorksheet();

                const selected = context.workbook.getSelectedRange();
                selected.load("rowIndex,columnIndex");

                await context.sync();

                const startRow = selected.rowIndex;
                const col = selected.columnIndex;

                const scanRange = sheet.getRangeByIndexes(
                    startRow,
                    col,
                    100000,
                    1
                );

                const visibleRange = scanRange.getSpecialCells(
                    Excel.SpecialCellType.visible
                );

                visibleRange.load("areas/items/rowCount");

                await context.sync();

                let valueIndex = 0;

                for (const area of visibleRange.areas.items) {

                    for (let r = 0; r < area.rowCount; r++) {

                        if (stopRequested) {
                            status.innerText = "⛔ Stopped";
                            break;
                        }

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

                countEl.innerText = valueIndex;

                progressBar.style.width = "100%";

                if (!stopRequested) {
                    status.innerText = "Done 🚀";
                }

            });

        } catch (err) {

            console.log(err);

            status.innerText = "Error ❌ " + err.message;
        }
    };

});