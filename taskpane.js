Office.onReady(() => {

    const uploadBtn = document.getElementById("uploadBtn");
    const fileInput = document.getElementById("fileInput");
    const dataBox = document.getElementById("dataBox");
    const status = document.getElementById("status");
    const countEl = document.getElementById("count");
    const progressBar = document.getElementById("progressBar");

    function updateCount() {
        const values = dataBox.value.split("\n").map(v => v.trim()).filter(v => v);
        countEl.innerText = values.length;
    }

    dataBox.addEventListener("input", updateCount);

    // ======================
    // 📂 Upload Excel (FIXED clean read)
    // ======================
    uploadBtn.onclick = function () {
        fileInput.value = "";
        fileInput.click();
    };

    fileInput.onchange = function (e) {

        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();

        reader.onload = function (event) {

            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: "array" });

            const sheet = workbook.Sheets[workbook.SheetNames[0]];

            const json = XLSX.utils.sheet_to_json(sheet, {
                header: 1,
                defval: ""
            });

            const cleaned = json
                .flat()
                .map(v => String(v).replace(/^\uFEFF/, "").trim())
                .filter(v => v !== "");

            dataBox.value = cleaned.join("\n");

            status.innerText = "Loaded ✅";
            updateCount();
        };

        reader.readAsArrayBuffer(file);

        fileInput.value = "";
    };

    // ======================
    // 🧹 Clear
    // ======================
    document.getElementById("clearBtn").onclick = function () {
        dataBox.value = "";
        status.innerText = "Cleared 🧹";
        updateCount();
        progressBar.style.width = "0%";
    };

    // ======================
    // ⚡ LOOP PASTE (FIXED FOR FILTER)
    // ======================
    document.getElementById("run").onclick = async function () {

        const values = dataBox.value
            .split("\n")
            .map(v => v.trim())
            .filter(v => v);

        if (!values.length) {
            alert("اكتب أو ارفع بيانات");
            return;
        }

        progressBar.style.width = "0%";

        try {

            await Excel.run(async (context) => {

                const sheet = context.workbook.worksheets.getActiveWorksheet();

                const usedRange = sheet.getUsedRange();
                usedRange.load("rowCount");

                const selected = context.workbook.getSelectedRange();
                selected.load("rowIndex, columnIndex");

                await context.sync();

                let startRow = selected.rowIndex;
                let col = selected.columnIndex;

                let valueIndex = 0;
                const total = values.length;

                for (let i = startRow; i < usedRange.rowCount; i++) {

                    const cell = sheet.getCell(i, col);
                    cell.load("rowHidden");

                    await context.sync();

                    // 🔥 نتجاهل الصفوف المخفية (Filter safe)
                    if (!cell.rowHidden) {

                        if (valueIndex >= total) break;

                        cell.values = [[values[valueIndex]]];
                        valueIndex++;

                        // Progress
                        const percent = Math.round((valueIndex / total) * 100);
                        progressBar.style.width = percent + "%";
                    }
                }

                await context.sync();
            });

            progressBar.style.width = "100%";
            status.innerText = `🎉 Done! ${values.length} values pasted successfully`;

        } catch (err) {
            console.log(err);
            status.innerText = "ERROR ❌ " + err.message;
        }
    };

});