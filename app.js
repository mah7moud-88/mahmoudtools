Office.onReady(() => {

    const uploadBtn = document.getElementById("uploadBtn");
    const fileInput = document.getElementById("fileInput");
    const dataBox = document.getElementById("dataBox");
    const status = document.getElementById("status");

    // ======================
    // 📂 Upload Excel file
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

            status.innerText = "Loaded ✅";
        };

        reader.readAsArrayBuffer(file);
    };

    // ======================
    // 🧹 Clear
    // ======================
    document.getElementById("clearBtn").onclick = () => {
        dataBox.value = "";
        status.innerText = "Cleared 🧹";
    };

    // ======================
    // ⚡ Paste to Excel (FIXED)
    // ======================
    document.getElementById("run").onclick = async () => {

        const values = dataBox.value
            .split("\n")
            .map(v => v.trim())
            .filter(v => v);

        if (!values.length) {
            alert("اكتب أو ارفع بيانات");
            return;
        }

        try {

            await Excel.run(async (context) => {

                const sheet = context.workbook.worksheets.getActiveWorksheet();

                const selected = context.workbook.getSelectedRange();
                selected.load("rowIndex, columnIndex");

                const usedRange = sheet.getUsedRange();
                usedRange.load("rowCount");

                await context.sync();

                let startRow = selected.rowIndex;
                let col = selected.columnIndex;

                let valueIndex = 0;
                const total = values.length;

                for (let i = 0; i < usedRange.rowCount; i++) {

                    if (valueIndex >= total) break;

                    const cell = sheet.getCell(startRow + i, col);
                    cell.load("rowHidden");

                    await context.sync();

                    // ✔️ تجاهل الصفوف المخفية
                    if (!cell.rowHidden) {
                        cell.values = [[values[valueIndex]]];
                        valueIndex++;
                    }
                }

                await context.sync();
            });

            status.innerText = "Done 🎉 (Visible only)";

        } catch (err) {
            console.log(err);
            status.innerText = "Error ❌ " + err.message;
        }
    };

});