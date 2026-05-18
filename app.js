Office.onReady(() => {

    const uploadBtn = document.getElementById("uploadBtn");
    const fileInput = document.getElementById("fileInput");
    const dataBox = document.getElementById("dataBox");
    const status = document.getElementById("status");
    const countEl = document.getElementById("count");
    const progressBar = document.getElementById("progressBar");

    // ======================
    // 📂 Upload Excel File
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

            const data =
                new Uint8Array(event.target.result);

            const workbook = XLSX.read(data, {
                type: "array"
            });

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

        countEl.innerText = "0";

        progressBar.style.width = "0%";
    };

    // ======================
    // ⚡ Paste To Visible Rows Only
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

            status.innerText = "Processing... ⏳";

            progressBar.style.width = "0%";

            await Excel.run(async (context) => {

                const sheet =
                    context.workbook
                        .worksheets
                        .getActiveWorksheet();

                const selected =
                    context.workbook
                        .getSelectedRange();

                selected.load(
                    "rowIndex,columnIndex"
                );

                const usedRange =
                    sheet.getUsedRange();

                usedRange.load("rowCount");

                await context.sync();

                const startRow =
                    selected.rowIndex;

                const col =
                    selected.columnIndex;

                // ======================
                // تجهيز الصفوف
                // ======================

                const rowsData = [];

                for (let i = 0; i < usedRange.rowCount; i++) {

                    const rowRange =
                        sheet.getRangeByIndexes(
                            startRow + i,
                            col,
                            1,
                            1
                        );

                    const entireRow =
                        rowRange.getEntireRow();

                    entireRow.load("hidden");

                    rowsData.push({
                        row: startRow + i,
                        rowObj: entireRow
                    });
                }

                // تحميل hidden مرة واحدة
                await context.sync();

                // ======================
                // لصق القيم
                // ======================

                let valueIndex = 0;

                for (let i = 0; i < rowsData.length; i++) {

                    if (valueIndex >= values.length) {
                        break;
                    }

                    const rowData =
                        rowsData[i];

                    // تجاهل الصفوف المخفية
                    if (!rowData.rowObj.hidden) {

                        const targetCell =
                            sheet.getCell(
                                rowData.row,
                                col
                            );

                        targetCell.values = [
                            [values[valueIndex]]
                        ];

                        valueIndex++;

                        // Progress
                        const percent =
                            Math.round(
                                (valueIndex / values.length) * 100
                            );

                        progressBar.style.width =
                            percent + "%";
                    }
                }

                // تنفيذ الكتابة مرة واحدة
                await context.sync();

                // ======================
                // UI
                // ======================

                countEl.innerText =
                    valueIndex;

                progressBar.style.width =
                    "100%";

                status.innerText =
                    "Done 🚀 (Visible Rows Only)";
            });

        } catch (err) {

            console.log(err);

            status.innerText =
                "Error ❌ " + err.message;
        }
    };

});