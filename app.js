// ===============================
// فتح/غلق بوكس البحث
// ===============================
document.getElementById("openSearchBtn").addEventListener("click", function () {

    const panel = document.getElementById("searchPanel");

    panel.style.display = (panel.style.display === "block") ? "none" : "block";

});


// ===============================
// زر البحث
// ===============================
document.getElementById("searchBtn").addEventListener("click", searchAccount);


// ===============================
// البحث في Excel
// ===============================
async function searchAccount() {

    const key = document.getElementById("searchInput").value.trim();

    if (!key || key.length !== 6) {
        document.getElementById("status").innerText = "❌ لازم 6 أرقام";
        return;
    }

    document.getElementById("status").innerText = "🔍 جاري البحث...";

    try {

        await Excel.run(async (context) => {

            const sheet = context.workbook.worksheets.getActiveWorksheet();

            const range = sheet.getRange("C:C");
            range.load("values");

            await context.sync();

            const values = range.values;

            let foundRow = -1;

            for (let i = 0; i < values.length; i++) {

                let v = values[i][0];

                if (v && v.toString().slice(-6) === key) {
                    foundRow = i + 1;
                    break;
                }
            }

            if (foundRow === -1) {
                document.getElementById("status").innerText = "❌ لم يتم العثور";
                return;
            }

            const row = sheet.getRange(`A${foundRow}:C${foundRow}`);
            row.select();

            await context.sync();

            document.getElementById("status").innerText = "✅ تم العثور وتحديد الصف";

        });

    } catch (err) {
        console.log(err);
        document.getElementById("status").innerText = "❌ خطأ في البحث";
    }
}