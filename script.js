let excelData = [];

const excelFile = document.getElementById("excelFile");
const searchBtn = document.getElementById("searchBtn");
const searchInput = document.getElementById("searchInput");
const resultDiv = document.getElementById("result");

// 上傳 Excel
excelFile.addEventListener("change", (e) => {
    const file = e.target.files[0];

    if (!file) return;

    const reader = new FileReader();

    reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);

        const workbook = XLSX.read(data, {
            type: "array"
        });

        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

        excelData = XLSX.utils.sheet_to_json(firstSheet);

        alert(`成功載入 ${excelData.length} 筆資料`);

        console.log(excelData);
    };

    reader.readAsArrayBuffer(file);
});

// 查詢功能
searchBtn.addEventListener("click", () => {
    const keyword = searchInput.value.trim();

    if (!keyword) {
        alert("請輸入查詢內容");
        return;
    }

    if (excelData.length === 0) {
        alert("請先上傳 Excel 檔案");
        return;
    }

    const found = excelData.find((item) => {
        return (
            String(item["姓名"] || "").includes(keyword) ||
            String(item["職別碼(學號/員編)"] || "").includes(keyword) ||
            String(item["e-mail"] || "").includes(keyword) ||
            String(item["聯絡電話"] || "").includes(keyword)
        );
    });

    if (found) {
        resultDiv.innerHTML = `
            <div class="success">
                <h2>✅ 驗證成功</h2>

                <div class="result-item">
                    <strong>姓名：</strong>
                    ${found["姓名"] || "無資料"}
                </div>

                <div class="result-item">
                    <strong>職別碼：</strong>
                    ${found["職別碼(學號/員編)"] || "無資料"}
                </div>

                <div class="result-item">
                    <strong>單位：</strong>
                    ${found["單位別(班級)"] || "無資料"}
                </div>

                <div class="result-item">
                    <strong>Email：</strong>
                    ${found["e-mail"] || "無資料"}
                </div>

                <div class="result-item">
                    <strong>聯絡電話：</strong>
                    ${found["聯絡電話"] || "無資料"}
                </div>

                <div class="result-item">
                    <strong>是否參與：</strong>
                    ${found["是否參與"] || "無資料"}
                </div>
            </div>
        `;
    } else {
        resultDiv.innerHTML = `
            <div class="error">
                <h2>❌ 查無資料</h2>
                <p>此使用者不在報名名單內</p>
            </div>
        `;
    }
});

// Enter 查詢
searchInput.addEventListener("keypress", (e) => {
    if (e.key === "Enter") {
        searchBtn.click();
    }
});