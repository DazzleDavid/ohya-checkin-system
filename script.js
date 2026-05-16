let excelData = [];

const excelFile = document.getElementById("excelFile");
const fileName = document.getElementById("fileName");
const searchBtn = document.getElementById("searchBtn");
const searchInput = document.getElementById("searchInput");
const resultDiv = document.getElementById("result");

excelFile.addEventListener("change", (e) => {
    const file = e.target.files[0];

    if (!file) return;

    fileName.textContent = `已選擇：${file.name}`;

    const reader = new FileReader();

    reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);

        const workbook = XLSX.read(data, {
            type: "array"
        });

        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        excelData = XLSX.utils.sheet_to_json(firstSheet);

        resultDiv.innerHTML = `
            <div class="success">
                <h2>✅ Excel 載入成功</h2>
                <p>共載入 ${excelData.length} 筆報名資料。</p>
            </div>
        `;

        console.log(excelData);
    };

    reader.readAsArrayBuffer(file);
});

searchBtn.addEventListener("click", () => {
    const keyword = searchInput.value.trim();

    if (!keyword) {
        resultDiv.innerHTML = `
            <div class="warning">
                <h2>⚠️ 請輸入查詢內容</h2>
                <p>可以輸入姓名、職別碼、Email 或電話。</p>
            </div>
        `;
        return;
    }

    if (excelData.length === 0) {
        resultDiv.innerHTML = `
            <div class="warning">
                <h2>⚠️ 尚未上傳 Excel</h2>
                <p>請先上傳活動報名名單。</p>
            </div>
        `;
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
                <p>此使用者不在報名名單內。</p>
            </div>
        `;
    }
});

searchInput.addEventListener("keypress", (e) => {
    if (e.key === "Enter") {
        searchBtn.click();
    }
});