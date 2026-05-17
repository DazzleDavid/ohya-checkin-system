let excelData = [];
let checkinRecords = [];

const excelFile = document.getElementById("excelFile");
const fileName = document.getElementById("fileName");

const searchBtn = document.getElementById("searchBtn");
const searchInput = document.getElementById("searchInput");

const exportBtn = document.getElementById("exportBtn");
const recordCount = document.getElementById("recordCount");

const recordFiles = document.getElementById("recordFiles");
const recordFileName = document.getElementById("recordFileName");

const drawBtn = document.getElementById("drawBtn");
const drawCount = document.getElementById("drawCount");

const verifyResult = document.getElementById("verifyResult");
const drawResult = document.getElementById("drawResult");

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

        excelData = XLSX.utils.sheet_to_json(firstSheet, {
            defval: ""
        });

        verifyResult.innerHTML = `
            <div class="success">
                <h2>✅ Excel 載入成功</h2>
                <p>共載入 ${excelData.length} 筆報名資料。</p>
            </div>
        `;
    };

    reader.readAsArrayBuffer(file);
});

searchBtn.addEventListener("click", () => {
    const keyword = searchInput.value.trim();

    if (!keyword) {
        showVerifyWarning("請輸入查詢內容", "可以輸入姓名、職別碼、Email 或電話。");
        return;
    }

    if (excelData.length === 0) {
        showVerifyWarning("尚未上傳 Excel", "請先上傳活動報名名單。");
        return;
    }

    const cleanKeyword = keyword.replace(/\s/g, "").toLowerCase();

    const found = excelData.find((item) => {
        const name = String(item["姓名"] || "").trim();
        const id = String(item["職別碼(學號/員編)"] || "").trim();
        const email = String(item["e-mail"] || "").trim().toLowerCase();
        const phone = String(item["聯絡電話"] || "").replace(/\s/g, "");

        return (
            name.includes(keyword) ||
            id.includes(keyword) ||
            email.includes(cleanKeyword) ||
            phone.includes(cleanKeyword)
        );
    });

    if (!found) {
        verifyResult.innerHTML = `
            <div class="error">
                <h2>❌ 查無資料</h2>
                <p>此使用者不在報名名單內。</p>
            </div>
        `;
        return;
    }

    const recordStatus = addCheckinRecord(found);
    updateRecordCount();

    verifyResult.innerHTML = `
        <div class="success">
            <h2>✅ 驗證成功</h2>
            <p>${recordStatus}</p>

            <div class="result-item"><strong>姓名：</strong>${found["姓名"] || "無資料"}</div>
            <div class="result-item"><strong>職別碼：</strong>${found["職別碼(學號/員編)"] || "無資料"}</div>
            <div class="result-item"><strong>單位：</strong>${found["單位別(班級)"] || "無資料"}</div>
            <div class="result-item"><strong>身份別：</strong>${found["身份別"] || "無資料"}</div>
            <div class="result-item"><strong>Email：</strong>${found["e-mail"] || "無資料"}</div>
            <div class="result-item"><strong>聯絡電話：</strong>${found["聯絡電話"] || "無資料"}</div>
            <div class="result-item"><strong>是否參與：</strong>${found["是否參與"] || "無資料"}</div>
        </div>
    `;
});

function addCheckinRecord(person) {
    const record = {
        uniqueKey: getUniqueKey(person),
        name: String(person["姓名"] || "").trim(),
        id: String(person["職別碼(學號/員編)"] || "").trim(),
        unit: String(person["單位別(班級)"] || "").trim(),
        identity: String(person["身份別"] || "").trim(),
        email: String(person["e-mail"] || "").trim(),
        phone: String(person["聯絡電話"] || "").trim(),
        checkinTime: formatDateTime(new Date())
    };

    const exists = checkinRecords.some(item => item.uniqueKey === record.uniqueKey);

    if (exists) {
        return "⚠️ 此人已驗證過，不重複紀錄。";
    }

    checkinRecords.push(record);
    return "✅ 已加入驗證紀錄。";
}

function getUniqueKey(person) {
    const id = String(person["職別碼(學號/員編)"] || "").trim();
    const email = String(person["e-mail"] || "").trim().toLowerCase();
    const phone = String(person["聯絡電話"] || "").replace(/\s/g, "");
    const name = String(person["姓名"] || "").trim();

    return id || email || phone || name;
}

recordFiles.addEventListener("change", async (e) => {
    const files = Array.from(e.target.files);

    if (files.length === 0) return;

    recordFileName.textContent = `已選擇 ${files.length} 個紀錄檔`;

    let importedCount = 0;

    for (const file of files) {
        const text = await file.text();
        const records = parseRecordTxt(text);

        records.forEach(record => {
            const added = addImportedRecord(record);
            if (added) importedCount++;
        });
    }

    updateRecordCount();

    drawResult.innerHTML = `
        <div class="success">
            <h2>✅ 紀錄檔匯入完成</h2>
            <p>本次新增 ${importedCount} 筆，去重後共有 ${checkinRecords.length} 筆紀錄。</p>
        </div>
    `;
});

function parseRecordTxt(text) {
    const blocks = text.split("------------------------------------");
    const records = [];

    blocks.forEach(block => {
        const name = getTxtValue(block, "姓名");
        const id = getTxtValue(block, "職別碼");
        const unit = getTxtValue(block, "單位");
        const identity = getTxtValue(block, "身份別");
        const email = getTxtValue(block, "Email");
        const phone = getTxtValue(block, "聯絡電話");
        const checkinTime = getTxtValue(block, "驗證時間");

        if (name || id || email || phone) {
            records.push({
                uniqueKey: id || email.toLowerCase() || phone.replace(/\s/g, "") || name,
                name,
                id,
                unit,
                identity,
                email,
                phone,
                checkinTime
            });
        }
    });

    return records;
}

function getTxtValue(block, label) {
    const regex = new RegExp(`${label}：(.+)`);
    const match = block.match(regex);
    return match ? match[1].trim() : "";
}

function addImportedRecord(record) {
    const index = checkinRecords.findIndex(item => item.uniqueKey === record.uniqueKey);

    if (index === -1) {
        checkinRecords.push(record);
        return true;
    }

    const oldTime = new Date(checkinRecords[index].checkinTime);
    const newTime = new Date(record.checkinTime);

    if (newTime < oldTime) {
        checkinRecords[index] = record;
    }

    return false;
}

exportBtn.addEventListener("click", () => {
    if (checkinRecords.length === 0) {
        showVerifyWarning("尚無驗證紀錄", "請先完成驗證或匯入紀錄檔。");
        return;
    }

    let txtContent = "";

    txtContent += "歐耶~不能舒在壓力上 活動驗證紀錄\n";
    txtContent += "====================================\n";
    txtContent += `匯出時間：${formatDateTime(new Date())}\n`;
    txtContent += `總驗證人數：${checkinRecords.length}\n`;
    txtContent += "====================================\n\n";

    checkinRecords.forEach((record, index) => {
        txtContent += `第 ${index + 1} 筆\n`;
        txtContent += `姓名：${record.name}\n`;
        txtContent += `職別碼：${record.id}\n`;
        txtContent += `單位：${record.unit}\n`;
        txtContent += `身份別：${record.identity}\n`;
        txtContent += `Email：${record.email}\n`;
        txtContent += `聯絡電話：${record.phone}\n`;
        txtContent += `驗證時間：${record.checkinTime}\n`;
        txtContent += "------------------------------------\n";
    });

    downloadTxt(txtContent);
});

drawBtn.addEventListener("click", () => {
    const n = Number(drawCount.value);

    if (!n || n <= 0) {
        showDrawWarning("請輸入抽獎人數", "請輸入大於 0 的數字。");
        return;
    }

    if (checkinRecords.length === 0) {
        showDrawWarning("尚無驗證紀錄", "請先驗證或匯入紀錄檔。");
        return;
    }

    const winners = shuffleArray([...checkinRecords]).slice(0, Math.min(n, checkinRecords.length));

    drawResult.innerHTML = `
        <div class="draw-result">
            <h2>🎉 隨機抽出 ${winners.length} 位</h2>

            ${winners.map((winner, index) => `
                <div class="winner-item">
                    <strong>${index + 1}. ${winner.name}</strong><br>
                    職別碼：${winner.id || "無資料"}<br>
                    單位：${winner.unit || "無資料"}<br>
                    電話：${winner.phone || "無資料"}<br>
                    Email：${winner.email || "無資料"}
                </div>
            `).join("")}
        </div>
    `;
});

function shuffleArray(array) {
    for (let i = array.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [array[i], array[j]] = [array[j], array[i]];
    }

    return array;
}

function downloadTxt(content) {
    const blob = new Blob([content], {
        type: "text/plain;charset=utf-8"
    });

    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = `ohya_checkin_records_${getFileTime()}.txt`;
    a.click();

    URL.revokeObjectURL(url);
}

function updateRecordCount() {
    recordCount.textContent = checkinRecords.length;
}

function showVerifyWarning(title, message) {
    verifyResult.innerHTML = `
        <div class="warning">
            <h2>⚠️ ${title}</h2>
            <p>${message}</p>
        </div>
    `;
}

function showDrawWarning(title, message) {
    drawResult.innerHTML = `
        <div class="warning">
            <h2>⚠️ ${title}</h2>
            <p>${message}</p>
        </div>
    `;
}

function formatDateTime(date) {
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
}

function getFileTime() {
    const now = new Date();
    return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
}

function pad(num) {
    return String(num).padStart(2, "0");
}

searchInput.addEventListener("keypress", (e) => {
    if (e.key === "Enter") {
        searchBtn.click();
    }
});