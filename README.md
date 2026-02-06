復興高中選班群系統
這是一個專為高中選修班群設計的 Flask 系統。除了提供學生選填介面，新版更加入了即時統計後台，能自動比對全校名單，快速抓出尚未繳交的學生，並支援匯出 Excel 報表。

🌟 核心功能
學生登入驗證：對接 students.xlsx 進行學號與身分證後六碼驗證。

班群選填與存檔：學生選擇後自動寫入 selections.xlsx，支援重複修改。

個人表單下載：自動將學生的選擇帶入 template.docx，產生符合排版格式的 PDF/Word 列印表單。

全校進度統計 (New!)：

網頁監控面板：/admin 頁面即時顯示「已繳交」與「未繳交」名單。

比對功能：系統自動拿全校名單進行 Left Join，未填寫者以紅字標示。

記憶體匯出技術：下載 Excel 時不佔用硬碟檔案，有效防止「Permission denied」權限鎖定問題。

📂 核心檔案清單
app.py：系統主程式（內含統計與下載邏輯）。

students.xlsx：【必備】 原始學生資料庫（需包含：班級、座號、學號、姓名、身份證號後6碼）。

selections.xlsx：系統自動產生，記錄所有學生的選填結果。

template.docx：Word 列印範本，內含 {{name}} 等變數標籤。

docker-compose.yml & Dockerfile：用於快速環境建置。

📊 管理者功能使用說明
1. 進入統計後台
在瀏覽器輸入以下網址（假設在本地運行）：

http://localhost:5000/admin

畫面會顯示：

總人數統計：總計、已完成、待補交人數。

即時名單表格：直接看到全校每位學生的選填狀況，沒填的人會標註為紅色的「尚未填寫」。

2. 匯出 Excel 統計報表
在管理頁面點擊 「下載完整 Excel 報表」：

系統會生成一份名為 全校選課統計_日期時間.xlsx 的檔案。

此檔案包含全校學生的基本資料，並額外增加一欄「填寫狀態」，方便老師直接依據狀態進行篩選與催繳。

🚀 部署指令
啟動系統：

Bash
docker-compose up -d --build
更新名單後重啟：

Bash
docker-compose restart
停止並移除容器：

Bash
docker-compose down
