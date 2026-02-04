復興高中選班群系統 (Docker 部署版)
這是一個基於 Flask 開發的極簡風選班群系統，支援自動存檔至 Excel 並產生個人化 Word 確認表單。透過 Docker 技術，本系統可以快速部署於任何電腦。

📂 核心檔案清單
搬移系統至新電腦時，請確保以下檔案存在於同一個資料夾中：

app.py：系統邏輯與介面 (內含 CSS)。

Dockerfile：環境映像檔定義。

docker-compose.yml：容器啟動與資料掛載配置。

students.xlsx：學生基本資料庫（需包含姓名與身分證字號欄位）。

template.docx：學生下載用的 Word 列印範本。

🚀 新電腦部署步驟
1. 環境準備
在新電腦上安裝並啟動以下軟體：

Docker Desktop: 官方下載連結

2. 啟動系統
開啟終端機 (PowerShell 或 CMD)，進入專案資料夾後執行：

PowerShell
# 建立映像檔並在背景啟動容器
docker-compose up -d --build
3. 開始使用
本地測試：在瀏覽器輸入 http://localhost:5000

對外連線：查詢該主機的 IP (如 192.168.x.x)，其他電腦輸入 http://192.168.x.x:5000 即可登入。

📊 資料管理說明
學生選填結果
系統會自動在資料夾下產生 selections.xlsx。

由於我們在 docker-compose.yml 中設定了 volumes，該檔案會即時同步出現在你的主機資料夾內。

注意：請勿在系統執行期間用 Excel 軟體「鎖定」開啟該檔案，以免寫入失敗。

更新學生名單
若要更新 students.xlsx，只需覆蓋舊檔並執行 docker-compose restart 即可生效。

🛠️ 常用維護指令
查看運行狀態：docker ps

查看即時錯誤日誌：docker-compose logs -f

停止系統：docker-compose down

重新啟動：docker-compose restart

🎨 風格調整
本系統採用 黑白灰極簡風格。若需修改介面顏色或文字，請編輯 app.py 中的 STYLE 變數與 HTML 字串，修改後需執行 docker-compose up -d --build 重新編譯。
