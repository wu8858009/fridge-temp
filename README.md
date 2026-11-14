# 冰箱溫度記錄表系統

冷藏／冷凍／熱藏設備的每日溫度記錄（早、晚等自訂時段）、超溫警示與異常處置、A4 月報表（列印／PDF／Excel）、溫度趨勢圖。
同一份網頁程式（`web/`）同時是手機 PWA 與 Windows 桌面版（`desktop/main.py`，pywebview）。

## 功能
- **每日記錄**：每台設備每個時段輸入溫度，自動記下記錄人與時間；超出標準（預設冷藏 ≤7°C、冷凍 ≤-18°C、熱藏 ≥60°C）會標紅並要求填寫處置與複測溫度。到了時段還沒記錄會出現提醒（導覽列紅點）。
- **月報表**：每台設備一頁 A4（日期 × 時段 溫度／記錄人、異常處置、主管確認欄、簽名欄），可列印、存 PDF、匯出 Excel（含「異常處置紀錄」分頁）。
- **溫度趨勢**：每台設備一張折線圖，綠色為合格範圍、紅色虛線為上下限、紅點為異常；平均／最低／最高／異常次數／完成率。
- **設定**：店名、設備（類型、上下限、停用、排序）、記錄時段、記錄人名單、Google 雲端同步、JSON 備份／還原、示範資料。

## Google 雲端同步
- `cloud/Code.gs` 綁定在試算表「冰箱溫度記錄表 雲端資料」（分頁：設備、溫度紀錄、設定）。試算表上手動修改也會同步回 App（`onEdit` 會重新判斷正常／異常）。
- 第一次部署後，擁有者要在指令碼編輯器執行一次 `setup` 完成授權。
- 第一台連線的裝置輸入的密碼會成為雲端密碼；忘記密碼可在編輯器執行 `resetPassword`。
- 更新後端：`cd cloud && clasp push --force && clasp update-deployment <部署ID>`

## 開發與測試
- 網頁：`cd web && python -m http.server 8797`
- 模擬雲端：`node tools/gasmock.js 8796 state.json`（在本機執行真正的 Code.gs），App 雲端網址填 `http://127.0.0.1:8796/`；`127.0.0.1` 與 `localhost` 兩個分頁可當作兩台裝置。
- 打包：`python build.py` → `dist\冰箱溫度記錄表.exe`（程式開著時可用 `python build.py <其他資料夾>`）
- 桌面版自我測試：`FRIDGETEMP_DATA=<資料夾> FRIDGETEMP_SELFTEST=<結果.json> FRIDGETEMP_EXPORT_DIR=<資料夾>` 執行 exe。
- 修改 `web/` 後：index.html 與 sw.js 的 `?v=` 和 `CACHE` 一起加 1，再重新打包。
