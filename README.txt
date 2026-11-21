ChiliScan V7 - 光譜資料分析平台

這個程式是什麼
----------------
ChiliScan 是一個純前端的光譜資料分析工具。使用者匯入辣椒或其他物件的 Excel / CSV
量測資料後，程式會自動配對工作表中的波長欄與光強欄，繪製光譜曲線，並計算指定波段的
波峰、平均、最低強度與有效量測點數。

一般分析都在瀏覽器中執行，不會上傳原始量測資料。只有使用者設定安全 API 並按下 Kumo AI 判斷時，才會傳送勾選的待測光譜與未勾選且已標籤的參考光譜。

吸光度新鮮／腐爛實驗判斷
----------------------
首頁頂端的「新鮮／腐爛判斷」會開啟 `classification.html`。此頁在瀏覽器本機以吸光度資料比對，不連接 Kumo API，也不會上傳光譜。

1. 首次使用時載入含「新鮮」與「腐爛」兩個工作表的吸光度參考檔；參考資料會存入目前瀏覽器的 localStorage。清除瀏覽器網站資料或改用其他瀏覽器時，需重新載入參考檔。
2. 待測工作表需有標題含「波長」的欄，波長數值只比對 620–680 nm。樣本可使用一欄已平均的光譜，或在樣本名稱下放多組「波長／光強」重複量測欄，程式會先依波長平均光強。只有數值、沒有波長欄的矩陣無法判斷每列的波長。資料可直接提供吸光度；若提供原始光強，還需在同一工作表包含「標準白」欄，程式會計算 `A = -log10(樣本平均光強 / 標準白光強度)`。
3. 只比對 620–680 nm，並只使用參考與待測都有效的共同波長。待測檔每個樣本至少需有 40 個共同有效點。
4. 判斷使用吸光度 RMSE 的 3 個最近參考樣本；每筆結果會顯示新鮮／腐爛的近鄰投票比例，這是參考樣本的票數占比，不是該筆樣本的準確率。頁面也顯示參考資料的留一交叉驗證準確率。以目前 20 筆新鮮、6 筆腐爛參考資料測試，整體為 22/26（84.6%），新鮮辨識率 19/20（95%）、腐爛辨識率 3/6（50%）；資料仍不足，特別容易漏判腐爛，結果僅供實驗參考，不可單獨作為食品安全依據。

執行流程
--------
1. 使用者選擇 Excel / CSV 或拖曳檔案。
2. load() 讀取檔案：
   - CSV：由 parseCsv() 轉成二維陣列。
   - Excel：由 SheetJS 讀取工作表。
3. parseSheet() 或 readSheetSeries() 找出標題列、波長欄及樣本欄。
4. analyze() 或 analyzeCross() 依指定波段建立系列曲線資料。
   - replicateMode 可選「每筆量測」、「自動平均」或「每筆＋平均」。
   - 一般格式可用相同樣本名稱加 1、2、3 標示重複量測。
   - 若每個品項有多組相鄰的「波長＋光強」欄位，程式會依品項標題分組；例如「糯米辣椒一」標題下的三組量測只互相平均，不會與其他品項混合。
   - 平均依相同波長逐點計算；分析狀態會回報偵測到的品項組數與量測筆數。
   - 匯入資料及修改工作表、波長欄、波段或平均方式後，結果會自動更新。
5. feature() 計算每個樣本的波段特徵。
6. render() 更新摘要卡、特徵表、原始數據表與圖例。
7. draw() 使用 Canvas 繪製座標軸、刻度、曲線與滑鼠游標提示。
8. 匯出按鈕呼叫 downloadCsv() 建立並下載 CSV。

專案結構導覽
------------
VS Code 左側檔案總管顯示的是專案原始碼與設定。各資料夾和主要檔案用途如下：

- `.github/workflows/`
  - `pages.yml`：GitHub Actions 部署設定。將網站發布到 GitHub Pages；不會啟動 Kumo 後端。

- `.vscode/`
  - `launch.json`：VS Code 的啟動設定，提供本機預覽入口。
  - `tasks.json`：VS Code 工作清單，透過 `server.ps1` 啟動本機預覽。

- `大腦/`
  - `kumo_pepper_predict.py`：選用的 Python 命令列實驗程式，讀取同資料夾 CSV 後呼叫 NVIDIA Kumo；與網站頁面及 Jetson API 分開執行。
  - `reference.csv`：命令列程式使用的已標記參考光譜。
  - `new_samples.csv`：命令列程式要送去分類的待測光譜。這兩個 CSV 被 `.gitignore` 排除，不會提交至 Git。

- `css/`
  - `style.css`：首頁的共用外觀、版面、表格與響應式樣式。
  - `classifier.css`：新鮮／腐爛判斷頁的上傳區、結果卡片與比例顯示樣式。

- `jetson/`
  - `jetson_api.py`：安全後端 API，負責健康檢查、權限驗證、資料檢查，再把分類請求轉送至 NVIDIA Kumo。金鑰由伺服器環境變數提供。

- `js/`
  - `app.js`：首頁主要功能，包括讀取與分析光譜、繪圖、匯出，以及使用者啟動時的 Kumo API 連線與分類。
  - `classifier.js`：新鮮／腐爛頁的吸光度處理、3 近鄰判斷、結果呈現及中英文切換；語言選擇會與首頁共用並記住。
  - `language.js`：未由目前頁面載入的舊翻譯工具；首頁由 `app.js` 切換語言，判斷頁由 `classifier.js` 切換語言。
  - `xlsx.full.min.js`：SheetJS 函式庫，用來讀取 Excel 檔案；屬於第三方套件，不是 ChiliScan 的業務邏輯。

- 根目錄主要檔案
  - `index.html`：光譜分析首頁的頁面結構。
  - `classification.html`：新鮮／腐爛吸光度判斷頁的頁面結構，包含中英文切換按鈕。
  - `README.txt`：本專案的功能、操作方式與部署說明。
  - `render.yaml`：Render 的 Kumo 後端部署設定；與 GitHub Pages 的靜態網站部署分開。
  - `server.ps1`：本機 HTTP 預覽伺服器，預設網址為 `http://127.0.0.1:5500/`。
  - `.gitignore`：指定不提交至 Git 的檔案；目前排除 `大腦/reference.csv` 與 `大腦/new_samples.csv`。

JavaScript主要區塊
-----------------
1. 共用狀態與工具函式
   workbook、rows、headers、series、selected等資料，以及中英文、數值格式與HTML防護。

2. Excel / CSV讀取與工作表解析
   parseCsv()、parseSheet()、readSheetSeries()與sheetToRows()。

3. 單一工作表與跨工作表分析
   analyze()處理單一工作表；analyzeCross()處理多工作表疊加；
   兩者都會依指定起始／結束波長篩選數據。

4. 特徵、表格與圖表
   feature()計算統計值；render()更新頁面；renderRawData()建立逐波長表格；
   draw()繪製座標軸、曲線、平滑值與游標提示。

5. 檔案匯入與示範資料
   load()讀取檔案；demo按鈕載入模擬數據。
   示範數據不代表實際量測結果。

6. CSV匯出
   特徵CSV包含波峰與統計值。
   原始數據CSV包含目前勾選樣本在指定波段內的逐波長數據。

7. NVIDIA Kumo AI
   將未勾選且指定品質標籤的光譜作為參考資料，勾選光譜作為待測資料；明確按下操作後才會送至安全後端代理。

重要資料變數
------------
- workbook：目前載入的 Excel / CSV 工作簿物件。
- headers：目前工作表的欄位名稱。
- rows：目前工作表的數據列。
- widx：使用者選擇的主要波長欄。
- series：分析後產生的曲線資料，每筆包含名稱與波長／強度配對。
- selected：目前勾選要繪製的series索引。
- crossCache：跨工作表解析結果快取。
- hover：圖表游標位置，用於顯示附近數值。

NVIDIA Kumo AI 連線
-------------------
`jetson/jetson_api.py` 是安全後端代理，可部署在 Jetson 或其他受控伺服器。它呼叫 NVIDIA Kumo-Tabular；不是在本機訓練模型，也不代表分類準確度已驗證。

1. 在安全後端設定 `CHILISCAN_TOKEN`、`NVIDIA_API_KEY`、`NVIDIA_KUMO_PREDICT_URL`。Kumo Catalog 預測網址為 `https://ai.api.nvidia.com/v1/structured-data/nvidia/kumo-tabular/predictions`。NVIDIA 金鑰只能保存在後端環境變數，絕不可貼入網頁、提交 Git 或放在公開靜態網站。
2. 設定 `CHILISCAN_ORIGIN` 為 ChiliScan 網站來源。服務預設綁定 `127.0.0.1`；部署到網際網路時必須使用受信任的 HTTPS 反向代理，不要直接暴露未加密服務。
3. 網頁輸入後端 HTTPS 網址及 `CHILISCAN_TOKEN`，測試連線。網址與權杖會儲存在目前瀏覽器的 localStorage，之後開啟頁面會自動測試連線；清除網站資料或換瀏覽器／裝置後需重新輸入。權杖只適合保存在個人可信任的瀏覽器，避免在共用裝置使用。
4. 匯入已知樣本資料，在樣本清單將未勾選的參考樣本指定品質標籤，按「保存未勾選且已標記樣本為參考資料」建立本機參考庫。參考光譜及標籤只存在瀏覽器 localStorage；刪除網站資料會一併刪除參考庫。
5. 後續匯入待測光譜時，勾選要預測的樣本；已保存的參考庫會自動作為 Kumo context。每類至少要有一筆參考光譜，且待測光譜的波長位置須與參考庫完全一致。按分類時，參考與待測光譜會送到安全後端及 NVIDIA API。
6. 確認資料已去識別且可分享後，再按 Kumo 分類。只有該操作會將參考與待測光譜傳送到你設定的後端及 NVIDIA API。

GitHub Pages 只託管靜態網站，不會執行 `jetson/jetson_api.py`。網站上的 Kumo 分類功能要等你另外部署並設定安全 HTTPS 後端後才能使用。

Render 後端部署
---------------
根目錄的 `render.yaml` 已準備好 Render Blueprint，使用免費方案、新加坡區域、HTTPS，並由 Render 指派的 `PORT` 啟動 API。

1. 登入 Render，選擇建立 Blueprint，連結 `rgwvb/ChiliScan_V7_VSCode` 儲存庫並套用 `render.yaml`。
2. 在 Render 的秘密環境變數設定 `NVIDIA_API_KEY`（已輪替的新金鑰）及 `CHILISCAN_TOKEN`（自訂的隨機長權杖）。這兩個值只輸入 Render 控制台，不要放進 YAML、GitHub 或對話。其餘 API 網址與網站來源已在 Blueprint 設定。
3. 部署成功後複製 Render 提供的 `https://....onrender.com` 網址；在 ChiliScan 的 Kumo 區塊輸入該 API 網址，並輸入同一個 `CHILISCAN_TOKEN`，再測試連線。

Render 帳號登入、設定秘密值及第一次建立服務需要由帳號持有人在 Render 控制台完成；本機 GitHub Pages 部署不會自動建立 Render 服務。免費服務閒置後可能需要等待喚醒。按下分類時，光譜會經過 Render 後端送往 NVIDIA。

健康檢查 `GET /api/v1/health` 與分類 `POST /api/v1/predict` 都需要 `Authorization: Bearer <CHILISCAN_TOKEN>`。分類請求使用 `api_version: 1`、`context_spectra` 與 `spectra`：每筆光譜包含 `sample_name`、`wavelengths_nm`、`intensities`、`replicate_count`；context 另需 `label`（`fresh`、`pest` 或 `rotten`）。回應包含 `api_version: 1` 和 `predictions` 陣列，每筆結果包含 `sample_name`、`label` 與 0 到 1 之間的 `confidence`。

目前可用的實驗樣本不足以證明三分類效能。重複量測不是新的獨立樣本；請增加各品質類別的獨立樣本並進行獨立驗證，不要將試用分類結果視為已驗證診斷。

命令列實驗腳本
--------------
`大腦/kumo_pepper_predict.py` 需要同資料夾中的 `reference.csv`（含 `sample_id`、波長特徵欄及 `quality`）與 `new_samples.csv`（含 `sample_id` 及相同順序的波長特徵欄）。已從指定活頁簿的「大數據」工作表產生本機 CSV：8 筆參考資料（6 新鮮、1 蟲害、1 腐爛）及 1 筆新鮮待測資料「辣椒三」，原始資料每筆含 508 個波長特徵。Kumo 每表最多 500 欄（包含樣本 ID 與標籤），命令列腳本與網站後端都會均勻取樣 498 個波長特徵送出。這只是功能試跑，不能代表分類準確度。CSV 已列入 `.gitignore`，不會提交到 Git。

PowerShell 從專案根目錄執行時，切換到 `大腦` 資料夾並啟動腳本：

    cd .\大腦
    python .\kumo_pepper_predict.py

腳本執行後會在終端機提示輸入已輪替的新金鑰，輸入時不會顯示；按 Enter 後會直接呼叫 NVIDIA，不會把金鑰寫入環境變數、指令歷史或檔案。請勿把真實金鑰貼到對話或截圖。若 `python` 指令無法使用，請在 VS Code 選取 Python 解譯器，再使用「Run Python File in Terminal」。

本機預覽
--------
在 VS Code 執行「啟動 ChiliScan 預覽」，或執行終端機中的預覽工作。
請使用 http://127.0.0.1:5500/。
不要直接以 file:///index.html 開啟網頁。
