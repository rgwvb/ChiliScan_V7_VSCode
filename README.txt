ChiliScan V7 - 光譜資料分析平台

這個程式是什麼
----------------
ChiliScan 是一個純前端的光譜資料分析工具。使用者匯入辣椒或其他物件的 Excel / CSV
量測資料後，程式會自動配對工作表中的波長欄與光強欄，繪製光譜曲線，並計算指定波段的
波峰、平均、最低強度與有效量測點數。

一般分析都在瀏覽器中執行，不會上傳原始量測資料。只有使用者設定安全 API 並按下 Kumo AI 判斷時，才會傳送勾選的待測光譜與未勾選且已標籤的參考光譜。

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

檔案分工
--------
- index.html
  頁面結構與介面元素，例如上傳區、分析控制、統計卡、圖表、樣本清單與資料表。

- css/style.css
  深色介面、版面配置、表格、輸入控制、跨工作表卡片及手機版響應式樣式。

- js/app.js
  檔案讀取、語言切換、Excel / CSV解析、單一工作表與跨工作表分析、
  特徵計算、Canvas繪圖、原始數據分頁及CSV匯出；另負責安全 API 連線、送出 Kumo 參考／待測光譜與顯示分類回應。

- js/xlsx.full.min.js
  SheetJS Excel解析套件，用於讀取.xlsx與.xls。

- server.ps1
  啟動本機HTTP預覽伺服器，預設位址為 http://127.0.0.1:5500/

- .github/workflows/pages.yml
  使用GitHub Actions自動部署到GitHub Pages。

- jetson/jetson_api.py
  安全後端代理，提供健康檢查及 Kumo 光譜分類端點；ChiliScan 權杖與 NVIDIA 金鑰都由伺服器環境變數提供，不會傳給瀏覽器。

- 大腦/kumo_pepper_predict.py
  可選的命令列實驗工具，從同資料夾的 reference.csv 與 new_samples.csv 直接呼叫 Kumo；它不會啟動網站 API，也不會使用瀏覽器內的分類介面。

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
3. 網頁輸入後端 HTTPS 網址及 `CHILISCAN_TOKEN`，測試連線。此權杖與 NVIDIA 金鑰不同，只暫存在頁面記憶體。
4. 在樣本清單中將已知品質標籤指定給未勾選的參考樣本，並勾選要預測的樣本。每類至少要有一筆參考光譜；所有光譜的波長位置必須一致。
5. 確認資料已去識別且可分享後，再按 Kumo 分類。只有該操作會將參考與待測光譜傳送到你設定的後端及 NVIDIA API。

GitHub Pages 只託管靜態網站，不會執行 `jetson/jetson_api.py`。網站上的 Kumo 分類功能要等你另外部署並設定安全 HTTPS 後端後才能使用。

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
