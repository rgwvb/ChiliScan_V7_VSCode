ChiliScan V7 - 光譜資料分析平台

這個程式是什麼
----------------
ChiliScan 是一個純前端的光譜資料分析工具。使用者匯入辣椒或其他物件的 Excel / CSV
量測資料後，程式會自動配對工作表中的波長欄與光強欄，繪製光譜曲線，並計算指定波段的
波峰、平均、最低強度與有效量測點數。

所有分析都在瀏覽器中執行，不會把原始量測資料上傳到伺服器。

執行流程
--------
1. 使用者選擇 Excel / CSV 或拖曳檔案。
2. load() 讀取檔案：
   - CSV：由 parseCsv() 轉成二維陣列。
   - Excel：由 SheetJS 讀取工作表。
3. parseSheet() 或 readSheetSeries() 找出標題列、波長欄及樣本欄。
4. analyze() 或 analyzeCross() 依指定波段建立系列曲線資料。
   - replicateMode 可選「每筆量測」、「自動平均」或「每筆＋平均」。
   - 同一樣本的三次量測放在同一工作表不同欄，欄名用共同名稱加 1、2、3，例如「辣椒1、辣椒2、辣椒3」。
   - 預設逐波長自動平均；分析狀態會回報偵測到的量測組數與筆數。
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
  特徵計算、Canvas繪圖、原始數據分頁及CSV匯出。

- js/xlsx.full.min.js
  SheetJS Excel解析套件，用於讀取.xlsx與.xls。

- server.ps1
  啟動本機HTTP預覽伺服器，預設位址為 http://127.0.0.1:5500/

- .github/workflows/pages.yml
  使用GitHub Actions自動部署到GitHub Pages。

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

本機預覽
--------
在 VS Code 執行「啟動 ChiliScan 預覽」，或執行終端機中的預覽工作。
請使用 http://127.0.0.1:5500/。
不要直接以 file:///index.html 開啟網頁。
