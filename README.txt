ChiliScan V7 - VS Code 專案

資料夾結構：
- index.html：網頁內容
- css/style.css：頁面樣式
- js/app.js：光譜資料分析與圖表程式

使用方式：
1. 解壓縮後，在 VS Code 選擇「檔案 > 開啟資料夾」，開啟 ChiliScan_V7_VSCode。
2. 編輯 index.html、css/style.css 或 js/app.js。
3. 在 VS Code 按 F5，選擇「啟動 ChiliScan 預覽」，使用固定預覽位址：http://127.0.0.1:5500/
4. 若未自動啟動伺服器，先執行「終端機 > 執行工作 > 預覽 ChiliScan」。
5. 不要直接用瀏覽器開啟舊的 file:///index.html；請使用上面的 HTTP 預覽位址。
6. 匯入 Excel 時需要網路連線以載入 SheetJS；CSV 可在離線時直接匯入。

程式用途與作用：
- ChiliScan 是純前端光譜資料分析工具，用於匯入辣椒或其他物件的 Excel / CSV 光譜量測資料。
- 自動辨識「波長」欄與其左側相鄰的樣本光強欄。
- 支援單一工作表分析及跨工作表疊加比較，可同時比較新鮮、蟲害、腐爛等平均資料。
- 可設定分析波段、X/Y 軸刻度、曲線平滑方式；曲線預設採 5 點移動平均，僅影響顯示，不會更動原始數據或統計結果。
- 預設分析 620–680 nm；「設定 620–680 nm」可快速套用，起始與結束波長也可自行修改。
- 自動 Y 軸在數據皆為非負時從 0 開始；數據含負值時仍保留負值範圍。
- 統計指定波段中的波峰波長、波峰強度、平均強度、最低強度及有效量測點數。
- 在頁面下方以逐波長表格呈現目前分析範圍內所有樣本的強度數據，每頁最多100個波長點。
- 可匯出特徵摘要 CSV，或將目前波段（例如 620–680 nm）與勾選樣本的原始數據整理成 CSV。
- 所有資料都在瀏覽器中處理，不會上傳至伺服器。

程式區塊：
- index.html：頁面結構、表單、統計卡、圖表、樣本清單及資料表。
- css/style.css：全域樣式、控制元件、圖表區、跨工作表選擇與手機版響應式版面。
- js/app.js：共用工具、檔案讀取、單一/跨工作表分析、特徵計算、Canvas繪圖、波段原始數據表及CSV匯出。
- server.ps1：本機 HTTP 預覽伺服器。
- .github/workflows/pages.yml：GitHub Actions 自動部署到 GitHub Pages。

公開上架（GitHub Pages）：
1. 在 GitHub 建立公開儲存庫，並將本資料夾中的網站檔案推送至 main 分支。
2. 在儲存庫 Settings > Pages，將 Build and deployment 的 Source 設為 GitHub Actions。
3. 推送後，GitHub Actions 會自動部署；完成後網址會顯示在 Actions 的部署紀錄及 Settings > Pages。
4. 此網站是純前端靜態頁面；Excel 分析仍在訪客的瀏覽器中執行，不會上傳使用者資料。
