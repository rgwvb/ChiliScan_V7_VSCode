/*
 * ChiliScan 光譜資料分析平台
 * =========================================
 * 用途：
 * 1. 將瀏覽器匯入的 Excel / CSV 光譜資料轉成可分析的數值。
 * 2. 自動辨識波長欄與樣本欄，支援單一工作表及跨工作表疊加。
 * 3. 計算指定波段中的波峰、平均強度、最低強度與有效量測點數。
 * 4. 使用 Canvas 繪製光譜曲線，並在頁面表格或 CSV 顯示波段數據。
 *
 * 程式流程：
 * 匯入檔案 -> 辨識工作表欄位 -> 篩選波段 -> 選取樣本 -> 計算特徵 -> 繪圖/匯出
 *
 * 注意：
 * 一般分析在訪客瀏覽器中處理；只有使用者明確呼叫 Jetson AI 時才傳送所選光譜。
 */

// =====================================================
// 區塊一：共用狀態與工具函式
// =====================================================
// 這裡負責：
// 1. 保存使用中的 Excel / CSV、工作表、樣本與選取狀態。
// 2. 提供中英文切換、數值轉換、HTML跳脈防護與格式化工具。
// 3. 不直接處理 Excel/CSV讀取或畫圖；這些工作交給後面的區塊。
const $=id=>document.getElementById(id);let crossSpecs=[],crossCache=new Map();let workbook=null,rows=[],headers=[],sourceHeaders=[],widx=-1,series=[],selected=new Set(),hover=null,rawPage=0,demoDataLoaded=false,csvImported=false,replicateSummary={groups:0,measurements:0};let jetsonConnection=null,jetsonPhase='idle',jetsonDetail='',jetsonPredictions=[];const rawPageSize=100;const palette=['#60a5fa','#34d399','#fbbf24','#fb7185','#a78bfa','#22d3ee','#f472b6','#bef264','#f97316','#93c5fd'];
const staticTranslations={
'pageTitle':'ChiliScan | Spectral Data Analysis Platform V7',
'brandSuffix':'/ Spectral Analysis Platform',
'intro':'Import Excel / CSV spectral data, automatically identify wavelength columns, plot sample curves, and calculate spectral features for a specified band. All analysis is performed in your browser; no data is uploaded to a server.',
'01　匯入光譜資料':'01  Import Spectral Data',
'點擊選擇 Excel / CSV，或將檔案拖曳到這裡':'Click to select Excel / CSV, or drag a file here',
'支援 .xlsx、.xls、.csv；匯入後可切換工作表；優先開啟「新鮮平均」，否則選擇「光強」；支援跨工作表光譜疊加與多組波長欄；預設分析 620–680 nm，曲線可切換平滑／原始顯示':'Supports .xlsx, .xls, and .csv; switch worksheets after import; prefers "Fresh Average", otherwise "Intensity"; supports cross-sheet overlays and multiple wavelength columns; default band is 620–680 nm; curves can be smoothed or displayed as raw data',
'工作表（優先：新鮮平均／光強）':'Worksheet (preferred: Fresh Average / Intensity)',
'尚未匯入':'Not imported',
'波長欄':'Wavelength Column',
'自動辨識':'Auto-detect',
'起始波長（nm）':'Start wavelength (nm)',
'結束波長（nm）':'End wavelength (nm)',
'X軸刻度間距（nm）':'X-axis tick spacing (nm)',
'設定 620–680 nm':'Set 620–680 nm',
'重新分析':'Re-analyze',
'載入示範資料':'Load Demo Data',
'等待資料匯入。':'Waiting for data import.',
'02　跨工作表光譜疊加比較':'02  Cross-Worksheet Overlay Comparison',
'（可同時選擇新鮮平均、蟲害平均、腐爛平均）':'(Select Fresh, Pest-Damaged, and Rotten averages together)',
'分析模式':'Analysis Mode',
'單一工作表':'Single Worksheet',
'跨工作表疊加比較':'Cross-Worksheet Overlay',
'選取三種品質平均工作表':'Select Three Quality Average Worksheets',
'勾選要比較的工作表，再於每張工作表內選擇樣本欄位。每個樣本會使用其左側最近的「波長」欄，不會按照資料列號強制配對。工作表名稱會加入圖例。':'Check worksheets to compare, then select sample columns in each worksheet. Each sample uses the nearest wavelength column to its left; rows are not matched by row number. Worksheet names are included in the legend.',
'請先匯入 Excel 檔案':'Import an Excel file first',
'更新疊加圖表':'Update Overlay Chart',
'勾選所有可用樣本':'Select All Available Samples',
'03　波段特徵摘要':'03  Band Feature Summary',
'（以目前勾選的第一個樣本為主）':'(Based on the first currently selected sample)',
'波峰波長':'Peak Wavelength',
'波峰強度':'Peak Intensity',
'區間平均強度':'Band Mean Intensity',
'有效量測點數':'Valid Measurement Points',
'點':'pts',
'04　光譜強度曲線':'04  Spectral Intensity Curve',
'縱軸刻度模式':'Y-axis Scale Mode',
'自動調整（非負數據從 0 開始）':'Auto (non-negative data starts at 0)',
'手動設定':'Manual',
'Y軸最小值':'Y-axis Minimum',
'Y軸最大值':'Y-axis Maximum',
'重設Y軸':'Reset Y-axis',
'例如 0':'e.g. 0',
'例如 4500':'e.g. 4500',
'曲線顯示方式':'Curve Display',
'平滑顯示（3點移動平均）':'Smoothed (3-point moving average)',
'平滑顯示（5點移動平均）':'Smoothed (5-point moving average)',
'原始數據（不平滑）':'Raw Data (no smoothing)',
'橫軸：實際波長（nm）；縱軸：原始訊號強度（未進行反射率校正）。平滑僅作用於曲線顯示，不影響原始資料、波峰與平均強度計算。滑鼠移至曲線區域可查看游標位置，以及各樣本最近實測波長與原始強度。':'X-axis: measured wavelength (nm); Y-axis: raw signal intensity (without reflectance correction). Smoothing affects only the curve display and not raw data or feature calculations. Hover over the chart to view the cursor position and each sample’s nearest measured wavelength and raw intensity.',
'樣本選擇':'Sample Selection',
'匯入資料後顯示樣本欄位':'Sample columns appear after import',
'勾選樣本即可疊加比較；按「單獨」可只查看該樣本。':'Check samples to overlay them; use "Only" to view a sample alone.',
'全部顯示／取消':'Show All / Clear',
'05　各樣本光譜特徵比較':'05  Spectral Feature Comparison',
'匯出特徵 CSV':'Export Features CSV',
'匯出選定波段原始數據 CSV':'Export Raw Band CSV',
'原始數據匯出會依上方起始／結束波長篩選，並只匯出目前勾選的樣本；設定 620–680 nm 後，結果僅包含此範圍內的量測點。':'Raw-data export uses the start/end wavelengths above and exports only currently checked samples. Setting 620–680 nm limits results to measurement points in that range.',
'樣本名稱（含工作表）':'Sample Name (with Worksheet)',
'波峰波長 (nm)':'Peak Wavelength (nm)',
'最低強度':'Minimum Intensity',
'量測點數':'Measurement Points',
'尚無分析資料':'No analysis data',
'特徵僅依指定波長範圍計算。若樣本未落入範圍，該樣本不列入分析。不同樣本量測條件不一致時，強度不可直接視為品質差異。':'Features are calculated only within the specified wavelength range. Samples outside the range are excluded. Intensities should not be treated directly as quality differences when measurement conditions differ.',
'06　分析範圍內的原始數據':'06  Raw Data in the Analysis Range',
'依目前起始與結束波長，顯示所有可用樣本的逐波長強度；沒有對應量測值時會留白。資料較多時可使用分頁瀏覽。':'Displays each available sample’s intensity by wavelength within the current start and end wavelengths. Missing measurements are left blank. Use pagination for large datasets.',
'波長 (nm)':'Wavelength (nm)',
'匯入資料後顯示波段數據':'Band data appears after import',
'上一頁':'Previous',
'下一頁':'Next',
'尚無數據':'No data',
'使用提示：支援範例.xlsx 的多組波長欄：每個樣本自動對應其左側最近的波長欄；重複樣本名稱會標示為第幾個同名欄位。若檔案只有圖片而非原始數值，請先提供 Excel / CSV 檔案。':'Tip: Multiple wavelength columns are supported. Each sample is paired with the nearest wavelength column to its left; duplicate names are numbered by occurrence. If a file contains only an image instead of numeric values, provide an Excel / CSV file.',
'沒有可繪製的樣本':'No drawable samples',
'沒有資料':'No data',
'單獨':'Only',
'請勾選至少一個樣本':'Check at least one sample'
};
let currentLanguage=localStorage.getItem('chiliscan-language')||'zh';
const uiTranslations={
  pageTitle:['ChiliScan｜光譜數據分析平台 V7','ChiliScan | Spectral Data Analysis Platform V7'],
  brandSubtitle:['／ 光譜分析平台','/ Spectral Analysis Platform'],
  intro:['匯入 Excel / CSV 光譜資料，自動辨識波長欄、繪製樣本曲線，計算指定波段的光譜特徵。一般分析在瀏覽器內進行；只有按下 Jetson AI 判斷時，勾選的光譜才會送到你設定的裝置。','Import Excel / CSV spectral data to identify wavelength columns, plot sample curves, and calculate spectral features in a selected band. Routine analysis stays in your browser; selected spectra are sent to your configured Jetson only when you request AI classification.'],
  importTitle:['01　匯入光譜資料','01  Import Spectral Data'],
  uploadTitle:['選擇 Excel / CSV，或將檔案拖曳到這裡','Choose an Excel / CSV file, or drag it here'],
  uploadHint:['支援 .xlsx、.xls、.csv。三次量測放在同一工作表不同欄，例如：辣椒1、辣椒2、辣椒3。','Supports .xlsx, .xls, and .csv. Put three measurements in separate columns on one worksheet, e.g. Pepper1, Pepper2, Pepper3.'],
  notImported:['尚未匯入','Not imported'],
  autoDetect:['自動辨識','Auto-detect'],
  sheetLabel:['工作表（優先：新鮮平均／光強）','Worksheet (preferred: Fresh Average / Intensity)'],
  wavelengthLabel:['波長欄','Wavelength Column'],
  startWavelength:['起始波長（nm）','Start Wavelength (nm)'],
  endWavelength:['結束波長（nm）','End Wavelength (nm)'],
  xStep:['X軸刻度間距（nm）','X-axis Tick Interval (nm)'],
  rangeSettings:['資料與波段','Data & Wavelength Range'],
  autoUpdateHint:['匯入後會自動分析；修改欄位或波段後會自動重新計算。','Analysis starts on import and recalculates when worksheet, wavelength, or range settings change.'],
  replicateSettings:['重複量測處理','Replicate Measurements'],
  replicateHint:['同一品項的重複量測請使用相同名稱或放在同一組欄位標題下；程式會依相同波長逐點平均。','Use the same sample name or place repeated measurements under one group heading; values are averaged at matching wavelengths.'],
  setBand:['設定 620–680 nm','Set 620–680 nm'],
  reanalyze:['重新分析','Re-analyze'],
  demo:['載入示範資料','Load Demo Data'],
  waiting:['等待資料匯入。','Waiting for data import.'],
  crossTitle:['02　跨工作表光譜疊加比較','02  Cross-Worksheet Spectral Overlay'],
  crossSubtitle:['（可同時選擇新鮮平均、蟲害平均、腐爛平均）','(Compare Fresh, Pest-Damaged, and Rotten averages together)'],
  analysisMode:['分析模式','Analysis Mode'],
  singleSheet:['單一工作表','Single Worksheet'],
  crossSheet:['跨工作表疊加比較','Cross-Worksheet Overlay'],
  selectThree:['選取三種品質平均工作表','Select the Three Quality-Average Worksheets'],
  crossHint:['勾選要比較的工作表，再於每張工作表內選擇樣本欄位。每個樣本會使用其左側最近的「波長」欄，不會按照資料列號強制配對。工作表名稱會加入圖例。','Select worksheets to compare and choose sample columns in each. Each sample uses the nearest wavelength column to its left; rows are not paired by row number. Worksheet names appear in the legend.'],
  importFirst:['請先匯入 Excel 檔案','Import an Excel file first'],
  updateChart:['更新疊加圖表','Update Overlay Chart'],
  selectAll:['勾選所有可用樣本','Select All Available Samples'],
  summaryTitle:['03　波段特徵摘要','03  Band Feature Summary'],
  summarySubtitle:['（以目前勾選的第一個樣本為主）','(Based on the first selected sample)'],
  peakWavelength:['波峰波長','Peak Wavelength'],
  peakIntensity:['波峰強度','Peak Intensity'],
  meanIntensity:['區間平均強度','Mean Intensity'],
  validPoints:['有效量測點數','Valid Measurement Points'],
  points:['點','points'],
  curveTitle:['04　光譜強度曲線','04  Spectral Intensity Curve'],
  yMode:['縱軸刻度模式','Y-axis Scale Mode'],
  yAuto:['自動調整（非負數據從 0 開始）','Auto (start at 0 for non-negative data)'],
  yManual:['手動設定','Manual'],
  yMin:['Y軸最小值','Y-axis Minimum'],
  yMax:['Y軸最大值','Y-axis Maximum'],
  yMinPlaceholder:['例如 0','e.g. 0'],
  yMaxPlaceholder:['例如 4500','e.g. 4500'],
  replicateModeLabel:['平均方式','Averaging Method'],
  replicateIndividual:['只顯示每筆量測','Show Individual Measurements'],
  replicateAverage:['只顯示自動平均','Show Automatic Average'],
  replicateBoth:['顯示每筆＋平均','Show Individuals + Average'],
  resetY:['重設Y軸','Reset Y-axis'],
  curveDisplay:['曲線顯示方式','Curve Display'],
  smooth3:['平滑顯示（3點移動平均）','Smoothed (3-point moving average)'],
  smooth5:['平滑顯示（5點移動平均）','Smoothed (5-point moving average)'],
  rawCurve:['原始數據（不平滑）','Raw Data (No Smoothing)'],
  chartHint:['橫軸：實際波長（nm）；縱軸：原始訊號強度（未進行反射率校正）。平滑僅作用於曲線顯示，不影響原始資料、波峰與平均強度計算。滑鼠移至曲線區域可查看游標位置，以及各樣本最近實測波長與原始強度。','X-axis: measured wavelength (nm); Y-axis: raw signal intensity (without reflectance correction). Smoothing affects display only, not the raw data or calculations. Hover over the chart to see the cursor position and each sample’s nearest measured wavelength and raw intensity.'],
  sampleTitle:['樣本選擇','Sample Selection'],
  showAfterImport:['匯入資料後顯示樣本欄位','Sample columns appear after import'],
  sampleHint:['勾選樣本即可疊加比較；按「單獨」可只查看該樣本。','Check samples to overlay them; click "Only" to view a single sample.'],
  toggleAll:['全部顯示／取消','Show All / Clear'],
  featureTableTitle:['05　各樣本光譜特徵比較','05  Spectral Feature Comparison'],
  exportFeatures:['匯出特徵 CSV','Export Features CSV'],
  exportRaw:['匯出選定波段原始數據 CSV','Export Selected-Band Raw Data CSV'],
  exportHint:['原始數據匯出會依上方起始／結束波長篩選，並只匯出目前勾選的樣本；設定 620–680 nm 後，結果僅包含此範圍內的量測點。','Raw-data export uses the start/end wavelengths above and includes only selected samples. Set 620–680 nm to export measurements in that range only.'],
  tableSample:['樣本名稱（含工作表）','Sample Name (with Worksheet)'],
  tablePeakW:['波峰波長 (nm)','Peak Wavelength (nm)'],
  tablePeakI:['波峰強度','Peak Intensity'],
  tableMean:['平均強度','Mean Intensity'],
  tableMin:['最低強度','Minimum Intensity'],
  tablePoints:['量測點數','Measurement Points'],
  noAnalysis:['尚無分析資料','No analysis data'],
  featureHint:['特徵僅依指定波長範圍計算。若樣本未落入範圍，該樣本不列入分析。不同樣本量測條件不一致時，強度不可直接視為品質差異。','Features are calculated only in the selected wavelength range; samples outside it are excluded. Intensity should not be interpreted as a quality difference when measurement conditions vary.'],
  rawTitle:['06　分析範圍內的原始數據','06  Raw Data in Analysis Range'],
  rawHint:['依目前起始與結束波長，顯示所有可用樣本的逐波長強度；沒有對應量測值時會留白。資料較多時可使用分頁瀏覽。','Shows each available sample’s intensity by wavelength within the selected range. Missing measurements are blank; use pagination for large datasets.'],
  rawWavelength:['波長 (nm)','Wavelength (nm)'],
  rawWaiting:['匯入資料後顯示波段數據','Band data appears after import'],
  jetsonTitle:['NVIDIA Jetson Orin Nano 光譜 AI 判斷','NVIDIA Jetson Orin Nano Spectral AI'],
  jetsonHint:['只會傳送目前勾選的分析曲線。GitHub Pages 需要 HTTPS Jetson API；存取權杖只保留在此頁面，不會儲存。','Only selected analysis curves are sent. GitHub Pages requires an HTTPS Jetson API; the access token stays in this page and is not stored.'],
  jetsonUrlLabel:['Jetson API 網址（HTTPS）','Jetson API URL (HTTPS)'],
  jetsonTokenLabel:['存取權杖','Access Token'],
  jetsonConnect:['測試 Jetson 連線','Test Jetson Connection'],
  jetsonPredict:['判斷已選的平均光譜','Classify Selected Average Spectra'],
  jetsonIdle:['尚未連線。需要 Jetson 分析服務與已載入的分類模型。','Not connected. A Jetson API service and a loaded classification model are required.'],
  jetsonSample:['樣本','Sample'],
  jetsonClass:['判斷結果','Classification'],
  jetsonConfidence:['信心度','Confidence'],
  previous:['上一頁','Previous'],
  next:['下一頁','Next'],
  noData:['尚無數據','No data'],
  finalHint:['使用提示：支援多組波長欄；每個樣本自動對應其左側最近的波長欄；重複樣本名稱會標示為第幾個同名欄位。若檔案只有圖片而非原始數值，請先提供 Excel / CSV 檔案。','Tip: Multiple wavelength columns are supported. Each sample uses the nearest wavelength column to its left; duplicate sample names are numbered. If your file contains only an image, provide the original Excel / CSV data.'],
};
const reverseStaticTranslations=Object.fromEntries(Object.entries(staticTranslations).map(([zh,en])=>[en,zh]));
function tStatic(text){if(currentLanguage==='en')return staticTranslations[text]||text;return reverseStaticTranslations[text]||text}
function applyLanguage(lang){
 currentLanguage=lang;localStorage.setItem('chiliscan-language',lang);
 document.documentElement.lang=lang==='zh'?'zh-Hant':'en';
 document.querySelectorAll('[data-lang]').forEach(btn=>{const active=btn.dataset.lang===lang;btn.classList.toggle('active',active);btn.setAttribute('aria-pressed',String(active))});
 document.querySelectorAll('[data-i18n]').forEach(el=>{
   const pair=uiTranslations[el.dataset.i18n];
   if(pair)el.textContent=pair[lang==='en'?1:0];
 });
 document.querySelectorAll('[data-i18n-placeholder]').forEach(el=>{
   const pair=uiTranslations[el.dataset.i18nPlaceholder];
   if(pair)el.placeholder=pair[lang==='en'?1:0];
 });
 document.title=uiTranslations.pageTitle[lang==='en'?1:0];
 if(workbook){
   if(csvImported){
     const oldName=workbook.SheetNames[0],newName=lang==='en'?'CSV Data':'CSV資料';
     if(oldName!==newName){workbook.Sheets[newName]=workbook.Sheets[oldName];delete workbook.Sheets[oldName];workbook.SheetNames[0]=newName;$('sheet').innerHTML=`<option value="${escapeHtml(newName)}">${escapeHtml(newName)}</option>`;$('sheet').value=newName;crossCache.clear()}
   }
   if($('compareMode').value==='cross'){populateCross();analyzeCross()}
   else parseSheet($('sheet').value);
 }else if(demoDataLoaded){
   headers=lang==='en'?['Wavelength','Fresh Pepper 1','Fresh Pepper 2','Fresh Pepper 3','Pest-Damaged Pepper 1','Rotten Pepper 1']:['波長','新鮮辣椒1','新鮮辣椒2','新鮮辣椒3','蟲害辣椒1','腐爛辣椒1'];
   $('sheet').innerHTML=`<option>${lang==='en'?'Demo Data (simulated)':'內建示範（非實測）'}</option>`;
   $('wave').innerHTML=headers.map((header,i)=>`<option value="${i}">${header}</option>`).join('');
   $('wave').value='0';analyze();
   $('status').innerHTML+=`<span class="warning">${lang==='en'?'　Simulated data, not experimental measurements.':'　目前為模擬曲線，並非你的實驗數據。'}</span>`;
 }else if(typeof render==='function')render();
 renderJetsonStatus();
 renderJetsonResults();
}
document.querySelectorAll('[data-lang]').forEach(btn=>btn.addEventListener('click',()=>applyLanguage(btn.dataset.lang)));
function fmt(x,d=2){return Number.isFinite(x)?x.toLocaleString(currentLanguage==='zh'?'zh-TW':'en-US',{maximumFractionDigits:d}):'—'}
function num(v){if(typeof v==='number')return Number.isFinite(v)?v:NaN;if(typeof v!=='string')return NaN;let s=v.trim().replace(/,/g,'');return s!==''&&Number.isFinite(Number(s))?Number(s):NaN}
function sampleLabel(name,occurrence){return occurrence>1?(currentLanguage==='en'?`${name} (duplicate ${occurrence})`:`${name}（第 ${occurrence} 個同名欄）`):name}

// 判斷表頭是否為重複量測，例如「新鮮辣椒1」與「新鮮辣椒2」。
function replicateBaseName(name){
 let base=String(name||'').trim()
  .replace(/\s*[（(](?:第\s*)?\d+\s*(?:個)?(?:同名欄)?[)）]\s*$/u,'')
  .replace(/\s*(?:_|-|\s)?(?:rep(?:licate)?|measurement|量測)\s*[_\-]?\s*\d+\s*$/iu,'')
  .replace(/\s*[_\-.\s]+\d+\s*$/u,'')
  .replace(/\d+\s*$/u,'')
  .trim();
 return base||String(name||'').trim();
}

function isWavelengthHeader(header){return /^(波長\s*\d*|wavelength\s*\d*|nm)$/i.test(String(header||'').trim())}
function isWavelengthColumn(dataRows,index){
 const values=dataRows.slice(0,150).map(row=>num(row?.[index])).filter(Number.isFinite);
 if(values.length<8)return false;
 const inRange=values.filter(value=>value>=200&&value<=1100).length/values.length;
 const rising=values.slice(1).filter((value,i)=>value>values[i]).length/(values.length-1);
 return inRange>.8&&rising>.75;
}
function pairedWavelengthColumns(rawHeaders,axisCandidates){
 const runs=[];
 axisCandidates.forEach(index=>{
  const previous=runs.at(-1),hasHeader=Boolean(String(rawHeaders[index]||'').trim());
  if(previous&&index===previous.at(-1)+2&&!hasHeader)previous.push(index);
  else runs.push([index]);
 });
 return runs.filter(run=>{
  const group=String(rawHeaders[run[0]]||'').trim();
  return run.length>=3&&group&&!isWavelengthHeader(group)&&run.slice(1).every(index=>!String(rawHeaders[index]||'').trim());
 }).flat();
}
function extractSpectra(rawHeaders,displayHeaders,dataRows,fallbackAxis,from=-Infinity,to=Infinity){
 const namedAxes=rawHeaders.map((header,index)=>isWavelengthHeader(header)?index:-1).filter(index=>index>=0);
 const inferredAxes=rawHeaders.map((_,index)=>isWavelengthColumn(dataRows,index)?index:-1).filter(index=>index>=0);
 const candidates=[...new Set([...namedAxes,...inferredAxes])].sort((a,b)=>a-b);
 const pairAxes=pairedWavelengthColumns(rawHeaders,candidates);
 const pairAxisSet=new Set(pairAxes);
 const normalAxes=pairAxes.length
  ?namedAxes.filter(index=>!pairAxisSet.has(index))
  :namedAxes.length?namedAxes:[Number.isInteger(fallbackAxis)?fallbackAxis:inferredAxes[0]].filter(Number.isInteger);
 const pairedColumns=new Set(pairAxes.flatMap(index=>[index,index+1]));
 const counts=new Map(),items=[];
 const makeData=(waveColumn,intensityColumn)=>dataRows.map(row=>({w:num(row?.[waveColumn]),y:num(row?.[intensityColumn])}))
  .filter(point=>Number.isFinite(point.w)&&Number.isFinite(point.y)&&point.w>=from&&point.w<=to)
  .sort((a,b)=>a.w-b.w);
 pairAxes.forEach(axis=>{
  const column=axis+1;
  if(column>=displayHeaders.length||pairAxisSet.has(column))return;
  let group='';
  for(let i=axis;i>=0;i--){if(String(rawHeaders[i]||'').trim()){group=String(rawHeaders[i]).trim();break}}
  if(!group)return;
  const data=makeData(axis,column);
  if(!data.length)return;
  const occurrence=(counts.get(group)||0)+1;
  counts.set(group,occurrence);
  const name=currentLanguage==='en'?`${group} (measurement ${occurrence})`:`${group}（量測${occurrence}）`;
  items.push({name,baseName:group,column,axis:group,data});
 });
 const ordinaryCounts={};
 for(let i=0;i<displayHeaders.length;i++){
  if(pairedColumns.has(i)||normalAxes.includes(i))continue;
  const axis=normalAxes.filter(index=>index<i).at(-1);
  if(axis===undefined)continue;
  const data=makeData(axis,i);
  if(!data.length)continue;
  const base=displayHeaders[i];
  ordinaryCounts[base]=(ordinaryCounts[base]||0)+1;
  items.push({name:sampleLabel(base,ordinaryCounts[base]),baseName:replicateBaseName(base),column:i,axis:displayHeaders[axis],data});
 }
 return {items,axes:[...new Set([...normalAxes,...pairAxes])].length};
}

// 將同名量測欄合併成逐波長平均。
// individual：只保留原曲線；average：只保留平均；both：原曲線＋平均。
function applyReplicateAveraging(rawSeries){
 const mode=$('replicateMode')?.value||'average';
 replicateSummary={groups:0,measurements:0};
 if(mode==='individual')return rawSeries;
 const groups=new Map();
 rawSeries.forEach(s=>{
  const base=s.baseName||replicateBaseName(s.name);
  if(!groups.has(base))groups.set(base,[]);
  groups.get(base).push(s);
 });
 const processed=[];
 groups.forEach((members,base)=>{
  if(members.length<2){
   processed.push(members[0]);
   return;
  }
  replicateSummary.groups++;
  replicateSummary.measurements+=members.length;
  if(mode==='both')processed.push(...members);
  const wavelengthMap=new Map();
  members.forEach(s=>{
   s.data.forEach(p=>{
    if(!wavelengthMap.has(p.w))wavelengthMap.set(p.w,[]);
    wavelengthMap.get(p.w).push(p.y);
   });
  });
  const data=[...wavelengthMap.entries()].map(([w,values])=>({
   w,
   y:values.reduce((sum,y)=>sum+y,0)/values.length
  })).sort((a,b)=>a.w-b.w);
  const suffix=currentLanguage==='en'?` (${members.length}-measurement mean)`:`（${members.length}筆平均）`;
  processed.push({
   ...members[0],
   name:`${base}${suffix}`,
   baseName:base,
   key:`average::${base}`,
   data,
   replicates:members.length
  });
 });
 return processed;
}
function replicateStatus(){
 const mode=$('replicateMode').value;
 if(mode==='individual')return '';
 if(replicateSummary.groups){
  return currentLanguage==='en'
   ?`Averaged ${replicateSummary.measurements} measurements across ${replicateSummary.groups} sample group(s).`
   :`已將 ${replicateSummary.groups} 組、共 ${replicateSummary.measurements} 筆重複量測逐波長平均。`;
 }
 return currentLanguage==='en'
  ?'No replicate columns detected. Check that repeated measurements share a sample name or group heading.'
  :'未偵測到重複量測欄；請確認重複量測使用相同樣本名稱或同一組欄位標題。';
}

function selectedJetsonSpectra(){
 const mode=$('replicateMode').value;
 if(mode==='individual')return [];
 return [...selected].map(index=>series[index]).filter(sample=>sample&&sample.data.length&&(mode!=='both'||sample.replicates>1));
}
function renderJetsonStatus(){
 const status=$('jetsonStatus');
 if(!status)return;
 const en=currentLanguage==='en';
 const messages={
  idle:en?'Not connected. A Jetson API service and a loaded classification model are required.':'尚未連線。需要 Jetson 分析服務與已載入的分類模型。',
  checking:en?'Checking the Jetson API…':'正在測試 Jetson API 連線…',
  ready:en?`Connected${jetsonDetail?`: ${jetsonDetail}`:''}. Ready to classify selected spectra.`:`已連線${jetsonDetail?`：${jetsonDetail}`:''}。可以判斷已選光譜。`,
  'no-model':en?'API is reachable, but no classification model is loaded. No prediction will be shown.':'API 已連線，但尚未載入分類模型；不會顯示分類結果。',
  predicting:en?'Sending selected spectra to Jetson for classification…':'正在將勾選的光譜送到 Jetson 判斷…',
  complete:en?`Classification complete: ${jetsonPredictions.length} sample(s).`:`分類完成：${jetsonPredictions.length} 個樣本。`,
  error:en?`Jetson request failed: ${jetsonDetail}`:`Jetson 連線失敗：${jetsonDetail}`
 };
 status.textContent=messages[jetsonPhase]||messages.idle;
 status.dataset.state=jetsonPhase;
 const busy=jetsonPhase==='checking'||jetsonPhase==='predicting';
 $('jetsonConnect').disabled=busy;
 $('jetsonUrl').disabled=busy;
 $('jetsonToken').disabled=busy;
 $('jetsonPredict').disabled=busy||!jetsonConnection?.ready||!selectedJetsonSpectra().length;
}
function renderJetsonResults(){
 const container=$('jetsonResults');
 if(!container)return;
 container.hidden=!jetsonPredictions.length;
 $('jetsonResultBody').innerHTML=jetsonPredictions.map(item=>`<tr><td>${escapeHtml(item.sample_name)}</td><td>${escapeHtml(item.label)}</td><td>${fmt(item.confidence*100,1)}%</td></tr>`).join('');
}
function clearJetsonPredictions(){
 jetsonPredictions=[];
 if(jetsonPhase==='complete')jetsonPhase=jetsonConnection?.ready?'ready':jetsonConnection?'no-model':'idle';
 renderJetsonStatus();
 renderJetsonResults();
}
function jetsonApiBase(){
 const raw=$('jetsonUrl').value.trim();
 if(!raw)throw Error(currentLanguage==='en'?'Enter the Jetson API URL.':'請輸入 Jetson API 網址。');
 let url;
 try{url=new URL(raw)}catch{throw Error(currentLanguage==='en'?'The Jetson API URL is invalid.':'Jetson API 網址格式不正確。')}
 if(!['https:','http:'].includes(url.protocol)||url.username||url.password||url.search||url.hash)
  throw Error(currentLanguage==='en'?'Use an HTTP(S) URL without credentials, query, or fragment.':'請使用不含帳密、查詢參數或片段的 HTTP(S) 網址。');
 const localHost=['localhost','127.0.0.1','[::1]'].includes(location.hostname);
 if(url.protocol==='http:'&&url.origin!==location.origin&&!localHost)
  throw Error(currentLanguage==='en'?'A secure page requires an HTTPS Jetson API.':'安全的 HTTPS 網頁需要使用 HTTPS Jetson API。');
 return url.href.replace(/\/+$/,'');
}
async function requestJetson(base,path,options={}){
 const token=$('jetsonToken').value.trim();
 if(!token)throw Error(currentLanguage==='en'?'Enter the access token.':'請輸入存取權杖。');
 const controller=new AbortController();
 const timeout=setTimeout(()=>controller.abort(),15000);
 try{
  const response=await fetch(`${base}${path}`,{
   ...options,
   headers:{Accept:'application/json',Authorization:`Bearer ${token}`,...options.headers},
   credentials:'omit',
   signal:controller.signal
  });
  let payload;
  try{payload=await response.json()}catch{throw Error(currentLanguage==='en'?'The Jetson API returned invalid JSON.':'Jetson API 回傳的 JSON 格式無效。')}
  if(!response.ok){
   const message=payload&&typeof payload==='object'?(payload.error||payload.message):'';
   throw Error(message||`HTTP ${response.status}`);
  }
  return payload;
 }catch(error){
  if(error.name==='AbortError')throw Error(currentLanguage==='en'?'The Jetson API request timed out.':'Jetson API 連線逾時。');
  throw error;
 }finally{clearTimeout(timeout)}
}
async function connectJetson(){
 jetsonConnection=null;jetsonPredictions=[];jetsonDetail='';jetsonPhase='checking';renderJetsonStatus();renderJetsonResults();
 try{
  const base=jetsonApiBase();
  const health=await requestJetson(base,'/api/v1/health');
  if(!health||typeof health!=='object'||health.api_version!==1||typeof health.model_ready!=='boolean')throw Error(currentLanguage==='en'?'Unsupported health response.':'Jetson 健康檢查回傳格式不受支援。');
  jetsonConnection={base,ready:health.model_ready};
  jetsonDetail=typeof health.model_name==='string'?health.model_name:'';
  jetsonPhase=health.model_ready?'ready':'no-model';
 }catch(error){jetsonPhase='error';jetsonDetail=error.message||String(error)}
 renderJetsonStatus();
}
async function classifyWithJetson(){
 const spectra=selectedJetsonSpectra();
 if(!jetsonConnection?.ready||!spectra.length)return;
 jetsonPredictions=[];jetsonDetail='';jetsonPhase='predicting';renderJetsonStatus();renderJetsonResults();
 try{
  const requestSpectra=spectra.map(sample=>({
   sample_name:sample.name,
   wavelengths_nm:sample.data.map(point=>point.w),
   intensities:sample.data.map(point=>point.y),
   replicate_count:sample.replicates||1
  }));
  const result=await requestJetson(jetsonConnection.base,'/api/v1/predict',{
   method:'POST',
   headers:{'Content-Type':'application/json'},
   body:JSON.stringify({api_version:1,spectra:requestSpectra})
  });
  if(!result||typeof result!=='object'||result.api_version!==1||!Array.isArray(result.predictions)||result.predictions.length!==requestSpectra.length)
   throw Error(currentLanguage==='en'?'Prediction response does not match the request.':'分類回傳筆數與送出樣本不一致。');
  const predictions=new Map();
  result.predictions.forEach(item=>{
   if(!item||typeof item.sample_name!=='string'||typeof item.label!=='string'||!item.label.trim()||
      !Number.isFinite(item.confidence)||item.confidence<0||item.confidence>1||predictions.has(item.sample_name))
    throw Error(currentLanguage==='en'?'Prediction response contains invalid fields.':'分類結果欄位格式無效。');
   predictions.set(item.sample_name,{sample_name:item.sample_name,label:item.label,confidence:item.confidence});
  });
  if(requestSpectra.some(sample=>!predictions.has(sample.sample_name)))
   throw Error(currentLanguage==='en'?'Prediction response is missing a requested sample.':'分類結果缺少送出的樣本。');
  jetsonPredictions=requestSpectra.map(sample=>predictions.get(sample.sample_name));
  jetsonPhase='complete';
 }catch(error){jetsonPhase='error';jetsonDetail=error.message||String(error)}
 renderJetsonStatus();
 renderJetsonResults();
}
$('jetsonConnect').addEventListener('click',connectJetson);
$('jetsonPredict').addEventListener('click',classifyWithJetson);
['jetsonUrl','jetsonToken'].forEach(id=>$(id).addEventListener('input',()=>{
 jetsonConnection=null;jetsonDetail='';jetsonPhase='idle';clearJetsonPredictions();
}));

// =====================================================
// 區塊二：Excel / CSV 讀取與工作表解析
// =====================================================
// parseSheet：讀取使用者選中的單一工作表，找出標題列、波長欄及全部數據列。
function parseSheet(name){
 const data=sheetToRows(workbook.Sheets[name]);
 if(!data.length)throw Error('工作表沒有資料');
 const head=data.findIndex(r=>r&&r.filter(v=>v!==null&&v!=='').length>=2);
 if(head<0)throw Error('找不到欄位標題');
 const rowHeaders=data[head],max=Math.max(...data.slice(head,head+5).map(r=>r.length));
 sourceHeaders=Array.from({length:max},(_,i)=>String(rowHeaders[i]??'').trim());
 headers=sourceHeaders.map((header,i)=>header||(currentLanguage==='en'?`Column ${i+1}`:`欄位${i+1}`));
 rows=data.slice(head+1).filter(r=>r&&r.some(v=>v!==null&&v!==''));
 const opts=headers.map((h,i)=>`<option value="${i}">${escapeHtml(h)}${currentLanguage==='en'?` (column ${i+1})`:`（第${i+1}欄）`}</option>`).join('');
 $('wave').innerHTML=opts;$('wave').disabled=false;
 let likely=headers.findIndex(isWavelengthHeader);
 if(likely<0)likely=headers.findIndex((_,i)=>isWavelengthColumn(rows,i));
 widx=likely>=0?likely:0;$('wave').value=String(widx);
 if($('compareMode').value==='single')analyze();
}
// escapeHtml：將檔案或工作表名稱中的特殊字元轉成安全字串，避免HTML被注入。
function escapeHtml(s){return String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
// parseCsv：將CSV文字逐字元轉成二維陣列，支援逗號、雙引號欄位及Windows換行。
function parseCsv(text){
 const rows=[];
 let row=[],cell='',quoted=false;
 text=text.replace(/^\uFEFF/, '');
 for(let i=0;i<text.length;i++){
  const c=text[i];
  if(c==='"' ){
   if(quoted&&text[i+1]==='"'){cell+='"';i++}else quoted=!quoted;
  }else if(c===','&&!quoted){row.push(cell);cell=''}
  else if((c==='\n'||c==='\r')&&!quoted){
   if(c==='\r'&&text[i+1]==='\n')i++;
   row.push(cell);rows.push(row);row=[];cell='';
  }else cell+=c;
 }
 row.push(cell);rows.push(row);
 return rows;
}
// 統一 Excel / CSV 工作表格式：回傳二維陣列。
function sheetToRows(sheet){
 if(sheet&&Array.isArray(sheet._rows))return sheet._rows;
 if(typeof XLSX==='undefined'||!XLSX.utils)return [];
 return XLSX.utils.sheet_to_json(sheet,{header:1,defval:null,raw:true});
}
// =====================================================
// 區塊三：單一工作表分析與跨工作表分析
// =====================================================
// analyze：單一工作表模式。
// 將所有數值欄配對到左側最近的波長欄，篩選波段，再建立 series 曲線資料。
function analyze(){
 if($('compareMode').value==='cross'){analyzeCross();return}
 let from=num($('from').value),to=num($('to').value);
 if(!Number.isFinite(from)||!Number.isFinite(to)||from>=to){$('status').textContent=tStatic('請輸入有效的波長範圍');return}
 clearJetsonPredictions();
 rawPage=0;
 widx=Number($('wave').value);
 const previousSelection=[...selected].map(i=>series[i]).filter(Boolean);
 const oldNames=new Set(previousSelection.map(s=>s.name));
 const oldBases=new Set(previousSelection.map(s=>s.baseName));
 const analysisHeaders=demoDataLoaded?headers:sourceHeaders;
 const extracted=extractSpectra(analysisHeaders,headers,rows,widx,from,to);
 series=extracted.items.map(item=>({...item,i:item.column}));
 series=applyReplicateAveraging(series);
 selected=new Set(series.map((v,i)=>oldNames.size?(oldNames.has(v.name)||oldBases.has(v.baseName)?i:-1):(i<Math.min(3,series.length)?i:-1)).filter(i=>i>=0));
 if(!selected.size&&series.length)selected.add(0);
 const waveCount=new Set(series.flatMap(s=>s.data.map(p=>p.w))).size;
 $('status').innerHTML=series.length?(currentLanguage==='en'?`<span class="ok">✓ Identified ${extracted.axes} wavelength column(s), ${series.length} spectral sample(s), and ${waveCount} distinct wavelength positions; analysis range ${fmt(from)}–${fmt(to)} nm.</span>`:`<span class="ok">✓ 已辨識 ${extracted.axes} 組波長欄、${series.length} 個光譜樣本；${waveCount} 個不同波長位置，分析範圍 ${fmt(from)}–${fmt(to)} nm。</span>`):(currentLanguage==='en'?'<span class="warning">No usable samples were found in the selected band. Check the worksheet or wavelength column.</span>':'<span class="warning">指定波段內沒有可用樣本，請檢查工作表或波長欄。</span>');
 if(series.length&&replicateStatus())$('status').innerHTML+=` <span class="replicate-result">${escapeHtml(replicateStatus())}</span>`;
 $('all').disabled=!series.length;$('export').disabled=!series.length;$('exportRaw').disabled=!series.length;render()
}
// readSheetSeries：跨工作表模式。
// 讀取指定工作表中的所有樣本；每個樣本只配對左側最近的波長欄。
function readSheetSeries(name){
 if(crossCache.has(name))return crossCache.get(name);
 const sh=workbook?.Sheets[name];if(!sh)return {error:'工作表不存在',items:[]};
 const arr=sheetToRows(sh);
 const head=arr.findIndex(r=>r&&r.filter(v=>v!==null&&v!==''&&v!==undefined).length>=2);
 if(head<0)return {error:'工作表沒有可用的欄位資料',items:[]};
 const max=Math.max(...arr.slice(head,head+5).map(r=>r?.length||0));
 const raw=Array.from({length:max},(_,i)=>String(arr[head][i]??'').trim());
 const hs=raw.map((header,i)=>header||`欄位${i+1}`);
 const rs=arr.slice(head+1).filter(r=>r&&r.some(v=>v!==null&&v!==''));
 const firstAxis=raw.findIndex((header,index)=>isWavelengthHeader(header)||isWavelengthColumn(rs,index));
 const extracted=extractSpectra(raw,hs,rs,firstAxis>=0?firstAxis:undefined);
 const items=extracted.items.map(item=>({...item,key:name+'::'+item.column,sheet:name}));
 const out={items,axes:extracted.axes,error:items.length?'':'找不到有效光譜數據'};
 crossCache.set(name,out);
 return out;
}
// populateCross：依目前匯入的工作表建立跨工作表勾選介面。
// 已勾選狀態會保留，工作表名稱則會出現在最終曲線名稱中。
function populateCross(){
 if(!workbook)return;
 const previousSheets=new Map([...document.querySelectorAll('.cross-sheet')].map(el=>[el.dataset.sheet,el.checked]));
 const previousSamples=new Map([...document.querySelectorAll('.cross-sample')].map(el=>[el.dataset.sheet+'::'+el.dataset.col,el.checked]));
 const preferred=['新鮮平均','蟲害平均','腐爛平均'];
 const hasPreferred=preferred.some(x=>workbook.SheetNames.includes(x));
 const html=workbook.SheetNames.map(name=>{
  const parsed=readSheetSeries(name),isDefault=hasPreferred?preferred.includes(name):name===$('sheet').value;
  const options=parsed.items.map((item,i)=>{const key=name+'::'+item.column,checked=previousSamples.has(key)?previousSamples.get(key):isDefault&&i===0,axisHint=item.axis===item.baseName?'':` <span class="muted">· ${escapeHtml(item.axis)}</span>`;return `<label><input type="checkbox" class="cross-sample" data-sheet="${escapeHtml(name)}" data-col="${item.column}" ${checked?'checked':''}>${escapeHtml(item.name)}${axisHint}</label>`}).join('');
  const sheetChecked=previousSheets.has(name)?previousSheets.get(name):isDefault&&parsed.items.length>0;
  const countText=currentLanguage==='en'?`${parsed.items.length} sample(s) · ${parsed.axes} wavelength column(s)`:`${parsed.items.length} 個樣本 · ${parsed.axes} 組波長欄`;
  return `<div class="cross-card"><h3><label><input class="cross-sheet" type="checkbox" data-sheet="${escapeHtml(name)}" ${sheetChecked?'checked':''} ${!parsed.items.length?'disabled':''}>${escapeHtml(name)}</label></h3>${parsed.items.length?`<div class="muted">${countText}</div><div class="cross-samples">${options}</div>`:`<div class="warning">${escapeHtml(currentLanguage==='en'?(parsed.error==='找不到波長欄'?'Wavelength column not found':parsed.error):parsed.error)}</div>`}</div>`;
 }).join('');$('crossSheets').innerHTML=html;
 $('selectRecommended').disabled=!hasPreferred;
 $('crossStatus').textContent=currentLanguage==='en'?(hasPreferred?'Available quality-average worksheets are preselected. Review the samples, then update the chart.':'No recommended quality-average worksheets found. Select worksheets manually.'):(hasPreferred?'已預選存在的品質平均工作表；請檢查各表樣本欄位，再點「更新疊加圖表」。':'沒有找到三種品質平均工作表，可手動勾選其他工作表。');
}
// analyzeCross：合併使用者勾選的跨工作表樣本，依波段篩選後交給 render() 顯示。
function analyzeCross(){
 if(!workbook){$('status').textContent=tStatic('請先匯入 Excel 檔案');return}
 const from=num($('from').value),to=num($('to').value);
 if(!Number.isFinite(from)||!Number.isFinite(to)||from>=to){$('status').textContent=tStatic('請輸入有效的波長範圍');return}
 clearJetsonPredictions();
 rawPage=0;
 const chosen=[...document.querySelectorAll('.cross-sheet:checked')].map(el=>el.dataset.sheet);
 const wanted=new Set([...document.querySelectorAll('.cross-sample:checked')].filter(el=>chosen.includes(el.dataset.sheet)).map(el=>el.dataset.sheet+'::'+el.dataset.col));
 const previousSelection=[...selected].map(i=>series[i]).filter(Boolean);
 const oldKeys=new Set(previousSelection.map(s=>s.key));
 const oldBases=new Set(previousSelection.map(s=>s.baseName));
 series=chosen.flatMap(name=>readSheetSeries(name).items.filter(item=>wanted.has(item.key)).map(item=>({...item,name:`${name}／${item.name}`,baseName:`${name}／${item.baseName}`,data:item.data.filter(p=>p.w>=from&&p.w<=to)})).filter(item=>item.data.length));
 series=applyReplicateAveraging(series);
 selected=new Set(series.map((s,i)=>oldKeys.size?(oldKeys.has(s.key)||oldBases.has(s.baseName)?i:-1):i).filter(i=>i>=0));
 if(!selected.size)selected=new Set(series.map((_,i)=>i));
 const totalWave=new Set(series.flatMap(s=>s.data.map(p=>p.w))).size;
 $('status').innerHTML=series.length?(currentLanguage==='en'?`<span class="ok">✓ Cross-worksheet: ${chosen.length} worksheet(s), ${series.length} curve(s), ${totalWave} distinct wavelengths; ${fmt(from)}–${fmt(to)} nm.</span>`:`<span class="ok">✓ 跨工作表：${chosen.length} 張工作表、${series.length} 條曲線、${totalWave} 個不同波長；${fmt(from)}–${fmt(to)} nm。</span>`):(currentLanguage==='en'?'<span class="warning">No curves match the selected band. Check worksheets, samples, or wavelength range.</span>':'<span class="warning">沒有符合波段的曲線，請勾選工作表及樣本，或檢查波長範圍。</span>');
 if(series.length&&replicateStatus())$('status').innerHTML+=` <span class="replicate-result">${escapeHtml(replicateStatus())}</span>`;
 $('crossStatus').textContent=series.length?(currentLanguage==='en'?'Overlaid: '+series.map(s=>s.name).join(', '):'已疊加 '+series.map(s=>s.name).join('、')):(currentLanguage==='en'?'No available samples selected':'尚未選取可用樣本');
 $('all').disabled=!series.length;$('export').disabled=!series.length;$('exportRaw').disabled=!series.length;render();
}
$('compareMode').onchange=()=>{const cross=$('compareMode').value==='cross';$('crossControls').hidden=!cross;$('sheet').disabled=cross||!workbook;$('wave').disabled=cross||!workbook;if(cross){if(workbook)populateCross();analyzeCross()}else{if(workbook)parseSheet($('sheet').value);else{series=[];selected.clear();render()}}};
$('selectRecommended').onclick=()=>{const wanted=new Set(['新鮮平均','蟲害平均','腐爛平均']);document.querySelectorAll('.cross-sheet').forEach(el=>{el.checked=wanted.has(el.dataset.sheet)&&!el.disabled});document.querySelectorAll('.cross-sample').forEach(el=>{let first=true;document.querySelectorAll('.cross-sample').forEach(candidate=>{if(candidate.dataset.sheet===el.dataset.sheet&&!el.checked&&first){first=false}});const isFirst=[...document.querySelectorAll('.cross-sample')].find(candidate=>candidate.dataset.sheet===el.dataset.sheet)===el;el.checked=wanted.has(el.dataset.sheet)&&isFirst});analyzeCross()};
$('crossApply').onclick=analyzeCross;
$('crossAll').onclick=()=>{document.querySelectorAll('.cross-sheet:not(:disabled)').forEach(el=>el.checked=true);document.querySelectorAll('.cross-sample').forEach(el=>el.checked=true);analyzeCross()};
$('crossSheets').addEventListener('click',e=>{
 if(!e.target.matches('.cross-sheet')||e.target.checked)return;
 const samples=[...document.querySelectorAll('.cross-sample')].filter(el=>el.dataset.sheet===e.target.dataset.sheet);
 const checkedCount=samples.filter(el=>el.checked).length;
 if(checkedCount>0&&checkedCount<samples.length){
  e.preventDefault();e.target.checked=true;samples.forEach(el=>el.checked=true);analyzeCross();
 }
});
$('crossSheets').addEventListener('change',e=>{
 if(e.target.matches('.cross-sheet')){
  document.querySelectorAll('.cross-sample').forEach(el=>{if(el.dataset.sheet===e.target.dataset.sheet)el.checked=e.target.checked});
 }else if(e.target.matches('.cross-sample')&&e.target.checked){
  document.querySelectorAll('.cross-sheet').forEach(el=>{if(el.dataset.sheet===e.target.dataset.sheet)el.checked=true});
 }
 analyzeCross();
});
// 快速套用指定波段；若已有資料，立即重新分析並更新匯出內容。
$('presetBand').onclick=()=>{$('from').value='620';$('to').value='680';if(workbook||demoDataLoaded)analyze()};
$('replicateMode').addEventListener('change',()=>{
 if(workbook)analyze();
 else if(demoDataLoaded)analyze();
});
// =====================================================
// 區塊四：特徵計算、表格、樣本選擇與圖表繪製
// =====================================================
// feature：從單一樣本在指定波段內找出波峰、平均、最低與有效點數。
function feature(s){const a=s.data;if(!a.length)return null;let peak=a.reduce((p,c)=>c.y>p.y?c:p,a[0]);return {name:s.name,w:peak.w,max:peak.y,mean:a.reduce((v,p)=>v+p.y,0)/a.length,min:Math.min(...a.map(p=>p.y)),n:a.length}}
// render：統一更新摘要卡、特徵表、原始數據表、圖例、樣本清單與Canvas。
function render(){const first=[...selected][0],f=first===undefined?null:feature(series[first]);$('peakW').textContent=f?fmt(f.w,3):'—';$('peakI').textContent=f?fmt(f.max):'—';$('mean').textContent=f?fmt(f.mean):'—';$('points').textContent=f?f.n:'—';const onlyLabel=currentLanguage==='en'?'Only':'單獨';$('samples').innerHTML=series.length?series.map((s,i)=>`<label class="sample"><input type="checkbox" data-i="${i}" ${selected.has(i)?'checked':''}><span>${escapeHtml(s.name)}</span><button class="mini" data-only="${i}" type="button">${onlyLabel}</button></label>`).join(''):(currentLanguage==='en'?'<div class="empty">No drawable samples</div>':'<div class="empty">沒有可繪製的樣本</div>');$('legend').innerHTML=[...selected].map(i=>`<label><i style="background:${palette[i%palette.length]}"></i>${escapeHtml(series[i].name)}</label>`).join('');$('tbody').innerHTML=series.length?series.map(s=>{const f=feature(s);return f?`<tr><td>${escapeHtml(f.name)}</td><td>${fmt(f.w,3)}</td><td>${fmt(f.max)}</td><td>${fmt(f.mean)}</td><td>${fmt(f.min)}</td><td>${f.n}</td></tr>`:''}).join(''):(currentLanguage==='en'?'<tr><td colspan="6">No data</td></tr>':'<tr><td colspan="6">沒有資料</td></tr>');renderRawData();draw();renderJetsonStatus()}
// 將目前波段內所有樣本整理成逐波長表格，最多每頁顯示100筆。
function renderRawData(){
 const usable=series.filter(s=>s.data.length);
 const wavelengths=[...new Set(usable.flatMap(s=>s.data.map(p=>p.w)))].sort((a,b)=>a-b);
 const pageCount=Math.ceil(wavelengths.length/rawPageSize);
 rawPage=Math.max(0,Math.min(rawPage,pageCount-1));
 $('rawDataHead').innerHTML=`<tr><th>${currentLanguage==='en'?'Wavelength (nm)':'波長 (nm)'}</th>${usable.map(s=>`<th>${escapeHtml(s.name)}</th>`).join('')}</tr>`;
 if(!wavelengths.length){
  $('rawDataBody').innerHTML=`<tr><td colspan="${Math.max(1,usable.length+1)}" class="empty">${currentLanguage==='en'?'Band data appears after import and analysis':'匯入資料並完成分析後，這裡會顯示波段數據'}</td></tr>`;
  $('rawPageStatus').textContent=currentLanguage==='en'?'No data':'尚無數據';
  $('rawPrev').disabled=true;$('rawNext').disabled=true;
  return;
 }
 const values=usable.map(s=>new Map(s.data.map(p=>[p.w,p.y])));
 const pageWaves=wavelengths.slice(rawPage*rawPageSize,(rawPage+1)*rawPageSize);
 $('rawDataBody').innerHTML=pageWaves.map(w=>`<tr><td>${fmt(w,3)}</td>${values.map(map=>`<td>${map.has(w)?fmt(map.get(w),6):'—'}</td>`).join('')}</tr>`).join('');
 $('rawPageStatus').textContent=currentLanguage==='en'?`Page ${rawPage+1} / ${pageCount}; ${wavelengths.length} distinct wavelengths, ${usable.length} ${usable.length===1?'sample':'samples'}`:`第 ${rawPage+1} / ${pageCount} 頁；共 ${wavelengths.length} 個不同波長、${usable.length} 個樣本`;
 $('rawPrev').disabled=rawPage===0;$('rawNext').disabled=rawPage>=pageCount-1;
}
$('rawPrev').onclick=()=>{if(rawPage>0){rawPage--;renderRawData()}};
$('rawNext').onclick=()=>{rawPage++;renderRawData()};
$('samples').addEventListener('change',e=>{if(e.target.matches('input[data-i]')){let i=Number(e.target.dataset.i);e.target.checked?selected.add(i):selected.delete(i);clearJetsonPredictions();render()}});$('samples').addEventListener('click',e=>{if(e.target.dataset.only!==undefined){selected=new Set([Number(e.target.dataset.only)]);clearJetsonPredictions();render()}});$('all').onclick=()=>{selected=selected.size===series.length?new Set():new Set(series.map((_,i)=>i));clearJetsonPredictions();render()};
// draw：Canvas繪圖入口，負責座標軸、刻度、曲線平滑、Y軸縮放與游標提示。
function draw(){const c=$('chart'),r=c.getBoundingClientRect(),dpr=window.devicePixelRatio||1;c.width=Math.max(300,Math.round(r.width*dpr));c.height=Math.max(220,Math.round(r.height*dpr));const ctx=c.getContext('2d');ctx.scale(dpr,dpr);const W=c.width/dpr,H=c.height/dpr;ctx.clearRect(0,0,W,H);const L=62,R=18,T=20,B=44,pw=W-L-R,ph=H-T-B;ctx.font='12px sans-serif';ctx.textBaseline='middle';const vis=[...selected].map(i=>({s:series[i],i})).filter(o=>o.s&&o.s.data.length);if(!vis.length){ctx.fillStyle='#91a5bc';ctx.textAlign='center';ctx.fillText(currentLanguage==='en'?'Check at least one sample':'請勾選至少一個樣本',W/2,H/2);return}const x0=Number($('from').value),x1=Number($('to').value);let vals=vis.flatMap(o=>o.s.data.map(p=>p.y));const dataMin=Math.min(...vals),dataMax=Math.max(...vals);let y0=Math.min(0,dataMin),y1=Math.max(0,dataMax);if(y0===y1)y1=y0+1;let pad=(y1-y0)*.07;y0-=pad;y1+=pad;if(dataMin>=0)y0=0;else if(dataMax<=0)y1=0;if($('ymode').value==='manual'){const low=num($('ymin').value),high=num($('ymax').value);if(Number.isFinite(low)&&Number.isFinite(high)&&low<high){y0=low;y1=high;$('yerror').style.display='none'}else{$('yerror').textContent=currentLanguage==='en'?'Y-axis minimum must be less than maximum, and both values must be numbers. Auto scale is shown for now.':'Y軸最小值必須小於最大值，且兩者皆需填寫數字；目前暫以自動刻度顯示。';$('yerror').style.display='block'}}else{$('yerror').style.display='none'}const X=v=>L+(v-x0)/(x1-x0)*pw,Y=v=>T+(y1-v)/(y1-y0)*ph;ctx.lineWidth=1;ctx.strokeStyle='#314255';ctx.fillStyle='#9eb1c8';ctx.textAlign='right';for(let i=0;i<=5;i++){let y=T+ph*i/5;ctx.beginPath();ctx.moveTo(L,y);ctx.lineTo(W-R,y);ctx.stroke();ctx.fillText(fmt(y1-(y1-y0)*i/5,0),L-10,y)}ctx.textAlign='center';const requestedStep=num($('xstep').value);const step=Number.isFinite(requestedStep)&&requestedStep>0?requestedStep:(x1-x0)/6;const tickCount=Math.floor((x1-x0)/step)+1;if(tickCount>150){$('status').textContent=currentLanguage==='en'?'X-axis ticks are too dense; increase the tick spacing.':'X軸刻度過密，請調大刻度間距';}const safeStep=tickCount>150?(x1-x0)/6:step;let ticks=[];const start=Math.ceil((x0-1e-9)/safeStep)*safeStep;for(let v=start,n=0;v<=x1+1e-8&&n<151;v+=safeStep,n++)ticks.push(Number(v.toFixed(9)));if(!ticks.length)ticks=[x0,x1];const labelStride=Math.max(1,Math.ceil(ticks.length/Math.max(2,Math.floor(pw/62))));ticks.forEach((v,i)=>{let x=X(v);ctx.beginPath();ctx.moveTo(x,T);ctx.lineTo(x,T+ph);ctx.stroke();if(i%labelStride===0)ctx.fillText(fmt(v,2),x,H-23)});vis.forEach(({s,i})=>{ctx.beginPath();ctx.strokeStyle=palette[i%palette.length];ctx.lineWidth=2.4;const n=Number($('smooth').value);const pts=s.data.map((p,j)=>{if(n===1)return p;let half=Math.floor(n/2),part=s.data.slice(Math.max(0,j-half),Math.min(s.data.length,j+half+1));return {w:p.w,y:part.reduce((a,b)=>a+b.y,0)/part.length}});pts.forEach((p,j)=>j?ctx.lineTo(X(p.w),Y(p.y)):ctx.moveTo(X(p.w),Y(p.y)));ctx.stroke()});if(hover!==null&&hover>=L&&hover<=W-R){let wave=x0+(hover-L)/pw*(x1-x0);ctx.strokeStyle='#b9c9dc';ctx.setLineDash([4,5]);ctx.beginPath();ctx.moveTo(hover,T);ctx.lineTo(hover,T+ph);ctx.stroke();ctx.setLineDash([]);let text=vis.slice(0,4).map(({s})=>{let p=s.data.reduce((a,b)=>Math.abs(b.w-wave)<Math.abs(a.w-wave)?b:a);return `${s.name}: ${fmt(p.y)} @ ${fmt(p.w,2)} nm`});let boxW=Math.min(W-20,Math.max(155,...text.map(s=>ctx.measureText(s).width+18)));let bx=Math.min(Math.max(hover+10,8),W-boxW-8),by=T+10;ctx.fillStyle='#24364a';ctx.fillRect(bx,by,boxW,22+text.length*18);ctx.fillStyle='#fff';ctx.textAlign='left';ctx.fillText(`${currentLanguage==='en'?'Cursor':'游標位置'} ${fmt(wave,2)} nm`,bx+9,by+13);text.forEach((s,i)=>ctx.fillText(s,bx+9,by+32+i*18))}}
$('smooth').addEventListener('change',draw);$('ymode').addEventListener('change',()=>{const manual=$('ymode').value==='manual';$('ymin').disabled=!manual;$('ymax').disabled=!manual;if(manual){const vals=[...selected].flatMap(i=>(series[i]?.data||[]).map(p=>p.y));if(vals.length){const lo=Math.min(0,...vals),hi=Math.max(...vals),pad=Math.max(1,(hi-lo)*.07);if(!$('ymin').value)$('ymin').value=String(Math.floor(lo-pad));if(!$('ymax').value)$('ymax').value=String(Math.ceil(hi+pad))}}draw()});['ymin','ymax'].forEach(id=>$(id).addEventListener('input',draw));$('yreset').addEventListener('click',()=>{$('ymode').value='auto';$('ymin').value='';$('ymax').value='';$('ymin').disabled=true;$('ymax').disabled=true;draw()});$('chart').addEventListener('mousemove',e=>{hover=e.clientX-$('chart').getBoundingClientRect().left;draw()});$('chart').addEventListener('mouseleave',()=>{hover=null;draw()});window.addEventListener('resize',draw);
// =====================================================
// 區塊五：檔案匯入、拖曳與示範資料
// =====================================================
// drag / file：負責接收使用者選擇或拖曳進來的 CSV 或 Excel 檔案。
$('drop').onclick=e=>{if(e.target!==$('file'))$('file').click()};$('drop').ondragover=e=>{e.preventDefault();$('drop').style.borderColor='#4ce2b6'};$('drop').ondragleave=()=>$('drop').style.borderColor='';$('drop').ondrop=e=>{e.preventDefault();$('drop').style.borderColor='';if(e.dataTransfer.files.length)load(e.dataTransfer.files[0])};$('file').onchange=e=>{if(e.target.files[0])load(e.target.files[0])};
$('drop').addEventListener('keydown',e=>{
 if(e.key==='Enter'||e.key===' '){
  e.preventDefault();
  $('file').click();
 }
});
// load()：讀取CSV或Excel；Excel解析使用SheetJS。
async function load(file){try{
 demoDataLoaded=false;csvImported=false;
 const lower=file.name.toLowerCase();
 if(lower.endsWith('.csv')){
  const text=await file.text();
  const csvRows=parseCsv(text);
  if(!csvRows.length)throw Error('CSV 沒有資料');
  csvImported=true;
  workbook={SheetNames:['CSV資料'],Sheets:{'CSV資料':{_rows:csvRows}}};
 }else{
  if(typeof XLSX==='undefined')throw Error('Excel解析套件尚未載入，請連上網路後重新整理。');
  const ab=await file.arrayBuffer();
  workbook=XLSX.read(ab,{type:'array'});
 }
 crossCache.clear();$('sheet').innerHTML=workbook.SheetNames.map(s=>`<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join('');$('sheet').disabled=false;
 const preferred=workbook.SheetNames.find(s=>s.trim()==='新鮮平均')||workbook.SheetNames.find(s=>s.trim()==='光強')||workbook.SheetNames[0];
 $('sheet').value=preferred;populateCross();
 if($('compareMode').value==='cross')analyzeCross();else parseSheet(preferred);
 $('status').innerHTML+=`<span class="sub">　檔案：${escapeHtml(file.name)}</span>`;
}catch(e){$('status').innerHTML=`<span class="warning">無法匯入：${escapeHtml(e.message)}</span>`}}$('sheet').onchange=()=>{try{parseSheet($('sheet').value)}catch(e){$('status').innerHTML='<span class="warning">工作表讀取失敗：'+escapeHtml(e.message)+'</span>'}};$('wave').onchange=analyze;
['from','to'].forEach(id=>$(id).addEventListener('change',()=>{if(workbook||demoDataLoaded)analyze()}));
$('xstep').addEventListener('input',draw);
$('demo').onclick=()=>{demoDataLoaded=true;headers=currentLanguage==='en'?['Wavelength','Fresh Pepper 1','Fresh Pepper 2','Fresh Pepper 3','Pest-Damaged Pepper 1','Rotten Pepper 1']:['波長','新鮮辣椒1','新鮮辣椒2','新鮮辣椒3','蟲害辣椒1','腐爛辣椒1'];rows=Array.from({length:61},(_,i)=>{const w=600+i;return [w,420+3650*Math.exp(-.5*((w-630)/8.8)**2),380+3380*Math.exp(-.5*((w-632)/9.5)**2),460+3500*Math.exp(-.5*((w-629)/9.2)**2),460+2580*Math.exp(-.5*((w-626)/10)**2),540+1600*Math.exp(-.5*((w-635)/11)**2)]});$('sheet').innerHTML=`<option>${currentLanguage==='en'?'Demo Data (simulated)':'內建示範（非實測）'}</option>`;$('sheet').disabled=true;$('wave').innerHTML=headers.map((s,i)=>`<option value="${i}">${s}</option>`).join('');$('wave').disabled=false;$('wave').value='0';analyze();$('status').innerHTML+=`<span class="warning">${currentLanguage==='en'?'　Simulated data, not experimental measurements.':'　目前為模擬曲線，並非你的實驗數據。'}</span>`};
// =====================================================
// 區塊六：CSV匯出
// =====================================================
// export：特徵與原始數據分開匯出，避免把統計值與逐波長原始值混在同一份檔案。
function csvEscape(v){let s=String(v??'');return /[",\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s}function downloadCsv(name,data){const text='\ufeff'+data.map(row=>row.map(csvEscape).join(',')).join('\r\n');const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([text],{type:'text/csv;charset=utf-8'}));a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}$('export').onclick=()=>{const headers=currentLanguage==='en'?['Sample Name','Peak Wavelength(nm)','Peak Intensity','Mean Intensity','Minimum Intensity','Measurement Points']:['樣本名稱','波峰波長(nm)','波峰強度','平均強度','最低強度','量測點數'];downloadCsv(currentLanguage==='en'?'ChiliScan_Spectral_Features.csv':'ChiliScan_光譜特徵.csv',[headers,...series.map(s=>{const f=feature(s);return [f.name,f.w,f.max,f.mean,f.min,f.n]})])};$('exportRaw').onclick=()=>{const picked=[...selected].map(i=>series[i]).filter(Boolean);if(!picked.length){alert(currentLanguage==='en'?'Please select at least one sample':'請先勾選樣本');return}let wavelengths=[...new Set(picked.flatMap(s=>s.data.map(p=>p.w)))].sort((a,b)=>a-b);downloadCsv(currentLanguage==='en'?'ChiliScan_Selected_Band.csv':'ChiliScan_選定波段.csv',[[(currentLanguage==='en'?'Wavelength(nm)':'波長(nm)'),...picked.map(s=>s.name)],...wavelengths.map(w=>[w,...picked.map(s=>s.data.find(p=>p.w===w)?.y??'')])])};applyLanguage(currentLanguage);draw();
