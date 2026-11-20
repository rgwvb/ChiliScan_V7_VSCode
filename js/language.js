(() => {
  const translations = {
    '／ 光譜分析平台': '/ Spectral Analysis Platform',
    '匯入 Excel / CSV 光譜資料，自動辨識波長欄、繪製樣本曲線，計算指定波段的光譜特徵。所有分析在瀏覽器內進行，不會上傳資料至伺服器。': 'Import Excel / CSV spectrum data to detect wavelength columns, plot sample curves, and calculate spectral features in a selected range. All analysis runs in your browser; no data is uploaded.',
    '01　匯入光譜資料': '01 Import Spectrum Data',
    '⬆ 點擊選擇 Excel / CSV，或將檔案拖曳到這裡': '⬆ Choose an Excel / CSV file, or drag and drop it here',
    '支援 .xlsx、.xls、.csv；匯入後可切換工作表；優先開啟「新鮮平均」，否則選擇「光強」；支援跨工作表光譜疊加與多組波長欄；預設600–680 nm，曲線可切換平滑／原始顯示': 'Supports .xlsx, .xls, and .csv. Switch worksheets after import; prefers "Fresh Average", then "Light Intensity". Supports cross-sheet overlays and multiple wavelength columns. Default range: 600–680 nm; choose smoothed or raw curves.',
    '工作表（優先：新鮮平均／光強）': 'Worksheet (prefers Fresh Average / Light Intensity)',
    '尚未匯入': 'Not imported',
    '內建示範（非實測）': 'Built-in demo (simulated)',
    '波長欄': 'Wavelength column',
    '自動辨識': 'Auto-detect',
    '起始波長（nm）': 'Start wavelength (nm)',
    '結束波長（nm）': 'End wavelength (nm)',
    'X軸刻度間距（nm）': 'X-axis interval (nm)',
    '重新分析': 'Reanalyze',
    '載入示範資料': 'Load demo data',
    '等待資料匯入。': 'Waiting for data import.',
    '02　跨工作表光譜疊加比較': '02 Cross-Worksheet Spectrum Comparison',
    '（可同時選擇新鮮平均、蟲害平均、腐爛平均）': '(Select Fresh Average, Pest-Damaged Average, and Rotten Average together)',
    '分析模式': 'Analysis mode',
    '單一工作表': 'Single worksheet',
    '跨工作表疊加比較': 'Cross-worksheet overlay',
    '選取三種品質平均工作表': 'Select the three quality-average worksheets',
    '勾選要比較的工作表，再於每張工作表內選擇樣本欄位。每個樣本會使用其左側最近的「波長」欄，不會按照資料列號強制配對。工作表名稱會加入圖例。': 'Select worksheets to compare, then choose sample columns in each. Each sample uses the nearest "Wavelength" column to its left; rows are not forcibly paired. Worksheet names are included in the legend.',
    '請先匯入 Excel 檔案': 'Import an Excel file first',
    '更新疊加圖表': 'Update overlay chart',
    '勾選所有可用樣本': 'Select all available samples',
    '03　波段特徵摘要': '03 Wavelength-Range Feature Summary',
    '（以目前勾選的第一個樣本為主）': '(Based on the first currently selected sample)',
    '波峰波長': 'Peak wavelength',
    '波峰強度': 'Peak intensity',
    '區間平均強度': 'Range mean intensity',
    '有效量測點數': 'Valid measurement points',
    '點': 'points',
    '04　光譜強度曲線': '04 Spectrum Intensity Curve',
    '縱軸刻度模式': 'Y-axis scale',
    '自動調整': 'Auto',
    '手動設定': 'Manual',
    'Y軸最小值': 'Y-axis minimum',
    '例如 0': 'e.g. 0',
    'Y軸最大值': 'Y-axis maximum',
    '例如 4500': 'e.g. 4500',
    '重設Y軸': 'Reset Y-axis',
    '曲線顯示方式': 'Curve display',
    '平滑顯示（3點移動平均）': 'Smoothed (3-point moving average)',
    '平滑顯示（5點移動平均）': 'Smoothed (5-point moving average)',
    '原始數據（不平滑）': 'Raw data (not smoothed)',
    '橫軸：實際波長（nm）；縱軸：原始訊號強度（未進行反射率校正）。平滑僅作用於曲線顯示，不影響原始資料、波峰與平均強度計算。滑鼠移至曲線區域可查看最近波長的數值。': 'X-axis: actual wavelength (nm); Y-axis: raw signal intensity (without reflectance correction). Smoothing affects display only, not the source data or peak/mean calculations. Hover over the chart to see values at the nearest wavelength.',
    '樣本選擇': 'Sample Selection',
    '匯入資料後顯示樣本欄位': 'Sample columns appear after importing data',
    '勾選樣本即可疊加比較；按「單獨」可只查看該樣本。': 'Select samples to compare; choose "Solo" to view only that sample.',
    '單獨': 'Solo',
    '全部顯示／取消': 'Select all / clear',
    '05　各樣本光譜特徵比較': '05 Sample Spectrum Feature Comparison',
    '匯出特徵 CSV': 'Export feature CSV',
    '匯出選定波段 CSV': 'Export selected-range CSV',
    '樣本名稱（含工作表）': 'Sample (including worksheet)',
    '波峰波長 (nm)': 'Peak wavelength (nm)',
    '平均強度': 'Mean intensity',
    '最低強度': 'Minimum intensity',
    '量測點數': 'Measurement points',
    '尚無分析資料': 'No analysis data',
    '特徵僅依指定波長範圍計算。若樣本未落入範圍，該樣本不列入分析。不同樣本量測條件不一致時，強度不可直接視為品質差異。': 'Features are calculated only within the selected wavelength range. Samples outside the range are excluded. Intensity should not be treated as a quality difference when measurement conditions vary.',
    '使用提示：支援範例.xlsx 的多組波長欄：每個樣本自動對應其左側最近的波長欄；重複樣本名稱自動編號。若檔案只有圖片而非原始數值，請先提供 Excel / CSV 檔案。': 'Tip: Multiple wavelength columns are supported. Each sample is paired with the nearest wavelength column to its left; duplicate sample names are numbered automatically. If your file contains images instead of raw values, provide an Excel / CSV file.',
    '請勾選至少一個樣本': 'Select at least one sample',
    '沒有可分析的樣本': 'No analyzable samples',
    'Excel解析套件尚未載入，請連上網路後重新整理。': 'The Excel parser did not load. Refresh the page or contact the administrator.',
    '請先匯入資料': 'Import data first',
    '請先勾選樣本': 'Select samples first',
    '目前為模擬曲線，並非你的實驗數據。': 'These are simulated curves, not your experimental data.',
    '波長欄': 'wavelength column',
    '個光譜樣本': 'spectral sample(s)',
    '個不同波長位置，分析範圍': ' distinct wavelength positions; analysis range',
    '指定波段內沒有可用樣本，請檢查工作表或波長欄。': 'No samples are available in the selected range. Check the worksheet or wavelength column.',
    '工作表讀取失敗：': 'Failed to read worksheet: ',
    '無法匯入：': 'Import failed: ',
    '　檔案：': ' File: ',
    'X軸刻度間距必須大於0': 'The X-axis interval must be greater than 0',
    'Y軸最小值必須小於最大值': 'The Y-axis minimum must be less than the maximum',
    '沒有可匯出的分析資料': 'No analysis data to export',
    '已預選存在的品質平均工作表；請檢查各表樣本欄位，再點「更新疊加圖表」。': 'Available quality-average worksheets were preselected. Check the sample columns, then click "Update overlay chart".',
    '沒有找到三種品質平均工作表，可手動勾選其他工作表。': 'The three quality-average worksheets were not found. You can select other worksheets manually.',
    '尚未選取可用樣本': 'No available samples selected',
    '沒有符合波段的曲線，請勾選工作表及樣本，或檢查波長範圍。': 'No curves match the selected range. Select worksheets and samples, or check the wavelength range.',
    'X軸刻度過密，請調大刻度間距': 'X-axis ticks are too dense. Increase the interval.',
    'Y軸最小值必須小於最大值，且兩者皆需填寫數字；目前暫以自動刻度顯示。': 'The Y-axis minimum must be less than the maximum, and both must be numeric. The chart is temporarily using automatic scaling.',
    ' 組波長欄、': ' wavelength column(s), ',
    '✓ 已辨識 ': '✓ Identified ',
  };
  const englishEntries = Object.entries(translations).sort((a, b) => b[0].length - a[0].length);
  const chineseByEnglish = new Map(englishEntries.map(([zh, en]) => [en, zh]));
  const chineseEntries = [...chineseByEnglish.entries()].sort((a, b) => b[0].length - a[0].length);
  let language = localStorage.getItem('chiliscan-language') === 'en' ? 'en' : 'zh';
  let updating = false;

  function translateText(text) {
    const entries = language === 'en' ? englishEntries : chineseEntries;
    for (const [source, target] of entries) text = text.split(source).join(target);
    return text;
  }

  function translateNode(node) {
    let value = translateText(node.nodeValue);
    if (node.parentElement && node.parentElement.closest('#wave option')) {
      value = language === 'en'
        ? value.replace(/（第(\d+)欄）/g, ' (column $1)')
        : value.replace(/ \(column (\d+)\)/g, '（第$1欄）');
    }
    if (node.parentElement && node.parentElement.closest('.status')) {
      value = language === 'en'
        ? value.replace(/；/g, '; ').replace(/。/g, '.')
        : value.replace(/; /g, '；').replace(/nm\./g, 'nm。');
    }
    if (value !== node.nodeValue) node.nodeValue = value;
  }

  function translateTree(root) {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode(node) {
        return node.parentElement && !['SCRIPT', 'STYLE'].includes(node.parentElement.tagName)
          ? NodeFilter.FILTER_ACCEPT
          : NodeFilter.FILTER_REJECT;
      },
    });
    const nodes = [];
    while (walker.nextNode()) nodes.push(walker.currentNode);
    updating = true;
    for (const node of nodes) translateNode(node);
    const elements = root.querySelectorAll ? root.querySelectorAll('[placeholder], [title]') : [];
    for (const element of elements) {
      for (const attribute of ['placeholder', 'title']) {
        if (element.hasAttribute(attribute)) {
          element.setAttribute(attribute, translateText(element.getAttribute(attribute)));
        }
      }
    }
    updating = false;
  }

  function setLanguage(nextLanguage) {
    language = nextLanguage;
    localStorage.setItem('chiliscan-language', language);
    document.documentElement.lang = language === 'en' ? 'en' : 'zh-Hant';
    document.title = language === 'en' ? 'ChiliScan | Spectral Data Analysis Platform V7' : 'ChiliScan｜光譜數據分析平台 V7';
    for (const button of document.querySelectorAll('.language-switch button')) {
      button.setAttribute('aria-pressed', String(button.dataset.language === language));
    }
    translateTree(document.body);
    window.dispatchEvent(new Event('languagechange'));
  }

  const header = document.querySelector('header');
  const tag = header.querySelector('.tag');
  const actions = document.createElement('div');
  actions.className = 'header-actions';
  const languageSwitch = document.createElement('div');
  languageSwitch.className = 'language-switch';
  languageSwitch.setAttribute('aria-label', 'Language');
  languageSwitch.innerHTML = '<button type="button" data-language="zh">中文</button><button type="button" data-language="en">English</button>';
  languageSwitch.addEventListener('click', event => {
    const button = event.target.closest('button[data-language]');
    if (button) setLanguage(button.dataset.language);
  });
  actions.append(tag, languageSwitch);
  header.append(actions);

  const style = document.createElement('style');
  style.textContent = '.header-actions{display:flex;align-items:center;gap:9px}.language-switch{display:inline-flex;padding:3px;background:#101b29;border:1px solid #304057;border-radius:9px}.language-switch button{border:0;background:transparent;color:#9caec2;padding:7px 9px;border-radius:6px;font:inherit;font-size:12px;cursor:pointer}.language-switch button[aria-pressed=true]{background:#3cb996;color:#071c18;font-weight:700}@media(max-width:520px){header{align-items:flex-start;gap:12px}.header-actions{flex-wrap:wrap;justify-content:flex-end}.brand{font-size:19px}}';
  document.head.append(style);

  const originalFillText = CanvasRenderingContext2D.prototype.fillText;
  CanvasRenderingContext2D.prototype.fillText = function (text, ...args) {
    return originalFillText.call(this, translateText(String(text)), ...args);
  };

  const observer = new MutationObserver(records => {
    if (updating) return;
    for (const record of records) {
      if (record.type === 'characterData') {
        translateNode(record.target);
      } else {
        for (const node of record.addedNodes) {
          if (node.nodeType === Node.TEXT_NODE) {
            translateNode(node);
          } else if (node.nodeType === Node.ELEMENT_NODE && !['SCRIPT', 'STYLE'].includes(node.tagName)) {
            translateTree(node);
          }
        }
      }
    }
  });
  observer.observe(document.body, { childList: true, characterData: true, subtree: true });
  window.addEventListener('languagechange', () => {
    if (typeof draw === 'function') draw();
  });
  setLanguage(language);
})();
