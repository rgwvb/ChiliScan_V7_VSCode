(() => {
  'use strict';

  ['chiliscan-kumo-reference-library', 'chiliscan-kumo-api-url', 'chiliscan-kumo-api-token']
    .forEach((key) => localStorage.removeItem(key));

  // 本模組處理新鮮／腐爛吸光度頁：讀取參考與待測檔、計算 RMSE 相對相似度並更新畫面。
  // 光譜處理與分類在瀏覽器執行；使用者可選擇儲存快照到雲端。
  const BAND_MIN = 600;
  const BAND_MAX = 660;
  const MIN_SHARED_WAVELENGTHS = 40;
  const REFERENCE_VERSION = 3;
  const CLASS_LABELS = ['fresh', 'pest', 'rotten'];
  const STORAGE_KEY = 'chiliscan-absorbance-binary-reference-v1';
  // 頁面各狀態訊息的繁中與英文版本集中在此，方便切換語言時重新顯示動態內容。
  const messages = {
    pageTitle: ['ChiliScan｜三類品質初步篩選實驗版', 'ChiliScan | Three-Class Preliminary Screening Experiment'],
    brandSuffix: ['／ 三類品質初步篩選實驗版', '/ Three-Class Preliminary Screening Experiment'],
    backToSpectrum: ['返回光譜分析', 'Back to Spectral Analysis'],
    factoryIncomplete: ['工廠端改版尚未完成資料追溯與量測檢查，請勿用於實際工廠判定。', 'The factory redesign lacks complete data traceability and measurement checks. Do not use it for factory decisions.'],
    intro: [
      '使用 600–660 nm 衍生吸光度參考曲線進行新鮮、蟲害與腐爛比對。所有結果都需要人工複核，尚未完成獨立驗證；分析在瀏覽器進行，使用者可選擇儲存至雲端。不能作為食品安全或產品放行依據。',
      'Compare derived absorbance spectra at 600–660 nm against fresh, pest-damaged and rotten references. All results require manual review; independent validation is incomplete. Analysis runs in the browser; saving to the cloud is optional. Do not use results for food-safety or product-release decisions.'
    ],
    referenceTitle: ['01　載入參考資料（首次使用或更新時）', '01  Load Reference Data (first use or when updating)'],
    referenceHint: [
      '選擇含「新鮮」、「蟲害」與「腐爛」工作表的衍生吸光度參考檔，資料會儲存在此瀏覽器。每類至少 3 筆、至少 40 個共同波長是程式條件，不是驗證標準。',
      'Choose a derived-absorbance reference file with Fresh, Pest and Rotten worksheets. It is saved in this browser. At least 3 samples per class and 40 shared wavelengths are program requirements, not validated standards.'
    ],
    chooseReference: ['選擇三類衍生吸光度參考檔', 'Choose Three-Class Derived Absorbance Reference File'],
    fileTypes: ['支援 .xlsx、.xls、.csv', 'Supports .xlsx, .xls, and .csv'],
    loadReference: ['載入並儲存參考資料', 'Load and Save Reference Data'],
    referenceVersion: ['參考資料版本（使用者提供）', 'Reference data version (provided by user)'],
    referenceCorrection: ['參考資料校正方式（使用者確認）', 'Reference correction method (confirmed by user)'],
    queryCorrection: ['待測資料校正方式（使用者確認）', 'Query correction method (confirmed by user)'],
    chooseCorrection: ['請選擇', 'Choose a method'],
    whiteCorrection: ['白板比值，未做暗訊號校正', 'White-reference ratio, no dark correction'],
    darkCorrection: ['白板與暗訊號校正', 'White and dark correction'],
    clearReference: ['清除本機參考資料', 'Clear Local Reference Data'],
    referenceNotLoaded: ['尚未載入參考資料。', 'Reference data has not been loaded.'],
    queryTitle: ['02　上傳待測資料', '02  Upload Query Data'],
    queryHint: [
      '每個工作表需有「波長 (nm)」欄，只比對 600–660 nm。原始光強需有「標準白」欄；選擇暗訊號校正時也需要「暗訊號」欄。以校正後的比值 R 計算衍生吸光度 A = −log10(R)，並確認參考與待測資料校正方式一致。',
      'Each worksheet needs a wavelength column; only 600–660 nm is compared. Raw intensity requires a white reference, and dark correction also requires a dark signal. Derived absorbance is A = −log10(R); confirm that reference and query correction methods match.'
    ],
    chooseQuery: ['選擇待測 Excel／CSV', 'Choose Query Excel / CSV'],
    queryFilePlaceholder: ['可包含一個或多個待測樣本', 'One or more query samples can be included'],
    queryModeLabel: ['待測資料格式', 'Query Data Format'],
    modeAuto: ['自動：有標準白欄就換算，否則視為吸光度', 'Auto: Convert if a white reference column exists; otherwise treat as absorbance'],
    modeAbsorbance: ['已是衍生吸光度', 'Already Derived Absorbance'],
    modeIntensity: ['原始光強（需有標準白欄）', 'Raw Intensity (requires a white reference column)'],
    modeIntensityDark: ['原始光強（需有標準白及暗訊號欄）', 'Raw Intensity (requires white and dark reference columns)'],
    classify: ['比對三類參考資料', 'Compare Three Reference Classes'],
    queryWaiting: ['請先載入參考資料，再選擇待測檔案。', 'Load reference data, then choose a query file.'],
    resultsTitle: ['03　比對結果', '03  Classification Results'],
    referenceLoading: ['正在載入參考資料…', 'Loading reference data…'],
    referenceLoaded: [
      '參考資料版本：{dataVersion}；校正方式：{correction}。已載入新鮮 {freshCount} 筆、蟲害 {pestCount} 筆、腐爛 {rottenCount} 筆；共同有效波長 {wavelengthCount} 點（{bandMin}–{bandMax} nm）。{ignoredNotice}',
      'Reference version: {dataVersion}; correction: {correction}. Loaded {freshCount} fresh, {pestCount} pest-damaged and {rottenCount} rotten samples; {wavelengthCount} shared valid wavelengths ({bandMin}–{bandMax} nm). {ignoredNotice}'
    ],
    referenceRestored: [
      '已從此瀏覽器載入參考資料版本 {dataVersion}；校正方式：{correction}。新鮮 {freshCount} 筆、蟲害 {pestCount} 筆、腐爛 {rottenCount} 筆；共同有效波長 {wavelengthCount} 點。',
      'Reference version {dataVersion} loaded from this browser; correction: {correction}. {freshCount} fresh, {pestCount} pest-damaged and {rottenCount} rotten samples; {wavelengthCount} shared valid wavelengths.'
    ],
    ignoredSheetsNotice: ['未使用工作表：{names}。', 'Ignored worksheets: {names}.'],
    referenceCleared: ['已清除此瀏覽器儲存的參考資料。', 'Reference data saved in this browser has been cleared.'],
    referenceLoadFailed: ['載入失敗：{detail}', 'Load failed: {detail}'],
    referenceRestoreFailed: ['無法載入本機參考資料：{detail}', 'Could not load local reference data: {detail}'],
    referenceClearFailed: ['無法清除參考資料：{detail}', 'Could not clear reference data: {detail}'],
    referenceStorageError: ['參考資料無法儲存到此瀏覽器：{detail}', 'Could not save reference data in this browser: {detail}'],
    validationStatus: [
      '參考資料留一內部測試：{correct}/{count}；新鮮 {freshCorrect}/{freshCount}、蟲害 {pestCorrect}/{pestCount}、腐爛 {rottenCorrect}/{rottenCount}。尚未完成獨立驗證。',
      'Internal leave-one-out reference test: {correct}/{count}; fresh {freshCorrect}/{freshCount}, pest {pestCorrect}/{pestCount}, rotten {rottenCorrect}/{rottenCount}. Independent validation has not been completed.'
    ],
    resultWarning: [
      '各類以最近 RMSE 的倒數正規化為相對相似度；零距離類別平均分配 100%。數值不是準確率或機率。尚未建立接近程度的判定門檻，所有結果均須人工複核；尚未完成獨立驗證。',
      'Relative similarity normalizes inverse nearest-class RMSE; classes with zero distance share 100% equally. It is not accuracy or probability. No closeness threshold has been established, so every result requires manual review. Independent validation has not been completed.'
    ],
    missingWavelengthHeader: ['{sheet}：找不到「波長」欄位標題。', '{sheet}: wavelength column header not found.'],
    intensityNeedsWhite: ['{sheet}：原始光強資料必須在同一工作表包含「標準白」欄。', '{sheet}: raw intensity data requires a “White Reference” column on the same worksheet.'],
    intensityNeedsDark: ['{sheet}：白板與暗訊號校正需要「暗訊號」欄。', '{sheet}: white and dark correction requires a dark-reference column.'],
    correctionRequired: ['請填寫參考資料版本並選擇校正方式。', 'Provide a reference data version and select its correction method.'],
    incompatibleCorrection: ['請確認待測與參考資料使用相同校正方式；自動模式無法確認方法，請改用明確格式。', 'Confirm the query and references use the same correction method. Auto mode cannot establish provenance; choose an explicit format.'],
    referenceMustBeAbsorbance: ['{sheet}：參考檔應是已換算的衍生吸光度，不要含標準白或暗訊號光強欄。', '{sheet}: reference data must be precomputed derived absorbance without white or dark intensity columns.'],
    noSampleColumn: ['{sheet}：沒有可用的樣本欄。', '{sheet}: no usable sample column found.'],
    duplicateWavelength: ['{sheet}：{wavelength} nm 有重複波長。', '{sheet}: duplicate wavelength {wavelength} nm.'],
    mismatchedWavelengths: ['{sheet}：重複量測的波長欄不一致，無法依列號平均；請先依實際波長整理資料。', '{sheet}: replicate wavelength columns differ and cannot be averaged by row; align the data by actual wavelength first.'],
    noUsableSampleValues: ['{sheet}：沒有找到 600–660 nm 的可用樣本數值。', '{sheet}: no usable sample values found in 600–660 nm.'],
    minimumClassSamples: [
      '參考資料不足／無法判定：每類至少 3 筆；目前新鮮 {freshCount}、蟲害 {pestCount}、腐爛 {rottenCount}。',
      'Insufficient reference data / unable to classify: at least 3 samples per class are required; currently fresh {freshCount}, pest {pestCount}, rotten {rottenCount}.'
    ],
    minimumCommonWavelengths: [
      '所有參考樣本共同有效波長只有 {count} 點，至少需要 {minimum} 點。',
      'Reference samples have only {count} shared valid wavelengths; at least {minimum} are required.'
    ],
    nonFiniteDistance: [
      '光譜數值範圍過大，無法計算 RMSE 相似度。',
      'Spectrum values are too large to compute RMSE similarity.'
    ],
    invalidStoredReference: ['本機參考資料格式無效，請重新載入參考檔。', 'Local reference data is invalid. Reload the reference file.'],
    outdatedStoredReference: [
      '舊版參考資料無法用於三類比對，請重新載入涵蓋 600–660 nm 的三類原始參考檔。',
      'Previous reference data cannot be used for three-class comparison. Reload a three-class reference file covering 600–660 nm.'
    ],
    queryProcessing: ['正在本機讀取並比對吸光度資料…', 'Reading and comparing absorbance data in this browser…'],
    queryFileReady: ['待測檔案已選擇，可開始比對。', 'Query file selected. Ready to classify.'],
    queryNoSamples: ['沒有找到可用的待測樣本欄。', 'No usable query sample columns found.'],
    queryNoSamplesWithSheets: [
      '找不到含「波長」欄位的待測工作表（略過：{sheets}）。',
      'No query worksheets with a wavelength column were found (ignored: {sheets}).'
    ],
    queryNeedsWavelength: [
      '請提供「波長 (nm)」欄及各樣本數值欄；原始光強另需「標準白」欄。',
      'Provide a “Wavelength (nm)” column and sample value columns. Raw intensity data also requires a “White Reference” column.'
    ],
    insufficientWavelengths: [
      '只有 {count} 個共同有效波長；至少需要 {minimum} 點。',
      'Only {count} shared valid wavelengths; at least {minimum} are required.'
    ],
    queryComplete: [
      '完成：{classified}/{total} 個樣本可比對，均需人工複核；{insufficient} 個樣本資料不足。分析在瀏覽器進行。{ignoredNotice}',
      'Complete: {classified}/{total} samples compared, all require manual review; {insufficient} had insufficient data. Analysis runs in this browser. {ignoredNotice}'
    ],
    queryFailed: ['無法判斷：{detail}', 'Classification failed: {detail}'],
    fresh: ['新鮮', 'Fresh'],
    pest: ['蟲害', 'Pest-damaged'],
    rotten: ['腐爛', 'Rotten'],
    cannotClassify: ['無法判斷', 'Unable to classify'],
    provisionalLabel: ['需人工複核；最接近參考類別：{label}', 'Manual review required; nearest reference class: {label}'],
    similarityTitle: ['相對相似度（不是準確率／機率）', 'Relative similarity (not accuracy / probability)'],
    freshSimilarity: ['新鮮：{rate}%', 'Fresh: {rate}%'],
    pestSimilarity: ['蟲害：{rate}%', 'Pest-damaged: {rate}%'],
    rottenSimilarity: ['腐爛：{rate}%', 'Rotten: {rate}%'],
    sharedWavelengths: ['有效共同波長：{count}/{total} 點', 'Shared valid wavelengths: {count}/{total} points'],
    partialWavelengths: ['部分波長缺失或無效，未納入比對，請檢查原始資料與校正欄位。', 'Some wavelengths were missing or invalid and were excluded; check the input and correction columns.']
  };
  let language = localStorage.getItem('chiliscan-language') === 'en' ? 'en' : 'zh';
  const $ = (id) => document.getElementById(id);

  let referenceSet = null;
  let queryFile = null;
  let latestResults = null;
  let referenceSheets = null;
  let querySheets = null;
  let referenceStatusState = { key: 'referenceNotLoaded', values: {}, isError: false };
  let queryStatusState = { key: 'queryWaiting', values: {}, isError: false };

  // 依目前語言格式化訊息；錯誤也保留翻譯代碼，切換語言後可以重新呈現。
  function formatMessage(key, values = {}) {
    const template = messages[key]?.[language === 'en' ? 1 : 0] || key;
    return template.replace(/\{(\w+)\}/g, (_, name) => {
      const value = values[name];
      if (value && value.translationKey) {
        return formatMessage(value.translationKey, value.translationValues);
      }
      if (value instanceof Error) return value.message;
      if (Array.isArray(value)) return value.join(language === 'en' ? ', ' : '、');
      return value === undefined || value === null ? '' : String(value);
    });
  }

  function localizedError(key, values = {}) {
    const error = new Error(formatMessage(key, values));
    error.translationKey = key;
    error.translationValues = values;
    return error;
  }

  function showReferenceState(key, values = {}, isError = false) {
    referenceStatusState = { key, values, isError };
    const displayValues = { ...values };
    displayValues.ignoredNotice = values.ignoredSheets?.length
      ? formatMessage('ignoredSheetsNotice', { names: values.ignoredSheets })
      : '';
    $('referenceStatus').textContent = formatMessage(key, displayValues);
    $('referenceStatus').dataset.state = isError ? 'error' : 'ready';
  }

  function showQueryState(key, values = {}, isError = false) {
    queryStatusState = { key, values, isError };
    const displayValues = { ...values };
    displayValues.ignoredNotice = values.ignoredSheets?.length
      ? formatMessage('ignoredSheetsNotice', { names: values.ignoredSheets })
      : '';
    $('queryStatus').textContent = formatMessage(key, displayValues);
    $('queryStatus').dataset.state = isError ? 'error' : '';
  }

  function applyLanguage(nextLanguage = language) {
    language = nextLanguage === 'en' ? 'en' : 'zh';
    localStorage.setItem('chiliscan-language', language);
    document.documentElement.lang = language === 'en' ? 'en' : 'zh-Hant';
    document.title = formatMessage('pageTitle');
    document.querySelectorAll('[data-i18n]').forEach((element) => {
      element.textContent = formatMessage(element.dataset.i18n);
    });
    document.querySelectorAll('.language-switch button[data-lang]').forEach((button) => {
      const selected = button.dataset.lang === language;
      button.setAttribute('aria-pressed', String(selected));
      button.classList.toggle('active', selected);
    });
    $('referenceFilename').textContent =
      $('referenceFile').files[0]?.name || formatMessage('fileTypes');
    $('queryFilename').textContent =
      queryFile?.name || formatMessage('queryFilePlaceholder');
    showReferenceState(referenceStatusState.key, referenceStatusState.values, referenceStatusState.isError);
    showQueryState(queryStatusState.key, queryStatusState.values, queryStatusState.isError);
    if (referenceSet) renderValidation(referenceSet.validation);
    if (latestResults) renderResults(latestResults);
  }

  // 光譜檔解析：將儲存格轉成有限數值，辨識波長／標準白欄，並取得各樣本吸光度。
  function numberOrNull(value) {
    if (typeof value === 'number') return Number.isFinite(value) ? value : null;
    if (typeof value !== 'string' || value.trim() === '') return null;
    const parsed = Number(value.trim().replace(/,/g, ''));
    return Number.isFinite(parsed) ? parsed : null;
  }

  function cellValue(sheet, rowIndex, columnIndex) {
    const cell = sheet[XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })];
    if (!cell || cell.t === 'e' || cell.v === undefined || cell.v === null) return null;
    return numberOrNull(cell.v);
  }

  function cellText(sheet, rowIndex, columnIndex) {
    const cell = sheet[XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })];
    if (!cell || cell.t === 'e' || cell.v === undefined || cell.v === null) return '';
    return String(cell.v).trim();
  }

  function isWavelengthHeader(value) {
    return /波長|wavelength|lambda/i.test(String(value || ''));
  }

  function isWhiteHeader(value) {
    return /標準白|白光|white\s*(reference|ref)?/i.test(String(value || ''));
  }

  function isDarkHeader(value) {
    return /暗訊號|暗電流|dark\s*(reference|signal|ref)?/i.test(String(value || ''));
  }

  function isMetadataHeader(value) {
    return /積分|integration|時間|備註|quality|標籤|sample[_ ]?id/i.test(String(value || ''));
  }

  function findHeaderRow(rows) {
    return rows.findIndex((row) =>
      Array.isArray(row) &&
      row.filter((value) => value !== null && value !== undefined && String(value).trim() !== '').length >= 2 &&
      row.some(isWavelengthHeader)
    );
  }

  function findWavelengthSeriesColumns(sheet, headerRowIndex, wavelengthIndex, columnCount, rowCount) {
    const matches = new Set([wavelengthIndex]);
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
      if (columnIndex === wavelengthIndex) continue;
      let eligible = 0;
      let matching = 0;
      for (let rowIndex = headerRowIndex + 1; rowIndex < rowCount; rowIndex++) {
        const wavelength = cellValue(sheet, rowIndex, wavelengthIndex);
        if (wavelength === null || wavelength < BAND_MIN || wavelength > BAND_MAX) continue;
        eligible++;
        const candidate = cellValue(sheet, rowIndex, columnIndex);
        if (candidate !== null && Math.abs(candidate - wavelength) <= 0.01) matching++;
      }
      if (eligible >= MIN_SHARED_WAVELENGTHS && matching / eligible >= 0.9) {
        matches.add(columnIndex);
      }
    }
    return matches;
  }

  function classifySheetName(name) {
    if (/新鮮|fresh/i.test(name)) return 'fresh';
    if (/蟲害|pest/i.test(name)) return 'pest';
    if (/腐爛|腐敗|rotten/i.test(name)) return 'rotten';
    return '';
  }

  function labelName(label) {
    return formatMessage(label);
  }

  function readSpectraFromSheet(sheetName, sheet, mode, label = '') {
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
    const headerRowIndex = findHeaderRow(rows);
    if (headerRowIndex < 0) throw localizedError('missingWavelengthHeader', { sheet: sheetName });

    const headers = rows[headerRowIndex].map((value) => String(value ?? '').trim());
    const wavelengthIndex = headers.findIndex(isWavelengthHeader);
    const whiteIndex = headers.findIndex(isWhiteHeader);
    const darkIndex = headers.findIndex(isDarkHeader);
    const convertIntensity = mode === 'intensity' || mode === 'intensity-dark' || (mode === 'auto' && whiteIndex >= 0);

    if (convertIntensity && whiteIndex < 0) {
      throw localizedError('intensityNeedsWhite', { sheet: sheetName });
    }
    if (mode === 'intensity-dark' && darkIndex < 0) {
      throw localizedError('intensityNeedsDark', { sheet: sheetName });
    }
    if (label && (whiteIndex >= 0 || darkIndex >= 0)) {
      throw localizedError('referenceMustBeAbsorbance', { sheet: sheetName });
    }

    const wavelengthSeriesColumns = findWavelengthSeriesColumns(
      sheet,
      headerRowIndex,
      wavelengthIndex,
      headers.length,
      rows.length
    );
    // 避免將位置相同但實際波長不同的重複量測當作同一點。
    for (const columnIndex of wavelengthSeriesColumns) {
      if (columnIndex === wavelengthIndex) continue;
      for (let rowIndex = headerRowIndex + 1; rowIndex < rows.length; rowIndex++) {
        const wavelength = cellValue(sheet, rowIndex, wavelengthIndex);
        const other = cellValue(sheet, rowIndex, columnIndex);
        if (wavelength !== null && other !== null && wavelength !== other) {
          throw localizedError('mismatchedWavelengths', { sheet: sheetName });
        }
      }
    }
    const sampleHeaders = headers
      .map((header, index) => ({ header, index }))
      .filter(({ header, index }) =>
        header &&
        index !== wavelengthIndex &&
        index !== whiteIndex &&
        index !== darkIndex &&
        !isMetadataHeader(header)
      );
    if (!sampleHeaders.length) throw localizedError('noSampleColumn', { sheet: sheetName });
    const sampleColumns = sampleHeaders.map((column, sampleIndex) => {
      const groupEnd = sampleHeaders[sampleIndex + 1]?.index ?? headers.length;
      // 樣本標題若位於光強欄，第一筆量測也要參與平均。
      const valueIndexes = wavelengthSeriesColumns.has(column.index) ? [] : [column.index];
      for (let index = column.index; index < groupEnd; index++) {
        if (wavelengthSeriesColumns.has(index) && index + 1 < groupEnd) {
          valueIndexes.push(index + 1);
        }
      }
      return {
        ...column,
        valueIndexes: valueIndexes.length ? valueIndexes : [column.index]
      };
    });

    const samples = [];
    for (const column of sampleColumns) {
      const points = new Map();
      for (let rowIndex = headerRowIndex + 1; rowIndex < rows.length; rowIndex++) {
        const wavelength = cellValue(sheet, rowIndex, wavelengthIndex);
        if (wavelength === null || wavelength < BAND_MIN || wavelength > BAND_MAX) continue;

        const values = column.valueIndexes
          .map((index) => cellValue(sheet, rowIndex, index))
          .filter((value) => value !== null);
        if (!values.length) continue;
        const value = values.reduce((sum, item) => sum + item, 0) / values.length;

        let absorbance = value;
        if (convertIntensity) {
          const white = cellValue(sheet, rowIndex, whiteIndex);
          const dark = mode === 'intensity-dark' ? cellValue(sheet, rowIndex, darkIndex) : 0;
          if (white === null || dark === null || white - dark <= 0 || value - dark <= 0) continue;
          absorbance = -Math.log10((value - dark) / (white - dark));
        }
        if (!Number.isFinite(absorbance)) continue;
        if (points.has(wavelength)) throw localizedError('duplicateWavelength', { sheet: sheetName, wavelength });
        points.set(wavelength, absorbance);
      }

      if (points.size) {
        samples.push({
          name: `${sheetName}／${column.header}`,
          label,
          points
        });
      }
    }
    if (!samples.length) throw localizedError('noUsableSampleValues', { sheet: sheetName });
    return { samples, converted: convertIntensity };
  }

  // 分類與驗證：把共同有效波長轉成向量，以 RMSE 選出最近參考樣本並計算留一驗證。
  function getCommonWavelengths(samples) {
    if (!samples.length) return [];
    return [...samples[0].points.keys()]
      .filter((wavelength) => samples.every((sample) => sample.points.has(wavelength)))
      .sort((left, right) => left - right);
  }

  function makeVector(sample, wavelengths) {
    return wavelengths.map((wavelength) => sample.points.get(wavelength));
  }

  function rmse(left, right, rightIndexes = null) {
    let squaredError = 0;
    for (let index = 0; index < left.length; index++) {
      const rightIndex = rightIndexes ? rightIndexes[index] : index;
      const difference = left[index] - right[rightIndex];
      squaredError += difference * difference;
    }
    return Math.sqrt(squaredError / left.length);
  }

  // 以最小距離除以各距離，再正規化；零距離類別平均分配相似度。
  function getRelativeSimilarities(distances) {
    if (CLASS_LABELS.some((label) => !Number.isFinite(distances[label]))) {
      throw localizedError('nonFiniteDistance');
    }
    const zeroLabels = CLASS_LABELS.filter((label) => distances[label] === 0);
    if (zeroLabels.length) {
      return Object.fromEntries(CLASS_LABELS.map((label) => [label, zeroLabels.includes(label) ? 100 / zeroLabels.length : 0]));
    }
    const minimum = Math.min(...CLASS_LABELS.map((label) => distances[label]));
    const weights = CLASS_LABELS.map((label) => minimum / distances[label]);
    const total = weights.reduce((sum, weight) => sum + weight, 0);
    return Object.fromEntries(CLASS_LABELS.map((label, index) => [label, 100 * weights[index] / total]));
  }

  function predictByClassSimilarity(vector, samples, excludedIndex = -1, featureIndexes = null) {
    const distances = { fresh: Infinity, pest: Infinity, rotten: Infinity };
    let nearestSample = null;
    for (let index = 0; index < samples.length; index++) {
      if (index === excludedIndex) continue;
      const sample = samples[index];
      const distance = rmse(vector, sample.values, featureIndexes);
      if (!Number.isFinite(distance)) throw localizedError('nonFiniteDistance');
      if (distance < distances[sample.label]) distances[sample.label] = distance;
      if (!nearestSample || distance < nearestSample.distance) nearestSample = { sample, distance };
    }

    const similarities = getRelativeSimilarities(distances);
    const label = nearestSample.sample.label;
    return { label, similarities };
  }

  function validateReferences(samples) {
    const freshCount = samples.filter((sample) => sample.label === 'fresh').length;
    const pestCount = samples.filter((sample) => sample.label === 'pest').length;
    const rottenCount = samples.filter((sample) => sample.label === 'rotten').length;
    if (freshCount < 3 || pestCount < 3 || rottenCount < 3) {
      throw localizedError('minimumClassSamples', { freshCount, pestCount, rottenCount });
    }

    const wavelengths = getCommonWavelengths(samples);
    if (wavelengths.length < MIN_SHARED_WAVELENGTHS) {
      throw localizedError('minimumCommonWavelengths', { count: wavelengths.length, minimum: MIN_SHARED_WAVELENGTHS });
    }
    return wavelengths;
  }

  function leaveOneOutValidation(samples) {
    const totals = {
      fresh: { correct: 0, count: 0 },
      pest: { correct: 0, count: 0 },
      rotten: { correct: 0, count: 0 }
    };
    samples.forEach((sample, index) => {
      const prediction = predictByClassSimilarity(sample.values, samples, index);
      totals[sample.label].count++;
      if (prediction.label === sample.label) totals[sample.label].correct++;
    });
    const correct = CLASS_LABELS.reduce((sum, label) => sum + totals[label].correct, 0);
    const count = CLASS_LABELS.reduce((sum, label) => sum + totals[label].count, 0);
    return { totals, correct, count };
  }

  // 參考資料生命週期：檢查 localStorage 格式、載入／還原參考曲線並顯示驗證結果。
  function validStoredReference(value) {
    return value &&
      value.version === REFERENCE_VERSION &&
      typeof value.dataVersion === 'string' && value.dataVersion.trim() !== '' &&
      ['white', 'dark'].includes(value.correction) &&
      Array.isArray(value.wavelengths) &&
      value.wavelengths.length >= MIN_SHARED_WAVELENGTHS &&
      value.wavelengths.every((wavelength) => Number.isFinite(wavelength) && wavelength >= BAND_MIN && wavelength <= BAND_MAX) &&
      Array.isArray(value.samples) &&
      value.samples.length >= 9 &&
      value.samples.every((sample) =>
        sample &&
        typeof sample.name === 'string' &&
        CLASS_LABELS.includes(sample.label) &&
        Array.isArray(sample.values) &&
        sample.values.length === value.wavelengths.length &&
        sample.values.every(Number.isFinite)
      ) &&
      value.samples.filter((sample) => sample.label === 'fresh').length >= 3 &&
      value.samples.filter((sample) => sample.label === 'pest').length >= 3 &&
      value.samples.filter((sample) => sample.label === 'rotten').length >= 3;
  }

  function renderValidation(validation) {
    const fresh = validation.totals.fresh;
    const pest = validation.totals.pest;
    const rotten = validation.totals.rotten;
    $('validationStatus').textContent = formatMessage('validationStatus', {
      correct: validation.correct,
      count: validation.count,
      accuracy: (100 * validation.correct / validation.count).toFixed(1),
      freshCorrect: fresh.correct,
      freshCount: fresh.count,
      freshRate: (100 * fresh.correct / fresh.count).toFixed(1),
      pestCorrect: pest.correct,
      pestCount: pest.count,
      rottenCorrect: rotten.correct,
      rottenCount: rotten.count,
      rottenRate: (100 * rotten.correct / rotten.count).toFixed(1)
    });
    $('validationStatus').hidden = false;
  }

  function updateButtons() {
    $('classify').disabled = !referenceSet || !queryFile;
    $('clearReference').disabled = !referenceSet;
  }

  async function loadReferenceFile(file) {
    const dataVersion = $('referenceVersion').value.trim();
    const correction = $('referenceCorrection').value;
    if (!dataVersion || !['white', 'dark'].includes(correction)) throw localizedError('correctionRequired');
    const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    const samples = [];
    const ignoredSheets = [];
    for (const sheetName of workbook.SheetNames) {
      const label = classifySheetName(sheetName);
      if (!label) {
        ignoredSheets.push(sheetName);
        continue;
      }
      const parsed = readSpectraFromSheet(sheetName, workbook.Sheets[sheetName], 'absorbance', label);
      samples.push(...parsed.samples);
    }

    const wavelengths = validateReferences(samples);
    const vectors = samples.map((sample) => ({
      name: sample.name,
      label: sample.label,
      values: makeVector(sample, wavelengths)
    }));
    const validation = leaveOneOutValidation(vectors);
    const nextReference = { version: REFERENCE_VERSION, dataVersion, correction, wavelengths, samples: vectors, validation };

    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(nextReference));
    } catch (error) {
      throw localizedError('referenceStorageError', { detail: error.message || String(error) });
    }

    referenceSet = nextReference;
    referenceSheets = workbook.SheetNames.map((name) => ({
      name, rows: XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: null, raw: true })
    }));
    renderValidation(validation);
    const freshCount = vectors.filter((sample) => sample.label === 'fresh').length;
    const pestCount = vectors.filter((sample) => sample.label === 'pest').length;
    const rottenCount = vectors.filter((sample) => sample.label === 'rotten').length;
    showReferenceState('referenceLoaded', {
      dataVersion,
      correction: { translationKey: correction === 'white' ? 'whiteCorrection' : 'darkCorrection' },
      freshCount,
      pestCount,
      rottenCount,
      wavelengthCount: wavelengths.length,
      bandMin: BAND_MIN,
      bandMax: BAND_MAX,
      ignoredSheets
    });
    updateButtons();
  }

  function restoreReference() {
    try {
      const stored = localStorage.getItem(STORAGE_KEY);
      if (!stored) return;
      const parsed = JSON.parse(stored);
      // 舊版本只儲存舊波段向量，無法補回新波段缺少的波長。
      if (parsed && (parsed.version === 1 || parsed.version === 2)) throw localizedError('outdatedStoredReference');
      if (!validStoredReference(parsed)) throw localizedError('invalidStoredReference');
      const validation = leaveOneOutValidation(parsed.samples);
      referenceSet = { ...parsed, validation };
      renderValidation(validation);
      const freshCount = parsed.samples.filter((sample) => sample.label === 'fresh').length;
      const pestCount = parsed.samples.filter((sample) => sample.label === 'pest').length;
      const rottenCount = parsed.samples.filter((sample) => sample.label === 'rotten').length;
      showReferenceState('referenceRestored', {
        dataVersion: parsed.dataVersion,
        correction: { translationKey: parsed.correction === 'white' ? 'whiteCorrection' : 'darkCorrection' },
        freshCount,
        pestCount,
        rottenCount,
        wavelengthCount: parsed.wavelengths.length
      });
      updateButtons();
    } catch (error) {
      showReferenceState('referenceRestoreFailed', { detail: error }, true);
    }
  }

  function renderResults(results) {
    latestResults = results;
    const validation = referenceSet.validation;
    const fresh = validation.totals.fresh;
    const rotten = validation.totals.rotten;
    $('resultWarning').textContent = formatMessage('resultWarning', {
      correct: validation.correct,
      count: validation.count,
      accuracy: (100 * validation.correct / validation.count).toFixed(1),
      freshCorrect: fresh.correct,
      freshCount: fresh.count,
      freshRate: (100 * fresh.correct / fresh.count).toFixed(1),
      rottenCorrect: rotten.correct,
      rottenCount: rotten.count,
      rottenRate: (100 * rotten.correct / rotten.count).toFixed(1)
    });
    $('resultList').innerHTML = results.map((result) => {
      if (result.error) {
        return `<article class="result-card"><h3>${escapeHtml(result.name)}</h3><span class="result-label rotten">${formatMessage('cannotClassify')}</span><p>${escapeHtml(formatMessage(result.error.translationKey, result.error.translationValues))}</p></article>`;
      }
      const displayedLabel = labelName(result.prediction.label);
      const className = result.prediction.label === 'rotten' ? 'result-label rotten' : 'result-label';
      const similarities = result.prediction.similarities;
      return `<article class="result-card">
        <h3>${escapeHtml(result.name)}</h3>
        <span class="${className}">${formatMessage('provisionalLabel', { label: displayedLabel })}</span>
        <div class="result-similarity">
          <p class="result-similarity-title">${formatMessage('similarityTitle')}</p>
          <div class="result-similarity-values">
            <span>${formatMessage('freshSimilarity', { rate: similarities.fresh.toFixed(1) })}</span>
            <span>${formatMessage('pestSimilarity', { rate: similarities.pest.toFixed(1) })}</span>
            <span>${formatMessage('rottenSimilarity', { rate: similarities.rotten.toFixed(1) })}</span>
          </div>
        </div>
        <p>${formatMessage('sharedWavelengths', { count: result.sharedCount, total: referenceSet.wavelengths.length })}</p>
        ${result.sharedCount < referenceSet.wavelengths.length ? `<p>${formatMessage('partialWavelengths')}</p>` : ''}
      </article>`;
    }).join('');
    $('resultsPanel').hidden = false;
  }

  async function classifyFile() {
    if (!referenceSet || !queryFile) return;
    showQueryState('queryProcessing');
    $('classify').disabled = true;
    $('resultsPanel').hidden = true;
    latestResults = null;
    querySheets = null;
    try {
      const workbook = XLSX.read(await queryFile.arrayBuffer(), { type: 'array' });
      const mode = $('queryMode').value;
      const correction = $('queryCorrection').value;
      if (correction !== referenceSet.correction || mode === 'auto' ||
          (mode === 'intensity' && correction !== 'white') ||
          (mode === 'intensity-dark' && correction !== 'dark')) {
        throw localizedError('incompatibleCorrection');
      }
      const querySamples = [];
      const ignoredSheets = [];
      for (const sheetName of workbook.SheetNames) {
        try {
          const parsed = readSpectraFromSheet(sheetName, workbook.Sheets[sheetName], mode);
          querySamples.push(...parsed.samples);
        } catch (error) {
          if (error.translationKey === 'missingWavelengthHeader') {
            ignoredSheets.push(sheetName);
            continue;
          }
          throw error;
        }
      }
      if (!querySamples.length) {
        const detail = ignoredSheets.length
          ? localizedError('queryNoSamplesWithSheets', { sheets: ignoredSheets })
          : localizedError('queryNoSamples');
        throw localizedError('queryNeedsWavelength', { detail });
      }

      const referenceIndexes = new Map(referenceSet.wavelengths.map((wavelength, index) => [wavelength, index]));
      const results = querySamples.map((sample) => {
        const sharedWavelengths = referenceSet.wavelengths.filter((wavelength) => sample.points.has(wavelength));
        if (sharedWavelengths.length < MIN_SHARED_WAVELENGTHS) {
          return {
            name: sample.name,
            error: localizedError('insufficientWavelengths', {
              count: sharedWavelengths.length,
              minimum: MIN_SHARED_WAVELENGTHS
            })
          };
        }
        const sharedIndexes = sharedWavelengths.map((wavelength) => referenceIndexes.get(wavelength));
        const queryVector = sharedWavelengths.map((wavelength) => sample.points.get(wavelength));
        const prediction = predictByClassSimilarity(queryVector, referenceSet.samples, -1, sharedIndexes);
        return { name: sample.name, sharedCount: sharedWavelengths.length, prediction };
      });

      const classifiedCount = results.filter((result) => !result.error).length;
      querySheets = workbook.SheetNames.map((name) => ({
        name, rows: XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: null, raw: true })
      }));
      showQueryState('queryComplete', {
        classified: classifiedCount,
        total: results.length,
        insufficient: results.length - classifiedCount,
        ignoredSheets
      });
      renderResults(results);
    } catch (error) {
      showQueryState('queryFailed', { detail: error }, true);
    } finally {
      updateButtons();
    }
  }

  // 瀏覽器事件：連接檔案選取、參考資料管理、語言切換與分類按鈕。
  function escapeHtml(value) {
    return String(value).replace(/[&<>"']/g, (character) => ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    })[character]);
  }

  $('referenceFile').addEventListener('change', (event) => {
    const file = event.target.files[0];
    $('referenceFilename').textContent = file ? file.name : formatMessage('fileTypes');
    $('loadReference').disabled = !file;
  });

  $('loadReference').addEventListener('click', async () => {
    const file = $('referenceFile').files[0];
    if (!file) return;
    $('loadReference').disabled = true;
    referenceSet = null;
    referenceSheets = null;
    querySheets = null;
    latestResults = null;
    $('validationStatus').hidden = true;
    $('resultsPanel').hidden = true;
    updateButtons();
    showReferenceState('referenceLoading');
    try {
      await loadReferenceFile(file);
    } catch (error) {
      showReferenceState('referenceLoadFailed', { detail: error }, true);
    } finally {
      $('loadReference').disabled = false;
    }
  });

  $('clearReference').addEventListener('click', () => {
    try {
      localStorage.removeItem(STORAGE_KEY);
      referenceSet = null;
      referenceSheets = null;
      querySheets = null;
      latestResults = null;
      $('validationStatus').hidden = true;
      $('resultsPanel').hidden = true;
      showReferenceState('referenceCleared');
      updateButtons();
    } catch (error) {
      showReferenceState('referenceClearFailed', { detail: error }, true);
    }
  });

  $('queryFile').addEventListener('change', (event) => {
    queryFile = event.target.files[0] || null;
    latestResults = null;
    querySheets = null;
    $('resultsPanel').hidden = true;
    $('queryFilename').textContent = queryFile ? queryFile.name : formatMessage('queryFilePlaceholder');
    showQueryState(queryFile ? 'queryFileReady' : 'queryWaiting');
    updateButtons();
  });

  ['queryMode', 'queryCorrection'].forEach((id) => {
    $(id).addEventListener('change', () => {
      latestResults = null;
      $('resultsPanel').hidden = true;
    });
  });

  function cloudAdapterError(zh, en) {
    const error = new Error(language === 'en' ? en : zh);
    error.cloudMessage = [zh, en];
    return error;
  }

  window.chiliCloudAdapter = {
    page: 'classification',
    capture() {
      if (!referenceSet || !referenceSheets) {
        throw cloudAdapterError('請先重新載入原始參考檔，再儲存至雲端。', 'Reload the original reference file before saving.');
      }
      return {
        version: 1,
        name: queryFile && latestResults ? queryFile.name : referenceSet.dataVersion,
        reference: { dataVersion: referenceSet.dataVersion, correction: referenceSet.correction, sheets: referenceSheets },
        query: latestResults && querySheets ? {
          name: queryFile.name, mode: $('queryMode').value, correction: $('queryCorrection').value, sheets: querySheets
        } : null,
        results: latestResults ? latestResults.map((result) => ({
          name: result.name,
          sharedCount: result.sharedCount ?? null,
          prediction: result.prediction ?? null,
          error: result.error ? { key: result.error.translationKey, values: result.error.translationValues } : null
        })) : null
      };
    },
    async restore(snapshot) {
      const validSheets = (sheets) => Array.isArray(sheets) && sheets.length > 0 &&
        sheets.every((sheet) => sheet && typeof sheet.name === 'string' &&
          Array.isArray(sheet.rows) && sheet.rows.every(Array.isArray)) &&
        new Set(sheets.map((sheet) => sheet.name)).size === sheets.length;
      if (!snapshot || snapshot.version !== 1 || !snapshot.reference ||
          typeof snapshot.reference.dataVersion !== 'string' || !snapshot.reference.dataVersion.trim() ||
          !['white', 'dark'].includes(snapshot.reference.correction) || !validSheets(snapshot.reference.sheets) ||
          (snapshot.query && (typeof snapshot.query.name !== 'string' || !validSheets(snapshot.query.sheets) ||
            !['absorbance', 'intensity', 'intensity-dark'].includes(snapshot.query.mode) ||
            snapshot.query.correction !== snapshot.reference.correction))) {
        throw cloudAdapterError('儲存的比對資料格式無效。', 'Invalid saved classification.');
      }
      const makeFile = (name, sheets) => {
        const workbook = XLSX.utils.book_new();
        sheets.forEach((sheet) => XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(sheet.rows), sheet.name));
        return new File([XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })], name, {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
      };
      $('referenceVersion').value = snapshot.reference.dataVersion;
      $('referenceCorrection').value = snapshot.reference.correction;
      await loadReferenceFile(makeFile('reference.xlsx', snapshot.reference.sheets));
      latestResults = null;
      querySheets = null;
      queryFile = null;
      $('resultsPanel').hidden = true;
      if (snapshot.query) {
        queryFile = makeFile(snapshot.query.name, snapshot.query.sheets);
        $('queryMode').value = snapshot.query.mode;
        $('queryCorrection').value = snapshot.query.correction;
        $('queryFilename').textContent = queryFile.name;
        await classifyFile();
        if (!latestResults) throw cloudAdapterError('無法還原比對結果，請檢查資料。', 'Could not restore classification results.');
      } else {
        $('queryFilename').textContent = formatMessage('queryFilePlaceholder');
        showQueryState('queryWaiting');
        updateButtons();
      }
    }
  };
  document.querySelectorAll('.language-switch button[data-lang]').forEach((button) => {
    button.addEventListener('click', () => applyLanguage(button.dataset.lang));
  });
  $('classify').addEventListener('click', classifyFile);
  applyLanguage(language);
  restoreReference();
})();
