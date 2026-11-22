(() => {
  'use strict';

  ['chiliscan-kumo-reference-library', 'chiliscan-kumo-api-url', 'chiliscan-kumo-api-token']
    .forEach((key) => localStorage.removeItem(key));

  // 本模組處理新鮮／腐爛吸光度頁：讀取參考與待測檔、計算 RMSE 相對相似度並更新畫面。
  // 光譜處理與分類都在瀏覽器執行，不會上傳資料。
  const BAND_MIN = 620;
  const BAND_MAX = 680;
  const MIN_SHARED_WAVELENGTHS = 40;
  const STORAGE_KEY = 'chiliscan-absorbance-binary-reference-v1';
  // 頁面各狀態訊息的繁中與英文版本集中在此，方便切換語言時重新顯示動態內容。
  const messages = {
    pageTitle: ['ChiliScan｜吸光度新鮮／腐爛判斷', 'ChiliScan | Absorbance Fresh / Rotten Classification'],
    brandSuffix: ['／ 新鮮・腐爛判斷', '/ Fresh / Rotten Classification'],
    backToSpectrum: ['返回光譜分析', 'Back to Spectral Analysis'],
    intro: [
      '使用 620–680 nm 吸光度參考曲線進行本機比對。資料只在此瀏覽器處理，不會上傳；此功能為實驗版，結果不是已驗證的食品安全判定。',
      'Compare absorbance spectra against reference curves at 620–680 nm. Data is processed in this browser and is not uploaded. This experimental feature is not a validated food-safety assessment.'
    ],
    referenceTitle: ['01　載入參考資料（首次使用或更新時）', '01  Load Reference Data (first use or when updating)'],
    referenceHint: [
      '選擇含「新鮮」與「腐爛」兩個工作表的吸光度參考檔。載入後會儲存在此瀏覽器，之後可直接上傳待測資料。',
      'Choose an absorbance reference file with “Fresh” and “Rotten” worksheets. The data is saved in this browser so you can upload query data later.'
    ],
    chooseReference: ['選擇新鮮／腐爛吸光度參考檔', 'Choose Fresh / Rotten Absorbance Reference File'],
    fileTypes: ['支援 .xlsx、.xls、.csv', 'Supports .xlsx, .xls, and .csv'],
    loadReference: ['載入並儲存參考資料', 'Load and Save Reference Data'],
    clearReference: ['清除本機參考資料', 'Clear Local Reference Data'],
    referenceNotLoaded: ['尚未載入參考資料。', 'Reference data has not been loaded.'],
    queryTitle: ['02　上傳待測資料', '02  Upload Query Data'],
    queryHint: [
      '每個工作表需有「波長 (nm)」欄；只比對 620–680 nm。樣本可用一欄已平均光譜，或在樣本名稱下放多組「波長／光強」重複量測欄，系統會先依波長平均光強。原始光強資料還需在同一工作表含「標準白」欄，再換算吸光度 A = −log10(樣本平均光強／標準白光強度)。只有數值、沒有波長欄的矩陣無法比對。',
      'Each worksheet needs a “Wavelength (nm)” column; only 620–680 nm is compared. Use one column for an averaged spectrum, or place repeated “wavelength / intensity” measurements under a sample name; intensities are averaged by wavelength. Raw intensity data also needs a “White Reference” column on the same worksheet to calculate absorbance A = −log10(mean sample intensity / white reference intensity). A matrix without wavelength values cannot be compared.'
    ],
    chooseQuery: ['選擇待測 Excel／CSV', 'Choose Query Excel / CSV'],
    queryFilePlaceholder: ['可包含一個或多個待測樣本', 'One or more query samples can be included'],
    queryModeLabel: ['待測資料格式', 'Query Data Format'],
    modeAuto: ['自動：有標準白欄就換算，否則視為吸光度', 'Auto: Convert if a white reference column exists; otherwise treat as absorbance'],
    modeAbsorbance: ['已是吸光度', 'Already Absorbance'],
    modeIntensity: ['原始光強（需有標準白欄）', 'Raw Intensity (requires a white reference column)'],
    classify: ['比對新鮮／腐爛', 'Classify Fresh / Rotten'],
    queryWaiting: ['請先載入參考資料，再選擇待測檔案。', 'Load reference data, then choose a query file.'],
    resultsTitle: ['03　比對結果', '03  Classification Results'],
    referenceLoading: ['正在載入參考資料…', 'Loading reference data…'],
    referenceLoaded: [
      '已載入並儲存：新鮮 {freshCount} 筆、腐爛 {rottenCount} 筆；共同有效波長 {wavelengthCount} 點（{bandMin}–{bandMax} nm）。{ignoredNotice}',
      'Loaded and saved: {freshCount} fresh and {rottenCount} rotten samples; {wavelengthCount} shared valid wavelengths ({bandMin}–{bandMax} nm). {ignoredNotice}'
    ],
    referenceRestored: [
      '已從此瀏覽器載入參考資料：新鮮 {freshCount} 筆、腐爛 {rottenCount} 筆；共同有效波長 {wavelengthCount} 點。',
      'Reference data loaded from this browser: {freshCount} fresh and {rottenCount} rotten samples; {wavelengthCount} shared valid wavelengths.'
    ],
    ignoredSheetsNotice: ['未使用工作表：{names}。', 'Ignored worksheets: {names}.'],
    referenceCleared: ['已清除此瀏覽器儲存的參考資料。', 'Reference data saved in this browser has been cleared.'],
    referenceLoadFailed: ['載入失敗：{detail}', 'Load failed: {detail}'],
    referenceRestoreFailed: ['無法載入本機參考資料：{detail}', 'Could not load local reference data: {detail}'],
    referenceClearFailed: ['無法清除參考資料：{detail}', 'Could not clear reference data: {detail}'],
    referenceStorageError: ['參考資料無法儲存到此瀏覽器：{detail}', 'Could not save reference data in this browser: {detail}'],
    validationStatus: [
      '留一樣本測試（依兩類最近 RMSE 相似度判斷）：整體 {correct}/{count}（{accuracy}%）；新鮮 {freshCorrect}/{freshCount}（{freshRate}%），腐爛 {rottenCorrect}/{rottenCount}（{rottenRate}%）。這只是目前參考資料的內部測試，不代表外部準確率。',
      'Leave-one-out validation (nearest-class RMSE similarity): overall {correct}/{count} ({accuracy}%); fresh {freshCorrect}/{freshCount} ({freshRate}%), rotten {rottenCorrect}/{rottenCount} ({rottenRate}%). This is an internal test of the current reference data, not an external accuracy result.'
    ],
    resultWarning: [
      '參考資料留一交叉驗證：整體準確率 {correct}/{count}（{accuracy}%）；新鮮辨識率 {freshCorrect}/{freshCount}（{freshRate}%），腐爛辨識率 {rottenCorrect}/{rottenCount}（{rottenRate}%）。卡片百分比是依兩類最近參考光譜 RMSE 換算的相對相似度，不是單筆準確率或機率；結果不可單獨作為食品安全依據。',
      'Reference leave-one-out validation: overall accuracy {correct}/{count} ({accuracy}%); fresh detection {freshCorrect}/{freshCount} ({freshRate}%), rotten detection {rottenCorrect}/{rottenCount} ({rottenRate}%). Card percentages are relative similarities derived from the nearest reference RMSE for each class, not per-sample accuracy or probability. Do not rely on them alone for food-safety decisions.'
    ],
    missingWavelengthHeader: ['{sheet}：找不到「波長」欄位標題。', '{sheet}: wavelength column header not found.'],
    intensityNeedsWhite: ['{sheet}：原始光強資料必須在同一工作表包含「標準白」欄。', '{sheet}: raw intensity data requires a “White Reference” column on the same worksheet.'],
    referenceMustBeAbsorbance: ['{sheet}：參考檔應是已換算的吸光度資料，不要再含標準白光強欄。', '{sheet}: reference data must already be converted to absorbance and must not contain a white-reference intensity column.'],
    noSampleColumn: ['{sheet}：沒有可用的樣本欄。', '{sheet}: no usable sample column found.'],
    duplicateWavelength: ['{sheet}：{wavelength} nm 有重複波長。', '{sheet}: duplicate wavelength {wavelength} nm.'],
    noUsableSampleValues: ['{sheet}：沒有找到 620–680 nm 的可用樣本數值。', '{sheet}: no usable sample values found in 620–680 nm.'],
    minimumClassSamples: [
      '至少需要新鮮、腐爛各 3 個參考樣本；目前為 {freshCount} 個新鮮、{rottenCount} 個腐爛。',
      'At least 3 fresh and 3 rotten reference samples are required; currently there are {freshCount} fresh and {rottenCount} rotten.'
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
      '完成：{classified}/{total} 個樣本可比對；{insufficient} 個樣本資料不足。全程在此瀏覽器處理。{ignoredNotice}',
      'Complete: {classified}/{total} samples classified; {insufficient} had insufficient data. All processing stays in this browser. {ignoredNotice}'
    ],
    queryFailed: ['無法判斷：{detail}', 'Classification failed: {detail}'],
    fresh: ['新鮮', 'Fresh'],
    rotten: ['腐爛', 'Rotten'],
    cannotClassify: ['無法判斷', 'Unable to classify'],
    provisionalLabel: ['較相似類別：{label}', 'Closer match: {label}'],
    similarityTitle: ['相對相似度（不是準確率／機率）', 'Relative similarity (not accuracy / probability)'],
    freshSimilarity: ['新鮮：{rate}%', 'Fresh: {rate}%'],
    rottenSimilarity: ['腐爛：{rate}%', 'Rotten: {rate}%'],
    sharedWavelengths: ['有效共同波長：{count}/{total} 點', 'Shared valid wavelengths: {count}/{total} points']
  };
  let language = localStorage.getItem('chiliscan-language') === 'en' ? 'en' : 'zh';
  const $ = (id) => document.getElementById(id);

  let referenceSet = null;
  let queryFile = null;
  let latestResults = null;
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
    if (/腐爛|腐敗|rotten/i.test(name)) return 'rotten';
    return '';
  }

  function labelName(label) {
    return label === 'fresh' ? formatMessage('fresh') : formatMessage('rotten');
  }

  function readSpectraFromSheet(sheetName, sheet, mode, label = '') {
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
    const headerRowIndex = findHeaderRow(rows);
    if (headerRowIndex < 0) throw localizedError('missingWavelengthHeader', { sheet: sheetName });

    const headers = rows[headerRowIndex].map((value) => String(value ?? '').trim());
    const wavelengthIndex = headers.findIndex(isWavelengthHeader);
    const whiteIndex = headers.findIndex(isWhiteHeader);
    const convertIntensity = mode === 'intensity' || (mode === 'auto' && whiteIndex >= 0);

    if (convertIntensity && whiteIndex < 0) {
      throw localizedError('intensityNeedsWhite', { sheet: sheetName });
    }
    if (label && whiteIndex >= 0) {
      throw localizedError('referenceMustBeAbsorbance', { sheet: sheetName });
    }

    const wavelengthSeriesColumns = findWavelengthSeriesColumns(
      sheet,
      headerRowIndex,
      wavelengthIndex,
      headers.length,
      rows.length
    );
    const sampleHeaders = headers
      .map((header, index) => ({ header, index }))
      .filter(({ header, index }) =>
        header &&
        index !== wavelengthIndex &&
        index !== whiteIndex &&
        !isMetadataHeader(header)
      );
    if (!sampleHeaders.length) throw localizedError('noSampleColumn', { sheet: sheetName });
    const sampleColumns = sampleHeaders.map((column, sampleIndex) => {
      const groupEnd = sampleHeaders[sampleIndex + 1]?.index ?? headers.length;
      const valueIndexes = [];
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
          if (white === null || white <= 0 || value <= 0) continue;
          absorbance = -Math.log10(value / white);
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

  // 將兩類最近 RMSE 換算為合計 100% 的相對相似度，並非準確率或機率。
  function getRelativeSimilarities(freshDistance, rottenDistance) {
    if (!Number.isFinite(freshDistance) || !Number.isFinite(rottenDistance)) {
      throw localizedError('nonFiniteDistance');
    }
    if (freshDistance === rottenDistance) return { fresh: 50, rotten: 50 };
    if (freshDistance < rottenDistance) {
      const fresh = 100 / (1 + freshDistance / rottenDistance);
      return { fresh, rotten: 100 - fresh };
    }
    const rotten = 100 / (1 + rottenDistance / freshDistance);
    return { fresh: 100 - rotten, rotten };
  }

  function predictByClassSimilarity(vector, samples, excludedIndex = -1, featureIndexes = null) {
    const distances = { fresh: Infinity, rotten: Infinity };
    let nearestSample = null;
    for (let index = 0; index < samples.length; index++) {
      if (index === excludedIndex) continue;
      const sample = samples[index];
      const distance = rmse(vector, sample.values, featureIndexes);
      if (!Number.isFinite(distance)) throw localizedError('nonFiniteDistance');
      if (distance < distances[sample.label]) distances[sample.label] = distance;
      if (!nearestSample || distance < nearestSample.distance) nearestSample = { sample, distance };
    }

    const similarities = getRelativeSimilarities(distances.fresh, distances.rotten);
    // 距離相同時沿用最先找到的最近參考樣本類別。
    const label = distances.fresh < distances.rotten
      ? 'fresh'
      : distances.rotten < distances.fresh ? 'rotten' : nearestSample.sample.label;
    return { label, similarities };
  }

  function validateReferences(samples) {
    const freshCount = samples.filter((sample) => sample.label === 'fresh').length;
    const rottenCount = samples.filter((sample) => sample.label === 'rotten').length;
    if (freshCount < 3 || rottenCount < 3) {
      throw localizedError('minimumClassSamples', { freshCount, rottenCount });
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
      rotten: { correct: 0, count: 0 }
    };
    samples.forEach((sample, index) => {
      const prediction = predictByClassSimilarity(sample.values, samples, index);
      totals[sample.label].count++;
      if (prediction.label === sample.label) totals[sample.label].correct++;
    });
    const correct = totals.fresh.correct + totals.rotten.correct;
    const count = totals.fresh.count + totals.rotten.count;
    return { totals, correct, count };
  }

  // 參考資料生命週期：檢查 localStorage 格式、載入／還原參考曲線並顯示驗證結果。
  function validStoredReference(value) {
    return value &&
      value.version === 1 &&
      Array.isArray(value.wavelengths) &&
      value.wavelengths.length >= MIN_SHARED_WAVELENGTHS &&
      value.wavelengths.every(Number.isFinite) &&
      Array.isArray(value.samples) &&
      value.samples.length >= 6 &&
      value.samples.every((sample) =>
        sample &&
        typeof sample.name === 'string' &&
        ['fresh', 'rotten'].includes(sample.label) &&
        Array.isArray(sample.values) &&
        sample.values.length === value.wavelengths.length &&
        sample.values.every(Number.isFinite)
      ) &&
      value.samples.filter((sample) => sample.label === 'fresh').length >= 3 &&
      value.samples.filter((sample) => sample.label === 'rotten').length >= 3;
  }

  function renderValidation(validation) {
    const fresh = validation.totals.fresh;
    const rotten = validation.totals.rotten;
    $('validationStatus').textContent = formatMessage('validationStatus', {
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
    $('validationStatus').hidden = false;
  }

  function updateButtons() {
    $('classify').disabled = !referenceSet || !queryFile;
    $('clearReference').disabled = !referenceSet;
  }

  async function loadReferenceFile(file) {
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
    const nextReference = { version: 1, wavelengths, samples: vectors, validation };

    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(nextReference));
    } catch (error) {
      throw localizedError('referenceStorageError', { detail: error.message || String(error) });
    }

    referenceSet = nextReference;
    renderValidation(validation);
    const freshCount = vectors.filter((sample) => sample.label === 'fresh').length;
    const rottenCount = vectors.filter((sample) => sample.label === 'rotten').length;
    showReferenceState('referenceLoaded', {
      freshCount,
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
      if (!validStoredReference(parsed)) throw localizedError('invalidStoredReference');
      const validation = leaveOneOutValidation(parsed.samples);
      referenceSet = { ...parsed, validation };
      renderValidation(validation);
      const freshCount = parsed.samples.filter((sample) => sample.label === 'fresh').length;
      const rottenCount = parsed.samples.filter((sample) => sample.label === 'rotten').length;
      showReferenceState('referenceRestored', {
        freshCount,
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
            <span>${formatMessage('rottenSimilarity', { rate: similarities.rotten.toFixed(1) })}</span>
          </div>
        </div>
        <p>${formatMessage('sharedWavelengths', { count: result.sharedCount, total: referenceSet.wavelengths.length })}</p>
      </article>`;
    }).join('');
    $('resultsPanel').hidden = false;
  }

  async function classifyFile() {
    if (!referenceSet || !queryFile) return;
    showQueryState('queryProcessing');
    $('classify').disabled = true;
    $('resultsPanel').hidden = true;
    try {
      const workbook = XLSX.read(await queryFile.arrayBuffer(), { type: 'array' });
      const mode = $('queryMode').value;
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
    $('queryFilename').textContent = queryFile ? queryFile.name : formatMessage('queryFilePlaceholder');
    showQueryState(queryFile ? 'queryFileReady' : 'queryWaiting');
    updateButtons();
  });

  document.querySelectorAll('.language-switch button[data-lang]').forEach((button) => {
    button.addEventListener('click', () => applyLanguage(button.dataset.lang));
  });
  $('classify').addEventListener('click', classifyFile);
  applyLanguage(language);
  restoreReference();
})();
