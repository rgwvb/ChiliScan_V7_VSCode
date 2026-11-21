(() => {
  'use strict';

  const BAND_MIN = 620;
  const BAND_MAX = 680;
  const MIN_SHARED_WAVELENGTHS = 40;
  const NEIGHBOR_COUNT = 3;
  const STORAGE_KEY = 'chiliscan-absorbance-binary-reference-v1';
  const $ = (id) => document.getElementById(id);

  let referenceSet = null;
  let queryFile = null;

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
    return label === 'fresh' ? '新鮮' : '腐爛';
  }

  function readSpectraFromSheet(sheetName, sheet, mode, label = '') {
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
    const headerRowIndex = findHeaderRow(rows);
    if (headerRowIndex < 0) throw new Error(`${sheetName}：找不到「波長」欄位標題。`);

    const headers = rows[headerRowIndex].map((value) => String(value ?? '').trim());
    const wavelengthIndex = headers.findIndex(isWavelengthHeader);
    const whiteIndex = headers.findIndex(isWhiteHeader);
    const convertIntensity = mode === 'intensity' || (mode === 'auto' && whiteIndex >= 0);

    if (convertIntensity && whiteIndex < 0) {
      throw new Error(`${sheetName}：原始光強資料必須在同一工作表包含「標準白」欄。`);
    }
    if (label && whiteIndex >= 0) {
      throw new Error(`${sheetName}：參考檔應是已換算的吸光度資料，不要再含標準白光強欄。`);
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
    if (!sampleHeaders.length) throw new Error(`${sheetName}：沒有可用的樣本欄。`);
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
        if (points.has(wavelength)) throw new Error(`${sheetName}：${wavelength} nm 有重複波長。`);
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
    if (!samples.length) throw new Error(`${sheetName}：沒有找到 620–680 nm 的可用樣本數值。`);
    return { samples, converted: convertIntensity };
  }

  function getCommonWavelengths(samples) {
    if (!samples.length) return [];
    return [...samples[0].points.keys()]
      .filter((wavelength) => samples.every((sample) => sample.points.has(wavelength)))
      .sort((left, right) => left - right);
  }

  function makeVector(sample, wavelengths) {
    return wavelengths.map((wavelength) => sample.points.get(wavelength));
  }

  function rmse(left, right) {
    let squaredError = 0;
    for (let index = 0; index < left.length; index++) {
      const difference = left[index] - right[index];
      squaredError += difference * difference;
    }
    return Math.sqrt(squaredError / left.length);
  }

  function predictByNeighbors(vector, samples, excludedIndex = -1) {
    const neighbors = samples
      .map((sample, index) => ({
        sample,
        index,
        distance: rmse(vector, sample.values)
      }))
      .filter((item) => item.index !== excludedIndex)
      .sort((left, right) => left.distance - right.distance)
      .slice(0, Math.min(NEIGHBOR_COUNT, samples.length - (excludedIndex >= 0 ? 1 : 0)));

    const votes = { fresh: 0, rotten: 0 };
    neighbors.forEach(({ sample }) => votes[sample.label]++);
    const label = votes.fresh === votes.rotten
      ? neighbors[0].sample.label
      : votes.fresh > votes.rotten ? 'fresh' : 'rotten';
    return { label, votes, neighbors };
  }

  function validateReferences(samples) {
    const freshCount = samples.filter((sample) => sample.label === 'fresh').length;
    const rottenCount = samples.filter((sample) => sample.label === 'rotten').length;
    if (freshCount < 3 || rottenCount < 3) {
      throw new Error(`至少需要新鮮、腐爛各 3 個參考樣本；目前為 ${freshCount} 個新鮮、${rottenCount} 個腐爛。`);
    }

    const wavelengths = getCommonWavelengths(samples);
    if (wavelengths.length < MIN_SHARED_WAVELENGTHS) {
      throw new Error(`所有參考樣本共同有效波長只有 ${wavelengths.length} 點，至少需要 ${MIN_SHARED_WAVELENGTHS} 點。`);
    }
    return wavelengths;
  }

  function leaveOneOutValidation(samples) {
    const totals = {
      fresh: { correct: 0, count: 0 },
      rotten: { correct: 0, count: 0 }
    };
    samples.forEach((sample, index) => {
      const prediction = predictByNeighbors(sample.values, samples, index);
      totals[sample.label].count++;
      if (prediction.label === sample.label) totals[sample.label].correct++;
    });
    const correct = totals.fresh.correct + totals.rotten.correct;
    const count = totals.fresh.count + totals.rotten.count;
    return { totals, correct, count };
  }

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
      );
  }

  function showReferenceState(message, isError = false) {
    $('referenceStatus').textContent = message;
    $('referenceStatus').dataset.state = isError ? 'error' : 'ready';
  }

  function renderValidation(validation) {
    const fresh = validation.totals.fresh;
    const rotten = validation.totals.rotten;
    const rate = (item) => `${item.correct}/${item.count}（${(100 * item.correct / item.count).toFixed(1)}%）`;
    $('validationStatus').textContent =
      `留一樣本測試（3 個近鄰）：整體 ${validation.correct}/${validation.count}（${(100 * validation.correct / validation.count).toFixed(1)}%）；` +
      `新鮮 ${rate(fresh)}，腐爛 ${rate(rotten)}。這只是目前參考資料的內部測試，不代表外部準確率。`;
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
      throw new Error(`參考資料無法儲存到此瀏覽器：${error.message || String(error)}`);
    }

    referenceSet = nextReference;
    renderValidation(validation);
    const freshCount = vectors.filter((sample) => sample.label === 'fresh').length;
    const rottenCount = vectors.filter((sample) => sample.label === 'rotten').length;
    let message = `已載入並儲存：新鮮 ${freshCount} 筆、腐爛 ${rottenCount} 筆；共同有效波長 ${wavelengths.length} 點（${BAND_MIN}–${BAND_MAX} nm）。`;
    if (ignoredSheets.length) message += ` 未使用工作表：${ignoredSheets.join('、')}。`;
    showReferenceState(message);
    updateButtons();
  }

  function restoreReference() {
    try {
      const stored = localStorage.getItem(STORAGE_KEY);
      if (!stored) return;
      const parsed = JSON.parse(stored);
      if (!validStoredReference(parsed)) throw new Error('本機參考資料格式無效，請重新載入參考檔。');
      referenceSet = parsed;
      renderValidation(parsed.validation);
      const freshCount = parsed.samples.filter((sample) => sample.label === 'fresh').length;
      const rottenCount = parsed.samples.filter((sample) => sample.label === 'rotten').length;
      showReferenceState(`已從此瀏覽器載入參考資料：新鮮 ${freshCount} 筆、腐爛 ${rottenCount} 筆；共同有效波長 ${parsed.wavelengths.length} 點。`);
      updateButtons();
    } catch (error) {
      showReferenceState(error.message || String(error), true);
    }
  }

  function renderResults(results) {
    const validation = referenceSet.validation;
    const formatRate = (item) =>
      `${item.correct}/${item.count}（${(100 * item.correct / item.count).toFixed(1)}%）`;
    $('resultWarning').textContent =
      `參考資料留一交叉驗證：整體準確率 ${validation.correct}/${validation.count}（${(100 * validation.correct / validation.count).toFixed(1)}%）；` +
      `新鮮辨識率 ${formatRate(validation.totals.fresh)}，腐爛辨識率 ${formatRate(validation.totals.rotten)}。` +
      `這是參考資料的整體測試結果，不代表單筆樣本準確率；目前分類仍依 3 個近鄰判斷，請勿單獨作為食品安全依據。`;
    $('resultList').innerHTML = results.map((result) => {
      if (result.error) {
        return `<article class="result-card"><h3>${escapeHtml(result.name)}</h3><span class="result-label rotten">無法判斷</span><p>${escapeHtml(result.error)}</p></article>`;
      }
      const displayedLabel = labelName(result.prediction.label);
      const className = result.prediction.label === 'rotten' ? 'result-label rotten' : 'result-label';
      const votes = result.prediction.votes;
      const voteCount = votes.fresh + votes.rotten;
      const voteRate = (count) => `${(100 * count / voteCount).toFixed(1)}%`;
      return `<article class="result-card">
        <h3>${escapeHtml(result.name)}</h3>
        <span class="${className}">暫定較接近：${displayedLabel}</span>
        <div class="result-votes">
          <p class="result-votes-title">近鄰投票比例（不是準確率）</p>
          <div class="result-vote-values">
            <span>新鮮：${votes.fresh}/${voteCount}（${voteRate(votes.fresh)}）</span>
            <span>腐爛：${votes.rotten}/${voteCount}（${voteRate(votes.rotten)}）</span>
          </div>
        </div>
        <p>有效共同波長：${result.sharedCount}/${referenceSet.wavelengths.length} 點</p>
      </article>`;
    }).join('');
    $('resultsPanel').hidden = false;
  }

  async function classifyFile() {
    if (!referenceSet || !queryFile) return;
    $('queryStatus').textContent = '正在本機讀取並比對吸光度資料…';
    $('queryStatus').dataset.state = '';
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
          if (String(error.message || error).includes('找不到「波長」欄位標題')) {
            ignoredSheets.push(sheetName);
            continue;
          }
          throw error;
        }
      }
      if (!querySamples.length) {
        const detail = ignoredSheets.length
          ? `找不到含「波長」欄位的待測工作表（略過：${ignoredSheets.join('、')}）。`
          : '沒有找到可用的待測樣本欄。';
        throw new Error(`${detail} 請提供「波長 (nm)」欄及各樣本數值欄；原始光強另需「標準白」欄。`);
      }

      const results = querySamples.map((sample) => {
        const sharedWavelengths = referenceSet.wavelengths.filter((wavelength) => sample.points.has(wavelength));
        if (sharedWavelengths.length < MIN_SHARED_WAVELENGTHS) {
          return {
            name: sample.name,
            error: `只有 ${sharedWavelengths.length} 個共同有效波長；至少需要 ${MIN_SHARED_WAVELENGTHS} 點。`
          };
        }
        const queryVector = sharedWavelengths.map((wavelength) => sample.points.get(wavelength));
        const alignedReferences = referenceSet.samples.map((reference) => ({
          ...reference,
          values: sharedWavelengths.map((wavelength) =>
            referenceSet.wavelengths.indexOf(wavelength) >= 0
              ? reference.values[referenceSet.wavelengths.indexOf(wavelength)]
              : NaN
          )
        }));
        const prediction = predictByNeighbors(queryVector, alignedReferences);
        return { name: sample.name, sharedCount: sharedWavelengths.length, prediction };
      });

      const classifiedCount = results.filter((result) => !result.error).length;
      let status = `完成：${classifiedCount}/${results.length} 個樣本可比對；${results.length - classifiedCount} 個樣本資料不足。全程在此瀏覽器處理。`;
      if (ignoredSheets.length) status += `略過無波長欄的工作表：${ignoredSheets.join('、')}。`;
      $('queryStatus').textContent = status;
      renderResults(results);
    } catch (error) {
      $('queryStatus').textContent = `無法判斷：${error.message || String(error)}`;
      $('queryStatus').dataset.state = 'error';
    } finally {
      updateButtons();
    }
  }

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
    $('referenceFilename').textContent = file ? file.name : '支援 .xlsx、.xls、.csv';
    $('loadReference').disabled = !file;
  });

  $('loadReference').addEventListener('click', async () => {
    const file = $('referenceFile').files[0];
    if (!file) return;
    $('loadReference').disabled = true;
    showReferenceState('正在載入參考資料…');
    try {
      await loadReferenceFile(file);
    } catch (error) {
      showReferenceState(`載入失敗：${error.message || String(error)}`, true);
    } finally {
      $('loadReference').disabled = false;
    }
  });

  $('clearReference').addEventListener('click', () => {
    try {
      localStorage.removeItem(STORAGE_KEY);
      referenceSet = null;
      $('validationStatus').hidden = true;
      $('resultsPanel').hidden = true;
      showReferenceState('已清除此瀏覽器儲存的參考資料。');
      updateButtons();
    } catch (error) {
      showReferenceState(`無法清除參考資料：${error.message || String(error)}`, true);
    }
  });

  $('queryFile').addEventListener('change', (event) => {
    queryFile = event.target.files[0] || null;
    $('queryFilename').textContent = queryFile ? queryFile.name : '可包含一個或多個待測樣本';
    $('queryStatus').textContent = queryFile ? '待測檔案已選擇，可開始比對。' : '請先載入參考資料，再選擇待測檔案。';
    $('queryStatus').dataset.state = '';
    updateButtons();
  });

  $('classify').addEventListener('click', classifyFile);
  restoreReference();
})();
