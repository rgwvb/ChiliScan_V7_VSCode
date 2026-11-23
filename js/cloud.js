/*
 * 個人雲端紀錄：只有使用者明確按下儲存時，才會上傳目前的資料快照。
 * Firestore 規則必須另外設定，前端登入狀態本身不是存取權限。
 */
(() => {
  'use strict';

  const config = {
    apiKey: 'AIzaSyBLOsmkbgqJcChLQmzA1ittqqNOrH6gv3k',
    authDomain: 'chili-pepper--system.firebaseapp.com',
    projectId: 'chili-pepper--system',
    appId: '1:836535969443:web:c400a18a4d352235430116'
  };
  const SDK = 'https://www.gstatic.com/firebasejs/13.4.0';
  const SPREADSHEET_ID = '12IW-9bbp5j31Ih1rJgsjFZhuY7ndBPLrsgU5EnCBIZQ';
  const SPREADSHEET_TAB = 'ChiliScan Records';
  const PICKER_API_KEY = 'AIzaSyAx6cAnTHsaHbKsUeCB04R7mamuuynC7nI';
  const GOOGLE_PROJECT_NUMBER = config.appId.split(':')[1];
  const SHEETS_API = 'https://sheets.googleapis.com/v4';
  const SHEET_HEADERS = [
    '紀錄ID', '項目序號', '項目數', '建立時間', '分析頁', '資料名稱',
    '樣本編號', '量測時間', '時區', '樣本狀態', '備註', '分析摘要'
  ];
  // 只有在管理員部署並測試權限規則後才能開放雲端讀寫。
  const SECURITY_RULES_VERIFIED = true;
  const CHUNK_BYTES = 192 * 1024;
  const MAX_CHUNKS = 24;
  let auth, db, firestore, authSdk, user;
  let authorized = false;
  let recordsOpen = false;
  let sheetsAccessToken = '';
  let selectedSpreadsheetId = '';
  let sheetStatus = ['尚未連結 Google 試算表。', 'Google Sheet is not connected.'];
  let pickerPromise;
  let measurementTimeZone = '';
  let busy = false;
  let status = ['尚未登入。', 'Not signed in.'];

  const panel = document.createElement('section');
  panel.className = 'panel';
  panel.innerHTML = `
    <h2 id="cloudTitle" tabindex="-1" data-cloud="title"></h2>
    <div class="tools"><button class="btn secondary" type="button" id="cloudBack" data-cloud="back"></button></div>
    <p class="note" data-cloud="description"></p>
    <p class="note" data-cloud="manualHint"></p>
    <div class="tools"><button class="btn secondary" type="button" id="cloudSheetConnect" data-cloud="connectSheet"></button></div>
    <p class="status" id="cloudSheetStatus" role="status"></p>
    <div class="tools">
      <label class="field"><span data-cloud="sampleId"></span><input id="cloudSampleId" type="text" maxlength="120"></label>
      <label class="field"><span data-cloud="measurementTime"></span><input id="cloudMeasurementTime" type="datetime-local"></label>
      <label class="field"><span data-cloud="sampleStatus"></span><select id="cloudSampleStatus">
        <option value="" data-cloud="notProvided"></option>
        <option value="undetermined" data-cloud="undetermined"></option>
        <option value="fresh" data-cloud="fresh"></option>
        <option value="pest" data-cloud="pest"></option>
        <option value="rotten" data-cloud="rotten"></option>
      </select></label>
      <label class="field"><span data-cloud="notes"></span><input id="cloudNotes" type="text" maxlength="1000"></label>
    </div>
    <p class="note" id="cloudTimeZone"></p>
    <div class="tools"><button class="btn" type="button" id="cloudSave" data-cloud="save" disabled></button></div>
    <div class="tools">
      <label class="field"><span data-cloud="records"></span><select id="cloudRecords" disabled></select></label>
      <button class="btn secondary" type="button" id="cloudRefresh" data-cloud="refresh" disabled></button>
      <button class="btn secondary" type="button" id="cloudLoad" data-cloud="load" disabled></button>
      <button class="btn secondary" type="button" id="cloudSync" data-cloud="sync" hidden disabled></button>
    </div>
    <p class="status" id="cloudStatus" role="status"></p>`;
  const analysisView = document.getElementById('analysisView');
  analysisView.tabIndex = -1;
  analysisView.after(panel);
  const account = document.createElement('div');
  account.className = 'cloud-account';
  account.innerHTML = `
    <button class="btn secondary" type="button" id="cloudLogin" data-cloud="login"></button>
    <button class="btn secondary" type="button" id="cloudOpen" data-cloud="open" hidden></button>
    <button class="btn secondary" type="button" id="cloudLogout" data-cloud="logout" hidden></button>
    <span class="cloud-account-status" id="cloudAccountStatus" role="status"></span>`;
  document.querySelector('.header-tools').append(account);

  const text = {
    title: ['個人量測紀錄', 'Personal Measurement Records'],
    description: [
      '登入並連結試算表後，手動儲存會將解析後資料寫入 Firebase 與「辣椒工廠」；管理員須先部署 Firestore 安全規則。示範資料不得上傳。',
      'After signing in and connecting the sheet, manual saves write parsed data to Firebase and “辣椒工廠”. An administrator must deploy Firestore rules first. Demo data cannot be uploaded.'
    ],
    manualHint: ['欄位由使用者填寫，樣本狀態不代表分類結果。持有試算表連結的人可看到樣本編號、備註與分析摘要，請勿填入機密資料。', 'Fields are user-provided; sample status is not a classification result. Anyone with the sheet link can see sample IDs, notes, and summaries. Do not enter confidential data.'],
    sampleId: ['樣本編號（選填）', 'Sample ID (optional)'],
    measurementTime: ['量測時間（選填）', 'Measurement time (optional)'],
    sampleStatus: ['樣本狀態（使用者填寫）', 'Sample status (user-provided)'],
    notes: ['備註（選填）', 'Notes (optional)'],
    timeZone: ['量測時間時區：', 'Measurement time zone: '],
    notProvided: ['未提供', 'Not provided'],
    undetermined: ['待判定', 'Undetermined'],
    fresh: ['新鮮', 'Fresh'],
    pest: ['蟲害', 'Pest-damaged'],
    rotten: ['腐爛', 'Rotten'],
    login: ['使用 Google 登入', 'Sign in with Google'],
    open: ['量測紀錄', 'Measurement records'],
    back: ['返回分析', 'Back to analysis'],
    logout: ['登出', 'Sign out'],
    save: ['儲存目前資料', 'Save current data'],
    records: ['已儲存紀錄', 'Saved records'],
    refresh: ['更新清單', 'Refresh'],
    load: ['載入所選紀錄', 'Load selected record'],
    sync: ['同步至試算表', 'Sync to sheet'],
    empty: ['尚無紀錄', 'No records'],
    working: ['處理中，請稍候…', 'Working…'],
    signedOut: ['尚未登入。', 'Not signed in.'],
    signedOutReminder: ['已登出；已載入的分析及本機參考資料仍可能留在瀏覽器。', 'Signed out; loaded analysis and local reference data may remain in the browser.'],
    httpRequired: ['雲端登入需要透過本機預覽網址或網站網址開啟，不能使用 file://。', 'Cloud sign-in requires a local preview or website URL, not file://.'],
    savedBoth: ['已同步儲存至 Firebase 與 Google 試算表。', 'Saved to Firebase and Google Sheets.'],
    savedPending: ['Firebase 已儲存；尚未同步到試算表。連結試算表後，選取這筆紀錄並按「同步至試算表」。', 'Saved to Firebase, but not synced to Sheets. Connect the sheet, select this record, and choose Sync to sheet.'],
    restored: ['已載入雲端紀錄。', 'Cloud record loaded.'],
    savedListFailed: ['紀錄已儲存，但清單更新失敗，請按「更新清單」。', 'The record was saved, but the list could not be refreshed. Use Refresh.'],
    notAuthorized: ['無法讀取雲端紀錄。請確認帳號已獲授權且 Firestore 規則已部署。', 'Cannot read cloud records. Verify that your account is authorized and Firestore rules have been deployed.'],
    setupRequired: ['Google 登入完成，但雲端讀寫尚未啟用；請先完成並測試 Firestore 安全規則。', 'Google sign-in succeeded, but cloud access is disabled until Firestore security rules are deployed and tested.'],
    connectSheet: ['連結 Google 試算表', 'Connect Google Sheet'],
    sheetNotConnected: ['尚未連結試算表。', 'Google Sheet is not connected.'],
    sheetConnected: ['已連結「辣椒工廠」；寫入範圍限於「ChiliScan Records」分頁。', 'Connected to “辣椒工廠”; writes are limited to the “ChiliScan Records” tab.'],
    sheetKeyMissing: ['尚未設定受限制的 Google Picker API key。請依 README 設定後再連結。', 'The restricted Google Picker API key is not configured. Follow the README setup before connecting.'],
    sheetWrongFile: ['請在 Picker 中選擇指定的「辣椒工廠」試算表。', 'Select the specified “辣椒工廠” spreadsheet in Picker.'],
    sheetWrongAccount: ['Google 試算表必須使用目前登入 ChiliScan 的同一個 Google 帳號連結。', 'Connect the sheet with the same Google account currently signed in to ChiliScan.'],
    sheetCancelled: ['已取消試算表連結。', 'Sheet connection was cancelled.'],
    sheetPermission: ['Google 授權或 Sheets API 操作失敗：', 'Google authorization or Sheets API request failed: '],
    sheetStructure: ['試算表分頁表頭不符，為避免覆寫資料已停止同步。', 'The sheet headers do not match; sync stopped to avoid overwriting data.'],
    sheetConflict: ['這筆紀錄在試算表中的摘要資料不一致，已停止重複寫入。', 'This record has conflicting summary rows in the sheet; duplicate writes were stopped.'],
    syncedPending: ['已將待同步紀錄寫入 Google 試算表。', 'Pending record synced to Google Sheets.'],
    tooLarge: ['資料過大，無法一次安全儲存。請分成較小的檔案。', 'Data is too large to save safely in one operation. Split it into smaller files.'],
    invalid: ['雲端紀錄不完整或格式無效。', 'The cloud record is incomplete or invalid.'],
    differentPage: ['此紀錄屬於另一個分析頁，請前往該頁載入。', 'This record belongs to the other analysis page. Load it there.'],
    error: ['操作失敗：', 'Operation failed: ']
  };
  const el = (id) => document.getElementById(id);
  const lang = () => localStorage.getItem('chiliscan-language') === 'en' ? 1 : 0;
  const tr = (key) => text[key][lang()];
  function cloudError(key, detail = '') {
    const error = new Error(tr(key));
    error.cloudMessage = text[key].map((message) => detail ? `${message} (${detail})` : message);
    return error;
  }
  function show(pair) {
    status = pair;
    el('cloudStatus').textContent = pair[lang()];
    el('cloudAccountStatus').textContent = pair[lang()];
  }
  function showSheet(pair) {
    sheetStatus = pair;
    el('cloudSheetStatus').textContent = pair[lang()];
  }
  function translate() {
    document.querySelectorAll('[data-cloud]').forEach((item) => {
      item.textContent = tr(item.dataset.cloud);
    });
    el('cloudStatus').textContent = status[lang()];
    el('cloudAccountStatus').textContent = status[lang()];
    el('cloudSheetStatus').textContent = sheetStatus[lang()];
    if (!el('cloudRecords').value) el('cloudRecords').options[0].textContent = tr('empty');
    [...el('cloudRecords').options].filter((option) => option.dataset.record).forEach((option) => {
      option.textContent = recordLabel(JSON.parse(option.dataset.record));
    });
    updateTimeZone();
  }
  function updateTimeZone() {
    el('cloudTimeZone').textContent = el('cloudMeasurementTime').value
      ? `${tr('timeZone')}${measurementTimeZone || tr('notProvided')}` : '';
  }
  function clearInputs() {
    ['cloudSampleId', 'cloudMeasurementTime', 'cloudSampleStatus', 'cloudNotes'].forEach((id) => { el(id).value = ''; });
    measurementTimeZone = '';
    updateTimeZone();
  }
  function recordLabel(data) {
    const sampleStatus = Object.hasOwn(text, data.sampleStatus) ? tr(data.sampleStatus) : tr('notProvided');
    const time = data.measurementTime ? `${data.measurementTime} ${data.measurementTimeZone || tr('notProvided')}` : tr('notProvided');
    const page = data.page === 'spectrum' ? (lang() ? 'Spectrum' : '光譜分析')
      : data.page === 'classification' ? (lang() ? 'Classification' : '三類比對') : data.page || '';
    return `${data.createdAt || ''} · ${page} · ${data.name || ''} · ${data.sampleId || tr('notProvided')} · ${time} · ${sampleStatus}`;
  }
  function controls() {
    if (!authorized) recordsOpen = false;
    panel.hidden = !authorized || !recordsOpen;
    analysisView.hidden = authorized && recordsOpen;
    el('cloudLogin').hidden = authorized;
    el('cloudOpen').hidden = !authorized || recordsOpen;
    el('cloudOpen').disabled = busy;
    el('cloudBack').disabled = busy;
    el('cloudLogout').hidden = !authorized;
    el('cloudLogin').disabled = busy;
    el('cloudLogout').disabled = busy || !user;
    el('cloudSheetConnect').disabled = busy || !user;
    el('cloudSave').disabled = busy || !user;
    el('cloudRefresh').disabled = busy || !user;
    el('cloudRecords').disabled = busy || !user || el('cloudRecords').options.length < 2;
    el('cloudLoad').disabled = busy || !user || !el('cloudRecords').value;
    const selected = el('cloudRecords').selectedOptions[0];
    const selectedData = selected && selected.dataset.record ? JSON.parse(selected.dataset.record) : null;
    const pending = selectedData && selectedData.sheetSyncStatus !== 'synced';
    el('cloudSync').hidden = !pending;
    el('cloudSync').disabled = busy || !user || !sheetsAccessToken || selectedSpreadsheetId !== SPREADSHEET_ID || !pending;
  }
  function setRecordsOpen(open) {
    recordsOpen = authorized && open;
    controls();
    if (recordsOpen) {
      el('cloudTitle').focus();
    } else {
      window.dispatchEvent(new Event('resize'));
      if (authorized && !busy) el('cloudOpen').focus();
      else analysisView.focus();
    }
  }
  function emptyRecords() {
    const option = document.createElement('option');
    option.value = '';
    option.textContent = tr('empty');
    el('cloudRecords').replaceChildren(option);
  }
  function recordsCollection() {
    return firestore.collection(db, 'users', user.uid, 'chiliscanRecords');
  }
  async function refresh() {
    const snapshot = await firestore.getDocsFromServer(firestore.query(
      recordsCollection(), firestore.orderBy('createdAt', 'desc')
    ));
    emptyRecords();
    snapshot.forEach((doc) => {
      const data = doc.data();
      const option = document.createElement('option');
      option.value = doc.id;
      option.dataset.record = JSON.stringify(data);
      option.textContent = recordLabel(data);
      el('cloudRecords').append(option);
    });
    controls();
  }
  async function initialize() {
    if (auth) return;
    const [appSdk, nextAuthSdk, nextFirestore] = await Promise.all([
      import(`${SDK}/firebase-app.js`),
      import(`${SDK}/firebase-auth.js`),
      import(`${SDK}/firebase-firestore.js`)
    ]);
    const app = appSdk.initializeApp(config);
    const nextAuth = nextAuthSdk.getAuth(app);
    await nextAuthSdk.setPersistence(nextAuth, nextAuthSdk.browserSessionPersistence);
    await nextAuth.authStateReady();
    authSdk = nextAuthSdk;
    firestore = nextFirestore;
    auth = nextAuth;
    db = firestore.getFirestore(app);
  }
  function sheetRange(range) {
    return `'${SPREADSHEET_TAB.replace(/'/g, "''")}'!${range}`;
  }
  async function sheetsRequest(path, options = {}) {
    const response = await fetch(`${SHEETS_API}/spreadsheets/${SPREADSHEET_ID}${path}`, {
      method: options.method || 'GET',
      headers: {
        Authorization: `Bearer ${sheetsAccessToken}`,
        ...(options.body ? { 'Content-Type': 'application/json' } : {})
      },
      ...(options.body ? { body: JSON.stringify(options.body) } : {})
    });
    const result = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(result.error && result.error.message ? result.error.message : `HTTP ${response.status}`);
    }
    return result;
  }
  async function prepareSheet() {
    const workbook = await sheetsRequest('?fields=sheets.properties(sheetId,title)');
    const sheet = (workbook.sheets || []).find((item) => item.properties.title === SPREADSHEET_TAB);
    if (!sheet) {
      await sheetsRequest(':batchUpdate', {
        method: 'POST',
        body: { requests: [{ addSheet: { properties: { title: SPREADSHEET_TAB } } }] }
      });
    }
    const headersRange = sheetRange('A1:L1');
    const current = await sheetsRequest(`/values/${encodeURIComponent(headersRange)}`);
    const headers = current.values && current.values[0] ? current.values[0] : [];
    if (!headers.length) {
      await sheetsRequest(`/values/${encodeURIComponent(headersRange)}?valueInputOption=RAW`, {
        method: 'PUT',
        body: { values: [SHEET_HEADERS] }
      });
    } else if (JSON.stringify(headers) !== JSON.stringify(SHEET_HEADERS)) {
      throw cloudError('sheetStructure');
    }
  }
  function loadPicker() {
    if (pickerPromise) return pickerPromise;
    pickerPromise = new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = 'https://apis.google.com/js/api.js';
      script.async = true;
      script.onload = () => {
        if (!window.gapi || typeof window.gapi.load !== 'function') {
          reject(new Error('Google API client failed to load.'));
          return;
        }
        window.gapi.load('picker', { callback: resolve, onerror: reject });
      };
      script.onerror = () => reject(new Error('Google Picker failed to load.'));
      document.head.append(script);
    }).catch((error) => {
      pickerPromise = null;
      throw error;
    });
    return pickerPromise;
  }
  async function showSpreadsheetPicker(accessToken) {
    await loadPicker();
    return new Promise((resolve, reject) => {
      const picker = new window.google.picker.PickerBuilder()
        .setAppId(GOOGLE_PROJECT_NUMBER)
        .setDeveloperKey(PICKER_API_KEY)
        .setOAuthToken(accessToken)
        .addView(window.google.picker.ViewId.SPREADSHEETS)
        .setCallback((data) => {
          if (data.action === window.google.picker.Action.CANCEL) {
            reject(cloudError('sheetCancelled'));
          } else if (data.action === window.google.picker.Action.PICKED) {
            const selected = data.docs && data.docs[0];
            if (!selected || selected.id !== SPREADSHEET_ID) {
              reject(cloudError('sheetWrongFile'));
            } else {
              resolve(selected);
            }
          }
        })
        .build();
      picker.setVisible(true);
    });
  }
  async function connectSheet() {
    if (PICKER_API_KEY.startsWith('REPLACE_')) throw cloudError('sheetKeyMissing');
    const provider = new authSdk.GoogleAuthProvider();
    provider.addScope('https://www.googleapis.com/auth/drive.file');
    if (user.email) provider.setCustomParameters({ login_hint: user.email });
    const result = await authSdk.signInWithPopup(auth, provider);
    if (result.user.uid !== user.uid) {
      await authSdk.signOut(auth);
      user = null;
      authorized = false;
      setRecordsOpen(false);
      emptyRecords();
      throw cloudError('sheetWrongAccount');
    }
    const credential = authSdk.GoogleAuthProvider.credentialFromResult(result);
    if (!credential || !credential.accessToken) throw cloudError('sheetPermission', 'No Google access token was returned.');
    sheetsAccessToken = credential.accessToken;
    try {
      await showSpreadsheetPicker(sheetsAccessToken);
      selectedSpreadsheetId = SPREADSHEET_ID;
      await prepareSheet();
      showSheet(text.sheetConnected);
      controls();
    } catch (error) {
      sheetsAccessToken = '';
      selectedSpreadsheetId = '';
      throw error;
    }
  }
  function formatSheetNumber(value, digits = 3) {
    return Number.isFinite(value) ? Number(value).toFixed(digits) : '未提供';
  }
  function sampleStatusLabel(value) {
    const labels = {
      undetermined: ['待判定', 'Undetermined'],
      fresh: ['新鮮', 'Fresh'],
      pest: ['蟲害', 'Pest-damaged'],
      rotten: ['腐爛', 'Rotten']
    };
    return labels[value] ? labels[value][lang()] : tr('notProvided');
  }
  function analysisSummaryRows(snapshot) {
    if (snapshot && snapshot.analysis && Array.isArray(snapshot.analysis.features)) {
      const features = snapshot.analysis.features.filter(Boolean);
      if (features.length) {
        return features.map((feature) => ({
          name: feature.name || '',
          summary: `波峰 ${formatSheetNumber(feature.w, 1)} nm / ${formatSheetNumber(feature.max)}；平均 ${formatSheetNumber(feature.mean)}；最低 ${formatSheetNumber(feature.min)}；有效點 ${Number.isInteger(feature.n) ? feature.n : 0}`
        }));
      }
      return [{ name: snapshot.name || '', summary: '沒有可用的光譜特徵。' }];
    }
    if (snapshot && snapshot.reference) {
      const results = Array.isArray(snapshot.results) ? snapshot.results : [];
      if (!results.length) {
        return [{ name: snapshot.name || '', summary: '僅保存參考資料，尚無待測結果。' }];
      }
      return results.map((result) => {
        if (result.error || !result.prediction) {
          return { name: result.name || '', summary: '未能分類；請檢查資料並人工複核。' };
        }
        const similarities = result.prediction.similarities || {};
        const label = {
          fresh: '新鮮',
          pest: '蟲害',
          rotten: '腐爛'
        }[result.prediction.label] || '未判定';
        return {
          name: result.name || '',
          summary: `暫定分類：${label}；相對相似度（非準確率／機率）新鮮 ${formatSheetNumber(similarities.fresh, 1)}%、蟲害 ${formatSheetNumber(similarities.pest, 1)}%、腐爛 ${formatSheetNumber(similarities.rotten, 1)}%；共同波長 ${Number.isInteger(result.sharedCount) ? result.sharedCount : 0} 點`
        };
      });
    }
    throw cloudError('invalid');
  }
  async function writeRecordToSheet(recordId, metadata, summaryRows) {
    if (!sheetsAccessToken || selectedSpreadsheetId !== SPREADSHEET_ID) throw cloudError('sheetNotConnected');
    await prepareSheet();
    const range = sheetRange('A2:C');
    const existing = await sheetsRequest(`/values/${encodeURIComponent(range)}`);
    const rows = existing.values || [];
    const found = [];
    rows.forEach((row, index) => {
      if (row[0] === recordId) {
        found.push({ rowNumber: index + 2, itemIndex: Number(row[1]), itemCount: Number(row[2]) });
      }
    });
    const seen = new Set();
    for (const item of found) {
      if (!Number.isInteger(item.itemIndex) || item.itemIndex < 0 ||
          item.itemIndex >= summaryRows.length || item.itemCount !== summaryRows.length ||
          seen.has(item.itemIndex)) throw cloudError('sheetConflict');
      seen.add(item.itemIndex);
    }
    if (found.length) {
      const query = new URLSearchParams();
      found.forEach((item) => query.append('ranges', sheetRange(`L${item.rowNumber}`)));
      const stored = await sheetsRequest(`/values:batchGet?${query.toString()}`);
      const storedChunks = stored.valueRanges || [];
      if (storedChunks.length !== found.length || storedChunks.some((item, index) =>
        !item.values || !item.values[0] || item.values[0][0] !== summaryRows[found[index].itemIndex].summary
      )) throw cloudError('sheetConflict');
    }
    const missing = summaryRows.map((item, itemIndex) => ({ item, itemIndex }))
      .filter((item) => !seen.has(item.itemIndex));
    if (!missing.length) return;
    const startRow = rows.length + 2;
    const values = missing.map(({ item, itemIndex }) => [
      recordId,
      itemIndex,
      summaryRows.length,
      metadata.createdAt,
      metadata.page === 'spectrum' ? '光譜分析' : '三類品質比對',
      metadata.name || '',
      metadata.sampleId || '',
      metadata.measurementTime || '',
      metadata.measurementTimeZone || '',
      sampleStatusLabel(metadata.sampleStatus),
      metadata.notes || '',
      `${item.name ? `${item.name}：` : ''}${item.summary}`
    ]);
    await sheetsRequest('/values:batchUpdate', {
      method: 'POST',
      body: {
        valueInputOption: 'RAW',
        data: [{
          range: sheetRange(`A${startRow}:L${startRow + values.length - 1}`),
          values
        }]
      }
    });
  }
  async function readRecordBytes(record, data) {
    if (!Number.isInteger(data.chunkCount) || data.chunkCount < 1 || data.chunkCount > MAX_CHUNKS) {
      throw cloudError('invalid');
    }
    const parts = await Promise.all(Array.from({ length: data.chunkCount }, (_, index) =>
      firestore.getDocFromServer(firestore.doc(record, 'chunks', String(index).padStart(4, '0')))
    ));
    if (parts.some((part) => !part.exists() || typeof part.data().value !== 'string')) throw cloudError('invalid');
    const decoded = parts.map((part) => Uint8Array.from(atob(part.data().value), (character) => character.charCodeAt(0)));
    const bytes = new Uint8Array(decoded.reduce((sum, part) => sum + part.length, 0));
    let offset = 0;
    decoded.forEach((part) => { bytes.set(part, offset); offset += part.length; });
    return bytes;
  }
  async function syncRecord(recordId) {
    const record = firestore.doc(recordsCollection(), recordId);
    const document = await firestore.getDocFromServer(record);
    if (!document.exists()) throw cloudError('invalid');
    const metadata = document.data();
    const bytes = await readRecordBytes(record, metadata);
    const snapshot = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes));
    await writeRecordToSheet(recordId, metadata, analysisSummaryRows(snapshot));
    await firestore.updateDoc(record, { sheetSyncStatus: 'synced', sheetSyncedAt: new Date().toISOString() });
  }
  function splitBytes(bytes) {
    const chunks = [];
    for (let offset = 0; offset < bytes.length; offset += CHUNK_BYTES) {
      const part = bytes.subarray(offset, offset + CHUNK_BYTES);
      let binary = '';
      for (let i = 0; i < part.length; i++) binary += String.fromCharCode(part[i]);
      chunks.push(btoa(binary));
    }
    return chunks;
  }
  async function save() {
    const adapter = window.chiliCloudAdapter;
    if (!adapter) throw new Error('Analysis adapter is unavailable.');
    const snapshot = adapter.capture();
    snapshot.measurement = {
      sampleId: el('cloudSampleId').value.trim(),
      time: el('cloudMeasurementTime').value,
      timeZone: el('cloudMeasurementTime').value ? measurementTimeZone : '',
      status: el('cloudSampleStatus').value,
      notes: el('cloudNotes').value.trim()
    };
    const bytes = new TextEncoder().encode(JSON.stringify(snapshot));
    const chunks = splitBytes(bytes);
    if (!chunks.length || chunks.length > MAX_CHUNKS) throw cloudError('tooLarge');
    const record = firestore.doc(recordsCollection());
    const metadata = {
      format: 1,
      page: adapter.page,
      name: snapshot.name || '',
      createdAt: new Date().toISOString(),
      sampleId: snapshot.measurement.sampleId,
      measurementTime: snapshot.measurement.time,
      measurementTimeZone: snapshot.measurement.timeZone,
      sampleStatus: snapshot.measurement.status,
      notes: snapshot.measurement.notes,
      chunkCount: chunks.length,
      sheetSyncStatus: 'pending'
    };
    const batch = firestore.writeBatch(db);
    batch.set(record, metadata);
    chunks.forEach((value, index) => {
      batch.set(firestore.doc(record, 'chunks', String(index).padStart(4, '0')), { value });
    });
    await batch.commit();
    let synced = false;
    let syncError = null;
    if (sheetsAccessToken && selectedSpreadsheetId === SPREADSHEET_ID) {
      try {
        await writeRecordToSheet(record.id, metadata, analysisSummaryRows(snapshot));
        await firestore.updateDoc(record, { sheetSyncStatus: 'synced', sheetSyncedAt: new Date().toISOString() });
        synced = true;
      } catch (error) {
        syncError = error;
      }
    }
    try {
      await refresh();
    } catch (error) {
      throw cloudError('savedListFailed', error.code || error.message);
    }
    el('cloudRecords').value = record.id;
    if (synced) {
      show(text.savedBoth);
    } else if (syncError) {
      const detail = syncError.message || String(syncError);
      show(text.savedPending.map((message) => `${message} (${detail})`));
    } else {
      show(text.savedPending);
    }
  }
  async function load() {
    const record = firestore.doc(recordsCollection(), el('cloudRecords').value);
    const metadata = await firestore.getDocFromServer(record);
    if (!metadata.exists()) throw cloudError('invalid');
    const data = metadata.data();
    if (data.page !== window.chiliCloudAdapter.page) throw cloudError('differentPage');
    if (data.format !== 1) throw cloudError('invalid');
    const bytes = await readRecordBytes(record, data);
    const snapshot = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes));
    await window.chiliCloudAdapter.restore(snapshot);
    const measurement = snapshot.measurement || {};
    el('cloudSampleId').value = measurement.sampleId || '';
    el('cloudMeasurementTime').value = measurement.time || '';
    measurementTimeZone = measurement.timeZone || '';
    updateTimeZone();
    el('cloudSampleStatus').value = ['undetermined', 'fresh', 'pest', 'rotten'].includes(measurement.status) ? measurement.status : '';
    el('cloudNotes').value = measurement.notes || '';
    show(text.restored);
    setRecordsOpen(false);
  }
  async function syncPending() {
    const recordId = el('cloudRecords').value;
    if (!recordId) throw cloudError('invalid');
    await syncRecord(recordId);
    await refresh();
    el('cloudRecords').value = recordId;
    show(text.syncedPending);
  }
  async function run(action) {
    busy = true;
    controls();
    show(text.working);
    try {
      await action();
    } catch (error) {
      const message = error && error.message ? error.message : String(error);
      const pair = error && Array.isArray(error.cloudMessage) ? error.cloudMessage : [message, message];
      show([`${text.error[0]}${pair[0]}`, `${text.error[1]}${pair[1]}`]);
    } finally {
      busy = false;
      controls();
    }
  }
  async function authorize(nextUser) {
    user = nextUser;
    if (!SECURITY_RULES_VERIFIED) {
      await authSdk.signOut(auth);
      user = null;
      throw cloudError('setupRequired');
    }
    try {
      await refresh();
      authorized = true;
      show([`已登入：${user.email || user.uid}`, `Signed in: ${user.email || user.uid}`]);
    } catch (error) {
      await authSdk.signOut(auth);
      user = null;
      authorized = false;
      setRecordsOpen(false);
      clearInputs();
      emptyRecords();
      throw cloudError('notAuthorized', error.code || error.message);
    }
  }
  el('cloudLogin').addEventListener('click', () => run(async () => {
    await initialize();
    const result = await authSdk.signInWithPopup(auth, new authSdk.GoogleAuthProvider());
    await authorize(result.user);
  }));
  el('cloudOpen').addEventListener('click', () => setRecordsOpen(true));
  el('cloudBack').addEventListener('click', () => setRecordsOpen(false));
  el('cloudSheetConnect').addEventListener('click', () => run(async () => {
    try {
      await connectSheet();
    } catch (error) {
      const pair = error && Array.isArray(error.cloudMessage)
        ? error.cloudMessage
        : text.sheetPermission.map((message) => `${message}${error && error.message ? error.message : String(error)}`);
      showSheet(pair);
      throw error;
    }
  }));
  el('cloudLogout').addEventListener('click', () => run(async () => {
    await authSdk.signOut(auth);
    user = null;
    authorized = false;
    sheetsAccessToken = '';
    selectedSpreadsheetId = '';
    setRecordsOpen(false);
    clearInputs();
    emptyRecords();
    showSheet(text.sheetNotConnected);
    show(text.signedOutReminder);
  }));
  el('cloudSave').addEventListener('click', () => run(save));
  el('cloudRefresh').addEventListener('click', () => run(async () => {
    await refresh();
    show([`已登入：${user.email || user.uid}`, `Signed in: ${user.email || user.uid}`]);
  }));
  el('cloudLoad').addEventListener('click', () => run(load));
  el('cloudSync').addEventListener('click', () => run(syncPending));
  el('cloudRecords').addEventListener('change', controls);
  el('cloudMeasurementTime').addEventListener('input', () => {
    measurementTimeZone = el('cloudMeasurementTime').value ? (Intl.DateTimeFormat().resolvedOptions().timeZone || '') : '';
    updateTimeZone();
  });
  document.querySelectorAll('[data-lang]').forEach((button) => button.addEventListener('click', translate));
  emptyRecords();
  translate();
  controls();
  if (!['http:', 'https:'].includes(location.protocol)) {
    show(text.httpRequired);
    el('cloudLogin').disabled = true;
  } else {
    run(async () => {
      await initialize();
      if (auth.currentUser) await authorize(auth.currentUser);
      else show(text.signedOut);
    });
  }
})();
