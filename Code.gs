/**
 * 領収書・請求書データ化ツール - バックエンド
 * Google Apps Script (Code.gs)
 * 
 * 【セットアップ手順】
 * 1. GASエディタの「プロジェクトの設定」→「スクリプトプロパティ」で
 *    プロパティ名「GEMINI_API_KEY」にAPIキーを設定してください。
 * 2. 「SHEET_ID」は未設定でも自動作成されます（任意設定可）。
 * 3. 「デプロイ」→「新しいデプロイ」→「ウェブアプリ」として公開してください。
 */

/**
 * Webアプリのエントリポイント（ページルーティング対応）
 */
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'dashboard';
  var pageMap = {
    'dashboard': { file: 'dashboard', title: 'ダッシュボード - 領収書データ化ツール' },
    'tool':      { file: 'index',     title: '領収書・請求書データ化ツール' },
    'clients':   { file: 'clients',   title: '関与先管理 - 領収書データ化ツール' },
    'search':    { file: 'search',    title: 'データ検索 - 領収書データ化ツール' }
  };
  var config = pageMap[page] || pageMap['dashboard'];
  return HtmlService.createHtmlOutputFromFile(config.file)
    .setTitle(config.title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 現在のWebアプリのベースURLを取得
 */
function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

// ==========================================
// ダッシュボード用データ
// ==========================================

/**
 * ダッシュボードに表示するサマリーデータを取得
 * @returns {Object} { totalThisMonth, amountThisMonth, clientCount, recentItems }
 */
function getDashboardData() {
  var now = new Date();
  var ym = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM');
  var result = {
    totalThisMonth: 0,
    amountThisMonth: 0,
    clientCount: 0,
    recentItems: []
  };

  var sheet;
  try { sheet = getOrCreateSheet(); } catch (e) { return result; }
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    // 関与先数はlocalStorageから取得不可なのでスクリプトプロパティから
    result.clientCount = getClientsList().length;
    return result;
  }

  var data = sheet.getRange(2, 1, lastRow - 1, SHEET_HEADERS.length).getValues();
  var clients = {};
  var allItems = [];

  data.forEach(function(row) {
    var rowYm = String(row[11] || '');
    var clientName = String(row[1] || '');
    if (clientName) clients[clientName] = true;

    allItems.push({
      id: row[0],
      clientName: clientName,
      date: row[2] instanceof Date ? Utilities.formatDate(row[2], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(row[2] || ''),
      payee: String(row[3] || ''),
      amount: Number(row[4]) || 0,
      account: String(row[7] || ''),
      createdAt: row[12] instanceof Date ? Utilities.formatDate(row[12], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : String(row[12] || '')
    });

    if (rowYm === ym) {
      result.totalThisMonth++;
      result.amountThisMonth += (Number(row[4]) || 0);
    }
  });

  // 関与先数: スプレッドシートのデータ + スクリプトプロパティの関与先リスト
  var savedClients = getClientsList();
  savedClients.forEach(function(c) { clients[c] = true; });
  result.clientCount = Object.keys(clients).length;

  // 直近500件を取得（新しい順）
  result.recentItems = allItems.reverse().slice(0, 500);

  return result;
}

// ==========================================
// 関与先管理
// ==========================================

/**
 * 関与先リストを取得（スクリプトプロパティに保存）
 * @returns {Array} 関与先名の配列
 */
function getClientsList() {
  var props = PropertiesService.getScriptProperties();
  var json = props.getProperty('CLIENTS_LIST');
  if (json) {
    try { return JSON.parse(json); } catch(e) { }
  }
  return ['A社', 'B社', 'C社']; // デフォルト
}

/**
 * 関与先リストを保存
 * @param {Array} clients - 関与先名の配列
 */
function saveClientsList(clients) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('CLIENTS_LIST', JSON.stringify(clients));
}

/**
 * 関与先ごとの統計情報を取得
 * @returns {Array} [{name, count, totalAmount, lastDate}]
 */
function getClientsWithStats() {
  var clients = getClientsList();
  var stats = {};
  clients.forEach(function(c) {
    stats[c] = { name: c, count: 0, totalAmount: 0, lastDate: '' };
  });

  var sheet;
  try { sheet = getOrCreateSheet(); } catch(e) {
    return clients.map(function(c) { return stats[c]; });
  }

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return clients.map(function(c) { return stats[c]; });
  }

  var data = sheet.getRange(2, 1, lastRow - 1, SHEET_HEADERS.length).getValues();
  data.forEach(function(row) {
    var cn = String(row[1] || '');
    if (!stats[cn]) {
      // CLIENTS_LIST に含まれない関与先は統計のみ記録（リストには追加しない）
      stats[cn] = { name: cn, count: 0, totalAmount: 0, lastDate: '' };
    }
    stats[cn].count++;
    stats[cn].totalAmount += (Number(row[4]) || 0);
    var dt = row[12] instanceof Date
      ? Utilities.formatDate(row[12], Session.getScriptTimeZone(), 'yyyy/MM/dd')
      : String(row[12] || '').substring(0, 10);
    if (dt > stats[cn].lastDate) stats[cn].lastDate = dt;
  });

  return clients.map(function(c) { return stats[c]; });
}

/**
 * 関与先を追加
 * @param {string} name - 関与先名
 * @returns {boolean} 成功/失敗
 */
function addClientServer(name) {
  if (!name || !name.trim()) return false;
  name = name.trim();
  var clients = getClientsList();
  if (clients.indexOf(name) !== -1) return false;
  clients.push(name);
  saveClientsList(clients);
  return true;
}

/**
 * 関与先名を変更
 * @param {string} oldName - 旧名称
 * @param {string} newName - 新名称
 * @returns {boolean} 成功/失敗
 */
function renameClientServer(oldName, newName) {
  if (!oldName || !newName || !newName.trim()) return false;
  newName = newName.trim();
  var clients = getClientsList();
  var idx = clients.indexOf(oldName);
  if (idx === -1) return false;
  if (oldName !== newName && clients.indexOf(newName) !== -1) return false;
  clients[idx] = newName;
  saveClientsList(clients);

  // スプレッドシート内の関与先名も更新
  if (oldName !== newName) {
    try {
      var sheet = getOrCreateSheet();
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var range = sheet.getRange(2, 2, lastRow - 1, 1);
        var values = range.getValues();
        values.forEach(function(row, i) {
          if (String(row[0]) === oldName) values[i][0] = newName;
        });
        range.setValues(values);
      }
    } catch(e) { Logger.log('renameClient シート更新エラー: ' + e.message); }
  }
  return true;
}

/**
 * 関与先を削除（データは残す）
 * @param {string} name - 関与先名
 * @returns {boolean} 成功/失敗
 */
function deleteClientServer(name) {
  var clients = getClientsList();
  var idx = clients.indexOf(name);
  if (idx === -1) return false;
  clients.splice(idx, 1);
  saveClientsList(clients);
  return true;
}

// ==========================================
// Gemini API
// ==========================================

/**
 * Gemini APIを使用して画像/PDFから領収書・請求書データを抽出する
 * @param {string} imageBase64 - Base64エンコードされた画像/PDFデータ
 * @param {string} mimeType - MIMEタイプ (image/jpeg, image/png, application/pdf 等)
 * @returns {Object} 抽出されたデータ {date, payee, amount, account, invoiceNo}
 */
function processImageWithGemini(imageBase64, mimeType) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

  if (!apiKey) {
    throw new Error(
      'GEMINI_API_KEY がスクリプトプロパティに設定されていません。\n' +
      'GASエディタの「プロジェクトの設定」→「スクリプトプロパティ」で設定してください。'
    );
  }

  var model = 'gemini-2.0-flash';
  var url =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    model +
    ':generateContent?key=' +
    apiKey;

  var prompt = [
    'この画像は領収書または請求書です。',
    '以下の情報をJSON形式のみで返してください。',
    '余計なテキストやMarkdownの装飾（```json 等）は絶対に含めないでください。',
    '純粋なJSONオブジェクトのみを返してください。',
    '',
    '{',
    '  "date": "YYYYMMDD形式（8桁の数字、ハイフンなし）。和暦は西暦に変換：令和6年=2024年、令和5年=2023年、令和4年=2022年、令和3年=2021年、令和2年=2020年、令和元年=2019年",',
    '  "payee": "支払先・発行元の会社名や店舗名",',
    '  "payee": "支払先・発行元の会社名や店舗名",',
    '  "amount": 金額を数値のみで入力（カンマなし、税込総額）,',
    '  "amount10": 10%対象の金額を数値のみで入力（カンマなし、税込）、記載がなければ0,',
    '  "amount8": 8%対象の金額を数値のみで入力（カンマなし、税込）、記載がなければ0,',
    '  "account": "最も適切な勘定科目",',
    '  "invoiceNo": "インボイス番号・登録番号（T+13桁の数字、例：T1234567890123）",',
    '  "memo": "取引内容・品目の説明（30字以内の日本語）"',
    '}',
    '',
    '勘定科目は以下から最も適切なものを1つ選んでください：',
    '旅費交通費、消耗品費、接待交際費、通信費、水道光熱費、地代家賃、',
    '事務用品費、広告宣伝費、修繕費、外注費、福利厚生費、租税公課、',
    '保険料、支払手数料、新聞図書費、会議費、車両費、荷造運賃、雑費',
    '',
    '読み取れない項目がある場合：dateは空文字列""、payeeは"不明"、amountは0、accountは"雑費"、invoiceNoは空文字列""、memoは空文字列""としてください。',
    'インボイス番号（適格請求書発行事業者登録番号）が記載されていない場合はinvoiceNoを空文字列""としてください。'
  ].join('\n');

  var payload = {
    contents: [
      {
        parts: [
          { text: prompt },
          {
            inlineData: {
              mimeType: mimeType,
              data: imageBase64
            }
          }
        ]
      }
    ],
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: 1024
    }
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();

    if (responseCode !== 200) {
      Logger.log('Gemini API Error (' + responseCode + '): ' + responseText);
      throw new Error('Gemini API エラー (HTTP ' + responseCode + ')');
    }

    var json = JSON.parse(responseText);

    if (!json.candidates || json.candidates.length === 0) {
      throw new Error('Gemini APIからの応答が空です。画像を確認してください。');
    }

    var text = json.candidates[0].content.parts[0].text;
    Logger.log('Gemini Raw Response: ' + text);

    // JSONを抽出（```json ... ``` のマークダウン記法にも対応）
    var jsonStr = text;
    var codeBlockMatch = text.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (codeBlockMatch) {
      jsonStr = codeBlockMatch[1].trim();
    }

    var jsonMatch = jsonStr.match(/\{[\s\S]*?\}/);
    if (jsonMatch) {
      var result = JSON.parse(jsonMatch[0]);
      // 日付のハイフン除去（念のため）
      var dateStr = String(result.date || '').replace(/-/g, '');
      return {
        date: dateStr,
        payee: String(result.payee || '不明'),
        amount: Number(result.amount) || 0,
        amount10: Number(result.amount10) || 0,
        amount8: Number(result.amount8) || 0,
        account: String(result.account || '雑費'),
        invoiceNo: String(result.invoiceNo || ''),
        memo: String(result.memo || '')
      };
    }

    Logger.log('JSON抽出に失敗: ' + text);
    return { date: '', payee: '不明', amount: 0, amount10: 0, amount8: 0, account: '雑費', invoiceNo: '', memo: '' };

  } catch (error) {
    Logger.log('processImageWithGemini Error: ' + error.message);
    throw new Error('解析エラー: ' + error.message);
  }
}

// ==========================================
// スプレッドシート永続化
// ==========================================

var SHEET_HEADERS = ['ID', '関与先', '日付', '支払先', '金額', '金額(10%)', '金額(8%)', '勘定科目', 'インボイス番号', '摘要', 'ファイル名', '年月', '作成日時', 'ファイルURL'];

/**
 * スプレッドシートを取得（なければ自動作成）
 */
function getOrCreateSheet() {
  var props = PropertiesService.getScriptProperties();
  var sheetId = props.getProperty('SHEET_ID');
  var ss;

  if (sheetId) {
    try {
      ss = SpreadsheetApp.openById(sheetId);
    } catch (e) {
      Logger.log('SHEET_ID のスプレッドシートが見つかりません。新規作成します。');
      ss = null;
    }
  }

  if (!ss) {
    ss = SpreadsheetApp.create('領収書データ化ツール_データ');
    props.setProperty('SHEET_ID', ss.getId());
    Logger.log('新規スプレッドシート作成: ' + ss.getId());
  }

  var sheet = ss.getSheetByName('データ一覧');
  if (!sheet) {
    sheet = ss.getActiveSheet();
    sheet.setName('データ一覧');
    sheet.getRange(1, 1, 1, SHEET_HEADERS.length).setValues([SHEET_HEADERS]);
    sheet.getRange(1, 1, 1, SHEET_HEADERS.length).setFontWeight('bold').setBackground('#e8eaf6');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 60);   // ID
    sheet.setColumnWidth(2, 100);  // 関与先
    sheet.setColumnWidth(3, 100);  // 日付
    sheet.setColumnWidth(4, 180);  // 支払先
    sheet.setColumnWidth(5, 100);  // 金額
    sheet.setColumnWidth(6, 100);  // 金額(10%)
    sheet.setColumnWidth(7, 100);  // 金額(8%)
    sheet.setColumnWidth(8, 120);  // 勘定科目
    sheet.setColumnWidth(9, 180);  // インボイス番号
    sheet.setColumnWidth(10, 160); // 摘要
    sheet.setColumnWidth(11, 200); // ファイル名
    sheet.setColumnWidth(12, 80);  // 年月
    sheet.setColumnWidth(13, 160); // 作成日時
    sheet.setColumnWidth(14, 200); // ファイルURL
  }

  return sheet;
}

/**
 * データをGoogleドライブに保存
 */
function saveFileToDrive(data) {
  var ROOT_FOLDER_ID = '0AGL-c1wNelXkUk9PVA';
  try {
    var rootFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);
  } catch (e) {
    Logger.log('Drive folder not found: ' + ROOT_FOLDER_ID + ' - ' + e.message);
    return '';
  }

  var clientName = data.clientName || '未分類';
  var yearMonth = data.yearMonth || '不明';

  // 関与先フォルダを取得または作成
  var clientFolder;
  var cFolders = rootFolder.getFoldersByName(clientName);
  if (cFolders.hasNext()) {
    clientFolder = cFolders.next();
  } else {
    clientFolder = rootFolder.createFolder(clientName);
  }

  // 年月フォルダを取得または作成
  var ymFolder;
  var yFolders = clientFolder.getFoldersByName(yearMonth);
  if (yFolders.hasNext()) {
    ymFolder = yFolders.next();
  } else {
    ymFolder = clientFolder.createFolder(yearMonth);
  }

  // ファイル名を生成: YYYYMMDD_支払先_金額.ext
  var dateForFile = (data.date || '日付不明').replace(/-/g, '');
  var payee = String(data.payee || '不明').replace(/[\\/:*?"<>|]/g, '_');
  var amount = data.amount || 0;
  var ext = data.ext || 'jpg';
  var baseFileName = dateForFile + '_' + payee + '_' + amount;
  
  // 重複回避
  var fileName = baseFileName + '.' + ext;
  var existingFiles = ymFolder.getFilesByName(fileName);
  var counter = 1;
  while (existingFiles.hasNext()) {
    counter++;
    fileName = baseFileName + '_' + counter + '.' + ext;
    existingFiles = ymFolder.getFilesByName(fileName);
  }

  var mimeType = data.mimeType;
  if (!mimeType) {
    if (ext === 'pdf') mimeType = 'application/pdf';
    else if (ext === 'png') mimeType = 'image/png';
    else mimeType = 'image/jpeg';
  }

  var blob;
  if (data.fileBase64) {
    var decoded = Utilities.base64Decode(data.fileBase64);
    blob = Utilities.newBlob(decoded, mimeType, fileName);
  } else {
    blob = Utilities.newBlob('', mimeType, fileName);
  }

  var file = ymFolder.createFile(blob);
  return file.getUrl();
}

/**
 * データをスプレッドシートに保存
 * @param {Object} data - 保存するデータ
 * @returns {string} 生成されたID
 */
function saveToSheet(data) {
  var sheet = getOrCreateSheet();
  var id = Utilities.getUuid().substring(0, 8);
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  var fileUrl = '';
  if (data.fileBase64) {
    try {
      fileUrl = saveFileToDrive(data);
    } catch (e) {
      Logger.log('Drive save error: ' + e.message);
    }
  }

  // スプレッドシートのヘッダー更新処理（既存シートにファイルURL列がない場合を考慮）
  var lastCol = sheet.getLastColumn();
  if (lastCol < SHEET_HEADERS.length) {
    sheet.getRange(1, 1, 1, SHEET_HEADERS.length).setValues([SHEET_HEADERS]);
    sheet.getRange(1, 1, 1, SHEET_HEADERS.length).setFontWeight('bold').setBackground('#e8eaf6');
    sheet.setColumnWidth(14, 200);
  }

  var row = [
    id,
    data.clientName || '',
    data.date || '',
    data.payee || '',
    data.amount || 0,
    data.amount10 || 0,
    data.amount8 || 0,
    data.account || '',
    data.invoiceNo || '',
    data.memo || '',
    data.fileName || '',
    data.yearMonth || '',
    now,
    fileUrl
  ];

  sheet.appendRow(row);

  // 金額列をフォーマット
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 5, 1, 3).setNumberFormat('#,##0');

  return id;
}

/**
 * 関与先ごとの保存データを取得
 * @param {string} clientName - 関与先名
 * @returns {Array} データ配列
 */
function getSheetData(clientName) {
  var sheet;
  try {
    sheet = getOrCreateSheet();
  } catch (e) {
    return [];
  }

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, SHEET_HEADERS.length).getValues();
  var results = [];

  data.forEach(function (row) {
    if (row[1] === clientName) {
      results.push({
        id: row[0],
        clientName: row[1],
        date: row[2] instanceof Date ? Utilities.formatDate(row[2], Session.getScriptTimeZone(), 'yyyyMMdd') : String(row[2] || ''),
        payee: String(row[3] || ''),
        amount: Number(row[4]) || 0,
        amount10: Number(row[5]) || 0,
        amount8: Number(row[6]) || 0,
        account: String(row[7] || ''),
        invoiceNo: String(row[8] || ''),
        memo: String(row[9] || ''),
        fileName: String(row[10] || ''),
        yearMonth: String(row[11] || ''),
        createdAt: row[12] instanceof Date ? Utilities.formatDate(row[12], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : String(row[12] || ''),
        fileUrl: String(row[13] || '')
      });
    }
  });

  return results;
}

/**
 * 保存データをすべて取得（検索用）
 * @returns {Array} データ配列
 */
function getAllSheetData() {
  var sheet;
  try {
    sheet = getOrCreateSheet();
  } catch (e) {
    return [];
  }

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, SHEET_HEADERS.length).getValues();
  var results = [];

  data.forEach(function (row) {
    results.push({
      id: row[0],
      clientName: row[1],
      date: row[2] instanceof Date ? Utilities.formatDate(row[2], Session.getScriptTimeZone(), 'yyyyMMdd') : String(row[2] || ''),
      payee: String(row[3] || ''),
      amount: Number(row[4]) || 0,
      amount10: Number(row[5]) || 0,
      amount8: Number(row[6]) || 0,
      account: String(row[7] || ''),
      invoiceNo: String(row[8] || ''),
      memo: String(row[9] || ''),
      fileName: String(row[10] || ''),
      yearMonth: String(row[11] || ''),
      createdAt: row[12] instanceof Date ? Utilities.formatDate(row[12], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : String(row[12] || ''),
      fileUrl: String(row[13] || '')
    });
  });

  return results;
}

/**
 * スプレッドシートから行を削除
 * @param {string} rowId - 削除するデータのID
 */
function deleteFromSheet(rowId) {
  var sheet = getOrCreateSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]) === String(rowId)) {
      sheet.deleteRow(i + 2);
      return;
    }
  }
}

/**
 * 担当者リストと関与先の紐付けを取得
 * @returns {Object} 担当者データ { "担当者名": ["A社", "B社"] }
 */
function getStaffList() {
  var props = PropertiesService.getScriptProperties();
  var json = props.getProperty('STAFF_LIST');
  return json ? JSON.parse(json) : {};
}

/**
 * 担当者リストと関与先の紐付けを保存
 * @param {Object} staffObj - 担当者データ
 */
function saveStaffList(staffObj) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('STAFF_LIST', JSON.stringify(staffObj));
}

/**
 * スプレッドシートのURLを取得
 * @returns {string} スプレッドシートURL
 */
/**
 * 画像/PDFから日付のみを抽出する（軽量版）
 * @param {string} imageBase64 - Base64エンコードされた画像/PDFデータ
 * @param {string} mimeType - MIMEタイプ
 * @returns {string} YYYYMMDD形式の日付、または空文字列
 */
function processDateOnly(imageBase64, mimeType) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が設定されていません');

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + apiKey;

  var prompt = [
    'この画像から「日付」だけを読み取り、YYYYMMDD形式（8桁の数字のみ、ハイフンなし）で返してください。',
    '余計なテキストは一切不要です。数字8桁のみを返してください。',
    '',
    '変換例：',
    '・2024年3月15日 → 20240315',
    '・令和6年3月15日 → 20240315（令和6年=2024年）',
    '・令和5年10月1日 → 20231001（令和5年=2023年）',
    '・R6.3.15 → 20240315',
    '・2024/03/15 → 20240315',
    '・2024-03-15 → 20240315',
    '・24.3.15 → 20240315',
    '',
    '日付が読み取れない場合は空文字列のみを返してください。'
  ].join('\n');

  var payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inlineData: { mimeType: mimeType, data: imageBase64 } }
      ]
    }],
    generationConfig: { temperature: 0, maxOutputTokens: 32 }
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) throw new Error('HTTP ' + response.getResponseCode());
    var json = JSON.parse(response.getContentText());
    if (!json.candidates || !json.candidates[0]) return '';
    var text = json.candidates[0].content.parts[0].text.trim();
    Logger.log('processDateOnly raw: ' + text);
    var m = text.match(/\d{8}/);
    return m ? m[0] : '';
  } catch (e) {
    Logger.log('processDateOnly error: ' + e.message);
    throw new Error('日付解析エラー: ' + e.message);
  }
}


/**
 * 画像/PDFから指定フィールドのみを抽出する汎用関数
 * @param {string} imageBase64 - Base64エンコードされた画像/PDFデータ
 * @param {string} mimeType - MIMEタイプ
 * @param {string} fieldType - 抽出フィールド ('payee' | 'amount' | 'memo' | 'invoiceNo')
 * @returns {string} 抽出値
 */
function processFieldOnly(imageBase64, mimeType, fieldType) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が設定されていません');

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + apiKey;

  var prompts = {
    payee: 'この画像から店舗名・会社名（支払先・発行元）のみを読み取り、その名前だけを返してください。余計なテキストは不要です。読み取れなければ「不明」と返してください。',
    amount: 'この画像から合計金額（税込総額）を数字のみで返してください。カンマ・消費税記号不要。数字のみ。読み取れなければ 0 のみ。',
    amount10: 'この画像から「10%対象」の金額（税込）を数字のみで返してください。カンマ不要。記載がなければ 0 のみ。',
    amount8: 'この画像から「8%対象（軽減税率）」の金額（税込）を数字のみで返してください。カンマ不要。記載がなければ 0 のみ。',
    memo: 'このレシート・領収書の取引内容をで2 0字以内の日本語で簡潔に説明してください。例：飲食代（3名）、文具購入、タクシー代。説明文のみを返してください。',
    invoiceNo: 'この画像からインボイス番号（適格請求書発行事業者登録番号）を読み取り、T+13桁の数字のみを返してください。例：T1234567890123。記載がなければ空文字列のみ。'
  };

  var prompt = prompts[fieldType];
  if (!prompt) throw new Error('不明なfieldType: ' + fieldType);

  var payload = {
    contents: [{
      parts: [{ text: prompt }, { inlineData: { mimeType: mimeType, data: imageBase64 } }]
    }],
    generationConfig: { temperature: 0, maxOutputTokens: 128 }
  };
  var options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };

  try {
    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) throw new Error('HTTP ' + response.getResponseCode());
    var json = JSON.parse(response.getContentText());
    if (!json.candidates || !json.candidates[0]) return '';
    var text = json.candidates[0].content.parts[0].text.trim();
    Logger.log('processFieldOnly [' + fieldType + '] raw: ' + text);
    if (fieldType === 'amount' || fieldType === 'amount10' || fieldType === 'amount8') {
      var m = text.match(/\d[\d,]*/);
      return m ? m[0].replace(/,/g, '') : '0';
    }
    if (fieldType === 'invoiceNo') {
      var tm = text.match(/T\d{13}/);
      return tm ? tm[0] : '';
    }
    return text;
  } catch (e) {
    Logger.log('processFieldOnly error: ' + e.message);
    throw new Error('取得エラー: ' + e.message);
  }
}

/**
 * 国税庁インボイス制度適格請求書発行事業者公表サイトから登録状況を確認
 * @param {string} invoiceNo - Tから始まる13桁の登録番号
 * @returns {Object} { isValid: boolean, name: string, message: string }
 */
function checkInvoiceNumber(invoiceNo) {
  if (!invoiceNo) return { isValid: false, name: '', message: '番号が未入力です' };
  
  // Tを除外した13桁の数字を取得
  var numberOnly = invoiceNo.replace(/[^0-9]/g, '');
  if (numberOnly.length !== 13) {
    return { isValid: false, name: '', message: '13桁の数字で入力してください' };
  }

  var url = 'https://www.invoice-kohyo.nta.go.jp/regno-search/detail?selRegNo=' + numberOnly;
  
  try {
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var html = response.getContentText('UTF-8');
    
    // エラーメッセージがあるか確認
    if (html.indexOf('検索対象の登録番号は存在しません') !== -1) {
      return { isValid: false, name: '', message: '登録が確認できませんでした' };
    }
    if (html.indexOf('エラー情報') !== -1 && html.indexOf('メッセージ') !== -1) {
       return { isValid: false, name: '', message: '無効な番号形式です' };
    }
    
    // 氏名又は名称を取得 (HTML構造から抽出: `<h3>氏名又は名称</h3> <p class="...">株式会社〇〇</p>`)
    var nameMatch = html.match(/氏名又は名称<\/h3>[\s\S]*?<p[^>]*>([^<]+)<\/p>/i);
    
    // 古い構造等のフォールバック
    if (!nameMatch) {
       nameMatch = html.match(/<th[^>]*>氏名又は名称<\/th>\s*<td[^>]*>([^<]+)<\/td>/i);
    }
    if (!nameMatch) {
       nameMatch = html.match(/<td class="sp-mb2"[^>]*>([^<]+)<\/td>/i);
    }
    
    if (nameMatch && nameMatch[1]) {
      var name = nameMatch[1].trim()
          .replace(/&amp;/g, '&')
          .replace(/&lt;/g, '<')
          .replace(/&gt;/g, '>')
          .replace(/&quot;/g, '"');
      return { isValid: true, name: name, message: '有効な登録番号です' };
    }
    
    // htmlとして取得はできたが、パースに失敗した場合
    return { isValid: true, name: '名称取得エラー (登録は有効)', message: '有効な登録番号です' };

  } catch (e) {
    Logger.log('checkInvoiceNumber エラー: ' + e.message);
    return { isValid: false, name: '', message: '確認に失敗しました: ' + e.message };
  }
}

function getSheetUrl() {
  var props = PropertiesService.getScriptProperties();
  var sheetId = props.getProperty('SHEET_ID');
  if (sheetId) {
    return 'https://docs.google.com/spreadsheets/d/' + sheetId;
  }
  return '';
}
