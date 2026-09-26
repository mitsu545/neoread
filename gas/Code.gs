// ══════════════════════════════════════════════════════
//  NeoRead — GAS Web App
//  ① バックアップ: スプシの「Backup」シートにJSONを1セルで保存・復元
//  ② 文字起こし : 画像をドライブに保存 → Geminiで文字起こし
// ══════════════════════════════════════════════════════
//
//  【セットアップ手順】
//  1. スプシ → 拡張機能 → Apps Script
//  2. このコードを全部貼り付けて保存（Ctrl+S）
//  3. ⚙️プロジェクトの設定 → スクリプト プロパティ に追加
//     - GEMINI_API_KEY : GeminiのAPIキー（必須）
//     - GEMINI_MODEL   : 使うモデル名（任意。未設定なら gemini-flash-latest）
//  4. エディタで「testGemini」を選択して▶実行 → 権限を許可
//  5. デプロイ → デプロイを管理 → ✏️編集 → バージョン「新バージョン」→ デプロイ
//     （※「新しいデプロイ」だとURLが変わるので注意）
//
// ══════════════════════════════════════════════════════

var IMG_FOLDER = 'NeoRead画像';

// POST: action で振り分け（actionなし = 従来のバックアップ）
function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return respond(false, 'No POST data');
  }
  try {
    var raw = e.postData.contents;
    var body = JSON.parse(raw); // JSON検証
    if (body.action === 'ocr') return json(ocr(body));

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Backup') || ss.insertSheet('Backup');
    sheet.getRange('A1').setValue(raw);
    sheet.getRange('B1').setValue(new Date().toISOString());

    return respond(true);
  } catch(err) {
    return respond(false, err.message);
  }
}

// 文字起こし: Gemini成功 → ドライブ保存 の順（失敗時にゴミ画像を残さない）
function ocr(body) {
  if (!body.image) throw new Error('画像がありません');
  var mime = body.mime || 'image/jpeg';
  var r = geminiOcr(body.image, mime);

  var it = DriveApp.getFoldersByName(IMG_FOLDER);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(IMG_FOLDER);
  var name = (body.title || 'book') + '_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss') + '.jpg';
  var file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(body.image), mime, name));

  return { success: true, text: r.text, page: r.page, fileId: file.getId(), url: file.getUrl() };
}

function geminiOcr(base64, mime) {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('GEMINI_API_KEY が未設定です');
  var model = props.getProperty('GEMINI_MODEL') || 'gemini-flash-latest';

  var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': key },
    muteHttpExceptions: true,
    payload: JSON.stringify({
      contents: [{ parts: [
        { text: 'これは本のページの写真です。次の形のJSONだけを出力してください。\n' +
                '{"page": ページ番号の数字（写っていなければ null。見開きなら小さい方）, "text": "本文"}\n' +
                'text には本文をそのまま正確に文字起こしし、ページ番号・柱（ページ上部の章タイトル等）は含めず、段落の区切りは改行で表してください。' },
        { inline_data: { mime_type: mime, data: base64 } }
      ]}],
      generationConfig: { responseMimeType: 'application/json' }
    })
  });
  var j = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200) throw new Error('Gemini: ' + (j.error ? j.error.message : res.getResponseCode()));
  var parts = (j.candidates && j.candidates[0].content && j.candidates[0].content.parts) || [];
  var out = parts.map(function(p) { return p.text || ''; }).join('').trim();
  var r;
  try { r = JSON.parse(out); } catch(err) { r = { text: out }; } // JSONで返らなかったら全体を本文として扱う
  var text = String(r.text || '').trim();
  if (!text) throw new Error('文字を読み取れませんでした');
  return { text: text, page: parseInt(r.page, 10) || null };
}

// 復元: A1セルのJSONを返す
function doGet(e) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Backup');
    if (!sheet) return respond(false, 'Backupシートがありません');

    var raw = sheet.getRange('A1').getValue();
    if (!raw) return respond(false, 'バックアップデータがありません');

    var d = JSON.parse(raw);
    d.success = true;
    return json(d);
  } catch(err) {
    return respond(false, err.message);
  }
}

// レスポンス共通化
function respond(success, error) {
  var obj = { success: success };
  if (error) obj.error = error;
  return json(obj);
}
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════
//  動作テスト（エディタで関数を選択して▶実行）
// ══════════════════════════════════════════════════════

// スプシのA1セルにサンプルJSONが入ればOK
function testBackup() {
  var fakeEvent = {
    postData: {
      contents: JSON.stringify({ readDates: ['2025-01-01'], books: [] })
    }
  };
  Logger.log(doPost(fakeEvent).getContent()); // {"success":true}
}

// APIキーとモデルが使えるか確認（小さな画像で試す）→ ログにエラーが出なければOK
// ※初回実行時に「ドライブ」「外部サービス」の権限を許可してください
function testGemini() {
  var png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
  try { Logger.log(JSON.stringify(geminiOcr(png, 'image/png'))); }
  catch(err) { Logger.log(err.message); } // 「文字を読み取れませんでした」はキー・モデルOKの意味
  DriveApp.getRootFolder(); // ドライブ権限の許可を促す
}

// 使えるモデル名の一覧（GEMINI_MODEL を変えたいとき用）
function listModels() {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=100', { headers: { 'x-goog-api-key': key } });
  JSON.parse(res.getContentText()).models
    .filter(function(m) { return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0; })
    .forEach(function(m) { Logger.log(m.name.replace('models/', '')); });
}
