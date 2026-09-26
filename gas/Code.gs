// ══════════════════════════════════════════════════════
//  NeoRead — GAS Web App
//  ① バックアップ: スプシの「Backup」シートA列にJSONを分割して保存・復元
//  ② 文字起こし : 画像をドライブに保存 → Geminiで文字起こし
//  ③ 本の登録   : 表紙の写真 → 書名・著者・ジャンル・総ページ数
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
var CHUNK = 40000; // 1セルの上限は5万文字なので、余裕をもって4万文字ずつ分ける

// POST: action で振り分け（actionなし = 従来のバックアップ）
function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return respond(false, 'No POST data');
  }
  try {
    var raw = e.postData.contents;
    var body = JSON.parse(raw); // JSON検証
    if (body.action === 'ocr') return json(ocr(body));
    if (body.action === 'book') return json(bookInfo(body));

    // 先頭に 'x' を付けて保存（数字・日付・数式として解釈されるのを防ぐ）
    var rows = [];
    for (var i = 0; i < raw.length; i += CHUNK) rows.push(['x' + raw.slice(i, i + CHUNK)]);
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Backup') || ss.insertSheet('Backup');
    sheet.getRange('A:A').clearContent();
    sheet.getRange(1, 1, rows.length, 1).setValues(rows);
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
  var out = callGemini([
    { text: 'これは本のページの写真です。次の形のJSONだけを出力してください。\n' +
            '{"page": ページ番号の数字（写っていなければ null。見開きなら小さい方）, "text": "本文"}\n' +
            'text には本文をそのまま正確に文字起こしし、ページ番号・柱（ページ上部の章タイトル等）は含めず、段落の区切りは改行で表してください。' },
    { inline_data: { mime_type: mime, data: base64 } }
  ], { generationConfig: { responseMimeType: 'application/json' } });
  var r = parseJson(out) || { text: out }; // JSONで返らなかったら全体を本文として扱う
  var text = String(r.text || '').trim();
  if (!text) throw new Error('文字を読み取れませんでした');
  return { text: text, page: parseInt(r.page, 10) || null };
}

// 本の登録: 表紙から書名・著者を読み、Google検索でジャンルと総ページ数を調べる
function bookInfo(body) {
  if (!body.image) throw new Error('画像がありません');
  var parts = [
    { text: 'これは本の表紙の写真です。\n' +
            '1. 表紙から書名と著者名を読み取ってください。\n' +
            '2. Google検索でその本のジャンルと総ページ数を調べてください。\n' +
            '次の形のJSONだけを出力してください（前後に文章は付けない）。\n' +
            '{"title": "書名", "author": "著者名（なければ空文字）", "genre": "ジャンルを短く（例: ビジネス、自己啓発、小説、技術書）", "pages": 総ページ数の数字（わからなければ null）}' },
    { inline_data: { mime_type: body.mime || 'image/jpeg', data: body.image } }
  ];
  var out;
  try { out = callGemini(parts, { tools: [{ google_search: {} }] }); }
  catch (err) { out = callGemini(parts); } // 検索が使えないとき（上限・非対応モデル）は検索なしで再挑戦
  var r = parseJson(out) || {};
  if (!r.title) throw new Error('書名を読み取れませんでした');
  return { success: true, title: String(r.title), author: String(r.author || ''), genre: String(r.genre || ''), pages: parseInt(r.pages, 10) || null };
}

function callGemini(parts, opts) {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('GEMINI_API_KEY が未設定です');
  var model = props.getProperty('GEMINI_MODEL') || 'gemini-flash-latest';

  var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': key },
    muteHttpExceptions: true,
    payload: JSON.stringify(Object.assign({ contents: [{ parts: parts }] }, opts || {}))
  });
  var j = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200) throw new Error('Gemini: ' + (j.error ? j.error.message : res.getResponseCode()));
  var ps = (j.candidates && j.candidates[0] && j.candidates[0].content && j.candidates[0].content.parts) || [];
  return ps.map(function(p) { return p.text || ''; }).join('').trim();
}

// 文章の中から { ... } を取り出してJSONとして読む（読めなければ null）
function parseJson(s) {
  var m = String(s).match(/\{[\s\S]*\}/);
  if (!m) return null;
  try { return JSON.parse(m[0]); } catch (err) { return null; }
}

// 復元: A列をつなげてJSONを返す（昔のA1だけの形式もそのまま読める）
function doGet(e) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Backup');
    if (!sheet) return respond(false, 'Backupシートがありません');

    var n = sheet.getLastRow();
    var raw = n < 1 ? '' : sheet.getRange(1, 1, n, 1).getValues().map(function(r) {
      var v = String(r[0]);
      return v.charAt(0) === 'x' ? v.slice(1) : v;
    }).join('');
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

// 12万文字のデータで保存→復元して、同じ内容に戻ればOK（ログに「OK」）
// ※実行するとBackupシートがテスト用データで上書きされます。最後にアプリから「今すぐバックアップ」してください
function testBackup() {
  var big = JSON.stringify({ readDates: ['2025-01-01'], books: [{ id: '1', title: 'テスト', memos: [{ text: new Array(120001).join('あ') }] }] });
  Logger.log(doPost({ postData: { contents: big } }).getContent()); // {"success":true}
  var back = JSON.parse(doGet().getContent());
  delete back.success;
  Logger.log(JSON.stringify(back) === big ? 'OK：分割保存と復元が正しく動いています' : 'NG：復元した内容が一致しません');
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
