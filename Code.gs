const SHEET_ID = '13OhYf8z-drBqHY1NujsgtxLKppkgVw8WtMTal3Brvxk';

// ★ パスワード管理 ★
const ADMIN_PASS = 'admin'; 
// ★ ゲストがサーバーを使うときの名前(index.html の GUEST_ID と同じにする)。パスワードは要らない ★
const GUEST_ID = '__guest__';

// ゲストか、ログインしたユーザーなら true。ゲストに許す機能の関数だけで使う
function checkUserOrGuest_(id, pass) {
  return String(id) === GUEST_ID || checkUser_(id, pass);
}
const NO_PASS_USERS = ['test'];

// ★ ゲームのスコア送信用の合言葉(ゲーム側のコードにも同じ値を入れる。好きな文字列に変えてOK) ★
const SCORE_KEY = 'score_juki2023_k3y';

// ★ 「小さいほど上位」のゲーム名(タイムアタックなど)。それ以外は大きいほど上位 ★
//   例: const LOWER_IS_BETTER = ['ドッジレース'];
const LOWER_IS_BETTER = [];

// ★ ルーム(チャンネル)の一覧。好きな名前に変えてOK。1番目が「最初に開くルーム」で、
//   ルーム情報のない古いメッセージもここに入る。名前を変えると、旧名のメッセージは1番目に移る ★
const ROOMS = ['雑談', 'ゲーム', '連絡'];

// ★ 「入力中…」を表示し続ける秒数(キャッシュの寿命) ★
const MAX_PER_ROOM = 100;   // 各ルームで画面に送るメッセージ数(新しい順に100件。古いものはシートに残る)
const TYPING_TTL = 7;

function normRoom_(r) {
  r = String(r == null ? '' : r);
  return ROOMS.indexOf(r) !== -1 ? r : ROOMS[0];
}

function typingKey_(id) {
  return 'ty_' + String(id).slice(0, 100);
}

// 入力中のユーザー一覧 [{id, room}] を返す(自分は除く)
function getTyping_(ids, exceptId) {
  const keys = [];
  const owners = {};
  ids.forEach(function(u) {
    const sid = String(u);
    if (sid === String(exceptId)) return;
    const k = typingKey_(sid);
    keys.push(k);
    owners[k] = sid;
  });
  if (!keys.length) return [];
  const found = CacheService.getScriptCache().getAll(keys);
  const out = [];
  Object.keys(found).forEach(function(k) {
    out.push({ id: owners[k], room: normRoom_(found[k]) });
  });
  return out;
}

// 入力中になったことを記録する(画面から入力中に数秒おきに呼ばれる)
function setTyping(id, pass, room) {
  try {
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Users');
    if (!sh) return { error: '「Users」シートがありません。' };
    const d = sh.getDataRange().getValues();
    let ok = false;
    const ids = [];
    for (let i = 1; i < d.length; i++) {
      if (d[i][0] === '') continue;
      ids.push(String(d[i][0]));
      if (passOk_(d[i][0], d[i][1], id, pass)) ok = true;
    }
    if (!ok) return { error: '認証エラー' };
    CacheService.getScriptCache().put(typingKey_(id), normRoom_(room), TYPING_TTL);
    return { typing: getTyping_(ids, id) };
  } catch(e) {
    return { error: '入力中エラー: ' + e.toString() };
  }
}

function doGet(e) {
  // ゲームからのスコア送信: ?action=score&game=..&player=..&score=..&key=..
  if (e && e.parameter && e.parameter.action === 'score') {
    return jsonOut_(submitScore(e.parameter.game, e.parameter.player, e.parameter.score, e.parameter.key));
  }
    if (e && e.parameter && e.parameter.bridge === '1') {
    return HtmlService.createHtmlOutputFromFile('bridge')
      .setTitle('bridge')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('juki2023 base')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let p = (e && e.parameter) || {};
  try {
    if (e && e.postData && e.postData.contents && String(e.postData.type).indexOf('json') !== -1) {
      p = Object.assign({}, p, JSON.parse(e.postData.contents));
    }
  } catch(err) {}
  return jsonOut_(submitScore(p.game, p.player, p.score, p.key));
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// 🔧 自動修復: Googleのサーバーが一時的に失敗したときは、少し待ってもう一度だけやり直す
//   (パスワード違いなど「やり直しても直らないエラー」は、やり直さない)
// ============================================================
function isTransientError_(res) {
  return !!(res && res.error && /エラー:/.test(String(res.error)));
}

function loginAndSync(id, pass, tiKey) {
  let res = loginAndSyncOnce_(id, pass, tiKey);
  if (isTransientError_(res)) {
    Utilities.sleep(600);
    res = loginAndSyncOnce_(id, pass, tiKey);
  }
  return res;
}

function loginAndSyncOnce_(id, pass, tiKey) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const userSheet = ss.getSheetByName('Users');
    if (!userSheet) return { error: "「Users」シートがありません。" };
    
    const userData = userSheet.getDataRange().getValues();
    let userIndex = -1;
    for (let i = 1; i < userData.length; i++) {
       if (passOk_(userData[i][0], userData[i][1], id, pass)) {
        userIndex = i + 1;
        break;
      }
    }

    if (userIndex === -1) return { error: "IDまたはパスワードが違います" };

    const now = new Date();
    userSheet.getRange(userIndex, 3).setValue(now);
    return getAllData(id, tiKey);
  } catch(e) { return { error: "ログインエラー: " + e.toString() }; }
}

function sendMessage(id, pass, message, adminKey, tiKey, isPin, imageBase64, imageMime, quote, room) {
  try {
    return sendMessageCore_(id, pass, message, adminKey, tiKey, isPin, imageBase64, imageMime, quote, room);
  } catch(e) {
    return { error: "送信エラー: " + e.toString() };
  }
}

function sendMessageCore_(id, pass, message, adminKey, tiKey, isPin, imageBase64, imageMime, quote, room) {
  const auth = loginAndSync(id, pass, tiKey);
  if (auth.error) return auth;

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const msgSheet = getOrCreateMessagesSheet_(ss);

  let role = 0; 
   if (isTiOk_(id, tiKey)) role = 2;
  else if (adminKey === ADMIN_PASS) role = 1;

  // ★ 画像が添付されていればDriveに保存してURLを作る ★
  //   保存に失敗したときは、画像なしで黙って送らず、エラーを返す(画面に理由が出る)
  let imageUrl = "";
  if (imageBase64) {
    try {
      const bytes = Utilities.base64Decode(imageBase64);
      const blob = Utilities.newBlob(bytes, imageMime || 'image/jpeg', 'chat_' + new Date().getTime() + '.jpg');
      const file = getOrCreateImageFolder().createFile(blob);
      // 共有設定は学校の設定で禁止されていることがあるので、失敗しても続行(表示はサーバー経由で行う)
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) { Logger.log('共有設定スキップ: ' + e); }
      imageUrl = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
    } catch(e) {
      Logger.log('画像アップロードエラー: ' + e.toString());
      return { error: '画像の保存に失敗しました: ' + String(e).slice(0, 120) };
    }
  }

  if (role >= 1 && isPin) {
    PropertiesService.getScriptProperties().setProperty('PINNED_MSG', message);
  }

  // ★ 引用(返信)があれば、送信者・本文の抜粋・元メッセージの時刻だけを安全な形で保存する(G列) ★
  let quoteJson = "";
  if (quote && typeof quote === 'object') {
    const q = {
      s: String(quote.s || '').slice(0, 40),
      t: String(quote.t || '').slice(0, 80),
      ts: String(quote.ts || '').slice(0, 40)
    };
    if (q.s || q.t) quoteJson = JSON.stringify(q);
  }

  // ★ E列=リアクションJSON、F列=画像URL、G列=引用JSON、H列=ルーム ★
  msgSheet.appendRow([new Date(), id, message, role, "{}", imageUrl, quoteJson, normRoom_(room)]);
  pushForChat_(id, message, normRoom_(room), !!imageUrl); // 📲 プッシュ通知(失敗してもチャットは止めない)
  try { CacheService.getScriptCache().remove(typingKey_(id)); } catch(e) {} // 送信したら入力中を消す
  return getAllData(id, tiKey);
}

// ★ チャット画像の保存先フォルダを取得（なければ作成） ★
function getOrCreateImageFolder() {
  const FOLDER_NAME = 'juki2023_chat_images';
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(FOLDER_NAME);
}

// ============================================================
// 🖼️ 画像をサーバー経由で表示する(共有設定がなくても、スクリプトの持ち主が読めるDriveの画像なら表示できる)
//   Messages シートのF列には、Driveの「共有リンク」でも「ファイルID」だけでもOK
// ============================================================
function extractDriveId_(s) {
  s = String(s || '').trim();
  const m = s.match(/\/d\/([A-Za-z0-9_-]{20,})/) || s.match(/[?&]id=([A-Za-z0-9_-]{20,})/);
  if (m) return m[1];
  return /^[A-Za-z0-9_-]{25,}$/.test(s) ? s : '';
}

function getImageData(id, pass, urlOrId) {
  try {
    if (!checkUser_(id, pass)) return { error: '認証エラー' };
    const fid = extractDriveId_(urlOrId);
    if (!fid) return { error: '画像のURLからIDを読み取れません' };
    const blob = DriveApp.getFileById(fid).getBlob();
    const mime = String(blob.getContentType() || '');
    if (mime.indexOf('image/') !== 0) return { error: '画像ファイルではありません' };
    const bytes = blob.getBytes();
    if (bytes.length > 8 * 1024 * 1024) return { error: '画像が大きすぎます(8MBまで)' };
    return { ok: true, mime: mime, data: Utilities.base64Encode(bytes) };
  } catch (e) {
    return { error: '画像取得エラー: ' + String(e).slice(0, 100) };
  }
}

// エディタから1回だけ手動で実行して、Driveの権限を許可するための関数
function authorizeDrive() {
  Logger.log('OK: ' + getOrCreateImageFolder().getUrl());
}

function deleteMessage(id, pass, rowIndex, adminKey, tiKey) {
  const auth = loginAndSync(id, pass, tiKey);
  if (auth.error) return auth;

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const msgSheet = ss.getSheetByName('Messages');
  if (!msgSheet) return { error: "「Messages」シートがありません。" };

    if (isTiOk_(id, tiKey)) {
    const data = msgSheet.getDataRange().getValues();
    if (rowIndex <= 1 || rowIndex > data.length) return getAllData(id, tiKey);
    msgSheet.deleteRow(rowIndex);
  } else {
    return { error: "このメッセージを削除するには TrustedInstaller の権限が必要です。" };
  }

  return getAllData(id, tiKey);
}

// ★ リアクション（スタンプ）を処理する関数 ★
function toggleReaction(id, pass, tiKey, rowIndex, emoji) {
  const auth = loginAndSync(id, pass, tiKey);
  if (auth.error) return auth;

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const msgSheet = ss.getSheetByName('Messages');
  
  // 5列目（E列）のデータを取得
  let reactionCell = msgSheet.getRange(rowIndex, 5);
  let data = reactionCell.getValue();
  let reactions = {};
  if (data) {
    try { reactions = JSON.parse(data); } catch(e) {}
  }

  // 絵文字の配列を準備
  if (!reactions[emoji]) reactions[emoji] = [];

  // すでに押していれば取り消し、押していなければ追加
  let userIndex = reactions[emoji].indexOf(id);
  if (userIndex > -1) {
    reactions[emoji].splice(userIndex, 1);
  } else {
    reactions[emoji].push(id);
  }

  reactionCell.setValue(JSON.stringify(reactions));
  return getAllData(id, tiKey);
}

function getAllData(myId, tiKey) {
  let res = getAllDataOnce_(myId, tiKey);
  if (isTransientError_(res)) {
    Utilities.sleep(600);
    res = getAllDataOnce_(myId, tiKey);
  }
  return res;
}

// 「Messages」シートが消えていたら、見出し付きで自動的に作り直す
function getOrCreateMessagesSheet_(ss) {
  let sh = ss.getSheetByName('Messages');
  if (!sh) {
    sh = ss.insertSheet('Messages');
    sh.appendRow(['時刻', '送信者', '本文', '権限', 'リアクション', '画像', '引用', 'ルーム', '編集']);
  }
  return sh;
}

function getAllDataOnce_(myId, tiKey) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const msgSheet = getOrCreateMessagesSheet_(ss);
    const userSheet = ss.getSheetByName('Users');
    if (!msgSheet || !userSheet) return { error: "シート名を確認してください" };
    
     let isTI = isTiOk_(myId, tiKey);
    
    const msgValues = msgSheet.getDataRange().getValues();
    const userValues = userSheet.getDataRange().getValues();
    const pinned = PropertiesService.getScriptProperties().getProperty('PINNED_MSG') || 'ログイン成功';

    let messages = [];
    let onlineUsers = []; 
    const nowMs = new Date().getTime(); 

    // ★ 軽くするため、ルームごとに「最新 MAX_PER_ROOM 件」だけを送る(古い順に並べたまま)
    const keep = {};
    const perRoom = {};
    for (let k = msgValues.length - 1; k >= 1; k--) {
      const r0 = msgValues[k];
      if (r0[0] === "" || r0[1] === "") continue;
      if (String(r0[2]) === '!deleted') continue; // 本人が削除したメッセージは送らない
      const rm = normRoom_(r0[7]);
      perRoom[rm] = (perRoom[rm] || 0) + 1;
      if (perRoom[rm] <= MAX_PER_ROOM) keep[k] = true;
    }

    const more = {};
    Object.keys(perRoom).forEach(function(rm) { more[rm] = perRoom[rm] > MAX_PER_ROOM; });

    for (let i = 1; i < msgValues.length; i++) {
      if (!keep[i]) continue;
      try { // ★ 壊れた行が1つあっても、他のメッセージは表示する
      let row = msgValues[i];
      if (row[0] === "" || row[1] === "") continue;

      let msgDate = new Date(row[0]);
      let rawTimeStr = isNaN(msgDate.getTime()) ? "" : msgDate.toISOString(); // 相対時間計算用に生の時間を渡す
      
      let readUsers = [];
      for (let j = 1; j < userValues.length; j++) {
        let lastAccess = new Date(userValues[j][2]);
        if (userValues[j][0] !== row[1] && !isNaN(lastAccess.getTime()) && !isNaN(msgDate.getTime())) {
          if (lastAccess.getTime() >= msgDate.getTime()) {
            readUsers.push(userValues[j][0]);
          }
        }
      }

      let rawRole = row[3];
      let finalRole = 0;
      if (rawRole === 2) finalRole = 2;
      else if (rawRole === 1 || rawRole === true || rawRole === "true") finalRole = 1;

      // ★ リアクションデータを読み込む ★
      let reactionStr = row[4];
      let reactionsObj = {};
      if (reactionStr) {
        try { reactionsObj = JSON.parse(reactionStr); } catch(e) {}
      }

      // ★ 画像URLを読み込む（F列） ★
      let imageUrl = row[5] || "";

      // ★ 引用を読み込む（G列） ★
      let quoteObj = null;
      if (row[6]) {
        try { quoteObj = JSON.parse(row[6]); } catch(e) {}
      }

      // ★ 編集した時刻を読み込む（I列） ★
      let editedMs = 0;
      if (row[8]) {
        const ed = new Date(row[8]);
        if (!isNaN(ed.getTime())) editedMs = ed.getTime();
      }

      messages.push({
        rowIndex: i + 1, 
        timeStr: rawTimeStr,
        sender: String(row[1]), // ★ 数値扱いされたセル対策で文字列化
        text: String(row[2]),   // ★ 同上(例: "123"だけのメッセージが数値セルになるケース)
        role: finalRole,
        readBy: readUsers,
        reactions: reactionsObj,
        imageUrl: String(imageUrl),
        quote: quoteObj,
        room: normRoom_(row[7]),
        edited: editedMs
      });
      } catch(rowErr) {
        Logger.log('壊れた行をスキップ: ' + (i + 1) + ' ' + rowErr.toString());
      }
    }

    const allIds = [];
    for (let j = 1; j < userValues.length; j++) {
      if (userValues[j][0] !== "") {
        allIds.push(userValues[j][0]);
        let lastAccess = new Date(userValues[j][2]).getTime();
        if (!isNaN(lastAccess) && (nowMs - lastAccess) <= 15000) {
          onlineUsers.push(userValues[j][0]);
        }
      }
    }

    let typing = [];
    try { typing = getTyping_(allIds, myId); } catch(e) {} // 入力中表示が失敗してもチャット自体は動かす

    let nowPlaying = [];
    try { nowPlaying = getNowPlayingSafe_(allIds, myId); } catch(e) {} // 「いま聴いている曲」が失敗してもチャット自体は動かす

    return {
      messages: messages, pinned: pinned, onlineUsers: onlineUsers, isTI: isTI, rooms: ROOMS, typing: typing, more: more,
      users: allIds.map(String), nowPlaying: nowPlaying
    };
  } catch(e) { return { error: "データ取得エラー: " + e.toString() }; }
}

function fetchVideoFromYoutube(youtubeUrl) {
  try {
    var videoId = youtubeUrl.match(/(?:v=|youtu\.be\/|embed\/)([^&\?]+)/) ? youtubeUrl.match(/(?:v=|youtu\.be\/|embed\/)([^&\?]+)/)[1] : youtubeUrl;
    
    // ★最新のAPIエンドポイント（末尾に /api/json をつけない）
    var apiUrl = "https://api.cobalt.tools/"; 
    
    var payload = {
      url: "https://www.youtube.com/watch?v=" + videoId,
      vQuality: "720",
      filenameStyle: "classic"
    };
    
    // ★ブラウザのフリをするための強力なヘッダー
    var options = {
      method: "post",
      contentType: "application/json",
      headers: { 
        "Accept": "application/json",
        "Origin": "https://cobalt.tools",
        "Referer": "https://cobalt.tools/",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    var response = UrlFetchApp.fetch(apiUrl, options);
    var responseText = response.getContentText();
    
    Logger.log("API Response: " + responseText);
    
    var json = JSON.parse(responseText);
    
    // 新仕様では結果が "url" ではなく "result" の中に入っている場合があります
    if (json.url) {
      return { success: true, message: "成功！", videoUrl: json.url };
    } else if (json.result && json.result.url) {
      return { success: true, message: "成功！", videoUrl: json.result.url };
    } else {
      return { success: false, message: "抽出失敗: " + (json.text || JSON.stringify(json)) };
    }
  } catch (e) {
    return { success: false, message: "エラー: " + e.toString() };
  }
}


// ============================================================
// 🩺 画面側で起きたエラーの自動記録(ErrorLog シート。自動作成)
//   「いつ・誰の画面で・何が起きたか」が溜まるので、バグの原因探しに使える
// ============================================================
const ERROR_LOG_MAX_ROWS = 300;

function reportClientError(id, pass, msg, ua, ver) {
  try {
    if (!checkUser_(id, pass)) return { ok: false };
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName('ErrorLog');
    if (!sh) {
      sh = ss.insertSheet('ErrorLog');
      sh.appendRow(['時刻', 'ユーザー', '内容', '端末', 'バージョン']);
    }
    sh.appendRow([
      new Date(),
      String(id).slice(0, 40),
      String(msg || '').slice(0, 500),
      String(ua || '').slice(0, 200),
      String(ver || '').slice(0, 20)
    ]);
    // 古いものから消して、行が増えすぎないようにする
    const last = sh.getLastRow();
    if (last > ERROR_LOG_MAX_ROWS + 1) sh.deleteRows(2, last - (ERROR_LOG_MAX_ROWS + 1));
    return { ok: true };
  } catch(e) {
    return { ok: false };
  }
}

// ============================================================
// 🏆 ゲームのスコアランキング
//   Scores シート(自動作成): 更新日時 / ゲーム / プレイヤー / スコア
//   1人1ゲームにつき「ベストスコア」だけを1行で保持する
// ============================================================

function getScoreSheet_(create) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName('Scores');
  if (!sh && create) {
    sh = ss.insertSheet('Scores');
    sh.appendRow(['更新日時', 'ゲーム', 'プレイヤー', 'スコア']);
  }
  return sh;
}

function checkUser_(id, pass) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Users');
  if (!sh) return false;
  const d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (passOk_(d[i][0], d[i][1], id, pass)) return true;
  }
  return false;
}

function userExists_(id) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Users');
  if (!sh) return false;
  const d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][0]) === String(id)) return true;
  }
  return false;
}

// ゲームから呼ばれる。ベスト更新のときだけ書き込む
function submitScore(game, player, score, key) {
  try {
    if (String(key) !== SCORE_KEY) return { ok: false, error: '合言葉(key)が違います' };
    game = String(game || '').trim().slice(0, 40);
    player = String(player || '').trim().slice(0, 40);
    const num = Number(score);
    if (!game || !player || score === '' || score === null || !isFinite(num) || Math.abs(num) > 1e12) {
      return { ok: false, error: 'game / player / score が正しくありません' };
    }
    if (!userExists_(player)) return { ok: false, error: 'プレイヤー「' + player + '」はUsersシートにいません' };

    const lower = LOWER_IS_BETTER.indexOf(game) !== -1;
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = getScoreSheet_(true);
      const values = sh.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][1]) === game && String(values[i][2]) === player) {
          const old = Number(values[i][3]);
          const better = !isFinite(old) || (lower ? num < old : num > old);
          if (better) {
            sh.getRange(i + 1, 1).setValue(new Date());
            sh.getRange(i + 1, 4).setValue(num);
            return { ok: true, updated: true, best: num };
          }
          return { ok: true, updated: false, best: old };
        }
      }
      sh.appendRow([new Date(), game, player, num]);
      return { ok: true, updated: true, best: num };
    } finally {
      lock.releaseLock();
    }
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// チャット画面から呼ばれる。全ゲームの上位 limit 人を返す
function getAllRankings(id, pass, limit) {
  try {
    if (!checkUser_(id, pass)) return { error: '認証エラー' };
    const n = Math.min(Math.max(Number(limit) || 10, 1), 50);
    const sh = getScoreSheet_(false);
    if (!sh) return { rankings: {}, lower: LOWER_IS_BETTER };

    const values = sh.getDataRange().getValues();
    const byGame = {};
    for (let i = 1; i < values.length; i++) {
      const g = String(values[i][1]);
      const p = String(values[i][2]);
      const sc = Number(values[i][3]);
      if (!g || !p || !isFinite(sc)) continue;
      const d = new Date(values[i][0]);
      (byGame[g] = byGame[g] || []).push({
        player: p,
        score: sc,
        time: isNaN(d.getTime()) ? '' : d.toISOString()
      });
    }
    const rankings = {};
    Object.keys(byGame).forEach(function(g) {
      const lower = LOWER_IS_BETTER.indexOf(g) !== -1;
      byGame[g].sort(function(a, b) { return lower ? a.score - b.score : b.score - a.score; });
      rankings[g] = byGame[g].slice(0, n);
    });
    return { rankings: rankings, lower: LOWER_IS_BETTER };
  } catch(e) {
    return { error: 'ランキング取得エラー: ' + e.toString() };
  }
}

// エディタから手動で実行して、ランキング機能が動くか試す用(Usersシートの最初の人でテストスコアを登録)
function testScore() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Users');
  const first = sh.getDataRange().getValues()[1][0];
  Logger.log(JSON.stringify(submitScore('テスト', first, 123, SCORE_KEY)));
}

// ============================================================
// ⬆ 過去のメッセージ読み込み(「過去のメッセージを読み込む」ボタン用)
//   古いメッセージを、新しい方から limit 件ずつ返す(beforeRow より前の行・指定ルームだけ)
// ============================================================
function getOlderMessages(id, pass, room, beforeRow, limit) {
  try {
    if (!checkUser_(id, pass)) return { error: '認証エラー' };
    const n = Math.min(Math.max(Number(limit) || 100, 1), 200);
    const before = Number(beforeRow) || 0;
    if (before <= 2) return { messages: [], hasMore: false };
    const rm = normRoom_(room);

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName('Messages');
    if (!sh) return { messages: [], hasMore: false };
    const values = sh.getDataRange().getValues();
    // 既読の計算用(最新メッセージと同じ方法: 最終アクセス時刻がメッセージより後の人)
    const userSheet = ss.getSheetByName('Users');
    const userValues = userSheet ? userSheet.getDataRange().getValues() : [];

    const out = [];
    let hasMore = false;
    // 行番号 = 配列の添字 + 1。beforeRow より小さい行を、下(新しい方)から上へ見ていく
    for (let i = Math.min(before - 2, values.length - 1); i >= 1; i--) {
      const row = values[i];
      if (row[0] === '' || row[1] === '') continue;
      if (String(row[2]) === '!update' || String(row[2]) === '!deleted') continue;
      if (normRoom_(row[7]) !== rm) continue;
      if (out.length >= n) { hasMore = true; break; }

      const d = new Date(row[0]);
      let role = 0;
      if (row[3] === 2) role = 2;
      else if (row[3] === 1 || row[3] === true || row[3] === 'true') role = 1;
      let reactions = {};
      if (row[4]) { try { reactions = JSON.parse(row[4]); } catch(e) {} }
      let quote = null;
      if (row[6]) { try { quote = JSON.parse(row[6]); } catch(e) {} }
      let editedMs = 0;
      if (row[8]) { const ed = new Date(row[8]); if (!isNaN(ed.getTime())) editedMs = ed.getTime(); }

      // 既読の人を調べる
      const readUsers = [];
      if (!isNaN(d.getTime())) {
        for (let j = 1; j < userValues.length; j++) {
          const la = new Date(userValues[j][2]);
          if (String(userValues[j][0]) !== String(row[1]) && !isNaN(la.getTime()) && la.getTime() >= d.getTime()) {
            readUsers.push(String(userValues[j][0]));
          }
        }
      }

      out.push({
        rowIndex: i + 1,
        timeStr: isNaN(d.getTime()) ? '' : d.toISOString(),
        sender: String(row[1]),
        text: String(row[2]),
        role: role,
        readBy: readUsers,
        reactions: reactions,
        imageUrl: String(row[5] || ''),
        quote: quote,
        room: rm,
        edited: editedMs
      });
    }
    out.reverse(); // 古い順に並べ直す
    return { messages: out, hasMore: hasMore };
  } catch(e) {
    return { error: '過去メッセージ取得エラー: ' + String(e).slice(0, 80) };
  }
}

// ============================================================
// 🎵 TikTok(公式の仕組みだけを使う)
//   ・短縮URL(vm.tiktok.com/… や tiktok.com/t/…)の転送先を調べて、動画IDと正式なURLに直す
//   ・動画の「タイトル・投稿者・サムネイルのURL」は、TikTok公式の oEmbed(公開されている情報取得)で読む
//   ・動画そのものは、取得も保存も中継もしない(再生は、画面側がTikTok公式の埋め込みプレイヤーで行う)
//   ・TikTok以外のサイトには接続しない。同じURLは6時間キャッシュ
// ============================================================
const TT_HOST_RE_ = /^https:\/\/(?:[a-z0-9-]+\.)?tiktok\.com\//i;
const TT_UA_ = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36';

function ttIdOf_(x) { const m = String(x).match(/\/(?:video|v)\/(\d{10,25})/); return m ? m[1] : ''; }
function ttCacheKey_(prefix, s) {
  return prefix + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s, Utilities.Charset.UTF_8));
}

// URL → { id, url }(正式なURL)または { error }
function ttResolve_(url) {
  let u = String(url || '').trim().slice(0, 300).replace(/^http:\/\//i, 'https://'); // httpのURLもhttpsにそろえる
  if (!TT_HOST_RE_.test(u)) return { error: 'TikTokのURLではありません' };
  let vid = ttIdOf_(u);
  if (vid) return { id: vid, url: u };

  const cache = CacheService.getScriptCache();
  const ck = ttCacheKey_('tt_', u);
  const hit = cache.get(ck);
  if (hit) { try { const o = JSON.parse(hit); if (o && o.id) return o; } catch (e) {} }

  for (let hop = 0; hop < 4 && !vid; hop++) {
    const r = UrlFetchApp.fetch(u, { followRedirects: false, muteHttpExceptions: true, headers: { 'User-Agent': TT_UA_ } });
    const code = r.getResponseCode();
    const hdr = r.getHeaders();
    const loc = String(hdr['Location'] || hdr['location'] || '');
    if (code >= 300 && code < 400 && loc) {
      u = loc.charAt(0) === '/' ? 'https://www.tiktok.com' + loc : loc;
      if (!TT_HOST_RE_.test(u)) return { error: 'TikTokのURLを読み取れませんでした' }; // 外部サイトには飛ばない
      vid = ttIdOf_(u);
    } else {
      vid = ttIdOf_(r.getContentText().slice(0, 20000)); // 転送でなければ、ページ中のURLから探す
      break;
    }
  }
  if (!vid) return { error: '動画IDを読み取れませんでした' };
  const out = { id: vid, url: u.split('?')[0] };
  cache.put(ck, JSON.stringify(out), 21600);
  return out;
}

// 画面から呼ばれる: 動画IDだけを返す(上の動画欄・「ここで再生」用)
function resolveTikTok(id, pass, url) {
  try {
    if (!checkUserOrGuest_(id, pass)) return { error: '認証エラー' };
    const r = ttResolve_(url);
    return r.error ? { error: r.error } : { id: r.id };
  } catch (e) {
    return { error: 'TikTok取得エラー: ' + String(e).slice(0, 80) };
  }
}

// 画面から呼ばれる: 動画カードに出す情報(タイトル・投稿者・サムネイルのURL)を、公式のoEmbedで読む
function getTikTokInfo(id, pass, url) {
  try {
    if (!checkUserOrGuest_(id, pass)) return { error: '認証エラー' };
    const r = ttResolve_(url);
    if (r.error) return { error: r.error };

    const cache = CacheService.getScriptCache();
    const ck = ttCacheKey_('ti_', r.id);
    const hit = cache.get(ck);
    if (hit) { try { return JSON.parse(hit); } catch (e) {} }

    const res = UrlFetchApp.fetch('https://www.tiktok.com/oembed?url=' + encodeURIComponent(r.url), { muteHttpExceptions: true, headers: { 'User-Agent': TT_UA_ } });
    const code = res.getResponseCode();
    // 取得できなくても、動画IDは返す(画面側は「ここで再生」「TikTokで開く」だけのカードにできる)
    const base = { id: r.id, url: r.url, title: '', author: '', thumb: '' };
    if (code !== 200) return Object.assign(base, { info: false, note: 'TikTokから情報を取得できませんでした(' + code + ')' });
    let j = null;
    try { j = JSON.parse(res.getContentText()); } catch (e) {}
    if (!j) return Object.assign(base, { info: false, note: 'TikTokの返事を読めませんでした' });
    const th = String(j.thumbnail_url || '');
    const out = Object.assign(base, {
      info: true,
      title: String(j.title || '').slice(0, 150),
      author: String(j.author_name || '').slice(0, 60),
      thumb: /^https:\/\//.test(th) ? th : ''
    });
    cache.put(ck, JSON.stringify(out), 21600);
    return out;
  } catch (e) {
    return { error: 'TikTok情報の取得エラー: ' + String(e).slice(0, 80) };
  }
}

// 𝕏 ポストの内容を取得して、カード表示に使う(6時間キャッシュ)
//   ※ Xの埋め込み表示が裏で使っている取得口(非公式)を使います。Xの都合で止まる可能性があります。
// 動画の候補(variants)から、音つきmp4のいちばんよいものを選ぶ
//   variants: [{ url, type, bitrate }]。重すぎない画質(約2.2Mbps以下)の中で最高のものを選ぶ
function twBestMp4_(variants) {
  const mp4 = (variants || []).filter(function(v) {
    return v && /^https:\/\//.test(String(v.url || '')) &&
      (/mp4/i.test(String(v.type || '')) || /\.mp4(\?|$)/i.test(String(v.url)));
  });
  if (!mp4.length) return '';
  mp4.sort(function(a, b) { return (b.bitrate || 0) - (a.bitrate || 0); });
  const ok = mp4.filter(function(v) { return !v.bitrate || v.bitrate <= 2200000; })[0];
  return String((ok || mp4[mp4.length - 1]).url);
}

// syndication の返事(j)から動画を探す → { url, w, h, gif } または null
function twPickVideo_(j) {
  try {
    const variants = [];
    let w = 0, h = 0, gif = false;
    (j.mediaDetails || []).forEach(function(m) {
      if (!m || !m.video_info) return;
      if (m.type === 'animated_gif') gif = true;
      if (m.original_info) { w = w || Number(m.original_info.width) || 0; h = h || Number(m.original_info.height) || 0; }
      (m.video_info.variants || []).forEach(function(v) {
        if (v && v.url) variants.push({ url: v.url, type: v.content_type, bitrate: Number(v.bitrate) || 0 });
      });
    });
    if (!variants.length && j.video && Array.isArray(j.video.variants)) {
      j.video.variants.forEach(function(v) {
        if (v && v.src) variants.push({ url: v.src, type: v.type, bitrate: 0 });
      });
    }
    const url = twBestMp4_(variants);
    return url ? { url: url, w: w, h: h, gif: gif } : null;
  } catch (e) { return null; }
}
function getTweetCard(id, pass, tweetId) {
  try {
    if (!checkUserOrGuest_(id, pass)) return { ok: false, error: '認証エラー' };
    const tid = String(tweetId || '').replace(/\D/g, '').slice(0, 25);
    if (!tid) return { ok: false, notfound: true };

    const cache = CacheService.getScriptCache();
    const ck = 'tw2_' + tid;
    const hit = cache.get(ck);
    if (hit) { try { return JSON.parse(hit); } catch(e) {} }

    const token = ((Number(tid) / 1e15) * Math.PI).toString(36).replace(/(0+|\.)/g, '');
    const url = 'https://cdn.syndication.twimg.com/tweet-result?id=' + tid + '&lang=ja&token=' + encodeURIComponent(token);
    const r = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36' }
    });
    const code = r.getResponseCode();
    if (code === 404) return { ok: false, notfound: true };
    if (code !== 200) return { ok: false, error: 'Xの応答エラー ' + code };

    let j = null;
    try { j = JSON.parse(r.getContentText()); } catch(e) {}
    if (!j || j.__typename === 'TweetTombstone' || !j.user) return { ok: false, notfound: true };

    let text = String(j.text || '');
    const ent = j.entities || {};
    (ent.urls || []).forEach(function(u) {
      if (u && u.url) text = text.split(u.url).join(u.expanded_url || u.display_url || u.url);
    });
    (ent.media || []).forEach(function(m) {
      if (m && m.url) text = text.split(m.url).join('');
    });

    const photos = (j.photos || []).slice(0, 4).map(function(p) { return p.url; }).filter(String);
    const vd = twPickVideo_(j);
    const res = {
      ok: true,
      tweet: {
        name: String(j.user.name || '').slice(0, 60),
        handle: String(j.user.screen_name || '').slice(0, 40),
        avatar: String(j.user.profile_image_url_https || ''),
        text: text.trim().slice(0, 600),
        photos: photos,
        poster: (j.video && j.video.poster) ? String(j.video.poster) : '',
        video: vd ? vd.url : '',
        gif: vd ? !!vd.gif : false,
        vw: vd ? vd.w : 0,
        vh: vd ? vd.h : 0,
        date: j.created_at ? Utilities.formatDate(new Date(j.created_at), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') : '',
        likes: Number(j.favorite_count) || 0
      }
    };
    cache.put(ck, JSON.stringify(res), 21600);
    return res;
  } catch(e) {
    return { ok: false, error: '取得エラー: ' + String(e).slice(0, 80) };
  }
}


// ============================================================
// 🔍 𝕏 ポスト検索(外部サービス twitterapi.io を使う)
//   ・APIキーは「プロジェクトの設定 → スクリプトプロパティ」に TWITTERAPI_KEY という名前で入れる
//     (コードには書かない。画面にも出ない)
//   ・無料分(約0.1ドル ≒ 20件の検索が約33回ぶん)を長持ちさせるための設定が、下の4つ
//   ・同じ検索は6時間キャッシュ。その間は何度見ても、回数も課金も増えない
//   ・回数に数えるのは「実際に検索できたとき」だけ(エラーのときは数えない)
//   ※ twitterapi.io はXとは無関係の外部サービスです。料金や仕様は先方の都合で変わる可能性があります。
// ============================================================
const SEARCH_ALLOWED_USERS = [`juki2023`,'あつき'];   // 検索を使えるユーザーID。[] なら全員。自分だけにするなら ['juki'] のように書く
const SEARCH_USER_PER_HOUR = 2;    // 1人が1時間に使える回数
const SEARCH_DAILY_MAX = 5;        // 全体で1日に使える回数(「さらに表示」も1回に数える)
const SEARCH_CACHE_SECONDS = 21600; // 同じ検索を覚えておく時間(秒)。最長6時間

// APIキーを読む。名前の前後の空白や、大文字小文字・全角のちがいがあっても見つけられるようにする
function getSearchKey_() {
  const all = PropertiesService.getScriptProperties().getProperties();
  const want = 'TWITTERAPI_KEY';
  const norm = function(x) { return String(x).replace(/[\s\u3000]/g, '').toUpperCase(); };
  const names = Object.keys(all);
  for (let i = 0; i < names.length; i++) {
    if (norm(names[i]) === want) return String(all[names[i]]).replace(/[\s\u3000]/g, '');
  }
  return '';
}

// エディタから手動で実行して、キーが見えているかを確かめる用(キーそのものはログに出さない)
function checkKey() {
  const names = Object.keys(PropertiesService.getScriptProperties().getProperties());
  Logger.log('このプロジェクトのスクリプトプロパティ名: ' + JSON.stringify(names));
  const v = getSearchKey_();
  Logger.log(v ? 'APIキーが見つかりました(長さ ' + v.length + ' 文字)' : 'APIキーが見つかりません');
}

function readSearchDaily_(props, today) {
  let d = {};
  try { d = JSON.parse(props.getProperty('SEARCH_DAILY') || '{}') || {}; } catch(e) {}
  if (d.day !== today) d = { day: today, n: 0 };
  return d;
}

function searchPosts(id, pass, query, cursor) {
  try {
    if (!checkUser_(id, pass)) return { error: '認証エラー' };
      if (SEARCH_ALLOWED_USERS.length && SEARCH_ALLOWED_USERS.indexOf(String(id)) === -1 && String(id) !== GUEST_ID) {
      return { error: 'この検索は、許可されたユーザーだけが使えます' };
    }
    const props = PropertiesService.getScriptProperties();
    const key = getSearchKey_();
    if (!key) return { error: 'TWITTERAPI_KEY が未設定です(スクリプトプロパティに入れてください)' };
    const q = String(query || '').trim().slice(0, 200);
    if (!q) return { error: '検索ワードを入れてください' };
    const cur = String(cursor || '').slice(0, 400);

    // 同じ検索はキャッシュから返す(回数にも課金にも数えない)
    const cache = CacheService.getScriptCache();
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, q + '|' + cur, Utilities.Charset.UTF_8);
    const ck = 'sr_' + digest.map(function(b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
    const hit = cache.get(ck);
    if (hit) { try { const o = JSON.parse(hit); o.cached = true; return o; } catch(e) {} }

    // 回数の確認(ここではまだ数えない)
    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
    const uk = 'su_' + String(id).slice(0, 60);
    if (readSearchDaily_(props, today).n >= SEARCH_DAILY_MAX) {
      return { error: '今日の検索回数の上限(全体で' + SEARCH_DAILY_MAX + '回)に達しました。明日また使えます' };
    }
    if (Number(cache.get(uk) || 0) >= SEARCH_USER_PER_HOUR) {
      return { error: '検索のしすぎです。しばらくしてからにしてください(1時間に' + SEARCH_USER_PER_HOUR + '回まで)' };
    }

    let url = 'https://api.twitterapi.io/twitter/tweet/advanced_search?query=' + encodeURIComponent(q) + '&queryType=Latest';
    if (cur) url += '&cursor=' + encodeURIComponent(cur);
    const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'X-API-Key': key } });
    const code = r.getResponseCode();
    if (code === 401 || code === 403) return { error: 'APIキーが正しくないか、クレジットが足りません(' + code + ')' };
    if (code === 429) return { error: '混み合っています。少し待ってからやり直してください' };
    if (code !== 200) return { error: '検索サービスの応答エラー ' + code };

    let j = null;
    try { j = JSON.parse(r.getContentText()); } catch(e) {}
    if (!j) return { error: '検索サービスの返事を読めませんでした' };
    if (j.status === 'error') return { error: String(j.message || j.msg || '検索に失敗しました').slice(0, 100) };

    const list = Array.isArray(j.tweets) ? j.tweets : [];
    const out = {
      tweets: list.slice(0, 20).map(mapSearchTweet_).filter(function(t) { return t && t.id; }),
      next: (j.has_next_page && j.next_cursor) ? String(j.next_cursor) : ''
    };
    cache.put(ck, JSON.stringify(out), Math.min(SEARCH_CACHE_SECONDS, 21600));

    // 検索できたので、ここで回数を数える
    let left = null;
    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      const d = readSearchDaily_(props, today);
      d.n++;
      props.setProperty('SEARCH_DAILY', JSON.stringify(d));
      cache.put(uk, String(Number(cache.get(uk) || 0) + 1), 3600);
      left = Math.max(0, SEARCH_DAILY_MAX - d.n);
    } finally {
      lock.releaseLock();
    }
    out.left = left;
    return out;
  } catch(e) {
    return { error: '検索エラー: ' + String(e).slice(0, 80) };
  }
}

// 外部サービスの1件分を、画面のカード表示用の形に直す(項目名の違いに備えて、いくつかの名前を見る)
function mapSearchTweet_(t) {
  if (!t) return null;
  const a = t.author || t.user || {};
  let text = String(t.text || t.full_text || '');
  const ent = t.entities || {};
  (ent.urls || []).forEach(function(u) {
    if (u && u.url) text = text.split(u.url).join(u.expanded_url || u.display_url || u.url);
  });
  const mediaList = (t.extendedEntities && t.extendedEntities.media) || (t.extended_entities && t.extended_entities.media) || ent.media || [];
  const photos = [];
  let poster = '', video = '', gif = false;
  mediaList.forEach(function(m) {
    if (!m) return;
    const u = m.media_url_https || m.mediaUrlHttps || '';
    if (m.type !== 'photo' && !video) {
      const vi = m.video_info || m.videoInfo;
      if (vi && Array.isArray(vi.variants)) {
        video = twBestMp4_(vi.variants.map(function(v) {
          return { url: v.url, type: v.content_type || v.contentType, bitrate: Number(v.bitrate) || 0 };
        }));
        gif = m.type === 'animated_gif';
      }
    }
    if (!u) return;
    if (m.type === 'photo') { if (photos.length < 4) photos.push(String(u)); }
    else if (!poster) poster = String(u);
    if (m.url) text = text.split(m.url).join('');
  });
  let date = '';
  const dt = new Date(t.createdAt || t.created_at || '');
  if (!isNaN(dt.getTime())) date = Utilities.formatDate(dt, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  return {
    id: String(t.id || t.id_str || ''),
    name: String(a.name || '').slice(0, 60),
    handle: String(a.userName || a.screen_name || a.username || '').slice(0, 40),
    avatar: String(a.profilePicture || a.profile_image_url_https || ''),
    text: text.trim().slice(0, 600),
    photos: photos,
    poster: poster,
    video: video,
    gif: gif,
    date: date,
    likes: Number(t.likeCount || t.favorite_count) || 0
  };
}

// エディタから手動で実行して、外部サービスの返事の形を確かめる用(ログに最初の1件が出る)
function testSearchPosts() {
  const key = getSearchKey_();
  const r = UrlFetchApp.fetch('https://api.twitterapi.io/twitter/tweet/advanced_search?query=' + encodeURIComponent('nhk') + '&queryType=Latest', { muteHttpExceptions: true, headers: { 'X-API-Key': key } });
  Logger.log('code=' + r.getResponseCode());
  let j = null;
  try { j = JSON.parse(r.getContentText()); } catch(e) {}
  Logger.log('keys=' + (j ? Object.keys(j).join(',') : '(JSONではない)'));
  if (j && j.tweets && j.tweets[0]) Logger.log(JSON.stringify(j.tweets[0]).slice(0, 1800));
  else Logger.log(r.getContentText().slice(0, 500));
}

// アカウントのTLを表示する(searchPosts を使うので、回数制限・キャッシュも共通)
function getAccountTimeline(id, pass, handle, cursor, includeReplies) {
  const h = String(handle || '').replace(/^@/, '').replace(/[^A-Za-z0-9_]/g, '').slice(0, 15);
  if (!h) return { error: 'アカウント名を入れてください(例: nhk)' };
  const q = 'from:' + h + (includeReplies ? '' : ' -filter:replies');
  return searchPosts(id, pass, q, cursor);
}
// ============================================================
// 📥 TikTok動画を保存して再生(tikwm.com を使う。非公式サービス)
//   同じ動画は一度だけ保存して、Drive のファイルIDを覚えておく
// ============================================================
const TT_VIDEO_FOLDER = 'juki2023_tt_videos';
const TT_MAX_BYTES = 100 * 1024 * 1024; // UrlFetchApp の上限(約50MB)より少し小さめに
const TT_DL_PER_HOUR = 5;              // 1人が1時間に保存できる本数

function getTTFolder_() {
  const it = DriveApp.getFoldersByName(TT_VIDEO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(TT_VIDEO_FOLDER);
}

// tikwmを呼ぶ。「利用制限」の応答のときだけ、間隔をあけて最大3回まで試す
function tikwmFetch_(url) {
  let j = null, http = 0;
  for (let i = 0; i < 3; i++) {
    const api = UrlFetchApp.fetch('https://www.tikwm.com/api/', {
      method: 'post',
      payload: { url: url, hd: 0 },
      headers: { 'User-Agent': TT_UA_ },
      muteHttpExceptions: true
    });
    http = api.getResponseCode();
    j = null;
    try { j = JSON.parse(api.getContentText()); } catch (e) {}
    const limited = j && j.code !== 0 && /limit/i.test(String(j.msg || ''));
    if (http !== 200 || !limited) break;
    if (i < 2) Utilities.sleep(2000 * (i + 1));
  }
  return { j: j, http: http };
}

function getTikTokVideo(id, pass, url) {
  try {
    if (!checkUserOrGuest_(id, pass)) return { error: '認証エラー' };
    const r = ttResolve_(url); // 短縮URLも動画IDに直す(既存の関数)
    if (r.error) return { error: r.error };

    // 保存済みなら、それを返す
    const props = PropertiesService.getScriptProperties();
    const pk = 'ttv_' + r.id;
    const saved = props.getProperty(pk);
    if (saved) {
      try { DriveApp.getFileById(saved); return { fid: saved, cached: true }; }
      catch (e) { props.deleteProperty(pk); } // ファイルが消えていたら取り直す
    }

    // 1人あたりの回数制限
    const cache = CacheService.getScriptCache();
    const uk = 'ttdl_' + String(id).slice(0, 60);
    if (Number(cache.get(uk) || 0) >= TT_DL_PER_HOUR) {
      return { error: '保存のしすぎです。しばらくしてからにしてください(1時間に' + TT_DL_PER_HOUR + '本まで)' };
    }

    // tikwmで直リンクを取得
    const t = tikwmFetch_(r.url);
    if (t.http !== 200) return { error: 'tikwmの応答エラー ' + t.http };
    const j = t.j;
    if (j && j.code !== 0 && /limit/i.test(String(j.msg || ''))) {
      return { error: '動画取得サービスが混み合っています。少し待ってからもう一度押してください' };
    }
    if (!j || j.code !== 0 || !j.data) {
      return { error: String((j && j.msg) || 'tikwmで解析できませんでした').slice(0, 100) };
    }

    // サイズが分かる場合は、ダウンロード前に確認する
    const sizeBytes = Number(j.data.size) || 0;
    if (sizeBytes > TT_MAX_BYTES) {
      return { error: '動画が大きすぎます(約' + Math.round(sizeBytes / 1048576) + 'MB)。「TikTokで開く」から見てください' };
    }

    let play = String(j.data.play || '');
    if (!play) return { error: '動画のURLを取得できませんでした' };
    if (play.charAt(0) === '/') play = 'https://www.tikwm.com' + play;

    const res = UrlFetchApp.fetch(play, { muteHttpExceptions: true, headers: { 'User-Agent': TT_UA_ } });
    if (res.getResponseCode() !== 200) return { error: 'ダウンロードに失敗しました(' + res.getResponseCode() + ')' };
    const blob = res.getBlob();
    if (blob.getBytes().length > TT_MAX_BYTES) {
      return { error: '動画が大きすぎます。「TikTokで開く」から見てください' };
    }
    blob.setName(r.id + '.mp4');

    const file = getTTFolder_().createFile(blob);
    // 共有設定は学校の設定で制限されていることがあるので、段階的に試す
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
    catch (e1) {
      try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); }
      catch (e2) { Logger.log('共有設定スキップ: ' + e2); }
    }

    props.setProperty(pk, file.getId());
    cache.put(uk, String(Number(cache.get(uk) || 0) + 1), 3600);
    return { fid: file.getId() };
  } catch (e) {
    return { error: '保存エラー: ' + String(e).slice(0, 100) };
  }
}

// 古い動画を消して容量を空ける。エディタから1回実行して、毎日のトリガーに設定する
function cleanupTikTokVideos() {
  const limit = Date.now() - 3 * 24 * 3600 * 1000; // 3日より古いもの
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  Object.keys(all).forEach(function(k) {
    if (k.indexOf('ttv_') !== 0 && k.indexOf('nicov_') !== 0 && k.indexOf('ytv_') !== 0) return; // TikTok・ニコニコ・YouTubeの保存分
    try {
      const f = DriveApp.getFileById(all[k]);
      if (f.getDateCreated().getTime() < limit) { f.setTrashed(true); props.deleteProperty(k); }
    } catch (e) { props.deleteProperty(k); }
  });
}
// 保存した動画が、Driveで再生できる状態になったかを調べる
//   Drive API の videoMediaMetadata は、動画の処理が終わってから入る(アップロード直後は空)
function checkVideoReady(id, pass, fid) {
  try {
    if (!checkUserOrGuest_(id, pass)) return { error: '認証エラー' };
    fid = String(fid || '');
    if (!/^[A-Za-z0-9_-]{20,}$/.test(fid)) return { error: 'IDが正しくありません' };
    const res = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + fid + '?fields=id,videoMediaMetadata&supportsAllDrives=true',
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }
    );
    if (res.getResponseCode() !== 200) {
      return { ready: false, unknown: true, note: 'status ' + res.getResponseCode() };
    }
    const m = JSON.parse(res.getContentText()).videoMediaMetadata;
    return { ready: !!(m && (m.durationMillis || m.width)) };
  } catch (e) {
    return { ready: false, unknown: true, note: String(e).slice(0, 80) };
  }
}
// ============================================================
// 📥 ニコニコ動画を保存して再生(非公式)
//   ニコニコの動画は、暗号化された細切れの形で配られていて、Apps Script だけでは mp4 にできません。
//   そこで、yt-dlp を動かす小さなサーバー(nico_server フォルダ)に作ってもらい、
//   出来上がった mp4 を Drive に保存します(TikTokと同じ)。
//   スクリプトプロパティに、次の2つを入れておく:
//     NICO_DL_URL … サーバーのアドレス(例: https://xxxx.onrender.com)
//     NICO_DL_KEY … サーバーの NICO_KEY と同じ合言葉
// ============================================================
const NICO_MAX_BYTES = 45 * 1024 * 1024;   // UrlFetchApp の上限(約50MB)より少し小さめに
const NICO_DL_PER_HOUR = 5;                // 1人が1時間に保存できる本数
const NICO_WAIT_MS = 250 * 1000;           // サーバーが取り込むのを待つ時間(画面からの呼び出しは6分まで)
const NICO_ID_RE_ = /^(?:sm|so|nm)\d{1,12}$/;

// サーバーに頼む。無料のサーバーは、眠っていると起きるまで時間がかかるので、通信エラーなら少し待って、数回やり直す
function nicoFetch_(base, key, path, retries) {
  let last = null;
  for (let i = 0; i < (retries || 1); i++) {
    try {
      return UrlFetchApp.fetch(base + path, { headers: { 'X-Nico-Key': key }, muteHttpExceptions: true });
    } catch (e) { last = e; if (i < retries - 1) Utilities.sleep(2000); }
  }
  throw last;
}

function getNicoVideo(id, pass, nicoId) {
  try {
    if (!checkUserOrGuest_(id, pass)) return { error: '認証エラー' };
    const vid = String(nicoId || '').trim().toLowerCase();
    if (!NICO_ID_RE_.test(vid)) return { error: 'ニコニコの動画ID(sm・so・nm のあとに数字)を読み取れませんでした' };

    // 保存済みなら、それを返す
    const props = PropertiesService.getScriptProperties();
    const pk = 'nicov_' + vid;
    const saved = props.getProperty(pk);
    if (saved) {
      try { DriveApp.getFileById(saved); return { fid: saved, cached: true }; }
      catch (e) { props.deleteProperty(pk); }
    }

    const base = String(props.getProperty('NICO_DL_URL') || '').trim().replace(/\/+$/, '');
    const key = String(props.getProperty('NICO_DL_KEY') || '').trim();
    if (!base || !key) return { error: 'ニコニコのダウンローダーが設定されていません(スクリプトプロパティ NICO_DL_URL と NICO_DL_KEY)' };

    const cache = CacheService.getScriptCache();
    const uk = 'nicodl_' + String(id).slice(0, 60);
    if (Number(cache.get(uk) || 0) >= NICO_DL_PER_HOUR) {
      return { error: '保存のしすぎです。しばらくしてからにしてください(1時間に' + NICO_DL_PER_HOUR + '本まで)' };
    }

    // 1) 取り込みを始めてもらう
    let r;
    try { r = nicoFetch_(base, key, '/start?id=' + encodeURIComponent(vid), 3); }
    catch (e) { return { error: 'ダウンローダーに接続できません(サーバーが起動中かもしれません。少し待ってからもう一度)' }; }
    let j = null;
    try { j = JSON.parse(r.getContentText()); } catch (e) {}
    if (r.getResponseCode() === 403) return { error: 'ダウンローダーの合言葉(NICO_DL_KEY)が違います' };
    if (r.getResponseCode() !== 200 || !j || !j.job) {
      return { error: 'ダウンローダーの応答エラー ' + r.getResponseCode() + (j && j.error ? '(' + String(j.error).slice(0, 80) + ')' : '') };
    }

    // 2) 出来上がるまで待つ
    const t0 = Date.now();
    let st = null;
    while (true) {
      const s = nicoFetch_(base, key, '/status?job=' + encodeURIComponent(j.job), 2);
      try { st = JSON.parse(s.getContentText()); } catch (e) { st = null; }
      if (s.getResponseCode() !== 200 || !st) return { error: 'ダウンローダーの応答エラー ' + s.getResponseCode() };
      if (st.state === 'done') break;
      if (st.state === 'error') return { error: 'ニコニコから取れませんでした: ' + String(st.error || '理由不明').slice(0, 120) };
      if (Date.now() - t0 > NICO_WAIT_MS) return { error: '取り込みに時間がかかっています。少し待ってから、もう一度押してください(続きから取り込みます)' };
      Utilities.sleep(3000);
    }
    if (Number(st.size) > NICO_MAX_BYTES) return { error: '動画が大きすぎます(約' + Math.round(st.size / 1048576) + 'MB)' };

    // 3) mp4 を受け取って、Drive に保存
    const f = nicoFetch_(base, key, '/file?job=' + encodeURIComponent(j.job), 2);
    if (f.getResponseCode() !== 200) return { error: '動画の受け取りに失敗しました(' + f.getResponseCode() + ')' };
    const blob = f.getBlob();
    if (blob.getBytes().length > NICO_MAX_BYTES) return { error: '動画が大きすぎます' };
    blob.setName(vid + '.mp4');

    const file = getTTFolder_().createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
    catch (e1) {
      try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); }
      catch (e2) { Logger.log('共有設定スキップ: ' + e2); }
    }
    props.setProperty(pk, file.getId());
    cache.put(uk, String(Number(cache.get(uk) || 0) + 1), 3600);
    return { fid: file.getId() };
  } catch (e) {
    return { error: '保存エラー: ' + String(e).slice(0, 100) };
  }
}

// 小さい動画は、動画データをそのまま返す(TikTokと同じ)。ニコニコは大きいことが多いので、たいていは fid だけを返して、画面が小分けで受け取る
function getNicoVideoFast(id, pass, nicoId) {
  const r = getNicoVideo(id, pass, nicoId);
  if (!r || r.error) return r;
  if (r.cached) return r;
  try {
    const file = DriveApp.getFileById(r.fid);
    if (file.getSize() <= TT_DIRECT_MAX) {
      return { fid: r.fid, mime: 'video/mp4', data: Utilities.base64Encode(file.getBlob().getBytes()) };
    }
  } catch (e) { Logger.log('直接返せませんでした: ' + e); }
  return { fid: r.fid };
}

// エディタから実行して、ダウンローダーにつながるかを確かめる(結果は実行ログに出ます)
function testNicoServer() {
  const props = PropertiesService.getScriptProperties();
  const base = String(props.getProperty('NICO_DL_URL') || '').trim().replace(/\/+$/, '');
  const key = String(props.getProperty('NICO_DL_KEY') || '').trim();
  if (!base || !key) { Logger.log('NICO_DL_URL と NICO_DL_KEY を、スクリプトプロパティに入れてください'); return; }
  const r = nicoFetch_(base, key, '/health', 3);
  Logger.log('/health → ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 200));
  const s = nicoFetch_(base, key, '/start?id=sm9', 1);
  Logger.log('/start(合言葉の確認) → ' + s.getResponseCode() + ' ' + s.getContentText().slice(0, 200));
}

// 小さい動画は、動画データをそのまま返して、Driveの処理待ちなしで再生できるようにする
const TT_DIRECT_MAX = 8 * 1024 * 1024; // これ以下なら、直接返す

function getTikTokVideoFast(id, pass, url) {
  const r = getTikTokVideo(id, pass, url); // 保存まで(既存の関数)
  if (!r || r.error) return r;
  if (r.cached) return r; // 保存済みなら、Driveのプレーヤーでそのまま見られる
  try {
    const file = DriveApp.getFileById(r.fid);
    if (file.getSize() <= TT_DIRECT_MAX) {
      return { fid: r.fid, mime: 'video/mp4', data: Utilities.base64Encode(file.getBlob().getBytes()) };
    }
  } catch (e) { Logger.log('直接返せませんでした: ' + e); }
  return { fid: r.fid };
}
// 保存した動画を、小さく分けて返す(Driveの処理待ちなしで再生するため)
const TT_CHUNK = 3 * 1024 * 1024; // 1回に返す大きさ(3MB)

function getVideoChunk(id, pass, fid, index) {
  try {
    if (!checkUserOrGuest_(id, pass)) return { error: '認証エラー' };
    fid = String(fid || '');
    if (!/^[A-Za-z0-9_-]{20,}$/.test(fid)) return { error: 'IDが正しくありません' };

    // 保存用フォルダの中のファイルだけ読める(他のファイルを読まれないようにする)
    const file = DriveApp.getFileById(fid);
    const folderId = getTTFolder_().getId();
    let inFolder = false;
    const parents = file.getParents();
    while (parents.hasNext()) { if (parents.next().getId() === folderId) { inFolder = true; break; } }
    if (!inFolder) return { error: '対象外のファイルです' };

    const size = file.getSize();
    const i = Number(index) || 0;
    const start = i * TT_CHUNK;
    if (start >= size) return { error: '範囲外です' };
    const end = Math.min(start + TT_CHUNK, size) - 1;

    const r = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + fid + '?alt=media&supportsAllDrives=true',
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), Range: 'bytes=' + start + '-' + end }, muteHttpExceptions: true }
    );
    const code = r.getResponseCode();
    if (code !== 206 && code !== 200) return { error: 'Driveの読み込みエラー ' + code };
    let bytes = r.getBlob().getBytes();
    if (code === 200) bytes = bytes.slice(start, end + 1); // 範囲指定が効かなかったとき
    return { data: Utilities.base64Encode(bytes), chunks: Math.ceil(size / TT_CHUNK), index: i };
  } catch (e) {
    return { error: 'チャンク取得エラー: ' + String(e).slice(0, 80) };
  }
}
// このアプリ自身のURLを返す(画面を開き直すときに使う)
function getAppUrl() {
  try { return ScriptApp.getService().getUrl(); } catch (e) { return ''; }
}
// 保存したTikTok動画を、1回の通信でまとめて返す(小さめの動画用。受け取りを速くする)
const TT_ONECALL_MAX = 12 * 1024 * 1024;

function getSavedVideoFast(id, pass, fid) {
  try {
    if (!checkUserOrGuest_(id, pass)) return { error: '認証エラー' };
    fid = String(fid || '');
    if (!/^[A-Za-z0-9_-]{20,}$/.test(fid)) return { error: 'IDが正しくありません' };

    const cache = CacheService.getScriptCache();
    const file = DriveApp.getFileById(fid);
    // 保存用フォルダの中のファイルだけ読める(一度確認したら6時間は確認を省く)
    if (!cache.get('tvok_' + fid)) {
      const folderId = getTTFolder_().getId();
      let inFolder = false;
      const parents = file.getParents();
      while (parents.hasNext()) { if (parents.next().getId() === folderId) { inFolder = true; break; } }
      if (!inFolder) return { error: '対象外のファイルです' };
      cache.put('tvok_' + fid, '1', 21600);
    }
    if (file.getSize() > TT_ONECALL_MAX) return { big: true };
    return { mime: 'video/mp4', data: Utilities.base64Encode(file.getBlob().getBytes()) };
  } catch (e) {
    return { error: '取得エラー: ' + String(e).slice(0, 80) };
  }
}// ============================================================
// 📜 みんなの「見た動画」(Videos シート。自動作成)
// ============================================================
const VIDEO_SHEET = 'Videos';
const VIDEO_MAX_ROWS = 200;

function getVideoSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(VIDEO_SHEET);
  if (!sh) {
    sh = ss.insertSheet(VIDEO_SHEET);
    sh.appendRow(['更新日時', 'ユーザー', 'キー', '内容']);
  }
  return sh;
}

// 画面から届いた動画の情報を、安全な形に整える(合わないものは null)
function cleanVideoEntry_(e) {
  e = e || {};
  const type = String(e.type || '');
  const o = { type: type };
  if (type === 'yt') {
    const id = String(e.id || '');
    if (!/^[A-Za-z0-9_-]{11}$/.test(id)) return null;
    o.id = id; o.shorts = !!e.shorts;
    return { key: 'yt:' + id, data: o };
  }
  if (type === 'drive') {
    const fid = String(e.fid || '');
    if (!/^[A-Za-z0-9_-]{20,}$/.test(fid)) return null;
    o.fid = fid;
    return { key: 'drive:' + fid, data: o };
  }
  if (type === 'tt') {
    const url = String(e.url || '').slice(0, 300).replace(/^http:\/\//i, 'https://');
    if (!TT_HOST_RE_.test(url)) return null;
    const vid = ttIdOf_(url) || String(e.id || '').replace(/\D/g, '').slice(0, 25);
    o.url = url; o.id = vid;
    return { key: 'tt:' + (vid || url), data: o };
  }
  if (type === 'nico') {
    const id = String(e.id || '');
    if (!/^(?:(?:sm|so|nm)\d{1,12}|\d{1,12})$/.test(id)) return null;
    o.id = id;
    return { key: 'nico:' + id, data: o };
  }
  if (type === 'mp4') {
    const url = String(e.url || '').slice(0, 300);
    if (!/^https?:\/\/[^\s]+$/i.test(url)) return null;
    o.url = url;
    return { key: 'mp4:' + url, data: o };
  }
  return null;
}

// 動画を見たときに、画面から呼ばれる。同じ動画は1行にして、最新の時刻と見た人に更新する
function shareVideo(id, pass, entry) {
  try {
    if (!checkUser_(id, pass)) return { ok: false };
    const c = cleanVideoEntry_(entry);
    if (!c) return { ok: false };
    const lock = LockService.getScriptLock();
    lock.waitLock(8000);
    try {
      const sh = getVideoSheet_();
      const last = sh.getLastRow();
      if (last >= 2) {
        const keys = sh.getRange(2, 3, last - 1, 1).getValues();
        for (let i = keys.length - 1; i >= 0; i--) {
          if (String(keys[i][0]) === c.key) { sh.deleteRow(i + 2); break; }
        }
      }
      sh.appendRow([new Date(), String(id).slice(0, 40), c.key, JSON.stringify(c.data)]);
      const now = sh.getLastRow();
      if (now > VIDEO_MAX_ROWS + 1) sh.deleteRows(2, now - (VIDEO_MAX_ROWS + 1));
    } finally {
      lock.releaseLock();
    }
    return { ok: true };
  } catch (e) {
    return { ok: false };
  }
}

// 「見た動画」を開いたときに、画面から呼ばれる。みんなの動画を新しい順に返す
function getSharedVideos(id, pass, limit) {
  try {
    if (!checkUser_(id, pass)) return { error: '認証エラー' };
    const n = Math.min(Math.max(Number(limit) || 60, 1), 100);
    const sh = getVideoSheet_();
    const last = sh.getLastRow();
    if (last < 2) return { videos: [] };
    const from = Math.max(2, last - n + 1);
    const rows = sh.getRange(from, 1, last - from + 1, 4).getValues();
    const props = PropertiesService.getScriptProperties().getProperties();
    const out = [];
    for (let i = rows.length - 1; i >= 0; i--) {
      let d = null;
      try { d = JSON.parse(rows[i][3]); } catch (e) {}
      if (!d || !d.type) continue;
      const t = new Date(rows[i][0]);
      // TikTokは、だれかがDriveに保存していれば、その保存先を教える(消えていれば空)
      const fid = d.type === 'tt' ? (props['ttv_' + d.id] || '') : (d.type === 'nico' ? (props['nicov_' + d.id] || '') : (d.fid || ''));
      out.push({
        key: String(rows[i][2]), type: d.type, id: d.id || '', fid: fid,
        url: d.url || '', shorts: !!d.shorts,
        user: String(rows[i][1]), ts: isNaN(t.getTime()) ? 0 : t.getTime()
      });
    }
    return { videos: out };
  } catch (e) {
    return { error: '取得エラー: ' + String(e).slice(0, 80) };
  }
}
// ============================================================
// 📺 YouTubeチャンネルの最新動画(公式のRSSを読む。APIキー不要)
//   ・チャンネルのURL / @名前 / チャンネルID のどれでもOK
//   ・同じチャンネルの一覧は10分、@名前から調べたチャンネルIDは6時間覚えておく
// ============================================================
const YT_UA_ = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36';

// ページのHTMLから、チャンネルIDを探す
function ytFindCid_(html) {
  const pats = [
    /<link rel="canonical" href="https:\/\/www\.youtube\.com\/channel\/(UC[A-Za-z0-9_-]{22})"/,
    /<meta itemprop="(?:channelId|identifier)" content="(UC[A-Za-z0-9_-]{22})"/,
    /"externalId":"(UC[A-Za-z0-9_-]{22})"/,
    /"channelId":"(UC[A-Za-z0-9_-]{22})"/,
    /"browseId":"(UC[A-Za-z0-9_-]{22})"/,
    /channel_id=(UC[A-Za-z0-9_-]{22})/
  ];
  for (let i = 0; i < pats.length; i++) { const m = html.match(pats[i]); if (m) return m[1]; }
  return '';
}

// 入力 → チャンネルID(UC…)。見つからなければ ''
//   対応: チャンネルのURL(/channel/ /@名前 /c/名前 /user/名前)、@名前(日本語もOK)、チャンネルID、(APIキーがあれば)チャンネル名
function ytResolveChannel_(input) {
  let s = String(input || '').trim().slice(0, 300);
  try { s = decodeURIComponent(s); } catch (e) {}
  s = s.replace(/[?#].*$/, '').replace(/\/(videos|shorts|streams|featured|live|playlists|about|community)\/?$/, '');
  let m = s.match(/\/channel\/(UC[A-Za-z0-9_-]{22})/) || s.match(/^(UC[A-Za-z0-9_-]{22})$/);
  if (m) return m[1];

  // 調べる方法を決める: [ページのURL, APIの引数]
  let pageUrl = '', apiQuery = '', key = '';
  m = s.match(/youtube\.com\/@([^\/\s]{1,80})/) || s.match(/^@([^\/\s]{1,80})$/);
  if (m) {
    key = 'h_' + m[1].toLowerCase();
    pageUrl = 'https://www.youtube.com/@' + encodeURIComponent(m[1]);
    apiQuery = 'forHandle=' + encodeURIComponent('@' + m[1]);
  } else if ((m = s.match(/youtube\.com\/user\/([^\/\s]{1,80})/))) {
    key = 'u_' + m[1].toLowerCase();
    pageUrl = 'https://www.youtube.com/user/' + encodeURIComponent(m[1]);
    apiQuery = 'forUsername=' + encodeURIComponent(m[1]);
  } else if ((m = s.match(/youtube\.com\/c\/([^\/\s]{1,80})/))) {
    key = 'c_' + m[1].toLowerCase();
    pageUrl = 'https://www.youtube.com/c/' + encodeURIComponent(m[1]);
  } else if (s && !/[\/\s]/.test(s) && s.length <= 80) { // 名前だけ(@なし) → ハンドルとして試す
    key = 'h_' + s.toLowerCase();
    pageUrl = 'https://www.youtube.com/@' + encodeURIComponent(s);
    apiQuery = 'forHandle=' + encodeURIComponent('@' + s);
  } else if (s && !/youtube\.com|youtu\.be/.test(s)) { // 空白のある文章 → チャンネル名の検索
    key = 'q_' + s.toLowerCase().slice(0, 60);
  } else return '';

  const cache = CacheService.getScriptCache();
  const ck = 'ytch2_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, key, Utilities.Charset.UTF_8)).slice(0, 22);
  const hit = cache.get(ck);
  if (hit) return hit;

  let cid = '';
  // 1) ページを読む
  if (pageUrl) {
    try {
      const r = UrlFetchApp.fetch(pageUrl, {
        muteHttpExceptions: true,
        headers: { 'User-Agent': YT_UA_, 'Accept-Language': 'ja', 'Cookie': 'CONSENT=YES+1' }
      });
      if (r.getResponseCode() === 200) cid = ytFindCid_(r.getContentText());
    } catch (e) {}
  }
  // 2) YouTubeの内部API(非公式)に、URLからチャンネルIDを聞く
  if (!cid && pageUrl) {
    try {
      const r = UrlFetchApp.fetch('https://www.youtube.com/youtubei/v1/navigation/resolve_url?prettyPrint=false', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        payload: JSON.stringify({ context: { client: { clientName: 'WEB', clientVersion: '2.20250620.01.00', hl: 'ja', gl: 'JP' } }, url: pageUrl }),
        headers: { 'User-Agent': YT_UA_, 'Origin': 'https://www.youtube.com' }
      });
      if (r.getResponseCode() === 200) {
        const j = JSON.parse(r.getContentText());
        const b = j && j.endpoint && j.endpoint.browseEndpoint && j.endpoint.browseEndpoint.browseId;
        if (b && /^UC[A-Za-z0-9_-]{22}$/.test(b)) cid = b;
      }
    } catch (e) {}
  }
  // 3) 公式API(キーがあるとき)
  const apiKey = PropertiesService.getScriptProperties().getProperty('YT_API_KEY');
  if (!cid && apiKey) {
    try {
      if (apiQuery) {
        const r = UrlFetchApp.fetch('https://www.googleapis.com/youtube/v3/channels?part=id&' + apiQuery + '&key=' + encodeURIComponent(apiKey), { muteHttpExceptions: true });
        if (r.getResponseCode() === 200) {
          const j = JSON.parse(r.getContentText());
          if (j.items && j.items[0] && j.items[0].id) cid = j.items[0].id;
        }
      }
      if (!cid) { // 名前・(/c/の)カスタムURLは、検索で探す(1回で100消費するので、最後の手段)
        const q = (m && m[1]) ? m[1] : s;
        const r = UrlFetchApp.fetch('https://www.googleapis.com/youtube/v3/search?part=snippet&type=channel&maxResults=1&q=' + encodeURIComponent(q) + '&key=' + encodeURIComponent(apiKey), { muteHttpExceptions: true });
        if (r.getResponseCode() === 200) {
          const j = JSON.parse(r.getContentText());
          const it = j.items && j.items[0];
          const c = it && ((it.id && it.id.channelId) || (it.snippet && it.snippet.channelId));
          if (c) cid = c;
        }
      }
    } catch (e) {}
  }
  if (!/^UC[A-Za-z0-9_-]{22}$/.test(cid)) return '';
  cache.put(ck, cid, 21600);
  return cid;
}

// RSS(Atom)の文章 → { name, videos }。読めなければ null
function ytParseAtom_(xml) {
  try {
    const root = XmlService.parse(xml).getRootElement();
    const atom = XmlService.getNamespace('http://www.w3.org/2005/Atom');
    const yt = XmlService.getNamespace('yt', 'http://www.youtube.com/xml/schemas/2015');
    const videos = [];
    root.getChildren('entry', atom).slice(0, 15).forEach(function(e) {
      const vid = e.getChildText('videoId', yt);
      if (!vid) return;
      const link = e.getChild('link', atom);
      const href = link && link.getAttribute('href') ? link.getAttribute('href').getValue() : '';
      const pub = e.getChildText('published', atom) || '';
      const t = new Date(pub);
      videos.push({
        id: vid,
        title: String(e.getChildText('title', atom) || '').slice(0, 100),
        published: pub,
        ts: isNaN(t.getTime()) ? 0 : t.getTime(),
        shorts: /\/shorts\//.test(href)
      });
    });
    return { name: String(root.getChildText('title', atom) || '').slice(0, 60), videos: videos };
  } catch (e) { return null; }
}

// チャンネルのページ(HTML)から、動画のIDとタイトルを拾う(RSSがだめなときの予備。日付は取れない)
// 「3 日前」「2 weeks ago」のような文章 → 投稿時刻(ミリ秒)の目安。読めなければ 0
function ytRelToTs_(t) {
  t = String(t || '');
  let m = t.match(/(\d+)\s*(秒|分|時間|日|週間|週|か月|ヶ月|カ月|年)\s*前/), n, u;
  const ja = { '秒': 1e3, '分': 6e4, '時間': 36e5, '日': 864e5, '週間': 6048e5, '週': 6048e5, 'か月': 2592e6, 'ヶ月': 2592e6, 'カ月': 2592e6, '年': 31536e6 };
  if (m) { n = Number(m[1]); u = ja[m[2]]; }
  else {
    m = t.match(/(\d+)\s*(second|minute|hour|day|week|month|year)s?\s*ago/i);
    if (!m) return 0;
    const en = { second: 1e3, minute: 6e4, hour: 36e5, day: 864e5, week: 6048e5, month: 2592e6, year: 31536e6 };
    n = Number(m[1]); u = en[m[2].toLowerCase()];
  }
  return u ? Date.now() - n * u : 0;
}

// ページのHTMLまたはJSONの文章から、JSON(オブジェクト)を取り出す。だめなら null
function ytJsonFrom_(text) {
  text = String(text || '');
  const t = text.replace(/^\s+/, '');
  if (t.charAt(0) === '{') { try { return JSON.parse(t); } catch (e) { return null; } }
  const k = text.indexOf('ytInitialData');
  if (k < 0) return null;
  const st = text.indexOf('{', k);
  if (st < 0) return null;
  let depth = 0, inStr = false, esc = false;
  for (let i = st; i < text.length; i++) {
    const c = text.charAt(i);
    if (inStr) { if (esc) esc = false; else if (c === '\\') esc = true; else if (c === '"') inStr = false; continue; }
    if (c === '"') inStr = true;
    else if (c === '{') depth++;
    else if (c === '}') { depth--; if (depth === 0) { try { return JSON.parse(text.slice(st, i + 1)); } catch (e) { return null; } } }
  }
  return null;
}

// JSONの中から、動画(通常・新しい形式・ショート・モバイル)を、見つけた順に集める。項目の順番には頼らない
function ytCollectVideos_(root) {
  const videos = [], seen = {};
  const tx = function(o) {
    if (!o) return '';
    if (typeof o === 'string') return o;
    if (o.simpleText) return String(o.simpleText);
    if (Array.isArray(o.runs)) return o.runs.map(function(r) { return (r && r.text) || ''; }).join('');
    if (o.content) return String(o.content);
    return '';
  };
  const add = function(id, title, timeText, shorts, extra) {
    if (!/^[A-Za-z0-9_-]{11}$/.test(String(id || '')) || seen[id]) return;
    seen[id] = 1;
    const ts = ytRelToTs_(timeText);
    const o = { id: id, title: String(title || '').replace(/\s+/g, ' ').trim().slice(0, 100), published: ts ? new Date(ts).toISOString() : '', ts: ts, shorts: !!shorts };
    if (extra) { if (extra.channel) o.channel = String(extra.channel).slice(0, 60); if (extra.duration) o.duration = String(extra.duration).slice(0, 12); }
    videos.push(o);
  };
  const walk = function(n, depth) {
    if (!n || typeof n !== 'object' || depth > 60 || videos.length >= 15) return;
    if (Array.isArray(n)) { for (let i = 0; i < n.length; i++) walk(n[i], depth + 1); return; }
    if (n.videoRenderer && n.videoRenderer.videoId) {
      const v = n.videoRenderer;
      const url = (((v.navigationEndpoint || {}).commandMetadata || {}).webCommandMetadata || {}).url || '';
      add(v.videoId, tx(v.title), tx(v.publishedTimeText), /\/shorts\//.test(url), { channel: tx(v.ownerText) || tx(v.longBylineText) || tx(v.shortBylineText), duration: tx(v.lengthText) });
    } else if (n.videoWithContextRenderer && n.videoWithContextRenderer.videoId) {
      const v = n.videoWithContextRenderer;
      add(v.videoId, tx(v.headline) || tx(v.title), tx(v.publishedTimeText), false);
    } else if (n.lockupViewModel && n.lockupViewModel.contentId && /VIDEO/.test(String(n.lockupViewModel.contentType || ''))) {
      const l = n.lockupViewModel, lm = (l.metadata || {}).lockupMetadataViewModel || {};
      let timeText = '';
      const rows = ((lm.metadata || {}).contentMetadataViewModel || {}).metadataRows || [];
      rows.forEach(function(r) { ((r && r.metadataParts) || []).forEach(function(pt) { const c = tx((pt || {}).text); if (!timeText && ytRelToTs_(c)) timeText = c; }); });
      add(l.contentId, tx(lm.title), timeText, false);
    } else if (n.shortsLockupViewModel) {
      const l = n.shortsLockupViewModel;
      const id = (((l.onTap || {}).innertubeCommand || {}).reelWatchEndpoint || {}).videoId || String(l.entityId || '').replace(/^shorts-shelf-item-/, '');
      add(id, tx((l.overlayMetadata || {}).primaryText), '', true);
    } else {
      for (const k in n) { if (Object.prototype.hasOwnProperty.call(n, k)) walk(n[k], depth + 1); }
    }
  };
  walk(root, 0);
  return videos;
}

// ページ・内部APIの返事 → { name, videos }。JSONとして読めればそちら、だめなら文字の拾い読み
function ytParsePage_(text) {
  const j = ytJsonFrom_(text);
  if (j) {
    const videos = ytCollectVideos_(j);
    if (videos.length) {
      let name = '';
      const cm = (j.metadata || {}).channelMetadataRenderer;
      const mf = ((j.microformat || {}).microformatDataRenderer) || {};
      if (cm && cm.title) name = String(cm.title);
      else if (mf.title) name = String(mf.title);
      else { const nm = String(text).match(/<meta property="og:title" content="([^"]*)"/); if (nm) name = nm[1].replace(/&amp;/g, '&'); }
      return { name: name.replace(/ - YouTube$/, '').slice(0, 60), videos: videos };
    }
  }
  return ytParsePageRegex_(text);
}

function ytParsePageRegex_(html) {
  const videos = [], seen = {};
  const re = /"videoRenderer":\{"videoId":"([A-Za-z0-9_-]{11})"/g;
  let m;
  while ((m = re.exec(html)) && videos.length < 15) {
    if (seen[m[1]]) continue;
    seen[m[1]] = 1;
    const seg = html.slice(m.index, m.index + 1500);
    const tm = seg.match(/"title":\{"runs":\[\{"text":"((?:[^"\\]|\\.)*)"/);
    let title = '';
    if (tm) { try { title = JSON.parse('"' + tm[1] + '"'); } catch (e) { title = tm[1]; } }
    videos.push({ id: m[1], title: String(title).slice(0, 100), published: '', ts: 0, shorts: false });
  }
  // 新しい表示形式(lockupViewModel)の動画
  const re2 = /"contentId":"([A-Za-z0-9_-]{11})","contentType":"LOCKUP_CONTENT_TYPE_VIDEO"/g;
  while ((m = re2.exec(html)) && videos.length < 15) {
    if (seen[m[1]]) continue;
    seen[m[1]] = 1;
    const seg = html.slice(m.index, m.index + 3000);
    const tm = seg.match(/"lockupMetadataViewModel":\{"title":\{"content":"((?:[^"\\]|\\.)*)"/);
    let title = '';
    if (tm) { try { title = JSON.parse('"' + tm[1] + '"'); } catch (e) { title = tm[1]; } }
    videos.push({ id: m[1], title: String(title).slice(0, 100), published: '', ts: 0, shorts: false });
  }
  let name = '';
  const nm = html.match(/<meta property="og:title" content="([^"]*)"/);
  if (nm) name = nm[1].replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;/g, "'").slice(0, 60);
  return { name: name, videos: videos };
}

// ★ 非公式の公開サーバー(YouTubeが断ったときの予備。止まることがあります。動かないものは消して、新しいものを足してください)
//   Invidious: https://docs.invidious.io/instances/ / Piped: https://github.com/TeamPiped/Piped/wiki
const YT_INVIDIOUS_HOSTS = [
  'https://inv.nadeko.net', 'https://invidious.nerdvpn.de', 'https://invidious.tiekoetter.com',
  'https://invidious.f5.si', 'https://inv.zoomerville.com', 'https://yt.chocolatemoo53.com'
];
const YT_PIPED_HOSTS = [
  'https://pipedapi.kavin.rocks', 'https://pipedapi.leptons.xyz', 'https://pipedapi.nosebs.ru',
  'https://piped-api.privacy.com.de', 'https://pipedapi.adminforge.de', 'https://api.piped.yt',
  'https://pipedapi.drgns.space', 'https://pipedapi.owo.si'
];

// YouTube Data API v3(公式)。スクリプトプロパティ YT_API_KEY にキーを入れると使える。なければ何もしない
function ytFromApi_(cid) {
  const key = PropertiesService.getScriptProperties().getProperty('YT_API_KEY');
  if (!key) return { skipped: true };
  try {
    const r = UrlFetchApp.fetch('https://www.googleapis.com/youtube/v3/playlistItems?part=snippet,contentDetails&maxResults=15&playlistId=UU'
      + encodeURIComponent(cid.slice(2)) + '&key=' + encodeURIComponent(key), { muteHttpExceptions: true });
    const code = r.getResponseCode();
    let j = null;
    try { j = JSON.parse(r.getContentText()); } catch (e) {}
    if (code !== 200 || !j) {
      const reason = j && j.error && j.error.errors && j.error.errors[0] ? j.error.errors[0].reason : '';
      return { note: 'API: ' + code + (reason ? ' ' + reason : '') };
    }
    const videos = [];
    let name = '';
    (j.items || []).forEach(function(it) {
      const sn = it.snippet || {}, cd = it.contentDetails || {};
      const vid = cd.videoId || (sn.resourceId && sn.resourceId.videoId) || '';
      if (!/^[A-Za-z0-9_-]{11}$/.test(vid)) return;
      if (sn.title === 'Private video' || sn.title === 'Deleted video') return;
      if (!name && sn.channelTitle) name = String(sn.channelTitle).slice(0, 60);
      const pub = cd.videoPublishedAt || sn.publishedAt || '';
      const t = new Date(pub);
      videos.push({ id: vid, title: String(sn.title || '').slice(0, 100), published: pub, ts: isNaN(t.getTime()) ? 0 : t.getTime(), shorts: false });
    });
    if (!videos.length) return { note: 'API: 動画なし' };
    return { videos: videos, name: name, source: 'YouTube Data API' };
  } catch (e) {
    return { note: 'API: 接続できません' };
  }
}

// YouTube自身の内部API(youtubei。非公式。アプリやサイトが使っているもの)。ブラウザ用とAndroid用を順に試す
function ytFromInnertube_(cid) {
  const clients = [
    { ctx: { clientName: 'WEB', clientVersion: '2.20250620.01.00', hl: 'ja', gl: 'JP' }, ua: YT_UA_, name: 'WEB' },
    { ctx: { clientName: 'ANDROID', clientVersion: '19.44.38', androidSdkVersion: 34, hl: 'ja', gl: 'JP' }, ua: 'com.google.android.youtube/19.44.38 (Linux; U; Android 14) gzip', name: 'ANDROID' }
  ];
  const notes = [];
  for (let i = 0; i < clients.length; i++) {
    try {
      const r = UrlFetchApp.fetch('https://www.youtube.com/youtubei/v1/browse?prettyPrint=false', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        payload: JSON.stringify({ context: { client: clients[i].ctx }, browseId: cid, params: 'EgZ2aWRlb3PyBgQKAjoA' }),
        headers: { 'User-Agent': clients[i].ua, 'Origin': 'https://www.youtube.com', 'X-Youtube-Client-Name': clients[i].name === 'WEB' ? '1' : '3', 'X-Youtube-Client-Version': clients[i].ctx.clientVersion }
      });
      const code = r.getResponseCode();
      if (code !== 200) { notes.push('内部API(' + clients[i].name + '): ' + code); continue; }
      const text = r.getContentText();
      const p = ytParsePage_(text);
      if (p.videos.length) {
        const nm = text.match(/"channelMetadataRenderer":\{"title":"((?:[^"\\]|\\.)*)"/);
        let name = '';
        if (nm) { try { name = JSON.parse('"' + nm[1] + '"'); } catch (e) { name = nm[1]; } }
        return { videos: p.videos, name: String(name).slice(0, 60), source: '内部API(' + clients[i].name + ')' };
      }
      notes.push('内部API(' + clients[i].name + '): 動画なし');
    } catch (e) { notes.push('内部API(' + clients[i].name + '): 接続できません'); }
  }
  return { notes: notes };
}

function ytFromMirrors_(cid) {
  const reqs = [], names = [];
  const opt = { muteHttpExceptions: true, headers: { 'User-Agent': YT_UA_, 'Accept': 'application/json' } };
  YT_INVIDIOUS_HOSTS.forEach(function(h) {
    reqs.push(Object.assign({ url: h + '/api/v1/channels/' + cid + '/latest' }, opt)); names.push('Invidious(' + ttHostName_(h) + ')');
  });
  YT_PIPED_HOSTS.forEach(function(h) {
    reqs.push(Object.assign({ url: h + '/channel/' + cid }, opt)); names.push('Piped(' + ttHostName_(h) + ')');
  });
  if (!reqs.length) return { notes: [] };
  let res;
  try { res = UrlFetchApp.fetchAll(reqs); }
  catch (e) { res = reqs.map(function(q) { try { return UrlFetchApp.fetch(q.url, opt); } catch (e2) { return null; } }); }
  const notes = [];
  for (let i = 0; i < reqs.length; i++) {
    const r = res[i];
    if (!r) { notes.push(names[i] + ': 接続できません'); continue; }
    if (r.getResponseCode() !== 200) { notes.push(names[i] + ': ' + r.getResponseCode()); continue; }
    let j = null;
    try { j = JSON.parse(r.getContentText()); } catch (e) {}
    const videos = [];
    let name = '';
    if (Array.isArray(j)) { // Invidious
      j.slice(0, 15).forEach(function(v) {
        if (!v || !/^[A-Za-z0-9_-]{11}$/.test(String(v.videoId || ''))) return;
        const t = Number(v.published) * 1000;
        if (!name && v.author) name = String(v.author).slice(0, 60);
        videos.push({ id: v.videoId, title: String(v.title || '').slice(0, 100), published: t > 0 ? new Date(t).toISOString() : '', ts: t > 0 ? t : 0, shorts: false });
      });
    } else if (j && Array.isArray(j.relatedStreams)) { // Piped
      name = String(j.name || '').slice(0, 60);
      j.relatedStreams.slice(0, 15).forEach(function(v) {
        const m = String((v && v.url) || '').match(/[?&]v=([A-Za-z0-9_-]{11})/);
        if (!m) return;
        const t = Number(v.uploaded);
        videos.push({ id: m[1], title: String(v.title || '').slice(0, 100), published: t > 0 ? new Date(t).toISOString() : '', ts: t > 0 ? t : 0, shorts: !!v.isShort });
      });
    }
    if (videos.length) return { videos: videos, name: name, source: names[i] };
    notes.push(names[i] + ': 動画なし');
  }
  return { notes: notes };
}

// ============================================================
// 🔎 YouTubeの動画を、言葉で検索する(音楽プレイヤー用)
//   順番(取れたところで止まる): 公式API(YT_API_KEY があるとき) → 非公式の公開サーバー → YouTubeの内部API(非公式)
// ============================================================
const YT_SEARCH_PER_HOUR = 40;

function ytDecode_(s) {
  return String(s || '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'");
}
function ytSecToText_(sec) {
  sec = Number(sec) || 0;
  if (sec <= 0) return '';
  const h = Math.floor(sec / 3600), m = Math.floor((sec % 3600) / 60), s = sec % 60;
  return (h ? h + ':' + ('0' + m).slice(-2) : m) + ':' + ('0' + s).slice(-2);
}

function ytSearchApi_(q) {
  const key = PropertiesService.getScriptProperties().getProperty('YT_API_KEY');
  if (!key) return { skipped: true };
  try {
    const r = UrlFetchApp.fetch('https://www.googleapis.com/youtube/v3/search?part=snippet&type=video&maxResults=15&q=' + encodeURIComponent(q) + '&key=' + encodeURIComponent(key), { muteHttpExceptions: true });
    let j = null; try { j = JSON.parse(r.getContentText()); } catch (e) {}
    if (r.getResponseCode() !== 200 || !j) {
      const reason = j && j.error && j.error.errors && j.error.errors[0] ? j.error.errors[0].reason : '';
      return { note: '検索API: ' + r.getResponseCode() + (reason ? ' ' + reason : '') };
    }
    const videos = (j.items || []).map(function(it) {
      const sn = it.snippet || {};
      return { id: (it.id && it.id.videoId) || '', title: ytDecode_(sn.title).slice(0, 100), channel: ytDecode_(sn.channelTitle).slice(0, 60), duration: '' };
    }).filter(function(v) { return /^[A-Za-z0-9_-]{11}$/.test(v.id); });
    return videos.length ? { videos: videos, source: 'YouTube Data API' } : { note: '検索API: 結果なし' };
  } catch (e) { return { note: '検索API: 接続できません' }; }
}

function ytSearchMirrors_(q) {
  const reqs = [], names = [];
  const opt = { muteHttpExceptions: true, headers: { 'User-Agent': YT_UA_, 'Accept': 'application/json' } };
  YT_INVIDIOUS_HOSTS.forEach(function(h) { reqs.push(Object.assign({ url: h + '/api/v1/search?type=video&q=' + encodeURIComponent(q) }, opt)); names.push('Invidious(' + ttHostName_(h) + ')'); });
  YT_PIPED_HOSTS.forEach(function(h) { reqs.push(Object.assign({ url: h + '/search?filter=videos&q=' + encodeURIComponent(q) }, opt)); names.push('Piped(' + ttHostName_(h) + ')'); });
  if (!reqs.length) return { notes: [] };
  let res;
  try { res = UrlFetchApp.fetchAll(reqs); }
  catch (e) { res = reqs.map(function(x) { try { return UrlFetchApp.fetch(x.url, opt); } catch (e2) { return null; } }); }
  const notes = [];
  for (let i = 0; i < reqs.length; i++) {
    const r = res[i];
    if (!r) { notes.push(names[i] + ': 接続できません'); continue; }
    if (r.getResponseCode() !== 200) { notes.push(names[i] + ': ' + r.getResponseCode()); continue; }
    let j = null; try { j = JSON.parse(r.getContentText()); } catch (e) {}
    const videos = [];
    if (Array.isArray(j)) { // Invidious
      j.forEach(function(v) {
        if (!v || v.type !== 'video' || !/^[A-Za-z0-9_-]{11}$/.test(String(v.videoId || ''))) return;
        videos.push({ id: v.videoId, title: String(v.title || '').slice(0, 100), channel: String(v.author || '').slice(0, 60), duration: ytSecToText_(v.lengthSeconds) });
      });
    } else if (j && Array.isArray(j.items)) { // Piped
      j.items.forEach(function(v) {
        const m = String((v && v.url) || '').match(/[?&]v=([A-Za-z0-9_-]{11})/);
        if (!m || (v.type && v.type !== 'stream') || v.isShort) return;
        videos.push({ id: m[1], title: String(v.title || '').slice(0, 100), channel: String(v.uploaderName || '').slice(0, 60), duration: ytSecToText_(v.duration) });
      });
    }
    if (videos.length) return { videos: videos.slice(0, 15), source: names[i] };
    notes.push(names[i] + ': 結果なし');
  }
  return { notes: notes };
}

function ytSearchInnertube_(q) {
  try {
    const r = UrlFetchApp.fetch('https://www.youtube.com/youtubei/v1/search?prettyPrint=false', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ context: { client: { clientName: 'WEB', clientVersion: '2.20250620.01.00', hl: 'ja', gl: 'JP' } }, query: q, params: 'EgIQAQ%3D%3D' }),
      headers: { 'User-Agent': YT_UA_, 'Origin': 'https://www.youtube.com' }
    });
    if (r.getResponseCode() !== 200) return { note: '内部API(検索): ' + r.getResponseCode() };
    const j = JSON.parse(r.getContentText());
    const videos = ytCollectVideos_(j).filter(function(v) { return !v.shorts; }).map(function(v) {
      return { id: v.id, title: v.title, channel: v.channel || '', duration: v.duration || '' };
    });
    return videos.length ? { videos: videos, source: '内部API' } : { note: '内部API(検索): 結果なし' };
  } catch (e) { return { note: '内部API(検索): 接続できません' }; }
}

// 画面から呼ばれる。言葉でYouTubeの動画を探して、最大15本を返す
function searchYouTubeVideos(id, pass, query) {
  try {
    if (!checkUser_(id, pass)) return { error: '認証エラー' };
    const q = String(query || '').replace(/\s+/g, ' ').trim().slice(0, 100);
    if (!q) return { error: '検索ワードを入れてください' };

    const cache = CacheService.getScriptCache();
    const ck = 'yts_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, q.toLowerCase(), Utilities.Charset.UTF_8)).slice(0, 22);
    const hit = cache.get(ck);
    if (hit) { try { return JSON.parse(hit); } catch (e) {} }

    const uk = 'ytsc_' + String(id).slice(0, 60);
    if (Number(cache.get(uk) || 0) >= YT_SEARCH_PER_HOUR) return { error: '検索のしすぎです。しばらくしてからにしてください(1時間に' + YT_SEARCH_PER_HOUR + '回まで)' };

    const notes = [];
    let found = null;
    const take = function(r) {
      if (found || !r) return;
      if (r.videos && r.videos.length) found = r;
      else { if (r.note) notes.push(r.note); (r.notes || []).forEach(function(n) { notes.push(n); }); }
    };
    take(ytSearchApi_(q));
    if (!found) take(ytSearchMirrors_(q));
    if (!found) take(ytSearchInnertube_(q));
    if (!found) return { error: '検索できませんでした(' + notes.join(' / ') + ')。時間をおいて試してください' };

    const out = { videos: found.videos, source: found.source };
    cache.put(ck, JSON.stringify(out), 600);
    cache.put(uk, String(Number(cache.get(uk) || 0) + 1), 3600);
    return out;
  } catch (e) {
    return { error: '検索エラー: ' + String(e).slice(0, 80) };
  }
}

// 画面から呼ばれる。チャンネルの最新動画(最大15本)を返す
//   順番(取れたところで止まる):
//   1. YouTube Data API(公式。YT_API_KEY があるときだけ)
//   2. 非公式の公開サーバー(Invidious・Piped)
//   3. YouTubeの内部API(非公式)
//   4. 公式のRSS(チャンネル用・アップロード一覧用)
//   5. チャンネルのページ
//   6. 前に取れたときの一覧(最長6時間)
function getChannelVideos(id, pass, input) {
  try {
    if (!checkUser_(id, pass)) return { error: '認証エラー' };
    const cid = ytResolveChannel_(input);
    if (!cid) return { error: 'チャンネルを見つけられませんでした。「youtube.com/channel/UC…」の形のURLなら確実です。@名前やチャンネル名は、YouTubeが断ることがあります(APIキーを入れると探せます)' };

    const cache = CacheService.getScriptCache();
    const ck = 'ytc_' + cid;
    const hit = cache.get(ck);
    if (hit) { try { return JSON.parse(hit); } catch (e) {} }

    const notes = [];
    let found = null;
    const take = function(r) { // 取れていれば found に入れる。取れなければ理由を notes に足す
      if (found || !r) return;
      if (r.videos && r.videos.length) found = { name: r.name || '', videos: r.videos, source: r.source || '' };
      else { if (r.note) notes.push(r.note); (r.notes || []).forEach(function(n) { notes.push(n); }); }
    };

    take(ytFromApi_(cid));
    if (!found) take(ytFromMirrors_(cid));
    if (!found) take(ytFromInnertube_(cid));

    // 公式のRSS(いまはApps Scriptからだと断られることが多い)
    const feeds = [
      ['RSS', 'https://www.youtube.com/feeds/videos.xml?channel_id=' + cid],
      ['アップロード一覧RSS', 'https://www.youtube.com/feeds/videos.xml?playlist_id=UU' + cid.slice(2)]
    ];
    for (let i = 0; i < feeds.length && !found; i++) {
      try {
        const r = UrlFetchApp.fetch(feeds[i][1], { muteHttpExceptions: true });
        const code = r.getResponseCode();
        if (code === 200) {
          const p = ytParseAtom_(r.getContentText());
          if (p && p.videos.length) { found = { name: p.name, videos: p.videos, source: 'RSS' }; break; }
          notes.push(feeds[i][0] + ': 動画なし');
        } else notes.push(feeds[i][0] + ': ' + code);
      } catch (e) { notes.push(feeds[i][0] + ': 接続できません'); }
    }

    if (!found) {
      try {
        const r = UrlFetchApp.fetch('https://www.youtube.com/channel/' + cid + '/videos?hl=ja', {
          muteHttpExceptions: true,
          headers: { 'User-Agent': YT_UA_, 'Accept-Language': 'ja', 'Cookie': 'CONSENT=YES+1' }
        });
        const code = r.getResponseCode();
        if (code === 200) {
          const p = ytParsePage_(r.getContentText());
          if (p.videos.length) found = { name: p.name, videos: p.videos, source: 'ページ' };
          else {
            const body = r.getContentText();
            notes.push(p.name ? 'ページ: チャンネルはありますが、動画が見つかりません'
              : (/consent\.youtube\.com/.test(body) ? 'ページ: 同意画面が出ました' : 'ページ: チャンネルの中身が入っていないページが返りました(' + body.length + '文字)'));
          }
        } else notes.push('ページ: ' + code);
      } catch (e) { notes.push('ページ: 接続できません'); }
    }

    if (!found) {
      const old = cache.get('ytcl_' + cid);
      if (old) { try { const o = JSON.parse(old); o.stale = true; return o; } catch (e) {} }
      return { error: 'チャンネルの動画を取れませんでした(' + notes.join(' / ') + ')。時間をおいて試してください' };
    }

    const out = { channel: { id: cid, name: found.name }, videos: found.videos };
    if (found.source) out.source = found.source;
    cache.put(ck, JSON.stringify(out), 600);
    cache.put('ytcl_' + cid, JSON.stringify(out), 21600);
    return out;
  } catch (e) {
    return { error: '取得エラー: ' + String(e).slice(0, 80) };
  }
}
// ============================================================
// 🎵 TikTokアカウントの最新動画(複数の取得元を、順に試す)
//   1. tikwm.com(非公式)
//   2. RSSHub の公開サーバー(/tiktok/user/@名前 のRSS)
//   3. ProxiTok の公開サーバー(/@名前/rss のRSS)
//   4. Apify(最後の手段。スクリプトプロパティに APIFY_TOKEN が必要。1日の回数制限つき)
//   ・どれか1つで取れれば、それを使う。同じアカウントの一覧は10分(Apifyは30分)覚えておく
//   ・全部だめだったときは、1分だけ「失敗」を覚えて、同じところに続けて頼まないようにする
//   ・1〜3は非公式・無料の公開サーバーなので、止まることがある。
//     動かなくなったサーバーは、下の2つの一覧から消して、新しいものを足してください
//     (どのサーバーが生きているかは、エディタから testTikTokHosts() を実行すると分かります)
// ============================================================
const TT_USER_PER_HOUR = 20;

// ★ 取得元の一覧(編集してOK。上にあるものほど先に試す)
const TT_RSSHUB_HOSTS = [
  'https://rsshub.app',
  'https://rsshub.rssforever.com',
  'https://hub.slarker.me'
];
const TT_PROXITOK_HOSTS = [];   // 接続できなかったので空にしています。生きているサーバーが見つかったら足してください

// ★ Apify を1日に使える回数(全体)。無料枠を使い切らないための安全装置
const APIFY_DAILY_MAX = 10;

// TikTokの動画IDには投稿時刻(秒)が入っている(上位の桁)。RSSに日付がないときの代わりに使う
function ttIdToMs_(vid) {
  try { return Number(BigInt(String(vid)) >> BigInt(32)) * 1000; } catch (e) { return 0; }
}

function ttStripHtml_(s) {
  return String(s || '')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/<[^>]*>/g, ' ')
    .replace(/&nbsp;/g, ' ').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&')
    .replace(/\s+/g, ' ').trim();
}

function ttHostName_(h) { return String(h).replace(/^https?:\/\//, '').replace(/\/.*$/, ''); }

// RSS(RSSHub・ProxiTok)の文章から、動画の一覧を作る
function ttParseFeed_(text, uid) {
  const out = [], seen = {};
  const push = function(vid, title, pub, cover) {
    if (!/^\d{10,25}$/.test(vid) || seen[vid]) return;
    seen[vid] = 1;
    let ts = 0;
    if (pub) { const d = new Date(pub); if (!isNaN(d.getTime())) ts = d.getTime(); }
    if (!ts) ts = ttIdToMs_(vid);
    out.push({
      id: vid,
      title: String(title || '').slice(0, 100),
      ts: ts,
      published: ts ? new Date(ts).toISOString() : '',
      cover: /^https:\/\//.test(cover || '') ? cover : ''
    });
  };
  text = String(text || '');
  try {
    const root = XmlService.parse(text).getRootElement();
    const ch = root.getChild('channel');
    const items = ch ? ch.getChildren('item') : [];
    items.forEach(function(it) {
      const link = it.getChildText('link') || '';
      const guid = it.getChildText('guid') || '';
      const m = (link + ' ' + guid).match(/\/video\/(\d{10,25})/);
      if (!m) return;
      const desc = it.getChildText('description') || '';
      let cover = '';
      const im = desc.match(/<img[^>]+src=["']([^"']+)["']/i);
      if (im) cover = im[1].replace(/&amp;/g, '&');
      push(m[1], ttStripHtml_(it.getChildText('title') || desc), it.getChildText('pubDate'), cover);
    });
  } catch (e) {}
  if (!out.length) { // XMLとして読めない・形が違うとき: 文章の中の動画URLから、IDだけ拾う
    const re = /\/video\/(\d{10,25})/g;
    let m;
    while ((m = re.exec(text)) && out.length < 30) push(m[1], '', '', '');
  }
  out.sort(function(a, b) { return b.ts - a.ts; });
  return out.slice(0, 12);
}

// RSSHub と ProxiTok を、まとめて同時に頼んで、上にあるものから順に、使えるものを採用する
function ttUserFromFeeds_(uid) {
  const u = encodeURIComponent(uid);
  const srcs = [];
  TT_RSSHUB_HOSTS.forEach(function(h) {
    srcs.push({ name: 'RSSHub(' + ttHostName_(h) + ')', url: String(h).replace(/\/$/, '') + '/tiktok/user/@' + u });
  });
  TT_PROXITOK_HOSTS.forEach(function(h) {
    srcs.push({ name: 'ProxiTok(' + ttHostName_(h) + ')', url: String(h).replace(/\/$/, '') + '/@' + u + '/rss' });
  });
  if (!srcs.length) return { notes: ['RSSHub / ProxiTok の取得元が設定されていません'] };

  const opt = {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { 'User-Agent': TT_UA_, 'Accept': 'application/rss+xml, application/xml, text/xml, */*' }
  };
  let responses;
  try {
    responses = UrlFetchApp.fetchAll(srcs.map(function(s) { return Object.assign({ url: s.url }, opt); }));
  } catch (e) {
    // 1つでも接続できないものがあると、まとめての取得は全体が失敗することがある → 1つずつ試す
    responses = srcs.map(function(s) {
      try { return UrlFetchApp.fetch(s.url, opt); } catch (e2) { return null; }
    });
  }

  const notes = [];
  for (let i = 0; i < srcs.length; i++) {
    const r = responses[i];
    if (!r) { notes.push(srcs[i].name + ': 接続できません'); continue; }
    const code = r.getResponseCode();
    if (code !== 200) { notes.push(srcs[i].name + ': ' + code); continue; }
    let videos = [];
    try { videos = ttParseFeed_(r.getContentText(), uid); } catch (e) {}
    if (videos.length) return { videos: videos, source: srcs[i].name };
    notes.push(srcs[i].name + ': 動画が見つかりません');
  }
  return { notes: notes };
}

// tikwm(非公式)で取る
function ttUserFromTikwm_(uid) {
  try {
    let j = null, http = 0;
    const hosts = ['https://www.tikwm.com', 'https://tikwm.com'];
    outer:
    for (let h = 0; h < hosts.length; h++) {
      for (let i = 0; i < 2; i++) {
        const res = UrlFetchApp.fetch(hosts[h] + '/api/user/posts', {
          method: 'post',
          contentType: 'application/x-www-form-urlencoded; charset=UTF-8',
          payload: { unique_id: '@' + uid, count: 12, cursor: 0 },
          headers: {
            'User-Agent': TT_UA_,
            'Referer': 'https://www.tikwm.com/',
            'Origin': 'https://www.tikwm.com',
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Cookie': 'current_language=en',
            'X-Requested-With': 'XMLHttpRequest'
          },
          muteHttpExceptions: true
        });
        http = res.getResponseCode();
        j = null;
        try { j = JSON.parse(res.getContentText()); } catch (e) {}
        const limited = j && j.code !== 0 && /limit/i.test(String(j.msg || ''));
        if (http === 200 && !limited) break outer;
        if (limited && i === 0) { Utilities.sleep(1500); continue; }
        break; // 断られた・エラー → 次のアドレスを試す
      }
    }
    if (http !== 200) return { note: 'tikwm: ' + http };
    if (!j || j.code !== 0 || !j.data) return { note: 'tikwm: ' + String((j && j.msg) || '取得できません').slice(0, 60) };

    const list = Array.isArray(j.data.videos) ? j.data.videos : [];
    const videos = list.slice(0, 12).map(function(v) {
      const vid = String(v.video_id || v.id || '');
      let cover = String(v.cover || v.origin_cover || '');
      if (cover.charAt(0) === '/') cover = 'https://www.tikwm.com' + cover;
      const t = Number(v.create_time) * 1000;
      return {
        id: vid,
        title: String(v.title || '').slice(0, 100),
        ts: isFinite(t) ? t : 0,
        published: isFinite(t) && t > 0 ? new Date(t).toISOString() : '',
        cover: /^https:\/\//.test(cover) ? cover : ''
      };
    }).filter(function(v) { return /^\d{10,25}$/.test(v.id); });
    if (!videos.length) return { note: 'tikwm: 動画が見つかりません' };
    return { videos: videos, source: 'tikwm' };
  } catch (e) {
    return { note: 'tikwm: ' + String(e).slice(0, 60) };
  }
}

// ------------------------------------------------------------
// 🔑 Apify のトークンを読む
//   スクリプトプロパティの名前は APIFY_TOKEN がおすすめ。
//   API_TOKEN / APIFY_API_TOKEN / 名前に APIFY が入っているものでも見つけられる
//   (名前の前後の空白や、大文字小文字・全角のちがいは無視する)
// ------------------------------------------------------------
function getApifyToken_() {
  const all = PropertiesService.getScriptProperties().getProperties();
  const norm = function(x) { return String(x).replace(/[\s\u3000]/g, '').toUpperCase(); };
  const clean = function(v) { return String(v).replace(/[\s\u3000"']/g, ''); };
  const wants = ['APIFY_TOKEN', 'APIFY_API_TOKEN', 'APIFY_KEY', 'API_TOKEN'];
  const names = Object.keys(all);
  for (let w = 0; w < wants.length; w++) {
    for (let i = 0; i < names.length; i++) {
      if (norm(names[i]) === wants[w]) {
        const v = clean(all[names[i]]);
        if (v) return v;
      }
    }
  }
  for (let i = 0; i < names.length; i++) {
    if (norm(names[i]).indexOf('APIFY') !== -1) {
      const v = clean(all[names[i]]);
      if (v) return v;
    }
  }
  return '';
}

// 最後の手段: Apify の TikTok スクレイパー(無料の取得元が全部だめなときだけ使う)
function ttUserFromApify_(uid) {
  const token = getApifyToken_();
  if (!token) return { note: 'Apify: トークンが未設定です(スクリプトプロパティに APIFY_TOKEN を入れてください)' };
  try {
    // 1日の回数制限(無料枠を使い切らないため)
    const props = PropertiesService.getScriptProperties();
    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
    let d = {};
    try { d = JSON.parse(props.getProperty('APIFY_DAILY') || '{}') || {}; } catch (e) {}
    if (d.day !== today) d = { day: today, n: 0 };
    if (d.n >= APIFY_DAILY_MAX) return { note: 'Apify: 今日の上限(' + APIFY_DAILY_MAX + '回)に達しました' };
    d.n++;
    props.setProperty('APIFY_DAILY', JSON.stringify(d)); // 実行する前に数える(失敗しても料金がかかることがあるため)

    const url = 'https://api.apify.com/v2/acts/clockworks~tiktok-scraper/run-sync-get-dataset-items'
      + '?timeout=90&token=' + encodeURIComponent(token);
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        profiles: [uid],
        resultsPerPage: 12,
        shouldDownloadVideos: false,
        shouldDownloadCovers: false
      }),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code === 401 || code === 403) return { note: 'Apify: トークンが正しくありません(' + code + ')' };
    if (code === 402) return { note: 'Apify: 無料枠を使い切りました' };
    if (code !== 200 && code !== 201) return { note: 'Apify: ' + code };

    let list = null;
    try { list = JSON.parse(res.getContentText()); } catch (e) {}
    if (!Array.isArray(list) || !list.length) return { note: 'Apify: 動画が見つかりません' };
    Logger.log('Apify 1件目のキー: ' + Object.keys(list[0]).join(','));

    const videos = list.slice(0, 12).map(function(v) {
      const vid = String(v.id || '');
      let ts = 0;
      if (v.createTimeISO) { const dt = new Date(v.createTimeISO); if (!isNaN(dt.getTime())) ts = dt.getTime(); }
      if (!ts && v.createTime) ts = Number(v.createTime) * 1000 || 0;
      if (!ts) ts = ttIdToMs_(vid);
      const cover = String((v.videoMeta && v.videoMeta.coverUrl) || v.coverUrl || '');
      return {
        id: vid,
        title: String(v.text || '').slice(0, 100),
        ts: ts,
        published: ts ? new Date(ts).toISOString() : '',
        cover: /^https:\/\//.test(cover) ? cover : ''
      };
    }).filter(function(v) { return /^\d{10,25}$/.test(v.id); });

    if (!videos.length) return { note: 'Apify: 動画IDを読み取れません' };
    videos.sort(function(a, b) { return b.ts - a.ts; });
    return { videos: videos, source: 'Apify' };
  } catch (e) {
    return { note: 'Apify: ' + String(e).slice(0, 60) };
  }
}

// 画面から呼ばれる。TikTokアカウントの最新動画(最大12本)を返す
function getTikTokUserVideos(id, pass, input) {
  try {
    if (!checkUser_(id, pass)) return { error: '認証エラー' };
    const s = String(input || '').trim().slice(0, 200);
    const m = s.match(/tiktok\.com\/@([A-Za-z0-9._]{2,30})/) || s.match(/^@?([A-Za-z0-9._]{2,30})$/);
    if (!m) return { error: '@アカウント名か、プロフィールのURLを入れてください' };
    const uid = m[1];

    const cache = CacheService.getScriptCache();
    const ck = 'ttu_' + uid.toLowerCase();
    const hit = cache.get(ck);
    if (hit) { try { return JSON.parse(hit); } catch (e) {} }

    // 直前に全部失敗したばかりなら、同じところに続けて頼まない(1分)
    const fk = 'ttuf_' + uid.toLowerCase();
    const failHit = cache.get(fk);
    if (failHit) return { error: failHit };

    const uk = 'ttuc_' + String(id).slice(0, 60);
    if (Number(cache.get(uk) || 0) >= TT_USER_PER_HOUR) {
      return { error: '取得のしすぎです。しばらくしてからにしてください(1時間に' + TT_USER_PER_HOUR + '回まで)' };
    }

    const notes = [];
    let found = ttUserFromTikwm_(uid);
    if (!(found.videos && found.videos.length)) {
      if (found.note) notes.push(found.note);
      found = ttUserFromFeeds_(uid);
      if (found.notes) found.notes.forEach(function(n) { notes.push(n); });
    }

    // 無料の取得元が全部だめだったときだけ、最後の手段として Apify を使う
    if (!(found.videos && found.videos.length)) {
      const ap = ttUserFromApify_(uid);
      if (ap.videos && ap.videos.length) found = ap;
      else if (ap.note) notes.push(ap.note);
    }

    if (!(found.videos && found.videos.length)) {
      const msg = '最新動画を取れませんでした(' + notes.join(' / ') + ')。時間をおくか、取得元の一覧(TT_RSSHUB_HOSTS / TT_PROXITOK_HOSTS)を見直してください';
      cache.put(fk, msg.slice(0, 900), 60);
      return { error: msg };
    }

    const out = { account: { id: uid }, videos: found.videos, source: found.source };
    cache.put(ck, JSON.stringify(out), found.source === 'Apify' ? 1800 : 600); // Apifyは課金を減らすため長めに覚える
    cache.put(uk, String(Number(cache.get(uk) || 0) + 1), 3600);
    return out;
  } catch (e) {
    return { error: '取得エラー: ' + String(e).slice(0, 80) };
  }
}

// ============================================================
// 🧪 動作確認用(エディタから手動で実行する。画面からは呼ばれない)
// ============================================================

// Apify のトークンが見えているかを確かめる(トークンそのものはログに出さない)
function checkApify() {
  const names = Object.keys(PropertiesService.getScriptProperties().getProperties());
  Logger.log('このプロジェクトのスクリプトプロパティ名: ' + JSON.stringify(names));
  const v = getApifyToken_();
  Logger.log(v ? 'Apifyトークンが見つかりました(長さ ' + v.length + ' 文字)' : 'Apifyトークンが見つかりません');
}

// 生きているRSSHub / ProxiTok を探す。✅ が出たものだけを TT_RSSHUB_HOSTS / TT_PROXITOK_HOSTS に入れる
function testTikTokHosts() {
  const uid = 'tiktok'; // 試すアカウント名
  const rsshubs = [
    'https://hub.slarker.me',
    'https://rsshub.ktachibana.party',
  ];
  rsshubs.forEach(function(h) {
    try {
      const r = UrlFetchApp.fetch(h + '/tiktok/user/@' + uid, {
        muteHttpExceptions: true,
        headers: { 'User-Agent': TT_UA_, 'Accept': 'application/rss+xml, application/xml, */*' }
      });
      const body = r.getContentText();
      const ok = r.getResponseCode() === 200 && /\/video\/\d{10,}/.test(body);
      Logger.log((ok ? '✅ 使える ' : '❌ だめ ') + h + ' (' + r.getResponseCode() + ')');
    } catch (e) {
      Logger.log('❌ だめ ' + h + ' 接続できません');
    }
  });
}

// ============================================================
// 💬 YouTubeのコメントを読む
//   順番(取れたところで止まる): 公式API(YT_API_KEY があるとき) → YouTubeの内部API(非公式) → 非公式の公開サーバー(Invidious・Piped)
//   つづき(もっと見る)は、同じ取得元から取る。コメントがオフの動画は、その旨を返す
// ============================================================
const YT_COMMENTS_PER_HOUR = 60;

function ytcWalk_(o, fn, depth) {
  depth = depth || 0;
  if (!o || typeof o !== 'object' || depth > 40) return;
  if (Array.isArray(o)) { for (let i = 0; i < o.length; i++) ytcWalk_(o[i], fn, depth + 1); return; }
  fn(o);
  const ks = Object.keys(o);
  for (let i = 0; i < ks.length; i++) ytcWalk_(o[ks[i]], fn, depth + 1);
}
function ytcText_(t) {
  if (!t) return '';
  if (typeof t === 'string') return t;
  if (t.simpleText) return String(t.simpleText);
  if (t.content) return String(t.content);
  if (Array.isArray(t.runs)) return t.runs.map(function(r) { return r.text || ''; }).join('');
  return '';
}
function ytcTokenOf_(o) { // continuationItemRenderer などの中から、続きのトークンを探す
  let tok = '';
  ytcWalk_(o, function(x) { if (!tok && x.continuationCommand && x.continuationCommand.token) tok = String(x.continuationCommand.token); });
  return tok;
}
function ytcAgo_(iso) {
  const t = Date.parse(iso); if (!isFinite(t)) return '';
  const s = Math.max(0, (Date.now() - t) / 1000);
  if (s < 3600) return Math.max(1, Math.floor(s / 60)) + '分前';
  if (s < 86400) return Math.floor(s / 3600) + '時間前';
  if (s < 86400 * 30) return Math.floor(s / 86400) + '日前';
  if (s < 86400 * 365) return Math.floor(s / 86400 / 30) + 'か月前';
  return Math.floor(s / 86400 / 365) + '年前';
}
function ytcOne_(author, text, likes, time, replies) {
  return { author: String(author || '').slice(0, 60), text: String(text || '').replace(/\r/g, '').slice(0, 1000), likes: String(likes == null ? '' : likes).slice(0, 12), time: String(time || '').slice(0, 20), replies: Number(replies) || 0 };
}

function ytcApi_(videoId, sort, tok) {
  const key = PropertiesService.getScriptProperties().getProperty('YT_API_KEY');
  if (!key) return { skipped: true };
  try {
    const r = UrlFetchApp.fetch('https://www.googleapis.com/youtube/v3/commentThreads?part=snippet&maxResults=20&textFormat=plainText&order=' + (sort === 'new' ? 'time' : 'relevance') +
      '&videoId=' + encodeURIComponent(videoId) + (tok ? '&pageToken=' + encodeURIComponent(tok) : '') + '&key=' + encodeURIComponent(key), { muteHttpExceptions: true });
    let j = null; try { j = JSON.parse(r.getContentText()); } catch (e) {}
    if (r.getResponseCode() !== 200 || !j) {
      const reason = j && j.error && j.error.errors && j.error.errors[0] ? j.error.errors[0].reason : '';
      if (reason === 'commentsDisabled') return { disabled: true, source: 'YouTube Data API' };
      return { note: 'コメントAPI: ' + r.getResponseCode() + (reason ? ' ' + reason : '') };
    }
    const comments = (j.items || []).map(function(it) {
      const s = (it.snippet && it.snippet.topLevelComment && it.snippet.topLevelComment.snippet) || {};
      return ytcOne_(s.authorDisplayName, s.textDisplay || s.textOriginal, s.likeCount, ytcAgo_(s.publishedAt), it.snippet && it.snippet.totalReplyCount);
    });
    return { comments: comments, next: j.nextPageToken ? 'api:' + j.nextPageToken : '', total: '', source: 'YouTube Data API' };
  } catch (e) { return { note: 'コメントAPI: 接続できません' }; }
}

function ytcInnerPost_(body) {
  const r = UrlFetchApp.fetch('https://www.youtube.com/youtubei/v1/next?prettyPrint=false', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify(Object.assign({ context: { client: { clientName: 'WEB', clientVersion: '2.20250620.01.00', hl: 'ja', gl: 'JP' } } }, body)),
    headers: { 'User-Agent': YT_UA_, 'Origin': 'https://www.youtube.com' }
  });
  if (r.getResponseCode() !== 200) return { status: r.getResponseCode() };
  try { return { json: JSON.parse(r.getContentText()) }; } catch (e) { return { status: 0 }; }
}
function ytcSectionToken_(j) { // 動画のページから、コメント欄の入口のトークンを取る
  let tok = '', seen = false;
  ytcWalk_(j, function(o) {
    const s = o.itemSectionRenderer;
    if (!tok && s && s.sectionIdentifier === 'comment-item-section') { seen = true; tok = ytcTokenOf_(s.contents); }
  });
  return { token: tok, sectionSeen: seen };
}
function ytcParseInner_(j) {
  const ents = {};
  const muts = (((j.frameworkUpdates || {}).entityBatchUpdate || {}).mutations) || [];
  muts.forEach(function(m) { const p = m.payload && m.payload.commentEntityPayload; if (p && p.properties && p.properties.commentId) ents[p.properties.commentId] = p; });
  const threads = [], out = { comments: [], next: '', sortTokens: [], disabled: false };
  (j.onResponseReceivedEndpoints || []).forEach(function(e) {
    const cmd = e.reloadContinuationItemsCommand || e.appendContinuationItemsAction; if (!cmd) return;
    (cmd.continuationItems || []).forEach(function(it) {
      if (it.commentThreadRenderer) threads.push(it.commentThreadRenderer);
      else if (it.continuationItemRenderer) { const t = ytcTokenOf_(it.continuationItemRenderer); if (t) out.next = t; }
      else if (it.commentsHeaderRenderer) {
        const sub = (((it.commentsHeaderRenderer.sortMenu || {}).sortFilterSubMenuRenderer || {}).subMenuItems) || [];
        out.sortTokens = sub.map(function(x) { return ytcTokenOf_(x); });
        out.total = ytcText_(it.commentsHeaderRenderer.countText).replace(/\s*(件)?のコメント.*$/, '').replace(/\s*Comments?$/i, '');
      } else if (it.messageRenderer) out.disabled = true;
    });
  });
  threads.forEach(function(ct) {
    const vm = ct.commentViewModel && (ct.commentViewModel.commentViewModel || ct.commentViewModel);
    const cid = vm && vm.commentId;
    const rep = ct.replies && ct.replies.commentRepliesRenderer ? (Number(String(ytcText_(ct.replies.commentRepliesRenderer.moreText)).replace(/[^0-9]/g, '')) || 0) : 0;
    if (cid && ents[cid]) {
      const p = ents[cid], tb = p.toolbar || {};
      out.comments.push(ytcOne_((p.author && p.author.displayName) || '', ytcText_(p.properties.content), tb.likeCountNotliked || tb.likeCountA11y || '', p.properties.publishedTime, tb.replyCount || rep));
    } else if (ct.comment && ct.comment.commentRenderer) {
      const c = ct.comment.commentRenderer;
      out.comments.push(ytcOne_(ytcText_(c.authorText), ytcText_(c.contentText), ytcText_(c.voteCount), ytcText_(c.publishedTimeText), c.replyCount || rep));
    }
  });
  return out;
}
function ytcInner_(videoId, sort, tok) {
  try {
    let cont = tok;
    if (!cont) {
      const a = ytcInnerPost_({ videoId: videoId });
      if (!a.json) return { note: '内部API(コメント): ' + a.status };
      const sec = ytcSectionToken_(a.json);
      if (!sec.token) return sec.sectionSeen ? { disabled: true, source: '内部API' } : { note: '内部API(コメント): 入口が見つかりません' };
      cont = sec.token;
    }
    let b = ytcInnerPost_({ continuation: cont });
    if (!b.json) return { note: '内部API(コメント): ' + b.status };
    let p = ytcParseInner_(b.json);
    if (!tok && sort === 'new' && p.sortTokens[1]) { // 先頭ページだけ、「新しい順」に切り替える
      const c2 = ytcInnerPost_({ continuation: p.sortTokens[1] });
      if (c2.json) { const p2 = ytcParseInner_(c2.json); if (p2.comments.length) { p2.total = p.total; p = p2; } }
    }
    if (!p.comments.length) return p.disabled ? { disabled: true, source: '内部API' } : { note: '内部API(コメント): 結果なし' };
    return { comments: p.comments, next: p.next ? 'it:' + p.next : '', total: p.total || '', source: '内部API' };
  } catch (e) { return { note: '内部API(コメント): 接続できません' }; }
}

function ytcMirrors_(videoId, sort, tok) {
  const hostsI = YT_INVIDIOUS_HOSTS.slice(0, 3), hostsP = YT_PIPED_HOSTS.slice(0, 3);
  const opt = { muteHttpExceptions: true, headers: { 'User-Agent': YT_UA_, 'Accept': 'application/json' } };
  const parseI = function(j, host) {
    if (!j || !Array.isArray(j.comments)) return null;
    return { comments: j.comments.map(function(c) { return ytcOne_(c.author, c.content, c.likeCount, c.publishedText, c.replies && c.replies.replyCount); }), next: j.continuation ? 'inv:' + host + '|' + j.continuation : '', total: j.commentCount ? String(j.commentCount) : '' };
  };
  const parseP = function(j, host) {
    if (!j || !Array.isArray(j.comments)) return null;
    return { comments: j.comments.map(function(c) { return ytcOne_(c.author, c.commentText, c.likeCount, c.commentedTime, c.replyCount); }), next: j.nextpage ? 'pi:' + host + '|' + j.nextpage : '', total: j.commentCount ? String(j.commentCount) : '' };
  };
  const urlI = function(h, t) { return h + '/api/v1/comments/' + encodeURIComponent(videoId) + '?sort_by=' + (sort === 'new' ? 'new' : 'top') + (t ? '&continuation=' + encodeURIComponent(t) : ''); };
  const urlP = function(h, t) { return t ? h + '/nextpage/comments/' + encodeURIComponent(videoId) + '?nextpage=' + encodeURIComponent(t) : h + '/comments/' + encodeURIComponent(videoId); };
  const get = function(url) { try { return UrlFetchApp.fetch(url, opt); } catch (e) { return null; } };
  const J = function(r) { if (!r || r.getResponseCode() !== 200) return null; try { return JSON.parse(r.getContentText()); } catch (e) { return null; } };
  if (tok) { // つづき: 最初に取れた公開サーバーに、そのままお願いする
    const m = String(tok).match(/^(inv|pi):([^|]+)\|(.+)$/);
    if (!m) return { note: '公開サーバー(コメント): つづきの形が違います' };
    const isI = m[1] === 'inv', r = get(isI ? urlI(m[2], m[3]) : urlP(m[2], m[3]));
    const p = (isI ? parseI : parseP)(J(r), m[2]);
    return p && p.comments.length ? Object.assign(p, { source: (isI ? 'Invidious(' : 'Piped(') + ttHostName_(m[2]) + ')' }) : { note: '公開サーバー(コメント): つづきを取れませんでした' };
  }
  const reqs = [], meta = [];
  hostsI.forEach(function(h) { reqs.push(Object.assign({ url: urlI(h) }, opt)); meta.push({ h: h, i: true }); });
  hostsP.forEach(function(h) { reqs.push(Object.assign({ url: urlP(h) }, opt)); meta.push({ h: h, i: false }); });
  let res;
  try { res = UrlFetchApp.fetchAll(reqs); } catch (e) { res = reqs.map(function(x) { return get(x.url); }); }
  const notes = [];
  for (let k = 0; k < reqs.length; k++) {
    const p = (meta[k].i ? parseI : parseP)(J(res[k]), meta[k].h);
    const name = (meta[k].i ? 'Invidious(' : 'Piped(') + ttHostName_(meta[k].h) + ')';
    if (p && p.comments.length) return Object.assign(p, { source: name });
    notes.push(name + ': ' + (res[k] ? res[k].getResponseCode() : '接続できません'));
  }
  return { notes: notes };
}

// 画面から呼ばれる。動画のコメントを20件ずつ返す(next が空なら先頭から)。sort は 'top'(人気順)か 'new'(新しい順)
function getYouTubeComments(id, pass, videoId, sort, next) {
  try {
    if (!checkUserOrGuest_(id, pass)) return { error: '認証エラー' };
    videoId = String(videoId || '').trim();
    if (!/^[A-Za-z0-9_-]{11}$/.test(videoId)) return { error: '動画のIDが正しくありません' };
    sort = sort === 'new' ? 'new' : 'top';
    next = String(next || '').slice(0, 3000);

    const cache = CacheService.getScriptCache();
    const ck = 'ytcm_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, videoId + '|' + sort + '|' + next, Utilities.Charset.UTF_8)).slice(0, 22);
    const hit = cache.get(ck);
    if (hit) { try { return JSON.parse(hit); } catch (e) {} }

    const uk = 'ytcmc_' + String(id).slice(0, 60);
    if (Number(cache.get(uk) || 0) >= YT_COMMENTS_PER_HOUR) return { error: 'コメントを見すぎです。しばらくしてからにしてください(1時間に' + YT_COMMENTS_PER_HOUR + '回まで)' };

    const notes = [];
    let found = null, disabled = false;
    const take = function(r) {
      if (found || !r) return;
      if (r.comments && r.comments.length) found = r;
      else if (r.disabled) disabled = true;
      else { if (r.note) notes.push(r.note); (r.notes || []).forEach(function(n) { notes.push(n); }); }
    };
    const colon = next.indexOf(':');
    const kind = colon > 0 ? next.slice(0, colon) : '', tok = colon > 0 ? next.slice(colon + 1) : '';
    if (next && kind === 'api') take(ytcApi_(videoId, sort, tok));
    else if (next && kind === 'it') take(ytcInner_(videoId, sort, tok));
    else if (next && (kind === 'inv' || kind === 'pi')) take(ytcMirrors_(videoId, sort, next));
    else if (next) return { error: 'つづきの形が正しくありません' };
    else {
      take(ytcApi_(videoId, sort, ''));
      if (!found && !disabled) take(ytcInner_(videoId, sort, ''));
      if (!found && !disabled) take(ytcMirrors_(videoId, sort, ''));
    }
    if (!found) {
      if (disabled) return { comments: [], next: '', disabled: true };
      return { error: 'コメントを取れませんでした(' + notes.join(' / ') + ')。時間をおいて試してください' };
    }
    const out = { comments: found.comments, next: found.next || '', total: found.total || '', source: found.source };
    cache.put(ck, JSON.stringify(out), 600);
    cache.put(uk, String(Number(cache.get(uk) || 0) + 1), 3600);
    return out;
  } catch (e) {
    return { error: 'コメントのエラー: ' + String(e).slice(0, 80) };
  }
}


// ============================================================
// ⚡ YouTubeの動画をDriveに保存する(非公式・「超貫通モード」用)
//   YouTubeの内部API(アプリ用の窓口)で、音つきのmp4の直リンクをもらい、Apps Script から取って Drive に保存する。
//   サーバーはいらない。ただし、YouTube側の変更・ボット対策で、取れなくなることがある(取れない理由は、画面に出す)。
//   ・音と映像が一体になった形式(360p・720p)だけが対象。上限は約100MB。長い動画・高画質のみの動画は保存できない
//   ・保存したものは、他の動画と同じく、cleanupTikTokVideos で古いものから消える
//   ・1人が1時間に保存できる本数: YTD_PER_HOUR
// ============================================================
const YTD_PER_HOUR = 6;
const YTD_MAX_BYTES = 100 * 1024 * 1024;
const YTD_CHUNK = 16 * 1024 * 1024; // 256KBの倍数(Driveの分割アップロードの決まり)
const YTD_DIRECT_MAX = 40 * 1024 * 1024; // これ以下は、1回で保存する

// 取得の窓口(アプリのふりをする)。取れなかったら、次の窓口を試す
const YTD_CLIENTS_ = [
  { label: 'ANDROID_VR', num: 28, ver: '1.60.19', ua: 'com.google.android.apps.youtube.vr.oculus/1.60.19 (Linux; U; Android 12L; eureka-user Build/SQ3A.220605.009.A1) gzip',
    client: { clientName: 'ANDROID_VR', clientVersion: '1.60.19', deviceMake: 'Oculus', deviceModel: 'Quest 3', androidSdkVersion: 32, osName: 'Android', osVersion: '12L' } },
  { label: 'IOS', num: 5, ver: '19.45.4', ua: 'com.google.ios.youtube/19.45.4 (iPhone16,2; U; CPU iOS 18_1_0 like Mac OS X;)',
    client: { clientName: 'IOS', clientVersion: '19.45.4', deviceMake: 'Apple', deviceModel: 'iPhone16,2', osName: 'iPhone', osVersion: '18.1.0.22B83' } },
  { label: 'TV_EMBEDDED', num: 85, ver: '2.0', ua: 'Mozilla/5.0 (PlayStation; PlayStation 4/12.00) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/15.4 Safari/605.1.15',
    client: { clientName: 'TVHTML5_SIMPLY_EMBEDDED_PLAYER', clientVersion: '2.0' }, thirdParty: { embedUrl: 'https://www.youtube.com/' } },
  { label: 'ANDROID', num: 3, ver: '19.44.38', ua: 'com.google.android.youtube/19.44.38 (Linux; U; Android 11) gzip',
    client: { clientName: 'ANDROID', clientVersion: '19.44.38', androidSdkVersion: 30, osName: 'Android', osVersion: '11' } }
];

// 窓口に「この動画の再生用の情報をください」と頼む。返事から、音つきmp4のいちばんよいものを選ぶ
function ytdPlayer_(cl, vid) {
  try {
    const ctx = { client: Object.assign({ hl: 'ja', gl: 'JP' }, cl.client) };
    if (cl.thirdParty) ctx.thirdParty = cl.thirdParty;
    const r = UrlFetchApp.fetch('https://www.youtube.com/youtubei/v1/player?prettyPrint=false', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ context: ctx, videoId: vid, contentCheckOk: true, racyCheckOk: true }),
      headers: { 'User-Agent': cl.ua, 'X-YouTube-Client-Name': String(cl.num), 'X-YouTube-Client-Version': cl.ver, 'Origin': 'https://www.youtube.com' }
    });
    if (r.getResponseCode() !== 200) return { note: cl.label + ': ' + r.getResponseCode() };
    const j = JSON.parse(r.getContentText());
    const ps = j.playabilityStatus || {};
    if (ps.status && ps.status !== 'OK') return { note: cl.label + ': ' + ps.status + (ps.reason ? '(' + String(ps.reason).slice(0, 40) + ')' : '') };
    const sd = j.streamingData || {};
    const dur = Number((j.videoDetails && j.videoDetails.lengthSeconds) || 0);
    const cands = (sd.formats || []).filter(function(f) {
      return f && f.url && /^video\/mp4/.test(f.mimeType || '') && /mp4a|opus/.test(f.mimeType || ''); // 映像と音が一体のもの
    }).map(function(f) {
      const size = Number(f.contentLength) || (f.bitrate && dur ? Math.round(f.bitrate / 8 * dur) : 0);
      return { url: f.url, height: Number(f.height) || 0, size: size, itag: f.itag };
    });
    if (!cands.length) return { note: cl.label + ': 音つきの形式がありません(' + (sd.adaptiveFormats ? '映像と音が別々の形式だけ' : '再生情報なし') + ')' };
    // 高画質から順に、上限に入るものを選ぶ(720pまで)
    cands.sort(function(a, b) { return b.height - a.height; });
    let pick = null;
    for (let i = 0; i < cands.length; i++) { const c = cands[i]; if (c.height <= 720 && (!c.size || c.size <= YTD_MAX_BYTES)) { pick = c; break; } }
    if (!pick) return { note: cl.label + ': 動画が大きすぎます(約' + Math.round((cands[cands.length - 1].size || 0) / 1048576) + 'MB)', tooBig: true };
    pick.ua = cl.ua; pick.label = cl.label; pick.title = String((j.videoDetails && j.videoDetails.title) || '').slice(0, 100);
    return { pick: pick };
  } catch (e) { return { note: cl.label + ': ' + String(e).slice(0, 40) }; }
}

function ytdRange_(pick, start, end) {
  const r = UrlFetchApp.fetch(pick.url + (pick.url.indexOf('?') === -1 ? '?' : '&') + 'range=' + start + '-' + end, { headers: { 'User-Agent': pick.ua }, muteHttpExceptions: true });
  const code = r.getResponseCode();
  if (code !== 200 && code !== 206) throw new Error('取得 ' + code);
  const got = r.getBlob().getBytes().length;
  if (got !== end - start + 1) throw new Error('範囲指定が効きませんでした(' + got + 'バイト)');
  return r;
}
// 小さい動画は1回で、大きい動画は分割して、Driveに入れる。返すのは Driveのファイルのid
function ytdSave_(pick, name) {
  let total = pick.size;
  if (!total) { // 大きさが分からないときは、先頭だけ取って、返事から全体の大きさを知る
    const r0 = ytdRange_(pick, 0, 1023);
    const cr = String((r0.getAllHeaders()['Content-Range'] || r0.getAllHeaders()['content-range'] || '')).match(/\/(\d+)$/);
    total = cr ? Number(cr[1]) : 0;
    if (!total) throw new Error('大きさが分かりません');
  }
  if (total > YTD_MAX_BYTES) throw new Error('動画が大きすぎます(約' + Math.round(total / 1048576) + 'MB)');
  if (total <= YTD_DIRECT_MAX) {
    const blob = ytdRange_(pick, 0, total - 1).getBlob().setName(name).setContentType('video/mp4');
    if (blob.getBytes().length < Math.min(total, 1024)) throw new Error('受け取った動画が小さすぎます');
    return getTTFolder_().createFile(blob).getId();
  }
  const token = ScriptApp.getOAuthToken();
  const init = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&supportsAllDrives=true', {
    method: 'post', contentType: 'application/json; charset=UTF-8', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token, 'X-Upload-Content-Type': 'video/mp4', 'X-Upload-Content-Length': String(total) },
    payload: JSON.stringify({ name: name, mimeType: 'video/mp4', parents: [getTTFolder_().getId()] })
  });
  if (init.getResponseCode() !== 200) throw new Error('Driveへの保存の開始に失敗(' + init.getResponseCode() + ')');
  const hd = init.getAllHeaders();
  const loc = String(hd['Location'] || hd['location'] || '');
  if (!loc) throw new Error('Driveの保存先が分かりません');
  let fileId = '';
  for (let start = 0; start < total; start += YTD_CHUNK) {
    const end = Math.min(total, start + YTD_CHUNK) - 1;
    const part = ytdRange_(pick, start, end).getBlob();
    const up = UrlFetchApp.fetch(loc, { method: 'put', contentType: 'video/mp4', muteHttpExceptions: true, headers: { 'Content-Range': 'bytes ' + start + '-' + end + '/' + total }, payload: part });
    const c = up.getResponseCode();
    if (end + 1 < total) { if (c !== 308) throw new Error('Driveへの保存に失敗(' + c + ')'); }
    else {
      if (c !== 200 && c !== 201) throw new Error('Driveへの保存の仕上げに失敗(' + c + ')');
      try { fileId = JSON.parse(up.getContentText()).id || ''; } catch (e) {}
    }
  }
  if (!fileId) throw new Error('Driveのファイルidを受け取れませんでした');
  return fileId;
}

// 公開サーバー(Piped・Invidious)に動画を出してもらう。取得はサーバー側がやるので、Apps Script の送り元がボット扱いされても関係ない
//   戻り: { cands:[{url, height, size, label, ua}], notes:[…] }(高画質→低画質の順。720pまで・上限内のもの)
function ytdMirrorCands_(vid) {
  const opt = { muteHttpExceptions: true, headers: { 'User-Agent': YT_UA_, 'Accept': 'application/json' } };
  const reqs = [], meta = [];
  YT_PIPED_HOSTS.forEach(function(h) { reqs.push(Object.assign({ url: h + '/streams/' + vid }, opt)); meta.push({ h: h, piped: true }); });
  YT_INVIDIOUS_HOSTS.forEach(function(h) { reqs.push(Object.assign({ url: h + '/api/v1/videos/' + vid + '?local=true' }, opt)); meta.push({ h: h, piped: false }); });
  let res;
  try { res = UrlFetchApp.fetchAll(reqs); }
  catch (e) { res = reqs.map(function(x) { try { return UrlFetchApp.fetch(x.url, opt); } catch (e2) { return null; } }); }
  const cands = [], notes = [];
  const sizeOf = function(url, declared) {
    const n = Number(declared);
    if (n > 0) return n;
    const m = String(url).match(/[?&]clen=(\d+)/);
    return m ? Number(m[1]) : 0;
  };
  for (let i = 0; i < reqs.length; i++) {
    const m = meta[i], r = res[i], name = (m.piped ? 'Piped(' : 'Invidious(') + ttHostName_(m.h) + ')';
    if (!r) { notes.push(name + ': 接続できません'); continue; }
    if (r.getResponseCode() !== 200) { notes.push(name + ': ' + r.getResponseCode()); continue; }
    let j = null; try { j = JSON.parse(r.getContentText()); } catch (e) {}
    if (!j) { notes.push(name + ': 読めません'); continue; }
    if (j.error) { notes.push(name + ': ' + String(j.error).slice(0, 30)); continue; }
    let found = 0;
    const add = function(url, height, size) {
      if (!url) return;
      if (url.charAt(0) === '/') url = m.h + url;
      if (height > 720 || (size && size > YTD_MAX_BYTES)) return;
      cands.push({ url: url, height: height, size: size, label: name, ua: YT_UA_ }); found++;
    };
    if (m.piped) {
      (j.videoStreams || []).forEach(function(v) { // 音つき(videoOnly が false)のmp4だけ
        if (v && v.url && v.videoOnly === false && /mp4/i.test(String(v.mimeType || v.format || ''))) add(v.url, Number(v.height) || parseInt(v.quality, 10) || 0, sizeOf(v.url, v.contentLength));
      });
    } else {
      (j.formatStreams || []).forEach(function(v) { // Invidiousの formatStreams は、音つきの形式
        if (v && v.url && /mp4/i.test(String(v.container || v.type || ''))) add(v.url, parseInt(v.resolution || v.qualityLabel, 10) || 0, sizeOf(v.url, v.clen));
      });
    }
    if (!found) notes.push(name + ': 音つきの形式なし');
  }
  cands.sort(function(a, b) { return b.height - a.height; });
  return { cands: cands, notes: notes };
}
// ============================================================
// 🚀 RapidAPI のYouTubeダウンローダーで、直リンクをもらう(超貫通モード用)
//   返ってきたJSONの中から、音つきmp4のURLを探す(APIごとに形が違っても動くように、全体を探索する)
// ============================================================
const RAPID_DAILY_MAX = 20; // 1日に使える回数(全体)。無料枠を使い切らないための安全装置

function ytdRapidCands_(vid) {
  const props = PropertiesService.getScriptProperties();
  const key = String(props.getProperty('RAPIDAPI_KEY') || '').replace(/[\s\u3000"']/g, '');
  if (!key) return { cands: [], notes: ['RapidAPI: RAPIDAPI_KEY が未設定です'] };
  const host = String(props.getProperty('RAPIDAPI_HOST') || 'youtube-media-downloader.p.rapidapi.com').trim();
  const path = String(props.getProperty('RAPIDAPI_PATH') || '/v2/video/details?videoId={id}').trim().replace('{id}', encodeURIComponent(vid));

  // 1日の回数制限(実行前に数える)
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
  let d = {};
  try { d = JSON.parse(props.getProperty('RAPID_DAILY') || '{}') || {}; } catch (e) {}
  if (d.day !== today) d = { day: today, n: 0 };
  if (d.n >= RAPID_DAILY_MAX) return { cands: [], notes: ['RapidAPI: 今日の上限(' + RAPID_DAILY_MAX + '回)に達しました'] };
  d.n++;
  props.setProperty('RAPID_DAILY', JSON.stringify(d));

  let r;
  try {
    r = UrlFetchApp.fetch('https://' + host + path, {
      muteHttpExceptions: true,
      headers: { 'x-rapidapi-key': key, 'x-rapidapi-host': host }
    });
  } catch (e) { return { cands: [], notes: ['RapidAPI: 接続できません'] }; }
  const code = r.getResponseCode();
  if (code === 401 || code === 403) return { cands: [], notes: ['RapidAPI: キーが違うか、このAPIを購読していません(' + code + ')'] };
  if (code === 429) return { cands: [], notes: ['RapidAPI: 無料枠の上限です(429)'] };
  if (code !== 200) return { cands: [], notes: ['RapidAPI: ' + code] };
  let j = null;
  try { j = JSON.parse(r.getContentText()); } catch (e) {}
  if (!j) return { cands: [], notes: ['RapidAPI: 返事を読めません'] };

  const cands = [], seen = {};
  const walk = function(n, depth) {
    if (!n || typeof n !== 'object' || depth > 12) return;
    if (Array.isArray(n)) { n.forEach(function(x) { walk(x, depth + 1); }); return; }
    const url = typeof n.url === 'string' ? n.url : '';
    if (/^https:\/\//.test(url) && !seen[url]) {
      const mime = String(n.mimeType || n.mime_type || n.type || '');
      const ext = String(n.extension || n.ext || n.container || '');
      const isMp4 = /mp4/i.test(mime) || /^mp4$/i.test(ext) || /mime=video%2Fmp4/i.test(url);
      const isAudioOnly = /^audio\//i.test(mime) || /^(m4a|mp3|webm)$/i.test(ext) && !/video/i.test(mime);
      const noAudio = n.hasAudio === false || n.isVideoOnly === true || n.videoOnly === true ||
        (/codecs=/.test(mime) && !/mp4a|opus/.test(mime));
      if (isMp4 && !isAudioOnly && !noAudio) {
        seen[url] = 1;
        let h = Number(n.height) || parseInt(n.qualityLabel || n.quality || n.resolution, 10) || 0;
        let size = Number(n.contentLength || n.size || n.clen) || 0;
        cands.push({ url: url, height: h, size: size, label: 'RapidAPI', ua: YT_UA_ });
      }
    }
    Object.keys(n).forEach(function(k) { walk(n[k], depth + 1); });
  };
  walk(j, 0);

  // 720p以下で、いちばん高画質のものから順に。サイズ不明のものも残す
  const ok = cands.filter(function(c) { return (!c.height || c.height <= 720) && (!c.size || c.size <= YTD_MAX_BYTES); });
  ok.sort(function(a, b) { return b.height - a.height; });
  if (!ok.length) return { cands: [], notes: ['RapidAPI: 音つきmp4が見つかりません(RAPIDAPI_PATH や使うAPIを見直してください)'] };
  return { cands: ok, notes: [] };
}

// Range指定が効かないURL(中継サーバー経由など)用: そのまま丸ごと取って保存する(UrlFetchAppの上限があるので約50MBまで)
function ytdSaveWhole_(c, name) {
  const r = UrlFetchApp.fetch(c.url, { muteHttpExceptions: true, headers: { 'User-Agent': c.ua || YT_UA_ } });
  if (r.getResponseCode() !== 200) throw new Error('取得 ' + r.getResponseCode());
  const blob = r.getBlob().setName(name).setContentType('video/mp4');
  const len = blob.getBytes().length;
  if (len < 10 * 1024) throw new Error('受け取った動画が小さすぎます');
  if (len > YTD_MAX_BYTES) throw new Error('動画が大きすぎます');
  return getTTFolder_().createFile(blob).getId();
}
function getYouTubeVideo(id, pass, videoId) {
  try {
    if (!checkUser_(id, pass)) return { error: '認証エラー' };
    const vid = String(videoId || '').trim();
    if (!/^[A-Za-z0-9_-]{11}$/.test(vid)) return { error: 'YouTubeの動画IDが正しくありません' };

    const props = PropertiesService.getScriptProperties();
    const pk = 'ytv_' + vid;
    const saved = props.getProperty(pk);
    if (saved) {
      try { DriveApp.getFileById(saved); return { fid: saved, cached: true }; }
      catch (e) { props.deleteProperty(pk); }
    }

    const cache = CacheService.getScriptCache();
    const uk = 'ytdl_' + String(id).slice(0, 60);
    if (Number(cache.get(uk) || 0) >= YTD_PER_HOUR) return { error: '保存のしすぎです。しばらくしてからにしてください(1時間に' + YTD_PER_HOUR + '本まで)' };

    const notes = [];
    let fid = '';

    // ① RapidAPI(キーがあれば最初に使う)
    const rp = ytdRapidCands_(vid);
    for (let i = 0; i < rp.cands.length && i < 3 && !fid; i++) {
      try { fid = ytdSave_(rp.cands[i], vid + '.mp4'); }          // まず、分割して保存する方法
      catch (e1) {
        try { fid = ytdSaveWhole_(rp.cands[i], vid + '.mp4'); }   // だめなら、丸ごと取る方法
        catch (e2) { notes.push('RapidAPI: ' + String(e2.message || e2).slice(0, 50)); }
      }
    }
    if (!fid) rp.notes.forEach(function(n) { notes.push(n); });

    // ② 従来の方法(アプリ用の窓口 → 公開サーバー)
    for (let i = 0; i < YTD_CLIENTS_.length && !fid; i++) {
      const p = ytdPlayer_(YTD_CLIENTS_[i], vid);
      if (!p.pick) { notes.push(p.note); if (p.tooBig) break; continue; }
      try { fid = ytdSave_(p.pick, vid + '.mp4'); }
      catch (e) { notes.push(p.pick.label + ': ' + String(e.message || e).slice(0, 60)); }
    }
    if (!fid) {
      const mc = ytdMirrorCands_(vid);
      for (let i = 0; i < mc.cands.length && i < 3 && !fid; i++) {
        try { fid = ytdSave_(mc.cands[i], vid + '.mp4'); }
        catch (e) { notes.push(mc.cands[i].label + ': ' + String(e.message || e).slice(0, 50)); }
      }
      if (!fid) mc.notes.slice(0, 6).forEach(function(n) { notes.push(n); });
    }
    if (!fid) return { error: 'YouTubeから取れませんでした(' + notes.join(' / ') + ')' };
    const file = DriveApp.getFileById(fid);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
    catch (e1) {
      try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); }
      catch (e2) { Logger.log('共有設定スキップ: ' + e2); }
    }
    props.setProperty(pk, fid);
    cache.put(uk, String(Number(cache.get(uk) || 0) + 1), 3600);
    return { fid: fid };
  } catch (e) {
    return { error: '保存エラー: ' + String(e).slice(0, 100) };
  }
}

// 小さい動画は、動画データをそのまま返す(TikTokと同じ)。それ以外は fid だけ返して、画面が小分けで受け取る
function getYouTubeVideoFast(id, pass, videoId) {
  const r = getYouTubeVideo(id, pass, videoId);
  if (!r || r.error || r.cached) return r;
  try {
    const file = DriveApp.getFileById(r.fid);
    if (file.getSize() <= TT_DIRECT_MAX) return { fid: r.fid, mime: 'video/mp4', data: Utilities.base64Encode(file.getBlob().getBytes()) };
  } catch (e) { Logger.log('直接返せませんでした: ' + e); }
  return { fid: r.fid };
}

// エディタから実行して、YouTubeから取れるかを確かめる(保存はしません。結果は実行ログに出ます)
function testYtDownload() {
  const vid = 'jNQXAC9IVRw'; // YouTube最初の動画(短い)
  YTD_CLIENTS_.forEach(function(cl) {
    const p = ytdPlayer_(cl, vid);
    Logger.log(p.pick ? ('○ ' + cl.label + ' → 取れる(' + p.pick.height + 'p, 約' + Math.round((p.pick.size || 0) / 1024) + 'KB, itag ' + p.pick.itag + ')') : ('× ' + p.note));
    if (p.pick) { try { const r = ytdRange_(p.pick, 0, 1023); Logger.log('   先頭1KBの取得 → ' + r.getResponseCode()); } catch (e) { Logger.log('   先頭1KBの取得 → 失敗: ' + e.message); } }
  });
}


// エディタから実行して、公開サーバー経由で取れるかを確かめる(保存はしません。結果は実行ログに出ます)
//   「○」が1つでもあれば、超貫通モードは使えます。全部「×」なら、いまは公開サーバーも使えない状態です
function testYtMirrors() {
  const vid = 'jNQXAC9IVRw'; // YouTube最初の動画(短い)
  const mc = ytdMirrorCands_(vid);
  Logger.log('音つきの候補: ' + mc.cands.length + '件');
  const shown = {};
  mc.cands.slice(0, 6).forEach(function(c) {
    const key = c.label + c.height; if (shown[key]) return; shown[key] = 1;
    let t = '';
    try { const r = ytdRange_(c, 0, 1023); t = '先頭1KBの取得 → ' + r.getResponseCode() + ' ○'; } catch (e) { t = '先頭1KBの取得 → 失敗: ' + String(e.message || e).slice(0, 60) + ' ×'; }
    Logger.log((t.indexOf('○') !== -1 ? '○ ' : '× ') + c.label + ' ' + c.height + 'p 約' + Math.round((c.size || 0) / 1024) + 'KB  ' + t);
  });
  mc.notes.forEach(function(n) { Logger.log('× ' + n); });
}
function testRapid() {
  const props = PropertiesService.getScriptProperties();
  const key = String(props.getProperty('RAPIDAPI_KEY') || '').replace(/[\s\u3000"']/g, '');
  const host = String(props.getProperty('RAPIDAPI_HOST') || 'youtube-media-downloader.p.rapidapi.com').trim();
  Logger.log('キーの長さ: ' + key.length + ' / ホスト: ' + host);
  const r = UrlFetchApp.fetch('https://' + host + '/v2/video/details?videoId=jNQXAC9IVRw', {
    muteHttpExceptions: true,
    headers: { 'x-rapidapi-key': key, 'x-rapidapi-host': host }
  });
  Logger.log('code=' + r.getResponseCode());
  Logger.log(r.getContentText().slice(0, 500));
}

// ============================================================
// 📐 動画の向き(プレイヤー枠の縦横比)を決める  ── Code.gs に追加するブロック
//
//   優先順位
//     ① 動画ファイルの実際の幅・高さ
//          Render が測った値(result.width / result.height。無ければ result.quality.width / .height)
//          → ファイルIDごとに、スクリプトプロパティ yt_size_<ファイルID> に覚える
//          → Render から届かないとき(保存済みの動画など)は、Drive が持っている videoMediaMetadata
//     ② YouTube の URL が /shorts/ 形式 → 縦長
//     ③ それでも分からなければ 横長(16:9)
//
//   返す項目: orientation('landscape' | 'portrait' | 'square')
//             orientationSource('file' | 'drive' | 'url' | 'default')
//             width / height(分かったときだけ数値、分からなければ null)
//   しきい値(幅÷高さが 0.9〜1.1 なら正方形)は、main.py の orientation_of、
//   index.html の vfSizeToOrientation と同じ。変えるときは3か所そろえる。
//
//   ここの関数は、どれも例外を投げない(失敗しても 横長 16:9 を返す)。
//   なので、getYouTubeVideoViaRender の fid / cached / name の返し方には影響しない。
// ============================================================

function vidOrientationFromSize_(w, h) {
  w = Number(w); h = Number(h);
  if (!(w > 0) || !(h > 0)) return '';
  const r = w / h;
  if (r > 1.1) return 'landscape';
  if (r < 0.9) return 'portrait';
  return 'square';
}

// size: { w, h, src }(実測)または null。URL は /shorts/ 判定にだけ使う
function vidResolveOrientation_(size, url) {
  try {
    if (size) {
      const o = vidOrientationFromSize_(size.w, size.h);
      if (o) return { orientation: o, orientationSource: size.src || 'file', width: Number(size.w), height: Number(size.h) };
    }
    if (/youtube(?:-nocookie)?\.com\/shorts\//i.test(String(url || ''))) {
      return { orientation: 'portrait', orientationSource: 'url', width: null, height: null };
    }
  } catch (e) {}
  return { orientation: 'landscape', orientationSource: 'default', width: null, height: null };
}

// Drive が持っている動画の幅・高さ(アップロード直後は空のことがある)。取れなければ null
function vidDriveSize_(fid) {
  try {
    const r = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fid) + '?fields=videoMediaMetadata&supportsAllDrives=true',
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }
    );
    if (r.getResponseCode() !== 200) return null;
    const m = JSON.parse(r.getContentText()).videoMediaMetadata;
    return (m && Number(m.width) > 0 && Number(m.height) > 0) ? { w: Number(m.width), h: Number(m.height), src: 'drive' } : null;
  } catch (e) { return null; }
}

// 覚えている寸法 → 無ければ Drive の寸法(見つかれば覚える)
function vidSizeFor_(props, fid) {
  try {
    const key = 'yt_size_' + fid;
    const raw = props.getProperty(key);
    if (raw) {
      const m = JSON.parse(raw);
      if (m && m.w > 0 && m.h > 0) return { w: m.w, h: m.h, src: m.src || 'file' };
    }
    const d = vidDriveSize_(fid);
    if (d) { try { props.setProperty(key, JSON.stringify(d)); } catch (e) {} }
    return d;
  } catch (e) { return null; }
}

// ① 新しく保存した動画用: Render の返事(result)から向きを決める
//    getYouTubeVideoViaRender の「最後の return」で使う
function vidOrientationFromRender_(props, fileId, result, url) {
  try {
    const q = (result && result.quality && typeof result.quality === 'object') ? result.quality : {};
    const pair = (result && Number(result.width) > 0 && Number(result.height) > 0)
      ? [result.width, result.height]
      : [q.width, q.height];
    let size = null;
    if (Number(pair[0]) > 0 && Number(pair[1]) > 0) {
      size = { w: Number(pair[0]), h: Number(pair[1]), src: 'file' };
      try { props.setProperty('yt_size_' + fileId, JSON.stringify(size)); } catch (e) {}
    } else {
      size = vidSizeFor_(props, fileId); // Render から届かなければ、Drive が持つ寸法で補う
    }
    return vidResolveOrientation_(size, url);
  } catch (e) {
    return vidResolveOrientation_(null, url);
  }
}

// ② 保存済みの動画用: 覚えている寸法 → Drive の寸法 → URL → 16:9
//    getYouTubeVideoViaRender の「以前保存した動画を再利用する return」で使う
function vidOrientationForSaved_(props, fileId, url) {
  try {
    return vidResolveOrientation_(vidSizeFor_(props, fileId), url);
  } catch (e) {
    return vidResolveOrientation_(null, url);
  }
}

// 画面から呼ぶ(任意)。ファイルIDしか分からない場所(「見た動画」など)から向きを知りたいとき用
function getVideoOrientation(id, pass, fid, url) {
  try {
    if (!checkUser_(id, pass)) return { error: '認証エラー' };
    fid = String(fid || '');
    if (!/^[A-Za-z0-9_-]{20,}$/.test(fid)) return { error: 'IDが正しくありません' };
    return vidOrientationForSaved_(PropertiesService.getScriptProperties(), fid, url);
  } catch (e) {
    return { error: '向きの取得エラー: ' + String(e).slice(0, 80) };
  }
}

/**
 * 超貫通モード用：
 * Render経由でYouTube動画をGoogle Driveへ保存し、
 * 保存された動画のファイルIDをプレイヤーへ返す。
 *
 * スクリプトプロパティ:
 * RENDER_DOWNLOAD_URL
 * YOUTUBE_DRIVE_FOLDER_ID
 */
function getYouTubeVideoViaRender(id, pass, youtubeUrl) {
  try {
    // 1. 既存のユーザー認証
    if (!checkUser_(id, pass)) {
      return { error: '認証エラー' };
    }

    // 2. YouTube動画IDを取り出す
    const url = String(youtubeUrl || '').trim();
    const match =
      url.match(/(?:youtube(?:-nocookie)?\.com\/(?:watch\?v=|shorts\/|live\/|embed\/|v\/)|youtu\.be\/)([A-Za-z0-9_-]{11})/) ||
      url.match(/[?&]v=([A-Za-z0-9_-]{11})/);

    if (!match) {
      return { error: 'YouTubeの動画URLを読み取れませんでした' };
    }

    const videoId = match[1];
    const canonicalUrl =
      'https://www.youtube.com/watch?v=' + videoId;

    const props = PropertiesService.getScriptProperties();

    // 3. 以前保存した動画があれば再利用
    const cacheKey = 'yt_render_' + videoId;
    const savedId = props.getProperty(cacheKey);

    if (savedId) {
      try {
        DriveApp.getFileById(savedId);
        return Object.assign({ fid: savedId, cached: true }, vidOrientationForSaved_(props, savedId, url));
      } catch (e) {
        props.deleteProperty(cacheKey);
      }
    }

    // 4. Renderの設定を取得
    const endpoint = String(
      props.getProperty('RENDER_DOWNLOAD_URL') || ''
    ).trim();

    const folderId = String(
      props.getProperty('YOUTUBE_DRIVE_FOLDER_ID') || ''
    ).trim();

    if (!/^https:\/\/[^/]+\/download\/?$/.test(endpoint)) {
      return {
        error: 'RENDER_DOWNLOAD_URLが未設定か、URLが正しくありません'
      };
    }

    if (!folderId) {
      return {
        error: 'YOUTUBE_DRIVE_FOLDER_IDが未設定です'
      };
    }

    // 5. Renderへ動画の保存を依頼
    const response = UrlFetchApp.fetch(endpoint, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        Accept: 'application/json'
      },
      payload: JSON.stringify({
        url: canonicalUrl,
        folder_id: folderId,
        access_token: ScriptApp.getOAuthToken()
      }),
      muteHttpExceptions: true
    });

    const status = response.getResponseCode();
    const body = response.getContentText();

   let result = {};

try {
  result = JSON.parse(body);
} catch (e) {
  const headers = response.getHeaders();
  const contentType =
    headers['Content-Type'] ||
    headers['content-type'] ||
    '不明';

  const preview = String(body || '')
    .replace(/\s+/g, ' ')
    .slice(0, 500);

  return {
    error:
      'Render応答がJSONではありません（HTTP ' +
      status +
      '）。Content-Type=' +
      contentType +
      '。応答内容=' +
      (preview || '(空の応答)')
  };
}

    if (status < 200 || status >= 300 || result.error) {
      return {
        error: 'Render保存エラー（HTTP ' + status + '）: ' +
          String(result.error || result.message || '').slice(0, 180)
      };
    }

    // 6. Renderの応答からDriveのファイルIDを取得
    const data = result.data || {};
    const file = result.file || data.file || {};

    const fileId = String(
      result.file_id ||
      result.fileId ||
      result.fid ||
      result.id ||
      data.file_id ||
      data.fileId ||
      data.id ||
      file.id ||
      ''
    ).trim();

    if (!fileId) {
      return {
        error: '保存応答にDriveのファイルIDがありません。Renderの応答JSONを確認してください'
      };
    }

    // 7. 保存されたファイルがDriveから読めることを確認
    const driveFile = DriveApp.getFileById(fileId);

    // リンクから再生できるように共有を試す
    try {
      driveFile.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW
      );
    } catch (e) {
      // 組織の共有制限などで変更できない場合は既存設定のまま続行
      Logger.log('動画の共有設定を変更できませんでした: ' + e);
    }

    // 8. 次回は同じ動画を再ダウンロードしない
    props.setProperty(cacheKey, fileId);

    // 9. 動画の向きを決める(実際の幅・高さ → /shorts/ のURL → 16:9)
    return Object.assign({
      fid: fileId,
      cached: false,
      name: driveFile.getName()
    }, vidOrientationFromRender_(props, fileId, result, url));

  } catch (e) {
    return {
      error: 'Render経由の動画保存に失敗しました: ' +
        String(e).slice(0, 200)
    };
  }
}
// ============================================================
// ▼▼▼ ここから下を、Code.gs の一番最後に貼り付ける ▼▼▼
// 💌 DM(1対1のダイレクトメッセージ)と、🔗 URLまとめ
//   ・DMs シート(自動作成): 時刻 / 送信者 / 宛先 / 本文 / 既読時刻
//   ・DMは、当事者の2人にだけ返す(チャットの Messages シートとは完全に別)
//   ・URLまとめは、Messages シートの全履歴から URL を拾う(画面に届いていない古い分も含む)
// ============================================================
const DM_SHEET = 'DMs';
const DM_MAX_LEN = 1000;     // 1通の最大文字数
const DM_READ_ROWS = 1500;   // 読み込む最新の行数(シートが大きくなっても軽くするため)
const DM_SHOW = 100;         // 1つの会話で画面に送る件数
const DM_PER_MIN = 20;       // 1人が1分間に送れる通数

// ★ DM専用スプレッドシートのID。空のままなら、今まで通り SHEET_ID の中の DMs シートを使う
//   スクリプトプロパティ DM_SHEET_ID があれば、そちらを優先する(コードを書き換えずにIDを入れられる)
const DM_SHEET_ID = '';

function dmSpreadsheetId_() {
  const p = String(PropertiesService.getScriptProperties().getProperty('DM_SHEET_ID') || '').replace(/[\s\u3000"']/g, '');
  return p || String(DM_SHEET_ID || '').trim() || SHEET_ID;
}

function getDmSheet_() {
  const ss = SpreadsheetApp.openById(dmSpreadsheetId_());
  let sh = ss.getSheetByName(DM_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DM_SHEET);
    sh.appendRow(['時刻', '送信者', '宛先', '本文', '既読時刻']);
    sh.getRange('D:D').setNumberFormat('@'); // 本文は常に「文字」として保存
  }
  return sh;
}

// エディタから【1回だけ】実行: 今のスプレッドシートの DMs を、新しいファイルへコピーする(古い方は消さない)
function migrateDmsToNewSheet() {
  if (dmSpreadsheetId_() === SHEET_ID) { Logger.log('先に DM_SHEET_ID を設定してください'); return; }
  const old = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DM_SHEET);
  if (!old || old.getLastRow() < 2) { Logger.log('移すDMはありません'); return; }
  const dst = getDmSheet_();
  if (dst.getLastRow() >= 2) { Logger.log('新しい方に、もうDMがあります。二重に移さないよう中止しました'); return; }
  const n = old.getLastRow() - 1;
  dst.getRange(2, 4, n, 1).setNumberFormat('@');
  dst.getRange(2, 1, n, 5).setValues(old.getRange(2, 1, n, 5).getValues());
  Logger.log('コピーしました: ' + n + '件。新しい方を開いて確認してから、deleteOldDmSheet を実行してください');
}

// 新しい方に同じ件数が入っているのを確かめてから、古い DMs シートを削除する
function deleteOldDmSheet() {
  if (dmSpreadsheetId_() === SHEET_ID) { Logger.log('DM_SHEET_ID が未設定です'); return; }
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const old = ss.getSheetByName(DM_SHEET);
  if (!old) { Logger.log('古い DMs シートは、もうありません'); return; }
  if (getDmSheet_().getLastRow() < old.getLastRow()) { Logger.log('件数が合わないので、削除しません'); return; }
  ss.deleteSheet(old);
  Logger.log('古い DMs シートを削除しました');
}

function allUserIds_() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Users');
  if (!sh) return [];
  const d = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < d.length; i++) {
    if (d[i][0] !== '') out.push(String(d[i][0]));
  }
  return out;
}

// DMを送る。成功したら、その相手との会話(getDMData と同じ形)を返す
function sendDM(id, pass, to, text) {
  try {
    if (!checkUser_(id, pass)) return { error: '認証エラー' };
    const me = String(id);
    to = String(to || '').trim();
    text = String(text || '').replace(/\r/g, '').trim().slice(0, DM_MAX_LEN);
    if (!to || to === me) return { error: '送り先が正しくありません' };
    if (!userExists_(to)) return { error: '「' + to + '」というユーザーはいません' };
    if (!text) return { error: '本文が空です' };

    const cache = CacheService.getScriptCache();
    const rk = 'dmr_' + me.slice(0, 60);
    const n = Number(cache.get(rk) || 0);
    if (n >= DM_PER_MIN) return { error: '送りすぎです。少し待ってからにしてください(1分に' + DM_PER_MIN + '通まで)' };
    cache.put(rk, String(n + 1), 60);

    const lock = LockService.getScriptLock();
    lock.waitLock(8000);
    try {
      const sh = getDmSheet_();
      const row = sh.getLastRow() + 1;
      sh.getRange(row, 4).setNumberFormat('@');
      sh.getRange(row, 1, 1, 4).setValues([[new Date(), me, to, text]]);
    } finally {
      lock.releaseLock();
    }
    pushForDm_(me, to, text); // 📲 プッシュ通知
    return getDMData(id, pass, to);
  } catch (e) {
    return { error: 'DM送信エラー: ' + String(e).slice(0, 100) };
  }
}

// DMの一覧(会話ごとの最新・未読数)と、peer を指定したときはその会話の中身を返す。
//   peer を指定すると、その相手から届いた未読は「既読」になる
function getDMData(id, pass, peer) {
  try {
    if (!checkUser_(id, pass)) return { error: '認証エラー' };
    const me = String(id);
    const users = allUserIds_().filter(function(u) { return u !== me; });
    peer = String(peer || '');
    if (users.indexOf(peer) === -1) peer = '';

    const sh = getDmSheet_();
    const last = sh.getLastRow();
    let rows = [], startRow = 2;
    if (last >= 2) {
      startRow = Math.max(2, last - DM_READ_ROWS + 1);
      rows = sh.getRange(startRow, 1, last - startRow + 1, 5).getValues();
    }

    const threads = {};
    const conv = [];
    const toMark = [];
    rows.forEach(function(r, i) {
      const from = String(r[1]), to = String(r[2]);
      if (from !== me && to !== me) return;
      const other = from === me ? to : from;
      const t = new Date(r[0]);
      const ms = isNaN(t.getTime()) ? 0 : t.getTime();
      const th = threads[other] || (threads[other] = { peer: other, ts: 0, text: '', mine: false, unread: 0 });
      if (ms >= th.ts) {
        th.ts = ms;
        th.text = String(r[3]).replace(/\s+/g, ' ').slice(0, 40);
        th.mine = (from === me);
      }
      const isUnread = (to === me && !r[4]);
      if (peer && other === peer) {
        conv.push({ from: from, text: String(r[3]), ts: ms, read: !!r[4] });
        if (isUnread) toMark.push(i); // 開いている会話の未読は、これから既読にする
      } else if (isUnread) {
        th.unread++;
      }
    });

    if (toMark.length) { // 既読時刻(E列)を、まとめて1回で書き込む
      const now = new Date();
      const col = rows.map(function(r) { return [r[4]]; });
      toMark.forEach(function(i) { col[i] = [now]; });
      sh.getRange(startRow, 5, rows.length, 1).setValues(col);
    }

    const list = Object.keys(threads).map(function(k) { return threads[k]; })
      .sort(function(a, b) { return b.ts - a.ts; });
    let unread = 0;
    list.forEach(function(t) { unread += t.unread; });
    return { users: users, threads: list, messages: conv.slice(-DM_SHOW), peer: peer, unread: unread };
  } catch (e) {
    return { error: 'DM取得エラー: ' + String(e).slice(0, 100) };
  }
}

// URLの取り出し(英数字などURLに使える文字だけ。日本語の文章が続いていても、そこで切れる)
function extractUrls_(text) {
  const m = String(text || '').match(/https?:\/\/[A-Za-z0-9\-._~:\/?#\[\]@!$&'()*+,;=%]+/g) || [];
  return m.map(function(u) {
    u = u.replace(/[.,;:!?']+$/, '');
    while (u.charAt(u.length - 1) === ')' && (u.match(/\(/g) || []).length < (u.match(/\)/g) || []).length) u = u.slice(0, -1);
    return u;
  }).filter(function(u) { return u.length > 8; });
}

// チャットの全履歴から URL を集める(新しい順・同じURLは1つにまとめて、何回貼られたかを数える)
//   room が空なら、すべてのルーム
function getUrlList(id, pass, room, limit) {
  try {
    if (!checkUser_(id, pass)) return { error: '認証エラー' };
    const n = Math.min(Math.max(Number(limit) || 300, 1), 500);
    const want = room ? normRoom_(room) : '';
    const sh = getOrCreateMessagesSheet_(SpreadsheetApp.openById(SHEET_ID));
    const last = sh.getLastRow();
    if (last < 2) return { urls: [], rooms: ROOMS };
    const values = sh.getRange(2, 1, last - 1, 8).getValues();

    const map = {};
    const out = [];
    for (let i = values.length - 1; i >= 0; i--) { // 新しい方から
      const r = values[i];
      if (r[0] === '' || r[1] === '') continue;
      const text = String(r[2]);
      if (text === '!update' || text === '!deleted') continue;
      const rm = normRoom_(r[7]);
      if (want && rm !== want) continue;
      const urls = extractUrls_(text);
      if (!urls.length) continue;
      const d = new Date(r[0]);
      const ts = isNaN(d.getTime()) ? 0 : d.getTime();
      urls.forEach(function(u) {
        if (map[u]) { map[u].count++; return; }
        if (out.length >= n) return;
        map[u] = { url: u, sender: String(r[1]), ts: ts, room: rm, count: 1 };
        out.push(map[u]);
      });
    }
    return { urls: out, rooms: ROOMS };
  } catch (e) {
    return { error: 'URL取得エラー: ' + String(e).slice(0, 100) };
  }
}
// ▲▲▲ ここまで ▲▲▲


// ============================================================
// 🔑 9. パスワードのハッシュ化と、パスワード変更
//   ・Users シートのB列に「h1:ソルト:ハッシュ」の形で保存する(元のパスワードは、シートに残らない)
//   ・まだ平文のままのパスワードも、これまで通りログインできる(古い形式として扱う)
//   ・平文を全員ぶんハッシュ化するのは、エディタから hashAllPasswords を1回だけ実行する
// ============================================================
const PASS_MIN_LEN = 6;
const PASS_MAX_LEN = 64;

function hashPass_(salt, pass) {
  return Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + ':' + String(pass), Utilities.Charset.UTF_8)
  );
}

function makePassHash_(pass) {
  const salt = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  return 'h1:' + salt + ':' + hashPass_(salt, pass);
}

// 保存されている値(stored)と、入力されたパスワード(pass)が合うか
function verifyPass_(stored, pass) {
  stored = String(stored);
  pass = String(pass);
  if (stored.indexOf('h1:') === 0) {
    const p = stored.split(':');
    return p.length === 3 && hashPass_(p[1], pass) === p[2];
  }
  return stored === pass; // まだハッシュ化していない古い形式
}

// ★ 旧 passOk_ の置き換え(動きは同じ。ハッシュ形式にも対応しただけ)
function passOk_(rowId, rowPass, id, pass) {
  if (String(rowId) !== String(id)) return false;
  if (NO_PASS_USERS.indexOf(String(id)) !== -1) return true;
  return verifyPass_(rowPass, pass);
}

// エディタから【1回だけ】手動で実行する。Users シートの平文のパスワードを、全員ぶんハッシュに置き換える
//   ・実行前に、Users シートをコピー(右クリック → コピーを作成)してバックアップしておくと安心
//   ・あとから新しいユーザーを平文で追加したときは、もう一度実行すればハッシュ化される
function hashAllPasswords() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Users');
  if (!sh) { Logger.log('「Users」シートがありません'); return; }
  const d = sh.getDataRange().getValues();
  let n = 0, skipped = 0;
  for (let i = 1; i < d.length; i++) {
    if (d[i][0] === '') continue;
    const p = String(d[i][1]);
    if (p === '' || p.indexOf('h1:') === 0 || NO_PASS_USERS.indexOf(String(d[i][0])) !== -1) { skipped++; continue; }
    sh.getRange(i + 1, 2).setNumberFormat('@').setValue(makePassHash_(p));
    n++;
  }
  Logger.log('ハッシュ化: ' + n + '人 / そのまま: ' + skipped + '人');
}

// 画面から呼ばれる。今のパスワードを確かめてから、新しいパスワードをハッシュで保存する
function changePassword(id, pass, newPass) {
  try {
    id = String(id || '');
    if (NO_PASS_USERS.indexOf(id) !== -1) return { error: 'このユーザーはパスワードを変更できません' };
    newPass = String(newPass == null ? '' : newPass);
    if (newPass.length < PASS_MIN_LEN || newPass.length > PASS_MAX_LEN) {
      return { error: '新しいパスワードは' + PASS_MIN_LEN + '〜' + PASS_MAX_LEN + '文字にしてください' };
    }
    if (newPass !== newPass.trim()) return { error: '新しいパスワードの前後に空白は入れられません' };
    if (newPass === id) return { error: 'IDと同じパスワードは使えません' };
    if (newPass === String(pass)) return { error: '今のパスワードと同じです' };

    const cache = CacheService.getScriptCache();
    const fk = 'pwf_' + id.slice(0, 60);
    if (Number(cache.get(fk) || 0) >= 5) return { error: '失敗が続いたため、10分ほど使えません' };

    const lock = LockService.getScriptLock();
    lock.waitLock(8000);
    try {
      const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Users');
      if (!sh) return { error: '「Users」シートがありません' };
      const d = sh.getDataRange().getValues();
      for (let i = 1; i < d.length; i++) {
        if (String(d[i][0]) !== id) continue;
        if (!verifyPass_(d[i][1], pass)) {
          cache.put(fk, String(Number(cache.get(fk) || 0) + 1), 600);
          return { error: '今のパスワードが違います' };
        }
        sh.getRange(i + 1, 2).setNumberFormat('@').setValue(makePassHash_(newPass));
        cache.remove(fk);
        return { ok: true };
      }
      return { error: 'ユーザーが見つかりません' };
    } finally {
      lock.releaseLock();
    }
  } catch (e) {
    return { error: 'パスワード変更エラー: ' + String(e).slice(0, 80) };
  }
}


// ============================================================
// ✏️ 2. 自分のメッセージの編集・削除
//   ・編集 … 送信から EDIT_WINDOW_MIN 分以内。I列に編集時刻を記録(画面に「編集済み」と出る)
//   ・削除 … いつでも本人が削除できる。本文を '!deleted' に置き換えて、画面には送らない
//            (行を消さないので、他の人の画面の行番号がずれない)。自分で送った画像も、Driveのごみ箱に入れる
// ============================================================
const EDIT_WINDOW_MIN = 15;
const EDIT_MAX_LEN = 2000;

function ownMessageRow_(sh, id, rowIndex) {
  const row = Number(rowIndex);
  if (!sh || !isFinite(row) || row < 2 || row > sh.getLastRow()) return { error: 'メッセージが見つかりません' };
  const vals = sh.getRange(row, 1, 1, 8).getValues()[0];
  if (String(vals[1]) !== String(id)) return { error: '自分のメッセージだけ操作できます' };
  return { row: row, vals: vals };
}

function editMessage(id, pass, tiKey, rowIndex, newText) {
  try {
    if (!checkUser_(id, pass)) return { error: 'IDまたはパスワードが違います' };
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName('Messages');
    const own = ownMessageRow_(sh, id, rowIndex);
    if (own.error) return own;

    const old = String(own.vals[2]);
    if (old === '!deleted' || old === '!update' || old.indexOf('!poll ') === 0) return { error: 'このメッセージは編集できません' };
    const sent = new Date(own.vals[0]);
    if (isNaN(sent.getTime()) || Date.now() - sent.getTime() > EDIT_WINDOW_MIN * 60000) {
      return { error: '送信から' + EDIT_WINDOW_MIN + '分を過ぎたので、編集できません' };
    }

    const text = String(newText == null ? '' : newText).replace(/\r/g, '').trim().slice(0, EDIT_MAX_LEN);
    if (!text && !own.vals[5]) return { error: '本文が空です' };
    if (/^!(update|deleted)$/.test(text) || text.indexOf('!poll ') === 0) return { error: 'この文章には変更できません' };

    if (sh.getMaxColumns() < 9) sh.insertColumnsAfter(sh.getMaxColumns(), 1);
    if (!sh.getRange(1, 9).getValue()) sh.getRange(1, 9).setValue('編集');
    const cell = sh.getRange(own.row, 3);
    cell.setNumberFormat('@');
    cell.setValue(text);
    sh.getRange(own.row, 9).setValue(new Date());
    return getAllData(id, tiKey);
  } catch (e) {
    return { error: '編集エラー: ' + String(e).slice(0, 100) };
  }
}

function deleteOwnMessage(id, pass, tiKey, rowIndex) {
  try {
    if (!checkUser_(id, pass)) return { error: 'IDまたはパスワードが違います' };
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName('Messages');
    const own = ownMessageRow_(sh, id, rowIndex);
    if (own.error) return own;
    if (String(own.vals[2]) === '!deleted') return getAllData(id, tiKey);

    // 送った画像は、チャット用のフォルダにあるものだけ、ごみ箱へ(失敗しても、削除は続ける)
    const fid = extractDriveId_(own.vals[5]);
    if (fid) {
      try {
        const f = DriveApp.getFileById(fid);
        const folderId = getOrCreateImageFolder().getId();
        const ps = f.getParents();
        let inFolder = false;
        while (ps.hasNext()) { if (ps.next().getId() === folderId) { inFolder = true; break; } }
        if (inFolder) f.setTrashed(true);
      } catch (e2) { Logger.log('画像を消せませんでした: ' + e2); }
    }

    const cell = sh.getRange(own.row, 3);
    cell.setNumberFormat('@');
    cell.setValue('!deleted');
    sh.getRange(own.row, 5).setValue('{}');   // リアクション
    sh.getRange(own.row, 6).setValue('');     // 画像
    sh.getRange(own.row, 7).setValue('');     // 引用
    return getAllData(id, tiKey);
  } catch (e) {
    return { error: '削除エラー: ' + String(e).slice(0, 100) };
  }
}


// ============================================================
// 🎧 7-a. いま聴いている曲(短く覚えておくだけ。シートには書かない)
//   ・画面が、曲を聴いている間 約1分ごとに reportNowPlaying を呼ぶ(止めたら null を送って消す)
//   ・2.5分で自動的に消える。結果は、5秒ごとの同期(loginAndSync)に乗って、全員の画面に届く
// ============================================================
function reportNowPlaying(id, pass, track) {
  try {
    if (!checkUser_(id, pass)) return { ok: false };
    const key = 'np_' + String(id).slice(0, 100);
    const cache = CacheService.getScriptCache();
    if (!track || !/^[A-Za-z0-9_-]{11}$/.test(String(track.id || ''))) { cache.remove(key); return { ok: true }; }
    cache.put(key, JSON.stringify({
      id: String(track.id),
      t: String(track.title || '').slice(0, 100),
      c: String(track.ch || '').slice(0, 60)
    }), 150);
    return { ok: true };
  } catch (e) {
    return { ok: false };
  }
}

// getAllDataOnce_ の返り値に付ける。失敗しても、チャットは止めない
function getNowPlayingSafe_(ids, exceptId) {
  try {
    const keys = [];
    const owner = {};
    ids.forEach(function(u) {
      const s = String(u);
      if (s === String(exceptId)) return;
      const k = 'np_' + s.slice(0, 100);
      keys.push(k);
      owner[k] = s;
    });
    if (!keys.length) return [];
    const found = CacheService.getScriptCache().getAll(keys);
    const out = [];
    Object.keys(found).forEach(function(k) {
      try {
        const o = JSON.parse(found[k]);
        if (o && o.id) out.push({ user: owner[k], id: o.id, title: o.t || '', ch: o.c || '' });
      } catch (e) {}
    });
    return out;
  } catch (e) {
    return [];
  }
}


// ============================================================
// 🎶 7-b. みんなのプレイリスト(SharedPlaylists シート。自動作成)
//   ・だれでも、プレイリストを作れる / 曲を追加できる
//   ・名前の変更・プレイリストの削除は、作った人だけ
//   ・曲の削除は、その曲を追加した人か、プレイリストを作った人
//   ・列: ID / 名前 / 作成者 / 曲(JSON) / 更新
// ============================================================
const SHP_SHEET = 'SharedPlaylists';
const SHP_MAX_LISTS = 30;
const SHP_MAX_TRACKS = 150;
const SHP_MAX_JSON = 45000; // セル1つの上限(5万文字)に収まるように

function getShpSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHP_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHP_SHEET);
    sh.appendRow(['ID', '名前', '作成者', '曲(JSON)', '更新']);
    sh.getRange('A:D').setNumberFormat('@');
  }
  return sh;
}

function shpRead_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const v = sh.getRange(2, 1, last - 1, 5).getValues();
  return v.map(function(r, i) {
    let tracks = [];
    try { tracks = JSON.parse(String(r[3] || '[]')) || []; } catch (e) {}
    if (!Array.isArray(tracks)) tracks = [];
    const d = new Date(r[4]);
    return { row: i + 2, id: String(r[0]), name: String(r[1]), owner: String(r[2]), tracks: tracks, updated: isNaN(d.getTime()) ? 0 : d.getTime() };
  }).filter(function(l) { return l.id; });
}

function shpPublic_(lists) {
  return lists.map(function(l) { return { id: l.id, name: l.name, owner: l.owner, tracks: l.tracks, updated: l.updated }; });
}

function shpFind_(lists, listId) {
  for (let i = 0; i < lists.length; i++) if (lists[i].id === String(listId)) return lists[i];
  return null;
}

// 書き込み系は、ここを通す(ログイン確認 → 排他制御 → 変更 → 最新の一覧を返す)。fn は、エラー文を返すか、成功なら null を返す
function shpMutate_(id, pass, fn) {
  try {
    if (!checkUser_(id, pass)) return { error: '認証エラー' };
    const lock = LockService.getScriptLock();
    lock.waitLock(8000);
    try {
      const sh = getShpSheet_();
      const err = fn(sh, shpRead_(sh), String(id));
      if (err) return { error: err };
      return { lists: shpPublic_(shpRead_(sh)) };
    } finally {
      lock.releaseLock();
    }
  } catch (e) {
    return { error: 'プレイリストのエラー: ' + String(e).slice(0, 80) };
  }
}

function shpGetAll(id, pass) {
  try {
    if (!checkUser_(id, pass)) return { error: '認証エラー' };
    return { lists: shpPublic_(shpRead_(getShpSheet_())) };
  } catch (e) {
    return { error: 'プレイリストのエラー: ' + String(e).slice(0, 80) };
  }
}

function shpCreate(id, pass, name) {
  return shpMutate_(id, pass, function(sh, lists, me) {
    const nm = String(name || '').replace(/\s+/g, ' ').trim().slice(0, 40);
    if (!nm) return '名前を入れてください';
    if (lists.length >= SHP_MAX_LISTS) return 'プレイリストが多すぎます(' + SHP_MAX_LISTS + 'こまで)';
    const row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, 4).setNumberFormat('@');
    sh.getRange(row, 1, 1, 5).setValues([[Utilities.getUuid().replace(/-/g, '').slice(0, 10), nm, me, '[]', new Date()]]);
    return null;
  });
}

function shpRename(id, pass, listId, name) {
  return shpMutate_(id, pass, function(sh, lists, me) {
    const l = shpFind_(lists, listId);
    if (!l) return 'プレイリストが見つかりません';
    if (l.owner !== me) return '名前を変えられるのは、作った人(' + l.owner + ')だけです';
    const nm = String(name || '').replace(/\s+/g, ' ').trim().slice(0, 40);
    if (!nm) return '名前を入れてください';
    sh.getRange(l.row, 2).setValue(nm);
    sh.getRange(l.row, 5).setValue(new Date());
    return null;
  });
}

function shpDelete(id, pass, listId) {
  return shpMutate_(id, pass, function(sh, lists, me) {
    const l = shpFind_(lists, listId);
    if (!l) return 'プレイリストが見つかりません';
    if (l.owner !== me) return '削除できるのは、作った人(' + l.owner + ')だけです';
    sh.deleteRow(l.row);
    return null;
  });
}

function shpAddTrack(id, pass, listId, track) {
  return shpMutate_(id, pass, function(sh, lists, me) {
    const l = shpFind_(lists, listId);
    if (!l) return 'プレイリストが見つかりません';
    track = track || {};
    const vid = String(track.id || '');
    if (!/^[A-Za-z0-9_-]{11}$/.test(vid)) return '曲の情報が正しくありません';
    if (l.tracks.some(function(t) { return t && t.id === vid; })) return 'この曲は、もう入っています';
    if (l.tracks.length >= SHP_MAX_TRACKS) return 'このプレイリストはいっぱいです(' + SHP_MAX_TRACKS + '曲まで)';
    l.tracks.push({
      id: vid,
      title: String(track.title || '').slice(0, 100),
      ch: String(track.ch || '').slice(0, 60),
      dur: String(track.dur || '').slice(0, 12),
      by: me.slice(0, 40)
    });
    const json = JSON.stringify(l.tracks);
    if (json.length > SHP_MAX_JSON) return 'このプレイリストはいっぱいです';
    sh.getRange(l.row, 4, 1, 2).setValues([[json, new Date()]]);
    return null;
  });
}

function shpRemoveTrack(id, pass, listId, videoId) {
  return shpMutate_(id, pass, function(sh, lists, me) {
    const l = shpFind_(lists, listId);
    if (!l) return 'プレイリストが見つかりません';
    const idx = l.tracks.findIndex(function(t) { return t && t.id === String(videoId); });
    if (idx === -1) return null; // もう無い
    if (l.tracks[idx].by !== me && l.owner !== me) return '消せるのは、追加した人か、プレイリストを作った人だけです';
    l.tracks.splice(idx, 1);
    sh.getRange(l.row, 4, 1, 2).setValues([[JSON.stringify(l.tracks), new Date()]]);
    return null;
  });
}
// ============================================================
// 🖥 画面共有のシグナリング(CacheServiceだけで動く。シートは使わない)
//   ・部屋: 共有している人の一覧   ・メッセージ箱: 接続情報の受け渡し(各 受信者ごとに箱が1つ)
// ============================================================
const SCR_TTL = 90;
const SCR_MSG_TTL = 120;
const SCR_MAX_BOX = 40;

function scrKey_(id) { return 'scr_' + String(id).slice(0, 60); }
// 受信箱は「共有する側用(host)」と「見る側用(view)」に分ける(同じ人が両方になっても、取り合わない)
function scrBoxKey_(id, role) { return (role === 'host' ? 'scbh_' : 'scbv_') + String(id).slice(0, 60); }
function scrFlagKey_(id) { return 'scx_' + String(id).slice(0, 60); }
function scrAuth_(id, pass) { return String(id) !== GUEST_ID && checkUser_(id, pass); }

// role: 'host'(共有する側の別ウィンドウ)/ 'view'(見る側の親画面)
function scrPoll(id, pass, hosting, role) {
  try {
    if (!scrAuth_(id, pass)) return { error: '認証エラー' };
    const cache = CacheService.getScriptCache();
    const me = String(id);
    const isHost = role === 'host';
    let stop = false;

    if (isHost) {
      if (cache.get(scrFlagKey_(me))) { cache.remove(scrFlagKey_(me)); stop = true; }  // 親画面の「共有をやめる」
      else if (hosting) cache.put(scrKey_(me), JSON.stringify({ host: me, t: Date.now() }), SCR_TTL);
    }

    const ids = allUserIds_();
    const keys = ids.map(scrKey_);
    const found = keys.length ? cache.getAll(keys) : {};
    const rooms = [];
    ids.forEach(function(u) {
      const v = found[scrKey_(u)];
      if (!v) return;
      try { const o = JSON.parse(v); if (o && o.host) rooms.push({ host: String(o.host), t: o.t || 0 }); } catch (e) {}
    });

    let msgs = [];
    const lock = LockService.getScriptLock();
    lock.waitLock(3000);
    try {
      const bk = scrBoxKey_(me, role);
      const raw = cache.get(bk);
      if (raw) { try { msgs = JSON.parse(raw) || []; } catch (e) {} cache.remove(bk); }
    } finally { lock.releaseLock(); }
    return { rooms: rooms, msgs: msgs, stop: stop };
  } catch (e) {
    return { error: '共有の確認エラー: ' + String(e).slice(0, 80) };
  }
}

function scrStart(id, pass) {
  try {
    if (!scrAuth_(id, pass)) return { error: '認証エラー' };
    const cache = CacheService.getScriptCache();
    cache.remove(scrFlagKey_(id));
    cache.put(scrKey_(id), JSON.stringify({ host: String(id), t: Date.now() }), SCR_TTL);
    return { ok: true };
  } catch (e) { return { error: String(e).slice(0, 80) }; }
}

// 共有をやめる。別ウィンドウに「やめて」と伝えるための印も残す(20秒)
function scrStop(id, pass) {
  try {
    if (!scrAuth_(id, pass)) return { error: '認証エラー' };
    const cache = CacheService.getScriptCache();
    cache.remove(scrKey_(id));
    cache.put(scrFlagKey_(id), '1', 20);
    return { ok: true };
  } catch (e) { return { error: String(e).slice(0, 80) }; }
}

function scrSend(id, pass, to, msg) {
  try {
    if (!scrAuth_(id, pass)) return { error: '認証エラー' };
    to = String(to || '');
    if (!to || !userExists_(to)) return { error: '相手がいません' };
    const s = JSON.stringify(msg || {});
    if (s.length > 20000) return { error: '長すぎます' };
    const role = (msg && msg.dir === 'toHost') ? 'host' : 'view';   // 共有する側宛て → host の箱 / 見る側宛て → view の箱

    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      const cache = CacheService.getScriptCache();
      const bk = scrBoxKey_(to, role);
      let box = [];
      const raw = cache.get(bk);
      if (raw) { try { box = JSON.parse(raw) || []; } catch (e) {} }
      box.push({ from: String(id), msg: msg, ts: Date.now() });
      if (box.length > SCR_MAX_BOX) box = box.slice(-SCR_MAX_BOX);
      const out = JSON.stringify(box);
      if (out.length > 90000) return { error: '混み合っています' };
      cache.put(bk, out, SCR_MSG_TTL);
    } finally {
      lock.releaseLock();
    }
    return { ok: true };
  } catch (e) { return { error: String(e).slice(0, 80) }; }
}
// ============================================================
// 🔑 TIキーの確認(キーそのものはコードに書かない。ハッシュだけをスクリプトプロパティに入れる)
// ============================================================
const TI_USERS = ['juki2023'];   // ← TIを使ってよいユーザーID(自分がログインに使っているID)に書き換える

function isTiOk_(id, tiKey) {
  try {
    if (!tiKey) return false;
    if (TI_USERS.indexOf(String(id)) === -1) return false;
    const stored = String(PropertiesService.getScriptProperties().getProperty('TI_KEY_HASH') || '');
    if (stored.indexOf('h1:') !== 0) return false;
    const p = stored.split(':');
    if (p.length !== 3) return false;
    return hashPass_(p[1], String(tiKey)) === p[2];
  } catch (e) {
    return false;
  }
}

// エディタから実行して、設定できているか確かめる用
function checkKeys() {
  const v = String(PropertiesService.getScriptProperties().getProperty('TI_KEY_HASH') || '');
  Logger.log(v.indexOf('h1:') === 0 && v.split(':').length === 3
    ? '設定済み' : '未設定、または形が違います(h1:…:… の形にしてください)');
}


// ============================================================
// 📲 プッシュ通知(アプリを閉じていても届く通知)
//   ・外枠サイト(PWA)が端末の「通知トークン」を作り、アプリ本体が registerPushToken でここへ登録する
//   ・メッセージ / DM が送られたとき、FCM(Firebase Cloud Messaging)で相手の端末に送る
//   ・送る先の判断(ルームのミュート・@だけ・静かな時間帯)は、端末の設定を一緒に登録して、サーバー側で行う
//   ・Firebase のサービスアカウントの鍵(JSON)を、スクリプトプロパティ「FCM_SA」に入れると有効になる(入れなければ何もしない)
//   ・PushTokens シート(自動作成): ID / 通知トークン / 設定(JSON) / 更新
// ============================================================
const PUSH_SHEET = 'PushTokens';
const PUSH_MAX_PER_USER = 5;        // 1人が登録できる端末の数(古いものから消える)
const PUSH_TZ = 'Asia/Tokyo';       // 静かな時間帯を判断する時間帯
const PUSH_TOKEN_RE = /^[A-Za-z0-9_:\-.]{50,4096}$/;

function getPushSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(PUSH_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PUSH_SHEET);
    sh.appendRow(['ID', '通知トークン', '設定(JSON)', '更新']);
    sh.getRange('A:C').setNumberFormat('@');
  }
  return sh;
}

// 端末から届いた通知設定を、安全な形に整える
function pushPrefsClean_(raw) {
  let p = {};
  try { p = typeof raw === 'string' ? JSON.parse(raw) : (raw || {}); } catch (e) { p = {}; }
  const out = { off: !!p.off, rooms: {}, quiet: { on: false, from: '23:00', to: '07:00' } };
  const rooms = (p.rooms && typeof p.rooms === 'object') ? p.rooms : {};
  ROOMS.forEach(function(r) {
    const m = rooms[r];
    if (m === 'mention' || m === 'mute') out.rooms[r] = m;
  });
  const q = (p.quiet && typeof p.quiet === 'object') ? p.quiet : {};
  const hm = function(s, d) { return /^\d{1,2}:\d{2}$/.test(String(s)) ? String(s) : d; };
  out.quiet = { on: !!q.on, from: hm(q.from, '23:00'), to: hm(q.to, '07:00') };
  return out;
}

// アプリ本体から呼ばれる。通知トークンと設定を保存する(同じトークンなら上書き。端末を共有していたら持ち主を入れ替える)
function registerPushToken(id, pass, token, prefs) {
  try {
    id = String(id || '');
    if (id === GUEST_ID || !checkUser_(id, pass)) return { error: '認証エラー' };
    token = String(token || '');
    if (!PUSH_TOKEN_RE.test(token)) return { error: '通知トークンが正しくありません' };
    const cleaned = JSON.stringify(pushPrefsClean_(prefs));

    const lock = LockService.getScriptLock();
    lock.waitLock(8000);
    try {
      const sh = getPushSheet_();
      const last = sh.getLastRow();
      const vals = last >= 2 ? sh.getRange(2, 1, last - 1, 4).getValues() : [];
      let found = -1;
      const mine = [];
      for (let i = 0; i < vals.length; i++) {
        if (String(vals[i][1]) === token) found = i;
        else if (String(vals[i][0]) === id) mine.push(i);
      }
      if (found !== -1) {
        const row = found + 2;
        sh.getRange(row, 1, 1, 4).setNumberFormat('@').setValues([[id, token, cleaned, new Date().toISOString()]]);
      } else {
        const row = last + 1;
        sh.getRange(row, 1, 1, 4).setNumberFormat('@').setValues([[id, token, cleaned, new Date().toISOString()]]);
        // 多すぎる分は、古い順に消す(下の行から消す)
        const over = mine.length + 1 - PUSH_MAX_PER_USER;
        if (over > 0) {
          mine.slice(0, over).sort(function(a, b) { return b - a; }).forEach(function(i) { sh.deleteRow(i + 2); });
        }
      }
    } finally {
      lock.releaseLock();
    }
    return { ok: true };
  } catch (e) {
    return { error: '通知の登録エラー: ' + String(e).slice(0, 100) };
  }
}

// ---- FCM への認証(サービスアカウント → アクセストークン。50分キャッシュ) ----
function pushServiceAccount_() {
  const raw = PropertiesService.getScriptProperties().getProperty('FCM_SA');
  if (!raw) return null;
  try {
    const sa = JSON.parse(raw);
    return (sa && sa.client_email && sa.private_key && sa.project_id) ? sa : null;
  } catch (e) { return null; }
}

function b64url_(x) {
  const s = typeof x === 'string' ? Utilities.base64EncodeWebSafe(x, Utilities.Charset.UTF_8) : Utilities.base64EncodeWebSafe(x);
  return s.replace(/=+$/, '');
}

function fcmAccessToken_(sa) {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('fcm_at');
  if (hit) return hit;
  const now = Math.floor(Date.now() / 1000);
  const head = b64url_(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const claim = b64url_(JSON.stringify({
    iss: sa.client_email,
    scope: 'https://www.googleapis.com/auth/firebase.messaging',
    aud: 'https://oauth2.googleapis.com/token',
    iat: now, exp: now + 3600
  }));
  const unsigned = head + '.' + claim;
  const sig = b64url_(Utilities.computeRsaSha256Signature(unsigned, sa.private_key));
  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: unsigned + '.' + sig },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    Logger.log('FCM認証に失敗: ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200));
    return '';
  }
  const tok = JSON.parse(res.getContentText()).access_token || '';
  if (tok) cache.put('fcm_at', tok, 3000);
  return tok;
}

// targets: [{ token, row }]。送れなかったトークン(もう使われていない端末)は、シートから消す
function pushSend_(targets, title, body, tag) {
  if (!targets.length) return;
  const sa = pushServiceAccount_();
  if (!sa) return;
  const at = fcmAccessToken_(sa);
  if (!at) return;
  const url = 'https://fcm.googleapis.com/v1/projects/' + encodeURIComponent(sa.project_id) + '/messages:send';
  const reqs = targets.map(function(t) {
    return {
      url: url, method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + at },
      muteHttpExceptions: true,
      payload: JSON.stringify({ message: {
        token: t.token,
        data: { title: String(title).slice(0, 80), body: String(body).slice(0, 120), tag: String(tag || 'chat') },
        webpush: { headers: { Urgency: 'high', TTL: '86400' } }
      } })
    };
  });
  const res = UrlFetchApp.fetchAll(reqs);
  const dead = [];
  res.forEach(function(r, i) {
    const code = r.getResponseCode();
    if (code === 200) return;
    const txt = r.getContentText();
    Logger.log('FCM送信失敗 ' + code + ': ' + txt.slice(0, 160));
    if (code === 404 || /UNREGISTERED|not a valid FCM registration token/i.test(txt)) dead.push(targets[i].token);
  });
  if (dead.length) pushDropTokens_(dead);
}

function pushDropTokens_(tokens) {
  const lock = LockService.getScriptLock();
  lock.waitLock(8000);
  try {
    const sh = getPushSheet_();
    const last = sh.getLastRow();
    if (last < 2) return;
    const vals = sh.getRange(2, 2, last - 1, 1).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      if (tokens.indexOf(String(vals[i][0])) !== -1) sh.deleteRow(i + 2);
    }
  } finally {
    lock.releaseLock();
  }
}

// 全員ぶんの登録(行)を読む
//   ※ 画面を開いている人にも送る。見ている(前面にある)ときは、外枠サイトのサービスワーカーが自動で表示を控える
function pushLoad_() {
  const sh = getPushSheet_();
  const last = sh.getLastRow();
  return { rows: last >= 2 ? sh.getRange(2, 1, last - 1, 3).getValues() : [] };
}

// 静かな時間帯か(hm は "HH:mm")
function pushInQuiet_(q, hm) {
  if (!q || !q.on) return false;
  const toMin = function(s) { const m = String(s).match(/^(\d{1,2}):(\d{2})$/); return m ? Number(m[1]) * 60 + Number(m[2]) : 0; };
  const cur = toMin(hm), f = toMin(q.from), t = toMin(q.to);
  if (f === t) return false;
  return f < t ? (cur >= f && cur < t) : (cur >= f || cur < t);
}

// この人に、このメッセージを通知してよいか。画面の通知設定(index.html の notifyMentionPlan / notifyAllowNormal)と同じ考え方
//   自分宛て → いつでも / @all → ミュートと静かな時間帯は止める / 普通のメッセージ → 「すべて」で、静かな時間帯でないときだけ
function pushShouldSend_(prefs, room, isMe, isAll, hm) {
  if (prefs.off) return false;
  if (isMe) return true;
  const mode = prefs.rooms[room] || 'all';
  const quiet = pushInQuiet_(prefs.quiet, hm);
  if (isAll) return mode !== 'mute' && !quiet;
  return mode === 'all' && !quiet;
}

// チャットにメッセージが入ったとき(sendMessageCore_ から呼ぶ。失敗してもチャットは止めない)
function pushForChat_(from, text, room, hasImage) {
  try {
    if (!pushServiceAccount_()) return;
    text = String(text || '');
    if (text === '!update' || text === '!deleted' || text.indexOf('!poll ') === 0) return;
    if (!text && !hasImage) return;

    const named = {};
    (text.match(/[@＠][^\s　@＠]+/g) || []).forEach(function(m) { named[m.slice(1).toLowerCase()] = true; });
    const isAll = !!named['all'];

    const ld = pushLoad_();
    const hm = Utilities.formatDate(new Date(), PUSH_TZ, 'HH:mm');
    const targets = [];
    ld.rows.forEach(function(r) {
      const uid = String(r[0]);
      if (!uid || uid === String(from)) return;
      const prefs = pushPrefsClean_(r[2]);
      const isMe = !!named[uid.toLowerCase()];
      if (pushShouldSend_(prefs, room, isMe, isAll, hm)) targets.push({ token: String(r[1]), isMe: isMe });
    });
    if (!targets.length) return;

    const body = text ? text.replace(/\s+/g, ' ') : '📷 画像を送信しました';
    const mk = function(prefix) { return prefix + from + ' # ' + room; };
    const groups = [
      { list: targets.filter(function(t) { return t.isMe; }), title: mk('🔔 ') },
      { list: targets.filter(function(t) { return !t.isMe; }), title: isAll ? mk('📢 ') : mk('') }
    ];
    groups.forEach(function(g) { if (g.list.length) pushSend_(g.list, g.title, body, 'chat-' + room); });
  } catch (e) {
    Logger.log('プッシュ通知(チャット)エラー: ' + e);
  }
}

// DMが届いたとき(sendDM から呼ぶ)
function pushForDm_(from, to, text) {
  try {
    if (!pushServiceAccount_()) return;
    const ld = pushLoad_();
    const targets = [];
    ld.rows.forEach(function(r) {
      if (String(r[0]) !== String(to)) return;
      if (pushPrefsClean_(r[2]).off) return;
      targets.push({ token: String(r[1]) });
    });
    if (!targets.length) return;
    pushSend_(targets, '💌 ' + from + ' さんからDM', String(text).replace(/\s+/g, ' '), 'dm-' + from);
  } catch (e) {
    Logger.log('プッシュ通知(DM)エラー: ' + e);
  }
}

// ---- エディタから手動で実行する確認用 ----
// 1) 設定が正しいか(鍵が読めて、FCMの認証が通るか)
function checkPushSetup() {
  const sa = pushServiceAccount_();
  if (!sa) { Logger.log('❌ スクリプトプロパティ「FCM_SA」がありません(サービスアカウントの鍵JSONを丸ごと貼ってください)'); return; }
  PropertiesService.getScriptProperties(); // 権限の確認用
  CacheService.getScriptCache().remove('fcm_at');
  const at = fcmAccessToken_(sa);
  Logger.log(at ? '✅ FCMの認証に成功しました(プロジェクト: ' + sa.project_id + ')' : '❌ FCMの認証に失敗しました(上のログを見てください)');
  const sh = getPushSheet_();
  Logger.log('登録されている端末: ' + Math.max(0, sh.getLastRow() - 1) + '台');
}
// 2) 自分の端末にテスト通知を送る(下の TEST_ID を、自分のIDに変えて実行)
function testPush() {
  const TEST_ID = 'juki2023';
  const sh = getPushSheet_();
  const last = sh.getLastRow();
  const rows = last >= 2 ? sh.getRange(2, 1, last - 1, 2).getValues() : [];
  const targets = rows.filter(function(r) { return String(r[0]) === TEST_ID; }).map(function(r) { return { token: String(r[1]) }; });
  Logger.log('送り先: ' + targets.length + '台');
  pushSend_(targets, '📲 テスト通知', 'これが見えれば、プッシュ通知は使えます', 'test');
}
