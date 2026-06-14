/**
 * judo-seifukuin / 予約バックエンド
 *
 *  機能:
 *    1) doGet              — LIFFからの予約一覧取得API (本人のみ)
 *    2) sendDailyReminders — 毎日定時のリマインダー送信 (時間トリガー)
 *    3) sendConfirmations  — 新規予約の確定通知 (時間トリガー / 数分おき)
 *
 *  セットアップ:
 *    A. プロジェクト設定 → スクリプトプロパティ に下記を登録
 *         LINE_CHANNEL_ACCESS_TOKEN  : Messaging APIのチャネルアクセストークン (長期)
 *         LIFF_CHANNEL_ID            : LIFFを動かすLINEログインチャネルのID
 *    B. トリガー設定:
 *         sendDailyReminders : 時間主導型 → 日タイマー → 午後6〜7時
 *         sendConfirmations  : 時間主導型 → 分タイマー → 5分おき
 *    C. デプロイ → 新しいデプロイ → ウェブアプリ
 *         実行ユーザー   : 自分
 *         アクセス       : 全員
 *         発行URLを reservations.html の API_ENDPOINT に貼る
 *
 *  シート列構成 (1行目はヘッダー):
 *    A: userId      LINEのuserId (Uxxxxxxxx...)
 *    B: datetime    予約日時 (Date型)
 *    C: menu        メニュー名
 *    D: note        備考
 *    E: status      "確定" / "キャンセル" など
 *    F: reminderSentAt    リマインド送信日時 (Date / 空欄なら未送信)
 *    G: confirmationSentAt 確定通知送信日時 (Date / 空欄なら未送信)
 *    H: customerName 表示用の顧客名 (任意)
 */

const SHEET_NAME = '予約一覧';
const TZ = 'Asia/Tokyo';

const COL = {
  userId:              0, // A
  datetime:            1, // B
  menu:                2, // C
  note:                3, // D
  status:              4, // E
  reminderSentAt:      5, // F
  confirmationSentAt:  6, // G
  customerName:        7, // H
};

// ────────────────────────────────────────────────
//  1) LIFFからの予約一覧取得API
// ────────────────────────────────────────────────
function doGet(e) {
  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);

  try {
    const idToken = e.parameter.idToken;
    if (!idToken) throw new Error('idToken required');

    const verified = verifyIdToken_(idToken);
    const userId = verified.sub;

    const list = readReservations_()
      .filter(r => r.userId === userId && r.status !== 'キャンセル')
      .map(r => ({
        datetime: Utilities.formatDate(r.datetime, TZ, "yyyy-MM-dd'T'HH:mm:ss"),
        menu:     r.menu,
        note:     r.note,
        status:   r.status,
      }));

    out.setContent(JSON.stringify(list));
  } catch (err) {
    out.setContent(JSON.stringify({ error: String(err.message || err) }));
  }
  return out;
}

function verifyIdToken_(idToken) {
  const channelId = getProp_('LIFF_CHANNEL_ID');
  const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post',
    payload: { id_token: idToken, client_id: channelId },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('invalid id token');
  }
  return JSON.parse(res.getContentText());
}

// ────────────────────────────────────────────────
//  2) 翌日のリマインダー送信 (1日1回トリガー)
// ────────────────────────────────────────────────
function sendDailyReminders() {
  const sheet = getSheet_();
  const all = readReservations_();
  const now = new Date();
  const tomorrowStart = startOfDay_(addDays_(now, 1));
  const tomorrowEnd   = startOfDay_(addDays_(now, 2));

  all.forEach(r => {
    if (r.status === 'キャンセル') return;
    if (r.reminderSentAt) return;
    if (!(r.datetime >= tomorrowStart && r.datetime < tomorrowEnd)) return;

    const text = buildReminderText_(r);
    const ok = pushLine_(r.userId, text);
    if (ok) {
      sheet.getRange(r.row, COL.reminderSentAt + 1).setValue(new Date());
    }
  });
}

function buildReminderText_(r) {
  const dt = Utilities.formatDate(r.datetime, TZ, 'M月d日(E) HH:mm');
  const name = r.customerName ? r.customerName + '様' : '';
  return [
    name ? name : 'こんにちは。',
    '',
    '明日のご予約のお知らせです。',
    '',
    '📅 ' + dt,
    r.menu ? '🩹 ' + r.menu : '',
    '',
    'お会いできるのを楽しみにしております。',
    '変更・キャンセルはこのトークからご連絡ください。',
  ].filter(Boolean).join('\n');
}

// ────────────────────────────────────────────────
//  3) 新規予約の確定通知 (数分おきトリガー)
// ────────────────────────────────────────────────
function sendConfirmations() {
  const sheet = getSheet_();
  const all = readReservations_();
  const now = new Date();

  all.forEach(r => {
    if (r.confirmationSentAt) return;
    if (r.status === 'キャンセル') return;
    if (!r.userId || !r.datetime) return;
    if (r.datetime < now) return; // 過去のものは無視

    const text = buildConfirmationText_(r);
    const ok = pushLine_(r.userId, text);
    if (ok) {
      sheet.getRange(r.row, COL.confirmationSentAt + 1).setValue(new Date());
    }
  });
}

function buildConfirmationText_(r) {
  const dt = Utilities.formatDate(r.datetime, TZ, 'M月d日(E) HH:mm');
  const name = r.customerName ? r.customerName + '様' : '';
  return [
    name ? name : 'ご予約ありがとうございます。',
    '',
    '下記の内容で承りました。',
    '',
    '📅 ' + dt,
    r.menu ? '🩹 ' + r.menu : '',
    '',
    'お気をつけてお越しください。',
    '変更・キャンセルはこのトークから承ります。',
  ].filter(Boolean).join('\n');
}

// ────────────────────────────────────────────────
//  共通: シート読み書き
// ────────────────────────────────────────────────
function getSheet_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('sheet not found: ' + SHEET_NAME);
  return sheet;
}

function readReservations_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1);
  return rows.map((r, i) => ({
    row:                 i + 2,
    userId:              String(r[COL.userId] || '').trim(),
    datetime:            r[COL.datetime] instanceof Date ? r[COL.datetime] : new Date(r[COL.datetime]),
    menu:                String(r[COL.menu] || ''),
    note:                String(r[COL.note] || ''),
    status:              String(r[COL.status] || ''),
    reminderSentAt:      r[COL.reminderSentAt],
    confirmationSentAt:  r[COL.confirmationSentAt],
    customerName:        String(r[COL.customerName] || ''),
  })).filter(r => r.userId);
}

// ────────────────────────────────────────────────
//  共通: LINE Messaging API
// ────────────────────────────────────────────────
function pushLine_(userId, text) {
  const token = getProp_('LINE_CHANNEL_ACCESS_TOKEN');
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify({
      to: userId,
      messages: [{ type: 'text', text: text }],
    }),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code !== 200) {
    console.error('LINE push failed', code, res.getContentText());
    return false;
  }
  return true;
}

function getProp_(key) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v) throw new Error('missing script property: ' + key);
  return v;
}

// ────────────────────────────────────────────────
//  日時ユーティリティ
// ────────────────────────────────────────────────
function startOfDay_(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}
function addDays_(d, n) {
  const x = new Date(d);
  x.setDate(x.getDate() + n);
  return x;
}

// ────────────────────────────────────────────────
//  動作確認用 (手動実行)
// ────────────────────────────────────────────────
function _testPushToMe() {
  // 自分のuserIdを入れて手動で1通テスト送信
  const MY_USER_ID = 'Uxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';
  pushLine_(MY_USER_ID, '通知テストです。');
}
