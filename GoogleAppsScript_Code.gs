/**
 * 車両管理アプリ用バックエンド（Google Apps Script）
 * ------------------------------------------------------
 * スプレッドシートを本体データベースとして使い、車両管理アプリ（HTML版）
 * からの読み書き（一覧取得・追加・編集・削除・並び替え・事業所管理）を
 * 受け付けるWebアプリです。
 *
 * 【使い方】
 * 1. 新規のGoogleスプレッドシートを用意する（シート名は何でもよい）
 * 2. 拡張機能 → Apps Script を開き、このコード全体を貼り付けて保存
 * 3. 下の API_TOKEN を書き換える（他人に推測されない適当な文字列に）
 *    ※ index.html 側の設定パネルに、ここで決めたのと同じ値を入力します
 * 4. 上部の「デプロイ」→「新しいデプロイ」→種類「ウェブアプリ」を選択
 * 5. 「実行するユーザー」＝自分（Me）
 *    「アクセスできるユーザー」＝
 *       ・組織のGoogle Workspaceがあれば「〇〇（組織名）内の全員」
 *       ・まだ無ければ、暫定的に「全員」を選択（URLとトークンを知る人だけがアクセス可）
 * 6. デプロイ後に表示される「ウェブアプリのURL」（.../exec で終わるもの）を
 *    index.html の設定パネルに入力する
 *
 * 【初回のみ】このスクリプトを一度手動実行（関数 seedIfEmpty を選んで実行）
 * すると、車両シートが空の場合に元データ21台分を自動投入できます。
 * 実行時にスプレッドシートへのアクセス許可を求められるので許可してください。
 * 既にデータがある場合は何もしません（二重投入の心配はありません）。
 *
 * 【コードを更新した場合の再デプロイ】
 * コードを保存しただけでは、公開中のURLには反映されません。
 * 「デプロイ」→「デプロイを管理」→ 編集（鉛筆アイコン）→
 * バージョン「新しいバージョン」を選んで「デプロイ」を押してください
 * （URLは変わりません）。
 *
 * 【複数人での同時利用について】
 * 書き込み処理はすべて LockService でロックしてから行うため、
 * 複数人が同時に編集しても行が壊れたり重複したりしません。
 */

const VEHICLE_SHEET_NAME = "Vehicles";
const OFFICE_SHEET_NAME = "Offices";

// APIトークン（合言葉）。この値と一致しないリクエストは拒否します。
// index.html の設定パネルで入力する値と必ず同じにしてください。
const API_TOKEN = "CHANGE_ME_TO_A_RANDOM_SECRET"; // ←必ず自分で決めた値に書き換えてください（このリポジトリはGitHub上でpublic公開されるため、既定値のままデプロイしないこと）

// 書き込みを許可するメールアドレス（空欄なら誰でも書き込み可＝暫定運用向け）
// 組織のGoogleアカウントが決まったら、ここに列挙してアクセスを絞ってください。
const ALLOWED_EMAILS = [];

const VEHICLE_FIELDS = [
  "id", "sortOrder", "office", "type", "plate", "owner",
  "leaseEnd", "first", "purchase", "shaken", "summer", "studless", "note"
];

function checkToken_(token) {
  return API_TOKEN && token === API_TOKEN;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getVehicleSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(VEHICLE_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(VEHICLE_SHEET_NAME);
    sh.appendRow(VEHICLE_FIELDS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getOfficeSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(OFFICE_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(OFFICE_SHEET_NAME);
    sh.appendRow(["name"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readVehicles_() {
  const sh = getVehicleSheet_();
  const values = sh.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[0] && row[0] !== 0) continue;
    const obj = {};
    VEHICLE_FIELDS.forEach(function (f, idx) {
      obj[f] = row[idx] === undefined ? "" : row[idx];
    });
    obj.id = Number(obj.id);
    obj.sortOrder = Number(obj.sortOrder) || 0;
    ["first", "purchase", "shaken", "leaseEnd"].forEach(function (f) {
      obj[f] = formatDateValue_(obj[f]);
    });
    rows.push(obj);
  }
  rows.sort(function (a, b) { return a.sortOrder - b.sortOrder; });
  return rows;
}

function formatDateValue_(v) {
  if (!v) return "";
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(v);
}

function readOffices_() {
  const sh = getOfficeSheet_();
  const values = sh.getDataRange().getValues();
  const names = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i][0]) names.push(String(values[i][0]));
  }
  return names;
}

function findVehicleRow_(sh, id) {
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (Number(values[i][0]) === Number(id)) return i + 1; // 1-indexed sheet row
  }
  return -1;
}

function nextVehicleId_(sh) {
  const values = sh.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < values.length; i++) {
    const v = Number(values[i][0]);
    if (v > max) max = v;
  }
  return max + 1;
}

function nextSortOrder_(sh) {
  const values = sh.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < values.length; i++) {
    const v = Number(values[i][1]);
    if (v > max) max = v;
  }
  return max + 1;
}

function currentUserEmail_() {
  try { return Session.getActiveUser().getEmail(); } catch (err) { return ""; }
}

function isAllowed_() {
  if (!ALLOWED_EMAILS.length) return true;
  return ALLOWED_EMAILS.indexOf(currentUserEmail_()) !== -1;
}

// ---- GET: 一覧取得 ?token=API_TOKEN ----
function doGet(e) {
  try {
    const token = (e && e.parameter && e.parameter.token) || "";
    if (!checkToken_(token)) return jsonOut_({ ok: false, error: "forbidden" });
    return jsonOut_({ ok: true, vehicles: readVehicles_(), offices: readOffices_() });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

// ---- POST: 追加・編集・削除・並び替え・事業所管理 ----
// body: { token, action, payload }
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (!checkToken_(body && body.token)) return jsonOut_({ ok: false, error: "forbidden" });
    if (!isAllowed_()) return jsonOut_({ ok: false, error: "forbidden", email: currentUserEmail_() });

    const action = body.action;
    const payload = body.payload || {};

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      switch (action) {
        case "addVehicle": return jsonOut_(addVehicle_(payload));
        case "updateVehicle": return jsonOut_(updateVehicle_(payload));
        case "updateField": return jsonOut_(updateField_(payload));
        case "deleteVehicle": return jsonOut_(deleteVehicle_(payload));
        case "moveVehicle": return jsonOut_(moveVehicle_(payload));
        case "addOffice": return jsonOut_(addOffice_(payload));
        case "removeOffice": return jsonOut_(removeOffice_(payload));
        case "seedIfEmpty": return jsonOut_(seedIfEmptyInternal_());
        default: return { ok: false, error: "unknown action: " + action };
      }
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function addVehicle_(payload) {
  const sh = getVehicleSheet_();
  const id = nextVehicleId_(sh);
  const sortOrder = nextSortOrder_(sh);
  const row = VEHICLE_FIELDS.map(function (f) {
    if (f === "id") return id;
    if (f === "sortOrder") return sortOrder;
    return payload[f] || "";
  });
  sh.appendRow(row);
  SpreadsheetApp.flush();
  const vehicle = {};
  VEHICLE_FIELDS.forEach(function (f, idx) { vehicle[f] = row[idx]; });
  return { ok: true, vehicle: vehicle };
}

// 入力フォームからの一括更新: id以外の全項目をまとめて上書きする
function updateVehicle_(payload) {
  const sh = getVehicleSheet_();
  const rowIndex = findVehicleRow_(sh, payload.id);
  if (rowIndex < 0) return { ok: false, error: "vehicle not found" };
  const editableFields = VEHICLE_FIELDS.filter(function (f) { return f !== "id" && f !== "sortOrder"; });
  const values = editableFields.map(function (f) { return payload[f] || ""; });
  const firstCol = VEHICLE_FIELDS.indexOf(editableFields[0]) + 1;
  sh.getRange(rowIndex, firstCol, 1, editableFields.length).setValues([values]);
  SpreadsheetApp.flush();
  return { ok: true };
}

function updateField_(payload) {
  const sh = getVehicleSheet_();
  const rowIndex = findVehicleRow_(sh, payload.id);
  if (rowIndex < 0) return { ok: false, error: "vehicle not found" };
  const colIndex = VEHICLE_FIELDS.indexOf(payload.field);
  if (colIndex < 0 || payload.field === "id" || payload.field === "sortOrder") {
    return { ok: false, error: "invalid field" };
  }
  sh.getRange(rowIndex, colIndex + 1).setValue(payload.value);
  SpreadsheetApp.flush();
  return { ok: true };
}

function deleteVehicle_(payload) {
  const sh = getVehicleSheet_();
  const rowIndex = findVehicleRow_(sh, payload.id);
  if (rowIndex < 0) return { ok: false, error: "vehicle not found" };
  sh.deleteRow(rowIndex);
  SpreadsheetApp.flush();
  return { ok: true };
}

function moveVehicle_(payload) {
  const sh = getVehicleSheet_();
  const rowA = findVehicleRow_(sh, payload.idA);
  const rowB = findVehicleRow_(sh, payload.idB);
  if (rowA < 0 || rowB < 0) return { ok: false, error: "vehicle not found" };
  const sortColIndex = VEHICLE_FIELDS.indexOf("sortOrder") + 1;
  const sortA = sh.getRange(rowA, sortColIndex).getValue();
  const sortB = sh.getRange(rowB, sortColIndex).getValue();
  sh.getRange(rowA, sortColIndex).setValue(sortB);
  sh.getRange(rowB, sortColIndex).setValue(sortA);
  SpreadsheetApp.flush();
  return { ok: true };
}

function addOffice_(payload) {
  const name = String(payload.name || "").trim();
  if (!name) return { ok: false, error: "name is required" };
  const sh = getOfficeSheet_();
  const existing = readOffices_();
  if (existing.indexOf(name) === -1) {
    sh.appendRow([name]);
    SpreadsheetApp.flush();
  }
  return { ok: true };
}

function removeOffice_(payload) {
  const name = String(payload.name || "").trim();
  const sh = getOfficeSheet_();
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === name) {
      sh.deleteRow(i + 1);
      SpreadsheetApp.flush();
      break;
    }
  }
  return { ok: true };
}

// 初回セットアップ用: 車両シートが空のときだけ、元データ21台を投入する
function seedIfEmptyInternal_() {
  const sh = getVehicleSheet_();
  if (sh.getLastRow() > 1) return { ok: true, skipped: true, reason: "already has data" };

  var seed = [
    {office:"リハビリサロンひだまり", type:"ハイエースレジアス", plate:"一宮830 す・510", owner:"CENTO ANNI", leaseEnd:"", first:"", purchase:"", shaken:"", summer:"195/80/15", studless:"195/80/15", note:"6穴"},
    {office:"リハビリサロンひだまり", type:"ハイエース", plate:"一宮830 さ・510", owner:"CENTO ANNI", leaseEnd:"", first:"", purchase:"", shaken:"", summer:"195/80/15", studless:"195/80/15", note:"6穴"},
    {office:"リハビリサロンひだまり", type:"ヴォクシー", plate:"一宮533 も・510", owner:"CENTO ANNI", leaseEnd:"", first:"", purchase:"", shaken:"", summer:"195/65/15", studless:"195/65/15", note:"5穴"},
    {office:"リハビリサロンひだまり", type:"セレナ", plate:"一宮533 ま・510", owner:"CENTO ANNI", leaseEnd:"", first:"", purchase:"", shaken:"", summer:"195/65/15", studless:"", note:""},
    {office:"リハビリサロンひだまり", type:"セレナ", plate:"一宮533 ら・510", owner:"CENTO ANNI", leaseEnd:"", first:"", purchase:"", shaken:"", summer:"195/65/15", studless:"", note:""},
    {office:"リハビリサロンひだまり", type:"タント", plate:"一宮585 め・510", owner:"CENTO ANNI", leaseEnd:"", first:"", purchase:"", shaken:"", summer:"155/70/13", studless:"", note:""},
    {office:"幸の鳥", type:"ミライース", plate:"一宮585 ほ・510", owner:"清流会", leaseEnd:"", first:"", purchase:"", shaken:"", summer:"155/70/13", studless:"155/70/13", note:"4穴"},
    {office:"幸の鳥", type:"ミライース", plate:"一宮585 た・510", owner:"清流会", leaseEnd:"", first:"", purchase:"", shaken:"", summer:"155/70/13", studless:"155/70/13", note:"4穴"},
    {office:"幸の鳥", type:"ミライース", plate:"一宮585 ひ・510", owner:"清流会", leaseEnd:"", first:"", purchase:"", shaken:"", summer:"155/70/13", studless:"155/70/13", note:"4穴"},
    {office:"幸の鳥", type:"ミライース", plate:"一宮585 ぬ・510", owner:"清流会", leaseEnd:"", first:"", purchase:"", shaken:"", summer:"155/70/13", studless:"155/70/13", note:"4穴"},
    {office:"幸の鳥", type:"ミライース", plate:"一宮585 は・510", owner:"清流会", leaseEnd:"", first:"", purchase:"", shaken:"", summer:"155/70/13", studless:"155/70/13", note:"4穴"},
    {office:"おむすび", type:"ミライース", plate:"一宮585 な・510", owner:"清流会", leaseEnd:"", first:"", purchase:"", shaken:"", summer:"155/70/13", studless:"155/70/13", note:"4穴"},
    {office:"おむすび", type:"ミライース", plate:"一宮585 ち・510", owner:"清流会", leaseEnd:"", first:"", purchase:"", shaken:"", summer:"155/70/13", studless:"155/70/13", note:"4穴"},
    {office:"おむすび", type:"アルト", plate:"一宮586 ぬ・510", owner:"CENTO ANNI", leaseEnd:"", first:"", purchase:"", shaken:"", summer:"", studless:"", note:""},
    {office:"幸の鳥", type:"アルト", plate:"一宮586 に・510", owner:"CENTO ANNI", leaseEnd:"", first:"", purchase:"", shaken:"", summer:"", studless:"", note:""},
    {office:"幸の鳥", type:"ミライース", plate:"一宮586 た・510", owner:"CENTO ANNI", leaseEnd:"", first:"", purchase:"", shaken:"", summer:"", studless:"", note:""},
    {office:"幸の鳥", type:"ミライース", plate:"一宮586 そ・510", owner:"CENTO ANNI", leaseEnd:"", first:"", purchase:"", shaken:"", summer:"", studless:"", note:""},
    {office:"五藤医院", type:"アクア", plate:"一宮532 み・510", owner:"清流会", leaseEnd:"", first:"", purchase:"", shaken:"", summer:"", studless:"", note:""},
    {office:"五藤医院", type:"ルーミー", plate:"一宮533 ろ・510", owner:"CENTO ANNI", leaseEnd:"", first:"", purchase:"", shaken:"", summer:"", studless:"", note:""},
    {office:"五藤医院", type:"アルト", plate:"一宮586 く・510", owner:"CENTO ANNI", leaseEnd:"", first:"", purchase:"", shaken:"", summer:"", studless:"", note:""},
    {office:"五藤医院", type:"アクア", plate:"一宮532 さ・510", owner:"CENTO ANNI", leaseEnd:"", first:"", purchase:"", shaken:"", summer:"", studless:"", note:""}
  ];

  const rows = seed.map(function (v, i) {
    return VEHICLE_FIELDS.map(function (f) {
      if (f === "id") return i + 1;
      if (f === "sortOrder") return i + 1;
      return v[f] || "";
    });
  });
  sh.getRange(2, 1, rows.length, VEHICLE_FIELDS.length).setValues(rows);
  SpreadsheetApp.flush();
  return { ok: true, seeded: rows.length };
}

// Apps Scriptエディタから直接実行する用（引数なしで呼べる関数）
function seedIfEmpty() {
  const result = seedIfEmptyInternal_();
  Logger.log(JSON.stringify(result));
}
