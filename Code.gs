/***** 設定ここから *****/
// Web アプリの公開 URL（/exec）。空欄ならこのプロジェクトのデプロイ URL を自動で使う
const BASE_WEBAPP_URL = '';

// ---- 作業依頼（既存フォーム） ----
const SPREADSHEET_ID = '1oyVI671g6kZHmJRAqawLHS7EXqfKM_3wrfR-GwYZy8U';
const SHEET_NAME     = 'フォーム回答';

// アップロード先フォルダ（任意・自動作成可）
const DRIVE_FOLDER_FILES_ID  = '1vwrV-hyRPnFddqlmHpzvCOHstnPQD5xl'; // 画像以外ファイル
const DRIVE_FOLDER_HTML_ID   = ''; // 本文HTML保存用（未設定なら自動作成）

// 自動作成時のフォルダ名
const AUTO_FOLDER_FILES_NAME = 'NEO-問い合わせ-その他';
const AUTO_FOLDER_HTML_NAME  = 'NEO-問い合わせ-本文HTML';

// ---- 新規開発・仕様変更依頼 ----
const DEV_SPREADSHEET_ID   = '1XKe47vEB_e351nmqOZsVwzfGg6TDzOOhPPdPuqs7cNc';
const DEV_SHEET_NAME       = '小規模案件';
const DEV_DRIVE_FOLDER_ID  = '1mgH317-UZ1JBcxyOt_spx84VhbSJm1Ye'; // 資料アップ先（F列見出しのフォルダ）
const AUTO_FOLDER_DEV_NAME = 'NEO-開発依頼-資料';

// ---- 不具合報告 ----
const BUG_SPREADSHEET_ID   = '1odxPfKhfA3T48dEgPHv1F7j4AdpWxSU4l6J-Yqb-pOU';
const BUG_SHEET_NAME       = 'フィードバック一覧';
const BUG_DRIVE_FOLDER_ID  = ''; // スクリーンショット等の保存先（未設定なら自動作成）
const AUTO_FOLDER_BUG_NAME = 'NEO-不具合報告-添付';

// Script Properties に保存するキー
const PROP_FILES_FOLDER_ID = 'UPLOAD_FILES_FOLDER_ID';
const PROP_HTML_FOLDER_ID  = 'UPLOAD_HTML_FOLDER_ID';
const PROP_DEV_FOLDER_ID   = 'UPLOAD_DEV_FOLDER_ID';
const PROP_BUG_FOLDER_ID   = 'UPLOAD_BUG_FOLDER_ID';
const PROP_JOB_PREFIX      = 'JOB_';      // 非同期ジョブ（1ジョブ1キー）
const PROP_ASYNC_JOBS_OLD  = 'ASYNC_JOBS'; // 旧形式キュー（移行用に読み取りのみ）

// Drive の共有リンク（Anyone with the link）を自動付与する場合 true
const MAKE_FILE_PUBLIC = true;

// 回答者へ控えメールを送るかどうか
const SEND_COPY_TO_RESPONDENT = true;

// メール送信メソッド: "auto" | "gmail" | "mail"
const COPY_MAIL_METHOD = 'auto';

// 添付ファイル上限（1ファイルあたり）
const MAX_FILE_BYTES = 15 * 1024 * 1024;

// 非同期ジョブ：トリガーのハンドラ名と最大リトライ回数
const JOB_TRIGGER_HANDLER = 'runQueuedJobs';
const JOB_MAX_ATTEMPTS    = 5;
/***** 設定ここまで *****/


/** 画面配信：URL パラメータをテンプレートへ渡す */
function doGet(e) {
  // ▼ プレビュー表示（?mode=preview&id=DriveFileId）
  if (e && e.parameter && e.parameter.mode === 'preview') {
    const fileId = (e.parameter.id || '').trim();
    if (!fileId) {
      return HtmlService.createHtmlOutput('<p>プレビュー対象のIDが指定されていません。</p>');
    }
    try {
      const file = DriveApp.getFileById(fileId);
      const html = file.getBlob().getDataAsString('UTF-8');

      const page = `
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <div style="font-family:system-ui,Segoe UI,Roboto,Arial,sans-serif;line-height:1.7;padding:16px;max-width:880px;margin:auto;">
          <h2 style="margin:0 0 12px;">NEO 問い合わせ：本文プレビュー</h2>
          <div style="border:1px solid #ddd;border-radius:8px;padding:14px;background:#fff;">
            ${html}
          </div>
        </div>
      `;
      return HtmlService.createHtmlOutput(page)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch (err) {
      return HtmlService.createHtmlOutput('<p>プレビュー用ファイルが見つかりませんでした。</p>');
    }
  }

  // 通常フォーム画面（?type=work|bug|dev|pc|material で初期タブを指定）
  const t = HtmlService.createTemplateFromFile('index');
  t.params = (e && e.parameter) ? e.parameter : {};
  return t.evaluate()
    .setTitle('NEO 総合依頼フォーム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** 本文プレビュー等のリンクに使う Web アプリ URL */
function webAppUrl_() {
  return BASE_WEBAPP_URL || ScriptApp.getService().getUrl();
}

/** テンプレートから部品 HTML を読み込む */
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/** テンプレートへ埋め込む JSON（</script> 混入対策） */
function jsonForScript_(obj) {
  return JSON.stringify(obj || {}).replace(/</g, '\\u003c');
}


/** ========== 非同期（後段）処理：公開化・URL反映・控えメール送信 ========== */
/*
 * ジョブは Script Properties に 1件1キー（JOB_xxx）で保存する。
 *   - submit 系関数の最後でインライン実行（通常はここで完了）
 *   - 失敗・未処理分はワンショットトリガー（runQueuedJobs）で再実行
 * ジョブの種類（type）: 'work' | 'dev' | 'bug'
 */

/** キューにジョブを積む */
function enqueueJob_(job) {
  const key = PROP_JOB_PREFIX + Date.now() + '_' + Utilities.getUuid().slice(0, 8);
  const value = JSON.stringify(Object.assign({ attempts: 0, enqueuedAt: Date.now() }, job));
  PropertiesService.getScriptProperties().setProperty(key, value);
}

/** 1回だけの時間主導トリガーを作る（なければ） */
function ensureOneShotTrigger_(delayMs) {
  try {
    const exists = ScriptApp.getProjectTriggers()
      .some(t => t.getHandlerFunction() === JOB_TRIGGER_HANDLER);
    if (!exists) {
      ScriptApp.newTrigger(JOB_TRIGGER_HANDLER).timeBased().after(delayMs || 1000).create();
    }
  } catch (e) {
    // トリガー作成に失敗しても、インライン実行しているため致命傷にはしない
    console.warn('ensureOneShotTrigger_ failed:', e);
  }
}

/** トリガーから呼ばれる入口（末尾 _ の関数はトリガーから呼べないため公開名にする） */
function runQueuedJobs() {
  // 実行済みのワンショットトリガーを掃除（旧ハンドラ名 processJobs_ も含む）
  try {
    ScriptApp.getProjectTriggers()
      .filter(t => [JOB_TRIGGER_HANDLER, 'processJobs_'].includes(t.getHandlerFunction()))
      .forEach(t => ScriptApp.deleteTrigger(t));
  } catch (e) { console.warn('trigger cleanup failed:', e); }

  processJobs_();

  // 失敗分が残っていれば 1分後に再実行を予約
  if (listJobKeys_().length) ensureOneShotTrigger_(60 * 1000);
}

function listJobKeys_() {
  return PropertiesService.getScriptProperties().getKeys()
    .filter(k => k.indexOf(PROP_JOB_PREFIX) === 0);
}

/** ジョブ処理本体（トリガー／インラインから呼ばれる） */
function processJobs_() {
  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  let jobs = [];
  try {
    // 取り出しつつ削除（重複実行防止）
    for (const key of listJobKeys_()) {
      const raw = props.getProperty(key);
      props.deleteProperty(key);
      if (raw) { try { jobs.push(JSON.parse(raw)); } catch (_) {} }
    }
    // 旧形式キューが残っていれば取り込む
    const old = props.getProperty(PROP_ASYNC_JOBS_OLD);
    if (old) {
      props.deleteProperty(PROP_ASYNC_JOBS_OLD);
      try {
        JSON.parse(old).forEach(it => jobs.push(Object.assign({ type: 'work', rowIndex: it.rowIndex, attempts: 0 }, it.jobData)));
      } catch (_) {}
    }
  } finally {
    lock.releaseLock();
  }

  for (const job of jobs) {
    try {
      if (job.type === 'dev')      runDevJob_(job);
      else if (job.type === 'bug') runBugJob_(job);
      else                         runWorkJob_(job);
    } catch (err) {
      console.error('job failed:', JSON.stringify(job).slice(0, 2000), err);
      job.attempts = (job.attempts || 0) + 1;
      if (job.attempts < JOB_MAX_ATTEMPTS) enqueueJob_(job); // 次回リトライ
    }
  }
}

/** ファイルを「リンクを知っている全員が閲覧可」にする */
function makeFilesPublic_(fileIds) {
  if (!MAKE_FILE_PUBLIC) return;
  for (const id of (fileIds || [])) {
    try { DriveApp.getFileById(id).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (_) {}
  }
}


/** ========== ユーティリティ ========== */

/** スクリプトロック下で fn を実行 */
function withScriptLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

/** 指定列の数値の最大値 + 1（見出しや空欄は無視） */
function nextNumericId_(ws, col) {
  const last = ws.getLastRow();
  if (last < 1) return 1;
  let max = 0;
  for (const [v] of ws.getRange(1, col, last, 1).getValues()) {
    const n = (typeof v === 'number') ? v : Number(String(v).trim() || NaN);
    if (isFinite(n) && n > max) max = n;
  }
  return max + 1;
}

/** シート取得（なければエラー） */
function getSheet_(spreadsheetId, sheetName) {
  const ws = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  if (!ws) throw new Error(`シートが見つかりません: ${sheetName}`);
  return ws;
}

/** 数式・数値として解釈されないよう文字列として書き込む */
function asSheetText_(v) {
  const s = (v === undefined || v === null) ? '' : String(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

/** メール本文用に長文を切り詰める（Script Properties の容量対策） */
function truncate_(s, max) {
  s = String(s || '');
  return s.length > max ? s.slice(0, max) + '\n…（以下省略）' : s;
}

/** 添付ファイルを指定フォルダに保存し、ファイルIDの配列を返す */
function saveFiles_(files, folder) {
  const ids = [];
  for (const [idx, f] of (files || []).entries()) {
    const fname = f.name || `file-${Date.now()}-${idx}`;
    ids.push(folder.createFile(dataUrlToBlob_(f.dataUrl, fname)).getId());
  }
  return ids;
}

/** 添付ファイルの事前検査（サイズのみ。画像可否は呼び出し側で判定） */
function checkFileSizes_(files) {
  for (const f of (files || [])) {
    const m = String((f && f.dataUrl) || '').replace(/\s/g, '').match(/^data:([^;]*);base64,([A-Za-z0-9+/=]+)$/);
    if (!m) throw new Error(`ファイル「${(f && f.name) || ''}」を読み込めませんでした。`);
    if (Math.floor(m[2].length * 0.75) > MAX_FILE_BYTES) {
      throw new Error(`添付ファイル「${f.name}」のサイズが上限(15MB)を超えています。`);
    }
  }
}

/**
 * 添付ファイルの事前検査（作業依頼用）：
 *  - 画像(mime: image/*)は不可
 *  - 15MB超は不可
 * ここで throw したら submit は何も書き込まず終了
 */
function prevalidateFiles_(files) {
  for (const f of (files || [])) {
    const mime = (f && f.mime) || '';
    const dataUrl = (f && f.dataUrl) || '';

    // 画像の排除（mime か dataURL プレフィックスで判定）
    if ((mime && mime.startsWith('image/')) || /^data:image\//i.test(dataUrl)) {
      throw new Error('「添付ファイル」に画像は指定できません。画像は本文欄に直接貼り付けてください。');
    }

    // サイズ上限（約 15MB）。base64長×0.75 で概算
    const m = String(dataUrl).replace(/\s/g,'').match(/^data:([^;]+);base64,([A-Za-z0-9+/=]+)$/);
    if (m) {
      const approxBytes = Math.floor(m[2].length * 0.75);
      if (approxBytes > MAX_FILE_BYTES) {
        throw new Error('添付ファイルのサイズが上限(15MB)を超えています。');
      }
    }
  }
}

/** Drive フォルダ取得（ID → ScriptProp → 自動作成） */
function ensureFolder_(candidateId, propKey, createName) {
  const props = PropertiesService.getScriptProperties();

  const cand = (candidateId || '').trim();
  if (cand) {
    const f = getFolderByIdSafe_(cand);
    if (f) { props.setProperty(propKey, cand); return f; }
  }

  const saved = (props.getProperty(propKey) || '').trim();
  if (saved) {
    const f2 = getFolderByIdSafe_(saved);
    if (f2) return f2;
  }

  const created = DriveApp.createFolder(createName);
  props.setProperty(propKey, created.getId());
  return created;
}

function getFolderByIdSafe_(id) {
  try {
    if (!id) return null;
    const f = DriveApp.getFolderById(id);
    f.getName();
    return f;
  } catch (e) {
    return null;
  }
}

/** DataURL -> Blob（15MB上限） */
function dataUrlToBlob_(dataUrl, filename) {
  if (typeof dataUrl !== 'string') throw new Error('Invalid dataUrl');
  dataUrl = dataUrl.replace(/\s/g, '');
  const m = dataUrl.match(/^data:([^;]*);base64,([A-Za-z0-9+/=]+)$/);
  if (!m) throw new Error('Invalid data URL format');
  const mime = m[1] || 'application/octet-stream';
  const base64 = m[2];
  const approxBytes = Math.floor(base64.length * 0.75);
  if (approxBytes > MAX_FILE_BYTES) throw new Error('ファイルサイズ上限(15MB)を超えています。');
  const bytes = Utilities.base64Decode(base64);
  return Utilities.newBlob(bytes, mime, filename);
}

/** YYYY-MM-DD or YYYY/MM/DD or Date -> YYYY/MM/DD */
function normalizeDateYmd_(v) {
  if (!v) return '';
  try {
    let d;
    if (v instanceof Date) {
      d = v;
    } else if (/^\d{4}-\d{2}-\d{2}$/.test(v)) {
      const [y,m,day] = v.split('-').map(Number);
      d = new Date(y, m-1, day);
    } else if (/^\d{4}\/\d{2}\/\d{2}$/.test(v)) {
      const [y,m,day] = v.split('/').map(Number);
      d = new Date(y, m-1, day);
    } else {
      d = new Date(v);
    }
    return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
  } catch (_) {
    return '';
  }
}

/** 本日（JST）の日付のみの Date */
function todayJst_() {
  const [y, m, d] = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd').split('/').map(Number);
  return new Date(y, m - 1, d);
}

/** 回答者への控えメール送信 */
function sendCopy_(to, subject, body) {
  var method = (COPY_MAIL_METHOD || 'auto').toLowerCase();

  if (method === 'gmail') { GmailApp.sendEmail(to, subject, body); return; }
  if (method === 'mail')  { MailApp.sendEmail(to, subject, body);  return; }

  try {
    GmailApp.sendEmail(to, subject, body);
  } catch (e1) {
    try { MailApp.sendEmail(to, subject, body); }
    catch (e2) {
      throw new Error('控えメール送信に失敗しました（GmailApp/MailApp いずれも不可）: '
        + (e1 && e1.message ? e1.message : e1) + ' / ' + (e2 && e2.message ? e2.message : e2));
    }
  }
}

function safe_(s){ return String(s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }

/** 接続テスト */
function ping() { return 'pong'; }
