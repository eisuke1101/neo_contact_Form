/** ========== 不具合報告 ========== */
/*
 * 書き込み先：NEO V2 UAT フィードバック管理表 ＞ フィードバック一覧
 *   A: ID（最大値+1） / B: ステータス「新規 / New」 / C: HT Redmine「N/A」
 *   D: Tracker（不具合→Bug / 改善要望→CR） / E: タイムスタンプ / F: メールアドレス
 *   G: カテゴリ / H: 深刻度・優先度 / I: 画面のURL / J: 操作手順
 *   K: 問題の動作 / L: あるべき動作 / M: スクリーンショット・録画（リンク、改行区切り）
 *   N: アカウント名 / O: パスワード
 *   P: フロー番号・シナリオID、Q: TC番号 は入力不要のため空欄。R列以降は書き込まない。
 *
 * G・H・B 列は「リストを直接指定」の入力規則（不正な値を拒否）があるため、
 * 下記の値はシートの選択肢と完全一致させること。
 */
const BUG_STATUS_NEW = '新規 / New';
const BUG_REDMINE_DEFAULT = 'N/A';
const BUG_CATEGORIES = {
  '不具合 / Bug report':      'Bug',
  '改善要望 / Change request': 'CR'
};
const BUG_PRIORITIES = [
  'P1: テストを進めることができない深刻な問題 / Critical',
  'P2: 業務に支障がある問題（ロジック・入出力の間違いなど） / Important',
  'P3: 業務に支障がない問題（推測可能な誤字など） / Medium',
  'P4: 問題ではないが、あったらよい改善要望・アイデア / Low'
];

function submitBug(payload) {
  try {
    if (!FORMS_ENABLED.bug) throw new Error('この受付は現在準備中です。');
    if (!payload) throw new Error('no payload');
    const req = payload.fields || {};
    const v = k => String(req[k] || '').trim();

    const email    = v('email');
    const category = v('category');
    const priority = v('priority');
    const steps    = v('steps');
    const actual   = v('actual');
    const expected = v('expected');
    const links    = v('extraLinks');

    if (!email)                                throw new Error('メールアドレスは必須です。');
    if (!(category in BUG_CATEGORIES))         throw new Error('カテゴリを選択してください。');
    if (BUG_PRIORITIES.indexOf(priority) < 0)  throw new Error('深刻度・優先度を選択してください。');
    if (!steps)                                throw new Error('問題に至るまでの操作手順は必須です。');
    if (!actual)                               throw new Error('問題の動作は必須です。');
    if (!expected)                             throw new Error('あるべき動作は必須です。');

    const files = payload.files || [];
    checkFileSizes_(files);

    const ws = getSheet_(BUG_SPREADSHEET_ID, BUG_SHEET_NAME);
    const folder = ensureFolder_(BUG_DRIVE_FOLDER_ID, PROP_BUG_FOLDER_ID, AUTO_FOLDER_BUG_NAME);

    // 先にファイルを保存（失敗したらシートには何も書かない）
    const fileIds = saveFiles_(files, folder);
    const captures = fileIds.map(id => 'https://drive.google.com/open?id=' + id)
      .concat(links ? links.split(/\s+/).filter(Boolean) : [])
      .join('\n');

    const now = new Date();
    const id = withScriptLock_(() => {
      const id = nextNumericId_(ws, 1);
      const rowIndex = ws.getLastRow() + 1;
      ws.getRange(rowIndex, 1, 1, 17).setValues([[
        id,
        BUG_STATUS_NEW,
        BUG_REDMINE_DEFAULT,
        BUG_CATEGORIES[category],
        now,
        email,
        category,
        priority,
        asSheetText_(v('url')),
        asSheetText_(steps),
        asSheetText_(actual),
        asSheetText_(expected),
        captures,
        asSheetText_(v('username')),
        asSheetText_(v('password')),
        '', // フロー番号・シナリオID（不要）
        ''  // TC番号（不要）
      ]]);
      return id;
    });

    enqueueJob_({
      type: 'bug',
      id,
      fileIds,
      fields: {
        email, category, priority,
        url: v('url'),
        timestamp: Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
        steps: truncate_(steps, 600),
        actual: truncate_(actual, 600),
        expected: truncate_(expected, 600),
        captures
      }
    });
    try { processJobs_(); } catch (e) { console.warn('inline processJobs_ failed:', e); }
    if (listJobKeys_().length) ensureOneShotTrigger_();

    return { ok: true, id: `BUG-${id}`, message: `受付番号 BUG-${id} で受け付けました。` };

  } catch (e) {
    return { ok: false, error: (e && e.message) || String(e) };
  }
}

/** 後段処理（不具合報告）：公開化・控えメール（パスワードはメールに載せない） */
function runBugJob_(job) {
  const { id, fileIds, fields } = job;
  makeFilesPublic_(fileIds);

  if (SEND_COPY_TO_RESPONDENT && fields.email && !job.mailSent) {
    const subject = `【BUG-${id}】不具合報告を受け付けました`;
    const body =
`お疲れ様です。

以下の内容で不具合報告を受け付けました。
確認のうえ、順次対応いたします。

◆ ID: ${id}
◆ 受付日時: ${fields.timestamp}
◆ カテゴリ: ${fields.category}
◆ 深刻度・優先度: ${fields.priority}
◆ 画面のURL: ${fields.url}
◆ 問題に至るまでの操作手順:
${fields.steps}
◆ 問題の動作:
${fields.actual}
◆ あるべき動作:
${fields.expected}
◆ スクリーンショット・録画など:
${fields.captures || '（なし）'}

何卒よろしくお願いいたします。

システム課`;
    sendCopy_(fields.email, subject, body);
    job.mailSent = true;
  }
}
