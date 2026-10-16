/** ========== 新規開発・仕様変更依頼 ========== */
/*
 * 書き込み先：【リスト】システム開発依頼一覧 ＞ 小規模案件
 *   A: 管理No.（最大値+1） / B: 部署 / C: 担当 / D: 日付（受付日）
 *   E: 詳細 / F: 資料（アップロードしたファイルのリンク）
 *   G〜J（状況・完了日・備考・通知）は空欄のまま。
 *   ※ J列「通知」が空欄の行を既存の Chatwork 通知処理が拾う前提。
 */

function submitDev(payload) {
  try {
    if (!FORMS_ENABLED.dev) throw new Error('この受付は現在準備中です。');
    if (!payload) throw new Error('no payload');
    const req = payload.fields || {};

    const email  = String(req.email || '').trim();
    const person = String(req.person || '').trim();
    const dept   = String(req.dept || '').trim();
    const detail = String(req.detail || '').trim();
    const links  = String(req.extraLinks || '').trim();

    if (!email)  throw new Error('メールアドレスは必須です。');
    if (!person) throw new Error('ご担当者様氏名は必須です。');
    if (!dept)   throw new Error('部署は必須です。');
    if (!detail) throw new Error('詳細は必須です。');

    const files = payload.files || [];
    checkFileSizes_(files);

    const ws = getSheet_(DEV_SPREADSHEET_ID, DEV_SHEET_NAME);
    const folder = ensureFolder_(DEV_DRIVE_FOLDER_ID, PROP_DEV_FOLDER_ID, AUTO_FOLDER_DEV_NAME);

    // 先にファイルを保存（失敗したらシートには何も書かない）
    const fileIds = saveFiles_(files, folder);
    const fileUrls = fileIds.map(id => `https://drive.google.com/file/d/${id}/view`);
    const materials = fileUrls.concat(links ? links.split(/\s+/).filter(Boolean) : []).join('\n');

    const today = todayJst_();
    const { no } = withScriptLock_(() => {
      const no = nextNumericId_(ws, 1);
      const rowIndex = ws.getLastRow() + 1;
      ws.getRange(rowIndex, 1, 1, 6).setValues([[
        no,
        asSheetText_(dept),
        asSheetText_(person),
        today,
        asSheetText_(detail),
        materials
      ]]);
      return { no, rowIndex };
    });

    enqueueJob_({
      type: 'dev',
      no,
      fileIds,
      fields: {
        email, person, dept,
        date: Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd'),
        detail: truncate_(detail, 1500),
        materials
      }
    });
    try { processJobs_(); } catch (e) { console.warn('inline processJobs_ failed:', e); }
    if (listJobKeys_().length) ensureOneShotTrigger_();

    return { ok: true, id: `DEV-${no}`, message: `受付番号 DEV-${no} で受け付けました。` };

  } catch (e) {
    return { ok: false, error: (e && e.message) || String(e) };
  }
}

/** 後段処理（開発依頼）：公開化・控えメール */
function runDevJob_(job) {
  const { no, fileIds, fields } = job;
  makeFilesPublic_(fileIds);

  if (SEND_COPY_TO_RESPONDENT && fields.email && !job.mailSent) {
    const subject = `【DEV-${no}】新規開発・仕様変更のご依頼を受け付けました`;
    const body =
`${fields.dept}　　${fields.person}　様

お疲れ様です。

以下の内容で新規開発・仕様変更のご依頼を受け付けました。
内容を確認のうえ、担当よりご連絡いたします。

◆ 管理No.: ${no}
◆ 部署: ${fields.dept}
◆ 担当: ${fields.person}
◆ 日付: ${fields.date}
◆ 詳細:
${fields.detail}
◆ 資料:
${fields.materials || '（なし）'}

何卒よろしくお願いいたします。

システム課`;
    sendCopy_(fields.email, subject, body);
    job.mailSent = true;
  }
}
