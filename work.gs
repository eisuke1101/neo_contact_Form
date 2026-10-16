/** ========== 作業依頼（既存フォーム） ========== */

/** 送信本体（軽量化：事前検証→最低限の保存→後段をキューで処理） */
function submit(payload) {
  try {
    if (!payload) throw new Error('no payload');

    // --- 必須チェック ---
    const req = payload.fields || {};
    if (!req.email)                       throw new Error('メールアドレスは必須です。');
    if (!req.schoolDept)                  throw new Error('校舎名・部署名は必須です。');
    if (!req.person)                      throw new Error('ご担当者様氏名は必須です。');
    if (!req.topic)                       throw new Error('お問い合わせ内容①は必須です。');
    if (!(req.detail || req.detailHtml))  throw new Error('お問い合わせ内容②は必須です。');

    // === 添付ファイルの事前検査（ここでNGなら以降の保存は一切しない） ===
    prevalidateFiles_(payload.files || []);

    // --- 保存先準備 ---
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ws = SHEET_NAME
      ? (ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME))
      : ss.getSheets()[0];
    if (!ws) throw new Error(`シートが見つかりません: ${SHEET_NAME}`);

    // アップロード先フォルダ
    const folderHtml  = ensureFolder_(DRIVE_FOLDER_HTML_ID,  PROP_HTML_FOLDER_ID,  AUTO_FOLDER_HTML_NAME);
    const folderFiles = ensureFolder_(DRIVE_FOLDER_FILES_ID, PROP_FILES_FOLDER_ID, AUTO_FOLDER_FILES_NAME);

    // --- 本文（HTML）を Drive に保存（プレビュー用） ---
    const bodyHtml = buildPreviewHtml_(req);
    const htmlName = `detail-${Date.now()}.html`;
    const htmlFile = folderHtml.createFile(htmlName, bodyHtml, MimeType.HTML);
    const previewUrl = `${BASE_WEBAPP_URL}?mode=preview&id=${htmlFile.getId()}`;

    // --- 期日（YYYY/MM/DD）整形 ---
    const dueDateStr = normalizeDateYmd_(req.dueDate); // 空なら空文字

    // --- シート書き込み ---
    const submittedTopic = req.topic === 'その他' ? `その他: ${req.topicOther || ''}` : req.topic;
    const now = new Date();
    const asText = v => (v === undefined || v === null || String(v) === '') ? '' : "'" + String(v);

    // 列マッピング（列9=プレビューURL）
    // 1: Timestamp / 2: Email / 3: スクール・部署名 / 4: スクール・部署コード
    // 5: ご担当者様氏名 / 6: 電話番号 / 7: 問い合わせ区分 / 8: 問い合わせ内容（プレーン）
    // 9: プレビューURL / 10: 対応希望期日(YYYY/MM/DD) / 11: 期日(補足) / 12: その他ファイルURL
    const row = [
      now,
      req.email,
      req.schoolDept,
      asText(req.schoolDeptId),
      req.person,
      asText(req.phone || ''),
      submittedTopic,
      (req.detail || '').trim(),
      previewUrl,
      dueDateStr,
      req.dueDateNote || '',
      '' // 添付URLは後段で更新
    ];

    // 同時送信で行番号を取り違えないよう、追記と行番号取得をロック内で行う
    const rowIndex = withScriptLock_(() => {
      ws.appendRow(row);
      return ws.getLastRow();
    });

    // 送信直後に「処理中…」を明示（空欄残しを防止）
    if ((payload.files || []).length > 0) {
      try { ws.getRange(rowIndex, 12).setValue('（添付URL処理中…）'); } catch (_) {}
    }

    // --- 添付ファイルは作成のみ（公開設定やURL整備は後段） ---
    const savedFileIds = saveFiles_(payload.files || [], folderFiles);

    // --- 後段処理キュー（公開設定・URL整備・控えメール送信） ---
    enqueueJob_({
      type: 'work',
      rowIndex,
      htmlFileId: htmlFile.getId(),
      fileIds: savedFileIds,
      fields: {
        email: req.email,
        schoolDept: req.schoolDept,
        schoolDeptId: req.schoolDeptId || '',
        person: req.person,
        phone: req.phone || '',
        topic: submittedTopic,
        detail: truncate_((req.detail || '').trim(), 1500),
        dueDate: dueDateStr,
        dueDateNote: req.dueDateNote || ''
      }
    });

    // ベストエフォートで即時実行（戻り値は変えず、失敗しても無視）
    try { processJobs_(); } catch (e) { console.warn('inline processJobs_ failed:', e); }

    // 未処理分があれば1秒後に再実行
    if (listJobKeys_().length) ensureOneShotTrigger_();

    return { ok: true, rowIndex, message: '受付しました（添付URL等は数秒後に反映されます）' };

  } catch (e) {
    return { ok: false, error: (e && e.message) || String(e) };
  }
}

/** 後段処理（作業依頼）：公開化・添付URL反映・控えメール */
function runWorkJob_(job) {
  const { rowIndex, htmlFileId, fileIds, fields } = job;
  const ws = getSheet_(SPREADSHEET_ID, SHEET_NAME);

  makeFilesPublic_([htmlFileId].concat(fileIds || []));

  // 添付URL作成してシート反映
  const fileUrls = (fileIds || []).map(id => 'https://drive.google.com/uc?id=' + id);
  ws.getRange(rowIndex, 12).setValue(fileUrls.join(', '));

  // 控えメール送信（設定有効時）
  if (SEND_COPY_TO_RESPONDENT && fields.email && !job.mailSent) {
    const subject = `【NEO-${rowIndex}】お問い合わせを受け付けました`;
    const body =
`${fields.schoolDept}　　${fields.person}　様

お疲れ様です。

以下の内容でお問い合わせを受け付けました。
対応完了まで今しばらくお待ちください。

◆ スクール・部署コード: ${fields.schoolDeptId}
◆ スクール・部署名: ${fields.schoolDept}
◆ ご担当者様氏名: ${fields.person}
◆ 電話番号: ${fields.phone}
◆ お問い合わせ内容①: ${fields.topic}
◆ お問い合わせ内容②:
   ${fields.detail}
◆ 対応希望期日: ${fields.dueDate}
◆ 補足: ${fields.dueDateNote}

◆ プレビュー: ${BASE_WEBAPP_URL}?mode=preview&id=${htmlFileId}

何卒よろしくお願いいたします。

システム課`;
    sendCopy_(fields.email, subject, body);
    job.mailSent = true;
  }
}

/** 本文プレビュー用のHTML（本文の見出しなど最低限の体裁を付与） */
function buildPreviewHtml_(req) {
  const title = 'お問い合わせ内容';
  const html = (req.detailHtml || '').trim();
  const plain = (req.detail || '').trim();

  const body = html || ('<pre style="white-space:pre-wrap;">' + safe_(plain) + '</pre>');

  return `
    <div style="font-family:system-ui,Segoe UI,Roboto,Arial,sans-serif;line-height:1.7;">
      <h3 style="margin:0 0 12px;">${title}</h3>
      <div>${body}</div>
    </div>
  `;
}
