/**
 * AI Pulse Webhook (Google Apps Script)
 * KONNEKT INTERNATIONAL 社員のAI活用度をスプシで管理＋Gemini APIでAIコメント生成
 *
 * セットアップ手順は `00_🏢 company/ai/20260519_AI効果数値化/GAS構築手順_v1.md` を参照
 *
 * スクリプトプロパティ必須:
 *   - SHEET_ID         : スプレッドシートID
 *   - GEMINI_API_KEY   : Gemini APIキー（無料・https://aistudio.google.com/apikey で発行）
 *   - ADMIN_PASSWORD   : 管理画面パスワード（既定: admin2026）
 *   - EXEC_PASSWORD    : 経営層パスワード（既定: exec2026）
 *   - MODEL            : Geminiモデル（既定: gemini-2.0-flash）
 */

const PROPS = PropertiesService.getScriptProperties();
const SHEET_ID = PROPS.getProperty('SHEET_ID');
const GEMINI_API_KEY = PROPS.getProperty('GEMINI_API_KEY');
const ADMIN_PW = PROPS.getProperty('ADMIN_PASSWORD') || 'admin2026';
const EXEC_PW = PROPS.getProperty('EXEC_PASSWORD') || 'exec2026';
const MODEL = PROPS.getProperty('MODEL') || 'gemini-2.0-flash';

// ============ ルーティング ============
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || 'get_master';
  try {
    if (action === 'get_master') return jsonOut(getMaster());
    if (action === 'get_dashboard') return jsonOut(getDashboard());
    return jsonOut({ error: 'unknown action: ' + action });
  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const action = data.action;
    if (action === 'submit_response') return jsonOut(submitResponse(data));
    if (action === 'generate_cat_comment') return jsonOut(generateCatComment(data));
    if (action === 'generate_recommendation') return jsonOut(generateRecommendation(data));
    if (action === 'generate_exec_analysis') return jsonOut(generateExecAnalysis(data));
    if (action === 'generate_next_actions') return jsonOut(generateNextActions(data));
    if (action === 'generate_monthly_digest') return jsonOut(generateMonthlyDigest(data));
    if (action === 'send_to_slack') return jsonOut(sendToSlack(data));
    if (action === 'update_master') return jsonOut(updateMaster(data));
    if (action === 'update_pin') return jsonOut(updatePin(data));
    return jsonOut({ error: 'unknown action: ' + action });
  } catch (err) {
    return jsonOut({ error: err.message, stack: err.stack });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============ マスター取得 ============
function getMaster() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const namesSheet = ss.getSheetByName('名簿');
  const initSheet = ss.getSheetByName('施策');
  const settingsSheet = ss.getSheetByName('設定');

  const names = sheetToObjects(namesSheet, ['name', 'department', 'email', 'active', 'pin']);
  const initiatives = sheetToObjects(initSheet, ['id', 'name', 'month', 'detail', 'active']);
  const settings = {};
  if (settingsSheet) {
    const rows = settingsSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0]) settings[rows[i][0]] = rows[i][1];
    }
  }
  const departments = [...new Set(names.map(n => n.department).filter(Boolean))];

  // 統合「選択肢」タブから全質問の選択肢を読み取り、質問キーごとに分類
  const choicesSheet = ss.getSheetByName('選択肢');
  const choices = {};
  if (choicesSheet) {
    const rows = choicesSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      const key = String(r[0]);
      const item = {
        value: r[1] || '',
        emoji: r[2] || '',
        label: r[3] || '',
        detail: r[4] || '',
        score: parseFloat(r[5]) || 0,
        order: parseFloat(r[6]) || 0,
        active: (r[7] === true || r[7] === 'TRUE' || r[7] === 'true' || r[7] === 1)
      };
      if (!choices[key]) choices[key] = [];
      choices[key].push(item);
    }
    // 並び順でソート
    Object.keys(choices).forEach(k => {
      choices[k].sort((a, b) => a.order - b.order);
    });
  }

  return { names, initiatives, settings, departments, choices };
}

function sheetToObjects(sheet, keys) {
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  return rows.slice(1).filter(r => r[0]).map(r => {
    const obj = {};
    keys.forEach((k, i) => {
      let v = r[i];
      if (k === 'active') v = (v === true || v === 'TRUE' || v === 'true' || v === 1);
      obj[k] = v;
    });
    return obj;
  });
}

// ============ 回答送信 ============
function submitResponse(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('回答');
  if (!sheet) throw new Error('「回答」タブが見つかりません');

  const score = data.score || calcScore(data);
  const month = data.month || (new Date()).toISOString().slice(0, 7);
  // 新列構造：質問の備考は該当質問の右隣・経営判断材料を明確化
  const row = [
    new Date().toISOString(),                   // A タイムスタンプ
    month,                                      // B 対象月
    data.name || '',                            // C 名前
    data.department || '',                      // D 部署
    data.q1 || '',                              // E Q1 使用頻度
    JSON.stringify(data.q2 || []),              // F Q2 場面
    data.q3 || 0,                               // G Q3 効率化%
    data.q4 || '',                              // H Q4 成長実感
    data.q4Note || '',                          // I Q4備考
    data.q5 || '',                              // J Q5 方向性
    data.q6 || '',                              // K Q6 共有
    data.q6Note || '',                          // L Q6備考
    data.q7 || '',                              // M Q7 ベスプラ
    JSON.stringify(data.initiatives || []),     // N 施策実施
    JSON.stringify(data.didableTags || []),     // O できたことタグ
    data.didable || '',                         // P できたこと（自由記述）
    data.troubleLevel || '',                    // Q 困り度
    data.troubleNote || '',                     // R 困り度の具体
    data.comment || '',                         // S コメント
    score                                       // T スコア(XP)
  ];
  sheet.appendRow(row);
  // B列「対象月」をDate型自動変換から守るため、書式を文字列に強制
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 2).setNumberFormat('@').setValue(month);
  return { ok: true, score };
}

function calcScore(ans) {
  // 月最大30XPの骨太モデル
  let xp = 0;
  xp += { 'ほぼ毎日': 10, '週2-3日': 5, '週1日': 2, 'ほぼ使ってない': 0 }[ans.q1] || 0;
  xp += Math.min(6, (ans.q2 || []).length);
  xp += Math.floor((ans.q3 || 0) / 14);
  xp += { '😫': 0, '😐': 1, '😊': 2, '🔥': 3 }[ans.q4] || 0;
  if (ans.q5) xp += (ans.q5 === '現状維持') ? 1 : 2;
  xp += { '😫': 0, '😐': 1, '😊': 2, '🔥': 2 }[ans.q6] || 0;
  return Math.max(0, Math.min(30, xp));
}

// ============ ダッシュボード取得 ============
function getDashboard() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('回答');
  if (!sheet) return { responses: [] };
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { responses: [] };

  const responses = rows.slice(1).filter(r => r[0]).map(r => {
    const ts = new Date(r[0]);
    // 対象月（B列）はDate型自動変換されてる可能性があるため正規化
    let month;
    if (r[1] instanceof Date) {
      month = Utilities.formatDate(r[1], 'JST', 'yyyy-MM');
    } else if (r[1]) {
      const s = String(r[1]);
      month = /^\d{4}-\d{2}/.test(s) ? s.slice(0, 7) : Utilities.formatDate(ts, 'JST', 'yyyy-MM');
    } else {
      month = Utilities.formatDate(ts, 'JST', 'yyyy-MM');
    }
    return {
      timestamp: ts.toISOString(),
      month: month,
      name: r[2],
      department: r[3],
      q1: r[4],
      q2: tryParseJSON(r[5]) || [],
      q3: parseFloat(r[6]) || 0,
      q4: r[7],
      q4Note: r[8] || '',
      q5: r[9],
      q6: r[10],
      q6Note: r[11] || '',
      q7: r[12],
      initiatives: tryParseJSON(r[13]) || [],
      didableTags: tryParseJSON(r[14]) || [],
      didable: r[15] || '',
      troubleLevel: r[16] || '',
      troubleNote: r[17] || '',
      comment: r[18] || '',
      score: parseInt(r[19]) || 0
    };
  });
  return { responses };
}

function tryParseJSON(s) {
  try { return JSON.parse(s); } catch { return null; }
}

// ============ 猫キャラコメント生成（Gemini API） ============
function generateCatComment(data) {
  const { name, score, cumulativeXp, currentMonth, currentAnswers } = data;
  if (!GEMINI_API_KEY) return { comment: '今月もおつかれにゃ！' };

  // 当人の今月以前のデータを全部取得（昇順）
  const dashboard = getDashboard();
  const myHist = dashboard.responses
    .filter(r => r.name === name)
    .sort((a, b) => a.month.localeCompare(b.month));

  // 最新の保存済みエントリ（既に今月分が保存されてる場合はそれの1個前、まだなら最新を「先月」とみなす）
  const lastEntry = myHist.length > 0
    ? (myHist[myHist.length - 1].month === currentMonth
        ? (myHist.length >= 2 ? myHist[myHist.length - 2] : null)
        : myHist[myHist.length - 1])
    : null;

  const cur = currentAnswers || {};
  const summarize = (r) => {
    if (!r) return '(なし)';
    return [
      `使用頻度=${r.q1 || '-'}`,
      `場面=${(r.q2 || []).join('・') || '-'}`,
      `効率化=${r.q3 || 0}%`,
      `成長実感=${r.q4 || '-'}${r.q4Note ? `(${r.q4Note})` : ''}`,
      `来月の方向=${r.q5 || '-'}`,
      `共有=${r.q6 || '-'}${r.q6Note ? `(${r.q6Note})` : ''}`,
      `ベスプラ=${r.q7 || 'なし'}`,
      `できたこと=${r.didable || 'なし'}`
    ].join(' / ');
  };

  const prompt = `あなたは社員のAI活用を応援する猫キャラ「にゃんこ」です。
${name}さんの「先月→今月」の変化を見て、振り返り＋来月のアドバイスを3-4行で返してください。

【今月（${currentMonth}）】月次+${score}XP / 累積${cumulativeXp}XP
${summarize(cur)}

【先月】
${lastEntry ? `(${lastEntry.month}) ${summarize(lastEntry)}` : '初回の回答なのでデータなし'}

ルール:
- 語尾は「にゃ」
- 先月データがあれば、必ず具体的な変化に触れる（例:「先月は週1だったのが今月はほぼ毎日になってるにゃ！」「先月の効率化15%から、今月は40%に上がったにゃ✨」「先月『議事録』だけだったのが、今月は『企画書』にも広がってるにゃ」）
- 来月の具体行動を1つ提案
- 3-4行・絵文字1〜2個OK・読みやすく改行
- 励まし口調・厳しすぎないで`;

  const comment = callGemini(prompt, 500);
  return { comment };
}

// ============ 個別レコメンド生成（次の一歩） ============
function generateRecommendation(data) {
  const { name, currentAnswers } = data;
  if (!GEMINI_API_KEY) {
    return { recommendation: '来月もこの調子で続けてみるにゃ！新しい使い方も1つ試してみてね💡' };
  }

  // 個人の過去回答全部取得（昇順）
  const dashboard = getDashboard();
  const myHist = dashboard.responses
    .filter(r => r.name === name)
    .sort((a, b) => a.month.localeCompare(b.month));

  // 社内アクティブ施策
  const master = getMaster();
  const activeInitiatives = (master.initiatives || []).filter(i => i.active);

  const cur = currentAnswers || {};

  // 履歴サマリー
  const histSummary = myHist.slice(-6).map(r =>
    `${r.month}: 頻度=${r.q1 || '-'}/効率化=${r.q3 || 0}%/場面=${(r.q2 || []).join(',') || '-'}/できた=${(r.didableTags || []).join(',') || '-'}`
  ).join('\n');

  const prompt = `あなたはKONNEKT INTERNATIONALのAI推進アシスタントです。
${name}さんの活用パターンから、来月試すべき具体的アクションを1つ提案してください。

【今月の状況】
- 使用頻度: ${cur.q1 || '-'}
- 使ってる場面: ${(cur.q2 || []).join('・') || '-'}
- 効率化%: ${cur.q3 || 0}%
- 成長実感: ${cur.q4 || '-'}${cur.q4Note ? ` (${cur.q4Note})` : ''}
- 来月やりたい方向: ${cur.q5 || '-'}
- チーム共有: ${cur.q6 || '-'}${cur.q6Note ? ` (${cur.q6Note})` : ''}
- できるようになった: ${(cur.didableTags || []).join('・') || '-'}${cur.didable ? ` (${cur.didable})` : ''}
- 困り度: ${cur.troubleLevel || '-'}${cur.troubleNote ? ` (${cur.troubleNote})` : ''}

【過去${myHist.length}回の履歴】
${histSummary || '初回回答'}

【今月の社内施策候補】
${activeInitiatives.map(i => `- ${i.name}: ${i.detail || ''}`).join('\n') || '施策なし'}

【出力ルール（厳守・必ず全部入れる）】
1. **具体的なツール名を必ず1つ以上**入れる：Claude / ChatGPT / Cowork / NotebookLM / Slack MCP / Claude Projects / CircleBack 等
2. **「何をすると何ができるか」の動作レベル**で書く（抽象的な「新しい使い方を試そう」は禁止）
3. **プロンプト例を1つ必ず含める**（「" 〜 "」の引用形式で・実際に貼って使える形）
4. **所要時間を明示**（例：「最初の試行5分」「準備10分・運用1分」）
5. **期待効果を数字で**（例：「議事録30分→3分」「メール5分→30秒」）
6. 既に使いこなしてる領域は除外
7. 困り度が高い人（😟😣）はまず「困り解消」を最優先
8. ${name}さんの「来月やりたい方向」（${cur.q5 || '不明'}）に合わせる
9. 絵文字1-2個OK・フレンドリー口調・敬語不要

出力フォーマット（4-6行・各行で必要情報を埋める）：
💡 来月のおすすめ：[見出し1行]
[何をすると何ができるか・期待効果を数字で示す 1-2行]
プロンプト例："[実際に貼って使える具体プロンプト]"
[所要時間・気軽さの動機づけを1行]

出力例:
💡 来月のおすすめ：Claude Projectsで「企画書専用」スペースを作る
今月企画書がAIで作れるようになったね！次は「Claude Projects」で企画書専用スペース作って、過去の良い企画書3本を添付すると、文体まで学習してくれて初稿の精度が爆上がり。企画書作成30分→5分に。
プロンプト例："過去のAM 26SS企画書のトンマナで、kemio抹茶26AWの企画書たたき台を3案"
準備10分・運用1分。試しに金曜夕方やってみない？✨`;

  const rawText = callGemini(prompt, 800);
  // 空チェック・JSONじゃないただのテキスト想定（既存仕様）
  const cleaned = (rawText || '').trim();
  if (!cleaned || cleaned.length < 20) {
    Logger.log('Recommendation empty: ' + rawText);
    return {
      recommendation: `💡 来月のおすすめ：${name}さんの今月の活用パターンを伸ばす\n今月の使い方をベースに、Claude/Coworkで次の業務にもAIを広げてみよう。具体的にはメール返信や議事録要約のような短時間タスクで「これも頼める」と気づくのがコツ。\nプロンプト例："このタスクをAIに頼むとしたら、どんな指示文がベスト？"\n所要時間5分。気軽に試してみてね✨`
    };
  }
  return { recommendation: cleaned };
}

// ============ 経営層：総合サマリー＋分析＋打ち手（JSON出力で確実化） ============
function generateExecAnalysis(data) {
  if (!GEMINI_API_KEY) return { summary: '（APIキー未設定）', analysis: '（APIキー未設定）', nextActions: '（APIキー未設定）' };
  const { currentMonth, totalSaved, lastSaved, heavyRate, respRate, deptStats, initStats, troubled, unanswered, masters, trend, targets } = data;

  const prompt = `あなたはKONNEKT INTERNATIONAL の経営アドバイザー兼AI推進顧問です。
社長（由羽さん）への月次報告書として、「総合サマリー（経営報告書レベル長文）」＋「分析（簡潔）」＋「来月の打ち手」の3つを生成してください。

【KONNEKT のAI推進の経緯・ミッション】
- 社長（由羽さん）からAI推進チーム（藤本ゆりこ・引地瑞生）への直接ミッション
- 目標：「AI活用による粗利向上・作業時間削減を可視化」「全社員（18名）がAIを日常業務で使いこなせる状態にする」
- AI推進フェーズ：Phase 0（組織設計）→ Phase 1（基盤整備）→ Phase 2（ツール連携：Slack/GWS/Shopify）→ Phase 3（業務AI化）→ Phase 4（定着・KPI測定）
- 現在はPhase 4の入口・このAI Pulseアプリで月次KPI測定を開始した段階
- 旧Forms設計→アプリ化で運用工数削減＋データ蓄積を実現

【今月（${currentMonth || '今月'}）のデータ】
- 合計削減時間: ${totalSaved}h / 目標 ${targets?.savedH || 100}h
- AI浸透率: ${heavyRate}% / 目標 ${targets?.heavyRate || 70}%
- 回答率: ${respRate}%
- 困ってる人: ${(troubled || []).join('・') || 'なし'}
- 未回答者: ${(unanswered || []).join('・') || 'なし'}
- AI伝説/マスター達成者: ${(masters || []).join('・') || 'なし'}
- 先月比: ${(totalSaved - lastSaved).toFixed(1)}h
- 部署別: ${(deptStats || []).map(d => `${d.dept}:${d.saved.toFixed(1)}h(${d.count}名)`).join(', ')}
- 施策実施率: ${(initStats || []).map(i => `${i.name}:${i.rate}%`).join(', ')}
${trend ? `- 3ヶ月推移: ${trend}` : ''}

【出力フォーマット（厳守・必ずJSON形式のみ）】
以下のJSON形式だけを返してください。前後に説明文・コードブロック記号は不要。

{
  "summary": "総合サマリー本文（8-12行の長文経営報告書レベル）",
  "analysis": "分析テキスト（3-4行・要点）",
  "nextActions": "1. アクション1\\n2. アクション2\\n3. アクション3"
}

【summary（総合サマリー）の作り方】
- 単なる事実報告ではなく、AI推進チーム（藤本ゆりこ・引地瑞生）に **寄り添う経営アドバイザー** の視点で書く
- 「今月のAI活用は○○な月だった」というナラティブで始める
- KONNEKTのAI推進フェーズ（現在Phase 4=定着・KPI測定）に対する今月の位置付けを明示
- 目標vs実績の達成率を必ず触れる（削減h・浸透率・新規創出）
- 数字の意味を読み解いた上で「**だからこそ来月はこうしよう**」「ここに気をつけよう」を必ず含める
- ROIの観点で投資対効果を経営判断材料として提示
- AI推進チームの努力を労う一文も入れる（経営層視点）
- 来月への具体的な方向性・推進チームへの提案を1-2個示す
- **段落間は \\n 1個だけ（\\n\\n は使わない）**・各段落2-3行・読みやすく
- 社長報告として完成度の高い文章（敬語・丁寧）・全体8-12行

【analysis（分析）の作り方】
- 数字の動き・良かったこと・課題を箇条書き風に
- 数字を必ず入れる
- 短く要点だけ・3-4行
- 改行は\\n

【nextActions（来月の打ち手）の作り方】
- 番号付き3つ
- 「誰に・何を・いつ・どうやって」を明確に
- 困ってる人への個別サポートを必ず1つ含める
- 各1-2行で簡潔に・経営層に「動いてる」と思わせる
- ツール名・所要時間を明示
- 改行は\\n

【重要】
- 純粋なJSON文字列のみ返す
- 余計な説明・前置き・コードブロック記号（\`\`\`）禁止
- summary・analysis・nextActions の3キー必須`;

  const fullText = callGemini(prompt, 4000);

  // JSON抽出（コードブロック記号や前置きを除去）
  let cleaned = fullText.trim();
  // ```json ... ``` を取り除く
  cleaned = cleaned.replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/i, '').trim();
  // 最初の { から最後の } までを抽出
  const firstBrace = cleaned.indexOf('{');
  const lastBrace = cleaned.lastIndexOf('}');
  if (firstBrace >= 0 && lastBrace > firstBrace) {
    cleaned = cleaned.substring(firstBrace, lastBrace + 1);
  }

  try {
    const obj = JSON.parse(cleaned);
    return {
      summary: obj.summary || '',
      analysis: obj.analysis || '',
      nextActions: obj.nextActions || ''
    };
  } catch (e) {
    // JSONパース失敗時：原文を summary に丸ごと入れる（少しでも何か表示）
    Logger.log('JSON parse failed: ' + e.message + ' | raw: ' + fullText.substring(0, 500));
    return {
      summary: fullText,
      analysis: '（JSON出力フォーマットエラー・summaryに原文を表示中）',
      nextActions: '（再生成してください）'
    };
  }
}

// ============ Gemini API 呼び出し ============
function callGemini(prompt, maxTokens) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent?key=${GEMINI_API_KEY}`;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      maxOutputTokens: maxTokens || 500,
      temperature: 0.7
    }
  };
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const body = JSON.parse(res.getContentText());
  if (body.error) throw new Error(body.error.message || JSON.stringify(body.error));
  if (!body.candidates || !body.candidates[0]) {
    throw new Error('Geminiレスポンスが空です: ' + JSON.stringify(body));
  }
  return body.candidates[0].content.parts[0].text;
}

// ============ 来月の打ち手（AI推進チームのアクション）生成 ============
function generateNextActions(data) {
  if (!GEMINI_API_KEY) {
    return { actions: '1. 困ってる人にDMで個別サポート\n2. 未回答者に火曜DMで声がけ\n3. マスター達成者にナレッジ共有依頼' };
  }
  const { currentMonth, totalSaved, heavyRate, respRate, troubled, unanswered, masters, deptStats, trend } = data;

  const prompt = `あなたはKONNEKT INTERNATIONALのAI推進チームに助言する経営アドバイザーです。
今月のデータを見て、AI推進チーム（藤本ゆりこ・引地瑞生）が「来月やるべき具体的アクション」を3つ提案してください。

【今月のデータ】
- 月：${currentMonth}
- 合計削減時間: ${totalSaved}h
- AI浸透率（毎日活用）: ${heavyRate}%
- 回答率: ${respRate}%
- 困ってる人: ${(troubled || []).join('・') || 'なし'}
- 未回答者: ${(unanswered || []).join('・') || 'なし'}
- AI伝説/マスター達成者: ${(masters || []).join('・') || 'なし'}
- 部署別: ${(deptStats || []).map(d => `${d.dept}=${d.saved.toFixed(1)}h`).join(', ')}
${trend ? `- 3ヶ月推移: ${trend}` : ''}

【出力ルール】
- 番号付きで3つの具体アクション
- 各アクション1-2行で簡潔に
- 「誰に・何を・いつ」を明確に
- AI推進チームが実際にやれること（DM・MTG・配布資料など）
- 困ってる人への個別サポートを必ず1つ含める
- 経営層に「AI推進チーム動いてるな」と思わせる内容

出力例：
1. 困ってる3名（菊田さん・植田さん・佐生さん）に火曜10:00個別DM → プロンプト集の活用ハンズオン誘導
2. マスター達成の藤本さん・川田さんに「ベスプラ共有会（30分）」を6月第2週で打診
3. 未回答5名向け：6月最初の水曜にリマインドDM＋「3分版簡易フォーム」用意`;

  const actions = callGemini(prompt, 700);
  return { actions };
}

// ============ 月次ダイジェスト生成 ============
function generateMonthlyDigest(data) {
  const {
    currentMonth, totalSaved, lastSaved, monthlySaving, roi, respRate, heavyRate,
    deptStats = [], initStats = [], voices = [], totalCount, responseCount
  } = data;

  const fallback = `🐱 AI Pulse 月次ダイジェスト（${currentMonth}）

📊 主要KPI
・合計時短時間：${(totalSaved || 0).toFixed(1)}h（先月比 ${totalSaved - lastSaved > 0 ? '+' : ''}${(totalSaved - lastSaved).toFixed(1)}h）
・月間コスト削減：¥${(monthlySaving || 0).toLocaleString()}（ROI ${roi > 0 ? '+' : ''}${roi || 0}%）
・回答率：${respRate || 0}%（${responseCount || 0}/${totalCount || 0}名）
・ヘビーユーザー率：${heavyRate || 0}%

🏆 部署別トップ3
${[...deptStats].sort((a,b) => b.saved - a.saved).slice(0,3).map((d,i) => `${i+1}. ${d.dept}：${d.saved.toFixed(1)}h（${d.count}名）`).join('\n')}

💡 主な「できるようになったこと」
${voices.slice(0,3).map(v => `・${v.didable}（${v.department}）`).join('\n')}

（AI Pulse / KONNEKT INTERNATIONAL）`;

  if (!ANTHROPIC_API_KEY) return { digest: fallback };

  const prompt = `あなたはKONNEKT INTERNATIONALのAI推進担当アシスタントです。経営層（社長・役員・AI推進リーダー）にSlack DMで送る月次ダイジェスト本文を作成してください。

【今月のデータ】
- 月：${currentMonth}
- 合計時短時間：${totalSaved}h（先月比 ${(totalSaved - lastSaved).toFixed(1)}h）
- 月間コスト削減：¥${monthlySaving?.toLocaleString()}（ROI ${roi}%）
- 回答率：${respRate}%（${responseCount}/${totalCount}名）
- ヘビーユーザー率：${heavyRate}%
- 部署別：${deptStats.map(d => `${d.dept}=${d.saved.toFixed(1)}h(${d.count}名)`).join(', ')}
- 施策実施率：${initStats.map(i => `${i.name}=${i.rate}%`).join(', ')}
- できるようになったこと：${voices.slice(0,5).map(v => `「${v.didable}」(${v.department})`).join(' / ')}

【出力ルール】
- Slack DMで読みやすい構成（絵文字・改行・箇条書き）
- 冒頭に「🐱 AI Pulse 月次ダイジェスト（${currentMonth}）」
- 数字を必ず入れる・前月比に触れる
- 3セクション程度：①主要KPI ②今月のハイライト ③経営への提言（1-2行）
- 全体で15行以内・600字以内
- 文末に「（AI Pulse / KONNEKT INTERNATIONAL）」`;

  try {
    const digest = callClaude(prompt, 1200);
    return { digest };
  } catch (e) {
    return { digest: fallback, warning: e.message };
  }
}

// ============ Slack配信 ============
function sendToSlack(data) {
  const slackToken = PROPS.getProperty('SLACK_BOT_TOKEN');
  if (!slackToken) throw new Error('SLACK_BOT_TOKEN が未設定です（スクリプトプロパティで設定してください）');

  // 配信先：スプシ「設定」タブの SLACK_RECIPIENTS を「,」区切りで取得
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const settingsSheet = ss.getSheetByName('設定');
  let recipientsRaw = '';
  if (settingsSheet) {
    const rows = settingsSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === 'SLACK_RECIPIENTS') { recipientsRaw = String(rows[i][1] || ''); break; }
    }
  }
  if (!recipientsRaw) throw new Error('スプシ「設定」タブに SLACK_RECIPIENTS（カンマ区切りメンバーID）を登録してください');

  const recipients = recipientsRaw.split(',').map(s => s.trim()).filter(Boolean);
  const message = data.message || '';
  const sent = [];
  const failed = [];

  for (const userId of recipients) {
    // im.open でDMチャンネル取得
    const openRes = UrlFetchApp.fetch('https://slack.com/api/conversations.open', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + slackToken, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ users: userId }),
      muteHttpExceptions: true
    });
    const openBody = JSON.parse(openRes.getContentText());
    if (!openBody.ok) { failed.push(`${userId}: ${openBody.error}`); continue; }
    const channelId = openBody.channel.id;

    // chat.postMessage
    const postRes = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + slackToken, 'Content-Type': 'application/json; charset=utf-8' },
      payload: JSON.stringify({ channel: channelId, text: message }),
      muteHttpExceptions: true
    });
    const postBody = JSON.parse(postRes.getContentText());
    if (!postBody.ok) { failed.push(`${userId}: ${postBody.error}`); continue; }
    sent.push(userId);
  }

  return { ok: failed.length === 0, sentTo: sent, failed };
}

// ============ 月次自動ダイジェスト（GASトリガー用） ============
/**
 * 月初の朝8:00 に GAS時間トリガーで自動実行する関数。
 * トリガー設定方法：
 *   左サイドバー「トリガー」→「+ トリガーを追加」
 *   関数：monthlyAutoDigest
 *   イベント：時間主導型 / 月タイマー / 1日 / 午前8〜9時
 */
function monthlyAutoDigest() {
  try {
    const dashboard = getDashboard();
    const responses = dashboard.responses;
    const now = new Date();
    // 前月をターゲットにする
    const target = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const currentMonth = Utilities.formatDate(target, 'JST', 'yyyy-MM');
    const lastDate = new Date(target.getFullYear(), target.getMonth() - 1, 1);
    const lastMonth = Utilities.formatDate(lastDate, 'JST', 'yyyy-MM');

    const thisM = responses.filter(r => r.month === currentMonth);
    const lastM = responses.filter(r => r.month === lastMonth);

    const master = getMaster();
    const totalCount = master.names.filter(n => n.active).length;
    const hourlyRate = master.settings['人件費単価_時給'] || 5000;
    const aiCost = master.settings['AI月額コスト'] || 50000;

    // q3は業務効率化%。月160h想定で時間換算
    const PCT_TO_HOURS = 1.6;
    const totalSaved = thisM.reduce((s, r) => s + (r.q3 || 0) * PCT_TO_HOURS, 0);
    const lastSaved = lastM.reduce((s, r) => s + (r.q3 || 0) * PCT_TO_HOURS, 0);
    const heavyCount = thisM.filter(r => r.q1 === 'ほぼ毎日').length;
    const heavyRate = thisM.length ? Math.round(heavyCount / thisM.length * 100) : 0;
    const respRate = totalCount ? Math.round(thisM.length / totalCount * 100) : 0;
    const monthlySaving = totalSaved * hourlyRate;
    const roi = aiCost > 0 ? Math.round((monthlySaving - aiCost) / aiCost * 100) : 0;

    const deptStats = master.departments.map(d => {
      const m = thisM.filter(r => r.department === d);
      return {
        dept: d,
        count: m.length,
        saved: m.reduce((s, r) => s + (r.q3 || 0) * PCT_TO_HOURS, 0)
      };
    });

    const initStats = master.initiatives.filter(i => i.active).map(i => {
      const did = thisM.filter(r => (r.initiatives || []).includes(i.id)).length;
      return { name: i.name, did, rate: thisM.length ? Math.round(did / thisM.length * 100) : 0 };
    });

    const voices = thisM.filter(r => r.didable && r.didable.length > 5).slice(0, 5);

    const { digest } = generateMonthlyDigest({
      currentMonth, totalSaved, lastSaved, monthlySaving, roi, respRate, heavyRate,
      deptStats, initStats, voices, totalCount, responseCount: thisM.length
    });

    const result = sendToSlack({ message: digest });
    Logger.log('Monthly digest sent: ' + JSON.stringify(result));
    return result;
  } catch (e) {
    Logger.log('monthlyAutoDigest error: ' + e.message);
    throw e;
  }
}

// ============ PIN更新（マイページから個人で変更） ============
function updatePin(data) {
  const { name, oldPin, newPin } = data;
  if (!name || !newPin) throw new Error('名前と新PINが必要です');
  if (!/^\d{4}$/.test(String(newPin))) throw new Error('新PINは4桁の数字で');

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('名簿');
  if (!sheet) throw new Error('名簿タブが見つかりません');

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === name) {
      // 旧PIN照合（指定があれば）
      if (oldPin && String(rows[i][4] || '1111') !== String(oldPin)) {
        throw new Error('現在のPINが違います');
      }
      // E列（5列目）にnewPinを書き込み
      sheet.getRange(i + 1, 5).setValue(String(newPin));
      return { ok: true, name };
    }
  }
  throw new Error('名前が見つかりません: ' + name);
}

// ============ マスター更新 ============
function updateMaster(data) {
  if (data.password !== ADMIN_PW) throw new Error('パスワードが違います');
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // 施策タブを上書き
  if (data.master?.initiatives) {
    const sheet = ss.getSheetByName('施策');
    sheet.clear();
    sheet.appendRow(['施策ID', '施策名', '月', '詳細', '有効フラグ']);
    data.master.initiatives.forEach(i => {
      sheet.appendRow([i.id, i.name, i.month, i.detail || '', i.active]);
    });
  }

  // 設定タブを上書き
  if (data.master?.settings) {
    const sheet = ss.getSheetByName('設定');
    sheet.clear();
    sheet.appendRow(['キー', '値']);
    Object.entries(data.master.settings).forEach(([k, v]) => {
      sheet.appendRow([k, v]);
    });
  }

  return { ok: true };
}
