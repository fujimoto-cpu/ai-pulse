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
    if (action === 'generate_monthly_digest') return jsonOut(generateMonthlyDigest(data));
    if (action === 'send_to_slack') return jsonOut(sendToSlack(data));
    if (action === 'update_master') return jsonOut(updateMaster(data));
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

  return { names, initiatives, settings, departments };
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
    return {
      timestamp: ts.toISOString(),
      month: r[1] || Utilities.formatDate(ts, 'JST', 'yyyy-MM'),
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

【レコメンド作成ルール】
- 「次のあなたにおすすめ」として、具体的なアクション1つだけ提案
- 既にこの人が使いこなしてる領域は除外（過去の場面・できたことで判断）
- 伸びしろある領域・難易度ステップアップする方向で
- 困り度が高い人（😟😣）にはまず「困り解消」のアドバイスを優先
- 社内施策候補からも選んでOK（その人にマッチする場合）
- ${name}さんの「来月やりたい方向」（${cur.q5 || '不明'}）も考慮
- 出力フォーマット：見出し1行＋具体的に何するか2-3行
- 絵文字1-2個OK・フレンドリー口調・敬語不要
- 「これ試してみない？」みたいな提案口調

出力例（イメージ）:
💡 来月のおすすめ：議事録から提案書へステップアップ
今月の議事録自動化バッチリにゃ！来月は同じノリで「提案書のたたき台作り」を試してみない？
Claudeに「○○の提案書、議事録を元に作って」って投げるだけで初稿出てくるよ✨`;

  const recommendation = callGemini(prompt, 600);
  return { recommendation };
}

// ============ 経営層分析生成（Claude API） ============
function generateExecAnalysis(data) {
  if (!GEMINI_API_KEY) return { analysis: '（APIキー未設定）' };
  const { totalSaved, lastSaved, heavyRate, respRate, deptStats, initStats } = data;

  const prompt = `あなたはAI経営アナリストです。KONNEKT INTERNATIONAL の今月のAI活用データを見て、経営層向けに前月比の変化点を3-5個、定量的に指摘してください。

【今月】合計時短=${totalSaved}h / ヘビーユーザー率=${heavyRate}% / 回答率=${respRate}%
【先月】合計時短=${lastSaved}h
【差分】${(totalSaved - lastSaved).toFixed(1)}h

【部署別】${deptStats.map(d => `${d.dept}:${d.saved.toFixed(1)}h(${d.count}名)`).join(', ')}
【施策実施率】${initStats.map(i => `${i.name}:${i.rate}%(${i.did}名)`).join(', ')}

出力ルール:
- 経営判断に資する事実ベースの分析（褒め言葉や精神論は不要）
- 数字を必ず入れる
- 3-5個の箇条書き
- 最後に1行で「経営への提言」`;

  const analysis = callGemini(prompt, 1500);
  return { analysis };
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
