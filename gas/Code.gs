/**
 * AI Pulse Webhook (Google Apps Script)
 * KONNEKT INTERNATIONAL 社員のAI活用度をスプシで管理＋Claude APIでAIコメント生成
 *
 * セットアップ手順は `00_🏢 company/ai/20260519_AI効果数値化/GAS構築手順_v1.md` を参照
 *
 * スクリプトプロパティ必須:
 *   - SHEET_ID            : スプレッドシートID
 *   - ANTHROPIC_API_KEY   : Claude APIキー（sk-ant-...）
 *   - ADMIN_PASSWORD      : 管理画面パスワード（既定: admin2026）
 *   - EXEC_PASSWORD       : 経営層パスワード（既定: exec2026）
 *   - MODEL               : Claudeモデル（既定: claude-haiku-4-5-20251001）
 */

const PROPS = PropertiesService.getScriptProperties();
const SHEET_ID = PROPS.getProperty('SHEET_ID');
const ANTHROPIC_API_KEY = PROPS.getProperty('ANTHROPIC_API_KEY');
const ADMIN_PW = PROPS.getProperty('ADMIN_PASSWORD') || 'admin2026';
const EXEC_PW = PROPS.getProperty('EXEC_PASSWORD') || 'exec2026';
const MODEL = PROPS.getProperty('MODEL') || 'claude-haiku-4-5-20251001';

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
    if (action === 'generate_exec_analysis') return jsonOut(generateExecAnalysis(data));
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

  const names = sheetToObjects(namesSheet, ['name', 'department', 'email', 'active']);
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
  const row = [
    new Date().toISOString(),     // タイムスタンプ
    data.name || '',
    data.department || '',
    data.q1 || '',
    JSON.stringify(data.q2 || []),
    data.q3 || 0,
    data.q4 || '',
    data.q5 || '',
    data.q6 || '',
    data.q7 || '',
    JSON.stringify(data.initiatives || []),
    data.didable || '',
    data.comment || '',
    score
  ];
  sheet.appendRow(row);
  return { ok: true, score };
}

function calcScore(ans) {
  let s = 0;
  const q1Map = { 'ほぼ毎日': 25, '週2-3日': 18, '週1日': 10, 'ほぼ使ってない': 3 };
  s += q1Map[ans.q1] || 0;
  s += Math.min(20, (ans.q2 || []).length * 5);
  s += Math.min(20, Math.round((ans.q3 || 0) * 2));
  const q4Map = { '😫': 0, '😐': 5, '😊': 10, '🔥': 15 };
  s += q4Map[ans.q4] || 0;
  const q5Map = { 'プロンプト改善': 10, '新ツール試す': 8, 'チーム共有': 10, '現状維持': 3 };
  s += q5Map[ans.q5] || 0;
  const q6Map = { '😫': 0, '😐': 3, '😊': 7, '🔥': 10 };
  s += q6Map[ans.q6] || 0;
  return Math.max(0, Math.min(100, s));
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
      month: Utilities.formatDate(ts, 'JST', 'yyyy-MM'),
      name: r[1],
      department: r[2],
      q1: r[3],
      q2: tryParseJSON(r[4]) || [],
      q3: parseFloat(r[5]) || 0,
      q4: r[6],
      q5: r[7],
      q6: r[8],
      q7: r[9],
      initiatives: tryParseJSON(r[10]) || [],
      didable: r[11],
      comment: r[12],
      score: parseInt(r[13]) || 0
    };
  });
  return { responses };
}

function tryParseJSON(s) {
  try { return JSON.parse(s); } catch { return null; }
}

// ============ 猫キャラコメント生成（Claude API） ============
function generateCatComment(data) {
  const { name, score, currentMonth } = data;
  if (!ANTHROPIC_API_KEY) return { comment: '今月もおつかれにゃ！' };

  // 当人の前月データを取得
  const dashboard = getDashboard();
  const myHist = dashboard.responses.filter(r => r.name === name).sort((a, b) => a.month.localeCompare(b.month));
  const lastEntry = myHist.length >= 2 ? myHist[myHist.length - 2] : null;
  const diff = lastEntry ? score - lastEntry.score : 0;

  const prompt = `あなたは社員のAI活用を応援する猫キャラ「にゃんこ」です。${name}さんの今月のAI活用結果を見て、振り返り＋来月のアドバイスを2-3行で返してください。

【今月のスコア】${score}/100pt
${lastEntry ? `【先月のスコア】${lastEntry.score}/100pt（差分 ${diff > 0 ? '+' : ''}${diff}pt）` : '【先月】データなし（今月が初回）'}
【最新の回答】使用頻度=${data.q1 || '不明'} / 節約時間=${data.q3 || 0}h / 成長実感=${data.q4 || '不明'}

ルール:
- 語尾に「にゃ」をつける
- 前月比に触れる（伸びてれば褒める・下がってれば優しく後押し）
- 来月の具体的アクションを1つだけ提案
- 2-3行・絵文字1-2個OK・改行入れて読みやすく`;

  const comment = callClaude(prompt, 300);
  return { comment };
}

// ============ 経営層分析生成（Claude API） ============
function generateExecAnalysis(data) {
  if (!ANTHROPIC_API_KEY) return { analysis: '（APIキー未設定）' };
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

  const analysis = callClaude(prompt, 1500);
  return { analysis };
}

// ============ Claude API 呼び出し ============
function callClaude(prompt, maxTokens) {
  const url = 'https://api.anthropic.com/v1/messages';
  const payload = {
    model: MODEL,
    max_tokens: maxTokens || 500,
    messages: [{ role: 'user', content: prompt }]
  };
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const body = JSON.parse(res.getContentText());
  if (body.error) throw new Error(body.error.message || JSON.stringify(body.error));
  return body.content[0].text;
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
