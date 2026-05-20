# AI Pulse — KONNEKT INTERNATIONAL

社員のAI活用度を月次でチェックイン・可視化するWebアプリ。

## 公開URL

https://fujimoto-cpu.github.io/ai-pulse/

（社内限定共有・noindex設定）

## 構成

```
ブラウザ（index.html / React+Tailwind CDN）
    ↓ fetch
Google Apps Script Webhook
    ↓
Google Sheets（KONNEKT_AIPulse_データ）
    ↓ プロキシ
Anthropic Claude API（haiku-4-5）
```

## 画面構成

| 画面 | アクセス |
|---|---|
| 全社ダッシュボード | メインナビ・誰でも閲覧可 |
| アンケート（チェックイン） | メインナビ・社員18名 |
| 管理画面 | 右上🛠アイコン・パスワード `admin2026` |
| 経営層ダッシュボード | 右上🔒アイコン・パスワード `exec2026` |

## ローカル動作確認

```bash
cd ~/Documents/Claude/Projects/ai-pulse
python3 -m http.server 8000
# ブラウザで http://localhost:8000 を開く
```

## デプロイ

```bash
gh repo create fujimoto-cpu/ai-pulse --public --source=. --push
gh api -X POST /repos/fujimoto-cpu/ai-pulse/pages -f source[branch]=main
```

## Apps Script設定（ゆりこ側で1回だけ実施）

詳細は `00_🏢 company/ai/20260519_AI効果数値化/GAS構築手順_v1.md` を参照。

1. Google Apps Script で新規プロジェクト作成
2. `gas/Code.gs` をコピペ
3. スクリプトプロパティに以下を設定
   - `SHEET_ID` : KONNEKT_AIPulse_データ のスプシID
   - `ANTHROPIC_API_KEY` : Claude APIキー
   - `ADMIN_PASSWORD` : admin2026
   - `EXEC_PASSWORD` : exec2026
4. ウェブアプリとしてデプロイ（「自分」として実行・「全員」がアクセス可）
5. 発行されたWebhook URLをアプリ起動時に管理画面で登録

## 仕様書

- `00_🏢 company/ai/20260519_AI効果数値化/アプリ仕様書_v1.md`
- `00_🏢 company/ai/20260519_AI効果数値化/GAS構築手順_v1.md`
