# Apple Design Like Majin Prompt

## 概要
Google Apps Script を用いて Apple らしいミニマルな Google スライド資料を自動生成するスクリプトを管理するリポジトリです。レイアウト、タイポグラフィ、テーマ適用を自動化し、一定品質のプレゼン資料を素早く作成できます。

## 主な特徴
- カスタムメニューからライト / ダークテーマおよびアクセントカラーをワンクリックで切り替え可能。
- 表紙、セクション区切り、ヒーロー、カード型、比較、タイムライン、引用、数値ハイライトなど多彩なレイアウトを生成。
- `CONFIG` 配下のデザイントークンで余白・行間・フォントスケールを集中管理し、全体の一貫性を担保。
- スクリプトプロパティ経由でフォント、プライマリカラー、フッターテキスト、ロゴなどを柔軟に差し替え可能。
- ドキュメント類 (`docs/`) と Apple 風デザインの参考画像 (`Apple_like_design_template/`) を同梱。

## リポジトリ構成
| パス | 説明 |
| --- | --- |
| `コード.js` | メインの Google Apps Script。本体設定、レイアウト生成ロジック、カスタムメニューを実装。 |
| `appsscript.json` | Apps Script プロジェクト設定 (タイムゾーン・実行環境など)。 |
| `docs/` | 要件定義、ロードマップ、デザイン要件、Google スライド MVP 設計の資料。 |
| `Apple_like_design_template/` | Apple プレゼンの参考レイアウト画像 (JPEG)。 |
| `AGENTS.md` | チームロールと開発フローのガイドライン。 |
| `Apple_like_system_prompt.md`, `system_prompt.md`, `GEMINI.md` | プロンプト設計や他モデル向けガイドライン。 |

## 前提条件
- Node.js (推奨: LTS) と npm
- `@google/clasp` CLI (`npm install -g @google/clasp`)
- Google アカウント (Apps Script / Google スライドの権限が必要)

## セットアップ手順
1. `clasp login` でブラウザ認証を行う。
2. 既存プロジェクトに接続する場合は `.clasp.json` をリポジトリ直下に配置する。例:
   ```json
   {
     "scriptId": "<YOUR_SCRIPT_ID>",
     "rootDir": "."
   }
   ```
3. Apps Script 側の最新版を取得する場合は `clasp pull` を実行。
4. ローカルで `コード.js` などを編集し、必要に応じて `clasp push` で反映。

## 利用方法
1. `clasp push` で変更を Google Apps Script にデプロイ。
2. 対象の Google スライドを開き、初回実行時は認可ダイアログで権限を付与。
3. スライド上部の「カスタム設定」メニューから以下を利用:
   - `🎨 スライドを生成`: `slideData` 定義にもとづきプレゼン資料を一括生成。
   - テーマ切り替え (ライト / ダーク)、プライマリカラー、フォント、フッターテキスト、ロゴの設定。
   - `🔄 リセット`: スクリプトプロパティを既定値へ戻し再生成を容易にする。
4. 出力されたスライドをプレビューし、必要に応じて Google スライド上で微調整。

## 開発フロー
- 作業前に `clasp pull` で最新状態を取得。
- コード編集後はローカルで可能なチェック (レビュー、静的解析など) を実施。
- 問題なければ `clasp push` で同期し、必要に応じて Git でコミット。リモートへ push する際は関係者に確認。
- 認証情報や共有アカウントはハードコードせず、スクリプトプロパティや `CONFIG` のトークンで管理。

## 設定と拡張のヒント
- `SETTINGS` で既存スライド削除 (`SHOULD_CLEAR_ALL_SLIDES`) やターゲットプレゼン ID を制御。
- `CONFIG.APPLE_TOKENS` にタイポグラフィとスペーシングの基準値がまとまっており、新しいレイアウト追加時はこのトークンを参照。
- 新規レイアウトを追加する際は `slideGenerators` に関数を定義し、`slideData` から呼び出す構成。詳細は `docs/component_requirements.md` を参照。
- カラー・フォント・ロゴの設定は `PropertiesService` (スクリプトプロパティ) に永続化され、ユーザーごとにカスタマイズ可能。

## 参考資料
- 要件: `docs/requirements.md`
- ロードマップ: `docs/ROADMAP.md`
- デザイン仕様: `docs/component_requirements.md`, `docs/google_slide_mvp_design.md`
- テキスト/画像計測レポート: `docs/テキストサイズレポート.md`, `docs/画像計測レポート.md`
- Apple 風デザイン参考画像: `Apple_like_design_template/*.jpeg`

## 注意事項
- Apps Script の実行時間・クォータ制限を考慮してスライド構成を調整する。
- `.clasp.json` は共有しないため、各開発者がローカルで管理する。
- リポジトリの差分を共有する際は、認証情報や個人情報が含まれていないか必ず確認する。

