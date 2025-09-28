# Apple Design Like Majin Prompt

## 概要
Google Apps Script を用いて Apple らしいミニマルな Google スライド資料を自動生成するスクリプトを管理するリポジトリです。レイアウト、タイポグラフィ、テーマ適用を自動化し、一定品質のプレゼン資料を素早く作成できます。

## 主な特徴
- カスタムメニューからライト / ダークテーマおよびアクセントカラーをワンクリックで切り替え可能。
- 表紙、セクション区切り、ヒーロー、カード型、比較、タイムライン、引用、数値ハイライトなど多彩なレイアウトを生成。
- `CONFIG` 配下のデザイントークンで余白・行間・フォントスケールを集中管理し、全体の一貫性を担保。
- スクリプトプロパティ経由でフォント、プライマリカラー、フッターテキスト、ロゴなどを柔軟に差し替え可能。
- ドキュメント類 (`docs/`) と Apple 風レイアウトのサンプル SVG (`img/`) を同梱。

## リポジトリ構成
| パス | 説明 |
| --- | --- |
| `src/` | Apps Script の主要モジュール (`config.js`, `presentation.js`, `slides_*.js`, `webapp.js`) とダイアログ UI (`index.html`) を格納。 |
| `src/appsscript.json` | Apps Script プロジェクト設定 (実行権限・タイムゾーン・ランタイム)。 |
| `docs/` | 設計資料群。要件は `docs/Requirements/`、プロンプトは `docs/prompt/` に整理。 |
| `img/` | スライドレイアウトの参考 SVG。配色や構図の確認に利用可能。 |
| `ProductSans-Regular.ttf` | デモ用フォント。実運用では Google Fonts 等で代替可能。 |

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
4. ローカルで `src/` 配下のスクリプトを編集し、必要に応じて `clasp push` で反映。

## 利用方法
1. `clasp push` で変更を Google Apps Script にデプロイ。
2. 対象の Google スライドを開き、初回実行時は認可ダイアログで権限を付与。
3. スライド上部の「カスタム設定」メニューから以下を利用:
   - `🎨 スライドを生成`: `slideData` 定義にもとづきプレゼン資料を一括生成。
   - テーマ切り替え (ライト / ダーク)、プライマリカラー、フォント、フッターテキスト、ロゴの設定。
   - `🔄 リセット`: スクリプトプロパティを既定値へ戻し再生成を容易にする。
4. 出力されたスライドをプレビューし、必要に応じて Google スライド上で微調整。

### 初回動作確認のサンプル
1. GAS エディタで次の `slideData` を一時的に貼り付け、`createPresentation(slideData, settings)` を実行。
   ```javascript
   const slideData = [
     { type: 'title', title: 'Majin Prompt Kickoff', date: '2024.05.01' },
     { type: 'agenda', title: '今日の流れ', items: ['背景', '進め方', '次の一歩'] },
     {
       type: 'content',
       title: '背景',
       subhead: 'Why Apple-like?',
       points: ['スライドの世界観を統一したい', 'GAS で量産を自動化したい']
     },
     { type: 'closing', title: 'Thanks for watching' }
   ];

   const settings = {
     primaryColor: '#0A84FF',
     footerText: '© 2024 Majin Prompt Team',
     fontFamily: 'Product Sans',
     showBottomBar: true,
     showDateColumn: false
   };

   function runQuickStartDemo() {
     return createPresentation(slideData, settings);
   }
   ```
2. `runQuickStartDemo` を実行し、生成された 3 枚のスライドで配色・余白の雰囲気を確認。
   - スライド上のカスタムメニューを使う場合は、同じ `slideData` をダイアログに貼り付けて `🎨 スライドを生成` を実行。
3. `settings` の色やロゴ URL を差し替えて再実行し、どこが変わるかを把握する。

### よくあるつまづきと対処
- `Exception: Request had insufficient authentication scopes.` が出た場合は、`src/appsscript.json` の `oauthScopes` が最新か確認し、`clasp login` → `clasp push` で再デプロイする。
- ロゴや背景画像が表示されない場合は、共有ドライブのファイル権限 (スライドを実行するアカウントに閲覧権限があるか) と URL/ファイル ID の指定ミスを確認する。
- スライド生成中に処理が止まる場合は、Apps Script の実行ログを開き、どの `type` のスライドでエラーが出ているかを確認して `slideData` を修正する。
- `clasp` の開発者サーバーやローカルプレビューで開くと HTML テンプレートが評価されないため UI だけが表示されます。実際に生成を行うには Apps Script の Web アプリとして実行してください。

## 開発フロー
- 作業前に `clasp pull` で最新状態を取得。
- コード編集後はローカルで可能なチェック (レビュー、静的解析など) を実施。
- 問題なければ `clasp push` で同期し、必要に応じて Git でコミット。リモートへ push する際は関係者に確認。
- 認証情報や共有アカウントはハードコードせず、スクリプトプロパティや `CONFIG` のトークンで管理。

## 設定と拡張のヒント
- `SETTINGS` で既存スライド削除 (`SHOULD_CLEAR_ALL_SLIDES`) やターゲットプレゼン ID を制御。
- `CONFIG.APPLE_TOKENS` にタイポグラフィとスペーシングの基準値がまとまっており、新しいレイアウト追加時はこのトークンを参照。
- 新規レイアウトを追加する際は `slideGenerators` に関数を定義し、`slideData` から呼び出す構成。詳細は `docs/Requirements/component_requirements.md` を参照。
- カラー・フォント・ロゴの設定は `PropertiesService` (スクリプトプロパティ) に永続化され、ユーザーごとにカスタマイズ可能。

## 参考資料
- 要件: `docs/Requirements/requirements.md`
- ロードマップ: `docs/ROADMAP.md`
- デザイン仕様: `docs/Requirements/component_requirements.md`, `docs/google_slide_mvp_design.md`
- 課題/議論ログ: `docs/issues/`, `docs/TODO.md`
- 更新履歴: `docs/log/CHANGELOG.md`
- Apple 風デザイン参考画像: `img/*.svg`

## 注意事項
- Apps Script の実行時間・クォータ制限を考慮してスライド構成を調整する。
- `.clasp.json` は共有しないため、各開発者がローカルで管理する。
- リポジトリの差分を共有する際は、認証情報や個人情報が含まれていないか必ず確認する。
