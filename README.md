# Apple Design Like Majin Prompt

## 概要 / Overview
- 日本語: Google Apps Script で Apple らしいミニマルな Google スライド資料を自動生成するスクリプトを管理するリポジトリです。
- English: Repository for a Google Apps Script that generates Apple-inspired, minimalist Google Slides decks by automating layout, typography, and theming.

## 主な特徴 / Key Features
- ライト/ダークテーマやアクセントカラーをワンクリックで切り替えるカスタムメニューを提供します。
- Cover、セクション、ヒーロー、カードレスフィーチャー、比較、タイムライン、引用、数値ハイライトなど複数のレイアウトを自動生成します。
- `CONFIG` 内の design token (`APPLE_TOKENS` など) で余白・ラインハイト・フォントスケールを一元管理し、プレゼン全体で一貫性を担保します。
- スクリプトプロパティ経由でフォント、プライマリカラー、フッターテキスト、ロゴなどを動的に差し替え可能です。
- 豊富なドキュメント (`docs/`) と Apple 風デザインの参考画像群 (`Apple_like_design_template/`) を同梱しています。

## リポジトリ構成 / Repository Contents
| パス | 説明 |
| --- | --- |
| `コード.js` | メインの Google Apps Script。本体の設定、レイアウト生成ロジック、カスタムメニューを実装。 |
| `appsscript.json` | Apps Script プロジェクト設定 (タイムゾーン、実行環境など)。 |
| `docs/` | 要件定義、ロードマップ、デザイン要件、Google スライド MVP 設計などの資料。 |
| `Apple_like_design_template/` | Apple プレゼンの参考レイアウト画像 (JPEG)。 |
| `AGENTS.md` | チームロールと開発フローのガイドライン。 |
| `Apple_like_system_prompt.md`, `system_prompt.md`, `GEMINI.md` | プロンプト設計や他モデル向けガイドライン。 |

## 前提条件 / Prerequisites
- Node.js (推奨: LTS) と npm
- `@google/clasp` CLI (`npm install -g @google/clasp`)
- Google アカウント (Apps Script と Google スライドの権限が必要)

## セットアップ手順 / Setup
1. `clasp login` でブラウザ認証を行います。
2. (既存プロジェクトに接続する場合) `.clasp.json` をリポジトリ直下に配置します。例:
   ```json
   {
     "scriptId": "<YOUR_SCRIPT_ID>",
     "rootDir": "."
   }
   ```
3. 最新の Apps Script と同期する場合は `clasp pull` を実行します。
4. ローカルで `コード.js` や関連ファイルを編集し、必要であれば `clasp push` で Apps Script 側へ反映します。

## 利用方法 / Usage
1. `clasp push` で変更を Google Apps Script にデプロイします。
2. 対象の Google スライドを開き、初回実行時は認可ダイアログに従って権限を付与します。
3. スライド上部の「カスタム設定」メニューから以下の機能にアクセスできます。
   - `🎨 スライドを生成`: `slideData` 定義に基づきプレゼンを一括生成。
   - テーマ切替 (`ライト/ダーク`)、プライマリカラー、フォント、フッターテキスト、ロゴの設定。
   - `🔄 リセット`: スクリプトプロパティを既定値に戻し、再生成を容易にします。
4. 出力結果をプレビューし、必要に応じて Google スライド上で微調整します。

## 開発フロー / Development Flow
- 作業前に `clasp pull` で Apps Script の最新状態を取得します。
- コード編集後はローカルで可能なチェック (コードレビュー、静的検証など) を行います。
- 問題がなければ `clasp push` で同期し、必要に応じて Git でコミットします。Git リモートへ push する際は事前に関係者へ確認してください。
- 認証情報や共有アカウントはハードコードしない方針です。設定値はスクリプトプロパティや `CONFIG` のトークン経由で管理します。

## 設定と拡張のヒント / Configuration Notes
- `SETTINGS` で既存スライドの削除可否 (`SHOULD_CLEAR_ALL_SLIDES`) やターゲットのプレゼン ID を制御できます。
- `CONFIG.APPLE_TOKENS` にタイポグラフィとスペーシングの基準値がまとまっており、新しいレイアウトを追加する場合はこのトークンを参照して整合性を保ちます。
- 新規レイアウトを追加する際は `slideGenerators` へ関数を定義し、`slideData` から呼び出す構成です。詳細は `docs/component_requirements.md` を参照してください。
- カラー・フォント・ロゴ設定は `PropertiesService` (スクリプトプロパティ) に永続化されるため、ユーザーごとにカスタマイズ可能です。

## 参考資料 / References
- 要件: `docs/requirements.md`
- ロードマップ: `docs/ROADMAP.md`
- デザイン仕様: `docs/component_requirements.md`, `docs/google_slide_mvp_design.md`
- テキスト/画像計測レポート: `docs/テキストサイズレポート.md`, `docs/画像計測レポート.md`
- Apple 風デザインの視覚例: `Apple_like_design_template/*.jpeg`

## 注意事項 / Notes
- Apps Script の実行時間・クォータ制限を考慮してスライド構成を調整してください。
- `.clasp.json` は共有しないため、開発者それぞれがローカルで管理します。
- リポジトリの差分を第三者と共有する際は、認証情報や個人情報が含まれていないか必ず確認してください。

