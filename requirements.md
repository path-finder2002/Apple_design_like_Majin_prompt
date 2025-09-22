# 要件定義 / Requirements Definition

## プロジェクト概要 / Project Overview
- 日本語: Google Apps Script を用いて、まじん式スライドの仕組みを継承しつつ Apple らしいミニマルで洗練された Google スライド資料を自動生成するプロジェクトです。
- English: This project uses Google Apps Script to evolve the Majin-style automation into Apple-inspired, minimal, and refined Google Slides presentations.

## 目的 / Objectives
- 日本語: 余白を生かしたタイポグラフィ主導のスライドを一貫した品質で短時間に生成する。
- English: Deliver typography-led slides with generous white space quickly and consistently.
- 日本語: Inter（SF Pro 代替）を駆使したモノトーン中心のテーマで Apple らしいブランド体験を再現する設定を提供する。
- English: Provide configuration that leverages Inter (SF Pro alternative) and monochromatic themes to recreate an Apple-like brand experience.
- 日本語: 既存コード資産を保ちつつ、レイアウトやトークンを Apple デザイン原則に合わせてモジュール化し将来の拡張を容易にする。
- English: Modularize layouts and tokens around Apple design principles while preserving existing code assets for future extensibility.

## ステークホルダーとロール / Stakeholders and Roles
- 日本語: Developer Agent – Apple 風デザイン向けレイアウト実装、トークン管理、テストを担当。
- English: Developer Agent – Implements Apple-style layouts, manages design tokens, and handles testing.
- 日本語: Reviewer Agent – デザインポリシー適合性とコード品質のレビューを実施。
- English: Reviewer Agent – Reviews for design policy compliance and code quality.
- 日本語: Ops Agent – clasp 運用、Apps Script への反映、フォントや画像アセットの管理を支援。
- English: Ops Agent – Supports clasp operations, Apps Script deployment, and font/asset management.
- 日本語: デザインリード – Apple らしさの観点からトーン&マナーを監修し、要件の承認を行う。
- English: Design Lead – Oversees tone and manner for Apple authenticity and approves requirements.
- 日本語: 最終利用者 – Apple テイストの資料を求めるビジネスユーザーやデザインチーム。
- English: End users – Business and design teams requiring Apple-flavored presentations.

## スコープ / Scope
- 日本語: Apple HIG を踏まえたレイアウトテンプレート、カラーパレット、タイポグラフィトークンの設計と自動配置。
- English: Design and automated placement of layouts, color palettes, and typography tokens grounded in the Apple HIG.
- 日本語: ライト/ダークモード切替、アクセントカラー調整、デバイスフレーム挿入など Apple らしい演出コントロール。
- English: Controls for light/dark modes, accent color tuning, and device frame insertion to convey an Apple look.
- 日本語: カスタムメニュー経由での設定編集、スクリプトプロパティ保存、clasp を用いた同期フローの維持。
- English: Maintain custom menu configuration, script property persistence, and clasp-based synchronization.
- 日本語（外スコープ）: 画像編集や 3D アニメーション生成、Apple 公式アセットの配布、外部 API 連携。
- English (Out of scope): Image retouching, 3D animation generation, distribution of official Apple assets, or external API integrations.

## 機能要件 / Functional Requirements
| ID | 日本語要件 | English Requirement |
|----|------------|---------------------|
| FR-01 | Google スライド起動時に「Apple Presentation」メニューを表示し、デザイン生成・設定機能へアクセスできること。 | Show an "Apple Presentation" menu on open to access generation and configuration features. |
| FR-02 | 生成実行時、ライトまたはダークテーマを選択し、選択結果に応じて背景・文字色・アクセントカラーを適用すること。 | Allow choosing light or dark theme at generation and apply matching backgrounds, typography, and accent colors. |
| FR-03 | カバースライド、ミニマルセクション、ワイドヒーロー、カードレス・フィーチャー、シンプル比較、ナラティブタイムライン、ステートメント引用、数値ハイライトの各レイアウトを提供すること。 | Provide layouts for cover, minimal section, wide hero, cardless feature, simple comparison, narrative timeline, statement quote, and metric highlight slides. |
| FR-04 | 余白・行間・フォントサイズを design token として `CONFIG.APPLE_TOKENS` に集約し、スライド描画時に参照すること。 | Centralize spacing, line height, and type scale tokens in `CONFIG.APPLE_TOKENS` and reference them during rendering. |
| FR-05 | San Francisco フォントが利用できない環境では自動的に Noto Sans / Helvetica へフォールバックし、利用者へメッセージを表示すること。 | Fall back to Noto Sans / Helvetica when San Francisco fonts are unavailable and notify the user. |
| FR-06 | アクセントカラー設定時に WCAG AA のコントラストを満たさない場合は警告を表示し、既定値に戻すオプションを提供すること。 | Warn when chosen accent colors violate WCAG AA contrast and offer reverting to defaults. |
| FR-07 | キーヴィジュアル用のデバイス枠（iPhone・MacBook 風）を任意で挿入し、横幅比率と配置を自動調整すること。 | Optionally insert device frames (iPhone/MacBook style) with automated scaling and placement. |
| FR-08 | スピーカーノート、アクセシビリティ用代替テキスト、生成ログをメニューから確認・再出力できること。 | Provide menu access to speaker notes, accessibility alt text, and generation logs for review or re-export. |
| FR-09 | 設定リセット時は Apple テーマ用の初期トークン（ライトモード／ミッドグレイ背景／アクセント#0A84FF）へ戻すこと。 | Reset settings to Apple theme defaults (light mode, mid-gray backgrounds, accent `#0A84FF`). |
| FR-10 | 生成結果をプレビュー用サンドボックスプレゼンに出力後、利用者が承認すると本番プレゼンへ適用する 2 段階フローを提供すること。 | Support a two-step flow that renders to a sandbox deck first and applies to the target deck upon user approval. |
| FR-11 | テーブルレイアウトは chatgpt.com のように横線のみで区切り、ヘッダー太字・行間を広く確保し、垂直線や塗りつぶし背景を使用しないこと。 | Render table layouts with horizontal separators only, bold headers, generous row spacing, and no vertical rules or filled backgrounds to mirror chatgpt.com styling. |
| FR-12 | 各スライドの表示要素は 3〜4 オブジェクト以内に制限し、超過する場合は自動でスライドを分割すること。 | Limit visible objects per slide to 3–4 and automatically split slides when the threshold is exceeded. |
| FR-13 | タイトルスライドは 1 つのテキストオブジェクトのみを中央配置で描画し、ロゴ・サブタイトル・影・自動改行を禁止すること。 | Render title slides with exactly one centered text object, prohibiting logos, subtitles, shadows, or auto line breaks. |
| FR-14 | タイトル文字サイズはスライド幅に応じて安全余白（上下 7.5%、左右 6%）内で最大化し、Inter 600 / letter-spacing 0〜0.5px を維持すること。 | Maximize title font size within safe margins (top/bottom 7.5%, left/right 6%) using Inter weight 600 and 0–0.5px letter spacing. |
| FR-15 | カード、図、プログレスバーは半径 20〜24px の角丸と控えめな影、最大 3 要素までのレイアウトを用意すること。 | Provide card, diagram, and progress layouts with 20–24px rounded corners, subtle shadows, and at most three elements. |
| FR-16 | 生成結果は `style`（デザイントークン）と `slideData`（スライド配列）の 2 オブジェクトとして出力し、Apps Script から直接読み込める JSON 等価形式にすること。 | Output generation results as two objects—`style` design tokens and `slideData` slide array—in a JSON-equivalent structure consumable by Apps Script. |

## 非機能要件 / Non-Functional Requirements
| ID | 日本語要件 | English Requirement |
|----|------------|---------------------|
| NFR-01 | スクリプトは Google Apps Script V8 ランタイムで 6 分以内に完了し、100 スライド以内の構成を想定する。 | Complete within six minutes on Google Apps Script V8 and target decks up to 100 slides. |
| NFR-02 | Apple HIG のグリッド（8pt グリッド）とモジュール比率を破らないようレイアウトを算出する。 | Compute layouts that honor the Apple HIG grid (8pt) and modular ratios. |
| NFR-03 | デザインパラメータは JSON 形式のトークン定義で管理し、Pull Request ベースで変更を可視化できること。 | Manage design parameters via JSON-like token definitions tracked through pull requests. |
| NFR-04 | 全スライドで最低 16px の余白と 28px の見出しサイズを確保し、可読性を保証する。 | Guarantee minimum 16px padding and 28px heading size across slides for readability. |
| NFR-05 | ログ・警告は日本語と英語で出力し、デザインチームと開発チーム双方が理解できるようにする。 | Output logs and warnings bilingually so both design and engineering teams can understand. |
| NFR-06 | 依存ライブラリを追加する場合は Apps Script の制限内で軽量に保ち、ビルドステップを増やさない。 | Keep any new dependencies lightweight within Apps Script limits and avoid introducing build steps. |

## データおよび設定 / Data and Configuration
- 日本語: `slideData` は Apple 風テンプレート向けフィールド（ライト/ダーク指定、アクセントカラー、デバイス枠有無、短いステートメント）を含む構造へ更新する。
- English: Update `slideData` to include Apple-themed fields (light/dark selection, accent color, device frame flags, succinct statements).
- 日本語: `CONFIG.APPLE_TOKENS` にタイポグラフィスケール、余白、カラーセット、影設定を定義する。
- English: Define typography scale, spacing, color sets, and elevation settings in `CONFIG.APPLE_TOKENS`.
- 日本語: テーブル描画は横罫線のみを用い、行間に余白を設けて chatgpt.com 風の軽やかな見た目を再現する。
- English: Draw tables using horizontal rules only with airy row spacing to emulate the chatgpt.com look.
- 日本語: `style` オブジェクトにテーマ別トークン（フォント、カラー、スペーシング、アルゴリズム）を保持し、`slideData` とセットで出力する。
- English: Store theme-specific tokens (fonts, colors, spacing, algorithms) inside a `style` object and output it alongside `slideData`.
- 日本語: タイトル自動フィット用パラメータ（安全余白、最大行数、letter-spacing）を `style.algorithms.titleFit` に構造化して定義する。
- English: Define title auto-fit parameters (safe margins, max lines, letter spacing) under `style.algorithms.titleFit`.
- 日本語: デフォルトアセットはローカルの `img/apple` ディレクトリで管理し、公開 URL に変換して利用する。
- English: Store default assets under `img/apple` locally and expose them via public URLs for use.
- 日本語: ライト/ダークテーマ差分は設定プロパティとスライド ID の組み合わせで記録する。
- English: Persist light/dark variations by pairing configuration properties with slide identifiers.

## 業務フロー / Operational Flow
- 日本語: 作業開始時に `clasp pull` を実行し、Apple テーマ用トークン更新後に `clasp push` で Apps Script と同期する。
- English: Run `clasp pull` before work and push after Apple theme token updates to sync with Apps Script.
- 日本語: デザインリードがプレビュー資料を確認し、承認後 Ops Agent が本番デッキへ適用する。
- English: Design lead reviews the preview deck and, once approved, the Ops Agent applies it to the production deck.
- 日本語: アクセシビリティ検証（コントラストチェック、フォント確認）を完了後にリリースとする。
- English: Release only after accessibility verification covering contrast and font availability.

## 制約と前提 / Constraints and Assumptions
- 日本語: サンフランシスコ系フォントはローカル環境に依存し、第三者配布は行わない前提とする。
- English: Assume San Francisco fonts depend on local availability and will not be redistributed.
- 日本語: Apple 公式ロゴやデバイス画像は使用せず、独自に作成した汎用アセットを用いる。
- English: Use custom-created generic assets instead of official Apple logos or device imagery.
- 日本語: テーマ切替やプレビューのため、対象プレゼンとは別にサンドボックス用プレゼン ID を設定する必要がある。
- English: Require a sandbox presentation ID separate from the target deck to enable theme switching and previews.
- 日本語: タイトルスライドは 1 テキストのみ許容されるため、slideData に複数要素を定義しないこと。
- English: Restrict title slides to a single text node—do not wire multiple elements in slideData.
- 日本語: Apps Script の API 制限上、画像挿入は 30 枚以内を目安とし、それ以上は手動対応とする。
- English: Due to Apps Script API quotas, limit automated image insertions to roughly 30; handle extras manually.

## リスクと対応策 / Risks and Mitigations
- 日本語: フォント未導入によるデザイン崩れ → フォールバック適用と導入ガイドの表示で対応。
- English: Font absence causing layout shifts → Apply fallbacks and surface installation guidance.
- 日本語: 過度なアクセントカラー指定で Apple らしさが損なわれる → 設定画面で推奨色プリセットを提示。
- English: Excessive accent colors diluting Apple tone → Present recommended presets in the settings dialog.
- 日本語: 余白トークン変更による崩れ → デザインスナップショットとビジュアル差分チェックで回帰検知。
- English: Spacing token regressions → Maintain visual snapshots and run diff checks to detect breakages.

## 未解決事項 / Open Issues
- 日本語: San Francisco フォント利用に関するライセンス表現をメニュー上にどう明記するか。
- English: How to document San Francisco font licensing notices within the menu interface remains undecided.
- 日本語: 2 段階プレビュー承認フローでの UX（自動遷移か手動遷移か）を要検討。
- English: UX for the two-step preview approval flow (automatic vs manual transition) needs refinement.
- 日本語: Apple らしいマイクロアニメーションをどの範囲で許容するかの基準が未定義。
- English: Criteria for acceptable Apple-style micro-animations have yet to be defined.
