# TODO

- [ ] `docs/ROADMAP.md` のフェーズ1/2を実作業で裏付けできる状態にする。現状、Apple HIG 調査・適用分析・テンプレート試作の成果が `コード.js` などへ反映されていない。
- [ ] `コード.js` の `CONFIG` を Apple 仕様へ移行する。`CONFIG.FONTS`/`CONFIG.COLORS` が Google 配色と Arial のままなので、要件で求める `CONFIG.APPLE_TOKENS` や Apple らしいカラーパレット・Inter などへ置き換える。
- [ ] Apple 向けレイアウトを生成するロジックを実装する。`docs/component_requirements.md` に記載されたテンプレートが `slideGenerators` / `slideData` でまだ扱われていない。
- [ ] テーマ切替のUXを要件通りにする。FR-02 の「生成時にライト/ダークを選択」が未実装で、メニュー名も `カスタム設定` のまま（FR-01 は "Apple Presentation" を要求）。
- [ ] `darkMode.js` と `themeToggler.js` の役割を整理する。現在は両方が `onOpen` を定義し、Google 配色と Apple 風配色が混在しているため、メニュー競合とテーマ不整合が発生する。
