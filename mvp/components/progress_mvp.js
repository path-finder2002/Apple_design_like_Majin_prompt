/**
 * MVP dataset factory for progress slide component.
 */
function createProgressSlideData(overrides = {}) {
  return Object.assign({
    type: 'progress',
    title: 'タスク進捗',
    subhead: '主要タスクを可視化',
    items: [
      { label: 'デザイン言語定義', percent: 85 },
      { label: 'テンプレート実装', percent: 60 },
      { label: 'レビュー＆微調整', percent: 35 }
    ],
    notes: '進捗バーの割合と数値表示を検証します。'
  }, overrides || {});
}
