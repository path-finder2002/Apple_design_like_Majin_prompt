/**
 * MVP dataset factory for diagram slide component.
 */
function createDiagramSlideData(overrides = {}) {
  return Object.assign({
    type: 'diagram',
    title: 'エクスペリエンスフロー',
    subhead: '主要なタッチポイントを整理',
    lanes: [
      {
        title: 'Discover',
        items: ['ヒアリング', '課題定義']
      },
      {
        title: 'Design',
        items: ['トークン設計', 'レイアウト検証']
      },
      {
        title: 'Deliver',
        items: ['リハーサル', '本番運用']
      }
    ],
    notes: 'レーンとカードの間隔、矢印接続をチェックするMVPです。'
  }, overrides || {});
}
