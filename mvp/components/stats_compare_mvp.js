/**
 * MVP dataset factory for stats compare slide component.
 */
function createStatsCompareSlideData(overrides = {}) {
  return Object.assign({
    type: 'statsCompare',
    title: '主要指標の比較',
    subhead: '現状指標と目標指標の差分を視覚化',
    leftTitle: '現状',
    rightTitle: '目標',
    stats: [
      { label: 'NPS', leftValue: '42', rightValue: '65', trend: 'up' },
      { label: '体験満足度', leftValue: '3.8', rightValue: '4.5', trend: 'up' },
      { label: '導入コスト', leftValue: '120万', rightValue: '80万', trend: 'down' }
    ],
    notes: '数値比較テンプレートの色と整列を確認するMVPです。'
  }, overrides || {});
}
