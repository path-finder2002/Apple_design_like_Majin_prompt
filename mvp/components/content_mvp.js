/**
 * MVP dataset factory for content slide component.
 */
function createContentSlideData(overrides = {}) {
  return Object.assign({
    type: 'content',
    title: 'デザイン原則の要約',
    points: [
      '意図的な余白で呼吸感をつくる',
      'タイポグラフィの階層で視線を制御する',
      'アクセントカラーは一点に集約する'
    ],
    notes: 'コンテンツスライドでの箇条書きと余白のバランスを検証します。'
  }, overrides || {});
}
