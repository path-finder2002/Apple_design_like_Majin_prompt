/**
 * MVP dataset factory for header cards slide component.
 */
function createHeaderCardsSlideData(overrides = {}) {
  return Object.assign({
    type: 'headerCards',
    title: '導入効果',
    subhead: 'ステークホルダー別の価値訴求',
    items: [
      { title: '経営陣', desc: 'ブランドトーンを短時間で整え、意思決定を加速。' },
      { title: 'PM', desc: 'テンプレート化でメッセージ整理が容易に。' },
      { title: 'デザイナー', desc: '細部の調整に集中でき、反復作業を削減。' }
    ],
    notes: 'ヘッダー一体型カードの色とテキスト余白を検証します。'
  }, overrides || {});
}
