/**
 * MVP dataset factory for cards slide component.
 */
function createCardsSlideData(overrides = {}) {
  return Object.assign({
    type: 'cards',
    title: '特徴のハイライト',
    subhead: 'Appleらしい余白と影響力のあるキーメッセージ',
    items: [
      { title: 'Minimal', desc: '不要な装飾を排し、余白で魅せる構成。' },
      { title: 'Focused', desc: '一枚一メッセージで訴求ポイントを明確化。' },
      { title: 'Consistent', desc: 'トークン化したスタイルで再現性を担保。' },
      { title: 'Adaptive', desc: 'ライト／ダークテーマを簡単に切替可能。' }
    ],
    notes: 'シンプルカードの列バランスとテキストスタイルを確認します。'
  }, overrides || {});
}
