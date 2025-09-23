/**
 * MVP dataset factory for quote slide component.
 */
function createQuoteSlideData(overrides = {}) {
  return Object.assign({
    type: 'quote',
    title: '体験を導く言葉',
    text: 'デザインは見た目だけでなく、どのように機能するかだ。',
    author: 'Steve Jobs',
    notes: '引用スライドの余白とアクセントのバランス確認用です。'
  }, overrides || {});
}
