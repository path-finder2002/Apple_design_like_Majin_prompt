/**
 * MVP dataset factory for closing slide component.
 */
function createClosingSlideData(overrides = {}) {
  return Object.assign({
    type: 'closing',
    notes: 'MVPのご確認ありがとうございました。フィードバックをお待ちしています。'
  }, overrides || {});
}
