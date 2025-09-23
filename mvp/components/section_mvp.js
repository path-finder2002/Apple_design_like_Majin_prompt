/**
 * MVP dataset factory for section slide component.
 */
function createSectionSlideData(overrides = {}) {
  return Object.assign({
    type: 'section',
    title: '01. 体験ビジョン',
    notes: 'セクションスライドの余白とゴースト番号を確認する最小構成です。'
  }, overrides || {});
}
