/**
 * MVP dataset factory for title slide component.
 */
function createTitleSlideData(overrides = {}) {
  return Object.assign({
    type: 'title',
    title: 'Apple Design MVP',
    date: '2025.09.23',
    notes: 'AppleテイストのジェネレーターMVPをカバー単体で検証するためのサンプルです。'
  }, overrides || {});
}
