/**
 * MVP dataset factory for bullet cards slide component.
 */
function createBulletCardsSlideData(overrides = {}) {
  return Object.assign({
    type: 'bulletCards',
    title: '導入のメリット',
    items: [
      { title: 'ブランド整合性', desc: 'Appleらしいトーンを即座に再現できるプリセット。' },
      { title: '制作スピード', desc: '30分以内で10枚の資料を組み上げるワークフロー。' },
      { title: '再利用性', desc: 'トークン変更で他案件にも横展開しやすい設計。' }
    ],
    notes: 'カード型箇条書きの高さとテキストスタイルを検証します。'
  }, overrides || {});
}
