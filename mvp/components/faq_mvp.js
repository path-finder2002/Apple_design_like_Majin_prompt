/**
 * MVP dataset factory for FAQ slide component.
 */
function createFaqSlideData(overrides = {}) {
  return Object.assign({
    type: 'faq',
    title: 'よくある質問',
    items: [
      { q: '生成されたスライドは編集できますか？', a: 'はい、生成後は通常のSlides編集と同じように変更できます。' },
      { q: 'ライト/ダークテーマの切替は可能ですか？', a: '将来的な対応を見据え、設定メニューでトグル予定です。' },
      { q: '画像は自動で挿入されますか？', a: 'MVPではプレースホルダーでの対応となります。' }
    ],
    notes: 'FAQ配置と質疑応答のフォーマットを確認します。'
  }, overrides || {});
}
