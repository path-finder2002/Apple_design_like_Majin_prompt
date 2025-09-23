/**
 * MVP dataset factory for compare slide component.
 */
function createCompareSlideData(overrides = {}) {
  return Object.assign({
    type: 'compare',
    title: 'プラン比較',
    leftTitle: 'Standard プラン',
    rightTitle: 'Pro プラン',
    leftItems: [
      '基本的なAppleスタイルテーマ',
      '最大10枚の自動生成',
      'アクセントカラー固定'
    ],
    rightItems: [
      '拡張テーマとトークン編集',
      '最大30枚＆スプレッドシート連携',
      'アクセントカラーを自由指定'
    ],
    notes: '対比レイアウトの余白とヘッダーバーのスタイルを検証します。'
  }, overrides || {});
}
