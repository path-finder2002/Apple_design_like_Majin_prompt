/**
 * MVP dataset factory for process slide component.
 */
function createProcessSlideData(overrides = {}) {
  return Object.assign({
    type: 'process',
    title: 'デザインプロセス',
    subhead: 'Discovery から Launch までの流れ',
    steps: [
      'Discovery スプリント',
      'ビジュアル方向性レビュー',
      'トークン適用と実装',
      'ユーザーテストとローンチ準備'
    ],
    notes: '縦プロセスの矢印と番号スタイルを検証します。'
  }, overrides || {});
}
