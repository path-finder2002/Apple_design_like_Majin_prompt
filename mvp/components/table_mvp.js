/**
 * MVP dataset factory for table slide component.
 */
function createTableSlideData(overrides = {}) {
  return Object.assign({
    type: 'table',
    title: 'ロードマップ進捗表',
    subhead: '各フェーズの状態とアクション',
    headers: ['フェーズ', 'ステータス', '次のアクション'],
    rows: [
      ['Discovery', '完了', '成果の共有資料を整える'],
      ['Design', '進行中', 'トークン適用の最終調整'],
      ['Pilot', '未着手', 'パイロットユーザー選定']
    ],
    notes: '表レイアウトでのヘッダー色とセル整列をチェックします。'
  }, overrides || {});
}
