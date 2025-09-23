/**
 * MVP dataset factory for timeline slide component.
 */
function createTimelineSlideData(overrides = {}) {
  return Object.assign({
    type: 'timeline',
    title: 'リリースロードマップ',
    milestones: [
      { label: 'Kickoff', date: '2025 Q1' },
      { label: 'Design Freeze', date: '2025 Q2' },
      { label: 'Pilot', date: '2025 Q3' },
      { label: 'Launch', date: '2025 Q4' }
    ],
    notes: 'タイムラインの配色とラベル位置を確認するMVPです。'
  }, overrides || {});
}
