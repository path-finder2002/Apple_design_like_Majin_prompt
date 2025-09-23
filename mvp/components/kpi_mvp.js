/**
 * MVP dataset factory for KPI slide component.
 */
function createKpiSlideData(overrides = {}) {
  return Object.assign({
    type: 'kpi',
    title: '注力すべきKPI',
    items: [
      { label: 'Weekly Active Users', value: '1.2M', change: '+12% vs LW', status: 'good' },
      { label: 'Conversion Rate', value: '8.4%', change: '+1.1pt', status: 'good' },
      { label: 'Time to Deck', value: '18min', change: '-7min vs baseline', status: 'good' }
    ],
    notes: 'KPIカードの縦配置と色味を確認するMVPです。'
  }, overrides || {});
}
