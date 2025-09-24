/**
 * Theme toggling utilities for light/dark support.
 */
const THEME_PROPERTY_KEY = 'themeMode';
const DEFAULT_THEME_KEY = 'light';

const PRESENTATION_THEMES = {
  light: {
    key: 'light',
    label: 'ライト',
    colors: {
      canvas: '#FFFFFF',
      primary_color: '#4285F4',
      text_primary: '#333333',
      background_white: '#FFFFFF',
      background_gray: '#f8f9fa',
      faint_gray: '#e8eaed',
      lane_title_bg: '#f8f9fa',
      table_header_bg: '#f8f9fa',
      lane_border: '#dadce0',
      card_bg: '#ffffff',
      card_border: '#dadce0',
      neutral_gray: '#9e9e9e',
      ghost_gray: '#efefed',
      text_on_primary: '#FFFFFF'
    }
  },
  dark: {
    key: 'dark',
    label: 'ダーク',
    colors: {
      canvas: '#121212',
      primary_color: '#8ab4f8',
      text_primary: '#e8eaed',
      background_white: '#FFFFFF',
      background_gray: '#1c1d1f',
      faint_gray: '#2a2b2f',
      lane_title_bg: '#1e1f22',
      table_header_bg: '#1e1f22',
      lane_border: '#3c4043',
      card_bg: '#1f2023',
      card_border: '#3c4043',
      neutral_gray: '#bdc1c6',
      ghost_gray: '#2a2b2d',
      text_on_primary: '#FFFFFF'
    }
  }
};

const THEME_ROTATION = { light: 'dark', dark: 'light' };

function sanitizeThemeKey(raw) {
  if (!raw) return DEFAULT_THEME_KEY;
  const key = String(raw).toLowerCase();
  return PRESENTATION_THEMES[key] ? key : DEFAULT_THEME_KEY;
}

function getStoredThemeKey() {
  try {
    return PropertiesService.getScriptProperties().getProperty(THEME_PROPERTY_KEY);
  } catch (e) {
    return DEFAULT_THEME_KEY;
  }
}

function applyThemePalette(themeKey) {
  const resolved = sanitizeThemeKey(themeKey);
  const palette = PRESENTATION_THEMES[resolved].colors;
  const target = CONFIG.COLORS;
  Object.keys(palette).forEach(key => {
    target[key] = palette[key];
  });
  CONFIG.__activeTheme = resolved;
  return resolved;
}

function ensureTheme(themeKey) {
  const resolved = applyThemePalette(themeKey || getStoredThemeKey());
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(THEME_PROPERTY_KEY) !== resolved) {
    props.setProperty(THEME_PROPERTY_KEY, resolved);
  }
  return resolved;
}

function getActiveThemeKey() {
  if (CONFIG.__activeTheme) {
    return CONFIG.__activeTheme;
  }
  return sanitizeThemeKey(getStoredThemeKey());
}

function getThemeToggleMenuLabel() {
  const active = getActiveThemeKey();
  const next = THEME_ROTATION[active] || DEFAULT_THEME_KEY;
  const activeLabel = PRESENTATION_THEMES[active].label;
  const nextLabel = PRESENTATION_THEMES[next].label;
  return `🌗 テーマ切替（現在: ${activeLabel} → ${nextLabel}）`;
}

function applyThemeForGeneration(themeKey) {
  return ensureTheme(themeKey);
}

function toggleTheme() {
  const current = getActiveThemeKey();
  const next = THEME_ROTATION[current] || DEFAULT_THEME_KEY;
  ensureTheme(next);
  const ui = SlidesApp.getUi();
  ui.alert(`テーマを「${PRESENTATION_THEMES[next].label}」に切り替えました。スライド生成時に新しいテーマが適用されます。`);
}
