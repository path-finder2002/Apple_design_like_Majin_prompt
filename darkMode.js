// テーマ設定とダークモード切り替えロジック
const THEMES = {
  light: {
    key: 'light',
    label: 'ライトテーマ',
    colors: {
      canvas: '#FFFFFF',
      primary_color: '#4285F4',
      text_primary: '#000000',
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
    label: 'ダークテーマ',
    colors: {
      canvas: '#202124',
      primary_color: '#8AB4F8',
      text_primary: '#FFFFFF',
      background_white: '#1f1f1f',
      background_gray: '#2b2c2f',
      faint_gray: '#3c4043',
      lane_title_bg: '#2b2c2f',
      table_header_bg: '#2b2c2f',
      lane_border: '#44464a',
      card_bg: '#1f1f1f',
      card_border: '#3c4043',
      neutral_gray: '#b0b0b0',
      ghost_gray: '#2f3032',
      text_on_primary: '#202124'
    }
  }
};

let __ACTIVE_THEME = THEMES.light.key;

function applyTheme(themeKey) {
  const target = THEMES[themeKey] || THEMES.light;
  Object.keys(target.colors).forEach(key => {
    CONFIG.COLORS[key] = target.colors[key];
  });
  __ACTIVE_THEME = target.key;
  return target.key;
}

function applyThemeForGeneration(themeMode) {
  const requested = themeMode && THEMES[themeMode] ? themeMode : THEMES.light.key;
  return applyTheme(requested);
}

function ensureTheme() {
  const props = PropertiesService.getScriptProperties();
  const storedTheme = props.getProperty('themeMode');
  return applyThemeForGeneration(storedTheme);
}

function getActiveTheme() {
  return __ACTIVE_THEME;
}

function getThemeToggleMenuLabel() {
  return getActiveTheme() === THEMES.dark.key ? '☀️ ライトテーマに切替' : '🌙 ダークテーマに切替';
}

function toggleTheme() {
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty('themeMode');
  const current = THEMES[stored] ? stored : getActiveTheme();
  const next = current === THEMES.dark.key ? THEMES.light.key : THEMES.dark.key;

  props.setProperty('themeMode', next);
  applyTheme(next);

  try {
    refreshPresentationTheme(current, next);
  } catch (e) {
    logError('toggleTheme:refreshFailed', e);
  }

  try {
    onOpen();
  } catch (e) {
    logError('toggleTheme:onOpenRefreshFailed', e);
  }

  SlidesApp.getUi().alert(`${THEMES[next].label}に切り替えました。`);
  logInfo('toggleTheme:updated', { from: current, to: next });
}

function refreshPresentationTheme(fromThemeKey, toThemeKey) {
  const fromTheme = THEMES[fromThemeKey] || THEMES.light;
  const toTheme = THEMES[toThemeKey] || THEMES.light;
  const presentation = SETTINGS.TARGET_PRESENTATION_ID
    ? SlidesApp.openById(SETTINGS.TARGET_PRESENTATION_ID)
    : SlidesApp.getActivePresentation();

  if (!presentation) return;

  const fromColorIndex = buildThemeColorIndex(fromTheme.colors);

  presentation.getSlides().forEach(slide => {
    updateSlideBackground(slide, fromColorIndex, toTheme.colors);
    updateSlideElements(slide, fromTheme, toTheme, fromColorIndex);
  });

  logInfo('refreshPresentationTheme:completed', { from: fromTheme.key, to: toTheme.key });
}

function buildThemeColorIndex(colors) {
  const index = {};
  Object.keys(colors).forEach(key => {
    index[normalizeHex(colors[key])] = key;
  });
  return index;
}

function updateSlideBackground(slide, fromColorIndex, toColors) {
  try {
    const background = slide.getBackground();
    const fill = background.getFill();
    const hex = getFillHex(fill);
    const colorKey = hex ? fromColorIndex[hex] : null;
    if (colorKey && toColors[colorKey]) {
      background.setSolidFill(toColors[colorKey]);
    }
  } catch (e) {
    logError('updateSlideBackground', e);
  }
}

function updateSlideElements(slide, fromTheme, toTheme, fromColorIndex) {
  const elements = slide.getPageElements();
  elements.forEach(element => updatePageElement(element, fromTheme, toTheme, fromColorIndex));
}

function updatePageElement(element, fromTheme, toTheme, fromColorIndex) {
  try {
    switch (element.getPageElementType()) {
      case SlidesApp.PageElementType.SHAPE:
        refreshShapeElement(element.asShape(), fromTheme, toTheme, fromColorIndex);
        break;
      case SlidesApp.PageElementType.TABLE:
        refreshTableElement(element.asTable(), fromTheme, toTheme, fromColorIndex);
        break;
      case SlidesApp.PageElementType.GROUP:
        element.asGroup().getChildren().forEach(child => updatePageElement(child, fromTheme, toTheme, fromColorIndex));
        break;
      default:
        break;
    }
  } catch (e) {
    logError('updatePageElement', e);
  }
}

function refreshShapeElement(shape, fromTheme, toTheme, fromColorIndex) {
  try {
    const fill = shape.getFill();
    updateFillColor(fill, fromTheme, toTheme, fromColorIndex);
  } catch (e) {
    // 一部の図形では Fill 情報が取得できないケースがあるため、ログのみ出力
    logError('refreshShapeElement:fill', e);
  }

  try {
    const text = shape.getText ? shape.getText() : null;
    if (text && text.getLength() > 0) {
      applyThemeToTextRange(text, fromTheme, toTheme, fromColorIndex);
    }
  } catch (e) {
    logError('refreshShapeElement:text', e);
  }
}

function refreshTableElement(table, fromTheme, toTheme, fromColorIndex) {
  for (let r = 0; r < table.getNumRows(); r++) {
    for (let c = 0; c < table.getNumColumns(); c++) {
      const cell = table.getCell(r, c);
      if (!cell) continue;
      updateFillColor(cell.getFill(), fromTheme, toTheme, fromColorIndex);
      applyThemeToTextRange(cell.getText(), fromTheme, toTheme, fromColorIndex);
    }
  }
}

function updateFillColor(fill, fromTheme, toTheme, fromColorIndex) {
  if (!fill || fill.getType() !== SlidesApp.FillType.SOLID) return;
  try {
    const hex = getFillHex(fill);
    const colorKey = hex ? fromColorIndex[hex] : null;
    if (colorKey && toTheme.colors[colorKey]) {
      fill.setSolidFill(toTheme.colors[colorKey]);
      return;
    }

    const derivedColor = transformDerivedColor(hex, fromTheme, toTheme);
    if (derivedColor) {
      fill.setSolidFill(derivedColor);
    }
  } catch (e) {
    logError('updateFillColor', e);
  }
}

function applyThemeToTextRange(textRange, fromTheme, toTheme, fromColorIndex) {
  if (!textRange) return;
  const runs = textRange.getRuns();
  if (runs && runs.length > 0) {
    runs.forEach(run => applyThemeToSingleTextRange(run, fromTheme, toTheme, fromColorIndex));
  } else {
    applyThemeToSingleTextRange(textRange, fromTheme, toTheme, fromColorIndex);
  }
}

function applyThemeToSingleTextRange(range, fromTheme, toTheme, fromColorIndex) {
  const style = range.getTextStyle();
  if (!style) return;
  try {
    if (!style.isForegroundColorSet()) return;
    const color = style.getForegroundColor();
    if (!color) return;
    const rgb = color.asRgbColor();
    if (!rgb) return;
    const hex = normalizeHex(rgb.asHexString());
    const colorKey = fromColorIndex[hex];
    if (colorKey && toTheme.colors[colorKey]) {
      style.setForegroundColor(toTheme.colors[colorKey]);
      return;
    }

    const derivedColor = transformDerivedColor(hex, fromTheme, toTheme);
    if (derivedColor) {
      style.setForegroundColor(derivedColor);
    }
  } catch (e) {
    logError('applyThemeToSingleTextRange', e);
  }
}

function transformDerivedColor(hex, fromTheme, toTheme) {
  if (!hex) return null;
  const primary = normalizeHex(fromTheme.colors.primary_color);
  const factor = inferBrightnessFactor(hex, primary);
  if (factor === null) return null;
  return adjustColorBrightness(toTheme.colors.primary_color, factor);
}

function inferBrightnessFactor(targetHex, baseHex) {
  if (!targetHex || !baseHex) return null;
  const base = hexToRgb(baseHex);
  const rgb = hexToRgb(targetHex);
  const ratios = [];

  ['r', 'g', 'b'].forEach(channel => {
    const baseValue = base[channel];
    const rgbValue = rgb[channel];
    if (baseValue === 0) {
      if (rgbValue === 0) return;
      ratios.push(null);
      return;
    }
    if (rgbValue === 255) {
      // 飽和している成分は比率計算に使用しない
      return;
    }
    ratios.push(rgbValue / baseValue);
  });

  const validRatios = ratios.filter(v => typeof v === 'number' && isFinite(v));
  if (!validRatios.length) return null;

  const factor = validRatios.reduce((sum, value) => sum + value, 0) / validRatios.length;
  if (!isFinite(factor) || factor <= 0.1 || factor > 3) return null;

  const reconstructed = adjustColorBrightness(baseHex, factor);
  if (!isWithinColorTolerance(reconstructed, targetHex)) return null;

  return Number(factor.toFixed(2));
}

function isWithinColorTolerance(colorA, colorB) {
  const a = hexToRgb(colorA);
  const b = hexToRgb(colorB);
  const distance = Math.abs(a.r - b.r) + Math.abs(a.g - b.g) + Math.abs(a.b - b.b);
  return distance <= 6; // 小さな丸め誤差を許容
}

function getFillHex(fill) {
  if (!fill || fill.getType() !== SlidesApp.FillType.SOLID) return null;
  try {
    const color = fill.getColor();
    if (!color) return null;
    const rgb = color.asRgbColor();
    if (!rgb) return null;
    return normalizeHex(rgb.asHexString());
  } catch (e) {
    logError('getFillHex', e);
    return null;
  }
}

function normalizeHex(hex) {
  if (!hex) return null;
  let value = hex.trim();
  if (value.charAt(0) !== '#') value = `#${value}`;
  return value.toUpperCase();
}

function hexToRgb(hex) {
  const value = normalizeHex(hex).replace('#', '');
  return {
    r: parseInt(value.substring(0, 2), 16),
    g: parseInt(value.substring(2, 4), 16),
    b: parseInt(value.substring(4, 6), 16),
  };
}
