const THEME_STORAGE_KEY = 'theme';
const DEFAULT_THEME_KEY = 'light';
const FILL_COLOR_PRIORITY = [
  'background_white',
  'background_gray',
  'canvas',
  'card_bg',
  'faint_gray',
  'lane_title_bg',
  'table_header_bg',
  'card_border',
  'lane_border',
  'neutral_gray',
  'ghost_gray',
  'primary_color',
  'text_on_primary',
  'text_primary'
];
const TEXT_COLOR_PRIORITY = [
  'text_primary',
  'neutral_gray',
  'ghost_gray',
  'text_on_primary',
  'primary_color'
];
const STROKE_COLOR_PRIORITY = [
  'card_border',
  'lane_border',
  'neutral_gray',
  'primary_color',
  'text_on_primary'
];

let ACTIVE_THEME_KEY = null;

function ensureTheme() {
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty(THEME_STORAGE_KEY);
  const themeKey = CONFIG.THEMES[stored] ? stored : DEFAULT_THEME_KEY;
  const appliedTheme = applyThemeToConfig(themeKey);
  if (stored !== themeKey) {
    props.setProperty(THEME_STORAGE_KEY, themeKey);
  }
  ACTIVE_THEME_KEY = themeKey;
  return appliedTheme;
}

function toggleTheme() {
  const props = PropertiesService.getScriptProperties();
  const currentKey = getActiveThemeKey();
  const nextKey = currentKey === 'dark' ? 'light' : 'dark';
  const oldTheme = getThemeWithOverrides(currentKey);
  const newTheme = applyThemeToConfig(nextKey);
  props.setProperty(THEME_STORAGE_KEY, nextKey);
  ACTIVE_THEME_KEY = nextKey;

  try {
    refreshPresentationTheme(oldTheme, newTheme);
    logInfo('toggleTheme:completed', { from: currentKey, to: nextKey });
  } catch (error) {
    logError('toggleTheme:refresh_failed', error);
  }

  try {
    SlidesApp.getUi().alert(`テーマを${newTheme.displayName || newTheme.key}に切り替えました。`);
  } catch (uiError) {
    logError('toggleTheme:ui_alert_failed', uiError);
  }
}

function getThemeToggleMenuLabel() {
  const key = getActiveThemeKey();
  return key === 'dark' ? '🌞 ライトモードに切り替え' : '🌙 ダークモードに切り替え';
}

function getActiveThemeKey() {
  if (ACTIVE_THEME_KEY) {
    return ACTIVE_THEME_KEY;
  }
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty(THEME_STORAGE_KEY);
  if (stored && CONFIG.THEMES[stored]) {
    ACTIVE_THEME_KEY = stored;
    return stored;
  }
  ACTIVE_THEME_KEY = DEFAULT_THEME_KEY;
  return ACTIVE_THEME_KEY;
}

function applyThemeToConfig(themeKey) {
  const theme = getThemeWithOverrides(themeKey);
  const targetColors = theme.colors || {};
  Object.keys(CONFIG.COLORS).forEach((colorKey) => {
    if (Object.prototype.hasOwnProperty.call(targetColors, colorKey)) {
      CONFIG.COLORS[colorKey] = targetColors[colorKey];
    }
  });
  return theme;
}

function getThemeWithOverrides(themeKey) {
  const baseTheme = CONFIG.THEMES[themeKey] || CONFIG.THEMES[DEFAULT_THEME_KEY];
  const clone = {
    key: baseTheme.key,
    displayName: baseTheme.displayName,
    colors: Object.assign({}, baseTheme.colors)
  };

  const props = PropertiesService.getScriptProperties();
  const customPrimary = normalizeHexColor(props.getProperty('primaryColor'));
  if (customPrimary) {
    clone.colors.primary_color = customPrimary;
  }

  return clone;
}

function refreshPresentationTheme(oldTheme, newTheme) {
  let presentation;
  try {
    presentation = SlidesApp.getActivePresentation();
  } catch (error) {
    logError('refreshPresentationTheme:active_presentation_failed', error);
    return;
  }

  if (!presentation) {
    logError('refreshPresentationTheme:no_presentation', null);
    return;
  }

  const slides = presentation.getSlides();
  const colorIndex = buildThemeColorIndex(oldTheme);

  slides.forEach((slide) => {
    try {
      updateSlideElements(slide, oldTheme, newTheme, colorIndex);
    } catch (error) {
      logError('refreshPresentationTheme:update_slide_failed', error);
    }
  });

  logInfo('refreshPresentationTheme:completed', { slides: slides.length, theme: newTheme.key });
}

function buildThemeColorIndex(theme) {
  const index = {};
  if (!theme || !theme.colors) return index;

  Object.keys(theme.colors).forEach((key) => {
    const normalized = normalizeHexColor(theme.colors[key]);
    if (!normalized) return;
    if (!index[normalized]) {
      index[normalized] = [];
    }
    index[normalized].push(key);
  });

  return index;
}

function updateSlideElements(slide, oldTheme, newTheme, colorIndex) {
  try {
    slide.getBackground().setSolidFill(newTheme.colors.background_white);
  } catch (error) {
    logError('updateSlideElements:background_failed', error);
  }

  const elements = slide.getPageElements();
  for (let i = 0; i < elements.length; i++) {
    refreshPageElement(elements[i], oldTheme, newTheme, colorIndex);
  }
}

function refreshPageElement(element, oldTheme, newTheme, colorIndex) {
  try {
    switch (element.getPageElementType()) {
      case SlidesApp.PageElementType.SHAPE:
        refreshShapeElement(element.asShape(), oldTheme, newTheme, colorIndex);
        break;
      case SlidesApp.PageElementType.TABLE:
        refreshTableElement(element.asTable(), oldTheme, newTheme, colorIndex);
        break;
      case SlidesApp.PageElementType.GROUP:
        const groupChildren = element.asGroup().getChildren();
        groupChildren.forEach((child) => refreshPageElement(child, oldTheme, newTheme, colorIndex));
        break;
      case SlidesApp.PageElementType.LINE:
        applyThemeToLineFill(element.asLine().getLineFill(), oldTheme, newTheme, colorIndex);
        break;
      default:
        break;
    }
  } catch (error) {
    logError('refreshPageElement:failed', { type: element.getPageElementType(), error: error.message });
  }
}

function refreshShapeElement(shape, oldTheme, newTheme, colorIndex) {
  try {
    applyThemeToFill(shape.getFill(), oldTheme, newTheme, colorIndex);
  } catch (error) {
    logError('refreshShapeElement:fill_failed', error);
  }

  try {
    const border = shape.getBorder();
    if (border) {
      applyThemeToLineFill(border.getLineFill(), oldTheme, newTheme, colorIndex);
    }
  } catch (error) {
    logError('refreshShapeElement:border_failed', error);
  }

  try {
    const textRange = shape.getText();
    if (textRange && textRange.getLength) {
      applyThemeToTextRange(textRange, oldTheme, newTheme, colorIndex);
    }
  } catch (error) {
    logError('refreshShapeElement:text_failed', error);
  }
}

function refreshTableElement(table, oldTheme, newTheme, colorIndex) {
  try {
    applyThemeToFill(table.getFill(), oldTheme, newTheme, colorIndex, 'background_white');
  } catch (error) {
    logError('refreshTableElement:table_fill_failed', error);
  }

  const rows = table.getNumRows();
  const cols = table.getNumColumns();

  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const cell = table.getCell(r, c);
      if (!cell) continue;
      try {
        applyThemeToFill(cell.getFill(), oldTheme, newTheme, colorIndex);
      } catch (error) {
        logError('refreshTableElement:cell_fill_failed', error);
      }

      try {
        applyThemeToTextRange(cell.getText(), oldTheme, newTheme, colorIndex);
      } catch (error) {
        logError('refreshTableElement:cell_text_failed', error);
      }

      try {
        const border = cell.getBorder();
        if (border) {
          applyThemeToLineFill(border.getLineFill(), oldTheme, newTheme, colorIndex);
        }
      } catch (error) {
        logError('refreshTableElement:cell_border_failed', error);
      }
    }
  }
}

function applyThemeToFill(fill, oldTheme, newTheme, colorIndex, fallbackKey) {
  if (!fill) return;
  let targetColor = null;
  const hex = getHexFromFill(fill);
  if (hex) {
    targetColor = mapThemeColor(hex, oldTheme, newTheme, colorIndex, 'fill');
  }
  if (!targetColor && fallbackKey && newTheme.colors[fallbackKey]) {
    targetColor = newTheme.colors[fallbackKey];
  }
  if (targetColor && (!hex || targetColor !== hex)) {
    try {
      fill.setSolidFill(targetColor);
    } catch (error) {
      logError('applyThemeToFill:setSolidFill_failed', error);
    }
  }
}

function applyThemeToLineFill(lineFill, oldTheme, newTheme, colorIndex) {
  if (!lineFill) return;
  const hex = getHexFromFill(lineFill);
  const targetColor = hex ? mapThemeColor(hex, oldTheme, newTheme, colorIndex, 'stroke') : null;
  if (targetColor && (!hex || targetColor !== hex)) {
    try {
      lineFill.setSolidFill(targetColor);
    } catch (error) {
      logError('applyThemeToLineFill:setSolidFill_failed', error);
    }
  }
}

function applyThemeToTextRange(textRange, oldTheme, newTheme, colorIndex) {
  if (!textRange || typeof textRange.getLength !== 'function') return;
  applyThemeToSingleTextRange(textRange, oldTheme, newTheme, colorIndex);

  const length = textRange.getLength();
  if (!length || length <= 1) return;

  for (let i = 0; i < length; i++) {
    try {
      const segment = textRange.getRange(i, i);
      applyThemeToSingleTextRange(segment, oldTheme, newTheme, colorIndex);
    } catch (error) {
      logError('applyThemeToTextRange:segment_failed', error);
      return;
    }
  }
}

function applyThemeToSingleTextRange(range, oldTheme, newTheme, colorIndex) {
  if (!range) return;
  let style;
  try {
    style = range.getTextStyle();
  } catch (error) {
    logError('applyThemeToSingleTextRange:style_failed', error);
    return;
  }

  if (!style) return;

  let hex = null;
  try {
    hex = getHexFromTextStyle(style);
  } catch (error) {
    logError('applyThemeToSingleTextRange:get_hex_failed', error);
  }

  const mapped = hex ? mapThemeColor(hex, oldTheme, newTheme, colorIndex, 'text') : null;
  if (mapped) {
    if (mapped !== hex) {
      style.setForegroundColor(mapped);
    }
    return;
  }

  if (!hex) {
    style.setForegroundColor(newTheme.colors.text_primary);
    return;
  }

  const normalizedTextPrimary = normalizeHexColor(oldTheme.colors.text_primary);
  if (normalizedTextPrimary && normalizedTextPrimary === normalizeHexColor(hex)) {
    style.setForegroundColor(newTheme.colors.text_primary);
  }
}

function mapThemeColor(hex, oldTheme, newTheme, colorIndex, context) {
  if (!oldTheme || !oldTheme.colors || !newTheme || !newTheme.colors) return null;
  const normalized = normalizeHexColor(hex);
  if (!normalized) return null;

  const keys = colorIndex[normalized];
  if (keys && keys.length) {
    const resolvedKey = chooseColorKeyForContext(keys, context);
    if (resolvedKey && newTheme.colors[resolvedKey]) {
      return newTheme.colors[resolvedKey];
    }
  }

  const oldPrimary = normalizeHexColor(oldTheme.colors.primary_color);
  if (oldPrimary && isWithinColorTolerance(normalized, oldPrimary)) {
    return transformDerivedColor(normalized, oldTheme, newTheme);
  }

  return null;
}

function chooseColorKeyForContext(keys, context) {
  const priority = context === 'text'
    ? TEXT_COLOR_PRIORITY
    : context === 'stroke'
      ? STROKE_COLOR_PRIORITY
      : FILL_COLOR_PRIORITY;

  for (let i = 0; i < priority.length; i++) {
    if (keys.indexOf(priority[i]) !== -1) {
      return priority[i];
    }
  }

  return keys[0];
}

function transformDerivedColor(derivedHex, oldTheme, newTheme) {
  if (!oldTheme || !oldTheme.colors || !newTheme || !newTheme.colors) return null;
  const delta = inferBrightnessDelta(derivedHex, oldTheme.colors.primary_color);
  if (!delta) return null;
  const baseRgb = hexToRgb(newTheme.colors.primary_color);
  if (!baseRgb) return null;

  const baseHsl = rgbToHsl(baseRgb);
  const targetHsl = {
    h: baseHsl.h,
    s: clamp(baseHsl.s + delta.saturation, 0, 1),
    l: clamp(baseHsl.l + delta.lightness, 0, 1)
  };

  return rgbToHex(hslToRgb(targetHsl));
}

function inferBrightnessDelta(targetHex, baseHex) {
  if (!targetHex || !baseHex) return null;
  const targetRgb = hexToRgb(targetHex);
  const baseRgb = hexToRgb(baseHex);
  if (!targetRgb || !baseRgb) return null;

  const targetHsl = rgbToHsl(targetRgb);
  const baseHsl = rgbToHsl(baseRgb);
  const hueDiff = Math.min(Math.abs(targetHsl.h - baseHsl.h), 1 - Math.abs(targetHsl.h - baseHsl.h));
  if (hueDiff > 0.08) return null;

  return {
    lightness: targetHsl.l - baseHsl.l,
    saturation: targetHsl.s - baseHsl.s
  };
}

function isWithinColorTolerance(targetHex, baseHex) {
  const targetRgb = hexToRgb(targetHex);
  const baseRgb = hexToRgb(baseHex);
  if (!targetRgb || !baseRgb) return false;

  const targetHsl = rgbToHsl(targetRgb);
  const baseHsl = rgbToHsl(baseRgb);
  const hueDiff = Math.min(Math.abs(targetHsl.h - baseHsl.h), 1 - Math.abs(targetHsl.h - baseHsl.h));
  const satDiff = Math.abs(targetHsl.s - baseHsl.s);
  return hueDiff <= 0.05 && satDiff <= 0.25;
}

function getHexFromFill(fill) {
  try {
    if (!fill || fill.getType() !== SlidesApp.FillType.SOLID) return null;
    const solid = fill.getSolidFill();
    if (!solid) return null;
    return getHexFromColor(solid.getColor());
  } catch (error) {
    logError('getHexFromFill:failed', error);
    return null;
  }
}

function getHexFromTextStyle(style) {
  try {
    return getHexFromColor(style.getForegroundColor());
  } catch (error) {
    logError('getHexFromTextStyle:failed', error);
    return null;
  }
}

function getHexFromColor(color) {
  if (!color) return null;
  try {
    const type = color.getColorType();
    if (type === SlidesApp.ColorType.RGB) {
      return normalizeHexColor(color.asRgbColor().asHexString());
    }
    return null;
  } catch (error) {
    logError('getHexFromColor:failed', error);
    return null;
  }
}

function normalizeHexColor(color) {
  if (!color || typeof color !== 'string') return null;
  let value = color.trim();
  if (!value) return null;
  if (value[0] === '#') value = value.substring(1);
  if (value.length === 3) {
    value = value.split('').map((ch) => ch + ch).join('');
  }
  if (!/^[0-9a-fA-F]{6}$/.test(value)) return null;
  return `#${value.toUpperCase()}`;
}

function hexToRgb(hex) {
  const normalized = normalizeHexColor(hex);
  if (!normalized) return null;
  const value = normalized.substring(1);
  return {
    r: parseInt(value.substring(0, 2), 16),
    g: parseInt(value.substring(2, 4), 16),
    b: parseInt(value.substring(4, 6), 16)
  };
}

function rgbToHex(rgb) {
  if (!rgb) return null;
  const toHex = (component) => {
    const clamped = clamp(Math.round(component), 0, 255);
    return clamped.toString(16).padStart(2, '0');
  };
  return `#${toHex(rgb.r)}${toHex(rgb.g)}${toHex(rgb.b)}`.toUpperCase();
}

function rgbToHsl(rgb) {
  const r = rgb.r / 255;
  const g = rgb.g / 255;
  const b = rgb.b / 255;

  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let h = 0;
  let s = 0;
  const l = (max + min) / 2;

  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    switch (max) {
      case r:
        h = (g - b) / d + (g < b ? 6 : 0);
        break;
      case g:
        h = (b - r) / d + 2;
        break;
      case b:
        h = (r - g) / d + 4;
        break;
      default:
        break;
    }
    h /= 6;
  }

  return { h, s, l };
}

function hslToRgb(hsl) {
  const hue2rgb = (p, q, t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };

  const l = clamp(hsl.l, 0, 1);
  const s = clamp(hsl.s, 0, 1);
  let r;
  let g;
  let b;

  if (s === 0) {
    r = g = b = l; // achromatic
  } else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r = hue2rgb(p, q, hsl.h + 1 / 3);
    g = hue2rgb(p, q, hsl.h);
    b = hue2rgb(p, q, hsl.h - 1 / 3);
  }

  return {
    r: Math.round(r * 255),
    g: Math.round(g * 255),
    b: Math.round(b * 255)
  };
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}
