// THEME TOGGLER MODULE

/**
 * 注意:
 * - 既存要素の色指定はこのモジュールによって上書きされます。適用後はモード依存で色管理してください。
 * - 下記コメントに受け入れテスト手順を記載しています。運用はメニュー駆動のみを想定しています。
 *
 * 受け入れテスト:
 * 1) 既存のダーク配色スライドで「ライト適用」を実行し、背景・本文・罫線・表ヘッダがライトパレットへ更新されること。
 * 2) 続けて「ダーク適用」を実行し、背景が濃色 / 本文が白 / 罫線が中間灰 / 表ヘッダが暗所でも読める配色になること。
 * 3) 画像・ロゴ・配置・サイズが変わらないこと。
 * 4) 12pt 以下の文字が textSecondary（薄め）で読みやすくなること。
 * 5) 同モードを連続適用しても見た目が変わらないこと（冪等）。
 */

/**
 * パレット定義
 * @param {'light'|'dark'} mode
 * @return {{ background:string, textPrimary:string, textSecondary:string, accent:string, border:string, faintBg:string }}
 */
function getPalette(mode) {
  const light = {
    background: '#FFFFFF',
    textPrimary: '#111111',
    textSecondary: '#5F6368',
    accent: '#007AFF',
    border: '#DADCE0',
    faintBg: '#F5F7F9',
  };

  const dark = {
    background: '#0B0B0D',
    textPrimary: '#FFFFFF',
    textSecondary: '#C7C7CC',
    accent: '#0A84FF',
    border: '#3A3A3C',
    faintBg: '#1C1C1E',
  };

  return mode === 'dark' ? dark : light;
}

/**
 * プレゼン全体へテーマ適用
 * @param {'light'|'dark'} mode
 */
function applyThemeToPresentation(mode) {
  const palette = getPalette(mode);
  let pres;
  try {
    pres = SlidesApp.getActivePresentation();
  } catch (err) {
    Logger.log('[ThemeToggle] Failed to open presentation: ' + err);
    return;
  }
  if (!pres) {
    Logger.log('[ThemeToggle] Presentation not found.');
    return;
  }

  const slides = pres.getSlides();
  for (let i = 0; i < slides.length; i++) {
    try {
      recolorSlide(slides[i], palette);
    } catch (err) {
      Logger.log('[ThemeToggle] slide ' + (i + 1) + ' error: ' + err);
    }
  }
}

/**
 * 単一スライドの再配色
 * @param {GoogleAppsScript.Slides.Slide} slide
 * @param {{background:string,textPrimary:string,textSecondary:string,accent:string,border:string,faintBg:string}} palette
 */
function recolorSlide(slide, palette) {
  applyBackground(slide, palette.background);

  const elements = slide.getPageElements();
  for (let i = 0; i < elements.length; i++) {
    try {
      applyPaletteToElement(elements[i], palette);
    } catch (err) {
      Logger.log('[ThemeToggle] element error: ' + err);
    }
  }
}

/**
 * テキストカラーを安全に更新
 * @param {GoogleAppsScript.Slides.Shape} shape
 * @param {{textPrimary:string,textSecondary:string}} palette
 */
function safeSetTextColor(shape, palette) {
  if (!shape || typeof shape.getText !== 'function') return;
  let text;
  try {
    text = shape.getText();
  } catch (err) {
    return;
  }
  if (!text || text.getLength() === 0) return;

  const fontSize = inferFontSize(text);
  const targetColor = fontSize <= 12 ? palette.textSecondary : palette.textPrimary;
  const style = text.getTextStyle();
  if (!style) return;

  const current = safeGetForegroundColor(style);
  if (colorsEqual(current, targetColor)) return;

  style.setForegroundColor(targetColor);
}

/**
 * テーブルの再配色
 * @param {GoogleAppsScript.Slides.Table} table
 * @param {{textPrimary:string,textSecondary:string,border:string,faintBg:string}} palette
 */
function safeRecolorTable(table, palette) {
  const rows = table.getNumRows();
  const cols = table.getNumColumns();
  if (!rows || !cols) return;

  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const cell = table.getCell(r, c);
      if (!cell) continue;

      try {
        recolorTableCell(cell, palette, r === 0);
      } catch (err) {
        Logger.log('[ThemeToggle] table cell error: ' + err);
      }
    }
  }
}

/**
 * メニュー追加
 */
function onOpen(e) {
  const ui = SlidesApp.getUi();
  ui.createMenu('テーマ切替')
    .addItem('ライト適用', 'applyLightTheme')
    .addItem('ダーク適用', 'applyDarkTheme')
    .addToUi();
}

// --- メニューアクション ---

function applyLightTheme() {
  applyThemeToPresentation('light');
}

function applyDarkTheme() {
  applyThemeToPresentation('dark');
}

// --- 内部ユーティリティ ---

function applyBackground(slide, color) {
  try {
    const background = slide.getBackground();
    const fill = background.getFill();
    if (fill.getType() === SlidesApp.FillType.SOLID) {
      const current = safeGetFillColor(fill);
      if (!colorsEqual(current, color)) {
        background.setSolidFill(color);
      }
    } else if (fill.getType() === SlidesApp.FillType.NONE) {
      background.setSolidFill(color);
    }
  } catch (err) {
    Logger.log('[ThemeToggle] background error: ' + err);
  }
}

function applyPaletteToElement(element, palette) {
  const type = element.getPageElementType();
  switch (type) {
    case SlidesApp.PageElementType.SHAPE:
      refreshShape(element.asShape(), palette);
      break;
    case SlidesApp.PageElementType.TABLE:
      safeRecolorTable(element.asTable(), palette);
      break;
    case SlidesApp.PageElementType.LINE:
      refreshLine(element.asLine(), palette.border);
      break;
    case SlidesApp.PageElementType.GROUP:
      const children = element.asGroup().getChildren();
      for (let i = 0; i < children.length; i++) {
        applyPaletteToElement(children[i], palette);
      }
      break;
    default:
      break;
  }
}

function refreshShape(shape, palette) {
  safeSetTextColor(shape, palette);

  const fill = shape.getFill();
  if (fill) {
    const fillType = fill.getType();
    if (fillType === SlidesApp.FillType.SOLID) {
      if (!isHeroShape(shape)) {
        const current = safeGetFillColor(fill);
        if (!colorsEqual(current, palette.faintBg)) {
          fill.setSolidFill(palette.faintBg);
        }
      }
    }
  }

  try {
    const border = shape.getBorder();
    if (border) {
      const lineFill = border.getLineFill();
      if (lineFill && lineFill.getType() === SlidesApp.FillType.SOLID) {
        const current = safeGetFillColor(lineFill);
        if (!colorsEqual(current, palette.border)) {
          lineFill.setSolidFill(palette.border);
        }
      } else if (lineFill) {
        lineFill.setSolidFill(palette.border);
      }
    }
  } catch (err) {
    Logger.log('[ThemeToggle] border refresh error: ' + err);
  }
}

function refreshLine(line, targetColor) {
  if (!line) return;
  const lineFill = line.getLineFill();
  if (!lineFill) return;

  const current = safeGetFillColor(lineFill);
  if (!colorsEqual(current, targetColor)) {
    lineFill.setSolidFill(targetColor);
  }
}

function recolorTableCell(cell, palette, isHeader) {
  try {
    const fill = cell.getFill();
    if (fill) {
      const fillType = fill.getType();
      if (fillType === SlidesApp.FillType.SOLID) {
        const target = isHeader ? palette.faintBg : palette.background;
        const current = safeGetFillColor(fill);
        if (!colorsEqual(current, target)) {
          fill.setSolidFill(target);
        }
      } else if (fillType === SlidesApp.FillType.NONE && isHeader) {
        fill.setSolidFill(palette.faintBg);
      }
    }
  } catch (err) {
    Logger.log('[ThemeToggle] table cell fill error: ' + err);
  }

  try {
    const borders = [
      cell.getBorderTop(),
      cell.getBorderBottom(),
      cell.getBorderLeft(),
      cell.getBorderRight(),
    ];
    for (let i = 0; i < borders.length; i++) {
      const border = borders[i];
      if (!border) continue;
      const color = safeGetBorderColor(border);
      if (!colorsEqual(color, palette.border)) {
        border.setColor(palette.border);
      }
    }
  } catch (err) {
    Logger.log('[ThemeToggle] table border error: ' + err);
  }

  try {
    const text = cell.getText();
    if (text && text.getLength() > 0) {
      const style = text.getTextStyle();
      if (style) {
        const target = isHeader ? palette.textPrimary : palette.textSecondary;
        const current = safeGetForegroundColor(style);
        if (!colorsEqual(current, target)) {
          style.setForegroundColor(target);
        }
      }
    }
  } catch (err) {
    Logger.log('[ThemeToggle] table text error: ' + err);
  }
}

function safeGetFillColor(fill) {
  try {
    if (!fill) return null;
    if (fill.getType() !== SlidesApp.FillType.SOLID) return null;
    const color = fill.getColor();
    if (!color) return null;
    const rgb = color.asRgbColor();
    if (!rgb) return null;
    return normalizeColor(rgb.asHexString());
  } catch (err) {
    return null;
  }
}

function safeGetBorderColor(border) {
  try {
    const color = border.getColor();
    if (!color) return null;
    return normalizeColor(color);
  } catch (err) {
    return null;
  }
}

function safeGetForegroundColor(style) {
  try {
    if (!style.isForegroundColorSet()) return null;
    const color = style.getForegroundColor();
    if (!color) return null;
    const rgb = color.asRgbColor();
    if (!rgb) return null;
    return normalizeColor(rgb.asHexString());
  } catch (err) {
    return null;
  }
}

function normalizeColor(color) {
  if (!color) return null;
  let value = color.toString().trim();
  if (value.charAt(0) !== '#') value = '#' + value;
  return value.toUpperCase();
}

function colorsEqual(colorA, colorB) {
  if (!colorA && !colorB) return true;
  if (!colorA || !colorB) return false;
  return normalizeColor(colorA) === normalizeColor(colorB);
}

function inferFontSize(textRange) {
  try {
    const uniformSize = textRange.getTextStyle().getFontSize();
    if (uniformSize) return uniformSize;
  } catch (err) {
    // 続行
  }

  const length = textRange.getLength();
  for (let i = 0; i < length; i++) {
    try {
      const charRange = textRange.getRange(i, i + 1);
      const style = charRange.getTextStyle();
      if (style) {
        const size = style.getFontSize();
        if (size) return size;
      }
    } catch (err) {
      // 続行
    }
  }
  return 14;
}

function isHeroShape(shape) {
  try {
    const page = shape.getPage();
    const slideWidth = page.getPageWidth();
    const slideHeight = page.getPageHeight();
    const bounds = shape.getGeometry();
    const area = bounds ? bounds.getWidth() * bounds.getHeight() : 0;
    const slideArea = slideWidth * slideHeight;
    if (!slideArea) return false;
    const coverage = area / slideArea;
    return coverage > 0.6;
  } catch (err) {
    return false;
  }
}
