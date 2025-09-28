/**
 * ========================================
 * 色彩操作ヘルパー関数
 * ========================================
 */
function hexToRgb(hex) {
  const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result ? {
    r: parseInt(result[1], 16),
    g: parseInt(result[2], 16),
    b: parseInt(result[3], 16)
  } : null;
}

function rgbToHsl(r, g, b) {
  r /= 255;
  g /= 255;
  b /= 255;
  const max = Math.max(r, g, b),
    min = Math.min(r, g, b);
  let h, s, l = (max + min) / 2;
  if (max === min) {
    h = s = 0;
  } else {
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
    }
    h /= 6;
  }
  return {
    h: h * 360,
    s: s * 100,
    l: l * 100
  };
}

function hslToHex(h, s, l) {
  s /= 100;
  l /= 100;
  const c = (1 - Math.abs(2 * l - 1)) * s,
    x = c * (1 - Math.abs((h / 60) % 2 - 1)),
    m = l - c / 2;
  let r = 0,
    g = 0,
    b = 0;
  if (0 <= h && h < 60) {
    r = c;
    g = x;
    b = 0;
  } else if (60 <= h && h < 120) {
    r = x;
    g = c;
    b = 0;
  } else if (120 <= h && h < 180) {
    r = 0;
    g = c;
    b = x;
  } else if (180 <= h && h < 240) {
    r = 0;
    g = x;
    b = c;
  } else if (240 <= h && h < 300) {
    r = x;
    g = 0;
    b = c;
  } else if (300 <= h && h < 360) {
    r = c;
    g = 0;
    b = x;
  }
  r = Math.round((r + m) * 255);
  g = Math.round((g + m) * 255);
  b = Math.round((b + m) * 255);
  return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
}

function generateTintedGray(tintColorHex, saturation, lightness) {
  const rgb = hexToRgb(tintColorHex);
  if (!rgb) return '#F8F9FA';
  const hsl = rgbToHsl(rgb.r, rgb.g, rgb.b);
  return hslToHex(hsl.h, saturation, lightness);
}

/**
 * ピラミッド用カラーグラデーション生成
 * @param {string} baseColor - ベースとなるプライマリカラー
 * @param {number} levels - レベル数
 * @return {string[]} 上から下へのグラデーションカラー配列
 */
function generatePyramidColors(baseColor, levels) {
  const colors = [];
  for (let i = 0; i < levels; i++) {
    // 上から下に向かって徐々に薄くなる (0% → 60%まで)
    const lightenAmount = (i / Math.max(1, levels - 1)) * 0.6;
    colors.push(lightenColor(baseColor, lightenAmount));
  }
  return colors;
}

/**
 * 色を明るくする関数
 * @param {string} color - 元の色 (#RRGGBB形式)
 * @param {number} amount - 明るくする量 (0.0-1.0)
 * @return {string} 明るくした色
 */
function lightenColor(color, amount) {
  const rgb = hexToRgb(color);
  if (!rgb) return color;
  
  const lighten = (c) => Math.min(255, Math.round(c + (255 - c) * amount));
  const newR = lighten(rgb.r);
  const newG = lighten(rgb.g);
  const newB = lighten(rgb.b);
  
  return `#${newR.toString(16).padStart(2, '0')}${newG.toString(16).padStart(2, '0')}${newB.toString(16).padStart(2, '0')}`;
}

/**
 * 色を暗くする関数
 * @param {string} color - 元の色 (#RRGGBB形式)
 * @param {number} amount - 暗くする量 (0.0-1.0)
 * @return {string} 暗くした色
 */
function darkenColor(color, amount) {
  const rgb = hexToRgb(color);
  if (!rgb) return color;
  
  const darken = (c) => Math.max(0, Math.round(c * (1 - amount)));
  const newR = darken(rgb.r);
  const newG = darken(rgb.g);
  const newB = darken(rgb.b);
  
  return `#${newR.toString(16).padStart(2, '0')}${newG.toString(16).padStart(2, '0')}${newB.toString(16).padStart(2, '0')}`;
}

/**
 * StepUp用カラーグラデーション生成（左から右に濃くなる）
 * @param {string} baseColor - ベースとなるプライマリカラー
 * @param {number} steps - ステップ数
 * @return {string[]} 左から右へのグラデーションカラー配列（薄い→濃い）
 */
function generateStepUpColors(baseColor, steps) {
  const colors = [];
  for (let i = 0; i < steps; i++) {
    // 左から右に向かって徐々に濃くなる (60% → 0%)
    const lightenAmount = 0.6 * (1 - (i / Math.max(1, steps - 1)));
    colors.push(lightenColor(baseColor, lightenAmount));
  }
  return colors;
}

/**
 * Process用カラーグラデーション生成（上から下に濃くなる）
 * @param {string} baseColor - ベースとなるプライマリカラー
 * @param {number} steps - ステップ数
 * @return {string[]} 上から下へのグラデーションカラー配列（薄い→濃い）
 */
function generateProcessColors(baseColor, steps) {
  const colors = [];
  for (let i = 0; i < steps; i++) {
    // 上から下に向かって徐々に濃くなる (50% → 0%)
    const lightenAmount = 0.5 * (1 - (i / Math.max(1, steps - 1)));
    colors.push(lightenColor(baseColor, lightenAmount));
  }
  return colors;
}

/**
 * Timeline用カードグラデーション生成（左から右に濃くなる）
 * @param {string} baseColor - ベースとなるプライマリカラー
 * @param {number} milestones - マイルストーン数
 * @return {string[]} 左から右へのグラデーションカラー配列（薄い→濃い）
 */
function generateTimelineCardColors(baseColor, milestones) {
  const colors = [];
  for (let i = 0; i < milestones; i++) {
    // 左から右に向かって徐々に濃くなる (40% → 0%)
    const lightenAmount = 0.4 * (1 - (i / Math.max(1, milestones - 1)));
    colors.push(lightenColor(baseColor, lightenAmount));
  }
  return colors;
}

/**
 * Compare系用の左右対比色生成
 * @param {string} baseColor - ベースとなるプライマリカラー
 * @return {Object} {left: 濃い色, right: 元の色}の組み合わせ
 */
function generateCompareColors(baseColor) {
  return {
    left: darkenColor(baseColor, 0.3),   // 左側：30%暗く（Before/導入前）- 視認性向上
    right: baseColor                     // 右側：元の色（After/導入後）
  };
}
function adjustColorBrightness(hex, factor) {
  const c = hex.replace('#', '');
  const rgb = parseInt(c, 16);
  let r = (rgb >> 16) & 0xff,
    g = (rgb >> 8) & 0xff,
    b = (rgb >> 0) & 0xff;
  r = Math.min(255, Math.round(r * factor));
  g = Math.min(255, Math.round(g * factor));
  b = Math.min(255, Math.round(b * factor));
  return '#' + (0x1000000 + (r << 16) + (g << 8) + b).toString(16).slice(1);
