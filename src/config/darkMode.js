// Dark mode theme refresh utilities and operations

// Theme color mappings between light and dark
var THEME_COLOR_MAPPINGS = (function(){
  var t = (CONFIG && CONFIG.THEMES) ? CONFIG.THEMES : {};
  return {
    lightToDark: {
      primary: t.light ? t.light.primary : '#6750A4',
      secondary: t.light ? t.light.secondary : '#625B71',
      background: t.light ? t.light.background : '#FDF8FF',
      surface: t.light ? t.light.surface : '#FEF7FF',
      text: t.light ? t.light.text : '#1C1B1F',
      accent: t.light ? t.light.accent : '#03DAC5',
      mapTo: {
        primary: t.dark ? t.dark.primary : '#D0BCFF',
        secondary: t.dark ? t.dark.secondary : '#CCC2DC',
        background: t.dark ? t.dark.background : '#101014',
        surface: t.dark ? t.dark.surface : '#141218',
        text: t.dark ? t.dark.text : '#E6E0E9',
        accent: t.dark ? t.dark.accent : '#66FFF9'
      }
    },
    darkToLight: {
      primary: t.dark ? t.dark.primary : '#D0BCFF',
      secondary: t.dark ? t.dark.secondary : '#CCC2DC',
      background: t.dark ? t.dark.background : '#101014',
      surface: t.dark ? t.dark.surface : '#141218',
      text: t.dark ? t.dark.text : '#E6E0E9',
      accent: t.dark ? t.dark.accent : '#66FFF9',
      mapTo: {
        primary: t.light ? t.light.primary : '#6750A4',
        secondary: t.light ? t.light.secondary : '#625B71',
        background: t.light ? t.light.background : '#FDF8FF',
        surface: t.light ? t.light.surface : '#FEF7FF',
        text: t.light ? t.light.text : '#1C1B1F',
        accent: t.light ? t.light.accent : '#03DAC5'
      }
    }
  };
})();

function buildThemeColorIndex(themeKey) {
  var theme = CONFIG && CONFIG.THEMES && CONFIG.THEMES[themeKey] ? CONFIG.THEMES[themeKey] : CONFIG.THEMES.light;
  var idx = {};
  idx[(theme.primary || '').toUpperCase()] = 'primary';
  idx[(theme.secondary || '').toUpperCase()] = 'secondary';
  idx[(theme.background || '').toUpperCase()] = 'background';
  idx[(theme.surface || '').toUpperCase()] = 'surface';
  idx[(theme.text || '').toUpperCase()] = 'text';
  if (theme.accent) idx[(theme.accent || '').toUpperCase()] = 'accent';
  return idx;
}

function hexNorm(c){ return (c || '').toString().trim().toUpperCase(); }

function isWithinColorTolerance(color1, color2, tolerance) {
  var c1 = hexToRgb(color1), c2 = hexToRgb(color2);
  if (!c1 || !c2) return false;
  var dr = Math.abs(c1.r - c2.r);
  var dg = Math.abs(c1.g - c2.g);
  var db = Math.abs(c1.b - c2.b);
  return (dr + dg + db) <= (typeof tolerance === 'number' ? tolerance : 30);
}

function inferBrightnessFactor(derivedColor, baseColor) {
  var c1 = hexToRgb(derivedColor), c2 = hexToRgb(baseColor);
  if (!c1 || !c2) return 1.0;
  var hsl1 = rgbToHsl(c1.r, c1.g, c1.b);
  var hsl2 = rgbToHsl(c2.r, c2.g, c2.b);
  var delta = hsl1.l - hsl2.l; // percentage
  var factor = 1 + (delta / 100);
  if (factor < 0.2) factor = 0.2;
  if (factor > 2.0) factor = 2.0;
  return factor;
}

function transformDerivedColor(originalColor, oldPrimary, newPrimary) {
  if (!originalColor) return originalColor;
  var oc = hexNorm(originalColor);
  var op = hexNorm(oldPrimary);
  var np = hexNorm(newPrimary);
  if (!op || !np) return oc;
  if (oc === op) return np;
  var tol = 48;
  if (!isWithinColorTolerance(oc, op, tol)) {
    // derive by brightness relationship
    var bf = inferBrightnessFactor(oc, op);
    var rgbNP = hexToRgb(np);
    var hslNP = rgbToHsl(rgbNP.r, rgbNP.g, rgbNP.b);
    var l2 = Math.max(0, Math.min(100, hslNP.l * bf));
    return hslToHex(hslNP.h, hslNP.s, l2);
  }
  return np;
}

function applyThemeToTextRange(textRange, colorIndex, newTheme) {
  try {
    var style = textRange.getTextStyle();
    style.setForegroundColor(newTheme.text || '#000000');
  } catch (e) {}
}

function refreshShapeElement(shape, colorIndex, newTheme) {
  try {
    var fill = shape.getFill();
    // Update solid fills only
    try {
      var colorObj = fill.getSolidFill && fill.getSolidFill();
      var oldColor = colorObj && colorObj.getColor && colorObj.getColor();
      var oldHex = oldColor && oldColor.asRgbColor && oldColor.asRgbColor().asHexString ? oldColor.asRgbColor().asHexString() : null;
      if (oldHex) {
        var mapped = transformDerivedColor(oldHex, CONFIG.THEMES.light.primary, newTheme.primary);
        fill.setSolidFill(mapped);
      } else {
        // fallback to theme surface
        fill.setSolidFill(newTheme.surface);
      }
    } catch (e) {}
    // Text color
    try {
      var tr = shape.getText();
      if (tr) applyThemeToTextRange(tr, colorIndex, newTheme);
    } catch (e) {}
  } catch (e) {}
}

function refreshTableElement(table, colorIndex, newTheme) {
  try {
    var rows = table.getNumRows();
    var cols = table.getNumColumns();
    for (var r = 0; r < rows; r++) {
      for (var c = 0; c < cols; c++) {
        var cell = table.getCell(r, c);
        try { cell.getFill().setSolidFill(newTheme.surface); } catch (e) {}
        try { applyThemeToTextRange(cell.getText(), colorIndex, newTheme); } catch (e) {}
      }
    }
  } catch (e) {}
}

function updateSlideElements(slide, oldTheme, newTheme, colorIndex) {
  try { slide.getBackground().setSolidFill(newTheme.background); } catch (e) {}
  var elements = slide.getPageElements();
  for (var i = 0; i < elements.length; i++) {
    var el = elements[i];
    var type = el.getPageElementType && el.getPageElementType();
    if (type === SlidesApp.PageElementType.SHAPE) {
      refreshShapeElement(el.asShape(), colorIndex, newTheme);
    } else if (type === SlidesApp.PageElementType.TABLE) {
      refreshTableElement(el.asTable(), colorIndex, newTheme);
    } else if (type === SlidesApp.PageElementType.GROUP) {
      var group = el.asGroup().getChildren();
      for (var j = 0; j < group.length; j++) {
        var g = group[j];
        var gt = g.getPageElementType && g.getPageElementType();
        if (gt === SlidesApp.PageElementType.SHAPE) refreshShapeElement(g.asShape(), colorIndex, newTheme);
        if (gt === SlidesApp.PageElementType.TABLE) refreshTableElement(g.asTable(), colorIndex, newTheme);
      }
    }
  }
}

function refreshPresentationTheme(newThemeKey) {
  var oldKey = (typeof getCurrentTheme === 'function') ? getCurrentTheme() : 'light';
  var newKey = (newThemeKey === 'dark' || newThemeKey === 'light') ? newThemeKey : (oldKey === 'dark' ? 'light' : 'dark');
  if (typeof setCurrentTheme === 'function') setCurrentTheme(newKey);
  var pres = SlidesApp.getActivePresentation && SlidesApp.getActivePresentation();
  if (!pres) {
    return { status: 'error', message: 'No active presentation found.' };
  }
  var slides = pres.getSlides();
  var colorIndex = buildThemeColorIndex(oldKey);
  var oldTheme = CONFIG.THEMES[oldKey];
  var newTheme = CONFIG.THEMES[newKey];
  for (var i = 0; i < slides.length; i++) {
    updateSlideElements(slides[i], oldTheme, newTheme, colorIndex);
  }
  return { status: 'success', message: 'Theme refreshed', theme: newKey, updatedSlides: slides.length };
}

