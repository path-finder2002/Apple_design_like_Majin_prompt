function renderSingleImageInArea(slide, layout, area, imageUrl, caption = '', position = 'left') {
  if (!imageUrl) return;
  
  try {
    const imageData = insertImageFromUrlOrFileId(imageUrl);
    if (!imageData) return null;
    
    const img = slide.insertImage(imageData);
    
    // 固定フレーム内にフィット（アスペクト比維持）
    const scale = Math.min(area.width / img.getWidth(), area.height / img.getHeight());
    const w = img.getWidth() * scale;
    const h = img.getHeight() * scale;
    
    // エリア中央に配置
    img.setWidth(w).setHeight(h)
       .setLeft(area.left + (area.width - w) / 2)
       .setTop(area.top + (area.height - h) / 2);
    
    // キャプション追加（画像の実際のサイズに応じて動的に配置）
    if (caption && caption.trim()) {
      // 画像の実際の位置とサイズを取得
      const imageBottom = area.top + (area.height - h) / 2 + h;
      const captionMargin = layout.pxToPt(8); // キャプションと画像の間隔
      const captionHeight = layout.pxToPt(30);
      
      // キャプションを画像の下に配置
      const captionTop = imageBottom + captionMargin;
      const captionShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
        area.left, captionTop, area.width, captionHeight);
      captionShape.getFill().setTransparent();
      captionShape.getBorder().setTransparent();
      setStyledText(captionShape, caption.trim(), {
        size: CONFIG.FONTS.sizes.small,
        color: CONFIG.COLORS.neutral_gray,
        align: SlidesApp.ParagraphAlignment.CENTER
      });
    }
       
    return img;
  } catch (e) {
    Logger.log(`Image insertion failed: ${e.message}. URL: ${imageUrl}`);
    return null;
  }
}

// ========================================
// 7. ユーティリティ関数群
// ========================================
function setMainSlideBackground(slide, layout) {
  setBackgroundImageFromUrl(slide, layout, CONFIG.BACKGROUND_IMAGES.main, CONFIG.COLORS.background_white);
}

function setBackgroundImageFromUrl(slide, layout, imageUrl, fallbackColor) {
  slide.getBackground().setSolidFill(fallbackColor);
  if (!imageUrl) return;
  try {
    const image = insertImageFromUrlOrFileId(imageUrl);
    if (!image) return;
    
    slide.insertImage(image).setWidth(layout.pageW_pt).setHeight(layout.pageH_pt).setLeft(0).setTop(0).sendToBack();
  } catch (e) {
    Logger.log(`Background image failed: ${e.message}. URL: ${imageUrl}`);
  }
}

/**
 * URLまたはGoogle Drive FileIDから画像を取得
 * @param {string} urlOrFileId - 画像URLまたはGoogle Drive FileID
 * @return {Blob|string} 画像データまたはURL
 */
function insertImageFromUrlOrFileId(urlOrFileId) {
  if (!urlOrFileId) return null;
  
  // URLからFileIDを抽出する関数
  function extractFileIdFromUrl(url) {
    const patterns = [
      /\/file\/d\/([a-zA-Z0-9_-]+)/,
      /id=([a-zA-Z0-9_-]+).*file/,
      /file\/([a-zA-Z0-9_-]+)/
    ];
    for (const pattern of patterns) {
      const match = url.match(pattern);
      if (match && match[1]) return match[1];
    }
    return null;
  }
  
  // FileIDの形式かチェック（GoogleドライブのFileIDは通常28-33文字の英数字）
  const fileIdPattern = /^[a-zA-Z0-9_-]{25,}$/;
  
  // URLからFileIDを抽出
  const extractedFileId = extractFileIdFromUrl(urlOrFileId);
  
  if (extractedFileId && fileIdPattern.test(extractedFileId)) {
    // Google Drive FileIDとして処理
    try {
      const file = DriveApp.getFileById(extractedFileId);
      return file.getBlob();
    } catch (e) {
      Logger.log(`Drive file access failed: ${e.message}. FileID: ${extractedFileId}`);
      return null;
    }
  } else if (fileIdPattern.test(urlOrFileId)) {
    // 直接FileIDとして処理
    try {
      const file = DriveApp.getFileById(urlOrFileId);
      return file.getBlob();
    } catch (e) {
      Logger.log(`Drive file access failed: ${e.message}. FileID: ${urlOrFileId}`);
      return null;
    }
  } else {
    // URLとして処理
    return urlOrFileId;
  }
}

function normalizeImages(arr) {
  return (arr || []).map(v => typeof v === 'string' ? {
    url: v
  } : (v && v.url ? v : null)).filter(Boolean).slice(0, 6);
}

function renderImagesInArea(slide, layout, area, images) {
  if (!images || !images.length) return;
  const n = Math.min(6, images.length);
  let cols = n === 1 ? 1 : (n <= 4 ? 2 : 3);
  const rows = Math.ceil(n / cols);
  const gap = layout.pxToPt(10);
  const cellW = (area.width - gap * (cols - 1)) / cols,
    cellH = (area.height - gap * (rows - 1)) / rows;
  for (let i = 0; i < n; i++) {
    const r = Math.floor(i / cols),
      c = i % cols;
    try {
      const img = slide.insertImage(images[i].url);
      const scale = Math.min(cellW / img.getWidth(), cellH / img.getHeight());
      const w = img.getWidth() * scale,
        h = img.getHeight() * scale;
      img.setWidth(w).setHeight(h).setLeft(area.left + c * (cellW + gap) + (cellW - w) / 2).setTop(area.top + r * (cellH + gap) + (cellH - h) / 2);
    } catch (e) {}
  }
}

function createGradientRectangle(slide, x, y, width, height, colors) {
  const numStrips = Math.max(20, Math.floor(width / 2));
  const stripWidth = width / numStrips;
  const startColor = hexToRgb(colors[0]),
    endColor = hexToRgb(colors[1]);
  if (!startColor || !endColor) return null;
  const shapes = [];
  for (let i = 0; i < numStrips; i++) {
    const ratio = i / (numStrips - 1);
    const r = Math.round(startColor.r + (endColor.r - startColor.r) * ratio);
    const g = Math.round(startColor.g + (endColor.g - startColor.g) * ratio);
    const b = Math.round(startColor.b + (endColor.b - startColor.b) * ratio);
    const strip = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x + (i * stripWidth), y, stripWidth + 0.5, height);
    strip.getFill().setSolidFill(r, g, b);
    strip.getBorder().setTransparent();
    shapes.push(strip);
  }
  if (shapes.length > 1) {
    return slide.group(shapes);
  }
  return shapes[0] || null;
}

function applyFill(slide, x, y, width, height, settings) {
  if (settings.enableGradient) {
    createGradientRectangle(slide, x, y, width, height, [settings.gradientStart, settings.gradientEnd]);
  } else {
    const shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, width, height);
    shape.getFill().setSolidFill(settings.primaryColor);
    shape.getBorder().setTransparent();
