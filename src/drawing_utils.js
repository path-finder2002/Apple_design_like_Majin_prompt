function estimateTextWidthPt(text, fontSizePt) {
  const multipliers = {
    ascii: 0.62,
    japanese: 1.0,
    other: 0.85
  };
  return String(text || '').split('').reduce((acc, char) => {
    if (char.match(/[ -~]/)) {
      return acc + multipliers.ascii;
    } else if (char.match(/[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF]/)) {
      return acc + multipliers.japanese;
    } else {
      return acc + multipliers.other;
    }
  }, 0) * fontSizePt;
}
function drawStandardTitleHeader(slide, layout, key, title, settings) {
  const logoRect = safeGetRect(layout, `${key}.headerLogo`);
  
  try {
    if (CONFIG.LOGOS.header && logoRect) {
      const imageData = insertImageFromUrlOrFileId(CONFIG.LOGOS.header);
      if (imageData) {
        const logo = slide.insertImage(imageData);
        const asp = logo.getHeight() / logo.getWidth();
        logo.setLeft(logoRect.left).setTop(logoRect.top).setWidth(logoRect.width).setHeight(logoRect.width * asp);
      }
    }
  } catch (e) {
    Logger.log(`Header logo error: ${e.message}`);
  }
  const titleRect = safeGetRect(layout, `${key}.title`);
  if (!titleRect) {
    Logger.log(`[rect-missing] key=${key}.title`);
    return;
  }
  
  const fontSize = CONFIG.FONTS.sizes.contentTitle;
  const optimalHeight = layout.pxToPt(fontSize + 8); // フォントサイズ + 8px余白
  
  const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
    titleRect.left, 
    titleRect.top, 
    titleRect.width, 
    optimalHeight
  );
  setStyledText(titleShape, title || '', {
    size: fontSize,
    bold: true
  });
  
  // 上揃えで元の位置を維持
  try {
    titleShape.setContentAlignment(SlidesApp.ContentAlignment.TOP);
  } catch (e) {}
  
  if (settings.showTitleUnderline && title) {
    const uRect = safeGetRect(layout, `${key}.titleUnderline`);
    if (!uRect) {
      Logger.log(`[rect-missing] key=${key}.titleUnderline`);
      return;
    }
    const estimatedWidthPt = estimateTextWidthPt(title, fontSize);
    
    // アンダーライン幅の上限制限（スライド枠内に収める）
    const maxUnderlineWidth = layout.pageW_pt - uRect.left - layout.pxToPt(25); // 右余白25px確保
    const finalWidth = Math.min(estimatedWidthPt, maxUnderlineWidth);
    
    applyFill(slide, uRect.left, uRect.top, finalWidth, uRect.height, settings);
  }
}

/**
 * テキスト高さ概算（pt）。日本語を含む一般文で「1em ≒ fontSizePt」前提の簡易推定。
 * 既存コードで px を使っていた場合のズレもここで吸収。
 */
function estimateTextHeightPt(text, widthPt, fontSizePt, lineHeight) {
  var paragraphs = String(text).split(/\r?\n/);
  // 1行に入る概算文字数：幅 / (0.95em) として保守的に
  var charsPerLine = Math.max(1, Math.floor(widthPt / (fontSizePt * 0.95)));
  var lines = 0;
  for (var i = 0; i < paragraphs.length; i++) {
    var s = paragraphs[i].replace(/\s+/g, ' ').trim();
    // 空行も最低1行とみなす
    var len = s.length || 1;
    lines += Math.ceil(len / charsPerLine);
  }
  var lineH = fontSizePt * (lineHeight || 1.2);
  return Math.max(lineH, lines * lineH);
}

/**
 * Subhead を描画して、次要素の基準Yを返すユーティリティ。
 * 互換用に dy も返す（reserved 分との差分）。全単位 pt。
 *
 * @param {SlidesApp.Slide} slide
 * @param {{x:number,y:number,width:number}} frame   // コンテンツ基準枠（pt）
 * @param {string} subheadText
 * @param {{
 *   fontFamily?: string,
 *   fontSizePt?: number,     // 例: 18
 *   lineHeight?: number,     // 例: 1.25 （ParagraphStyle.setLineSpacing に %換算で反映）
 *   color?: string,          // 例: '#444444'
 *   bold?: boolean,
 *   marginAfterPt?: number,  // subhead 下の外余白（例: 12）
 *   reservedPt?: number,     // 既存レイアウトが想定していた予約高さ（例: 22）
 * }} [style]
 * @param {string} [debugKey] // ログ表示用キー
 * @return {{dy:number,nextY:number,height:number,shape:SlidesApp.Shape|null}}
 */
function drawSubheadIfAny(slide, layout, key, subhead) {
  if (!subhead) return 0;
  
  const rect = safeGetRect(layout, `${key}.subhead`);
  if (!rect) {
    Logger.log(`[drawSubheadIfAny] Could not get rect for ${key}.subhead`);
    return 0;
  }
  
  const box = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rect.left, rect.top, rect.width, rect.height);
  setStyledText(box, subhead, {
    size: CONFIG.FONTS.sizes.subhead,
    color: CONFIG.COLORS.text_primary
  });
  
  return layout.pxToPt(36); // 固定値を返す
}


function drawBottomBar(slide, layout, settings) {
  const barRect = layout.getRect('bottomBar');
  applyFill(slide, barRect.left, barRect.top, barRect.width, barRect.height, settings);
}

function addCucFooter(slide, layout, pageNum) {
  const leftRect = layout.getRect('footer.leftText');
  const leftShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, leftRect.left, leftRect.top, leftRect.width, leftRect.height);
  leftShape.getText().setText(CONFIG.FOOTER_TEXT);
  applyTextStyle(leftShape.getText(), {
    size: CONFIG.FONTS.sizes.footer,
    color: CONFIG.COLORS.text_primary
  });
  if (pageNum > 0) {
    const rightRect = layout.getRect('footer.rightPage');
    const rightShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rightRect.left, rightRect.top, rightRect.width, rightRect.height);
    rightShape.getText().setText(String(pageNum));
    applyTextStyle(rightShape.getText(), {
      size: CONFIG.FONTS.sizes.footer,
      color: CONFIG.COLORS.primary_color,
      align: SlidesApp.ParagraphAlignment.END
    });
  }
}

function drawBottomBarAndFooter(slide, layout, pageNum, settings) {
  if (settings.showBottomBar) {
    drawBottomBar(slide, layout, settings);
  }
  addCucFooter(slide, layout, pageNum);
}

function drawCompareBox(slide, layout, rect, title, items, settings, isLeft = false) {
  const box = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, rect.left, rect.top, rect.width, rect.height);
  box.getFill().setSolidFill(CONFIG.COLORS.background_gray);
  box.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.lane_border);
  const th = layout.pxToPt(40);
  const titleBarBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, rect.left, rect.top, rect.width, th);
  
  // 左右対比色の適用
  const compareColors = generateCompareColors(settings.primaryColor);
  const headerColor = isLeft ? compareColors.left : compareColors.right;
  titleBarBg.getFill().setSolidFill(headerColor);
  titleBarBg.getBorder().setTransparent();
  const titleTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rect.left, rect.top, rect.width, th);
  titleTextShape.getFill().setTransparent();
  titleTextShape.getBorder().setTransparent();
  setStyledText(titleTextShape, title, {
    size: CONFIG.FONTS.sizes.laneTitle,
    bold: true,
    color: CONFIG.COLORS.background_white,
    align: SlidesApp.ParagraphAlignment.CENTER
  });
  const pad = layout.pxToPt(12);
  const body = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rect.left + pad, rect.top + th + pad, rect.width - pad * 2, rect.height - th - pad * 2);
  setBulletsWithInlineStyles(body, items);
}

/**
 * [修正4] レーン図の矢印をカード間を結ぶ線（コネクタ）に変更
 */
function drawArrowBetweenRects(slide, a, b, arrowH, arrowGap, settings) {
  const fromX = a.left + a.width;
  const fromY = a.top + a.height / 2;
  const toX = b.left;
  const toY = b.top + b.height / 2;

  // 描画するスペースがある場合のみ線を描画
  if (toX - fromX <= 0) return;

  const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, fromX, fromY, toX, toY);
  line.getLineFill().setSolidFill(settings.primaryColor);
  line.setWeight(1.5);
  line.setEndArrow(SlidesApp.ArrowStyle.FILL_ARROW);
}


function drawNumberedItems(slide, layout, area, items, settings) {
  // アジェンダ用の座布団を作成
  createContentCushion(slide, area, settings, layout);
  
  const n = Math.max(1, items.length);
  const topPadding = layout.pxToPt(30);
  const bottomPadding = layout.pxToPt(10);
  const drawableHeight = area.height - topPadding - bottomPadding;
  const gapY = drawableHeight / Math.max(1, n - 1);
  const cx = area.left + layout.pxToPt(44);
  const top0 = area.top + topPadding;

  const line = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, cx - layout.pxToPt(1), top0 + layout.pxToPt(6), layout.pxToPt(2), gapY * (n - 1));
  line.getFill().setSolidFill(CONFIG.COLORS.faint_gray);
  line.getBorder().setTransparent();

  for (let i = 0; i < n; i++) {
    const cy = top0 + gapY * i;
    const sz = layout.pxToPt(28);
    const numBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, cx - sz/2, cy - sz/2, sz, sz);
    numBox.getFill().setSolidFill(settings.primaryColor);
    numBox.getBorder().setTransparent();
    const num = numBox.getText(); num.setText(String(i + 1));
    applyTextStyle(num, { size: 12, bold: true, color: CONFIG.COLORS.background_white, align: SlidesApp.ParagraphAlignment.CENTER });

    // 元の箇条書きテキストから先頭の数字を除去
    let cleanText = String(items[i] || '');
    cleanText = cleanText.replace(/^\s*\d+[\.\s]*/, '');

    const txt = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, cx + layout.pxToPt(28), cy - layout.pxToPt(16), area.width - layout.pxToPt(70), layout.pxToPt(32));
    setStyledText(txt, cleanText, { size: CONFIG.FONTS.sizes.processStep });
    try { txt.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch(e){}
  }
}

// ========================================
// 9. ヘルパー関数群
// ========================================

