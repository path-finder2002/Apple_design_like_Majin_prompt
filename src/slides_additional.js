// --- Slide Generators (Additional) ---
var slideGenerators = this.slideGenerators || (this.slideGenerators = {});

function createCardsSlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
  drawStandardTitleHeader(slide, layout, 'cardsSlide', data.title);
  const dy = drawSubheadIfAny(slide, layout, 'cardsSlide', data.subhead);

  const area = offsetRect(layout.getRect('cardsSlide.gridArea'), 0, dy);
  const items = Array.isArray(data.items) ? data.items : [];
  const cols = Math.min(3, Math.max(2, Number(data.columns) || (items.length <= 4 ? 2 : 3)));
  const gap = layout.pxToPt(16);
  const rows = Math.ceil(items.length / cols);
  const cardW = (area.width - gap * (cols - 1)) / cols;
  const cardH = Math.max(layout.pxToPt(92), (area.height - gap * (rows - 1)) / rows);

  for (let idx = 0; idx < items.length; idx++) {
    const r = Math.floor(idx / cols), c = idx % cols;
    const left = area.left + c * (cardW + gap);
    const top  = area.top  + r * (cardH + gap);

    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, left, top, cardW, cardH);
    card.getFill().setSolidFill(CONFIG.COLORS.card_bg);
    card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    card.getBorder().setWeight(1);

    const obj = items[idx];
    if (typeof obj === 'string') {
      setStyledText(card, obj, { size: CONFIG.FONTS.sizes.body });
    } else {
      const title = String(obj.title || '');
      const desc  = String(obj.desc || '');
      
      if (title.length > 0 && desc.length > 0) {
        // タイトル + 改行 + 説明文
        const combined = `${title}\n\n${desc}`;
        setStyledText(card, combined, { size: CONFIG.FONTS.sizes.body });
        try { 
          card.getText().getRange(0, title.length).getTextStyle().setBold(true);
        } catch(e){}
      } else if (title.length > 0) {
        // タイトルのみ
        setStyledText(card, title, { size: CONFIG.FONTS.sizes.body, bold: true });
      } else {
        // 説明文のみ（稀なケース）
        setStyledText(card, desc, { size: CONFIG.FONTS.sizes.body });
      }
    }
    try { card.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch(e) {}
  }

  drawBottomBarAndFooter(slide, layout, pageNum);
}

// headerCards（ヘッダー付きカード）
function createHeaderCardsSlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
  drawStandardTitleHeader(slide, layout, 'cardsSlide', data.title);
  const dy = drawSubheadIfAny(slide, layout, 'cardsSlide', data.subhead);

  const area = offsetRect(layout.getRect('cardsSlide.gridArea'), 0, dy);
  const items = Array.isArray(data.items) ? data.items : [];
  const cols = Math.min(3, Math.max(2, Number(data.columns) || (items.length <= 4 ? 2 : 3)));
  const gap = layout.pxToPt(16);
  const rows = Math.ceil(items.length / cols);
  const cardW = (area.width - gap * (cols - 1)) / cols;
  const cardH = Math.max(layout.pxToPt(92), (area.height - gap * (rows - 1)) / rows);

  for (let idx = 0; idx < items.length; idx++) {
    const r = Math.floor(idx / cols), c = idx % cols;
    const left = area.left + c * (cardW + gap);
    const top  = area.top  + r * (cardH + gap);

    const obj = items[idx];
    const titleText = (typeof obj === 'string') ? '' : String(obj.title || '');
    const descText = (typeof obj === 'string') ? String(obj) : String(obj.desc || '');
    
    const headerHeight = layout.pxToPt(40);
    const headerShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top, cardW, headerHeight);
    headerShape.getFill().setSolidFill(CONFIG.COLORS.primary_color);
    headerShape.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    headerShape.getBorder().setWeight(1);
    
    const bodyShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top + headerHeight, cardW, cardH - headerHeight);
    bodyShape.getFill().setSolidFill(CONFIG.COLORS.card_bg);
    bodyShape.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    bodyShape.getBorder().setWeight(1);
    
    const headerTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left, top, cardW, headerHeight);
    setStyledText(headerTextShape, titleText, { size: CONFIG.FONTS.sizes.body, bold: true, color: CONFIG.COLORS.text_on_primary, align: SlidesApp.ParagraphAlignment.CENTER });
    try { headerTextShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch(e){}

    const bodyTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left, top + headerHeight, cardW, cardH - headerHeight);
    setStyledText(bodyTextShape, descText, { size: CONFIG.FONTS.sizes.body, align: SlidesApp.ParagraphAlignment.CENTER });
    try { bodyTextShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch(e){}
  }

  drawBottomBarAndFooter(slide, layout, pageNum);
}

// table（表）
function createBulletCardsSlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
  drawStandardTitleHeader(slide, layout, 'contentSlide', data.title);
  const dy = drawSubheadIfAny(slide, layout, 'contentSlide', data.subhead);

  const area = offsetRect(layout.getRect('contentSlide.body'), 0, dy);
  const items = Array.isArray(data.items) ? data.items : [];
  const totalItems = Math.min(items.length, 3);
  if (totalItems === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum);
    return;
  }

  const gap = layout.pxToPt(16);
  const minCardHeight = layout.pxToPt(90);
  const maxCardHeight = layout.pxToPt(120);
  const idealCardHeight = (area.height - (totalItems - 1) * gap) / totalItems;
  const cardHeight = Math.max(minCardHeight, Math.min(maxCardHeight, idealCardHeight));
  
  let currentY = area.top;

  for (let i = 0; i < totalItems; i++) {
    const item = items[i];
    const card = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, currentY, area.width, cardHeight);
    card.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    card.getBorder().setWeight(1);

    const padding = layout.pxToPt(20);
    const title = String(item.title || '');
    const desc = String(item.desc || '');
    
    if (title.length > 0 && desc.length > 0) {
      const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left + padding, currentY + layout.pxToPt(12), area.width - padding * 2, layout.pxToPt(18));
      titleShape.getFill().setTransparent();
      titleShape.getBorder().setTransparent();
      setStyledText(titleShape, title, { size: 14, bold: true });
      
      const descShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left + padding, currentY + layout.pxToPt(38), area.width - padding * 2, cardHeight - layout.pxToPt(48));
      descShape.getFill().setTransparent();
      descShape.getBorder().setTransparent();
      setStyledText(descShape, desc, { size: 14, color: CONFIG.COLORS.text_primary });
    } else if (title.length > 0) {
      const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left + padding, currentY, area.width - padding * 2, cardHeight);
      titleShape.getFill().setTransparent();
      titleShape.getBorder().setTransparent();
      titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      setStyledText(titleShape, title, { size: 14, bold: true });
    } else {
      const descShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left + padding, currentY, area.width - padding * 2, cardHeight);
      descShape.getFill().setTransparent();
      descShape.getBorder().setTransparent();
      descShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      setStyledText(descShape, desc, { size: 14, color: CONFIG.COLORS.text_primary });
    }

    currentY += cardHeight + gap;
  }

  drawBottomBarAndFooter(slide, layout, pageNum);
}

// hybridContent（箇条書き＋カード統合）
function createHybridContentSlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
  drawStandardTitleHeader(slide, layout, 'hybridContentSlide', data.title);
  const dy = drawSubheadIfAny(slide, layout, 'hybridContentSlide', data.subhead);

  // 上部: 箇条書きエリア
  const bulletArea = offsetRect(layout.getRect('hybridContentSlide.bulletArea'), 0, dy);
  const points = Array.isArray(data.points) ? data.points : [];
  
  if (points.length > 0) {
    const isTwoColumn = points.length > 3;
    if (isTwoColumn) {
      const mid = Math.ceil(points.length / 2);
      const leftPoints = points.slice(0, mid);
      const rightPoints = points.slice(mid);
      
      const leftArea = { left: bulletArea.left, top: bulletArea.top, width: bulletArea.width * 0.48, height: bulletArea.height };
      const rightArea = { left: bulletArea.left + bulletArea.width * 0.52, top: bulletArea.top, width: bulletArea.width * 0.48, height: bulletArea.height };
      
      const leftShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, leftArea.left, leftArea.top, leftArea.width, leftArea.height);
      const rightShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rightArea.left, rightArea.top, rightArea.width, rightArea.height);
      setBulletsWithInlineStyles(leftShape, leftPoints);
      setBulletsWithInlineStyles(rightShape, rightPoints);
    } else {
      const bulletShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, bulletArea.left, bulletArea.top, bulletArea.width, bulletArea.height);
      setBulletsWithInlineStyles(bulletShape, points);
    }
  }

  // 下部: カードエリア
  const cardArea = offsetRect(layout.getRect('hybridContentSlide.cardArea'), 0, dy);
  const cards = Array.isArray(data.cards) ? data.cards : [];
  
  if (cards.length > 0) {
    const cardCount = Math.min(cards.length, 3);
    const cardWidth = (cardArea.width - layout.pxToPt(16) * (cardCount - 1)) / cardCount;
    
    for (let i = 0; i < cardCount; i++) {
      const card = cards[i];
      const x = cardArea.left + i * (cardWidth + layout.pxToPt(16));
      
      const cardShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, cardArea.top, cardWidth, cardArea.height);
      cardShape.getFill().setSolidFill(CONFIG.COLORS.background_gray);
      cardShape.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
      cardShape.getBorder().setWeight(1);
      
      const padding = layout.pxToPt(16);
      const title = String(card.title || '');
      const desc = String(card.desc || '');
      
      if (title.length > 0 && desc.length > 0) {
        const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + padding, cardArea.top + layout.pxToPt(12), cardWidth - padding * 2, layout.pxToPt(18));
        titleShape.getFill().setTransparent();
        titleShape.getBorder().setTransparent();
        setStyledText(titleShape, title, { size: 12, bold: true });
        
        const descShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + padding, cardArea.top + layout.pxToPt(36), cardWidth - padding * 2, cardArea.height - layout.pxToPt(48));
        descShape.getFill().setTransparent();
        descShape.getBorder().setTransparent();
        setStyledText(descShape, desc, { size: 11, color: CONFIG.COLORS.text_primary });
      } else if (title.length > 0) {
        const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + padding, cardArea.top, cardWidth - padding * 2, cardArea.height);
        titleShape.getFill().setTransparent();
        titleShape.getBorder().setTransparent();
        titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        setStyledText(titleShape, title, { size: 12, bold: true });
      } else if (desc.length > 0) {
        const descShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + padding, cardArea.top, cardWidth - padding * 2, cardArea.height);
        descShape.getFill().setTransparent();
        descShape.getBorder().setTransparent();
        descShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        setStyledText(descShape, desc, { size: 11, color: CONFIG.COLORS.text_primary });
      }
    }
  }

  drawBottomBarAndFooter(slide, layout, pageNum);
}

// faq（よくある質問）
function createFaqSlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
  drawStandardTitleHeader(slide, layout, 'contentSlide', data.title || 'よくあるご質問');
  const dy = 0; // FAQパターンでは小見出しを使用しない

  const area = offsetRect(layout.getRect('contentSlide.body'), 0, dy);
  const items = Array.isArray(data.items) ? data.items.slice(0, 4) : [];
  if (items.length === 0) { drawBottomBarAndFooter(slide, layout, pageNum); return; }

  let currentY = area.top;

  items.forEach(item => {
    const qShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, currentY, area.width, layout.pxToPt(40));
    const qText = qShape.getText();
    const fullQText = `Q. ${item.q || ''}`;
    qText.setText(fullQText);
    applyTextStyle(qText.getRange(0, 2), { size: 16, bold: true, color: CONFIG.COLORS.primary_color });
    if (fullQText.length > 2) {
      applyTextStyle(qText.getRange(2, fullQText.length), { size: 14, bold: true });
    }
    const qBounds = qShape.getHeight();
    currentY += qBounds + layout.pxToPt(4);

    const aShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left + layout.pxToPt(20), currentY, area.width - layout.pxToPt(20), layout.pxToPt(60));
    setStyledText(aShape, `A. ${item.a || ''}`, { size: 14 });
    const aBounds = aShape.getHeight();
    currentY += aBounds + layout.pxToPt(15);
  });

  drawBottomBarAndFooter(slide, layout, pageNum);
}

// compareCards（対比＋カード）
function createCompareCardsSlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
  drawStandardTitleHeader(slide, layout, 'compareCardsSlide', data.title);
  const dy = drawSubheadIfAny(slide, layout, 'compareCardsSlide', data.subhead);

  const leftArea = offsetRect(layout.getRect('compareCardsSlide.leftArea'), 0, dy);
  const rightArea = offsetRect(layout.getRect('compareCardsSlide.rightArea'), 0, dy);
  
  // 左側のタイトルヘッダー
  const leftTitleHeight = layout.pxToPt(40);
  const leftTitleBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, leftArea.left, leftArea.top, leftArea.width, leftTitleHeight);
  leftTitleBar.getFill().setSolidFill(CONFIG.COLORS.primary_color);
  leftTitleBar.getBorder().setTransparent();
  setStyledText(leftTitleBar, data.leftTitle || '選択肢A', { size: CONFIG.FONTS.sizes.laneTitle, bold: true, color: CONFIG.COLORS.text_on_primary, align: SlidesApp.ParagraphAlignment.CENTER });

  // 右側のタイトルヘッダー
  const rightTitleBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, rightArea.left, rightArea.top, rightArea.width, leftTitleHeight);
  rightTitleBar.getFill().setSolidFill(CONFIG.COLORS.primary_color);
  rightTitleBar.getBorder().setTransparent();
  setStyledText(rightTitleBar, data.rightTitle || '選択肢B', { size: CONFIG.FONTS.sizes.laneTitle, bold: true, color: CONFIG.COLORS.text_on_primary, align: SlidesApp.ParagraphAlignment.CENTER });

  // 左側のカード
  const leftCards = Array.isArray(data.leftCards) ? data.leftCards : [];
  const leftCardArea = { left: leftArea.left, top: leftArea.top + leftTitleHeight, width: leftArea.width, height: leftArea.height - leftTitleHeight };
  drawCardList(slide, layout, leftCardArea, leftCards);

  // 右側のカード
  const rightCards = Array.isArray(data.rightCards) ? data.rightCards : [];
  const rightCardArea = { left: rightArea.left, top: rightArea.top + leftTitleHeight, width: rightArea.width, height: rightArea.height - leftTitleHeight };
  drawCardList(slide, layout, rightCardArea, rightCards);

  drawBottomBarAndFooter(slide, layout, pageNum);
}

// contentProgress（コンテンツ＋進捗）
function createContentProgressSlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
  drawStandardTitleHeader(slide, layout, 'contentProgressSlide', data.title);
  const dy = drawSubheadIfAny(slide, layout, 'contentProgressSlide', data.subhead);

  // 左側: コンテンツエリア（箇条書きまたはカード）
  const contentArea = offsetRect(layout.getRect('contentProgressSlide.contentArea'), 0, dy);
  const points = Array.isArray(data.points) ? data.points : [];
  const cards = Array.isArray(data.cards) ? data.cards : [];
  
  // カード形式を優先、なければ箇条書き
  if (cards.length > 0) {
    drawCardList(slide, layout, contentArea, cards);
  } else if (points.length > 0) {
    const contentShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, contentArea.left, contentArea.top, contentArea.width, contentArea.height);
    setBulletsWithInlineStyles(contentShape, points);
  }

  // 右側: 進捗エリア
  const progressArea = offsetRect(layout.getRect('contentProgressSlide.progressArea'), 0, dy);
  const progressItems = Array.isArray(data.progress) ? data.progress : [];
  
  if (progressItems.length > 0) {
    const n = progressItems.length;
    const rowHeight = progressArea.height / n;
    
    for (let i = 0; i < n; i++) {
      const item = progressItems[i];
      const rowCenterY = progressArea.top + i * rowHeight + rowHeight / 2;
      const textHeight = layout.pxToPt(18);
      const barHeight = layout.pxToPt(14);
      
      // 全要素を行の中央に配置するための基準Y座標を計算
      const textY = rowCenterY - textHeight / 2;
      const barY = rowCenterY - barHeight / 2;
      
      // ラベル
      const label = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, progressArea.left, textY, layout.pxToPt(120), textHeight);
      setStyledText(label, String(item.label || ''), { size: CONFIG.FONTS.sizes.body });
      try { label.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch(e){}
      
      // プログレスバー
      const barLeft = progressArea.left + layout.pxToPt(130);
      const barWidth = progressArea.width - layout.pxToPt(200);
      const barBG = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, barLeft, barY, barWidth, barHeight);
      barBG.getFill().setSolidFill(CONFIG.COLORS.faint_gray);
      barBG.getBorder().setTransparent();
      
      const percent = Math.max(0, Math.min(100, Number(item.percent || 0)));
      if (percent > 0) {
        const barFG = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, barLeft, barY, barWidth * (percent/100), barHeight);
        barFG.getFill().setSolidFill(CONFIG.COLORS.primary_color);
        barFG.getBorder().setTransparent();
      }
      
      // パーセント表示
      const pctShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, barLeft + barWidth + layout.pxToPt(10), textY, layout.pxToPt(60), textHeight);
      pctShape.getText().setText(String(percent) + '%');
      applyTextStyle(pctShape.getText(), { size: CONFIG.FONTS.sizes.small, color: CONFIG.COLORS.neutral_gray });
      try { pctShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch(e){}
    }
  }

  drawBottomBarAndFooter(slide, layout, pageNum);
}

// カードリスト描画のヘルパー関数
function drawCardList(slide, layout, area, cards) {
  if (!cards || cards.length === 0) return;
  
  const cardCount = Math.min(cards.length, 4);
  const gap = layout.pxToPt(12);
  const cardHeight = (area.height - gap * (cardCount - 1)) / cardCount;
  
  for (let i = 0; i < cardCount; i++) {
    const card = cards[i] || {};
    const y = area.top + i * (cardHeight + gap);
    
    const cardShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, y, area.width, cardHeight);
    cardShape.getFill().setSolidFill(CONFIG.COLORS.card_bg);
    cardShape.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    cardShape.getBorder().setWeight(1);
    
    const padding = layout.pxToPt(12);
    const title = String(card.title || '');
    const desc = String(card.desc || '');
    
    if (title.length > 0 && desc.length > 0) {
      const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left + padding, y + layout.pxToPt(8), area.width - padding * 2, layout.pxToPt(16));
      titleShape.getFill().setTransparent();
      titleShape.getBorder().setTransparent();
      setStyledText(titleShape, title, { size: 12, bold: true });
      
      const descShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left + padding, y + layout.pxToPt(28), area.width - padding * 2, cardHeight - layout.pxToPt(36));
      descShape.getFill().setTransparent();
      descShape.getBorder().setTransparent();
      setStyledText(descShape, desc, { size: 11, color: CONFIG.COLORS.text_primary });
    } else if (title.length > 0) {
      const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left + padding, y, area.width - padding * 2, cardHeight);
      titleShape.getFill().setTransparent();
      titleShape.getBorder().setTransparent();
      titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      setStyledText(titleShape, title, { size: 12, bold: true });
    }
  }
}

// timelineCards（タイムライン＋カード）
function createTimelineCardsSlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
  drawStandardTitleHeader(slide, layout, 'timelineCardsSlide', data.title);
  const dy = drawSubheadIfAny(slide, layout, 'timelineCardsSlide', data.subhead);

  // 上部: タイムラインエリア（横幅フル活用）
  const timelineArea = offsetRect(layout.getRect('timelineCardsSlide.timelineArea'), 0, dy);
  const timeline = Array.isArray(data.timeline) ? data.timeline : [];
  
  if (timeline.length > 0) {
    const inner = layout.pxToPt(80);
    const baseY = timelineArea.top + timelineArea.height * 0.65;
    const leftX = timelineArea.left + inner;
    const rightX = timelineArea.left + timelineArea.width - inner;
    
    // タイムライン描画
    const line = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, leftX, baseY - layout.pxToPt(1), rightX - leftX, layout.pxToPt(2));
    line.getFill().setSolidFill(CONFIG.COLORS.neutral_gray);
    line.getBorder().setTransparent();
    
    const dotR = layout.pxToPt(8);
    const gap = (rightX - leftX) / Math.max(1, (timeline.length - 1));
    
    timeline.forEach((m, i) => {
      const x = leftX + gap * i - dotR / 2;
      const dot = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x, baseY - dotR / 2, dotR, dotR);
      
      // 時系列順で徐々に濃くなる色計算
      const progress = timeline.length > 1 ? i / (timeline.length - 1) : 0;
      const brightness = 1.5 - (progress * 0.8);
      dot.getFill().setSolidFill(adjustColorBrightness(CONFIG.COLORS.primary_color, brightness));
      dot.getBorder().setTransparent();
      
      // ラベルテキスト（上部）
      const labelShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x - layout.pxToPt(50) + dotR/2, baseY - layout.pxToPt(50), layout.pxToPt(100), layout.pxToPt(18));
      labelShape.getFill().setTransparent();
      labelShape.getBorder().setTransparent();
      setStyledText(labelShape, String(m.label || ''), { size: CONFIG.FONTS.sizes.small, bold: true, align: SlidesApp.ParagraphAlignment.CENTER });
      
      // 日付テキスト（下部）
      const dateShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x - layout.pxToPt(50) + dotR/2, baseY + layout.pxToPt(15), layout.pxToPt(100), layout.pxToPt(16));
      dateShape.getFill().setTransparent();
      dateShape.getBorder().setTransparent();
      setStyledText(dateShape, String(m.date || ''), { size: CONFIG.FONTS.sizes.chip, color: CONFIG.COLORS.neutral_gray, align: SlidesApp.ParagraphAlignment.CENTER });
    });
  }

  // 下部: カードエリア（横並びカード）
  const cardArea = offsetRect(layout.getRect('timelineCardsSlide.cardArea'), 0, dy);
  const cards = Array.isArray(data.cards) ? data.cards : [];
  
  if (cards.length > 0) {
    drawTimelineCardGrid(slide, layout, cardArea, cards);
  }

  drawBottomBarAndFooter(slide, layout, pageNum);
}

// timelineCards用の横並びカードグリッド描画
function drawTimelineCardGrid(slide, layout, area, cards) {
  if (!cards || cards.length === 0) return;
  
  const cardCount = Math.min(cards.length, 4);
  const gap = layout.pxToPt(16);
  const cardWidth = (area.width - gap * (cardCount - 1)) / cardCount;
  
  for (let i = 0; i < cardCount; i++) {
    const card = cards[i] || {};
    const x = area.left + i * (cardWidth + gap);
    
    const cardShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, area.top, cardWidth, area.height);
    cardShape.getFill().setSolidFill(CONFIG.COLORS.card_bg);
    cardShape.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    cardShape.getBorder().setWeight(1);
    
    const padding = layout.pxToPt(12);
    const title = String(card.title || '');
    const desc = String(card.desc || '');
    
    if (title.length > 0 && desc.length > 0) {
      const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + padding, area.top + layout.pxToPt(12), cardWidth - padding * 2, layout.pxToPt(20));
      titleShape.getFill().setTransparent();
      titleShape.getBorder().setTransparent();
      setStyledText(titleShape, title, { size: 12, bold: true });
      
      const descShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + padding, area.top + layout.pxToPt(35), cardWidth - padding * 2, area.height - layout.pxToPt(48));
      descShape.getFill().setTransparent();
      descShape.getBorder().setTransparent();
      setStyledText(descShape, desc, { size: 11, color: CONFIG.COLORS.text_primary });
    } else if (title.length > 0) {
      const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x + padding, area.top, cardWidth - padding * 2, area.height);
      titleShape.getFill().setTransparent();
      titleShape.getBorder().setTransparent();
      titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      setStyledText(titleShape, title, { size: 12, bold: true });
    }
  }
}

// iconCards（アイコン付きカード）
function createIconCardsSlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
  drawStandardTitleHeader(slide, layout, 'iconCardsSlide', data.title);
  const dy = drawSubheadIfAny(slide, layout, 'iconCardsSlide', data.subhead);

  const area = offsetRect(layout.getRect('iconCardsSlide.gridArea'), 0, dy);
  const items = Array.isArray(data.items) ? data.items : [];
  const cols = Math.min(3, Math.max(2, items.length <= 4 ? 2 : 3));
  const gap = layout.pxToPt(16);
  const rows = Math.ceil(items.length / cols);
  const cardW = (area.width - gap * (cols - 1)) / cols;
  const cardH = Math.max(layout.pxToPt(100), (area.height - gap * (rows - 1)) / rows);

  for (let idx = 0; idx < items.length; idx++) {
    const r = Math.floor(idx / cols), c = idx % cols;
    const left = area.left + c * (cardW + gap);
    const top = area.top + r * (cardH + gap);

    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, left, top, cardW, cardH);
    card.getFill().setSolidFill(CONFIG.COLORS.card_bg);
    card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    card.getBorder().setWeight(1);

    const item = items[idx];
    const icon = String(item.icon || '🔧');
    const title = String(item.title || '');
    const desc = String(item.desc || '');

    // アイコン（上部中央）
    const iconShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left + layout.pxToPt(10), top + layout.pxToPt(12), cardW - layout.pxToPt(20), layout.pxToPt(32));
    iconShape.getFill().setTransparent();
    iconShape.getBorder().setTransparent();
    setStyledText(iconShape, icon, { size: 24, align: SlidesApp.ParagraphAlignment.CENTER });

    // タイトル（中央）
    const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left + layout.pxToPt(10), top + layout.pxToPt(56), cardW - layout.pxToPt(20), layout.pxToPt(20));
    titleShape.getFill().setTransparent();
    titleShape.getBorder().setTransparent();
    setStyledText(titleShape, title, { size: 14, bold: true, align: SlidesApp.ParagraphAlignment.CENTER });

    // 説明文（下部）
    const descShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left + layout.pxToPt(10), top + layout.pxToPt(84), cardW - layout.pxToPt(20), cardH - layout.pxToPt(94));
    descShape.getFill().setTransparent();
    descShape.getBorder().setTransparent();
    setStyledText(descShape, desc, { size: 11, color: CONFIG.COLORS.text_primary, align: SlidesApp.ParagraphAlignment.CENTER });
  }

  drawBottomBarAndFooter(slide, layout, pageNum);
}

// statsCompare（数値比較）
function createRoadmapTimelineSlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
  drawStandardTitleHeader(slide, layout, 'roadmapTimelineSlide', data.title);
  const dy = drawSubheadIfAny(slide, layout, 'roadmapTimelineSlide', data.subhead);

  const timelineArea = offsetRect(layout.getRect('roadmapTimelineSlide.timelineArea'), 0, dy);
  const detailArea = offsetRect(layout.getRect('roadmapTimelineSlide.detailArea'), 0, dy);
  const phases = Array.isArray(data.phases) ? data.phases : [];

  if (phases.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum);
    return;
  }

  // 上部：フェーズタイムライン
  const phaseWidth = timelineArea.width / phases.length;
  const baseY = timelineArea.top + timelineArea.height * 0.6;

  // タイムライン横線
  const line = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, timelineArea.left, baseY - layout.pxToPt(1), timelineArea.width, layout.pxToPt(2));
  line.getFill().setSolidFill(CONFIG.COLORS.neutral_gray);
  line.getBorder().setTransparent();

  phases.forEach((phase, i) => {
    const x = timelineArea.left + i * phaseWidth + phaseWidth / 2;
    const dotR = layout.pxToPt(10);

    // フェーズドット
    const dot = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x - dotR / 2, baseY - dotR / 2, dotR, dotR);
    let dotColor = CONFIG.COLORS.primary_color;
    if (phase.status === 'completed') dotColor = adjustColorBrightness(CONFIG.COLORS.primary_color, 1.2);
    if (phase.status === 'planned') dotColor = adjustColorBrightness(CONFIG.COLORS.primary_color, 0.6);
    dot.getFill().setSolidFill(dotColor);
    dot.getBorder().setTransparent();

    // フェーズラベル（上部）
    const labelShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x - layout.pxToPt(60), baseY - layout.pxToPt(50), layout.pxToPt(120), layout.pxToPt(18));
    labelShape.getFill().setTransparent();
    labelShape.getBorder().setTransparent();
    setStyledText(labelShape, String(phase.label || ''), { size: 12, bold: true, align: SlidesApp.ParagraphAlignment.CENTER });

    // 期間（下部）
    const periodShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x - layout.pxToPt(60), baseY + layout.pxToPt(15), layout.pxToPt(120), layout.pxToPt(16));
    periodShape.getFill().setTransparent();
    periodShape.getBorder().setTransparent();
    setStyledText(periodShape, String(phase.period || ''), { size: 10, color: CONFIG.COLORS.neutral_gray, align: SlidesApp.ParagraphAlignment.CENTER });
  });

  // 下部：マイルストーン詳細
  const currentPhase = phases.find(p => p.status === 'current') || phases[0];
  if (currentPhase && Array.isArray(currentPhase.milestones)) {
    const milestones = currentPhase.milestones.slice(0, 4); // 最大4項目
    const milestoneHeight = detailArea.height / Math.max(1, milestones.length);

    milestones.forEach((milestone, i) => {
      const y = detailArea.top + i * milestoneHeight;
      const padding = layout.pxToPt(15);

      const milestoneCard = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, detailArea.left, y + layout.pxToPt(2), detailArea.width, milestoneHeight - layout.pxToPt(4));
      milestoneCard.getFill().setSolidFill(CONFIG.COLORS.background_gray);
      milestoneCard.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
      milestoneCard.getBorder().setWeight(1);

      const milestoneText = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, detailArea.left + padding, y + layout.pxToPt(5), detailArea.width - padding * 2, milestoneHeight - layout.pxToPt(6));
      milestoneText.getFill().setTransparent();
      milestoneText.getBorder().setTransparent();
      setStyledText(milestoneText, `• ${String(milestone || '')}`, { size: 12, color: CONFIG.COLORS.text_primary });
      try { milestoneText.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch(e){}
    });
  }

  drawBottomBarAndFooter(slide, layout, pageNum);
}

// imageGallery（画像ギャラリー）
function createImageGallerySlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
  drawStandardTitleHeader(slide, layout, 'imageGallerySlide', data.title);
  const dy = drawSubheadIfAny(slide, layout, 'imageGallerySlide', data.subhead);

  const images = normalizeImages(data.images || []);
  if (images.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum);
    return;
  }

  const layoutType = data.layout || 'grid';
  
  if (layoutType === 'single') {
    // 単一画像（中央大きく表示）
    const area = offsetRect(layout.getRect('imageGallerySlide.singleImage'), 0, dy);
    renderSingleImage(slide, layout, area, images[0]);
  } else if (layoutType === 'showcase') {
    // ショーケース（メイン1枚 + サイド複数）
    const mainArea = offsetRect(layout.getRect('imageGallerySlide.showcaseMain'), 0, dy);
    const sideArea = offsetRect(layout.getRect('imageGallerySlide.showcaseSide'), 0, dy);
    
    // メイン画像
    renderSingleImage(slide, layout, mainArea, images[0]);
    
    // サイド画像（最大3枚）
    const sideImages = images.slice(1, 4);
    if (sideImages.length > 0) {
      renderImageGrid(slide, layout, sideArea, sideImages, 1);
    }
  } else {
    // グリッド（デフォルト）
    const area = offsetRect(layout.getRect('imageGallerySlide.gridArea'), 0, dy);
    const cols = images.length === 1 ? 1 : (images.length <= 4 ? 2 : 3);
    renderImageGrid(slide, layout, area, images, cols);
  }

  drawBottomBarAndFooter(slide, layout, pageNum);
}

// 単一画像の描画
function renderSingleImage(slide, layout, area, imageData) {
  if (!imageData || !imageData.url) return;
  
  try {
    const img = slide.insertImage(imageData.url);
    const imgAspect = img.getHeight() / img.getWidth();
    const areaAspect = area.height / area.width;
    
    let imgWidth, imgHeight;
    if (imgAspect > areaAspect) {
      // 画像が縦長 → 高さ基準
      imgHeight = area.height;
      imgWidth = imgHeight / imgAspect;
    } else {
      // 画像が横長 → 幅基準  
      imgWidth = area.width;
      imgHeight = imgWidth * imgAspect;
    }
    
    const left = area.left + (area.width - imgWidth) / 2;
    const top = area.top + (area.height - imgHeight) / 2;
    
    img.setLeft(left).setTop(top).setWidth(imgWidth).setHeight(imgHeight);
    
    // キャプション追加
    if (imageData.caption) {
      const captionShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
        area.left, area.top + area.height + layout.pxToPt(8), 
        area.width, layout.pxToPt(20));
      captionShape.getFill().setTransparent();
      captionShape.getBorder().setTransparent();
      setStyledText(captionShape, imageData.caption, { 
        size: CONFIG.FONTS.sizes.small, 
        color: CONFIG.COLORS.neutral_gray, 
        align: SlidesApp.ParagraphAlignment.CENTER 
      });
    }
  } catch(e) {
    // 画像読み込み失敗時のフォールバック
    const placeholder = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, area.top, area.width, area.height);
    placeholder.getFill().setSolidFill(CONFIG.COLORS.faint_gray);
    placeholder.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    placeholder.getBorder().setWeight(1);
    
    const errorText = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, area.top, area.width, area.height);
    errorText.getFill().setTransparent();
    errorText.getBorder().setTransparent();
    setStyledText(errorText, '画像を読み込めませんでした', { 
      size: CONFIG.FONTS.sizes.body, 
      color: CONFIG.COLORS.neutral_gray, 
      align: SlidesApp.ParagraphAlignment.CENTER 
    });
    try { errorText.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch(e){}
  }
}

// グリッド画像の描画
function renderImageGrid(slide, layout, area, images, cols) {
  if (!images || images.length === 0) return;
  
  const gap = layout.pxToPt(12);
  const rows = Math.ceil(images.length / cols);
  const cellW = (area.width - gap * (cols - 1)) / cols;
  const cellH = (area.height - gap * (rows - 1)) / rows;
  
  for (let i = 0; i < images.length; i++) {
    const r = Math.floor(i / cols);
    const c = i % cols;
    const left = area.left + c * (cellW + gap);
    const top = area.top + r * (cellH + gap);
    
    const cellArea = { left, top, width: cellW, height: cellH };
    renderSingleImage(slide, layout, cellArea, images[i]);
  }
}


Object.assign(slideGenerators, {
  cards: createCardsSlide,
  headerCards: createHeaderCardsSlide,
  bulletCards: createBulletCardsSlide,
  faq: createFaqSlide,
});
