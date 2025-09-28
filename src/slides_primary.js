function createTitleSlide(slide, data, layout, pageNum, settings) {
  setBackgroundImageFromUrl(slide, layout, CONFIG.BACKGROUND_IMAGES.title, CONFIG.COLORS.background_white);
  const logoRect = layout.getRect('titleSlide.logo');
  try {
    if (CONFIG.LOGOS.header) {
      const imageData = insertImageFromUrlOrFileId(CONFIG.LOGOS.header);
      if (imageData) {
        const logo = slide.insertImage(imageData);
        const aspect = logo.getHeight() / logo.getWidth();
        logo.setLeft(logoRect.left).setTop(logoRect.top).setWidth(logoRect.width).setHeight(logoRect.width * aspect);
      }
    }
  } catch (e) {
    Logger.log(`Title logo error: ${e.message}`);
  }
  const titleRect = layout.getRect('titleSlide.title');
  const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, titleRect.left, titleRect.top, titleRect.width, titleRect.height);
  setStyledText(titleShape, data.title, {
    size: CONFIG.FONTS.sizes.title,
    bold: true
  });
  
  // 日付カラムの条件付き生成
  if (settings.showDateColumn) {
    const dateRect = layout.getRect('titleSlide.date');
    const dateShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, dateRect.left, dateRect.top, dateRect.width, dateRect.height);
    dateShape.getText().setText(data.date || '');
    applyTextStyle(dateShape.getText(), {
      size: CONFIG.FONTS.sizes.date
    });
  }
  
  if (settings.showBottomBar) {
    drawBottomBar(slide, layout, settings);
  }
}

function createSectionSlide(slide, data, layout, pageNum, settings) {
  setBackgroundImageFromUrl(slide, layout, CONFIG.BACKGROUND_IMAGES.section, CONFIG.COLORS.background_gray);
  __SECTION_COUNTER++;
  const parsedNum = (() => {
    if (Number.isFinite(data.sectionNo)) {
      return Number(data.sectionNo);
    }
    const m = String(data.title || '').match(/^\s*(\d+)[\.．]/);
    return m ? Number(m[1]) : __SECTION_COUNTER;
  })();
  const num = String(parsedNum).padStart(2, '0');
  const ghostRect = layout.getRect('sectionSlide.ghostNum');
  const ghost = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, ghostRect.left, ghostRect.top, ghostRect.width, ghostRect.height);
  // ゴースト数字に半透明効果を適用
  ghost.getText().setText(num);
  const ghostTextStyle = ghost.getText().getTextStyle();
  ghostTextStyle.setFontFamily(CONFIG.FONTS.family)
    .setFontSize(CONFIG.FONTS.sizes.ghostNum)
    .setBold(true);
  
  // 透明度を適用（座布団と同様の15%不透明度）
  try {
    // setForegroundColorWithAlphaを使用して透明度付きの色を設定
    ghostTextStyle.setForegroundColorWithAlpha(CONFIG.COLORS.ghost_gray, 0.15);
  } catch (e) {
    // フォールバック：通常の色設定
    ghostTextStyle.setForegroundColor(CONFIG.COLORS.ghost_gray);
  }
  try {
    ghost.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  } catch (e) {}
  const titleRect = layout.getRect('sectionSlide.title');
  const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, titleRect.left, titleRect.top, titleRect.width, titleRect.height);
  titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  setStyledText(titleShape, data.title, {
    size: CONFIG.FONTS.sizes.sectionTitle,
    bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER
  });
  addCucFooter(slide, layout, pageNum, settings);
}

function createClosingSlide(slide, data, layout, pageNum, settings) {
  setBackgroundImageFromUrl(slide, layout, CONFIG.BACKGROUND_IMAGES.closing, CONFIG.COLORS.background_white);
  try {
    if (CONFIG.LOGOS.closing) {
      const imageData = insertImageFromUrlOrFileId(CONFIG.LOGOS.closing);
      if (imageData) {
        const image = slide.insertImage(imageData);
        const imgW_pt = layout.pxToPt(450) * layout.scaleX;
        const aspect = image.getHeight() / image.getWidth();
        image.setWidth(imgW_pt).setHeight(imgW_pt * aspect);
        image.setLeft((layout.pageW_pt - imgW_pt) / 2).setTop((layout.pageH_pt - (imgW_pt * aspect)) / 2);
      }
    }
  } catch (e) {
    Logger.log(`Closing logo error: ${e.message}`);
  }
}

function createContentSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'contentSlide', data.title, settings);
  
  const dy = drawSubheadIfAny(slide, layout, 'contentSlide', data.subhead);
  const isAgenda = isAgendaTitle(data.title || '');
  let points = Array.isArray(data.points) ? data.points.slice(0) : [];
  if (isAgenda && points.length === 0) {
    points = buildAgendaFromSlideData();
    if (points.length === 0) {
      points = ['本日の目的', '進め方', '次のアクション'];
    }
  }
  const hasImages = Array.isArray(data.images) && data.images.length > 0;
  const isTwo = !!(data.twoColumn || data.columns);
  if ((isTwo && (data.columns || points)) || (!isTwo && points && points.length > 0)) {
    if (isTwo) {
      let L = [],
        R = [];
      if (Array.isArray(data.columns) && data.columns.length === 2) {
        L = data.columns[0] || [];
        R = data.columns[1] || [];
      } else {
        const mid = Math.ceil(points.length / 2);
        L = points.slice(0, mid);
        R = points.slice(mid);
      }
      // 小見出しの高さに応じて2カラムエリアを動的に調整
      const baseLeftRect = layout.getRect('contentSlide.twoColLeft');
      const baseRightRect = layout.getRect('contentSlide.twoColRight');
      const adjustedLeftRect = adjustAreaForSubhead(baseLeftRect, data.subhead, layout);
      const adjustedRightRect = adjustAreaForSubhead(baseRightRect, data.subhead, layout);
      
      const leftRect = offsetRect(adjustedLeftRect, 0, dy);
      const rightRect = offsetRect(adjustedRightRect, 0, dy);
      
      createContentCushion(slide, leftRect, settings, layout);
      createContentCushion(slide, rightRect, settings, layout);
      
      // テキストボックスを座布団の内側に配置（パディングを追加）
      const padding = layout.pxToPt(20); // 20pxのパディング
      const leftTextRect = {
        left: leftRect.left + padding,
        top: leftRect.top + padding,
        width: leftRect.width - (padding * 2),
        height: leftRect.height - (padding * 2)
      };
      const rightTextRect = {
        left: rightRect.left + padding,
        top: rightRect.top + padding,
        width: rightRect.width - (padding * 2),
        height: rightRect.height - (padding * 2)
      };
      const leftShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, leftTextRect.left, leftTextRect.top, leftTextRect.width, leftTextRect.height);
      const rightShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rightTextRect.left, rightTextRect.top, rightTextRect.width, rightTextRect.height);
      setBulletsWithInlineStyles(leftShape, L);
      setBulletsWithInlineStyles(rightShape, R);
    } else {
      // 小見出しの高さに応じて本文エリアを動的に調整
      const baseBodyRect = layout.getRect('contentSlide.body');
      const adjustedBodyRect = adjustAreaForSubhead(baseBodyRect, data.subhead, layout);
      const bodyRect = offsetRect(adjustedBodyRect, 0, dy);
      
      createContentCushion(slide, bodyRect, settings, layout);
      
      if (isAgenda) {
        drawNumberedItems(slide, layout, bodyRect, points, settings);
      } else {
        // テキストボックスを座布団の内側に配置（パディングを追加）
        const padding = layout.pxToPt(20); // 20pxのパディング
        const textRect = {
          left: bodyRect.left + padding,
          top: bodyRect.top + padding,
          width: bodyRect.width - (padding * 2),
          height: bodyRect.height - (padding * 2)
        };
        const bodyShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textRect.left, textRect.top, textRect.width, textRect.height);
        setBulletsWithInlineStyles(bodyShape, points);
      }
    }
  }
  // 画像はテキストがない場合のみ表示（imageTextパターンを推奨）
  if (hasImages && !points.length && !isTwo) {
    const baseArea = layout.getRect('contentSlide.body');
    const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
    const area = offsetRect(adjustedArea, 0, dy);
    
    // 画像表示時も座布団を作成
    createContentCushion(slide, area, settings, layout);
    renderImagesInArea(slide, layout, area, normalizeImages(data.images));
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

function createCompareSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'compareSlide', data.title, settings);
  
  const dy = drawSubheadIfAny(slide, layout, 'compareSlide', data.subhead);
  
  // 小見出しの高さに応じて比較ボックスエリアを動的に調整
  const baseLeftBox = layout.getRect('compareSlide.leftBox');
  const baseRightBox = layout.getRect('compareSlide.rightBox');
  const adjustedLeftBox = adjustAreaForSubhead(baseLeftBox, data.subhead, layout);
  const adjustedRightBox = adjustAreaForSubhead(baseRightBox, data.subhead, layout);
  
  const leftBox = offsetRect(adjustedLeftBox, 0, dy);
  const rightBox = offsetRect(adjustedRightBox, 0, dy);
  drawCompareBox(slide, layout, leftBox, data.leftTitle || '選択肢A', data.leftItems || [], settings, true);  // 左側
  drawCompareBox(slide, layout, rightBox, data.rightTitle || '選択肢B', data.rightItems || [], settings, false); // 右側
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

function createProcessSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'processSlide', data.title, settings);
  const dy = drawSubheadIfAny(slide, layout, 'processSlide', data.subhead);
  
  // 小見出しの高さに応じてプロセスエリアを動的に調整
  const baseArea = layout.getRect('processSlide.area');
  const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
  const area = offsetRect(adjustedArea, 0, dy);
  
  const steps = Array.isArray(data.steps) ? data.steps.slice(0, 4) : []; // 4ステップまで対応
  if (steps.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }

  // このスライド専用の背景色を定義
  const processBodyBgColor = generateTintedGray(settings.primaryColor, 30, 94);

  // ステップ数に応じてサイズを可変調整
  const n = steps.length;
  let boxHPx, arrowHPx, fontSize;
  
  if (n <= 2) {
    boxHPx = 100; // 2ステップ以下は大きめ
    arrowHPx = 25;
    fontSize = 16;
  } else if (n === 3) {
    boxHPx = 80; // 3ステップは標準サイズ
    arrowHPx = 20;
    fontSize = 16;
  } else {
    boxHPx = 65; // 4ステップは小さめ
    arrowHPx = 15;
    fontSize = 14;
  }

  // Processカラーグラデーション生成（上から下に濃くなる）
  const processColors = generateProcessColors(settings.primaryColor, n);

  const startY = area.top + layout.pxToPt(10);
  let currentY = startY;
  const boxHPt = layout.pxToPt(boxHPx),
    arrowHPt = layout.pxToPt(arrowHPx);
  const headerWPt = layout.pxToPt(120);
  const bodyLeft = area.left + headerWPt;
  const bodyWPt = area.width - headerWPt;

  for (let i = 0; i < n; i++) {
    const cleanText = String(steps[i] || '').replace(/^\s*\d+[\.\s]*/, '');
    const header = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, currentY, headerWPt, boxHPt);
    header.getFill().setSolidFill(processColors[i]); // グラデーションカラー適用
    header.getBorder().setTransparent();
    setStyledText(header, `STEP ${i + 1}`, {
      size: fontSize,
      bold: true,
      color: CONFIG.COLORS.background_white,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
    try {
      header.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
    const body = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, bodyLeft, currentY, bodyWPt, boxHPt);
    
    // 専用色を背景に設定
    body.getFill().setSolidFill(processBodyBgColor);
    
    body.getBorder().setTransparent();
    const textShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, bodyLeft + layout.pxToPt(20), currentY, bodyWPt - layout.pxToPt(40), boxHPt);
    setStyledText(textShape, cleanText, {
      size: fontSize
    });
    try {
      textShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
    currentY += boxHPt;
    if (i < n - 1) {
      const arrowLeft = area.left + headerWPt / 2 - layout.pxToPt(8);
      const arrow = slide.insertShape(SlidesApp.ShapeType.DOWN_ARROW, arrowLeft, currentY, layout.pxToPt(16), arrowHPt);
      arrow.getFill().setSolidFill(CONFIG.COLORS.process_arrow);
      arrow.getBorder().setTransparent();
      currentY += arrowHPt;
    }
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

function createProcessListSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'processSlide', data.title, settings);
  const dy = drawSubheadIfAny(slide, layout, 'processSlide', data.subhead);

  const area = offsetRect(layout.getRect('processSlide.area'), 0, dy);
  const steps = Array.isArray(data.steps) ? data.steps : [];
  if (steps.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }

  const n = Math.max(1, steps.length);

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

    // 元のプロセステキストから先頭の数字を除去
    let cleanText = String(steps[i] || '');
    cleanText = cleanText.replace(/^\s*\d+[\.\s]*/, '');

    const txt = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, cx + layout.pxToPt(28), cy - layout.pxToPt(16), area.width - layout.pxToPt(70), layout.pxToPt(32));
    setStyledText(txt, cleanText, { size: CONFIG.FONTS.sizes.processStep });
    try { txt.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch(e){}
  }

  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

function createTimelineSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'timelineSlide', data.title, settings);
  const dy = drawSubheadIfAny(slide, layout, 'timelineSlide', data.subhead);
  
  // 小見出しの高さに応じてタイムラインエリアを動的に調整
  const baseArea = layout.getRect('timelineSlide.area');
  const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
  const area = offsetRect(adjustedArea, 0, dy);
  
  const milestones = Array.isArray(data.milestones) ? data.milestones : [];
  if (milestones.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }
  const inner = layout.pxToPt(80),
    baseY = area.top + area.height * 0.50;
  const leftX = area.left + inner,
    rightX = area.left + area.width - inner;
  const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, leftX, baseY, rightX, baseY);
  line.getLineFill().setSolidFill(CONFIG.COLORS.faint_gray);
  line.setWeight(2);
  const dotR = layout.pxToPt(10);
  const gap = (milestones.length > 1) ? (rightX - leftX) / (milestones.length - 1) : 0;
  const cardW_pt = layout.pxToPt(180); // カード幅
  const vOffset = layout.pxToPt(40); // 縦オフセット
  const headerHeight = layout.pxToPt(28); // ヘッダー高さを縮小
  const bodyHeight = layout.pxToPt(80); // ボディ高さを固定化（十分な余裕を確保）
  const cardPadding = layout.pxToPt(8);
  milestones.forEach((m, i) => {
    const x = leftX + gap * i;
    const isAbove = i % 2 === 0;

    // テキスト取得（サイズは固定）
    const dateText = String(m.date || '');
    const labelText = String(m.label || '');
    
    // カード全体の高さを固定
    const cardH_pt = headerHeight + bodyHeight;
    
    // カード位置計算
    const cardLeft = x - (cardW_pt / 2);
    const cardTop = isAbove ? (baseY - vOffset - cardH_pt) : (baseY + vOffset);
    
    // タイムラインカラーを取得
    const timelineColors = generateTimelineCardColors(settings.primaryColor, milestones.length);
    
    // ヘッダー部分（日付表示）
    const headerShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, cardLeft, cardTop, cardW_pt, headerHeight);
    headerShape.getFill().setSolidFill(timelineColors[i]);
    headerShape.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    
    // ボディ部分（テキスト表示）
    const bodyShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, cardLeft, cardTop + headerHeight, cardW_pt, bodyHeight);
    bodyShape.getFill().setSolidFill(CONFIG.COLORS.background_white);
    bodyShape.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    const connectorY_start = isAbove ? (cardTop + cardH_pt) : baseY;
    const connectorY_end = isAbove ? baseY : cardTop;
    const connector = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x, connectorY_start, x, connectorY_end);
    connector.getLineFill().setSolidFill(CONFIG.COLORS.neutral_gray);
    connector.setWeight(1);
    // タイムライン上のドット
    const dot = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x - dotR / 2, baseY - dotR / 2, dotR, dotR);
    dot.getFill().setSolidFill(timelineColors[i]);
    dot.getBorder().setTransparent();
    // ヘッダーテキスト（日付）- 中央寄せ強化
    const headerTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
      cardLeft, cardTop, 
      cardW_pt, headerHeight);
    setStyledText(headerTextShape, dateText, {
      size: CONFIG.FONTS.sizes.body, // ボディと同じサイズに統一
      bold: true,
      color: CONFIG.COLORS.background_white,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
    try {
      headerTextShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
    
    // ボディテキスト（説明）- 動的フォントサイズ調整 + 中央寄せ強化
    let bodyFontSize = CONFIG.FONTS.sizes.body; // 14が標準
    const textLength = labelText.length;
    
    if (textLength > 40) bodyFontSize = 10;      // 長文（40文字超）は小さく
    else if (textLength > 30) bodyFontSize = 11; // やや長文（30文字超）は少し小さく
    else if (textLength > 20) bodyFontSize = 12; // 中文（20文字超）は標準より小さく
    // 短文（20文字以下）は標準サイズ(14)のまま
    
    const bodyTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
      cardLeft, cardTop + headerHeight, 
      cardW_pt, bodyHeight);
    setStyledText(bodyTextShape, labelText, {
      size: bodyFontSize,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
    try {
      bodyTextShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
  });
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

function createDiagramSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'diagramSlide', data.title, settings);
  const dy = drawSubheadIfAny(slide, layout, 'diagramSlide', data.subhead);
  const lanes = Array.isArray(data.lanes) ? data.lanes : [];
  
  // 小見出しの高さに応じてダイアグラムエリアを動的に調整
  const baseArea = layout.getRect('diagramSlide.lanesArea');
  const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
  const area = offsetRect(adjustedArea, 0, dy);
  
  const px = (p) => layout.pxToPt(p);
  const {
    laneGap_px,
    lanePad_px,
    laneTitle_h_px,
    cardGap_px,
    cardMin_h_px,
    cardMax_h_px,
    arrow_h_px,
    arrowGap_px
  } = CONFIG.DIAGRAM;
  const n = Math.max(1, lanes.length);
  const laneW = (area.width - px(laneGap_px) * (n - 1)) / n;
  const cardBoxes = [];
  for (let j = 0; j < n; j++) {
    const lane = lanes[j] || {
      title: '',
      items: []
    };
    const left = area.left + j * (laneW + px(laneGap_px));
    const lt = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, area.top, laneW, px(laneTitle_h_px));
    lt.getFill().setSolidFill(settings.primaryColor);
    lt.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.lane_border);
    setStyledText(lt, lane.title || '', {
      size: CONFIG.FONTS.sizes.laneTitle,
      bold: true,
      color: CONFIG.COLORS.background_white,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
    try {
      lt.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
    const laneBodyTop = area.top + px(laneTitle_h_px),
      laneBodyHeight = area.height - px(laneTitle_h_px);
    const laneBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, laneBodyTop, laneW, laneBodyHeight);
    laneBg.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    laneBg.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.lane_border);
    const items = Array.isArray(lane.items) ? lane.items : [];
    const availH = laneBodyHeight - px(lanePad_px) * 2,
      rows = Math.max(1, items.length);
    const idealH = (availH - px(cardGap_px) * (rows - 1)) / rows;
    const cardH = Math.max(px(cardMin_h_px), Math.min(px(cardMax_h_px), idealH));
    const firstTop = laneBodyTop + px(lanePad_px) + Math.max(0, (availH - (cardH * rows + px(cardGap_px) * (rows - 1))) / 2);
    cardBoxes[j] = [];
    for (let i = 0; i < rows; i++) {
      const cardTop = firstTop + i * (cardH + px(cardGap_px));
      const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, left + px(lanePad_px), cardTop, laneW - px(lanePad_px) * 2, cardH);
      card.getFill().setSolidFill(CONFIG.COLORS.background_white);
      card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
      setStyledText(card, items[i] || '', {
        size: CONFIG.FONTS.sizes.body
      });
      try {
        card.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      } catch (e) {}
      cardBoxes[j][i] = {
        left: left + px(lanePad_px),
        top: cardTop,
        width: laneW - px(lanePad_px) * 2,
        height: cardH
      };
    }
  }
  const maxRows = Math.max(0, ...cardBoxes.map(a => a ? a.length : 0));
  for (let j = 0; j < n - 1; j++) {
    for (let i = 0; i < maxRows; i++) {
      if (cardBoxes[j] && cardBoxes[j][i] && cardBoxes[j + 1] && cardBoxes[j + 1][i]) {
        drawArrowBetweenRects(slide, cardBoxes[j][i], cardBoxes[j + 1][i], px(arrow_h_px), px(arrowGap_px), settings);
      }
    }
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

function createCycleSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'contentSlide', data.title, settings);
  const dy = drawSubheadIfAny(slide, layout, 'contentSlide', data.subhead);
  const area = offsetRect(layout.getRect('contentSlide.body'), 0, dy);
  const items = Array.isArray(data.items) && data.items.length === 4 ? data.items : [];
  if (items.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }

  // 各アイテムのテキスト長を分析
  const textLengths = items.map(item => {
    const labelLength = (item.label || '').length;
    const subLabelLength = (item.subLabel || '').length;
    return labelLength + subLabelLength;
  });
  const maxLength = Math.max(...textLengths);
  const avgLength = textLengths.reduce((sum, len) => sum + len, 0) / textLengths.length;

  const centerX = area.left + area.width / 2;
  const centerY = area.top + area.height / 2;
  
  // 楕円半径は固定（安定した配置）
  const radiusX = area.width / 3.2;
  const radiusY = area.height / 2.6;
  
  // カードサイズの上限制限（楕円枠内に収める）
  const maxCardW = Math.min(layout.pxToPt(220), radiusX * 0.8); // 楕円半径の80%まで
  const maxCardH = Math.min(layout.pxToPt(100), radiusY * 0.6); // 楕円半径の60%まで
  
  // 文字数に基づいてカードサイズを動的調整（適度な範囲内）
  let cardW, cardH, fontSize;
  if (maxLength > 25 || avgLength > 18) {
    // 中長文対応：適度なサイズ拡張＋フォント縮小
    cardW = Math.min(layout.pxToPt(230), maxCardW);
    cardH = Math.min(layout.pxToPt(105), maxCardH);
    fontSize = 13;  // フォント縮小で文字収容力向上
  } else if (maxLength > 15 || avgLength > 10) {
    // 短中文対応：軽微なサイズ拡張
    cardW = Math.min(layout.pxToPt(215), maxCardW);
    cardH = Math.min(layout.pxToPt(95), maxCardH);
    fontSize = 14;
  } else {
    // 短文対応：従来サイズ
    cardW = layout.pxToPt(200);
    cardH = layout.pxToPt(90);
    fontSize = 16;
  }

  if (data.centerText) {
    const centerTextBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, centerX - layout.pxToPt(100), centerY - layout.pxToPt(50), layout.pxToPt(200), layout.pxToPt(100));
    setStyledText(centerTextBox, data.centerText, { size: 20, bold: true, align: SlidesApp.ParagraphAlignment.CENTER, color: CONFIG.COLORS.text_primary });
    try { centerTextBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) {}
  }

  const positions = [
    { x: centerX + radiusX, y: centerY }, // 右
    { x: centerX, y: centerY + radiusY }, // 下
    { x: centerX - radiusX, y: centerY }, // 左
    { x: centerX, y: centerY - radiusY }  // 上
  ];

  positions.forEach((pos, i) => {
    const cardX = pos.x - cardW / 2;
    const cardY = pos.y - cardH / 2;
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, cardX, cardY, cardW, cardH);
    card.getFill().setSolidFill(settings.primaryColor);
    card.getBorder().setTransparent();
    const item = items[i] || {};
    const subLabelText = item.subLabel || `${i + 1}番目`;
    const labelText = item.label || '';

    setStyledText(card, `${subLabelText}\n${labelText}`, { size: fontSize, bold: true, color: CONFIG.COLORS.background_white, align: SlidesApp.ParagraphAlignment.CENTER });
    try {
      card.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      const textRange = card.getText();
      const subLabelEnd = subLabelText.length;
      if (textRange.asString().length > subLabelEnd) {
        // subLabelのフォントサイズを少し小さく
        textRange.getRange(0, subLabelEnd).getTextStyle().setFontSize(Math.max(10, fontSize - 2));
      }
    } catch (e) {}
  });

  // 矢印の座標を決める半径の値を調整（固定配置）
  const arrowRadiusX = radiusX * 0.75;
  const arrowRadiusY = radiusY * 0.80;
  const arrowSize = layout.pxToPt(80);

  const arrowPositions = [
    // 上から右へ向かう矢印 (右上)
    { left: centerX + arrowRadiusX, top: centerY - arrowRadiusY, rotation: 90 },
    // 右から下へ向かう矢印 (右下)
    { left: centerX + arrowRadiusX, top: centerY + arrowRadiusY, rotation: 180 },
    // 下から左へ向かう矢印 (左下)
    { left: centerX - arrowRadiusX, top: centerY + arrowRadiusY, rotation: 270 },
    // 左から上へ向かう矢印 (左上)
    { left: centerX - arrowRadiusX, top: centerY - arrowRadiusY, rotation: 0 }
  ];

  arrowPositions.forEach(pos => {
    const arrow = slide.insertShape(SlidesApp.ShapeType.BENT_ARROW, pos.left - arrowSize / 2, pos.top - arrowSize / 2, arrowSize, arrowSize);
    arrow.getFill().setSolidFill(CONFIG.COLORS.ghost_gray);
    arrow.getBorder().setTransparent();
    arrow.setRotation(pos.rotation);
  });

  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

function createCardsSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'cardsSlide', data.title, settings);
  const dy = drawSubheadIfAny(slide, layout, 'cardsSlide', data.subhead);
  
  // 小見出しの高さに応じてカードグリッドエリアを動的に調整
  const baseArea = layout.getRect('cardsSlide.gridArea');
  const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
  const area = offsetRect(adjustedArea, 0, dy);
  
  const items = Array.isArray(data.items) ? data.items : [];
  const cols = Math.min(3, Math.max(2, Number(data.columns) || (items.length <= 4 ? 2 : 3)));
  const gap = layout.pxToPt(16),
    rows = Math.ceil(items.length / cols);
  const cardW = (area.width - gap * (cols - 1)) / cols,
    cardH = Math.max(layout.pxToPt(92), (area.height - gap * (rows - 1)) / rows);
  for (let idx = 0; idx < items.length; idx++) {
    const r = Math.floor(idx / cols),
      c = idx % cols;
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, area.left + c * (cardW + gap), area.top + r * (cardH + gap), cardW, cardH);
    card.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    const obj = items[idx];
    if (typeof obj === 'string') {
      setStyledText(card, obj, {
        size: CONFIG.FONTS.sizes.body
      });
    } else {
      const title = String(obj.title || ''),
        desc = String(obj.desc || '');
      if (title && desc) {
        const combined = `${title}\n\n${desc}`;
        setStyledText(card, combined, {
          size: CONFIG.FONTS.sizes.body
        });
        try {
          card.getText().getRange(0, title.length).getTextStyle().setBold(true);
        } catch (e) {}
      } else if (title) {
        setStyledText(card, title, {
          size: CONFIG.FONTS.sizes.body,
          bold: true
        });
      } else {
        setStyledText(card, desc, {
          size: CONFIG.FONTS.sizes.body
        });
      }
    }
    try {
      card.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

function createHeaderCardsSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'cardsSlide', data.title, settings);
  const dy = drawSubheadIfAny(slide, layout, 'cardsSlide', data.subhead);
  
  // 小見出しの高さに応じてヘッダーカードグリッドエリアを動的に調整
  const baseArea = layout.getRect('cardsSlide.gridArea');
  const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
  const area = offsetRect(adjustedArea, 0, dy);
  
  const items = Array.isArray(data.items) ? data.items : [];
  const cols = Math.min(3, Math.max(2, Number(data.columns) || (items.length <= 4 ? 2 : 3)));
  const gap = layout.pxToPt(16),
    rows = Math.ceil(items.length / cols);
  const cardW = (area.width - gap * (cols - 1)) / cols,
    cardH = Math.max(layout.pxToPt(92), (area.height - gap * (rows - 1)) / rows);
  for (let idx = 0; idx < items.length; idx++) {
    const r = Math.floor(idx / cols),
      c = idx % cols;
    const left = area.left + c * (cardW + gap),
      top = area.top + r * (cardH + gap);
    const titleText = String(items[idx].title || ''),
      descText = String(items[idx].desc || '');
    const headerHeight = layout.pxToPt(40);
    const headerShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top, cardW, headerHeight);
    headerShape.getFill().setSolidFill(settings.primaryColor);
    headerShape.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    const bodyShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top + headerHeight, cardW, cardH - headerHeight);
    bodyShape.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    bodyShape.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    const headerTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left, top, cardW, headerHeight);
    setStyledText(headerTextShape, titleText, {
      size: CONFIG.FONTS.sizes.body,
      bold: true,
      color: CONFIG.COLORS.background_white,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
    try {
      headerTextShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
    const bodyTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left + layout.pxToPt(12), top + headerHeight, cardW - layout.pxToPt(24), cardH - headerHeight);
    setStyledText(bodyTextShape, descText, {
      size: CONFIG.FONTS.sizes.body,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
    try {
      bodyTextShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

function createTableSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'tableSlide', data.title, settings);
  const dy = drawSubheadIfAny(slide, layout, 'tableSlide', data.subhead);
  
  // 小見出しの高さに応じてテーブルエリアを動的に調整
  const baseArea = layout.getRect('tableSlide.area');
  const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
  const area = offsetRect(adjustedArea, 0, dy);
  
  const headers = Array.isArray(data.headers) ? data.headers : [];
  const rows = Array.isArray(data.rows) ? data.rows : [];
  try {
    if (headers.length > 0) {
      const table = slide.insertTable(rows.length + 1, headers.length, area.left, area.top, area.width, area.height);
      for (let c = 0; c < headers.length; c++) {
        const cell = table.getCell(0, c);
        cell.getFill().setSolidFill(CONFIG.COLORS.table_header_bg);
        setStyledText(cell, String(headers[c] || ''), {
          bold: true,
          color: CONFIG.COLORS.text_primary,
          align: SlidesApp.ParagraphAlignment.CENTER
        });
        try {
          cell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        } catch (e) {}
      }
      for (let r = 0; r < rows.length; r++) {
        for (let c = 0; c < headers.length; c++) {
          const cell = table.getCell(r + 1, c);
          cell.getFill().setSolidFill(CONFIG.COLORS.background_white);
          setStyledText(cell, String((rows[r] || [])[c] || ''), {
            align: SlidesApp.ParagraphAlignment.CENTER
          });
          try {
            cell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
          } catch (e) {}
        }
      }
    }
  } catch (e) {
    Logger.log(`Table creation error: ${e.message}`);
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

function createProgressSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'progressSlide', data.title, settings);
  const dy = drawSubheadIfAny(slide, layout, 'progressSlide', data.subhead);
  
  // 小見出しの高さに応じてプログレスエリアを動的に調整
  const baseArea = layout.getRect('progressSlide.area');
  const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
  const area = offsetRect(adjustedArea, 0, dy);
  
  const items = Array.isArray(data.items) ? data.items : [];
  const n = Math.max(1, items.length);
  // カード型レイアウトの設定
  const cardGap = layout.pxToPt(12); // カード間の間隔
  const cardHeight = Math.max(layout.pxToPt(80), (area.height - cardGap * (n - 1)) / n);
  const cardPadding = layout.pxToPt(15);
  const barHeight = layout.pxToPt(12);
  
  const percentHeight = layout.pxToPt(30);
  const percentWidth = layout.pxToPt(120); // 幅を拡大（三桁対応）
  
  for (let i = 0; i < n; i++) {
    const cardTop = area.top + i * (cardHeight + cardGap);
    const p = Math.max(0, Math.min(100, Number(items[i].percent || 0)));
    
    // カード背景
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 
      area.left, cardTop, area.width, cardHeight);
    card.getFill().setSolidFill(CONFIG.COLORS.background_white);
    card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    
    // タスク名（左上）
    const labelHeight = layout.pxToPt(20);
    const labelWidth = area.width - percentWidth - cardPadding * 3; // パーセンテージ幅を考慮
    const label = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
      area.left + cardPadding, cardTop + cardPadding, 
      labelWidth, labelHeight);
    setStyledText(label, String(items[i].label || ''), {
      size: CONFIG.FONTS.sizes.body,
      bold: true,
      align: SlidesApp.ParagraphAlignment.LEFT
    });
    
    // パーセンテージ（右上、大きく表示）
    const pct = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
      area.left + area.width - percentWidth - cardPadding, 
      cardTop + cardPadding - layout.pxToPt(2), 
      percentWidth, percentHeight);
    
    setStyledText(pct, `${p}%`, {
      size: 20, // 大きく表示
      bold: true,
      color: settings.primaryColor, // プライマリカラーに統一
      align: SlidesApp.ParagraphAlignment.RIGHT // 右詰め
    });
    
    // 進捗バー（カード下部）
    const barTop = cardTop + cardHeight - cardPadding - barHeight;
    const barWidth = area.width - cardPadding * 2;
    
    // バー背景
    const barBG = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 
      area.left + cardPadding, barTop, barWidth, barHeight);
    barBG.getFill().setSolidFill(CONFIG.COLORS.faint_gray);
    barBG.getBorder().setTransparent();
    
    // 進捗バー
    if (p > 0) {
      const filledBarWidth = Math.max(layout.pxToPt(6), barWidth * (p / 100));
      const barFG = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 
        area.left + cardPadding, barTop, filledBarWidth, barHeight);
      barFG.getFill().setSolidFill(settings.primaryColor); // プライマリカラーに統一
      barFG.getBorder().setTransparent();
    }
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

function createQuoteSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'quoteSlide', data.title || '引用', settings);
  const dy = drawSubheadIfAny(slide, layout, 'quoteSlide', data.subhead);
  
  // 小見出しの高さに応じて座布団の位置を調整
  const baseTop = 120;
  const subheadHeight = data.subhead ? layout.pxToPt(40) : 0; // 小見出しの高さ
  const margin = layout.pxToPt(10); // 小見出しと座布団の間隔
  
  const area = offsetRect(layout.getRect({
    left: 40,
    top: baseTop + subheadHeight + margin,
    width: 880,
    height: 320 - subheadHeight - margin // 高さも調整
  }), 0, dy);
  const bgCard = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, area.left, area.top, area.width, area.height);
  bgCard.getFill().setSolidFill(CONFIG.COLORS.background_white);
  const border = bgCard.getBorder();
  border.getLineFill().setSolidFill(CONFIG.COLORS.card_border);
  border.setWeight(2);
  const padding = layout.pxToPt(40);
  
  // 引用符を削除し、テキストエリアを全幅で使用
  const textLeft = area.left + padding,
    textTop = area.top + padding;
  const textWidth = area.width - (padding * 2),
    textHeight = area.height - (padding * 2);
  const quoteTextHeight = textHeight - layout.pxToPt(30);
  const textShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textLeft, textTop, textWidth, quoteTextHeight);
  setStyledText(textShape, data.text || '', {
    size: 24,
    align: SlidesApp.ParagraphAlignment.CENTER, // 中央揃えに変更
    color: CONFIG.COLORS.text_primary
  });
  try {
    textShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  } catch (e) {}
  const authorTop = textTop + quoteTextHeight;
  const authorShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textLeft, authorTop, textWidth, layout.pxToPt(30));
  setStyledText(authorShape, `— ${data.author || ''}`, {
    size: 16,
    color: CONFIG.COLORS.neutral_gray,
    align: SlidesApp.ParagraphAlignment.END
  });
  try {
    authorShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  } catch (e) {}
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

function createKpiSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'kpiSlide', data.title || '主要指標', settings);
  const dy = drawSubheadIfAny(slide, layout, 'kpiSlide', data.subhead);
  
  // 小見出しの高さに応じてKPIグリッドエリアを動的に調整
  const baseArea = layout.getRect('kpiSlide.gridArea');
  const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
  const area = offsetRect(adjustedArea, 0, dy);
  
  const items = Array.isArray(data.items) ? data.items : [];
  const cols = Math.min(4, Math.max(2, Number(data.columns) || (items.length <= 4 ? items.length : 4)));
  const gap = layout.pxToPt(16);
  const cardW = (area.width - gap * (cols - 1)) / cols,
    cardH = layout.pxToPt(240);
  for (let idx = 0; idx < items.length; idx++) {
    const c = idx % cols,
      r = Math.floor(idx / cols);
    const left = area.left + c * (cardW + gap),
      top = area.top + r * (cardH + gap);
    const card = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top, cardW, cardH);
    card.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    const item = data.items[idx] || {};
    const pad = layout.pxToPt(15);
    const labelShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left + pad, top + layout.pxToPt(25), cardW - pad * 2, layout.pxToPt(35));
    setStyledText(labelShape, item.label || 'KPI', {
      size: 14,
      color: CONFIG.COLORS.neutral_gray
    });
    const valueShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left + pad, top + layout.pxToPt(80), cardW - pad * 2, layout.pxToPt(80));
    setStyledText(valueShape, item.value || '0', {
      size: 32,
      bold: true,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
    try {
      valueShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
    const changeShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left + pad, top + layout.pxToPt(180), cardW - pad * 2, layout.pxToPt(40));
    let changeColor = CONFIG.COLORS.text_primary;
    if (item.status === 'bad') changeColor = '#d93025';
    if (item.status === 'good') changeColor = '#1e8e3e';
    if (item.status === 'neutral') changeColor = CONFIG.COLORS.neutral_gray;
    setStyledText(changeShape, item.change || '', {
      size: 14,
      color: changeColor,
      bold: true,
      align: SlidesApp.ParagraphAlignment.END
    });
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

function createBulletCardsSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'contentSlide', data.title, settings);
  const dy = drawSubheadIfAny(slide, layout, 'contentSlide', data.subhead);
  const area = offsetRect(layout.getRect('contentSlide.body'), 0, dy);
  const items = Array.isArray(data.items) ? data.items.slice(0, 3) : [];
  if (items.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }
  
  const gap = layout.pxToPt(16);
  const cardHeight = (area.height - gap * (items.length - 1)) / items.length;
  
  for (let i = 0; i < items.length; i++) {
    const card = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, area.top + i * (cardHeight + gap), area.width, cardHeight);
    card.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    
    const padding = layout.pxToPt(20);
    const title = String(items[i].title || '');
    const desc = String(items[i].desc || '');
    
    if (title && desc) {
      // 縦積みレイアウト（統一）
      const titleFontSize = 14;
      const titleHeight = layout.pxToPt(titleFontSize + 4); // 14pxフォント + 4px余白
      
      // タイトル部分
        const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
          area.left + padding, 
        area.top + i * (cardHeight + gap) + layout.pxToPt(12), 
        area.width - padding * 2, 
        titleHeight
        );
        setStyledText(titleShape, title, {
        size: titleFontSize,
        bold: true
      });
      
      // 説明文の位置と高さを動的に計算
      const descTop = area.top + i * (cardHeight + gap) + layout.pxToPt(12) + titleHeight + layout.pxToPt(8);
      const descHeight = cardHeight - layout.pxToPt(12) - titleHeight - layout.pxToPt(8);
      
      // 文字数に応じてフォントサイズを自動調整
      let descFontSize = 14;
      if (desc.length > 100) {
        descFontSize = 12; // 100文字超えは12px
      } else if (desc.length > 80) {
        descFontSize = 13; // 80文字超えは13px
      }
      
        const descShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
        area.left + padding, 
        descTop,
        area.width - padding * 2, 
        descHeight
        );
        setStyledText(descShape, desc, {
        size: descFontSize
        });
        
        try {
          descShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        } catch (e) {}
      } else {
      // titleまたはdescのみの場合 - フォントサイズに応じた高さ調整
      const singleFontSize = 14;
      const singleHeight = layout.pxToPt(singleFontSize + 8); // 14pxフォント + 8px余白
      const shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
        area.left + padding, 
        area.top + i * (cardHeight + gap) + (cardHeight - singleHeight) / 2, // 中央配置
        area.width - padding * 2, 
        singleHeight
      );
      setStyledText(shape, title || desc, {
        size: singleFontSize,
        bold: !!title
      });
      try {
        shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      } catch (e) {}
    }
  }
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

function createAgendaSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'processSlide', data.title, settings);
  const dy = drawSubheadIfAny(slide, layout, 'processSlide', data.subhead);

  const area = offsetRect(layout.getRect('processSlide.area'), 0, dy);
  
  // アジェンダ項目の取得（安全装置付き）
  let items = Array.isArray(data.items) ? data.items : [];
  if (items.length === 0) {
    // アジェンダ生成の安全装置
    items = buildAgendaFromSlideData();
    if (items.length === 0) {
      items = ['本日の目的', '進め方', '次のアクション'];
    }
  }

  const n = Math.max(1, items.length);

  const topPadding = layout.pxToPt(30);
  const bottomPadding = layout.pxToPt(10);
  const drawableHeight = area.height - topPadding - bottomPadding;
  const gapY = drawableHeight / Math.max(1, n - 1);
  const cx = area.left + layout.pxToPt(44);
  const top0 = area.top + topPadding;


  for (let i = 0; i < n; i++) {
    const cy = top0 + gapY * i;
    const sz = layout.pxToPt(28);
    
    // 番号ボックス
    const numBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, cx - sz/2, cy - sz/2, sz, sz);
    numBox.getFill().setSolidFill(settings.primaryColor);
    numBox.getBorder().setTransparent();
    
    const num = numBox.getText(); 
    num.setText(String(i + 1));
    applyTextStyle(num, { 
      size: 12, 
      bold: true, 
      color: CONFIG.COLORS.background_white, 
      align: SlidesApp.ParagraphAlignment.CENTER 
    });

    // テキストボックス
    let cleanText = String(items[i] || '');
    cleanText = cleanText.replace(/^\s*\d+[\.\s]*/, '');

    const txt = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, cx + layout.pxToPt(28), cy - layout.pxToPt(16), area.width - layout.pxToPt(70), layout.pxToPt(32));
    setStyledText(txt, cleanText, { size: CONFIG.FONTS.sizes.processStep });
    try { 
      txt.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); 
    } catch(e){}
  }

  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

/**
 * FAQスライドを生成
 * @param {Slide} slide - スライドオブジェクト
 * @param {Object} data - スライドデータ
 * @param {Object} layout - レイアウトマネージャー
 * @param {number} pageNum - ページ番号
 * @param {Object} settings - ユーザー設定
 */
function createFaqSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'contentSlide', data.title || 'よくあるご質問', settings);
  const dy = drawSubheadIfAny(slide, layout, 'contentSlide', data.subhead);
  const area = offsetRect(layout.getRect('contentSlide.body'), 0, dy);
  const items = Array.isArray(data.items) ? data.items.slice(0, 4) : [];
  if (items.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }

  let currentY = area.top;
  const cardGap = layout.pxToPt(12); // カード間の余白
  
  // 項目数に応じた動的カードサイズ計算
  const totalGaps = cardGap * (items.length - 1);
  const availableHeight = area.height - totalGaps;
  const cardHeight = availableHeight / items.length;

  items.forEach((item, index) => {
    // カード背景
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 
      area.left, currentY, area.width, cardHeight);
    card.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    card.getBorder().setWeight(1);

    // 項目数に応じた余白とレイアウト調整
    let cardPadding, qAreaRatio, qAGap;
    
    if (items.length <= 2) {
      // 2項目以下：ゆったりレイアウト
      cardPadding = layout.pxToPt(16);  // 大きめの余白
      qAreaRatio = 0.30;               // Q部分30%
      qAGap = layout.pxToPt(6);        // Q-A間6px
    } else if (items.length === 3) {
      // 3項目：バランス重視
      cardPadding = layout.pxToPt(12);  // 標準余白
      qAreaRatio = 0.35;               // Q部分35%
      qAGap = layout.pxToPt(4);        // Q-A間4px
    } else {
      // 4項目以上：コンパクトレイアウト
      cardPadding = layout.pxToPt(8);   // 小さめの余白
      qAreaRatio = 0.40;               // Q部分40%（質問文が長い場合への対応）
      qAGap = layout.pxToPt(2);        // Q-A間2px
    }
    
    // フォントサイズ
    const baseFontSize = items.length >= 4 ? 12 : 14;
    
    // 固定レイアウト：カード内をQ部分とA部分に分割
    const availableHeight = cardHeight - cardPadding * 2;
    const qAreaHeight = Math.floor(availableHeight * qAreaRatio);
    const aAreaHeight = availableHeight - qAreaHeight - qAGap;
    
    // Q部分のレイアウト（Q.とテキストを統合）
    const qTop = currentY + cardPadding;
    const qText = item.q || '';
    
    // 強調語を解析してからプレフィックスを追加
    const qParsed = parseInlineStyles(qText);
    const qFullText = `Q. ${qParsed.output}`;
    
    const qTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
      area.left + cardPadding, qTop, 
      area.width - cardPadding * 2, qAreaHeight - layout.pxToPt(2));
    qTextShape.getFill().setTransparent();
    qTextShape.getBorder().setTransparent();
    
    // Q.部分だけスタイル適用
    const qTextRange = qTextShape.getText().setText(qFullText);
    applyTextStyle(qTextRange, {
      size: baseFontSize,
      color: CONFIG.COLORS.text_primary,
      align: SlidesApp.ParagraphAlignment.LEFT
    });
    
    // Q.部分（最初の2文字）に特別なスタイルを適用
    try {
      const qPrefixRange = qTextRange.getRange(0, 2); // "Q."の部分
      qPrefixRange.getTextStyle()
        .setBold(true)
        .setForegroundColor(settings.primaryColor);
        
      // 質問文部分を太字に
      if (qFullText.length > 3) {
        const qContentRange = qTextRange.getRange(3, qFullText.length); // "Q. "以降の部分
        qContentRange.getTextStyle().setBold(true);
      }
      
      // 強調語のスタイルを適用（オフセットを調整）
      qParsed.ranges.forEach(r => {
        const adjustedRange = qTextRange.getRange(r.start + 3, r.end + 3); // "Q. "の3文字分
        if (r.bold) adjustedRange.getTextStyle().setBold(true);
        if (r.color) adjustedRange.getTextStyle().setForegroundColor(r.color);
      });
    } catch (e) {}
    
    try {
      qTextShape.setContentAlignment(SlidesApp.ContentAlignment.TOP);
    } catch (e) {}

    // A部分のレイアウト（A.とテキストを統合・インデント風にするため右にずらす）
    const aTop = qTop + qAreaHeight + qAGap;
    const aText = item.a || '';
    
    // 強調語を解析してからプレフィックスを追加
    const aParsed = parseInlineStyles(aText);
    const aFullText = `A. ${aParsed.output}`;
    
    // A部分を右にインデント（約16px程度）
    const aIndent = layout.pxToPt(16);
    
    const aTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
      area.left + cardPadding + aIndent, aTop, 
      area.width - cardPadding * 2 - aIndent, aAreaHeight - layout.pxToPt(2));
    aTextShape.getFill().setTransparent();
    aTextShape.getBorder().setTransparent();
    
    // A.部分だけスタイル適用
    const aTextRange = aTextShape.getText().setText(aFullText);
    applyTextStyle(aTextRange, {
      size: baseFontSize,
      color: CONFIG.COLORS.text_primary,
      align: SlidesApp.ParagraphAlignment.LEFT
    });
    
    // A.部分（最初の2文字）に特別なスタイルを適用
    try {
      const aPrefixRange = aTextRange.getRange(0, 2); // "A."の部分
      aPrefixRange.getTextStyle()
        .setBold(true)
        .setForegroundColor(generateTintedGray(settings.primaryColor, 15, 70));
        
      // 強調語のスタイルを適用（オフセットを調整）
      aParsed.ranges.forEach(r => {
        const adjustedRange = aTextRange.getRange(r.start + 3, r.end + 3); // "A. "の3文字分
        if (r.bold) adjustedRange.getTextStyle().setBold(true);
        if (r.color) adjustedRange.getTextStyle().setForegroundColor(r.color);
      });
    } catch (e) {}
    
    try {
      aTextShape.setContentAlignment(SlidesApp.ContentAlignment.TOP);
    } catch (e) {}

    currentY += cardHeight + cardGap;
  });

  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

/**
 * FAQの各項目を描画（Q/A記号を前提としたレイアウトに修正）
 * @param {Slide} slide - スライドオブジェクト
 * @param {Array} items - FAQのQ&Aオブジェクト配列
 * @param {Object} layout - レイアウトマネージャー
 * @param {Object} listArea - 描画エリア
 * @param {Object} settings - ユーザー設定
 */
function drawFaqItems(slide, items, layout, listArea, settings) {
  if (!items || !items.length) return;

  const px = v => layout.pxToPt(v);
  const GAP_ITEM = px(16); // カード間の垂直マージン
  const PADDING = px(20); // カード内部の余白

  // 各カードの高さを均等に分配
  const totalCardHeight = listArea.height - (GAP_ITEM * (items.length - 1));
  const cardHeight = totalCardHeight / items.length;
  
  let currentY = listArea.top;

  items.forEach((qa) => {
    // カード背景 (bulletCardsと同様のスタイル)
    const card = slide.insertShape(
      SlidesApp.ShapeType.RECTANGLE,
      listArea.left, currentY,
      listArea.width, cardHeight
    );
    card.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);

    const q = qa.q || '';
    const a = qa.a || '';

    const qIconWidth = px(30);
    const qTextLeft = listArea.left + PADDING + qIconWidth;
    const qTextWidth = listArea.width - PADDING * 2 - qIconWidth;
    
    // Qアイコン
    const qIcon = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX,
      listArea.left + PADDING, currentY + PADDING, qIconWidth, px(24));
    setStyledText(qIcon, 'Q.', { size: 18, bold: true, color: settings.primaryColor });

    // Qテキスト
    const qBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX,
      qTextLeft, currentY + PADDING, qTextWidth, px(40));
    setStyledText(qBox, q, { size: 14, bold: true, color: CONFIG.COLORS.text_primary });
    
    const aTop = currentY + PADDING + px(35); // Qの下に配置
    const aHeight = cardHeight - (PADDING * 2) - px(35); // 残りの高さを回答エリアに

    // 回答の文字数に応じてフォントサイズを動的に変更
    let answerFontSize = 14;
    if (a.length > 100) {
      answerFontSize = 11;
    } else if (a.length > 60) {
      answerFontSize = 12.5;
    }

    // Aアイコン
    const aIcon = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX,
      listArea.left + PADDING, aTop, qIconWidth, aHeight);
    const tintedGrayColor = generateTintedGray(settings.primaryColor, 15, 70);
    setStyledText(aIcon, 'A.', { size: 18, bold: true, color: tintedGrayColor });

    // Aテキスト
    const aBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX,
      qTextLeft, aTop, qTextWidth, aHeight);
    setStyledText(aBox, a, { size: answerFontSize, color: CONFIG.COLORS.text_primary });

    // テキストボックスの縦揃え設定
    try {
      [qIcon, qBox, aIcon, aBox].forEach(s => {
        s.setContentAlignment(SlidesApp.ContentAlignment.TOP);
        // テキストのはみ出しを最終的に防ぐための自動調整
        s.setAutofit(SlidesApp.AutofitType.SHRINK_ON_OVERFLOW);
      });
    } catch(e){}
    
    currentY += cardHeight + GAP_ITEM;
  });
}

