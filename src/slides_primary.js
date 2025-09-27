// --- Slide Generators (Primary) ---
var slideGenerators = this.slideGenerators || (this.slideGenerators = {});

function createTitleSlide(slide, data, layout) {
  const colors = CONFIG.COLORS;
  slide.getBackground().setSolidFill(colors.title_bg || '#000000');

  const titleRect = layout.getRect('titleSlide.title');
  const titleShape = slide.insertShape(
    SlidesApp.ShapeType.TEXT_BOX,
    titleRect.left,
    titleRect.top,
    titleRect.width,
    titleRect.height
  );
  titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  setStyledText(titleShape, String(data.title || ''), {
    size: CONFIG.FONTS.sizes.title,
    bold: true,
    color: colors.title_text || '#FFFFFF',
    align: SlidesApp.ParagraphAlignment.CENTER
  });

  if (data.subtitle) {
    const subtitleRect = layout.getRect('titleSlide.subtitle');
    const subtitleShape = slide.insertShape(
      SlidesApp.ShapeType.TEXT_BOX,
      subtitleRect.left,
      subtitleRect.top,
      subtitleRect.width,
      subtitleRect.height
    );
    subtitleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    setStyledText(subtitleShape, String(data.subtitle || ''), {
      size: CONFIG.FONTS.sizes.titleSubtitle,
      color: data.subtitleColor || colors.title_subtitle || colors.neutral_gray,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
  }

  // 右下の折り返しアクセント
  const foldSizePt = layout.pxToPt(CONFIG.POS_PX.titleSlide.foldSize || 150);
  const highlightSizePt = layout.pxToPt(CONFIG.POS_PX.titleSlide.foldHighlightSize || 90);
  const foldLeft = layout.pageW_pt - foldSizePt;
  const foldTop = layout.pageH_pt - foldSizePt;
  const foldShape = slide.insertShape(
    SlidesApp.ShapeType.RIGHT_TRIANGLE,
    foldLeft,
    foldTop,
    foldSizePt,
    foldSizePt
  );
  foldShape.setRotation(270);
  foldShape.getFill().setSolidFill(colors.title_fold_shadow || '#BDBDBD');
  foldShape.getBorder().setTransparent();

  const highlightLeft = layout.pageW_pt - highlightSizePt;
  const highlightTop = layout.pageH_pt - highlightSizePt;
  const highlightShape = slide.insertShape(
    SlidesApp.ShapeType.RIGHT_TRIANGLE,
    highlightLeft,
    highlightTop,
    highlightSizePt,
    highlightSizePt
  );
  highlightShape.setRotation(270);
  highlightShape.getFill().setSolidFill(colors.title_fold_highlight || '#F5F5F5');
  highlightShape.getBorder().setTransparent();
}

function createSectionSlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_gray);

  // 透かし番号：sectionNo > タイトル先頭の数字 > 自動連番
  __SECTION_COUNTER++;
  const parsedNum = (() => {
    if (Number.isFinite(data.sectionNo)) return Number(data.sectionNo);
    const m = String(data.title || '').match(/^\s*(\d+)[\.．]/);
    return m ? Number(m[1]) : __SECTION_COUNTER;
  })();
  const num = String(parsedNum).padStart(2, '0');

  const ghostRect = layout.getRect('sectionSlide.ghostNum');
  const ghost = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, ghostRect.left, ghostRect.top, ghostRect.width, ghostRect.height);
  ghost.getText().setText(num);
  applyTextStyle(ghost.getText(), { size: CONFIG.FONTS.sizes.ghostNum, color: CONFIG.COLORS.ghost_gray, bold: true });
  try { ghost.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch(e) {}

  const titleRect = layout.getRect('sectionSlide.title');
  const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, titleRect.left, titleRect.top, titleRect.width, titleRect.height);
  titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  setStyledText(titleShape, data.title, { size: CONFIG.FONTS.sizes.sectionTitle, bold: true, align: SlidesApp.ParagraphAlignment.CENTER });

  addCucFooter(slide, layout, pageNum);
}

// content（1/2カラム + 小見出し + 画像）
function createContentSlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
  drawStandardTitleHeader(slide, layout, 'contentSlide', data.title);
  const dy = 0; // アジェンダパターンでは小見出しを使用しない

  // アジェンダ安全装置
  const isAgenda = isAgendaTitle(data.title || '');
  let points = Array.isArray(data.points) ? data.points.slice(0) : [];
  if (isAgenda && (!points || points.length === 0)) {
    points = buildAgendaFromSlideData();
    if (points.length === 0) points = ['本日の目的', '進め方', '次のアクション'];
  }

  const hasImages = Array.isArray(data.images) && data.images.length > 0;
  const isTwo = !!(data.twoColumn || data.columns);

  if ((isTwo && (data.columns || points)) || (!isTwo && points && points.length > 0)) {
    if (isTwo) {
      let L = [], R = [];
      if (Array.isArray(data.columns) && data.columns.length === 2) {
        L = data.columns[0] || []; R = data.columns[1] || [];
      } else {
        const mid = Math.ceil(points.length / 2);
        L = points.slice(0, mid); R = points.slice(mid);
      }
      const leftRect = offsetRect(layout.getRect('contentSlide.twoColLeft'), 0, dy);
      const rightRect = offsetRect(layout.getRect('contentSlide.twoColRight'), 0, dy);
      const leftShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, leftRect.left, leftRect.top, leftRect.width, leftRect.height);
      const rightShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rightRect.left, rightRect.top, rightRect.width, rightRect.height);
      setBulletsWithInlineStyles(leftShape, L);
      setBulletsWithInlineStyles(rightShape, R);
    } else {
      const bodyRect = offsetRect(layout.getRect('contentSlide.body'), 0, dy);
      if (isAgenda) {
        drawNumberedItems(slide, layout, bodyRect, points);
      } else {
        const bodyShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, bodyRect.left, bodyRect.top, bodyRect.width, bodyRect.height);
        setBulletsWithInlineStyles(bodyShape, points);
      }
    }
  }

  // 画像（任意）
  if (hasImages) {
    const area = offsetRect(layout.getRect('contentSlide.body'), 0, dy);
    renderImagesInArea(slide, layout, area, normalizeImages(data.images));
  }

  drawBottomBarAndFooter(slide, layout, pageNum);
}

// compare（左右ボックス：ヘッダー色＋白文字）＋インライン装飾対応
function createCompareSlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
  drawStandardTitleHeader(slide, layout, 'compareSlide', data.title);
  const dy = drawSubheadIfAny(slide, layout, 'compareSlide', data.subhead);

  const leftBox = offsetRect(layout.getRect('compareSlide.leftBox'), 0, dy);
  const rightBox = offsetRect(layout.getRect('compareSlide.rightBox'), 0, dy);
  drawCompareBox(slide, leftBox, data.leftTitle || '選択肢A', data.leftItems || []);
  drawCompareBox(slide, rightBox, data.rightTitle || '選択肢B', data.rightItems || []);

  drawBottomBarAndFooter(slide, layout, pageNum);
}
function drawCompareBox(slide, rect, title, items) {
  const box = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, rect.left, rect.top, rect.width, rect.height);
  box.getFill().setSolidFill(CONFIG.COLORS.lane_title_bg);
  box.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.lane_border);
  box.getBorder().setWeight(1);

  const th = 0.75 * 40; // 約30px相当
  const titleBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, rect.left, rect.top, rect.width, th);
  titleBar.getFill().setSolidFill(CONFIG.COLORS.primary_color);
  titleBar.getBorder().setTransparent();
  setStyledText(titleBar, title, { size: CONFIG.FONTS.sizes.laneTitle, bold: true, color: CONFIG.COLORS.text_on_primary, align: SlidesApp.ParagraphAlignment.CENTER });

  const pad = 0.75 * 12;
  const textRect = { left: rect.left + pad, top: rect.top + th + pad, width: rect.width - pad * 2, height: rect.height - th - pad * 2 };
  const body = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textRect.left, textRect.top, textRect.width, textRect.height);
  setBulletsWithInlineStyles(body, items);
}

function createBigFactSlide(slide, data, layout) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);

  const mainRect = layout.getRect('bigFactSlide.mainValue');
  const mainShape = slide.insertShape(
    SlidesApp.ShapeType.TEXT_BOX,
    mainRect.left,
    mainRect.top,
    mainRect.width,
    mainRect.height
  );

  const mainText = data.value || data.fact || data.title || '2x';
  const mainRange = mainShape.getText();
  mainRange.setText(mainText);
  applyTextStyle(mainRange, {
    size: CONFIG.FONTS.sizes.bigFactMain,
    bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    color: CONFIG.COLORS.text_primary
  });
  try {
    mainRange.getTextStyle().setFontFamily('Inter');
  } catch (e) {
    // フォント未対応環境では既定フォントを使用
  }
  try { mainShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) {}

  const captionRect = layout.getRect('bigFactSlide.caption');
  const gapPt = layout.pxToPt(100);
  const captionTop = Math.min(
    mainRect.top + mainRect.height + gapPt,
    layout.pageH_pt - captionRect.height - layout.pxToPt(40)
  );
  const captionShape = slide.insertShape(
    SlidesApp.ShapeType.TEXT_BOX,
    captionRect.left,
    captionTop,
    captionRect.width,
    captionRect.height
  );

  const captionText = data.caption || data.label || data.subhead || 'Encoding Faster';
  const captionRange = captionShape.getText();
  captionRange.setText(captionText);
  applyTextStyle(captionRange, {
    size: CONFIG.FONTS.sizes.bigFactCaption,
    color: CONFIG.COLORS.bigFact_caption,
    align: SlidesApp.ParagraphAlignment.CENTER
  });
  try {
    const st = captionRange.getTextStyle();
    st.setFontFamily('Inter');
    st.setBold(false);
  } catch (e) {
    // フォント未対応環境では既定フォントを使用
  }
  try { captionShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) {}
}

function createFullBreedSlide(slide, data, layout) {
  const bgUrl = data.backgroundImage || CONFIG.IMAGES.fullBreedBackground;
  try {
    if (bgUrl) {
      slide.getBackground().setPictureFill(bgUrl);
    } else {
      slide.getBackground().setSolidFill('#1A223A');
    }
  } catch (e) {
    slide.getBackground().setSolidFill('#1A223A');
  }

  const overlayOpacity = typeof data.overlayOpacity === 'number'
    ? data.overlayOpacity
    : CONFIG.POS_PX.fullBreedSlide.overlayOpacity;

  if (typeof overlayOpacity === 'number' && overlayOpacity >= 0 && overlayOpacity < 1) {
    const overlay = slide.insertShape(
      SlidesApp.ShapeType.RECTANGLE,
      0,
      0,
      layout.pageW_pt,
      layout.pageH_pt
    );
    overlay.getFill().setSolidFill(CONFIG.COLORS.fullBreed_overlay || '#000000');
    try {
      const transparency = Math.min(1, Math.max(0, overlayOpacity));
      overlay.getFill().setTransparency(transparency);
    } catch (e) {}
    overlay.getBorder().setTransparent();
  }

  const textArea = layout.getRect('fullBreedSlide.textArea');
  const items = Array.isArray(data.items) ? data.items : [];
  const itemGapPt = layout.pxToPt(CONFIG.POS_PX.fullBreedSlide.itemGap || 180);
  const lineHeightPt = layout.pxToPt(140);
  const fontSize = data.itemSize || CONFIG.FONTS.sizes.fullBreedItem;
  const textColor = data.textColor || CONFIG.COLORS.fullBreed_text || '#FFFFFF';

  let currentTop = textArea.top;

  items.forEach((raw, index) => {
    if (index > 0) currentTop += itemGapPt;
    const entry = String(raw || '');
    const box = slide.insertShape(
      SlidesApp.ShapeType.TEXT_BOX,
      textArea.left,
      currentTop,
      textArea.width,
      lineHeightPt
    );
    box.getFill().setTransparent();
    box.getBorder().setTransparent();
    const tr = box.getText();
    tr.setText(entry);
    applyTextStyle(tr, { size: fontSize, bold: true, color: textColor });
    try {
      const style = tr.getTextStyle();
      style.setFontFamily('Inter');
    } catch (e) {}
  });

}

// process（角枠1px＋一桁数字）
function createProcessSlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
  drawStandardTitleHeader(slide, layout, 'processSlide', data.title);
  const dy = drawSubheadIfAny(slide, layout, 'processSlide', data.subhead);

  const area = offsetRect(layout.getRect('processSlide.area'), 0, dy);
  const steps = Array.isArray(data.steps) ? data.steps : [];
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
    numBox.getFill().setSolidFill(CONFIG.COLORS.primary_color);
    numBox.getBorder().setTransparent();
    const num = numBox.getText(); num.setText(String(i + 1));
    applyTextStyle(num, { size: 12, bold: true, color: CONFIG.COLORS.text_on_primary, align: SlidesApp.ParagraphAlignment.CENTER });

    // 元のプロセステキストから先頭の数字を除去
    let cleanText = String(steps[i] || '');
    cleanText = cleanText.replace(/^\s*\d+[\.\s]*/, '');

    const txt = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, cx + layout.pxToPt(28), cy - layout.pxToPt(16), area.width - layout.pxToPt(70), layout.pxToPt(32));
    setStyledText(txt, cleanText, { size: CONFIG.FONTS.sizes.processStep });
    try { txt.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch(e){}
  }

  drawBottomBarAndFooter(slide, layout, pageNum);
}

// timeline（左右余白広め）
function createTimelineSlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
  drawStandardTitleHeader(slide, layout, 'timelineSlide', data.title);
  const dy = drawSubheadIfAny(slide, layout, 'timelineSlide', data.subhead);

  const area = offsetRect(layout.getRect('timelineSlide.area'), 0, dy);
  const milestones = Array.isArray(data.milestones) ? data.milestones : [];
  if (milestones.length === 0) { drawBottomBarAndFooter(slide, layout, pageNum); return; }

  const inner = layout.pxToPt(80);
  const baseY = area.top + area.height * 0.50;
  const leftX = area.left + inner;
  const rightX = area.left + area.width - inner;

  const line = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, leftX, baseY - layout.pxToPt(1), rightX - leftX, layout.pxToPt(2));
  line.getFill().setSolidFill(CONFIG.COLORS.neutral_gray);
  line.getBorder().setTransparent();

  const dotR = layout.pxToPt(8);
  const gap = (rightX - leftX) / Math.max(1, (milestones.length - 1));

  milestones.forEach((m, i) => {
    const x = leftX + gap * i - dotR / 2;
    const dot = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x, baseY - dotR / 2, dotR, dotR);
    
    // 時系列順で徐々に濃くなる色計算
    const progress = milestones.length > 1 ? i / (milestones.length - 1) : 0;
    const brightness = 1.5 - (progress * 0.8); // 1.5 → 0.7 の範囲で徐々に濃くなる
    dot.getFill().setSolidFill(adjustColorBrightness(CONFIG.COLORS.primary_color, brightness));
    dot.getBorder().setTransparent();

    // ラベルテキスト（図形の上部、重ならない位置）
    const labelShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x - layout.pxToPt(50) + dotR/2, baseY - layout.pxToPt(70), layout.pxToPt(100), layout.pxToPt(18));
    labelShape.getFill().setTransparent();
    labelShape.getBorder().setTransparent();
    setStyledText(labelShape, String(m.label || ''), { size: CONFIG.FONTS.sizes.small, bold: true, align: SlidesApp.ParagraphAlignment.CENTER });

    // 日付テキスト（図形の下部、より小さいフォント）
    const dateShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x - layout.pxToPt(50) + dotR/2, baseY + layout.pxToPt(15), layout.pxToPt(100), layout.pxToPt(18));
    dateShape.getFill().setTransparent();
    dateShape.getBorder().setTransparent();
    setStyledText(dateShape, String(m.date || ''), { size: CONFIG.FONTS.sizes.small, color: CONFIG.COLORS.neutral_gray, align: SlidesApp.ParagraphAlignment.CENTER });

  });

  drawBottomBarAndFooter(slide, layout, pageNum);
}

// diagram（Mermaid風・レーン＋カード＋自動矢印）
function createDiagramSlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
  drawStandardTitleHeader(slide, layout, 'diagramSlide', data.title);
  const dy = drawSubheadIfAny(slide, layout, 'diagramSlide', data.subhead);

  const lanes = Array.isArray(data.lanes) ? data.lanes : [];
  const area0 = layout.getRect('diagramSlide.lanesArea');
  const area = offsetRect(area0, 0, dy);

  const px = (p)=> layout.pxToPt(p);
  const laneGap = px(CONFIG.DIAGRAM.laneGap_px);
  const lanePad = px(CONFIG.DIAGRAM.lanePad_px);
  const laneTitleH = px(CONFIG.DIAGRAM.laneTitle_h_px);
  const cardGap = px(CONFIG.DIAGRAM.cardGap_px);
  const cardMinH = px(CONFIG.DIAGRAM.cardMin_h_px);
  const cardMaxH = px(CONFIG.DIAGRAM.cardMax_h_px);
  const arrowH = px(CONFIG.DIAGRAM.arrow_h_px);
  const arrowGap = px(CONFIG.DIAGRAM.arrowGap_px);

  const n = Math.max(1, lanes.length);
  const laneW = (area.width - laneGap * (n - 1)) / n;

  const cardBoxes = [];

  for (let j = 0; j < n; j++) {
    const lane = lanes[j] || { title: '', items: [] };
    const left = area.left + j * (laneW + laneGap);
    const top = area.top;

    const lt = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top, laneW, laneTitleH);
    lt.getFill().setSolidFill(CONFIG.COLORS.lane_title_bg);
    lt.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.lane_border);
    lt.getBorder().setWeight(1);
    lt.getText().setText(lane.title || '');
    applyTextStyle(lt.getText(), { size: CONFIG.FONTS.sizes.laneTitle, bold: true, align: SlidesApp.ParagraphAlignment.CENTER });

    const items = Array.isArray(lane.items) ? lane.items : [];
    const availH = area.height - laneTitleH - lanePad * 2;
    const rows = Math.max(1, items.length);
    const idealH = (availH - cardGap * (rows - 1)) / rows;
    const cardH = Math.max(cardMinH, Math.min(cardMaxH, idealH));
    const totalH = cardH * rows + cardGap * (rows - 1);
    const firstTop = top + laneTitleH + lanePad + Math.max(0, (availH - totalH) / 2);

    cardBoxes[j] = [];
    for (let i = 0; i < rows; i++) {
      const cardTop = firstTop + i * (cardH + cardGap);
      const cardLeft = left + lanePad;
      const cardWidth = laneW - lanePad * 2;

      const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, cardLeft, cardTop, cardWidth, cardH);
      card.getFill().setSolidFill(CONFIG.COLORS.card_bg);
      card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
      card.getBorder().setWeight(1);
      setStyledText(card, items[i] || '', { size: CONFIG.FONTS.sizes.body });

      try { card.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch(e){}
      cardBoxes[j][i] = { left: cardLeft, top: cardTop, width: cardWidth, height: cardH };
    }
  }

  // 同行カード間を矢印で接続
  const maxRows = Math.max(...cardBoxes.map(a => a.length));
  for (let j = 0; j < n - 1; j++) {
    const L = cardBoxes[j], R = cardBoxes[j + 1];
    for (let i = 0; i < maxRows; i++) {
      const a = L[i], b = R[i];
      if (a && b) drawArrowBetweenRects(slide, a, b, arrowH, arrowGap);
    }
  }

  drawBottomBarAndFooter(slide, layout, pageNum);
}

// cards（シンプルカード）
function createTableSlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
  drawStandardTitleHeader(slide, layout, 'tableSlide', data.title);
  const dy = drawSubheadIfAny(slide, layout, 'tableSlide', data.subhead);

  const area = offsetRect(layout.getRect('tableSlide.area'), 0, dy);
  const headers = Array.isArray(data.headers) ? data.headers : [];
  const rows = Array.isArray(data.rows) ? data.rows : [];

  try {
    if (headers.length > 0) {
      const table = slide.insertTable(rows.length + 1, headers.length);
      table.setLeft(area.left).setTop(area.top).setWidth(area.width);
      
      // ヘッダー行の背景色設定とテキスト設定
      for (let c = 0; c < headers.length; c++) {
        const cell = table.getCell(0, c);
        cell.getFill().setSolidFill(CONFIG.COLORS.table_header_bg);
        setStyledText(cell, String(headers[c] || ''), { bold: true, align: SlidesApp.ParagraphAlignment.CENTER });
      }
      
      // データ行の設定
      for (let r = 0; r < rows.length; r++) {
        const row = rows[r] || [];
        for (let c = 0; c < headers.length; c++) {
          const cell = table.getCell(r + 1, c);
          setStyledText(cell, String(row[c] || ''), { align: SlidesApp.ParagraphAlignment.CENTER });
        }
      }
    } else {
      throw new Error('headers is empty');
    }
  } catch (e) {
    // フォールバック：矩形シェイプで表を作成
    const cols = Math.max(1, headers.length || 3);
    const rcount = rows.length + 1;
    const gap = layout.pxToPt(1);
    const cellW = (area.width - gap * (cols - 1)) / cols;
    const cellH = (area.height - gap * (rcount - 1)) / rcount;
    
    const drawCell = (r, c, text, isHeader) => {
      const left = area.left + c * (cellW + gap);
      const top  = area.top  + r * (cellH + gap);
      const cell = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top, cellW, cellH);
      cell.getFill().setSolidFill(isHeader ? CONFIG.COLORS.table_header_bg : CONFIG.COLORS.background_white);
      cell.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
      cell.getBorder().setWeight(1);
      setStyledText(cell, String(text || ''), { bold: !!isHeader, align: SlidesApp.ParagraphAlignment.CENTER });
      try { cell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch(e){}
    };
    
    // ヘッダー行の描画
    (headers.length ? headers : ['項目','値1','値2']).forEach((h, c) => drawCell(0, c, h, true));
    
    // データ行の描画
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r] || [];
      for (let c = 0; c < (headers.length || 3); c++) drawCell(r + 1, c, row[c], false);
    }
  }
  drawBottomBarAndFooter(slide, layout, pageNum);
}

// progress（進捗バー）
function createProgressSlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
  drawStandardTitleHeader(slide, layout, 'progressSlide', data.title);
  const dy = drawSubheadIfAny(slide, layout, 'progressSlide', data.subhead);

  const area = offsetRect(layout.getRect('progressSlide.area'), 0, dy);
  const items = Array.isArray(data.items) ? data.items : [];
  const n = Math.max(1, items.length);
  const rowH = area.height / n;

  for (let i = 0; i < n; i++) {
    const rowCenterY = area.top + i * rowH + rowH / 2;
    const textHeight = layout.pxToPt(18);
    const barHeight = layout.pxToPt(14);
    
    // 全要素を行の中央に配置するための基準Y座標を計算
    const textY = rowCenterY - textHeight / 2;
    const barY = rowCenterY - barHeight / 2;

    const label = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, textY, layout.pxToPt(150), textHeight);
    setStyledText(label, String(items[i].label || ''), { size: CONFIG.FONTS.sizes.body });
    try { label.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch(e){}

    const barLeft = area.left + layout.pxToPt(160);
    const barW    = area.width - layout.pxToPt(300);
    const barBG = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, barLeft, barY, barW, barHeight);
    barBG.getFill().setSolidFill(CONFIG.COLORS.faint_gray); barBG.getBorder().setTransparent();

    const p = Math.max(0, Math.min(100, Number(items[i].percent || 0)));
    if (p > 0) {
      const barFG = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, barLeft, barY, barW * (p/100), barHeight);
      barFG.getFill().setSolidFill(CONFIG.COLORS.primary_color); barFG.getBorder().setTransparent();
    }

    const pct = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, barLeft + barW + layout.pxToPt(10), textY, layout.pxToPt(80), textHeight);
    pct.getText().setText(String(p) + '%');
    applyTextStyle(pct.getText(), { size: CONFIG.FONTS.sizes.small, color: CONFIG.COLORS.neutral_gray });
    try { pct.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch(e){}
  }

  drawBottomBarAndFooter(slide, layout, pageNum);
}

// quote（引用）
function createQuoteSlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
  drawStandardTitleHeader(slide, layout, 'quoteSlide', data.title || '引用');
  const dy = drawSubheadIfAny(slide, layout, 'quoteSlide', data.subhead);

  const markRect = offsetRect(layout.getRect('quoteSlide.quoteMark'), 0, dy);
  const markShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, markRect.left, markRect.top, markRect.width, markRect.height);
  markShape.getText().setText('“');
  applyTextStyle(markShape.getText(), { size: 120, color: CONFIG.COLORS.ghost_gray, bold: true });

  const textRect = offsetRect(layout.getRect('quoteSlide.quoteText'), 0, dy);
  const textShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textRect.left, textRect.top, textRect.width, textRect.height);
  setStyledText(textShape, data.text || '', { size: 24, align: SlidesApp.ParagraphAlignment.START });
  try { textShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch(e){}

  const authorRect = offsetRect(layout.getRect('quoteSlide.author'), 0, dy);
  const authorShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, authorRect.left, authorRect.top, authorRect.width, authorRect.height);
  setStyledText(authorShape, `— ${data.author || ''}`, { size: 16, color: CONFIG.COLORS.neutral_gray, align: SlidesApp.ParagraphAlignment.END });

  drawBottomBarAndFooter(slide, layout, pageNum);
}

// kpi（KPIカード）
function createKpiSlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
  drawStandardTitleHeader(slide, layout, 'kpiSlide', data.title || '主要指標');
  const dy = drawSubheadIfAny(slide, layout, 'kpiSlide', data.subhead);

  const area = offsetRect(layout.getRect('kpiSlide.gridArea'), 0, dy);
  const items = Array.isArray(data.items) ? data.items : [];
  const cols = Math.min(4, Math.max(2, Number(data.columns) || (items.length <= 4 ? items.length : 4)));
  const gap = layout.pxToPt(16);
  const cardW = (area.width - gap * (cols - 1)) / cols;
  const cardH = layout.pxToPt(240);  // 200px → 240px に拡大

  for (let idx = 0; idx < items.length; idx++) {
    const c = idx % cols;
    const r = Math.floor(idx / cols);
    const left = area.left + c * (cardW + gap);
    const top  = area.top  + r * (cardH + gap);

    const card = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top, cardW, cardH);
    card.getFill().setSolidFill(CONFIG.COLORS.card_bg);
    card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    card.getBorder().setWeight(1);
    
    const item = data.items[idx] || {};
    const pad = layout.pxToPt(15);
    const contentWidth = cardW - (pad * 2);
    
    // 3つの要素を均等配置
    const labelShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left + pad, top + layout.pxToPt(25), contentWidth, layout.pxToPt(35));
    setStyledText(labelShape, item.label || 'KPI', { size: 14, color: CONFIG.COLORS.neutral_gray });

    const valueShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left + pad, top + layout.pxToPt(80), contentWidth, layout.pxToPt(80));
    setStyledText(valueShape, item.value || '0', { size: 32, bold: true, align: SlidesApp.ParagraphAlignment.CENTER });
    try { valueShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch(e){}

    const changeShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left + pad, top + layout.pxToPt(180), contentWidth, layout.pxToPt(40));
    let changeColor = CONFIG.COLORS.text_primary;
    if (item.status === 'bad') changeColor = adjustColorBrightness(CONFIG.COLORS.primary_color, 0.7);
    if (item.status === 'good') changeColor = adjustColorBrightness(CONFIG.COLORS.primary_color, 1.3);
    setStyledText(changeShape, item.change || '', { size: 14, color: changeColor, bold: true, align: SlidesApp.ParagraphAlignment.END });
  }

  drawBottomBarAndFooter(slide, layout, pageNum);
}

// closing（結び）
function createClosingSlide(slide, data, layout) {
slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
try {
  const image = slide.insertImage(CONFIG.LOGOS.closing);
  const imgW_pt = layout.pxToPt(450) * layout.scaleX;
  const aspect = image.getHeight() / image.getWidth();
  image.setWidth(imgW_pt).setHeight(imgW_pt * aspect);
  image.setLeft((layout.pageW_pt - imgW_pt) / 2).setTop((layout.pageH_pt - (imgW_pt * aspect)) / 2);
} catch (e) {
  // 画像挿入に失敗した場合はスキップして他の要素を描画
}
}

// bulletCards（箇条書きカード）
function createStatsCompareSlide(slide, data, layout, pageNum) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);
  drawStandardTitleHeader(slide, layout, 'statsCompareSlide', data.title);
  const dy = drawSubheadIfAny(slide, layout, 'statsCompareSlide', data.subhead);

  const leftArea = offsetRect(layout.getRect('statsCompareSlide.leftArea'), 0, dy);
  const rightArea = offsetRect(layout.getRect('statsCompareSlide.rightArea'), 0, dy);

  // 左右のタイトルヘッダー
  const headerHeight = layout.pxToPt(40);
  const leftHeader = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, leftArea.left, leftArea.top, leftArea.width, headerHeight);
  leftHeader.getFill().setSolidFill(CONFIG.COLORS.primary_color);
  leftHeader.getBorder().setTransparent();
  setStyledText(leftHeader, data.leftTitle || '現在', { size: CONFIG.FONTS.sizes.laneTitle, bold: true, color: CONFIG.COLORS.text_on_primary, align: SlidesApp.ParagraphAlignment.CENTER });

  const rightHeader = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, rightArea.left, rightArea.top, rightArea.width, headerHeight);
  rightHeader.getFill().setSolidFill(CONFIG.COLORS.primary_color);
  rightHeader.getBorder().setTransparent();
  setStyledText(rightHeader, data.rightTitle || '目標', { size: CONFIG.FONTS.sizes.laneTitle, bold: true, color: CONFIG.COLORS.text_on_primary, align: SlidesApp.ParagraphAlignment.CENTER });

  // 統計データエリア
  const stats = Array.isArray(data.stats) ? data.stats : [];
  const statHeight = (leftArea.height - headerHeight) / Math.max(1, stats.length);

  stats.forEach((stat, i) => {
    const y = leftArea.top + headerHeight + i * statHeight;
    const padding = layout.pxToPt(15);

    // ラベル（左右共通、左側に表示）
    const labelShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, leftArea.left + padding, y + layout.pxToPt(5), leftArea.width - padding * 2, layout.pxToPt(18));
    labelShape.getFill().setTransparent();
    labelShape.getBorder().setTransparent();
    setStyledText(labelShape, String(stat.label || ''), { size: 12, bold: true, color: CONFIG.COLORS.neutral_gray });

    // 左側の値
    const leftValueShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, leftArea.left + padding, y + layout.pxToPt(25), leftArea.width - padding * 2, layout.pxToPt(30));
    leftValueShape.getFill().setTransparent();
    leftValueShape.getBorder().setTransparent();
    setStyledText(leftValueShape, String(stat.leftValue || ''), { size: 20, bold: true, align: SlidesApp.ParagraphAlignment.CENTER });

    // 右側の値
    const rightValueShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rightArea.left + padding, y + layout.pxToPt(25), rightArea.width - padding * 2, layout.pxToPt(30));
    rightValueShape.getFill().setTransparent();
    rightValueShape.getBorder().setTransparent();
    
    let trendColor = CONFIG.COLORS.text_primary;
    let trendSymbol = '';
    if (stat.trend === 'up') {
      trendColor = adjustColorBrightness(CONFIG.COLORS.primary_color, 1.3);
      trendSymbol = ' ↗';
    } else if (stat.trend === 'down') {
      trendColor = adjustColorBrightness(CONFIG.COLORS.primary_color, 0.7);
      trendSymbol = ' ↘';
    }
    
    setStyledText(rightValueShape, String(stat.rightValue || '') + trendSymbol, { size: 20, bold: true, color: trendColor, align: SlidesApp.ParagraphAlignment.CENTER });
  });

  drawBottomBarAndFooter(slide, layout, pageNum);
}


Object.assign(slideGenerators, {
  title: createTitleSlide,
  section: createSectionSlide,
  content: createContentSlide,
  compare: createCompareSlide,
  process: createProcessSlide,
  timeline: createTimelineSlide,
  diagram: createDiagramSlide,
  table: createTableSlide,
  progress: createProgressSlide,
  quote: createQuoteSlide,
  kpi: createKpiSlide,
  closing: createClosingSlide,
  statsCompare: createStatsCompareSlide,
  bigFact: createBigFactSlide,
  fullBreed: createFullBreedSlide,
});
