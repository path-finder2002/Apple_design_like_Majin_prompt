// --- 6. スライド生成ディスパッチャ ---
const slideGenerators = {
  title: createTitleSlide,
  section: createSectionSlide,
  content: createContentSlide,
  statsCompare: createStatsCompareSlide,
  compare: createCompareSlide,
  process: createProcessSlide,
  timeline: createTimelineSlide,
  diagram: createDiagramSlide,
  cards: createCardsSlide,
  headerCards: createHeaderCardsSlide,
  table: createTableSlide,
  progress: createProgressSlide,
  quote: createQuoteSlide,
  kpi: createKpiSlide,
  closing: createClosingSlide,
  bulletCards: createBulletCardsSlide,
  faq: createFaqSlide,
};

// --- 7. スライド生成関数群 ---
function createTitleSlide(slide, data, layout) {
  slide.getBackground().setSolidFill(CONFIG.COLORS.background_white);

  const logoRect = layout.getRect('titleSlide.logo');
  try {
    const logo = slide.insertImage(CONFIG.LOGOS.header);
    const aspect = logo.getHeight() / logo.getWidth();
    logo.setLeft(logoRect.left).setTop(logoRect.top).setWidth(logoRect.width).setHeight(logoRect.width * aspect);
  } catch (e) {
    // 画像挿入に失敗した場合はスキップして他の要素を描画
  }

  const titleRect = layout.getRect('titleSlide.title');
  const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, titleRect.left, titleRect.top, titleRect.width, titleRect.height);
  setStyledText(titleShape, data.title, { size: CONFIG.FONTS.sizes.title, bold: true });

  const dateRect = layout.getRect('titleSlide.date');
  const dateShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, dateRect.left, dateRect.top, dateRect.width, dateRect.height);
  dateShape.getText().setText(data.date || '');
  applyTextStyle(dateShape.getText(), { size: CONFIG.FONTS.sizes.date });

  drawBottomBar(slide, layout);
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

// roadmapTimeline（詳細ロードマップ）
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

// --- 8. ユーティリティ関数群 ---
function createLayoutManager(pageW_pt, pageH_pt) {
const pxToPt = (px) => px * 0.75;
const baseW_pt = pxToPt(CONFIG.BASE_PX.W);
const baseH_pt = pxToPt(CONFIG.BASE_PX.H);
const scaleX = pageW_pt / baseW_pt;
const scaleY = pageH_pt / baseH_pt;

const getPositionFromPath = (path) => path.split('.').reduce((obj, key) => obj[key], CONFIG.POS_PX);
return {
scaleX, scaleY, pageW_pt, pageH_pt, pxToPt,
getRect: (spec) => {
const pos = typeof spec === 'string' ? getPositionFromPath(spec) : spec;
let left_px = pos.left;
if (pos.right !== undefined && pos.left === undefined) {
left_px = CONFIG.BASE_PX.W - pos.right - pos.width;
}
return {
left:   left_px !== undefined ? pxToPt(left_px) * scaleX : undefined,
top:    pos.top !== undefined ? pxToPt(pos.top) * scaleY : undefined,
width:  pos.width !== undefined ? pxToPt(pos.width) * scaleX : undefined,
height: pos.height !== undefined ? pxToPt(pos.height) * scaleY : undefined,
};
}
};
}

function offsetRect(rect, dx, dy) {
return { left: rect.left + (dx || 0), top: rect.top + (dy || 0), width: rect.width, height: rect.height };
}

function drawStandardTitleHeader(slide, layout, key, title) {
const logoRect = layout.getRect(`${key}.headerLogo`);
try {
  const logo = slide.insertImage(CONFIG.LOGOS.header);
  const asp = logo.getHeight() / logo.getWidth();
  logo.setLeft(logoRect.left).setTop(logoRect.top).setWidth(logoRect.width).setHeight(logoRect.width * asp);
} catch (e) {
  // 画像挿入に失敗した場合はスキップして他の要素を描画
}

const titleRect = layout.getRect(`${key}.title`);
const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, titleRect.left, titleRect.top, titleRect.width, titleRect.height);
setStyledText(titleShape, title || '', { size: CONFIG.FONTS.sizes.contentTitle, bold: true });

const uRect = layout.getRect(`${key}.titleUnderline`);
const u = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, uRect.left, uRect.top, uRect.width, uRect.height);
u.getFill().setSolidFill(CONFIG.COLORS.primary_color);
u.getBorder().setTransparent();
}

function drawSubheadIfAny(slide, layout, key, subhead) {
if (!subhead) return 0;
const rect = layout.getRect(`${key}.subhead`);
const box = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rect.left, rect.top, rect.width, rect.height);
setStyledText(box, subhead, { size: CONFIG.FONTS.sizes.subhead, color: CONFIG.COLORS.text_primary });
return layout.pxToPt(36);
}

function drawBottomBar(slide, layout) {
const barRect = layout.getRect('bottomBar');
const bar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, barRect.left, barRect.top, barRect.width, barRect.height);
bar.getFill().setSolidFill(CONFIG.COLORS.primary_color);
bar.getBorder().setTransparent();
}

function drawBottomBarAndFooter(slide, layout, pageNum) {
drawBottomBar(slide, layout);
addCucFooter(slide, layout, pageNum);
}

function addCucFooter(slide, layout, pageNum) {
const leftRect = layout.getRect('footer.leftText');
const leftShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, leftRect.left, leftRect.top, leftRect.width, leftRect.height);
leftShape.getText().setText(CONFIG.FOOTER_TEXT);
applyTextStyle(leftShape.getText(), { size: CONFIG.FONTS.sizes.footer, color: CONFIG.COLORS.text_primary });

if (pageNum > 0) {
const rightRect = layout.getRect('footer.rightPage');
const rightShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rightRect.left, rightRect.top, rightRect.width, rightRect.height);
rightShape.getText().setText(String(pageNum));
applyTextStyle(rightShape.getText(), { size: CONFIG.FONTS.sizes.footer, color: CONFIG.COLORS.primary_color, align: SlidesApp.ParagraphAlignment.END });
}
}

function applyTextStyle(textRange, opt) {
const style = textRange.getTextStyle();
style.setFontFamily(CONFIG.FONTS.family);
style.setForegroundColor(opt.color || CONFIG.COLORS.text_primary);
style.setFontSize(opt.size || CONFIG.FONTS.sizes.body);
style.setBold(opt.bold || false);
if (opt.align) {
try { textRange.getParagraphs().forEach(p => p.getRange().getParagraphStyle().setParagraphAlignment(opt.align)); } catch (e) {}
}
}

function setStyledText(shapeOrCell, rawText, baseOpt) {
const parsed = parseInlineStyles(rawText || '');
const tr = shapeOrCell.getText();
tr.setText(parsed.output);
applyTextStyle(tr, baseOpt || {});
applyStyleRanges(tr, parsed.ranges);
}

function setBulletsWithInlineStyles(shape, points) {
const joiner = '\n\n';
let combined = '';
const ranges = [];

(points || []).forEach((pt, idx) => {
const parsed = parseInlineStyles(String(pt || ''));
const bullet = '• ' + parsed.output;
if (idx > 0) combined += joiner;
const start = combined.length;
combined += bullet;

parsed.ranges.forEach(r => {
  ranges.push({ start: start + 2 + r.start, end: start + 2 + r.end, bold: r.bold, color: r.color });
});
});

const tr = shape.getText();
tr.setText(combined || '• —');
applyTextStyle(tr, { size: CONFIG.FONTS.sizes.body });

try {
tr.getParagraphs().forEach(p => {
const ps = p.getRange().getParagraphStyle();
ps.setLineSpacing(100);
ps.setSpaceBelow(6);
});
} catch (e) {}

applyStyleRanges(tr, ranges);
}

function parseInlineStyles(s) {
const ranges = [];
let out = '';
for (let i = 0; i < s.length; ) {
if (s[i] === '[' && s[i+1] === '[') {
const close = s.indexOf(']]', i + 2);
if (close !== -1) {
const content = s.substring(i + 2, close);
const start = out.length;
out += content;
const end = out.length;
ranges.push({ start, end, bold: true, color: CONFIG.COLORS.primary_color });
i = close + 2; continue;
}
}
if (s[i] === '*' && s[i+1] === '*') {
const close = s.indexOf('**', i + 2);
if (close !== -1) {
const content = s.substring(i + 2, close);
const start = out.length;
out += content;
const end = out.length;
ranges.push({ start, end, bold: true });
i = close + 2; continue;
}
}
out += s[i]; i++;
}
return { output: out, ranges };
}

function applyStyleRanges(textRange, ranges) {
ranges.forEach(r => {
try {
const sub = textRange.getRange(r.start, r.end);
if (!sub) return;
const st = sub.getTextStyle();
if (r.bold) st.setBold(true);
if (r.color) st.setForegroundColor(r.color);
} catch (e) {}
});
}

function normalizeImages(arr) {
return (arr || []).map(v => {
if (typeof v === 'string') return { url: v };
if (v && typeof v.url === 'string') return { url: v.url, caption: v.caption || '' };
return null;
}).filter(Boolean).slice(0, 6);
}

function renderImagesInArea(slide, layout, area, images) {
if (!images || images.length === 0) return;
const n = Math.min(6, images.length);
let cols = 1, rows = 1;
if (n === 1) { cols = 1; rows = 1; }
else if (n === 2) { cols = 2; rows = 1; }
else if (n <= 4) { cols = 2; rows = 2; }
else { cols = 3; rows = 2; }

const gap = layout.pxToPt(10);
const cellW = (area.width - gap * (cols - 1)) / cols;
const cellH = (area.height - gap * (rows - 1)) / rows;

for (let i = 0; i < n; i++) {
const r = Math.floor(i / cols), c = i % cols;
const left = area.left + c * (cellW + gap);
const top  = area.top  + r * (cellH + gap);
try {
const img = slide.insertImage(images[i].url);
const scale = Math.min(cellW / img.getWidth(), cellH / img.getHeight());
const w = img.getWidth() * scale;
const h = img.getHeight() * scale;
img.setWidth(w).setHeight(h);
img.setLeft(left + (cellW - w) / 2).setTop(top + (cellH - h) / 2);
} catch(e) {}
}
}

function isAgendaTitle(title) {
const t = String(title || '').toLowerCase();
return /(agenda|アジェンダ|目次|本日お伝えすること)/.test(t);
}

function buildAgendaFromSlideData() {
const pts = [];
for (const d of slideData) {
if (d && d.type === 'section' && typeof d.title === 'string' && d.title.trim()) pts.push(d.title.trim());
}
if (pts.length > 0) return pts.slice(0, 5);
const alt = [];
for (const d of slideData) {
if (d && d.type === 'content' && typeof d.title === 'string' && d.title.trim()) alt.push(d.title.trim());
}
return alt.slice(0, 5);
}

function drawArrowBetweenRects(slide, a, b, arrowH, arrowGap) {
const fromX = a.left + a.width;
const toX   = b.left;
const width = Math.max(0, toX - fromX - arrowGap * 2);
if (width < 8) return;
const yMid = a.top + a.height/2;
const top = yMid - arrowH / 2;
const left = fromX + arrowGap;
const arr = slide.insertShape(SlidesApp.ShapeType.RIGHT_ARROW, left, top, width, arrowH);
arr.getFill().setSolidFill(CONFIG.COLORS.primary_color);
arr.getBorder().setTransparent();
}

function drawNumberedItems(slide, layout, area, items) {
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
const numBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, cx - sz/2, cy - sz/2, sz, sz);
numBox.getFill().setSolidFill(CONFIG.COLORS.primary_color);
numBox.getBorder().setTransparent();
const num = numBox.getText(); num.setText(String(i + 1));
applyTextStyle(num, { size: 12, bold: true, color: CONFIG.COLORS.text_on_primary, align: SlidesApp.ParagraphAlignment.CENTER });

// 元の箇条書きテキストから先頭の数字を除去
let cleanText = String(items[i] || '');
cleanText = cleanText.replace(/^\s*\d+[\.\s]*/, '');

const txt = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, cx + layout.pxToPt(28), cy - layout.pxToPt(16), area.width - layout.pxToPt(70), layout.pxToPt(32));
setStyledText(txt, cleanText, { size: CONFIG.FONTS.sizes.processStep });
try { txt.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch(e){}
}
}

function adjustColorBrightness(hexColor, factor) {
const hex = hexColor.replace('#', '');
const r = parseInt(hex.substr(0, 2), 16);
const g = parseInt(hex.substr(2, 2), 16);
const b = parseInt(hex.substr(4, 2), 16);
const newR = Math.max(0, Math.min(255, Math.round(r * factor)));
const newG = Math.max(0, Math.min(255, Math.round(g * factor)));
const newB = Math.max(0, Math.min(255, Math.round(b * factor)));
return '#' + ((1 << 24) + (newR << 16) + (newG << 8) + newB).toString(16).slice(1);
}

function logInfo(message, meta) {
  if (meta && typeof meta === 'object') {
    try {
      Logger.log(`[Majin] ${message} :: ${JSON.stringify(meta)}`);
      return;
    } catch (e) {
      Logger.log(`[Majin] ${message} :: ${meta}`);
      return;
    }
  }
  Logger.log(`[Majin] ${message}`);
}

function logError(message, error) {
  const payload = error && error.stack ? error.stack : (error && error.message) ? error.message : error;
  Logger.log(`[Majin][Error] ${message} :: ${payload}`);
}
