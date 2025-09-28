function safeAlignTop(box){
  try { box.setContentAlignment(SlidesApp.ContentAlignment.TOP);
  } catch(e){}
}

// トレンドアイコンを挿入
function insertTrendIcon(slide, position, trend, settings) {
  const iconSize = 20;
  const icon = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
    position.left, position.top - iconSize/2, iconSize, iconSize);
  
  let iconText = '';
  let iconColor = CONFIG.COLORS.text_primary;
  
  switch (trend) {
    case 'up':
      iconText = '↑';
      iconColor = CONFIG.COLORS.success_green;
      break;
    case 'down':
      iconText = '↓';
      iconColor = CONFIG.COLORS.error_red;
      break;
    case 'neutral':
      iconText = '→';
      iconColor = CONFIG.COLORS.neutral_gray;
      break;
    default:
      iconText = '→';
      iconColor = CONFIG.COLORS.neutral_gray;
  }
  
  setStyledText(icon, iconText, {
    size: 16,
    bold: true,
    color: iconColor,
    align: SlidesApp.ParagraphAlignment.CENTER
  });
  
  try {
    icon.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  } catch (e) {}
  
  return icon;
}

/**
 * 文字列から数値のみを抽出するヘルパー関数
 * @param {string} str - 例: "100万円", "75.5%"
 * @return {number} 抽出された数値。見つからない場合は0。
 */
function parseNumericValue(str) {
  if (typeof str !== 'string') return 0;
  const match = str.match(/(\d+(\.\d+)?)/);
  return match ? parseFloat(match[1]) : 0;
}

/**
 * statsCompareスライドの描画関数
 * 中央項目列と白い背景のデザインで統計データを比較表示
 */
function createStatsCompareSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'compareSlide', data.title, settings);
  const dy = drawSubheadIfAny(slide, layout, 'compareSlide', data.subhead);

  const area = offsetRect(layout.getRect({
    left: 25,
    top: 130,
    width: 910,
    height: 330
  }), 0, dy);
  const stats = Array.isArray(data.stats) ? data.stats : [];
  if (stats.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }

  // テーブル全体の背景に白い座布団を追加
  const cushion = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, area.left, area.top, area.width, area.height);
  cushion.getFill().setSolidFill(CONFIG.COLORS.background_white);
  cushion.getBorder().setTransparent();

  // 3列構成の幅配分を最適化（矢印スペース削除により各列を拡大）
  const headerHeight = layout.pxToPt(40);
  const totalContentWidth = area.width;
  const centerColWidth = totalContentWidth * 0.25; // 20% → 25%に拡大
  const sideColWidth = (totalContentWidth - centerColWidth) / 2; // 残りを左右で等分

  const leftValueColX = area.left;
  const centerLabelColX = leftValueColX + sideColWidth;
  const rightValueColX = centerLabelColX + centerColWidth;

  // 項目ラベル用の色を生成
  const labelColor = generateTintedGray(settings.primaryColor, 35, 70);

  const compareColors = generateCompareColors(settings.primaryColor);
  
  const leftHeader = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, leftValueColX, area.top, sideColWidth, headerHeight);
  leftHeader.getFill().setSolidFill(compareColors.left); // 左側：暗い色（視認性向上）
  leftHeader.getBorder().setTransparent();
  setStyledText(leftHeader, data.leftTitle || '', {
    size: 14,
    bold: true,
    color: CONFIG.COLORS.background_white,
    align: SlidesApp.ParagraphAlignment.CENTER
  });
  try { leftHeader.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) {}

  const rightHeader = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, rightValueColX, area.top, sideColWidth, headerHeight);
  rightHeader.getFill().setSolidFill(compareColors.right); // 右側：元の色
  rightHeader.getBorder().setTransparent();
  setStyledText(rightHeader, data.rightTitle || '', {
    size: 14,
    bold: true,
    color: CONFIG.COLORS.background_white,
    align: SlidesApp.ParagraphAlignment.CENTER
  });
  try { rightHeader.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) {}

  const contentAreaHeight = area.height - headerHeight;
  const rowHeight = contentAreaHeight / stats.length;
  let currentY = area.top + headerHeight;

  stats.forEach((stat, index) => {
    const centerY = currentY + rowHeight / 2;
    const valueHeight = layout.pxToPt(40);

    // 項目ラベル (中央列)
    const labelShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, centerLabelColX, centerY - valueHeight / 2, centerColWidth, valueHeight);
    setStyledText(labelShape, stat.label || '', {
      size: 14,
      align: SlidesApp.ParagraphAlignment.CENTER,
      color: labelColor,
      bold: true
    });
    try { labelShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) {}

    // 左の値（拡大されたスペースを活用）
    const leftValueShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, leftValueColX, centerY - valueHeight / 2, sideColWidth, valueHeight);
    setStyledText(leftValueShape, stat.leftValue || '', {
      size: 22,
      bold: true,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
    try { leftValueShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) {}

    // 右の値（拡大されたスペースを活用・矢印スペース不要）
    const rightValueShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rightValueColX, centerY - valueHeight / 2, sideColWidth, valueHeight);
    setStyledText(rightValueShape, stat.rightValue || '', {
      size: 22,
      bold: true,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
    try { rightValueShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch (e) {}

    // 矢印描画処理を完全に削除（この部分が不要）

    // アンダーラインを描画
    if (index < stats.length - 1) {
      const lineY = currentY + rowHeight;
      const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, area.left + layout.pxToPt(15), lineY, area.left + area.width - layout.pxToPt(15), lineY);
      line.getLineFill().setSolidFill(CONFIG.COLORS.faint_gray);
      line.setWeight(1);
    }
    
    currentY += rowHeight;
  });

  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

/**
 * 新しいスライドタイプ：バーチャートでの数値比較 (レイアウト調整版)
 */
function createBarCompareSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'compareSlide', data.title, settings);
  const dy = drawSubheadIfAny(slide, layout, 'compareSlide', data.subhead);

  const area = offsetRect(layout.getRect({
    left: 40,
    top: 130,
    width: 880,
    height: 340
  }), 0, dy);
  const stats = Array.isArray(data.stats) ? data.stats : [];
  if (stats.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }

  // showTrendsオプション（デフォルトはfalse）
  const showTrends = !!data.showTrends;

  // 各ブロック間のマージンを定義
  const blockMargin = layout.pxToPt(20);
  const totalContentHeight = area.height - (blockMargin * (stats.length - 1));
  const blockHeight = totalContentHeight / stats.length;
  let currentY = area.top;

  stats.forEach(stat => {
    const blockTop = currentY;
    const titleHeight = layout.pxToPt(35); // タイトルエリアを少し調整
    const barAreaHeight = blockHeight - titleHeight;
    const barRowHeight = barAreaHeight / 2;

    // トレンド表示の判定
    const shouldShowTrend = showTrends && stat.trend;

    // 項目タイトル
    const statTitleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, blockTop, area.width, titleHeight);
    // フォントサイズを調整
    setStyledText(statTitleShape, stat.label || '', {
      size: 16,
      bold: true
    });
    try { statTitleShape.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM); } catch(e){}

    const asIsY = blockTop + titleHeight;
    const toBeY = asIsY + barRowHeight;

    const labelWidth = layout.pxToPt(90);  // ラベル幅を少し調整
    const valueWidth = layout.pxToPt(140); // 値幅を少し調整
    const barWidth = Math.max(layout.pxToPt(50), area.width - labelWidth - valueWidth - layout.pxToPt(10));
    const barLeft = area.left + labelWidth;
    const barHeight = layout.pxToPt(18); // バーの高さを調整

    const val1 = parseNumericValue(stat.leftValue);
    const val2 = parseNumericValue(stat.rightValue);
    const maxValue = Math.max(val1, val2, 1);

    // 1. 現状 (As-Is) の行
    const asIsLabel = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, asIsY, labelWidth, barRowHeight);
    setStyledText(asIsLabel, '現状', { size: 13, color: CONFIG.COLORS.neutral_gray }); // フォントサイズを調整
    const asIsValue = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, barLeft + barWidth, asIsY, valueWidth, barRowHeight);
    setStyledText(asIsValue, stat.leftValue || '', { size: 18, bold: true, align: SlidesApp.ParagraphAlignment.END }); // フォントサイズを調整
    
    // バーの形状を角丸四角形で描画
    const asIsTrack = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, barLeft, asIsY + barRowHeight/2 - barHeight/2, barWidth, barHeight);
    asIsTrack.getFill().setSolidFill(CONFIG.COLORS.faint_gray);
    asIsTrack.getBorder().setTransparent();
    const asIsFillWidth = Math.max(layout.pxToPt(2), barWidth * (val1 / maxValue));
    const asIsFill = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, barLeft, asIsY + barRowHeight/2 - barHeight/2, asIsFillWidth, barHeight);
    asIsFill.getFill().setSolidFill(CONFIG.COLORS.neutral_gray);
    asIsFill.getBorder().setTransparent();

    // 2. 導入後 (To-Be) の行
    const toBeLabel = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, area.left, toBeY, labelWidth, barRowHeight);
    setStyledText(toBeLabel, '導入後', { size: 13, color: settings.primaryColor, bold: true }); // フォントサイズを調整
    const toBeValue = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, barLeft + barWidth, toBeY, valueWidth, barRowHeight);
    setStyledText(toBeValue, stat.rightValue || '', { size: 18, bold: true, color: settings.primaryColor, align: SlidesApp.ParagraphAlignment.END }); // フォントサイズを調整
    
    // バーの形状を角丸四角形で描画
    const toBeTrack = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, barLeft, toBeY + barRowHeight/2 - barHeight/2, barWidth, barHeight);
    toBeTrack.getFill().setSolidFill(generateTintedGray(settings.primaryColor, 20, 96));
    toBeTrack.getBorder().setTransparent();
    const toBeFillWidth = Math.max(layout.pxToPt(2), barWidth * (val2 / maxValue));
    
    // プライマリカラーで塗りつぶし
    const shape = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, barLeft, toBeY + barRowHeight/2 - barHeight/2, toBeFillWidth, barHeight);
    shape.getFill().setSolidFill(settings.primaryColor);
    shape.getBorder().setTransparent();

    try {
        [asIsLabel, asIsValue, toBeLabel, toBeValue].forEach(shape => shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE));
    } catch(e){}

    if (shouldShowTrend) {
      const trendIcon = insertTrendIcon(slide, { left: barLeft + barWidth + layout.pxToPt(10), top: toBeY + barRowHeight/2 }, stat.trend, settings);
    }

    // 次のブロックへのマージンを追加
    currentY += blockHeight + blockMargin;
  });

  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

/**
 * 新しいスライドタイプ：3要素のトライアングル循環図 (直線矢印版)
 */
/**
 * Triangle専用：テキストをヘッダー/ボディに自動分離
 */
function smartFormatTriangleText(text) {
  if (!text || text.length <= 30) {
    // 短文はそのまま
    return { text: text, isSimple: true, headerLength: 0 };
  }
  
  // 分離候補パターンを試行
  const separators = [
    { pattern: '：', priority: 1 },
    { pattern: ':', priority: 2 },
    { pattern: '。', priority: 3 },
    { pattern: 'について', priority: 4, keepSeparator: true },
    { pattern: 'における', priority: 5, keepSeparator: true }
  ];
  
  for (let sep of separators) {
    const index = text.indexOf(sep.pattern);
    if (index > 5 && index < text.length * 0.6) { // 5文字以上、前60%以内
      const headerEnd = sep.keepSeparator ? index + sep.pattern.length : index;
      const header = text.substring(0, headerEnd).trim();
      const body = text.substring(index + sep.pattern.length).trim();
      
      if (header.length >= 3 && body.length >= 3) {
        return {
          text: `${header}\n${body}`,
          isSimple: false,
          headerLength: header.length
        };
      }
    }
  }
  
  // フォールバック：バランス良く分割
  if (text.length > 50) {
    const midPoint = Math.floor(text.length * 0.4);
    const header = text.substring(0, midPoint).trim();
    const body = text.substring(midPoint).trim();
    
    return {
      text: `${header}\n${body}`,
      isSimple: false,
      headerLength: header.length
    };
  }
  
  // それでも短い場合はそのまま
  return { text: text, isSimple: true, headerLength: 0 };
}

function createTriangleSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'triangleSlide', data.title, settings);
  const dy = drawSubheadIfAny(slide, layout, 'triangleSlide', data.subhead);
  
  // 小見出しの高さに応じてトライアングルエリアを動的に調整
  const baseArea = layout.getRect('triangleSlide.area');
  const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
  const area = offsetRect(adjustedArea, 0, dy);

  // 表示するアイテムを3つに限定
  const items = Array.isArray(data.items) ? data.items.slice(0, 3) : [];
  if (items.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }

  // 各アイテムの文字数を分析
  const textLengths = items.map(item => (item || '').length);
  const maxLength = Math.max(...textLengths);
  const avgLength = textLengths.reduce((sum, len) => sum + len, 0) / textLengths.length;
  
  // 文字数に基づいてカードサイズとフォントサイズを動的調整
  let cardW, cardH, fontSize;
  if (maxLength > 60 || avgLength > 40) {
    // 長文対応：大きめカード + 小さめフォント
    cardW = layout.pxToPt(340);
    cardH = layout.pxToPt(160);
    fontSize = 13; // 長文用小さめフォント
  } else if (maxLength > 35 || avgLength > 25) {
    // 中文対応：標準カード + 標準フォント
    cardW = layout.pxToPt(290);
    cardH = layout.pxToPt(135);
    fontSize = 14; // 中文用標準フォント
  } else {
    // 短文対応：コンパクトカード + 大きめフォント
    cardW = layout.pxToPt(250);
    cardH = layout.pxToPt(115);
    fontSize = 15; // 短文用大きめフォント
  }
  
  // 利用可能エリアに基づく最大サイズ制限
  const maxCardW = (area.width - layout.pxToPt(160)) / 1.5; // 左右余白考慮
  const maxCardH = (area.height - layout.pxToPt(80)) / 2;   // 上下余白考慮
  
  cardW = Math.min(cardW, maxCardW);
  cardH = Math.min(cardH, maxCardH);

  // 3つの頂点の中心座標を計算
  const positions = [
    // 頂点0: 上
    {
      x: area.left + area.width / 2,
      y: area.top + layout.pxToPt(40) + cardH / 2
    },
    // 頂点1: 右下
    {
      x: area.left + area.width - layout.pxToPt(80) - cardW / 2,
      y: area.top + area.height - cardH / 2
    },
    // 頂点2: 左下
    {
      x: area.left + layout.pxToPt(80) + cardW / 2,
      y: area.top + area.height - cardH / 2
    }
  ];

  positions.forEach((pos, i) => {
    if (!items[i]) return; // アイテムが3つ未満の場合に対応

    const cardX = pos.x - cardW / 2;
    const cardY = pos.y - cardH / 2;
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, cardX, cardY, cardW, cardH);
    card.getFill().setSolidFill(CONFIG.COLORS.background_gray); // ティントグレー背景
    card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    
    // 構造化データまたは文字列データに対応
    const item = items[i] || {};
    const itemTitle = typeof item === 'string' ? '' : (item.title || '');
    const itemDesc = typeof item === 'string' ? item : (item.desc || '');
    
    if (typeof item === 'string' || !itemTitle) {
      // 従来の文字列形式または構造化されていない場合：自動分離を試行
      const itemText = typeof item === 'string' ? item : itemDesc;
      const processedText = smartFormatTriangleText(itemText);

      if (processedText.isSimple) {
        // シンプルテキスト：従来通り
        let appliedFontSize = fontSize;
        if ((processedText.text || '').length > 35) {
          appliedFontSize = Math.max(fontSize - 1, 12);
        }

        setStyledText(card, processedText.text, {
          size: appliedFontSize,
          bold: true,
          color: CONFIG.COLORS.text_primary, // 黒文字に変更
          align: SlidesApp.ParagraphAlignment.CENTER
        });
      } else {
        // 自動分離：見出し+本文構造
        const lines = processedText.text.split('\n');
        const header = lines[0] || '';
        const body = lines.slice(1).join('\n') || '';
        
        const enhancedText = `${header}\n${body}`;
        const headerFontSize = Math.max(fontSize - 1, 13);
        const bodyFontSize = Math.max(fontSize - 3, 11);
        
        setStyledText(card, enhancedText, {
          size: bodyFontSize,
          bold: false,
          color: CONFIG.COLORS.text_primary, // 本文：黒文字
          align: SlidesApp.ParagraphAlignment.CENTER
        });
        
        try {
          const textRange = card.getText();
          const headerEndIndex = header.length;
          if (headerEndIndex > 0) {
            const headerRange = textRange.getRange(0, headerEndIndex);
            headerRange.getTextStyle()
              .setBold(true)
              .setFontSize(headerFontSize)
              .setForegroundColor(settings.primaryColor); // 見出し：プライマリカラー
          }
        } catch (e) {}
      }
    } else {
      // 新しい構造化形式：title + desc
      const enhancedText = itemDesc ? `${itemTitle}\n${itemDesc}` : itemTitle;
      const headerFontSize = Math.max(fontSize - 1, 13); // 見出し
      const bodyFontSize = Math.max(fontSize - 3, 11);   // 本文
      
      setStyledText(card, enhancedText, {
        size: bodyFontSize, // ベースは本文サイズ
        bold: false,        // ベースは通常フォント
        color: CONFIG.COLORS.text_primary, // 本文：黒文字
        align: SlidesApp.ParagraphAlignment.CENTER
      });
      
      // 見出し部分のスタイリング
      try {
        const textRange = card.getText();
        const headerEndIndex = itemTitle.length;
        if (headerEndIndex > 0) {
          const headerRange = textRange.getRange(0, headerEndIndex);
          headerRange.getTextStyle()
            .setBold(true)
            .setFontSize(headerFontSize)
            .setForegroundColor(settings.primaryColor); // 見出し：プライマリカラー
        }
      } catch (e) {}
    }
    
    try {
      card.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
  });

  // 動的サイズに応じた余白調整
  const arrowPadding = cardW > layout.pxToPt(300) ? layout.pxToPt(25) : layout.pxToPt(20);

  // カードの辺の中央座標を計算
  const cardEdges = [
    // 上のカード
    {
      rightCenter: { x: positions[0].x + cardW / 2, y: positions[0].y },
      leftCenter: { x: positions[0].x - cardW / 2, y: positions[0].y },
      bottomCenter: { x: positions[0].x, y: positions[0].y + cardH / 2 }
    },
    // 右下のカード
    {
      leftCenter: { x: positions[1].x - cardW / 2, y: positions[1].y },
      rightCenter: { x: positions[1].x + cardW / 2, y: positions[1].y },
      topCenter: { x: positions[1].x, y: positions[1].y - cardH / 2 }
    },
    // 左下のカード
    {
      rightCenter: { x: positions[2].x + cardW / 2, y: positions[2].y },
      leftCenter: { x: positions[2].x - cardW / 2, y: positions[2].y },
      topCenter: { x: positions[2].x, y: positions[2].y - cardH / 2 }
    }
  ];

  // 自然な曲線の矢印を描画
  const arrowCurves = [
    // 上→右下：上のカードの右辺中央から右下のカードの上辺中央へ
    {
      startX: cardEdges[0].rightCenter.x + arrowPadding,
      startY: cardEdges[0].rightCenter.y,
      endX: cardEdges[1].topCenter.x,
      endY: cardEdges[1].topCenter.y - arrowPadding,
      controlX: (cardEdges[0].rightCenter.x + cardEdges[1].topCenter.x) / 2 + arrowPadding,
      controlY: (cardEdges[0].rightCenter.y + cardEdges[1].topCenter.y) / 2
    },
    // 右下→左下：右下のカードの左辺中央から左下のカードの右辺中央へ
    {
      startX: cardEdges[1].leftCenter.x - arrowPadding,
      startY: cardEdges[1].leftCenter.y,
      endX: cardEdges[2].rightCenter.x + arrowPadding,
      endY: cardEdges[2].rightCenter.y,
      controlX: (cardEdges[1].leftCenter.x + cardEdges[2].rightCenter.x) / 2,
      controlY: (cardEdges[1].leftCenter.y + cardEdges[2].rightCenter.y) / 2
    },
    // 左下→上：左下のカードの上辺中央から上のカードの左辺中央へ
    {
      startX: cardEdges[2].topCenter.x,
      startY: cardEdges[2].topCenter.y - arrowPadding,
      endX: cardEdges[0].leftCenter.x - arrowPadding,
      endY: cardEdges[0].leftCenter.y,
      controlX: (cardEdges[2].topCenter.x + cardEdges[0].leftCenter.x) / 2,
      controlY: (cardEdges[2].topCenter.y + cardEdges[0].leftCenter.y) / 2
    }
  ];

  arrowCurves.forEach(curve => {
    const line = slide.insertLine(
      SlidesApp.LineCategory.STRAIGHT,
      curve.startX,
      curve.startY,
      curve.endX,
      curve.endY
    );
    line.getLineFill().setSolidFill(CONFIG.COLORS.ghost_gray);
    line.setWeight(4);
    line.setEndArrow(SlidesApp.ArrowStyle.FILL_ARROW);
  });

  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

/**
 * ピラミッドスライドを生成
 * @param {Slide} slide - スライドオブジェクト
 * @param {Object} data - スライドデータ
 * @param {Object} layout - レイアウトマネージャー
 * @param {number} pageNum - ページ番号
 * @param {Object} settings - ユーザー設定
 */
function createPyramidSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'pyramidSlide', data.title, settings);
  const dy = drawSubheadIfAny(slide, layout, 'pyramidSlide', data.subhead);
  
  // 小見出しの高さに応じてピラミッドエリアを動的に調整
  const baseArea = layout.getRect('pyramidSlide.pyramidArea');
  const adjustedArea = adjustAreaForSubhead(baseArea, data.subhead, layout);
  const area = offsetRect(adjustedArea, 0, dy);

  // ピラミッドのレベルデータを取得（最大4レベル）
  const levels = Array.isArray(data.levels) ? data.levels.slice(0, 4) : [];
  if (levels.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }

  const levelHeight = layout.pxToPt(70); // 高さ調整
  const levelGap = layout.pxToPt(2); // 余白を大幅縮小（5px→2px）
  const totalHeight = (levelHeight * levels.length) + (levelGap * (levels.length - 1));
  const startY = area.top + (area.height - totalHeight) / 2;

  // ピラミッドとテキストカラムのレイアウト
  const pyramidWidth = layout.pxToPt(480); // 幅調整
  const textColumnWidth = layout.pxToPt(400); // テキストエリア拡大
  const gap = layout.pxToPt(30); // ピラミッドとテキスト間の間隔
  
  const pyramidLeft = area.left;
  const textColumnLeft = pyramidLeft + pyramidWidth + gap;
  

  // カラーグラデーション生成
  const pyramidColors = generatePyramidColors(settings.primaryColor, levels.length);
  
  // 各レベルの幅を計算（上から下に向かって広がる）
  const baseWidth = pyramidWidth;
  const widthIncrement = baseWidth / levels.length;
  const centerX = pyramidLeft + pyramidWidth / 2; // ピラミッドの中央基準

  levels.forEach((level, index) => {
    const levelWidth = baseWidth - (widthIncrement * (levels.length - 1 - index)); // 逆順で計算
    const levelX = centerX - levelWidth / 2; // 中央揃え
    const levelY = startY + index * (levelHeight + levelGap);

    // ピラミッドレベルボックス（グラデーションカラー適用）
    const levelBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, levelX, levelY, levelWidth, levelHeight);
    levelBox.getFill().setSolidFill(pyramidColors[index]); // グラデーションカラー
    levelBox.getBorder().setTransparent();

    // ピラミッド内のタイトルテキスト（簡潔に）
    const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, levelX, levelY, levelWidth, levelHeight);
    titleShape.getFill().setTransparent();
    titleShape.getBorder().setTransparent();

    const levelTitle = level.title || `レベル${index + 1}`;
    setStyledText(titleShape, levelTitle, {
      size: CONFIG.FONTS.sizes.body,
      bold: true,
      color: CONFIG.COLORS.background_white,
      align: SlidesApp.ParagraphAlignment.CENTER
    });

    try {
      titleShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}

    // 接続線の描画（ピラミッド右端からテキストエリア左端へ）
    const connectionStartX = levelX + levelWidth;
    const connectionEndX = textColumnLeft;
    const connectionY = levelY + levelHeight / 2;
    
    if (connectionEndX > connectionStartX) {
      const connectionLine = slide.insertLine(
        SlidesApp.LineCategory.STRAIGHT,
        connectionStartX,
        connectionY,
        connectionEndX,
        connectionY
      );
      connectionLine.getLineFill().setSolidFill('#D0D7DE'); // 薄いグレー接続線
      connectionLine.setWeight(1.5);
    }

    // 右側のテキストカラムに詳細説明を配置
    const textShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textColumnLeft, levelY, textColumnWidth, levelHeight);
    textShape.getFill().setTransparent();
    textShape.getBorder().setTransparent();

    const levelDesc = level.description || '';
    
    // description内容の整理（箇条書き対応）
    let formattedText;
    if (levelDesc.includes('•') || levelDesc.includes('・')) {
      // 既に箇条書きの場合はそのまま
      formattedText = levelDesc;
    } else if (levelDesc.includes('\n')) {
      // 改行区切りの場合は箇条書きに変換
      const lines = levelDesc.split('\n').filter(line => line.trim()).slice(0, 2);
      formattedText = lines.map(line => `• ${line.trim()}`).join('\n');
    } else {
      // 単一文の場合はそのまま
      formattedText = levelDesc;
    }

    setStyledText(textShape, formattedText, {
      size: CONFIG.FONTS.sizes.body - 1, // 少し小さめ
      align: SlidesApp.ParagraphAlignment.LEFT,
      color: CONFIG.COLORS.text_primary
    });

    try {
      textShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
  });

  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

/**
 * flowChartスライドを生成
 * @param {Slide} slide - スライドオブジェクト
 * @param {Object} data - スライドデータ
 * @param {Object} layout - レイアウトマネージャー
 * @param {number} pageNum - ページ番号
 * @param {Object} settings - ユーザー設定
 */
function createFlowChartSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'flowChartSlide', data.title, settings);
  const dy = drawSubheadIfAny(slide, layout, 'flowChartSlide', data.subhead);
  
  const flows = Array.isArray(data.flows) ? data.flows : [{ steps: data.steps || [] }];
  
  // 2行レイアウトの判定: 複数フローがあるか、単一フローでも5ステップ以上の場合
  let isDouble = flows.length > 1;
  let upperFlow, lowerFlow, maxStepsPerRow;
  
  if (!isDouble && flows[0] && flows[0].steps && flows[0].steps.length >= 5) {
    // 単一フローでも5ステップ以上なら2行に分割
    isDouble = true;
    const allSteps = flows[0].steps;
    const midPoint = Math.ceil(allSteps.length / 2);
    upperFlow = { steps: allSteps.slice(0, midPoint) };
    lowerFlow = { steps: allSteps.slice(midPoint) };
    maxStepsPerRow = midPoint; // 上段の枚数を基準とする
  } else {
    upperFlow = flows[0];
    lowerFlow = flows.length > 1 ? flows[1] : null; // null を明示的に設定
    maxStepsPerRow = Math.max(
      upperFlow?.steps?.length || 0, 
      lowerFlow?.steps?.length || 0
    );
  }
  
  if (isDouble) {
    // 2行レイアウト（統一カード幅）
    const upperArea = offsetRect(layout.getRect('flowChartSlide.upperRow'), 0, dy);
    const lowerArea = offsetRect(layout.getRect('flowChartSlide.lowerRow'), 0, dy);
    drawFlowRow(slide, upperFlow, upperArea, settings, layout, maxStepsPerRow);
    if (lowerFlow && lowerFlow.steps && lowerFlow.steps.length > 0) { // より厳密なチェック
      drawFlowRow(slide, lowerFlow, lowerArea, settings, layout, maxStepsPerRow);
    }
  } else {
    // 1行レイアウト
    const singleArea = offsetRect(layout.getRect('flowChartSlide.singleRow'), 0, dy);
    drawFlowRow(slide, flows[0], singleArea, settings, layout);
  }
  
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

/**
 * フローチャートの1行を描画
 * @param {Slide} slide - スライドオブジェクト
 * @param {Object} flow - フローデータ
 * @param {Object} area - 描画エリア
 * @param {Object} settings - ユーザー設定
 * @param {Object} layout - レイアウトマネージャー
 * @param {number} maxStepsPerRow - 1行あたりの最大ステップ数（統一カード幅用）
 */
function drawFlowRow(slide, flow, area, settings, layout, maxStepsPerRow = null) {
  // 安全性チェックを強化
  if (!flow || !flow.steps || !Array.isArray(flow.steps)) {
    return; // 早期リターン
  }
  
  const steps = flow.steps.filter(step => step && String(step).trim()); // 空要素を除去
  if (steps.length === 0) return;
  
  // 統一カード幅の計算（2行レイアウト時）
  const actualSteps = maxStepsPerRow || steps.length;
  
  // カード重視のレイアウト調整（矢印間隔を最小限に）
  const baseArrowSpace = layout.pxToPt(25); // 矢印間隔を縮小してカード幅を拡大
  const arrowSpace = Math.max(baseArrowSpace, area.width * 0.04); // エリア幅の4%を最小値に縮小
  const totalArrowSpace = (actualSteps - 1) * arrowSpace; // 統一基準で計算
  const cardW = (area.width - totalArrowSpace) / actualSteps; // 統一カード幅
  const cardH = area.height;
  
  // カードサイズに応じて矢印サイズを動的調整（コンパクト設計）
  const arrowHeight = Math.min(cardH * 0.3, layout.pxToPt(40)); // 高さを少し縮小
  const arrowWidth = arrowSpace; // 矢印幅をスペース全体に設定（カード間を完全に埋める）
  
  steps.forEach((step, index) => {
    const cardX = area.left + index * (cardW + arrowSpace);
    
    // カード描画（既存のcardsデザイン流用）
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, cardX, area.top, cardW, cardH);
    card.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    card.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    
    // テキスト設定の安全性を向上
    const stepText = String(step || '').trim() || 'ステップ'; // フォールバック値を設定
    setStyledText(card, stepText, {
      size: CONFIG.FONTS.sizes.body,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
    
    try {
      card.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    } catch (e) {}
    
    // 矢印描画（最後以外）- カード間を直接つなぐ
    if (index < steps.length - 1) {
      const arrowStartX = cardX + cardW; // 現在のカードの右端
      const arrowCenterY = area.top + cardH / 2;
      const arrowTop = arrowCenterY - (arrowHeight / 2);
      
      // 右向き矢印図形を使用（カードの右端から次のカードまで）
      const arrow = slide.insertShape(SlidesApp.ShapeType.RIGHT_ARROW, arrowStartX, arrowTop, arrowWidth, arrowHeight);
      arrow.getFill().setSolidFill(settings.primaryColor);
      arrow.getBorder().setTransparent();
    }
  });
}

/**
 * stepUpスライドを生成（階段状に成長するヘッダー付きカード）
 * @param {Slide} slide - スライドオブジェクト
 * @param {Object} data - スライドデータ
 * @param {Object} layout - レイアウトマネージャー
 * @param {number} pageNum - ページ番号
 * @param {Object} settings - ユーザー設定
 */
function createStepUpSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'stepUpSlide', data.title, settings);
  const dy = drawSubheadIfAny(slide, layout, 'stepUpSlide', data.subhead);
  const area = offsetRect(layout.getRect('stepUpSlide.stepArea'), 0, dy);
  const items = Array.isArray(data.items) ? data.items : [];
  
  if (items.length === 0) {
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
    return;
  }
  
  const numSteps = Math.min(5, items.length); // 最大5ステップ
  const gap = 0; // 余白なし（くっつける）
  const headerHeight = layout.pxToPt(40);
  
  const maxHeight = area.height * 0.95; // エリアの95%を最大高さに
  
  // ステップ数に応じて最小高さを動的調整（全体のバランスを考慮）
  let minHeightRatio;
  if (numSteps <= 2) {
    minHeightRatio = 0.70; // 2ステップ：最小でも70%の高さ
  } else if (numSteps === 3) {
    minHeightRatio = 0.60; // 3ステップ：最小60%の高さ
  } else {
    minHeightRatio = 0.50; // 4-5ステップ：最小50%の高さ
  }
  const minHeight = maxHeight * minHeightRatio;
  
  // 総幅から各カードの幅を計算
  const totalWidth = area.width;
  const cardW = totalWidth / numSteps;
  
  // StepUpカラーグラデーション生成（左から右に濃くなる）
  const stepUpColors = generateStepUpColors(settings.primaryColor, numSteps);
  
  for (let idx = 0; idx < numSteps; idx++) {
    const item = items[idx] || {};
    const titleText = String(item.title || `STEP ${idx + 1}`);
    const descText = String(item.desc || '');
    
    // 階段状に高さを計算（線形に増加）
    const heightRatio = (idx / Math.max(1, numSteps - 1)); // 0から1の比率
    const cardH = minHeight + (maxHeight - minHeight) * heightRatio;
    
    const left = area.left + idx * cardW;
    const top = area.top + area.height - cardH; // 下端揃え
    
    // ヘッダー部分（グラデーションカラー適用）
    const headerShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top, cardW, headerHeight);
    headerShape.getFill().setSolidFill(stepUpColors[idx]); // グラデーションカラー
    headerShape.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    
    // ボディ部分
    const bodyShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top + headerHeight, cardW, cardH - headerHeight);
    bodyShape.getFill().setSolidFill(CONFIG.COLORS.background_gray);
    bodyShape.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.card_border);
    
    // ヘッダーテキスト
    const headerTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left, top, cardW, headerHeight);
    setStyledText(headerTextShape, titleText, {
      size: CONFIG.FONTS.sizes.body,
      bold: true,
      color: CONFIG.COLORS.background_white,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
    
    try {
      headerTextShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      headerTextShape.setAutofit(SlidesApp.AutofitType.SHRINK_ON_OVERFLOW);
    } catch (e) {}
    
    // ボディテキスト
    const bodyTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
      left + layout.pxToPt(8), top + headerHeight, 
      cardW - layout.pxToPt(16), cardH - headerHeight);
    setStyledText(bodyTextShape, descText, {
      size: CONFIG.FONTS.sizes.body,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
    
    try {
      bodyTextShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      bodyTextShape.setAutofit(SlidesApp.AutofitType.SHRINK_ON_OVERFLOW);
    } catch (e) {}
  }
  
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}

/**
 * imageTextスライドを生成（画像とテキストの2カラム）
 * @param {Slide} slide - スライドオブジェクト
 * @param {Object} data - スライドデータ
 * @param {Object} layout - レイアウトマネージャー
 * @param {number} pageNum - ページ番号
 * @param {Object} settings - ユーザー設定
 */
function createImageTextSlide(slide, data, layout, pageNum, settings) {
  setMainSlideBackground(slide, layout);
  drawStandardTitleHeader(slide, layout, 'imageTextSlide', data.title, settings);
  const dy = drawSubheadIfAny(slide, layout, 'imageTextSlide', data.subhead);
  
  const imageUrl = data.image || '';
  const imageCaption = data.imageCaption || '';
  const points = Array.isArray(data.points) ? data.points : [];
  const imagePosition = data.imagePosition === 'right' ? 'right' : 'left'; // デフォルトは左
  
  if (imagePosition === 'left') {
    // 左に画像、右にテキスト
    const imageArea = offsetRect(layout.getRect('imageTextSlide.leftImage'), 0, dy);
    const textArea = offsetRect(layout.getRect('imageTextSlide.rightText'), 0, dy);
    
    if (imageUrl) {
      renderSingleImageInArea(slide, layout, imageArea, imageUrl, imageCaption, 'left');
    }
    
    if (points.length > 0) {
      // テキストエリアに座布団を作成
      createContentCushion(slide, textArea, settings, layout);
      
      // テキストボックスを座布団の内側に配置（パディングを追加）
      const padding = layout.pxToPt(20); // 20pxのパディング
      const textRect = {
        left: textArea.left + padding,
        top: textArea.top + padding,
        width: textArea.width - (padding * 2),
        height: textArea.height - (padding * 2)
      };
      const textShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textRect.left, textRect.top, textRect.width, textRect.height);
      setBulletsWithInlineStyles(textShape, points);
    }
  } else {
    // 左にテキスト、右に画像
    const textArea = offsetRect(layout.getRect('imageTextSlide.leftText'), 0, dy);
    const imageArea = offsetRect(layout.getRect('imageTextSlide.rightImage'), 0, dy);
    
    if (points.length > 0) {
      // テキストエリアに座布団を作成
      createContentCushion(slide, textArea, settings, layout);
      
      // テキストボックスを座布団の内側に配置（パディングを追加）
      const padding = layout.pxToPt(20); // 20pxのパディング
      const textRect = {
        left: textArea.left + padding,
        top: textArea.top + padding,
        width: textArea.width - (padding * 2),
        height: textArea.height - (padding * 2)
      };
      const textShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, textRect.left, textRect.top, textRect.width, textRect.height);
      setBulletsWithInlineStyles(textShape, points);
    }
    
    if (imageUrl) {
      renderSingleImageInArea(slide, layout, imageArea, imageUrl, imageCaption, 'right');
    }
  }
  
  drawBottomBarAndFooter(slide, layout, pageNum, settings);
}
