function createLayoutManager(pageW_pt, pageH_pt) {
  const pxToPt = (px) => px * 0.75;
  const baseW_pt = pxToPt(CONFIG.BASE_PX.W),
    baseH_pt = pxToPt(CONFIG.BASE_PX.H);
  const scaleX = pageW_pt / baseW_pt,
    scaleY = pageH_pt / baseH_pt;
  const getPositionFromPath = (path) => path.split('.').reduce((obj, key) => obj[key], CONFIG.POS_PX);

  return {
    scaleX,
    scaleY,
    pageW_pt,
    pageH_pt,
    pxToPt,
    getRect: (spec) => {
      const pos = typeof spec === 'string' ? getPositionFromPath(spec) : spec;
      let left_px = pos.left;
      if (pos.right !== undefined && pos.left === undefined) {
        left_px = CONFIG.BASE_PX.W - pos.right - pos.width;
      }
      
      // left_px が undefined の場合でも、0 を返すようにする
      if (left_px === undefined && pos.right === undefined) {
        Logger.log(`Warning: Neither left nor right defined for spec: ${JSON.stringify(spec)}`);
        left_px = 0; // デフォルト値
      }
      
      return {
        left: left_px !== undefined ? pxToPt(left_px) * scaleX : 0,
        top: pos.top !== undefined ? pxToPt(pos.top) * scaleY : 0,
        width: pos.width !== undefined ? pxToPt(pos.width) * scaleX : 0,
        height: pos.height !== undefined ? pxToPt(pos.height) * scaleY : 0,
      };
    }
  };
}

// ========================================
// 6. スライド生成関数群
// ========================================

/**
 * 小見出しの高さに応じて本文エリアを動的に調整するヘルパー関数
 * @param {Object} area - 元のエリア定義
 * @param {string} subhead - 小見出しテキスト
 * @param {Object} layout - レイアウトマネージャー
 * @returns {Object} 調整されたエリア定義
 */
function adjustAreaForSubhead(area, subhead, layout) {
  return area;
}

/**
 * コンテンツスライド用の座布団を作成するヘルパー関数（修正版）
 * @param {Object} slide - スライドオブジェクト
 * @param {Object} area - 座布団のエリア定義
 * @param {Object} settings - ユーザー設定
 * @param {Object} layout - レイアウトマネージャー
 */
function createContentCushion(slide, area, settings, layout) {
  if (!area || !area.width || !area.height || area.width <= 0 || area.height <= 0) {
    Logger.log(`[Warning] Invalid area for cushion was provided. Skipping creation. Area: ${JSON.stringify(area)}`);
    return;
  }
  
  // セクションスライドと同じティントグレーの座布団を作成
  const cushionColor = CONFIG.COLORS.background_gray;
  const cushion = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 
    area.left, area.top, area.width, area.height);
  
  cushion.getFill().setSolidFill(cushionColor, 0.50);
  
  // 枠線を完全に削除
  const border = cushion.getBorder();
  border.setTransparent();
}
function offsetRect(rect, dx, dy) {
  return {
    left: rect.left + (dx || 0),
    top: rect.top + (dy || 0),
    width: rect.width,
    height: rect.height
  };
}
function safeGetRect(layout, path) {
  try {
    const rect = layout.getRect(path);
    if (rect && 
        (typeof rect.left === 'number' || rect.left === undefined) && 
        typeof rect.top === 'number' && 
        typeof rect.width === 'number' && 
        typeof rect.height === 'number') {
      
      // leftがundefinedの場合、rightから計算されるべき値が入っていない可能性がある
      // その場合は null を返す
      if (rect.left === undefined) {
        Logger.log(`[safeGetRect] Warning: rect.left is undefined for path ${path}`);
        return null;
      }
      
      return rect;
    }
    // headerLogoパスについてはログを出力しない（よく使われるが存在しない場合が多い）
    if (!path.includes('headerLogo')) {
      Logger.log(`[safeGetRect] Invalid rect for path ${path}:`, JSON.stringify(rect));
    }
    return null;
  } catch (e) {
    // headerLogoパスについてはログを出力しない  
    if (!path.includes('headerLogo')) {
      Logger.log(`[safeGetRect] Error for path ${path}:`, e.message);
    }
    return null;
  }
function findContentRect(layout, key) {
  const candidates = [
    'body',        // contentSlide 等
    'area',        // timeline / process / table / progress 等
    'gridArea',    // cards / kpi / headerCards 等
    'lanesArea',   // diagram
    'pyramidArea', // pyramid
    'stepArea',    // stepUp
    'singleRow',   // flowChart（1行）
    'twoColLeft',  // content 2カラム
    'leftBox',     // compare 左
    'leftText'     // imageText 左テキスト 等
  ];
  for (const name of candidates) {
    const r = safeGetRect(layout, `${key}.${name}`);
    if (r && r.top != null) return r;
  }
  return null;
