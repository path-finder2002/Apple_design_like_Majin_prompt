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
