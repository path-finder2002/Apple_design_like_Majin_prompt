function applyTextStyle(textRange, opt) {
  const style = textRange.getTextStyle();
  style.setFontFamily(CONFIG.FONTS.family).setForegroundColor(opt.color || CONFIG.COLORS.text_primary).setFontSize(opt.size || CONFIG.FONTS.sizes.body).setBold(opt.bold || false);
  if (opt.align) {
    try {
      textRange.getParagraphs().forEach(p => {
        p.getRange().getParagraphStyle().setParagraphAlignment(opt.align);
      });
    } catch (e) {}
  }
}

function setStyledText(shapeOrCell, rawText, baseOpt) {
  const parsed = parseInlineStyles(rawText || '');
  const tr = shapeOrCell.getText().setText(parsed.output);
  applyTextStyle(tr, baseOpt || {});
  applyStyleRanges(tr, parsed.ranges);
}

function setBulletsWithInlineStyles(shape, points) {
  const joiner = '\n\n';
  let combined = '';
  const ranges = [];
  (points || []).forEach((pt, idx) => {
    const parsed = parseInlineStyles(String(pt || ''));
    // 中黒を追加しない、またはオフセット計算を修正
    const bullet = parsed.output;  // '• ' を削除
    if (idx > 0) combined += joiner;
    const start = combined.length;
    combined += bullet;
    parsed.ranges.forEach(r => ranges.push({
      start: start + r.start,  // オフセットを削除
      end: start + r.end,
      bold: r.bold,
      color: r.color
    }));
  });
  const tr = shape.getText().setText(combined || '—');
  applyTextStyle(tr, {
    size: CONFIG.FONTS.sizes.body
  });
  // 箇条書きスタイルを別途適用する場合はここで
  try {
    tr.getParagraphs().forEach(p => {
      p.getRange().getParagraphStyle().setLineSpacing(100).setSpaceBelow(6);
      // 必要に応じて箇条書きプリセットを適用
      // p.getRange().getListStyle().applyListPreset(...);
    });
  } catch (e) {}
  applyStyleRanges(tr, ranges);
}

function parseInlineStyles(s) {
  const ranges = [];
  let out = '';
  let i = 0;
  
  while (i < s.length) {
    // **[[]] 記法を優先的に処理
    if (s[i] === '*' && s[i + 1] === '*' && 
        s[i + 2] === '[' && s[i + 3] === '[') {
      const contentStart = i + 4;
      const close = s.indexOf(']]**', contentStart);
      if (close !== -1) {
        const content = s.substring(contentStart, close);
        const start = out.length;
        out += content;
        const end = out.length;
        const rangeObj = {
          start,
          end,
          bold: true,
          color: CONFIG.COLORS.primary_color
        };
        ranges.push(rangeObj);
        i = close + 4;
        continue;
      }
    }
    
    // [[]] 記法の処理
    if (s[i] === '[' && s[i + 1] === '[') {
      const close = s.indexOf(']]', i + 2);
      if (close !== -1) {
        const content = s.substring(i + 2, close);
        const start = out.length;
        out += content;
        const end = out.length;
        const rangeObj = {
          start,
          end,
          bold: true,
          color: CONFIG.COLORS.primary_color
        };
        ranges.push(rangeObj);
        i = close + 2;
        continue;
      }
    }
    
    // ** 記法の処理
    if (s[i] === '*' && s[i + 1] === '*') {
      const close = s.indexOf('**', i + 2);
      if (close !== -1) {
        const content = s.substring(i + 2, close);
        
        // [[]] が含まれていない場合のみ処理
        if (content.indexOf('[[') === -1) {
          const start = out.length;
          out += content;
          const end = out.length;
          ranges.push({
            start,
            end,
            bold: true
          });
          i = close + 2;
          continue;
        } else {
          // [[]] が含まれている場合は ** をスキップ
          i += 2;
          continue;
        }
      }
    }
    
    out += s[i];
    i++;
  }
  
  return {
    output: out,
    ranges
  };
}

/**
 * スピーカーノートから強調記法を除去する関数
 * @param {string} notesText - 元のノートテキスト
 * @return {string} クリーンなテキスト
 */
function cleanSpeakerNotes(notesText) {
  if (!notesText) return '';
  
  let cleaned = notesText;
  
  // **太字** を除去
  cleaned = cleaned.replace(/\*\*([^*]+)\*\*/g, '$1');
  
  // [[強調語]] を除去
  cleaned = cleaned.replace(/\[\[([^\]]+)\]\]/g, '$1');
  
  // *イタリック* を除去（念のため）
  cleaned = cleaned.replace(/\*([^*]+)\*/g, '$1');
  
  // _下線_ を除去（念のため）
  cleaned = cleaned.replace(/_([^_]+)_/g, '$1');
  
  // ~~取り消し線~~ を除去（念のため）
  cleaned = cleaned.replace(/~~([^~]+)~~/g, '$1');
  
  // `コード` を除去（念のため）
  cleaned = cleaned.replace(/`([^`]+)`/g, '$1');
  
  return cleaned;
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

function isAgendaTitle(title) {
  return /(agenda|アジェンダ|目次|本日お伝えすること)/i.test(String(title || ''));
}

function buildAgendaFromSlideData() {
  return __SLIDE_DATA_FOR_AGENDA.filter(d => d && d.type === 'section' && d.title).map(d => d.title.trim());
}

