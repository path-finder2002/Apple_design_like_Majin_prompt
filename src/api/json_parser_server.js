/**
 * JSON文字列を厳格に検証して slideData 配列へパースします（Apps Script サーバー側）
 * - 空文字チェック
 * - JSON構文チェック
 * - ルートが配列であること
 * - 各要素がオブジェクトであり、type(string) を必須とすること
 *
 * 失敗時は SlideDataParseError を投げます（message はクライアントと揃えています）。
 */
function parseSlideDataStrict(jsonText) {
  var text = String(jsonText || '').trim();
  if (!text) {
    var err1 = new Error('JSON文字列が空です');
    err1.name = 'SlideDataParseError';
    throw err1;
  }
  var data;
  try {
    data = JSON.parse(text);
  } catch (e) {
    var err2 = new Error('JSONの構文エラー: ' + e.message);
    err2.name = 'SlideDataParseError';
    throw err2;
  }
  if (!Array.isArray(data)) {
    var err3 = new Error('slideDataは配列である必要があります');
    err3.name = 'SlideDataParseError';
    throw err3;
  }
  for (var i = 0; i < data.length; i++) {
    var item = data[i];
    if (item === null || typeof item !== 'object' || Array.isArray(item)) {
      var err4 = new Error('スライド#' + (i + 1) + ' はオブジェクトである必要があります');
      err4.name = 'SlideDataParseError';
      throw err4;
    }
    if (!item.type || typeof item.type !== 'string') {
      var err5 = new Error('スライド#' + (i + 1) + ' のtypeは必須です');
      err5.name = 'SlideDataParseError';
      throw err5;
    }
  }
  return data;
}

