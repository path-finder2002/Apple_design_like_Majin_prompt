function logDebug(message, payload) {
  const prefix = `[SlideGenerator] ${message}`;
  if (typeof payload === 'undefined') {
    Logger.log(prefix);
    return;
  }
  try {
    Logger.log(`${prefix} :: ${JSON.stringify(payload)}`);
  } catch (error) {
    Logger.log(`${prefix} :: (payload serialization failed: ${error.message})`);
  }
}

function doGet(e) {
  logDebug('doGet invoked', {
    queryKeys: e && e.parameter ? Object.keys(e.parameter) : []
  });
  const htmlTemplate = HtmlService.createTemplateFromFile('index');
  htmlTemplate.settings = loadSettings();
  const result = htmlTemplate.evaluate().setTitle('Google Slide Generator').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
  logDebug('doGet template evaluated');
  return result;
}

// HTML テンプレートから部分テンプレートを読み込むための include ヘルパー
function include(filename) {
  const normalized = String(filename).replace(/[\\/]+/g, '_');
  return HtmlService.createHtmlOutputFromFile(normalized).getContent();
}

function saveSettings(settings) {
  try {
    logDebug('saveSettings called', {
      providedKeys: Object.keys(settings || {})
    });
    const storableSettings = Object.assign({}, settings);
    storableSettings.showTitleUnderline = String(storableSettings.showTitleUnderline);
    storableSettings.showBottomBar = String(storableSettings.showBottomBar);
    storableSettings.showDateColumn = String(storableSettings.showDateColumn); // 日付カラム設定を追加
    storableSettings.enableGradient = String(storableSettings.enableGradient);
    // UIテーマ設定の永続化（存在する場合）
    if (typeof storableSettings.currentTheme === 'string') {
      storableSettings.currentTheme = String(storableSettings.currentTheme);
    }
    ['uiPrimaryColor','uiSecondaryColor','uiBackgroundColor','uiSurfaceColor','uiTextColor','uiFontFamily','displayFontFamily','bodyFontFamily']
      .forEach(function(k){ if (typeof storableSettings[k] === 'undefined') delete storableSettings[k]; });
    PropertiesService.getUserProperties().setProperties(storableSettings, false);
    logDebug('saveSettings persisted', {
      storeSize: Object.keys(storableSettings).length
    });
    return {
      status: 'success',
      message: '設定を保存しました。'
    };
  } catch (e) {
    Logger.log(`設定の保存エラー: ${e.message}`);
    return {
      status: 'error',
      message: `設定の保存中にエラーが発生しました: ${e.message}`
    };
  }
}

function saveSelectedPreset(presetName) {
  try {
    logDebug('saveSelectedPreset called', { presetName });
    PropertiesService.getUserProperties().setProperty('selectedPreset', presetName);
    return {
      status: 'success',
      message: 'プリセット選択を保存しました。'
    };
  } catch (e) {
    Logger.log(`プリセット保存エラー: ${e.message}`);
    return {
      status: 'error',
      message: `プリセットの保存中にエラーが発生しました: ${e.message}`
    };
  }
}

function loadSettings() {
  const properties = PropertiesService.getUserProperties().getProperties();
  logDebug('loadSettings read properties', { availableKeys: Object.keys(properties) });
  return {
    primaryColor: properties.primaryColor || '#4285F4',
    gradientStart: properties.gradientStart || '#4285F4',
    gradientEnd: properties.gradientEnd || '#ff52df',
    fontFamily: properties.fontFamily || 'Noto Sans JP',
    showTitleUnderline: properties.showTitleUnderline === 'false' ? false : true,
    showBottomBar: properties.showBottomBar === 'false' ? false : true,
    showDateColumn: properties.showDateColumn === 'false' ? false : true, // 日付カラム設定を追加（デフォルトtrue）
    enableGradient: properties.enableGradient === 'true' ? true : false,
    footerText: properties.footerText || '© Google Inc.',
    headerLogoUrl: properties.headerLogoUrl || 'https://upload.wikimedia.org/wikipedia/commons/thumb/8/8a/Google_Gemini_logo.svg/2560px-Google_Gemini_logo.svg.png',
    closingLogoUrl: properties.closingLogoUrl || 'https://upload.wikimedia.org/wikipedia/commons/thumb/8/8a/Google_Gemini_logo.svg/2560px-Google_Gemini_logo.svg.png',
    titleBgUrl: properties.titleBgUrl || '',
    sectionBgUrl: properties.sectionBgUrl || '',
    mainBgUrl: properties.mainBgUrl || '',
    closingBgUrl: properties.closingBgUrl || '',
    driveFolderUrl: properties.driveFolderUrl || '',
    selectedPreset: properties.selectedPreset || 'default',
    // UIテーマ情報
    currentTheme: properties.currentTheme || (typeof getCurrentTheme === 'function' ? getCurrentTheme() : 'light'),
    uiPrimaryColor: properties.uiPrimaryColor || (CONFIG && CONFIG.THEMES && CONFIG.THEMES.light ? CONFIG.THEMES.light.primary : '#6750A4'),
    uiSecondaryColor: properties.uiSecondaryColor || (CONFIG && CONFIG.THEMES && CONFIG.THEMES.light ? CONFIG.THEMES.light.secondary : '#625B71'),
    uiBackgroundColor: properties.uiBackgroundColor || (CONFIG && CONFIG.THEMES && CONFIG.THEMES.light ? CONFIG.THEMES.light.background : '#FDF8FF'),
    uiSurfaceColor: properties.uiSurfaceColor || (CONFIG && CONFIG.THEMES && CONFIG.THEMES.light ? CONFIG.THEMES.light.surface : '#FEF7FF'),
    uiTextColor: properties.uiTextColor || (CONFIG && CONFIG.THEMES && CONFIG.THEMES.light ? CONFIG.THEMES.light.text : '#1C1B1F'),
    uiFontFamily: properties.uiFontFamily || (CONFIG && CONFIG.FONTS ? CONFIG.FONTS.ui_family : "Roboto, 'Noto Sans JP', Arial, sans-serif"),
    displayFontFamily: properties.displayFontFamily || (CONFIG && CONFIG.FONTS ? CONFIG.FONTS.display_family : 'Product Sans'),
    bodyFontFamily: properties.bodyFontFamily || (CONFIG && CONFIG.FONTS ? CONFIG.FONTS.body_family : 'Noto Sans JP')
  };
}

// ========================================
// 3. スライド生成メイン処理
// ========================================

/** settingsオブジェクトに基づき、CONFIG内の動的カラーを更新します */
function updateDynamicColors(settings) {
  logDebug('updateDynamicColors invoked', {
    primaryColor: settings.primaryColor,
    currentTheme: settings && settings.currentTheme
  });
  const primary = settings.primaryColor;
  const isDark = (settings && settings.currentTheme === 'dark');
  // UIテーマの同期（存在すれば）
  try {
    if (settings.uiPrimaryColor) CONFIG.COLORS.ui_primary_color = settings.uiPrimaryColor;
    if (settings.uiSecondaryColor) CONFIG.COLORS.ui_secondary_color = settings.uiSecondaryColor;
    if (settings.uiBackgroundColor) CONFIG.COLORS.ui_background_color = settings.uiBackgroundColor;
    if (settings.uiSurfaceColor) CONFIG.COLORS.ui_surface_color = settings.uiSurfaceColor;
    if (settings.uiTextColor) CONFIG.COLORS.ui_text_color = settings.uiTextColor;
    if (typeof setCurrentTheme === 'function' && settings.currentTheme) {
      setCurrentTheme(settings.currentTheme);
    }
  } catch (e) {
    logDebug('updateDynamicColors UI sync failed', { error: e.message });
  }
  // フォントの同期
  try {
    if (CONFIG && CONFIG.FONTS) {
      if (settings.uiFontFamily) CONFIG.FONTS.ui_family = settings.uiFontFamily;
      if (settings.displayFontFamily) CONFIG.FONTS.display_family = settings.displayFontFamily;
      if (settings.bodyFontFamily) CONFIG.FONTS.body_family = settings.bodyFontFamily;
      if (settings.fontFamily) CONFIG.FONTS.family = settings.fontFamily;
    }
  } catch (e) {}

  // テーマに応じてティントグレーを調整
  if (!primary) return;
  if (isDark) {
    CONFIG.COLORS.background_gray = generateTintedGray(primary, 12, 18);
    CONFIG.COLORS.faint_gray = generateTintedGray(primary, 12, 22);
    CONFIG.COLORS.ghost_gray = generateTintedGray(primary, 25, 28);
    CONFIG.COLORS.table_header_bg = generateTintedGray(primary, 18, 20);
    CONFIG.COLORS.lane_border = generateTintedGray(primary, 12, 35);
    CONFIG.COLORS.card_border = generateTintedGray(primary, 12, 35);
    CONFIG.COLORS.neutral_gray = generateTintedGray(primary, 6, 68);
  } else {
    CONFIG.COLORS.background_gray = generateTintedGray(primary, 10, 98); 
    CONFIG.COLORS.faint_gray = generateTintedGray(primary, 10, 93);
    CONFIG.COLORS.ghost_gray = generateTintedGray(primary, 38, 88);
    CONFIG.COLORS.table_header_bg = generateTintedGray(primary, 20, 94);
    CONFIG.COLORS.lane_border = generateTintedGray(primary, 15, 85);
    CONFIG.COLORS.card_border = generateTintedGray(primary, 15, 85);
    CONFIG.COLORS.neutral_gray = generateTintedGray(primary, 5, 62);
  }
  CONFIG.COLORS.process_arrow = CONFIG.COLORS.ghost_gray;
}

function generateSlidesFromWebApp(slideDataString, settings) {
  try {
    logDebug('generateSlidesFromWebApp started', {
      payloadBytes: slideDataString ? slideDataString.length : 0,
      hasSettings: Boolean(settings)
    });
    // 共通パーサで厳格に検証・パース
    const slideData = parseSlideDataStrict(slideDataString);
    const result = createPresentation(slideData, settings);
    logDebug('generateSlidesFromWebApp completed', {
      slideCount: Array.isArray(slideData) ? slideData.length : 0
    });
    return result;
  } catch (e) {
    // 改行のエスケープを修正（実際の改行を出力）
    Logger.log(`Error: ${e.message}
Stack: ${e.stack}`);
    logDebug('generateSlidesFromWebApp failed', { error: e.message });
    throw new Error(`Server error: ${e.message}`);
  }
}

/**
 * 既存のプレゼンテーションのテーマを新しいテーマに合わせて更新
 * UIから呼ばれるエントリーポイント
 */
function updateExistingSlides(settings) {
  try {
    logDebug('updateExistingSlides invoked', { theme: settings && settings.currentTheme });
    if (settings && settings.currentTheme && typeof setCurrentTheme === 'function') {
      setCurrentTheme(settings.currentTheme);
    }
    // 色とフォントを同期
    updateDynamicColors(settings || {});
    if (typeof refreshPresentationTheme === 'function') {
      return refreshPresentationTheme(settings && settings.currentTheme ? settings.currentTheme : getCurrentTheme());
    }
    // darkMode.js が未ロードの場合のフォールバック
    return { status: 'info', message: 'テーマ刷新モジュールが見つかりませんでした。' };
  } catch (e) {
    logDebug('updateExistingSlides failed', { error: e.message });
    throw e;
  }
}

/**
 * 現在のテーマをトグルして適用
 */
function togglePresentationTheme() {
  const props = PropertiesService.getUserProperties();
  const current = (props.getProperty('currentTheme')) || (typeof getCurrentTheme === 'function' ? getCurrentTheme() : 'light');
  const next = current === 'dark' ? 'light' : 'dark';
  if (typeof setCurrentTheme === 'function') setCurrentTheme(next);
  const result = (typeof refreshPresentationTheme === 'function') ? refreshPresentationTheme(next) : { status: 'info', message: 'テーマ刷新モジュールが見つかりませんでした。' };
  props.setProperty('currentTheme', next);
  return result;
}
function clearLegacyUserProperties() {
  try {
    // ユーザープロパティを全て取得
    const properties = PropertiesService.getUserProperties().getProperties();
    logDebug('clearLegacyUserProperties fetched props', { availableKeys: Object.keys(properties) });
    
    // 削除対象のキー（プリセット機能追加前の設定）
    const legacyKeys = [
      'primaryColor',
      'gradientStart', 
      'gradientEnd',
      'fontFamily',
      'showTitleUnderline',
      'showBottomBar',
      'enableGradient',
      'footerText',
      'headerLogoUrl',
      'closingLogoUrl',
      'titleBgUrl',
      'sectionBgUrl',
      'mainBgUrl',
      'closingBgUrl',
      'driveFolderUrl',
      'driveFolderId'
    ];
    
    // レガシーキーを削除
    const keysToDelete = [];
    legacyKeys.forEach(key => {
      if (properties.hasOwnProperty(key)) {
        keysToDelete.push(key);
      }
    });
    
    if (keysToDelete.length > 0) {
      // 個別にプロパティを削除
      const userProperties = PropertiesService.getUserProperties();
      keysToDelete.forEach(key => {
        userProperties.deleteProperty(key);
      });
      logDebug('clearLegacyUserProperties removed keys', { keys: keysToDelete });
      return {
        status: 'success',
        message: `${keysToDelete.length}個のレガシープロパティを削除しました。`,
        deletedKeys: keysToDelete
      };
    } else {
      return {
        status: 'info',
        message: '削除対象のレガシープロパティは見つかりませんでした。'
      };
    }
    
  } catch (e) {
    Logger.log(`レガシープロパティ削除エラー: ${e.message}`);
    return {
      status: 'error',
      message: `レガシープロパティの削除中にエラーが発生しました: ${e.message}`
    };
  }
}
