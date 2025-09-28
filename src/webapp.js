function doGet(e) {
  const htmlTemplate = HtmlService.createTemplateFromFile('index');
  htmlTemplate.settings = loadSettings();
  return htmlTemplate.evaluate().setTitle('Google Slide Generator').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function saveSettings(settings) {
  try {
    const storableSettings = Object.assign({}, settings);
    storableSettings.showTitleUnderline = String(storableSettings.showTitleUnderline);
    storableSettings.showBottomBar = String(storableSettings.showBottomBar);
    storableSettings.showDateColumn = String(storableSettings.showDateColumn); // 日付カラム設定を追加
    storableSettings.enableGradient = String(storableSettings.enableGradient);
    PropertiesService.getUserProperties().setProperties(storableSettings, false);
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
    selectedPreset: properties.selectedPreset || 'default'
  };
}

// ========================================
// 3. スライド生成メイン処理
// ========================================

/** settingsオブジェクトに基づき、CONFIG内の動的カラーを更新します */
function updateDynamicColors(settings) {
  const primary = settings.primaryColor;
  CONFIG.COLORS.background_gray = generateTintedGray(primary, 10, 98); 
  CONFIG.COLORS.faint_gray = generateTintedGray(primary, 10, 93);
  CONFIG.COLORS.ghost_gray = generateTintedGray(primary, 38, 88);
  CONFIG.COLORS.table_header_bg = generateTintedGray(primary, 20, 94);
  CONFIG.COLORS.lane_border = generateTintedGray(primary, 15, 85);
  CONFIG.COLORS.card_border = generateTintedGray(primary, 15, 85);
  CONFIG.COLORS.neutral_gray = generateTintedGray(primary, 5, 62);
  CONFIG.COLORS.process_arrow = CONFIG.COLORS.ghost_gray;
}

function generateSlidesFromWebApp(slideDataString, settings) {
  try {
    const slideData = JSON.parse(slideDataString);
    return createPresentation(slideData, settings);
  } catch (e) {
    Logger.log(`Error: ${e.message}\nStack: ${e.stack}`);
    throw new Error(`Server error: ${e.message}`);
  }
}
function clearLegacyUserProperties() {
  try {
    // ユーザープロパティを全て取得
    const properties = PropertiesService.getUserProperties().getProperties();
    
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
      Logger.log(`削除されたレガシープロパティ: ${keysToDelete.join(', ')}`);
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