/**
 * Google Slide Generator Web Application
 * 
 * JSONデータからGoogleスライドを自動生成するWebアプリケーション
 * 
 * @author まじん
 * @version 3.0.0
 * @requires Google Apps Script
 * @requires Google Slides API
 * @requires Google Drive API
 * @license CC BY-NC 4.0
 * @see https://creativecommons.org/licenses/by-nc/4.0/
 */

/**
 * ========================================
 * 色彩操作ヘルパー関数
 * ========================================
 */
function hexToRgb(hex) {
  const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result ? {
    r: parseInt(result[1], 16),
    g: parseInt(result[2], 16),
    b: parseInt(result[3], 16)
  } : null;
}

function rgbToHsl(r, g, b) {
  r /= 255;
  g /= 255;
  b /= 255;
  const max = Math.max(r, g, b),
    min = Math.min(r, g, b);
  let h, s, l = (max + min) / 2;
  if (max === min) {
    h = s = 0;
  } else {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    switch (max) {
      case r:
        h = (g - b) / d + (g < b ? 6 : 0);
        break;
      case g:
        h = (b - r) / d + 2;
        break;
      case b:
        h = (r - g) / d + 4;
        break;
    }
    h /= 6;
  }
  return {
    h: h * 360,
    s: s * 100,
    l: l * 100
  };
}

function hslToHex(h, s, l) {
  s /= 100;
  l /= 100;
  const c = (1 - Math.abs(2 * l - 1)) * s,
    x = c * (1 - Math.abs((h / 60) % 2 - 1)),
    m = l - c / 2;
  let r = 0,
    g = 0,
    b = 0;
  if (0 <= h && h < 60) {
    r = c;
    g = x;
    b = 0;
  } else if (60 <= h && h < 120) {
    r = x;
    g = c;
    b = 0;
  } else if (120 <= h && h < 180) {
    r = 0;
    g = c;
    b = x;
  } else if (180 <= h && h < 240) {
    r = 0;
    g = x;
    b = c;
  } else if (240 <= h && h < 300) {
    r = x;
    g = 0;
    b = c;
  } else if (300 <= h && h < 360) {
    r = c;
    g = 0;
    b = x;
  }
  r = Math.round((r + m) * 255);
  g = Math.round((g + m) * 255);
  b = Math.round((b + m) * 255);
  return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
}

function generateTintedGray(tintColorHex, saturation, lightness) {
  const rgb = hexToRgb(tintColorHex);
  if (!rgb) return '#F8F9FA';
  const hsl = rgbToHsl(rgb.r, rgb.g, rgb.b);
  return hslToHex(hsl.h, saturation, lightness);
}

/**
 * ピラミッド用カラーグラデーション生成
 * @param {string} baseColor - ベースとなるプライマリカラー
 * @param {number} levels - レベル数
 * @return {string[]} 上から下へのグラデーションカラー配列
 */
function generatePyramidColors(baseColor, levels) {
  const colors = [];
  for (let i = 0; i < levels; i++) {
    // 上から下に向かって徐々に薄くなる (0% → 60%まで)
    const lightenAmount = (i / Math.max(1, levels - 1)) * 0.6;
    colors.push(lightenColor(baseColor, lightenAmount));
  }
  return colors;
}

/**
 * 色を明るくする関数
 * @param {string} color - 元の色 (#RRGGBB形式)
 * @param {number} amount - 明るくする量 (0.0-1.0)
 * @return {string} 明るくした色
 */
function lightenColor(color, amount) {
  const rgb = hexToRgb(color);
  if (!rgb) return color;
  
  const lighten = (c) => Math.min(255, Math.round(c + (255 - c) * amount));
  const newR = lighten(rgb.r);
  const newG = lighten(rgb.g);
  const newB = lighten(rgb.b);
  
  return `#${newR.toString(16).padStart(2, '0')}${newG.toString(16).padStart(2, '0')}${newB.toString(16).padStart(2, '0')}`;
}

/**
 * 色を暗くする関数
 * @param {string} color - 元の色 (#RRGGBB形式)
 * @param {number} amount - 暗くする量 (0.0-1.0)
 * @return {string} 暗くした色
 */
function darkenColor(color, amount) {
  const rgb = hexToRgb(color);
  if (!rgb) return color;
  
  const darken = (c) => Math.max(0, Math.round(c * (1 - amount)));
  const newR = darken(rgb.r);
  const newG = darken(rgb.g);
  const newB = darken(rgb.b);
  
  return `#${newR.toString(16).padStart(2, '0')}${newG.toString(16).padStart(2, '0')}${newB.toString(16).padStart(2, '0')}`;
}

/**
 * StepUp用カラーグラデーション生成（左から右に濃くなる）
 * @param {string} baseColor - ベースとなるプライマリカラー
 * @param {number} steps - ステップ数
 * @return {string[]} 左から右へのグラデーションカラー配列（薄い→濃い）
 */
function generateStepUpColors(baseColor, steps) {
  const colors = [];
  for (let i = 0; i < steps; i++) {
    // 左から右に向かって徐々に濃くなる (60% → 0%)
    const lightenAmount = 0.6 * (1 - (i / Math.max(1, steps - 1)));
    colors.push(lightenColor(baseColor, lightenAmount));
  }
  return colors;
}

/**
 * Process用カラーグラデーション生成（上から下に濃くなる）
 * @param {string} baseColor - ベースとなるプライマリカラー
 * @param {number} steps - ステップ数
 * @return {string[]} 上から下へのグラデーションカラー配列（薄い→濃い）
 */
function generateProcessColors(baseColor, steps) {
  const colors = [];
  for (let i = 0; i < steps; i++) {
    // 上から下に向かって徐々に濃くなる (50% → 0%)
    const lightenAmount = 0.5 * (1 - (i / Math.max(1, steps - 1)));
    colors.push(lightenColor(baseColor, lightenAmount));
  }
  return colors;
}

/**
 * Timeline用カードグラデーション生成（左から右に濃くなる）
 * @param {string} baseColor - ベースとなるプライマリカラー
 * @param {number} milestones - マイルストーン数
 * @return {string[]} 左から右へのグラデーションカラー配列（薄い→濃い）
 */
function generateTimelineCardColors(baseColor, milestones) {
  const colors = [];
  for (let i = 0; i < milestones; i++) {
    // 左から右に向かって徐々に濃くなる (40% → 0%)
    const lightenAmount = 0.4 * (1 - (i / Math.max(1, milestones - 1)));
    colors.push(lightenColor(baseColor, lightenAmount));
  }
  return colors;
}

/**
 * Compare系用の左右対比色生成
 * @param {string} baseColor - ベースとなるプライマリカラー
 * @return {Object} {left: 濃い色, right: 元の色}の組み合わせ
 */
function generateCompareColors(baseColor) {
  return {
    left: darkenColor(baseColor, 0.3),   // 左側：30%暗く（Before/導入前）- 視認性向上
    right: baseColor                     // 右側：元の色（After/導入後）
  };
}

// ========================================
// 1. マスターデザイン設定
// ========================================
const CONFIG = {
  BASE_PX: {
    W: 960,
    H: 540
  },
  BACKGROUND_IMAGES: {
    title: '',
    closing: '',
    section: '',
    main: ''
  },
  POS_PX: {
    titleSlide: {
      logo: {
        left: 55,
        top: 60,    // 105 → 60 に変更（45px上に移動）
        width: 135
      },
      title: {
        left: 50,
        top: 200,
        width: 830,
        height: 90
      },
      date: {
        left: 50,
        top: 450,
        width: 250,
        height: 40
      }
    },
    contentSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      body: {
        left: 25,
        top: 132,
        width: 910,
        height: 330
      },
      twoColLeft: {
        left: 25,
        top: 132,
        width: 440,
        height: 330
      },
      twoColRight: {
        left: 495,
        top: 132,
        width: 440,
        height: 330
      }
    },
    compareSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      leftBox: {
        left: 25,
        top: 112,
        width: 445,
        height: 350
      },
      rightBox: {
        left: 490,
        top: 112,
        width: 445,
        height: 350
      }
    },
    processSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      area: {
        left: 25,
        top: 132,
        width: 910,
        height: 330
      }
    },
    timelineSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      area: {
        left: 25,
        top: 132,
        width: 910,
        height: 330
      }
    },
    diagramSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      lanesArea: {
        left: 25,
        top: 132,
        width: 910,
        height: 330
      }
    },
    cardsSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      gridArea: {
        left: 25,
        top: 120,
        width: 910,
        height: 340
      }
    },
    tableSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      area: {
        left: 25,
        top: 130,
        width: 910,
        height: 330
      }
    },
    progressSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      area: {
        left: 25,
        top: 132,
        width: 910,
        height: 330
      }
    },
    quoteSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 88,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 100,
        width: 910,
        height: 40
      }
    },
    kpiSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      gridArea: {
        left: 25,
        top: 132,
        width: 910,
        height: 330
      }
    },
    triangleSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      area: {
        left: 25,
        top: 110,
        width: 910,
        height: 350
      }
    },
    flowChartSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      singleRow: {
        left: 25,
        top: 160,
        width: 910,
        height: 180
      },
      upperRow: {
        left: 25,
        top: 150,
        width: 910,
        height: 120
      },
      lowerRow: {
        left: 25,
        top: 290,
        width: 910,
        height: 120
      }
    },
    stepUpSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      stepArea: {
        left: 25,
        top: 130,
        width: 910,
        height: 330
      }
    },
    imageTextSlide: {
      headerLogo: {
        right: 20,
        top: 20,
        width: 75
      },
      title: {
        left: 25,
        top: 20,
        width: 830,
        height: 65
      },
      titleUnderline: {
        left: 25,
        top: 80,
        width: 260,
        height: 4
      },
      subhead: {
        left: 25,
        top: 90,
        width: 910,
        height: 40
      },
      leftImage: {
        left: 25,
        top: 150,
        width: 440,
        height: 270  // キャプション分減算
      },
      leftImageCaption: {
        left: 25,
        top: 430,
        width: 440,
        height: 30
      },
      rightText: {
        left: 485,
        top: 150,
        width: 450,
        height: 310
      },
      leftText: {
        left: 25,
        top: 150,
        width: 450,
        height: 310
      },
      rightImage: {
        left: 495,
        top: 150,
        width: 440,
        height: 270  // キャプション分減算
      },
      rightImageCaption: {
        left: 495,
        top: 430,
        width: 440,
        height: 30
      }
    },
      pyramidSlide: {
        headerLogo: {
          right: 20,
          top: 20,
          width: 75
        },
        title: {
          left: 25,
          top: 20,
          width: 830,
          height: 65
        },
        titleUnderline: {
          left: 25,
          top: 88,
          width: 260,
          height: 4
        },
        subhead: {
          left: 25,
          top: 100,
          width: 910,
          height: 40
        },
        pyramidArea: {
          left: 25,
          top: 120,
          width: 910,
          height: 360
        }
      },
    sectionSlide: {
      title: {
        left: 55,
        top: 230,
        width: 840,
        height: 80
      },
      ghostNum: {
        left: 35,
        top: 120,
        width: 400,
        height: 200
      }
    },
    footer: {
      leftText: {
        left: 15,
        top: 505,
        width: 250,
        height: 20
      },
      rightPage: {
        right: 15,
        top: 505,
        width: 50,
        height: 20
      }
    },
    bottomBar: {
      left: 0,
      top: 534,
      width: 960,
      height: 6
    }
  },
  FONTS: {
    family: 'Noto Sans JP',
    sizes: {
      title: 40,
      date: 16,
      sectionTitle: 38,
      contentTitle: 24,
      subhead: 16,
      body: 14,
      footer: 9,
      chip: 11,
      laneTitle: 13,
      small: 10,
      processStep: 14,
      axis: 12,
      ghostNum: 180
    }
  },
  COLORS: {
    primary_color: '#4285F4',
    text_primary: '#333333',
    background_white: '#FFFFFF',
    card_bg: '#f6e9f0',
    background_gray: '',
    faint_gray: '',
    ghost_gray: '',
    table_header_bg: '',
    lane_border: '',
    card_border: '',
    neutral_gray: '',
    process_arrow: ''
  },
  DIAGRAM: {
    laneGap_px: 24,
    lanePad_px: 10,
    laneTitle_h_px: 30,
    cardGap_px: 12,
    cardMin_h_px: 48,
    cardMax_h_px: 70,
    arrow_h_px: 10,
    arrowGap_px: 8
  },
  LOGOS: {
    header: '',
    closing: ''
  },
  FOOTER_TEXT: `© ${new Date().getFullYear()} Your Company`
};

// ========================================
// 2. Webアプリケーションのメイン関数
// ========================================

function doGet(e) {
  const htmlTemplate = HtmlService.createTemplateFromFile('index.html');
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

let __SECTION_COUNTER = 0;
let __SLIDE_DATA_FOR_AGENDA = [];

function createPresentation(slideData, settings) {
  updateDynamicColors(settings);
  CONFIG.COLORS.primary_color = settings.primaryColor || CONFIG.COLORS.primary_color;
  CONFIG.FOOTER_TEXT = settings.footerText;
  CONFIG.FONTS.family = settings.fontFamily || CONFIG.FONTS.family;
  CONFIG.LOGOS.header = settings.headerLogoUrl;
  CONFIG.LOGOS.closing = settings.closingLogoUrl;
  CONFIG.BACKGROUND_IMAGES.title = settings.titleBgUrl;
  CONFIG.BACKGROUND_IMAGES.closing = settings.closingBgUrl;
  CONFIG.BACKGROUND_IMAGES.section = settings.sectionBgUrl;
  CONFIG.BACKGROUND_IMAGES.main = settings.mainBgUrl;

  __SLIDE_DATA_FOR_AGENDA = slideData;

  // ファイル名の生成（日付カラムの設定に応じて日付を付与）
  const rawTitle = (slideData[0] && slideData[0].type === 'title' ? String(slideData[0].title || '') : 'Google Slide Generator Presentation');
  const singleLineTitle = rawTitle.replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim();
  
  let finalName;
  if (settings.showDateColumn) {
    // 日付カラムがオンの場合：ファイル名に日付を付与
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy.MM.dd');
    finalName = singleLineTitle ? (singleLineTitle + ' ' + dateStr) : ('Google Slide Generator Presentation ' + dateStr);
  } else {
    // 日付カラムがオフの場合：ファイル名に日付を付与しない
    finalName = singleLineTitle || 'Google Slide Generator Presentation';
  }
  const presentation = SlidesApp.create(finalName);
  presentation.getSlides()[0].remove();

  if (settings.driveFolderId && settings.driveFolderId.trim()) {
    try {
      DriveApp.getFileById(presentation.getId()).moveTo(DriveApp.getFolderById(settings.driveFolderId.trim()));
    } catch (e) {
      Logger.log(`フォルダ移動エラー: ${e.message}`);
    }
  }

  __SECTION_COUNTER = 0;
  const layout = createLayoutManager(presentation.getPageWidth(), presentation.getPageHeight());
  let pageCounter = 0;

  for (const data of slideData) {
    try {
      const generator = slideGenerators[data.type];
      if (data.type !== 'title' && data.type !== 'closing') {
        pageCounter++;
      }
      if (generator) {
        const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
        generator(slide, data, layout, pageCounter, settings);
        
        // スピーカーノートのクリーニング処理
        if (data.notes) {
          const cleanedNotes = cleanSpeakerNotes(data.notes);
          slide.getNotesPage().getSpeakerNotesShape().getText().setText(cleanedNotes);
        }
      }
    } catch (e) {
      Logger.log(`Slide generation skipped: Type: ${data.type}, Error: ${e.message}`);
    }
  }
  return presentation.getUrl();
}

// ========================================
// 4. スライドジェネレーター定義
// ========================================
const slideGenerators = {
  title: createTitleSlide,
  section: createSectionSlide,
  content: createContentSlide,
  agenda: createAgendaSlide,
  compare: createCompareSlide,
  process: createProcessSlide,
  processList: createProcessListSlide,
  timeline: createTimelineSlide,
  diagram: createDiagramSlide,
  cycle: createCycleSlide,
  cards: createCardsSlide,
  headerCards: createHeaderCardsSlide,
  table: createTableSlide,
  progress: createProgressSlide,
  quote: createQuoteSlide,
  kpi: createKpiSlide,
  closing: createClosingSlide,
  bulletCards: createBulletCardsSlide,
  faq: createFaqSlide,
  statsCompare: createStatsCompareSlide,
  barCompare: createBarCompareSlide,
  triangle: createTriangleSlide,
  pyramid: createPyramidSlide,
  flowChart: createFlowChartSlide,
  stepUp: createStepUpSlide,
  imageText: createImageTextSlide
};

// ========================================
// 5. レイアウト管理システム
// ========================================
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

// ユーティリティ（上揃え）- この関数は不要になったため削除しても問題ありません
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

/**
 * 単一画像を固定フレーム内に配置（キャプション対応）
 * @param {Slide} slide - スライドオブジェクト
 * @param {Object} layout - レイアウトマネージャー
 * @param {Object} area - 画像配置エリア
 * @param {string} imageUrl - 画像URL
 * @param {string} caption - キャプション（任意）
 * @param {string} position - 画像位置（'left'または'right'）
 */
function renderSingleImageInArea(slide, layout, area, imageUrl, caption = '', position = 'left') {
  if (!imageUrl) return;
  
  try {
    const imageData = insertImageFromUrlOrFileId(imageUrl);
    if (!imageData) return null;
    
    const img = slide.insertImage(imageData);
    
    // 固定フレーム内にフィット（アスペクト比維持）
    const scale = Math.min(area.width / img.getWidth(), area.height / img.getHeight());
    const w = img.getWidth() * scale;
    const h = img.getHeight() * scale;
    
    // エリア中央に配置
    img.setWidth(w).setHeight(h)
       .setLeft(area.left + (area.width - w) / 2)
       .setTop(area.top + (area.height - h) / 2);
    
    // キャプション追加（画像の実際のサイズに応じて動的に配置）
    if (caption && caption.trim()) {
      // 画像の実際の位置とサイズを取得
      const imageBottom = area.top + (area.height - h) / 2 + h;
      const captionMargin = layout.pxToPt(8); // キャプションと画像の間隔
      const captionHeight = layout.pxToPt(30);
      
      // キャプションを画像の下に配置
      const captionTop = imageBottom + captionMargin;
      const captionShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
        area.left, captionTop, area.width, captionHeight);
      captionShape.getFill().setTransparent();
      captionShape.getBorder().setTransparent();
      setStyledText(captionShape, caption.trim(), {
        size: CONFIG.FONTS.sizes.small,
        color: CONFIG.COLORS.neutral_gray,
        align: SlidesApp.ParagraphAlignment.CENTER
      });
    }
       
    return img;
  } catch (e) {
    Logger.log(`Image insertion failed: ${e.message}. URL: ${imageUrl}`);
    return null;
  }
}

// ========================================
// 7. ユーティリティ関数群
// ========================================

function estimateTextWidthPt(text, fontSizePt) {
  const multipliers = {
    ascii: 0.62,
    japanese: 1.0,
    other: 0.85
  };
  return String(text || '').split('').reduce((acc, char) => {
    if (char.match(/[ -~]/)) {
      return acc + multipliers.ascii;
    } else if (char.match(/[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF]/)) {
      return acc + multipliers.japanese;
    } else {
      return acc + multipliers.other;
    }
  }, 0) * fontSizePt;
}

function offsetRect(rect, dx, dy) {
  return {
    left: rect.left + (dx || 0),
    top: rect.top + (dy || 0),
    width: rect.width,
    height: rect.height
  };
}

function drawStandardTitleHeader(slide, layout, key, title, settings) {
  const logoRect = safeGetRect(layout, `${key}.headerLogo`);
  
  try {
    if (CONFIG.LOGOS.header && logoRect) {
      const imageData = insertImageFromUrlOrFileId(CONFIG.LOGOS.header);
      if (imageData) {
        const logo = slide.insertImage(imageData);
        const asp = logo.getHeight() / logo.getWidth();
        logo.setLeft(logoRect.left).setTop(logoRect.top).setWidth(logoRect.width).setHeight(logoRect.width * asp);
      }
    }
  } catch (e) {
    Logger.log(`Header logo error: ${e.message}`);
  }
  const titleRect = safeGetRect(layout, `${key}.title`);
  if (!titleRect) {
    Logger.log(`[rect-missing] key=${key}.title`);
    return;
  }
  
  const fontSize = CONFIG.FONTS.sizes.contentTitle;
  const optimalHeight = layout.pxToPt(fontSize + 8); // フォントサイズ + 8px余白
  
  const titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 
    titleRect.left, 
    titleRect.top, 
    titleRect.width, 
    optimalHeight
  );
  setStyledText(titleShape, title || '', {
    size: fontSize,
    bold: true
  });
  
  // 上揃えで元の位置を維持
  try {
    titleShape.setContentAlignment(SlidesApp.ContentAlignment.TOP);
  } catch (e) {}
  
  if (settings.showTitleUnderline && title) {
    const uRect = safeGetRect(layout, `${key}.titleUnderline`);
    if (!uRect) {
      Logger.log(`[rect-missing] key=${key}.titleUnderline`);
      return;
    }
    const estimatedWidthPt = estimateTextWidthPt(title, fontSize);
    
    // アンダーライン幅の上限制限（スライド枠内に収める）
    const maxUnderlineWidth = layout.pageW_pt - uRect.left - layout.pxToPt(25); // 右余白25px確保
    const finalWidth = Math.min(estimatedWidthPt, maxUnderlineWidth);
    
    applyFill(slide, uRect.left, uRect.top, finalWidth, uRect.height, settings);
  }
}

/**
 * テキスト高さ概算（pt）。日本語を含む一般文で「1em ≒ fontSizePt」前提の簡易推定。
 * 既存コードで px を使っていた場合のズレもここで吸収。
 */
function estimateTextHeightPt(text, widthPt, fontSizePt, lineHeight) {
  var paragraphs = String(text).split(/\r?\n/);
  // 1行に入る概算文字数：幅 / (0.95em) として保守的に
  var charsPerLine = Math.max(1, Math.floor(widthPt / (fontSizePt * 0.95)));
  var lines = 0;
  for (var i = 0; i < paragraphs.length; i++) {
    var s = paragraphs[i].replace(/\s+/g, ' ').trim();
    // 空行も最低1行とみなす
    var len = s.length || 1;
    lines += Math.ceil(len / charsPerLine);
  }
  var lineH = fontSizePt * (lineHeight || 1.2);
  return Math.max(lineH, lines * lineH);
}

/**
 * Subhead を描画して、次要素の基準Yを返すユーティリティ。
 * 互換用に dy も返す（reserved 分との差分）。全単位 pt。
 *
 * @param {SlidesApp.Slide} slide
 * @param {{x:number,y:number,width:number}} frame   // コンテンツ基準枠（pt）
 * @param {string} subheadText
 * @param {{
 *   fontFamily?: string,
 *   fontSizePt?: number,     // 例: 18
 *   lineHeight?: number,     // 例: 1.25 （ParagraphStyle.setLineSpacing に %換算で反映）
 *   color?: string,          // 例: '#444444'
 *   bold?: boolean,
 *   marginAfterPt?: number,  // subhead 下の外余白（例: 12）
 *   reservedPt?: number,     // 既存レイアウトが想定していた予約高さ（例: 22）
 * }} [style]
 * @param {string} [debugKey] // ログ表示用キー
 * @return {{dy:number,nextY:number,height:number,shape:SlidesApp.Shape|null}}
 */
function drawSubheadIfAny(slide, layout, key, subhead) {
  if (!subhead) return 0;
  
  const rect = safeGetRect(layout, `${key}.subhead`);
  if (!rect) {
    Logger.log(`[drawSubheadIfAny] Could not get rect for ${key}.subhead`);
    return 0;
  }
  
  const box = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rect.left, rect.top, rect.width, rect.height);
  setStyledText(box, subhead, {
    size: CONFIG.FONTS.sizes.subhead,
    color: CONFIG.COLORS.text_primary
  });
  
  return layout.pxToPt(36); // 固定値を返す
}


function drawBottomBar(slide, layout, settings) {
  const barRect = layout.getRect('bottomBar');
  applyFill(slide, barRect.left, barRect.top, barRect.width, barRect.height, settings);
}

function addCucFooter(slide, layout, pageNum) {
  const leftRect = layout.getRect('footer.leftText');
  const leftShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, leftRect.left, leftRect.top, leftRect.width, leftRect.height);
  leftShape.getText().setText(CONFIG.FOOTER_TEXT);
  applyTextStyle(leftShape.getText(), {
    size: CONFIG.FONTS.sizes.footer,
    color: CONFIG.COLORS.text_primary
  });
  if (pageNum > 0) {
    const rightRect = layout.getRect('footer.rightPage');
    const rightShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rightRect.left, rightRect.top, rightRect.width, rightRect.height);
    rightShape.getText().setText(String(pageNum));
    applyTextStyle(rightShape.getText(), {
      size: CONFIG.FONTS.sizes.footer,
      color: CONFIG.COLORS.primary_color,
      align: SlidesApp.ParagraphAlignment.END
    });
  }
}

function drawBottomBarAndFooter(slide, layout, pageNum, settings) {
  if (settings.showBottomBar) {
    drawBottomBar(slide, layout, settings);
  }
  addCucFooter(slide, layout, pageNum);
}

// ========================================
// 8. テキストスタイリング関数群
// ========================================

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

function drawCompareBox(slide, layout, rect, title, items, settings, isLeft = false) {
  const box = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, rect.left, rect.top, rect.width, rect.height);
  box.getFill().setSolidFill(CONFIG.COLORS.background_gray);
  box.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.lane_border);
  const th = layout.pxToPt(40);
  const titleBarBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, rect.left, rect.top, rect.width, th);
  
  // 左右対比色の適用
  const compareColors = generateCompareColors(settings.primaryColor);
  const headerColor = isLeft ? compareColors.left : compareColors.right;
  titleBarBg.getFill().setSolidFill(headerColor);
  titleBarBg.getBorder().setTransparent();
  const titleTextShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rect.left, rect.top, rect.width, th);
  titleTextShape.getFill().setTransparent();
  titleTextShape.getBorder().setTransparent();
  setStyledText(titleTextShape, title, {
    size: CONFIG.FONTS.sizes.laneTitle,
    bold: true,
    color: CONFIG.COLORS.background_white,
    align: SlidesApp.ParagraphAlignment.CENTER
  });
  const pad = layout.pxToPt(12);
  const body = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, rect.left + pad, rect.top + th + pad, rect.width - pad * 2, rect.height - th - pad * 2);
  setBulletsWithInlineStyles(body, items);
}

/**
 * [修正4] レーン図の矢印をカード間を結ぶ線（コネクタ）に変更
 */
function drawArrowBetweenRects(slide, a, b, arrowH, arrowGap, settings) {
  const fromX = a.left + a.width;
  const fromY = a.top + a.height / 2;
  const toX = b.left;
  const toY = b.top + b.height / 2;

  // 描画するスペースがある場合のみ線を描画
  if (toX - fromX <= 0) return;

  const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, fromX, fromY, toX, toY);
  line.getLineFill().setSolidFill(settings.primaryColor);
  line.setWeight(1.5);
  line.setEndArrow(SlidesApp.ArrowStyle.FILL_ARROW);
}


function drawNumberedItems(slide, layout, area, items, settings) {
  // アジェンダ用の座布団を作成
  createContentCushion(slide, area, settings, layout);
  
  const n = Math.max(1, items.length);
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

    // 元の箇条書きテキストから先頭の数字を除去
    let cleanText = String(items[i] || '');
    cleanText = cleanText.replace(/^\s*\d+[\.\s]*/, '');

    const txt = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, cx + layout.pxToPt(28), cy - layout.pxToPt(16), area.width - layout.pxToPt(70), layout.pxToPt(32));
    setStyledText(txt, cleanText, { size: CONFIG.FONTS.sizes.processStep });
    try { txt.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); } catch(e){}
  }
}

// ========================================
// 9. ヘルパー関数群
// ========================================

/**
 * 安全にレイアウト矩形を取得するヘルパー関数
 * @param {Object} layout - レイアウトマネージャー
 * @param {string} path - レイアウトパス
 * @return {Object|null} レイアウト矩形またはnull
 */
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
}

/**
 * 小見出しの下で「本文開始位置」として使える候補を順に探す
 * @param {Object} layout - レイアウトマネージャー
 * @param {string} key - スライドキー
 * @return {Object|null} コンテンツ矩形またはnull
 */
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
}

function adjustColorBrightness(hex, factor) {
  const c = hex.replace('#', '');
  const rgb = parseInt(c, 16);
  let r = (rgb >> 16) & 0xff,
    g = (rgb >> 8) & 0xff,
    b = (rgb >> 0) & 0xff;
  r = Math.min(255, Math.round(r * factor));
  g = Math.min(255, Math.round(g * factor));
  b = Math.min(255, Math.round(b * factor));
  return '#' + (0x1000000 + (r << 16) + (g << 8) + b).toString(16).slice(1);
}

function setMainSlideBackground(slide, layout) {
  setBackgroundImageFromUrl(slide, layout, CONFIG.BACKGROUND_IMAGES.main, CONFIG.COLORS.background_white);
}

function setBackgroundImageFromUrl(slide, layout, imageUrl, fallbackColor) {
  slide.getBackground().setSolidFill(fallbackColor);
  if (!imageUrl) return;
  try {
    const image = insertImageFromUrlOrFileId(imageUrl);
    if (!image) return;
    
    slide.insertImage(image).setWidth(layout.pageW_pt).setHeight(layout.pageH_pt).setLeft(0).setTop(0).sendToBack();
  } catch (e) {
    Logger.log(`Background image failed: ${e.message}. URL: ${imageUrl}`);
  }
}

/**
 * URLまたはGoogle Drive FileIDから画像を取得
 * @param {string} urlOrFileId - 画像URLまたはGoogle Drive FileID
 * @return {Blob|string} 画像データまたはURL
 */
function insertImageFromUrlOrFileId(urlOrFileId) {
  if (!urlOrFileId) return null;
  
  // URLからFileIDを抽出する関数
  function extractFileIdFromUrl(url) {
    const patterns = [
      /\/file\/d\/([a-zA-Z0-9_-]+)/,
      /id=([a-zA-Z0-9_-]+).*file/,
      /file\/([a-zA-Z0-9_-]+)/
    ];
    for (const pattern of patterns) {
      const match = url.match(pattern);
      if (match && match[1]) return match[1];
    }
    return null;
  }
  
  // FileIDの形式かチェック（GoogleドライブのFileIDは通常28-33文字の英数字）
  const fileIdPattern = /^[a-zA-Z0-9_-]{25,}$/;
  
  // URLからFileIDを抽出
  const extractedFileId = extractFileIdFromUrl(urlOrFileId);
  
  if (extractedFileId && fileIdPattern.test(extractedFileId)) {
    // Google Drive FileIDとして処理
    try {
      const file = DriveApp.getFileById(extractedFileId);
      return file.getBlob();
    } catch (e) {
      Logger.log(`Drive file access failed: ${e.message}. FileID: ${extractedFileId}`);
      return null;
    }
  } else if (fileIdPattern.test(urlOrFileId)) {
    // 直接FileIDとして処理
    try {
      const file = DriveApp.getFileById(urlOrFileId);
      return file.getBlob();
    } catch (e) {
      Logger.log(`Drive file access failed: ${e.message}. FileID: ${urlOrFileId}`);
      return null;
    }
  } else {
    // URLとして処理
    return urlOrFileId;
  }
}

function normalizeImages(arr) {
  return (arr || []).map(v => typeof v === 'string' ? {
    url: v
  } : (v && v.url ? v : null)).filter(Boolean).slice(0, 6);
}

function renderImagesInArea(slide, layout, area, images) {
  if (!images || !images.length) return;
  const n = Math.min(6, images.length);
  let cols = n === 1 ? 1 : (n <= 4 ? 2 : 3);
  const rows = Math.ceil(n / cols);
  const gap = layout.pxToPt(10);
  const cellW = (area.width - gap * (cols - 1)) / cols,
    cellH = (area.height - gap * (rows - 1)) / rows;
  for (let i = 0; i < n; i++) {
    const r = Math.floor(i / cols),
      c = i % cols;
    try {
      const img = slide.insertImage(images[i].url);
      const scale = Math.min(cellW / img.getWidth(), cellH / img.getHeight());
      const w = img.getWidth() * scale,
        h = img.getHeight() * scale;
      img.setWidth(w).setHeight(h).setLeft(area.left + c * (cellW + gap) + (cellW - w) / 2).setTop(area.top + r * (cellH + gap) + (cellH - h) / 2);
    } catch (e) {}
  }
}

function createGradientRectangle(slide, x, y, width, height, colors) {
  const numStrips = Math.max(20, Math.floor(width / 2));
  const stripWidth = width / numStrips;
  const startColor = hexToRgb(colors[0]),
    endColor = hexToRgb(colors[1]);
  if (!startColor || !endColor) return null;
  const shapes = [];
  for (let i = 0; i < numStrips; i++) {
    const ratio = i / (numStrips - 1);
    const r = Math.round(startColor.r + (endColor.r - startColor.r) * ratio);
    const g = Math.round(startColor.g + (endColor.g - startColor.g) * ratio);
    const b = Math.round(startColor.b + (endColor.b - startColor.b) * ratio);
    const strip = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x + (i * stripWidth), y, stripWidth + 0.5, height);
    strip.getFill().setSolidFill(r, g, b);
    strip.getBorder().setTransparent();
    shapes.push(strip);
  }
  if (shapes.length > 1) {
    return slide.group(shapes);
  }
  return shapes[0] || null;
}

function applyFill(slide, x, y, width, height, settings) {
  if (settings.enableGradient) {
    createGradientRectangle(slide, x, y, width, height, [settings.gradientStart, settings.gradientEnd]);
  } else {
    const shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, width, height);
    shape.getFill().setSolidFill(settings.primaryColor);
    shape.getBorder().setTransparent();
  }
}

/**
 * プリセット機能追加前のレガシーユーザープロパティを削除する関数
 */
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
