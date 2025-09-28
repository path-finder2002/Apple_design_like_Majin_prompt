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
    // UI向けフォントと役割別のファミリー（後方互換を維持してfamilyを優先使用）
    ui_family: "Roboto, 'Noto Sans JP', 'Helvetica Neue', Arial, sans-serif",
    display_family: 'Product Sans',
    body_family: 'Noto Sans JP',
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
    // 既存互換
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
    process_arrow: '',
    success_green: '#1e8e3e',
    error_red: '#d93025',
    // UIテーマと同期するための拡張
    ui_primary_color: '#6750A4',
    ui_secondary_color: '#625B71',
    ui_background_color: '#FDF8FF',
    ui_surface_color: '#FEF7FF',
    ui_text_color: '#1C1B1F'
  },
  // 現在のテーマ状態
  CURRENT_THEME: 'light',
  // Material Design 3 準拠のテーマ定義 + Googleブランドカラー
  THEMES: {
    light: {
      primary: '#6750A4',
      onPrimary: '#FFFFFF',
      secondary: '#625B71',
      onSecondary: '#FFFFFF',
      background: '#FDF8FF',
      surface: '#FEF7FF',
      text: '#1C1B1F',
      outline: '#79747E',
      accent: '#03DAC5',
      google: { blue: '#4285F4', red: '#DB4437', yellow: '#F4B400', green: '#0F9D58' }
    },
    dark: {
      primary: '#D0BCFF',
      onPrimary: '#381E72',
      secondary: '#CCC2DC',
      onSecondary: '#332D41',
      background: '#101014',
      surface: '#141218',
      text: '#E6E0E9',
      outline: '#938F99',
      accent: '#66FFF9',
      google: { blue: '#8AB4F8', red: '#F28B82', yellow: '#FDD663', green: '#81C995' }
    }
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

/**
 * 現在のテーマを取得（'light' | 'dark'）
 */
function getCurrentTheme() {
  try {
    return CONFIG.CURRENT_THEME || 'light';
  } catch (e) {
    return 'light';
  }
}

/**
 * テーマを設定し、CONFIG.COLORSのUI連携プロパティを更新
 * @param {('light'|'dark')} theme
 */
function setCurrentTheme(theme) {
  const next = (theme === 'dark') ? 'dark' : 'light';
  CONFIG.CURRENT_THEME = next;
  syncUIThemeWithConfig();
  return CONFIG.CURRENT_THEME;
}

/**
 * CONFIG.THEMES と CONFIG.COLORS, CONFIG.FONTS を同期
 * - UI用カラー（ui_*）をテーマから反映
 * - 後方互換のため primary_color は維持（必要に応じて同期）
 */
function syncUIThemeWithConfig() {
  const themeKey = getCurrentTheme();
  const theme = (CONFIG.THEMES && CONFIG.THEMES[themeKey]) ? CONFIG.THEMES[themeKey] : CONFIG.THEMES.light;
  // UI Colors
  CONFIG.COLORS.ui_primary_color = theme.primary;
  CONFIG.COLORS.ui_secondary_color = theme.secondary;
  CONFIG.COLORS.ui_background_color = theme.background;
  CONFIG.COLORS.ui_surface_color = theme.surface;
  CONFIG.COLORS.ui_text_color = theme.text;
  // 互換のため、primary_color が未定義/空なら同期
  if (!CONFIG.COLORS.primary_color) {
    CONFIG.COLORS.primary_color = theme.primary;
  }
  return {
    theme: themeKey,
    colors: CONFIG.COLORS
  };
}
