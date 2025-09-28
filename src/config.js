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
