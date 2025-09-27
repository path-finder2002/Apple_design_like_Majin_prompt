/**
* @OnlyCurrentDoc
* このスクリプトは、指定されたデザインテンプレートに基づきGoogleスライドを自動生成します。
* Version: 17.0 (Generalized Version)
* Prompt Design: まじん式プロンプト
* Author: Googleスライド自動生成マスター
* Description: 指定されたslideData配列とカスタムメニューの設定に基づき、Googleのデザイン原則に準拠したスライドを生成します。
*/


// --- 1. 実行設定 ---
const SETTINGS = {
SHOULD_CLEAR_ALL_SLIDES: true,
TARGET_PRESENTATION_ID: null
};

// --- 2. マスターデザイン設定 (Pixel Perfect Ver.) ---
const LIGHT_THEME_COLORS = {
  canvas: '#FFFFFF',
  primary_color: '#4285F4',
  text_primary: '#000000',
  background_white: '#FFFFFF',
  background_gray: '#F8F9FA',
  faint_gray: '#E8EAED',
  lane_title_bg: '#F8F9FA',
  table_header_bg: '#F8F9FA',
  lane_border: '#DADCE0',
  card_bg: '#FFFFFF',
  card_border: '#DADCE0',
  neutral_gray: '#9E9E9E',
  ghost_gray: '#EFEFED',
  text_on_primary: '#FFFFFF',
  bigFact_caption: '#6E6E73',
  fullBreed_overlay: '#000000',
  fullBreed_text: '#FFFFFF'
};

const DARK_THEME_COLORS = {
  canvas: '#1C1C1E',
  primary_color: '#0A84FF',
  text_primary: '#F5F5F7',
  background_white: '#1C1C1E',
  background_gray: '#2C2C2E',
  faint_gray: '#3A3A3C',
  lane_title_bg: '#2C2C2E',
  table_header_bg: '#2C2C2E',
  lane_border: '#3A3A3C',
  card_bg: '#2C2C2E',
  card_border: '#3A3A3C',
  neutral_gray: '#8E8E93',
  ghost_gray: '#636366',
  text_on_primary: '#FFFFFF',
  bigFact_caption: '#D1D1D6',
  fullBreed_overlay: '#000000',
  fullBreed_text: '#FFFFFF'
};

const CONFIG = {
BASE_PX: { W: 960, H: 540 },

// レイアウトの基準となる不変のpx値
POS_PX: {
titleSlide: {
logo:       { left: 55,  top: 105,  width: 135 },
title:      { left: 50,  top: 200, width: 830, height: 90 },
date:       { left: 50,  top: 450, width: 250, height: 40 },
},

// 共通ヘッダーを持つ各スライド  
contentSlide: {  
  headerLogo:     { right: 20, top: 20, width: 75 },  
  title:          { left: 25, top: 50,  width: 830, height: 65 },  
  titleUnderline: { left: 25, top: 118, width: 260, height: 4 },  
  subhead:        { left: 25, top: 130, width: 830, height: 40 },  
  body:           { left: 25, top: 172, width: 910, height: 290 },  
  twoColLeft:     { left: 25,  top: 172, width: 440, height: 290 },  
  twoColRight:    { left: 495, top: 172, width: 440, height: 290 }  
},  
bigFactSlide: {
  mainValue: { left: 0, top: 120, width: 960, height: 200 },
  caption:   { left: 0, top: 360, width: 960, height: 80 }
},
fullBreedSlide: {
  textArea: { left: 100, top: 160, width: 720, height: 520 },
  itemGap: 180,
  overlayOpacity: null
},
compareSlide: {  
  headerLogo:     { right: 20, top: 20, width: 75 },  
  title:          { left: 25, top: 50,  width: 830, height: 65 },  
  titleUnderline: { left: 25, top: 118, width: 260, height: 4 },  
  subhead:        { left: 25, top: 130, width: 830, height: 40 },  
  leftBox:        { left: 25,  top: 152, width: 430, height: 290 },  
  rightBox:       { left: 485, top: 152, width: 430, height: 290 }  
},  
processSlide: {  
  headerLogo:     { right: 20, top: 20, width: 75 },  
  title:          { left: 25, top: 50,  width: 830, height: 65 },  
  titleUnderline: { left: 25, top: 118, width: 260, height: 4 },  
  subhead:        { left: 25, top: 130, width: 830, height: 40 },  
  area:           { left: 25, top: 152, width: 910, height: 290 }  
},  
timelineSlide: {  
  headerLogo:     { right: 20, top: 20, width: 75 },  
  title:          { left: 25, top: 50,  width: 830, height: 65 },  
  titleUnderline: { left: 25, top: 118, width: 260, height: 4 },  
  subhead:        { left: 25, top: 130, width: 830, height: 40 },  
  area:           { left: 25, top: 172, width: 910, height: 290 }  
},  
diagramSlide: {  
  headerLogo:     { right: 20, top: 20, width: 75 },  
  title:          { left: 25, top: 50,  width: 830, height: 65 },  
  titleUnderline: { left: 25, top: 118, width: 260, height: 4 },  
  subhead:        { left: 25, top: 130, width: 830, height: 40 },  
  lanesArea:      { left: 25, top: 172, width: 910, height: 290 }  
},  
cardsSlide: { // This POS_PX is used by both cards and headerCards
  headerLogo:     { right: 20, top: 20, width: 75 },  
  title:          { left: 25, top: 50,  width: 830, height: 65 },  
  titleUnderline: { left: 25, top: 118, width: 260, height: 4 },  
  subhead:        { left: 25, top: 130, width: 830, height: 40 },  
  gridArea:       { left: 25, top: 160, width: 910, height: 290 }  
},  
tableSlide: {  
  headerLogo:     { right: 20, top: 20, width: 75 },  
  title:          { left: 25, top: 50,  width: 830, height: 65 },  
  titleUnderline: { left: 25, top: 118, width: 260, height: 4 },  
  subhead:        { left: 25, top: 130, width: 830, height: 40 },  
  area:           { left: 25, top: 160, width: 910, height: 290 }  
},  
progressSlide: {  
  headerLogo:     { right: 20, top: 20, width: 75 },  
  title:          { left: 25, top: 50,  width: 830, height: 65 },  
  titleUnderline: { left: 25, top: 118, width: 260, height: 4 },  
  subhead:        { left: 25, top: 130, width: 830, height: 40 },  
  area:           { left: 25, top: 172, width: 910, height: 290 }  
},

quoteSlide: {
  headerLogo:     { right: 20, top: 20, width: 75 },
  title:          { left: 25, top: 50,  width: 830, height: 65 },
  titleUnderline: { left: 25, top: 118, width: 260, height: 4 },
  subhead:        { left: 25, top: 130, width: 830, height: 40 },
  quoteMark:      { left: 40, top: 180, width: 100, height: 100 },
  quoteText:      { left: 150, top: 210, width: 700, height: 150 },
  author:         { right: 110, top: 370, width: 700, height: 30 }
},

kpiSlide: {
  headerLogo:     { right: 20, top: 20, width: 75 },
  title:          { left: 25, top: 50,  width: 830, height: 65 },
  titleUnderline: { left: 25, top: 118, width: 260, height: 4 },
  subhead:        { left: 25, top: 130, width: 830, height: 40 },
  gridArea:       { left: 25, top: 172, width: 910, height: 290 }
},

statsCompareSlide: {
  headerLogo:     { right: 20, top: 20, width: 75 },
  title:          { left: 25, top: 50,  width: 830, height: 65 },
  titleUnderline: { left: 25, top: 118, width: 260, height: 4 },
  subhead:        { left: 25, top: 130, width: 830, height: 40 },
  leftArea:       { left: 25, top: 172, width: 430, height: 290 },
  rightArea:      { left: 485, top: 172, width: 430, height: 290 }
},

sectionSlide: {  
  title:      { left: 55, top: 230, width: 840, height: 80 },  
  ghostNum:   { left: 35, top: 120, width: 400, height: 200 }
},

footer: {  
  leftText:  { left: 15, top: 505, width: 250, height: 20 },  
  rightPage: { right: 15, top: 505, width: 50,  height: 20 }  
},  
bottomBar: { left: 0, top: 534, width: 960, height: 6 }  

},

// フォントと色と図形サイズ
FONTS: {
family: 'Arial', // デフォルト、プロパティから動的に変更可能
sizes: {
title: 40, date: 16, sectionTitle: 38, contentTitle: 28, subhead: 18,
body: 14, footer: 9, chip: 11, laneTitle: 13, small: 10,
processStep: 14, axis: 12, ghostNum: 180,
bigFactMain: 180, bigFactCaption: 36,
fullBreedItem: 90
}
},
COLORS: Object.assign({}, LIGHT_THEME_COLORS),
THEMES: {
  light: {
    key: 'light',
    displayName: 'ライトモード',
    colors: LIGHT_THEME_COLORS
  },
  dark: {
    key: 'dark',
    displayName: 'ダークモード',
    colors: DARK_THEME_COLORS
  }
},
DIAGRAM: {
laneGap_px: 24, lanePad_px: 10, laneTitle_h_px: 30, cardGap_px: 12,
cardMin_h_px: 48, cardMax_h_px: 70, arrow_h_px: 10, arrowGap_px: 8
},

LOGOS: {
header: 'https://upload.wikimedia.org/wikipedia/commons/thumb/2/2f/Google_2015_logo.svg/640px-Google_2015_logo.svg.png',
closing: 'https://upload.wikimedia.org/wikipedia/commons/thumb/2/2f/Google_2015_logo.svg/640px-Google_2015_logo.svg.png'
},

IMAGES: {
  fullBreedBackground: 'https://unsplash.com/photos/g-d0I9CMnZw/download?force=true&w=2400'
},

FOOTER_TEXT: `© ${new Date().getFullYear()} Google Inc.`
};
