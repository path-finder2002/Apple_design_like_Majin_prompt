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
processStep: 14, axis: 12, ghostNum: 180
}
},
COLORS: {
canvas: '#FFFFFF',
primary_color: '#4285F4',
text_primary: '#000000',
background_white: '#FFFFFF',
background_gray: '#f8f9fa',
faint_gray: '#e8eaed',
lane_title_bg: '#f8f9fa',
table_header_bg: '#f8f9fa',
lane_border: '#dadce0',
card_bg: '#ffffff',
card_border: '#dadce0',
neutral_gray: '#9e9e9e',
ghost_gray: '#efefed',
text_on_primary: '#FFFFFF'
},
DIAGRAM: {
laneGap_px: 24, lanePad_px: 10, laneTitle_h_px: 30, cardGap_px: 12,
cardMin_h_px: 48, cardMax_h_px: 70, arrow_h_px: 10, arrowGap_px: 8
},

LOGOS: {
header: 'https://upload.wikimedia.org/wikipedia/commons/thumb/2/2f/Google_2015_logo.svg/640px-Google_2015_logo.svg.png',
closing: 'https://upload.wikimedia.org/wikipedia/commons/thumb/2/2f/Google_2015_logo.svg/640px-Google_2015_logo.svg.png'
},

FOOTER_TEXT: `© ${new Date().getFullYear()} Google Inc.`
};

// --- 3. スライドデータ（サンプル：必ず置換してください） ---
const slideData = [
  { type: 'title', title: 'Google Workspace 新機能提案', date: '2025.08.24', notes: '本日は、AIを活用した新しいコラボレーション機能についてご提案します。' },
  {
    type: 'bulletCards',
    title: '提案する3つの新機能',
    subhead: 'チームの生産性をさらに向上させるためのコンセプト',
    items: [
      {
        title: 'AIミーティングサマリー',
        desc: 'Google Meetでの会議内容をAIが自動で要約し、[[決定事項とToDoリストを自動生成]]します。'
      },
      {
        title: 'スマート・ドキュメント連携',
        desc: 'DocsやSheetsで関連するファイルやデータをAIが予測し、[[ワンクリックで参照・引用]]できるようにします。'
      },
      {
        title: 'インタラクティブ・チャット',
        desc: 'Google Chat内で簡易的なアンケートや投票、承認フローを[[コマンド一つで実行]]可能にします。'
      }
    ],
    notes: '今回ご提案するのは、この3つの新機能です。それぞれが日々の業務の非効率を解消し、チーム全体の生産性向上を目指しています。'
  },
  {
    type: 'faq',
    title: '想定されるご質問',
    subhead: '本提案に関するQ&A',
    items: [
      { q: '既存のプランで利用できますか？', a: 'はい、Business Standard以上のすべてのプランで、追加料金なしでご利用いただける想定です。' },
      { q: '対応言語はどうなりますか？', a: '初期リリースでは日本語と英語に対応し、順次対応言語を拡大していく計画です。' },
      { q: 'セキュリティは考慮されていますか？', a: 'もちろんです。すべてのデータは既存のGoogle Workspaceの[[堅牢なセキュリティ基準]]に準拠して処理されます。' }
    ],
    notes: 'ご提案にあたり、想定される質問をまとめました。ご不明な点がございましたら、お気軽にご質問ください。'
  },
  { type: 'closing', notes: '本日のご提案は以上です。ご清聴いただき、ありがとうございました。' }
];


// --- 4. メイン実行関数（エントリーポイント） ---
let __SECTION_COUNTER = 0; // 章番号カウンタ（ゴースト数字用）
