function onOpen(e) {
  const ui = SlidesApp.getUi();
  ensureTheme();

  ui.createMenu('カスタム設定')
    .addItem('🎨 スライドを生成', 'generatePresentation')
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ 設定')
      .addItem('プライマリカラー', 'setPrimaryColor')
      .addItem('フォント', 'setFont')
      .addItem('フッターテキスト', 'setFooterText')
      .addItem('ヘッダーロゴ', 'setHeaderLogo')
      .addItem('クロージングロゴ', 'setClosingLogo'))
    .addSeparator()
    .addSubMenu(ui.createMenu('テキスト編集')
      .addItem('太字/解除', 'toggleBold')
      .addItem('全スライド太字/解除', 'toggleBoldAllSlides'))
    .addSeparator()
    .addItem(getThemeToggleMenuLabel(), 'toggleTheme')
    .addItem('🔄 リセット', 'resetSettings')
    .addToUi();
}

// プライマリカラー設定
function setPrimaryColor() {
  const ui = SlidesApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const currentValue = props.getProperty('primaryColor') || '#4285F4';
  
  const result = ui.prompt(
    'プライマリカラー設定',
    `カラーコードを入力してください（例: #4285F4）\n現在値: ${currentValue}\n\n空欄で既定値にリセットされます。`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const value = result.getResponseText().trim();
    if (value === '') {
      props.deleteProperty('primaryColor');
      ui.alert('プライマリカラーをリセットしました。');
    } else {
      props.setProperty('primaryColor', value);
      ui.alert('プライマリカラーを保存しました。');
    }
  }
}

// フォント設定（プルダウン形式）
function setFont() {
  const ui = SlidesApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const currentValue = props.getProperty('fontFamily') || 'Arial';
  
  const fonts = [
    'Arial',
    'Noto Sans JP',
    'M PLUS 1p',
    'Noto Serif JP'
  ];
  
  const fontList = fonts.map((font, index) => 
    `${index + 1}. ${font}${font === currentValue ? ' (現在)' : ''}`
  ).join('\n');
  
  const result = ui.prompt(
    'フォント設定',
    `使用するフォントの番号を入力してください:\n\n${fontList}\n\n※ 空欄で既定値（Arial）にリセット`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const input = result.getResponseText().trim();
    if (input === '') {
      props.deleteProperty('fontFamily');
      ui.alert('フォントをリセットしました（Arial）。');
    } else {
      const index = parseInt(input) - 1;
      if (index >= 0 && index < fonts.length) {
        props.setProperty('fontFamily', fonts[index]);
        ui.alert(`フォントを「${fonts[index]}」に設定しました。`);
      } else {
        ui.alert('無効な番号です。設定をキャンセルしました。');
      }
    }
  }
}

// フッターテキスト設定
function setFooterText() {
  const ui = SlidesApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const currentValue = props.getProperty('footerText') || '未設定';
  
  const result = ui.prompt(
    'フッターテキスト設定',
    `フッターに表示するテキストを入力してください\n現在値: ${currentValue}\n\n空欄でリセットされます。`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const value = result.getResponseText().trim();
    if (value === '') {
      props.deleteProperty('footerText');
      ui.alert('フッターテキストをリセットしました。');
    } else {
      props.setProperty('footerText', value);
      ui.alert('フッターテキストを保存しました。');
    }
  }
}

// ヘッダーロゴ設定
function setHeaderLogo() {
  const ui = SlidesApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const currentValue = props.getProperty('headerLogoUrl') || '未設定';
  
  const result = ui.prompt(
    'ヘッダーロゴ設定',
    `ヘッダーロゴのURLを入力してください\n現在値: ${currentValue}\n\n空欄でリセットされます。`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const value = result.getResponseText().trim();
    if (value === '') {
      props.deleteProperty('headerLogoUrl');
      ui.alert('ヘッダーロゴをリセットしました。');
    } else {
      props.setProperty('headerLogoUrl', value);
      ui.alert('ヘッダーロゴを保存しました。');
    }
  }
}

// クロージングロゴ設定
function setClosingLogo() {
  const ui = SlidesApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const currentValue = props.getProperty('closingLogoUrl') || '未設定';
  
  const result = ui.prompt(
    'クロージングロゴ設定',
    `クロージングページのロゴURLを入力してください\n現在値: ${currentValue}\n\n空欄でリセットされます。`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const value = result.getResponseText().trim();
    if (value === '') {
      props.deleteProperty('closingLogoUrl');
      ui.alert('クロージングロゴをリセットしました。');
    } else {
      props.setProperty('closingLogoUrl', value);
      ui.alert('クロージングロゴを保存しました。');
    }
  }
}

function resetSettings() {
  const ui = SlidesApp.getUi();
  const result = ui.alert('設定のリセット', 'すべてのカスタム設定をリセットしますか？', ui.ButtonSet.YES_NO);
  
  if (result === ui.Button.YES) {
    PropertiesService.getScriptProperties().deleteAllProperties();
