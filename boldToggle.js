// Bold toggle utilities extracted from main script.

function toggleBold() {
  const ui = SlidesApp.getUi();

  try {
    const presentation = SlidesApp.getActivePresentation();
    const selection = presentation.getSelection();
    const targetSlides = resolveTargetSlides(selection);

    if (targetSlides.length === 0) {
      ui.alert('太字を切り替えるスライドを選択してください。');
      return;
    }

    handleBoldToggle(targetSlides, ui, '対象のテキストが見つかりませんでした。', 'スライド内のテキストを太字にしました。', 'スライド内の太字を解除しました。');
  } catch (error) {
    logError('toggleBold failed', error);
    ui.alert('太字の切り替えに失敗しました。もう一度お試しください。');
  }
}

function toggleBoldAllSlides() {
  const ui = SlidesApp.getUi();

  try {
    const presentation = SlidesApp.getActivePresentation();
    const slides = presentation.getSlides();

    if (!slides || slides.length === 0) {
      ui.alert('プレゼンテーション内にスライドがありません。');
      return;
    }

    handleBoldToggle(slides, ui, '対象のテキストが見つかりませんでした。', '全スライドのテキストを太字にしました。', '全スライドの太字を解除しました。');
  } catch (error) {
    logError('toggleBoldAllSlides failed', error);
    ui.alert('全スライドの太字の切り替えに失敗しました。もう一度お試しください。');
  }
}

function handleBoldToggle(slides, ui, emptyMessage, boldMessage, unboldMessage) {
  const textRanges = collectTextRangesFromSlides(slides);

  if (textRanges.length === 0) {
    ui.alert(emptyMessage);
    return;
  }

  const shouldBold = textRanges.some(range => {
    try {
      const style = range.getTextStyle();
      const state = style.isBold();
      return state === null || state === false;
    } catch (e) {
      return false;
    }
  });

  textRanges.forEach(range => {
    try {
      range.getTextStyle().setBold(shouldBold);
    } catch (e) {}
  });

  ui.alert(shouldBold ? boldMessage : unboldMessage);
}

function resolveTargetSlides(selection) {
  if (!selection) return [];

  const slides = [];

  try {
    const pageRange = selection.getPageRange();
    if (pageRange) {
      const rangeSlides = pageRange.getSlides();
      if (rangeSlides && rangeSlides.length) {
        rangeSlides.forEach(slide => slides.push(slide));
      }
    }
  } catch (e) {}

  if (!slides.length) {
    try {
      const currentPage = selection.getCurrentPage();
      if (currentPage && currentPage.getPageType && currentPage.getPageType() === SlidesApp.PageType.SLIDE) {
        slides.push(currentPage.asSlide());
      }
    } catch (e) {}
  }

  if (!slides.length) {
    try {
      const elementRange = selection.getPageElementRange();
      if (elementRange) {
        elementRange.getPageElements().forEach(element => {
          try {
            const parent = element.getParentPage();
            if (parent && parent.getPageType && parent.getPageType() === SlidesApp.PageType.SLIDE) {
              slides.push(parent.asSlide());
            }
          } catch (e) {}
        });
      }
    } catch (e) {}
  }

  if (slides.length <= 1) {
    return slides;
  }

  const unique = [];
  const seen = {};
  slides.forEach(slide => {
    try {
      const id = slide.getObjectId();
      if (!seen[id]) {
        seen[id] = true;
        unique.push(slide);
      }
    } catch (e) {}
  });
  return unique;
}

function collectTextRangesFromSlides(slides) {
  const bucket = [];
  (slides || []).forEach(slide => collectTextRangesFromSlide(slide, bucket));
  return bucket;
}

function collectTextRangesFromSlide(slide, bucket) {
  if (!slide) return;

  try {
    const elements = slide.getPageElements();
    elements.forEach(element => collectTextRangesFromElement(element, bucket));
  } catch (e) {}

  try {
    const notesShape = slide.getNotesPage().getSpeakerNotesShape();
    if (notesShape) {
      const notesText = notesShape.getText();
      pushRangeIfMeaningful(notesText, bucket);
    }
  } catch (e) {}
}

function collectTextRangesFromElement(element, bucket) {
  if (!element) return;

  try {
    const type = element.getPageElementType();

    if (type === SlidesApp.PageElementType.GROUP) {
      element.asGroup().getChildren().forEach(child => collectTextRangesFromElement(child, bucket));
      return;
    }

    if (type === SlidesApp.PageElementType.TABLE) {
      const table = element.asTable();
      for (let r = 0; r < table.getNumRows(); r++) {
        for (let c = 0; c < table.getNumColumns(); c++) {
          pushRangeIfMeaningful(table.getCell(r, c).getText(), bucket);
        }
      }
      return;
    }

    if (
      type === SlidesApp.PageElementType.SHAPE ||
      type === SlidesApp.PageElementType.TEXT_BOX ||
      type === SlidesApp.PageElementType.WORD_ART ||
      type === SlidesApp.PageElementType.PLACEHOLDER
    ) {
      pushRangeIfMeaningful(element.asShape().getText(), bucket);
      return;
    }
  } catch (e) {}
}

function pushRangeIfMeaningful(textRange, bucket) {
  if (!textRange) return;

  const text = extractTextFromRange(textRange);
  if (!text) return;

  if (text.replace(/[\s\u200B\u00A0]+/g, '') === '') return;

  bucket.push(textRange);
}

function extractTextFromRange(textRange) {
  if (!textRange) return '';

  try {
    if (typeof textRange.asString === 'function') {
      return textRange.asString();
    }
  } catch (e) {}

  try {
    if (typeof textRange.getText === 'function') {
      return textRange.getText();
    }
  } catch (e) {}

  return '';
}
