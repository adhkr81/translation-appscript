// New script
// This script Jigwan added the HTML to RICH TEXT script, to send to OSM for proofreadings and add back to vkrs
// vkrs html sheet -> translate with app script -> rich text with script




function setTranslateFormulas() {
  const sheetName = "English (HTML)";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if (!sheet) {
    console.log("Sheet not found: " + sheetName);
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  const languages2 = ["es"];  // Add more if needed
  const languages2Values = ["Spanish (Latin)"];

  for (let j = 0; j < languages2.length; j++) {
    const translatedValues = [];

    for (let row = 2; row <= lastRow; row++) {
      const originalValues = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];
      const translations = [];

      for (let i = 0; i < originalValues.length; i++) {
        let originalValue = originalValues[i];
          if (typeof originalValue === 'string') {
            originalValue = originalValue
              .replace(/<b>(.*?)<\/b>/gi, '[[B]]$1[[/B]]')
              .replace(/<strong>(.*?)<\/strong>/gi, '[[B]]$1[[/B]]')
              .replace(/<i>(.*?)<\/i>/gi, '[[I]]$1[[/I]]')
              .replace(/<em>(.*?)<\/em>/gi, '[[I]]$1[[/I]]')
              .replace(/<u>(.*?)<\/u>/gi, '[[U]]$1[[/U]]')
              .replace(/<br\s*\/?>/gi, '[[BR]]')
              .replace(/<p>/gi, '')
              .replace(/<\/p>/gi, '');
          } else {
            originalValue = String(originalValue || ''); // convert numbers/null to string
          }

          let translatedValue = LanguageApp.translate(originalValue, "en", languages2[j]);
          // Restore <br> from [[BR]]
          translatedValue = translatedValue.replace(/\[\[BR\]\]/g, '<br>');

          // Now remove any other unwanted raw HTML tags
          translatedValue = translatedValue.replace(/<(?!br\s*\/?>)[^>]+>/gi, '');


        translations.push(translatedValue);
      }

      translatedValues.push(translations);
      Utilities.sleep(500);
    }

    // Create sheet for translated content
    const translatedSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(languages2Values[j]);
    translatedSheet.getRange(1, 1).setValue(languages2[j]);

    for (let i = 0; i < translatedValues.length; i++) {
      translatedSheet.getRange(i + 2, 1, 1, lastColumn).setValues([translatedValues[i]]);
    }

    convertPlaceholdersToRichText(translatedSheet, translatedValues.length, lastColumn);
  }
}


function convertPlaceholdersToRichText(sheet, numRows, numCols) {
  const range = sheet.getRange(2, 1, numRows, numCols);
  const values = range.getValues();

  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      const rawText = values[i][j];
      if (typeof rawText === "string") {
        const richText = parsePlaceholdersToRichText(rawText);
        sheet.getRange(i + 2, j + 1).setRichTextValue(richText);
      }
    }
  }
}


function parsePlaceholdersToRichText(text) {
  const tagRegex = /\[\[(\/?)(B|I|U|BR)\]\]/g;
  let plainText = '';
  let tagStack = [];
  let indexMap = []; // Stores style info by char index

  // Replace tags and build index map
  let lastIndex = 0;
  let match;

  while ((match = tagRegex.exec(text)) !== null) {
    const tag = match[2];
    const isClosing = match[1] === '/';
    const tagStart = match.index;

    // Append normal text between last tag and this tag
    const textChunk = text.substring(lastIndex, tagStart);
    for (let k = 0; k < textChunk.length; k++) {
      const styles = tagStack.slice(); // copy current styles
      indexMap.push(styles);
      plainText += textChunk[k];
    }

    if (tag === 'BR') {
      plainText += '\n';
      indexMap.push(tagStack.slice());
    } else {
      if (isClosing) {
        // Remove tag from stack
        const pos = tagStack.indexOf(tag);
        if (pos !== -1) tagStack.splice(pos, 1);
      } else {
        // Add tag to stack
        tagStack.push(tag);
      }
    }

    lastIndex = tagRegex.lastIndex;
  }

  // Append remaining text after last tag
  const remaining = text.substring(lastIndex);
  for (let k = 0; k < remaining.length; k++) {
    const styles = tagStack.slice();
    indexMap.push(styles);
    plainText += remaining[k];
  }

  // Now apply styles
  const builder = SpreadsheetApp.newRichTextValue().setText(plainText);
  let start = 0;

  while (start < plainText.length) {
    const currentStyle = indexMap[start].sort().join();
    let end = start + 1;

    while (
      end < plainText.length &&
      indexMap[end].sort().join() === currentStyle
    ) {
      end++;
    }

    const styleBuilder = SpreadsheetApp.newTextStyle();
    if (currentStyle.includes('B')) styleBuilder.setBold(true);
    if (currentStyle.includes('I')) styleBuilder.setItalic(true);
    if (currentStyle.includes('U')) styleBuilder.setUnderline(true);

    builder.setTextStyle(start, end, styleBuilder.build());
    start = end;
  }

  return builder.build();
}

/* old script
function setTranslateFormulas() {
  var sheetName = "Sheet5";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if (sheet !== null) {
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();

    // Define language codes and their sheet names
    var languages2 = [ "es" ];  // Add more codes if needed (e.g., "fr", "pt")
    var languages2Values = [ "Spanish (Latin)" ];  // Match order

    for (var j = 0; j < languages2.length; j++) {
      var translatedValues = [];

      for (var row = 2; row <= lastRow; row++) {
        var originalValues = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];
        var translations = [];

        for (var i = 0; i < originalValues.length; i++) {
          var originalValue = originalValues[i];
          var translatedValue = LanguageApp.translate(originalValue, "en", languages2[j]);
          translations.push(translatedValue);
        }
        translatedValues.push(translations);
        Utilities.sleep(1000);  // Prevent quota issues
      }

      // Create a new sheet for the translations
      var translatedSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(languages2Values[j]);
      translatedSheet.getRange(1, 1).setValue(languages2[j]);

      for (var i = 0; i < translatedValues.length; i++) {
        translatedSheet.getRange(i + 2, 1, 1, lastColumn).setValues([translatedValues[i]]);
      }

      // ✅ Convert translated HTML strings to rich text
      convertHtmlToRichText(translatedSheet, translatedValues.length, lastColumn);
    }
  } else {
    console.log("Sheet not found: " + sheetName);
  }
}


function convertHtmlToRichText(sheet, numRows, numCols) {
  const range = sheet.getRange(2, 1, numRows, numCols);
  const values = range.getValues();

  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      const cellText = values[i][j];
      if (typeof cellText === 'string') {
        const richText = parseHtmlToRichText(cellText);
        sheet.getRange(i + 2, j + 1).setRichTextValue(richText);
      }
    }
  }
}


function parseHtmlToRichText(html) {
  // Replace <br>, <br/>, <br /> with newlines before stripping tags
  const htmlWithLineBreaks = html.replace(/<br\s*\/?>/gi, '\n');

  const tempDiv = HtmlService.createHtmlOutput(htmlWithLineBreaks).getContent();
  const cleanText = tempDiv.replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ');

  const boldMatches = [...html.matchAll(/<(strong|b)>(.*?)<\/\1>/gi)];
  const italicMatches = [...html.matchAll(/<(em|i)>(.*?)<\/\1>/gi)];
  const underlineMatches = [...html.matchAll(/<u>(.*?)<\/u>/gi)];

  let builder = SpreadsheetApp.newRichTextValue().setText(cleanText);
  let indexOffset = 0;

  [boldMatches, italicMatches, underlineMatches].forEach((matches, groupIdx) => {
    matches.forEach(match => {
      const tagContent = match[2] || match[1];
      const tagStart = cleanText.indexOf(tagContent, indexOffset);
      if (tagStart !== -1) {
        const tagEnd = tagStart + tagContent.length;
        const styleBuilder = SpreadsheetApp.newTextStyle();
        if (groupIdx === 0) styleBuilder.setBold(true);
        if (groupIdx === 1) styleBuilder.setItalic(true);
        if (groupIdx === 2) styleBuilder.setUnderline(true);
        builder.setTextStyle(tagStart, tagEnd, styleBuilder.build());
        indexOffset = tagEnd;
      }
    });
  });

  return builder.build();
}
*/
