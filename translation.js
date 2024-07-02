// var languages2 = [ "pt" ]; // Define languages2 globally
// var languages = [ "es", "fr", "pt", "de", "it", "nl", "vi", "th", "id", "ar" ];


function setTranslateFormulas() {
    var sheetName = "test";
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    
    if (sheet !== null) {
      var lastRow = 84; // Set the number of rows to translate
      var lastColumn = sheet.getLastColumn();
      var languages2 = [ "es", "fr", "pt", "de", "it", "nl", "vi", "th", "id", "ar" ];
      var languages2Values = [ "Spanish (Latin)", "French", "Portuguese (Brazil)", "German", "Italian", "Dutch (Netherlands)", "Vietnamese", "Thai", "Indonesian", "Arabic" ];
  
      for (var j = 0; j < languages2.length; j++) {
        var translatedValues = [];
  
        for (var row = 2; row <= lastRow; row++) {
          var originalValues = sheet.getRange(row, 1, 1, lastColumn).getValues()[0]; // Get the values from each row
  
          var translations = [];
          for (var i = 0; i < originalValues.length; i++) {
            var originalValue = originalValues[i];
            var languageAppTranslatedValue = LanguageApp.translate(originalValue, "en", languages2[j]);
            translations.push(languageAppTranslatedValue);
          }
          translatedValues.push(translations);
          Utilities.sleep(1000);
        }
  
        // Create a new sheet for the translations
        var translatedSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(languages2Values[j]);
        for (var i = 0; i < languages2.length; i++) {
          translatedSheet.getRange(1, i + 1).setValue(languages2[i]); // Set the language headers
        }
        for (var i = 0; i < translatedValues.length; i++) {
          translatedSheet.getRange(i + 2, 1, 1, lastColumn).setValues([translatedValues[i]]); // Set the translated values
        }
      }
    } else {
      console.log("Sheet not found: " + sheetName);
    }
  }
  
  
  
  
  
  