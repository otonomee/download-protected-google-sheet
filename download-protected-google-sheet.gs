function download_sheet(sheet_url) {
  const sh = SpreadsheetApp.openByUrl(sheet_url).getSheets()[0]
  
  const nSheet = SpreadsheetApp.create(sh.getName() + ": copy");
  const o_range = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns())

  const c_range = nSheet.getSheets()[0].getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns())

  // Get the values and styles from the original
  const b = o_range.getBackgrounds()
  const c = o_range.getFontColors()
  const f = o_range.getFontSizes()
  var values = o_range.getDisplayValues();
  var background = o_range.getBackgrounds();
  var fontColor = o_range.getFontColors();
  var fontFamily = o_range.getFontFamilies();
  var fontLine = o_range.getFontLines();
  var fontSize = o_range.getFontSizes();
  var fontStyle = o_range.getFontStyles();
  var fontWeight = o_range.getFontWeights();
  var textStyle = o_range.getTextStyles();
  var horAlign = o_range.getHorizontalAlignments();
  var vertAlign = o_range.getVerticalAlignments();
  var bandings = o_range.getBandings();
  var mergedRanges = o_range.getMergedRanges();
 

  c_range.setBackgrounds(background);
  c_range.setFontColors(fontColor);
  c_range.setFontFamilies(fontFamily);
  c_range.setFontLines(fontLine);
  c_range.setFontSizes(fontSize);
  c_range.setFontStyles(fontStyle);
  c_range.setFontWeights(fontWeight);
  c_range.setTextStyles(textStyle);
  c_range.setHorizontalAlignments(horAlign);
  c_range.setVerticalAlignments(vertAlign);
  c_range.setValues(values);
  const formulaA1Notation = ["L1", "L2", "L3", "L10", "L11", "L12", "L13", "L14", "L18", "L19", "L20", "L21", "L22", "L26", "L27", "L28", "L29", "L30", "A3", "C3", "F3", "G3", "H3", "F5", "G5", "H5", "A15", "C18"]

    formulaA1Notation.forEach(a1Notation => {
      var sourceFormulas = sh.getRange(a1Notation).getFormulas();
      nSheet.getRange(a1Notation).setFormulas(sourceFormulas);
    });

  for (let i in bandings){
    let srcBandA1 = bandings[i].getRange().getA1Notation();
    let destBandRange = nSheet.getRange(srcBandA1);

    destBandRange.applyRowBanding()
    .setFirstRowColor(bandings[i].getFirstRowColor())
    .setSecondRowColor(bandings[i].getSecondRowColor())
    .setHeaderRowColor(bandings[i].getHeaderRowColor())
    .setFooterRowColor(bandings[i].getFooterRowColor());
  }

  for (let i = 0; i < mergedRanges.length; i++) {
    nSheet.getRange(mergedRanges[i].getA1Notation()).merge();
  }
 
 try {
  for (let i = 1; i <= o_range.getWidth(); i++) {
    let width = sh.getColumnWidth(i);
    nSheet.setColumnWidth(i, width);
  }
 
  for (let i = 1; i <= o_range.getHeight(); i++){
    let height = sh.getRowHeight(i);
    nSheet.setRowHeight(i, height);
  }
 } catch(e) { console.log(e) }

  var URL = 'https://docs.google.com/spreadsheets/d/' + nSheet.getId() + '/export?format=xlsx';
  console.log(URL)
}

// Insert link to Google Sheets spreadsheet
let sheet_download = ''

/*
Running this in the App Scripts Editor should instantiate a 
copy of said spreadsheet to your local GDrive
*/
download_sheet(sheet_download)