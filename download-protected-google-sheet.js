function downloadXLS_GUI() {
    const numCols = 20
    //const sh = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1KfxRBuGh_Idq5zHLLTUhwhz0dFN9cTKVCYXb7WW15qk/edit#gid=0').getSheets()[0]
    const ssID = ''
    const sh = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/' + ssID).getSheets()[1]
    const nSheet = SpreadsheetApp.create(sh.getName() + ": copy");
  
    //const o_range = sh.getRange(1, 1, sh.getLastRow(), numCols)
    const o_range = sh.getRange(1, 1, 20, numCols)
  
    //const c_range = nSheet.getSheets()[0].getRange(1, 1, sh.getLastRow(), numCols)
    const c_range = nSheet.getSheets()[0].getRange(1, 1, 20, numCols)
  
    // Get the values and styles from the original
    const v = o_range.getValues()
    const b = o_range.getBackgrounds()
    const c = o_range.getFontColors()
    const f = o_range.getFontSizes()
    // Set the values and styles to the copy
    c_range.setValues(v)
    c_range.setBackgrounds(b)
    c_range.setFontColors(c)
    c_range.setFontSizes(f)
    const formulaA1Notation = ["A3", "C3", "F3", "G3", "H3", "F5", "G5", "H5", "A15", "C18"]
    formulaA1Notation.forEach(a1Notation => {
      var sourceFormulas = sh.getRange(a1Notation).getFormulas();
      nSheet.getRange(a1Notation).setFormulas(sourceFormulas);
    });
    var URL = 'https://docs.google.com/spreadsheets/d/' + nSheet.getId() + '/export?format=xlsx';
    console.log(URL)
  }
  
  
  