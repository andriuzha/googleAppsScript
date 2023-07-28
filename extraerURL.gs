/** 
 * Devuelve la URL de una celda con hipervínculo 
 * Requiere de un comando de hypervínculos. 
 * Compatible con rangos de celdas
 * @Requiere de la siguiente referencia:
 * @param {A1} 
 * @customfunction
 *
 */
function linkURL(reference) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var formula = SpreadsheetApp.getActiveRange().getFormula();
  var args = formula.match(/=\w+\((.*)\)/i);
  try {
    var range = sheet.getRange(args[1]);
  }
  catch(e) {
    throw new Error(args[1] + ' no es un rango válido');
  }

  var formulas = range.getRichTextValues();
  var output = [];
  for (var i = 0; i < formulas.length; i++) {
    var row = [];
    for (var j = 0; j < formulas[0].length; j++) {
      row.push(formulas[i][j].getLinkUrl());
    }
    output.push(row);
  }
  return output
}
