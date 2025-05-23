
# Este script puede iterar a través de las celdas y obtener la URL del hipervínculo asociado a cada una.
# Se scribirán los enlaces extraídos en las columnas siguientes a los datos originales.


function extraerHipervinculos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getActiveSheet();
  var rango = hoja.getDataRange(); // Obtiene todas las celdas con datos
  var valores = rango.getValues();
  var hipervinculos = [];

  for (var i = 0; i < valores.length; i++) {
    var filaHipervinculos = [];
    for (var j = 0; j < valores[i].length; j++) {
      var hipervinculo = hoja.getRange(i + 1, j + 1).getRichTextValue().getLinkUrl();
      filaHipervinculos.push(hipervinculo);
    }
    hipervinculos.push(filaHipervinculos);
  }
  hoja.getRange(1, valores[0].length + 1, hipervinculos.length, hipervinculos[0].length)
      .setValues(hipervinculos);
}
