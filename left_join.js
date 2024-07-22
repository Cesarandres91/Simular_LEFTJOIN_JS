function leftJoinMultipleTablesWithOptionalColumns(mainSheetName, lookupSheetNames, mainKeyColLetters, lookupKeyColLetters, selectedColumns, resultSheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Crear o limpiar la hoja de resultados
  var resultSheet = ss.getSheetByName(resultSheetName);
  if (resultSheet) {
    ss.deleteSheet(resultSheet);
  }
  resultSheet = ss.insertSheet(resultSheetName);

  // Convertir letras de columna a índices
  var mainKeyCols = mainKeyColLetters.map(letterToColumn);
  var lookupKeyCols = lookupKeyColLetters.map(letterToColumn);

  // Obtener los datos de la hoja principal
  var mainSheet = ss.getSheetByName(mainSheetName);
  var mainData = mainSheet.getDataRange().getValues();
  
  // Obtener los encabezados de la hoja principal
  var mainHeaders = mainData[0];
  var resultHeaders = mainHeaders.slice();
  var resultData = [];
  resultData.push(resultHeaders);

  // Iterar a través de cada hoja de búsqueda para hacer el join
  for (var n = 0; n < lookupSheetNames.length; n++) {
    var lookupSheetName = lookupSheetNames[n];
    var lookupSheet = ss.getSheetByName(lookupSheetName);
    var lookupData = lookupSheet.getDataRange().getValues();
    var lookupHeaders = lookupData[0];

    // Determinar las columnas seleccionadas
    var selectedColIndices = (selectedColumns && selectedColumns[n] && selectedColumns[n].length > 0) 
        ? selectedColumns[n].map(letterToColumn) 
        : lookupHeaders.map((_, index) => index);

    // Agregar encabezados de búsqueda seleccionados a los resultados
    var selectedLookupHeaders = selectedColIndices.map(i => lookupSheetName + "_" + lookupHeaders[i]);
    resultHeaders = resultHeaders.concat(selectedLookupHeaders);
    resultData[0] = resultHeaders;
    
    // Crear un índice para las columnas de búsqueda por clave
    var lookupIndex = {};
    for (var i = 1; i < lookupData.length; i++) {
      var key = lookupData[i][lookupKeyCols[n]];
      if (!lookupIndex[key]) {
        lookupIndex[key] = [];
      }
      lookupIndex[key].push(lookupData[i]);
    }

    // Realizar el LEFT JOIN
    var newResultData = [];
    for (var i = 1; i < mainData.length; i++) {
      var mainRow = mainData[i];
      var lookupRows = lookupIndex[mainRow[mainKeyCols[n]]];
      if (lookupRows) {
        for (var j = 0; j < lookupRows.length; j++) {
          var lookupRow = lookupRows[j];
          var selectedLookupRow = selectedColIndices.map(k => lookupRow[k]);
          newResultData.push(mainRow.concat(selectedLookupRow));
        }
      } else {
        // Si no se encuentra una coincidencia, agregar nulos para las columnas de búsqueda seleccionadas
        newResultData.push(mainRow.concat(Array(selectedColIndices.length).fill(null)));
      }
    }

    // Actualizar mainData con los nuevos datos combinados
    mainData = [resultData[0]].concat(newResultData);
  }
  
  // Establecer los datos en la hoja de resultados
  resultSheet.getRange(1, 1, mainData.length, mainData[0].length).setValues(mainData);
}

function letterToColumn(letter) {
  var column = 0;
  var length = letter.length;
  for (var i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column - 1;
}

function leftJoinExampleWithOptionalColumns() {
  // Unir H1, H2 y H3 usando columnas B, A y C respectivamente
  // y seleccionando columnas específicas de cada hoja. Si no se especifican columnas, se traen todas.
  leftJoinMultipleTablesWithOptionalColumns(
    'H1', 
    ['H2', 'H3'], 
    ['B', 'B'], 
    ['A', 'C'], 
    [['A', 'B', 'C'], null], // Columnas seleccionadas de H2 y todas las columnas de H3
    'Resultado'
  );
}
