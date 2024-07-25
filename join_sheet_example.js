function consolidateSheets() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var consolidatedSheet = ss.getSheetByName("Consolidado");

    // Crear o limpiar la hoja Consolidado
    if (consolidatedSheet) {
      consolidatedSheet.clear();
    } else {
      consolidatedSheet = ss.insertSheet("Consolidado");
    }

    var data = {};
    var sheetNames = [];
    var numberOfColumnsToConsolidate = 9; // Cambia este valor si necesitas consolidar más o menos columnas
    var useConcatenatedKey = 0; // Cambia a 1 para usar la clave concatenada

    // Función opcional para generar la clave concatenada
    function generateKey(row) {
      var columnsToConcatenate = [0, 2, 4, 6]; // Cambia estos índices para concatenar diferentes columnas
      return columnsToConcatenate.map(function(index) {
        return row[index];
      }).join("_");
    }

    // Recopilar datos de hojas que comienzan con "R_"
    sheets.forEach(function(sheet) {
      var sheetName = sheet.getName();
      if (sheetName.startsWith("R_")) {
        sheetNames.push(sheetName);
        var lastRow = sheet.getLastRow();
        var values = sheet.getRange(1, 1, lastRow, numberOfColumnsToConsolidate).getValues();
        values.forEach(function(row) {
          var key = useConcatenatedKey ? generateKey(row) : row[0];
          if (!data[key]) {
            data[key] = useConcatenatedKey ? [key].concat(row.slice()) : row.slice();
          }
        });
      }
    });

    // Ordenar nombres de hojas en orden Z-A
    sheetNames.sort().reverse();

    // Crear encabezados en la hoja Consolidado
    var headers = sheets.find(sheet => sheet.getName().startsWith("R_")).getRange(1, 1, 1, numberOfColumnsToConsolidate).getValues()[0];
    if (useConcatenatedKey) {
      headers = ["Clave Concatenada"].concat(headers);
    }
    sheetNames.forEach(function(sheetName) {
      headers.push(sheetName.replace("R_", ""));
    });
    consolidatedSheet.appendRow(headers);

    // Agregar datos a la hoja Consolidado
    var consolidatedData = [];
    for (var key in data) {
      var row = data[key];
      sheetNames.forEach(function(sheetName) {
        var sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
          Logger.log("Hoja no encontrada: " + sheetName);
          return;
        }
        var sheetValues = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
        var found = false;
        for (var i = 0; i < sheetValues.length; i++) {
          var sheetKey = useConcatenatedKey ? generateKey(sheetValues[i]) : sheetValues[i][0];
          if (sheetKey === key) {
            row.push(sheetValues[i][sheetValues[i].length - 1]);
            found = true;
            break;
          }
        }
        if (!found) {
          row.push("");
        }
      });
      consolidatedData.push(row);
    }

    // Eliminar duplicados basándonos en la primera columna (o clave concatenada)
    var uniqueData = [];
    var seenKeys = {};
    consolidatedData.forEach(function(row) {
      var key = row[0];
      if (!seenKeys[key]) {
        seenKeys[key] = true;
        uniqueData.push(row);
      }
    });

    // Escribir datos consolidados en la hoja Consolidado
    consolidatedSheet.getRange(2, 1, uniqueData.length, uniqueData[0].length).setValues(uniqueData);

  } catch (e) {
    Logger.log("Error: " + e.message);
    SpreadsheetApp.getUi().alert("Se produjo un error: " + e.message);
  }
}

// Para activar la función de clave concatenada, cambia el valor de useConcatenatedKey a 1
// Para cambiar las columnas que se concatenan, ajusta los índices en el array columnsToConcatenate en la función generateKey
