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
    var useConcatenatedKey = 1; // Cambia a 1 para usar la clave concatenada, o 0 para omitir
    var specificColumnToAdd = -1; // -1 para indicar que se debe usar la última columna

    // Función opcional para generar la clave concatenada
    function generateKey(row) {
      var columnsToConcatenate = [0, 2, 4, 8]; // Cambia estos índices para concatenar diferentes columnas
      return columnsToConcatenate.map(function(index) {
        return row[index];
      }).join("_");
    }

    // Función para convertir fecha en formato dd.mm.yy a objeto Date
    function parseDate(dateString) {
      var parts = dateString.split(".");
      return new Date(parts[2], parts[1] - 1, parts[0]);
    }

    // Recopilar datos de hojas que comienzan con "R_"
    sheets.forEach(function(sheet) {
      var sheetName = sheet.getName();
      if (sheetName.startsWith("R_")) {
        sheetNames.push(sheetName);
        var lastRow = sheet.getLastRow();
        var values = sheet.getRange(2, 1, lastRow - 1, numberOfColumnsToConsolidate).getValues(); // Saltar la primera fila (encabezados)
        values.forEach(function(row) {
          var key = useConcatenatedKey ? generateKey(row) : row[0].toString().trim(); // Asegurarse de que la clave no tenga espacios en blanco
          if (!data[key]) {
            data[key] = useConcatenatedKey ? [key].concat(row.slice()) : row.slice();
          }
        });
      }
    });

    // Ordenar nombres de hojas en orden cronológico (más reciente a más antigua)
    sheetNames.sort(function(a, b) {
      var dateA = parseDate(a.replace("R_", ""));
      var dateB = parseDate(b.replace("R_", ""));
      return dateB - dateA; // Ordenar de más reciente a más antigua
    });

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
        var sheetValues = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues(); // Saltar la primera fila (encabezados)
        var found = false;
        for (var i = 0; i < sheetValues.length; i++) {
          var sheetKey = useConcatenatedKey ? generateKey(sheetValues[i]) : sheetValues[i][0].toString().trim(); // Asegurarse de que la clave no tenga espacios en blanco
          Logger.log("Comparando claves: " + key + " con " + sheetKey); // Depuración de claves
          if (sheetKey === key) {
            var valueToAdd = specificColumnToAdd === -1 ? sheetValues[i][sheetValues[i].length - 1] : sheetValues[i][specificColumnToAdd];
            row.push(valueToAdd); // Usar la columna específica o la última columna
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
// Para cambiar la columna específica a agregar, ajusta el valor de specificColumnToAdd
// Usar -1 para indicar que se debe usar la última columna
