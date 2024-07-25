# Join_sheets

Editar el Código

#### Número de Columnas a Consolidar

El valor predeterminado es 9 columnas. Puedes cambiar este valor ajustando la variable `numberOfColumnsToConsolidate`:

```javascript
var numberOfColumnsToConsolidate = 9; // Cambia este valor si necesitas consolidar más o menos columnas
```
Activar la Clave Concatenada
Para activar la generación de una clave basada en la concatenación de varias columnas, cambia el valor de useConcatenatedKey a 1:

```javascript
var useConcatenatedKey = 1; // Cambia a 1 para usar la clave concatenada
```
Cambiar las Columnas para la Clave Concatenada
Por defecto, las columnas 1, 3, 5 y 7 se concatenan para formar la clave. Puedes cambiar estas columnas ajustando los índices en el array columnsToConcatenate dentro de la función generateKey:

```javascript
function generateKey(row) {
  var columnsToConcatenate = [0, 2, 4, 6]; // Cambia estos índices para concatenar diferentes columnas
  return columnsToConcatenate.map(function(index) {
    return row[index];
  }).join("_");
}
```

Funcionamiento del Código
Crear o Limpiar la Hoja Consolidado: Si la hoja "Consolidado" existe, se limpia. Si no, se crea una nueva hoja con ese nombre.
Recopilar Datos: Se recopilan los datos de las primeras numberOfColumnsToConsolidate columnas de todas las hojas cuyo nombre comience con "R_".
Generar Clave: Dependiendo del valor de useConcatenatedKey, se genera una clave basada en la primera columna o en la concatenación de varias columnas.
Ordenar Hojas: Los nombres de las hojas se ordenan en orden descendente (Z-A).
Crear Encabezados: Se crean los encabezados en la hoja "Consolidado", incluyendo los nombres de las hojas sin el prefijo "R_".
Agregar Datos: Se agregan los datos a la hoja "Consolidado", realizando un LEFT JOIN para agregar los datos de la última columna de cada hoja según la clave.
Eliminar Duplicados: Se eliminan los duplicados basándose en la clave generada.
Escribir Datos: Los datos consolidados se escriben en la hoja "Consolidado".
Ejemplo de Personalización
Cambiar el Número de Columnas a Consolidar
Si quieres consolidar solo las primeras 7 columnas en lugar de 9, cambia la variable numberOfColumnsToConsolidate:

```javascript
var numberOfColumnsToConsolidate = 7;
```

Activar y Configurar la Clave Concatenada
Si quieres usar una clave basada en la concatenación de las columnas 1, 2 y 4, activa la clave concatenada y ajusta los índices:

```javascript
var useConcatenatedKey = 1;

function generateKey(row) {
  var columnsToConcatenate = [0, 1, 3]; // Cambia estos índices para concatenar diferentes columnas
  return columnsToConcatenate.map(function(index) {
    return row[index];
  }).join("_");
}
```
Manejo de Errores
El script incluye manejo de errores para registrar y mostrar alertas en caso de problemas. Si ocurre un error, se registrará en el Logger y se mostrará una alerta en la interfaz del usuario.
```javascript
catch (e) {
  Logger.log("Error: " + e.message);
  SpreadsheetApp.getUi().alert("Se produjo un error: " + e.message);
}
```
Notas
Clave Concatenada Opcional: Para activar la generación de clave concatenada, cambia el valor de useConcatenatedKey a 1.
Cambio de Columnas para Concatenar: Para cambiar las columnas que se concatenan, ajusta los índices en el array columnsToConcatenate dentro de la función generateKey.

Número de Columnas a Consolidar: Puedes cambiar el valor de numberOfColumnsToConsolidate para ajustar el número de columnas que deseas consolidar.

