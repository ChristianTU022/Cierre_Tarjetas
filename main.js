//Funcion para Generar Alertas, Menus Personalizados, etc.
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Menú Personalizado')
      .addItem('Convertir Salida a Excel', 'convertToExcel')
      .addItem('Limpiar Hoja Close', 'confirmCleanClose')
      .addToUi();
  }

  //Funcion para Conectarse al Sheet
function conectionSheets() {
    //Conectar Sheets a AppScript
   const sheetId = '1EZPMR8xxaP860YffDhxuWr7cUaKVbYPPu8kqIoQRN-Y'; //1IfbxGR6tHOPCHc0r2oVb5R9B598clH6V5Fh5aNiZKqE Cod Original
   const sheet = SpreadsheetApp.openById(sheetId);
    //Conectar Hojas especificas
   const p_Close_Cards_Data = sheet.getSheetByName('Close_Cards_Data');

   return { sheet, p_Close_Cards_Data };
 }

function confirmCleanClose() {
//Se debe especificar hasta el numero de Columna que se desea eliminar (ultimo parametro)
    confirmAndCleanData('Close_Cards_Data', '¿Está seguro de que desea limpiar la hoja llamada "Close_Cards_Data"?', 'C');
}

function confirmAndCleanData(sheetName, confirmationMessage, lastColumn) {
    const ui = SpreadsheetApp.getUi();
    const respuesta = ui.alert(
      'Confirmación',
      confirmationMessage,
      ui.ButtonSet.YES_NO);
  
    if (respuesta == ui.Button.YES) {
      const { sheet } = conectionSheets();
      const targetSheet = sheet.getSheetByName(sheetName);
      const lastRow = targetSheet.getLastRow();
      const range = 'A2:' + lastColumn + lastRow;
      targetSheet.getRange(range).clearContent();
    }
}

function convertToExcel() {
    var hoja = SpreadsheetApp.getActiveSpreadsheet();
    var hojaSeleccionada = hoja.getSheetByName('Close_Cards_Data');
  
    // Verificar si la hoja 'CT_Output_Data' existe
    if (!hojaSeleccionada) {
      SpreadsheetApp.getUi().alert("La hoja 'Close_Cards_Data' no existe.");
      return;
    }
  
    // Obtener los datos de la hoja seleccionada
    var data = hojaSeleccionada.getDataRange().getValues();
  
    // Crear un nuevo archivo de Excel en Google Drive
    var newSpreadsheet = SpreadsheetApp.create('Close_Cards_Data');
    var newSheet = newSpreadsheet.getActiveSheet();
    newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  
    // Obtener el ID del archivo de Excel recién creado
    var fileId = newSpreadsheet.getId();
  
    // Obtener la URL de descarga del archivo de Excel
    var url = "https://docs.google.com/spreadsheets/d/" + fileId + "/export?format=xlsx";
  
    // Abrir la URL en una nueva ventana o pestaña
    var html = "<script>window.open('" + url + "');</script>";
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html), "Descargar archivo");
  }
  
  
  