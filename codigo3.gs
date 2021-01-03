/*funcion que envía por correo electrónico el DOCUMENTO del examen

function enviarDocMail(){
  var documentProperties = PropertiesService.getDocumentProperties(); 
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  if (documentProperties.getProperty('docURL') == null) { 
    Browser.msgBox('No hay formulario que enviar. Créelo o vincule un formulario en el menú correspondiente.');
  } else if (ss.getSheetByName('Alumnos').getRange(2, 1).getValue() == '') {
    Browser.msgBox('No hay alumnos a los que enviar el documento. Importe un listado de alumnos antes de realizar esta tarea.');
  } else {
    realizandoTareas();
    var hoja = ss.getSheetByName('Alumnos');
    var rango = hoja.getDataRange();
    var valores = rango.getValues();
    var numFilas = rango.getNumRows()-1;
      for (var i=0; i<numFilas; i++) {
        var email = GmailApp.sendEmail(valores[i+1][1], 'Enlace a documento de examen', 'Aquí tienes el enlace al documento de examen: ' + documentProperties.getProperty('docURL'));
      } 
  }
}


//funcion que envía por correo electrónico el FORMULARIO del examen

function enviarFormMail(){
  var documentProperties = PropertiesService.getDocumentProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (documentProperties.getProperty('formURL') == null) { 
    Browser.msgBox('No hay formulario que enviar. Créelo o vincule un formulario en el menú correspondiente.');
  } else if (ss.getSheetByName('Alumnos').getRange(2, 1).getValue() == ''){
    Browser.msgBox('No hay alumnos a los que enviar el formulario. Importe un listado de alumnos antes de realizar esta tarea.');
  } else {  
    realizandoTareas();  
    var hoja = ss.getSheetByName('Alumnos');
    var rango = hoja.getDataRange();
    var valores = rango.getValues();
    var numFilas = rango.getNumRows()-1;   
      for (var i=0; i<numFilas; i++) {        
          GmailApp.sendEmail(valores[i+1][1], 'Enlace a formulario de examen.', 'Aquí tienes el enlace al formulario de examen: ' + documentProperties.getProperty('formURL'));      
      } Browser.msgBox('Los correos han sido enviados.'); 
  } 
}

*/
