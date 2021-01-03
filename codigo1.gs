//función para instalación como addon
 
function onInstall(e) {
  onOpen(e);
}

//función de ejecución del menú al iniciarse

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  if (e && e.authMode == ScriptApp.AuthMode.NONE){
      ui.createAddonMenu().addItem('Activar complemento', 'activeAddon');
  } else {
      ui.createAddonMenu().addItem('Crear interfaz', 'iniciarInterfaz')
      .addItem('Importar listado alumnos', 'importStudents').addSeparator()
      .addSubMenu(ui.createMenu('Menú Formularios')
        .addItem('Crear nuevo formulario', 'menuForm')
        .addItem('Crear tarea en Google Classroom', 'menuFormClass').addSeparator()
        .addItem('Enviar formulario por email', 'enviarFormMail'))
      .addSubMenu(ui.createMenu('Menú Documentos')
        .addItem('Crear nuevo documento', 'menuDoc')
        .addItem('Crear tarea en Google Classroom', 'menuDocClass').addSeparator()
        .addItem('Enviar documento por email', 'enviarDocMail'))
      .addToUi();
  }
}

//activar el Addon al iniciar por primera vez.

function activeAddon(){
  Browser.msgBox('autoForms ha sido activado.', Browser.Buttons.OK);
  onOpen(e);
}

//función para crear la interfaz de trabajo

function iniciarInterfaz() {
  var mensaje = Browser.msgBox('Se va a proceder a crear una nueva interfaz. Si acepta, se eliminarán todas las hojas existentes y perderá todos los datos almacenados. ¿Está de acuerdo?', Browser.Buttons.YES_NO);
  if (mensaje == 'yes'){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var numeroHojas = ss.getNumSheets();
    var hojas = ss.getSheets();
      for (var i=0; i<numeroHojas-1; i++){
        ss.deleteSheet(hojas[i]);
      }
    ss.getActiveSheet().setName('Hoja0');
    var hoja0 = ss.getSheetByName('Hoja0');
    var source = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1ukB5VJ1q4IUxPJJVuNvukJ7hGJvSxGeZemcx2y3fm3s/edit#gid=1496564602');
    source.getSheetByName('Plantilla').copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName('Plantilla');
    var hojaAlumnos = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1ukB5VJ1q4IUxPJJVuNvukJ7hGJvSxGeZemcx2y3fm3s/edit#gid=778331155');
    hojaAlumnos.getSheetByName('Alumnos').copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName('Alumnos');
    ss.deleteSheet(hoja0);  
  } else {
    Browser.msgBox('Se ha suspendido la creación de la nueva interfaz.');
  }
}

//cuadro de Diálogo para seleccionar una clase de la que importar alumnos

function importStudents(){
  var respuesta = Browser.msgBox('Se va a proceder a importar un nuevo listado de alumnos. Si ya existe una lista previa, autoForms la eliminará. ¿Desea continuar?', Browser.Buttons.YES_NO);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (respuesta == 'yes'){
    var hoja = ss.getSheetByName('Alumnos');
    var numFilas = hoja.getDataRange().getNumRows(); 
    hoja.getRange(2, 1, numFilas, 2).clearContent();
    var html = HtmlService.createTemplateFromFile('importStudents').evaluate().setHeight(130);
    return SpreadsheetApp.getUi().showModalDialog(html, 'Importar alumnos desde Google Classroom');
  } else {
    Browser.msgBox('Se ha suspendido la importación de alumnos.');
  } 
}

//menú para avisar de que el examen se está creando.

function realizandoTareas() {
  var html = HtmlService.createHtmlOutputFromFile('realizandoTareas').setWidth(400).setHeight(200);
  return SpreadsheetApp.getUi().showModelessDialog(html, 'Realizando tareas');
}

//ventana para crear el formulario de examen

function menuForm() {
  var menuForm = HtmlService.createHtmlOutputFromFile('menuForm').setHeight(250).setWidth(400);
  SpreadsheetApp.getUi().showModalDialog(menuForm, 'Nuevo formulario de examen');
}

//ventana para crear el documento de examen

function menuDoc() {
  var menuDoc = HtmlService.createHtmlOutputFromFile('menuDoc').setHeight(250).setWidth(400);
  SpreadsheetApp.getUi().showModalDialog(menuDoc, 'Nuevo documento de examen');
}

// actualizar la URL del formulario de examen

function updateForm(){  
  var documentProperties = PropertiesService.getDocumentProperties();
  var formURL = Browser.inputBox('URL nuevo formulario', 'Escribe la URL del formulario a vincular: ', Browser.Buttons.OK_CANCEL); 
  if (formURL == true && formURL != '') {     
    documentProperties.setProperty('formURL', formURL);
    Browser.msgBox('El archivo ha sido actualizado con éxito.'); 
  } Browser.msgBox('Se ha producido un problema en la actualización');
}


// actualizar la URL del documento de examen

function updateDoc(){
  var documentProperties = PropertiesService.getDocumentProperties();
  var docURL = Browser.inputBox('URL nuevo documento', 'Escribe la URL del documento a vincular: ', Browser.Buttons.OK_CANCEL);  
  if (docURL == true && docURL != '') {      
    documentProperties.setProperty('docURL', docURL);
    Browser.msgBox('El archivo ha sido actualizado con éxito.'); 
    } Browser.msgBox('Se ha producido un problema en la actualización');
}

//menuDialog para crear tarea con documento de examen en classroom

function menuDocClass() { 
  var html = HtmlService.createTemplateFromFile('comDocClass').evaluate().setTitle('Crear tarea en Classroom');
  SpreadsheetApp.getUi().showSidebar(html);
}

//menuDialog para crear tarea con formulario de examen en classroom

function menuFormClass() {
  var html = HtmlService.createTemplateFromFile('comFormClass').evaluate().setTitle('Crear tarea en Classroom');
  SpreadsheetApp.getUi().showSidebar(html);
}
