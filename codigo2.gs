//funcion que crea hoja nueva con listado de alumnos importados y muestra nombres y correos electrónicos

function importarAlumnos(seleccion){
  if (seleccion == 0){
    Browser.msgBox('Error', 'La clase elegida no es correcta.', Browser.Buttons.OK);
  } else {
    realizandoTareas();
    var optionalArgs = { courseStates: 'ACTIVE' , teacherId: 'me' };
    var cursos = Classroom.Courses.list(optionalArgs).courses;
    for (var i=0; i<cursos.length; i++){
      if (cursos[i].name == seleccion) {
        var alumnos = Classroom.Courses.Students.list(cursos[i].id).students;
        for (var i=0; i<alumnos.length ; i++){
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Alumnos').getRange(i+2, 1).setValue(alumnos[i].profile.name.fullName); //filas empiezan en 1 y columnas en 1
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Alumnos').getRange(i+2, 2).setValue(alumnos[i].profile.emailAddress);
        }
      }
    } Browser.msgBox('Acción finalizada', 'La importación de alumnos de la clase ' + seleccion +' se ha realizado con éxito.', Browser.Buttons.OK)
  }
}


//función para crear tarea en Classroom - Formulario

function shareFormClassroom(clase,titulo,descripcion){
  var documentProperties = PropertiesService.getDocumentProperties();
  var formURL = documentProperties.getProperty('formURL');
  if (clase == '0' || titulo == '' || descripcion == '') {
    Browser.msgBox('Alguno de los datos no es correcto. Revísalos otra vez.');
  } else {
    realizandoTareas();
    var optionalArgs = { courseStates: 'ACTIVE' , teacherId: 'me' };  
    var cursos = Classroom.Courses.list(optionalArgs).courses;
    for (var i=0; i<cursos.length; i++){
      if (cursos[i].name == clase) {
        var tarea = Classroom.Courses.CourseWork.create({
            "workType": "ASSIGNMENT",
            "title": titulo,
            "description": descripcion,
            "materials": [
              {
                "link": {"url": formURL}
              }
            ], "state": "PUBLISHED"
          }, cursos[i].id);
      }
    } Browser.msgBox('La tarea se ha creado con éxito.');
  } 
}

//funnción para crear tarea en Classroom - Doc

function shareDocClassroom(clase,titulo,descripcion){
  var documentProperties = PropertiesService.getDocumentProperties();
  var docURL = documentProperties.getProperty('docURL');
  if (clase == '0' || titulo == '' || descripcion == '') {
    Browser.msgBox('Alguno de los datos no es correcto. Revísalos otra vez.');
  } else {
    realizandoTareas();
    var optionalArgs = { courseStates: 'ACTIVE' , teacherId: 'me' };  
    var cursos = Classroom.Courses.list(optionalArgs).courses;
    for (var i=0; i<cursos.length; i++){
      if (cursos[i].name == clase) {  
        var tarea = Classroom.Courses.CourseWork.create({
            "workType": "ASSIGNMENT",
            "title": titulo,
            "description": descripcion,
            "materials": [
              {
                "link": {"url": docURL}
              }
            ], "state": "PUBLISHED"
          }, cursos[i].id);
      }
    } Browser.msgBox('La tarea se ha creado con éxito.');
  } 
}

//seleccion del tipo de formulario a crear según valores introducidos

function createForm(examName,numPreg,checkAleatorio){
  if (examName == ''){
    Browser.msgBox('El título del examen está vacío. Escriba un título válido.', Browser.Buttons.OK);
  } else if (examName != '' && checkAleatorio == 'on') {
    FormAleatorio(examName,numPreg);
    } 
}

//creación del formulario con preguntas aleatorias

function FormAleatorio(examName,numPreg){
  realizandoTareas();
  var rango = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Plantilla').getDataRange();
  var valores = rango.getValues();
  var numFilas = rango.getNumRows();
  var numFilasReales = 0;

  for (var i=2; i<=numFilas; i++){
    if (valores[i][0] != '') {
      numFilasReales = numFilasReales++;
    } else {
      continue;
    }
  }
  
  var numColumnas = rango.getNumColumns();
  var numbers = [];
  const min = 1;
  if (numPreg > numFilas || numPreg <= 0) {
    Browser.msgBox('El número de preguntas elegido no es correcto.');
  } else {  
    
    var form = FormApp.create(examName).setTitle(examName).setIsQuiz(true);
    
    while(numbers.length < numPreg){
    var number = Math.round(Math.random()*(numFilas - min) + min);
    if(numbers.indexOf(number) != -1) continue;
    numbers.push(number);
    }
  Logger.log(numbers);
  for (var i = 0; i <= numbers.length; i++){
    var line = numbers[i];
    var tipoPregunta = valores[line][0];
    var titulo = valores[line][2];
    var descripcion = valores[line][3];
    var url = valores[line][4];
    var puntuacion = valores[line][9];
    var respuestas = [];
    var checks = [];
    for (var a = 11; a <= numColumnas; a+= 2){
      respuestas.push(respuestas[line][a]);
    }
    for (var b = 10; b <= numColumnas; b+= 2){
      checks.push(checks[line][b]);
    }
    if (tipoPregunta == 'RESP_MULT') {
      var choices = [];
      var item = form.addMultipleChoiceItem().setTitle(titulo).setHelpText(descripcion)
      .setFeedbackForCorrect(FormApp.createFeedback().setText(valores[line][5]).addLink(valores[line][6]).build())
      .setFeedbackForIncorrect(FormApp.createFeedback().setText(valores[line][7]).addLink(valores[line][8]).build());
      for (var k=0; k < respuestas.length; k++) {
          if (checks[k] == true && respuestas[k] != '') {
            choices.push(item.createChoice(respuestas[k],true));
          } else if (checks[k] == false && respuestas[k] != '') {
            choices.push(item.createChoice(respuestas[k],false));
          } else {
            continue;
          }
        }
      item.setChoices(choices);
      item.setPoints(puntuacion);
      choices.length = 0;
    } else if (tipoPregunta == 'SELEC_MULT') {
      var choices = [];
      var item = form.addCheckboxItem().setTitle(titulo).setHelpText(descripcion);
      for (var k=0; k < respuestas.length; k++) {
          if (checks[k] == true && respuestas[k] != '') {
            choices.push(item.createChoice(respuestas[k],true));
          } else if (checks[k] == false && respuestas[k] != '') {
            choices.push(item.createChoice(respuestas[k],false));
          } else {
            continue;
          }
        }
      item.setChoices(choices);
      item.setPoints(puntuacion);
      choices.length = 0;
    } else if (tipoPregunta == 'VERD_FALSO') {
      var choices = [];
      var item = form.addMultipleChoiceItem().setTitle(titulo).setHelpText(descripcion);
      for (var k=0; k < respuestas.length; k++) {
          if (checks[k] == true && respuestas[k] != '') {
            choices.push(item.createChoice(respuestas[k],true));
          } else if (checks[k] == false && respuestas[k] != '') {
            choices.push(item.createChoice(respuestas[k],false));
          } else {
            continue;
          }
        }
      item.setChoices(choices);
      item.setPoints(puntuacion);
      choices.length = 0;
    } else if (tipoPregunta == 'TEXTO_CORTO') { 
      var item = form.addTextItem().setTitle(titulo).setHelpText(descripcion).setPoints(puntuacion).setValidation(valores[line][4]);
    } else if (tipoPregunta == 'VIDEO') { 
      var item = form.addVideoItem().setTitle(titulo).setHelpText(descripcion).setAlignment(FormApp.Alignment.CENTER).setVideoUrl(url).setWidth(850);
    } else if (tipoPregunta == 'TEXTO_LARGO') {
      var item = form.addParagraphTextItem().setTitle(titulo).setHelpText(descripcion).setPoints(puntuacion);
    } else if (tipoPregunta == 'HORA') {
      var item = form.addTimeItem().setTitle(titulo).setHelpText(descripcion).setPoints(puntuacion);
    } else if (tipoPregunta == 'FECHA') {
      var item = form.addDateItem().setTitle(titulo).setHelpText(descripcion).setPoints(puntuacion);
    } else if (tipoPregunta == 'ESCALA') { 
      var item = form.addScaleItem().setTitle(titulo).setHelpText(descripcion).setPoints(puntuacion);
    } else {    
        SpreadsheetApp.getUi().alert('Algún tipo de pregunta no está bien seleccionado. Revísela, por favor.');
    }
   }
     var formLink = form.getPublishedUrl();
     var documentProperties = PropertiesService.getDocumentProperties();
     documentProperties.setProperty('formURL', formLink);
  }
}

//creación del documento de examen con las diferentes preguntas

function createDoc(examName) {
   
  if (examName == "") {
    Browser.msgBox('No ha indicado el título del examen.');
    } else {
    
    var doc = DocumentApp.create(examName);
    var body = doc.getBody();
    var docLink = doc.getUrl();
    var documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('docURL', docLink);
    
    var tituloTexto = body.appendParagraph(examName.toUpperCase()).setAttributes({ 
                      FONT_SIZE:16, 
                      HORIZONTAL_ALIGNMENT: DocumentApp.HorizontalAlignment.CENTER
                      });
    var cabecera = body.appendParagraph('iniciales:         curso:        número:      ').setAttributes({
                    FONT_SIZE:12,
                    HORIZONTAL_ALIGNMENT: DocumentApp.HorizontalAlignment.RIGHT
                    });
 
    var rango = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Plantilla').getDataRange();
    var valores = rango.getValues();
    var numFilas = rango.getNumRows()-1;
    var numColumnas = rango.getNumColumns();
    var contador1 = 1;
    var contador2 = 1;
    var pregTest = [];
    var pregTexto = [];
    
    for (var i=1; i<=numFilas; i++) { //cuento el número de preguntas Test marcadas
      if (valores[i][1] == true && (valores[i][0] == 'RESP_MULT' || valores[i][0] == 'SELEC_MULT' || valores[i][0] == 'VERD_FALSO'))
      pregTest.push(i);
    } pregTest = pregTest.length;
    
    for (var i=1; i<=numFilas; i++) { //cuento el número de preguntas Texto marcadas
      if (valores[i][1] == true && (valores[i][0] == 'TEXTO_CORTO' || valores[i][0] == 'TEXTO_LARGO'))
      pregTexto.push(i);
    } pregTexto = pregTexto.length;
    
    if (pregTest > 0) {
    
      for (var i=1; i<=numFilas; i++) {
        
        if (valores[i][1] == true && (valores[i][0] == 'RESP_MULT' || valores[i][0] == 'SELEC_MULT' || valores[i][0] == 'VERD_FALSO')) {
          
          var bloquesTablas = pregTest/15;
          var tabla = body.appendTable();
          
          for (var j=0; j<15; j++){
          
          var celdas = [j+1,'a','b','c','d']
          var filasTabla = tabla.appendTableRow(celdas);
          
          
          }
          
          var enunciado = valores[i][2];
          var descripcion = valores[i][3];
          var puntuacion = valores[i][5];
    } else if (valores[i][1] == true && (valores[i][0] == 'TEXTO_CORTO' || valores[i][0] == 'TEXTO_LARGO')) {
        var enunciado = valores[i][2];
        var descripcion = valores[i][3];
        var puntuacion = valores[i][5];
        
        body.appendParagraph(contador2 + '. ' + enunciado + ' ' + descripcion + ' (' + puntuacion + ' puntos.)');
        body.appendParagraph(' ');
        body.appendParagraph(' ');
        
        contador2++;
      } else {
        continue;
      }
    }
  }
}
}
