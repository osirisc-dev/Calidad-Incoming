var CONFIG = {
  SPREADSHEET_ID: "TU_ID_DEL_SHEET_AQUI", 
  CARPETA_FOTOS_ID: "TU_ID_DE_CARPETA_DRIVE_AQUI",
  TITULO_APP: "Reporte Incoming - [NOMBRE_DE_TU_EMPRESA]",
  HOJAS: {
    MAQUINAS: "BaseMaquinas",
    PENDIENTES: "RefaccionesPendientes",
    RESPUESTAS: "RespuestasRefacciones"
  }
};
// ==============================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle(CONFIG.TITULO_APP)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function obtenerMaquinas() {
  var sheetMaquinas = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.HOJAS.MAQUINAS);
  var datos = sheetMaquinas.getDataRange().getValues();
  var catalogo = {};
  
  for (var i = 1; i < datos.length; i++) {
    var maquina = datos[i][0]; var pieza = datos[i][2]; var criterio = datos[i][4];
    if (maquina) {
      if (!catalogo[maquina]) { catalogo[maquina] = []; }
      catalogo[maquina].push({ pieza: pieza, criterio: criterio });
    }
  }
  return catalogo;
}

function guardarDatos(datosForm) {
  try {
    var sheetDestino = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheets()[0];
    var carpeta = DriveApp.getFolderById(CONFIG.CARPETA_FOTOS_ID);
    var listadoResultados = [];

    for (var i = 0; i < datosForm.evaluaciones.length; i++) {
      var ev = datosForm.evaluaciones[i];
      var textoPieza = ev.pieza + ": " + ev.seleccion;
      if (ev.comentario) { textoPieza += " (Nota: " + ev.comentario + ")"; }

      if (ev.fotoBase64) {
        var splitBase = ev.fotoBase64.split(',');
        var tipo = splitBase[0].split(';')[0].replace('data:', '');
        var byteCharacters = Utilities.base64Decode(splitBase[1]);
        var blob = Utilities.newBlob(byteCharacters, tipo, ev.fotoNombre);
        var archivo = carpeta.createFile(blob);
        archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        textoPieza += " [Foto: " + archivo.getUrl() + "]";
        
        ev.fotoUrlDirecta = archivo.getUrl(); 
      }
      listadoResultados.push(textoPieza);
    }

    var textoResultadosFinal = listadoResultados.join(" | ");
    
    sheetDestino.appendRow([
      new Date(), 
      datosForm.correo, 
      datosForm.serie, 
      datosForm.contenedor,
      datosForm.referencia, 
      datosForm.nombreMaquina,
      textoResultadosFinal,
      datosForm.disposicion,
      datosForm.comentariosGenerales 
    ]);
    
    generarReporteIndividualAutomatico(datosForm);
    
    return "Éxito";
  } catch(error) {
    return "Error: " + error.toString();
  }
}

function obtenerRefacciones() {
  var sheetRefacciones = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.HOJAS.PENDIENTES);
  
  if (!sheetRefacciones) return [];
  
  var datos = sheetRefacciones.getDataRange().getValues();
  var lista = [];
  
  for (var i = 1; i < datos.length; i++) {
    var arCode = datos[i][0]; 
    var erpCode = datos[i][1]; 
    var descripcion = datos[i][2];
    var ordenado = datos[i][3];
    var enviado = datos[i][4];
    
    // Validamos que al menos exista el código ERP o AR
    if (arCode || erpCode) { 
      lista.push({ 
        arCode: arCode ? arCode.toString().trim() : "N/A", 
        erpCode: erpCode ? erpCode.toString().trim() : "N/A",
        descripcion: descripcion ? descripcion.toString().trim() : "Sin descripción",
        ordenado: ordenado ? ordenado.toString().trim() : "0",
        enviado: enviado ? enviado.toString().trim() : "0" // Esto es lo que mandaremos al PDF
      });
    }
  }
  
  return lista;
}

function guardarRefacciones(datosForm) {
  try {
    var sheetDestino = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.HOJAS.RESPUESTAS);
    var carpeta = DriveApp.getFolderById(CONFIG.CARPETA_FOTOS_ID);
    var listadoResultados = [];

    for (var i = 0; i < datosForm.evaluaciones.length; i++) {
      var ev = datosForm.evaluaciones[i];
      var textoPieza = ev.pieza + " [Cant: " + ev.cantidad + "] [Disp: " + ev.disposicionItem + "]: " + ev.seleccion;
      if (ev.comentario) { textoPieza += " (Nota: " + ev.comentario + ")"; }

      if (ev.fotoBase64) {
        var splitBase = ev.fotoBase64.split(',');
        var tipo = splitBase[0].split(';')[0].replace('data:', '');
        var byteCharacters = Utilities.base64Decode(splitBase[1]);
        var blob = Utilities.newBlob(byteCharacters, tipo, ev.fotoNombre);
        var archivo = carpeta.createFile(blob);
        archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        textoPieza += " [Foto: " + archivo.getUrl() + "]";
        ev.fotoUrlDirecta = archivo.getUrl(); 
      }
      listadoResultados.push(textoPieza);
    }

    var textoResultadosFinal = listadoResultados.join(" | ");
    
    sheetDestino.appendRow([
      new Date(), 
      datosForm.correo, 
      datosForm.contenedor,
      datosForm.referencia, 
      textoResultadosFinal
    ]);
    
    if (datosForm.codigoRefaccion) {
        var sheetPendientes = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.HOJAS.PENDIENTES);
        if (sheetPendientes) {
            var datosPendientes = sheetPendientes.getDataRange().getValues();
            for (var j = datosPendientes.length - 1; j >= 1; j--) {
                if (datosPendientes[j][1].toString().trim() === datosForm.codigoRefaccion.toString().trim()) {
                    sheetPendientes.deleteRow(j + 1); 
                    break; 
                }
            }
        }
    }
    
    return "Éxito";
  } catch(error) {
    return "Error: " + error.toString();
  }
}