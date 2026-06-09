var CONFIG_REPORTES = {
  EMPRESA_NOMBRE: "[NOMBRE_DE_TU_EMPRESA]",
  NOMBRES: {
    AUTORIZACION: "[NOMBRE_DEL_JEFE_O_AUTORIZADOR]",
    AUDITOR_DEFAULT: "[NOMBRE_DEL_AUDITOR_POR_DEFECTO]"
  },
  IDS: {
    LOGO: "TU_ID_DE_LOGO_DRIVE_AQUI",
    FIRMA_JEFE: "TU_ID_FIRMA_JEFE_DRIVE_AQUI",
    CARPETA_CONSOLIDADOS: "TU_ID_CARPETA_CONSOLIDADOS_AQUI",
    CARPETA_INDIVIDUALES: "TU_ID_CARPETA_INDIVIDUALES_AQUI"
  },
  INSPECTORES: {
    "inspector1@tudominio.com": "ID_FIRMA_1_AQUI",
    "inspector2@tudominio.com": "ID_FIRMA_2_AQUI"
  }
};
// ===================================

function traducirTexto(texto) {
  if (!texto || texto.toString().trim() === "") return "";
  
  var tBajo = texto.toString().toLowerCase().trim();
  if (tBajo === "sin observaciones" || tBajo === "sin comentarios" || tBajo === "sin comentarios generales registrados.") {
    return "";
  }
  
  try {
    return LanguageApp.translate(texto.toString().trim(), 'es', 'en');
  } catch (e) {
    Logger.log("Error en traducción automática: " + e.toString());
    return ""; // Si falla por red, devuelve vacío para no romper el reporte
  }
}

function generarReporteConsolidado() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaPanel = libro.getSheetByName("Panel");
  var hojaRespuestas = libro.getSheets()[0]; 
  var hojaRefacciones = libro.getSheetByName("RespuestasRefacciones");
  
  var refBuscada = hojaPanel.getRange("B1").getValue().toString().trim();
  var contenedor = hojaPanel.getRange("B2").getValue().toString().trim();
  var proveedor  = hojaPanel.getRange("B3").getValue().toString().trim();
  var factura    = hojaPanel.getRange("B4").getValue().toString().trim();
  
  if (refBuscada === "") {
    SpreadsheetApp.getUi().alert("⚠️ Por favor escribe una Referencia en la celda B1.");
    return;
  }

  var hojaMaquinas = libro.getSheetByName("BaseMaquinas");
  var datosMaq = hojaMaquinas.getDataRange().getValues();
  var diccionarioModelos = {};
  var diccionarioBilingue = {};

  for (var m = 1; m < datosMaq.length; m++) {
    var maqCorto = datosMaq[m][0] ? datosMaq[m][0].toString().trim().toUpperCase() : "";
    var maqTecnico = datosMaq[m][1] ? datosMaq[m][1].toString().trim() : "";
    var piezaEsp = datosMaq[m][2] ? datosMaq[m][2].toString().trim() : "";
    var piezaIng = datosMaq[m][3] ? datosMaq[m][3].toString().trim() : "";
    if (maqCorto) { diccionarioModelos[maqCorto] = maqTecnico ? maqTecnico : maqCorto; }
    if (piezaEsp && piezaIng) { diccionarioBilingue[piezaEsp] = piezaIng; }
  }

  var imgLogoHtml = ""; try { var fL = DriveApp.getFileById(CONFIG_REPORTES.IDS.LOGO); imgLogoHtml = "<img src='data:" + fL.getMimeType() + ";base64," + Utilities.base64Encode(fL.getBlob().getBytes()) + "' style='max-height:40px;' />"; } catch(e){}
  var imgJefeHtml = "<div style='height:100px;'></div>"; try { var fJ = DriveApp.getFileById(CONFIG_REPORTES.IDS.FIRMA_JEFE); imgJefeHtml = "<img src='data:" + fJ.getMimeType() + ";base64," + Utilities.base64Encode(fJ.getBlob().getBytes()) + "' style='max-height:100px; display:block; margin:0 auto 4px auto;' />"; } catch(e){}

  var datos = hojaRespuestas.getDataRange().getValues();
  var maquinasProcesadas = [];
  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    if (fila[4] && fila[4].toString().trim() === refBuscada) {
      
      var resultadosStr = fila[6] ? fila[6].toString().trim() : "";
      var bloquesPiezas = resultadosStr.split(" | ");
      var htmlFilasComponentes = "";

      for (var p = 0; p < bloquesPiezas.length; p++) {
        var bloque = bloquesPiezas[p].trim();
        if (!bloque) continue;

        var splitIndex = bloque.indexOf(":");
        if (splitIndex === -1) continue;

        var pieza = bloque.substring(0, splitIndex).trim();
        var detalles = bloque.substring(splitIndex + 1).trim();

        var notaLimpia = "";
        var matchNota = detalles.match(/\(Nota:\s*(.*?)\)/);
        if (matchNota) { notaLimpia = matchNota[1]; detalles = detalles.replace(matchNota[0], "").trim(); }

        var urlFoto = "";
        var matchFoto = detalles.match(/\[Foto:\s*(.*?)\]/);
        if (matchFoto) { urlFoto = matchFoto[1]; detalles = detalles.replace(matchFoto[0], "").trim(); }

        var statusTexto = detalles;
        var compBadgeStyle = "border: 1px solid #393D42; color: #393D42;";
        if (detalles.toLowerCase().includes("no cumple")) {
            statusTexto = "No Cumple / Fail";
            compBadgeStyle = "background-color: #E60404; color: #FFFFFF; border: 1px solid #E60404;";
        } else if (detalles.toLowerCase().includes("cumple")) {
            statusTexto = "Cumple / Pass";
        }

        var celdaEvidencia = "";
        if (notaLimpia) { 
          var notaLimpiaEn = traducirTexto(notaLimpia);
          var textoNotaBilingue = notaLimpia + (notaLimpiaEn ? " <br><i style='color:#718096;'>(" + notaLimpiaEn + ")</i>" : "");
          celdaEvidencia += "<div style='font-size:11px; margin-bottom:4px;'>" + textoNotaBilingue + "</div>"; 
        }

        if (urlFoto) {
            var matchId = urlFoto.match(/id=([^&]+)/) || urlFoto.match(/\/d\/([^/]+)/);
            if (matchId && matchId[1]) {
                try {
                    var fileF = DriveApp.getFileById(matchId[1]);
                    celdaEvidencia += "<div style='margin-top:4px;'><img src='data:" + fileF.getMimeType() + ";base64," + Utilities.base64Encode(fileF.getBlob().getBytes()) + "' style='max-height:105px; border:1px solid #CBD5E0; border-radius:4px;' /></div>";
                } catch(err) {
                    celdaEvidencia += "<div style='margin-top:4px;'><a href='" + urlFoto + "' target='_blank' style='font-size:10px; color:#E60404; font-weight:bold;'>Ver Foto Evidencia ↗</a></div>";
                }
            }
        }
        
        var transHeader = diccionarioBilingue[pieza];
        var headerFinal = transHeader ? (pieza + " / " + transHeader) : pieza;

        htmlFilasComponentes += "<tr>" +
          "<td style='padding:6px 10px; font-size:11px; border-bottom:1px solid #E2E8F0; color:#2D3748; font-weight:600;'>" + headerFinal + "</td>" +
          "<td style='padding:6px 25px; font-size:11px; border-bottom:1px solid #E2E8F0; text-align:center;'><span style='display:inline-block; padding:2px 8px; font-weight:700; border-radius:4px; font-size:9px; text-transform:uppercase; " + compBadgeStyle + "'>" + statusTexto + "</span></td>" +
          "<td style='padding:6px 10px 6px 35px; font-size:11px; border-bottom:1px solid #E2E8F0; background-color:#fafafa;'>" + celdaEvidencia + "</td></tr>";
      }

      var dispValInd = fila[7] ? fila[7].toString().trim() : "No Evaluado";
      var indDispBadgeStyle = "border: 1px solid #393D42; color: #393D42;";
      var indDispBilingue = dispValInd;
      if (dispValInd.toLowerCase().includes("desviacion") || dispValInd.toLowerCase().includes("desviación")) {
        indDispBadgeStyle = "background-color: #EF6C00; color: #FFFFFF; border: 1px solid #EF6C00;"; indDispBilingue = "Aceptado con Desviación / Accepted with Deviation";
      } else if (dispValInd.toLowerCase().includes("aceptado")) {
        indDispBadgeStyle = "background-color: #2E7D32; color: #FFFFFF; border: 1px solid #2E7D32;"; indDispBilingue = "Aceptado / Accepted";
      } else if (dispValInd.toLowerCase().includes("rechazado")) {
        indDispBadgeStyle = "background-color: #E60404; color: #FFFFFF; border: 1px solid #E60404;"; indDispBilingue = "Rechazado / Rejected";
      }

      maquinasProcesadas.push({
        correo: fila[1], serie: fila[2], contenedor: fila[3], referencia: fila[4], modelo: fila[5],
        htmlComponentes: htmlFilasComponentes, dispStyle: indDispBadgeStyle, dispBilingue: indDispBilingue,
        comentariosGenerales: fila[8] ? fila[8].toString().trim() : ""
      });
    }
  }

  var refaccionesProcesadas = [];
  if (hojaRefacciones) {
    var datosRef = hojaRefacciones.getDataRange().getValues();
    for (var r = 1; r < datosRef.length; r++) {
      var filaRef = datosRef[r];
      if (filaRef[3] && filaRef[3].toString().trim() === refBuscada) {
        refaccionesProcesadas.push({
          correo: filaRef[1], contenedor: filaRef[2], referencia: filaRef[3],
          resultados: filaRef[4] ? filaRef[4].toString().trim() : "",
          comentariosGenerales: filaRef[5] ? filaRef[5].toString().trim() : "",
          disposicion: filaRef[6] ? filaRef[6].toString().trim() : ""
        });
      }
    }
  }
  
  if (maquinasProcesadas.length === 0 && refaccionesProcesadas.length === 0) {
    SpreadsheetApp.getUi().alert("❌ No se encontraron registros de Máquinas ni Refacciones para la Referencia: " + refBuscada);
    return;
  }

  var fechaHoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  
  var htmlMaster = "<!DOCTYPE html><html><head><meta charset='utf-8'></head><body style='font-family:Arial, sans-serif; color:#393D42; background-color:#FFFFFF; line-height:1.4; margin:0; padding:0;'>" +
    "<div style='padding: 20px 25px; page-break-inside: avoid;'>" +
      "<table style='width:100%; border-collapse:collapse; margin-bottom:16px;'>" +
        "<tr><td><h1 style='font-family:\"Arial Black\", sans-serif; font-size:30px; font-weight:900; color:#E60404; margin:0;'>" + CONFIG_REPORTES.EMPRESA_NOMBRE + "</h1><div style='font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:1px; margin-top:2px;'>Reporte Consolidado de Inspección — CONSOLIDATED REPORT</div></td><td style='text-align:right;'>" + imgLogoHtml + "</td></tr>" +
      "</table>" +
      "<div style='width:100%; padding:8px 10px; background-color:#F8F9FA; border-top:2px solid #393D42; border-bottom:1px solid #E2E8F0; margin-bottom:20px; font-size:10px;'>" +
        "<table style='width:100%; font-size:10px; border-collapse:collapse;'>" +
          "<tr>" +
            "<td style='width:40%; padding-bottom:4px;'><b>AUDITOR / AUDITOR:</b> <span style='text-transform:uppercase;'>" + CONFIG_REPORTES.NOMBRES.AUDITOR_DEFAULT + "</span></td>" +
            "<td style='width:30%; padding-bottom:4px;'><b>FECHA / DATE:</b> " + fechaHoy + "</td>" +
            "<td style='width:30%; padding-bottom:4px;'><b>REFERENCIA / REFERENCE:</b> <span style='color:#E60404; font-weight:bold;'>" + refBuscada + "</span></td>" +
          "</tr>" +
          "<tr>" +
            "<td><b>CONTENEDOR / CONTAINER:</b> " + contenedor + "</td>" +
            "<td><b>PROVEEDOR / SUPPLIER:</b> <span style='text-transform:uppercase;'>" + proveedor + "</span></td>" +
            "<td><b>FACTURA / INVOICE NO:</b> " + factura + "</td>" +
          "</tr>" +
        "</table>" +
      "</div>";

  if (maquinasProcesadas.length > 0) {
      var htmlFilasResumen = "";
      for (var k = 0; k < maquinasProcesadas.length; k++) {
        var maq = maquinasProcesadas[k];
        var modeloOficial = diccionarioModelos[maq.modelo.toUpperCase()] || maq.modelo;
        htmlFilasResumen += "<tr>" +
          "<td style='padding:6px; font-size:11px; border-bottom:1px solid #E2E8F0; text-align:center;'>" + maq.referencia + "</td>" +
          "<td style='padding:6px; font-size:9px; border-bottom:1px solid #E2E8F0; text-transform:uppercase;'>" + modeloOficial + "</td>" +
          "<td style='padding:6px; font-size:11px; border-bottom:1px solid #E2E8F0; text-align:center; font-weight:bold;'>" + maq.serie + "</td>" +
          "<td style='padding:6px; font-size:11px; border-bottom:1px solid #E2E8F0; text-align:center;'><span style='display:inline-block; padding:2px 8px; font-weight:700; border-radius:4px; font-size:9px; text-transform:uppercase; " + maq.dispStyle + "'>" + maq.dispBilingue + "</span></td>" +
        "</tr>";
      }
      htmlMaster += "<div style='font-size:14px; font-weight:700; text-transform:uppercase; margin-bottom:6px; padding-bottom:4px; border-bottom:2px solid #E60404;'>Resumen de Equipos / Equipment Summary</div>" +
          "<table style='width:100%; border-collapse:collapse; margin-bottom:16px;'>" +
            "<thead><tr style='background-color:#393D42; color:#FFFFFF; font-size:10px; text-transform:uppercase;'><th style='padding:6px; text-align:center; width:20%;'>REF / Box</th><th style='padding:6px; text-align:left; width:40%;'>Descripción / Description</th><th style='padding:6px; text-align:center; width:20%;'>Serie / S/N</th><th style='padding:6px; text-align:center; width:20%;'>Disposición / Provision</th></tr></thead>" +
            "<tbody>" + htmlFilasResumen + "</tbody>" +
          "</table>";
  }

  if (refaccionesProcesadas.length > 0) {
      var htmlFilasRefacciones = "";
      for(var rf = 0; rf < refaccionesProcesadas.length; rf++) {
          var rep = refaccionesProcesadas[rf];
          var partesArray = rep.resultados.split(" | ");
          
          for (var pa = 0; pa < partesArray.length; pa++) {
              var bloqueTemp = partesArray[pa].trim();
              if(!bloqueTemp) continue;

              var numParte = "N/A", descripcion = "N/A", cant = "N/A", dispItem = "N/A", estatus = "", obs = "Sin observaciones", fotoUrl = "";

              var matchFoto = bloqueTemp.match(/\[Foto:\s*([\s\S]*?)\]$/);
              if (matchFoto) { fotoUrl = matchFoto[1].trim(); bloqueTemp = bloqueTemp.replace(matchFoto[0], "").trim(); }

              var matchNota = bloqueTemp.match(/\(Nota:\s*([\s\S]*?)\)$/);
              if (matchNota) { obs = matchNota[1].trim(); bloqueTemp = bloqueTemp.replace(matchNota[0], "").trim(); }

              var idxCant = bloqueTemp.indexOf("[Cant:");
              var idxDisp = bloqueTemp.indexOf("[Disp:");
              var idxEndDisp = bloqueTemp.indexOf("]:");

              if (idxCant !== -1 && idxDisp !== -1 && idxEndDisp !== -1) {
                  var tituloCompleto = bloqueTemp.substring(0, idxCant).trim();
                  cant = bloqueTemp.substring(idxCant + 6, bloqueTemp.indexOf("]", idxCant)).trim();
                  dispItem = bloqueTemp.substring(idxDisp + 6, idxEndDisp).trim();
                  estatus = bloqueTemp.substring(idxEndDisp + 2).trim();

                  var idxGuion = tituloCompleto.indexOf(" - ");
                  if (idxGuion !== -1) {
                      numParte = tituloCompleto.substring(0, idxGuion).replace("NP:", "").trim();
                      descripcion = tituloCompleto.substring(idxGuion + 3).trim();
                  } else {
                      numParte = tituloCompleto.replace("NP:", "").trim();
                      descripcion = "-";
                  }
              } else {
                  descripcion = bloqueTemp;
              }

              var badgeEstatus = estatus.toLowerCase().includes("no cumple") ? "color:#E60404;" : "color:#2E7D32;";
              var estatusBilingue = estatus.toLowerCase().includes("no cumple") ? "No cumple / Fail" : "Cumple / Pass";
              
              var badgeDisp = "color:#393D42;";
              var dispBilingue = dispItem;
              if (dispItem.toLowerCase().includes("rechazado")) { badgeDisp = "color:#E60404;"; dispBilingue = "Rechazado / Rejected"; }
              else if (dispItem.toLowerCase().includes("desviacion")) { badgeDisp = "color:#EF6C00;"; dispBilingue = "Aceptado con Desviación / Accpt. w/ Deviation"; }
              else if (dispItem.toLowerCase().includes("aceptado")) { badgeDisp = "color:#2E7D32;"; dispBilingue = "Aceptado / Accepted"; }

              var obsEn = traducirTexto(obs);
              var textoObsFinal = obs + (obsEn && obs.toLowerCase() !== "sin observaciones" ? " <br><i style='color:#718096;'>(" + obsEn + ")</i>" : "");

              var htmlEvidencia = "<span style='color:#4A5568;'>" + textoObsFinal + "</span>";
              if (fotoUrl) {
                  var matchId = fotoUrl.match(/id=([^&]+)/) || fotoUrl.match(/\/d\/([^/]+)/);
                  if (matchId && matchId[1]) {
                      try {
                          var fileF = DriveApp.getFileById(matchId[1]);
                          htmlEvidencia += "<br><img src='data:" + fileF.getMimeType() + ";base64," + Utilities.base64Encode(fileF.getBlob().getBytes()) + "' style='max-height:80px; border:1px solid #CBD5E0; border-radius:4px; margin-top:6px;' />";
                      } catch(err) {
                          htmlEvidencia += "<br><a href='" + fotoUrl + "' target='_blank' style='color:#3182ce; font-size:9px; text-decoration:none;'>📸 Ver Evidencia</a>";
                      }
                  }
              }

              htmlFilasRefacciones += "<tr>" +
                  "<td style='padding:6px; font-size:9px; border-bottom:1px solid #E2E8F0; color:#2D3748; font-weight:bold;'>" + numParte + "</td>" +
                  "<td style='padding:6px; font-size:9px; border-bottom:1px solid #E2E8F0; color:#4A5568;'>" + descripcion + "</td>" +
                  "<td style='padding:6px; font-size:10px; border-bottom:1px solid #E2E8F0; text-align:center; font-weight:bold;'>" + cant + "</td>" +
                  "<td style='padding:6px; font-size:10px; border-bottom:1px solid #E2E8F0; text-align:center; font-weight:bold; " + badgeEstatus + "'>" + estatusBilingue + "</td>" +
                  "<td style='padding:6px; font-size:10px; border-bottom:1px solid #E2E8F0; text-align:center; font-weight:bold; " + badgeDisp + "'>" + dispBilingue + "</td>" +
                  "<td style='padding:6px; font-size:9px; border-bottom:1px solid #E2E8F0;'>" + htmlEvidencia + "</td>" +
              "</tr>";
          }
      }

      htmlMaster += "<div style='font-size:14px; font-weight:700; text-transform:uppercase; margin-top:25px; margin-bottom:6px; padding-bottom:4px; border-bottom:2px solid #E60404;'>COMPONENTES Y ACCESORIOS / COMPONENTS AND ACCESSORIES</div>" +
          "<table style='width:100%; border-collapse:collapse; margin-bottom:16px;'>" +
            "<thead><tr style='background-color:#393D42; color:#FFFFFF; font-size:9px; text-transform:uppercase;'>" +
              "<th style='padding:6px; text-align:left; width:15%;'>Num. Parte / Part Number</th>" +
              "<th style='padding:6px; text-align:left; width:25%;'>Descripción / Description</th>" +
              "<th style='padding:6px; text-align:center; width:8%;'>Cant / Quantity.</th>" +
              "<th style='padding:6px; text-align:center; width:12%;'>Estatus / Status</th>" +
              "<th style='padding:6px; text-align:center; width:15%;'>Disposición / Provision</th>" +
              "<th style='padding:6px; text-align:left; width:25%;'>Observaciones / Observations</th>" +
            "</tr></thead>" +
            "<tbody>" + htmlFilasRefacciones + "</tbody>" +
          "</table>";
  }

  htmlMaster += "</div>";

  for (var x = 0; x < maquinasProcesadas.length; x++) {
    var maq = maquinasProcesadas[x];
    var emailInsp = maq.correo; 
    
    var nombreComercialInd = maq.modelo.toUpperCase();
    var modeloOficialInd = diccionarioModelos[nombreComercialInd] ? diccionarioModelos[nombreComercialInd] : maq.modelo;

    var inspectorNombre = "Inspector";
    if (emailInsp) {
      var partes = emailInsp.split('@')[0].split('.');
      if (partes.length >= 2) { inspectorNombre = partes[0].charAt(0).toUpperCase() + partes[0].slice(1) + " " + partes[1].charAt(0).toUpperCase() + partes[1].slice(1); } 
      else { inspectorNombre = emailInsp.split('@')[0]; }
    }
    var imgInspectorHtml = "<div style='height:100px;'></div>";
    var idFirmaInsp = CONFIG_REPORTES.INSPECTORES[emailInsp.toLowerCase()];
    if (idFirmaInsp) {
      try { var fI = DriveApp.getFileById(idFirmaInsp); imgInspectorHtml = "<img src='data:" + fI.getMimeType() + ";base64," + Utilities.base64Encode(fI.getBlob().getBytes()) + "' style='max-height:100px; display:block; margin:0 auto 4px auto;' />"; } catch(e){}
    }

    var comentariosGeneralesEn = traducirTexto(maq.comentariosGenerales);
    var htmlComentariosGenerales = "<span style='color:#a0aec0; font-style:italic;'>Sin comentarios generales registrados. / No general comments registered.</span>";
    if (maq.comentariosGenerales) {
      htmlComentariosGenerales = maq.comentariosGenerales + (comentariosGeneralesEn ? "<br><span style='color:#718096; font-style:italic; display:block; margin-top:4px;'>(" + comentariosGeneralesEn + ")</span>" : "");
    }

    htmlMaster += "<div style='page-break-before: always; padding: 20px 25px;'>" +
      "<table style='width:100%; border-collapse:collapse; margin-bottom:16px;'>" +
        "<tr><td><h1 style='font-family:\"Arial Black\", sans-serif; font-size:30px; font-weight:900; color:#E60404; margin:0;'>" + CONFIG_REPORTES.EMPRESA_NOMBRE + "</h1><div style='font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:1px; margin-top:2px;'>Reporte de Control de Calidad — INCOMING</div></td><td style='text-align:right;'>" + imgLogoHtml + "</td></tr>" +
      "</table>" +
      "<div style='width:100%; padding:8px 10px; background-color:#F8F9FA; border-top:2px solid #393D42; border-bottom:1px solid #E2E8F0; margin-bottom:12px; font-size:10px;'>" +
        "<div style='margin-bottom:6px;'><b>MÁQUINA/MODELO / MACHINE/MODEL:</b> <span style='color:#E60404; font-weight:bold;'>" + modeloOficialInd + " / " + maq.modelo + "</span></div>" +
        "<table style='width:100%; font-size:10px; border-collapse:collapse;'>" +
          "<tr>" +
            "<td style='width:40%; padding-bottom:4px;'><b>NÚMERO SERIE / SERIAL NUMBER:</b> " + maq.serie + "</td>" +
            "<td style='width:30%; padding-bottom:4px;'><b>FECHA / DATE:</b> " + fechaHoy + "</td>" +
            "<td style='width:30%; padding-bottom:4px;'><b>CONTENEDOR / CONTAINER:</b> " + maq.contenedor + "</td>" +
          "</tr>" +
          "<tr>" +
            "<td><b>REFERENCIA / REFERENCE:</b> " + maq.referencia + "</td>" +
            "<td colspan='2'><b>INSPECTOR / INSPECTOR:</b> " + emailInsp + "</td>" +
          "</tr>" +
        "</table>" +
      "</div>" +
      "<div style='font-size:15px; font-weight:700; text-transform:uppercase; margin:12px 0 6px 0; padding-bottom:4px; border-bottom:2px solid #E60404;'>1. Evaluación de Componentes / Components Evaluation</div>" +
      "<table style='width:100%; border-collapse:collapse; margin-bottom:16px;'>" +
        "<thead><tr style='background-color:#393D42; color:#FFFFFF; font-size:10px; text-transform:uppercase;'><th style='padding:6px 10px; text-align:left; width:35%;'>Componente Evaluado / Evaluated Component</th><th style='padding:6px 10px; text-align:center; width:20%;'>Estatus / Status</th><th style='padding:6px 10px 6px 35px; text-align:left; width:45%;'>Notas y Evidencias / Notes & Photos</th></tr></thead>" +
        "<tbody>" + maq.htmlComponentes + "</tbody>" +
      "</table>" +
      "<div style='margin:16px 0; padding:10px; background-color:#f8fafc; border:1px solid #e2e8f0; border-radius:6px; font-size:12px; font-weight:bold; color:#393D42; text-align:right;'>DISPOSICIÓN FINAL | FINAL DISPOSITION: <span style='display:inline-block; padding:4px 12px; font-weight:700; border-radius:4px; font-size:11px; text-transform:uppercase; " + maq.dispStyle + "'>" + maq.dispBilingue + "</span></div>" +
      "<div style='font-size:14px; font-weight:700; text-transform:uppercase; margin:16px 0 6px 0; padding-bottom:4px; border-bottom:2px solid #393D42;'>2. Comentarios Generales del Equipo / General Machine Comments</div>" +
      "<div style='background-color:#F8F9FA; border-left:4px solid #E60404; padding:10px; font-size:11px; color:#2D3748; min-height:25px; margin-bottom:16px;'>" + 
        htmlComentariosGenerales + 
      "</div>" +
      "<table style='width:100%; border-collapse:collapse; margin-top:30px; page-break-inside: avoid;'>" +
        "<tr>" +
          "<td style='width:50%; text-align:center; vertical-align:bottom; padding:0 30px;'>" + imgInspectorHtml + "<div style='width:100%; border-top:1px solid #393D42; margin-bottom:4px;'></div><div style='font-size:11px; font-weight:700; text-transform:uppercase;'>" + inspectorNombre + "</div><div style='font-size:9px; color:#718096;'>Inspector de Calidad / Quality Inspector</div></td>" +
          "<td style='width:50%; text-align:center; vertical-align:bottom; padding:0 30px;'>" + imgJefeHtml + "<div style='width:100%; border-top:1px solid #393D42; margin-bottom:4px;'></div><div style='font-size:11px; font-weight:700; color:#1A202C; text-transform:uppercase;'>" + CONFIG_REPORTES.NOMBRES.AUTORIZACION + "</div><div style='font-size:9px; color:#718096;'>Autorización / Authorized Signature</div></td>" +
        "</tr>" +
      "</table>" +
    "</div>";
  }

  htmlMaster += "</body></html>";

  var finalBlob = Utilities.newBlob(htmlMaster, "text/html", "Temporal.html").getAs(MimeType.PDF);
  finalBlob.setName("REPORTE_COMPLETO_" + refBuscada + ".pdf");
  
  var carpetaDestino = DriveApp.getFolderById(CONFIG_REPORTES.IDS.CARPETA_CONSOLIDADOS);
  carpetaDestino.createFile(finalBlob);

  SpreadsheetApp.getUi().alert("✅ ¡Reporte creado con éxito!");
}


function generarReporteIndividualAutomatico(datosForm) {
  try {
    var libro = SpreadsheetApp.getActiveSpreadsheet();
    var hojaMaquinas = libro.getSheetByName("BaseMaquinas");
    var datosMaq = hojaMaquinas.getDataRange().getValues();
    
    var diccionarioModelos = {};
    var diccionarioBilingue = {};

    for (var m = 1; m < datosMaq.length; m++) {
      var maqCorto = datosMaq[m][0] ? datosMaq[m][0].toString().trim().toUpperCase() : "";
      var maqTecnico = datosMaq[m][1] ? datosMaq[m][1].toString().trim() : "";
      var piezaEsp = datosMaq[m][2] ? datosMaq[m][2].toString().trim() : "";
      var piezaIng = datosMaq[m][3] ? datosMaq[m][3].toString().trim() : "";
      
      if (maqCorto) { diccionarioModelos[maqCorto] = maqTecnico ? maqTecnico : maqCorto; }
      if (piezaEsp && piezaIng) { diccionarioBilingue[piezaEsp] = piezaIng; }
    }

    var imgLogoHtml = ""; try { var fL = DriveApp.getFileById(CONFIG_REPORTES.IDS.LOGO); imgLogoHtml = "<img src='data:" + fL.getMimeType() + ";base64," + Utilities.base64Encode(fL.getBlob().getBytes()) + "' style='max-height:40px;' />"; } catch(e){}
    var imgJefeHtml = "<div style='height:100px;'></div>"; try { var fJ = DriveApp.getFileById(CONFIG_REPORTES.IDS.FIRMA_JEFE); imgJefeHtml = "<img src='data:" + fJ.getMimeType() + ";base64," + Utilities.base64Encode(fJ.getBlob().getBytes()) + "' style='max-height:100px; display:block; margin:0 auto 4px auto;' />"; } catch(e){}

    var nombreComercialInd = datosForm.nombreMaquina.toUpperCase();
    var modeloOficialInd = diccionarioModelos[nombreComercialInd] ? diccionarioModelos[nombreComercialInd] : datosForm.nombreMaquina;
    var fechaHoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    var htmlFilasComponentes = "";
    for (var p = 0; p < datosForm.evaluaciones.length; p++) {
      var ev = datosForm.evaluaciones[p];
      var statusTexto = ev.seleccion === "Cumple" ? "Cumple / Pass" : "No Cumple / Fail";
      var compBadgeStyle = ev.seleccion === "Cumple" ? "border: 1px solid #393D42; color: #393D42;" : "background-color: #E60404; color: #FFFFFF; border: 1px solid #E60404;";

      var celdaEvidencia = "";
      if (ev.comentario) { 
        var comentarioEn = traducirTexto(ev.comentario);
        var textoEvBilingue = ev.comentario + (comentarioEn ? " <br><i style='color:#718096;'>(" + comentarioEn + ")</i>" : "");
        celdaEvidencia += "<div style='font-size:11px; margin-bottom:4px;'>" + textoEvBilingue + "</div>"; 
      }
      if (ev.fotoUrlDirecta) {
        var matchId = ev.fotoUrlDirecta.match(/id=([^&]+)/) || ev.fotoUrlDirecta.match(/\/d\/([^/]+)/);
        if (matchId && matchId[1]) {
          try {
            var fileF = DriveApp.getFileById(matchId[1]);
            celdaEvidencia += "<div style='margin-top:4px;'><img src='data:" + fileF.getMimeType() + ";base64," + Utilities.base64Encode(fileF.getBlob().getBytes()) + "' style='max-height:105px; border:1px solid #CBD5E0; border-radius:4px;' /></div>";
          } catch(err) {
            celdaEvidencia += "<div style='margin-top:4px;'><a href='" + ev.fotoUrlDirecta + "' target='_blank' style='font-size:10px; color:#E60404; font-weight:bold;'>Ver Foto Evidencia ↗</a></div>";
          }
        }
      }

      var transHeader = diccionarioBilingue[ev.pieza];
      var headerFinal = transHeader ? (ev.pieza + " / " + transHeader) : ev.pieza;

      htmlFilasComponentes += "<tr>" +
        "<td style='padding:6px 10px; font-size:11px; border-bottom:1px solid #E2E8F0; color:#2D3748; font-weight:600;'>" + headerFinal + "</td>" +
        "<td style='padding:6px 25px; font-size:11px; border-bottom:1px solid #E2E8F0; text-align:center;'>" +
          "<span style='display:inline-block; padding:2px 8px; font-weight:700; border-radius:4px; font-size:9px; text-transform:uppercase; " + compBadgeStyle + "'>" + statusTexto + "</span>" +
        "</td>" +
        "<td style='padding:6px 10px 6px 35px; font-size:11px; border-bottom:1px solid #E2E8F0; background-color:#fafafa;'>" + celdaEvidencia + "</td>" +
      "</tr>";
    }

    var dispValInd = datosForm.disposicion;
    var indDispBadgeStyle = "border: 1px solid #393D42; color: #393D42;";
    var indDispBilingue = dispValInd;
    if (dispValInd.toLowerCase().includes("desviacion") || dispValInd.toLowerCase().includes("desviación")) {
      indDispBadgeStyle = "background-color: #EF6C00; color: #FFFFFF; border: 1px solid #EF6C00;"; indDispBilingue = "Aceptado con Desviación / Accepted with Deviation";
    } else if (dispValInd.toLowerCase().includes("aceptado")) {
      indDispBadgeStyle = "background-color: #2E7D32; color: #FFFFFF; border: 1px solid #2E7D32;"; indDispBilingue = "Aceptado / Accepted";
    } else if (dispValInd.toLowerCase().includes("rechazado")) {
      indDispBadgeStyle = "background-color: #E60404; color: #FFFFFF; border: 1px solid #E60404;"; indDispBilingue = "Rechazado / Rejected";
    }

    var emailInsp = datosForm.correo;
    var inspectorNombre = "Inspector";
    if (emailInsp) {
      var partes = emailInsp.split('@')[0].split('.');
      if (partes.length >= 2) { inspectorNombre = partes[0].charAt(0).toUpperCase() + partes[0].slice(1) + " " + partes[1].charAt(0).toUpperCase() + partes[1].slice(1); } 
      else { inspectorNombre = emailInsp.split('@')[0]; }
    }
    var imgInspectorHtml = "<div style='height:100px;'></div>";
    var idFirmaInsp = CONFIG_REPORTES.INSPECTORES[emailInsp.toLowerCase()];
    if (idFirmaInsp) {
      try { var fI = DriveApp.getFileById(idFirmaInsp); imgInspectorHtml = "<img src='data:" + fI.getMimeType() + ";base64," + Utilities.base64Encode(fI.getBlob().getBytes()) + "' style='max-height:100px; display:block; margin:0 auto 4px auto;' />"; } catch(e){}
    }

    var comentariosGeneralesEnForm = traducirTexto(datosForm.comentariosGenerales);
    var htmlComentariosGeneralesForm = "<span style='color:#a0aec0; font-style:italic;'>Sin comentarios generales registrados. / No general comments registered.</span>";
    if (datosForm.comentariosGenerales) {
      htmlComentariosGeneralesForm = datosForm.comentariosGenerales + (comentariosGeneralesEnForm ? "<br><span style='color:#4A5568; font-style:italic; display:block; margin-top:4px;'><b>Translation:</b> " + comentariosGeneralesEnForm + "</span>" : "");
    }

    var htmlMaster = "<!DOCTYPE html><html><head><meta charset='utf-8'></head><body style='font-family:Arial, sans-serif; color:#393D42; background-color:#FFFFFF; line-height:1.4; margin:0; padding:0;'>" +
      "<div style='padding: 20px 25px;'>" +
        "<table style='width:100%; border-collapse:collapse; margin-bottom:16px;'>" +
          "<tr><td><h1 style='font-family:\"Arial Black\", sans-serif; font-size:30px; font-weight:900; color:#E60404; margin:0;'>" + CONFIG_REPORTES.EMPRESA_NOMBRE + "</h1><div style='font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:1px; margin-top:2px;'>Reporte de Control de Calidad — INCOMING</div></td><td style='text-align:right;'>" + imgLogoHtml + "</td></tr>" +
        "</table>" +
        "<div style='width:100%; padding:8px 10px; background-color:#F8F9FA; border-top:2px solid #393D42; border-bottom:1px solid #E2E8F0; margin-bottom:12px; font-size:10px;'>" +
          "<div style='margin-bottom:6px;'><b>MÁQUINA/MODELO / MACHINE/MODEL:</b> <span style='color:#E60404; font-weight:bold;'>" + modeloOficialInd + " / " + datosForm.nombreMaquina + "</span></div>" +
          "<table style='width:100%; font-size:10px; border-collapse:collapse;'>" +
            "<tr>" +
              "<td style='width:40%; padding-bottom:4px;'><b>NÚMERO SERIE / SERIAL NUMBER:</b> " + datosForm.serie + "</td>" +
              "<td style='width:30%; padding-bottom:4px;'><b>FECHA / DATE:</b> " + fechaHoy + "</td>" +
              "<td style='width:30%; padding-bottom:4px;'><b>CONTENEDOR / CONTAINER:</b> " + datosForm.contenedor + "</td>" +
            "</tr>" +
            "<tr>" +
              "<td><b>REFERENCIA / REFERENCE:</b> " + datosForm.referencia + "</td>" +
              "<td colspan='2'><b>INSPECTOR / INSPECTOR:</b> " + emailInsp + "</td>" +
            "</tr>" +
          "</table>" +
        "</div>" +
        
        "<div style='font-size:15px; font-weight:700; text-transform:uppercase; margin:12px 0 6px 0; padding-bottom:4px; border-bottom:2px solid #E60404;'>1. Evaluación de Componentes / Components Evaluation</div>" +
        "<table style='width:100%; border-collapse:collapse; margin-bottom:16px;'>" +
          "<thead><tr style='background-color:#393D42; color:#FFFFFF; font-size:10px; text-transform:uppercase;'><th style='padding:6px 10px; text-align:left; width:35%;'>Componente Evaluado / Evaluated Component</th><th style='padding:6px 10px; text-align:center; width:20%;'>Estatus / Status</th><th style='padding:6px 10px 6px 35px; text-align:left; width:45%;'>Notas y Evidencias / Notes & Photos</th></tr></thead>" +
          "<tbody>" + htmlFilasComponentes + "</tbody>" +
        "</table>" +
        
        "<div style='margin:16px 0; padding:10px; background-color:#f8fafc; border:1px solid #e2e8f0; border-radius:6px; font-size:12px; font-weight:bold; color:#393D42; text-align:right;'>DISPOSICIÓN FINAL | FINAL DISPOSITION: <span style='display:inline-block; padding:4px 12px; font-weight:700; border-radius:4px; font-size:11px; text-transform:uppercase; " + indDispBadgeStyle + "'>" + indDispBilingue + "</span></div>" +
        
        "<div style='font-size:14px; font-weight:700; text-transform:uppercase; margin:16px 0 6px 0; padding-bottom:4px; border-bottom:2px solid #393D42;'>2. Comentarios Generales del Equipo / General Machine Comments</div>" +
        "<div style='background-color:#F8F9FA; border-left:4px solid #E60404; padding:10px; font-size:11px; color:#2D3748; min-height:25px; margin-bottom:16px;'>" + 
          htmlComentariosGeneralesForm + 
        "</div>" +
        "<table style='width:100%; border-collapse:collapse; margin-top:30px; page-break-inside: avoid;'>" +
          "<tr>" +
            "<td style='width:50%; text-align:center; vertical-align:bottom; padding:0 30px;'>" + imgInspectorHtml + "<div style='width:100%; border-top:1px solid #393D42; margin-bottom:4px;'></div><div style='font-size:11px; font-weight:700; text-transform:uppercase;'>" + inspectorNombre + "</div><div style='font-size:9px; color:#718096;'>Inspector de Calidad / Quality Inspector</div></td>" +
            "<td style='width:50%; text-align:center; vertical-align:bottom; padding:0 30px;'>" + imgJefeHtml + "<div style='width:100%; border-top:1px solid #393D42; margin-bottom:4px;'></div><div style='font-size:11px; font-weight:700; color:#1A202C; text-transform:uppercase;'>" + CONFIG_REPORTES.NOMBRES.AUTORIZACION + "</div><div style='font-size:9px; color:#718096;'>Autorización / Authorized Signature</div></td>" +
          "</tr>" +
        "</table>" +
      "</div></body></html>";

    var finalBlob = Utilities.newBlob(htmlMaster, "text/html", "TempInd.html").getAs(MimeType.PDF);
    
    var nombreArchivo = datosForm.referencia + "_" + datosForm.nombreMaquina + "_" + datosForm.serie + ".pdf";
    finalBlob.setName(nombreArchivo);
    
    var carpetaDestino = DriveApp.getFolderById(CONFIG_REPORTES.IDS.CARPETA_INDIVIDUALES);
    carpetaDestino.createFile(finalBlob);
    
  } catch(e) {
    Logger.log("Error en PDF automático: " + e.toString());
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 ' + CONFIG_REPORTES.EMPRESA_NOMBRE).addItem('📄 Generar Reporte de Referencia', 'generarReporteConsolidado').addToUi();
}