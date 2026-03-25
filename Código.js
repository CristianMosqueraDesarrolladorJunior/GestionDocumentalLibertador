function doGet(e) {
  var Template = HtmlService.createTemplateFromFile("Index").evaluate().setTitle("CRM El Libertador").setFaviconUrl("https://www.ellibertador.co/favicon.ico").addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return Template;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

const SheetConsolidado = SpreadsheetApp.openById("1oPlv7OhZMMYnNOqfprtmIEiZiOrKw8MAIAI5IDnGfXw").getSheetByName("Negocios Especiales Dev");
const SheetUsers = SpreadsheetApp.openById("1oPlv7OhZMMYnNOqfprtmIEiZiOrKw8MAIAI5IDnGfXw").getSheetByName("Table_Auto_Asignación");
const SheetAssignment = SpreadsheetApp.openById("1oPlv7OhZMMYnNOqfprtmIEiZiOrKw8MAIAI5IDnGfXw").getSheetByName("Tabla Asignación Usuarios");
const SheetConsolidadoBrokersYInmobiliarias = SpreadsheetApp.openById("1IBl08te0QJ27DYU9qGZ3qVRlOowQxWbc1Cu_Eu32g_8").getSheetByName("Consolidado")
const Espejo = SpreadsheetApp.openById("1ACFjJriwgFE-VOUHifx2Rr7zUNk_Ovy0OnQUMHKnvKY").getSheetByName("new_data_polizas- archivo David")
const Renovaciones = SpreadsheetApp.openById("1wxqoUCggSYXE0vOUHdgLnwDYfBnlATeyfTZ8CgQQEY4").getSheetByName("JSON")
const PolizasAntiguas = SpreadsheetApp.openById("1wxqoUCggSYXE0vOUHdgLnwDYfBnlATeyfTZ8CgQQEY4").getSheetByName("PolizasAntiguas")

const DataWereHouse = SpreadsheetApp.openById("1_Hi5iunWuSrsT4V2ApWKIka6sdYyz7Mo_atSrz_uxhc");
const DataGestion = DataWereHouse.getSheetByName("Gestion");


function resolveClientReference(policyNumber) {
  if (!policyNumber) return { found: false, ref: null };
  const cleanPolicy = String(policyNumber).trim();

  try {

    let finder = Espejo.getRange("BE2:BE").createTextFinder(cleanPolicy).matchEntireCell(true).findNext();

    if (finder) {
      let ref = Espejo.getRange(finder.getRow(), 2).getValue(); 
      return { found: true, ref: String(ref).trim(), source: 'Espejo' };
    }
    let finderAnt = PolizasAntiguas.getRange("C2:C").createTextFinder(cleanPolicy).matchEntireCell(true).findNext();

    if (finderAnt) {
      let ref = PolizasAntiguas.getRange(finderAnt.getRow(), 2).getValue(); 
      return { found: true, ref: String(ref).trim(), source: 'PolizasAntiguas' };
    }

    return { found: false, ref: null };

  } catch (e) {
    Logger.log("Error resolviendo referencia: " + e.toString());
    return { found: false, error: e.toString() };
  }
}


function findDriveFolderByRef(ref) {
  if (!ref) return null;
  const rootId = "1e05FPKAfrRnqBbUpOF1JP9ostCg2TsjC"; 
  const rootFolder = DriveApp.getFolderById(rootId);

  const iter = rootFolder.searchFolders(`title contains '${ref}' and trashed = false`);

  while (iter.hasNext()) {
    let candidate = iter.next();
    let name = candidate.getName();
    if (name.trim() === ref || new RegExp(`\\b${ref}\\b`, 'i').test(name)) {
      return candidate;
    }
  }
  return null;
}

function getPolicyFiles(policyRef) {
  // 1. Validación de entrada
  if (!policyRef) {
    return { success: false, message: "Referencia vacía enviada al servidor." };
  }

  try {
    console.log("🔍 [1/4] Iniciando búsqueda para:", policyRef);

    const resolution = resolveClientReference(policyRef);
    let targetRef = resolution.ref;

    if (!resolution.found) {
      console.warn("⚠️ Referencia no encontrada en BD, usando política como fallback.");
      targetRef = policyRef;
    }

    const clientFolder = findDriveFolderByRef(targetRef); // Asumo que tienes esta función auxiliar

    if (!clientFolder) {
      console.error("❌ Carpeta no encontrada para:", targetRef);
      return { success: false, message: "Carpeta no encontrada en Drive", files: [] };
    }

    let filesFound = [];

    const scanFolder = (folder, contextName) => {
      if (!folder) return;
      const files = folder.getFiles();
      while (files.hasNext()) {
        let f = files.next();
        filesFound.push({
          id: f.getId(),
          name: f.getName(),
          mimeType: f.getMimeType(),
          url: f.getUrl(),
          size: (f.getSize() / 1024 / 1024).toFixed(2) + " MB",
          date: f.getLastUpdated().toISOString(),
          context: contextName
        });
      }
    };

    scanFolder(clientFolder, 'General (Raíz)');
    const subFolders = clientFolder.getFolders();
    while (subFolders.hasNext()) {
      let sub = subFolders.next();
      let name = sub.getName().toLowerCase();

      if (name.includes("renova")) {
        scanFolder(sub, 'Renovaciones');
        let deepSubs = sub.getFolders();
        while (deepSubs.hasNext()) {
          let deep = deepSubs.next();
          if (deep.getName().toLowerCase().includes("especial")) {
            scanFolder(deep, 'Procesos Especiales');
          }
        }
      }
    }

    console.log(`[4/4] Búsqueda finalizada. Archivos: ${filesFound.length}`);

    return {
      success: true,
      files: filesFound,
      folderUrl: clientFolder.getUrl()
    };

  } catch (e) {
    console.error("ERROR FATAL en getPolicyFiles: " + e.stack);
    // Retorno de error controlado (para que el frontend no reciba null)
    return { success: false, error: e.toString(), files: [] };
  }
}


function buscarCarpetaYGuardarArchivos(archivosBase64, datos, ref) {
  const rootId = "1e05FPKAfrRnqBbUpOF1JP9ostCg2TsjC";
  const rootFolder = DriveApp.getFolderById(rootId);
  const urlsNuevas = {};

  const iteradorCandidatos = rootFolder.searchFolders(`title contains '${ref}' and trashed = false`);
  let carpetaCliente = null;

  while (iteradorCandidatos.hasNext()) {
    let candidato = iteradorCandidatos.next();
    let nombre = candidato.getName();
    if (nombre.includes(ref)) {
      carpetaCliente = candidato;
      break;
    }
  }

  if (!carpetaCliente) {
    return urlsNuevas;
  }

  let carpetaRenovacion = null;
  const subCarpetas = carpetaCliente.getFolders();

  while (subCarpetas.hasNext()) {
    let sub = subCarpetas.next();
    if (sub.getName().toLowerCase().includes("renova")) {
      carpetaRenovacion = sub;
      break;
    }
  }

  if (!carpetaRenovacion) {
    carpetaRenovacion = carpetaCliente.createFolder("Renovacion");
  }

  guardarArchivosEnCarpeta(archivosBase64, datos, carpetaRenovacion);
  return urlsNuevas;
}




function guardarArchivosEnCarpeta(archivosBase64, datos, carpetaRenovacion) {
  const mapaPrefijos = {
    'sarlaft': 'SARLAFT',
    'contrato': 'CONTRATO_ARR',
    'poliza': 'POLIZA_RENOVADA',
    'cedula': 'CEDULA_TOMADOR',
    'comprobante': 'SOPORTE_PAGO',
    'propietarioDoc': 'DOC_PROPIETARIO',
    'otroSi': 'OTRO_SI',
    'cesionCedula': 'CESION_ID',
    'polizaAjuste': 'POLIZA_CON_AJUSTE' 
  };

  let urlsGeneradas = {};
  if (!archivosBase64) return urlsGeneradas;

  const llavesEspeciales = ['otroSi', 'cesionCedula', 'cesionDocAdicional'];
  let carpetaEspeciales = null;
  const tieneEspeciales = Object.keys(archivosBase64).some(k => llavesEspeciales.includes(k));
  if (tieneEspeciales) {
    const subIter = carpetaRenovacion.getFoldersByName("Procesos Especiales");
    carpetaEspeciales = subIter.hasNext() ? subIter.next() : carpetaRenovacion.createFolder("Procesos Especiales");
  }
  for (const [key, fileObj] of Object.entries(archivosBase64)) {
    if (!fileObj || !fileObj.data) continue;
    try {
      const folderDestino = llavesEspeciales.includes(key) ? carpetaEspeciales : carpetaRenovacion;
      const prefijo = mapaPrefijos[key] || key.toUpperCase();
      const fileName = fileObj.name || "archivo.pdf";
      const ext = fileName.includes('.') ? fileName.split('.').pop() : 'pdf';
      const idParaNombre = datos.poliza || datos.solicitud || "SinID";
      const nombreFinal = `${prefijo}_${idParaNombre}.${ext}`;
      const existing = folderDestino.getFilesByName(nombreFinal);
      while (existing.hasNext()) { existing.next().setTrashed(true); }
      const blob = Utilities.newBlob(Utilities.base64Decode(fileObj.data), fileObj.mimeType || "application/pdf", nombreFinal);
      const file = folderDestino.createFile(blob);
      urlsGeneradas[key + 'URL'] = file.getUrl();
    } catch (err) {
      console.error("Error guardando " + key + ": " + err);
    }
  }

  return urlsGeneradas;
}

function extraerIdGoogleDrive(url) {
  if (!url) return null;
  const strUrl = String(url);
  const match = strUrl.match(/[-\w]{25,}/);
  if (match && match[0]) {
    return match[0];
  }
  return null;
}


function obtenerCarpetaRenovacionPorPoliza(poliza, documento, asegurado) {
  const rootId = "1e05FPKAfrRnqBbUpOF1JP9ostCg2TsjC"; // TU ID RAÍZ
  const rootFolder = DriveApp.getFolderById(rootId);
  let ref = null;
  if (Espejo) {
    let finderEspejo = Espejo.getRange("BE2:BE").createTextFinder(poliza).matchEntireCell(true).findNext();
    if (finderEspejo) {
      let row = finderEspejo.getRow();
      ref = Espejo.getRange(row, 2).getDisplayValue();
      console.log(`Referencia encontrada en Espejo para póliza ${poliza}: ${ref}`);
    }
  }

  if (!ref || ref.trim() === "") {
    if (PolizasAntiguas) {
      let finderAntiguas = PolizasAntiguas.getRange("C2:C").createTextFinder(poliza).matchEntireCell(true).findNext();
      if (finderAntiguas) {
        let row = finderAntiguas.getRow();
        ref = PolizasAntiguas.getRange(row, 2).getDisplayValue();
        console.log(`Referencia encontrada en PolizasAntiguas para póliza ${poliza}: ${ref}`);
      }
    }
  }

  // C. Si no existe en ninguna, CREAR NUEVA REFERENCIA
  if (!ref || ref.trim() === "") {
    // Usar documento si existe, sino poliza como fallback para la referencia
    const baseRef = documento || poliza;
    ref = "Ref" + baseRef;
    const fechaActual = new Date();
    if (PolizasAntiguas) {
      PolizasAntiguas.appendRow([
        fechaActual,
        ref,
        poliza,
        asegurado
      ]);
      console.log(`Nueva referencia creada y registrada en PolizasAntiguas: ${ref}`);
    } else {
      console.error("Hoja PolizasAntiguas no encontrada. No se pudo guardar la referencia.");
    }
  }

  // 2. GESTIÓN DE CARPETAS DRIVE (Usando la Ref encontrada/creada)
  // ----------------------------------------------------------------
  let carpetaCliente = null;

  // Buscar carpeta existente por nombre exacto o que contenga la referencia
  const iterador = rootFolder.searchFolders(`title contains '${ref}' and trashed = false`);

  while (iterador.hasNext()) {
    let folder = iterador.next();
    let nombre = folder.getName();

    // Verificación estricta o regex para asegurar coincidencia
    if (nombre.trim() === String(ref).trim() || new RegExp(`\\b${ref}\\b`, 'i').test(nombre)) {
      carpetaCliente = folder;
      break;
    }
  }

  // Si no existe la carpeta del cliente, la creamos
  if (!carpetaCliente) {
    carpetaCliente = rootFolder.createFolder(ref + " - " + (asegurado || "Cliente"));
    console.log(`Carpeta raíz creada: ${carpetaCliente.getName()}`);
  }

  // Buscar o crear subcarpeta "Renovacion"
  let carpetaRenovacion = null;
  const subIter = carpetaCliente.getFolders();
  while (subIter.hasNext()) {
    let sub = subIter.next();
    // Búsqueda flexible para "Renovacion", "Renovaciones", etc.
    if (sub.getName().toLowerCase().includes("renova")) {
      carpetaRenovacion = sub;
      break;
    }
  }

  if (!carpetaRenovacion) {
    carpetaRenovacion = carpetaCliente.createFolder("Renovacion");
    console.log("Subcarpeta Renovacion creada");
  }

  return {
    carpeta: carpetaRenovacion,
    referencia: ref,
    url: carpetaRenovacion.getUrl()
  };
}


function processAnalystDecision(payload) {
  const sheet = Renovaciones;

  const mapaNombres = {
    'sarlaft': 'SARLAFT',
    'contrato': 'CONTRATO_ARR',
    'poliza': 'POLIZA_RENOVADA',
    'cedula': 'CEDULA_TOMADOR',
    'comprobante': 'SOPORTE_PAGO',
    'polizaAjuste': 'POLIZA_CON_AJUSTE'  
  };

  try {
    const poliza = payload.dataLead.poliza;
    const documento = payload.dataLead.documento;
    const asegurado = payload.dataLead.asegurado;
    const resultadoCarpeta = obtenerCarpetaRenovacionPorPoliza(poliza, documento, asegurado);
    const carpetaRenovacion = resultadoCarpeta.carpeta;

    let urlsFinales = {};

    if (payload.files && Object.keys(payload.files).length > 0) {
      const urlsLocales = guardarArchivosEnCarpeta(payload.files, payload.dataLead, carpetaRenovacion);
      Object.assign(urlsFinales, urlsLocales);
    }

    if (payload.driveRefs && Object.keys(payload.driveRefs).length > 0) {
      for (const [key, meta] of Object.entries(payload.driveRefs)) {
        try {
          let fileId = extraerIdGoogleDrive(meta.url);
          if (fileId) {
            const originalFile = DriveApp.getFileById(fileId);
            const prefijo = mapaNombres[key] || key.toUpperCase();
            const nombreOriginal = originalFile.getName();
            const ext = nombreOriginal.includes('.') ? nombreOriginal.split('.').pop() : 'pdf';
            const nuevoNombre = `${prefijo}_${poliza}.${ext}`;
            const existing = carpetaRenovacion.getFilesByName(nuevoNombre);
            while (existing.hasNext()) { existing.next().setTrashed(true); }
            const copy = originalFile.makeCopy(nuevoNombre, carpetaRenovacion);
            urlsFinales[key + 'URL'] = copy.getUrl();

          } else {
            console.warn(`⚠️ No se pudo extraer ID válido de la URL: ${meta.url}`);
          }
        } catch (err) {
          console.error(`🔥 Error copiando Drive File (${key}): ${err.toString()}`);
        }
      }
    }

    const data = sheet.getDataRange().getDisplayValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      const rawJsonColB = String(data[i][1])
      if (rawJsonColB.includes(poliza)) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex !== -1) {
      let nuevoEstado = "CORRECCION";
      let nuevoAsesor = "";

      switch (payload.decision) {
        case 'APPROVE': nuevoEstado = "Expedido"; break;
        case 'CANCELADA': nuevoEstado = "Cancelado"; break;
        case 'REVIEWED': nuevoEstado = "Caso Revisado"; break;
        case 'CORRECTION':
          nuevoAsesor = getNewLeadAssignment(payload.segmento) || "sin.asignar@segurosbolivar.com";
          sheet.getRange(rowIndex, 3).setValue(nuevoAsesor);  
          break;
      }
      sheet.getRange(rowIndex, 5).setValue(nuevoEstado);

      let jsonCell = sheet.getRange(rowIndex, 6);
      let jsonActual = {};
      try {
        jsonActual = JSON.parse(jsonCell.getValue());
      } catch (e) {
        jsonActual = payload.dataLead || {};
      }

      Object.assign(jsonActual, urlsFinales);
      if (payload.manualUpdates) {
        Object.assign(jsonActual, payload.manualUpdates);
      }
      jsonCell.setValue(JSON.stringify(jsonActual));
      let cellObs = sheet.getRange(rowIndex, 7);
      let historial = [];
      try {
        let val = cellObs.getValue();
        if (val) historial = JSON.parse(val);
        if (!Array.isArray(historial)) historial = [];
      } catch (e) { historial = []; }

      historial.push({
        fecha: new Date().toISOString(),
        usuario: Session.getActiveUser().getEmail(),
        observacion: payload.observations,
        estado: nuevoEstado,
        accion: "GESTION_ANALISTA"
      });
      cellObs.setValue(JSON.stringify(historial));

      return { success: true, message: "Gestión procesada correctamente." };

    } else {
      throw new Error("No se encontró la fila en Google Sheets para el ID: " + payload.leadId);
    }

  } catch (e) {
    console.error(e);
    return { success: false, message: "Error Servidor: " + e.toString() };
  }
}

function GetDataUser() {
  var UserMail = Session.getActiveUser().getEmail();
  Logger.log(UserMail)
  let Status;
  let DataRange = SheetConsolidado.getRange("A2:BC" + SheetConsolidado.getLastRow()).getDisplayValues();
  let DataRangeUserPending = [];
  let Rows = SheetAssignment.getRange("A:A").createTextFinder(UserMail).ignoreDiacritics(true).matchEntireCell(true).ignoreDiacritics(true).findPrevious();



  if (Rows != null) {
    Status = "Autenticado";
    let FilaDataRol = SheetAssignment.getRange("A:A").createTextFinder(UserMail).matchEntireCell(true).ignoreDiacritics(true).findPrevious().getRow();
    let Rol = SheetAssignment.getRange("C" + FilaDataRol).getDisplayValue();
    console.log
    if (Rol == "Gestión Documental") {
      DataRange.filter(function (DataRange) {
        let Validation = DataRange.indexOf("Pendiente Validación Documental");
        let ValidationCorrecion = DataRange.indexOf("Pendiente Corrección Documental");
        let Validation2 = DataRange[54].indexOf(UserMail);
        if ((Validation > -1 || ValidationCorrecion > -1) && Validation2 > -1) {
          DataRangeUserPending.push([DataRange[0], DataRange[4], DataRange[8], DataRange[24], DataRange[26], "Gestión", DataRange[11], DataRange[12], DataRange[1], DataRange[3], DataRange[5], DataRange[9], DataRange[10], DataRange[13], DataRange[14], DataRange[15], DataRange[17], DataRange[19], DataRange[20], DataRange[21], DataRange[22], DataRange[23], DataRange[16], DataRange[45], DataRange[7], DataRange[37]]);
        }
      });
      console.log(DataRangeUserPending)
      return [Status, DataRangeUserPending, UserMail, Rol];
    } else if (Rol == "Gestión Documental 2") {
      let data = GetDataBrokersYInmobiliarias();
      return [Status, data, UserMail, Rol];


    } else if (Rol == "Analista Renovaciones") {
      let dataUpd = Renovaciones.getRange("A2:H" + Renovaciones.getLastRow()).getDisplayValues();
      let misRegistros = dataUpd.filter(row =>
        row[2] && row[2].toString().trim().toLowerCase() === UserMail.trim().toLowerCase()
      );
      let todosMisLeads = misRegistros.map(row => {
        const fechaIngreso = row[0];
        const registroRaw = row[1]; // Col B
        const nombreAgente = row[2];
        const etapaFunel = row[3];
        const estadoGestion = row[4]; // Col E (El estado clave)
        const dataGestionRaw = row[5]; // Col F
        const historiaRaw = row[6]; // Col G

        let historialGestion = [];
        if (historiaRaw && String(historiaRaw).trim() !== "") {
          let cleanHistoryString = String(historiaRaw).trim();
          if (cleanHistoryString.startsWith(")]}',")) {
            cleanHistoryString = cleanHistoryString.substring(5);
          }
          try {
            let parsed = JSON.parse(cleanHistoryString);
            historialGestion = Array.isArray(parsed) ? parsed : [parsed];
          } catch (e) {
            Logger.log("Error historial: " + e.message);
            historialGestion = [];
          }
        }

        let leadSelect = {};
        try {
          let cleanGestion = String(dataGestionRaw).trim().replace(/\bNaN\b/g, "null");
          leadSelect = JSON.parse(cleanGestion);
        } catch (e) {
          leadSelect = { error: "JSON Gestión inválido" };
        }
        let dataLead = {};
        try {
          let cleanLead = String(registroRaw).trim().replace(/\bNaN\b/g, "null");
          dataLead = JSON.parse(cleanLead);
        } catch (e) {
          dataLead = { error: "JSON Lead inválido" };
        }
        return {
          fechaIngreso: fechaIngreso,
          leadSelect: leadSelect,
          dataLead: dataLead, // OJO: Antes lo llamabas strLead o dataLead, lo unifiqué aquí
          nombreAgente: nombreAgente,
          etapaFunel: etapaFunel,
          estadoGestion: estadoGestion,
          historialGestiones: historialGestion
        };
      });

      let dataAnalitic = todosMisLeads.filter(item => item.estadoGestion === "Enviar a Expedicion" || item.estadoGestion === "Autogestionado" || item.estadoGestion === "Caso Revisado");
      let dataEspecial = todosMisLeads.filter(item => item.estadoGestion === "Caso Especial");
      let polizasRenovadas = todosMisLeads.filter(item => item.estadoGestion === "Poliza Renovada" || item.estadoGestion === "Expedido"); // 
      console.log(`Analitic: ${dataAnalitic.length}, Especial: ${dataEspecial.length}, Renovadas: ${polizasRenovadas.length}`);

      return [
        Status,
        {
          pendientes: dataAnalitic,
          especiales: dataEspecial,
          renovadas: polizasRenovadas
        },
        UserMail,
        Rol
      ];
    }
  } else {
    Status = "No Autenticado";
    return [Status, "Null"];
  }
}

function getNewLeadAssignment(segmento) {
  let gestion = "Renovations";
  let seg = String(segmento || "").toUpperCase();
  if (seg.includes("BROKER") || seg.includes("INMOBILIARIA")) {
    gestion = "CorreccionesBI";
  }
  let asignacion = AssignLead(gestion);
  return asignacion.email;
}

function AssignLead(deal = "CorreccionesBI") {

  let userDataleads = DataGestion.getRange("A1:K" + DataGestion.getLastRow()).getDisplayValues();

  if (deal === "Renovations") {
    userData = userDataleads.filter(function (row) {
      return row[3] === "Renovations";
    })
  } else if (deal === "Sales") {
    userData = userDataleads.filter(function (row) {
      return row[3] === "Seguro de Vida" || row[3] === "Seguro de Desempleo";
    })
  } else if (deal === "CorreccionesBI") {
    userData = userDataleads.filter(function (row) {
      return row[3] === "CorreccionesBI";
    })
  }
  let bestAgent = null;
  let highestEffectiveness = -1;
  let bestSortingKey = Number.POSITIVE_INFINITY;
  let equity = 0.5;

  for (let i = 0; i < userData.length; i++) {
    let id = userData[i][0];
    let agentName = userData[i][1];
    let email = userData[i][2];
    let novelty = userData[i][4];
    let productoAsignado = userData[i][3];
    let totalCapacity = Number(userData[i][5]);
    let totalInProcess = Number(userData[i][6]);
    let effectivenessStr = userData[i][10] ? userData[i][10].toString().replace("%", "").trim() : "0";
    let effectiveness = Number(effectivenessStr);


    let availability = totalCapacity - totalInProcess;

    if ((!novelty || novelty.toString().trim() === "") && availability > 0) {

      console.log(availability)
      let sortingKey = (-(1 - equity) * totalCapacity) + (equity * totalInProcess);

      if (sortingKey < bestSortingKey || (sortingKey === bestSortingKey && effectiveness > highestEffectiveness)) {
        bestSortingKey = sortingKey;
        highestEffectiveness = effectiveness;

        bestAgent = {
          name: agentName,
          email: email,
          productoAsignado: productoAsignado
        };
      }
    }
  }
  if (bestAgent !== null) {
    console.log("The most available agent is: " + bestAgent.name);
    console.log("gestion Asigned: " + bestAgent.productoAsignado)
    return bestAgent;
  } else {
    Logger.log("No agents available.");
    return null;
  }
}


function ValuesCotizacion(Canon, Admon, FuncionarioBolivar, MesesVigencia) {
  Logger.log(Canon + " - " + Admon + " - " + FuncionarioBolivar + " - " + MesesVigencia)
  var Tasa;
  if (FuncionarioBolivar == "Si") {
    Tasa = 0.03;
  } else if (FuncionarioBolivar == "No") {
    Tasa = 0.035;
  }
  var CanonAdmon = Canon + Admon;
  MesesVigencia = parseInt(MesesVigencia);
  var TotalAsegurado = CanonAdmon * MesesVigencia;
  var PrimaNeta = +TotalAsegurado * Tasa;
  var Iva19 = (PrimaNeta * 0.19);
  var PrimaPagar = parseInt(PrimaNeta + Iva19);
  return [PrimaPagar, TotalAsegurado]
}

const PROJECT_ID_GLOBAL = "671385553610";
const LOCATION = "us";
const PROCESSOR_ID = "7723f88834c656ed";
const SERVICE_ACCOUNT = {
  private_key: "-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQDAs2Icjjf0Eujr\nFYudm+/8yWJQjwE8jhXtZW1NFllvFh8gXc5tH0TeCny5KKaJSx5yxADQ2X9lr2Qq\nXVdX4vcuEVYMqgSAjfLJkmSt8AiO2rK1RdIf6yvpSH8LetNmydH8xibKxWDMWWuc\nyYuhoaT30TgxQ7M+xYi+lt2Mywu3twTPbGsKgb7dUPd4aLmwGIr74N90UkUcYgZm\nXc9SKKuPyLQZ2LRdF7JKaLdY4JTGLPdrGqaAsr4ZlXTmCmiZTUB1GaI7s99eL3fY\nIHM1j/GnS4T9TgypYKIZ5JKziCYQJsxEwHQG+FpUSXWBu21fJjlaIrKyCZuyKGCj\n/hOGvN+lAgMBAAECggEAKo0R25td6KPqUcrSpw1he3DeqEpDrCL13ZN5hL2sJvb8\nDZIZPIhclSk8rEg5KfTv9sioI3X7hzEpDZ/J4yrHiSEj3q0GTHrLw03ztGLeCOlq\n79NImGq+KgerohXPq5FiMI5yz3CxNL6EID1y+1Bt1Jka7unzoSdOUEORDX9iiYDa\nan8OfWqyMP567hK1w61En0faA8N99lh9jUaXLKImJQ8D2elBq9A+Lxe2FlG0bfXd\nQVjpcZ/BsySFZWDVlGqNV9oGlwyouqfxfzXtDwUDDacMdZTlTL8AcikIvflbxiFN\nWxs4nny/SO3cKFeA66NJkL1B6vEzjoSekG3y5KpBAQKBgQDzeT9yK502k9U4KaeK\n2Y1s8qGbdGYKgN1F2PMgov/QaNip376xF3Zo59Zq0VIYeFjx3HLx0HiiQeZxbcaC\njb+Bpcg+wv0zr2T0LFOEVOFNOT2/aKf09QPu2Mgakr4sAkxUK4c0qvpJa2WxOtaO\ntiKD953DfgdAJvFRL4YOQxNGJQKBgQDKnWneubSFG7vyCFie3tblT/1bRM6qQGhI\nMbWk9ap97CRuKKPFkXfkaexqy+WBv+WH62aH6h0gfOdffqEpJcJOce0a9YnJTIwD\nWHSqFRrnST+BTdRl6wlQtaurypoPJkprkRtkSI22YNn7je8ZSDzsTYP+eZKjBCGu\nPTACCKw7gQKBgEh2UJS5OEwTCYVymEOx5e6D8+chaHE90x1DqXCQMpSjb8B3L/ji\n48HrJhyaedWAk/A/zRH9Gron5N7jbg5TA6khXwyW2eb1D5XAT4b2ACwMmj0Kd9pm\nxanjaQLHo8PTV0ZBwjbBoEYTqatquIq22GTwYErbimrkbDPecgZynhzlAoGAFjj1\np6wOlJraHk20Cpi+USBY1W3SjPHLfj+VgKZBMNZ5mGt0qvKth6vmdkAux/BYKHQ1\nJqsSzsFkTyEAZBb0HM56Bv7vQdjXcnZ9NTpjXQK3qGL07Mi+mM+UKJ9sDkVQ3ENq\nEbGzeVFeFy0WEFvP8sr9syd6Yc7OMuIbJd31pgECgYAlQoJ3MkoVFUN5/IPe9EZB\n1lRDTXnGffM92MOFpWChjNeGdcGpXyk1k/0pRSmzTZXWmt67qqMC5iAL4LE/tf14\nNqd0cCpl+SWFTdC5AXP5eJRUaSPenQfN9dJM+qlKGN65lZkurStgYZkxK4ChWDcR\nK9S2TDdxm0Je7ge1DdSwKg==\n-----END PRIVATE KEY-----\n",
  client_email: "admin-buckets-ai@proyecto-ia-servicios-bolivar.iam.gserviceaccount.com",
};

function reset() {
  var service = getService();
  service.reset();
}

function getService() {
  return (
    OAuth2.createService("GCP")
      .setTokenUrl("https://accounts.google.com/o/oauth2/token")
      .setPrivateKey(SERVICE_ACCOUNT.private_key)
      .setIssuer(SERVICE_ACCOUNT.client_email)
      .setPropertyStore(PropertiesService.getScriptProperties())
      .setScope(["https://www.googleapis.com/auth/cloud-platform"])
  );
}

function ModelValidationCTL(Ref) {
  //var Ref = "Ref2"
  var FilaData = SheetConsolidado.getRange("B:B").createTextFinder(Ref).matchEntireCell(true).ignoreDiacritics(true).findPrevious().getRow();
  reset();
  var TipoCliente = SheetConsolidado.getRange("J" + FilaData).getDisplayValue();
  Logger.log(TipoCliente)
  if (TipoCliente == "Persona natural") {
    var IdFiles = SheetConsolidado.getRange("AT" + FilaData).getDisplayValue();
    IdFiles = JSON.parse(IdFiles);
    var FileIdCTL = DriveApp.getFileById(IdFiles.idCTL);
    var FileBase64CTL = Utilities.base64Encode(FileIdCTL.getBlob().getBytes());
    var FileTypeCTL = FileIdCTL.getMimeType();
    var FileIdCedula = DriveApp.getFileById(IdFiles.idCedula);
    var FileBase64Cedula = Utilities.base64Encode(FileIdCedula.getBlob().getBytes());
    var FileTypeCedula = FileIdCedula.getMimeType();
    var FileIdSarlaft = DriveApp.getFileById(IdFiles.idSARLAFT);
    var FileBase64Sarlaft = Utilities.base64Encode(FileIdSarlaft.getBlob().getBytes());
    var FileTypeSarlaft = FileIdSarlaft.getMimeType();
    var service = getService();
    var API_ENDPOINT = "us-central1-aiplatform.googleapis.com"
    var PROJECT_ID = "proyecto-ia-servicios-bolivar"
    var MODEL_ID = "gemini-2.0-flash-lite-001"//"gemini-1.0-pro-vision-001";
    var url = "https://" + API_ENDPOINT + "/v1/projects/" + PROJECT_ID + "/locations/us-central1/publishers/google/models/" + MODEL_ID + ":streamGenerateContent";
    var payloadAgent1 = {
      "contents": [
        {
          "role": "user",
          "parts": [
            {
              "text": "Documento 1:"
            },
            {
              "inlineData": {
                "mimeType": FileTypeCedula,
                "data": FileBase64Cedula
              }
            },
            {
              "text": "Documento 2:"
            },
            {
              "inlineData": {
                "mimeType": FileTypeCTL,
                "data": FileBase64CTL
              }
            },
            {
              "text": "Documento 3:"
            },
            {
              "inlineData": {
                "mimeType": FileTypeSarlaft,
                "data": FileBase64Sarlaft
              }
            },
            { "text": "Analiza los documentos" }
          ]
        }
      ],
      "systemInstruction": {
        "parts": [
          {
            "text": "Rol: Analista de documentos altamente riguroso. Tu principal prioridad es la extracción precisa de números de identificación y la comparación exacta. Analiza (1) Cédula (CC), (2) Certificado de Tradición (CTL), y (3) Formulario SARLAFT (B-14). Devuelve **SOLO UN ÚNICO OBJETO JSON** con esta estructura exacta: {'DocumentoEnCTL':'...','EstadoInmueble':'...','DocumentoFormularioB14':'...'}.\n\nInstrucciones:\n1. **EXTRACCIÓN Y NORMALIZACIÓN (BASE RIGUROSA):** \n    a. **Extracción:** Localiza el número de identificación del titular en cada uno de los tres documentos (CC, CTL, B-14). En el CTL y B-14, considera todas las variaciones de etiqueta: 'CC', 'No.', 'Documento', 'Identificación', 'Titular', 'Propietario'.\n    b. **Normalización:** Para cada número extraído, aplica la normalización de forma estricta: **ELIMINA COMPLETAMENTE** cualquier carácter que no sea un dígito numérico (0-9). (Ej: 'C.C. 51.984.657-X' debe ser '51984657'). La BASE de comparación es el número normalizado del Doc 1 (CC).\n2. **COMPARACIÓN (IGUALDAD DE CADENA ABSOLUTA):** Los campos 'DocumentoEnCTL' y 'DocumentoFormularioB14' deben reflejar una **igualdad de cadena de caracteres (STRING) BIT A BIT**. La respuesta debe ser 'Coincide' **SOLAMENTE SI** la BASE normalizada es idéntica *en longitud y secuencia de dígitos* al N° normalizado del CTL o B-14. Cualquier desviación, omisión, o caracter adicional/faltante debe resultar en 'No coincide'. Si el número no fue legible en el documento: 'No puedo leer [nombre del documento]'.\n3. **Campo 'EstadoInmueble':** Revisa en el CTL **EXCLUSIVAMENTE** la presencia de limitaciones de dominio (ej. hipotecas, embargos, leasing, patrimonio de familia o afectación a vivienda) **VIGENTES y ACTIVAS**. Si existe una o más: 'Inmueble Afectado: [detalle completo de la *naturaleza* y *número de anotación* de la limitación activa]'. **NO** incluya anotaciones de tradición, adquisición o correcciones. Si están anuladas, canceladas o no existen: 'Inmueble sin afectación'.\n4. **PROHIBICIÓN GLOBAL:** **BAJO NINGUNA CIRCUNSTANCIA** se debe incluir cualquier número de identificación, dígito o secuencia numérica relacionada con la CC, CTL o B-14 en el texto final de los campos JSON.\n\nDevuelve **SOLO EL JSON FINAL** sin preámbulos, explicaciones ni texto adicional."
          }]
      },
      "generationConfig": {
        "maxOutputTokens": 8192,
        "temperature": 0.5,
        "topP": 0.95,
        "seed": 0,
        "responseMimeType": "application/json",
        "responseSchema": { "type": "OBJECT", "properties": { "response": { "type": "STRING" } } }
      },
      "safetySettings": [
        {
          "category": "HARM_CATEGORY_HATE_SPEECH",
          "threshold": "OFF"
        },
        {
          "category": "HARM_CATEGORY_DANGEROUS_CONTENT",
          "threshold": "OFF"
        },
        {
          "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT",
          "threshold": "OFF"
        },
        {
          "category": "HARM_CATEGORY_HARASSMENT",
          "threshold": "OFF"
        }
      ],
    }
    /*     var payloadAgent2 = {
          "contents": [
            {
              "role": "user",
              "parts": [
                {
                  "text": "Certificado de tradición y libertad:"
                },
                {
                  "inlineData": {
                    "mimeType": FileTypeCTL,
                    "data": FileBase64CTL
                  }
                }
              ]
            }
          ],
          "systemInstruction": {
            "parts": [
              {
                "text": "Eres un Analista de Documentos, especializado en la revisión y validación de certificados de tradición y libertad (CTL). Se te proporciona un archivo adjunto que contiene el CTL de un inmueble. Tu tarea es analizar TODAS las anotaciones registradas en el documento para determinar si el inmueble tiene procesos abiertos que afecten su estado, tales como hipotecas, embargos o leasing.\nInstrucciones:\nRevisa todas las anotaciones del documento adjunto, desde la primera hasta la última.\nIdentifica cualquier anotación que indique la existencia de un proceso abierto como hipoteca, embargo o leasing.\nSi en una anotación se indica que un proceso fue cancelado o liberado, asegúrate de excluirlo de la lista de afectaciones vigentes.\nResponde exclusivamente en el formato especificado:\nSi todos los procesos están cancelados o liberados: Respuesta: \"Inmueble sin afectación\"\nSi hay afectaciones vigentes: Respuesta: \"Inmueble Afectado: [lista detallada de afectaciones vigentes]\"\nEjemplo: \"Inmueble Afectado: HIPOTECA-EMBARGO-LEASING\"\nConsideraciones:\nSi el contenido del CTL no permite completar la tarea, responde: Respuesta: \"No se pudo completar la tarea: [explicación del motivo]\"\nRealiza el análisis y responde únicamente en el formato indicado."
              }
            ]
          },
          "generationConfig": {
            "maxOutputTokens": 8192,
            "temperature": 0.5,
            "topP": 0.95,
            "seed": 0,
            "responseMimeType": "application/json",
            "responseSchema": { "type": "OBJECT", "properties": { "response": { "type": "STRING" } } }
          },
          "safetySettings": [
            {
              "category": "HARM_CATEGORY_HATE_SPEECH",
              "threshold": "OFF"
            },
            {
              "category": "HARM_CATEGORY_DANGEROUS_CONTENT",
              "threshold": "OFF"
            },
            {
              "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT",
              "threshold": "OFF"
            },
            {
              "category": "HARM_CATEGORY_HARASSMENT",
              "threshold": "OFF"
            }
          ],
        } */

    var headers = {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + service.getAccessToken()
    };

    var options = {
      method: "POST",
      headers: headers,
      payload: JSON.stringify(payloadAgent1)
    };

    /* var options2 = {
      method: "POST",
      headers: headers,
      payload: JSON.stringify(payloadAgent2),
      url: url
    }; */

    /* var requests = [options, options2]; */
    var response = UrlFetchApp.fetch(url, options);
    var responseText = response.getContentText();
    Logger.log("Respuesta Completa de la API (Texto): " + responseText);

    try {
      var responseArray = JSON.parse(responseText);
    } catch (e) {

      Logger.log("Error al parsear la respuesta de la API: " + e);
      return null;
    }
    //var responseArray = response;
    var validationResult = null;
    var validationResult = ResponseGeminiStream(responseArray);

    if (validationResult && validationResult.response) {
      Logger.log("Mensaje del modelo: " + validationResult.response);

    } else {
      Logger.log("Fallo en la validación o en el procesamiento.");
    }
    Logger.log(validationResult.response);
    /* var parsedResponse = response.map(r => JSON.parse(r.getContentText()));
    var fullText = cleanGeminiResponse(parsedResponse);
    Logger.log(fullText); */
    return validationResult.response;
  } else {
    var IdFiles = SheetConsolidado.getRange("AT" + FilaData).getDisplayValue();
    IdFiles = JSON.parse(IdFiles);
    var FileIdCTL = DriveApp.getFileById(IdFiles.idCTL);
    var FileBase64CTL = Utilities.base64Encode(FileIdCTL.getBlob().getBytes());
    var FileTypeCTL = FileIdCTL.getMimeType();
    var FileIdCedula = DriveApp.getFileById(IdFiles.idCedula);
    var FileBase64Cedula = Utilities.base64Encode(FileIdCedula.getBlob().getBytes());
    var FileTypeCedula = FileIdCedula.getMimeType();
    var FileIdRut = DriveApp.getFileById(IdFiles.idRUT);
    var FileBase64Rut = Utilities.base64Encode(FileIdRut.getBlob().getBytes());
    var FileTypeRut = FileIdRut.getMimeType();
    var FileIdRepLegal = DriveApp.getFileById(IdFiles.idCERTLEGAL);
    var FileBase64RepLegal = Utilities.base64Encode(FileIdRepLegal.getBlob().getBytes());
    var FileTypeRepLegal = FileIdRepLegal.getMimeType();
    var service = getService();
    var API_ENDPOINT = "us-central1-aiplatform.googleapis.com"
    var PROJECT_ID = "proyecto-ia-servicios-bolivar"
    var MODEL_ID = "gemini-2.0-flash-lite-001";//gemini-1.0-pro-vision-001
    var url = "https://" + API_ENDPOINT + "/v1/projects/" + PROJECT_ID + "/locations/us-central1/publishers/google/models/" + MODEL_ID + ":streamGenerateContent";

    var payloadAgent1 = {
      "contents": [
        {
          "role": "user",
          "parts": [
            {
              "text": "Cédula de Ciudadanía:"
            },
            {
              "inlineData": {
                "mimeType": FileTypeCedula,
                "data": FileBase64Cedula
              }
            },
            {
              "text": "Registro Único Tributario (RUT):"
            },
            {
              "inlineData": {
                "mimeType": FileTypeRut,
                "data": FileBase64Rut
              }
            },
            {
              "text": "Certificado de Existencia y Representación Legal:"
            },
            {
              "inlineData": {
                "mimeType": FileTypeRepLegal,
                "data": FileBase64RepLegal
              }
            }
          ]
        }
      ],
      "systemInstruction": {
        "parts": [
          {
            "text": "Eres un Analista de Documentos, especializado en la validación de datos de identificación. Se te proporcionan tres documentos: Cédula de Ciudadanía, Registro Único Tributario (RUT) y Certificado de Existencia y Representación Legal. Tu tarea es extraer el número de identificación de la Cédula de Ciudadanía y verificar si este número aparece en los otros dos documentos.\nInstrucciones:\nExtrae el número de identificación del documento de la Cédula de Ciudadanía.\nBusca el número de identificación en los otros documentos siguiendo estas guías:\nEn el RUT, utiliza como referencia las secciones donde se mencione \"Representación\" o \"REPRS LEGAL PRIN\".\nEn el Certificado de Existencia y Representación Legal, utiliza como referencia la sección que menciona \"REPRESENTANTES LEGALES\".\nCompara los resultados:\nSi el número de identificación aparece en ambos documentos (RUT y Certificado de Existencia y Representación Legal), responde:\nRespuesta: \"Si coinciden\"\nSi el número de identificación no aparece en uno de los documentos, responde:\nRespuesta: \"No coinciden\"\nConsideraciones:\nSi no puedes completar la tarea por falta de información o porque el contenido de los documentos no es claro, responde:\nRespuesta: \"No se pudo completar la tarea: [explicación del motivo]\"\nFormato de respuesta:\nResponde únicamente en el formato especificado:\nCoinciden\nNo coinciden\nRealiza el análisis y responde."
          }
        ]
      },
      "generationConfig": {
        "maxOutputTokens": 8192,
        "temperature": 0.5,
        "topP": 0.95,
        "seed": 0,
        "responseMimeType": "application/json",
        "responseSchema": { "type": "OBJECT", "properties": { "response": { "type": "STRING" } } }
      },
      "safetySettings": [
        {
          "category": "HARM_CATEGORY_HATE_SPEECH",
          "threshold": "OFF"
        },
        {
          "category": "HARM_CATEGORY_DANGEROUS_CONTENT",
          "threshold": "OFF"
        },
        {
          "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT",
          "threshold": "OFF"
        },
        {
          "category": "HARM_CATEGORY_HARASSMENT",
          "threshold": "OFF"
        }
      ],
    }

    var payloadAgent2 = {
      "contents": [
        {
          "role": "user",
          "parts": [
            {
              "text": "Certificado de tradición y libertad:"
            },
            {
              "inlineData": {
                "mimeType": FileTypeCTL,
                "data": FileBase64CTL
              }
            }
          ]
        }
      ],
      "systemInstruction": {
        "parts": [
          {
            "text": "Eres un Analista de Documentos, especializado en la revisión y validación de certificados de tradición y libertad (CTL). Se te proporciona un archivo adjunto que contiene el CTL de un inmueble. Tu tarea es analizar TODAS las anotaciones registradas en el documento para determinar si el inmueble tiene procesos abiertos que afecten su estado, tales como hipotecas, embargos o leasing.\nInstrucciones:\nRevisa todas las anotaciones del documento adjunto, desde la primera hasta la última.\nIdentifica cualquier anotación que indique la existencia de un proceso abierto como hipoteca, embargo o leasing.\nSi en una anotación se indica que un proceso fue cancelado o liberado, asegúrate de excluirlo de la lista de afectaciones vigentes.\nResponde exclusivamente en el formato especificado:\nSi todos los procesos están cancelados o liberados: Respuesta: \"Inmueble sin afectación\"\nSi hay afectaciones vigentes: Respuesta: \"Inmueble Afectado: [lista detallada de afectaciones vigentes]\"\nEjemplo: \"Inmueble Afectado: HIPOTECA-EMBARGO-LEASING\"\nConsideraciones:\nSi el contenido del CTL no permite completar la tarea, responde: Respuesta: \"No se pudo completar la tarea: [explicación del motivo]\"\nRealiza el análisis y responde únicamente en el formato indicado."
          }
        ]
      },
      "generationConfig": {
        "maxOutputTokens": 8192,
        "temperature": 0.5,
        "topP": 0.95,
        "seed": 0,
        "responseMimeType": "application/json",
        "responseSchema": { "type": "OBJECT", "properties": { "response": { "type": "STRING" } } }
      },
      "safetySettings": [
        {
          "category": "HARM_CATEGORY_HATE_SPEECH",
          "threshold": "OFF"
        },
        {
          "category": "HARM_CATEGORY_DANGEROUS_CONTENT",
          "threshold": "OFF"
        },
        {
          "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT",
          "threshold": "OFF"
        },
        {
          "category": "HARM_CATEGORY_HARASSMENT",
          "threshold": "OFF"
        }
      ],
    }

    var headers = {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + service.getAccessToken()
    };

    var options = {
      method: "POST",
      headers: headers,
      payload: JSON.stringify(payloadAgent1),
      url: url
    };

    var options2 = {
      method: "POST",
      headers: headers,
      payload: JSON.stringify(payloadAgent2),
      url: url
    };

    var requests = [options, options2];
    try {
      var response = UrlFetchApp.fetchAll(requests);
      var parsedResponse = response.map(r => JSON.parse(r.getContentText()));
      var fullText = cleanGeminiResponse(parsedResponse);
      Logger.log(fullText);
      return fullText;
    } catch (e) {
      Logger.log(e)
      return "Error al procesar respuesta"
    }
  }
}

///////// funcion retorno API Gemini /////////

function ResponseGeminiStream(responseArray) {

  // 1. Unir los fragmentos de texto
  // Usamos .reduce() para concatenar el campo 'text' de cada fragmento.
  const fullText = responseArray.reduce((accumulator, currentItem) => {
    try {
      // Navegamos al campo 'text' de forma segura
      const partText = currentItem?.candidates?.[0]?.content?.parts?.[0]?.text;

      if (partText) {
        return accumulator + partText;
      }
    } catch (e) {
      Logger.log("Error al procesar un fragmento: " + e.toString());
    }
    return accumulator;
  }, "");

  // 2. Parsear el texto unificado como JSON
  if (fullText.trim()) {
    try {
      // El texto unificado es el JSON completo (e.g., {"response": "..."})
      return JSON.parse(fullText);
    } catch (e) {
      Logger.log("Error al parsear el texto unificado como JSON: " + e.toString());
      Logger.log("Texto que causó el error: " + fullText);
      // Retornamos null o un objeto de error si el parseo falla
      return null;
    }
  }

  Logger.log("No se pudo extraer texto útil de la respuesta.");
  return null;
}

/////////////////////////////////////////////////////////////////////


function ModelValidationCTL2(Ref) {
  //var Ref = "BR-PRUEBA"
  var FilaData = SheetConsolidadoBrokersYInmobiliarias.getRange("C:C").createTextFinder(Ref).matchEntireCell(true).ignoreDiacritics(true).findPrevious().getRow();
  reset();
  var TipoCliente = SheetConsolidadoBrokersYInmobiliarias.getRange("I" + FilaData).getDisplayValue();
  if (TipoCliente == "Persona Natural") {
    var IdFiles = SheetConsolidadoBrokersYInmobiliarias.getRange("AZ" + FilaData).getDisplayValue();
    IdFiles = JSON.parse(IdFiles);
    var idCtlFinal = "";
    var idCedulaFinal = "";
    for (var key in IdFiles) {
      if (IdFiles[key].Tipo_Doc == "CTL Corregido") {
        var idFileCtl = IdFiles[key].IdFile;
        idCtlFinal = idFileCtl;
      } else if (IdFiles[key].Tipo_Doc == "CTL") {
        var idFileCtl = IdFiles[key].IdFile;
        idCtlFinal = idFileCtl;
      }
      if (IdFiles[key].Tipo_Doc == "Cédula Corregido") {
        var idFileCedula = IdFiles[key].IdFile;
        idCedulaFinal = idFileCedula;
      } else if (IdFiles[key].Tipo_Doc == "Cédula") {
        var idFileCedula = IdFiles[key].IdFile;
        idCedulaFinal = idFileCedula;
      }
    }
    if (idCtlFinal != "" && idFileCedula != "") {
      Logger.log(idCtlFinal)
      var FileIdCTL = DriveApp.getFileById(idCtlFinal);
      var FileBase64CTL = Utilities.base64Encode(FileIdCTL.getBlob().getBytes());
      var FileTypeCTL = FileIdCTL.getMimeType();
      var FileIdCedula = DriveApp.getFileById(idCedulaFinal);
      var FileBase64Cedula = Utilities.base64Encode(FileIdCedula.getBlob().getBytes());
      var FileTypeCedula = FileIdCedula.getMimeType();
      var service = getService();
      var API_ENDPOINT = "us-central1-aiplatform.googleapis.com"
      var PROJECT_ID = "proyecto-ia-servicios-bolivar"
      var MODEL_ID = "gemini-2.0-flash-lite-001"//"gemini-1.0-pro-vision-001";
      var url = "https://" + API_ENDPOINT + "/v1/projects/" + PROJECT_ID + "/locations/us-central1/publishers/google/models/" + MODEL_ID + ":streamGenerateContent";
      var payloadAgent1 = {
        "contents": [
          {
            "role": "user",
            "parts": [
              {
                "text": "Documento 1:"
              },
              {
                "inlineData": {
                  "mimeType": FileTypeCedula,
                  "data": FileBase64Cedula
                }
              },
              {
                "text": "Documento 2:"
              },
              {
                "inlineData": {
                  "mimeType": FileTypeCTL,
                  "data": FileBase64CTL
                }
              }
            ]
          }
        ],
        "systemInstruction": {
          "parts": [
            {
              "text": "Eres un experto analista de documentos, tu tarea es analizar dos documentos y dar una respuesta en formato json. El primer documento es la cédula de ciudadania, del cual debes extraer el numero de identificación, el segundo documento es un un Certificado de Tradición y Libertad (CTL), en este segundo docuemnto te debes concentrar en las anotaciones, generalmente inician con el texto \"ANOTACION: Nro\", debes validar si el número extraido del pirmer documento se encuentre en alguna anotación del segundo documento.\n\n    Instrucciones:\n    - Recibirás dos documentos.\n    - El primer documento es una Cédula de Ciudadanía.\n    - El segundo es el Certificado de tradición y libertad.\n    - Debes extraer el número de identificación del primer documento y buscarlo en alguna anotación del segundo documento.\n\n    FORMATO DE RESPUESTA:\n    - Respuesta tarea 1: \"Si coinciden\" (si se encuentra el número de identificación en alguna anotación.)\n    - Respuesta tarea 1: \"No coinciden\" (no registra el número en alguna anotación.)\n    - Las respuestas deben ser en español.\n    - Asegurate de responder correctamente con la forma establecida.\n\n    Si NO puedes procesar los documentos:\n    - \"No puedo leer los documentos: [RAZÓN ESPECÍFICA]\"\n\n    CONSIDERACIONES IMPORTANTES:\n    - Asegúrate de poder identificar claramente los números en ambos documentos\n    - No incluyas información adicional fuera del formato especificado\n    - La respuesta debe ser en formato JSON.\n\nEjemplo Respuesta:\n{\n\"Respuesta tarea\": \"Coincide\"\n}"
            }
          ]
        },
        "generationConfig": {
          "maxOutputTokens": 8192,
          "temperature": 0.5,
          "topP": 0.95,
          "seed": 0,
          "responseMimeType": "application/json",
          "responseSchema": { "type": "OBJECT", "properties": { "response": { "type": "STRING" } } }
        },
        "safetySettings": [
          {
            "category": "HARM_CATEGORY_HATE_SPEECH",
            "threshold": "OFF"
          },
          {
            "category": "HARM_CATEGORY_DANGEROUS_CONTENT",
            "threshold": "OFF"
          },
          {
            "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT",
            "threshold": "OFF"
          },
          {
            "category": "HARM_CATEGORY_HARASSMENT",
            "threshold": "OFF"
          }
        ],
      }
      var payloadAgent2 = {
        "contents": [
          {
            "role": "user",
            "parts": [
              {
                "text": "Certificado de tradición y libertad:"
              },
              {
                "inlineData": {
                  "mimeType": FileTypeCTL,
                  "data": FileBase64CTL
                }
              }
            ]
          }
        ],
        "systemInstruction": {
          "parts": [
            {
              "text": "Eres un Analista de Documentos, especializado en la revisión y validación de certificados de tradición y libertad (CTL). Se te proporciona un archivo adjunto que contiene el CTL de un inmueble. Tu tarea es analizar TODAS las anotaciones registradas en el documento para determinar si el inmueble tiene procesos abiertos que afecten su estado, tales como hipotecas, embargos o leasing.\nInstrucciones:\nRevisa todas las anotaciones del documento adjunto, desde la primera hasta la última.\nIdentifica cualquier anotación que indique la existencia de un proceso abierto como hipoteca, embargo o leasing.\nSi en una anotación se indica que un proceso fue cancelado o liberado, asegúrate de excluirlo de la lista de afectaciones vigentes.\nResponde exclusivamente en el formato especificado:\nSi todos los procesos están cancelados o liberados: Respuesta: \"Inmueble sin afectación\"\nSi hay afectaciones vigentes: Respuesta: \"Inmueble Afectado: [lista detallada de afectaciones vigentes]\"\nEjemplo: \"Inmueble Afectado: HIPOTECA-EMBARGO-LEASING\"\nConsideraciones:\nSi el contenido del CTL no permite completar la tarea, responde: Respuesta: \"No se pudo completar la tarea: [explicación del motivo]\"\nRealiza el análisis y responde únicamente en el formato indicado."
            }
          ]
        },
        "generationConfig": {
          "maxOutputTokens": 8192,
          "temperature": 0.5,
          "topP": 0.95,
          "seed": 0,
          "responseMimeType": "application/json",
          "responseSchema": { "type": "OBJECT", "properties": { "response": { "type": "STRING" } } }
        },
        "safetySettings": [
          {
            "category": "HARM_CATEGORY_HATE_SPEECH",
            "threshold": "OFF"
          },
          {
            "category": "HARM_CATEGORY_DANGEROUS_CONTENT",
            "threshold": "OFF"
          },
          {
            "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT",
            "threshold": "OFF"
          },
          {
            "category": "HARM_CATEGORY_HARASSMENT",
            "threshold": "OFF"
          }
        ],
      }

      var headers = {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + service.getAccessToken()
      };

      var options = {
        method: "POST",
        headers: headers,
        payload: JSON.stringify(payloadAgent1),
        url: url
      };

      var options2 = {
        method: "POST",
        headers: headers,
        payload: JSON.stringify(payloadAgent2),
        url: url
      };

      var requests = [options, options2];
      try {
        var response = UrlFetchApp.fetchAll(requests);
        var parsedResponse = response.map(r => JSON.parse(r.getContentText()));
        var fullText = cleanGeminiResponse(parsedResponse);
        Logger.log(fullText);
        //var fullText = '[{"response":"Si coinciden"},{"response":"Inmueble sin afectación"}]'
        //var fullText = '[{"response":"No coinciden"},{"response":"Inmueble Afectado: HIPOTECA"}]'
        return fullText;
      } catch (e) {
        Logger.log(e)
        return "Error al procesar respuesta"
      }
    } else {
      return "Error al procesar respuesta"
    }
  } else {
    var IdFiles = SheetConsolidadoBrokersYInmobiliarias.getRange("AZ" + FilaData).getDisplayValue();
    IdFiles = JSON.parse(IdFiles);
    var idCtlFinal = "";
    var idCedulaFinal = "";
    var idCertificacionFinal = "";
    var idRutFinal = "";
    for (var key in IdFiles) {
      if (IdFiles[key].Tipo_Doc == "CTL Corregido") {
        var idFileCtl = IdFiles[key].IdFile;
        idCtlFinal = idFileCtl;
      } else if (IdFiles[key].Tipo_Doc == "CTL") {
        var idFileCtl = IdFiles[key].IdFile;
        idCtlFinal = idFileCtl;
      }
      if (IdFiles[key].Tipo_Doc == "Cédula Corregido") {
        var idFileCedula = IdFiles[key].IdFile;
        idCedulaFinal = idFileCedula;
      } else if (IdFiles[key].Tipo_Doc == "Cédula") {
        var idFileCedula = IdFiles[key].IdFile;
        idCedulaFinal = idFileCedula;
      }
      if (IdFiles[key].Tipo_Doc == "Rut Corregido") {
        var idFileRut = IdFiles[key].IdFile;
        idRutFinal = idFileRut;
      } else if (IdFiles[key].Tipo_Doc == "Rut") {
        var idFileRut = IdFiles[key].IdFile;
        idRutFinal = idFileRut;
      }
      if (IdFiles[key].Tipo_Doc == "Certificacion Corregido") {
        var idFileCertificacion = IdFiles[key].IdFile;
        idCertificacionFinal = idFileCertificacion;
      } else if (IdFiles[key].Tipo_Doc == "Certificacion") {
        var idFileCertificacion = IdFiles[key].IdFile;
        idCertificacionFinal = idFileCertificacion;
      }
    }
    if (idCtlFinal != "" && idCedulaFinal != "" && idCertificacionFinal != "" && idRutFinal != "") {
      var FileIdCTL = DriveApp.getFileById(idCtlFinal);
      var FileBase64CTL = Utilities.base64Encode(FileIdCTL.getBlob().getBytes());
      var FileTypeCTL = FileIdCTL.getMimeType();
      var FileIdCedula = DriveApp.getFileById(idCedulaFinal);
      var FileBase64Cedula = Utilities.base64Encode(FileIdCedula.getBlob().getBytes());
      var FileTypeCedula = FileIdCedula.getMimeType();
      var FileIdRut = DriveApp.getFileById(idRutFinal);
      var FileBase64Rut = Utilities.base64Encode(FileIdRut.getBlob().getBytes());
      var FileTypeRut = FileIdRut.getMimeType();
      var FileIdRepLegal = DriveApp.getFileById(idCertificacionFinal);
      var FileBase64RepLegal = Utilities.base64Encode(FileIdRepLegal.getBlob().getBytes());
      var FileTypeRepLegal = FileIdRepLegal.getMimeType();
      var service = getService();
      var API_ENDPOINT = "us-central1-aiplatform.googleapis.com"
      var PROJECT_ID = "proyecto-ia-servicios-bolivar"
      var MODEL_ID = "gemini-2.0-flash-lite-001";//gemini-1.0-pro-vision-001
      var url = "https://" + API_ENDPOINT + "/v1/projects/" + PROJECT_ID + "/locations/us-central1/publishers/google/models/" + MODEL_ID + ":streamGenerateContent";

      var payloadAgent1 = {
        "contents": [
          {
            "role": "user",
            "parts": [
              {
                "text": "Cédula de Ciudadanía:"
              },
              {
                "inlineData": {
                  "mimeType": FileTypeCedula,
                  "data": FileBase64Cedula
                }
              },
              {
                "text": "Registro Único Tributario (RUT):"
              },
              {
                "inlineData": {
                  "mimeType": FileTypeRut,
                  "data": FileBase64Rut
                }
              },
              {
                "text": "Certificado de Existencia y Representación Legal:"
              },
              {
                "inlineData": {
                  "mimeType": FileTypeRepLegal,
                  "data": FileBase64RepLegal
                }
              }
            ]
          }
        ],
        "systemInstruction": {
          "parts": [
            {
              "text": "Eres un Analista de Documentos, especializado en la validación de datos de identificación. Se te proporcionan tres documentos: Cédula de Ciudadanía, Registro Único Tributario (RUT) y Certificado de Existencia y Representación Legal. Tu tarea es extraer el número de identificación de la Cédula de Ciudadanía y verificar si este número aparece en los otros dos documentos.\nInstrucciones:\nExtrae el número de identificación del documento de la Cédula de Ciudadanía.\nBusca el número de identificación en los otros documentos siguiendo estas guías:\nEn el RUT, utiliza como referencia las secciones donde se mencione \"Representación\" o \"REPRS LEGAL PRIN\".\nEn el Certificado de Existencia y Representación Legal, utiliza como referencia la sección que menciona \"REPRESENTANTES LEGALES\".\nCompara los resultados:\nSi el número de identificación aparece en ambos documentos (RUT y Certificado de Existencia y Representación Legal), responde:\nRespuesta: \"Si coinciden\"\nSi el número de identificación no aparece en uno de los documentos, responde:\nRespuesta: \"No coinciden\"\nConsideraciones:\nSi no puedes completar la tarea por falta de información o porque el contenido de los documentos no es claro, responde:\nRespuesta: \"No se pudo completar la tarea: [explicación del motivo]\"\nFormato de respuesta:\nResponde únicamente en el formato especificado:\nCoinciden\nNo coinciden\nRealiza el análisis y responde."
            }
          ]
        },
        "generationConfig": {
          "maxOutputTokens": 8192,
          "temperature": 0.5,
          "topP": 0.95,
          "seed": 0,
          "responseMimeType": "application/json",
          "responseSchema": { "type": "OBJECT", "properties": { "response": { "type": "STRING" } } }
        },
        "safetySettings": [
          {
            "category": "HARM_CATEGORY_HATE_SPEECH",
            "threshold": "OFF"
          },
          {
            "category": "HARM_CATEGORY_DANGEROUS_CONTENT",
            "threshold": "OFF"
          },
          {
            "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT",
            "threshold": "OFF"
          },
          {
            "category": "HARM_CATEGORY_HARASSMENT",
            "threshold": "OFF"
          }
        ],
      }

      var payloadAgent2 = {
        "contents": [
          {
            "role": "user",
            "parts": [
              {
                "text": "Certificado de tradición y libertad:"
              },
              {
                "inlineData": {
                  "mimeType": FileTypeCTL,
                  "data": FileBase64CTL
                }
              }
            ]
          }
        ],
        "systemInstruction": {
          "parts": [
            {
              "text": "Eres un Analista de Documentos, especializado en la revisión y validación de certificados de tradición y libertad (CTL). Se te proporciona un archivo adjunto que contiene el CTL de un inmueble. Tu tarea es analizar TODAS las anotaciones registradas en el documento para determinar si el inmueble tiene procesos abiertos que afecten su estado, tales como hipotecas, embargos o leasing.\nInstrucciones:\nRevisa todas las anotaciones del documento adjunto, desde la primera hasta la última.\nIdentifica cualquier anotación que indique la existencia de un proceso abierto como hipoteca, embargo o leasing.\nSi en una anotación se indica que un proceso fue cancelado o liberado, asegúrate de excluirlo de la lista de afectaciones vigentes.\nResponde exclusivamente en el formato especificado:\nSi todos los procesos están cancelados o liberados: Respuesta: \"Inmueble sin afectación\"\nSi hay afectaciones vigentes: Respuesta: \"Inmueble Afectado: [lista detallada de afectaciones vigentes]\"\nEjemplo: \"Inmueble Afectado: HIPOTECA-EMBARGO-LEASING\"\nConsideraciones:\nSi el contenido del CTL no permite completar la tarea, responde: Respuesta: \"No se pudo completar la tarea: [explicación del motivo]\"\nRealiza el análisis y responde únicamente en el formato indicado."
            }
          ]
        },
        "generationConfig": {
          "maxOutputTokens": 8192,
          "temperature": 0.5,
          "topP": 0.95,
          "seed": 0,
          "responseMimeType": "application/json",
          "responseSchema": { "type": "OBJECT", "properties": { "response": { "type": "STRING" } } }
        },
        "safetySettings": [
          {
            "category": "HARM_CATEGORY_HATE_SPEECH",
            "threshold": "OFF"
          },
          {
            "category": "HARM_CATEGORY_DANGEROUS_CONTENT",
            "threshold": "OFF"
          },
          {
            "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT",
            "threshold": "OFF"
          },
          {
            "category": "HARM_CATEGORY_HARASSMENT",
            "threshold": "OFF"
          }
        ],
      }

      var headers = {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + service.getAccessToken()
      };

      var options = {
        method: "POST",
        headers: headers,
        payload: JSON.stringify(payloadAgent1),
        url: url
      };

      var options2 = {
        method: "POST",
        headers: headers,
        payload: JSON.stringify(payloadAgent2),
        url: url
      };

      var requests = [options, options2];
      try {
        var response = UrlFetchApp.fetchAll(requests);
        var parsedResponse = response.map(r => JSON.parse(r.getContentText()));
        var fullText = cleanGeminiResponse(parsedResponse);
        Logger.log(fullText);
        return fullText;
      } catch (e) {
        Logger.log(e)
        return "Error al procesar respuesta"
      }
    } else {
      return "Error al procesar respuesta"
    }
  }
}

function cleanGeminiResponse(response) {
  function extractTextFromCandidates(candidatesArray) {
    return candidatesArray.flatMap(item =>
      item?.candidates?.[0]?.content?.parts?.map(part => part.text) || []
    ).join('');
  }
  let result = response.map(group => {
    try {
      const combinedText = extractTextFromCandidates(group);
      return JSON.parse(combinedText);
    } catch (e) {
      return { response: "Error al procesar respuesta" };
    }
  });
  return JSON.stringify(result);
}

function extractText(candidate) {
  if (candidate.content && candidate.content.parts) {
    return candidate.content.parts.map(part => part.text).join('');
  }
  return '';
}

function GenerarContrato(Ref, TipoContrato, NumeroMatricula) {
  var FilaData = SheetConsolidado.getRange("B:B").createTextFinder(Ref).matchEntireCell(true).ignoreDiacritics(true).findPrevious().getRow();
  if (TipoContrato == "Persona natural") {
    var DataFinal = [];
    var Dato1 = SheetConsolidado.getRange("E" + FilaData).getDisplayValue();
    var Dato2 = SheetConsolidado.getRange("F" + FilaData).getDisplayValue();
    var Dato3 = SheetConsolidado.getRange("AI" + FilaData).getDisplayValue();
    var Dato4 = SheetConsolidado.getRange("AJ" + FilaData).getDisplayValue();
    var Dato5 = SheetConsolidado.getRange("M" + FilaData).getDisplayValue();
    var Dato6 = SheetConsolidado.getRange("I" + FilaData).getDisplayValue();
    var Dato7 = SheetConsolidado.getRange("N" + FilaData).getDisplayValue();
    var Dato8 = NumeroALetras(Dato7);
    var Dato9 = ConvertirMeses(SheetConsolidado.getRange("Q" + FilaData).getDisplayValue());
    var Dato10 = SheetConsolidado.getRange("R" + FilaData).getDisplayValue().split("-")[2];
    var Dato11 = MesesAText(SheetConsolidado.getRange("R" + FilaData).getDisplayValue().split("-")[1]);
    var Dato12 = SheetConsolidado.getRange("R" + FilaData).getDisplayValue().split("-")[0];
    var Dato13 = SheetConsolidado.getRange("AL" + FilaData).getDisplayValue();
    var Dato14 = SheetConsolidado.getRange("AM" + FilaData).getDisplayValue();
    var Dato15 = SheetConsolidado.getRange("AN" + FilaData).getDisplayValue();
    var Dato16 = SheetConsolidado.getRange("AP" + FilaData).getDisplayValue();
    var Dato18 = SheetConsolidado.getRange("AO" + FilaData).getDisplayValue();
    DataFinal.push(Dato1, Dato2, Dato3, Dato4, Dato5, Dato6, Dato7, Dato8, Dato9, Dato10, Dato11, Dato12, NumeroMatricula, Dato13, Dato14, Dato15, Dato16, Dato18);
    return { IdContrato: "16kv23KUdb6Fdot2dFb3VKERgLNBz6rAr", DataFinal: DataFinal }
  } else {
    var DataFinal = [];
    var Dato1 = SheetConsolidado.getRange("E" + FilaData).getDisplayValue();
    var Dato2 = SheetConsolidado.getRange("F" + FilaData).getDisplayValue();
    var Dato3 = SheetConsolidado.getRange("AI" + FilaData).getDisplayValue();
    var Dato4 = SheetConsolidado.getRange("AJ" + FilaData).getDisplayValue();
    var Dato5 = SheetConsolidado.getRange("M" + FilaData).getDisplayValue();
    var Dato6 = SheetConsolidado.getRange("I" + FilaData).getDisplayValue();
    var Dato7 = SheetConsolidado.getRange("N" + FilaData).getDisplayValue();
    var Dato8 = NumeroALetras(Dato7);
    var Dato9 = ConvertirMeses(SheetConsolidado.getRange("Q" + FilaData).getDisplayValue());
    var Dato10 = SheetConsolidado.getRange("R" + FilaData).getDisplayValue().split("-")[2];
    var Dato11 = MesesAText(SheetConsolidado.getRange("R" + FilaData).getDisplayValue().split("-")[1]);
    var Dato12 = SheetConsolidado.getRange("R" + FilaData).getDisplayValue().split("-")[0];
    var Dato13 = SheetConsolidado.getRange("AL" + FilaData).getDisplayValue();
    var Dato14 = SheetConsolidado.getRange("AM" + FilaData).getDisplayValue();
    var Dato15 = SheetConsolidado.getRange("AN" + FilaData).getDisplayValue();
    var Dato16 = SheetConsolidado.getRange("AP" + FilaData).getDisplayValue();
    var Dato18 = SheetConsolidado.getRange("AO" + FilaData).getDisplayValue();
    DataFinal.push(Dato1, Dato2, Dato3, Dato4, Dato5, Dato6, Dato7, Dato8, Dato9, Dato10, Dato11, Dato12, NumeroMatricula, Dato13, Dato14, Dato15, Dato16, Dato18);
    return { IdContrato: "16kv23KUdb6Fdot2dFb3VKERgLNBz6rAr", DataFinal: DataFinal }
  }
}

function NumeroALetras(num) {
  var unidades = ["", "uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve"];
  var decenas = ["", "diez", "veinte", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa"];
  var centenas = ["", "cien", "doscientos", "trescientos", "cuatrocientos", "quinientos", "seiscientos", "setecientos", "ochocientos", "novecientos"];

  var especiales = {
    11: "once", 12: "doce", 13: "trece", 14: "catorce", 15: "quince",
    16: "dieciséis", 17: "diecisiete", 18: "dieciocho", 19: "diecinueve"
  };

  function convertir(num) {
    if (num === 0) return "cero";
    if (num < 10) return unidades[num];
    if (num < 20) return especiales[num];
    if (num < 100) {
      return decenas[Math.floor(num / 10)] + (num % 10 > 0 ? " y " + unidades[num % 10] : "");
    }
    if (num < 1000) {
      return (num === 100 ? "cien" : centenas[Math.floor(num / 100)] + (num % 100 > 0 ? " " + convertir(num % 100) : ""));
    }
    if (num < 1000000) {
      let miles = Math.floor(num / 1000);
      let resto = num % 1000;
      return (miles === 1 ? "mil" : convertir(miles) + " mil") + (resto > 0 ? " " + convertir(resto) : "");
    }
    if (num < 1000000000) {
      let millones = Math.floor(num / 1000000);
      let resto = num % 1000000;
      return (millones === 1 ? "un millón" : convertir(millones) + " millones") + (resto > 0 ? " " + convertir(resto) : "");
    }
    return "Número fuera de rango";
  }

  // Limpiar la entrada y convertir a número entero
  var numStr = num.toString().replace(/[^0-9]/g, '');
  var numInt = parseInt(numStr, 10);

  if (isNaN(numInt) || numInt < 0) return "Número inválido";

  var resultado = convertir(numInt) + " pesos";
  resultado = resultado.charAt(0).toUpperCase() + resultado.slice(1);
  return resultado;
}


function ConvertirMeses(meses) {
  meses = meses.split(" ")[0];
  var mesesTexto = [
    "UN (01) mes", "DOS (02) meses", "TRES (03) meses", "CUATRO (04) meses", "CINCO (05) meses",
    "SEIS (06) meses", "SIETE (07) meses", "OCHO (08) meses", "NUEVE (09) meses", "DIEZ (10) meses",
    "ONCE (11) meses", "DOCE (12) meses"
  ];
  var resultado = mesesTexto[meses - 1];
  return resultado;
}

function MesesAText(mes) {
  var mesesTexto = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO",
    "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE",
    "NOVIEMBRE", "DICIEMBRE"
  ];
  var resultado = mesesTexto[mes - 1];
  return resultado;
}

function EnviarCorreccionDocs(referenciasDocs, observacionesFinales, Ref) {
  var FilaData = SheetConsolidado.getRange("B:B").createTextFinder(Ref).matchEntireCell(true).ignoreDiacritics(true).findPrevious().getRow();
  var Correo = SheetConsolidado.getRange("D" + FilaData).getDisplayValue();
  var observacionesText = "";
  var htmlObservaciones = "";
  for (var i = 0; i < referenciasDocs.length; i++) {
    var docReferencia = referenciasDocs[i];
    var observacion = observacionesFinales[i];
    if (observacion) {
      if (observacionesText !== "") {
        observacionesText += ',';
      }
      observacionesText += `"{${docReferencia},${observacion}}"`;
      htmlObservaciones += `<p><b>${docReferencia}:</b> ${observacion}</p>`;
    }
  }
  SheetConsolidado.getRange("AU" + FilaData).setValue(observacionesText);
  var correoAsesor = SheetConsolidado.getRange("BC" + FilaData).getDisplayValue();
  var Pieza = HtmlService.createHtmlOutputFromFile("Pieza_Correcion_Docs_Correo").getContent();
  Pieza = Pieza.replace('"Reemplazar"', htmlObservaciones);
  GmailApp.sendEmail(Correo, "Corrección de Documentos Poliza Libertador " + Ref, "", { htmlBody: Pieza, noReply: true, cc: correoAsesor, bcc: "paola.garcia@aselibertador.com" });
  SheetConsolidado.getRange("AV" + FilaData).setValue(new Date());
  SheetConsolidado.getRange("Y" + FilaData).setValue('Pendiente Corrección Documental');
}


function EnviarContratoFirma(IdContrato, Ref, IdPoliza) {
  var FilaData = SheetConsolidado.getRange("B:B").createTextFinder(Ref).matchEntireCell(true).ignoreDiacritics(true).findPrevious().getRow();
  var Correo = SheetConsolidado.getRange("D" + FilaData).getDisplayValue();
  var correoClienteNormalizado = Correo.trim().replace(/ñ/g, 'n');
  var nombreUsuario = SheetConsolidado.getRange("E" + FilaData).getDisplayValue();
  var url = "https://www.googleapis.com/drive/v3/files/" + IdContrato + "/export?mimeType=application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  var opciones = {
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    }
  };
  var respuesta = UrlFetchApp.fetch(url, opciones);
  var blob = respuesta.getBlob().setName("Contrato_" + Ref + ".docx");
  Logger.log(IdPoliza)
  var Poliza = DriveApp.getFileById(IdPoliza);
  var Inventario = DriveApp.getFileById("1xy6OAB0QpCW86VD-qBJMn9Ku-NzxY6tR");
  var Clausulado = DriveApp.getFileById("1hX4qLiv20MBMWUVoj5UMyjvMdUge4-rt");
  var Files = [];
  Files.push(blob, Poliza, Inventario, Clausulado);
  var Pieza = HtmlService.createHtmlOutputFromFile("Pieza_Contrato_Correo").getContent();
  Pieza = Pieza.replace('[nombre usuario]', nombreUsuario);
  var correoAsesor = SheetConsolidado.getRange("BC" + FilaData).getDisplayValue();
  GmailApp.sendEmail(correoClienteNormalizado, "Firma de Contrato Poliza Libertador " + Ref, "", { htmlBody: Pieza, attachments: Files, noReply: true, cc: correoAsesor, bcc: "paola.garcia@aselibertador.com" });
  SheetConsolidado.getRange("Y" + FilaData).setValue("Expedido");
  SheetConsolidado.getRange("AV" + FilaData).setValue(new Date());
}

function CargarPoliza(form) {
  var Ref = form["RefCargarPoliza"];
  var FilaData = SheetConsolidado.getRange("B:B").createTextFinder(Ref).ignoreDiacritics(true).matchEntireCell(true).findPrevious().getRow();
  var Folder = SheetConsolidado.getRange("AS" + FilaData).getDisplayValue().split("/folders/")[1];
  var File1 = form["Dato20Contrato"];
  var TypeFile1 = form["Dato20Contrato"].name;
  var MimeTypeFile1 = TypeFile1.split(".")[1].toUpperCase();
  var NameFile1 = "Poliza-" + Ref;
  var resource = {
    title: NameFile1,
    mimeType: MimeTypeFile1,
    parents: [{ id: Folder }]
  };
  var FileCedula = Drive.Files.insert(resource, File1, {
    convert: false
  });
  var IdPoliza = FileCedula.id;
  return IdPoliza
}

function GenerarContratoFinal(deudores, contratoDatos, Ref) {
  var Dato1 = contratoDatos["Nombre Arrendador"];
  var Dato2 = contratoDatos["Documento Arrendador"];
  var Dato3 = contratoDatos["Nombre Arrendatario"];
  var Dato4 = contratoDatos["Documento Arrendatario"];
  var Dato5 = contratoDatos["Dirección Inmueble"];
  var Dato6 = contratoDatos["Ciudad Inmueble"];
  var Dato7 = contratoDatos["Ciudad Oficina de Registro"];
  var Dato8 = contratoDatos["Matricula Inmobiliaria"];
  var Dato9 = contratoDatos["Valor Canon Arrendamiento"];
  var Dato10 = contratoDatos["Valor Canon Arrendamiento(texto)"];
  var Dato11 = contratoDatos["Tipo Cuenta"];
  var Dato12 = contratoDatos["No. Cuenta"];
  var Dato13 = contratoDatos["Banco de la Cuenta"];
  var Dato14 = contratoDatos["Nombre del Titular de la Cuenta"];
  var Dato15 = contratoDatos["Vigencia de la poliza"];
  var Dato16 = contratoDatos["Día Inicio Póliza"];
  var Dato17 = contratoDatos["Mes Inicio Póliza"];
  var Dato18 = contratoDatos["Año Inicio Póliza"];
  var Dato19 = contratoDatos["Nro Solicitud Estudio"];
  var Dato20 = Utilities.formatDate(new Date(), "GMT-05:00", "yy");
  var FilaData = SheetConsolidado.getRange("B:B").createTextFinder(Ref).ignoreDiacritics(true).matchEntireCell(true).findPrevious().getRow();
  var ValidarTipoPoliza = SheetConsolidado.getRange("L" + FilaData).getDisplayValue();
  var ValorAdmin = SheetConsolidado.getRange("O" + FilaData).getDisplayValue();
  if (ValidarTipoPoliza == "Vivienda") {
    if (ValorAdmin == "0") {
      var IdFormatoContrato = "1rA7G86JBJQA-T60laqvHiMgH2oz0uy1-gwvzW4B8tWA";
    } else {
      var IdFormatoContrato = "16Qk_XtwfbUyZXk-ks0j8HqV_H8o_5Lq63gE0e03a0yI";
    }
  } else {
    if (ValorAdmin == "0") {
      var IdFormatoContrato = "1D414uYf-bPpLdXSgrpbopJHZuT12aG7rfHFSfukOm9E";
    } else {
      var IdFormatoContrato = "1TK8PJMoFXAQeHCItNj6RM_29dzjtLt002XwBSN6xLCU";
    }
  }
  var idFolder = SheetConsolidado.getRange("AS" + FilaData).getDisplayValue();
  idFolder = idFolder.split("/folders/")[1];
  var CopiaFormato = DriveApp.getFileById(IdFormatoContrato).makeCopy("Contrato-" + Ref, DriveApp.getFolderById(idFolder)).getId();
  var ContratoFinal = DocumentApp.openById(CopiaFormato);
  if (deudores.length < 1) {
    ContratoFinal.getBody().replaceText('“TITULO DEUDOR  SOLIDARIO FIRMA”', '');
    procesarDeudores(CopiaFormato, deudores);
    procesarSeccionFinalDeudores(CopiaFormato, deudores);
  } else {
    ContratoFinal.getBody().replaceText('“TITULO DEUDOR  SOLIDARIO FIRMA”', 'DEUDORES  SOLIDARIOS');
    procesarDeudores(CopiaFormato, deudores);
    procesarSeccionFinalDeudores(CopiaFormato, deudores);
  }
  ContratoFinal.getBody().replaceText('“Dato1”', Dato1)
  ContratoFinal.getBody().replaceText('“Dato2”', Dato2)
  ContratoFinal.getBody().replaceText('“Dato3”', Dato3)
  ContratoFinal.getBody().replaceText('“Dato4”', Dato4)
  ContratoFinal.getBody().replaceText('“Dato5”', Dato5)
  ContratoFinal.getBody().replaceText('“Dato6”', Dato6)
  ContratoFinal.getBody().replaceText('“Dato7”', Dato7)
  ContratoFinal.getBody().replaceText('“Dato8”', Dato8)
  ContratoFinal.getBody().replaceText('“Dato9”', Dato9);
  ContratoFinal.getBody().replaceText('“Dato10”', Dato10);
  ContratoFinal.getBody().replaceText('“Dato11”', Dato11);
  ContratoFinal.getBody().replaceText('“Dato12”', Dato12);
  ContratoFinal.getBody().replaceText('“Dato13”', Dato13);
  ContratoFinal.getBody().replaceText('“Dato14”', Dato14);
  ContratoFinal.getBody().replaceText('“Dato24”', Dato15);
  ContratoFinal.getBody().replaceText('“Dato15”', Dato16);
  ContratoFinal.getBody().replaceText('“Dato16”', Dato17);
  ContratoFinal.getBody().replaceText('“Dato17”', Dato18);
  ContratoFinal.getBody().replaceText('“Dato20”', Dato3);
  ContratoFinal.getBody().replaceText('“Dato21”', Dato3);
  ContratoFinal.getBody().replaceText('“Dato22”', Dato1);
  ContratoFinal.getBody().replaceText('“Dato23”', Dato1);
  ContratoFinal.getBody().replaceText('“Dato25”', Dato1);
  ContratoFinal.getBody().replaceText('“Dato26”', Dato2);
  ContratoFinal.getBody().replaceText('“Dato27”', Dato3);
  ContratoFinal.getBody().replaceText('“Dato28”', Dato4);
  ContratoFinal.getBody().replaceText('“Dato31”', Dato1);
  ContratoFinal.getBody().replaceText('“Dato32”', Dato3);
  ContratoFinal.getFooter().replaceText('“Dato33”', Dato19);
  ContratoFinal.getFooter().replaceText('“Dato34”', Dato20);
  var Fecha = formatearFechaEnEspanol();
  ContratoFinal.getBody().replaceText('“Dato35”', Fecha.dia);
  ContratoFinal.getBody().replaceText('“Dato36”', Fecha.mes);
  ContratoFinal.getBody().replaceText('“Dato37”', Fecha.year);
  if (ValorAdmin != "0") {
    var NumLetrasAdmin = NumeroALetras(ValorAdmin);
    ContratoFinal.getBody().replaceText('“DatoAdminPesos”', ValorAdmin);
    ContratoFinal.getBody().replaceText('“DatoAdminLetras”', NumLetrasAdmin);
  }
  if (ValidarTipoPoliza == "Comercio") {
    var Destino = SheetConsolidado.getRange("BD" + FilaData).getDisplayValue();
    ContratoFinal.getBody().replaceText('“DESTINOCOMERCIO”', Destino);
  }
  var Servicios = SheetConsolidado.getRange("AQ" + FilaData).getDisplayValue();
  ContratoFinal.getBody().replaceText('“servicios”', Servicios);
  ContratoFinal.saveAndClose();
  return { IdContratoFinal: ContratoFinal.getId() }
}

function procesarDeudores(id, deudores) {
  var body = DocumentApp.openById(id).getBody();
  var paragraphs = body.getParagraphs();
  var bloques = [];
  var inicio = null;
  for (var i = 0; i < paragraphs.length; i++) {
    var text = paragraphs[i].getText();
    if (text.indexOf('InicioDeudor') !== -1) {
      inicio = i;
    }
    if (text.indexOf('FinDeudor') !== -1 && inicio !== null) {
      bloques.push({ inicio: inicio, fin: i });
      inicio = null;
    }
  }
  for (var j = 0; j < bloques.length; j++) {
    var bloque = bloques[j];
    if (j < deudores.length) {
      var nombre = deudores[j][0];
      var cedula = deudores[j][1];
      var deudorTexto = `${nombre} con C.C. No. ${cedula}`;
      paragraphs[bloque.inicio + 1].setText(deudorTexto);
      paragraphs[bloque.fin].removeFromParent();
      paragraphs[bloque.inicio].removeFromParent();
    } else {
      for (var k = bloque.fin; k >= bloque.inicio; k--) {
        paragraphs[k].removeFromParent();
      }
    }
  }
}

function procesarSeccionFinalDeudores(id, deudores) {
  var body = DocumentApp.openById(id).getBody();
  for (var i = 0; i < deudores.length; i++) {
    var nombrePlaceholder = `NombreDeudor${i + 1}Final`;
    var cedulaPlaceholder = `DocumentoDeudor${i + 1}Final`;
    body.replaceText(nombrePlaceholder, deudores[i][0]);
    body.replaceText(cedulaPlaceholder, deudores[i][1]);
  }
  for (var j = deudores.length + 1; j <= 5; j++) {
    var inicioPlaceholder = `FirmaDeudorFinal${j}`;
    var finPlaceholder = `FirmaFinDeudorFinal${j}`;
    var startIndex = body.getText().indexOf(inicioPlaceholder);
    var endIndex = body.getText().indexOf(finPlaceholder) + finPlaceholder.length;
    if (startIndex !== -1 && endIndex !== -1) {
      body.editAsText().deleteText(startIndex, endIndex - 1);
    }
  }
  for (var j = 1; j <= 5; j++) {
    var inicioPlaceholder = `FirmaDeudorFinal${j}`;
    var finPlaceholder = `FirmaFinDeudorFinal${j}`;
    body.replaceText(inicioPlaceholder, '');
    body.replaceText(finPlaceholder, '');
  }
}

function DesistirCaso(Ref) {
  var FilaData = SheetConsolidado.getRange("B:B").createTextFinder(Ref).matchEntireCell(true).ignoreDiacritics(true).findPrevious().getRow();
  SheetConsolidado.getRange("Y" + FilaData).setValue("Desistido");
}

function DesistirCasoBrokerInmobiliaria(Ref) {
  var FilaData = SheetConsolidadoBrokersYInmobiliarias.getRange("C:C").createTextFinder(Ref).matchEntireCell(true).ignoreDiacritics(true).findPrevious().getRow();
  SheetConsolidadoBrokersYInmobiliarias.getRange("BA" + FilaData).setValue("Desistido");
  SheetConsolidadoBrokersYInmobiliarias.getRange("BG" + FilaData).setValue(new Date());
}


function GetDataBrokersYInmobiliarias() {
  var UserMail = Session.getActiveUser().getEmail();
  var DataRange = SheetConsolidadoBrokersYInmobiliarias.getRange("A2:BF" + SheetConsolidado.getLastRow()).getDisplayValues();
  var DataRangeUserPending = [];
  DataRange.filter(function (DataRange) {
    var Validation = DataRange[52].indexOf("Pendiente Validación Documental");
    var Validation2 = DataRange[57].indexOf(UserMail);
    if (Validation > -1 && Validation2 > -1) {
      var ValorDanos = formatearAEntero(DataRange[49]);
      var PrimaDanos = formatNumberInput(CalculatePrimaDanos(ValorDanos).toString())
      var ValorServicios = formatearAEntero(DataRange[48]);
      var PrimaServicios = formatNumberInput(CalculatePrimaServicios(DataRange[24], ValorServicios).toString());
      Logger.log(PrimaServicios)
      DataRangeUserPending.push([DataRange[0], DataRange[11], DataRange[19], DataRange[12], DataRange[1], "", DataRange[2], DataRange[24], DataRange[28], DataRange[45], DataRange[10], DataRange[8], DataRange[9], DataRange[13], DataRange[18], DataRange[33], DataRange[26], DataRange[27], DataRange[28], DataRange[25], DataRange[14], DataRange[48], DataRange[49], DataRange[51], DataRange[17], DataRange[56], DataRange[30], DataRange[31]]);
    }
  });
  return DataRangeUserPending;
}

function CalculatePrimaDanos(ValorDanosAsegurar) {
  var PrimaNetaDanos = ValorDanosAsegurar * 0.05;
  var Iva19Danos = (PrimaNetaDanos * 0.19);
  var PrimaPagarDanos = parseInt(PrimaNetaDanos + Iva19Danos);
  return PrimaPagarDanos;
}

function CalculatePrimaServicios(TypeInmueble, ValorServiciosAsegurado) {
  var PrimaPagarServicios;
  if (TypeInmueble.includes("Vivienda")) {
    // var ValorRegalo = 1000000;
    if (ValorServiciosAsegurado < 90000000) {
      var PrimaNetaServicios = ValorServiciosAsegurado * 0.05;
    } else {
      var PrimaNetaServicios = 0;
    }
    var Iva19Servicios = (PrimaNetaServicios * 0.19);
    PrimaPagarServicios = parseInt(PrimaNetaServicios + Iva19Servicios);
  } else if (TypeInmueble.includes("Comercio")) {
    // var ValorRegalo = 2000000;
    if (ValorServiciosAsegurado < 180000000) {
      var PrimaNetaServicios = ValorServiciosAsegurado * 0.05;
    } else {
      var PrimaNetaServicios = 0;
    }
    var Iva19Servicios = (PrimaNetaServicios * 0.19);
    PrimaPagarServicios = parseInt(PrimaNetaServicios + Iva19Servicios);
  }
  return PrimaPagarServicios;
}

function formatearAEntero(valor) {
  var formato = valor.replace(/[^0-9]/g, '');
  return parseInt(formato);
}

function formatNumberInput(input) {
  const value = input.replace(/[^0-9]/g, '');
  const formattedValue = '$' + new Intl.NumberFormat('es-CO').format(Number(value));
  return formattedValue;
}

function GenerarContratoBrokersInmobiliarias(Ref, NumeroMatricula) {
  var FilaData = SheetConsolidadoBrokersYInmobiliarias.getRange("C:C").createTextFinder(Ref).matchEntireCell(true).ignoreDiacritics(true).findPrevious().getRow();
  var DataFinal = [];
  var Dato1 = SheetConsolidadoBrokersYInmobiliarias.getRange("L" + FilaData).getDisplayValue();
  var Dato2 = SheetConsolidadoBrokersYInmobiliarias.getRange("K" + FilaData).getDisplayValue();
  var Dato3 = ""
  var Dato4 = ""
  var Dato5 = SheetConsolidadoBrokersYInmobiliarias.getRange("S" + FilaData).getDisplayValue();
  var Dato6 = SheetConsolidadoBrokersYInmobiliarias.getRange("T" + FilaData).getDisplayValue();
  var Dato7 = SheetConsolidadoBrokersYInmobiliarias.getRange("AA" + FilaData).getDisplayValue();
  var Dato8 = NumeroALetras(Dato7);
  var Dato9 = ConvertirMeses(SheetConsolidadoBrokersYInmobiliarias.getRange("Z" + FilaData).getDisplayValue());
  var Dato10 = SheetConsolidadoBrokersYInmobiliarias.getRange("AH" + FilaData).getDisplayValue().split("-")[2];
  var Dato11 = MesesAText(SheetConsolidadoBrokersYInmobiliarias.getRange("AH" + FilaData).getDisplayValue().split("-")[1]);
  var Dato12 = SheetConsolidadoBrokersYInmobiliarias.getRange("AH" + FilaData).getDisplayValue().split("-")[0];
  var Dato13 = SheetConsolidadoBrokersYInmobiliarias.getRange("H" + FilaData).getDisplayValue();
  var Dato14 = SheetConsolidadoBrokersYInmobiliarias.getRange("AN" + FilaData).getDisplayValue();
  var Dato15 = SheetConsolidadoBrokersYInmobiliarias.getRange("AO" + FilaData).getDisplayValue();
  var Dato16 = SheetConsolidadoBrokersYInmobiliarias.getRange("AM" + FilaData).getDisplayValue();
  var Dato18 = SheetConsolidadoBrokersYInmobiliarias.getRange("AQ" + FilaData).getDisplayValue();
  DataFinal.push(Dato1, Dato2, Dato3, Dato4, Dato5, Dato6, Dato7, Dato8, Dato9, Dato10, Dato11, Dato12, NumeroMatricula, Dato13, Dato14, Dato15, Dato16, Dato18);
  return { DataFinal: DataFinal }
}

function GenerarContratoFinalBrokersYInmobiliarias(deudores, contratoDatos, Ref) {
  var Dato1 = contratoDatos["Nombre Arrendador"];
  var Dato2 = contratoDatos["Documento Arrendador"];
  var Dato3 = contratoDatos["Nombre Arrendatario"];
  var Dato4 = contratoDatos["Documento Arrendatario"];
  var Dato5 = contratoDatos["Dirección Inmueble"];
  var Dato6 = contratoDatos["Ciudad Inmueble"];
  var Dato7 = contratoDatos["Ciudad Oficina de Registro"];
  var Dato8 = contratoDatos["Matricula Inmobiliaria"];
  var Dato9 = contratoDatos["Valor Canon Arrendamiento"];
  var Dato10 = contratoDatos["Valor Canon Arrendamiento(texto)"];
  var Dato11 = contratoDatos["Tipo Cuenta"];
  var Dato12 = contratoDatos["No. Cuenta"];
  var Dato13 = contratoDatos["Banco de la Cuenta"];
  var Dato14 = contratoDatos["Nombre del Titular de la Cuenta"];
  var Dato15 = contratoDatos["Vigencia de la poliza"];
  var Dato16 = contratoDatos["Día Inicio Póliza"];
  var Dato17 = contratoDatos["Mes Inicio Póliza"];
  var Dato18 = contratoDatos["Año Inicio Póliza"];
  var Dato19 = contratoDatos["Nro Solicitud Estudio"];
  var Dato20 = Utilities.formatDate(new Date(), "GMT-05:00", "yy");
  var FilaData = SheetConsolidadoBrokersYInmobiliarias.getRange("C:C").createTextFinder(Ref).ignoreDiacritics(true).matchEntireCell(true).findPrevious().getRow();
  var ValidarTipoPoliza = SheetConsolidadoBrokersYInmobiliarias.getRange("Y" + FilaData).getDisplayValue();
  var ValorAdmin = SheetConsolidadoBrokersYInmobiliarias.getRange("AB" + FilaData).getDisplayValue();
  if (ValidarTipoPoliza == "Vivienda") {
    if (ValorAdmin == "") {
      var IdFormatoContrato = "1rA7G86JBJQA-T60laqvHiMgH2oz0uy1-gwvzW4B8tWA";
    } else {
      var IdFormatoContrato = "16Qk_XtwfbUyZXk-ks0j8HqV_H8o_5Lq63gE0e03a0yI";
    }
  } else {
    if (ValorAdmin == "") {
      var IdFormatoContrato = "1D414uYf-bPpLdXSgrpbopJHZuT12aG7rfHFSfukOm9E";
    } else {
      var IdFormatoContrato = "1TK8PJMoFXAQeHCItNj6RM_29dzjtLt002XwBSN6xLCU";
    }
  }
  var idFolder = SheetConsolidadoBrokersYInmobiliarias.getRange("AY" + FilaData).getDisplayValue();
  idFolder = idFolder.split("/folders/")[1];
  var CopiaFormato = DriveApp.getFileById(IdFormatoContrato).makeCopy("Contrato-" + Ref, DriveApp.getFolderById(idFolder)).getId();
  var ContratoFinal = DocumentApp.openById(CopiaFormato);
  if (deudores.length < 1) {
    ContratoFinal.getBody().replaceText('“TITULO DEUDOR  SOLIDARIO FIRMA”', '');
    procesarDeudores(CopiaFormato, deudores);
    procesarSeccionFinalDeudores(CopiaFormato, deudores);
  } else {
    ContratoFinal.getBody().replaceText('“TITULO DEUDOR  SOLIDARIO FIRMA”', 'DEUDORES  SOLIDARIOS');
    procesarDeudores(CopiaFormato, deudores);
    procesarSeccionFinalDeudores(CopiaFormato, deudores);
  }
  ContratoFinal.getBody().replaceText('“Dato1”', Dato1)
  ContratoFinal.getBody().replaceText('“Dato2”', Dato2)
  ContratoFinal.getBody().replaceText('“Dato3”', Dato3)
  ContratoFinal.getBody().replaceText('“Dato4”', Dato4)
  ContratoFinal.getBody().replaceText('“Dato5”', Dato5)
  ContratoFinal.getBody().replaceText('“Dato6”', Dato6)
  ContratoFinal.getBody().replaceText('“Dato7”', Dato7)
  ContratoFinal.getBody().replaceText('“Dato8”', Dato8)
  ContratoFinal.getBody().replaceText('“Dato9”', Dato9);
  ContratoFinal.getBody().replaceText('“Dato10”', Dato10);
  ContratoFinal.getBody().replaceText('“Dato11”', Dato11);
  ContratoFinal.getBody().replaceText('“Dato12”', Dato12);
  ContratoFinal.getBody().replaceText('“Dato13”', Dato13);
  ContratoFinal.getBody().replaceText('“Dato14”', Dato14);
  ContratoFinal.getBody().replaceText('“Dato24”', Dato15);
  ContratoFinal.getBody().replaceText('“Dato15”', Dato16);
  ContratoFinal.getBody().replaceText('“Dato16”', Dato17);
  ContratoFinal.getBody().replaceText('“Dato17”', Dato18);
  ContratoFinal.getBody().replaceText('“Dato20”', Dato3);
  ContratoFinal.getBody().replaceText('“Dato21”', Dato3);
  ContratoFinal.getBody().replaceText('“Dato22”', Dato1);
  ContratoFinal.getBody().replaceText('“Dato23”', Dato1);
  ContratoFinal.getBody().replaceText('“Dato25”', Dato1);
  ContratoFinal.getBody().replaceText('“Dato26”', Dato2);
  ContratoFinal.getBody().replaceText('“Dato27”', Dato3);
  ContratoFinal.getBody().replaceText('“Dato28”', Dato4);
  ContratoFinal.getBody().replaceText('“Dato31”', Dato1);
  ContratoFinal.getBody().replaceText('“Dato32”', Dato3);
  ContratoFinal.getFooter().replaceText('“Dato33”', Dato19);
  ContratoFinal.getFooter().replaceText('“Dato34”', Dato20);
  var Fecha = formatearFechaEnEspanol();
  ContratoFinal.getBody().replaceText('“Dato35”', Fecha.dia);
  ContratoFinal.getBody().replaceText('“Dato36”', Fecha.mes);
  ContratoFinal.getBody().replaceText('“Dato37”', Fecha.year);
  if (ValorAdmin != "0" && ValorAdmin != "$0" && ValorAdmin != "") {
    var NumLetrasAdmin = NumeroALetras(ValorAdmin);
    ContratoFinal.getBody().replaceText('“DatoAdminPesos”', ValorAdmin);
    ContratoFinal.getBody().replaceText('“DatoAdminLetras”', NumLetrasAdmin);
  }
  if (ValidarTipoPoliza == "Comercio") {
    var Destino = SheetConsolidadoBrokersYInmobiliarias.getRange("BD" + FilaData).getDisplayValue();
    ContratoFinal.getBody().replaceText('“DESTINOCOMERCIO”', Destino);
  }
  var Servicios = SheetConsolidadoBrokersYInmobiliarias.getRange("AK" + FilaData).getDisplayValue();
  ContratoFinal.getBody().replaceText('“servicios”', Servicios);
  ContratoFinal.saveAndClose();
  return { IdContratoFinal: ContratoFinal.getId() }
}

function CargarPolizaBrokerYInmobiliaria(form) {
  var Ref = form["RefCargarPolizaModal2"];
  var FilaData = SheetConsolidadoBrokersYInmobiliarias.getRange("C:C").createTextFinder(Ref).ignoreDiacritics(true).matchEntireCell(true).findPrevious().getRow();
  var Folder = SheetConsolidadoBrokersYInmobiliarias.getRange("AY" + FilaData).getDisplayValue().split("/folders/")[1];
  var File1 = form["Dato20ContratoModal2"];
  var TypeFile1 = form["Dato20ContratoModal2"].name;
  var MimeTypeFile1 = TypeFile1.split(".")[1].toUpperCase();
  var NameFile1 = "Poliza-" + Ref;
  var resource = {
    title: NameFile1,
    mimeType: MimeTypeFile1,
    parents: [{ id: Folder }]
  };
  var FileCedula = Drive.Files.insert(resource, File1, {
    convert: false
  });
  var IdPoliza = FileCedula.id;
  return IdPoliza
}

function EnviarContratoFirmaBrokerYInmobiliaria(IdContrato, Ref, IdPoliza, TipoEnvio) {
  var FilaData = SheetConsolidadoBrokersYInmobiliarias.getRange("C:C").createTextFinder(Ref).matchEntireCell(true).ignoreDiacritics(true).findPrevious().getRow();
  var Correo = SheetConsolidadoBrokersYInmobiliarias.getRange("BH" + FilaData).getDisplayValue();
  Correo = Correo.replace(/\s/g, "");
  if (Correo.includes(";")) {
    Correo = Correo.replace(";", ",")
  }
  var correoAsesor = SheetConsolidadoBrokersYInmobiliarias.getRange("BF" + FilaData).getDisplayValue();
  var tipoCaso = SheetConsolidadoBrokersYInmobiliarias.getRange("B" + FilaData).getDisplayValue();
  var nombreUsuario = SheetConsolidadoBrokersYInmobiliarias.getRange("L" + FilaData).getDisplayValue();
  if (tipoCaso == "Brokers") {
    var correoResponsable = "wilber.barrera@segurosbolivar.com"
  } else {
    var correoResponsable = "magda.ramirez@segurosbolivar.com"
  }
  var correoCliente = SheetConsolidadoBrokersYInmobiliarias.getRange("N" + FilaData).getDisplayValue();
  var copiasCC = [];
  copiasCC.push(correoAsesor.trim())
  if (correoCliente && correoCliente.trim() !== '') {
    var correoClienteNormalizado = correoCliente.trim().replace(/ñ/g, 'n');
    copiasCC.push(correoClienteNormalizado);
  }
  if (TipoEnvio != "Solo Poliza") {
    var url = "https://www.googleapis.com/drive/v3/files/" + IdContrato + "/export?mimeType=application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    var opciones = {
      headers: {
        Authorization: "Bearer " + ScriptApp.getOAuthToken()
      }
    };
    var respuesta = UrlFetchApp.fetch(url, opciones);
    var blob = respuesta.getBlob().setName("Contrato_" + Ref + ".docx");
    Logger.log(IdPoliza)
    var Poliza = DriveApp.getFileById(IdPoliza);
    var Inventario = DriveApp.getFileById("1xy6OAB0QpCW86VD-qBJMn9Ku-NzxY6tR");
    var Clausulado = DriveApp.getFileById("1hX4qLiv20MBMWUVoj5UMyjvMdUge4-rt");
    var Files = [];
    Files.push(blob, Poliza, Inventario, Clausulado);
    var Pieza = HtmlService.createHtmlOutputFromFile("Pieza_Contrato_Correo").getContent();
    Pieza = Pieza.replace('[nombre usuario]', nombreUsuario)
    GmailApp.sendEmail(Correo, "Firma de Contrato Poliza Libertador " + Ref, "", { htmlBody: Pieza, attachments: Files, noReply: true, cc: copiasCC.join(','), bcc: correoResponsable });
    SheetConsolidadoBrokersYInmobiliarias.getRange("BG" + FilaData).setValue(new Date());
    SheetConsolidadoBrokersYInmobiliarias.getRange("BA" + FilaData).setValue("Expedido");
  } else {
    var Poliza = DriveApp.getFileById(IdPoliza);
    var Inventario = DriveApp.getFileById("1xy6OAB0QpCW86VD-qBJMn9Ku-NzxY6tR");
    var Clausulado = DriveApp.getFileById("1hX4qLiv20MBMWUVoj5UMyjvMdUge4-rt");
    var Files = [];
    Files.push(Poliza, Inventario, Clausulado);
    var Pieza = HtmlService.createHtmlOutputFromFile("Pieza_Solo_Poliza_Correo").getContent();
    Pieza = Pieza.replace('[nombre usuario]', nombreUsuario)
    GmailApp.sendEmail(Correo, "Poliza Libertador " + Ref, "", { htmlBody: Pieza, attachments: Files, noReply: true, cc: copiasCC.join(','), bcc: correoResponsable });
    SheetConsolidadoBrokersYInmobiliarias.getRange("BG" + FilaData).setValue(new Date());
    SheetConsolidadoBrokersYInmobiliarias.getRange("BA" + FilaData).setValue("Expedido");
  }
}

function EnviarCorreccionDocsBrokersInmobiliarias(referenciasDocs, observacionesFinales, Ref) {
  var FilaData = SheetConsolidadoBrokersYInmobiliarias.getRange("C:C").createTextFinder(Ref).matchEntireCell(true).ignoreDiacritics(true).findPrevious().getRow();
  var Correo = SheetConsolidadoBrokersYInmobiliarias.getRange("BH" + FilaData).getDisplayValue();
  Correo = Correo.replace(/\s/g, "");
  if (Correo.includes(";")) {
    Correo = Correo.replace(";", ",")
  }
  var observacionesText = "";
  var htmlObservaciones = "";
  for (var i = 0; i < referenciasDocs.length; i++) {
    var docReferencia = referenciasDocs[i];
    var observacion = observacionesFinales[i];
    if (observacion) {
      if (observacionesText !== "") {
        observacionesText += ',';
      }
      observacionesText += `"{${docReferencia},${observacion}}"`;
      htmlObservaciones += `<p><b>${docReferencia}:</b> ${observacion}</p>`;
    }
  }
  SheetConsolidadoBrokersYInmobiliarias.getRange("BB" + FilaData).setValue(observacionesText);
  var correoAsesor = SheetConsolidadoBrokersYInmobiliarias.getRange("BF" + FilaData).getDisplayValue();
  var tipoCaso = SheetConsolidadoBrokersYInmobiliarias.getRange("B" + FilaData).getDisplayValue();
  if (tipoCaso == "Brokers") {
    var correoResponsable = "wilber.barrera@segurosbolivar.com"
  } else {
    var correoResponsable = "magda.ramirez@segurosbolivar.com"
  }
  var Pieza = HtmlService.createHtmlOutputFromFile("Pieza_Correcion_Docs_Correo").getContent();
  Pieza = Pieza.replace('"Reemplazar"', htmlObservaciones);
  GmailApp.sendEmail(Correo, "Corrección de Documentos Poliza Libertador " + Ref, "", { htmlBody: Pieza, noReply: true, cc: correoAsesor, bcc: correoResponsable });
  SheetConsolidadoBrokersYInmobiliarias.getRange("BG" + FilaData).setValue(new Date());
  SheetConsolidadoBrokersYInmobiliarias.getRange("BA" + FilaData).setValue('Pendiente Corrección Documental');
}

function formatearFechaEnEspanol() {
  var meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
  var fecha = new Date();
  var dia = fecha.getDate();
  var mes = meses[fecha.getMonth()];
  var year = fecha.getFullYear();
  return { dia: dia, mes: mes, year: year };
}

function ocrPolizas() {
  var file = DriveApp.getFileById("1oyfkuO8YZtn8wJps2461erOTOOSqHOoF");
  var folderId = "1r33QZCu3RlumKQK3wj97W6lAYbHSzcGl";
  var resource = {
    title: file.getName(),
    parents: [{ id: folderId }]
  };
  var blob = file.getBlob();
  var docFile = Drive.Files.insert(resource, blob, {
    ocr: true,
    mimeType: MimeType.GOOGLE_DOCS
  });
  var doc = DocumentApp.openById(docFile.id);
  var text = doc.getBody().getText();
  var poliza = text.split("Póliza N°:")[1].trim().split("Certificado:")[0].trim();
  var valorNeto = text.split("TOTAL A PAGAR:")[1].trim().split("PERIODICIDAD")[0].trim();
  Logger.log(text);
  Logger.log(poliza);
  Logger.log(valorNeto)
  DriveApp.getFileById(docFile.id).setTrashed(true);
}


