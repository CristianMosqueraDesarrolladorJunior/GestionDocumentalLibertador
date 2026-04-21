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
/** Parametrización correos renovación (US-01): misma hoja WareHouse que Gestion / CRM Leads */
const SHEET_CORREO_RENOV_PARAM = "Correo_Renovacion_Param";
/** Columna I en hoja JSON renovaciones: historial de envíos de correo (US marca envío) */
const COL_LOG_CORREO_RENOV = 9;


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
          nuevoAsesor = getNewLeadAssignment(payload.segmento, payload.dataLead) || "sin.asignar@segurosbolivar.com";
          sheet.getRange(rowIndex, 3).setValue(nuevoAsesor);
          break;
      }
      sheet.getRange(rowIndex, 5).setValue(nuevoEstado);

      let jsonCell = sheet.getRange(rowIndex, 6);
      let hayManual = false;
      if (payload.manualUpdates && typeof payload.manualUpdates === "object") {
        for (const mk in payload.manualUpdates) {
          if (Object.prototype.hasOwnProperty.call(payload.manualUpdates, mk)) {
            hayManual = true;
            break;
          }
        }
      }
      const hayDeltaJson = Object.keys(urlsFinales).length > 0 || hayManual;
      const correccionSoloTipificacion = payload.decision === "CORRECTION" && !hayDeltaJson;

      if (!correccionSoloTipificacion) {
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
      }
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

/** Texto para comparar perfiles en "Tabla Asignación Usuarios" (acentos / mayúsculas). */
function normalizarTextoAsignacion_(s) {
  return String(s || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

/** Filas A:D de Tabla Asignación Usuarios (correo, estado, perfil, webhook). */
function leerTablaAsignacionUsuarios_() {
  try {
    var lr = SheetAssignment.getLastRow();
    if (lr < 2) return [];
    var vals = SheetAssignment.getRange(2, 1, lr, 4).getDisplayValues();
    var out = [];
    var reMail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    for (var i = 0; i < vals.length; i++) {
      var correo = String(vals[i][0] || "").trim().replace(/\s/g, "");
      if (!correo || !reMail.test(correo)) continue;
      out.push({
        correo: correo,
        estado: String(vals[i][1] || "").trim(),
        perfil: String(vals[i][2] || "").trim(),
        webhook: String(vals[i][3] || "").trim()
      });
    }
    return out;
  } catch (e) {
    Logger.log("leerTablaAsignacionUsuarios_: " + e);
    return [];
  }
}

/**
 * Metadatos para el cliente (preview correo renovación sin llamadas extra).
 * Destinatarios según hoja de asignación; colas de corrección según Gestion.
 */
function getMetaCargaCorreoRenovacionCliente_() {
  return {
    usuariosAsignacion: leerTablaAsignacionUsuarios_(),
    agentesRenovations: listarAgentesDisponiblesColaGestion_("Renovations"),
    agentesCorreccionesBI: listarAgentesDisponiblesColaGestion_("CorreccionesBI")
  };
}

function GetDataUser() {
  var UserMail = Session.getActiveUser().getEmail();
  let Status;
  
  // Nota: Asegúrate de que 'SheetConsolidado', 'SheetAssignment' y 'Renovaciones' 
  // estén definidas globalmente o inicialízalas aquí.
  let DataRange = SheetConsolidado.getRange("A2:BC" + SheetConsolidado.getLastRow()).getDisplayValues();

  let Rows = SheetAssignment.getRange("A:A")
    .createTextFinder(UserMail)
    .matchEntireCell(true)
    .ignoreDiacritics(true)
    .findAll();

  if (Rows == null || Rows.length === 0) return ["No Autenticado", "Null"];

  Status = "Autenticado";

  // Recolectar TODOS los roles del usuario
  let Roles = Rows.map(row => SheetAssignment.getRange("C" + row.getRow()).getDisplayValue());
  console.log("Roles encontrados:", Roles);

  let resultado = {};

  // ── SECCIÓN GESTIÓN DOCUMENTAL (UNIFICADA) ──────────────────
  if (Roles.includes("Gestión Documental")) {
    let leadsInternos = [];
    
    // 1. Filtramos los leads internos de la hoja Consolidado
    DataRange.filter(function(row) {
      let tieneEstado = row.indexOf("Pendiente Validación Documental") > -1 ||
                        row.indexOf("Pendiente Corrección Documental") > -1;
      let esDelUsuario = row[54].indexOf(UserMail) > -1;
      
      if (tieneEstado && esDelUsuario) {
        leadsInternos.push([
          row[0], row[4], row[8], row[24], row[2], "Gestión",
          row[11], row[12], row[1], row[3], row[5], row[9],
          row[10], row[13], row[14], row[15], row[17], row[19],
          row[20], row[21], row[22], row[23], row[16], row[45],
          row[7], row[37], "gestiondocumental"
        ]);
      }
    });

    // 2. Obtenemos los leads de Brokers llamando a la otra función
    let leadsBrokers = GetDataBrokersYInmobiliarias();

    // 3. Unimos ambos arrays en la misma propiedad del objeto resultado
    // .concat() junta los dos grupos en una sola lista larga
    resultado.gestionDocumental = leadsInternos.concat(leadsBrokers);
    
    console.log("Total leads (Internos + Brokers):", resultado.gestionDocumental.length);
  }

  // ── SECCIÓN RENOVACIONES ────────────────────────────────────
  if (Roles.includes("Analista Renovaciones")) {
    let dataUpd = Renovaciones.getRange("A2:H" + Renovaciones.getLastRow()).getDisplayValues();

    let misRegistros = dataUpd.filter(row =>
      row[2] && row[2].toString().trim().toLowerCase() === UserMail.trim().toLowerCase()
    );

    let todosMisLeads = misRegistros.map(row => {
      const fechaIngreso = row[0];
      const registroRaw = row[1];
      const nombreAgente = row[2];
      const etapaFunel = row[3];
      const estadoGestion = row[4];
      const dataGestionRaw = row[5];
      const historiaRaw = row[6];
      const emailNotificacionGestionAt = row[7] != null ? String(row[7]).trim() : "";

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
          historialGestion = [];
        }
      }

      let leadSelect = {};
      try {
        leadSelect = JSON.parse(String(dataGestionRaw).trim().replace(/\bNaN\b/g, "null"));
      } catch (e) {
        leadSelect = { error: "JSON Gestión inválido" };
      }

      let dataLead = {};
      try {
        dataLead = JSON.parse(String(registroRaw).trim().replace(/\bNaN\b/g, "null"));
      } catch (e) {
        dataLead = { error: "JSON Lead inválido" };
      }

      return {
        fechaIngreso, leadSelect, dataLead,
        nombreAgente, etapaFunel, estadoGestion,
        historialGestiones: historialGestion,
        emailNotificacionGestionAt
      };
    });

    resultado.renovaciones = {
      pendientes: todosMisLeads.filter(i => ["Enviar a Expedicion", "Autogestionado", "Caso Revisado"].includes(i.estadoGestion)),
      especiales: todosMisLeads.filter(i => i.estadoGestion === "Caso Especial"),
      renovadas: todosMisLeads.filter(i => ["Poliza Renovada", "Expedido"].includes(i.estadoGestion))
    };
  }

  // Retornar Status, el objeto con todos los leads, el mail y los roles
  return [Status, resultado, UserMail, Roles];
}

/**
 * Cola en hoja Gestion (WareHouse): CorreccionesBI vs Renovations.
 * Misma regla que segmento normalizado del flujo de renovaciones.
 */
function colaGestionCorreccionPorSegmentoNorm_(segmentoNorm) {
  var u = String(segmentoNorm || "").toUpperCase().trim();
  if (u === "BROKER" || u === "INMOBILIARIA") return "CorreccionesBI";
  return "Renovations";
}

/**
 * @param {string} [segmento] prioridad sobre dataLead.segmento para tipificar la cola
 * @param {object} [dataLead] respaldo si segmento viene vacío (misma lógica que normalizeSegmentoRenovacion_)
 */
function getNewLeadAssignment(segmento, dataLead) {
  var dl = dataLead && typeof dataLead === "object" ? Object.assign({}, dataLead) : {};
  if (segmento != null && String(segmento).trim() !== "") {
    dl.segmento = segmento;
  }
  var segNorm = normalizeSegmentoRenovacion_(dl, null);
  var gestion = colaGestionCorreccionPorSegmentoNorm_(segNorm);
  var asignacion = AssignLead(gestion);
  return asignacion && asignacion.email ? String(asignacion.email).trim() : null;
}

function getUserDataFilteredForDeal_(deal) {
  var userDataleads = DataGestion.getRange("A1:K" + DataGestion.getLastRow()).getDisplayValues();
  var userData = [];
  if (deal === "Renovations") {
    userData = userDataleads.filter(function (row) {
      return row[3] === "Renovations";
    });
  } else if (deal === "Sales") {
    userData = userDataleads.filter(function (row) {
      return row[3] === "Seguro de Vida" || row[3] === "Seguro de Desempleo";
    });
  } else if (deal === "CorreccionesBI") {
    userData = userDataleads.filter(function (row) {
      return row[3] === "CorreccionesBI";
    });
  }
  return userData;
}

/** Asesores con cupo en la cola (misma regla que AssignLead). Útil para mostrar en UI. */
function listarAgentesDisponiblesColaGestion_(deal) {
  var userData = getUserDataFilteredForDeal_(deal);
  var out = [];
  for (var i = 0; i < userData.length; i++) {
    var novelty = userData[i][4];
    var totalCapacity = Number(userData[i][5]);
    var totalInProcess = Number(userData[i][6]);
    var availability = totalCapacity - totalInProcess;
    if ((!novelty || String(novelty).trim() === "") && availability > 0) {
      out.push({
        nombre: String(userData[i][1] || ""),
        email: String(userData[i][2] || "").trim()
      });
    }
  }
  return out;
}

function AssignLead(deal = "CorreccionesBI") {
  let userData = getUserDataFilteredForDeal_(deal);
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
      DataRangeUserPending.push([DataRange[0], DataRange[11], DataRange[19], DataRange[12], DataRange[1], "", DataRange[2], DataRange[24], DataRange[28], DataRange[45], DataRange[10], DataRange[8], DataRange[9], DataRange[13], DataRange[18], DataRange[33], DataRange[26], DataRange[27], DataRange[28], DataRange[25], DataRange[14], DataRange[48], DataRange[49], DataRange[51], DataRange[17], DataRange[56], DataRange[30], DataRange[31], "broker-inmobiliaria"]);
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


/* =============================================================================
   Renovaciones — Correo por segmento (US-01, US-02, US-03)
   Sin segmento en datos = se trata como PROPIETARIO.
   ============================================================================= */

function _escapeHtmlRenovCorreo_(s) {
  if (s == null || s === undefined) return "";
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function getOrCreateCorreoRenovParamSheet_() {
  var sh = DataWereHouse.getSheetByName(SHEET_CORREO_RENOV_PARAM);
  if (!sh) {
    sh = DataWereHouse.insertSheet(SHEET_CORREO_RENOV_PARAM);
    sh.getRange(1, 1, 1, 4).setValues([["Clave", "Nombre", "Correo", "Especialidad"]]);
  }
  return sh;
}

/**
 * Lista filas para la vista CRM de administración (US-01).
 */
function listarParamCorreoRenovacion() {
  try {
    getOrCreateCorreoRenovParamSheet_();
    var sh = DataWereHouse.getSheetByName(SHEET_CORREO_RENOV_PARAM);
    var out = [];
    if (sh.getLastRow() >= 2) {
      var rows = sh.getRange(2, 1, sh.getLastRow(), 4).getDisplayValues();
      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i][0] || "").trim() === "") continue;
        out.push({
          clave: rows[i][0],
          nombre: rows[i][1],
          correo: rows[i][2],
          especialidad: rows[i][3]
        });
      }
    }
    return { success: true, filas: out };
  } catch (e) {
    return { success: false, message: String(e), filas: [] };
  }
}

/**
 * Guarda la tabla de parámetros (desde la vista CRM). Espera array de { clave, nombre, correo, especialidad }.
 */
function guardarParamCorreoRenovacion(filas) {
  try {
    var sh = getOrCreateCorreoRenovParamSheet_();
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow(), 4).clearContent();
    }
    if (!filas || !filas.length) {
      return { success: true, message: "Sin filas CRM; los correos de renovación usan Tabla Asignación Usuarios." };
    }
    var matrix = [];
    for (var i = 0; i < filas.length; i++) {
      var f = filas[i] || {};
      matrix.push([
        String(f.clave || "").trim(),
        String(f.nombre || "").trim(),
        String(f.correo || "").trim().replace(/\s/g, ""),
        String(f.especialidad || "").trim()
      ]);
    }
    sh.getRange(2, 1, 1 + matrix.length, 4).setValues(matrix);
    return { success: true, message: "Parámetros guardados." };
  } catch (e) {
    return { success: false, message: String(e) };
  }
}

function normalizeSegmentoRenovacion_(dataLead, segmentoSheetColumn) {
  var raw = "";
  if (dataLead && dataLead.segmento) raw = String(dataLead.segmento);
  else if (segmentoSheetColumn) raw = String(segmentoSheetColumn);
  var u = raw.toUpperCase().trim();
  if (!u || u.indexOf("SIN") === 0) return "PROPIETARIO";
  if (u.indexOf("INMOBILIARIA") !== -1) return "INMOBILIARIA";
  if (u.indexOf("BROKER") !== -1 || u.indexOf("BRÓKER") !== -1 || u.indexOf("BROK") !== -1) return "BROKER";
  if (u.indexOf("PROPIETARIO") !== -1) return "PROPIETARIO";
  return "PROPIETARIO";
}

function _validEmail_(e) {
  if (!e || typeof e !== "string") return false;
  var t = e.trim();
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(t);
}

function _emailPropietario_(dataLead) {
  var a = (dataLead && dataLead.email) ? String(dataLead.email).trim() : "";
  var b = (dataLead && dataLead.emailBrIn) ? String(dataLead.emailBrIn).trim() : "";
  if (_validEmail_(a)) return a.replace(/\s/g, "");
  if (_validEmail_(b)) return b.replace(/\s/g, "");
  return "";
}

function filasAsignacionActivas_(usuarios) {
  var out = [];
  for (var i = 0; i < usuarios.length; i++) {
    if (normalizarTextoAsignacion_(usuarios[i].estado) === "activo") out.push(usuarios[i]);
  }
  return out;
}

function filtrarPorPerfilNorm_(activos, testFn) {
  var r = [];
  for (var i = 0; i < activos.length; i++) {
    if (testFn(normalizarTextoAsignacion_(activos[i].perfil))) r.push(activos[i]);
  }
  return r;
}

function elegirPorCoincidenciaAnalista_(candidatos, emailAnalistaAsignado) {
  if (!candidatos || !candidatos.length) return null;
  var e = String(emailAnalistaAsignado || "").toLowerCase().trim().replace(/\s/g, "");
  if (!e) return candidatos[0];
  var j;
  for (j = 0; j < candidatos.length; j++) {
    if (candidatos[j].correo.toLowerCase() === e) return candidatos[j];
  }
  var local = e.indexOf("@") > 0 ? e.split("@")[0] : "";
  for (j = 0; j < candidatos.length; j++) {
    var c = candidatos[j].correo.toLowerCase();
    if (local && c.split("@")[0] === local) return candidatos[j];
  }
  return candidatos[0];
}

/**
 * Arma Para + CC según segmento leyendo perfiles activos en Tabla Asignación Usuarios.
 */
function buildRenovacionRecipientPlan_(segmentoNorm, dataLead, emailAnalistaAsignado) {
  var activos = filasAsignacionActivas_(leerTablaAsignacionUsuarios_());
  var paraRows = filtrarPorPerfilNorm_(activos, function (np) {
    return np.indexOf("cuenta comercial") >= 0 || (np.indexOf("notificacion") >= 0 && np.indexOf("comercial") >= 0);
  });
  var para = paraRows.length
    ? String(paraRows[0].correo).trim().replace(/\s/g, "")
    : "renovacionesarrendamiento@segurosbolivar.com";
  var cc = [];
  var ccMeta = [];

  if (segmentoNorm === "PROPIETARIO") {
    var ep = _emailPropietario_(dataLead);
    if (ep) {
      cc.push(ep);
      ccMeta.push({ nombre: "Propietario / tomador", correo: ep, rol: "CC segmento Propietario" });
    }
  } else if (segmentoNorm === "BROKER") {
    var segBrokers = filtrarPorPerfilNorm_(activos, function (np) {
      return np.indexOf("segmento") >= 0 && np.indexOf("broker") >= 0;
    });
    var kb;
    for (kb = 0; kb < segBrokers.length; kb++) {
      var emW = String(segBrokers[kb].correo).trim().replace(/\s/g, "");
      if (_validEmail_(emW) && cc.indexOf(emW) === -1) {
        cc.push(emW);
        ccMeta.push({ nombre: segBrokers[kb].perfil || "Notificación segmento Bróker", correo: emW, rol: "CC segmento Bróker" });
      }
    }
    var ejRows = filtrarPorPerfilNorm_(activos, function (np) {
      if (np.indexOf("inmobiliaria") >= 0) return false;
      return np.indexOf("ejecutivo") >= 0 && np.indexOf("broker") >= 0;
    });
    var ejPick = elegirPorCoincidenciaAnalista_(ejRows, emailAnalistaAsignado);
    if (ejPick && _validEmail_(ejPick.correo)) {
      var emE = String(ejPick.correo).trim().replace(/\s/g, "");
      if (cc.indexOf(emE) === -1) {
        cc.push(emE);
        ccMeta.push({ nombre: ejPick.perfil || "Ejecutivo de cuenta", correo: emE, rol: "CC ejecutivo cuenta Bróker" });
      }
    }
  } else if (segmentoNorm === "INMOBILIARIA") {
    var segInm = filtrarPorPerfilNorm_(activos, function (np) {
      return np.indexOf("segmento") >= 0 && np.indexOf("inmobiliaria") >= 0;
    });
    var ki;
    for (ki = 0; ki < segInm.length; ki++) {
      var emM = String(segInm[ki].correo).trim().replace(/\s/g, "");
      if (_validEmail_(emM) && cc.indexOf(emM) === -1) {
        cc.push(emM);
        ccMeta.push({ nombre: segInm[ki].perfil || "Notificación segmento Inmobiliaria", correo: emM, rol: "CC segmento Inmobiliaria" });
      }
    }
    var ejInm = filtrarPorPerfilNorm_(activos, function (np) {
      return np.indexOf("ejecutivo") >= 0 && np.indexOf("inmobiliaria") >= 0;
    });
    var ejInmPick = elegirPorCoincidenciaAnalista_(ejInm, emailAnalistaAsignado);
    if (ejInmPick && _validEmail_(ejInmPick.correo)) {
      var emEI = String(ejInmPick.correo).trim().replace(/\s/g, "");
      if (cc.indexOf(emEI) === -1) {
        cc.push(emEI);
        ccMeta.push({ nombre: ejInmPick.perfil || "Ejecutivo de cuenta", correo: emEI, rol: "CC ejecutivo cuenta Inmobiliaria" });
      }
    }
  }

  return { para: para, cc: cc, ccMeta: ccMeta, segmentoNorm: segmentoNorm };
}

function _parseExtraCcList_(s1, s2) {
  var out = [];
  var parts = [];
  if (s1) parts = parts.concat(String(s1).split(/[;,]/));
  if (s2) parts = parts.concat(String(s2).split(/[;,]/));
  for (var i = 0; i < parts.length; i++) {
    var x = String(parts[i] || "").trim();
    if (_validEmail_(x) && out.indexOf(x) === -1) out.push(x.replace(/\s/g, ""));
  }
  return out;
}

function _tituloCampoCorreoRenov_(key) {
  var k = String(key).toLowerCase();
  if (k === "email") return "Tomador / asegurado (correo principal)";
  if (k.indexOf("emailbrin") !== -1) return "Arrendatario / contacto BrIn";
  if (k.indexOf("correobroker") !== -1 || k.indexOf("emailbroker") !== -1 || k.indexOf("correo_broker") !== -1 || k.indexOf("email_broker") !== -1) return "Contacto bróker";
  if (k.indexOf("inmobiliaria") !== -1 && (k.indexOf("correo") !== -1 || k.indexOf("email") !== -1)) return "Contacto inmobiliaria";
  if (k.indexOf("ejecutivo") !== -1 && (k.indexOf("correo") !== -1 || k.indexOf("email") !== -1)) return "Ejecutivo comercial";
  if (k.indexOf("correo") !== -1 || k.indexOf("email") !== -1 || k.indexOf("mail") !== -1) return "Contacto en datos (" + key + ")";
  return key;
}

function _categoriaCampoCorreoRenov_(key) {
  var k = String(key).toLowerCase();
  if (k === "email") return "Propietario / tomador";
  if (k.indexOf("brin") !== -1) return "Arrendatario";
  if (k.indexOf("broker") !== -1 || k.indexOf("corredor") !== -1) return "Bróker";
  if (k.indexOf("inmobiliaria") !== -1) return "Inmobiliaria";
  if (k.indexOf("ejecutivo") !== -1) return "Ejecutivo";
  return "Dato póliza";
}

function _extraerCorreosDesdeObjetoPlano_(obj, fuente, seen, candidatos) {
  if (!obj || typeof obj !== "object") return;
  for (var key in obj) {
    if (!obj.hasOwnProperty(key)) continue;
    var val = obj[key];
    if (val == null) continue;
    if (typeof val === "string" && _validEmail_(val)) {
      var lk = String(key).toLowerCase();
      if (lk.indexOf("email") === -1 && lk.indexOf("correo") === -1 && lk.indexOf("mail") === -1) continue;
      var em = String(val).trim().replace(/\s/g, "");
      var low = em.toLowerCase();
      if (seen[low]) continue;
      seen[low] = true;
      candidatos.push({
        id: fuente + "_" + key + "_" + low.replace(/[^a-z0-9]/gi, "").substring(0, 14),
        correo: em,
        titulo: _tituloCampoCorreoRenov_(key),
        categoria: _categoriaCampoCorreoRenov_(key),
        campo: key,
        fuenteDatos: fuente
      });
    }
  }
}

function construirCandidatosCorreoRenovacion_(dataLead, leadSelect, plan) {
  var seen = {};
  var candidatos = [];
  _extraerCorreosDesdeObjetoPlano_(dataLead, "Póliza (JSON fila)", seen, candidatos);
  _extraerCorreosDesdeObjetoPlano_(leadSelect || {}, "Gestión / oferta", seen, candidatos);

  for (var i = 0; i < plan.ccMeta.length; i++) {
    var m = plan.ccMeta[i];
    var em = String(m.correo || "").trim().replace(/\s/g, "");
    if (!_validEmail_(em)) continue;
    var low = em.toLowerCase();
    if (seen[low]) {
      for (var j = 0; j < candidatos.length; j++) {
        if (candidatos[j].correo.toLowerCase() === low) {
          if (m.rol) candidatos[j].titulo += " · " + m.rol;
          break;
        }
      }
      continue;
    }
    seen[low] = true;
    candidatos.push({
      id: "regla_" + i + "_" + low.replace(/[^a-z0-9]/gi, "").substring(0, 12),
      correo: em,
      titulo: m.nombre || "Destinatario por regla de segmento",
      categoria: "Regla / segmento",
      rolDetalle: m.rol || "",
      fuenteDatos: "parametrizado"
    });
  }

  var planEmails = {};
  for (var p = 0; p < plan.cc.length; p++) {
    planEmails[String(plan.cc[p]).toLowerCase()] = true;
  }
  for (var c = 0; c < candidatos.length; c++) {
    candidatos[c].incluidoPorDefecto = !!planEmails[candidatos[c].correo.toLowerCase()];
  }

  return candidatos;
}

function _aplicarSeleccionCandidatos_(candidatos, ccSeleccionados, usarDefecto) {
  for (var i = 0; i < candidatos.length; i++) {
    var em = candidatos[i].correo.toLowerCase();
    if (usarDefecto) {
      candidatos[i].seleccionado = !!candidatos[i].incluidoPorDefecto;
    } else {
      candidatos[i].seleccionado = false;
      for (var s = 0; s < ccSeleccionados.length; s++) {
        if (String(ccSeleccionados[s] || "")
          .trim()
          .toLowerCase() === em) {
          candidatos[i].seleccionado = true;
          break;
        }
      }
    }
  }
}

function _ccFinalDesdeCandidatosExtras_(candidatos, extras, paraEmail) {
  var cc = [];
  var paraLow = String(paraEmail || "").toLowerCase().trim();
  for (var i = 0; i < candidatos.length; i++) {
    if (!candidatos[i].seleccionado) continue;
    var em = candidatos[i].correo;
    if (em.toLowerCase() === paraLow) continue;
    if (cc.indexOf(em) === -1) cc.push(em);
  }
  for (var e = 0; e < extras.length; e++) {
    var x = extras[e];
    if (!x || String(x).toLowerCase() === paraLow) continue;
    if (cc.indexOf(x) === -1) cc.push(x);
  }
  return cc;
}

function _ccMetaDesdeCandidatos_(candidatos, extras) {
  var meta = [];
  for (var i = 0; i < candidatos.length; i++) {
    if (!candidatos[i].seleccionado) continue;
    meta.push({
      nombre: candidatos[i].titulo,
      correo: candidatos[i].correo,
      rol: candidatos[i].categoria + " · " + (candidatos[i].fuenteDatos || "")
    });
  }
  for (var j = 0; j < extras.length; j++) {
    meta.push({
      nombre: "Adicional (formulario)",
      correo: extras[j],
      rol: "CC sugerido"
    });
  }
  return meta;
}

function _fillRenovCorreoTemplate_(fileName, vars) {
  var html = HtmlService.createHtmlOutputFromFile(fileName).getContent();
  for (var k in vars) {
    if (!vars.hasOwnProperty(k)) continue;
    var re = new RegExp("\\{\\{" + k + "\\}\\}", "g");
    html = html.replace(re, vars[k]);
  }
  return html;
}

function _subjectRenovacion_(tipoAccion, poliza) {
  var ref = poliza || "N/A";
  if (tipoAccion === "APPROVE") return "Renovación aprobada / expedición — Póliza " + ref;
  if (tipoAccion === "CORRECTION") return "Correcciones solicitadas — Renovación póliza " + ref;
  if (tipoAccion === "CANCELADA") return "Renovación declinada (imposible renovar) — Póliza " + ref;
  return "Notificación renovación — Póliza " + ref;
}

/**
 * Vista previa + metadatos para el modal (US-02, US-03). No envía correo.
 * payload: { tipoAccionCorreo, dataLead, leadSelect?, segmentoSheetColumn?, segmento? (prioridad para normalizar segmento / cola CorreccionesBI), emailAnalistaAsignado?, observations?, notasAnalista?, correoAdicional1?, correoAdicional2?, ccSeleccionados? }
 * Si ccSeleccionados es undefined/null → se usan los marcados por defecto (reglas de segmento). Si es array (puede estar vacío) → refleja la selección del analista.
 * Solo uso interno (envío / prueba en seco); la vista previa en webapp usa datos de GetDataUser.
 */
function buildRenovacionCorreoPreviewPayload_(payload) {
  try {
    var dataLead = payload.dataLead || {};
    var leadSelect = payload.leadSelect || {};
    var dataLeadSeg = dataLead;
    if (payload.segmento != null && String(payload.segmento).trim() !== "") {
      dataLeadSeg = Object.assign({}, dataLead, { segmento: payload.segmento });
    }
    var seg = normalizeSegmentoRenovacion_(dataLeadSeg, payload.segmentoSheetColumn);
    var plan = buildRenovacionRecipientPlan_(seg, dataLead, payload.emailAnalistaAsignado);
    var extra = _parseExtraCcList_(payload.correoAdicional1, payload.correoAdicional2);

    var candidatos = construirCandidatosCorreoRenovacion_(dataLead, leadSelect, plan);
    var rawSel = payload.ccSeleccionados;
    var usarDefecto = (rawSel === undefined || rawSel === null);
    var ccSelArr = usarDefecto ? [] : rawSel;
    _aplicarSeleccionCandidatos_(candidatos, ccSelArr, usarDefecto);

    var ccFinal = _ccFinalDesdeCandidatosExtras_(candidatos, extra, plan.para);
    var ccMeta = _ccMetaDesdeCandidatos_(candidatos, extra);

    var poliza = String(dataLead.poliza != null ? dataLead.poliza : "");
    var obsBase = payload.observations || payload.observaciones || "";
    var obs = payload.tipoAccionCorreo === "APPROVE"
      ? (payload.notasAnalista || obsBase || "")
      : obsBase;
    var fechaStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");

    var vars = {
      POLIZA: _escapeHtmlRenovCorreo_(poliza),
      ASEGURADO: _escapeHtmlRenovCorreo_(dataLead.asegurado || ""),
      DOCUMENTO: _escapeHtmlRenovCorreo_(dataLead.documento != null ? dataLead.documento : ""),
      SEGMENTO: _escapeHtmlRenovCorreo_(seg),
      NOTAS: _escapeHtmlRenovCorreo_(payload.notasAnalista || obsBase || "—"),
      OBSERVACIONES: _escapeHtmlRenovCorreo_(obs || "—"),
      FECHA_ENVIO: _escapeHtmlRenovCorreo_(fechaStr)
    };

    var fileTpl = "Pieza_Renov_Correo_Aprobada";
    if (payload.tipoAccionCorreo === "CORRECTION") fileTpl = "Pieza_Renov_Correo_Correccion";
    else if (payload.tipoAccionCorreo === "CANCELADA") fileTpl = "Pieza_Renov_Correo_Cancelada";

    var htmlBody = _fillRenovCorreoTemplate_(fileTpl, vars);
    var subject = _subjectRenovacion_(payload.tipoAccionCorreo, poliza);

    var asignacionCorreccion = null;
    if (payload.tipoAccionCorreo === "CORRECTION") {
      var colaCorr = colaGestionCorreccionPorSegmentoNorm_(seg);
      var agentePrev = AssignLead(colaCorr);
      asignacionCorreccion = {
        cola: colaCorr,
        colaLegible:
          colaCorr === "CorreccionesBI"
            ? "Correcciones BI (bróker / inmobiliaria) — Tabla Gestion"
            : "Renovaciones (propietario y demás) — Tabla Gestion",
        agentePrevisto: agentePrev
          ? { nombre: String(agentePrev.name || ""), email: String(agentePrev.email || "").trim() }
          : null,
        emailEnHoja:
          agentePrev && agentePrev.email
            ? String(agentePrev.email).trim()
            : "sin.asignar@segurosbolivar.com",
        agentesDisponibles: listarAgentesDisponiblesColaGestion_(colaCorr)
      };
    }

    return {
      success: true,
      subject: subject,
      para: plan.para,
      cc: ccFinal,
      ccMeta: ccMeta,
      candidatos: candidatos,
      segmentoNorm: seg,
      htmlBody: htmlBody,
      asignacionCorreccion: asignacionCorreccion
    };
  } catch (e) {
    return { success: false, message: String(e) };
  }
}

function _findRenovacionRowByPoliza_(poliza) {
  var sheet = Renovaciones;
  var data = sheet.getDataRange().getDisplayValues();
  var p = String(poliza);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).indexOf(p) !== -1) return i + 1;
  }
  return -1;
}

function appendLogCorreoRenovacion_(poliza, entry) {
  var row = _findRenovacionRowByPoliza_(poliza);
  if (row < 0) return false;
  var cell = Renovaciones.getRange(row, COL_LOG_CORREO_RENOV);
  var prev = cell.getValue();
  var arr = [];
  try {
    if (prev) arr = JSON.parse(prev);
  } catch (e) {
    arr = [];
  }
  if (!Array.isArray(arr)) arr = [];
  arr.push(entry);
  cell.setValue(JSON.stringify(arr));
  return true;
}

function _sanitizeNombreAdjuntoRenov_(nombre) {
  var n = String(nombre || "adjunto").trim();
  if (n.length > 180) n = n.substring(0, 180);
  return n.replace(/[^\w\.\- áéíóúñÁÉÍÓÚÑ]/g, "_");
}

/**
 * Procesa adjuntos una sola vez: blobs para Gmail + informe por ítem (reutilizado en envío real y en prueba en seco).
 * @returns {{ blobs: GoogleAppsScript.Base.Blob[], informe: Array<{tipo:string,nombre:string,ok:boolean,error?:string,sizeBytes?:number}> }}
 */
function _adjuntosRenovCorreoProcesar_(adjuntos) {
  var blobs = [];
  var informe = [];
  if (!adjuntos || !adjuntos.length) {
    return { blobs: blobs, informe: informe };
  }
  for (var i = 0; i < adjuntos.length; i++) {
    var a = adjuntos[i] || {};
    var nomRef = String(a.nombre || a.fileId || "archivo");
    try {
      if (a.tipo === "drive" && a.fileId) {
        var file = DriveApp.getFileById(String(a.fileId).trim());
        var b = file.getBlob();
        var nom = _sanitizeNombreAdjuntoRenov_(a.nombre || file.getName());
        b.setName(nom);
        blobs.push(b);
        informe.push({ tipo: "drive", nombre: nom, ok: true, sizeBytes: b.getBytes().length });
      } else if (a.tipo === "base64" && a.data) {
        var mime = a.mimeType || "application/octet-stream";
        var nom2 = _sanitizeNombreAdjuntoRenov_(a.nombre || "documento");
        var blob = Utilities.newBlob(Utilities.base64Decode(String(a.data)), mime, nom2);
        blobs.push(blob);
        informe.push({ tipo: "base64", nombre: nom2, ok: true, sizeBytes: blob.getBytes().length });
      } else {
        informe.push({ tipo: String(a.tipo || "?"), nombre: nomRef, ok: false, error: "Formato no reconocido o datos incompletos" });
      }
    } catch (err) {
      Logger.log("Adjunto correo renovación: " + err);
      informe.push({ tipo: String(a.tipo || "?"), nombre: nomRef, ok: false, error: String(err) });
    }
  }
  return { blobs: blobs, informe: informe };
}

function _adjuntosRenovCorreoABlobs_(adjuntos) {
  return _adjuntosRenovCorreoProcesar_(adjuntos).blobs;
}

/**
 * Prueba completa sin Gmail ni processAnalystDecision ni log en hoja.
 * Útil para validar desde el CRM (botón) o desde el editor: ejecutarPruebasAutomaticasCorreoRenovacion()
 */
function probarCorreoRenovacionSinEnvio(payload) {
  try {
    var tipo = payload.tipoAccionCorreo || payload.decision;
    if (tipo === "APPROVE") tipo = "APPROVE";
    else if (tipo === "CORRECTION") tipo = "CORRECTION";
    else if (tipo === "CANCELADA" || tipo === "CANCEL") tipo = "CANCELADA";

    if (tipo !== "APPROVE" && tipo !== "CORRECTION" && tipo !== "CANCELADA") {
      return { success: false, dryRun: true, message: "Tipo de acción de correo no válido." };
    }

    var dataLead = payload.dataLead || {};
    if ((tipo === "CORRECTION" || tipo === "CANCELADA") && !(payload.observations && String(payload.observations).trim())) {
      return { success: false, dryRun: true, message: "Debe indicar observaciones / motivo (prueba igual que envío real)." };
    }

    var prevIn = {
      tipoAccionCorreo: tipo,
      dataLead: dataLead,
      leadSelect: payload.leadSelect,
      segmentoSheetColumn: payload.segmentoSheetColumn,
      emailAnalistaAsignado: payload.emailAnalistaAsignado,
      observations: payload.observations,
      notasAnalista: payload.notasAnalista,
      correoAdicional1: payload.correoAdicional1,
      correoAdicional2: payload.correoAdicional2,
      ccSeleccionados: payload.ccSeleccionados
    };
    if (payload.segmento != null && String(payload.segmento).trim() !== "") {
      prevIn.segmento = payload.segmento;
    }
    var preview = buildRenovacionCorreoPreviewPayload_(prevIn);

    if (!preview.success) {
      return Object.assign({ dryRun: true }, preview);
    }

    var adjRes = _adjuntosRenovCorreoProcesar_(payload.emailAdjuntos || []);
    var fallos = adjRes.informe.filter(function (x) {
      return !x.ok;
    });

    return {
      success: fallos.length === 0,
      dryRun: true,
      message: fallos.length
        ? "Vista previa OK; revise adjuntos con error (no se envió nada)."
        : "Validación OK: destinatarios, cuerpo y adjuntos listos. No se envió correo ni se guardó gestión.",
      resumen: {
        para: preview.para,
        cc: preview.cc,
        asunto: preview.subject,
        segmento: preview.segmentoNorm,
        numCandidatos: (preview.candidatos || []).length,
        tamanoHtmlCuerpo: (preview.htmlBody || "").length,
        numAdjuntosSolicitados: (payload.emailAdjuntos || []).length,
        numAdjuntosBlobOk: adjRes.blobs.length,
        asignacionCorreccion: preview.asignacionCorreccion || null
      },
      adjuntosInforme: adjRes.informe
    };
  } catch (e) {
    return { success: false, dryRun: true, message: "Error: " + e.toString() };
  }
}

/**
 * Ejecutar en el editor de Apps Script (Ejecutar > ejecutarPruebasAutomaticasCorreoRenovacion).
 * No envía correos. Revisa rutas de preview + adjuntos con datos mínimos.
 */
function ejecutarPruebasAutomaticasCorreoRenovacion() {
  var resultados = [];
  var mockLead = {
    poliza: "PRUEBA_INTERNA_001",
    solicitud: "PRUEBA_INTERNA_001",
    asegurado: "Cliente Prueba QA",
    documento: "123",
    segmento: "PROPIETARIO",
    email: "qa.propietario@example.com"
  };
  var p1 = probarCorreoRenovacionSinEnvio({
    tipoAccionCorreo: "APPROVE",
    dataLead: mockLead,
    leadSelect: {},
    observations: "Prueba automática",
    notasAnalista: "Nota QA",
    ccSeleccionados: [],
    emailAdjuntos: []
  });
  resultados.push({ caso: "APPROVE propietario sin adjuntos", ok: p1.success, detalle: p1 });

  var p2 = probarCorreoRenovacionSinEnvio({
    tipoAccionCorreo: "CORRECTION",
    dataLead: mockLead,
    leadSelect: {},
    observations: "Motivo prueba corrección",
    segmento: "PROPIETARIO",
    ccSeleccionados: []
  });
  resultados.push({ caso: "CORRECTION con observaciones", ok: p2.success, detalle: p2 });

  var p3 = buildRenovacionCorreoPreviewPayload_({
    tipoAccionCorreo: "APPROVE",
    dataLead: { poliza: "P2", segmento: "BROKER", email: "b@example.com" },
    leadSelect: {},
    emailAnalistaAsignado: "jeison.sanchez@aselibertador.com",
    observations: "x",
    notasAnalista: ""
  });
  resultados.push({
    caso: "Preview solo segmento BROKER",
    ok: !!(p3 && p3.success),
    detalle: p3
  });

  var p4 = probarCorreoRenovacionSinEnvio({
    tipoAccionCorreo: "CORRECTION",
    dataLead: Object.assign({}, mockLead, { segmento: "BROKER" }),
    leadSelect: {},
    observations: "Motivo prueba cola BI",
    segmento: "BROKER",
    ccSeleccionados: []
  });
  var colaOk =
    p4.resumen &&
    p4.resumen.asignacionCorreccion &&
    p4.resumen.asignacionCorreccion.cola === "CorreccionesBI";
  resultados.push({
    caso: "CORRECTION segmento BROKER → cola CorreccionesBI",
    ok: !!p4.success && !!colaOk,
    detalle: p4
  });

  Logger.log(JSON.stringify(resultados, null, 2));
  return resultados;
}

/**
 * Envía el correo (solo tras aprobación manual en UI) y ejecuta processAnalystDecision.
 * payload: mismo que processAnalystDecision + tipoAccionCorreo ('APPROVE'|'CORRECTION'|'CANCELADA') + correoAdicional1/2 + emailAnalistaAsignado + segmentoSheetColumn (opcional) + notasAnalista (aprobar)
 * emailAdjuntos: [{ tipo:'drive', fileId, nombre }, { tipo:'base64', data, mimeType, nombre }]
 */
function procesarDecisionAnalistaConCorreo(payload) {
  try {
    var tipo = payload.tipoAccionCorreo || payload.decision;
    if (tipo === "APPROVE") tipo = "APPROVE";
    else if (tipo === "CORRECTION") tipo = "CORRECTION";
    else if (tipo === "CANCELADA" || tipo === "CANCEL") tipo = "CANCELADA";

    if (tipo !== "APPROVE" && tipo !== "CORRECTION" && tipo !== "CANCELADA") {
      return { success: false, message: "Tipo de acción de correo no válido." };
    }

    var dataLead = payload.dataLead || {};
    var poliza = dataLead.poliza || dataLead.solicitud;

    if ((tipo === "CORRECTION" || tipo === "CANCELADA") && !(payload.observations && String(payload.observations).trim())) {
      return { success: false, message: "Debe indicar observaciones / motivo." };
    }

    var prevPayload = {
      tipoAccionCorreo: tipo,
      dataLead: dataLead,
      leadSelect: payload.leadSelect,
      segmentoSheetColumn: payload.segmentoSheetColumn,
      emailAnalistaAsignado: payload.emailAnalistaAsignado,
      observations: payload.observations,
      notasAnalista: payload.notasAnalista,
      correoAdicional1: payload.correoAdicional1,
      correoAdicional2: payload.correoAdicional2,
      ccSeleccionados: payload.ccSeleccionados
    };
    if (payload.segmento != null && String(payload.segmento).trim() !== "") {
      prevPayload.segmento = payload.segmento;
    }
    var preview = buildRenovacionCorreoPreviewPayload_(prevPayload);

    if (!preview.success) return preview;

    var adjProc = _adjuntosRenovCorreoProcesar_(payload.emailAdjuntos || []);
    if (payload.emailAdjuntos && payload.emailAdjuntos.length) {
      var fallosAdj = adjProc.informe.filter(function (x) {
        return !x.ok;
      });
      if (fallosAdj.length) {
        return {
          success: false,
          message: "Adjuntos: " +
            fallosAdj
              .map(function (f) {
                return f.nombre + " — " + (f.error || "error");
              })
              .join(" | "),
          adjuntosInforme: adjProc.informe
        };
      }
    }

    var resProceso = processAnalystDecision(payload);
    if (!resProceso.success) {
      return resProceso;
    }

    var ccStr = preview.cc.length ? preview.cc.join(",") : "";
    var mailOpts = { htmlBody: preview.htmlBody, noReply: true };
    if (ccStr) mailOpts.cc = ccStr;
    if (adjProc.blobs.length) mailOpts.attachments = adjProc.blobs;
    try {
      GmailApp.sendEmail(preview.para, preview.subject, "", mailOpts);
    } catch (mailErr) {
      return {
        success: false,
        message: "La gestión se guardó en la hoja, pero falló el envío del correo: " + String(mailErr)
      };
    }

    var nombresAdj = [];
    if (payload.emailAdjuntos && payload.emailAdjuntos.length) {
      for (var ad = 0; ad < payload.emailAdjuntos.length; ad++) {
        var x = payload.emailAdjuntos[ad];
        nombresAdj.push(x.nombre || x.fileId || "archivo");
      }
    }

    var entry = {
      fecha: new Date().toISOString(),
      accion: tipo,
      para: preview.para,
      cc: preview.cc,
      asunto: preview.subject,
      usuario: Session.getActiveUser().getEmail(),
      segmento: preview.segmentoNorm,
      adjuntos: nombresAdj
    };
    appendLogCorreoRenovacion_(poliza, entry);

    return { success: true, message: "Gestión actualizada y correo enviado.", correo: entry };
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}


