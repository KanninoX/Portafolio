/**
 * ============================================================================
 * CARGAR DOCUMENTOS POR CÓDIGO — webhook de contratación (BUK)
 */

/**
 * Modelo Prevención de Delitos: aplica a todo trabajador nuevo, pero los
 * documentos cambian según la empresa (ver `carpeta_mpd` en `getEmployeeData`).
 * En BUK el parámetro `path` crea la carpeta si todavía no existe.
 */
const MPD_BUK_PATH  = "Modelo Prevención de Delitos";

/** Son informativos: visibles para el trabajador y sin firma de ninguna parte. */
const MPD_CONFIG = {
  empresa_name: MPD_BUK_PATH,
  visible_trabajador: true,
  firmable_por_el_trabajador: false
};


function uploadToBUK_v2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  ss.toast('Iniciando proceso de carga...', 'Proceso iniciado', 3);
  limpiarRegistrosAnteriores(sheet);

  const authToken = _authToken_();
  const sourceFolderName = "Documentos a cargar en BUK";

  const destinationFolder = sheet.getRange("D2").getValue();
  const visibilityValue = sheet.getRange("E2").getValue();
  const isVisible = visibilityValue.toString().toLowerCase() === "sí";

  const signatureValue = sheet.getRange("F2").getValue();
  const requiresSignature = signatureValue.toString().toLowerCase() === "sí";

  const deleteValue = sheet.getRange("G2").getValue();
  const deleteFromDrive = deleteValue.toString().toLowerCase() === "sí";

  let filesWithError = 0;

  // Leer datos desde hoja "Empleados"
  const empleadosSheet = ss.getSheetByName("Empleados");
  const empleadosData = empleadosSheet.getRange("A2:B" + empleadosSheet.getLastRow()).getValues();

  const codeSheetToId = {};
  empleadosData.forEach(([code, id]) => {
    if (code && id) {
      codeSheetToId[code.toString().trim()] = id.toString().trim();
    }
  });

  let sourceFolder;
  try {
    sourceFolder = DriveApp.getFoldersByName(sourceFolderName).next();
  } catch (e) {
    throw new Error(`La carpeta "${sourceFolderName}" no existe en Google Drive.`);
  }

  const files = sourceFolder.getFiles();
  const filesArray = [];
  while (files.hasNext()) {
    filesArray.push(files.next());
  }

  if (filesArray.length === 0) {
    throw new Error(`No hay archivos para procesar en la carpeta "${sourceFolderName}".`);
  }

  let filesProcessed = 0;
  filesArray.forEach(file => {
    const fileName = file.getName();
    let codeSheet = "";
    let resultado = "";
    const fechaEjecucion = new Date();

    try {
      console.log(`Procesando archivo: ${fileName}`);

      // Buscar código de 4 a 6 dígitos
      const match = fileName.match(/\d{4,6}/);
      codeSheet = match ? match[0] : null;

      if (!codeSheet) {
        resultado = "Error: Código no encontrado";
        filesWithError++;
        registrarResultado(sheet, fileName, "N/A", resultado, fechaEjecucion);
        return;
      }

      // Buscar ID del trabajador desde hoja local
      let employeeId = codeSheetToId[codeSheet] || codeSheetToId[codeSheet.padStart(6, '0')];

      // Si no se encuentra en la hoja local, consultar la API de BUK
      if (!employeeId) {
        console.log(`ID no encontrado localmente para código ${codeSheet}. Consultando API de BUK...`);

        const apiEmployeeData = getEmployeeByCodeSheet(codeSheet, authToken);

        if (apiEmployeeData && apiEmployeeData.id) {
          employeeId = apiEmployeeData.id.toString();

          // Actualizar la hoja local con el nuevo empleado encontrado
          updateEmployeesSheet(codeSheet, employeeId);

          console.log(`Empleado encontrado en BUK API: ID ${employeeId} para código ${codeSheet}`);
          resultado = `Encontrado en API BUK - Subido exitosamente`;
        } else {
          resultado = "Error: Trabajador no encontrado en BUK";
          filesWithError++;
          registrarResultado(sheet, fileName, codeSheet, resultado, fechaEjecucion);
          return;
        }
      }

      // Continuar con la subida del documento
      const apiUrl = `${_bukBase_()}/employees/${employeeId}/docs?path=${encodeURIComponent(destinationFolder)}&visible=${isVisible}&signable_by_employee=${requiresSignature}`;

      const payload = {
        file: file.getBlob(),
        file_url: file.getUrl(),
        original_filename: file.getName(),
        employee_file: {
          is_visible: isVisible,
          settings: {
            pdf_orientation: "Portrait",
            employee_sign: requiresSignature,
            legal_agent_sign: false,
            second_legal_agent_sign: false,
            is_external: true
          }
        }
      };

      const options = {
        method: "post",
        headers: {
          "Accept": "application/json",
          "auth_token": authToken
        },
        payload: payload,
        followRedirects: true,
        validateHttpsCertificates: true,
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(apiUrl, options);
      console.log(`Respuesta BUK: ${response.getContentText()}`);

      if (resultado !== `Encontrado en API BUK - Subido exitosamente`) {
        resultado = "Subido exitosamente";
      }
      filesProcessed++;

      // Eliminar el archivo si corresponde
      if (deleteFromDrive) {
        file.setTrashed(true);
      }

    } catch (error) {
      console.log(`Error procesando archivo: ${error.message}`);
      resultado = `Error: ${error.message}`;
      filesWithError++;
    }

    registrarResultado(sheet, fileName, codeSheet || "N/A", resultado, fechaEjecucion);
  });

  const message = `Proceso completado.\nArchivos procesados: ${filesProcessed}\nErrores: ${filesWithError}`;
  ss.toast(message, 'Proceso finalizado', 5);

  return { processed: filesProcessed, errors: filesWithError };
}

// Nueva función para consultar empleado por code_sheet en la API de BUK
function getEmployeeByCodeSheet(codeSheet, authToken) {
  const apiUrl = `${_bukBase_()}/employees?code_sheet=${codeSheet}`;

  try {
    console.log(`Consultando API BUK para código: ${codeSheet}`);
    console.log(`URL: ${apiUrl}`);

    const options = {
      method: "GET",
      headers: {
        "Accept": "application/json",
        "auth_token": authToken
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    console.log(`Respuesta API BUK: HTTP ${responseCode}`);
    console.log(`Contenido: ${responseText}`);

    if (responseCode === 200) {
      const responseData = JSON.parse(responseText);

      // Verificar que hay datos y al menos un empleado
      if (responseData.data && responseData.data.length > 0) {
        const employee = responseData.data[0]; // Tomar el primer empleado encontrado

        console.log(`Empleado encontrado: ID ${employee.id}, Nombre: ${employee.full_name}, Code Sheet: ${employee.code_sheet}`);

        return {
          id: employee.id,
          full_name: employee.full_name,
          code_sheet: employee.code_sheet,
          rut: employee.rut
        };
      } else {
        console.log(`No se encontraron empleados con code_sheet: ${codeSheet}`);
        return null;
      }
    } else {
      console.log(`Error HTTP ${responseCode} consultando empleado: ${responseText}`);
      return null;
    }

  } catch (error) {
    console.log(`Error consultando API BUK para código ${codeSheet}: ${error.toString()}`);
    return null;
  }
}

function registrarResultado(sheet, archivo, codigo, resultado, fecha) {
  const lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 4).setValue(archivo);   // D - Nombre de documento
  sheet.getRange(lastRow, 5).setValue(codigo);    // E - Código
  sheet.getRange(lastRow, 6).setValue(resultado); // F - Resultado
  sheet.getRange(lastRow, 7).setValue(fecha);     // G - Fecha de Carga
}

// ════════════════════════════════════════════════════════════════════════
//  CREDENCIALES — nunca en el código
// ════════════════════════════════════════════════════════════════════════
//
// El token y el subdominio se leen de las Propiedades de la secuencia de
// comandos. Cárgalos UNA vez y quedan guardados:
//
//   Configuración del proyecto  →  Propiedades de la secuencia de comandos
//   →  Añadir propiedad
//        BUK_AUTH_TOKEN   =  (el token NUEVO, después de rotarlo en BUK)
//        BUK_SUBDOMINIO   =  (el subdominio real, sin https:// ni .buk.cl)
//
// Se hace desde esa pantalla y no desde el código a propósito: así el valor
// no queda escrito en ninguna parte que se pueda copiar, compartir o commitear.
// ════════════════════════════════════════════════════════════════════════

function _cfg_(clave) {
  var v = PropertiesService.getScriptProperties().getProperty(clave);
  if (!v) {
    throw new Error('Falta la propiedad ' + clave + '. Cárgala en ' +
      'Configuración del proyecto → Propiedades de la secuencia de comandos.');
  }
  return String(v).trim();
}

/** Token de la API de BUK. */
function _authToken_() {
  return _cfg_('BUK_AUTH_TOKEN');
}

/** Base de la API: https://{subdominio}.buk.cl/api/v1/chile */
function _bukBase_() {
  var sub = _cfg_('BUK_SUBDOMINIO')
    .replace(/^https?:\/\//i, '')
    .split('/')[0]
    .replace(/\.buk\.cl$/i, '');
  return 'https://' + sub + '.buk.cl/api/v1/chile';
}

/** Comprueba que las credenciales estén cargadas. No imprime el token. */
function verificarCredenciales() {
  var props = PropertiesService.getScriptProperties();
  var tok   = props.getProperty('BUK_AUTH_TOKEN');
  var sub   = props.getProperty('BUK_SUBDOMINIO');
  var msg =
    'BUK_AUTH_TOKEN: ' + (tok ? 'cargado (' + tok.length + ' caracteres)' : '✗ FALTA') + '\n' +
    'BUK_SUBDOMINIO: ' + (sub ? sub : '✗ FALTA') + '\n' +
    (tok && sub ? 'Base de la API: ' + _bukBase_() : 'Cárgalas antes de usar el script.');
  console.log(msg);
  return msg;
}


// ════════════════════════════════════════════════════════════════════════
//  WEBHOOK — acuse inmediato + cola
// ════════════════════════════════════════════════════════════════════════

var COLA_PROP_ = 'COLA_WEBHOOK';

/**
 * Responde a BUK en milisegundos. Lo único que hace es anotar el evento.
 * El trabajo pesado corre después, en procesarColaWebhook().
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents).data || {};
    var tipo = data.event_type;

    if (tipo === 'employee_create' || tipo === 'job_hire') {
      var employeeId = (data.employee_id && data.employee_id.id)
        ? data.employee_id.id
        : data.employee_id;
      _encolarEvento_(tipo, employeeId);
      console.log('Encolado ' + tipo + ' para ' + employeeId);
    } else {
      console.log('Evento ignorado: ' + tipo);
    }
  } catch (error) {
    console.log('✗ Error leyendo webhook: ' + error);
  }

  // Siempre 200, siempre rápido. Si BUK no recibe esto a tiempo, reintenta.
  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Agrega el evento a la cola, sin duplicar lo que ya está pendiente. */
function _encolarEvento_(tipo, employeeId) {
  if (!employeeId) return;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    console.log('✗ No se pudo tomar el candado para encolar ' + employeeId);
    return;
  }
  try {
    var props = PropertiesService.getScriptProperties();
    var cola  = JSON.parse(props.getProperty(COLA_PROP_) || '[]');

    var yaEsta = cola.some(function (item) {
      return item.tipo === tipo && String(item.employeeId) === String(employeeId);
    });
    if (yaEsta) {
      console.log('Ya estaba en cola: ' + tipo + ' ' + employeeId);
      return;
    }

    cola.push({ tipo: tipo, employeeId: String(employeeId), ts: Date.now() });
    props.setProperty(COLA_PROP_, JSON.stringify(cola));
  } finally {
    lock.releaseLock();
  }
}

/**
 * Vacía la cola. Va en un activador de tiempo, cada 1 minuto.
 * Saca los eventos de la cola ANTES de procesarlos, para que dos ciclos
 * solapados no tomen el mismo.
 */
function procesarColaWebhook() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;

  var pendientes;
  try {
    var props = PropertiesService.getScriptProperties();
    pendientes = JSON.parse(props.getProperty(COLA_PROP_) || '[]');
    if (!pendientes.length) return;
    props.setProperty(COLA_PROP_, '[]');
  } finally {
    lock.releaseLock();
  }

  console.log('Procesando ' + pendientes.length + ' evento(s) en cola');

  pendientes.forEach(function (item) {
    try {
      if (item.tipo === 'employee_create') {
        addNewEmployee(item.employeeId);
      } else if (item.tipo === 'job_hire') {
        handleJobHire(item.employeeId);
      }
    } catch (error) {
      console.log('✗ Error procesando ' + item.tipo + ' ' + item.employeeId + ': ' + error);
    }
  });
}

function addNewEmployee(employeeId) {
  const authToken = _authToken_();
  const apiUrl = `${_bukBase_()}/employees/${employeeId}`;

  try {
    const options = {
      method: "GET",
      headers: {
        "Accept": "application/json",
        "auth_token": authToken
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseData = JSON.parse(response.getContentText());

    if (response.getResponseCode() === 200 && responseData.data) {
      const codeSheet = responseData.data.code_sheet || "";
      updateEmployeesSheet(codeSheet, employeeId);
    } else {
      console.log(`Error obteniendo datos del empleado ${employeeId}: ${response.getContentText()}`);
    }

  } catch (error) {
    console.log(`Error consultando API para empleado ${employeeId}: ${error.toString()}`);
  }
}

function updateEmployeesSheet(codeSheet, employeeId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const empleadosSheet = ss.getSheetByName('Empleados');

    if (!empleadosSheet) {
      console.log('No se encontró la hoja "Empleados"');
      return;
    }

    // Verificar si el empleado ya existe antes de agregarlo
    const existingData = empleadosSheet.getRange("A2:B" + empleadosSheet.getLastRow()).getValues();
    const employeeExists = existingData.some(([code, id]) =>
      code.toString().trim() === codeSheet.toString().trim() ||
      id.toString().trim() === employeeId.toString().trim()
    );

    if (!employeeExists) {
      const lastRow = empleadosSheet.getLastRow() + 1;
      empleadosSheet.getRange(lastRow, 1).setValue(codeSheet);
      empleadosSheet.getRange(lastRow, 2).setValue(employeeId);
      console.log(`Nuevo empleado agregado: Code Sheet: ${codeSheet}, Employee ID: ${employeeId}`);
    } else {
      console.log(`Empleado ya existe: Code Sheet: ${codeSheet}, Employee ID: ${employeeId}`);
    }

  } catch (error) {
    console.log(`Error actualizando hoja Empleados: ${error.toString()}`);
  }
}

/**
 * Reserva al empleado ANTES de tocar Drive. Ese es el punto: marcar al final
 * no sirve, porque las entregas duplicadas llegan mientras la primera todavía
 * está subiendo archivos.
 */
function handleJobHire(employeeId) {
  var VENTANA_MS = 24 * 60 * 60 * 1000;   // una recontratación real, meses después, sí pasa
  var props = PropertiesService.getScriptProperties();
  var key   = 'JOBHIRE_' + employeeId;
  var lock  = LockService.getScriptLock();

  if (!lock.tryLock(10000)) {
    console.log('job_hire ' + employeeId + ': candado ocupado, se ignora esta entrega');
    return;
  }
  try {
    var previo = props.getProperty(key);
    if (previo && (Date.now() - Number(previo)) < VENTANA_MS) {
      var hace = Math.round((Date.now() - Number(previo)) / 1000);
      console.log('job_hire ' + employeeId + ': DUPLICADO, ya procesado hace ' + hace + ' s. Ignorado.');
      return;
    }
    props.setProperty(key, String(Date.now()));   // ← reserva ANTES de subir nada
  } finally {
    lock.releaseLock();
  }

  try {
    console.log(`=== INICIANDO PROCESO JOB_HIRE ===`);
    console.log(`Procesando contratación para empleado ID: ${employeeId}`);

    const employeeData = getEmployeeData(employeeId);
    if (!employeeData) {
      console.log(`✗ FALLO: No se pudo obtener los datos del empleado ${employeeId}`);
      props.deleteProperty(key);   // no se subió nada: que un reintento pueda trabajar
      return;
    }

    console.log(`✓ Empresa obtenida exitosamente: ${employeeData.company.empresa_name}`);
    console.log(`✓ Empleado: ${employeeData.employeeName}`);

    copyDocumentsFromDrive(employeeId, employeeData.company);

    // Modelo Prevención de Delitos: para todos, con los documentos de su empresa
    copyMPDDocuments(employeeId, employeeData.company);

    // Carga adicional de documentos IRL
    processIRLDocument(employeeId, _authToken_(), employeeData);

    registrarEjecucionJobHire(
      employeeId,
      employeeData.employeeName,
      employeeData.company.empresa_name,
      new Date()
    );

    console.log(`=== PROCESO JOB_HIRE COMPLETADO ===`);

  } catch (error) {
    console.log(`✗ ERROR CRÍTICO en job_hire para empleado ${employeeId}: ${error.toString()}`);
    props.deleteProperty(key);     // permitir reintento manual tras un fallo real
  }
}

function getEmployeeData(employeeId) {
  const authToken = _authToken_();
  const apiUrl = `${_bukBase_()}/employees/${employeeId}`;

  try {
    console.log(`Obteniendo datos del empleado ${employeeId} desde BUK...`);

    const options = {
      method: "GET",
      headers: {
        "Accept": "application/json",
        "auth_token": authToken
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    console.log(`Respuesta BUK para empleado ${employeeId}: HTTP ${responseCode}`);
    console.log(`Contenido respuesta: ${responseText}`);

    if (responseCode === 200) {
      const responseData = JSON.parse(responseText);

      if (responseData.data) {
        const companyId = responseData.data.current_job ? responseData.data.current_job.company_id : responseData.data.company_id;
        const employeeName = `${responseData.data.first_name || ''} ${responseData.data.last_name || ''}`.trim() || 'Sin nombre';

        // Obtener código de cargo
        const roleData = responseData.data.role || (responseData.data.current_job ? responseData.data.current_job.role : null);
        const roleCode = roleData ? roleData.code : null;

        // Obtener Lugar de Trabajo (Atributo personalizado en current_job)
        const customAttrs = responseData.data.current_job ? responseData.data.current_job.custom_attributes : {};
        console.log(`[DEBUG] Atributos personalizados encontrados: ${JSON.stringify(customAttrs)}`);

        const workplace = customAttrs["Lugar de trabajo"] || null;

        console.log(`ID de empresa obtenido: ${companyId}`);
        console.log(`Nombre del empleado: ${employeeName}`);
        console.log(`Cargo (code): ${roleCode}`);
        console.log(`Lugar de trabajo: ${workplace}`);

        const companies = [
          {
            "id": 1,
            "empresa_name": "Invermar S.A.",
            "carpeta_drive": "https://drive.google.com/drive/folders/18r7MJALL05qZmtIAfM1taANT8lBL3rLg?usp=drive_link",
            "carpeta_mpd": "https://drive.google.com/drive/folders/1Lf390LzN38C9z3LIRoJs2jj81KrR5pDK?usp=drive_link",
            "visible_trabajador": true,
            "firmable_por_el_trabajador": false
          },
          {
            "id": 2,
            "empresa_name": "Pesquera La Portada S.A.",
            "carpeta_drive": "https://drive.google.com/drive/folders/1wBAYUyfhQWUk3FJAV_4iOVEhtkHBwQL3?usp=drive_link",
            "carpeta_mpd": "https://drive.google.com/drive/folders/1wykdnokuWXjAzkBP8_Nbw4l-ko5uq-A-?usp=drive_link",
            "visible_trabajador": true,
            "firmable_por_el_trabajador": true
          },
          {
            "id": 3,
            "empresa_name": "La Península S.A.",
            "carpeta_drive": "https://drive.google.com/drive/folders/1bKhbAfiXmayL_VjKJfNa4zxGLYo2rOXB?usp=drive_link",
            "carpeta_mpd": "https://drive.google.com/drive/folders/1dLJa0sdVpLajkXp229UJMJsZfnxarqkt?usp=drive_link",
            "visible_trabajador": true,
            "firmable_por_el_trabajador": true
          },
          {
            "id": 4,
            "empresa_name": "Astilleros Calbuco S.A.",
            "carpeta_drive": "https://drive.google.com/drive/folders/1fEJHIULau2LrkBisb6_ydFlQxxBohLEE?usp=drive_link",
            "carpeta_mpd": "https://drive.google.com/drive/folders/1aWyfnBHlAcY6d0u5vjjK93enl2xF5ttr?usp=drive_link",
            "visible_trabajador": true,
            "firmable_por_el_trabajador": true
          }
        ];

        const foundCompany = companies.find(company => company.id === companyId);

        if (foundCompany) {
          console.log(`✓ Empresa encontrada: ID ${foundCompany.id} - ${foundCompany.empresa_name}`);
          return {
            employeeName: employeeName,
            company: foundCompany,
            roleCode: roleCode,
            workplace: workplace
          };
        } else {
          console.log(`✗ No se encontró configuración para la empresa con ID: ${companyId}`);
          console.log(`IDs de empresas disponibles: ${companies.map(c => `${c.id} (${c.empresa_name})`).join(', ')}`);
          return null;
        }
      } else {
        console.log(`✗ No hay datos del empleado en la respuesta`);
        return null;
      }
    } else {
      console.log(`✗ Error HTTP ${responseCode}: ${responseText}`);
      return null;
    }

  } catch (error) {
    console.log(`✗ Error obteniendo datos del empleado ${employeeId}: ${error.toString()}`);
    return null;
  }
}

function getEmployeeCompany(employeeId) {
  const employeeData = getEmployeeData(employeeId);
  return employeeData ? employeeData.company : null;
}

function copyDocumentsFromDrive(employeeId, company) {
  try {
    console.log(`Iniciando copia de documentos para empleado ${employeeId} de empresa ${company.empresa_name}`);

    const folderId = extractFolderIdFromUrl(company.carpeta_drive);
    if (!folderId) {
      console.log(`✗ ID de carpeta no válido para empresa ${company.empresa_name}. URL: ${company.carpeta_drive}`);
      return;
    }

    console.log(`ID de carpeta extraído: ${folderId}`);

    const sourceFolder = DriveApp.getFolderById(folderId);
    const files = sourceFolder.getFiles();

    let fileCount = 0;
    while (files.hasNext()) {
      const file = files.next();
      fileCount++;
      console.log(`Procesando archivo ${fileCount}: ${file.getName()}`);
      uploadDocumentToBUK(employeeId, file, company);
    }

    console.log(`✓ Proceso completado. Total de archivos procesados: ${fileCount}`);

  } catch (error) {
    console.log(`✗ Error copiando documentos para empleado ${employeeId}: ${error.toString()}`);
  }
}

/**
 * Sube los documentos del Modelo Prevención de Delitos a su propia carpeta en
 * BUK. Van para todos los trabajadores, con los documentos de su empresa.
 */
function copyMPDDocuments(employeeId, company) {
  try {
    console.log(`Iniciando copia de documentos MPD para empleado ${employeeId} de empresa ${company.empresa_name}`);

    const folderId = extractFolderIdFromUrl(company.carpeta_mpd || "");
    if (!folderId) {
      console.log(`[MPD] ✗ No hay carpeta MPD configurada para ${company.empresa_name}`);
      return;
    }

    const sourceFolder = DriveApp.getFolderById(folderId);
    const files = sourceFolder.getFiles();

    let fileCount = 0;
    while (files.hasNext()) {
      const file = files.next();
      fileCount++;
      console.log(`[MPD] Procesando archivo ${fileCount}: ${file.getName()}`);
      uploadDocumentToBUK(employeeId, file, MPD_CONFIG, MPD_BUK_PATH);
    }

    if (fileCount === 0) {
      console.log(`[MPD] ⚠ La carpeta de origen no tiene archivos (ID: ${folderId})`);
    } else {
      console.log(`[MPD] ✓ Proceso completado. Total de archivos procesados: ${fileCount}`);
    }

  } catch (error) {
    console.log(`[MPD] ✗ Error copiando documentos MPD para empleado ${employeeId}: ${error.toString()}`);
  }
}

function extractFolderIdFromUrl(driveUrl) {
  const match = driveUrl.match(/\/folders\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

function uploadDocumentToBUK(employeeId, file, company, carpetaDestino) {
  const authToken = _authToken_();
  const destinationFolder = carpetaDestino || "Documentos entregados";

  try {
    console.log(`Iniciando subida de documento: ${file.getName()} para empleado ${employeeId} → carpeta "${destinationFolder}"`);

    const apiUrl = `${_bukBase_()}/employees/${employeeId}/docs?path=${encodeURIComponent(destinationFolder)}&visible=${company.visible_trabajador}&signable_by_employee=${company.firmable_por_el_trabajador}`;

    console.log(`URL de la API: ${apiUrl}`);

    const payload = {
      file: file.getBlob(),
      file_url: file.getUrl(),
      original_filename: file.getName(),
      employee_file: {
        is_visible: company.visible_trabajador,
        settings: {
          pdf_orientation: "Portrait",
          employee_sign: company.firmable_por_el_trabajador,
          legal_agent_sign: false,
          second_legal_agent_sign: false,
          is_external: true
        }
      }
    };

    const options = {
      method: "post",
      headers: {
        "Accept": "application/json",
        "auth_token": authToken
      },
      payload: payload,
      followRedirects: true,
      validateHttpsCertificates: true,
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    console.log(`Respuesta HTTP ${responseCode}: ${responseText}`);

    if (responseCode >= 200 && responseCode < 300) {
      console.log(`✓ Documento ${file.getName()} subido exitosamente para empleado ${employeeId}`);
    } else {
      console.log(`✗ Error al subir documento ${file.getName()}: HTTP ${responseCode} - ${responseText}`);
    }

  } catch (error) {
    console.log(`✗ Error crítico subiendo documento ${file.getName()} para empleado ${employeeId}: ${error.toString()}`);
  }
}

function registrarEjecucionJobHire(employeeId, nombreEmpleado, nombreEmpresa, fechaEjecucion) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName("Registro JOB_HIRE");

    if (!logSheet) {
      logSheet = ss.insertSheet("Registro JOB_HIRE");
      logSheet.getRange(1, 1, 1, 4).setValues([["ID", "Nombre Empleado", "Empresa", "Fecha Ejecución"]]);
      logSheet.getRange(1, 1, 1, 4).setFontWeight("bold");
    }

    const lastRow = logSheet.getLastRow() + 1;
    logSheet.getRange(lastRow, 1).setValue(employeeId);
    logSheet.getRange(lastRow, 2).setValue(nombreEmpleado);
    logSheet.getRange(lastRow, 3).setValue(nombreEmpresa);
    logSheet.getRange(lastRow, 4).setValue(fechaEjecucion);

    console.log(`Registro JOB_HIRE guardado: ID ${employeeId}, ${nombreEmpleado}, ${nombreEmpresa}`);

  } catch (error) {
    console.log(`Error registrando ejecución JOB_HIRE: ${error.toString()}`);
  }
}

function limpiarRegistrosAnteriores(sheet) {
  try {
    // Obtener la última fila con datos
    const lastRow = sheet.getLastRow();

    // Si hay más de 23 filas (ajusta este número según donde empiecen tus registros)
    if (lastRow > 23) {
      // Limpiar desde la fila 24 hasta la última fila, columnas D a G
      const rangeToDelete = sheet.getRange(24, 4, lastRow - 23, 4); // D:G
      rangeToDelete.clearContent();

      console.log(`Registros anteriores limpiados: filas 24-${lastRow}, columnas D-G`);
    }
  } catch (error) {
    console.log(`Error limpiando registros anteriores: ${error.toString()}`);
  }
}

/**
 * Función Carga IRL: Carga un documento en la Carpeta SSO de BUK
 * basándose en el cargo del trabajador.
 * Solo aplica para la empresa INVERMAR S.A. (ID 1)
 */
function processIRLDocument(employeeId, authToken, employeeData) {
  const IRL_FOLDER_ID = "1JHU0mY1nlYa25PuFpXS67M8sKAnlo8hF";

  // LOG INMEDIATO: Confirmamos que la función fue llamada
  registrarLogSSO(employeeId, "N/A", "INICIANDO", "Procesando documentos IRL...");

  try {
    // 1. Obtener mapeo dinámico desde el Excel
    const IRL_MAPPING = obtenerMapeoIRLExcel();

    // 2. Obtener datos si no fueron pasados
    if (!employeeData) {
      employeeData = getEmployeeData(employeeId);
    }

    if (!employeeData) {
      registrarLogSSO(employeeId, "N/A", "ERROR", "No se pudieron obtener datos del empleado desde la API");
      return;
    }

    const roleCodeRaw = employeeData.roleCode || "";
    const roleCode = roleCodeRaw.toString().trim().toLowerCase(); // Normalizamos para evitar fallos por mayúsculas/espacios
    const workplace = (employeeData.workplace || "").toString().trim();
    const companyId = employeeData.company ? employeeData.company.id : null;
    const employeeName = employeeData.employeeName;

    registrarLogSSO(employeeId, employeeName, "DATOS", `Cargo: ${roleCode}, Lugar: ${workplace}, Empresa ID: ${companyId}`);

    // VALIDACIÓN 1: Empresa
    if (companyId !== 1) {
      registrarLogSSO(employeeId, employeeName, "OMITIDO", `Empresa ID detectado: ${companyId}. Solo aplica para Invermar (ID 1).`);
      return;
    }

    // Buscamos el documento en el mapeo (usando el código normalizado)
    let documentName = "";
    for (let key in IRL_MAPPING) {
      if (key.trim().toLowerCase() === roleCode) {
        documentName = IRL_MAPPING[key];
        break;
      }
    }

    // VALIDACIÓN 2: Mapeo de Cargo
    if (!documentName) {
      registrarLogSSO(employeeId, employeeName, "OMITIDO", `El cargo "${roleCodeRaw}" no tiene un documento configurado en la hoja 'IRL'.`);
      return;
    }

    // VALIDACIÓN 3: Caso especial Operario (inv_118)
    if (roleCode === "inv_118") {
      const workplaceLower = workplace.toLowerCase();
      const isValidWorkplace = workplaceLower.includes("centro de cultivo") || workplaceLower.includes("centros de cultivo");

      if (!isValidWorkplace) {
        registrarLogSSO(employeeId, employeeName, "OMITIDO", `Operario en lugar no autorizado: "${workplace}"`);
        return;
      }
    }

    // 4. Carga del documento
    registrarLogSSO(employeeId, employeeName, "BUSCANDO", `Buscando archivo: ${documentName}`);

    const folder = DriveApp.getFolderById(IRL_FOLDER_ID);
    const files = folder.getFilesByName(documentName);

    if (!files.hasNext()) {
      registrarLogSSO(employeeId, employeeName, "ERROR", `No se encontró el archivo "${documentName}" dentro de la carpeta Drive.`);
      return;
    }

    const file = files.next();
    const apiUrl = `${_bukBase_()}/employees/${employeeId}/docs?path=${encodeURIComponent("SSO")}&visible=true&signable_by_employee=false`;

    const payload = {
      file: file.getBlob(),
      file_url: file.getUrl(),
      original_filename: file.getName(),
      employee_file: {
        is_visible: true,
        settings: {
          pdf_orientation: "Portrait",
          employee_sign: false,
          legal_agent_sign: false,
          second_legal_agent_sign: false,
          is_external: true
        }
      }
    };

    const options = {
      method: "post",
      headers: {
        "Accept": "application/json",
        "auth_token": authToken
      },
      payload: payload,
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const respCode = response.getResponseCode();

    if (respCode >= 200 && respCode < 300) {
      registrarLogSSO(employeeId, employeeName, "ÉXITO", `Subido correctamente: ${documentName}`);
      console.log(`[IRL] ✓ ÉXITO: Subido documento ${documentName}`);
    } else {
      registrarLogSSO(employeeId, employeeName, "ERROR API BUK", `HTTP ${respCode}: ${response.getContentText()}`);
    }

  } catch (error) {
    console.log(`[IRL] ✗ ERROR CRÍTICO: ${error.toString()}`);
    registrarLogSSO(employeeId, "N/A", "ERROR CRÍTICO", error.toString());
  }
}

/**
 * Registra cada paso de la carga IRL en una hoja dedicada
 */
function registrarLogSSO(id, nombre, estado, detalle) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Log_SSO");
    if (!sheet) {
      sheet = ss.insertSheet("Log_SSO");
      sheet.appendRow(["Fecha", "ID BUK", "Nombre", "Estado", "Detalle"]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#f3f3f3");
    }
    sheet.appendRow([new Date(), id, nombre, estado, detalle]);
  } catch (e) {
    console.log("Error al escribir log en hoja: " + e.message);
  }
}

/**
 * Lee la hoja "IRL" y devuelve un objeto de mapeo { "codigo": "documento.pdf" }
 */
function obtenerMapeoIRLExcel() {
  const mapping = {};
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("IRL");

    if (!sheet) {
      console.log("[IRL] Advertencia: No se encontró la hoja 'IRL'.");
      return mapping;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return mapping;

    const data = sheet.getRange("A2:B" + lastRow).getValues();
    data.forEach(([codigo, documento]) => {
      if (codigo && documento) {
        mapping[codigo.toString().trim()] = documento.toString().trim();
      }
    });

    console.log(`[IRL] Mapeo cargado desde Excel: ${Object.keys(mapping).length} cargos.`);
  } catch (e) {
    console.log("[IRL] Error leyendo hoja de configuración: " + e.message);
  }
  return mapping;
}


// ════════════════════════════════════════════════════════════════════════
//  UTILIDADES DE MANTENCIÓN (ejecutar a mano desde el editor)
// ════════════════════════════════════════════════════════════════════════

/** Muestra qué hay en la cola ahora mismo. */
function verColaWebhook() {
  var cola = PropertiesService.getScriptProperties().getProperty(COLA_PROP_) || '[]';
  console.log(cola);
  return cola;
}

/** Borra la reserva de un empleado, para poder reprocesarlo a propósito. */
function liberarJobHire(employeeId) {
  PropertiesService.getScriptProperties().deleteProperty('JOBHIRE_' + employeeId);
  console.log('Reserva liberada para ' + employeeId + '. El próximo job_hire volverá a subir documentos.');
}

/** Lista las reservas registradas. */
function verReservasJobHire() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var out = Object.keys(props)
    .filter(function (k) { return k.indexOf('JOBHIRE_') === 0; })
    .map(function (k) { return k.replace('JOBHIRE_', '') + ' → ' + new Date(Number(props[k])).toISOString(); });
  console.log(out.length ? out.join('\n') : 'Sin reservas.');
  return out;
}