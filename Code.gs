    // ================================================================
    // Code.gs — Backend DmV Control v3.0 (Versión Estable)
    // ================================================================

    var SPREADSHEET_ID = '14dnLZDpVTiBePINCT0sUefubzeJVryppqt-rTTYuUp0';
    var ROOT_FOLDER_ID = '1cTz7drePxyqsGFxNU_ymW5CnGd3yjArl';
    var LOGO_FILE_ID   = '1A0myW1PD0JIuutflbCgT_l0YbweG_UO1';

    var GLOBAL_SS_ID_OVERRIDE = null; 

    /**
    * Lógica principal para servir la interfaz y responder a la API
    */
    function doGet(e) {
    try {
        var action = e.parameter.action;
        var payload = e.parameter.payload ? JSON.parse(e.parameter.payload) : {};
        var callback = e.parameter.callback;
        GLOBAL_SS_ID_OVERRIDE = e.parameter.spreadsheetId || null;

        if (!action) {
        // Intenta cargar el archivo llamado 'Index'. Si tu archivo se llama diferente en Google, cámbialo aquí.
        return HtmlService.createTemplateFromFile('Index').evaluate()
            .setTitle('DmV Control v3.0')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
            .addMetaTag('viewport', 'width=device-width, initial-scale=1');
        }

        var result = (typeof this[action] === 'function') ? this[action](payload) : { success: false, error: 'Acción no encontrada' };
        if (callback) {
        return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
        }
        return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message })).setMimeType(ContentService.MimeType.JSON);
    }
    }

    function doPost(e) {
    try {
        var data = JSON.parse(e.postData.contents);
        GLOBAL_SS_ID_OVERRIDE = data.spreadsheetId || null;
        var result = (typeof this[data.action] === 'function') ? this[data.action](data.payload) : { success: false, error: 'Acción no encontrada' };
        return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message })).setMimeType(ContentService.MimeType.JSON);
    }
    }

    function getSheet(name) { 
    var id = GLOBAL_SS_ID_OVERRIDE || SPREADSHEET_ID;
    return SpreadsheetApp.openById(id).getSheetByName(name); 
    }

    function include(filename) {
    return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
    }

    /**
    * Función central para cargar todos los catálogos y datos iniciales.
    */
    function getAllData(activoFilter) {
    try {
        var data = {
        usuarios: getSheetData('Usuarios'),
        tramos: getSheetData('Tramos'),
        planTrabajo: getSheetData('Plan_Trabajo') || [],
        actividades: getSheetData('Actividades'),
        patrullajes: getSheetData('Patrullajes') || [],
        hallazgos: getSheetData('Hallazgos') || [],
        personalCampo: getSheetData('Personal_Campo') || [],
        inventario: getSheetData('Catalogo_Herramientas') || [],
        movimientosInventario: getSheetData('Movimientos_Inventario') || [],
        valesHerramienta: getSheetData('Herramientas_Asignacion') || [],
        incidentes: getSheetData('Incidentes') || [],
        platicas: getSheetData('Platicas_5min') || [],
        entregaEPP: getSheetData('Entrega_EPP') || [],
        actosInseguros: getSheetData('Actos_Inseguros') || [],
        config: {
            serverTime: new Date().toISOString(),
            activoFiltro: activoFilter || 'TODOS'
        }
        };
        return data;
    } catch (e) {
        Logger.log('Error en getAllData: ' + e.message);
        throw 'Error al obtener datos: ' + e.message;
    }
    }

    function getSheetData(name) {
    var sheet = getSheet(name);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var headers = data[0];
    return data.slice(1).map(function(row) {
        var obj = {};
        headers.forEach(function(h, i) {
        var val = row[i];
        // Convert Date objects to ISO strings — google.script.run
        // silently returns null if any Date object is in the response
        if (val instanceof Date) {
            obj[h] = Utilities.formatDate(val, 'America/Mexico_City', 'yyyy-MM-dd');
        } else {
            obj[h] = val;
        }
        });
        return obj;
    });
    }
    function appendRow(name, row) { getSheet(name).appendRow(row); }
    function safeJSON(val) { try { return JSON.parse(val || '[]'); } catch(e) { return []; } }
    function genId(prefix, sheetName) {
    var data = getSheetData(sheetName);
    var maxNum = 0;
    data.forEach(function(row) {
        var id = String(row.ID || '');
        var match = id.match(/(\d+)$/);
        if (match) {
        var num = parseInt(match[1], 10);
        if (num > maxNum) maxNum = num;
        }
    });
    return prefix + String(maxNum + 1).padStart(3, '0');
    }
    function logAction(usuario, accion, entidadId, detalles) {
    try {
        appendRow('Log_Acciones', [Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd HH:mm:ss'), usuario, accion, entidadId, detalles || '']);
    } catch(e) {
        Logger.log('Error logging action: ' + e.message);
    }
    }

    /**
    * Crea una nueva actividad en la hoja 'Actividades'.
    * Solo permitida para usuarios con rol ADMIN.
    */
    function crearActividad(datos) {
    // 1. Validar Sesión y Rol (SUPERVISOR o ADMIN)
    var auth = checkRole(datos.adminId, 'SUPERVISOR');
    if (!auth.allowed) return auth;

    try {
        // 2. CANDADO ANTI-DUPLICADOS: verificar que no exista actividad activa
        //    en el mismo tramo, misma fecha, mismo tipo
        var existentes = getSheetData('Actividades');
        var duplicado = existentes.find(function(a) {
        return a.Tramo === datos.tramoId
            && a.Fecha === datos.fecha
            && a.Tipo === datos.tipo
            && (a.Estado === 'Programada' || a.Estado === 'En Proceso');
        });
        if (duplicado) {
        return {
            success: false,
            error: 'Ya existe una actividad "' + datos.tipo + '" programada para el tramo ' + datos.tramoId
            + ' en fecha ' + datos.fecha + ' (Folio: ' + duplicado.ID + '). Cancele la existente o elija otra fecha/tramo.'
        };
        }

        // 3. Generar Folio
        var newId = genId('ACT-', 'Actividades');
        
        // 4. Preparar Fila
        // headers: ['ID','Tramo','Lider','Tipo','Periodicidad','Fecha','Estado','KmProgramados','ASTId','Timestamp','CreadoPor']
        var row = [
        newId,
        datos.tramoId,
        datos.liderId,
        datos.tipo,
        datos.periodicidad || 'Unica',
        datos.fecha,
        'Programada',
        datos.kmProg || 0,
        '', // ASTId (vacio al inicio)
        new Date().toISOString(),
        datos.adminId,
        datos.activoCatalogado || ''
        ];

        // 5. Insertar
        appendRow('Actividades', row);
        
        // 6. Log
        logAction(datos.adminId, 'CREAR ACTIVIDAD', newId, 'Tipo: ' + datos.tipo + ' | Tramo: ' + datos.tramoId);
        
        return { success: true, id: newId };
        
    } catch (e) {
        return { success: false, error: 'Error al crear actividad: ' + e.message };
    }
    }

    // ================================================================
    // GESTIÓN DE ESTADO DE ACTIVIDADES
    // ================================================================

    /**
    * Inicia una actividad programada — cambia estado de 'Programada' a 'En Curso'.
    * Requiere rol LIDER o superior.
    */
    function iniciarActividad(payload) {
    var auth = checkRole(payload.userId, 'LIDER');
    if (!auth.allowed) return auth;
    
    var sheet = getSheet('Actividades');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var estadoCol = headers.indexOf('Estado') + 1;
    
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
        var estadoActual = String(data[i][estadoCol - 1]).trim();
        if (estadoActual !== 'Programada' && estadoActual !== 'AST aprobado') {
            return { success: false, error: 'Solo se pueden iniciar actividades con estado Programada o AST aprobado. Estado actual: "' + estadoActual + '"' };
        }
        sheet.getRange(i + 1, estadoCol).setValue('En Curso');
        logAction(payload.userId, 'INICIAR ACTIVIDAD', payload.id, 'Programada → En Curso');
        return { success: true, id: payload.id };
        }
    }
    return { success: false, error: 'Actividad no encontrada: ' + payload.id };
    }

    /**
    * Completa una actividad — cambia estado a 'Completada'.
    * Requiere rol LIDER o superior.
    * Payload: { id, userId, observaciones? }
    */
    function completarActividad(payload) {
    var auth = checkRole(payload.userId, 'LIDER');
    if (!auth.allowed) return auth;
    
    var sheet = getSheet('Actividades');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var estadoCol = headers.indexOf('Estado') + 1;
    var colKmProg = headers.indexOf('KmProgramados') + 1;
    var colKmReal = headers.indexOf('KmRealizados') + 1;
    var colAvance = headers.indexOf('Avance') + 1;
    
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
        var estadoActual = String(data[i][estadoCol - 1]).trim();
        if (estadoActual === 'Completada') {
            return { success: false, error: 'Esta actividad ya está completada.' };
        }
        
        // Validar que el avance en km sea suficiente (>=90%)
        var kmProg = Number(data[i][colKmProg - 1]) || 0;
        var kmReal = Number(data[i][colKmReal - 1]) || 0;
        var avance = kmProg > 0 ? Math.round(kmReal / kmProg * 100) : 100;
        
        if (kmProg > 0 && avance < 90 && !payload.forzar) {
            var faltante = Math.round((kmProg - kmReal) * 100) / 100;
            return { 
            success: false, 
            error: 'No se puede completar: avance ' + avance + '% (' + kmReal + '/' + kmProg + ' km). Faltan ' + faltante + ' km. Mínimo requerido: 90%.',
            avanceActual: avance,
            kmFaltantes: faltante
            };
        }
        
        var nuevoEstado = 'Por Validar';
        // Si es Supervisor o Admin, cierra directamente. Si es Líder, va a revisión.
        if (auth.user.Rol === 'SUPERVISOR' || auth.user.Rol === 'ADMIN') {
            nuevoEstado = 'Completada';
        }
        
        // Actualizar avance final
        if (colAvance > 0) sheet.getRange(i + 1, colAvance).setValue(avance);
        
        sheet.getRange(i + 1, estadoCol).setValue(nuevoEstado);
        
        // VINCULACIÓN CON PLAN DE TRABAJO (SI SE CIERRA)
        if (nuevoEstado === 'Completada') {
            try {
            var tramo = data[i][headers.indexOf('Tramo')];
            var tipo = data[i][headers.indexOf('Tipo')];
            var kmDone = Number(data[i][colKmReal - 1]) || 0;
            if (kmDone > 0) sumarAvanceAlPlan(tramo, tipo, kmDone);
            } catch(e) { Logger.log('Error vinculando plan: ' + e.message); }
        }

        logAction(payload.userId, 'COMPLETAR ACTIVIDAD', payload.id, 'Estado: ' + nuevoEstado + ' | Avance: ' + avance + '% | Km: ' + kmReal + '/' + kmProg + ' | Obs: ' + (payload.observaciones || ''));
        return { success: true, id: payload.id, avance: avance };
        }
    }
    return { success: false, error: 'Actividad no encontrada: ' + payload.id };
    }

    // ================================================================
    // AVANCE POR FASES — Tracking de progreso en campo
    // Fases: Traslado -> Preparacion -> Ejecucion -> Cierre
    // ================================================================
    function actualizarAvance(payload) {
    var auth = checkRole(payload.userId, 'LIDER');
    if (!auth.allowed) return auth;
    var sheet = getSheet('Actividades');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var colKmReal = headers.indexOf('KmRealizados') + 1;
    var colAvance = headers.indexOf('Avance') + 1;
    var colFase   = headers.indexOf('Fase') + 1;
    var colKmProg = headers.indexOf('KmProgramados') + 1;
    var colEstado = headers.indexOf('Estado') + 1;
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
        var kmProg = Number(data[i][colKmProg - 1]) || 1;
        var kmReal = Number(payload.kmRealizados) || 0;
        var avance = Math.min(100, Math.round(kmReal / kmProg * 100));
        var fase = payload.fase || 'Ejecucion';
        if (colKmReal > 0) sheet.getRange(i + 1, colKmReal).setValue(kmReal);
        if (colAvance > 0) sheet.getRange(i + 1, colAvance).setValue(avance);
        if (colFase > 0)   sheet.getRange(i + 1, colFase).setValue(fase);
        // FIX: Aceptar tanto 'Programada' como 'AST aprobado' para transición a 'En Proceso'
        var estadoActual = String(data[i][colEstado - 1]).trim();
        if (colEstado > 0 && (estadoActual === 'Programada' || estadoActual === 'AST aprobado')) {
            sheet.getRange(i + 1, colEstado).setValue('En Proceso');
        }
        logAction(payload.userId, 'AVANCE', payload.id, 'Km:' + kmReal + '/' + kmProg + ' (' + avance + '%) Fase:' + fase);
        return { success: true, avance: avance, fase: fase };
        }
    }
    return { success: false, error: 'Actividad no encontrada' };
    }


    // ================================================================
    // SUSPENDER ACTIVIDAD — Clima, fuerza mayor, etc.
    // Genera folio extraordinario para reporte al cliente
    // ================================================================
    function suspenderActividad(payload) {
    var auth = checkRole(payload.userId, 'SUPERVISOR');
    if (!auth.allowed) return auth;
    var sheet = getSheet('Actividades');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var colEstado = headers.indexOf('Estado') + 1;
    var actInfo = null;
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
        var est = String(data[i][colEstado - 1]);
        if (est === 'Completada' || est === 'Cancelada') {
            return { success: false, error: 'No se puede suspender: estado ' + est };
        }
        sheet.getRange(i + 1, colEstado).setValue('Suspendida');
        actInfo = {};
        headers.forEach(function(h, idx) { actInfo[h] = data[i][idx]; });
        break;
        }
    }
    if (!actInfo) return { success: false, error: 'Actividad no encontrada' };
    var folioExt = genId('REX-', 'Reportes_Extraordinarios');
    var hoy = Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd');
    var hora = Utilities.formatDate(new Date(), 'America/Mexico_City', 'HH:mm');
    appendRow('Reportes_Extraordinarios', [
        folioExt, hoy, hora, payload.id, actInfo.Tipo || '', actInfo.Tramo || '',
        payload.motivo || 'Clima', payload.descripcion || '', payload.userId,
        'Abierto', '', '', payload.evidenciaFoto || '', '', new Date().toISOString()
    ]);
    logAction(payload.userId, 'SUSPENSION', payload.id, 'Motivo:' + (payload.motivo||'Clima') + ' Folio:' + folioExt);
    return { success: true, folioExtraordinario: folioExt,
        mensaje: 'Actividad suspendida. Folio extraordinario: ' + folioExt };
    }

    function reprogramarActividad(payload) {
    var auth = checkRole(payload.userId, 'SUPERVISOR');
    if (!auth.allowed) return auth;
    var sheetRex = getSheet('Reportes_Extraordinarios');
    var dataRex = sheetRex.getDataRange().getValues();
    var hRex = dataRex[0];
    var cEst = hRex.indexOf('Estado') + 1;
    var cFR = hRex.indexOf('FechaReprogramacion') + 1;
    var cAR = hRex.indexOf('ActividadReprogramadaId') + 1;
    var rexInfo = null, rexRow = -1;
    for (var i = 1; i < dataRex.length; i++) {
        if (dataRex[i][0] === payload.folioExt) {
        rexInfo = {}; hRex.forEach(function(h, idx) { rexInfo[h] = dataRex[i][idx]; });
        rexRow = i + 1; break;
        }
    }
    if (!rexInfo) return { success: false, error: 'Folio extraordinario no encontrado' };
    var nueva = crearActividad({
        adminId: payload.userId, tramoId: rexInfo.Tramo, liderId: payload.liderId || '',
        tipo: rexInfo.TipoActividad, fecha: payload.nuevaFecha, periodicidad: 'Reprogramada'
    });
    if (nueva.success) {
        sheetRex.getRange(rexRow, cEst).setValue('Reprogramado');
        sheetRex.getRange(rexRow, cFR).setValue(payload.nuevaFecha);
        sheetRex.getRange(rexRow, cAR).setValue(nueva.id);
        logAction(payload.userId, 'REPROGRAMAR', payload.folioExt, 'Nueva:' + nueva.id);
    }
    return nueva;
    }

    function getReportesExtraordinarios() { return getSheetData('Reportes_Extraordinarios'); }

    // ================================================================
    // PLAN DE TRABAJO — Alcance contractual por tramo
    // ================================================================
    function getPlanTrabajo(tramoId) {
    var data = getSheetData('Plan_Trabajo');
    if (tramoId) return data.filter(function(p) { return p.Tramo === tramoId; });
    return data;
    }

    function guardarPartidaPlan(payload) {
    var auth = checkRole(payload.userId, 'SUPERVISOR');
    if (!auth.allowed) return auth;
    var id = genId('PT', 'Plan_Trabajo');
    appendRow('Plan_Trabajo', [
        id, 
        payload.tramo, 
        payload.partida || payload.tipoTrabajo || 'S/N', 
        payload.descripcion || '',
        payload.unidad || payload.unidadMedida || 'km', 
        Number(payload.cantidadProgramada) || 0,
        0, // CantidadEjecutada
        0, // Avance
        payload.fechaInicio || '', 
        payload.fechaFin || '',
        'Pendiente', 
        payload.fase || 'Fase 1', 
        payload.observaciones || '',
        new Date().toISOString()
    ]);
    logAction(payload.userId, 'CREAR PARTIDA PLAN', id, (payload.partida || payload.tipoTrabajo) + ' ' + payload.tramo);
    return { success: true, id: id };
    }

    function actualizarAvancePlan(payload) {
    // payload: { id, userId, cantidadEjecutada, observaciones }
    var auth = checkRole(payload.userId, 'LIDER');
    if (!auth.allowed) return auth;
    var sheet = getSheet('Plan_Trabajo');
    var data = sheet.getDataRange().getValues();
    var h = data[0];
    var cEjec = h.indexOf('CantidadEjecutada') + 1;
    var cAvance = h.indexOf('Avance') + 1;
    var cEstado = h.indexOf('Estado') + 1;
    var cProg = h.indexOf('CantidadProgramada') + 1;
    var cObs = h.indexOf('Observaciones') + 1;
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
        var prog = Number(data[i][cProg - 1]) || 1;
        var ejec = Number(payload.cantidadEjecutada) || 0;
        var avance = Math.min(100, Math.round(ejec / prog * 100));
        sheet.getRange(i + 1, cEjec).setValue(ejec);
        sheet.getRange(i + 1, cAvance).setValue(avance);
        // Auto estado
        var estado = avance >= 100 ? 'Completado' : avance > 0 ? 'En Proceso' : 'Pendiente';
        sheet.getRange(i + 1, cEstado).setValue(estado);
        if (payload.observaciones && cObs > 0) sheet.getRange(i + 1, cObs).setValue(payload.observaciones);
        logAction(payload.userId, 'AVANCE PLAN', payload.id, ejec + '/' + prog + ' ' + data[i][h.indexOf('UnidadMedida')] + ' (' + avance + '%)');
        return { success: true, avance: avance, estado: estado };
        }
    }
    return { success: false, error: 'Partida no encontrada' };
    }

    function getResumenPlanTrabajo() {
    var plan = getSheetData('Plan_Trabajo');
    var tramos = {};
    plan.forEach(function(p) {
        var t = p.Tramo || '?';
        if (!tramos[t]) tramos[t] = { partidas: 0, completadas: 0, enProceso: 0, pendientes: 0, items: [] };
        tramos[t].partidas++;
        if (p.Estado === 'Completado') tramos[t].completadas++;
        else if (p.Estado === 'En Proceso') tramos[t].enProceso++;
        else tramos[t].pendientes++;
        tramos[t].items.push({
        id: p.ID, tipo: p.TipoTrabajo, unidad: p.UnidadMedida,
        programado: Number(p.CantidadProgramada) || 0,
        ejecutado: Number(p.CantidadEjecutada) || 0,
        avance: Number(p.Avance) || 0, estado: p.Estado, fase: p.Fase
        });
    });
    // Calcular avance general por tramo
    Object.keys(tramos).forEach(function(t) {
        var total = tramos[t].partidas;
        tramos[t].pctGeneral = total > 0 ? Math.round(tramos[t].completadas / total * 100) : 0;
    });
    return tramos;
    }

    // ================================================================
    // KPIs MENSUALES — Cumplimiento de metas
    // ================================================================
    function calcularKPIsMensuales(mesAnio) {
    if (!mesAnio) mesAnio = Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM');
    var acts = getSheetData('Actividades');
    var del = acts.filter(function(a) { return String(a.Fecha||'').substring(0,7) === mesAnio; });
    var total = del.length;
    var comp = del.filter(function(a){return a.Estado==='Completada';}).length;
    var susp = del.filter(function(a){return a.Estado==='Suspendida';}).length;
    var kmP = del.reduce(function(s,a){return s+(Number(a.KmProgramados)||0);},0);
    var kmR = del.reduce(function(s,a){return s+(Number(a.KmRealizados)||0);},0);
    var porTramo = {}, porTipo = {};
    del.forEach(function(a) {
        var t = a.Tramo||'?'; if(!porTramo[t]) porTramo[t]={total:0,comp:0,susp:0,kmP:0,kmR:0};
        porTramo[t].total++; if(a.Estado==='Completada') porTramo[t].comp++;
        if(a.Estado==='Suspendida') porTramo[t].susp++;
        porTramo[t].kmP+=Number(a.KmProgramados)||0; porTramo[t].kmR+=Number(a.KmRealizados)||0;
    });
    del.forEach(function(a) {
        var t = a.Tipo||'?'; if(!porTipo[t]) porTipo[t]={total:0,comp:0,susp:0};
        porTipo[t].total++; if(a.Estado==='Completada') porTipo[t].comp++;
        if(a.Estado==='Suspendida') porTipo[t].susp++;
    });
    var base = total - susp; // No penalizar suspendidas en cumplimiento
    var res = {
        mes: mesAnio,
        resumen: { total:total, completadas:comp, suspendidas:susp,
        pctCumplimiento: base>0 ? Math.round(comp/base*100) : 0,
        kmProgramados:kmP, kmRealizados:kmR,
        pctCobertura: kmP>0 ? Math.round(kmR/kmP*100) : 0 },
        porTramo: porTramo, porTipo: porTipo
    };
    appendRow('KPIs', [genId('KPI-','KPIs'),
        Utilities.formatDate(new Date(),'America/Mexico_City','yyyy-MM-dd'),
        'Mensual', JSON.stringify(res.resumen), JSON.stringify({porTramo:porTramo,porTipo:porTipo}),
        res.resumen.pctCumplimiento+'%', new Date().toISOString()
    ]);
    return res;
    }

    /**
    * Guarda la asistencia y el flag de "Hay Segurista".
    * Crea/Usa la hoja 'Asistencias'.
    */
    function guardarAsistencia(payload) {
    var auth = checkRole(payload.userId, 'LIDER');
    if (!auth.allowed) return { success: false, error: auth.error };
    // payload: { actividadId, asistentes: [], haySegurista: bool, userId }
    var sheet = getSheet('Asistencias');
    if (!sheet) {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        sheet = ss.insertSheet('Asistencias');
        sheet.appendRow(['ID', 'ActividadID', 'Lider', 'HaySegurista', 'AsistentesJSON', 'Timestamp']);
    }
    
    var id = genId('ASIST-', 'Asistencias');
    sheet.appendRow([
        id, 
        payload.actividadId, 
        payload.userId, 
        payload.haySegurista, 
        JSON.stringify(payload.asistentes || []), 
        new Date().toISOString()
    ]);
    
    return { success: true, id: id };
    }

    /**
    * Valida el cierre de una actividad (Supervisor).
    * Cambia de 'Por Validar' a 'Completada'.
    */
    function validarCierreSupervisor(payload) {
    var auth = checkRole(payload.userId, 'SUPERVISOR');
    if (!auth.allowed) return auth;

    var sheet = getSheet('Actividades');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var estadoCol = headers.indexOf('Estado') + 1;
    var tramoCol = headers.indexOf('Tramo');
    var tipoCol = headers.indexOf('Tipo');
    var kmRealCol = headers.indexOf('KmRealizados');

    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === payload.actividadId) {
        if (data[i][estadoCol - 1] === 'Completada') return { success: true, mensaje: 'Ya estaba completada' };
        
        var tramo = data[i][tramoCol];
        var tipo = data[i][tipoCol];
        var kmDone = Number(data[i][kmRealCol]) || 0;

        sheet.getRange(i + 1, estadoCol).setValue('Completada');
        
        // VINCULACIÓN AUTOMÁTICA CON PLAN DE TRABAJO
        if (kmDone > 0) {
            try { sumarAvanceAlPlan(tramo, tipo, kmDone); } 
            catch(e) { Logger.log('Error vincular plan: ' + e.message); }
        }

        logAction(payload.userId, 'VALIDAR CIERRE', payload.actividadId, payload.comentarios || '');
        return { success: true };
        }
    }
    return { success: false, error: 'Actividad no encontrada' };
    }

    /**
    * Helper para sumar avance al plan contractual
    */
    function sumarAvanceAlPlan(tramo, tipo, cantidad) {
    var sheet = getSheet('Plan_Trabajo');
    if (!sheet) return false;
    var data = sheet.getDataRange().getValues();
    var h = data[0];
    var cTramo = h.indexOf('Tramo');
    var cTipo = h.indexOf('TipoTrabajo');
    var cEjec = h.indexOf('CantidadEjecutada') + 1;
    var cProg = h.indexOf('CantidadProgramada') + 1;
    var cAvance = h.indexOf('Avance') + 1;
    var cEstado = h.indexOf('Estado') + 1;

    for (var i = 1; i < data.length; i++) {
        // Coincidencia por Tramo y Tipo de Trabajo (exacta o contenida)
        if (data[i][cTramo] === tramo && (data[i][cTipo] === tipo || tipo.indexOf(data[i][cTipo]) >= 0)) {
        var actualEjec = Number(data[i][cEjec - 1]) || 0;
        var prog = Number(data[i][cProg - 1]) || 1;
        var nuevoEjec = actualEjec + cantidad;
        var nuevoAvance = Math.min(100, Math.round(nuevoEjec / prog * 100));
        var nuevoEstado = nuevoAvance >= 100 ? 'Completado' : 'En Proceso';
        
        sheet.getRange(i + 1, cEjec).setValue(nuevoEjec);
        sheet.getRange(i + 1, cAvance).setValue(nuevoAvance);
        sheet.getRange(i + 1, cEstado).setValue(nuevoEstado);
        return true;
        }
    }
    return false;
    }

    // ================================================================
    // SUBIR EVIDENCIA INDEPENDIENTE
    // ================================================================

    /**
    * Sube una evidencia fotográfica independiente a la hoja 'Evidencias' y a Drive.
    * Puede vincularse a cualquier entidad (patrullaje, hallazgo, inspección, etc.)
    * Payload: { tipo, entidadId, foto (base64), descripcion?, userId }
    */
    function subirEvidencia(payload) {
    var auth = checkRole(payload.userId, 'LIDER');
    if (!auth.allowed) return auth;
    
    try {
        var id = genId('EV-', 'Evidencias');
        var fileUrl = '';
        var fileId = '';
        
        if (payload.foto) {
        var folder;
        try {
            folder = DriveApp.getFolderById(ROOT_FOLDER_ID).getFoldersByName('Evidencias').next();
        } catch(e) {
            folder = DriveApp.getFolderById(ROOT_FOLDER_ID).createFolder('Evidencias');
        }
        // FIX: Fallback si no tiene prefijo data:
        var b64data = payload.foto.indexOf(',') >= 0 ? payload.foto.split(',')[1] : payload.foto;
        var mimeType = 'image/jpeg';
        try { mimeType = payload.foto.split(';')[0].split(':')[1] || 'image/jpeg'; } catch(me) {}
        var ext = mimeType === 'image/png' ? '.png' : '.jpg';
        var blob = Utilities.newBlob(Utilities.base64Decode(b64data), mimeType, id + ext);
        var file = folder.createFile(blob);
        fileUrl = file.getUrl();
        fileId = file.getId();
        }
        
        appendRow('Evidencias', [
        id,
        payload.entidadId || '',
        payload.tipo || 'General',
        fileUrl,
        fileId,
        new Date().toISOString()
        ]);
        
        logAction(payload.userId, 'SUBIR EVIDENCIA', id, 'Tipo: ' + (payload.tipo || 'General') + ' | Entidad: ' + (payload.entidadId || 'N/A'));
        
        return { success: true, id: id, url: fileUrl, fileId: fileId };
    } catch(e) {
        return { success: false, error: 'Error al subir evidencia: ' + e.message };
    }
    }


    // ================================================================
    // getImgBase64 — CORE para evidencias en PDFs
    // ================================================================
    // Convierte archivo de Drive a base64 para insertar en PDFs/HTML
    // Llamado desde frontend: google.script.run.getImgBase64(fileId)
    // Llamado desde backend: getImgBase64(fileId) directo
    function getImgBase64(fileId) {
    if (!fileId) return null;
    try {
        var file = DriveApp.getFileById(fileId);
        var blob = file.getBlob();
        var bytes = blob.getBytes();
        return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(bytes);
    } catch (e) {
        Logger.log('Error getImgBase64 (' + fileId + '): ' + e.message);
        return null;
    }
    }

    // Batch version - multiple images
    function getImgsBase64(fileIds) {
    return fileIds.map(function(id) { return getImgBase64(id); });
    }

    // ================================================================
    // RBAC — Control de Acceso por Rol
    // ================================================================
    // Verifica que el usuario tenga el rol requerido antes de ejecutar
    function checkRole(userId, requiredRole) {
    var usuarios = getSheetData('Usuarios');
    var user = usuarios.find(function(u) { return u.ID === userId; });
    if (!user) return { allowed: false, error: 'Usuario no encontrado' };
    if (user.Rol === 'ADMIN') return { allowed: true, user: user }; // Acceso Total (God Mode)
    if (!(user.Activo === true || String(user.Activo).toUpperCase() === 'TRUE')) return { allowed: false, error: 'Usuario inactivo' };
    // Role hierarchy: ADMIN(4) > SUPERVISOR(3) > SEGURISTA(2) > LIDER(1)
    var ROLE_LEVEL = { 'ADMIN': 4, 'SUPERVISOR': 3, 'SEGURISTA': 2, 'LIDER': 1, 'CONSULTA': 0 };
    var userLevel = ROLE_LEVEL[user.Rol] || 0;
    var requiredLevel = ROLE_LEVEL[requiredRole] || 0;
    if (userLevel < requiredLevel) {
        logAction(userId, 'ACCESO DENEGADO', requiredRole, 'Rol: ' + user.Rol + ' < ' + requiredRole);
        return { allowed: false, error: 'Acceso denegado. Se requiere rol ' + requiredRole + ' o superior.' };
    }
    return { allowed: true, user: user };
    }

    // Verifica si el líder puede operar (PSST blocks)
    function checkOperability(userId) {
    var caps = getCapacitaciones(userId);
    var capVencida = caps.some(function(c) { return c.Estatus === 'Vencida'; });
    var apts = getAptitudesMedicas(userId);
    var aptOk = apts.length > 0 && apts.every(function(a) { return a.Estatus === 'Vigente'; });
    var swActivo = getStopWork().some(function(s) { return s.Estado === 'Activo'; });
    var blocks = [];
    if (capVencida) blocks.push('Capacitacion vencida');
    if (!aptOk) blocks.push('Sin aptitud medica vigente');
    if (swActivo) blocks.push('Stop Work activo');
    return { canOperate: blocks.length === 0, blocks: blocks };
    }

    // ================================================================
    // AUTENTICACIÓN
    // ================================================================
    function login(pin) {
    var usuarios = getSheetData('Usuarios');
    var user = usuarios.find(function(u) { return String(u.PIN) === String(pin) && (u.Activo === true || String(u.Activo).toUpperCase() === 'TRUE'); });
    if (!user) return { success: false, error: 'PIN incorrecto' };
    logAction(user.ID, 'Login', '', 'Rol: ' + user.Rol);
    return {
        success: true,
        usuario: {
        id: user.ID,
        nombre: user.Nombre,
        rol: user.Rol,
        iniciales: user.Iniciales,
        puesto: user.Puesto || '',
        activoAsignado: user.ActivoAsignado || '',
        numeroEco: user.NumeroEco || ''
        }
    };
    }

    function crearTramo(p) {
    var auth = checkRole(p.creadoPor, 'ADMIN');
    if (!auth.allowed) return { success: false, error: auth.error };
    try {
        var sheet = getSheet('Tramos');
        // FIX: Usar genId() consistente con el resto del sistema
        var id = genId('T', 'Tramos');
        
        var docUrl = '';
        
        if (p.documento && p.documento.indexOf('base64,') > -1) {
        var folder = DriveApp.getFolderById(ROOT_FOLDER_ID);
        var iter = folder.getFoldersByName('Tramos_Docs');
        var docFolder = iter.hasNext() ? iter.next() : folder.createFolder('Tramos_Docs');
        
        var parts = p.documento.split(',');
        var contentType = parts[0].split(':')[1].split(';')[0];
        var data = Utilities.base64Decode(parts[1]);
        var blob = Utilities.newBlob(data, contentType, 'Doc_Tramo_' + id + '_' + p.nombre);
        var file = docFolder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        docUrl = file.getUrl();
        } else {
        docUrl = p.documento || ''; 
        }

        sheet.appendRow([id, p.nombre, p.zona, p.kmInicio, p.kmFin, true, new Date().toISOString(), docUrl, p.creadoPor]);
        
        return { success: true, id: id, url: docUrl };
    } catch (e) {
        logAction(p.creadoPor || 'System', 'Error Crear Tramo', '', e.message);
        return { success: false, error: e.message };
    }
    }


    // ================================================================
    // LECTURAS
    // ================================================================
    function getDashboard() {
    var acts = getSheetData('Actividades');
    var hall = getSheetData('Hallazgos');
    var sas = getSheetData('SASISOPA');
    var sw = getSheetData('Stop_Work');
    var tramos = getSheetData('Tramos');
    var kpis = getSheetData('KPIs');
    var lastKpi = kpis.length ? kpis[kpis.length - 1] : {};
    
    var totalKm = acts.reduce(function(s, a) { return s + Number(a.KmProgramados || 0); }, 0);
    var compActs = acts.filter(function(a) { return a.Estado === 'Completada'; });
    var kmEjec = compActs.reduce(function(s, a) { return s + Number(a.KmProgramados || 0); }, 0);
    
    return {
        totalKmProg: totalKm, kmEjec: kmEjec,
        pctAvance: totalKm ? Math.round(kmEjec / totalKm * 100) : 0,
        hallAbiertos: hall.filter(function(h) { return h.Estado === 'Abierto'; }).length,
        swActivos: sw.filter(function(s) { return s.Estado === 'Activo'; }).length,
        sasisopa: sas.map(function(s) { return { elem: s.Elemento, cum: Number(s.Cumplimiento) }; }),
        kpis: lastKpi
    };
    }
    function getActividades(liderId) {
    var data = getSheetData('Actividades');
    if (liderId) data = data.filter(function(a) { return a.Lider === liderId; });
    return data;
    }
    function getTramos() { return getSheetData('Tramos'); }
    function getCatalogoActivos() { return getSheetData('Catalogo_Activos'); }
    function getHallazgos(filtro) {
    var data = getSheetData('Hallazgos');
    if (filtro && filtro !== 'todos') data = data.filter(function(h) { return h.Estado === filtro; });
    return data;
    }
    function getASTs() { return getSheetData('AST'); }
    function getStopWork() { return getSheetData('Stop_Work'); }
    function getSASISOPA() { return getSheetData('SASISOPA'); }
    function getNOMMatriz() { return getSheetData('NOM_Matriz'); }

    // ================================================================
    // MÓDULO: EVALUACIÓN AUTOMÁTICA DE CUMPLIMIENTO NOM
    // Cruza datos de todos los módulos para calcular % por NOM
    // ================================================================
    function evaluarCumplimientoNOM() {
    try {

    var patrullajes = getSheetData('Patrullajes');
    var inspecciones = getSheetData('Inspecciones');
    var permisos = getSheetData('Permisos_Trabajo');
    var capacitaciones = getSheetData('Capacitaciones');
    var aptitudes = getSheetData('Aptitudes_Medicas');
    var entregaEPP = getSheetData('Entrega_EPP');
    var programaSalud = getSheetData('Programa_Salud');
    var platicas = getSheetData('Platicas_5min');
    var emergencias = getSheetData('Emergencias');
    var incidentes = getSheetData('Incidentes');
    var listadoMaestro = getSheetData('Listado_Maestro');
    var ast = getSheetData('AST');
    var hallazgos = getSheetData('Hallazgos');
    var stopWork = getSheetData('Stop_Work');
    
    var docsVigentes = listadoMaestro.filter(function(d) { return d.Estado === 'Vigente'; });
    var capVigentes = capacitaciones.filter(function(c) { return c.Estatus === 'Vigente'; });
    var aptVigentes = aptitudes.filter(function(a) { return a.Estatus === 'Vigente'; });
    var inspEPP = inspecciones.filter(function(i) { return i.Tipo === 'EPP'; });
    var inspVehiculos = inspecciones.filter(function(i) { return i.Tipo === 'Utilitario' || i.Tipo === 'MaqPesada'; });
    var inspHerr = inspecciones.filter(function(i) { return i.Tipo === 'Herramienta'; });
    var ptAltura = permisos.filter(function(p) { return p.TipoPT === 'Altura'; });
    var ptElectrico = permisos.filter(function(p) { return p.TipoPT === 'Electrico'; });
    var ptEspConf = permisos.filter(function(p) { return p.TipoPT === 'EspacioConfinado'; });
    var ptCaliente = permisos.filter(function(p) { return p.TipoPT === 'Caliente'; });
    var saludRealizadas = programaSalud.filter(function(s) { return s.Estatus === 'Realizado'; });
    var simulacros = emergencias.filter(function(e) { return e.Tipo === 'Simulacro'; });
    
    function docExists(codigo) {
        return docsVigentes.some(function(d) { return d.Codigo && d.Codigo.indexOf(codigo) >= 0; });
    }
    
    function calcPct(checks) {
        var total = checks.length;
        var passed = checks.filter(function(c) { return c; }).length;
        return total > 0 ? Math.round(passed / total * 100) : 0;
    }
    
    var evaluacion = [
        {
        nom: 'NOM-020-ASEA-2016',
        titulo: 'Transporte de gas natural por ducto — DmV',
        modulo: 'Patrullajes, Hallazgos, Actividades',
        evidenciaRequerida: 'Registros de patrullaje, hallazgos, evidencia fotográfica, AST',
        checks: [
            patrullajes.length > 0,
            hallazgos.length > 0 || ast.length > 0, // Tiene registros de hallazgos O ASTs
            ast.length > 0,
            docExists('SAS-07-001'),
            docExists('SAS-02-002')
        ],
        evidenciaDigital: patrullajes.length + ' patrullajes, ' + hallazgos.length + ' hallazgos, ' + ast.length + ' ASTs'
        },
        {
        nom: 'NOM-001-STPS-2008',
        titulo: 'Condiciones de seguridad — Edificios y áreas',
        modulo: 'Inspecciones, Hallazgos',
        evidenciaRequerida: 'Inspecciones periódicas de instalaciones, control de hallazgos',
        checks: [
            inspecciones.length > 0,
            hallazgos.filter(function(h) { return h.Estado === 'Cerrado'; }).length > 0 || hallazgos.length === 0, // Tiene hallazgos cerrados O no hay hallazgos pendientes
            docExists('SAS-10-001')
        ],
        evidenciaDigital: inspecciones.length + ' inspecciones, hallazgos: ' + hallazgos.filter(function(h) { return h.Estado === 'Abierto'; }).length + ' abiertos'
        },
        {
        nom: 'NOM-002-STPS-2010',
        titulo: 'Prevención y protección contra incendios',
        modulo: 'Inspecciones, Emergencias, Capacitaciones',
        evidenciaRequerida: 'Inspección extintores, simulacros de incendio, brigada capacitada',
        checks: [
            simulacros.length > 0,
            docExists('SAS-09-001'),
            docExists('SAS-09-002'),
            capVigentes.length > 0
        ],
        evidenciaDigital: simulacros.length + ' simulacros, ' + capVigentes.length + ' capacitaciones vigentes'
        },
        {
        nom: 'NOM-004-STPS-1999',
        titulo: 'Maquinaria y equipo — Sistemas de protección',
        modulo: 'Inspecciones (MaqPesada), Permisos de Trabajo',
        evidenciaRequerida: 'Checklist maquinaria, operador calificado, protecciones',
        checks: [
            inspVehiculos.length > 0,
            docExists('SAS-05-003')
        ],
        evidenciaDigital: inspVehiculos.length + ' inspecciones de vehículos/maquinaria'
        },
        {
        nom: 'NOM-005-STPS-1998',
        titulo: 'Manejo de sustancias químicas peligrosas',
        modulo: 'Permisos de Trabajo, Capacitaciones',
        evidenciaRequerida: 'HDS disponibles, capacitación en manejo, kit derrames',
        checks: [
            docExists('SAS-16-002'),
            capVigentes.length > 0
        ],
        evidenciaDigital: 'Permisos de trabajo caliente: ' + ptCaliente.length
        },
        {
        nom: 'NOM-006-STPS-2014',
        titulo: 'Manejo y almacenamiento de materiales',
        modulo: 'Inspecciones (Herramienta), Capacitaciones',
        evidenciaRequerida: 'Inspección de herramienta, procedimientos de carga',
        checks: [
            inspHerr.length > 0,
            docExists('SAS-05-003')
        ],
        evidenciaDigital: inspHerr.length + ' inspecciones de herramienta'
        },
        {
        nom: 'NOM-009-STPS-2011',
        titulo: 'Trabajos en altura',
        modulo: 'Permisos de Trabajo, AST, Entrega EPP',
        evidenciaRequerida: 'PT altura firmado, AST, arnés inspeccionado, DC3 operador',
        checks: [
            ptAltura.length > 0 || permisos.length > 0, // Tiene PT altura O tiene sistema de permisos activo
            docExists('SAS-16-002'),
            entregaEPP.length > 0
        ],
        evidenciaDigital: ptAltura.length + ' permisos de altura, ' + entregaEPP.length + ' entregas EPP'
        },
        {
        nom: 'NOM-011-STPS-2001',
        titulo: 'Ruido — Condiciones de seguridad e higiene',
        modulo: 'Programa Salud, Entrega EPP',
        evidenciaRequerida: 'Audiometría, protección auditiva entregada, monitoreo ruido',
        checks: [
            saludRealizadas.some(function(s) { return s.Actividad && s.Actividad.indexOf('Audiometría') >= 0; }) || programaSalud.some(function(s) { return s.Actividad && s.Actividad.indexOf('Audiometría') >= 0; }),
            docExists('SAS-16-005')
        ],
        evidenciaDigital: 'Programa salud: ' + programaSalud.length + ' actividades'
        },
        {
        nom: 'NOM-017-STPS-2008',
        titulo: 'Equipo de protección personal',
        modulo: 'Entrega EPP, Inspecciones EPP',
        evidenciaRequerida: 'Análisis de riesgo para EPP, entrega con firma, inspección periódica',
        checks: [
            entregaEPP.length > 0,
            inspEPP.length > 0,
            docExists('SAS-16-003'),
            docExists('SAS-16-004')
        ],
        evidenciaDigital: entregaEPP.length + ' entregas, ' + inspEPP.length + ' inspecciones EPP'
        },
        {
        nom: 'NOM-019-STPS-2011',
        titulo: 'Comisiones de seguridad e higiene',
        modulo: 'Pláticas 5min, Capacitaciones',
        evidenciaRequerida: 'Acta constitutiva de comisión, recorridos de verificación, actas',
        checks: [
            platicas.length > 0,
            docExists('SAS-18-001')
        ],
        evidenciaDigital: platicas.length + ' pláticas de seguridad'
        },
        {
        nom: 'NOM-025-STPS-2008',
        titulo: 'Iluminación — Condiciones en centros de trabajo',
        modulo: 'Inspecciones',
        evidenciaRequerida: 'Estudio de iluminación (si aplica en bases operativas)',
        checks: [
            inspecciones.length > 0
        ],
        evidenciaDigital: 'Inspecciones generales: ' + inspecciones.length
        },
        {
        nom: 'NOM-026-STPS-2008',
        titulo: 'Señalización de seguridad e higiene',
        modulo: 'Inspecciones, Listado Maestro',
        evidenciaRequerida: 'Inventario señalización, inspección periódica, conformidad con NOM',
        checks: [
            docExists('SAS-17-001'),
            docExists('SAS-17-002')
        ],
        evidenciaDigital: 'Docs señalización: ' + (docExists('SAS-17-001') ? '✓' : '✗') + ' SAS-17-001, ' + (docExists('SAS-17-002') ? '✓' : '✗') + ' SAS-17-002'
        },
        {
        nom: 'NOM-029-STPS-2011',
        titulo: 'Mantenimiento de instalaciones eléctricas',
        modulo: 'Permisos de Trabajo, Capacitaciones',
        evidenciaRequerida: 'PT eléctrico, LOTO, personal calificado DC3',
        checks: [
            ptElectrico.length > 0 || permisos.length > 0, // Tiene PT eléctrico O sistema de permisos activo
            docExists('SAS-16-002')
        ],
        evidenciaDigital: ptElectrico.length + ' permisos eléctricos'
        },
        {
        nom: 'NOM-030-STPS-2009',
        titulo: 'Servicios preventivos de seguridad y salud',
        modulo: 'Programa Salud, Aptitudes Médicas, Capacitaciones',
        evidenciaRequerida: 'Programa de salud, EMOs, vigilancia epidemiológica, diagnóstico',
        checks: [
            programaSalud.length > 0,
            aptVigentes.length > 0,
            saludRealizadas.length > 0 || programaSalud.length > 0,
            docExists('SAS-16-005'),
            docExists('SAS-04-002')
        ],
        evidenciaDigital: programaSalud.length + ' actividades salud, ' + aptVigentes.length + ' aptitudes vigentes'
        },
        {
        nom: 'NOM-033-STPS-2015',
        titulo: 'Espacios confinados',
        modulo: 'Permisos de Trabajo, AST, Capacitaciones',
        evidenciaRequerida: 'PT espacio confinado, monitoreo atmósfera, rescate, DC3',
        checks: [
            ptEspConf.length > 0 || permisos.length > 0, // Tiene PT espacio confinado O sistema de permisos activo
            docExists('SAS-16-002')
        ],
        evidenciaDigital: ptEspConf.length + ' permisos espacio confinado'
        }
    ];
    
    var hoy = Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd');
    
    // Actualizar hoja NOM_Matriz
    var sheet = getSheet('NOM_Matriz');
    if (sheet) {
        // Limpiar datos existentes
        var lastRow = sheet.getLastRow();
        if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
        
        evaluacion.forEach(function(ev) {
        var cumplimiento = calcPct(ev.checks);
        var estado = cumplimiento >= 85 ? 'Cumple' : cumplimiento >= 50 ? 'Parcial' : 'No Cumple';
        sheet.appendRow([
            ev.nom,
            ev.titulo,
            ev.modulo,
            ev.evidenciaRequerida,
            ev.evidenciaDigital,
            cumplimiento,
            estado,
            hoy
        ]);
        });
    }
    
    // Calcular promedio general
    var promedioGeneral = Math.round(evaluacion.reduce(function(s, ev) {
        return s + calcPct(ev.checks);
    }, 0) / evaluacion.length);
    
    logAction('SISTEMA', 'Evaluación NOM automática', 'NOM_Matriz', promedioGeneral + '% cumplimiento promedio');
    
    return {
        success: true,
        evaluacion: evaluacion.map(function(ev) {
        var pct = calcPct(ev.checks);
        return {
            nom: ev.nom,
            titulo: ev.titulo,
            modulo: ev.modulo,
            evidenciaRequerida: ev.evidenciaRequerida,
            evidenciaDigital: ev.evidenciaDigital,
            cumplimiento: pct,
            estado: pct >= 85 ? 'Cumple' : pct >= 50 ? 'Parcial' : 'No Cumple',
            fecha: hoy
        };
        }),
        promedioGeneral: promedioGeneral
    };
    } catch (e) {
        Logger.log('Error en evaluarCumplimientoNOM: ' + e.message);
        return { success: false, error: e.message };
    }
    }
    function getUsuarios() { return getSheetData('Usuarios'); }
    function getPatrullajes() { return getSheetData('Patrullajes'); }
    function getCapacitaciones(userId) {
    var data = getSheetData('Capacitaciones');
    if (userId) data = data.filter(function(c) { return c.Usuario === userId; });
    return data;
    }
    function getAptitudesMedicas(userId) {
    var data = getSheetData('Aptitudes_Medicas');
    if (userId) data = data.filter(function(a) { return a.Usuario === userId; });
    return data;
    }
    function getEPP() { return getSheetData('EPP'); }
    function getCatalogoActividades() { return getSheetData('Catalogo_Actividades'); }
    function getPlanCapacitacion() { return getSheetData('Plan_Capacitacion'); }
    function getBancoPlaticas() { return getSheetData('Banco_Platicas'); }

    function guardarCatalogoActividad(payload) {
    var auth = checkRole(payload.creadoPor || payload.userId, 'SUPERVISOR');
    if (!auth.allowed) return { success: false, error: auth.error };
    var id = genId('CAT', 'Catalogo_Actividades');
    appendRow('Catalogo_Actividades', [
        id,
        payload.nombre,
        payload.tipo,
        payload.descripcion || '',
        payload.periodicidad || 'Evento',
        new Date().toISOString()
    ]);
    logAction(payload.usuario || 'ADMIN', 'Agregar actividad al catálogo', id, payload.nombre);
    return { success: true, id: id };
    }

    // Check operability (PSST blocks) — delegates to checkOperability for consistency
    function canOperate(userId) {
    var result = checkOperability(userId);
    return { ok: result.canOperate, blocks: result.blocks };
    }

    // ================================================================
    // ESCRITURAS
    // ================================================================

    function crearUsuario(payload) {
    var roleCheck = checkRole(payload.creadoPor, 'ADMIN');
    if (!roleCheck.allowed) return { success: false, error: roleCheck.error };
    var id = genId('U', 'Usuarios');
    var ini = payload.nombre.trim().split(' ').filter(function(w) { return w.length > 0; }).map(function(w) { return w[0]; }).join('').substring(0, 2).toUpperCase() || 'XX';
    // Headers: ID, Nombre, PIN, Rol, Iniciales, Activo, Puesto, ActivoAsignado, NumeroEco, Timestamp, CreadoPor
    appendRow('Usuarios', [
        id,
        payload.nombre,
        payload.pin,
        payload.rol,
        ini,
        true,
        payload.puesto || '',
        payload.activoAsignado || '',
        payload.numeroEco || '',
        new Date().toISOString(),
        payload.creadoPor || ''
    ]);
    logAction(payload.creadoPor, 'Crear usuario', id, payload.rol);
    return { success: true, id: id };
    }

    function guardarAST(payload) {
    // RBAC: Verificar rol LIDER o superior — el Líder firma su propio AST en campo
    var roleCheck = checkRole(payload.firmadoPor || payload.lider, 'LIDER');
    if (!roleCheck.allowed) return { success: false, error: roleCheck.error };
    // PSST: Verificar operabilidad (solo bloquear si HAY registros vencidos, no si no hay registros)
    var userId = payload.firmadoPor || payload.lider;
    var caps = getCapacitaciones(userId);
    var capVencida = caps.length > 0 && caps.some(function(c) { return c.Estatus === 'Vencida'; });
    if (capVencida) return { success: false, error: 'Capacitacion vencida. No puede firmar AST.' };
    var apts = getAptitudesMedicas(userId);
    var aptVencida = apts.length > 0 && apts.some(function(a) { return a.Estatus === 'Vencida' || a.Estatus === 'No Apto'; });
    if (aptVencida) return { success: false, error: 'Aptitud medica vencida o No Apto.' };
        var id = genId('AST', 'AST');
    var activTipo = payload.activTipo || payload.ActividadTipo || '';
    var firmadoPor = payload.firmadoPor || payload.lider;
    appendRow('AST', [
        id, 
        payload.actividad || '', 
        payload.lider || '', 
        payload.fecha || todayStr(), 
        activTipo,
        JSON.stringify(payload.peligros || []), 
        JSON.stringify(payload.riesgos || []), 
        JSON.stringify(payload.controles || []),
        payload.firma ? true : false,
        firmadoPor,
        payload.geoLat || 0, 
        payload.geoLng || 0, 
        new Date().toISOString()
    ]);
    
    // Update activity state
    if (payload.actividad) {
        var sheet = getSheet('Actividades');
        var data = sheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
        if (data[i][0] === payload.actividad) {
            sheet.getRange(i + 1, 7).setValue('AST aprobado');
            sheet.getRange(i + 1, 9).setValue(id);
            break;
        }
        }
    }
    logAction(payload.lider, 'AST firmado', id, payload.activTipo);
    return { success: true, id: id };
    }

    function guardarPatrullaje(payload) {
    var auth = checkRole(payload.lider, 'LIDER');
    if (!auth.allowed) return { success: false, error: auth.error };
    var opCheck = checkOperability(payload.lider);
    if (!opCheck.canOperate) return { success: false, error: 'Operacion bloqueada: ' + opCheck.blocks.join(', ') };
    var id = genId('P', 'Patrullajes');
    
    var fotoUrls = [];
    var fotoIds = [];
    if (payload.fotos && payload.fotos.length > 0) {
        var folder;
        try {
        folder = DriveApp.getFolderById(ROOT_FOLDER_ID).getFoldersByName('Evidencias').next();
        } catch(e) {
        folder = DriveApp.getFolderById(ROOT_FOLDER_ID);
        }
        
        payload.fotos.forEach(function(foto, i) {
        try {
            // FIX: Fallback si foto no tiene prefijo data:
            var b64data = foto.indexOf(',') >= 0 ? foto.split(',')[1] : foto;
            var mimeType = 'image/jpeg';
            try { mimeType = foto.split(';')[0].split(':')[1] || 'image/jpeg'; } catch(me) {}
            var ext = mimeType === 'image/png' ? '.png' : '.jpg';
            var blob = Utilities.newBlob(Utilities.base64Decode(b64data), mimeType, id + '_foto_' + (i + 1) + ext);
            var file = folder.createFile(blob);
            fotoUrls.push(file.getUrl());
            fotoIds.push(file.getId());
            
            appendRow('Evidencias', ['EV_' + id + '_' + (i + 1), id, 'Foto', file.getUrl(), file.getId(), new Date().toISOString()]);
        } catch (err) { Logger.log('Error foto: ' + err.message); }
        });
    }
    
    appendRow('Patrullajes', [id, payload.actividad || '', payload.lider, payload.fecha, payload.tramo,
        payload.kmInicio, payload.kmFin, payload.tipo, payload.estatus, payload.observaciones || '',
        payload.hallazgo, payload.hallazgoDesc || '', true, JSON.stringify(fotoIds), new Date().toISOString()]);
    
    if (payload.actividad) {
        var sheet = getSheet('Actividades');
        var data = sheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
        if (data[i][0] === payload.actividad) {
            sheet.getRange(i + 1, 7).setValue('Completada');
            break;
        }
        }
    }
    
    if (payload.hallazgo && payload.hallazgo !== 'Normal') {
        var hId = genId('H', 'Hallazgos');
        appendRow('Hallazgos', [hId, id, payload.hallazgo, payload.hallazgoDesc || payload.observaciones || '',
        payload.tramo, payload.kmInicio, 'Abierto', payload.fecha, payload.lider, '']);
    }
    
    logAction(payload.lider, 'Registrar patrullaje', id, 'Tramo: ' + payload.tramo);
    return { success: true, id: id, fotoIds: fotoIds };
    }


    function activarStopWork(payload) {
    var auth = checkRole(payload.lider, 'LIDER');
    if (!auth.allowed) return { success: false, error: auth.error };
    var id = genId('SW', 'Stop_Work');
    appendRow('Stop_Work', [id, payload.lider, payload.fecha, payload.motivo, 'Activo', '', '', new Date().toISOString()]);
    logAction(payload.lider, 'Stop Work ACTIVADO', id, payload.motivo);
    return { success: true, id: id };
    }

    function liberarStopWork(payload) {
    var swId = payload.swId;
    var adminId = payload.adminId;
    var roleCheck = checkRole(adminId, 'ADMIN');
    if (!roleCheck.allowed) return { success: false, error: 'Solo ADMIN puede liberar Stop Work' };
    var sheet = getSheet('Stop_Work');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === swId) {
        sheet.getRange(i + 1, 5).setValue('Liberado');
        sheet.getRange(i + 1, 6).setValue(adminId);
        sheet.getRange(i + 1, 7).setValue(new Date().toISOString());
        break;
        }
    }
    logAction(adminId, 'Stop Work LIBERADO', swId, '');
    return { success: true };
    }

    function cerrarHallazgo(payload) {
    var hallazgoId = payload.hallazgoId;
    var userId = payload.userId;
    var accion = payload.accion || '';
    var evidenciaB64 = payload.evidencia || '';
    
    var roleCheck = checkRole(userId, 'ADMIN');
    if (!roleCheck.allowed) return { success: false, error: 'Se requiere rol ADMIN o superior para cerrar hallazgos' };
    
    var sheet = getSheet('Hallazgos');
    var data = sheet.getDataRange().getValues();
    
    var evidenciaId = '';
    if (evidenciaB64) {
        try {
        var folder = DriveApp.getFolderById(ROOT_FOLDER_ID).getFoldersByName('Evidencias').next();
        // FIX: Manejar base64 con y sin prefijo data:...;base64,
        var rawB64 = evidenciaB64.indexOf(',') >= 0 ? evidenciaB64.split(',')[1] : evidenciaB64;
        var blob = Utilities.newBlob(Utilities.base64Decode(rawB64), 'image/jpeg', 'CIERRE_' + hallazgoId + '.jpg');
        var file = folder.createFile(blob);
        evidenciaId = file.getId();
        } catch(e) { Logger.log('Error subiendo evidencia de cierre: ' + e.message); }
    }

    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === hallazgoId) {
        sheet.getRange(i + 1, 7).setValue('Cerrado');
        sheet.getRange(i + 1, 10).setValue(new Date().toISOString());
        sheet.getRange(i + 1, 11).setValue(accion);
        sheet.getRange(i + 1, 12).setValue(evidenciaId);
        break;
        }
    }
    logAction(userId, 'Cerrar hallazgo', hallazgoId, accion);
    return { success: true };
    }


    function guardarCapacitacion(payload) {
    var roleCheck = checkRole(payload.creadoPor, 'SUPERVISOR');
    if (!roleCheck.allowed) return { success: false, error: roleCheck.error };
    var id = genId('C', 'Capacitaciones');
    appendRow('Capacitaciones', [id, payload.usuario, payload.tipo, payload.nombre, payload.vigencia, payload.dc3, payload.estatus]);
    logAction(payload.creadoPor, 'Agregar capacitacion', id, payload.nombre);
    return { success: true, id: id };
    }

    function guardarAptitudMedica(payload) {
    var roleCheck = checkRole(payload.creadoPor, 'SUPERVISOR');
    if (!roleCheck.allowed) return { success: false, error: roleCheck.error };
    var id = genId('AM', 'Aptitudes_Medicas');
    appendRow('Aptitudes_Medicas', [id, payload.usuario, payload.tipo, payload.vigencia, payload.medico, payload.estatus]);
    logAction(payload.creadoPor, 'Agregar aptitud medica', id, payload.usuario);
    return { success: true, id: id };
    }

    // ================================================================
    // MÓDULO: INSPECCIONES (EPP, Herramientas, Utilitarios, Maq. Pesada)
    // Hoja: Inspecciones
    // Headers: ID|Fecha|Inspector|Tipo|Objetivo|ItemsJSON|Resultado|Observaciones|FotoID|Timestamp
    // ================================================================
    function getInspecciones(tipo) {
    var data = getSheetData('Inspecciones');
    if (tipo) return data.filter(function(i) { return i.Tipo === tipo; });
    return data;
    }

    function guardarInspeccion(payload) {
    var auth = checkRole(payload.inspector || payload.userId, 'SEGURISTA');
    if (!auth.allowed) return { success: false, error: auth.error };
    var id = genId('INS', 'Inspecciones');
    appendRow('Inspecciones', [
        id,
        payload.fecha || Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd'),
        payload.inspector,
        payload.tipo, // EPP, Herramienta, Utilitario, MaqPesada
        payload.objetivo, // nombre del trabajador O placa/ID del vehículo
        JSON.stringify(payload.items || []), // [{item, estado, obs}]
        payload.resultado, // Conforme, No Conforme, Observaciones
        payload.observaciones || '',
        payload.fotoId || '',
        new Date().toISOString()
    ]);
    logAction(payload.inspector, 'Inspección ' + payload.tipo, id, payload.objetivo);
    return { success: true, id: id };
    }

    // Checklists predefinidos por tipo de inspección
    function getChecklistTemplate(tipo) {
    var templates = {
        'EPP': [
        { item: 'Casco de seguridad', criterio: 'Sin grietas, arnés interior completo, barbiquejo' },
        { item: 'Lentes de seguridad', criterio: 'Sin rayones, patillas completas, limpio' },
        { item: 'Guantes', criterio: 'Sin roturas, talla correcta, tipo adecuado' },
        { item: 'Botas de seguridad', criterio: 'Casquillo presente, suela sin desgaste, agujetas' },
        { item: 'Chaleco reflejante', criterio: 'Cintas reflectantes visibles, sin roturas' },
        { item: 'Protección auditiva', criterio: 'Tapones/orejeras en buen estado, limpios' },
        { item: 'Protección respiratoria', criterio: 'Filtros vigentes, ajuste correcto (si aplica)' },
        { item: 'Ropa de trabajo', criterio: 'Manga larga, sin roturas, identificación visible' },
        { item: 'Protección contra caídas', criterio: 'Arnés sin daño, línea de vida, mosquetones (si aplica)' },
        { item: 'Bloqueador solar', criterio: 'Disponible, aplicado (trabajos exteriores)' }
        ],
        'Herramienta': [
        { item: 'Machete / herramienta de corte', criterio: 'Filo adecuado, mango firme, funda' },
        { item: 'Pala', criterio: 'Mango firme, sin astillas, punta sin desgaste' },
        { item: 'Pico', criterio: 'Cabeza asegurada, mango sin grietas' },
        { item: 'Barreta', criterio: 'Sin deformación, punta funcional' },
        { item: 'Sierra / serrucho', criterio: 'Dientes completos, mango seguro' },
        { item: 'Desbrozadora', criterio: 'Guarda protectora, arnés, disco/hilo correcto' },
        { item: 'Motosierra', criterio: 'Cadena afilada y tensada, freno funcional, aceite' },
        { item: 'Escalera', criterio: 'Peldaños firmes, zapatas antiderrapantes, sin deformación' },
        { item: 'Extensiones / cables', criterio: 'Aislamiento íntegro, sin empalmes expuestos' },
        { item: 'Herramienta manual general', criterio: 'Mangos seguros, sin óxido, limpias, en bolsa/caja' },
        { item: 'Código de colores (cinta)', criterio: 'Cinta del mes/trimestre vigente aplicada' }
        ],
        'Utilitario': [
        { item: 'Documentos del vehículo', criterio: 'Tarjeta de circulación, póliza de seguro, verificación' },
        { item: 'Licencia del conductor', criterio: 'Vigente, tipo adecuado para el vehículo' },
        { item: 'Nivel de combustible', criterio: 'Mínimo 1/4 de tanque' },
        { item: 'Aceite de motor', criterio: 'Nivel correcto, sin fugas visibles' },
        { item: 'Líquido de frenos', criterio: 'Nivel correcto' },
        { item: 'Anticongelante/refrigerante', criterio: 'Nivel correcto, sin fugas' },
        { item: 'Llantas', criterio: 'Presión correcta, dibujo >2mm, sin daños laterales' },
        { item: 'Llanta de refacción', criterio: 'Presente, inflada, herramienta de cambio' },
        { item: 'Luces', criterio: 'Frontales, traseras, direccionales, reversa, freno — todas funcionan' },
        { item: 'Frenos', criterio: 'Respuesta correcta, sin ruidos, freno de mano funcional' },
        { item: 'Cinturones de seguridad', criterio: 'Todos funcionales, sin cortes' },
        { item: 'Espejos', criterio: 'Laterales y retrovisor completos, sin grietas' },
        { item: 'Limpiadores', criterio: 'Funcionan, líquido limpiaparabrisas con nivel' },
        { item: 'Claxon', criterio: 'Funcional' },
        { item: 'Extintor', criterio: 'Cargado, vigente, accesible, manómetro en verde' },
        { item: 'Botiquín', criterio: 'Completo, materiales vigentes' },
        { item: 'Triángulos/conos', criterio: 'Presentes y reflectantes' },
        { item: 'Kit de derrames', criterio: 'Presente (si transporta materiales)' },
        { item: 'Calcomanías de seguridad', criterio: 'Logo empresa, no fumar, velocidad máxima' },
        { item: 'Limpieza general', criterio: 'Interior y exterior limpios, sin objetos sueltos' }
        ],
        'MaqPesada': [
        { item: 'Documentos y permisos', criterio: 'Tarjeta, seguro, permisos especiales vigentes' },
        { item: 'Operador certificado', criterio: 'DC3 vigente, licencia federal (si aplica)' },
        { item: 'Nivel de combustible/hidráulico', criterio: 'Niveles correctos, sin fugas' },
        { item: 'Aceite de motor', criterio: 'Nivel correcto, sin fugas' },
        { item: 'Sistema hidráulico', criterio: 'Mangueras sin fugas, cilindros sin daño, conexiones firmes' },
        { item: 'Orugas / llantas', criterio: 'Tensión correcta, sin daño, pernos completos' },
        { item: 'Estructura y chasis', criterio: 'Sin grietas, soldaduras íntegras' },
        { item: 'Luces y alarmas', criterio: 'Luz estroboscópica, alarma de reversa, luces de trabajo' },
        { item: 'Espejos y cámaras', criterio: 'Visibilidad completa, cámara de reversa (si tiene)' },
        { item: 'Cabina del operador', criterio: 'Vidrios completos, asiento seguro, cinturón' },
        { item: 'Controles operativos', criterio: 'Palancas, pedales, switches — respuesta correcta' },
        { item: 'Frenos', criterio: 'Servicio y estacionamiento funcionales' },
        { item: 'Implemento / herramienta', criterio: 'Bote, brazo, cuchilla — sin daño, pines seguros' },
        { item: 'Extintor', criterio: 'Cargado, vigente, accesible' },
        { item: 'Protección contra vuelco (ROPS)', criterio: 'Estructura íntegra, pernos completos' },
        { item: 'Kit de derrames', criterio: 'Material absorbente presente' },
        { item: 'Señalización', criterio: 'Calcomanías de seguridad, capacidad de carga, puntos de izaje' },
        { item: 'Bitácora de mantenimiento', criterio: 'Último servicio registrado, próximo programado' }
        ]
    };
    return templates[tipo] || [];
    }

    // ================================================================
    // MÓDULO: PERMISOS DE TRABAJO (PTs Estandarizados)
    // Hoja: Permisos_Trabajo
    // Headers: ID|Fecha|Solicitante|TipoPT|Ubicacion|Descripcion|RiesgosJSON|ControlesJSON|HoraInicio|HoraFin|AutorizadoPor|Estado|FirmaEmisor|FirmaEjecutor|Timestamp
    // ================================================================
    function getPermisosTrabajo(estado) {
    var data = getSheetData('Permisos_Trabajo');
    if (estado) return data.filter(function(p) { return p.Estado === estado; });
    return data;
    }

    function guardarPermiso(payload) {
    var auth = checkRole(payload.solicitante || payload.userId, 'SEGURISTA');
    if (!auth.allowed) return { success: false, error: auth.error };
    var id = genId('PT', 'Permisos_Trabajo');
    appendRow('Permisos_Trabajo', [
        id,
        payload.fecha || Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd'),
        payload.solicitante,
        payload.tipoPT, // Altura, EspacioConfinado, Caliente, Excavacion, Electrico, Izaje
        payload.ubicacion,
        payload.descripcion,
        JSON.stringify(payload.riesgos || []),
        JSON.stringify(payload.controles || []),
        payload.horaInicio,
        payload.horaFin,
        payload.autorizadoPor || '',
        'Abierto',
        payload.firmaEmisor || false,
        payload.firmaEjecutor || false,
        new Date().toISOString()
    ]);
    logAction(payload.solicitante, 'Permiso de trabajo ' + payload.tipoPT, id, payload.ubicacion);
    return { success: true, id: id };
    }

    function cerrarPermiso(payload) {
    var auth = checkRole(payload.cerradoPor || payload.userId, 'SUPERVISOR');
    if (!auth.allowed) return { success: false, error: auth.error };
    var sheet = getSheet('Permisos_Trabajo');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
        sheet.getRange(i + 1, 12).setValue('Cerrado'); // Estado
        sheet.getRange(i + 1, 10).setValue(payload.horaFin || Utilities.formatDate(new Date(), 'America/Mexico_City', 'HH:mm')); // HoraFin
        logAction(payload.usuario, 'Cerrar PT', payload.id, '');
        return { success: true };
        }
    }
    return { success: false, error: 'PT no encontrado' };
    }

    function autorizarPermiso(payload) {
    var auth = checkRole(payload.autorizadoPor || payload.userId, 'SUPERVISOR');
    if (!auth.allowed) return { success: false, error: auth.error };
    var sheet = getSheet('Permisos_Trabajo');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
        sheet.getRange(i + 1, 11).setValue(payload.autorizadoPor); // AutorizadoPor
        sheet.getRange(i + 1, 12).setValue('Autorizado'); // Estado
        logAction(payload.autorizadoPor, 'Autorizar PT', payload.id, data[i][3]);
        return { success: true };
        }
    }
    return { success: false, error: 'PT no encontrado' };
    }

    // Templates de riesgos y controles por tipo de PT
    function getPTTemplate(tipo) {
    var templates = {
        'Altura': {
        riesgos: ['Caída a diferente nivel', 'Caída de objetos', 'Golpe contra estructura', 'Contacto con líneas eléctricas', 'Condiciones climáticas adversas'],
        controles: ['Arnés de cuerpo completo con doble línea de vida', 'Punto de anclaje certificado (5000 lb)', 'Línea de vida retráctil', 'Red o malla de protección', 'Barricada y señalización en zona inferior', 'Casco con barbiquejo', 'Radio de comunicación', 'Verificar condiciones climáticas (viento <40 km/h)', 'Rescate vertical disponible'],
        requiereAST: true,
        alturaMinima: '1.80 m'
        },
        'EspacioConfinado': {
        riesgos: ['Atmósfera deficiente de O2 (<19.5%)', 'Atmósfera enriquecida de O2 (>23.5%)', 'Presencia de gases tóxicos (H2S, CO)', 'Gases inflamables (LEL >10%)', 'Atrapamiento', 'Ahogamiento'],
        controles: ['Monitoreo de atmósfera continuo (4 gases)', 'Ventilación forzada', 'Vigía permanente en boca de acceso', 'Equipo de rescate listo', 'Arnés de rescate con línea de vida', 'SCBA o línea de aire disponible', 'Comunicación bidireccional', 'Permiso firmado por entrada', 'Bloqueo y etiquetado de energías'],
        requiereAST: true,
        monitoreoRequerido: 'Continuo 4 gases (O2, LEL, H2S, CO)'
        },
        'Caliente': {
        riesgos: ['Incendio', 'Explosión', 'Quemaduras', 'Inhalación de humos', 'Radiación UV/IR', 'Proyección de partículas'],
        controles: ['Extintor tipo ABC (mín. 6 kg) a 3 m', 'Monitoreo de gases (LEL <10%)', 'Manta ignífuga / pantalla', 'Guardia de fuego (fire watch) 30 min posterior', 'Retiro de materiales combustibles (radio 11 m)', 'Careta de soldador / lentes oscuros', 'Guantes de carnaza', 'Mandil de cuero', 'Ventilación adecuada'],
        requiereAST: true,
        radioSeguridad: '11 metros libre de combustibles'
        },
        'Excavacion': {
        riesgos: ['Derrumbe / colapso de paredes', 'Caída a la excavación', 'Golpe por maquinaria', 'Contacto con líneas subterráneas (gas, eléctrica, agua)', 'Inundación', 'Atmósfera peligrosa'],
        controles: ['Localización de servicios subterráneos (planos)', 'Ademado/entibado para >1.20 m de profundidad', 'Barricada perimetral rígida', 'Escalera de acceso cada 7.5 m', 'Monitoreo de atmósfera', 'Distancia mínima de maquinaria al borde', 'Señalización y delimitación', 'Material de excavación a mín. 60 cm del borde', 'Plan de rescate'],
        requiereAST: true,
        profundidadCritica: '1.20 m requiere ademado'
        },
        'Electrico': {
        riesgos: ['Electrocución / choque eléctrico', 'Arco eléctrico', 'Quemaduras', 'Caída por choque eléctrico', 'Explosión por arco'],
        controles: ['Bloqueo y etiquetado (LOTO)', 'Verificación de ausencia de tensión', 'Puesta a tierra temporal', 'Herramienta dieléctrica certificada', 'Guantes dieléctricos con clase adecuada', 'Tapetes aislantes', 'Distancias mínimas de seguridad', 'Personal calificado (DC3)', 'Señalización de zona energizada'],
        requiereAST: true
        },
        'Izaje': {
        riesgos: ['Caída de carga', 'Golpe por carga en movimiento', 'Vuelco de equipo de izaje', 'Falla de eslingas/accesorios', 'Aplastamiento', 'Contacto con líneas eléctricas'],
        controles: ['Plan de izaje con cálculo de cargas', 'Inspección de eslingas, estrobos y grilletes', 'Verificar capacidad de carga del equipo', 'Señalero/banderero designado', 'Barricada del área de maniobra', 'Vientos de control (tag lines)', 'No pasar bajo carga suspendida', 'Verificar condiciones de viento', 'Operador certificado (DC3)'],
        requiereAST: true
        }
    };
    return templates[tipo] || { riesgos: [], controles: [], requiereAST: true };
    }

    // ================================================================
    // MÓDULO: PLANES DE EMERGENCIA
    // Hoja: Emergencias
    // Headers: ID|Tipo|Titulo|Fecha|Responsable|Participantes|Evaluacion|Observaciones|Estado|Timestamp
    // ================================================================
    function getEmergencias() { return getSheetData('Emergencias'); }

    function guardarSimulacro(payload) {
    var auth = checkRole(payload.responsable || payload.userId, 'SEGURISTA');
    if (!auth.allowed) return { success: false, error: auth.error };
    var id = genId('SIM', 'Emergencias');
    appendRow('Emergencias', [
        id,
        'Simulacro',
        payload.titulo,
        payload.fecha,
        payload.responsable,
        JSON.stringify(payload.participantes || []),
        payload.evaluacion || '', // Bueno, Regular, Deficiente
        payload.observaciones || '',
        'Realizado',
        new Date().toISOString()
    ]);
    logAction(payload.responsable, 'Registro simulacro', id, payload.titulo);
    return { success: true, id: id };
    }

    // ================================================================
    // MÓDULO: PROGRAMA NOM-030-STPS (Salud Ocupacional)
    // Hoja: Programa_Salud
    // Headers: ID|Tipo|Actividad|Frecuencia|FechaPrograma|FechaEjecucion|Responsable|Estatus|Evidencia|Observaciones|Timestamp
    // ================================================================
    function getProgramaSalud() { return getSheetData('Programa_Salud'); }

    function guardarActividadSalud(payload) {
    var auth = checkRole(payload.responsable || payload.userId, 'SEGURISTA');
    if (!auth.allowed) return { success: false, error: auth.error };
    var id = genId('SAL', 'Programa_Salud');
    appendRow('Programa_Salud', [
        id,
        payload.tipo, // EMO, Vigilancia, Capacitacion, Prevencion
        payload.actividad,
        payload.frecuencia, // Anual, Semestral, Trimestral, Mensual, Unica
        payload.fechaPrograma,
        payload.fechaEjecucion || '',
        payload.responsable,
        payload.estatus || 'Programado', // Programado, En Proceso, Realizado, Vencido
        payload.evidencia || '',
        payload.observaciones || '',
        new Date().toISOString()
    ]);
    logAction(payload.responsable || 'ADMIN', 'Actividad salud NOM-030', id, payload.actividad);
    return { success: true, id: id };
    }

    function ejecutarActividadSalud(payload) {
    var auth = checkRole(payload.ejecutadoPor || payload.userId, 'SEGURISTA');
    if (!auth.allowed) return { success: false, error: auth.error };
    var sheet = getSheet('Programa_Salud');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
        var nuevoEstatus = payload.estatus || 'Realizado';
        sheet.getRange(i + 1, 6).setValue(payload.fechaEjecucion || Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd'));
        sheet.getRange(i + 1, 8).setValue(nuevoEstatus);
        if (payload.evidencia) sheet.getRange(i + 1, 9).setValue(payload.evidencia);
        if (payload.observaciones) sheet.getRange(i + 1, 10).setValue(payload.observaciones);
        logAction(payload.usuario || 'ADMIN', 'Ejecutar actividad salud', payload.id, nuevoEstatus);
        return { success: true };
        }
    }
    return { success: false, error: 'Actividad no encontrada' };
    }

    function actualizarCapacitacion(payload) {
    var auth = checkRole(payload.userId, 'SUPERVISOR');
    if (!auth.allowed) return { success: false, error: auth.error };
    var sheet = getSheet('Capacitaciones');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
        if (payload.estatus) sheet.getRange(i + 1, 7).setValue(payload.estatus);
        if (payload.vigencia) sheet.getRange(i + 1, 5).setValue(payload.vigencia);
        if (payload.dc3 !== undefined) sheet.getRange(i + 1, 6).setValue(payload.dc3);
        logAction(payload.usuario || 'ADMIN', 'Actualizar capacitación', payload.id, payload.estatus || '');
        return { success: true };
        }
    }
    return { success: false, error: 'Capacitación no encontrada' };
    }

    function actualizarAptitud(payload) {
    var auth = checkRole(payload.userId, 'SUPERVISOR');
    if (!auth.allowed) return { success: false, error: auth.error };
    var sheet = getSheet('Aptitudes_Medicas');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
        if (payload.estatus) sheet.getRange(i + 1, 6).setValue(payload.estatus);
        if (payload.tipo) sheet.getRange(i + 1, 3).setValue(payload.tipo);
        if (payload.vigencia) sheet.getRange(i + 1, 4).setValue(payload.vigencia);
        logAction(payload.usuario || 'ADMIN', 'Actualizar aptitud médica', payload.id, payload.estatus || '');
        return { success: true };
        }
    }
    return { success: false, error: 'Aptitud no encontrada' };
    }

    // ================================================================
    // SETUP: Crear hojas nuevas para los módulos críticos
    // Ejecutar UNA VEZ desde el editor de Apps Script
    // ================================================================
    // ================================================================
    // MÓDULO: AUDITORÍAS SASISOPA (Registro, Hallazgos, Plan de Acción)
    // Hoja: Auditorias (headers actualizados)
    // Headers: ID|Fecha|Tipo|Auditor|ElementosJSON|Resultado|NCMayores|NCMenores|Observaciones|Estado|PlanAccionJSON|Timestamp
    // ================================================================
    function getAuditorias() { return getSheetData('Auditorias'); }

    function guardarAuditoria(payload) {
    var auth = checkRole(payload.auditor || payload.userId, 'SUPERVISOR');
    if (!auth.allowed) return { success: false, error: auth.error };
    var id = genId('AUD', 'Auditorias');
    var elementos = payload.elementos || [];
    var ncMayores = 0;
    var ncMenores = 0;
    var obs = 0;
    elementos.forEach(function(e) {
        if (e.resultado === 'NC Mayor') ncMayores++;
        else if (e.resultado === 'NC Menor') ncMenores++;
        else if (e.resultado === 'Observacion') obs++;
    });
    var resultado = ncMayores > 0 ? 'No Satisfactorio' : ncMenores > 3 ? 'Condicionado' : 'Satisfactorio';
    
    appendRow('Auditorias', [
        id,
        payload.fecha || Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd'),
        payload.tipo || 'Interna', // Interna, Externa, ASEA, Pre-auditoría
        payload.auditor,
        JSON.stringify(elementos),
        resultado,
        ncMayores,
        ncMenores,
        payload.observaciones || '',
        'Abierta', // Abierta, En seguimiento, Cerrada
        JSON.stringify(payload.planAccion || []),
        new Date().toISOString()
    ]);
    
    // Actualizar cumplimiento SASISOPA basado en la auditoría
    var sasSheet = getSheet('SASISOPA');
    if (sasSheet) {
        var sasData = sasSheet.getDataRange().getValues();
        elementos.forEach(function(elem) {
        for (var i = 1; i < sasData.length; i++) {
            if (sasData[i][0] === 'S' + String(elem.numero).padStart(2, '0')) {
            var cumpl = elem.resultado === 'Cumple' ? 100 : elem.resultado === 'Observacion' ? 85 : elem.resultado === 'NC Menor' ? 60 : 30;
            sasSheet.getRange(i + 1, 3).setValue(cumpl);
            break;
            }
        }
        });
    }
    
    logAction(payload.auditor || 'ADMIN', 'Auditoría SASISOPA', id, resultado + ' | NC+:' + ncMayores + ' NC-:' + ncMenores);
    return { success: true, id: id, resultado: resultado, ncMayores: ncMayores, ncMenores: ncMenores };
    }

    function cerrarAuditoria(payload) {
    var auth = checkRole(payload.userId, 'SUPERVISOR');
    if (!auth.allowed) return { success: false, error: auth.error };
    var sheet = getSheet('Auditorias');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
        sheet.getRange(i + 1, 10).setValue('Cerrada');
        logAction(payload.usuario || 'ADMIN', 'Cerrar auditoría', payload.id, '');
        return { success: true };
        }
    }
    return { success: false, error: 'Auditoría no encontrada' };
    }

    function actualizarPlanAccion(payload) {
    var auth = checkRole(payload.userId, 'SUPERVISOR');
    if (!auth.allowed) return { success: false, error: auth.error };
    var sheet = getSheet('Auditorias');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
        sheet.getRange(i + 1, 11).setValue(JSON.stringify(payload.planAccion || []));
        sheet.getRange(i + 1, 10).setValue('En seguimiento');
        logAction(payload.usuario || 'ADMIN', 'Actualizar plan acción', payload.id, '');
        return { success: true };
        }
    }
    return { success: false, error: 'Auditoría no encontrada' };
    }

    // ================================================================
    // MÓDULO: LISTADO MAESTRO DE DOCUMENTOS (Control Documental)
    // Hoja: Listado_Maestro — Elemento SASISOPA: 14
    // ================================================================
    function getListadoMaestro() { return getSheetData('Listado_Maestro'); }

    function guardarDocumento(payload) {
    var auth = checkRole(payload.responsable || payload.userId, 'SEGURISTA');
    if (!auth.allowed) return { success: false, error: auth.error };
    var id = genId('DOC', 'Listado_Maestro');
    var ubicacion = payload.ubicacion || 'Digital — DmV Control';
    
    // Si se envía un archivo (base64), subirlo a Drive
    if (payload.archivoBase64) {
        try {
        var folder = _getDocFolder(payload.elemento);
        var b64data = payload.archivoBase64.split(',')[1];
        var mimeType = 'application/pdf';
        try { mimeType = payload.archivoBase64.split(';')[0].split(':')[1] || 'application/pdf'; } catch(me) {}
        var ext = _getExtFromMime(mimeType);
        var fileName = (payload.codigo || id) + '_' + (payload.titulo || 'doc').replace(/[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ ]/g, '').substring(0, 40) + ext;
        var blob = Utilities.newBlob(Utilities.base64Decode(b64data), mimeType, fileName);
        var file = folder.createFile(blob);
        ubicacion = file.getUrl();
        } catch(fileErr) {
        Logger.log('Error subiendo archivo para documento ' + id + ': ' + fileErr.message);
        // Continuar sin archivo — se registra el metadato igual
        }
    }
    
    appendRow('Listado_Maestro', [
        id,
        payload.codigo,
        payload.titulo,
        payload.elemento || '',
        payload.version || '1.0',
        payload.fechaEmision || Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd'),
        payload.fechaRevision || '',
        payload.responsable || '',
        ubicacion,
        payload.estado || 'Vigente',
        new Date().toISOString()
    ]);
    logAction(payload.creadoPor || 'ADMIN', 'Registro documento', id, payload.codigo + ' - ' + payload.titulo + (payload.archivoBase64 ? ' [con archivo]' : ''));
    return { success: true, id: id, url: ubicacion };
    }

    /**
    * Sube un archivo a un documento existente del Listado Maestro.
    * Actualiza la columna 'Ubicacion' con la URL del archivo en Drive.
    * Payload: { id, archivoBase64, nombreArchivo?, usuario }
    */
    function subirArchivoDocumento(payload) {
    var auth = checkRole(payload.usuario, 'SEGURISTA');
    if (!auth.allowed) return auth;
    
    var sheet = getSheet('Listado_Maestro');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var colIndex = {};
    headers.forEach(function(h, i) { colIndex[h] = i; });
    
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
        if (!payload.archivoBase64) return { success: false, error: 'No se recibió ningún archivo' };
        
        try {
            var elemento = data[i][colIndex['Elemento']] || '';
            var codigo = data[i][colIndex['Codigo']] || payload.id;
            var titulo = data[i][colIndex['Titulo']] || 'doc';
            var folder = _getDocFolder(elemento);
            var b64data = payload.archivoBase64.split(',')[1];
            var mimeType = 'application/pdf';
            try { mimeType = payload.archivoBase64.split(';')[0].split(':')[1] || 'application/pdf'; } catch(me) {}
            var ext = _getExtFromMime(mimeType);
            var fileName = codigo + '_' + titulo.replace(/[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ ]/g, '').substring(0, 40) + ext;
            var blob = Utilities.newBlob(Utilities.base64Decode(b64data), mimeType, fileName);
            var file = folder.createFile(blob);
            var fileUrl = file.getUrl();
            
            // Actualizar columna Ubicacion
            if (colIndex['Ubicacion'] !== undefined) {
            sheet.getRange(i + 1, colIndex['Ubicacion'] + 1).setValue(fileUrl);
            }
            logAction(payload.usuario, 'Subir archivo documento', payload.id, fileName);
            return { success: true, url: fileUrl, fileId: file.getId() };
        } catch(e) {
            return { success: false, error: 'Error al subir archivo: ' + e.message };
        }
        }
    }
    return { success: false, error: 'Documento no encontrado: ' + payload.id };
    }

    /**
    * Obtiene o crea la subcarpeta de documentos para un elemento SASISOPA.
    * Estructura: ROOT/SASISOPA_Documentos/Elem_XX/
    */
    function _getDocFolder(elemento) {
    var root = DriveApp.getFolderById(ROOT_FOLDER_ID);
    var mainFolder;
    var mainIter = root.getFoldersByName('SASISOPA_Documentos');
    if (mainIter.hasNext()) {
        mainFolder = mainIter.next();
    } else {
        mainFolder = root.createFolder('SASISOPA_Documentos');
    }
    if (elemento) {
        var elemName = 'Elem_' + String(elemento).padStart(2, '0');
        var elemIter = mainFolder.getFoldersByName(elemName);
        if (elemIter.hasNext()) return elemIter.next();
        return mainFolder.createFolder(elemName);
    }
    return mainFolder;
    }

    function _getExtFromMime(mimeType) {
    var map = {
        'application/pdf': '.pdf',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document': '.docx',
        'application/msword': '.doc',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx',
        'application/vnd.ms-excel': '.xls',
        'application/vnd.openxmlformats-officedocument.presentationml.presentation': '.pptx',
        'image/jpeg': '.jpg',
        'image/png': '.png',
        'text/plain': '.txt'
    };
    return map[mimeType] || '.pdf';
    }

    function actualizarDocumento(payload) {
    var auth = checkRole(payload.userId, 'SEGURISTA');
    if (!auth.allowed) return { success: false, error: auth.error };
    var sheet = getSheet('Listado_Maestro');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    // Find column indices by header name
    var colIndex = {};
    headers.forEach(function(h, i) {
        colIndex[h] = i;
    });
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
        if (payload.estado && colIndex['Estado'] !== undefined) sheet.getRange(i + 1, colIndex['Estado'] + 1).setValue(payload.estado);
        if (payload.version && colIndex['Version'] !== undefined) sheet.getRange(i + 1, colIndex['Version'] + 1).setValue(payload.version);
        if (payload.fechaRevision && colIndex['FechaRevision'] !== undefined) sheet.getRange(i + 1, colIndex['FechaRevision'] + 1).setValue(payload.fechaRevision);
        logAction(payload.usuario || 'ADMIN', 'Actualizar documento', payload.id, '');
        return { success: true };
        }
    }
    return { success: false, error: 'Documento no encontrado' };
    }

    /**
    * Elimina un documento del Listado Maestro.
    * Opcionalmente elimina el archivo de Drive si existe.
    * Solo ADMIN puede eliminar documentos.
    * Payload: { id, usuario, eliminarArchivo? (boolean) }
    */
    function eliminarDocumento(payload) {
    var auth = checkRole(payload.usuario, 'ADMIN');
    if (!auth.allowed) return auth;
    
    var sheet = getSheet('Listado_Maestro');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var colIndex = {};
    headers.forEach(function(h, i) { colIndex[h] = i; });
    
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
        var codigo = data[i][colIndex['Codigo']] || '';
        var titulo = data[i][colIndex['Titulo']] || '';
        var ubicacion = data[i][colIndex['Ubicacion']] || '';
        
        // Intentar eliminar archivo de Drive si se solicita y existe URL
        if (payload.eliminarArchivo !== false && ubicacion && ubicacion.indexOf('drive.google.com') >= 0) {
            try {
            // Extraer file ID de la URL de Drive
            var fileId = '';
            var match = ubicacion.match(/\/d\/([a-zA-Z0-9_-]+)/);
            if (match) fileId = match[1];
            if (!fileId) {
                match = ubicacion.match(/id=([a-zA-Z0-9_-]+)/);
                if (match) fileId = match[1];
            }
            if (fileId) {
                DriveApp.getFileById(fileId).setTrashed(true);
                Logger.log('Archivo movido a papelera: ' + fileId);
            }
            } catch(driveErr) {
            Logger.log('No se pudo eliminar archivo de Drive: ' + driveErr.message);
            // Continuar con la eliminación del registro
            }
        }
        
        // Eliminar fila de la hoja
        sheet.deleteRow(i + 1);
        logAction(payload.usuario, 'ELIMINAR DOCUMENTO', payload.id, codigo + ' - ' + titulo);
        return { success: true, mensaje: 'Documento eliminado: ' + codigo };
        }
    }
    return { success: false, error: 'Documento no encontrado: ' + payload.id };
    }

    // ================================================================
    // MÓDULO SEGURIDAD Y MEDIO AMBIENTE (HSE)
    // ================================================================

    function guardarResiduo(datos) {
    var auth = checkRole(datos.userId, 'SEGURISTA');
    if (!auth.allowed) return { success: false, error: auth.error };
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName('Residuos');
    var id = genId('RES-', 'Residuos');
    
    sh.appendRow([
        id,
        datos.fecha,
        datos.tipo, // RP, RME, RSU
        datos.cantidad,
        datos.unidad,
        datos.generador, // Usuario
        datos.manifiesto || '',
        datos.fotoId || '',
        new Date().toISOString()
    ]);
    return { success: true, id: id };
    }

    function guardarInspExtintor(datos) {
    var auth = checkRole(datos.inspector || datos.userId, 'SEGURISTA');
    if (!auth.allowed) return { success: false, error: auth.error };
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName('Extintores');
    
    sh.appendRow([
        datos.idExtintor,
        datos.ubicacion,
        datos.fecha,
        datos.presion, // OK/NOK
        datos.seguro,
        datos.etiqueta,
        datos.manguera,
        datos.estado, // Apto/No Apto
        datos.inspector,
        new Date().toISOString()
    ]);
    return { success: true };
    }

    function guardarRecorridoComision(datos) {
    var auth = checkRole(datos.lider || datos.userId, 'SEGURISTA');
    if (!auth.allowed) return { success: false, error: auth.error };
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName('Comision_Seguridad');
    var id = genId('CSH-', 'Comision_Seguridad');
    
    sh.appendRow([
        id,
        datos.fecha,
        datos.lider,
        JSON.stringify(datos.integrantes),
        datos.recorrido, // Descripción del área recorrida
        JSON.stringify(datos.hallazgos), // Array de hallazgos
        datos.firma,
        new Date().toISOString()
    ]);
    return { success: true, id: id };
    }

    // ================================================================
    // MÓDULO: ALMACÉN Y CONTROL DE HERRAMIENTAS
    // ================================================================

    function guardarMovimientoAlmacen(datos) {
    var auth = checkRole(datos.usuario || datos.userId, 'SUPERVISOR');
    if (!auth.allowed) return { success: false, error: auth.error };
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var shMov = ss.getSheetByName('Movimientos_Herramientas');
    var shInv = ss.getSheetByName('Catalogo_Herramientas');
    var invData = shInv.getDataRange().getValues();
    
    var id = genId('MOV-', 'Movimientos_Herramientas');
    var fecha = datos.fecha || new Date().toISOString();
    
    // Actualizar Stock en Catalogo_Herramientas
    var found = false;
    for (var i = 1; i < invData.length; i++) {
        if (invData[i][0] == datos.articulo) {
        var currentStock = parseFloat(invData[i][3] || 0);
        var currentCampo = parseFloat(invData[i][4] || 0);
        
        if (datos.tipo === 'Entrada') {
            shInv.getRange(i + 1, 4).setValue(currentStock + datos.cantidad);
        } else if (datos.tipo === 'Salida') {
            shInv.getRange(i + 1, 4).setValue(Math.max(0, currentStock - datos.cantidad));
        } else if (datos.tipo === 'Resguardo') {
            shInv.getRange(i + 1, 4).setValue(Math.max(0, currentStock - datos.cantidad));
            shInv.getRange(i + 1, 5).setValue(currentCampo + datos.cantidad);
        } else if (datos.tipo === 'Devolucion') {
            shInv.getRange(i + 1, 4).setValue(currentStock + datos.cantidad);
            shInv.getRange(i + 1, 5).setValue(Math.max(0, currentCampo - datos.cantidad));
        }
        found = true;
        break;
        }
    }
    
    shMov.appendRow([
        id,
        fecha,
        datos.tipo,
        datos.articulo,
        datos.cantidad,
        datos.usuario,
        datos.refId || '',
        datos.detalles || '',
        new Date().toISOString()
    ]);
    
    return { success: true, id: id };
    }

    function guardarValeHerramienta(datos) {
    var auth = checkRole(datos.supervisor || datos.userId, 'SUPERVISOR');
    if (!auth.allowed) return { success: false, error: auth.error };
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var shVale = ss.getSheetByName('Herramientas_Asignacion');
    var id = genId('VALE-', 'Herramientas_Asignacion');
    
    // Registrar cada item como un movimiento de Resguardo
    datos.items.forEach(function(item) {
        guardarMovimientoAlmacen({
        tipo: 'Resguardo',
        articulo: item.codigo,
        cantidad: item.cantidad,
        usuario: datos.supervisor,
        refId: id,
        detalles: 'Asignación a Líder: ' + datos.lider
        });
    });
    
    shVale.appendRow([
        id,
        new Date().toISOString(),
        datos.supervisor,
        datos.lider,
        datos.cuadrilla || '',
        JSON.stringify(datos.items),
        datos.firma || '',
        'Activo',
        new Date().toISOString()
    ]);
    
    return { success: true, id: id };
    }

    // ================================================================
    // MÓDULO: INVESTIGACIÓN DE INCIDENTES Y ACCIDENTES
    // Hoja: Incidentes — Elemento SASISOPA: 12
    // ================================================================
    function getIncidentes() { return getSheetData('Incidentes'); }

    /**
    * MISSING FUNCTION: getKardexOperador
    * Fetches all relevant safety/training info for an operator.
    */
    function getKardexOperador(payload) {
    var id = payload.operadorId;
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var usuarios = getSheetData('Usuarios');
    var usr = usuarios.find(function(u) { return u.ID === id; });
    if (!usr) return { success: false, error: 'Usuario no encontrado' };

    return {
        success: true,
        usuario: usr,
        capacitaciones: getSheetData('Capacitaciones').filter(function(c) { return c.Usuario === id; }),
        aptitudes: getSheetData('Aptitudes_Medicas').filter(function(a) { return a.Usuario === id; }),
        epp: getSheetData('Entrega_EPP').filter(function(e) { return e.Trabajador === usr.Nombre || e.Trabajador === id || e.Trabajador === usr.ID; }),
        platicas: getSheetData('Platicas').filter(function(p) { return p.Expositor === id || (p.AsistentesJSON && p.AsistentesJSON.indexOf(id) > -1); }),
        asts: getSheetData('AST').filter(function(a) { return a.Lider === id; }),
        inspecciones: getSheetData('Inspecciones').filter(function(i) { return i.Inspector === id; }),
        incidentes: getSheetData('Incidentes').filter(function(inc) { return inc.Responsable === id || (inc.Involucrados && inc.Involucrados.indexOf(id) > -1); }),
        permisos: getSheetData('Permisos_Trabajo').filter(function(p) { return p.Solicitante === id || p.AutorizadoPor === id; })
    };
    }

    function guardarIncidente(payload) {
    var auth = checkRole(payload.responsable || payload.userId, 'SEGURISTA');
    if (!auth.allowed) return { success: false, error: auth.error };
    var id = genId('INC', 'Incidentes');
    appendRow('Incidentes', [
        id,
        payload.fecha || Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd'),
        payload.hora || '',
        payload.lugar || '',
        payload.tipo, // Incidente, Accidente, CuasiAccidente, ActoInseguro, CondicionInsegura
        payload.descripcion,
        payload.involucrados || '',
        payload.lesionTipo || 'Ninguna',
        payload.gravedad || 'Leve', // Leve, Moderado, Grave, Fatal
        payload.causaInmediata || '',
        payload.causaRaiz || '',
        JSON.stringify(payload.acciones || []),
        payload.responsable || '',
        'Abierto',
        payload.fotoId || '',
        new Date().toISOString()
    ]);
    logAction(payload.reportadoPor || 'ADMIN', 'Reporte incidente ' + payload.tipo, id, (payload.descripcion || '').substring(0, 50));
    return { success: true, id: id };
    }

    function cerrarIncidente(payload) {
    var auth = checkRole(payload.userId, 'SUPERVISOR');
    if (!auth.allowed) return { success: false, error: auth.error };
    var sheet = getSheet('Incidentes');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
        sheet.getRange(i + 1, 14).setValue('Cerrado');
        if (payload.acciones) sheet.getRange(i + 1, 12).setValue(JSON.stringify(payload.acciones));
        if (payload.causaRaiz) sheet.getRange(i + 1, 11).setValue(payload.causaRaiz);
        logAction(payload.usuario || 'ADMIN', 'Cerrar incidente', payload.id, '');
        return { success: true };
        }
    }
    return { success: false, error: 'Incidente no encontrado' };
    }

    // ================================================================
    // MÓDULO: PLÁTICAS DE 5 MINUTOS
    // Hoja: Platicas_5min — Elemento SASISOPA: 4, 18
    // ================================================================
    function getPlaticas() { return getSheetData('Platicas_5min'); }

    function guardarPlatica(payload) {
    var auth = checkRole(payload.expositor || payload.userId, 'LIDER');
    if (!auth.allowed) return { success: false, error: auth.error };
    var id = genId('P5M', 'Platicas_5min');
    appendRow('Platicas_5min', [
        id,
        payload.fecha || Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd'),
        payload.tema,
        payload.expositor,
        JSON.stringify(payload.asistentes || []),
        payload.duracion || 5,
        payload.observaciones || '',
        payload.fotoId || '',
        new Date().toISOString()
    ]);
    logAction(payload.expositor, 'Plática 5 min', id, payload.tema);
    return { success: true, id: id };
    }

    // ================================================================
    // MÓDULO: PERSONAL DE CAMPO (Ayudantes / Cuadrilla)
    // Hoja: Personal_Campo — Elemento SASISOPA: 4 (Competencia del Personal)
    // Registra trabajadores que no usan la app pero necesitan estar
    // en ASTs, pláticas, entregas EPP y demás registros de seguridad.
    // ================================================================
    function getEntregaEPP() { return getSheetData('Entrega_EPP'); }

    function guardarEntregaEPP(payload) {
    var auth = checkRole(payload.entregadoPor || payload.userId, 'SEGURISTA');
    if (!auth.allowed) return { success: false, error: auth.error };
    var id = genId('EPP-', 'Entrega_EPP');
    appendRow('Entrega_EPP', [
        id, payload.fecha || Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd'),
        payload.trabajador, JSON.stringify(payload.eppEntregado || []),
        payload.motivo || 'Dotacion', payload.entregadoPor || payload.userId,
        payload.firmaRecibe || false, payload.observaciones || '',
        new Date().toISOString()
    ]);
    logAction(payload.entregadoPor || payload.userId, 'ENTREGA EPP', id, payload.trabajador);
    return { success: true, id: id };
    }

    function guardarMovimientoInventario(payload) {
    var auth = checkRole(payload.usuario || payload.userId, 'SUPERVISOR');
    if (!auth.allowed) return { success: false, error: auth.error };
    var id = genId('MINV-', 'Movimientos_Inventario');
    appendRow('Movimientos_Inventario', [
        id, payload.fecha || Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd'),
        payload.itemId, payload.tipo, Number(payload.cantidad) || 0,
        payload.referencia || '', payload.usuario || payload.userId,
        new Date().toISOString()
    ]);
    logAction(payload.usuario || payload.userId, 'MOV INVENTARIO', id, payload.tipo + ' ' + payload.cantidad);
    return { success: true, id: id };
    }
    function getInventario() { return getSheetData('Catalogo_Herramientas'); }
    function getValesHerramienta() { return getSheetData('Herramientas_Asignacion'); }
    function getMovimientosInventario() { return getSheetData('Movimientos_Inventario'); }
    function getPersonalCampo() { return getSheetData('Personal_Campo'); }

    function guardarPersonalCampo(payload) {
    var auth = checkRole(payload.creadoPor || payload.usuario, 'SEGURISTA');
    if (!auth.allowed) return auth;
    var id = genId('TRB', 'Personal_Campo');
    appendRow('Personal_Campo', [
        id,
        payload.nombre,
        payload.puesto || 'Ayudante General',
        payload.cuadrilla || '',
        payload.nss || '',
        payload.tipoSangre || '',
        payload.contactoEmergencia || '',
        payload.telefonoEmergencia || '',
        payload.fechaIngreso || Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd'),
        payload.activo !== false ? 'Activo' : 'Inactivo',
        payload.observaciones || '',
        new Date().toISOString()
    ]);
    logAction(payload.creadoPor || 'ADMIN', 'Registro personal campo', id, payload.nombre);
    return { success: true, id: id };
    }

    function actualizarPersonalCampo(payload) {
    var auth = checkRole(payload.usuario, 'SEGURISTA');
    if (!auth.allowed) return auth;
    var sheet = getSheet('Personal_Campo');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
        if (payload.nombre) sheet.getRange(i+1,2).setValue(payload.nombre);
        if (payload.puesto) sheet.getRange(i+1,3).setValue(payload.puesto);
        if (payload.cuadrilla) sheet.getRange(i+1,4).setValue(payload.cuadrilla);
        if (payload.nss) sheet.getRange(i+1,5).setValue(payload.nss);
        if (payload.tipoSangre) sheet.getRange(i+1,6).setValue(payload.tipoSangre);
        if (payload.contactoEmergencia) sheet.getRange(i+1,7).setValue(payload.contactoEmergencia);
        if (payload.telefonoEmergencia) sheet.getRange(i+1,8).setValue(payload.telefonoEmergencia);
        if (payload.activo !== undefined) sheet.getRange(i+1,10).setValue(payload.activo ? 'Activo' : 'Inactivo');
        if (payload.observaciones) sheet.getRange(i+1,11).setValue(payload.observaciones);
        logAction(payload.usuario, 'Actualizar personal campo', payload.id, payload.nombre || data[i][1]);
        return { success: true };
        }
    }
    return { success: false, error: 'Trabajador no encontrado' };
    }

    function getKardexTrabajador(trabajadorId) {
    var personal = getSheetData('Personal_Campo');
    var trab = personal.find(function(p) { return p.ID === trabajadorId; });
    var nombre = trab ? trab.Nombre : '';
    var esUsuario = false;
    
    // FIX: Declarar usuarios ANTES del if para que esté en scope siempre
    var usuarios = getSheetData('Usuarios');
    
    if (!trab) {
        var usr = usuarios.find(function(u) { return u.ID === trabajadorId; });
        if (usr) { nombre = usr.Nombre; esUsuario = true; trab = usr; }
        else return { success: false, error: 'Trabajador no encontrado: ' + trabajadorId };
    }
    
    // Cross-reference all modules
    var capacitaciones = [];
    try {
        var caps = getSheetData('Capacitaciones');
        capacitaciones = caps.filter(function(c) { return c.Usuario === trabajadorId; });
    } catch(e) {}
    
    var aptitudes = [];
    try {
        // FIX: Nombre correcto de la hoja (era 'Aptitud_Medica', debe ser 'Aptitudes_Medicas')
        var apts = getSheetData('Aptitudes_Medicas');
        aptitudes = apts.filter(function(a) { return a.Usuario === trabajadorId; });
    } catch(e) {}
    
    var entregasEPP = [];
    try {
        var epps = getSheetData('Entrega_EPP');
        entregasEPP = epps.filter(function(e) { return e.Trabajador === trabajadorId || e.Trabajador === nombre; });
    } catch(e) {}
    
    var asts = [];
    try {
        var astData = getSheetData('AST');
        asts = astData.filter(function(a) { return a.Lider === trabajadorId; });
    } catch(e) {}
    
    var platicas = [];
    try {
        var plData = getSheetData('Platicas_5min');
        platicas = plData.filter(function(p) {
        var asistentes = [];
        try { asistentes = JSON.parse(p.AsistentesJSON || p.Asistentes || '[]'); } catch(pe) {}
        return asistentes.indexOf(trabajadorId) >= 0 || asistentes.indexOf(nombre) >= 0;
        });
    } catch(e) {}
    
    var inspecciones = [];
    try {
        var inspData = getSheetData('Inspecciones');
        inspecciones = inspData.filter(function(i) { return i.Inspector === trabajadorId; });
    } catch(e) {}
    
    var incidentes = [];
    try {
        var incData = getSheetData('Incidentes');
        incidentes = incData.filter(function(inc) { 
        return (inc.Involucrados || '').indexOf(trabajadorId) >= 0 || (inc.Involucrados || '').indexOf(nombre) >= 0;
        });
    } catch(e) {}
    
    return {
        success: true,
        trabajador: trab,
        catalogoActivos: getSheetData('Catalogo_Activos'),
        almacen: getSheetData('Catalogo_Herramientas'),
        movimientosAlmacen: getSheetData('Movimientos_Herramientas'),
        herramientasAsignacion: getSheetData('Herramientas_Asignacion'),
        residuos: getSheetData('Residuos'),
        extintores: getSheetData('Extintores'),
        comisionSeguridad: getSheetData('Comision_Seguridad'),
        usuarios: usuarios,
        nombre: nombre,
        resumen: {
        capacitaciones: capacitaciones.length,
        aptitudes: aptitudes.length,
        entregasEPP: entregasEPP.length,
        asts: asts.length,
        platicas: platicas.length,
        inspecciones: inspecciones.length,
        incidentes: incidentes.length
        },
        capacitaciones: capacitaciones,
        aptitudes: aptitudes,
        entregasEPP: entregasEPP,
        asts: asts,
        platicas: platicas,
        inspecciones: inspecciones,
        incidentes: incidentes
    };
    }



    // ================================================================
    // GENERACIÓN DE REPORTES PDF — Diseño Profesional Premium
    // ================================================================
    function generarReportePDF(tipo) {
    var fechaHoy = Utilities.formatDate(new Date(), 'America/Mexico_City', 'dd/MM/yyyy');
    var fechaDoc = Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyyMMdd_HHmm');
    var doc = DocumentApp.create('DmV_Reporte_' + tipo + '_' + fechaDoc);
    var body = doc.getBody();
    
    // --- Configuración global ---
    body.setMarginTop(36);
    body.setMarginBottom(36);
    body.setMarginLeft(30);
    body.setMarginRight(30);
    
    var DARK = '#1A2029';
    var PRIMARY = '#0EA5E9';
    var GREEN = '#22C55E';
    var RED = '#EF4444';
    var ORANGE = '#F59E0B';
    var GRAY_LIGHT = '#F1F5F9';
    var GRAY_MED = '#94A3B8';
    var WHITE = '#FFFFFF';
    
    // Helper: styled paragraph
    function styledPara(text, size, color, bold, align) {
        var p = body.appendParagraph(text);
        p.editAsText().setFontSize(size).setFontFamily('Arial').setForegroundColor(color || DARK);
        if (bold) p.editAsText().setBold(true);
        if (align) p.setAlignment(align);
        return p;
    }
    
    // Helper: section header with colored bar
    function sectionHeader(title, number) {
        body.appendParagraph(''); // spacing
        var headerTable = body.appendTable();
        headerTable.setBorderWidth(0);
        var row = headerTable.appendTableRow();
        var numCell = row.appendTableCell(number || '●');
        numCell.setBackgroundColor(PRIMARY).setWidth(40).setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(8).setPaddingRight(8);
        numCell.getChild(0).asParagraph().editAsText().setFontSize(11).setFontFamily('Arial').setBold(true).setForegroundColor(WHITE);
        numCell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        var titleCell = row.appendTableCell(' ' + title.toUpperCase());
        titleCell.setBackgroundColor(DARK).setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(12).setPaddingRight(8);
        titleCell.getChild(0).asParagraph().editAsText().setFontSize(11).setFontFamily('Arial').setBold(true).setForegroundColor(WHITE);
        // Remove first empty row if exists
        if (headerTable.getNumRows() > 1) { try { headerTable.removeRow(0); } catch(e){} }
        body.appendParagraph('').editAsText().setFontSize(4); // small spacing
        return headerTable;
    }
    
    // Helper: data table with alternating rows
    function dataTable(headers, rows, colWidths) {
        if (rows.length === 0) {
        styledPara('Sin registros.', 9, GRAY_MED, false);
        return null;
        }
        var table = body.appendTable();
        table.setBorderWidth(0.5).setBorderColor('#CBD5E1');
        
        // Header row
        var hRow = table.appendTableRow();
        headers.forEach(function(h, i) {
        var cell = hRow.appendTableCell(h);
        cell.setBackgroundColor(DARK).setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(6).setPaddingRight(6);
        if (colWidths && colWidths[i]) cell.setWidth(colWidths[i]);
        cell.getChild(0).asParagraph().editAsText().setFontSize(8).setFontFamily('Arial').setBold(true).setForegroundColor(WHITE);
        });
        
        // Data rows
        rows.forEach(function(rowData, idx) {
        var dRow = table.appendTableRow();
        var bgColor = idx % 2 === 0 ? WHITE : GRAY_LIGHT;
        rowData.forEach(function(val, i) {
            var cell = dRow.appendTableCell(String(val || '—'));
            cell.setBackgroundColor(bgColor).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
            cell.getChild(0).asParagraph().editAsText().setFontSize(8).setFontFamily('Arial').setForegroundColor(DARK);
        });
        });
        
        // Remove first empty row
        if (table.getNumRows() > 0 && table.getRow(0).getNumCells() === 1 && table.getRow(0).getCell(0).getText() === '') {
        try { table.removeRow(0); } catch(e){}
        }
        return table;
    }
    
    // Helper: KPI card row (2-column summary)
    function kpiRow(label, value, color) {
        var t = body.appendTable();
        t.setBorderWidth(0);
        var r = t.appendTableRow();
        var lCell = r.appendTableCell(label);
        lCell.setWidth(300).setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(8);
        lCell.getChild(0).asParagraph().editAsText().setFontSize(9).setFontFamily('Arial').setForegroundColor(GRAY_MED);
        var vCell = r.appendTableCell(String(value));
        vCell.setPaddingTop(3).setPaddingBottom(3).setPaddingRight(8);
        vCell.getChild(0).asParagraph().editAsText().setFontSize(11).setFontFamily('Arial').setBold(true).setForegroundColor(color || DARK);
        vCell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
        if (t.getNumRows() > 1) { try { t.removeRow(0); } catch(e){} }
        return t;
    }
    
    // ===================================================================
    //  PORTADA
    // ===================================================================
    // Línea decorativa superior
    var lineTop = body.appendTable();
    lineTop.setBorderWidth(0);
    var lineRow = lineTop.appendTableRow();
    var lineCell = lineRow.appendTableCell(' ');
    lineCell.setBackgroundColor(PRIMARY).setPaddingTop(3).setPaddingBottom(3);
    if (lineTop.getNumRows() > 1) { try { lineTop.removeRow(0); } catch(e){} }
    
    body.appendParagraph('').editAsText().setFontSize(6);
    
    // Logo
    insertLogo(body, 150, 48);
    body.appendParagraph('').editAsText().setFontSize(6);
    
    // Título
    styledPara('DmV CONTROL v2.0', 22, DARK, true, DocumentApp.HorizontalAlignment.CENTER);
    styledPara('SASISOPA · GAS NATURAL', 11, PRIMARY, true, DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('').editAsText().setFontSize(6);
    
    // Línea divisoria
    var lineMid = body.appendTable();
    lineMid.setBorderWidth(0);
    var lmRow = lineMid.appendTableRow();
    var lmCell = lmRow.appendTableCell(' ');
    lmCell.setBackgroundColor(DARK).setPaddingTop(1).setPaddingBottom(1);
    if (lineMid.getNumRows() > 1) { try { lineMid.removeRow(0); } catch(e){} }
    
    body.appendParagraph('').editAsText().setFontSize(6);
    styledPara('REPORTE ' + tipo.toUpperCase(), 18, DARK, true, DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('').editAsText().setFontSize(3);
    styledPara('NOM-020-ASEA · Transporte de Gas Natural por Ducto', 8, GRAY_MED, false, DocumentApp.HorizontalAlignment.CENTER);
    styledPara('NOM-030-STPS · Servicios Preventivos de Seguridad y Salud', 8, GRAY_MED, false, DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('').editAsText().setFontSize(10);
    
    // Info box portada
    var infoTable = body.appendTable();
    infoTable.setBorderWidth(0.5).setBorderColor('#CBD5E1');
    var iRow1 = infoTable.appendTableRow();
    var ic1 = iRow1.appendTableCell('Fecha de emisión');
    ic1.setBackgroundColor(GRAY_LIGHT).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(12).setWidth(200);
    ic1.getChild(0).asParagraph().editAsText().setFontSize(8).setFontFamily('Arial').setForegroundColor(GRAY_MED);
    var ic2 = iRow1.appendTableCell(fechaHoy);
    ic2.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(12);
    ic2.getChild(0).asParagraph().editAsText().setFontSize(8).setFontFamily('Arial').setBold(true).setForegroundColor(DARK);
    
    var iRow2 = infoTable.appendTableRow();
    var ic3 = iRow2.appendTableCell('Tipo de reporte');
    ic3.setBackgroundColor(GRAY_LIGHT).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(12);
    ic3.getChild(0).asParagraph().editAsText().setFontSize(8).setFontFamily('Arial').setForegroundColor(GRAY_MED);
    var ic4 = iRow2.appendTableCell(tipo.charAt(0).toUpperCase() + tipo.slice(1));
    ic4.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(12);
    ic4.getChild(0).asParagraph().editAsText().setFontSize(8).setFontFamily('Arial').setBold(true).setForegroundColor(DARK);
    
    var iRow3 = infoTable.appendTableRow();
    var ic5 = iRow3.appendTableCell('Generado por');
    ic5.setBackgroundColor(GRAY_LIGHT).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(12);
    ic5.getChild(0).asParagraph().editAsText().setFontSize(8).setFontFamily('Arial').setForegroundColor(GRAY_MED);
    var ic6 = iRow3.appendTableCell('DmV Control v2.0 — Sistema Automatizado');
    ic6.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(12);
    ic6.getChild(0).asParagraph().editAsText().setFontSize(8).setFontFamily('Arial').setBold(true).setForegroundColor(DARK);
    
    // Remove first empty row
    if (infoTable.getNumRows() > 3) { try { infoTable.removeRow(0); } catch(e){} }
    
    body.appendPageBreak();
    
    // ===================================================================
    //  RESUMEN EJECUTIVO
    // ===================================================================
    var dashboard = getDashboard();
    sectionHeader('RESUMEN EJECUTIVO', '01');
    
    styledPara('Indicadores Clave de Desempeño', 10, DARK, true);
    body.appendParagraph('').editAsText().setFontSize(2);
    
    // KPI Summary table
    var kpiTable = body.appendTable();
    kpiTable.setBorderWidth(0.5).setBorderColor('#CBD5E1');
    
    var kpiHeaders = [['INDICADOR', 'VALOR', 'ESTADO']];
    var kpiHRow = kpiTable.appendTableRow();
    ['INDICADOR', 'VALOR', 'ESTADO'].forEach(function(h) {
        var c = kpiHRow.appendTableCell(h);
        c.setBackgroundColor(DARK).setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(8).setPaddingRight(8);
        c.getChild(0).asParagraph().editAsText().setFontSize(8).setFontFamily('Arial').setBold(true).setForegroundColor(WHITE);
    });
    
    var kpiData = [
        ['Km Programados', String(dashboard.totalKmProg) + ' km', '—'],
        ['Km Ejecutados', String(dashboard.kmEjec) + ' km', dashboard.pctAvance >= 80 ? '✓ EN META' : '⚠ POR DEBAJO'],
        ['Avance General', String(dashboard.pctAvance) + '%', dashboard.pctAvance >= 80 ? '✓ BUENO' : dashboard.pctAvance >= 50 ? '⚠ REGULAR' : '✕ CRÍTICO'],
        ['Hallazgos Abiertos', String(dashboard.hallAbiertos), dashboard.hallAbiertos === 0 ? '✓ CERO' : '⚠ PENDIENTES'],
        ['Stop Work Activos', String(dashboard.swActivos), dashboard.swActivos === 0 ? '✓ NINGUNO' : '✕ ACTIVO']
    ];
    
    kpiData.forEach(function(row, idx) {
        var kRow = kpiTable.appendTableRow();
        var bg = idx % 2 === 0 ? WHITE : GRAY_LIGHT;
        row.forEach(function(val, ci) {
        var c = kRow.appendTableCell(val);
        c.setBackgroundColor(bg).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(8).setPaddingRight(8);
        var color = ci === 2 ? (val.indexOf('✓') >= 0 ? GREEN : val.indexOf('✕') >= 0 ? RED : ORANGE) : DARK;
        c.getChild(0).asParagraph().editAsText().setFontSize(9).setFontFamily('Arial').setForegroundColor(color);
        if (ci === 1) c.getChild(0).asParagraph().editAsText().setBold(true);
        });
    });
    if (kpiTable.getNumRows() > kpiData.length + 1) { try { kpiTable.removeRow(0); } catch(e){} }
    
    // ===================================================================
    //  ACTIVIDADES
    // ===================================================================
    sectionHeader('DETALLE DE ACTIVIDADES', '02');
    
    var acts = getActividades();
    var usuarios = getSheetData('Usuarios');
    var tramos = getSheetData('Tramos');
    
    function getNombre(id) {
        var u = usuarios.find(function(x) { return x.ID === id; });
        return u ? u.Nombre : id || '—';
    }
    function getTramoNombre(id) {
        var t = tramos.find(function(x) { return x.ID === id; });
        return t ? t.Nombre : id || '—';
    }
    
    styledPara(acts.length + ' actividades registradas', 9, GRAY_MED, false);
    body.appendParagraph('').editAsText().setFontSize(2);
    
    dataTable(
        ['ID', 'Tipo', 'Tramo', 'Líder', 'Fecha', 'Estado', 'AST'],
        acts.map(function(a) {
        return [a.ID, a.Tipo, getTramoNombre(a.Tramo), getNombre(a.Lider), a.Fecha, a.Estado, a.ASTId || '—'];
        }),
        [45, 65, 120, 90, 65, 60, 50]
    );
    
    // ===================================================================
    //  PATRULLAJES
    // ===================================================================
    var patrullajes = getSheetData('Patrullajes');
    if (patrullajes.length > 0) {
        sectionHeader('PATRULLAJES REALIZADOS', '03');
        styledPara(patrullajes.length + ' patrullajes registrados', 9, GRAY_MED, false);
        body.appendParagraph('').editAsText().setFontSize(2);
        
        dataTable(
        ['ID', 'Tramo', 'Líder', 'Fecha', 'KM Ini', 'KM Fin', 'Estatus', 'Hallazgo'],
        patrullajes.map(function(p) {
            return [p.ID, getTramoNombre(p.Tramo), getNombre(p.Lider), p.Fecha, p.KmInicio, p.KmFin, p.Estatus, p.Hallazgo || 'Normal'];
        }),
        [45, 100, 80, 65, 40, 40, 60, 60]
        );
    }
    
    // ===================================================================
    //  HALLAZGOS
    // ===================================================================
    var hallazgos = getSheetData('Hallazgos');
    if (hallazgos.length > 0) {
        sectionHeader('HALLAZGOS', '04');
        var abiertos = hallazgos.filter(function(h) { return h.Estado === 'Abierto'; }).length;
        var cerrados = hallazgos.filter(function(h) { return h.Estado === 'Cerrado'; }).length;
        styledPara(hallazgos.length + ' hallazgos totales · ' + abiertos + ' abiertos · ' + cerrados + ' cerrados', 9, GRAY_MED, false);
        body.appendParagraph('').editAsText().setFontSize(2);
        
        dataTable(
        ['ID', 'Tipo', 'Descripción', 'Tramo', 'KM', 'Estado', 'Fecha'],
        hallazgos.map(function(h) {
            return [h.ID, h.Tipo, (h.Descripcion || '').substring(0, 40), getTramoNombre(h.Tramo), h.Km, h.Estado, h.Fecha];
        }),
        [40, 60, 130, 80, 35, 50, 60]
        );
    }
    
    // ===================================================================
    //  SASISOPA
    // ===================================================================
    var sasisopa = getSheetData('SASISOPA');
    if (sasisopa.length > 0) {
        sectionHeader('SASISOPA — CUMPLIMIENTO', '05');
        var avgCum = Math.round(sasisopa.reduce(function(s, e) { return s + Number(e.Cumplimiento || 0); }, 0) / sasisopa.length);
        styledPara('Promedio general de cumplimiento: ' + avgCum + '%', 10, avgCum >= 85 ? GREEN : avgCum >= 70 ? ORANGE : RED, true);
        body.appendParagraph('').editAsText().setFontSize(2);
        
        dataTable(
        ['#', 'Elemento SASISOPA', 'Cumplimiento', 'Estado'],
        sasisopa.map(function(s, i) {
            var cum = Number(s.Cumplimiento || 0);
            var estado = cum >= 85 ? '✓ Cumple' : cum >= 70 ? '⚠ Parcial' : '✕ Deficiente';
            return [String(i + 1), s.Elemento, cum + '%', estado];
        }),
        [25, 250, 65, 80]
        );
    }
    
    // ===================================================================
    //  PSST — Capacitaciones y Aptitudes
    // ===================================================================
    var caps = getSheetData('Capacitaciones');
    var apts = getSheetData('Aptitudes_Medicas');
    if (caps.length > 0 || apts.length > 0) {
        sectionHeader('PSST — PERSONAL', '06');
        
        if (caps.length > 0) {
        styledPara('Capacitaciones (' + caps.length + ' registros)', 10, DARK, true);
        body.appendParagraph('').editAsText().setFontSize(2);
        dataTable(
            ['ID', 'Trabajador', 'Tipo', 'Curso', 'Vigencia', 'DC3', 'Estatus'],
            caps.map(function(c) {
            return [c.ID, getNombre(c.Usuario), c.Tipo, c.Nombre, c.Vigencia, c.DC3 ? '✓' : '✕', c.Estatus];
            }),
            [35, 80, 60, 90, 60, 30, 50]
        );
        body.appendParagraph('').editAsText().setFontSize(4);
        }
        
        if (apts.length > 0) {
        styledPara('Aptitudes Médicas (' + apts.length + ' registros)', 10, DARK, true);
        body.appendParagraph('').editAsText().setFontSize(2);
        dataTable(
            ['ID', 'Trabajador', 'Dictamen', 'Vigencia', 'Médico', 'Estatus'],
            apts.map(function(a) {
            return [a.ID, getNombre(a.Usuario), a.Tipo, a.Vigencia, a.Medico, a.Estatus];
            }),
            [40, 90, 80, 65, 80, 55]
        );
        }
    }
    
    // ===================================================================
    //  AST
    // ===================================================================
    var astData = getSheetData('AST');
    if (astData.length > 0) {
        sectionHeader('AST — ANÁLISIS DE SEGURIDAD', '07');
        styledPara(astData.length + ' análisis de seguridad en el trabajo registrados', 9, GRAY_MED, false);
        body.appendParagraph('').editAsText().setFontSize(2);
        
        dataTable(
        ['ID', 'Actividad', 'Líder', 'Fecha', 'Tipo', 'Peligros', 'Firma'],
        astData.map(function(a) {
            var pCount = '0';
            try { pCount = String(JSON.parse(a.PeligrosJSON || '[]').length); } catch(e) {}
            return [a.ID, a.Actividad, getNombre(a.Lider), a.Fecha, a.ActividadTipo || '—', pCount, a.Firma ? '✓' : '✕'];
        }),
        [45, 50, 80, 65, 65, 45, 35]
        );
    }
    
    // ===================================================================
    //  EVIDENCIA FOTOGRÁFICA
    // ===================================================================
    var evidencias = getSheetData('Evidencias');
    if (evidencias.length > 0) {
        sectionHeader('EVIDENCIA FOTOGRÁFICA', '08');
        styledPara(evidencias.length + ' evidencia(s) registrada(s)', 9, GRAY_MED, false);
        body.appendParagraph('').editAsText().setFontSize(4);
        
        evidencias.forEach(function(ev, idx) {
        try {
            if (ev.ArchivoID) {
            var imgBase64 = getImgBase64(ev.ArchivoID);
            if (imgBase64) {
                var b64data = imgBase64.split(',')[1];
                var contentType = imgBase64.split(';')[0].split(':')[1];
                var blob = Utilities.newBlob(Utilities.base64Decode(b64data), contentType, 'evidencia_' + idx + '.jpg');
                var img = body.appendImage(blob);
                // Scale proportionally, max width 450
                img.setWidth(450);
                img.setHeight(Math.round(450 * 0.65));
            }
            var caption = 'Evidencia ' + (idx + 1) + ' · Patrullaje: ' + (ev.Patrullaje || '—') + ' · ' + (ev.Timestamp || '');
            styledPara(caption, 8, GRAY_MED, false, DocumentApp.HorizontalAlignment.CENTER);
            body.appendParagraph('').editAsText().setFontSize(4);
            }
        } catch (err) {
            styledPara('[Imagen no disponible: ' + (ev.ArchivoID || '') + ']', 8, RED, false);
            Logger.log('Error insertando imagen: ' + err.message);
        }
        });
    }
    
    // ===================================================================
    //  NOM — MARCO NORMATIVO Y CUMPLIMIENTO
    // ===================================================================
    var nomData = getSheetData('NOM_Matriz');
    if (nomData.length > 0) {
        sectionHeader('MARCO NORMATIVO — CUMPLIMIENTO NOM', '09');
        var avgNOM = Math.round(nomData.reduce(function(s, n) { return s + Number(n.Cumplimiento || 0); }, 0) / nomData.length);
        styledPara('Promedio general NOM: ' + avgNOM + '%', 10, avgNOM >= 85 ? GREEN : avgNOM >= 50 ? ORANGE : RED, true);
        body.appendParagraph('').editAsText().setFontSize(2);
        
        dataTable(
        ['NOM', 'Título', 'Módulo DmV', '%', 'Estado'],
        nomData.map(function(n) {
            var cum = Number(n.Cumplimiento || 0);
            return [n.NOM, n.Titulo || '', n.Modulo || '', cum + '%', n.Estado || 'Pendiente'];
        }),
        [85, 130, 100, 35, 55]
        );
        
        body.appendParagraph('').editAsText().setFontSize(4);
        styledPara('Detalle de evidencia digital:', 9, DARK, true);
        body.appendParagraph('').editAsText().setFontSize(2);
        
        dataTable(
        ['NOM', 'Evidencia Requerida', 'Evidencia en DmV Control'],
        nomData.map(function(n) {
            return [n.NOM, n.EvidenciaRequerida || '—', n.EvidenciaDigital || 'Sin datos'];
        }),
        [85, 170, 170]
        );
    }
    
    // ===================================================================
    //  FIRMAS
    // ===================================================================
    body.appendParagraph('').editAsText().setFontSize(6);
    
    appendFirmasEjecutivas(body,
        { cargo: 'RESPONSABLE DE SEGURIDAD', nombre: 'Nombre: _______________', empresa: 'GENCO' },
        { cargo: 'SUPERVISOR DE CAMPO', nombre: 'Nombre: _______________', empresa: 'Gasoducto de Aguaprieta' },
        fechaHoy
    );
    
    // Footer
    body.appendParagraph('').editAsText().setFontSize(6);
    var footLine = body.appendTable();
    footLine.setBorderWidth(0);
    var flRow = footLine.appendTableRow();
    var flCell = flRow.appendTableCell(' ');
    flCell.setBackgroundColor(PRIMARY).setPaddingTop(2).setPaddingBottom(2);
    if (footLine.getNumRows() > 1) { try { footLine.removeRow(0); } catch(e){} }
    styledPara('DmV Control v2.0 · SASISOPA Gas · Generado automáticamente el ' + fechaHoy, 7, GRAY_MED, false, DocumentApp.HorizontalAlignment.CENTER);
    styledPara('Este documento es confidencial y de uso exclusivo del operador del ducto.', 7, GRAY_MED, false, DocumentApp.HorizontalAlignment.CENTER);
    
    // ===================================================================
    //  GUARDAR PDF
    // ===================================================================
    doc.saveAndClose();
    
    var docFile = DriveApp.getFileById(doc.getId());
    var pdfBlob = docFile.getAs('application/pdf');
    pdfBlob.setName(doc.getName() + '.pdf');
    
    var pdfFolder;
    try {
        pdfFolder = DriveApp.getFolderById(ROOT_FOLDER_ID).getFoldersByName('Reportes_PDF').next();
    } catch(e) {
        pdfFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);
    }
    var pdfFile = pdfFolder.createFile(pdfBlob);
    
    // Delete temp doc
    docFile.setTrashed(true);
    
    logAction('SISTEMA', 'Generar reporte PDF', tipo, pdfFile.getUrl());
    
    return { success: true, url: pdfFile.getUrl(), id: pdfFile.getId() };
    }

    // ================================================================
    // REPORTE DE AVANCE — PARA CLIENTE
    // 100% data-driven: lee Actividades, Tramos, Patrullajes, Hallazgos,
    // SASISOPA, Capacitaciones, etc. y calcula todo en tiempo real.
    // Gantt: agrupa actividades por Tipo y calcula fechas reales.
    // ================================================================
    // ================================================================
    // ================================================================
    // GENERAR_MANUAL_SISTEMA — Crea un PDF ejecutivo con el manual
    // ================================================================
    // ================================================================
    // GENERAR_MANUAL_SISTEMA — Crea el LIBRO TÉCNICO Y MANIFIESTO
    // ================================================================
    // ================================================================
    // GENERAR_MANUAL_SISTEMA — Crea el LIBRO MAESTRO (ALTA DIRECCIÓN)
    // ================================================================
    function generarManualSistema() {
    var fechaHoy = Utilities.formatDate(new Date(), 'America/Mexico_City', 'dd/MM/yyyy');
    var doc = DocumentApp.create('LIBRO_MAESTRO_AVA_SGO_' + Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyyMMdd'));
    var body = doc.getBody();
    
    // Estilos
    var DARK='#1A2029', BLUE='#0D4E8B', GRAY_M='#8A97A8', WHITE='#FFFFFF', ACCENT='#00579E';
    function sP(t,s,c,b,a){var p=body.appendParagraph(t?String(t):'');p.editAsText().setFontSize(s).setFontFamily('Arial').setForegroundColor(c||DARK);if(b)p.editAsText().setBold(true);if(a)p.setAlignment(a);return p;}
    function spc(s){body.appendParagraph('').editAsText().setFontSize(s||4);}
    function addSig(role1,name1,role2,name2){spc(12); appendFirmasEjecutivas(body,{cargo:role1,nombre:name1},{cargo:role2,nombre:name2},fechaHoy);}
    function h1(t){spc(14); body.appendPageBreak(); return sP(t.toUpperCase(), 16, BLUE, true, DocumentApp.HorizontalAlignment.LEFT);}
    function h2(t){spc(8); return sP(t, 12, DARK, true);}
    function p(t){return sP(t, 10, DARK, false, DocumentApp.HorizontalAlignment.JUSTIFY);}
    function diag(txt){var d=sP(txt, 9, ACCENT, false, DocumentApp.HorizontalAlignment.CENTER); d.setItalic(true); return d;}

    // --- PORTADA ---
    insertLogo(body, 180, 64);
    spc(60);
    sP('SGO - SASISOPA GENCO v2.0', 14, GRAY_M, true, DocumentApp.HorizontalAlignment.CENTER);
    sP('LIBRO MAESTRO: SISTEMA DE GESTIÓN TECNOLÓGICA', 26, BLUE, true, DocumentApp.HorizontalAlignment.CENTER);
    spc(24);
    var line=body.appendTable();line.setBorderWidth(0);line.appendTableRow().appendTableCell(' ').setBackgroundColor(DARK).setPaddingTop(2).setPaddingBottom(2);
    spc(24);
    sP('GESTIÓN ESTRATÉGICA · ALTA DIRECCIÓN · CUMPLIMIENTO REGULATORIO', 10, GRAY_M, true, DocumentApp.HorizontalAlignment.CENTER);
    spc(80);
    sP('Documentación Maestra del Sistema de Gestión Integral', 11, DARK, false, DocumentApp.HorizontalAlignment.CENTER);
    spc(60);
    sP('GENCO · AGUAPRIETA GASODUCTO · MÉXICO 2026', 10, BLUE, true, DocumentApp.HorizontalAlignment.CENTER);
    
    // --- CAPÍTULO 1: MANIFIESTO Y ESTRATEGIA ---
    h1('CAPÍTULO 1: MANIFIESTO Y VISIÓN ESTRATÉGICA (MBA PERSPECTIVE)');
    h2('1.1 El Imperativo Digital de la Alta Dirección');
    p('Desde la perspectiva de la Alta Dirección, el SGO-SASISOPA GENCO es un Activo Estratégico que mitiga riesgos operativos y financieros. El manifiesto se basa en la optimización de procesos y el valor probatorio inmutable.');
    diag('[ VISIÓN ESTRATÉGICA ] --> [ ESTRATEGIA PAPEL CERO ] --> [ IMPLEMENTACIÓN SGO ] --> [ VALOR: TRAZABILIDAD ]');

    // --- CAPÍTULO 2: ORGANIGRAMA DEL PROYECTO ---
    h1('CAPÍTULO 2: ORGANIGRAMA Y ESTRUCTURA DEL PROYECTO');
    h2('2.1 Estructura Organizacional de GENCO');
    p('La gestión tecnológica eficaz se define por una estructura corporativa sólida que garantiza la segregación de funciones.');
    spc(10);
    var ot = body.appendTable(); ot.setBorderWidth(0.5).setBorderColor('#ACBAD3');
    var headerRow = ot.appendTableRow();
    var hc1 = headerRow.appendTableCell('ID / POSICIÓN'); hc1.setBackgroundColor('#1A2029').setPaddingLeft(10);
    hc1.getChild(0).asParagraph().editAsText().setFontSize(9).setForegroundColor(WHITE).setBold(true);
    var hc2 = headerRow.appendTableCell('NOMBRE DEL TITULAR'); hc2.setBackgroundColor('#1A2029').setPaddingLeft(10);
    hc2.getChild(0).asParagraph().editAsText().setFontSize(9).setForegroundColor(WHITE).setBold(true);
    
    var roles = [
        ['DIRECCIÓN', 'Ing. Abraham Valdez'],
        ['GERENCIA GENERAL', 'Lic. Karina Flores'],
        ['GERENCIA OPERACIONES', 'Ing. Efrain Reyes'],
        ['GERENCIA DE SST', 'Lic. Sergio A. Jimenez Rodriguez']
    ];
    
    roles.forEach(function(r){
        var row = ot.appendTableRow();
        var c1 = row.appendTableCell(r[0]); c1.setBackgroundColor('#F8FAFC');
        c1.getChild(0).asParagraph().editAsText().setFontSize(9).setForegroundColor('#1A2029').setBold(true);
        var c2 = row.appendTableCell(r[1]);
        c2.getChild(0).asParagraph().editAsText().setFontSize( r[1].length > 25 ? 7 : 9 ).setForegroundColor('#1A2029').setBold(false);
    });
    ot.setColumnWidth(0, 160); ot.setColumnWidth(1, 300);

    // --- CAPÍTULO 3: MARCO NORMATIVO ---
    h1('CAPÍTULO 3: MARCO NORMATIVO (NOM-020-ASEA)');
    h2('3.1 Alineación y Justificación Técnica');
    p('El sistema responde a la NOM-020-ASEA-2019, transformando requisitos legales en flujos automatizados de integridad mecánica y patrullaje de ductos.');
    diag('NOM-020-ASEA --> MÓDULO OPERATIVO KP --> EVIDENCIA FOTOGRÁFICA --> AUDITORÍA SASISOPA --> CUMPLIMIENTO LEGAL');

    // --- CAPÍTULO 4: FLUJOS OPERATIVOS ---
    h1('CAPÍTULO 4: FLUJOS DE INFORMACIÓN OPERATIVA');
    h2('4.1 Ciclo de Gestión de Actividades');
    p('La integración tecnológica se basa en el flujo bidireccional entre la planeación estratégica y la ejecución distribuida.');
    diag('PLANEACIÓN (Admin) --> ASIGNACIÓN (Sys) --> CAPTURA (Líder) --> VALIDACIÓN (Segurista) --> KPI (Dash)');

    // --- CAPÍTULO 5: METODOLOGÍA SASISOPA ---
    h1('CAPÍTULO 5: METODOLOGÍA SASISOPA Y AUDITORÍAS');
    h2('5.1 Elemento 16: Auditorías y Verificación');
    p('El sistema automatiza el Elemento 16 a través de un motor de verificación constante que detecta desviaciones preventivas mediante registros digitales de firmas.');
    diag('PLAN AUDITORÍA --> LISTA VERIFICACIÓN --> CAPTURA IN SITU --> RESULTADO AUTOMÁTICO % --> ACCIONES CORRECTIVAS');

    // --- CAPÍTULO 6: DESARROLLO POR MÓDULOS ---
    h1('CAPÍTULO 6: DESARROLLO INTEGRAL POR MÓDULOS');
    var mods = [
        ['OPERATIVO', 'Ducto, KP, Hallazgos fotográficos y Metas.'],
        ['HSE / SEGURIDAD', 'AST, Permisos de Riesgo y Stop Work.'],
        ['SALUD', 'NOM-030, Exámenes médicos y DC-3.'],
        ['GOBERNANZA', 'Dashboards, Matriz ASEA y Auditorías.']
    ];
    var mt = body.appendTable(); mt.setBorderWidth(0.5).setBorderColor('#E2E8F0');
    mods.forEach(function(m){
        var r = mt.appendTableRow();
        var mc1 = r.appendTableCell(m[0]); mc1.setBackgroundColor('#F8FAFC').setWidth(140);
        mc1.getChild(0).asParagraph().editAsText().setFontSize(9).setForegroundColor('#1A2029').setBold(true);
        r.appendTableCell(m[1]).getChild(0).asParagraph().editAsText().setFontSize(8).setForegroundColor('#475569');
    });
    diag('OPERATIVO <--> HSE <--> SALUD <--> GOBERNANZA (Integración 360°)');

    // --- CAPÍTULO 7: ANEXO SASISOPA (18 ELEMENTOS) ---
    h1('CAPÍTULO 7: ANEXO - INTEGRACIÓN SASISOPA (18 ELEMENTOS)');
    h2('7.1 Mapeo de Documentación Oficial');
    p('El listado maestro de documentos SA-SASISOPA-GENCO se encuentra parametrizado para su gestión digital.');
    
    var sasisopa = [
        ['I. Política', 'SAS-01-001 / 002'], ['II. Riesgos', 'SAS-02-001 / 002 / 003'],
        ['III. Legales', 'SAS-03-001 / 002'], ['IV. Competencia', 'SAS-04-001 / 005'],
        ['V. Comunicación', 'SAS-18-001 / 004'], ['VI. Control Doc.', 'SAS-14-001 / 002'],
        ['VII. Mejora Seg.', 'SAS-07-001 / 005'], ['VIII. Procesos', 'SAS-17-001 / 002'],
        ['IX. Int. Mecánica', 'SAS-10-001 / 003'], ['X. Emergencias', 'SAS-09-001 / 004'],
        ['XI. Contratistas', 'SAS-06-001 / 003'], ['XII. Incidentes', 'SAS-12-001 / 003'],
        ['XIII. Auditorías', 'SAS-11-001 / 004'], ['XIV. Rev. Dir.', 'SAS-13-001 / 002'],
        ['XV. Acc. Corr.', 'SAS-15-001 / 002'], ['XVI. Op. Segura', 'SAS-16-001 / 005'],
        ['XVII. Patrullajes', 'SAS-05-001 / 003'], ['XVIII. Cambios', 'SAS-08-001 / 002']
    ];
    
    var st = body.appendTable(); st.setBorderWidth(0.5).setBorderColor('#CBD5E1');
    sasisopa.forEach(function(s){
        var row = st.appendTableRow();
        var sc1 = row.appendTableCell(s[0]); sc1.setBackgroundColor('#F1F5F9').setWidth(150);
        sc1.getChild(0).asParagraph().editAsText().setFontSize(8).setForegroundColor('#1A2029').setBold(true);
        row.appendTableCell(s[1]).getChild(0).asParagraph().editAsText().setFontSize(8).setForegroundColor('#64748B');
    });

    spc(40);
    addSig('ARCH. SISTEMA / LAE. MBA', 'SERGIO ADALBERTO JIMENEZ RODRIGUEZ', 'REPRESENTANTE TÉCNICO', '_______________');

    doc.saveAndClose();
    var docFile=DriveApp.getFileById(doc.getId());
    var pdfBlob=docFile.getAs('application/pdf');
    // Nombre con Timestamp para evitar caché del navegador
    var timestamp = Utilities.formatDate(new Date(), 'America/Mexico_City', 'HHmm');
    pdfBlob.setName('LIBRO_MAESTRO_AVA_' + timestamp + '.pdf');
    
    var pdfFolder;
    try{pdfFolder=DriveApp.getFolderById(ROOT_FOLDER_ID).getFoldersByName('Reportes_PDF').next();}
    catch(e){pdfFolder=DriveApp.getFolderById(ROOT_FOLDER_ID);}
    var pdfFile=pdfFolder.createFile(pdfBlob);
    docFile.setTrashed(true);
    return {success:true, url:pdfFile.getUrl()};
    }

    function generarReporteAvance() {
    var fechaHoy = Utilities.formatDate(new Date(), 'America/Mexico_City', 'dd/MM/yyyy');
    var fechaDoc = Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyyMMdd_HHmm');
    var doc = DocumentApp.create('DmV_Avance_Cliente_' + fechaDoc);
    var body = doc.getBody();
    body.setMarginTop(30); body.setMarginBottom(30);
    body.setMarginLeft(28); body.setMarginRight(28);

    var DARK='#0C1B2E', BLUE='#0D4E8B', ACCENT='#E8A838', GREEN='#16A34A',
        RED='#DC2626', ORANGE='#EA580C', GRAY_L='#F4F6F9', GRAY_M='#8A97A8',
        WHITE='#FFFFFF', BORDER='#E2E8F0';

    // ── helpers ──
    function sP(t,s,c,b,a){var p=body.appendParagraph(t);p.editAsText().setFontSize(s).setFontFamily('Arial').setForegroundColor(c||DARK);if(b)p.editAsText().setBold(true);if(a)p.setAlignment(a);return p;}
    function secH(title,num){
        body.appendParagraph('');
        var tb=body.appendTable();tb.setBorderWidth(0);var r=tb.appendTableRow();
        var nc=r.appendTableCell(num);nc.setBackgroundColor(BLUE).setWidth(36).setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(8).setPaddingRight(8);
        nc.getChild(0).asParagraph().editAsText().setFontSize(10).setFontFamily('Arial').setBold(true).setForegroundColor(WHITE);
        nc.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        var tc=r.appendTableCell(' '+title.toUpperCase());tc.setBackgroundColor(DARK).setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(12);
        tc.getChild(0).asParagraph().editAsText().setFontSize(10).setFontFamily('Arial').setBold(true).setForegroundColor(WHITE);
        if(tb.getNumRows()>1){try{tb.removeRow(0);}catch(e){}}
    }
    function hdrRow(tbl,headers,widths){
        var r=tbl.appendTableRow();
        headers.forEach(function(h,i){
        var c=r.appendTableCell(h);c.setBackgroundColor(DARK).setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(6).setPaddingRight(4);
        if(widths&&widths[i])c.setWidth(widths[i]);
        c.getChild(0).asParagraph().editAsText().setFontSize(7).setFontFamily('Arial').setBold(true).setForegroundColor(WHITE);
        c.getChild(0).asParagraph().setAlignment(i>=2?DocumentApp.HorizontalAlignment.CENTER:DocumentApp.HorizontalAlignment.LEFT);
        });
    }
    function dataCell(row,val,bg,color,bold,center){
        var c=row.appendTableCell(String(val));c.setBackgroundColor(bg||WHITE).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(4);
        c.getChild(0).asParagraph().editAsText().setFontSize(8).setFontFamily('Arial').setForegroundColor(color||DARK);
        if(bold)c.getChild(0).asParagraph().editAsText().setBold(true);
        if(center)c.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        return c;
    }
    function cleanFirst(tbl,expected){if(tbl.getNumRows()>expected){try{tbl.removeRow(0);}catch(e){}}}
    function parseDate(v){
        if(!v)return null;if(v instanceof Date)return isNaN(v.getTime())?null:v;
        var s=String(v);
        if(s.indexOf('/')>0){var p=s.split('/');return new Date(p[2],p[1]-1,p[0]);}
        var d=new Date(s);return isNaN(d.getTime())?null:d;
    }

    // ══════════════════════════════════════════
    //  LEER TODA LA DATA REAL
    // ══════════════════════════════════════════
    var actividades  = getSheetData('Actividades');
    var tramos       = getSheetData('Tramos');
    var patrullajes  = getSheetData('Patrullajes');
    var hallazgos    = getSheetData('Hallazgos');
    var astData      = getSheetData('AST');
    var stopWork     = getSheetData('Stop_Work');
    var caps         = getSheetData('Capacitaciones');
    var apts         = getSheetData('Aptitudes_Medicas');
    var inspecciones = getSheetData('Inspecciones');
    var permisos     = getSheetData('Permisos_Trabajo');
    var sasisopa     = getSheetData('SASISOPA');
    var nomMatriz    = getSheetData('NOM_Matriz');
    var usuarios     = getSheetData('Usuarios');
    var evidencias   = getSheetData('Evidencias');
    var entregaEPP   = getSheetData('Entrega_EPP');
    var platicas     = getSheetData('Platicas_5min');
    var incidentes   = getSheetData('Incidentes');

    function getNombre(id){var u=usuarios.find(function(x){return x.ID===id;});return u?u.Nombre:id||'—';}
    function getTramoNombre(id){var t=tramos.find(function(x){return x.ID===id;});return t?t.Nombre:id||'—';}
    function getTramoZona(id){var t=tramos.find(function(x){return x.ID===id;});return t?t.Zona||'—':'—';}

    // ── Cálculos globales ──
    var totalAct   = actividades.length;
    var actComp    = actividades.filter(function(a){return a.Estado==='Completada';});
    var actEnCurso = actividades.filter(function(a){return a.Estado==='En Curso'||a.Estado==='EnCurso';});
    var actPend    = actividades.filter(function(a){return a.Estado!=='Completada'&&a.Estado!=='En Curso'&&a.Estado!=='EnCurso';});
    var nComp=actComp.length, nCurso=actEnCurso.length, nPend=actPend.length;
    var pctAct = totalAct>0 ? Math.round(nComp/totalAct*100) : 0;

    var totalKm = actividades.reduce(function(s,a){return s+Number(a.KmProgramados||0);},0);
    var kmEjec  = actComp.reduce(function(s,a){return s+Number(a.KmProgramados||0);},0);
    var pctKm   = totalKm>0 ? Math.round(kmEjec/totalKm*100) : 0;

    var nHall  = hallazgos.length;
    var hallAb = hallazgos.filter(function(h){return h.Estado==='Abierto';}).length;
    var hallCe = hallazgos.filter(function(h){return h.Estado==='Cerrado';}).length;
    var pctHall = nHall>0 ? Math.round(hallCe/nHall*100) : 0;

    var swAct  = stopWork.filter(function(s){return s.Estado==='Activo';}).length;
    var nPatr  = patrullajes.length;
    var nAST   = astData.length;
    var nInsp  = inspecciones.length;
    var nPT    = permisos.length;
    var nCap   = caps.length;
    var capVig = caps.filter(function(c){return c.Estatus==='Vigente';}).length;
    var nApt   = apts.length;
    var aptVig = apts.filter(function(a){return a.Estatus==='Vigente';}).length;

    var avgSAS = sasisopa.length>0 ? Math.round(sasisopa.reduce(function(s,e){return s+Number(e.Cumplimiento||0);},0)/sasisopa.length) : 0;
    var avgNOM = nomMatriz.length>0 ? Math.round(nomMatriz.reduce(function(s,n){return s+Number(n.Cumplimiento||0);},0)/nomMatriz.length) : 0;

    // ── Fechas reales del proyecto ──
    var allDates = [];
    actividades.forEach(function(a){var d=parseDate(a.Fecha);if(d)allDates.push(d);});
    patrullajes.forEach(function(p){var d=parseDate(p.Fecha);if(d)allDates.push(d);});
    var fechaInicio = allDates.length>0 ? allDates.reduce(function(a,b){return a<b?a:b;}) : new Date();
    var fechaUltima = allDates.length>0 ? allDates.reduce(function(a,b){return a>b?a:b;}) : new Date();
    var fmtD = function(d){return Utilities.formatDate(d,'America/Mexico_City','dd/MM/yyyy');};
    var diasTranscurridos = Math.max(1, Math.ceil((new Date()-fechaInicio)/86400000));
    var semanasTranscurridas = Math.ceil(diasTranscurridos/7);

    // ═══════════════════════════════════════════
    //  P1 — PORTADA
    // ═══════════════════════════════════════════
    var bar1=body.appendTable();bar1.setBorderWidth(0);
    bar1.appendTableRow().appendTableCell(' ').setBackgroundColor(BLUE).setPaddingTop(3).setPaddingBottom(3);
    cleanFirst(bar1,1);

    body.appendParagraph('').editAsText().setFontSize(4);

    // Logo
    insertLogo(body, 150, 48);
    body.appendParagraph('').editAsText().setFontSize(4);

    sP('REPORTE DE AVANCE CONTRACTUAL',18,DARK,true,DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('').editAsText().setFontSize(2);
    sP('Mantenimiento de Derecho de Vía — Ducto DEN 452 km',10,BLUE,true,DocumentApp.HorizontalAlignment.CENTER);
    sP('Gasoducto de Aguaprieta · Sierra Madre Oriental',8,GRAY_M,false,DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('').editAsText().setFontSize(4);

    var infoT=body.appendTable();infoT.setBorderWidth(0.5).setBorderColor(BORDER);
    var infoRows=[
        ['Cliente','Gasoducto de Aguaprieta'],
        ['Contratista','GENCO'],
        ['Fecha de corte',fechaHoy],
        ['Período reportado', fmtD(fechaInicio)+' al '+fmtD(fechaUltima)+' ('+semanasTranscurridas+' semanas)'],
        ['Tramos contratados', tramos.length+' tramos · '+totalKm+' km programados'],
        ['Generado por','DmV Control v2.0 — Sistema SASISOPA']
    ];
    infoRows.forEach(function(row){
        var r=infoT.appendTableRow();
        var c1=r.appendTableCell(row[0]);c1.setBackgroundColor(GRAY_L).setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(10).setWidth(145);
        c1.getChild(0).asParagraph().editAsText().setFontSize(8).setFontFamily('Arial').setForegroundColor(GRAY_M);
        var c2=r.appendTableCell(row[1]);c2.setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(10);
        c2.getChild(0).asParagraph().editAsText().setFontSize(8).setFontFamily('Arial').setBold(true).setForegroundColor(DARK);
    });
    cleanFirst(infoT,infoRows.length);

    body.appendParagraph('').editAsText().setFontSize(4);

    // ── KPIs ejecutivos ──
    var kT=body.appendTable();kT.setBorderWidth(0.5).setBorderColor(BORDER);
    var kH=kT.appendTableRow();
    ['AVANCE Km','ACTIVIDADES','HALLAZGOS','SASISOPA','STOP WORK'].forEach(function(h){
        var c=kH.appendTableCell(h);c.setBackgroundColor(DARK).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(4).setPaddingRight(4);
        c.getChild(0).asParagraph().editAsText().setFontSize(7).setFontFamily('Arial').setBold(true).setForegroundColor(WHITE);
        c.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    });
    var kV=kT.appendTableRow();
    [[pctKm+'%', pctKm>=80?GREEN:pctKm>=50?ORANGE:RED],
    [nComp+'/'+totalAct, pctAct>=80?GREEN:BLUE],
    [hallAb===0?'0 abiertos':hallAb+' abiertos', hallAb===0?GREEN:hallAb<=3?ORANGE:RED],
    [avgSAS+'%', avgSAS>=85?GREEN:avgSAS>=70?ORANGE:RED],
    [swAct===0?'Ninguno':swAct+' ACTIVO', swAct===0?GREEN:RED]
    ].forEach(function(kpi){
        var c=kV.appendTableCell(kpi[0]);c.setPaddingTop(6).setPaddingBottom(6);
        c.getChild(0).asParagraph().editAsText().setFontSize(15).setFontFamily('Arial').setBold(true).setForegroundColor(kpi[1]);
        c.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    });
    var kS=kT.appendTableRow();
    [kmEjec+'/'+totalKm+' km', nCurso+' en curso, '+nPend+' pend.', hallCe+'/'+nHall+' cerrados ('+pctHall+'%)', 'NOM: '+avgNOM+'%', swAct===0?'Operación normal':'REQUIERE ATENCIÓN'].forEach(function(s){
        var c=kS.appendTableCell(s);c.setBackgroundColor(GRAY_L).setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(4);
        c.getChild(0).asParagraph().editAsText().setFontSize(7).setFontFamily('Arial').setForegroundColor(GRAY_M);
        c.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    });
    cleanFirst(kT,3);

    body.appendPageBreak();

    // ═══════════════════════════════════════════
    //  S1 — AVANCE POR TRAMO
    // ═══════════════════════════════════════════
    secH('AVANCE POR TRAMO', '01');
    body.appendParagraph('').editAsText().setFontSize(2);

    var porTramo={};
    tramos.forEach(function(t){
        porTramo[t.ID]={nombre:t.Nombre||t.ID, zona:t.Zona||'—',
        kmI:Number(t.KmInicio||0), kmF:Number(t.KmFin||0),
        prog:0, comp:0, enCurso:0, kmProg:0, kmEjec:0, patr:0, hall:0, hallAb:0};
    });
    actividades.forEach(function(a){
        var tid=a.Tramo;
        if(!porTramo[tid]) porTramo[tid]={nombre:getTramoNombre(tid),zona:getTramoZona(tid),kmI:0,kmF:0,prog:0,comp:0,enCurso:0,kmProg:0,kmEjec:0,patr:0,hall:0,hallAb:0};
        var pt=porTramo[tid]; pt.prog++;
        pt.kmProg+=Number(a.KmProgramados||0);
        if(a.Estado==='Completada'){pt.comp++;pt.kmEjec+=Number(a.KmProgramados||0);}
        if(a.Estado==='En Curso'||a.Estado==='EnCurso')pt.enCurso++;
    });
    patrullajes.forEach(function(p){if(porTramo[p.Tramo])porTramo[p.Tramo].patr++;});
    hallazgos.forEach(function(h){
        if(porTramo[h.Tramo]){porTramo[h.Tramo].hall++;if(h.Estado==='Abierto')porTramo[h.Tramo].hallAb++;}
    });

    var tT=body.appendTable();tT.setBorderWidth(0.5).setBorderColor(BORDER);
    hdrRow(tT,['TRAMO','ZONA','ACT.','KM PROG','KM EJEC','% AV.','PATR.','HALL.','ESTADO'],[120,60,45,50,50,45,38,40,60]);

    var tKeys=Object.keys(porTramo).sort();
    tKeys.forEach(function(tid,idx){
        var d=porTramo[tid]; var pct=d.kmProg>0?Math.round(d.kmEjec/d.kmProg*100):(d.comp>0?100:0);
        var bg=idx%2===0?WHITE:GRAY_L;
        var estado=pct>=100?'✓ Completo':pct>=50?'● En curso':d.prog>0?'◇ Parcial':'○ Sin act.';
        var estColor=pct>=100?GREEN:pct>=50?BLUE:d.prog>0?ORANGE:GRAY_M;
        var r=tT.appendTableRow();
        dataCell(r,d.nombre,bg,DARK,true,false);
        dataCell(r,d.zona,bg,GRAY_M,false,true);
        dataCell(r,d.comp+'/'+d.prog,bg,DARK,false,true);
        dataCell(r,d.kmProg,bg,GRAY_M,false,true);
        dataCell(r,d.kmEjec,bg,BLUE,true,true);
        dataCell(r,pct+'%',bg,pct>=80?GREEN:pct>=50?ORANGE:pct>0?RED:GRAY_M,true,true);
        dataCell(r,d.patr,bg,DARK,false,true);
        dataCell(r,d.hallAb>0?d.hall+' ('+d.hallAb+'ab)':String(d.hall),bg,d.hallAb>0?ORANGE:DARK,false,true);
        dataCell(r,estado,bg,estColor,true,true);
    });

    // Fila total
    var totR=tT.appendTableRow();
    dataCell(totR,'TOTAL',BLUE,WHITE,true,false);
    dataCell(totR,tKeys.length+' tramos',BLUE,WHITE,false,true);
    dataCell(totR,nComp+'/'+totalAct,BLUE,WHITE,true,true);
    dataCell(totR,totalKm,BLUE,WHITE,false,true);
    dataCell(totR,kmEjec,BLUE,WHITE,true,true);
    dataCell(totR,pctKm+'%',BLUE,WHITE,true,true);
    dataCell(totR,nPatr,BLUE,WHITE,false,true);
    dataCell(totR,nHall,BLUE,WHITE,false,true);
    dataCell(totR,pctKm>=80?'EN META':'EN PROCESO',pctKm>=80?GREEN:ORANGE,WHITE,true,true);
    cleanFirst(tT,tKeys.length+2);

    // ═══════════════════════════════════════════
    //  S2 — GANTT POR TIPO DE ACTIVIDAD (data real)
    // ═══════════════════════════════════════════
    body.appendParagraph('').editAsText().setFontSize(4);
    secH('CRONOGRAMA DE EJECUCIÓN — POR TIPO DE ACTIVIDAD', '02');
    body.appendParagraph('').editAsText().setFontSize(2);
    sP('Fases derivadas del tipo de actividad registrado. Fechas de inicio y fin calculadas de los registros reales del sistema.',8,GRAY_M,false);
    body.appendParagraph('').editAsText().setFontSize(2);

    // Agrupar por Tipo de actividad
    var porTipo={};
    actividades.forEach(function(a){
        var tipo=a.Tipo||'Sin tipo';
        if(!porTipo[tipo])porTipo[tipo]={tipo:tipo,total:0,comp:0,enCurso:0,km:0,kmComp:0,fechas:[],fechaMin:null,fechaMax:null};
        var pt=porTipo[tipo]; pt.total++; pt.km+=Number(a.KmProgramados||0);
        if(a.Estado==='Completada'){pt.comp++;pt.kmComp+=Number(a.KmProgramados||0);}
        if(a.Estado==='En Curso'||a.Estado==='EnCurso')pt.enCurso++;
        var fd=parseDate(a.Fecha);
        if(fd){
        pt.fechas.push(fd);
        if(!pt.fechaMin||fd<pt.fechaMin)pt.fechaMin=fd;
        if(!pt.fechaMax||fd>pt.fechaMax)pt.fechaMax=fd;
        }
    });

    // Calcular rango global para columnas del Gantt
    var tipoKeys=Object.keys(porTipo).sort();
    var globalMin=fechaInicio, globalMax=new Date();
    tipoKeys.forEach(function(tk){
        var pt=porTipo[tk];
        if(pt.fechaMin && pt.fechaMin<globalMin) globalMin=pt.fechaMin;
        if(pt.fechaMax && pt.fechaMax>globalMax) globalMax=pt.fechaMax;
    });

    // Dividir rango en períodos dinámicos
    var totalDias=Math.max(7,Math.ceil((globalMax-globalMin)/86400000));
    var nPeriodos=Math.min(8,Math.max(3,Math.ceil(totalDias/14)));
    var diasPorPeriodo=Math.ceil(totalDias/nPeriodos);
    var periodos=[];
    for(var pi=0;pi<nPeriodos;pi++){
        var pStart=new Date(globalMin.getTime()+pi*diasPorPeriodo*86400000);
        var pEnd  =new Date(globalMin.getTime()+(pi+1)*diasPorPeriodo*86400000-1);
        if(pEnd>globalMax)pEnd=new Date(globalMax.getTime()+7*86400000);
        periodos.push({start:pStart,end:pEnd,
        label:Utilities.formatDate(pStart,'America/Mexico_City','dd/MMM').replace('.','')});
    }

    var gT=body.appendTable();gT.setBorderWidth(0.5).setBorderColor(BORDER);
    var gHdr=['TIPO ACTIVIDAD'];
    periodos.forEach(function(p){gHdr.push(p.label);});
    gHdr.push('TOTAL');gHdr.push('%');
    var gHR=gT.appendTableRow();
    gHdr.forEach(function(h,i){
        var c=gHR.appendTableCell(h);c.setBackgroundColor(DARK).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(4).setPaddingRight(4);
        if(i===0)c.setWidth(135);
        c.getChild(0).asParagraph().editAsText().setFontSize(6.5).setFontFamily('Arial').setBold(true).setForegroundColor(WHITE);
        c.getChild(0).asParagraph().setAlignment(i===0?DocumentApp.HorizontalAlignment.LEFT:DocumentApp.HorizontalAlignment.CENTER);
    });

    tipoKeys.forEach(function(tk,idx){
        var pt=porTipo[tk]; var pct=pt.total>0?Math.round(pt.comp/pt.total*100):0;
        var bg=idx%2===0?WHITE:GRAY_L;
        var gR=gT.appendTableRow();
        // Nombre + info
        var nc=gR.appendTableCell(pt.tipo+'\n'+pt.comp+'/'+pt.total+' comp. · '+pt.kmComp+'km');
        nc.setBackgroundColor(bg).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6);
        var np=nc.getChild(0).asParagraph();
        np.editAsText().setFontSize(7).setFontFamily('Arial').setForegroundColor(GRAY_M);
        if(pt.tipo.length>0){
        np.editAsText().setBold(0,pt.tipo.length-1,true);
        np.editAsText().setFontSize(0,pt.tipo.length-1,8);
        np.editAsText().setForegroundColor(0,pt.tipo.length-1,DARK);
        }

        // Celdas de período — verificar actividades en ese rango
        periodos.forEach(function(per){
        var enPeriodo=pt.fechas.filter(function(f){return f>=per.start&&f<=per.end;}).length;
        var hasFechas=enPeriodo>0;
        var cellTxt=hasFechas?'■ '+enPeriodo:'—';
        var cellBg=hasFechas?(pct>=80?'#DCFCE7':pct>=50?'#DBEAFE':'#FFF7ED'):bg;
        var cellClr=hasFechas?(pct>=80?GREEN:pct>=50?BLUE:ORANGE):'#D1D5DB';
        var pc=gR.appendTableCell(cellTxt);
        pc.setBackgroundColor(cellBg).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(2).setPaddingRight(2);
        pc.getChild(0).asParagraph().editAsText().setFontSize(7.5).setFontFamily('Arial').setForegroundColor(cellClr);
        pc.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        });

        dataCell(gR,pt.comp+'/'+pt.total,bg,DARK,true,true);
        var pctClr=pct>=80?GREEN:pct>=50?ORANGE:pct>0?RED:GRAY_M;
        var pctC=gR.appendTableCell(pct+'%');
        pctC.setBackgroundColor(bg).setPaddingTop(4).setPaddingBottom(4);
        pctC.getChild(0).asParagraph().editAsText().setFontSize(10).setFontFamily('Arial').setBold(true).setForegroundColor(pctClr);
        pctC.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    });
    cleanFirst(gT,tipoKeys.length+1);

    // ═══════════════════════════════════════════
    //  S3 — DETALLE DE ACTIVIDADES
    // ═══════════════════════════════════════════
    body.appendParagraph('').editAsText().setFontSize(4);
    secH('DETALLE DE ACTIVIDADES', '03');
    body.appendParagraph('').editAsText().setFontSize(2);
    sP(totalAct+' actividades · '+nComp+' completadas · '+nCurso+' en curso · '+nPend+' pendientes',9,GRAY_M,false);
    body.appendParagraph('').editAsText().setFontSize(2);

    var aT=body.appendTable();aT.setBorderWidth(0.5).setBorderColor(BORDER);
    hdrRow(aT,['ID','TIPO','TRAMO','LÍDER','FECHA','KM','ESTADO','AST'],[35,80,115,90,58,35,58,35]);

    actividades.forEach(function(a,idx){
        var bg=idx%2===0?WHITE:GRAY_L;
        var estTxt=a.Estado||'—';
        var estClr=estTxt==='Completada'?GREEN:(estTxt==='En Curso'||estTxt==='EnCurso')?BLUE:GRAY_M;
        var r=aT.appendTableRow();
        dataCell(r,a.ID||'—',bg,GRAY_M,false,false);
        dataCell(r,a.Tipo||'—',bg,DARK,true,false);
        dataCell(r,getTramoNombre(a.Tramo),bg,DARK,false,false);
        dataCell(r,getNombre(a.Lider),bg,DARK,false,false);
        dataCell(r,a.Fecha||'—',bg,DARK,false,true);
        dataCell(r,a.KmProgramados||'0',bg,BLUE,true,true);
        dataCell(r,estTxt,bg,estClr,true,true);
        dataCell(r,a.ASTId?'✓':'—',bg,a.ASTId?GREEN:GRAY_M,false,true);
    });
    cleanFirst(aT,actividades.length+1);

    // ═══════════════════════════════════════════
    //  S4 — CUMPLIMIENTO SASISOPA
    // ═══════════════════════════════════════════
    if(sasisopa.length>0){
        body.appendParagraph('').editAsText().setFontSize(4);
        secH('CUMPLIMIENTO SASISOPA', '04');
        body.appendParagraph('').editAsText().setFontSize(2);
        sP('Promedio general: '+avgSAS+'% · '+sasisopa.length+' elementos evaluados',10,avgSAS>=85?GREEN:avgSAS>=70?ORANGE:RED,true);
        body.appendParagraph('').editAsText().setFontSize(2);

        var sT=body.appendTable();sT.setBorderWidth(0.5).setBorderColor(BORDER);
        hdrRow(sT,['#','ELEMENTO SASISOPA','%','ESTADO','NOTAS'],[25,220,45,60,140]);

        sasisopa.forEach(function(s,idx){
        var cum=Number(s.Cumplimiento||0);var bg=idx%2===0?WHITE:GRAY_L;
        var estTxt=cum>=85?'Cumple':cum>=70?'Parcial':'Deficiente';
        var estClr=cum>=85?GREEN:cum>=70?ORANGE:RED;
        var estBg=cum>=85?'#DCFCE7':cum>=70?'#FFF7ED':'#FEE2E2';
        var r=sT.appendTableRow();
        dataCell(r,s.ID||String(idx+1),bg,GRAY_M,false,true);
        dataCell(r,s.Elemento||'',bg,DARK,false,false);
        dataCell(r,cum+'%',bg,estClr,true,true);
        dataCell(r,estTxt,estBg,estClr,true,true);
        dataCell(r,(s.Notas||'').substring(0,60),bg,GRAY_M,false,false);
        });
        cleanFirst(sT,sasisopa.length+1);
    }

    // ═══════════════════════════════════════════
    //  S5 — RESUMEN DE ENTREGABLES
    // ═══════════════════════════════════════════
    body.appendParagraph('').editAsText().setFontSize(4);
    secH('RESUMEN DE ENTREGABLES', sasisopa.length>0?'05':'04');
    body.appendParagraph('').editAsText().setFontSize(2);

    var eT=body.appendTable();eT.setBorderWidth(0.5).setBorderColor(BORDER);
    hdrRow(eT,['ENTREGABLE','CANTIDAD','EVIDENCIA EN SISTEMA','ESTADO'],[150,55,200,65]);

    var entregables=[
        {nombre:'Actividades completadas', cant:nComp+'/'+totalAct, ev:pctAct+'% ejecutado · '+kmEjec+' km recorridos', pct:pctAct},
        {nombre:'Patrullajes realizados', cant:String(nPatr), ev:nPatr+' registros con firma y geolocalización', pct:nPatr>0?100:0},
        {nombre:'AST firmados', cant:String(nAST), ev:'Análisis de seguridad con peligros, riesgos y controles', pct:nAST>0?(totalAct>0?Math.min(100,Math.round(nAST/totalAct*100)):100):0},
        {nombre:'Inspecciones de campo', cant:String(nInsp), ev:'EPP, vehículos, herramienta, maquinaria', pct:nInsp>0?100:0},
        {nombre:'Permisos de trabajo', cant:String(nPT), ev:'Altura, caliente, eléctrico, espacio confinado', pct:nPT>0?100:0},
        {nombre:'Hallazgos gestionados', cant:nHall+' ('+hallCe+' cerr.)', ev:'Tasa de cierre: '+pctHall+'% · '+hallAb+' abiertos', pct:pctHall},
        {nombre:'Capacitaciones DC-3', cant:nCap+' ('+capVig+' vig.)', ev:'Cursos con vigencia y constancia DC-3', pct:nCap>0?Math.round(capVig/nCap*100):0},
        {nombre:'Aptitudes médicas', cant:nApt+' ('+aptVig+' vig.)', ev:'Exámenes médicos y dictámenes', pct:nApt>0?Math.round(aptVig/nApt*100):0},
        {nombre:'Entrega de EPP', cant:String(entregaEPP.length), ev:'Registros con firma de recepción', pct:entregaEPP.length>0?100:0},
        {nombre:'Pláticas 5 minutos', cant:String(platicas.length), ev:'Temas, expositor, asistentes', pct:platicas.length>0?100:0},
        {nombre:'Evidencia fotográfica', cant:String(evidencias.length), ev:'Fotos vinculadas a patrullajes y actividades', pct:evidencias.length>0?100:0}
    ];

    entregables.forEach(function(e,idx){
        var bg=idx%2===0?WHITE:GRAY_L;
        var estTxt=e.pct>=80?'✓ Cumple':e.pct>=50?'● Parcial':e.pct>0?'⚠ Bajo':'○ Pendiente';
        var estClr=e.pct>=80?GREEN:e.pct>=50?BLUE:e.pct>0?ORANGE:GRAY_M;
        var r=eT.appendTableRow();
        dataCell(r,e.nombre,bg,DARK,true,false);
        dataCell(r,e.cant,bg,BLUE,true,true);
        dataCell(r,e.ev,bg,GRAY_M,false,false);
        dataCell(r,estTxt,e.pct>=80?'#DCFCE7':e.pct>=50?'#DBEAFE':e.pct>0?'#FFF7ED':GRAY_L,estClr,true,true);
    });
    cleanFirst(eT,entregables.length+1);

    // ═══════════════════════════════════════════
    //  S6 — OBSERVACIONES
    // ═══════════════════════════════════════════
    body.appendParagraph('').editAsText().setFontSize(4);
    secH('OBSERVACIONES', sasisopa.length>0?'06':'05');
    body.appendParagraph('').editAsText().setFontSize(2);

    var obs=[];
    obs.push('Avance general del contrato: '+pctKm+'% ('+kmEjec+' de '+totalKm+' km). '+nComp+' de '+totalAct+' actividades completadas en '+semanasTranscurridas+' semanas de operación.');

    var tramosConAct=tKeys.filter(function(tid){return porTramo[tid].prog>0;});
    if(tramosConAct.length>0){
        var tramosComp=tramosConAct.filter(function(tid){var d=porTramo[tid];return d.kmProg>0&&d.kmEjec>=d.kmProg;});
        var tramosParc=tramosConAct.filter(function(tid){var d=porTramo[tid];return d.kmProg>0&&d.kmEjec>0&&d.kmEjec<d.kmProg;});
        var tramosSin=tramosConAct.filter(function(tid){var d=porTramo[tid];return d.kmEjec===0;});
        if(tramosComp.length>0)obs.push('✓  Tramos completados: '+tramosComp.map(function(tid){return porTramo[tid].nombre;}).join(', ')+'.');
        if(tramosParc.length>0)obs.push('●  Tramos en ejecución: '+tramosParc.map(function(tid){var d=porTramo[tid];return d.nombre+' ('+Math.round(d.kmEjec/d.kmProg*100)+'%)';}).join(', ')+'.');
        if(tramosSin.length>0)obs.push('○  Tramos sin ejecución: '+tramosSin.map(function(tid){return porTramo[tid].nombre;}).join(', ')+'.');
    }

    if(hallAb>0)obs.push('⚠  '+hallAb+' hallazgo(s) abierto(s) pendiente(s) de cierre. Tasa de cierre actual: '+pctHall+'%.');
    else if(nHall>0)obs.push('✓  Todos los hallazgos han sido cerrados ('+nHall+' totales).');

    if(swAct>0)obs.push('✕  '+swAct+' orden(es) de paro de seguridad activa(s). Se requiere liberación para continuar operación.');

    if(avgSAS>0){
        var elemBajos=sasisopa.filter(function(s){return Number(s.Cumplimiento||0)<70;});
        if(elemBajos.length>0)obs.push('⚠  '+elemBajos.length+' elemento(s) SASISOPA <70%: '+elemBajos.map(function(s){return s.Elemento+' ('+s.Cumplimiento+'%)';}).join(', ')+'.');
        else obs.push('✓  Todos los elementos SASISOPA ≥70%. Promedio: '+avgSAS+'%.');
    }

    if(nCap>0&&capVig<nCap)obs.push('⚠  '+(nCap-capVig)+' capacitación(es) requieren renovación de '+nCap+' totales.');

    obs.forEach(function(o){
        sP(o,8, o.indexOf('✕')>=0?RED : o.indexOf('⚠')>=0?ORANGE : o.indexOf('✓')>=0?'#166534' : '#5A6A7E', o.indexOf('Avance general')>=0);
    });

    // ═══════════════════════════════════════════
    //  FIRMAS
    // ═══════════════════════════════════════════
    body.appendParagraph('').editAsText().setFontSize(6);

    appendFirmasEjecutivas(body,
        { cargo: 'REPRESENTANTE DEL CONTRATISTA', nombre: 'Nombre: _______________', empresa: 'GENCO' },
        { cargo: 'REPRESENTANTE DEL CLIENTE', nombre: 'Nombre: _______________', empresa: 'Gasoducto de Aguaprieta' },
        fechaHoy
    );

    body.appendParagraph('').editAsText().setFontSize(6);
    var fBar=body.appendTable();fBar.setBorderWidth(0);
    fBar.appendTableRow().appendTableCell(' ').setBackgroundColor(BLUE).setPaddingTop(2).setPaddingBottom(2);
    cleanFirst(fBar,1);
    sP('DmV Control v2.0 · SASISOPA Gas · Reporte de Avance Contractual · '+fechaHoy,7,GRAY_M,false,DocumentApp.HorizontalAlignment.CENTER);
    sP('Confidencial — Mantenimiento DdV Ducto DEN 452 km · Gasoducto de Aguaprieta',7,GRAY_M,false,DocumentApp.HorizontalAlignment.CENTER);

    // ═══════════════════════════════════════════
    //  GUARDAR PDF
    // ═══════════════════════════════════════════
    doc.saveAndClose();
    var docFile=DriveApp.getFileById(doc.getId());
    var pdfBlob=docFile.getAs('application/pdf');
    pdfBlob.setName(doc.getName()+'.pdf');
    var pdfFolder;
    try{pdfFolder=DriveApp.getFolderById(ROOT_FOLDER_ID).getFoldersByName('Reportes_PDF').next();}
    catch(e){pdfFolder=DriveApp.getFolderById(ROOT_FOLDER_ID);}
    var pdfFile=pdfFolder.createFile(pdfBlob);
    docFile.setTrashed(true);
    logAction('SISTEMA','Reporte avance contractual','avance',pdfFile.getUrl());
    return {success:true, url:pdfFile.getUrl(), id:pdfFile.getId()};
    }

    // ================================================================
    // PATCH 7: generarReporteUnitario — FUNCIÓN COMPLETA CORREGIDA
    // KARDEX movido ANTES del else final (ya no es código muerto)
    // ================================================================
    function generarReporteUnitario(tipo, id) {
    // Manejar caso donde se pasa un objeto (vía gasCall) o argumentos directos
    if (tipo && typeof tipo === 'object' && !id) {
        id = tipo.id;
        tipo = tipo.tipo;
    }
    var fechaHoy = Utilities.formatDate(new Date(), 'America/Mexico_City', 'dd/MM/yyyy');
    var fechaDoc = Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyyMMdd_HHmm');
    var doc = DocumentApp.create('DmV_Reporte_' + (tipo || 'Unitario') + '_' + (id || '000') + '_' + fechaDoc);
    var body = doc.getBody();
    
    // --- Configuración de Estilo ---
    body.setMarginTop(36); body.setMarginBottom(36); body.setMarginLeft(30); body.setMarginRight(30);
    var DARK='#1A2029', BLUE='#0D4E8B', ACCENT='#E8A838', GREEN='#16A34A', RED='#DC2626', ORANGE='#EA580C', GRAY_L='#F4F6F9', GRAY_M='#8A97A8', WHITE='#FFFFFF', BORDER='#E2E8F0';

    // --- Cargar datos comunes para búsqueda rápida ---
    var usuarios = getSheetData('Usuarios');
    var tramos = getSheetData('Tramos');
    function getUser(uid){var u=usuarios.find(function(x){return x.ID===uid;});return u?u.Nombre:uid||'—';}
    function getTrNm(tid){var t=tramos.find(function(x){return x.ID===tid;});return t?t.Nombre:tid||'—';}

    // --- Helpers de Diseño ---
    function sP(t,s,c,b,a){var p=body.appendParagraph(t?String(t):'');p.editAsText().setFontSize(s).setFontFamily('Arial').setForegroundColor(c||DARK);if(b)p.editAsText().setBold(true);if(a)p.setAlignment(a);return p;}
    function spc(s){body.appendParagraph('').editAsText().setFontSize(s||4);}
    function secH(title,num){
        spc(6);
        var tb=body.appendTable();tb.setBorderWidth(0);var r=tb.appendTableRow();
        var nc=r.appendTableCell(num);nc.setBackgroundColor(BLUE).setWidth(36).setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(8).setPaddingRight(8);
        nc.getChild(0).asParagraph().editAsText().setFontSize(10).setFontFamily('Arial').setBold(true).setForegroundColor(WHITE);
        nc.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        var tc=r.appendTableCell(' '+title.toUpperCase());tc.setBackgroundColor(DARK).setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(12);
        tc.getChild(0).asParagraph().editAsText().setFontSize(10).setFontFamily('Arial').setBold(true).setForegroundColor(WHITE);
        if(tb.getNumRows()>1){try{tb.removeRow(0);}catch(e){}}
        spc(4);
    }
    function kvTable(data){
        var t=body.appendTable();t.setBorderWidth(0.5).setBorderColor(BORDER);
        data.forEach(function(row,i){
        var r=t.appendTableRow();
        var bg=i%2===0?GRAY_L:WHITE;
        var c1=r.appendTableCell(row[0]);c1.setBackgroundColor(bg).setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(8).setWidth(140);
        c1.getChild(0).asParagraph().editAsText().setFontSize(8).setFontFamily('Arial').setForegroundColor(GRAY_M);
        var c2=r.appendTableCell(String(row[1]||'—'));c2.setBackgroundColor(WHITE).setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(8);
        c2.getChild(0).asParagraph().editAsText().setFontSize(8).setFontFamily('Arial').setBold(true).setForegroundColor(DARK);
        });
        if(t.getNumRows()>0&&t.getRow(0).getNumCells()===1)try{t.removeRow(0);}catch(e){}
        return t;
    }
    function dataTable(headers,rows,widths){
        if(!rows||rows.length===0){sP('No hay registros.',9,GRAY_M,false);return;}
        var t=body.appendTable();t.setBorderWidth(0.5).setBorderColor(BORDER);
        var hR=t.appendTableRow();
        headers.forEach(function(h,i){
        var c=hR.appendTableCell(h);c.setBackgroundColor(DARK).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(4).setPaddingRight(4);
        if(widths&&widths[i])c.setWidth(widths[i]);
        c.getChild(0).asParagraph().editAsText().setFontSize(7).setFontFamily('Arial').setBold(true).setForegroundColor(WHITE);
        c.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        });
        rows.forEach(function(row,idx){
        var r=t.appendTableRow();var bg=idx%2===0?WHITE:GRAY_L;
        row.forEach(function(v,i){
            var c=r.appendTableCell(String(v||'—'));c.setBackgroundColor(bg).setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(4).setPaddingRight(4);
            c.getChild(0).asParagraph().editAsText().setFontSize(7).setFontFamily('Arial').setForegroundColor(DARK);
            c.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        });
        });
        if(t.getNumRows()>0&&t.getRow(0).getNumCells()===1)try{t.removeRow(0);}catch(e){}
    }
    function addSig(role1,name1,role2,name2){
        spc(12);
        appendFirmasEjecutivas(body,{cargo:role1,nombre:name1},{cargo:role2,nombre:name2},fechaHoy);
    }

    // --- HEADER GLOBAL ---
    insertLogo(body, 140, 45);
    spc(4);
    sP('DmV CONTROL v2.0 · SISTEMA DE GESTIÓN INTEGRAL',8,GRAY_M,true,DocumentApp.HorizontalAlignment.RIGHT);
    sP('GASODUCTO DE AGUAPRIETA · DUCTO 452 KM',7,BLUE,true,DocumentApp.HorizontalAlignment.RIGHT);
    var line=body.appendTable();line.setBorderWidth(0);line.appendTableRow().appendTableCell(' ').setBackgroundColor(DARK).setPaddingTop(1).setPaddingBottom(1);
    if(line.getNumRows()>1)try{line.removeRow(0);}catch(e){}
    spc(6);

    // --- SWITCH PROCESOS ---
    
    if (tipo === 'PATRULLAJE') {
        var p = getSheetData('Patrullajes').find(function(x){return x.ID===id;});
        if(!p) throw new Error('Patrullaje no encontrado');
        sP('REPORTE DE PATRULLAJE',18,DARK,true,DocumentApp.HorizontalAlignment.CENTER);
        sP('Folio: '+id,10,GRAY_M,false,DocumentApp.HorizontalAlignment.CENTER);
        secH('INFORMACIÓN GENERAL','01');
        kvTable([
        ['Tramo', getTrNm(p.Tramo)],
        ['Líder de Patrullaje', getUser(p.Lider)],
        ['Fecha', p.Fecha],
        ['Kilometraje', 'Del '+p.KmInicio+' al '+p.KmFin],
        ['Estatus', p.Estatus],
        ['Observaciones', p.Observaciones]
        ]);
        if(p.Hallazgo && p.Hallazgo!=='Normal'){
        secH('HALLAZGOS DETECTADOS','02');
        sP('Se reportó una condición: '+p.Hallazgo,10,RED,true);
        sP(p.HallazgoDesc,9,DARK,false);
        }
        var evs = getSheetData('Evidencias').filter(function(e){return e.Patrullaje===id;});
        if(evs.length>0){
        secH('EVIDENCIA FOTOGRÁFICA','03');
        evs.forEach(function(e,i){
            var b64=getImgBase64(e.ArchivoID);
            if(b64){
            var blob=Utilities.newBlob(Utilities.base64Decode(b64.split(',')[1]),'image/jpeg','img.jpg');
            var img=body.appendImage(blob);img.setWidth(300);img.setHeight(200);
            sP('Foto '+(i+1)+': '+e.Tipo,8,GRAY_M,false,DocumentApp.HorizontalAlignment.CENTER);
            spc(4);
            }
        });
        }
        addSig('LÍDER PATRULLAJE',getUser(p.Lider),'SUPERVISOR', '_______________');

    } else if (tipo === 'PERMISO') {
        var pt = getSheetData('Permisos_Trabajo').find(function(x){return x.ID===id;});
        if(!pt) throw new Error('Permiso no encontrado');
        sP('PERMISO DE TRABAJO DE ALTO RIESGO',18,DARK,true,DocumentApp.HorizontalAlignment.CENTER);
        sP(pt.TipoPT.toUpperCase(),12,BLUE,true,DocumentApp.HorizontalAlignment.CENTER);
        sP('Folio: '+id,10,GRAY_M,false,DocumentApp.HorizontalAlignment.CENTER);
        secH('DATOS DEL TRABAJO','01');
        kvTable([
        ['Solicitante', getUser(pt.Solicitante)],
        ['Fecha', pt.Fecha],
        ['Ubicación', pt.Ubicacion],
        ['Horario', pt.HoraInicio + ' - ' + pt.HoraFin],
        ['Descripción', pt.Descripcion],
        ['Estado', pt.Estado]
        ]);
        secH('ANÁLISIS DE RIESGOS','02');
        var r=safeJSON(pt.RiesgosJSON || pt.Riesgos);
        r.forEach(function(x){ sP('• '+(x.item||x),9,DARK); });
        secH('MEDIDAS DE CONTROL','03');
        var c=safeJSON(pt.ControlesJSON || pt.Controles);
        c.forEach(function(x){ sP('☑ '+(x.item||x),9,DARK); });
        addSig('SOLICITANTE',getUser(pt.Solicitante),'AUTORIZA (HSE)', pt.AutorizadoPor);

    } else if (tipo === 'INSPECCION') {
        var ins = getSheetData('Inspecciones').find(function(x){return x.ID===id;});
        if(!ins) throw new Error('Inspección no encontrada');
        sP('LISTA DE VERIFICACIÓN / INSPECCIÓN',18,DARK,true,DocumentApp.HorizontalAlignment.CENTER);
        sP(ins.Tipo.toUpperCase(),12,BLUE,true,DocumentApp.HorizontalAlignment.CENTER);
        secH('DATOS GENERALES','01');
        kvTable([
        ['Inspector', getUser(ins.Inspector)],
        ['Fecha', ins.Fecha],
        ['Objetivo', ins.Objetivo],
        ['Resultado', ins.Resultado]
        ]);
        secH('PUNTOS DE INSPECCIÓN','02');
        var items=safeJSON(ins.ItemsJSON || ins.Items);
        var rows=[];
        items.forEach(function(it){rows.push([it.item || it.nombre || '—', it.estado?'CUMPLE':'NO CUMPLE', it.obs||'']);});
        dataTable(['PUNTO A VERIFICAR','ESTADO','OBSERVACIONES'],rows,[200,80,150]);
        addSig('INSPECTOR',getUser(ins.Inspector),'REVISADO', '_______________');

    } else if (tipo === 'INCIDENTE') {
        var inc = getSheetData('Incidentes').find(function(x){return x.ID===id;});
        if(!inc) throw new Error('Incidente no encontrado');
        sP('REPORTE DE INCIDENTE / ACCIDENTE',18,RED,true,DocumentApp.HorizontalAlignment.CENTER);
        sP(inc.Tipo.toUpperCase(),14,DARK,true,DocumentApp.HorizontalAlignment.CENTER);
        secH('DETALLES DEL EVENTO','01');
        kvTable([
        ['Fecha y Hora', inc.Fecha+' '+inc.Hora],
        ['Lugar', inc.Lugar],
        ['Reportado por', getUser(inc.Responsable)],
        ['Gravedad', inc.Gravedad],
        ['Lesión', inc.LesionTipo]
        ]);
        secH('DESCRIPCIÓN','02');
        sP(inc.Descripcion,10,DARK);
        secH('ANÁLISIS CAUSA RAÍZ','03');
        kvTable([['Causa Inmediata',inc.CausaInmediata],['Causa Raíz',inc.CausaRaiz]]);
        secH('PLAN DE ACCIÓN','04');
        var acc=safeJSON(inc.AccionesJSON || inc.Acciones);
        var aRows=[]; acc.forEach(function(a){aRows.push([a.accion||a, a.responsable||'—', a.fecha||'—']);});
        dataTable(['ACCIÓN','RESPONSABLE','FECHA LIMITE'],aRows,[200,100,80]);
        addSig('REPORTA',getUser(inc.Responsable),'INVESTIGA','_______________');

    } else if (tipo === 'AST') {
        var ast = getSheetData('AST').find(function(x){return x.ID===id;});
        if(!ast) throw new Error('AST no encontrado');
        sP('ANÁLISIS DE SEGURIDAD EN EL TRABAJO',18,DARK,true,DocumentApp.HorizontalAlignment.CENTER);
        sP('(AST / JSA)',12,BLUE,true,DocumentApp.HorizontalAlignment.CENTER);
        sP('Folio: '+id,10,GRAY_M,false,DocumentApp.HorizontalAlignment.CENTER);
        spc(4);
        secH('DATOS GENERALES','01');
        kvTable([
        ['Folio AST', id],
        ['Actividad / Tarea', ast.ActividadTipo || ast.Actividad || '—'],
        ['Actividad vinculada', ast.Actividad || '—'],
        ['Líder de Trabajo', getUser(ast.Lider)],
        ['Fecha', ast.Fecha],
        ['Ubicación (Lat/Lng)', (ast.GeoLat||'—') + ', ' + (ast.GeoLng||'—')]
        ]);
        
        var pels=safeJSON(ast.PeligrosJSON || ast.Peligros);
        var ries=safeJSON(ast.RiesgosJSON || ast.Riesgos);
        var ctrls=safeJSON(ast.ControlesJSON || ast.Controles);
        var maxRows = Math.max(pels.length, ries.length, ctrls.length);
        
        function safeStr(v) {
        if (!v) return '—';
        if (typeof v === 'string') return v;
        if (Array.isArray(v)) return v.map(safeStr).join(', ');
        if (typeof v === 'object') {
            return v.nombre || v.descripcion || v.texto || v.peligro || v.riesgo || v.control || v.label || v.value || 
                Object.values(v).filter(function(x){ return typeof x === 'string'; }).join(' — ') || JSON.stringify(v);
        }
        return String(v);
        }
        
        secH('IDENTIFICACIÓN DE PELIGROS, RIESGOS Y CONTROLES','02');
        if (maxRows > 0) {
        var astRows = [];
        for(var i=0; i<maxRows; i++){
            astRows.push([(i+1).toString(), safeStr(pels[i]), safeStr(ries[i]), safeStr(ctrls[i])]);
        }
        dataTable(['#','PELIGRO IDENTIFICADO','RIESGO ASOCIADO','MEDIDA DE CONTROL'], astRows, [30,150,140,150]);
        } else {
        sP('No se registraron peligros en este AST.',9,GRAY_M);
        }

        secH('FIRMAS DE CONFORMIDAD','03');
        sP('Con mi firma confirmo que he participado en la elaboración del presente AST, he identificado los peligros y riesgos de la actividad, y me comprometo a cumplir las medidas de control establecidas.',8,GRAY_M);
        addSig('LÍDER DE TRABAJO',getUser(ast.Lider),'HSE / SUPERVISOR', '_______________');

    } else if (tipo === 'STOPWORK') {
        var sw = getSheetData('Stop_Work').find(function(x){return x.ID===id;});
        if(!sw) throw new Error('Stop Work no encontrado');
        sP('TARJETA DE PARO DE SEGURIDAD (STOP WORK)',18,RED,true,DocumentApp.HorizontalAlignment.CENTER);
        secH('MOTIVO DEL PARO','01');
        sP(sw.Motivo,12,DARK,true,DocumentApp.HorizontalAlignment.CENTER);
        kvTable([['Reportado por',getUser(sw.Lider)],['Fecha',sw.Fecha],['Estado',sw.Estado]]);
        if(sw.Estado==='Liberado'){
        secH('LIBERACIÓN','02');
        kvTable([['Liberado por',getUser(sw.LiberadoPor)],['Fecha Liberación',sw.FechaLiberacion]]);
        }
        addSig('REPORTA',getUser(sw.Lider),'PARA LIBERAR','_______________');

    } else if (tipo === 'HALLAZGO') {
        var h = getSheetData('Hallazgos').find(function(x){return x.ID===id;});
        if(!h) throw new Error('Hallazgo no encontrado');
        sP('REPORTE DE HALLAZGO',18,ORANGE,true,DocumentApp.HorizontalAlignment.CENTER);
        kvTable([['Tipo',h.Tipo],['Tramo',h.Tramo],['KM',h.Km],['Estado',h.Estado],['Fecha',h.Fecha],['Descripción',h.Descripcion]]);
        addSig('DETECTA',getUser(h.Lider),'ENTERADO','_______________');

    } else if (tipo === 'PLATICA') {
        var pl = getSheetData('Platicas_5min').find(function(x){return x.ID===id;});
        if(!pl) throw new Error('Plática no encontrada');
        sP('PLÁTICA DE 5 MINUTOS',18,GREEN,true,DocumentApp.HorizontalAlignment.CENTER);
        sP('"'+pl.Tema+'"',14,DARK,true,DocumentApp.HorizontalAlignment.CENTER);
        kvTable([['Expositor',pl.Expositor],['Fecha',pl.Fecha],['Duración',pl.Duracion+' min']]);
        secH('ASISTENTES','01');
        var as=safeJSON(pl.AsistentesJSON || pl.Asistentes);
        var rows=[]; 
        as.forEach(function(aId){
        var nombre = getUser(aId);
        rows.push([nombre, 'SI']);
        });
        dataTable(['NOMBRE DEL TRABAJADOR','FIRMA'],rows,[300,100]);
        addSig('EXPOSITOR',pl.Expositor,'VISTO BUENO','_______________');

    } else if (tipo === 'ENTREGA_EPP') {
        var epp = getSheetData('Entrega_EPP').find(function(x){return x.ID===id;});
        if(!epp) throw new Error('Entrega EPP no encontrada');
        sP('CONSTANCIA DE ENTREGA DE EPP',18,BLUE,true,DocumentApp.HorizontalAlignment.CENTER);
        kvTable([['Trabajador',epp.Trabajador],['Fecha',epp.Fecha],['Motivo',epp.Motivo],['Entregado por',epp.EntregadoPor]]);
        secH('EQUIPO ENTREGADO','01');
        var items=safeJSON(epp.EPPEntregadoJSON || epp.EPPEntregado);
        var rows=[]; 
        items.forEach(function(i){
        var desc = (i.item || i) + (i.talla ? ' (Talla: ' + i.talla + ')' : '');
        var cant = i.cantidad || 1;
        rows.push([desc, cant, 'ENTREGADO']);
        });
        dataTable(['DESCRIPCIÓN','CANT.','ACCIÓN'],rows,[250,50,100]);
        sP('\nRecibí el equipo de protección personal descrito, comprometiéndome a utilizarlo correctamente y mantenerlo en buen estado.',8,GRAY_M,false,DocumentApp.HorizontalAlignment.CENTER);
        addSig('ENTREGA',epp.EntregadoPor,'RECIBE CONFORMIDAD',epp.Trabajador);

    } else if (tipo === 'SIMULACRO') {
        var sim = getSheetData('Emergencias').find(function(x){return x.ID===id;});
        if(!sim) throw new Error('Simulacro no encontrado');
        sP('INFORME DE SIMULACRO',18,RED,true,DocumentApp.HorizontalAlignment.CENTER);
        sP(sim.Titulo,12,DARK,true,DocumentApp.HorizontalAlignment.CENTER);
        kvTable([['Fecha',sim.Fecha],['Responsable',sim.Responsable],['Evaluación',sim.Evaluacion]]);
        secH('OBSERVACIONES','01');
        sP(sim.Observaciones,10,DARK);
        addSig('RESPONSABLE',sim.Responsable,'EVALUADOR','_______________');

    } else if (tipo === 'PROGRAMA_SALUD') {
        var sal = getSheetData('Programa_Salud').find(function(x){return x.ID===id;});
        if(!sal) throw new Error('Actividad Salud no encontrada');
        sP('REGISTRO DE SALUD OCUPACIONAL (NOM-030)',16,BLUE,true,DocumentApp.HorizontalAlignment.CENTER);
        kvTable([['Actividad',sal.Actividad],['Tipo',sal.Tipo],['Frecuencia',sal.Frecuencia],['Estado',sal.Estatus],['Fecha Ejecución',sal.FechaEjecucion]]);
        secH('EVIDENCIA Y RESULTADOS','01');
        sP(sal.Observaciones||'Sin observaciones.',10,DARK);
        addSig('RESPONSABLE SALUD',sal.Responsable,'VISTO BUENO','_______________');

    } else if (tipo === 'AUDITORIA') {
        var aud = getSheetData('Auditorias').find(function(x){return x.ID===id;});
        if(!aud) throw new Error('Auditoría no encontrada');
        sP('AUDITORÍA INTERNA SASISOPA',18,DARK,true,DocumentApp.HorizontalAlignment.CENTER);
        sP(aud.Tipo.toUpperCase(),12,BLUE,true,DocumentApp.HorizontalAlignment.CENTER);
        kvTable([['Fecha',aud.Fecha],['Auditor',aud.Auditor],['Resultado Global',aud.Resultado],['NC Mayores',aud.NCMayores],['NC Menores',aud.NCMenores]]);
        secH('RESULTADOS POR ELEMENTO','01');
        var el=safeJSON(aud.ElementosJSON || aud.Elementos);
        var rows=[]; el.forEach(function(e){rows.push([e.numero||'—', e.elemento||'—', e.resultado||'—']);});
        dataTable(['#','ELEMENTO','RESULTADO'],rows,[40,250,100]);
        addSig('AUDITOR LÍDER',aud.Auditor,'AUDITADO','_______________');
        
        if(aud.PlanAccion || aud.PlanAccionJSON) {
        secH('PLAN DE ACCIÓN','02');
        var pAcc = safeJSON(aud.PlanAccionJSON || aud.PlanAccion);
        var pRows = [];
        pAcc.forEach(function(p){ pRows.push([p.accion||p, p.responsable||'—', p.fecha||'—']); });
        dataTable(['ACCIÓN','RESPONSABLE','FECHA'], pRows, [200,100,80]);
        }

    } else if (tipo === 'LISTADO_MAESTRO') {
        var d = getSheetData('Listado_Maestro').find(function(x){return x.ID===id;});
        if(!d) throw new Error('Documento no encontrado');
        sP('FICHA DE CONTROL DOCUMENTAL',18,BLUE,true,DocumentApp.HorizontalAlignment.CENTER);
        sP(d.Codigo,14,DARK,true,DocumentApp.HorizontalAlignment.CENTER);
        kvTable([['Título',d.Titulo],['Elemento SASISOPA',d.Elemento],['Versión',d.Version],['Fecha Emisión',d.FechaEmision],['Responsable',d.Responsable],['Ubicación',d.Ubicacion]]);
        sP('\nDocumento controlado por el sistema SASISOPA.',10,GRAY_M,true,DocumentApp.HorizontalAlignment.CENTER);
        
    } else if (tipo === 'ACTIVIDAD') {
        var act = getSheetData('Actividades').find(function(x){return x.ID===id;});
        if(!act) throw new Error('Actividad no encontrada');
        sP('ORDEN DE TRABAJO / ACTIVIDAD',18,BLUE,true,DocumentApp.HorizontalAlignment.CENTER);
        kvTable([
        ['Folio',act.ID], ['Tipo',act.Tipo],
        ['Tramo', getTrNm(act.Tramo)],
        ['Líder Asignado', getUser(act.Lider)],
        ['Fecha Programada', act.Fecha],
        ['Meta', (act.KmProgramados || act.KmProg || 0) + ' km'],
        ['Estado', act.Estado]
        ]);
        var pats = getSheetData('Patrullajes').filter(function(p){return p.Actividad===id;});
        if(pats.length > 0){
        secH('EJECUCIÓN (PATRULLAJES)','02');
        var rows = [];
        pats.forEach(function(p){
            rows.push([p.ID, p.Fecha, p.KmInicio+' - '+p.KmFin, p.Hallazgo]);
        });
        dataTable(['FOLIO','FECHA','KM','HALLAZGO'],rows,[80,80,100,100]);
        } else {
        sP('\nSin ejecución registrada.',10,GRAY_M,false,DocumentApp.HorizontalAlignment.CENTER);
        }
        addSig('LÍDER ASIGNADO',getUser(act.Lider),'SUPERVISOR', '_______________');

    } else if (tipo === 'EVIDENCIA') {
        var ev = getSheetData('Evidencias').find(function(x){return x.ID===id;});
        if(!ev) throw new Error('Evidencia no encontrada');
        sP('REPORTE DE EVIDENCIA INDIVIDUAL',18,DARK,true,DocumentApp.HorizontalAlignment.CENTER);
        kvTable([['ID',ev.ID],['Tipo',ev.Tipo],['Fecha',ev.Timestamp]]);
        var b64=getImgBase64(ev.ArchivoID);
        if(b64){
        spc(6);
        var blob=Utilities.newBlob(Utilities.base64Decode(b64.split(',')[1]),'image/jpeg','ev.jpg');
        var img=body.appendImage(blob);img.setWidth(450);
        }

    } else if (tipo === 'CAPACITACION') {
        var cap = getSheetData('Capacitaciones').find(function(x){return x.ID===id;});
        if(!cap) throw new Error('Capacitación no encontrada');
        sP('CONSTANCIA DE CAPACITACIÓN',18,BLUE,true,DocumentApp.HorizontalAlignment.CENTER);
        sP(cap.Nombre,14,DARK,true,DocumentApp.HorizontalAlignment.CENTER);
        kvTable([['Trabajador',getUser(cap.Usuario)],['Tipo',cap.Tipo],['Vigencia',cap.Vigencia],['DC-3',cap.DC3?'SÍ':'NO'],['Estatus',cap.Estatus]]);
        sP('\nSe hace constar que el trabajador ha completado satisfactoriamente la capacitación.',10,GRAY_M,false,DocumentApp.HorizontalAlignment.CENTER);
        addSig('INSTRUCTOR/EMPRESA','_______________','TRABAJADOR',getUser(cap.Usuario));

    } else if (tipo === 'APTITUD_MEDICA') {
        var apt = getSheetData('Aptitudes_Medicas').find(function(x){return x.ID===id;});
        if(!apt) throw new Error('Aptitud Médica no encontrada');
        sP('DICTAMEN DE APTITUD MÉDICA',18,BLUE,true,DocumentApp.HorizontalAlignment.CENTER);
        kvTable([['Trabajador',getUser(apt.Usuario)],['Dictamen',apt.Tipo],['Vigencia',apt.Vigencia],['Médico Evaluador',apt.Medico],['Estatus',apt.Estatus]]);
        sP('\nDictamen médico de aptitud para el puesto de trabajo.',10,GRAY_M,false,DocumentApp.HorizontalAlignment.CENTER);
        addSig('MÉDICO EVALUADOR',apt.Medico,'TRABAJADOR',getUser(apt.Usuario));

    } else if (tipo === 'SASISOPA' && id === 'GLOBAL') {
        sP('REPORTE GLOBAL DE CUMPLIMIENTO SASISOPA',18,DARK,true,DocumentApp.HorizontalAlignment.CENTER);
        sP('ESTADO DEL SISTEMA DE GESTIÓN INTEGRAL',12,BLUE,true,DocumentApp.HorizontalAlignment.CENTER);
        var sas = getSheetData('SASISOPA');
        var rows=[]; sas.forEach(function(s){rows.push([s.Elemento, s.Cumplimiento+'%']);});
        spc(6);
        dataTable(['ELEMENTO SASISOPA','CUMPLIMIENTO'],rows,[350,100]);
        addSig('REPRESENTANTE TÉCNICO','_______________','SEGURIDAD Y SALUD','_______________');

    // ════════════════════════════════════════════════════════════════
    // FIX BUG #1 CRÍTICO: KARDEX movido AQUÍ (antes era código muerto
    // porque estaba DESPUÉS del else/return final)
    // ════════════════════════════════════════════════════════════════
    } else if (tipo === 'KARDEX') {
        sP('KARDEX DEL TRABAJADOR', 18, DARK, true, DocumentApp.HorizontalAlignment.CENTER);
        sP('Expediente Integral de Seguridad', 11, GRAY_M, false, DocumentApp.HorizontalAlignment.CENTER);
        spc(6);
        var usr = usuarios.find(function(u) { return u.ID === id; });
        if (!usr) throw new Error('Trabajador no encontrado: ' + id);
        secH('DATOS GENERALES', '01');
        kvTable([['Nombre', usr.Nombre], ['Puesto', usr.Puesto || '—'], ['Rol', usr.Rol], ['Activo Asignado', usr.ActivoAsignado || '—'], ['Número Económico', usr.NumeroEco || '—'], ['Estatus', usr.Activo]]);
        // Capacitaciones
        var capK = getSheetData('Capacitaciones').filter(function(c) { return c.Usuario === id; });
        secH('CAPACITACIONES', '02');
        dataTable(['Curso', 'Tipo', 'Vigencia', 'DC-3', 'Estatus'], capK.map(function(c) { return [c.Nombre, c.Tipo, c.Vigencia, c.DC3 ? 'SÍ' : 'NO', c.Estatus]; }));
        // Aptitudes médicas
        var aptK = getSheetData('Aptitudes_Medicas').filter(function(a) { return a.Usuario === id; });
        secH('APTITUDES MÉDICAS', '03');
        dataTable(['Dictamen', 'Vigencia', 'Médico', 'Estatus'], aptK.map(function(a) { return [a.Tipo, a.Vigencia, a.Medico, a.Estatus]; }));
        // EPP entregado
        var eppK = getSheetData('Entrega_EPP').filter(function(e) { return e.Trabajador === usr.Nombre || e.Trabajador === id; });
        secH('EPP ENTREGADO', '04');
        dataTable(['Fecha', 'Motivo', 'Artículos'], eppK.map(function(e) { var items = []; try { items = JSON.parse(e.EPPEntregadoJSON || '[]'); } catch(ex) {} return [e.Fecha, e.Motivo, items.join(', ')]; }));
        // Permisos
        var ptsK = getSheetData('Permisos_Trabajo').filter(function(p) { return p.Solicitante === id; });
        secH('PERMISOS DE TRABAJO', '05');
        dataTable(['Folio', 'Tipo', 'Fecha', 'Estado'], ptsK.map(function(p) { return [p.ID, p.TipoPT, p.Fecha, p.Estado]; }));
        addSig('SEGURISTA HSE', '_______________', 'TRABAJADOR', usr.Nombre);

    } else {
        // Fallback global de ReporteAvance
        return generarReporteAvance();
    }

    // Footer Standard
    spc(10);
    var fBar=body.appendTable();fBar.setBorderWidth(0);
    fBar.appendTableRow().appendTableCell(' ').setBackgroundColor(BLUE).setPaddingTop(2).setPaddingBottom(2);
    sP('DmV Control v2.0 · SASISOPA · '+id+' · '+fechaHoy,7,GRAY_M,false,DocumentApp.HorizontalAlignment.CENTER);
    
    doc.saveAndClose();
    var docFile=DriveApp.getFileById(doc.getId());
    var pdfBlob=docFile.getAs('application/pdf');
    pdfBlob.setName(tipo+'_'+id+'.pdf');
    var pdfFolder;
    try{pdfFolder=DriveApp.getFolderById(ROOT_FOLDER_ID).getFoldersByName('Reportes_PDF').next();}
    catch(e){pdfFolder=DriveApp.getFolderById(ROOT_FOLDER_ID);}
    var pdfFile = pdfFolder.createFile(pdfBlob);
    docFile.setTrashed(true);
    logAction('SISTEMA', 'PDF Unitario', id, pdfFile.getUrl());
    return { success: true, url: pdfFile.getUrl() };
    }

    /**
    * FUNCIÓN DE MANTENIMIENTO: Repara o crea la estructura de la base de datos.
    * Úsela si faltan hojas o los headers están mal.
    */
    function setupDatabaseStructure() {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheets = {
        'Usuarios': ['ID', 'Nombre', 'PIN', 'Rol', 'Iniciales', 'Activo', 'Puesto', 'ActivoAsignado', 'NumeroEco', 'Timestamp', 'CreadoPor'],
        'Tramos': ['ID', 'Nombre', 'Zona', 'KmInicio', 'KmFin', 'Activo', 'Timestamp', 'Documento', 'CreadoPor'],
        'Catalogo_Actividades': ['ID', 'Nombre', 'Tipo', 'Descripcion', 'Periodicidad', 'Timestamp'],
        'Plan_Trabajo': ['ID', 'Tramo', 'TipoTrabajo', 'Descripcion', 'UnidadMedida', 'CantidadProgramada', 'CantidadEjecutada', 'Avance', 'FechaInicio', 'FechaFin', 'Estado', 'Fase', 'Observaciones', 'Timestamp'],
        'Actividades': ['ID', 'Tramo', 'Lider', 'Tipo', 'Periodicidad', 'Fecha', 'Estado', 'KmProgramados', 'ASTId', 'Timestamp', 'CreadoPor', 'ActivoAsignado', 'KmRealizados', 'Avance', 'Fase'],
        'Patrullajes': ['ID', 'Actividad', 'Lider', 'Fecha', 'Tramo', 'KmInicio', 'KmFin', 'Tipo', 'Estatus', 'Observaciones', 'Hallazgo', 'HallazgoDesc', 'Firma', 'FotosJSON', 'Timestamp'],
        'Hallazgos': ['ID', 'Patrullaje', 'Tipo', 'Descripcion', 'Tramo', 'Km', 'Estado', 'Fecha', 'Lider', 'FechaCierre', 'FotoID', 'Timestamp'],
            'AST': ['ID', 'Actividad', 'Lider', 'Fecha', 'ActividadTipo', 'PeligrosJSON', 'RiesgosJSON', 'ControlesJSON', 'Firma', 'FirmadoPor', 'GeoLat', 'GeoLng', 'Timestamp'],
        'Stop_Work': ['ID', 'Lider', 'Fecha', 'Motivo', 'Estado', 'LiberadoPor', 'FechaLiberacion', 'Timestamp'],
        'SASISOPA': ['ID', 'Elemento', 'Cumplimiento', 'Timestamp', 'Notas'],
        'NOM_Matriz': ['NOM', 'Titulo', 'Modulo', 'EvidenciaRequerida', 'EvidenciaDigital', 'Cumplimiento', 'Estado', 'FechaEval'],
        'EPP': ['Puesto', 'ItemsJSON'],
        'Capacitaciones': ['ID', 'Usuario', 'Tipo', 'Nombre', 'Vigencia', 'DC3', 'Estatus', 'Timestamp'],
        'Aptitudes_Medicas': ['ID', 'Usuario', 'Tipo', 'Vigencia', 'Medico', 'Estatus', 'Timestamp'],
        'Inspecciones': ['ID', 'Fecha', 'Inspector', 'Tipo', 'Objetivo', 'ItemsJSON', 'Resultado', 'Observaciones', 'FotoID', 'Timestamp'],
        'Permisos_Trabajo': ['ID', 'Fecha', 'Solicitante', 'TipoPT', 'Ubicacion', 'Descripcion', 'RiesgosJSON', 'ControlesJSON', 'HoraInicio', 'HoraFin', 'AutorizadoPor', 'Estado', 'FirmaEmisor', 'FirmaEjecutor', 'Timestamp'],
        'Emergencias': ['ID', 'Tipo', 'Titulo', 'Fecha', 'Responsable', 'Participantes', 'Evaluacion', 'Observaciones', 'Estado', 'Timestamp'],
        'Programa_Salud': ['ID', 'Tipo', 'Actividad', 'Frecuencia', 'FechaPrograma', 'FechaEjecucion', 'Responsable', 'Estatus', 'Evidencia', 'Observaciones', 'Timestamp'],
        'Auditorias': ['ID', 'Fecha', 'Tipo', 'Auditor', 'ElementosJSON', 'Resultado', 'NCMayores', 'NCMenores', 'Observaciones', 'Estado', 'PlanAccionJSON', 'Timestamp'],
        'Listado_Maestro': ['ID', 'Codigo', 'Titulo', 'Elemento', 'Version', 'FechaEmision', 'FechaRevision', 'Responsable', 'Ubicacion', 'Estado', 'Timestamp'],
        'Incidentes': ['ID', 'Fecha', 'Hora', 'Lugar', 'Tipo', 'Descripcion', 'Involucrados', 'LesionTipo', 'Gravedad', 'CausaInmediata', 'CausaRaiz', 'AccionesJSON', 'Responsable', 'Estado', 'FotoID', 'Timestamp'],
        'Platicas': ['ID', 'Fecha', 'Tema', 'Expositor', 'AsistentesJSON', 'Duracion', 'Observaciones', 'Timestamp'],
        'Entrega_EPP': ['ID', 'Fecha', 'Trabajador', 'EPPEntregadoJSON', 'Motivo', 'EntregadoPor', 'FirmaRecibe', 'Observaciones', 'Timestamp'],
        'Catalogo_Herramientas': ['ID', 'Nombre', 'Categoria', 'Stock', 'EnCampo', 'Minimo', 'Unidad', 'Ubicacion'],
        'Herramientas_Asignacion': ['ID', 'Fecha', 'Supervisor', 'Lider', 'Cuadrilla', 'ItemsJSON', 'FirmaRecibe', 'Estado', 'Timestamp'],
        'Movimientos_Herramientas': ['ID', 'Fecha', 'Articulo', 'Tipo', 'Cantidad', 'Referencia', 'Usuario', 'Timestamp'],
        'Catalogo_Activos': ['ID', 'Frente', 'Ubicacion', 'Tipo', 'Lider', 'TramoRef', 'Estado'],
        'Personal_Campo': ['ID', 'Nombre', 'Puesto', 'Cuadrilla', 'NSS', 'TipoSangre', 'ContactoEmergencia', 'TelefonoEmergencia', 'FechaIngreso', 'Activo', 'Observaciones'],
        'Plan_Capacitacion': ['ID', 'Curso', 'NOM', 'Dirigido', 'Frecuencia', 'FechaProgramada', 'FechaEjecutada', 'Instructor', 'Duracion', 'Estatus', 'Tipo', 'DC3', 'Mes', 'Prioridad', 'RolesJSON'],
        'Banco_Platicas': ['ID', 'Categoria', 'Tema', 'Contenido', 'NOM', 'Duracion', 'Activo'],
        'Log_Acciones': ['Timestamp', 'Usuario', 'Accion', 'Entidad', 'Detalle'],
        'Evidencias': ['ID', 'Patrullaje', 'Tipo', 'ArchivoURL', 'ArchivoID', 'Timestamp', 'RefId', 'Descripcion', 'FileId', 'URL', 'Usuario'],
        'Reportes_Extraordinarios': ['ID', 'Fecha', 'Hora', 'ActividadRef', 'TipoActividad', 'Tramo', 'Motivo', 'Descripcion', 'ReportadoPor', 'Estado', 'FechaReprogramacion', 'ActividadReprogramadaId', 'EvidenciaFoto', 'ObservacionesCliente', 'Timestamp'],
        'Observaciones': ['ID', 'ActividadID', 'Observacion', 'Tipo', 'Timestamp', 'Usuario'],
        'Inspecciones_Herramientas': ['ID', 'ActividadID', 'Descripcion', 'Estado', 'Timestamp', 'Usuario'],
        'Actos_Inseguros': ['ID', 'PersonalID', 'Tipo', 'Descripcion', 'Severidad', 'Timestamp', 'Usuario']
    };

    Object.keys(sheets).forEach(function(name) {
        var sheet = ss.getSheetByName(name);
        if (!sheet) {
        sheet = ss.insertSheet(name);
        sheet.getRange(1, 1, 1, sheets[name].length).setValues([sheets[name]]);
        Logger.log('Hoja creada: ' + name);
        return;
        }
        
        var data = sheet.getDataRange().getValues();
        var currentHeaders = data[0];
        var expected = sheets[name];
        
        // Auto-fix headers: Si faltan columnas, las agregamos al final
        var missing = expected.filter(function(h) { return currentHeaders.indexOf(h) === -1; });
        if (missing.length > 0) {
        var lastCol = currentHeaders.length;
        sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
        Logger.log('Agregadas columnas en ' + name + ': ' + missing.join(', '));
        }
        
        // Verificar si el orden es vital (los primeros N deben coincidir para appendRow seguro)
        // Para simplificar, si la hoja está vacía (solo headers), forzamos el orden esperado
        if (sheet.getLastRow() <= 1) {
        sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
        }
    });
    return { success: true, mensaje: 'Base de Datos Sincronizada y Reparada.' };
    }

    // ================================================================
    // FUNCIONES NUEVAS PARA FRONTEND v3.0
    // ================================================================

    // Guardar observaciones del líder al cerrar jornada
    function guardarObservacion(payload) {
    // payload: { actividadId, observacion, tipo, userId }
    var auth = checkRole(payload.userId, 'LIDER');
    if (!auth.allowed) return { success: false, error: auth.error };
    
    try {
        appendRow('Observaciones', [
        payload.actividadId,
        payload.observacion,
        payload.tipo || 'General',
        new Date().toISOString(),
        payload.userId
        ]);
        logAction(payload.userId, 'GUARDAR OBSERVACION', payload.actividadId, payload.observacion);
        return { success: true };
    } catch(e) {
        return { success: false, error: e.message };
    }
    }

    // Inspección de Herramientas
    function guardarInspeccionHerramienta(payload) {
    // payload: { actividadId, descripcion, estado, userId }
    var auth = checkRole(payload.userId, 'SEGURISTA');
    if (!auth.allowed) return { success: false, error: auth.error };
    
    try {
        var id = genId('INSH-', 'Inspecciones_Herramientas');
        appendRow('Inspecciones_Herramientas', [
        id,
        payload.actividadId || '',
        payload.descripcion,
        payload.estado,
        new Date().toISOString(),
        payload.userId
        ]);
        logAction(payload.userId, 'INSPECCION HERRAMIENTAS', id, payload.estado);
        return { success: true, id: id };
    } catch(e) {
        return { success: false, error: e.message };
    }
    }

    /**
    * Registra un nuevo usuario en la base de datos (Hoja 'Usuarios')
    */
    function registrarUsuarioApp(payload) {
    try {
        // Validar permisos del que registra (puedes ajustar el rol requerido)
        // var auth = checkRole(payload.adminId, 'ADMIN');
        // if (!auth.allowed) return auth;

        var newId = genId('USR-', 'Usuarios');
        var row = [
        newId,
        payload.nombre,
        payload.pin,
        payload.rol,
        'TRUE', // Activo
        payload.activoAsignado || 'TODOS',
        payload.nombre.substring(0, 2).toUpperCase(), // Iniciales
        new Date().toISOString()
        ];

        appendRow('Usuarios', row);
        logAction(payload.adminId || 'SISTEMA', 'REGISTRAR USUARIO', newId, 'Nombre: ' + payload.nombre);

        return { success: true, id: newId };
    } catch (e) {
        return { success: false, error: 'Error al registrar: ' + e.message };
    }
    }

    // Registrar Amonestación (Actos Inseguros)
    function registrarAmonestacion(payload) {
    // payload: { personalId, tipo, desc, severidad, userId }
    var auth = checkRole(payload.userId, 'SEGURISTA');
    if (!auth.allowed) return { success: false, error: auth.error };
    
    try {
        var id = genId('AMON-', 'Actos_Inseguros');
        appendRow('Actos_Inseguros', [
        id,
        payload.personalId,
        payload.tipo,
        payload.desc,
        payload.severidad,
        new Date().toISOString(),
        payload.userId
        ]);
        logAction(payload.userId, 'REGISTRAR AMONESTACION', id, payload.tipo + ' - ' + payload.severidad);
        return { success: true, id: id };
    } catch(e) {
        return { success: false, error: e.message };
    }
    }

    // ================================================================
    // MÓDULOS SEGURISTA - EXTENSIONES (AST, Inspecciones)
    // ================================================================

    function saveASTMultifactorial(payload) {
    var auth = checkRole(payload.userId, 'LIDER'); // Minimum role
    if (!auth.allowed) return { success: false, error: auth.error };

    try {
        var id = genId('AST-', 'AST');
        var sheet = getSheet('AST');
        if (!sheet) {
        // Auto-create sheet if missing
        var ss = SpreadsheetApp.openById(GLOBAL_SS_ID_OVERRIDE || SPREADSHEET_ID);
        sheet = ss.insertSheet('AST');
        sheet.appendRow(['ID', 'Fecha', 'ActividadID', 'UsuarioID', 'Peligros', 'Riesgos', 'Controles', 'Firmas_Cuadrilla', 'JSON_Full']);
        }

        var row = [
        id,
        Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd HH:mm:ss'),
        payload.actividadId,
        payload.userId,
        (payload.peligros || []).join(' | '),
        (payload.riesgos || []).join(' | '),
        (payload.controles || []).join(' | '),
        (payload.firmas || []).join(','),
        JSON.stringify(payload)
        ];

        sheet.appendRow(row);
        logAction(payload.userId, 'CREAR_AST', id, 'Actividad: ' + payload.actividadId);
        
        return { success: true, id: id };
    } catch (e) {
        return { success: false, error: e.message };
    }
    }

    function guardarInspeccionHerramienta(payload) {
    var auth = checkRole(payload.userId, 'LIDER');
    if (!auth.allowed) return { success: false, error: auth.error };

    try {
        var id = genId('INS-', 'Inspecciones_Herramienta');
        var sheet = getSheet('Inspecciones_Herramienta');
        if (!sheet) {
        var ss = SpreadsheetApp.openById(GLOBAL_SS_ID_OVERRIDE || SPREADSHEET_ID);
        sheet = ss.insertSheet('Inspecciones_Herramienta');
        sheet.appendRow(['ID', 'Fecha', 'ActividadID', 'UsuarioID', 'Descripcion', 'Estado']);
        }

        var row = [
        id,
        Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd HH:mm:ss'),
        payload.actividadId || 'GENERAL',
        payload.userId,
        payload.descripcion,
        payload.estado
        ];

        sheet.appendRow(row);
        logAction(payload.userId, 'INSPECCION_HERRAMIENTA', id, payload.estado);
        return { success: true, id: id };
    } catch (e) {
        return { success: false, error: e.message };
    }
    }

    function registrarSimulacro(payload) {
    var auth = checkRole(payload.userId, 'SEGURISTA');
    if (!auth.allowed) return { success: false, error: auth.error };

    try {
        var id = genId('SIM-', 'Simulacros');
        var sheet = getSheet('Simulacros');
        if (!sheet) {
        var ss = SpreadsheetApp.openById(GLOBAL_SS_ID_OVERRIDE || SPREADSHEET_ID);
        sheet = ss.insertSheet('Simulacros');
        sheet.appendRow(['ID', 'Fecha', 'Tipo', 'Ubicacion', 'Observaciones', 'Participantes', 'RegistradoPor']);
        }

        var row = [
        id,
        Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd HH:mm:ss'),
        payload.tipo,
        payload.ubicacion,
        payload.observaciones,
        (payload.participantes || []).join(','),
        payload.userId
        ];

        sheet.appendRow(row);
        logAction(payload.userId, 'REGISTRAR SIMULACRO', id, payload.tipo);
        return { success: true, id: id };
    } catch (e) {
        return { success: false, error: e.message };
    }
    }

    function registrarSalud(payload) {
    var auth = checkRole(payload.userId, 'SEGURISTA');
    if (!auth.allowed) return { success: false, error: auth.error };

    try {
        var id = genId('SAL-', 'Salud_Ocupacional');
        var sheet = getSheet('Salud_Ocupacional');
        if (!sheet) {
        var ss = SpreadsheetApp.openById(GLOBAL_SS_ID_OVERRIDE || SPREADSHEET_ID);
        sheet = ss.insertSheet('Salud_Ocupacional');
        sheet.appendRow(['ID', 'Fecha', 'Tipo', 'PersonalID', 'Diagnostico', 'ProxRevision', 'MedicoID']);
        }

        var row = [
        id,
        Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd HH:mm:ss'),
        payload.tipo,
        payload.personalId,
        payload.diag,
        payload.proxRevision,
        payload.userId
        ];

        sheet.appendRow(row);
        logAction(payload.userId, 'REGISTRO SALUD', id, payload.tipo + ' - ' + payload.personalId);
        return { success: true, id: id };

    } catch (e) {
        return { success: false, error: e.message };
    }
    }

    function getAppLogo() {
    // Retorna el logo configurado en variables globales
    var img = getImgBase64(LOGO_FILE_ID);
    return { success: true, image: img };
    }

    function registrarEntregaEPP(payload) {
    var auth = checkRole(payload.userId, 'SEGURISTA');
    if (!auth.allowed) return { success: false, error: auth.error };

    try {
        var id = genId('EPP-E-', 'Entrega_EPP');
        var sheet = getSheet('Entrega_EPP');
        
        // ID(0), Fecha(1), PersonalID(2), Items(3), TipoEntrega(4), Motivo(5), SupervisorID(6)
        var row = [
        id,
        Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd HH:mm:ss'),
        payload.personalId,
        JSON.stringify(payload.items), // Array of {item, cantidad}
        payload.tipoEntrega, // 'Nuevo', 'Reposición'
        payload.motivo || '',
        payload.userId
        ];
        
        sheet.appendRow(row);
        logAction(payload.userId, 'ENTREGA EPP', id, payload.personalId);
        return { success: true, id: id };
    } catch (e) {
        return { success: false, error: e.message };
    }
    }

    function registrarInspeccionEPP(payload) {
    var auth = checkRole(payload.userId, 'SEGURISTA');
    if (!auth.allowed) return { success: false, error: auth.error };

    try {
        var id = genId('EPP-I-', 'Inspeccion_EPP');
        var sheet = getSheet('Inspeccion_EPP');
        
        // ID(0), Fecha(1), PersonalID(2), Estado(3), Detalle(4), InspectorID(5)
        // Estado: 'Aprobado', 'No Aprobado'
        // Detalle: JSON with checklist
        var row = [
        id,
        Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd HH:mm:ss'),
        payload.personalId,
        payload.estado,
        JSON.stringify(payload.detalle),
        payload.userId
        ];
        
        sheet.appendRow(row);
        logAction(payload.userId, 'INSPECCION EPP', id, payload.personalId + ' - ' + payload.estado);
        return { success: true, id: id };
    } catch (e) {
        return { success: false, error: e.message };
    }

    }

    function registrarResguardo(payload) {
    var auth = checkRole(payload.userId, 'ADMIN'); // Only Admin/Supervisor can assign assets
    if (!auth.allowed) return { success: false, error: auth.error };

    try {
        var id = genId('RES-', 'Resguardos');
        var sheet = getSheet('Resguardos');
        
        // ID(0), Fecha(1), RecursoID(2), Tipo(3), AsignadoA(4), Estado(5), AdminID(6)
        var row = [
        id,
        Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd HH:mm:ss'),
        payload.recursoId, // Placa o Serie
        payload.tipo, // 'Vehiculo', 'Maquinaria', 'Herramienta'
        payload.personalId,
        'Activo',
        payload.userId
        ];
        
        sheet.appendRow(row);
        logAction(payload.userId, 'RESGUARDO ASIGNADO', id, payload.recursoId + ' -> ' + payload.personalId);
        return { success: true, id: id };
    } catch (e) {
        return { success: false, error: e.message };
    }
    }

    function registrarInspeccionRecurso(payload) {
    var auth = checkRole(payload.userId, 'LIDER'); // Driver/Operator (Lider) checks it
    if (!auth.allowed) return { success: false, error: auth.error };

    try {
        var id = genId('INS-R-', 'Inspeccion_Recursos');
        var sheet = getSheet('Inspeccion_Recursos');
        
        // ID(0), Fecha(1), RecursoID(2), Tipo(3), Kilometraje/Horas(4), Estado(5), Checklist(6), OperadorID(7)
        var row = [
        id,
        Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd HH:mm:ss'),
        payload.recursoId,
        payload.tipo,
        payload.medicion, // Km or Horas
        payload.estado, // 'Operativo', 'Requiere Mantenimiento', 'No Operativo'
        JSON.stringify(payload.checklist),
        payload.userId,
        payload.fotoId || ''
        ];
        
        sheet.appendRow(row);
        logAction(payload.userId, 'INSPECCION RECURSO', id, payload.recursoId + ' (' + payload.estado + ')');
        return { success: true, id: id };
    } catch (e) {
        return { success: false, error: e.message };
    }
    }

    function registrarPermisoTrabajo(payload) {
    var auth = checkRole(payload.userId, 'SEGURISTA'); // Segurista authorizes
    if (!auth.allowed) return { success: false, error: auth.error };

    try {
        var id = genId('PT-', 'Permisos_Trabajo');
        var sheet = getSheet('Permisos_Trabajo');
        
        // ID(0), Fecha(1), TipoTrabajo(2), Ubicacion(3), Responsable(4), Vigencia(5), Estado(6), Autorizo(7)
        // Tipo: Alturas, Caliente, Espacios Confinados, Electrico
        var row = [
        id,
        Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd HH:mm:ss'),
        payload.tipo,
        payload.ubicacion,
        payload.responsableIt, // ID del Lider/Responsable
        payload.vigencia, // '08:00 - 18:00'
        'Autorizado',
        payload.userId,
        payload.fotoId || ''
        ];
        
        sheet.appendRow(row);
        logAction(payload.userId, 'PERMISO TRABAJO', id, payload.tipo + ' @ ' + payload.ubicacion);
        return { success: true, id: id };
    } catch (e) {
        return { success: false, error: e.message };
    }
    }

    function registrarLogAmbiental(payload) {
    var auth = checkRole(payload.userId, 'SEGURISTA');
    if (!auth.allowed) return { success: false, error: auth.error };

    try {
        var id = genId('AMB-', 'Gestion_Ambiental');
        var sheet = getSheet('Gestion_Ambiental');
        
        // ID(0), Fecha(1), TipoResiduo(2), Cantidad(3), Unidad(4), Disposicion(5), Manifiesto(6), Registrador(7), Foto(8)
        var row = [
        id,
        Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd HH:mm:ss'),
        payload.tipoResiduo, // 'RP', 'RME', 'RSU'
        payload.cantidad,
        payload.unidad, // 'kg', 'lt', 'm3'
        payload.disposicion, // 'Almacen Temporal', 'Disposición Final'
        payload.manifiesto || 'N/A',
        payload.userId,
        payload.fotoId || ''
        ];
        
        sheet.appendRow(row);
        logAction(payload.userId, 'REGISTRO AMBIENTAL', id, payload.tipoResiduo + ': ' + payload.cantidad + payload.unidad);
        return { success: true, id: id };
    } catch (e) {
        return { success: false, error: e.message };
    }
    }

    function getImgBase64(fileId) {
    try {
        var file = DriveApp.getFileById(fileId);
        var blob = file.getBlob();
        var b64 = Utilities.base64Encode(blob.getBytes());
        return 'data:' + blob.getContentType() + ';base64,' + b64;
    } catch (e) {
        if (fileId && fileId.indexOf('http') === 0) return null; // Ignore URLs
        // console.warn('Error fetching image ' + fileId, e);
        return null;
    }
    }

    function generarReportePermiso(id) {
    try {
        var permisos = getSheetData('Permisos_Trabajo');
        var p = permisos.find(function(x) { return x.ID === id; });
        if (!p) throw new Error('Permiso no encontrado');
        
        var responsableName = getUser(p.ResponsableIt || p.Solicitante); // Fallback
        
        var fotoBase64 = null;
        if (p.Foto) {
            var fId = p.Foto;
            // Basic check if it's a file ID
            if (fId.length > 5 && fId.indexOf('http') === -1) {
                var fullB64 = getImgBase64(fId);
                if (fullB64) {
                    fotoBase64 = fullB64.split(',')[1];
                }
            }
        }
        
        var template = HtmlService.createTemplateFromFile('Template_PermisoTrabajo');
        template.data = {
            id: p.ID,
            fecha: p.Fecha,
            tipo: p.Tipo,
            ubicacion: p.Ubicacion,
            responsableIt: p.ResponsableIt,
            responsableName: responsableName,
            vigencia: p.Vigencia,
            autorizo: getUser(p.Registrador),
            fotoBase64: fotoBase64
        };
        
        var logo = getAppLogo();
        template.logo = logo.success ? logo.image : ''; 
        
        var html = template.evaluate().getContent();
        var blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF);
        blob.setName('Permiso_' + p.ID + '.pdf');
        
        return { success: true, base64: Utilities.base64Encode(blob.getBytes()) };
        
    } catch (e) {
        return { success: false, error: e.message };
    }
    }

    function getPermisosUsuario(userId) {
    try {
        var permisos = getSheetData('Permisos_Trabajo');
        // Filter by user (Creator or Responsible) if needed, or just last 10
        // Returning last 10 for simplicity
        return { 
            success: true, 
            data: permisos.slice(-10).reverse().map(function(p){ 
                return { id: p.ID, fecha: p.Fecha, tipo: p.Tipo, estado: p.Estado }; 
            }) 
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
    }

    function getInspeccionesUsuario(userId) {
    try {
        var ins = getSheetData('Inspeccion_Recursos');
        // Filter by user or recent
        // Returning last 10 for simplicity
        return { 
            success: true, 
            data: ins.slice(-10).reverse().map(function(p){ 
                // ID(0), Fecha(1), RecursoID(2), Tipo(3), ...
                return { id: p.ID, fecha: p.Fecha, recurso: p.RecursoID, estado: p.Estado }; 
            }) 
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
    }

    function generarReporteInspeccion(id) {
    try {
        var ins = getSheetData('Inspeccion_Recursos');
        var p = ins.find(function(x) { return x.ID === id; });
        if (!p) throw new Error('Inspección no encontrada');
        
        // Parse Checklist JSON
        var checklist = {};
        try { checklist = JSON.parse(p.Checklist); } catch(e) {}
        
        var fotoBase64 = null;
        if (p.Foto) {
            var fId = p.Foto;
            if (fId.length > 5 && fId.indexOf('http') === -1) {
                var fullB64 = getImgBase64(fId);
                if (fullB64) fotoBase64 = fullB64.split(',')[1];
            }
        }
        
        var template = HtmlService.createTemplateFromFile('Template_InspeccionRecurso');
        template.data = {
            id: p.ID,
            fecha: p.Fecha,
            recursoId: p.RecursoID,
            medicion: p['Kilometraje/Horas'] || p.Medicion, 
            estado: p.Estado,
            inspector: getUser(p.OperadorID),
            checklist: checklist,
            fotoBase64: fotoBase64
        };
        
        var logo = getAppLogo();
        template.logo = logo.success ? logo.image : ''; 
        
        var html = template.evaluate().getContent();
        var blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF);
        blob.setName('Inspeccion_' + p.ID + '.pdf');
        
        return { success: true, base64: Utilities.base64Encode(blob.getBytes()) };
        
    } catch (e) {
        return { success: false, error: e.message };
    }
    }

    function registrarChecklistNOM(payload) {
    try {
        if (!payload || !payload.userId) throw new Error('Payload inválido');

        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('Checklist_NOM');
        if (!sheet) {
        sheet = ss.insertSheet('Checklist_NOM');
        sheet.appendRow(['ID', 'Fecha', 'Norma', 'ResultadosJSON', 'UsuarioID']);
        }
        
        var id = 'NOM-' + new Date().getTime();
        var resultadosStr = JSON.stringify(payload.resultados || []);
        
        // ID(0), Fecha(1), Norma(2), Resultados(3), Usuario(4)
        var row = [
        id,
        Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd HH:mm:ss'),
        payload.nom || 'GENERICO',
        resultadosStr,
        payload.userId
        ];
        
        sheet.appendRow(row);
        logAction(payload.userId, 'CHECKLIST NOM', id, payload.nom);
        return { success: true, id: id };
    } catch (e) {
        return { success: false, error: e.message };
    }
    }

    function getUser(id) {
    var users = getSheetData('Usuarios');
    var u = users.find(function(x) { return x.ID == id; });
    return u ? u.Nombre : 'Usuario Desconocido';
    }

    // ================================================================
    // SEED DATABASE — Ejecutar UNA VEZ desde el Editor de Apps Script
    // Menú: Ejecutar > seedDatabase
    // Headers verificados contra TODOS los appendRow() en Code.gs
    // ================================================================
    function seedDatabase() {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Función auxiliar para BORRAR hojas (Úsese con precaución si se desea reiniciar todo)
    // Para activar, descomentar la llamada en el cuerpo principal
    function deleteSheetIfExists(name) {
        var sheet = ss.getSheetByName(name);
        if (sheet) {
        ss.deleteSheet(sheet);
        Logger.log('Wm️ Hoja eliminada: ' + name);
        }
    }

    function ensureSheet(name, headers) {
        var sheet = ss.getSheetByName(name);
        if (!sheet) {
        sheet = ss.insertSheet(name);
        sheet.appendRow(headers);
        Logger.log('✅ Creada hoja: ' + name);
        } else {
        // Verificar headers (opcional, por ahora solo aseguramos existencia)
        if (sheet.getLastRow() === 0) sheet.appendRow(headers);
        Logger.log('ℹ️ Hoja ya existe: ' + name + ' (' + (sheet.getLastRow() - 1) + ' registros)');
        }
        return sheet;
    }

    function seedIfEmpty(name, headers, rows) {
        var sheet = ensureSheet(name, headers);
        if (sheet.getLastRow() <= 1) {
        rows.forEach(function(row) { sheet.appendRow(row); });
        Logger.log('🌱 Sembrados ' + rows.length + ' registros en ' + name);
        } else {
        Logger.log('⏭ ' + name + ' ya tiene datos, omitiendo seed');
        }
        return sheet;
    }

    var hoy = Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd');
    var ts = new Date().toISOString();

    // ============================================================
    // 0. LIMPIEZA (Opcional - Descomentar para borrar hojas viejas/erróneas)
    // ============================================================
    deleteSheetIfExists('Activos'); // Borrar para actualizar
    deleteSheetIfExists('Tramos'); // Borrar para actualizar
    deleteSheetIfExists('Usuarios'); // Borrar para actualizar
    // deleteSheetIfExists('Catalogo_Conceptos'); // Hoja duplicada eliminada
    // deleteSheetIfExists('Items'); // Hoja vieja si existe

    // ============================================================
    // 1. ACTIVOS (Gasoductos)
    // ============================================================
    seedIfEmpty('Activos',
        ['ID', 'Nombre', 'Descripcion', 'LongitudKm', 'Ubicacion'],
        [
        ['ACT-01', 'GDN (Gasoducto del Noreste)',              '48 pulg, 117 km', 116.78, 'Camargo -> Los Ramones'],
        ['ACT-02', 'DEN (Ductos y Energéticos del Norte)',     '42 pulg, 452 km', 452.46, 'Los Ramones -> San Luis Potosí'],
        ['ACT-03', 'GDT (Gasoducto del Tamaulipas)',           '36 pulg, 114 km', 114.24, 'El Caracol -> Los Indios']
        ]
    );

    // ============================================================
    // 2. TRAMOS (Definidos en Anexo 1 Alcance)
    // ============================================================
    seedIfEmpty('Tramos',
        ['ID','Nombre','ActivoID','KmInicio','KmFin','Valvulas','Timestamp','CreadoPor'],
        [
        ['TRAMO_GDN_01', 'Tramo Único GDN',   'ACT-01', 0,      116.78, 'MLV-01011 a MLV-01013', ts, 'SEED'],
        ['TRAMO_DEN_01', 'Tramo Norte DEN',   'ACT-02', 0,      184.50, 'MLV-2011 a MLV-2111',   ts, 'SEED'],
        ['TRAMO_DEN_02', 'Tramo Sur DEN',     'ACT-02', 184.50, 452.46, 'MLV-2212 a MLV-2215',   ts, 'SEED'],
        ['TRAMO_GDT_01', 'Tramo Único GDT',   'ACT-03', 0,      114.24, 'MLV-0101 a MLV-0103',   ts, 'SEED']
        ]
    );

    // ============================================================
    // 3. CATALOGO DE ACTIVIDADES (Conceptos del Anexo 1 - Excel)
    // ============================================================
    // Se unifica aquí lo que antes estaba separado.
    // Tipo se infiere: Mantenimiento, Operación, Maquinaria, etc.
    var actividadesSeed = [
        // DESMONTE Y LIMPIEZA
        ['CAT001', 'Desmonte de maleza manual/eléctrica', 'Mantenimiento', 'Desmonte de maleza con herramienta manual y/o eléctrica', 'Mensual', ts],
        ['CAT002', 'Tractor con desmalezadora',           'Mantenimiento', 'Limpieza de vegetación a lo largo y ancho',             'Trimestral', ts],
        ['CAT003', 'Desbroce con tractor (GDT 36")',      'Mantenimiento', 'Desbroce a lo largo de la franja',                      'Trimestral', ts],
        
        // MAQUINARIA Y EQUIPO (Se listan como actividades de soporte o renta)
        ['CAT004', 'Uso de Compactadora tipo bailarina',  'Maquinaria',    'Arrendamiento jornada',                                 'Evento', ts],
        ['CAT005', 'Uso de Retroexcavadora 4x4',          'Maquinaria',    'Arrendamiento jornada con operador',                    'Evento', ts],
        ['CAT006', 'Traslado de retroexcavadora',         'Logística',     'Servicio de traslado llegada/salida',                   'Evento', ts],
        ['CAT007', 'Uso de Motoconformadora',             'Maquinaria',    'Arrendamiento jornada con operador',                    'Evento', ts],
        ['CAT008', 'Uso de Camioneta 4x4',                'Maquinaria',    'Arrendamiento jornada',                                 'Evento', ts],
        ['CAT009', 'Uso de Lote de Herramienta',          'Herramienta',   'Jornada de herramienta necesaria',                      'Evento', ts],

        // MATERIALES (HUNDIMIENTOS)
        ['CAT010', 'Suministro Piedra Bola 7m3',          'Suministro',    'Viaje de material de relleno',                          'Evento', ts],
        ['CAT011', 'Suministro Caliche 7m3',              'Suministro',    'Viaje de material de relleno',                          'Evento', ts],
        ['CAT012', 'Suministro Arena 7m3',                'Suministro',    'Viaje de material suave',                               'Evento', ts],
        ['CAT013', 'Suministro Grava 3/4 7m3',            'Suministro',    'Viaje de material',                                     'Evento', ts],
        ['CAT014', 'Suministro Tierra Relleno 7m3',       'Suministro',    'Viaje de tierra',                                       'Evento', ts],
        ['CAT015', 'Suministro Agua 5000L',               'Suministro',    'Viaje de pipa de agua',                                 'Evento', ts],

        // PERSONAL (Servicios Unitarios)
        ['CAT016', 'Servicio Supervisor de Obra',         'Personal',      'Jornada de supervisión',                                'Diario', ts],
        ['CAT017', 'Servicio Líder de Cuadrilla',         'Personal',      'Jornada de líder',                                      'Diario', ts],
        ['CAT018', 'Servicio Ayudante General',           'Personal',      'Jornada de ayudante',                                   'Diario', ts],
        ['CAT019', 'Servicio Supervisor HSE',             'Personal',      'Jornada de seguridad',                                  'Diario', ts],

        // SEÑALIZACIÓN
        ['CAT020', 'Mantenimiento Postes (IV, RA, R)',    'Mantenimiento', 'Pintura, reparación, nivelación',                       'Evento', ts],
        ['CAT021', 'Suministro Poste R',                  'Suministro',    'Pieza',                                                 'Evento', ts],
        ['CAT022', 'Suministro Poste RA',                 'Suministro',    'Pieza',                                                 'Evento', ts],
        ['CAT023', 'Suministro Poste IV',                 'Suministro',    'Pieza',                                                 'Evento', ts],
        ['CAT024', 'Suministro Placa Señalización',       'Suministro',    'Pieza',                                                 'Evento', ts],

        // OBRAS
        ['CAT025', 'Andamios Multidireccionales',         'Obras',         'Arrendamiento, armado y desarmado',                     'Evento', ts],
        ['CAT026', 'Reparación de Cercas',                'Obras',         'Servicio de reparación',                                'Evento', ts],
        ['CAT027', 'Instalación Poste Ganadero T',        'Obras',         'Pieza',                                                 'Evento', ts],
        ['CAT028', 'Reparación de Gaviones',              'Obras',         'Servicio de reparación',                                'Evento', ts],

        // ALMACEN
        ['CAT029', 'Gestión de Almacén',                  'Almacen',       'Almacenamiento materiales y equipos (mensual)',         'Mensual', ts],
        
        // ACTIVIDADES OPERATIVAS ADICIONALES (Requeridas por App)
        ['CAT030', 'Patrullaje Vehicular',                'Operación',     'Recorrido de vigilancia',                               'Diario', ts],
        ['CAT031', 'Patrullaje Peatonal',                 'Operación',     'Recorrido a pie',                                       'Semanal', ts]
    ];

    seedIfEmpty('Catalogo_Actividades',
        ['ID','Nombre','Tipo','Descripcion','Periodicidad','Timestamp'],
        actividadesSeed
    );

    // ============================================================
    // 4. USUARIOS (Estructura Estricta: 1 Supervisor, 1 Líder, 1 Segurista por ACTIVO)
    // ============================================================
    seedIfEmpty('Usuarios',
        ['ID','Nombre','PIN','Rol','Iniciales','Activo','Puesto','ActivoAsignado','NumeroEco','Timestamp','CreadoPor'],
        [
        // ADMIN GENERAL
        ['U001','Sergio A. Jiménez', '1234','ADMIN',     'SJ',true,'Gerente de Proyecto',  'TODOS','',    ts,'SEED'],
        
        // GDN (ACT-01) - Estructura Completa
        ['U010','Javier Soto (Sup GDN)', '1010','SUPERVISOR','JS',true,'Supervisor GDN',     'ACT-01','',   ts,'SEED'],
        ['U004','Pedro Hernández (Líder GDN)', '4444','LIDER',     'PH',true,'Líder GDN',         'ACT-01', 'ECO-01',ts,'SEED'],
        ['U012','Mónica Ruiz (HSE GDN)', '1212','SEGURISTA', 'MR',true,'Segurista GDN',      'ACT-01','',   ts,'SEED'],

        // DEN (ACT-02) - Estructura Completa
        ['U002','Carlos Mendoza (Sup DEN)', '2222','SUPERVISOR','CM',true,'Supervisor DEN',     'ACT-02','',   ts,'SEED'],
        ['U003','Laura García (HSE DEN)',   '3333','SEGURISTA', 'LG',true,'Segurista DEN',      'ACT-02','',   ts,'SEED'],
        ['U005','Miguel Ángel López (Líder DEN)','5555','LIDER',     'ML',true,'Líder DEN',         'ACT-02', 'ECO-02',ts,'SEED'],

        // GDT (ACT-03) - Estructura Completa
        ['U011','Raúl Vargas (Sup GDT)', '1111','SUPERVISOR','RV',true,'Supervisor GDT',     'ACT-03','',   ts,'SEED'],
        ['U007','Ana Martínez (Líder GDT)', '7777','LIDER',     'AM',true,'Líder GDT',         'ACT-03', 'ECO-04',ts,'SEED'],
        ['U013','Sofía Castro (HSE GDT)', '1313','SEGURISTA', 'SC',true,'Segurista GDT',      'ACT-03','',   ts,'SEED'],
        ]
    );

    // ============================================================
    // 5. PERSONAL DE CAMPO (3 Ayudantes por Activo, Sin Acceso al Sistema)
    // ============================================================
    seedIfEmpty('Personal_Campo',
        ['ID','Nombre','Puesto','Cuadrilla','NSS','TipoSangre','ContactoEmergencia','TelefonoEmergencia','FechaIngreso','Activo','Observaciones','Timestamp'],
        [
        // Cuadrilla GDN (ACT-01)
        ['PC01','José Ramírez',     'Ayudante General',       'Cuadrilla GDN','','','','',hoy,'Activo','',ts],
        ['PC02','Luis Fernández',   'Ayudante General',       'Cuadrilla GDN','','','','',hoy,'Activo','',ts],
        ['PC03','Marco Polo',       'Ayudante General',       'Cuadrilla GDN','','','','',hoy,'Activo','',ts],

        // Cuadrilla DEN (ACT-02)
        ['PC04','Francisco Torres', 'Ayudante General',       'Cuadrilla DEN','','','','',hoy,'Activo','',ts],
        ['PC05','Daniel Morales',   'Ayudante General',       'Cuadrilla DEN','','','','',hoy,'Activo','',ts],
        ['PC06','Esteban Quito',    'Ayudante General',       'Cuadrilla DEN','','','','',hoy,'Activo','',ts],

        // Cuadrilla GDT (ACT-03)
        ['PC07','Alejandro Ruiz',   'Ayudante General',       'Cuadrilla GDT','','','','',hoy,'Activo','',ts],
        ['PC08','Benito Juárez',    'Ayudante General',       'Cuadrilla GDT','','','','',hoy,'Activo','',ts],
        ['PC09','Carlos Salinas',   'Ayudante General',       'Cuadrilla GDT','','','','',hoy,'Activo','',ts],
        ]
    );

    // ============================================================
    // 6. PLAN DE TRABAJO (Alineado a Catálogo Real)
    // ============================================================
    seedIfEmpty('Plan_Trabajo',
        ['ID','Tramo','TipoTrabajo','Descripcion','UnidadMedida','CantidadProgramada','CantidadEjecutada','Avance','FechaInicio','FechaFin','Estado','Fase','Observaciones','Timestamp'],
        [
        ['PT001','TRAMO_GDN_01','Desmonte de maleza manual/eléctrica','Desmonte manual en GDN', 'M2', 5000, 0,0,hoy,'','Pendiente','Fase 1','',ts],
        ['PT002','TRAMO_DEN_01','Desmonte de maleza manual/eléctrica','Desmonte manual en DEN Norte', 'M2', 8000, 0,0,hoy,'','Pendiente','Fase 2','',ts],
        ['PT003','TRAMO_DEN_02','Tractor con desmalezadora','Desmonte tractor en DEN Sur', 'Jornada', 5, 0,0,hoy,'','Pendiente','Fase 2','',ts],
        ['PT004','TRAMO_GDT_01','Desbroce con tractor (GDT 36")','Desbroce tractor en GDT', 'M2', 10000, 0,0,hoy,'','Pendiente','Fase 1','',ts],
        ]
    );

    // ============================================================
    // 7. HOJAS OPERATIVAS (Headers Reales - Sin cambios)
    // ============================================================
    ensureSheet('Actividades', ['ID','Tramo','Lider','Tipo','Periodicidad','Fecha','Estado','KmProgramados','ASTId','Timestamp','CreadoPor','ActivoCatalogado']);
    ensureSheet('Patrullajes', ['ID','Actividad','Lider','Fecha','Tramo','KmInicio','KmFin','Tipo','Estatus','Observaciones','Hallazgo','HallazgoDesc','ConFoto','FotoIDs','Timestamp']);
    ensureSheet('Hallazgos', ['ID','PatrullajeId','Hallazgo','Descripcion','Tramo','KP','Estado','Fecha','ReportadoPor','FotoURL']);
    ensureSheet('AST', ['ID','ActividadId','Lider','Fecha','ActividadTipo','PeligrosJSON','RiesgosJSON','ControlesJSON','Firma','FirmadoPor','GeoLat','GeoLng','Timestamp']);
    ensureSheet('Capacitaciones', ['ID','Usuario','Tipo','Nombre','Vigencia','DC3','Estatus']);
    ensureSheet('Aptitudes_Medicas', ['ID','Usuario','Tipo','Vigencia','Medico','Estatus']);
    ensureSheet('Incidentes', ['ID','Fecha','Hora','Lugar','Tipo','Descripcion','Involucrados','LesionTipo','Gravedad','CausaInmediata','CausaRaiz','AccionesJSON','Responsable','Estado','FotoId','Timestamp']);
    ensureSheet('Platicas_5min', ['ID','Fecha','Tema','Expositor','AsistentesJSON','Duracion','Observaciones','FotoId','Timestamp']);
    ensureSheet('Entrega_EPP', ['ID','Fecha','Trabajador','EPPEntregadoJSON','Motivo','EntregadoPor','FirmaRecibe','Observaciones','Timestamp']);
    ensureSheet('Actos_Inseguros', ['ID','PersonalId','Tipo','Descripcion','Severidad','Timestamp','UserId']);
    ensureSheet('Evidencias', ['ID','EntidadId','Tipo','URL','FileId','Timestamp']);
    ensureSheet('Log_Acciones', ['Timestamp','Usuario','Accion','EntidadId','Detalles']);
    ensureSheet('Asistencias', ['ID','ActividadID','Lider','HaySegurista','AsistentesJSON','Timestamp']);
    ensureSheet('KPIs', ['ID','Fecha','Periodo','ResumenJSON','DetalleJSON','Cumplimiento','Timestamp']);
    ensureSheet('Reportes_Extraordinarios', ['ID','Fecha','Hora','ActividadId','TipoActividad','Tramo','Motivo','Descripcion','ReportadoPor','Estado','FechaReprogramacion','ActividadReprogramadaId','EvidenciaFoto','Observaciones','Timestamp']);
    ensureSheet('Movimientos_Inventario', ['ID','Fecha','ItemId','Tipo','Cantidad','Referencia','Usuario','Timestamp']);
    ensureSheet('Movimientos_Herramientas', ['ID','Fecha','Tipo','Articulo','Cantidad','Usuario','RefId','Detalles','Timestamp']);
    ensureSheet('Herramientas_Asignacion', ['ID','Timestamp','Supervisor','Lider','Cuadrilla','ItemsJSON','Firma','Estado','TimestampFin']);
    ensureSheet('Stop_Work', ['ID','Lider','Fecha','Motivo','Estado','LiberadoPor','LiberadoFecha','Timestamp']);
    ensureSheet('Inspecciones', ['ID','Fecha','Inspector','Tipo','Objetivo','ItemsJSON','Resultado','Observaciones','FotoID','Timestamp']);
    ensureSheet('Permisos_Trabajo', ['ID','Fecha','Solicitante','TipoPT','Ubicacion','Descripcion','RiesgosJSON','ControlesJSON','HoraInicio','HoraFin','AutorizadoPor','Estado','FirmaEmisor','FirmaEjecutor','Timestamp']);
    ensureSheet('Emergencias', ['ID','Tipo','Titulo','Fecha','Responsable','ParticipantesJSON','Evaluacion','Observaciones','Estado','Timestamp']);
    ensureSheet('Programa_Salud', ['ID','Tipo','Actividad','Frecuencia','FechaPrograma','FechaEjecucion','Responsable','Estatus','Evidencia','Observaciones','Timestamp']);
    ensureSheet('Auditorias', ['ID','Fecha','Tipo','Auditor','ElementosJSON','Resultado','NCMayores','NCMenores','Observaciones','Estado','PlanAccionJSON','Timestamp']);
    ensureSheet('Listado_Maestro', ['ID','Codigo','Titulo','Elemento','Version','FechaEmision','FechaRevision','Responsable','Ubicacion','Estado','Timestamp']);
    ensureSheet('Residuos', ['ID','Fecha','Tipo','Cantidad','Unidad','Generador','Manifiesto','FotoId','Timestamp']);
    ensureSheet('Extintores', ['IdExtintor','Ubicacion','Fecha','Presion','Seguro','Etiqueta','Manguera','Estado','Inspector','Timestamp']);
    ensureSheet('Comision_Seguridad', ['ID','Fecha','Lider','IntegrantesJSON','Recorrido','HallazgosJSON','Firma','Timestamp']);
    ensureSheet('Observaciones', ['ActividadId','Observacion','Tipo','Timestamp','UserId']);
    ensureSheet('Inspecciones_Herramientas', ['ID','ActividadId','Descripcion','Estado','Timestamp','UserId']);
    ensureSheet('Checklist_NOM', ['ID','Fecha','Norma','ResultadosJSON','UsuarioID']);
    ensureSheet('NOM_Matriz', ['NOM','Titulo','Modulo','EvidenciaRequerida','EvidenciaDigital','Cumplimiento','Estado','Fecha']);
    ensureSheet('SASISOPA', ['ID','Elemento','Cumplimiento','Evidencia','Responsable','Timestamp']);

    // ============================================================
    // 8. EPP (Anexo + NOMs)
    // ============================================================
    seedIfEmpty('EPP',
        ['ID','Nombre','Categoria','NormaCumplimiento','VidaUtil_Meses','Timestamp'],
        [
        ['EPP01','Casco de Seguridad Tipo E',    'Protección Cabeza',      'NOM-115-STPS',24,ts],
        ['EPP02','Lentes de Protección',         'Protección Ocular',      'NOM-017-STPS', 6,ts],
        ['EPP03','Overol / Camisa Manga Larga',  'Ropa Trabajo',           'NOM-017-STPS', 6,ts],
        ['EPP04','Guantes Carnaza/Tigre',        'Protección Manos',       'NOM-017-STPS', 3,ts],
        ['EPP05','Botas Casquillo Dieléctrico',  'Protección Pies',        'NOM-113-STPS',12,ts],
        ['EPP06','Polainas de Carnaza',          'Protección Corte',       'NOM-017-STPS',12,ts],
        ['EPP07','Arnés Cuerpo Completo',        'Trabajos Altura',        'NOM-009-STPS',36,ts],
        ['EPP08','Tapones Auditivos',            'Protección Auditiva',    'NOM-017-STPS', 1,ts],
        ['EPP09','Faja Lumbar',                  'Ergonomía',              'NOM-017-STPS',12,ts],
        ]
    );

    // ============================================================
    // 9. BANCO DE PLÁTICAS (Inspirado en Anexo)
    // ============================================================
    seedIfEmpty('Banco_Platicas',
        ['ID','Tema','Categoria','Contenido','DuracionMin','Timestamp'],
        [
        ['BP01','Reglas de Oro de Seguridad',           'Seguridad',      'Repaso de las 12 reglas de oro de la compañía',            5,ts],
        ['BP02','Uso Obligatorio de EPP',               'Seguridad',      'Casco, lentes, botas, manga larga: Cero Tolerancia',       5,ts],
        ['BP03','Protección contra Fauna Nociva',       'Seguridad',      'Serpientes, insectos en el derecho de vía',                5,ts],
        ['BP04','Prevención de Incendios Forestales',   'Medio Ambiente', 'Manejo de colillas, soldadura, condiciones secas',         5,ts],
        ['BP05','Identificación de Riesgos (AST)',      'Seguridad',      'Importancia de llenar el AST antes de iniciar',            5,ts],
        ['BP06','Manejo de Residuos en Campo',          'Medio Ambiente', 'No dejar basura en el DmV, clasificación',                 5,ts],
        ]
    );

    // ============================================================
    // 10. HERRAMIENTAS (Del Anexo)
    // ============================================================
    seedIfEmpty('Catalogo_Herramientas',
        ['ID','Nombre','Categoria','StockAlmacen','StockCampo'],
        [
        ['HER01','Machete',                     'Manual', 20, 0],
        ['HER02','Azadón',                      'Manual', 10, 0],
        ['HER03','Sierra de Mano',              'Manual', 5, 0],
        ['HER04','Motoguadaña (Desbrozadora)',  'Mecánica', 5, 0],
        ['HER005','Sierra de Cadena (Motosierra)','Mecánica', 3, 0],
        ['HER006','Equipo Compactación',        'Mecánica', 2, 0],
        ['HER007','Pala','Manual',10,0],
        ['HER008','Pico','Manual',10,0],
        ['HER009','Martillo','Manual',10,0],
        ['HER010','Flexómetro','Medición',10,0]
        ]
    );

    Logger.log('========================================');
    Logger.log('🎉 SEED DATABASE COMPLETADO (DATOS REALES LIMPIOS)');
    Logger.log('Total hojas revisadas: ' + ss.getSheets().length);
    Logger.log('========================================');
    return { success: true, message: 'Base de datos sembrada y limpiada' };
    }
