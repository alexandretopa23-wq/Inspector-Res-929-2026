/* ---------- Configuración ---------- */
var SPREADSHEET_ID = '164c4ztO_Lb2Nwu6cq6yi9D4x4hmaHgNb5CkrDNwjST4'; // respaldo si el script no queda vinculado
var GID_DATOS       = 322693342;                // pestaña de datos confirmada
var RAIZ_DRIVE       = 'INSPECCIONES_PISCINAS';  // carpeta raíz en el Drive de la cuenta dedicada

/* Columnas confirmadas 1:1 contra la hoja (encabezado real, A→V — 22 columnas,
   sin "norma aplicable") */
var COL = {
  id:1, capitulo:2, item:3, enfoque:4, prioridad:5,
  sede:6, piscina:7, fecha:8, responsable:9,
  estado:10, riesgo:11, hallazgo:12, accion:13, respCierre:14,
  fechaCompromiso:15, fechaCierre:16, diasRest:17, avance:18,
  evidActual:19, evidCierre:20, linkCarpeta:21, observaciones:22
};
var TOTAL_COLS = 22;

/* Paleta corporativa (coherente con plan-accion-piscinas) */
var C_ENCABEZADO='#212121', C_TITULO='#424242', C_ACENTO='#E65100',
    C_CELDA='#EEEEEE', C_FONDO='#FAFAFA', C_OK='#2E7D32';

/* Traducción de los códigos que manda la PWA (selects de la ficha técnica)
   a texto legible para los anexos del informe. */
var TIPO_USO_LABEL = {
  colectivo:'Público colectivo (abierta al público)',
  restringido:'Uso restringido (no abierta al público)',
  infantil:'Infantil / profundidad menor a 0.6 m',
  similar:'Estructura similar'
};
var BOMBA_FAMILIA_LABEL = {
  superflo:'Pentair SuperFlo, velocidad única',
  superflo_vs:'Pentair SuperFlo VS 2.2 HP, velocidad seleccionable (dato de fábrica)',
  eq:'Pentair EQ Series comercial',
  manual:'Curva real medida en campo (mínimo 4 puntos Q-H)'
};
var BOMBA_VELOCIDAD_VS_LABEL = {
  '1':'Velocidad 1 — 3000 RPM', '2':'Velocidad 2 — 2200 RPM',
  '3':'Velocidad 3 — 1400 RPM', '4':'Velocidad 4 / máxima — 3450 RPM'
};
var FILTRO_TIPO_LABEL = {
  arena:'Arena / medio granular', cartucho:'Cartucho', de:'Tierra de diatomeas (D.E.)'
};

/* ============================================================================
   0. ACCESO A LA HOJA — por gid, no por nombre de pestaña
   ============================================================================ */
function _ss(){
  try{
    var activa = SpreadsheetApp.getActiveSpreadsheet();
    if(activa) return activa;
  }catch(e){ /* script no vinculado, cae al ID fijo */ }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}
function _hojaDatos(){
  var ss = _ss();
  var hojas = ss.getSheets();
  for(var i=0;i<hojas.length;i++){
    if(hojas[i].getSheetId() === GID_DATOS) return hojas[i];
  }
  return ss.getSheets()[0]; // respaldo si el gid cambiara
}

/* ============================================================================
   1. ROUTER doPost
   ============================================================================ */
function doPost(e){
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch(err){ body = {}; }

  if (body.accion === 'foto')      return _json(guardarFoto(body));
  if (body.accion === 'ficha')     return _json(guardarFicha(body));
  if (body.accion === 'informe')   return _json(generarInformeVaso(body.sede, body.piscina, body.fecha));
  if (body.accion === 'dashboard') return _json(obtenerDashboard());
  if (body.rows)                   return _json(guardarFilas(body.rows));

  return _json({ok:false, error:'payload no reconocido'});
}
function _json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================================
   2. CARGA DE FILAS (automática desde la PWA)
   Formato de cada fila (22 columnas, igual a CSV_HEADERS de index.html):
   [ID, capitulo, item, enfoque, prioridad, sede, piscina, fecha,
    responsable, estado, riesgo, hallazgo, accion, respCierre,
    fechaCompromiso, fechaCierre, diasRestantes, avance,
    evidenciaActual, evidenciaCierre, linkCarpeta, observaciones]

   La hoja NO se pre-carga con los ítems del checklist: arranca solo con el
   encabezado. Cada fila la crea la app la primera vez que alguien responde
   ese ítem en una inspección real (Sede+Piscina+Fecha reales), y desde ahí
   el UPSERT por clave (ID + Sede + Piscina + Fecha) la va actualizando.
   ============================================================================ */
function guardarFilas(rows){
  var lock = LockService.getScriptLock();
  lock.waitLock(30000); // evita que dos capítulos sincronizando en paralelo se pisen
  try{
    var sh = _hojaDatos();
    var last = sh.getLastRow();
    var datos = last>=2 ? sh.getRange(2,1,last-1,TOTAL_COLS).getValues() : [];

    // Índice por clave (ID+Sede+Piscina+Fecha) construido una sola vez.
    // Antes cada fila entrante disparaba un barrido lineal sobre TODA la
    // hoja (_indiceFila recorriendo `datos` fila por fila) — con 247 ítems
    // por sede/fecha acumulados durante meses, cada sync se iba volviendo
    // más lento y arriesgaba el timeout de 6 min de Apps Script. Con un
    // Map la búsqueda es O(1) por fila entrante.
    var indice = {};
    for(var i=0;i<datos.length;i++){
      indice[_claveFila(datos[i])] = i;
    }

    var nuevas=0, actualizadas=0;
    var porAgregar = [];
    var indicePorAgregar = {}; // dedup dentro del mismo lote (misma clave dos veces → gana la última)
    rows.forEach(function(row){
      row = _normalizarFila(row);
      var clave = _claveFila(row);
      if(indice.hasOwnProperty(clave)){
        var idx = indice[clave];
        sh.getRange(idx+2, 1, 1, TOTAL_COLS).setValues([row]);
        datos[idx] = row;
        actualizadas++;
      }else if(indicePorAgregar.hasOwnProperty(clave)){
        porAgregar[indicePorAgregar[clave]] = row;
      }else{
        indicePorAgregar[clave] = porAgregar.length;
        porAgregar.push(row);
      }
    });

    // Las filas nuevas se escriben en un solo bloque (setValues) en vez de
    // un appendRow() por fila — mismo motivo: menos llamadas a la hoja.
    if(porAgregar.length){
      sh.getRange(sh.getLastRow()+1, 1, porAgregar.length, TOTAL_COLS).setValues(porAgregar);
      nuevas = porAgregar.length;
    }

    return {ok:true, nuevas:nuevas, actualizadas:actualizadas};
  }catch(err){
    return {ok:false, error:String(err)};
  }finally{
    lock.releaseLock();
  }
}

/** Completa la fila a 22 columnas si llega más corta, sin inventar datos. */
function _normalizarFila(row){
  var r = row.slice(0, TOTAL_COLS);
  while(r.length < TOTAL_COLS) r.push('');
  return r;
}

/** Clave de upsert: ID + Sede + Piscina + Fecha, con la fecha normalizada a
 *  'yyyy-MM-dd' con _fechaStr() en ambos lados de la comparación.
 *
 *  FIX: la versión anterior (_indiceFila) comparaba con String(row[fecha]),
 *  y getValues() devuelve un objeto Date real cuando la columna H está
 *  formateada como fecha. String(Date) da algo como
 *  "Mon Aug 03 2026 00:00:00 GMT-0500 (COT)", que nunca coincide con el
 *  '2026-08-03' que manda la PWA. Resultado: el upsert nunca encontraba la
 *  fila existente, caía siempre a "nueva" y cada sincronización duplicaba
 *  el registro completo. _marcarEvidenciaEnHoja y generarInformeVaso ya
 *  usaban _fechaStr() para esto — faltaba aplicarlo aquí, que es la ruta
 *  más transitada (cada sync de cada capítulo). */
function _claveFila(row){
  return String(row[COL.id-1])+'|'+String(row[COL.sede-1])+'|'+String(row[COL.piscina-1])+'|'+_fechaStr(row[COL.fecha-1]);
}

/* ============================================================================
   3. INGESTA DE FOTOS
   Ruta: RAIZ / SEDE / PISCINA / FECHA / <ITEMID>__<tipo>__<n>.jpg
   Carpeta por sede, subcarpeta por piscina, subcarpeta de fecha — así el
   histórico de cada inspección queda separado y no se pisa mes a mes.

   Nota: esta función no distingue el estado del ítem (Cumple / No cumple /
   etc.) — solo necesita sede+piscina+fecha+itemId. Las fotos de evidencia
   de cumplimiento (nuevas en la PWA) llegan por la misma ruta, sin cambios
   aquí.
   ============================================================================ */
function guardarFoto(p){
  try{
    if(!p.sede || !p.piscina || !p.fecha || !p.itemId || !p.data){
      return {ok:false, error:'payload incompleto'};
    }
    var carpeta = _carpetaRuta([RAIZ_DRIVE, p.sede, p.piscina, p.fecha]);
    var nombre  = p.nombre || (p.itemId + '__actual__1.jpg');

    var previos = carpeta.getFilesByName(nombre);
    while(previos.hasNext()) previos.next().setTrashed(true); // reemplaza, no duplica

    var blob = Utilities.newBlob(Utilities.base64Decode(p.data), p.mime||'image/jpeg', nombre);
    var file = carpeta.createFile(blob);
    file.setDescription(JSON.stringify({
      item:p.itemId, sede:p.sede, piscina:p.piscina,
      fecha:p.fecha, responsable:p.responsable||'', subida:new Date().toISOString()
    }));

    _marcarEvidenciaEnHoja(p.sede, p.piscina, p.fecha, p.itemId, carpeta.getUrl());

    return {ok:true, fileId:file.getId(), url:file.getUrl(), carpeta:carpeta.getUrl()};
  }catch(err){
    return {ok:false, error:String(err)};
  }
}

function _carpetaRuta(partes){
  var actual = _carpetaHija(DriveApp.getRootFolder(), partes[0]);
  for(var i=1;i<partes.length;i++) actual = _carpetaHija(actual, String(partes[i]));
  return actual;
}
function _carpetaHija(padre, nombre){
  nombre = String(nombre).replace(/[\/\\]/g,'-').trim() || 'SIN_DATO';
  var it = padre.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : padre.createFolder(nombre);
}

/** Escribe el link de carpeta en la fila correspondiente, solo si está vacío. */
function _marcarEvidenciaEnHoja(sede, piscina, fecha, itemId, url){
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try{
    var sh = _hojaDatos();
    var last = sh.getLastRow();
    if(last<2) return;
    var datos = sh.getRange(2,1,last-1,TOTAL_COLS).getValues();
    for(var i=0;i<datos.length;i++){
      if(String(datos[i][COL.id-1])===itemId &&
         String(datos[i][COL.sede-1])===sede &&
         String(datos[i][COL.piscina-1])===piscina &&
         _fechaStr(datos[i][COL.fecha-1])===String(fecha)){
        if(!datos[i][COL.linkCarpeta-1]) sh.getRange(i+2, COL.linkCarpeta).setValue(url);
        return;
      }
    }
  }catch(err){ /* no bloquea la subida de la foto si la hoja falla */ }
  finally{ lock.releaseLock(); }
}
function _fechaStr(v){
  if(v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v||'').slice(0,10);
}

/* ============================================================================
   3.5. FICHA TÉCNICA DEL VASO (geometría, bomba, tubería, filtro, resultados
   de cálculo hidráulico y de aforo/dotación) — alimenta los Anexos A-D del
   informe. Se guarda como UN SOLO JSON por Sede+Piscina+Fecha en su propia
   hoja, no columna por columna: la ficha tiene ~25 campos que van a seguir
   creciendo (motor hidráulico, curvas manuales, etc.) y mantener ese
   esquema sincronizado 1:1 entre la PWA y esta hoja sería frágil. El
   generador de informes hace JSON.parse al leerla.
   ============================================================================ */
var NOMBRE_HOJA_FICHA = 'Ficha_Tecnica';
var COL_FICHA = { sede:1, piscina:2, fecha:3, responsable:4, json:5, actualizado:6 };
var TOTAL_COLS_FICHA = 6;

function _hojaFicha(){
  var ss = _ss();
  var sh = ss.getSheetByName(NOMBRE_HOJA_FICHA);
  if(!sh){
    sh = ss.insertSheet(NOMBRE_HOJA_FICHA);
    sh.getRange(1,1,1,TOTAL_COLS_FICHA).setValues([['Sede','Piscina','Fecha','Responsable','FichaJSON','ActualizadoEn']]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function _claveFicha(sede, piscina, fecha){
  return String(sede)+'|'+String(piscina)+'|'+_fechaStr(fecha);
}

/** Guarda/actualiza la ficha técnica completa. Upsert por Sede+Piscina+Fecha
 *  (misma clave que usa guardarFilas para el checklist, sin el ID de ítem). */
function guardarFicha(body){
  if(!body.sede || !body.piscina || !body.fecha){
    return {ok:false, error:'payload incompleto (sede/piscina/fecha requeridos)'};
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try{
    var sh = _hojaFicha();
    var last = sh.getLastRow();
    var datos = last>=2 ? sh.getRange(2,1,last-1,TOTAL_COLS_FICHA).getValues() : [];
    var claveNueva = _claveFicha(body.sede, body.piscina, body.fecha);
    var fila = [body.sede, body.piscina, body.fecha, body.responsable||'', JSON.stringify(body.ficha||{}), new Date().toISOString()];
    for(var i=0;i<datos.length;i++){
      var clave = _claveFicha(datos[i][COL_FICHA.sede-1], datos[i][COL_FICHA.piscina-1], datos[i][COL_FICHA.fecha-1]);
      if(clave===claveNueva){
        sh.getRange(i+2,1,1,TOTAL_COLS_FICHA).setValues([fila]);
        return {ok:true, accion:'actualizada'};
      }
    }
    sh.appendRow(fila);
    return {ok:true, accion:'creada'};
  }catch(err){
    return {ok:false, error:String(err)};
  }finally{
    lock.releaseLock();
  }
}

/** Lee y parsea la ficha de un vaso; null si no hay ficha guardada para esa
 *  clave (el informe sigue generándose igual, solo sin los Anexos A-D). */
function _obtenerFicha(sede, piscina, fecha){
  var sh = _hojaFicha();
  var last = sh.getLastRow();
  if(last<2) return null;
  var datos = sh.getRange(2,1,last-1,TOTAL_COLS_FICHA).getValues();
  var clave = _claveFicha(sede, piscina, fecha);
  for(var i=0;i<datos.length;i++){
    var claveFila = _claveFicha(datos[i][COL_FICHA.sede-1], datos[i][COL_FICHA.piscina-1], datos[i][COL_FICHA.fecha-1]);
    if(claveFila===clave){
      try{ return JSON.parse(datos[i][COL_FICHA.json-1] || '{}'); }
      catch(e){ return null; }
    }
  }
  return null;
}

/* ============================================================================
   4. GENERADOR DE INFORME POR VASO (Google Doc → PDF)
   ============================================================================ */
function generarInformeVaso(sede, piscina, fecha){
  var sh = _hojaDatos();
  var last = sh.getLastRow();
  if(last<2) return {ok:false, error:'La hoja no tiene datos aún'};
  var todo = sh.getRange(2,1,last-1,TOTAL_COLS).getValues();

  var filas = todo.filter(function(f){
    return String(f[COL.sede-1])===sede && String(f[COL.piscina-1])===piscina &&
           _fechaStr(f[COL.fecha-1])===String(fecha);
  });
  if(!filas.length) return {ok:false, error:'Sin registros para ese vaso y fecha'};

  var m = _metricas(filas);
  var fotos = _indiceFotos(sede, piscina, fecha);
  var ficha = _obtenerFicha(sede, piscina, fecha);

  var doc  = DocumentApp.create('Informe_Inspeccion_'+sede+'_'+piscina+'_'+fecha);
  var body = doc.getBody();
  body.setPageWidth(595).setPageHeight(842); // A4 en puntos
  body.clear();

  _portada(body, sede, piscina, fecha, filas[0][COL.responsable-1], m);
  _objetivo(body);
  _tableroKPI(body, m);
  _tablaCapitulos(body, m);
  _hallazgosConFotos(body, filas, fotos);
  _planAccion(body, filas);
  _conclusion(body, m, sede, piscina);
  _responsabilidades(body, filas[0][COL.responsable-1]);
  _anexoFotografico(body, filas, fotos);
  // Los Anexos A-D solo salen si el inspector diligenció la ficha técnica
  // en la PWA para este vaso+fecha (botón "📐 Ficha") — un informe sin
  // ficha se genera igual, simplemente sin estas 4 secciones finales.
  if(ficha){
    _anexoFichaEscenario(body, ficha);
    _anexoFichaHidraulica(body, ficha);
    _anexoDimensionamiento(body, ficha);
    _anexoMemoriaCalculo(body, ficha);
  }

  doc.saveAndClose();

  var carpeta = _carpetaRuta([RAIZ_DRIVE, sede, piscina, fecha]);
  var pdf = carpeta.createFile(DriveApp.getFileById(doc.getId()).getAs('application/pdf'))
                   .setName('Informe_'+sede+'_'+piscina+'_'+fecha+'.pdf');
  return {ok:true, docUrl:doc.getUrl(), pdfUrl:pdf.getUrl(), metricas:m};
}

/* Ítems que NO tienen numeral de anclaje en la Res. 929 (pH, cloro,
   alcalinidad, microbiológicos, dosificación, muestreo, manejo de
   desviaciones — dominio de la Res. 234, que esta app no evalúa). Se siguen
   registrando con total normalidad — aparecen en Hallazgos y Plan de acción
   igual que cualquier ítem — pero NO participan del % de cumplimiento
   normativo, porque no es un criterio que la 929 exija. Misma lista que
   FUERA_ALCANCE_929 en index.html (duplicada a propósito: son dos
   codebases distintas sin import compartido; si la lista cambia, hay que
   actualizar ambas). */
var FUERA_ALCANCE_929 = {
  'CHK-152':1,'CHK-153':1,'CHK-154':1,'CHK-157':1,'CHK-158':1,'CHK-159':1,'CHK-160':1,'CHK-161':1,
  'CHK-231':1,'CHK-232':1,'CHK-239':1,'CHK-240':1,'CHK-241':1,'CHK-242':1,'CHK-243':1,'CHK-244':1
};

/* ---------- Métricas ejecutivas ---------- */
function _metricas(filas){
  var m = {total:filas.length, cumple:0, noCumple:0, enProceso:0, pendiente:0, noAplica:0, fueraAlcance:0,
           cumpleEnAlcance:0, baseEnAlcance:0,
           critico:0, alto:0, medio:0, bajo:0, vencidos:0, sinFecha:0, avanceProm:0,
           porCapitulo:{}, criticosAltos:[]};
  var hoy = new Date(); hoy.setHours(0,0,0,0);
  var sumAvance=0, conAvance=0;

  filas.forEach(function(f){
    var id = String(f[COL.id-1]);
    var esFueraAlcance = !!FUERA_ALCANCE_929[id];
    var est = String(f[COL.estado-1]||''), rie = String(f[COL.riesgo-1]||'');
    var cap = String(f[COL.capitulo-1]||'Sin capítulo');
    // FIX: se agrega noAplica como contador propio por capítulo (antes iba
    // mezclado dentro de "otros" junto con En proceso/Pendiente), para que
    // _tablaCapitulos pueda excluirlo del denominador igual que el cálculo
    // global (ver _tablaCapitulos más abajo). fueraAlcance sigue el mismo
    // patrón: contador propio, excluido del % pero visible en el tablero.
    if(!m.porCapitulo[cap]) m.porCapitulo[cap] = {total:0, cumple:0, noCumple:0, noAplica:0, fueraAlcance:0, cumpleEnAlcance:0, otros:0};
    m.porCapitulo[cap].total++;
    if(esFueraAlcance){ m.fueraAlcance++; m.porCapitulo[cap].fueraAlcance++; }

    // Los conteos brutos (cumple/noCumple/...) incluyen TODO, fuera de
    // alcance o no — Hallazgos y Plan de acción filtran directamente sobre
    // `filas`, no sobre estos contadores, así que un ítem "No cumple" fuera
    // de alcance igual aparece ahí. Lo único que cambia con el alcance es
    // cumpleEnAlcance (numerador del %) y baseEnAlcance (denominador).
    if(est==='Cumple'){
      m.cumple++; m.porCapitulo[cap].cumple++;
      if(!esFueraAlcance){ m.cumpleEnAlcance++; m.porCapitulo[cap].cumpleEnAlcance++; }
    }
    else if(est==='No cumple'){ m.noCumple++; m.porCapitulo[cap].noCumple++; }
    else if(est==='En proceso'){ m.enProceso++; m.porCapitulo[cap].otros++; }
    else if(est==='Pendiente'){ m.pendiente++; m.porCapitulo[cap].otros++; }
    else if(est==='No aplica'){ m.noAplica++; m.porCapitulo[cap].noAplica++; }

    if(rie==='Critico'||rie==='Crítico') m.critico++;
    else if(rie==='Alto') m.alto++;
    else if(rie==='Medio') m.medio++;
    else if(rie==='Bajo') m.bajo++;

    if(rie==='Critico'||rie==='Crítico'||rie==='Alto') m.criticosAltos.push(f);

    var fc = f[COL.fechaCompromiso-1], cierre = f[COL.fechaCierre-1];
    if(est!=='Cumple' && est!=='No aplica'){
      if(!fc) m.sinFecha++;
      else if(!cierre && new Date(fc) < hoy) m.vencidos++;
    }
    var av = f[COL.avance-1];
    if(av!=='' && av!=null && !isNaN(av)){ sumAvance += Number(av); conAvance++; }
  });

  m.baseEnAlcance = m.total - m.noAplica - m.fueraAlcance;
  m.pctCumplimiento = m.baseEnAlcance>0 ? Math.round(100*m.cumpleEnAlcance/m.baseEnAlcance) : 0;
  m.avanceProm = conAvance? Math.round(sumAvance/conAvance) : 0;
  return m;
}

/* ---------- Dashboard de avances (todas las sedes/piscinas/fechas) ----------
   _metricas() ya es genérica sobre cualquier conjunto de filas (no asume una
   sola sede+piscina+fecha, como en el informe individual) — así que el
   dashboard la reutiliza tal cual sobre subconjuntos agrupados, en vez de
   reimplementar el conteo. Sirve TODA la hoja de una sola pasada: no hay
   filtro de fecha porque "avance" es del estado actual acumulado, no de un
   corte puntual. */
function obtenerDashboard(){
  var sh = _hojaDatos();
  var last = sh.getLastRow();
  var todo = last>=2 ? sh.getRange(2,1,last-1,TOTAL_COLS).getValues() : [];
  if(!todo.length) return {ok:true, vacio:true};

  var global = _metricas(todo);

  var porSede = {};
  todo.forEach(function(f){
    var sede = String(f[COL.sede-1]||'Sin sede');
    if(!porSede[sede]) porSede[sede] = [];
    porSede[sede].push(f);
  });
  var sedes = Object.keys(porSede).map(function(sede){
    var m = _metricas(porSede[sede]);
    return {sede:sede, total:m.total, pctCumplimiento:m.pctCumplimiento, noCumple:m.noCumple,
             vencidos:m.vencidos, critico:m.critico, alto:m.alto};
  }).sort(function(a,b){ return a.sede.localeCompare(b.sede); });

  // Cada combinación Sede+Piscina es un vaso — se resume con su fecha de
  // inspección más reciente vista en la hoja (upsert va actualizando la
  // misma fila, así que la fecha más alta es la vigente).
  var porVaso = {};
  todo.forEach(function(f){
    var key = String(f[COL.sede-1])+'␟'+String(f[COL.piscina-1]);
    if(!porVaso[key]) porVaso[key] = {sede:String(f[COL.sede-1]), piscina:String(f[COL.piscina-1]), filas:[], fechas:{}};
    porVaso[key].filas.push(f);
    var fecha = _fechaStr(f[COL.fecha-1]);
    if(fecha) porVaso[key].fechas[fecha] = 1;
  });
  var vasos = Object.keys(porVaso).map(function(key){
    var v = porVaso[key];
    var m = _metricas(v.filas);
    var fechas = Object.keys(v.fechas).sort();
    var ultima = fechas.length ? fechas[fechas.length-1] : null;
    // Responsable de la inspección más reciente de este vaso.
    var resp = '';
    v.filas.forEach(function(f){
      if(_fechaStr(f[COL.fecha-1])===ultima && !resp) resp = String(f[COL.responsable-1]||'');
    });
    return {
      sede:v.sede, piscina:v.piscina, responsable:resp,
      ultimaFecha: ultima,
      total:m.total, pctCumplimiento:m.pctCumplimiento, noCumple:m.noCumple,
      vencidos:m.vencidos, critico:m.critico, alto:m.alto
    };
  }).sort(function(a,b){ return String(b.ultimaFecha||'').localeCompare(String(a.ultimaFecha||'')); });

  // Tendencia: cuántos vasos se inspeccionaron en cada fecha. Es el único
  // eje temporal real que hay en la hoja — no hay timestamp por ítem.
  var porFecha = {};
  todo.forEach(function(f){
    var fecha = _fechaStr(f[COL.fecha-1]);
    if(!fecha) return;
    if(!porFecha[fecha]) porFecha[fecha] = {};
    porFecha[fecha][String(f[COL.sede-1])+'␟'+String(f[COL.piscina-1])] = 1;
  });
  var tendencia = Object.keys(porFecha).sort().map(function(fecha){
    return {fecha:fecha, vasos:Object.keys(porFecha[fecha]).length};
  });

  // Evidencia faltante: hallazgos "No cumple" sin foto del estado actual —
  // el dato que más le importa a un auditor externo.
  var hallazgosNoCumple = 0, evidenciaFaltante = 0;
  todo.forEach(function(f){
    if(String(f[COL.estado-1]||'')==='No cumple'){
      hallazgosNoCumple++;
      if(!String(f[COL.evidActual-1]||'').trim()) evidenciaFaltante++;
    }
  });

  var capitulos = Object.keys(global.porCapitulo).map(function(cap){
    var c = global.porCapitulo[cap];
    var base = c.total - c.noAplica - c.fueraAlcance;
    var pct = base>0 ? Math.round(100*c.cumpleEnAlcance/base) : 0;
    return {capitulo:cap, total:c.total, cumple:c.cumple, noCumple:c.noCumple,
             fueraAlcance:c.fueraAlcance||0, pctCumplimiento:pct};
  }).sort(function(a,b){ return a.pctCumplimiento - b.pctCumplimiento; });

  return {
    ok:true,
    actualizadoEn: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    global: {
      total:global.total, pctCumplimiento:global.pctCumplimiento,
      cumple:global.cumple, noCumple:global.noCumple, enProceso:global.enProceso,
      pendiente:global.pendiente, noAplica:global.noAplica, fueraAlcance:global.fueraAlcance,
      critico:global.critico, alto:global.alto, medio:global.medio, bajo:global.bajo,
      vencidos:global.vencidos, sinFecha:global.sinFecha, avanceProm:global.avanceProm,
      hallazgosNoCumple:hallazgosNoCumple, evidenciaFaltante:evidenciaFaltante
    },
    sedes: sedes,
    vasos: vasos,
    capitulos: capitulos,
    tendencia: tendencia
  };
}

/* ---------- Índice de fotos por ítem (máx 2 por tipo) ---------- */
function _indiceFotos(sede, piscina, fecha){
  var idx = {};
  try{
    var carpeta = _carpetaRuta([RAIZ_DRIVE, sede, piscina, fecha]);
    var it = carpeta.getFiles();
    while(it.hasNext()){
      var f = it.next();
      var n = f.getName();                       // CHK-047__actual__1.jpg
      var p = n.split('__');
      if(p.length < 3) continue;
      var itemId = p[0], tipo = p[1];
      if(!idx[itemId]) idx[itemId] = {actual:[], cierre:[]};
      if(idx[itemId][tipo] && idx[itemId][tipo].length < 2) idx[itemId][tipo].push(f);
    }
  }catch(err){ /* sin fotos para este vaso/fecha */ }
  return idx;
}

/* ---------- Secciones del documento ---------- */
function _portada(body, sede, piscina, fecha, responsable, m){
  var t = body.appendParagraph('INFORME DE INSPECCIÓN NORMATIVA');
  t.setHeading(DocumentApp.ParagraphHeading.TITLE)
   .setForegroundColor(C_ENCABEZADO).setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('Resolución 929 de 2026 — Ministerio de Salud y Protección Social')
      .setForegroundColor(C_TITULO).setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('');
  var tabla = body.appendTable([
    ['Sede', String(sede)],
    ['Vaso / estructura', String(piscina)],
    ['Fecha de inspección', String(fecha)],
    ['Responsable de la inspección', String(responsable||'—')],
    ['Cumplimiento global', m.pctCumplimiento + '%'],
    ['Fecha de generación', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')]
  ]);
  _estiloTabla(tabla, true);
  body.appendPageBreak();
}

function _objetivo(body){
  _h1(body, '1. Objetivo');
  body.appendParagraph(
    'Verificar el cumplimiento de los criterios técnicos constructivos y de seguridad ' +
    'establecidos en la Resolución 929 de 2026 y su Anexo Técnico para el vaso de piscina ' +
    'evaluado, documentar las condiciones observadas con evidencia fotográfica, clasificar ' +
    'los hallazgos por nivel de riesgo y establecer el plan de acción con responsables y ' +
    'fechas de cierre.'
  ).setForegroundColor(C_TITULO);
}

function _tableroKPI(body, m){
  _h1(body, '2. Tablero ejecutivo');
  var t = body.appendTable([
    ['Indicador','Valor','Lectura'],
    ['Cumplimiento global (Res. 929)', m.pctCumplimiento+'%', _semaforo(m.pctCumplimiento)],
    ['Ítems evaluados', String(m.total), '—'],
    ['Cumple', String(m.cumple), '—'],
    ['No cumple', String(m.noCumple), m.noCumple>0?'Requiere acción':'Sin desviaciones'],
    ['En proceso', String(m.enProceso), '—'],
    ['Pendiente', String(m.pendiente), '—'],
    ['No aplica', String(m.noAplica), 'Excluido del cálculo'],
    ['Fuera del alcance de la Res. 929', String(m.fueraAlcance), 'Excluido del cálculo — dominio de la Res. 234'],
    ['Riesgo crítico', String(m.critico), m.critico>0?'ATENCIÓN INMEDIATA':'Sin críticos'],
    ['Riesgo alto', String(m.alto), m.alto>0?'Prioritario':'—'],
    ['Hallazgos vencidos', String(m.vencidos), m.vencidos>0?'Fuera de plazo':'Al día'],
    ['Hallazgos sin fecha compromiso', String(m.sinFecha), m.sinFecha>0?'Asignar fecha':'Completo'],
    ['Avance promedio de cierre', m.avanceProm+'%', '—']
  ]);
  _estiloTabla(t, true);
}
function _semaforo(pct){
  if(pct>=95) return 'Conforme';
  if(pct>=85) return 'Aceptable con observaciones';
  if(pct>=70) return 'Deficiente';
  return 'Crítico';
}

function _tablaCapitulos(body, m){
  _h1(body, '3. Cumplimiento por capítulo normativo');
  var datos = [['Capítulo','Ítems','Cumple','No cumple','Fuera de alcance','% cumplimiento (Res. 929)']];
  Object.keys(m.porCapitulo).forEach(function(cap){
    var c = m.porCapitulo[cap];
    // FIX histórico: antes el % por capítulo se calculaba contra c.total
    // (incluyendo "No aplica"), mientras que el % global excluía "No
    // aplica" del denominador — no reconciliaban. Ahora ambos usan el mismo
    // criterio: base = total - noAplica - fueraAlcance, y el numerador es
    // cumpleEnAlcance (no cumple, que sí incluiría ítems sin numeral en la
    // 929 y podría superar el 100% frente a la base reducida).
    var base = c.total - c.noAplica - c.fueraAlcance;
    var pct = base>0 ? Math.round(100*c.cumpleEnAlcance/base) : 0;
    datos.push([cap, String(c.total), String(c.cumple), String(c.noCumple), String(c.fueraAlcance||0), pct+'%']);
  });
  var t = body.appendTable(datos);
  _estiloTabla(t, true);
}

function _hallazgosConFotos(body, filas, fotos){
  _h1(body, '4. Hallazgos');
  body.appendParagraph(
    'La evidencia fotográfica de cada ítem, incluidos los de esta sección, se presenta ' +
    'de forma consolidada en el Anexo técnico — Registro fotográfico (sección 8).'
  ).setForegroundColor(C_TITULO).setFontSize(9).setItalic(true);
  var conHallazgo = filas.filter(function(f){
    var est = String(f[COL.estado-1]||'');
    return est==='No cumple' || est==='En proceso' || est==='Pendiente';
  });

  if(!conHallazgo.length){
    body.appendParagraph('No se registraron hallazgos abiertos en esta inspección.')
        .setForegroundColor(C_OK);
    return;
  }

  var peso = {'Critico':0,'Crítico':0,'Alto':1,'Medio':2,'Bajo':3,'Sin riesgo':4,'':5};
  conHallazgo.sort(function(a,b){
    var pa = peso[String(a[COL.riesgo-1])], pb = peso[String(b[COL.riesgo-1])];
    return (pa!==undefined?pa:5) - (pb!==undefined?pb:5); // ojo: peso puede ser 0 (Crítico), "0||5" lo rompía
  });

  conHallazgo.forEach(function(f){
    var itemId = String(f[COL.id-1]);
    var riesgo = String(f[COL.riesgo-1]||'Sin clasificar');
    var esCritico = (riesgo==='Critico'||riesgo==='Crítico'||riesgo==='Alto');

    var h = body.appendParagraph(itemId + ' · ' + String(f[COL.item-1]));
    h.setHeading(DocumentApp.ParagraphHeading.HEADING3)
     .setForegroundColor(esCritico ? C_ACENTO : C_TITULO);

    var t = body.appendTable([
      ['Capítulo', String(f[COL.capitulo-1]||'—')],
      ['Estado', String(f[COL.estado-1]||'—')],
      ['Nivel de riesgo', riesgo],
      ['Hallazgo observado', String(f[COL.hallazgo-1]||'—')],
      ['Acción correctiva', String(f[COL.accion-1]||'—')],
      ['Responsable de cierre', String(f[COL.respCierre-1]||'—')],
      ['Fecha compromiso', _fechaStr(f[COL.fechaCompromiso-1])||'—'],
      ['% avance', String(f[COL.avance-1]!==''?f[COL.avance-1]+'%':'—')]
    ]);
    _estiloTabla(t, false);
    body.appendParagraph('');
  });

  // Esta sección es la narrativa de hallazgo (riesgo, acción, responsable,
  // plazo) para No cumple/En proceso/Pendiente — sin fotos incrustadas, para
  // no duplicar imágenes en el PDF. TODAS las fotos (de estos ítems y de los
  // que sí cumplen) quedan en un único lugar: el Anexo técnico — Registro
  // fotográfico (_anexoFotografico, sección 8), catálogo ordenado por ID con
  // fecha + observación/hallazgo según corresponda.
}

/* ---------- Anexo técnico: registro fotográfico ----------
   Catálogo ordenado por ID de ítem: uno por uno, con fecha de inspección y
   la observación (si el ítem cumple) o el hallazgo (si no cumple/está en
   proceso o pendiente) — luego las fotos. Cubre TODOS los ítems con al
   menos una foto adjunta, sin importar el estado; por eso los ítems "No
   cumple" pueden aparecer tanto aquí como en la sección 4 (que es la
   narrativa de riesgo/acción/responsable, no un registro fotográfico). */
function _anexoFotografico(body, filas, fotos){
  body.appendPageBreak();
  _h1(body, '8. Anexo técnico — Registro fotográfico');
  body.appendParagraph(
    'Evidencia fotográfica de los ítems verificados, en orden de ítem, con la fecha de ' +
    'inspección y la observación o hallazgo correspondiente según el estado registrado.'
  ).setForegroundColor(C_TITULO).setFontSize(9);

  var conFoto = filas.filter(function(f){
    var fi = fotos[String(f[COL.id-1])];
    return fi && (fi.actual.length || fi.cierre.length);
  });

  if(!conFoto.length){
    body.appendParagraph('No se registró evidencia fotográfica en esta inspección.')
        .setForegroundColor(C_TITULO).setItalic(true);
    return;
  }

  conFoto.sort(function(a,b){
    return String(a[COL.id-1]).localeCompare(String(b[COL.id-1]), 'es', {numeric:true});
  });

  conFoto.forEach(function(f){
    var itemId = String(f[COL.id-1]);
    var est = String(f[COL.estado-1]||'—');
    var esConforme = (est==='Cumple');
    var etiquetaObs = esConforme ? 'Observación' : 'Hallazgo / condición observada';
    var textoObs = esConforme ? String(f[COL.observaciones-1]||'—') : String(f[COL.hallazgo-1]||'—');

    var h = body.appendParagraph(itemId + ' · ' + String(f[COL.item-1]));
    h.setHeading(DocumentApp.ParagraphHeading.HEADING3).setForegroundColor(C_TITULO);

    var t = body.appendTable([
      ['Fecha de inspección', _fechaStr(f[COL.fecha-1])||'—'],
      ['Estado', est],
      [etiquetaObs, textoObs]
    ]);
    _estiloTabla(t, false);

    var fi = fotos[itemId];
    _insertarFotos(body, fi.actual, 'Fotografía — estado actual');
    _insertarFotos(body, fi.cierre, 'Fotografía — cierre');
    body.appendParagraph('');
  });
}

/** Inserta hasta 2 fotos en una fila de tabla de 2 columnas, escaladas a ancho A4. */
function _insertarFotos(body, archivos, rotulo){
  if(!archivos || !archivos.length) return;
  body.appendParagraph(rotulo).setForegroundColor(C_TITULO).setBold(true).setFontSize(9);
  var tabla = body.appendTable();
  var fila = tabla.appendTableRow();
  archivos.slice(0,2).forEach(function(file){
    var celda = fila.appendTableCell('');
    var img = celda.appendImage(file.getBlob());
    var maxW = archivos.length>1 ? 230 : 400;
    var esc = Math.min(1, maxW / img.getWidth());
    img.setWidth(Math.round(img.getWidth()*esc));
    img.setHeight(Math.round(img.getHeight()*esc));
    celda.appendParagraph(file.getName()).setFontSize(7).setForegroundColor(C_TITULO);
  });
  if(archivos.length===1) fila.appendTableCell('');
  tabla.setBorderColor(C_CELDA);
}

/* ---------- Anexos técnicos A-D: ficha del escenario ----------
   Leen ficha.motorResultado y ficha.aforoResultado tal como los calculó y
   persistió la PWA (mismo objeto, sin recalcular nada acá) — el informe
   documenta lo que el inspector vio en pantalla, no una versión distinta. */
function _anexoFichaEscenario(body, ficha){
  body.appendPageBreak();
  _h1(body, '9. Anexo técnico A — Ficha del escenario');
  var area = ficha.areaManual || ((ficha.largo && ficha.ancho) ? (ficha.largo*ficha.ancho) : null);
  var t = body.appendTable([
    ['Largo del espejo de agua', ficha.largo!=null ? ficha.largo+' m' : '—'],
    ['Ancho del espejo de agua', ficha.ancho!=null ? ficha.ancho+' m' : '—'],
    ['Área directa (si forma irregular)', ficha.areaManual!=null ? ficha.areaManual+' m²' : '—'],
    ['Área usada en los cálculos', area!=null ? area.toFixed(1)+' m²' : '— sin dato —'],
    ['Profundidad máxima', ficha.profMax!=null ? ficha.profMax+' m' : '—'],
    ['Profundidad mínima', ficha.profMin!=null ? ficha.profMin+' m' : '—'],
    ['Profundidad intermedia', ficha.profIntermedia!=null ? ficha.profIntermedia+' m' : '—'],
    ['Tipo de uso del estanque', TIPO_USO_LABEL[ficha.tipoUso] || '— sin dato —'],
    ['¿Bajo cubierta / recinto cerrado?', ficha.cubierta==='si' ? 'Sí' : (ficha.cubierta==='no' ? 'No' : '— sin dato —')]
  ]);
  _estiloTabla(t, false);
}

function _anexoFichaHidraulica(body, ficha){
  body.appendPageBreak();
  _h1(body, '10. Anexo técnico B — Ficha hidráulica');
  var r = ficha.motorResultado;
  if(!r || r.error){
    body.appendParagraph(r && r.error ? ('Sin resultado válido: '+r.error) : 'No se calculó el caudal y las velocidades para este vaso en la PWA.')
        .setForegroundColor(C_TITULO).setItalic(true);
    return;
  }
  var t = body.appendTable([
    ['Curva de bomba usada', String(r.origen||'—')],
    ['Fecha del cálculo', r.ts ? Utilities.formatDate(new Date(r.ts), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '—'],
    ['Caudal en el punto de operación', r.caudal!=null ? r.caudal.toFixed(2)+' m³/h' : '— sin dato —'],
    ['Cabezal en el punto de operación', r.cabezal!=null ? r.cabezal.toFixed(2)+' m c.a.' : '— sin dato —'],
    ['Velocidad en succión', r.vSuccion!=null ? r.vSuccion.toFixed(2)+' m/s' : '— sin dato —'],
    ['Velocidad en descarga/retorno', r.vDescarga!=null ? r.vDescarga.toFixed(2)+' m/s' : '— sin dato —'],
    ['Velocidad de filtración', r.vFiltracion!=null ? r.vFiltracion.toFixed(1)+' m³/h/m²' : '— sin dato —'],
    ['Volumen estimado del vaso', r.volumen!=null ? r.volumen.toFixed(1)+' m³' : '— sin dato —'],
    ['Tiempo de recirculación', r.tiempoRecirc!=null ? r.tiempoRecirc.toFixed(2)+' h' : '— sin dato —'],
    ['Rotaciones estimadas por día', r.rotacionesDia!=null ? r.rotacionesDia.toFixed(1) : '— sin dato —']
  ]);
  _estiloTabla(t, false);
  body.appendParagraph(
    'Límites normativos de referencia: succión ≤ 1.8 m/s, descarga ≤ 2.4 m/s (Numeral 10.1); ' +
    'filtración 20-40 m³/h/m² (50 en uso restringido, Numeral 10.2); tiempo de recirculación ' +
    'según Tabla No. 1 del Anexo Técnico.'
  ).setFontSize(8).setForegroundColor(C_TITULO).setItalic(true);
}

function _anexoDimensionamiento(body, ficha){
  body.appendPageBreak();
  _h1(body, '11. Anexo técnico C — Dimensionamiento normativo');
  var r = ficha.aforoResultado;
  if(!r || r.error){
    body.appendParagraph(r && r.error ? ('Sin resultado válido: '+r.error) : 'No se calculó el aforo ni la dotación sanitaria para este vaso en la PWA.')
        .setForegroundColor(C_TITULO).setItalic(true);
    return;
  }
  var resumen = body.appendTable([
    ['Área del espejo de agua', r.area.toFixed(1)+' m²'],
    ['Factor de uso aplicado (Tabla 4)', r.factorUso+' m²/bañista — '+String(r.criterioProf||'')],
    ['Aforo máximo (Numeral 10.9)', r.aforo+' bañistas'],
    ['Fecha del cálculo', r.ts ? Utilities.formatDate(new Date(r.ts), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '—']
  ]);
  _estiloTabla(resumen, false);
  body.appendParagraph('');

  var etiquetas = {duchas:'Duchas', inodoroH:'Inodoros hombres', inodoroM:'Inodoros mujeres',
                    orinal:'Orinales', lavamanos:'Lavamanos', vestier:'Vestieres'};
  var datos = [['Elemento (Tabla 5)','Ratio normativo','Requerido','Encontrado','Estado']];
  Object.keys(etiquetas).forEach(function(k){
    var d = r.dotacion && r.dotacion[k];
    if(!d) return;
    var encontrado = (d.encontrado==null) ? '— sin dato —' : String(d.encontrado);
    var estadoTxt = (d.encontrado==null) ? 'Pendiente de conteo' : (d.encontrado < d.requerido ? 'Faltan '+(d.requerido-d.encontrado) : 'Cumple');
    datos.push([etiquetas[k], d.ratio, String(d.requerido), encontrado, estadoTxt]);
  });
  var t = body.appendTable(datos);
  _estiloTabla(t, true);
}

function _anexoMemoriaCalculo(body, ficha){
  body.appendPageBreak();
  _h1(body, '12. Anexo técnico D — Memoria de cálculo');
  body.appendParagraph(
    'Estimaciones de apoyo a la decisión del inspector: no reemplazan la medición directa con ' +
    'caudalímetro ni el conteo físico de la dotación sanitaria. Los supuestos declarados a ' +
    'continuación permiten reproducir o auditar cada resultado de los Anexos B y C.'
  ).setForegroundColor(C_TITULO).setFontSize(9).setItalic(true);

  _h2(body, 'D.1 Datos de la bomba');
  var filasBomba = [
    ['Potencia', ficha.bombaHP!=null ? ficha.bombaHP+' HP' : '—'],
    ['Frecuencia de operación', ficha.bombaHz!=null ? ficha.bombaHz+' Hz' : '60 Hz (valor por defecto)'],
    ['Origen de la curva', BOMBA_FAMILIA_LABEL[ficha.bombaFamilia] || '— sin dato —']
  ];
  if(ficha.bombaFamilia==='superflo_vs'){
    filasBomba.push(['Velocidad configurada', BOMBA_VELOCIDAD_VS_LABEL[ficha.bombaVelocidadVS] || '— sin dato —']);
  }
  if(ficha.bombaFamilia==='manual'){
    filasBomba.push(['Frecuencia de medición de la curva manual', ficha.curvaManualHz!=null ? ficha.curvaManualHz+' Hz' : '— sin dato — se asume igual a la de operación']);
  }
  if(ficha.bombaFamilia==='manual' && ficha.curvaManual && ficha.curvaManual.length){
    var puntos = ficha.curvaManual
      .filter(function(p){ return p && p.q!=null && p.h!=null; })
      .map(function(p){ return 'Q='+p.q+' m³/h, H='+p.h+' m'; })
      .join('  |  ');
    filasBomba.push(['Puntos de la curva manual', puntos || '—']);
  }
  var tBomba = body.appendTable(filasBomba);
  _estiloTabla(tBomba, false);

  _h2(body, 'D.2 Datos de tubería y filtro');
  var tTub = body.appendTable([
    ['Diámetro de succión', ficha.tuberiaSuccionDiam!=null ? ficha.tuberiaSuccionDiam+' pulg' : '—'],
    ['Diámetro de descarga/retorno', ficha.tuberiaDescargaDiam!=null ? ficha.tuberiaDescargaDiam+' pulg' : '—'],
    ['Longitud total de tubería', ficha.tuberiaLongitud!=null ? ficha.tuberiaLongitud+' m' : '—'],
    ['Número de accesorios', ficha.tuberiaAccesorios!=null ? String(ficha.tuberiaAccesorios) : '—'],
    ['Reparto succión/descarga', ficha.tuberiaPctSuccion!=null ? ficha.tuberiaPctSuccion+'% / '+(100-ficha.tuberiaPctSuccion)+'%' : '50% / 50% (por defecto)'],
    ['Desnivel succión-descarga', ficha.desnivelSuccionDescarga!=null ? ficha.desnivelSuccionDescarga+' m' : '—'],
    ['Tipo de filtro', FILTRO_TIPO_LABEL[ficha.filtroTipo] || '—'],
    ['Área filtrante', ficha.filtroArea!=null ? ficha.filtroArea+' m²' : '—'],
    ['Presión de manómetro del filtro', ficha.presionManometro!=null ? ficha.presionManometro+' PSI' : '—']
  ]);
  _estiloTabla(tTub, false);

  var sup = (ficha.motorResultado && ficha.motorResultado.supuestos) || null;
  if(sup){
    _h2(body, 'D.3 Supuestos del cálculo hidráulico');
    var tSup = body.appendTable([
      ['Coeficiente de Hazen-Williams (C)', String(sup.C)],
      ['Longitud equivalente de accesorios', String(sup.LeqPorAccesorio)],
      ['Carga estática asumida', String(sup.Hgeo)],
      ['Reparto de tubería succión/descarga', String(sup.repartoLongitud)],
      ['Conversión de presión', String(sup.conversionPsi)]
    ]);
    _estiloTabla(tSup, false);
  }

  if(ficha.aforoResultado && !ficha.aforoResultado.error){
    _h2(body, 'D.4 Supuesto del aforo y dotación');
    body.appendParagraph(
      'La fila de la Tabla No. 4 (factor de uso) se determina con la profundidad MÁXIMA del vaso ' +
      '— criterio conservador. La dotación sanitaria requerida (Tabla No. 5) se calcula sobre el ' +
      '100% del aforo para cada elemento; en campo puede repartirse por género según la ' +
      'composición real de bañistas.'
    ).setForegroundColor(C_TITULO).setFontSize(9);
  }
}

function _planAccion(body, filas){
  body.appendPageBreak();
  _h1(body, '5. Plan de acción y matriz de cierre');
  var datos = [['ID','Ítem','Riesgo','Acción','Responsable','Compromiso','Avance']];
  filas.forEach(function(f){
    var est = String(f[COL.estado-1]||'');
    if(est==='Cumple' || est==='No aplica' || est==='') return;
    datos.push([
      String(f[COL.id-1]),
      String(f[COL.item-1]).slice(0,70),
      String(f[COL.riesgo-1]||'—'),
      String(f[COL.accion-1]||'—').slice(0,70),
      String(f[COL.respCierre-1]||'—'),
      _fechaStr(f[COL.fechaCompromiso-1])||'—',
      (f[COL.avance-1]!==''? f[COL.avance-1]+'%':'—')
    ]);
  });
  if(datos.length===1){ body.appendParagraph('Sin acciones abiertas.'); return; }
  var t = body.appendTable(datos);
  _estiloTabla(t, true);
}

function _conclusion(body, m, sede, piscina){
  _h1(body, '6. Conclusión ejecutiva');
  var txt = 'El vaso ' + piscina + ' de la sede ' + sede + ' presenta un cumplimiento global del ' +
    m.pctCumplimiento + '% frente a los criterios de la Resolución 929 de 2026, con ' +
    m.noCumple + ' ítem(s) en estado No cumple y ' + (m.critico + m.alto) +
    ' hallazgo(s) clasificados en riesgo crítico o alto. ';
  if(m.vencidos>0) txt += 'Se registran ' + m.vencidos + ' hallazgo(s) con fecha compromiso vencida, ' +
    'lo que constituye la desviación de gestión más relevante del período. ';
  if(m.sinFecha>0) txt += m.sinFecha + ' hallazgo(s) permanecen sin fecha compromiso asignada. ';
  txt += 'La condición general se califica como: ' + _semaforo(m.pctCumplimiento) + '.';
  body.appendParagraph(txt).setForegroundColor(m.critico>0 ? C_ACENTO : C_TITULO);
}

function _responsabilidades(body, responsable){
  _h1(body, '7. Responsabilidades');
  var t = body.appendTable([
    ['Rol','Responsabilidad'],
    ['Supervisor de piscinas','Validación del cumplimiento del protocolo de inspección, verificación de la evidencia fotográfica, autorización de acciones de refuerzo y reporte a administración.'],
    ['Operario certificado','Ejecución de las acciones correctivas asignadas, registro en bitácora y reporte de novedades.'],
    ['Facility Management','Seguimiento periódico al cierre de hallazgos, gestión de recursos y validación del cierre técnico.'],
    ['Administración de sede','Aprobación presupuestal de las intervenciones y cumplimiento de los plazos del régimen transitorio (Art. 12, Res. 929/2026).']
  ]);
  _estiloTabla(t, true);

  body.appendParagraph('');
  var f = body.appendTable([
    ['Elaboró', String(responsable||'—')],
    ['Revisó',''],
    ['Aprobó','']
  ]);
  _estiloTabla(f, false);
}

/* ---------- Utilidades de formato ---------- */
function _h1(body, txt){
  body.appendParagraph(txt)
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setForegroundColor(C_ENCABEZADO);
}
function _h2(body, txt){
  body.appendParagraph(txt)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2)
      .setForegroundColor(C_TITULO);
}
function _estiloTabla(tabla, conEncabezado){
  tabla.setBorderColor(C_CELDA);
  for(var r=0;r<tabla.getNumRows();r++){
    var fila = tabla.getRow(r);
    for(var c=0;c<fila.getNumCells();c++){
      var celda = fila.getCell(c);
      celda.setBackgroundColor((conEncabezado && r===0) ? C_TITULO : (r%2===0 ? C_FONDO : C_CELDA));
      var p = celda.getChild(0).asParagraph();
      p.setFontSize(9).setForegroundColor((conEncabezado && r===0) ? '#FFFFFF' : C_TITULO);
      if(conEncabezado && r===0) p.setBold(true);
    }
  }
}

/* ============================================================================
   5. MENÚ EN LA HOJA — generar informes sin salir de Sheets
   ============================================================================ */
function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('Inspección piscinas')
    .addItem('Generar informe del vaso seleccionado', 'menuGenerarInforme')
    .addItem('Generar informes de toda la sede', 'menuGenerarSede')
    .addToUi();
}

function menuGenerarInforme(){
  var sh = _hojaDatos();
  var fila = SpreadsheetApp.getActiveSheet().getActiveRange().getRow();
  if(fila<2){ SpreadsheetApp.getUi().alert('Selecciona una fila de datos.'); return; }
  var v = sh.getRange(fila,1,1,TOTAL_COLS).getValues()[0];
  var res = generarInformeVaso(String(v[COL.sede-1]), String(v[COL.piscina-1]), _fechaStr(v[COL.fecha-1]));
  SpreadsheetApp.getUi().alert(res.ok ? 'Informe generado:\n'+res.pdfUrl : 'Error: '+res.error);
}

function menuGenerarSede(){
  var sh = _hojaDatos();
  var fila = SpreadsheetApp.getActiveSheet().getActiveRange().getRow();
  var v = sh.getRange(fila,1,1,TOTAL_COLS).getValues()[0];
  var sede = String(v[COL.sede-1]), fecha = _fechaStr(v[COL.fecha-1]);
  var last = sh.getLastRow();
  var datos = last>=2 ? sh.getRange(2,1,last-1,TOTAL_COLS).getValues() : [];
  var vasos = {};
  datos.forEach(function(f){
    if(String(f[COL.sede-1])===sede && _fechaStr(f[COL.fecha-1])===fecha){
      vasos[String(f[COL.piscina-1])] = true;
    }
  });
  var urls = [];
  Object.keys(vasos).forEach(function(p){
    var r = generarInformeVaso(sede, p, fecha);
    if(r.ok) urls.push(p+': '+r.pdfUrl);
  });
  SpreadsheetApp.getUi().alert('Informes generados ('+urls.length+'):\n\n'+urls.join('\n'));
}

function TEST_escribirFilaDePrueba(){
  var filaPrueba = [
    'CHK-TEST', 'PRUEBA MANUAL', 'Ítem de prueba de conexión',
    'Prueba', 'Alta', 'SEDE_TEST', 'PISCINA_TEST', '2026-07-30', 'Test Manual',
    'Cumple', 'Sin riesgo', '', '', '', '', '', '', '', '', '', '', ''
  ];
  var resultado = guardarFilas([filaPrueba]);
  Logger.log(JSON.stringify(resultado));

  // Verifica en qué hoja y spreadsheet realmente está escribiendo
  Logger.log('Spreadsheet: ' + _ss().getName() + ' | ID: ' + _ss().getId());
  Logger.log('Hoja de datos: ' + _hojaDatos().getName() + ' | gid: ' + _hojaDatos().getSheetId());
}

/** NUEVO: valida específicamente el fix del bug de upsert por fecha.
 *  Escribe la misma fila dos veces seguidas (misma clave ID+Sede+Piscina+
 *  Fecha) y confirma que la segunda vez actualiza en vez de duplicar. */
function TEST_upsertPorFecha(){
  var fila = [
    'CHK-TEST-UPSERT', 'PRUEBA UPSERT', 'Ítem de prueba de upsert por fecha',
    'Prueba', 'Alta', 'SEDE_TEST', 'PISCINA_TEST', '2026-08-07', 'Test Manual',
    'Cumple', 'Sin riesgo', '', '', '', '', '', '', '', '', '', '', 'primera'
  ];
  var r1 = guardarFilas([fila]);
  fila[21] = 'segunda'; // observaciones, misma clave
  var r2 = guardarFilas([fila]);
  Logger.log('1ra escritura (debe ser nueva:1): ' + JSON.stringify(r1));
  Logger.log('2da escritura (debe ser actualizada:1, nueva:0): ' + JSON.stringify(r2));
}

/* ============================================================================
   6. PRUEBA DE INFORME — escribe 4 filas de ejemplo (Crítico, Alto, Medio,
   Cumple) repartidas en capítulos reales del checklist y genera el PDF
   completo, para validar de una vez: columnas correctas (22, sin "norma"),
   orden de hallazgos por riesgo (fix del bug 0||5) y armado del documento.
   ============================================================================ */
function TEST_generarInformeDePrueba(){
  var sede = 'SEDE_TEST', piscina = 'PISCINA_TEST', fecha = '2026-08-03';

  var filas = [
    // [id, capitulo, item, enfoque, prioridad, sede, piscina, fecha, responsable,
    //  estado, riesgo, hallazgo, accion, respCierre, fechaCompromiso, fechaCierre,
    //  diasRest, avance, evidActual, evidCierre, linkCarpeta, observaciones]
    ['CHK-132','17. Calidad del agua y operación sanitaria','Registro de pH.','Calidad de agua','Alta',
      sede, piscina, fecha, 'Responsable Prueba', 'Cumple','Sin riesgo','','','','','','','','','','',''],

    ['CHK-233','17. Calidad del agua y operación sanitaria','Medidor de CO2','Calidad de agua','Alta',
      sede, piscina, fecha, 'Responsable Prueba', 'No cumple','Alto',
      'No hay medidor de CO2 instalado en el área cubierta.','Comprar e instalar medidor CO2.',
      'Coordinador Mantenimiento','2026-08-15','', '', '0', '', '', '', ''],

    ['CHK-245','20. Primeros auxilios','desfibrilador externo automatico (DEA)','Emergencias','Alta',
      sede, piscina, fecha, 'Responsable Prueba', 'No cumple','Critico',
      'No hay DEA disponible en el área de primeros auxilios.','Adquirir e instalar DEA de forma inmediata.',
      'Coordinador Mantenimiento','2026-08-10', '', '', '0', '', '', '', ''],

    ['CHK-234','17. Calidad del agua y operación sanitaria','Cuenta con higrotermometro','Calidad de agua','Alta',
      sede, piscina, fecha, 'Responsable Prueba', 'Pendiente','Medio',
      'Instrumento solicitado, en trámite de compra.','Dar seguimiento a la orden de compra.',
      'Coordinador Mantenimiento','2026-08-20', '', '', '20', '', '', '', '']
  ];

  var resultadoEscritura = guardarFilas(filas);
  Logger.log('Escritura: ' + JSON.stringify(resultadoEscritura));

  var informe = generarInformeVaso(sede, piscina, fecha);
  Logger.log('Informe: ' + JSON.stringify(informe));

  return informe;
}

/** Borra únicamente las filas de prueba (SEDE_TEST) para dejar la hoja limpia. */
function TEST_limpiarFilasDePrueba(){
  var sh = _hojaDatos();
  var last = sh.getLastRow();
  if(last<2){ Logger.log('Hoja sin datos.'); return; }
  var datos = sh.getRange(2,1,last-1,TOTAL_COLS).getValues();
  var filasABorrar = [];
  for(var i=0;i<datos.length;i++){
    if(String(datos[i][COL.sede-1])==='SEDE_TEST') filasABorrar.push(i+2); // fila real en la hoja (1-indexed + encabezado)
  }
  // Borra de abajo hacia arriba para no desfasar los índices al eliminar
  filasABorrar.sort(function(a,b){ return b-a; }).forEach(function(fila){ sh.deleteRow(fila); });
  Logger.log('Filas de prueba eliminadas: ' + filasABorrar.length);
}

/* ============================================================================
   7. PRUEBA DE INFORME CON FICHA TÉCNICA — escribe una ficha técnica
   completa (geometría, bomba, tubería, filtro, resultado de caudal/
   velocidades y resultado de aforo/dotación) más 4 filas de checklist
   ligadas a esos mismos datos, y genera el informe completo para validar
   de una vez: la hoja Ficha_Tecnica se crea sola, el upsert por
   Sede+Piscina+Fecha funciona, y los 4 anexos (secciones 9-12) se arman
   con los datos correctos.

   Los números de la ficha son los MISMOS que se validaron en la PWA en
   sesión (largo 25m × ancho 12m, profundidad 1.8/0.9m, bomba SuperFlo 1.5
   HP a 60Hz, tubería 4"/4"/15m/6 accesorios) — así el Word se puede
   comparar directamente contra lo que se vio en pantalla.
   ============================================================================ */
function TEST_generarInformeConFicha(){
  var sede = 'SEDE_TEST', piscina = 'PISCINA_TEST', fecha = '2026-08-09';

  var ficha = {
    largo:25, ancho:12, areaManual:null,
    profMax:1.8, profMin:0.9, profIntermedia:null,
    tipoUso:'colectivo', cubierta:'no',
    bombaHP:1.5, bombaHz:60, bombaFamilia:'superflo',
    tuberiaSuccionDiam:4, tuberiaDescargaDiam:4, tuberiaLongitud:15, tuberiaAccesorios:6,
    tuberiaPctSuccion:40, desnivelSuccionDescarga:0.6,
    filtroTipo:'arena', filtroArea:0.8, presionManometro:8,
    duchasEncontradas:2, inodoroHEncontrados:3, inodoroMEncontrados:4,
    orinalesEncontrados:2, lavamanosEncontrados:3, vestieresEncontrados:3,
    motorResultado:{
      origen:'SuperFlo 1.5 HP (referencia) a 60 Hz', caudal:34.27, cabezal:0.41,
      vSuccion:1.17, vDescarga:1.17, vFiltracion:42.8, volumen:405.0,
      tiempoRecirc:11.82, rotacionesDia:2.03,
      supuestos:{
        C:150, LeqPorAccesorio:'30·D por accesorio',
        Hgeo:'0.6 m c.a. (desnivel succión-descarga declarado en campo, sin manómetro de descarga junto a la piscina)',
        repartoLongitud:'40/60 succión-descarga (elegido en campo, ítem CHK-078)', conversionPsi:'1 PSI = 0.703 m c.a.'
      },
      ts: Date.now()
    },
    aforoResultado:{
      area:300.0, factorUso:4.0, criterioProf:'profundidad máxima > 1.5 m', aforo:75,
      dotacion:{
        duchas:   {requerido:3, encontrado:2, ratio:'1 cada 30 bañistas'},
        inodoroH: {requerido:3, encontrado:3, ratio:'1 cada 25 bañistas'},
        inodoroM: {requerido:4, encontrado:4, ratio:'1 cada 20 bañistas'},
        orinal:   {requerido:2, encontrado:2, ratio:'1 cada 40 bañistas'},
        lavamanos:{requerido:3, encontrado:3, ratio:'1 cada 25 bañistas'},
        vestier:  {requerido:3, encontrado:3, ratio:'1 cada 25 bañistas'}
      },
      ts: Date.now()
    }
  };

  var rFicha = guardarFicha({sede:sede, piscina:piscina, fecha:fecha, responsable:'Responsable Prueba', ficha:ficha});
  Logger.log('Ficha guardada: ' + JSON.stringify(rFicha));

  // Filas de checklist coherentes con la ficha: el hallazgo de CHK-078 y
  // CHK-204 referencia los mismos números que van a salir en los anexos.
  // CHK-152 (pH) va en "No cumple" A PROPÓSITO: como está en
  // FUERA_ALCANCE_929, debe seguir apareciendo en Hallazgos/Plan de acción
  // (transparencia) pero NO debe bajar el % de "Cumplimiento global (Res.
  // 929)" del tablero — eso es justamente lo que valida esta fila de prueba.
  var filas = [
    ['CHK-152','17. Calidad del agua y operación sanitaria','Registro de pH.','Calidad de agua','Alta',
      sede, piscina, fecha, 'Responsable Prueba', 'No cumple','Bajo',
      'Sin registro de pH en bitácora — se deja consignado por control operativo interno; no es un criterio exigido por la Res. 929 (ver nota del ítem).',
      'Retomar el registro diario de pH como buena práctica operativa.',
      'Operario de mantenimiento','2026-08-18','', '', '0', '', '', '', ''],

    ['CHK-233','17. Calidad del agua y operación sanitaria','Medidor de CO2','Calidad de agua','Alta',
      sede, piscina, fecha, 'Responsable Prueba', 'No cumple','Alto',
      'No hay medidor de CO2 instalado en el área cubierta.','Comprar e instalar medidor CO2.',
      'Coordinador Mantenimiento','2026-08-15','', '', '0', '', '', '', ''],

    ['CHK-078','8. Sistema de recirculación','Caudal real verificado frente al diseño.','Operación / Calidad de agua','Alta',
      sede, piscina, fecha, 'Responsable Prueba', 'No cumple','Alto',
      'Velocidad de filtración (42.8 m³/h/m²) y tiempo de recirculación (11.82 h) fuera de rango según el cálculo del motor hidráulico — ver Anexo B.',
      'Evaluar aumento de área filtrante o bomba de mayor caudal.',
      'Coordinador Mantenimiento','2026-08-20','', '', '0', '', '', '', ''],

    ['CHK-204','24. Aforo y control de ingreso','Aforo máximo determinado por estanque.','Operación / Seguridad','Media',
      sede, piscina, fecha, 'Responsable Prueba', 'No cumple','Medio',
      'Faltan duchas frente a la dotación de la Tabla No. 5 (2 encontradas / 3 requeridas para 75 bañistas) — ver Anexo C.',
      'Instalar 1 ducha adicional en el área de uso exclusivo del bañista.',
      'Coordinador Mantenimiento','2026-08-25','', '', '0', '', '', '', '']
  ];
  var resultadoEscritura = guardarFilas(filas);
  Logger.log('Filas de checklist: ' + JSON.stringify(resultadoEscritura));

  var informe = generarInformeVaso(sede, piscina, fecha);
  Logger.log('Informe: ' + JSON.stringify(informe));
  return informe;
}

/** Borra la ficha técnica de prueba (SEDE_TEST) de la hoja Ficha_Tecnica.
 *  Complementa a TEST_limpiarFilasDePrueba, que solo limpia el checklist. */
function TEST_limpiarFichaDePrueba(){
  var sh = _hojaFicha();
  var last = sh.getLastRow();
  if(last<2){ Logger.log('Hoja de fichas sin datos.'); return; }
  var datos = sh.getRange(2,1,last-1,TOTAL_COLS_FICHA).getValues();
  var filasABorrar = [];
  for(var i=0;i<datos.length;i++){
    if(String(datos[i][COL_FICHA.sede-1])==='SEDE_TEST') filasABorrar.push(i+2);
  }
  filasABorrar.sort(function(a,b){ return b-a; }).forEach(function(fila){ sh.deleteRow(fila); });
  Logger.log('Filas de ficha de prueba eliminadas: ' + filasABorrar.length);
}
