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
/* Encabezado real de la hoja, en el mismo orden que COL — usado por
   _diagnosticarHoja() para detectar si alguien cambió/movió columnas. */
var COLUMNAS_ESPERADAS = [
  'ID','Capítulo / componente','Ítem de verificación','Enfoque técnico','Prioridad sugerida',
  'Sede','Piscina / estanque','Fecha inspección','Responsable inspección','Estado','Nivel de riesgo',
  'Hallazgo / condición observada','Acción correctiva / preventiva','Responsable cierre',
  'Fecha compromiso','Fecha cierre','Días restantes / vencido','% avance',
  'Evidencia fotográfica - estado actual','Evidencia fotográfica - cierre','Link carpeta / evidencia','Observaciones'
];

/* Paleta corporativa — alineada al diseño "Industrial Integrity" de la PWA
   (index.html/dashboard.html) y al mockup de Stitch para el informe: slate
   oscuro para texto/encabezados, azul para "conforme", coral para
   hallazgos, ámbar para en proceso, gris para pendiente/no aplica. Los
   nombres de las constantes se mantienen (C_TITULO, C_ACENTO, etc.) porque
   los usan ~20 funciones de anexos más abajo — solo cambian los valores. */
var C_ENCABEZADO='#0F172A',      // slate-900 — títulos H1/portada
    C_TITULO='#475569',          // slate-600 — texto secundario/H2
    C_ACENTO='#DC2626',          // rojo-600 — texto de alerta (crítico/no cumple)
    C_CELDA='#E2E8F0',           // slate-200 — bordes y filas alternas
    C_FONDO='#F8FAFC',           // slate-50 — filas pares
    C_OK='#2563EB';              // azul-600 — texto "conforme" (coherente con la app)

/* Colores de relleno para gráficas e insignias (más saturados que los de
   texto de arriba, pensados para áreas de color, no para letras). */
var C_FILL_CUMPLE='#3B82F6', C_FILL_NOCUMPLE='#F87171', C_FILL_PROCESO='#F59E0B',
    C_FILL_PENDIENTE='#94A3B8', C_FILL_NOAPLICA='#CBD5E1',
    C_HEADER_TABLA='#F1F5F9';    // slate-100 — fondo de encabezado de tabla
var C_AMBAR='#B45309';           // ámbar-700 — riesgo medio / lectura de advertencia

/* Tipografía y geometría de página. El mockup usa Inter + JetBrains Mono;
   Google Docs no trae ninguna de las dos por defecto, así que se mapean a
   las equivalentes disponibles en Docs (Arial para UI, Roboto Mono para
   valores numéricos). A4 son 595 pt de ancho: con márgenes de 45 pt quedan
   505 pt útiles, y ese es el ancho contra el que se dimensionan TODAS las
   tablas y columnas del informe. */
var FUENTE='Arial', FUENTE_MONO='Roboto Mono';
var MARGEN_PAG=45, ANCHO_UTIL=595-(MARGEN_PAG*2);   // 505 pt

/* Estilos por defecto del documento: se aplican una sola vez sobre el body
   y los niveles de encabezado, así cada párrafo/tabla los hereda en vez de
   tener que re-estilizar cada llamada. Equivale al bloque de fontSize/
   fontFamily del tailwind.config del mockup. */
function _configurarEstilosDoc(body){
  body.setMarginTop(MARGEN_PAG).setMarginBottom(MARGEN_PAG)
      .setMarginLeft(MARGEN_PAG).setMarginRight(MARGEN_PAG);

  var A = DocumentApp.Attribute;
  var normal = {};
  normal[A.FONT_FAMILY]=FUENTE; normal[A.FONT_SIZE]=9.5;
  normal[A.FOREGROUND_COLOR]=C_TITULO; normal[A.LINE_SPACING]=1.15;
  body.setAttributes(normal);

  var h1 = {};
  h1[A.FONT_FAMILY]=FUENTE; h1[A.FONT_SIZE]=18; h1[A.BOLD]=true;
  h1[A.FOREGROUND_COLOR]=C_ENCABEZADO;
  h1[A.SPACING_BEFORE]=20; h1[A.SPACING_AFTER]=8;
  body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING1, h1);

  var h2 = {};
  h2[A.FONT_FAMILY]=FUENTE; h2[A.FONT_SIZE]=12.5; h2[A.BOLD]=true;
  h2[A.FOREGROUND_COLOR]=C_ENCABEZADO;
  h2[A.SPACING_BEFORE]=14; h2[A.SPACING_AFTER]=6;
  body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING2, h2);

  var h3 = {};
  h3[A.FONT_FAMILY]=FUENTE; h3[A.FONT_SIZE]=11.5; h3[A.BOLD]=true;
  h3[A.FOREGROUND_COLOR]=C_ENCABEZADO;
  h3[A.SPACING_BEFORE]=14; h3[A.SPACING_AFTER]=2;
  body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING3, h3);
}

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
var PROP_KEY_HOJA_ANCLADA = 'SPREADSHEET_ID_ANCLADA';

/* El ancla (PropertiesService, a nivel de proyecto de Apps Script — la ve
   CUALQUIER dispositivo que llame al webhook, no solo el navegador que la
   fijó) tiene prioridad sobre el spreadsheet contenedor y sobre el ID fijo
   de respaldo. Es la forma de cambiar "la hoja de siempre" de toda la app
   —checklist, ficha técnica, dashboard, informes— sin tocar código ni
   volver a publicar el script. Si la hoja anclada dejó de ser accesible
   (se borró, se revocó el acceso), cae al ID fijo en vez de tronar. */
function _ss(){
  var anclada = PropertiesService.getScriptProperties().getProperty(PROP_KEY_HOJA_ANCLADA);
  if(anclada){
    try{ return SpreadsheetApp.openById(anclada); }
    catch(e){ /* la hoja anclada ya no es accesible — sigue con el respaldo de abajo */ }
  }
  try{
    var activa = SpreadsheetApp.getActiveSpreadsheet();
    if(activa) return activa;
  }catch(e){ /* script no vinculado, cae al ID fijo */ }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}
/* Variante de _ss() que acepta un ID de hoja distinto al de siempre — la
   usan el dashboard y el diagnóstico cuando el inspector pega el ID/URL de
   OTRA hoja (p.ej. si cambiaron de spreadsheet este año). Solo funciona si
   la cuenta de Google que ejecuta el script tiene acceso a esa hoja; si no,
   lanza un error legible en vez del genérico de Apps Script. */
function _ssPara(spreadsheetId){
  if(!spreadsheetId) return _ss();
  var id = _extraerIdHoja(spreadsheetId);
  try{
    return SpreadsheetApp.openById(id);
  }catch(err){
    throw new Error('No se pudo abrir la hoja con ese ID/URL. Verifica que sea correcto y que la cuenta del script tenga acceso (compártela con esa cuenta si es de otra persona).');
  }
}
/* Acepta tanto un ID pelado como una URL completa de Sheets pegada por el
   usuario (.../spreadsheets/d/<ID>/edit#gid=...). */
function _extraerIdHoja(valor){
  var s = String(valor||'').trim();
  var m = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : s;
}
function _hojaDatosDe(ss){
  var hojas = ss.getSheets();
  for(var i=0;i<hojas.length;i++){
    if(hojas[i].getSheetId() === GID_DATOS) return hojas[i];
  }
  return ss.getSheets()[0]; // respaldo si el gid cambiara o no existe en la hoja alternativa
}
function _hojaDatos(){
  return _hojaDatosDe(_ss());
}

/* ---------- Diagnóstico de la hoja conectada ----------
   Responde SIEMPRE con {ok, errores, avisos, ...datos}, nunca lanza — así
   el dashboard puede mostrar el motivo exacto por el que algo no carga en
   vez de un fetch fallido genérico. Se usa tanto para la hoja de siempre
   como para una alterna que el inspector quiera probar antes de adoptarla
   (accion:'diagnostico', body.spreadsheetId opcional). */
function _diagnosticarHoja(spreadsheetId){
  var r = {ok:true, errores:[], avisos:[]};
  var ss;
  try{ ss = _ssPara(spreadsheetId); }
  catch(err){ r.ok=false; r.errores.push(String(err.message||err)); return r; }

  r.spreadsheetNombre = ss.getName();
  r.spreadsheetId = ss.getId();
  r.spreadsheetUrl = ss.getUrl();

  var sh;
  try{ sh = _hojaDatosDe(ss); }
  catch(err){ r.ok=false; r.errores.push('No se pudo acceder a ninguna pestaña de la hoja: '+(err.message||err)); return r; }

  r.hojaNombre = sh.getName();
  r.hojaGid = sh.getSheetId();
  r.gidEsperado = GID_DATOS;
  if(sh.getSheetId() !== GID_DATOS){
    r.avisos.push('La pestaña usada no coincide con el GID configurado en el script (GID_DATOS='+GID_DATOS+') — se cayó a la primera pestaña como respaldo. Si esto es intencional (moviste/renombraste la pestaña), está bien; si no, revisa que la pestaña "'+sh.getName()+'" sea realmente la de datos.');
  }

  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  r.filas = Math.max(0, lastRow-1);
  r.columnas = lastCol;

  if(lastRow < 1){
    r.ok=false; r.errores.push('La pestaña está completamente vacía — ni siquiera tiene fila de encabezado.');
    return r;
  }
  if(lastCol !== TOTAL_COLS){
    r.avisos.push('La pestaña tiene '+lastCol+' columnas; la estructura esperada tiene '+TOTAL_COLS+'.');
  }

  var anchoLeer = Math.max(lastCol, COLUMNAS_ESPERADAS.length);
  var encabezado = sh.getRange(1,1,1,anchoLeer).getValues()[0].map(function(v){ return String(v||'').trim(); });
  var difs = [];
  COLUMNAS_ESPERADAS.forEach(function(esp, i){
    var real = encabezado[i]||'';
    if(real !== esp) difs.push({col:i+1, esperado:esp, encontrado: real || '(vacío)'});
  });
  if(difs.length){
    r.ok=false;
    r.errores.push('El encabezado no coincide con la estructura de 22 columnas en '+difs.length+' columna(s) — por eso el dashboard puede leer datos en el campo equivocado o simplemente no cargar.');
    r.diferenciasEncabezado = difs;
  }

  if(r.filas === 0){
    r.avisos.push('El encabezado está correcto pero todavía no hay filas de datos — es normal si aún no se ha sincronizado ninguna inspección desde la app.');
  } else if(!difs.length){
    var datos = sh.getRange(2,1,Math.min(lastRow-1, 5000),TOTAL_COLS).getValues();
    var sedesUnicas = {}, fechaMin=null, fechaMax=null, filasSinFecha=0, filasSinSede=0;
    datos.forEach(function(f){
      var sede = String(f[COL.sede-1]||'').trim();
      if(sede) sedesUnicas[sede]=1; else filasSinSede++;
      var fs = _fechaStr(f[COL.fecha-1]);
      if(fs){ if(!fechaMin||fs<fechaMin) fechaMin=fs; if(!fechaMax||fs>fechaMax) fechaMax=fs; }
      else filasSinFecha++;
    });
    r.sedesDetectadas = Object.keys(sedesUnicas).sort();
    r.rangoFechas = {desde:fechaMin, hasta:fechaMax};
    r.filasSinFecha = filasSinFecha;
    r.filasSinSede = filasSinSede;
    if(filasSinFecha>0) r.avisos.push(filasSinFecha+' fila(s) sin fecha de inspección válida.');
    if(filasSinSede>0) r.avisos.push(filasSinSede+' fila(s) sin sede.');
  }

  return r;
}

/* ---------- Anclar / desanclar la hoja de siempre ----------
   Guarda el ID en PropertiesService del proyecto de Apps Script — no en el
   navegador — así que aplica para TODOS los dispositivos que usen este
   webhook desde el momento en que se ancla, incluida la sincronización del
   checklist y la generación de informes, no solo el dashboard. Nunca ancla
   una hoja con estructura inválida: corre el diagnóstico primero y si no
   pasa, no guarda nada. */
function anclarHoja(spreadsheetId){
  if(!spreadsheetId) return {ok:false, error:'Falta el ID o URL de la hoja a anclar.'};
  var diag = _diagnosticarHoja(spreadsheetId);
  if(!diag.ok){
    return {ok:false, error:'No se ancló: la estructura de esa hoja no coincide con las 22 columnas esperadas (o hubo un error de acceso). Revisa el diagnóstico antes de anclar.', diagnostico:diag};
  }
  var id = _extraerIdHoja(spreadsheetId);
  PropertiesService.getScriptProperties().setProperty(PROP_KEY_HOJA_ANCLADA, id);
  return {ok:true, spreadsheetId:id, spreadsheetNombre:diag.spreadsheetNombre, hojaNombre:diag.hojaNombre};
}
function desanclarHoja(){
  PropertiesService.getScriptProperties().deleteProperty(PROP_KEY_HOJA_ANCLADA);
  return {ok:true};
}
/* Consulta de solo lectura: qué hoja está anclada ahora mismo (si alguna),
   para que el dashboard lo muestre aunque el inspector no haya tocado nada
   en este navegador — el ancla es del proyecto, no de la sesión. */
function obtenerHojaAnclada(){
  var id = PropertiesService.getScriptProperties().getProperty(PROP_KEY_HOJA_ANCLADA);
  if(!id) return {ok:true, anclada:false};
  try{
    var ss = SpreadsheetApp.openById(id);
    return {ok:true, anclada:true, spreadsheetId:id, spreadsheetNombre:ss.getName()};
  }catch(err){
    return {ok:true, anclada:true, spreadsheetId:id, spreadsheetNombre:null,
      aviso:'La hoja anclada ('+id+') ya no es accesible para la cuenta del script — la app está cayendo al respaldo mientras tanto.'};
  }
}

/* ============================================================================
   1. ROUTER doPost
   ============================================================================ */
function doPost(e){
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch(err){ body = {}; }

  if (body.accion === 'foto')       return _json(guardarFoto(body));
  if (body.accion === 'ficha')      return _json(guardarFicha(body));
  if (body.accion === 'informe')    return _json(generarInformeVaso(body.sede, body.piscina, body.fecha));
  if (body.accion === 'dashboard')    return _json(obtenerDashboard(body.spreadsheetId));
  if (body.accion === 'diagnostico')  return _json(_diagnosticarHoja(body.spreadsheetId));
  if (body.accion === 'anclarHoja')   return _json(anclarHoja(body.spreadsheetId));
  if (body.accion === 'desanclarHoja') return _json(desanclarHoja());
  if (body.accion === 'hojaAnclada')  return _json(obtenerHojaAnclada());
  if (body.rows)                      return _json(guardarFilas(body.rows));

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
  _configurarEstilosDoc(body);

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
function obtenerDashboard(spreadsheetId){
 try{
  var ss = _ssPara(spreadsheetId);
  var sh = _hojaDatosDe(ss);
  var last = sh.getLastRow();
  var todo = last>=2 ? sh.getRange(2,1,last-1,TOTAL_COLS).getValues() : [];
  if(!todo.length) return {ok:true, vacio:true, spreadsheetNombre:ss.getName(), hojaNombre:sh.getName()};

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
    spreadsheetNombre: ss.getName(),
    hojaNombre: sh.getName(),
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
 }catch(err){
  return {ok:false, error:String(err.message||err)};
 }
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
  // Encabezado: título+subtítulo a la izquierda y ficha de metadatos
  // compacta a la derecha, en la misma fila — igual que el header del
  // mockup de Stitch. DocumentApp no tiene flex/grid, así que el layout de
  // columnas se logra con una tabla de una fila sin bordes visibles
  // (_filaColumnas), y cada "tarjeta" es una tabla anidada de una sola
  // celda con borde propio (_tarjetaEnCelda).
  // 62% título / 38% ficha, sobre los 505 pt útiles.
  var header = _filaColumnas(body, [313, 192]);
  var colTitulo = header[0], colMeta = header[1];

  var pTitulo = colTitulo.getChild(0).asParagraph();
  pTitulo.setText('INFORME DE INSPECCIÓN NORMATIVA');
  pTitulo.setFontSize(19).setBold(true).setForegroundColor(C_ENCABEZADO)
         .setSpacingBefore(0).setSpacingAfter(2);
  colTitulo.appendParagraph('Resolución 929 de 2026')
    .setFontSize(10.5).setBold(false).setForegroundColor(C_TITULO);

  var metaCard = _tarjetaEnCelda(colMeta, C_FONDO);
  var metaTabla = metaCard.appendTable([
    ['SEDE', String(sede)],
    ['VASO', String(piscina)],
    ['FECHA', String(fecha)],
    ['RESPONSABLE', String(responsable||'—')]
  ]);
  metaTabla.setBorderWidth(0);
  for(var r=0;r<metaTabla.getNumRows();r++){
    var fila = metaTabla.getRow(r);
    var etq = fila.getCell(0), val = fila.getCell(1);
    etq.setWidth(68).setBackgroundColor(C_FONDO);
    val.setWidth(94).setBackgroundColor(C_FONDO);
    etq.setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(0).setPaddingRight(4);
    val.setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(0).setPaddingRight(0);
    etq.getChild(0).asParagraph().setFontSize(7.5).setBold(true).setForegroundColor(C_TITULO);
    val.getChild(0).asParagraph().setFontSize(9).setBold(true).setForegroundColor(C_ENCABEZADO)
       .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  }

  body.appendParagraph('');
  _h2(body, 'Resumen ejecutivo');

  // Dos tarjetas apiladas a ancho completo: primero el cumplimiento global
  // (dona + cifra), debajo la distribución de hallazgos (torta + tabla).
  // La dona va a la izquierda de la cifra dentro de su propia tarjeta, para
  // que el bloque no ocupe media página de alto.
  var donaCard = _tarjetaEnCelda(_filaColumnas(body, [ANCHO_UTIL])[0]);
  var pDonaTitulo = donaCard.getChild(0).asParagraph();
  pDonaTitulo.setText('CUMPLIMIENTO GLOBAL');
  pDonaTitulo.setFontSize(8).setBold(true).setForegroundColor(C_TITULO);

  var interior = _filaColumnas(donaCard, [150, 325]);
  _insertarImagenCentrada(interior[0], _chartDonutCumplimiento(m.pctCumplimiento), 130);
  var pPct = interior[1].getChild(0).asParagraph();
  pPct.setText(m.pctCumplimiento+'%');
  pPct.setBold(true).setFontSize(34).setForegroundColor(C_ENCABEZADO)
      .setSpacingBefore(18).setSpacingAfter(0);
  interior[1].appendParagraph(_semaforo(m.pctCumplimiento))
    .setFontSize(11).setBold(true).setForegroundColor(_colorPct(m.pctCumplimiento))
    .setSpacingBefore(0);
  interior[1].appendParagraph(m.cumpleEnAlcance+' de '+m.baseEnAlcance+' ítems conformes dentro del alcance de la Res. 929')
    .setFontSize(8.5).setForegroundColor(C_TITULO).setSpacingBefore(4);

  body.appendParagraph('').setFontSize(6);

  var distCard = _tarjetaEnCelda(_filaColumnas(body, [ANCHO_UTIL])[0]);
  var pDistTitulo = distCard.getChild(0).asParagraph();
  pDistTitulo.setText('DISTRIBUCIÓN DE HALLAZGOS');
  pDistTitulo.setFontSize(8).setBold(true).setForegroundColor(C_TITULO);
  _insertarImagenCentrada(distCard, _chartDistribucionEstados(m), 300);

  var distTabla = distCard.appendTable([
    ['Estado','Cantidad','%'],
    ['Cumple', String(m.cumple), (m.total>0?Math.round(100*m.cumple/m.total):0)+'%'],
    ['No cumple', String(m.noCumple), (m.total>0?Math.round(100*m.noCumple/m.total):0)+'%'],
    ['En proceso', String(m.enProceso), (m.total>0?Math.round(100*m.enProceso/m.total):0)+'%'],
    ['Pendiente', String(m.pendiente), (m.total>0?Math.round(100*m.pendiente/m.total):0)+'%'],
    ['No aplica', String(m.noAplica), '—']
  ]);
  _estiloTabla(distTabla, true, [230, 125, 120]);
  var coloresFila = [null, C_OK, C_ACENTO, C_AMBAR, C_TITULO, C_TITULO];
  for(var rr=1; rr<distTabla.getNumRows(); rr++){
    if(!coloresFila[rr]) continue;
    distTabla.getRow(rr).getCell(0).getChild(0).asParagraph().setForegroundColor(coloresFila[rr]).setBold(true);
    distTabla.getRow(rr).getCell(2).getChild(0).asParagraph().setForegroundColor(coloresFila[rr]);
  }

  body.appendPageBreak();
}

/* ---------- Layout: columnas y tarjetas sin CSS ----------
   DocumentApp no tiene flexbox/grid — el truco estándar para poner
   elementos lado a lado en un Google Doc es una tabla de layout con borde
   invisible (_filaColumnas) y, dentro de cada celda, una tabla anidada de
   1x1 con borde propio para simular una "tarjeta" (_tarjetaEnCelda). */
function _filaColumnas(body, widths){
  var vacio = widths.map(function(){ return ''; });
  var t = body.appendTable([vacio]);
  t.setBorderWidth(0).setBorderColor('#FFFFFF');
  var celdas = [];
  for(var i=0;i<widths.length;i++){
    var celda = t.getCell(0,i);
    celda.setWidth(widths[i]);
    celda.setPaddingTop(0).setPaddingBottom(0);
    celda.setPaddingLeft(i===0?0:10);
    celda.setPaddingRight(0);
    celdas.push(celda);
  }
  return celdas;
}
function _tarjetaEnCelda(celdaPadre, fondo){
  var t = celdaPadre.appendTable([['']]);
  t.setBorderColor(C_CELDA).setBorderWidth(1);
  var celda = t.getCell(0,0);
  celda.setBackgroundColor(fondo || '#FFFFFF');
  celda.setPaddingTop(10).setPaddingBottom(10).setPaddingLeft(10).setPaddingRight(10);
  return celda;
}

/* Color semántico de un porcentaje de cumplimiento — mismo umbral que
   _semaforo(), centralizado para no repetir el ternario en cada sección. */
function _colorPct(pct){
  return pct>=85 ? C_OK : (pct>=70 ? C_AMBAR : C_ACENTO);
}
/* Color semántico de un nivel de riesgo (acepta "Critico" y "Crítico"). */
function _colorRiesgo(riesgo){
  var r = String(riesgo||'');
  if(r==='Critico'||r==='Crítico'||r==='Alto') return C_ACENTO;
  if(r==='Medio') return C_AMBAR;
  return C_TITULO;
}

/* ---------- Tarjeta de datos etiqueta → valor ----------
   Es el componente que más se repite en el diseño: una "card" con borde,
   columna izquierda gris con la etiqueta en minúscula-negrita y columna
   derecha blanca con el valor destacado. Sustituye a _estiloTabla(t,false)
   en Hallazgos, Anexos A-D y Registro fotográfico.
   opts: {ancho, pctEtiqueta, mono} — `mono` pone el valor en Roboto Mono,
   como el font-mono que usa el mockup para magnitudes numéricas. */
function _tarjetaDatos(contenedor, filas, opts){
  opts = opts || {};
  var ancho = opts.ancho || ANCHO_UTIL;
  var anchoEtq = Math.round(ancho * (opts.pctEtiqueta || 0.34));
  var t = contenedor.appendTable(filas.map(function(f){
    return [String(f[0]), String(f[1])];
  }));
  t.setBorderColor(C_CELDA).setBorderWidth(0.75);
  for(var r=0;r<t.getNumRows();r++){
    var fila = t.getRow(r);
    var etq = fila.getCell(0), val = fila.getCell(1);
    etq.setWidth(anchoEtq).setBackgroundColor(C_HEADER_TABLA);
    val.setWidth(ancho-anchoEtq).setBackgroundColor('#FFFFFF');
    etq.setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(9).setPaddingRight(9);
    val.setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(9).setPaddingRight(9);
    etq.getChild(0).asParagraph().setFontSize(8.5).setBold(true).setForegroundColor(C_TITULO);
    var pv = val.getChild(0).asParagraph();
    pv.setFontSize(9.5).setBold(true).setForegroundColor(C_ENCABEZADO);
    if(opts.mono) pv.setFontFamily(FUENTE_MONO);
  }
  return t;
}

/* Pinta el valor de una fila concreta de una tarjeta creada con
   _tarjetaDatos (índice 0-based), para resaltar estado/riesgo en color. */
function _colorValorFila(tabla, indiceFila, color){
  tabla.getRow(indiceFila).getCell(1).getChild(0).asParagraph().setForegroundColor(color);
}

/* Tarjeta de datos con banda de encabezado propia: el ID del ítem (y su
   insignia de riesgo debajo) en la columna izquierda, y el nombre del ítem
   en la derecha. Al vivir el título DENTRO de la tabla, nunca queda
   separado de sus datos por un salto de página. */
function _tarjetaHallazgo(contenedor, itemId, nombreItem, riesgo, filas){
  var color = _colorRiesgo(riesgo);
  var todas = [[itemId, nombreItem]].concat(filas);
  var t = _tarjetaDatos(contenedor, todas);

  var cabId = t.getRow(0).getCell(0), cabNombre = t.getRow(0).getCell(1);
  cabId.setBackgroundColor(C_FONDO);
  cabNombre.setBackgroundColor(C_FONDO);

  var pId = cabId.getChild(0).asParagraph();
  pId.setFontSize(9).setBold(true).setFontFamily(FUENTE_MONO).setForegroundColor(color)
     .setSpacingAfter(0);
  cabId.appendParagraph('● RIESGO ' + String(riesgo).toUpperCase())
       .setFontSize(7.5).setBold(true).setForegroundColor(color)
       .setSpacingBefore(2).setSpacingAfter(0);

  cabNombre.getChild(0).asParagraph()
    .setFontSize(11).setBold(true).setForegroundColor(color);
  return t;
}

/* Rótulo de sección corto en mayúscula (el "label-bold uppercase" del
   mockup): más discreto que un HEADING, para encabezar galerías y bloques
   dentro de una sección ya titulada. */
function _rotulo(contenedor, txt){
  return contenedor.appendParagraph(txt.toUpperCase())
    .setFontSize(8).setBold(true).setForegroundColor(C_TITULO)
    .setSpacingBefore(10).setSpacingAfter(3);
}

/* Nota al pie de sección: texto pequeño, gris y en cursiva. */
function _nota(contenedor, txt){
  return contenedor.appendParagraph(txt)
    .setFontSize(8.5).setItalic(true).setForegroundColor(C_TITULO)
    .setSpacingAfter(8);
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
  // La distribución de hallazgos (gráfica + tabla) ya se muestra en el
  // "Resumen ejecutivo" de la portada (_portada) — esta sección cubre los
  // indicadores de gestión que no caben en esa tarjeta compacta, sin
  // repetir la misma gráfica dos veces en el informe.
  _h1(body, '2. Tablero ejecutivo — indicadores de gestión');
  var t = body.appendTable([
    ['Indicador','Valor','Lectura'],
    ['Ítems evaluados', String(m.total), '—'],
    ['Fuera del alcance de la Res. 929', String(m.fueraAlcance), 'Excluido del cálculo — dominio de la Res. 234'],
    ['Riesgo crítico', String(m.critico), m.critico>0?'ATENCIÓN INMEDIATA':'Sin críticos'],
    ['Riesgo alto', String(m.alto), m.alto>0?'Prioritario':'—'],
    ['Hallazgos vencidos', String(m.vencidos), m.vencidos>0?'Fuera de plazo':'Al día'],
    ['Hallazgos sin fecha compromiso', String(m.sinFecha), m.sinFecha>0?'Asignar fecha':'Completo'],
    ['Avance promedio de cierre', m.avanceProm+'%', '—']
  ]);
  _estiloTabla(t, true, [220, 70, 215]);
  // Realza en rojo las filas con lectura de alerta activa.
  for(var i=1;i<t.getNumRows();i++){
    var lectura = t.getRow(i).getCell(2).getText();
    if(/ATENCIÓN|Prioritario|Fuera de plazo|Asignar fecha/.test(lectura)){
      t.getRow(i).getCell(2).getChild(0).asParagraph().setForegroundColor(C_ACENTO).setBold(true);
    }
  }
}
function _semaforo(pct){
  if(pct>=95) return 'Conforme';
  if(pct>=85) return 'Aceptable con observaciones';
  if(pct>=70) return 'Deficiente';
  return 'Crítico';
}

function _tablaCapitulos(body, m){
  _h1(body, '3. Cumplimiento por capítulo normativo');

  _insertarImagenCentrada(body, _chartCapitulos(m), ANCHO_UTIL);
  body.appendParagraph('').setFontSize(6);

  var datos = [['Capítulo','Ítems','Cumple','No cumple','Fuera de alcance','% cumplimiento (Res. 929)']];
  var pcts = [null]; // paralelo a `datos`, guarda el % de cada fila para colorear después
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
    pcts.push(pct);
  });
  var t = body.appendTable(datos);
  _estiloTabla(t, true, [175, 45, 50, 60, 65, 110]);
  for(var r=1;r<t.getNumRows();r++){
    var celda = t.getRow(r).getCell(5).getChild(0).asParagraph();
    celda.setForegroundColor(_colorPct(pcts[r])).setBold(true);
  }
}

function _hallazgosConFotos(body, filas, fotos){
  _h1(body, '4. Hallazgos');
  _nota(body,
    'La evidencia fotográfica de cada ítem, incluidos los de esta sección, se presenta ' +
    'de forma consolidada en el Anexo técnico — Registro fotográfico (sección 8).');
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
    var riesgo = String(f[COL.riesgo-1]||'Sin clasificar');

    // Cada hallazgo es UNA sola tabla, con el ID y el nombre del ítem en su
    // primera fila a modo de encabezado. Antes eran un título suelto + dos
    // tablas: al saltar de página Docs dejaba el título huérfano o mandaba
    // la segunda tabla sola a la página siguiente con medio folio en blanco.
    // Con una tabla única el corte, si ocurre, es entre filas y se lee
    // continuo. (DocumentApp no expone "mantener junto con lo siguiente".)
    var t = _tarjetaHallazgo(body, String(f[COL.id-1]), String(f[COL.item-1]), riesgo, [
      ['Capítulo', String(f[COL.capitulo-1]||'—')],
      ['Estado', String(f[COL.estado-1]||'—')],
      ['Nivel de riesgo', riesgo],
      ['Hallazgo observado', String(f[COL.hallazgo-1]||'—')],
      ['Acción correctiva', String(f[COL.accion-1]||'—')],
      ['Responsable de cierre', String(f[COL.respCierre-1]||'—')],
      ['Fecha compromiso', _fechaStr(f[COL.fechaCompromiso-1])||'—'],
      ['% avance', String(f[COL.avance-1]!==''?f[COL.avance-1]+'%':'—')]
    ]);
    _colorValorFila(t, 2, C_ACENTO);              // Estado
    _colorValorFila(t, 3, _colorRiesgo(riesgo));  // Nivel de riesgo
    body.appendParagraph('').setFontSize(8);
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
  _nota(body,
    'Evidencia fotográfica de los ítems verificados, en orden de ítem, con la fecha de ' +
    'inspección y la observación o hallazgo correspondiente según el estado registrado.');

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

    // Mismo criterio que en Hallazgos: el título va dentro de la tabla para
    // que un salto de página no lo deje solo al final del folio.
    var t = _tarjetaDatos(body, [
      [itemId, String(f[COL.item-1])],
      ['Fecha de inspección', _fechaStr(f[COL.fecha-1])||'—'],
      ['Estado', est],
      [etiquetaObs, textoObs]
    ]);
    var cabId = t.getRow(0).getCell(0), cabNombre = t.getRow(0).getCell(1);
    cabId.setBackgroundColor(C_FONDO);
    cabNombre.setBackgroundColor(C_FONDO);
    cabId.getChild(0).asParagraph()
      .setFontSize(9).setBold(true).setFontFamily(FUENTE_MONO).setForegroundColor(C_ENCABEZADO);
    cabNombre.getChild(0).asParagraph()
      .setFontSize(11).setBold(true).setForegroundColor(C_ENCABEZADO);
    _colorValorFila(t, 2, esConforme ? C_OK : C_ACENTO);

    var fi = fotos[itemId];
    _insertarFotos(body, fi.actual, 'Fotografía — estado actual');
    _insertarFotos(body, fi.cierre, 'Fotografía — cierre');
    body.appendParagraph('').setFontSize(8);
  });
}

/** Galería de hasta 2 fotos en una fila de 2 columnas, con el nombre del
 *  archivo como pie en monoespaciada — la "photo card" del mockup. */
function _insertarFotos(body, archivos, rotulo){
  if(!archivos || !archivos.length) return;
  _rotulo(body, rotulo);
  var anchoCelda = Math.floor(ANCHO_UTIL/2);            // 252 pt por columna
  // Una sola foto no se estira a todo el ancho: se topa en 330 pt para que
  // una imagen vertical no ocupe la página entera.
  var maxW = archivos.length>1 ? anchoCelda-24 : 330;
  var tabla = body.appendTable();
  var fila = tabla.appendTableRow();
  archivos.slice(0,2).forEach(function(file){
    var celda = fila.appendTableCell('');
    celda.setWidth(archivos.length>1 ? anchoCelda : ANCHO_UTIL)
         .setBackgroundColor('#FFFFFF')
         .setPaddingTop(8).setPaddingBottom(8).setPaddingLeft(8).setPaddingRight(8);
    var img = celda.appendImage(file.getBlob());
    var esc = Math.min(1, maxW / img.getWidth());
    img.setWidth(Math.round(img.getWidth()*esc));
    img.setHeight(Math.round(img.getHeight()*esc));
    celda.appendParagraph(file.getName())
         .setFontSize(7).setFontFamily(FUENTE_MONO).setForegroundColor(C_TITULO)
         .setSpacingBefore(4).setSpacingAfter(0);
  });
  // Con una sola foto la tabla queda de una columna a ancho completo — no
  // se agrega celda vacía de relleno (Docs no acepta ancho 0).
  tabla.setBorderColor(C_CELDA).setBorderWidth(0.75);
}

/* ---------- Anexos técnicos A-D: ficha del escenario ----------
   Leen ficha.motorResultado y ficha.aforoResultado tal como los calculó y
   persistió la PWA (mismo objeto, sin recalcular nada acá) — el informe
   documenta lo que el inspector vio en pantalla, no una versión distinta. */
function _anexoFichaEscenario(body, ficha){
  body.appendPageBreak();
  _h1(body, '9. Anexo técnico A — Ficha del escenario');
  var area = ficha.areaManual || ((ficha.largo && ficha.ancho) ? (ficha.largo*ficha.ancho) : null);
  _h2(body, 'A.1 Geometría del vaso');
  _tarjetaDatos(body, [
    ['Largo del espejo de agua', ficha.largo!=null ? ficha.largo+' m' : '—'],
    ['Ancho del espejo de agua', ficha.ancho!=null ? ficha.ancho+' m' : '—'],
    ['Área directa (si forma irregular)', ficha.areaManual!=null ? ficha.areaManual+' m²' : '—'],
    ['Área usada en los cálculos', area!=null ? area.toFixed(1)+' m²' : '— sin dato —'],
    ['Profundidad máxima', ficha.profMax!=null ? ficha.profMax+' m' : '—'],
    ['Profundidad mínima', ficha.profMin!=null ? ficha.profMin+' m' : '—'],
    ['Profundidad intermedia', ficha.profIntermedia!=null ? ficha.profIntermedia+' m' : '—']
  ], {mono:true, pctEtiqueta:0.45});

  _h2(body, 'A.2 Clasificación de uso');
  _tarjetaDatos(body, [
    ['Tipo de uso del estanque', TIPO_USO_LABEL[ficha.tipoUso] || '— sin dato —'],
    ['¿Bajo cubierta / recinto cerrado?', ficha.cubierta==='si' ? 'Sí' : (ficha.cubierta==='no' ? 'No' : '— sin dato —')]
  ], {pctEtiqueta:0.45});
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
  _h2(body, 'B.1 Punto de operación');
  _tarjetaDatos(body, [
    ['Curva de bomba usada', String(r.origen||'—')],
    ['Fecha del cálculo', r.ts ? Utilities.formatDate(new Date(r.ts), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '—'],
    ['Caudal en el punto de operación', r.caudal!=null ? r.caudal.toFixed(2)+' m³/h' : '— sin dato —'],
    ['Cabezal en el punto de operación', r.cabezal!=null ? r.cabezal.toFixed(2)+' m c.a.' : '— sin dato —']
  ], {pctEtiqueta:0.45});

  _h2(body, 'B.2 Velocidades y recirculación');
  _tarjetaDatos(body, [
    ['Velocidad en succión', r.vSuccion!=null ? r.vSuccion.toFixed(2)+' m/s' : '— sin dato —'],
    ['Velocidad en descarga/retorno', r.vDescarga!=null ? r.vDescarga.toFixed(2)+' m/s' : '— sin dato —'],
    ['Velocidad de filtración', r.vFiltracion!=null ? r.vFiltracion.toFixed(1)+' m³/h/m²' : '— sin dato —'],
    ['Volumen estimado del vaso', r.volumen!=null ? r.volumen.toFixed(1)+' m³' : '— sin dato —'],
    ['Tiempo de recirculación', r.tiempoRecirc!=null ? r.tiempoRecirc.toFixed(2)+' h' : '— sin dato —'],
    ['Rotaciones estimadas por día', r.rotacionesDia!=null ? r.rotacionesDia.toFixed(1) : '— sin dato —']
  ], {mono:true, pctEtiqueta:0.45});

  if(r.tramos && r.tramos.length){
    _h2(body, 'B.3 Detalle de velocidad por tramo');
    var filasTramoV = [['Tramo','Ø (pulg)','Lado','Líneas','Velocidad (m/s)','Estado']];
    var excede = [null];
    r.tramos.forEach(function(t){
      var limite = t.lado==='succion' ? 1.8 : 2.4;
      filasTramoV.push([
        t.nombre || '(sin nombre)',
        String(t.diametro),
        t.lado==='succion' ? 'Succión' : 'Presión',
        String(t.nLineas>0 ? t.nLineas : 1),
        t.v.toFixed(2),
        t.v > limite ? 'Supera '+limite+' m/s' : 'Cumple'
      ]);
      excede.push(t.v > limite);
    });
    var tTramoV = body.appendTable(filasTramoV);
    _estiloTabla(tTramoV, true, [135, 60, 65, 50, 90, 105]);
    for(var i=1;i<tTramoV.getNumRows();i++){
      var col = excede[i] ? C_ACENTO : C_OK;
      tTramoV.getRow(i).getCell(4).getChild(0).asParagraph().setFontFamily(FUENTE_MONO).setBold(true).setForegroundColor(col);
      tTramoV.getRow(i).getCell(5).getChild(0).asParagraph().setBold(true).setForegroundColor(col);
    }
  }
  _nota(body,
    'Límites normativos de referencia: succión ≤ 1.8 m/s, descarga ≤ 2.4 m/s (Numeral 10.1); ' +
    'filtración 20-40 m³/h/m² (50 en uso restringido, Numeral 10.2); tiempo de recirculación ' +
    'según Tabla No. 1 del Anexo Técnico.');
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
  _h2(body, 'C.1 Aforo máximo');
  _tarjetaDatos(body, [
    ['Área del espejo de agua', r.area.toFixed(1)+' m²'],
    ['Factor de uso aplicado (Tabla 4)', r.factorUso+' m²/bañista — '+String(r.criterioProf||'')],
    ['Aforo máximo (Numeral 10.9)', r.aforo+' bañistas'],
    ['Fecha del cálculo', r.ts ? Utilities.formatDate(new Date(r.ts), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '—']
  ], {pctEtiqueta:0.45});

  _h2(body, 'C.2 Dotación sanitaria requerida (Tabla 5)');
  var etiquetas = {duchas:'Duchas', inodoroH:'Inodoros hombres', inodoroM:'Inodoros mujeres',
                    orinal:'Orinales', lavamanos:'Lavamanos', vestier:'Vestieres'};
  var datos = [['Elemento (Tabla 5)','Ratio normativo','Requerido','Encontrado','Estado']];
  var estados = [null];
  Object.keys(etiquetas).forEach(function(k){
    var d = r.dotacion && r.dotacion[k];
    if(!d) return;
    var encontrado = (d.encontrado==null) ? '— sin dato —' : String(d.encontrado);
    var falta = (d.encontrado!=null && d.encontrado < d.requerido);
    var estadoTxt = (d.encontrado==null) ? 'Pendiente de conteo' : (falta ? 'Faltan '+(d.requerido-d.encontrado) : 'Cumple');
    datos.push([etiquetas[k], d.ratio, String(d.requerido), encontrado, estadoTxt]);
    estados.push(d.encontrado==null ? 'sin dato' : (falta ? 'falta' : 'ok'));
  });
  var t = body.appendTable(datos);
  _estiloTabla(t, true, [125, 130, 80, 85, 85]);
  for(var i=1;i<t.getNumRows();i++){
    var col = estados[i]==='ok' ? C_OK : (estados[i]==='falta' ? C_ACENTO : C_AMBAR);
    t.getRow(i).getCell(4).getChild(0).asParagraph().setBold(true).setForegroundColor(col);
  }
}

function _anexoMemoriaCalculo(body, ficha){
  body.appendPageBreak();
  _h1(body, '12. Anexo técnico D — Memoria de cálculo');
  _nota(body,
    'Estimaciones de apoyo a la decisión del inspector: no reemplazan la medición directa con ' +
    'caudalímetro ni el conteo físico de la dotación sanitaria. Los supuestos declarados a ' +
    'continuación permiten reproducir o auditar cada resultado de los Anexos B y C.');

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
  _tarjetaDatos(body, filasBomba, {mono:true, pctEtiqueta:0.45});

  _h2(body, 'D.2 Datos de tubería y filtro');
  var tramosValidos = (ficha.tuberiaTramos||[]).filter(function(t){ return t && t.diametro>0; });
  if(tramosValidos.length){
    var filasTramos = [['Tramo','Ø (pulg)','Lado','Longitud (m)','Accesorios','Líneas paralelas']];
    tramosValidos.forEach(function(t){
      filasTramos.push([
        t.nombre || '(sin nombre)',
        String(t.diametro),
        t.lado==='succion' ? 'Succión' : 'Presión',
        t.longitud!=null ? String(t.longitud) : '—',
        t.accesorios!=null ? String(t.accesorios) : '—',
        String(t.nLineas>0 ? t.nLineas : 1)
      ]);
    });
    var tTramos = body.appendTable(filasTramos);
    _estiloTabla(tTramos, true, [125, 60, 65, 85, 80, 90]);
  } else {
    _tarjetaDatos(body, [
      ['Diámetro de succión', ficha.tuberiaSuccionDiam!=null ? ficha.tuberiaSuccionDiam+' pulg' : '—'],
      ['Diámetro de descarga/retorno', ficha.tuberiaDescargaDiam!=null ? ficha.tuberiaDescargaDiam+' pulg' : '—'],
      ['Longitud total de tubería', ficha.tuberiaLongitud!=null ? ficha.tuberiaLongitud+' m' : '—'],
      ['Número de accesorios', ficha.tuberiaAccesorios!=null ? String(ficha.tuberiaAccesorios) : '—'],
      ['Reparto succión/descarga', ficha.tuberiaPctSuccion!=null ? ficha.tuberiaPctSuccion+'% / '+(100-ficha.tuberiaPctSuccion)+'%' : '50% / 50% (por defecto)']
    ], {mono:true, pctEtiqueta:0.45});
  }
  body.appendParagraph('').setFontSize(4);
  _tarjetaDatos(body, [
    ['Desnivel succión-descarga', ficha.desnivelSuccionDescarga!=null ? ficha.desnivelSuccionDescarga+' m' : '—'],
    ['Tipo de filtro', FILTRO_TIPO_LABEL[ficha.filtroTipo] || '—'],
    ['Área filtrante', ficha.filtroArea!=null ? ficha.filtroArea+' m²' : '—'],
    ['Presión de manómetro del filtro', ficha.presionManometro!=null ? ficha.presionManometro+' PSI' : '—']
  ], {mono:true, pctEtiqueta:0.45});

  var sup = (ficha.motorResultado && ficha.motorResultado.supuestos) || null;
  if(sup){
    _h2(body, 'D.3 Supuestos del cálculo hidráulico');
    _tarjetaDatos(body, [
      ['Coeficiente de Hazen-Williams (C)', String(sup.C)],
      ['Longitud equivalente de accesorios', String(sup.LeqPorAccesorio)],
      ['Carga estática asumida', String(sup.Hgeo)],
      ['Reparto de tubería succión/descarga', String(sup.repartoLongitud)],
      ['Conversión de presión', String(sup.conversionPsi)]
    ], {mono:true, pctEtiqueta:0.45});
  }

  if(ficha.aforoResultado && !ficha.aforoResultado.error){
    _h2(body, 'D.4 Supuesto del aforo y dotación');
    _nota(body,
      'La fila de la Tabla No. 4 (factor de uso) se determina con la profundidad MÁXIMA del vaso ' +
      '— criterio conservador. La dotación sanitaria requerida (Tabla No. 5) se calcula sobre el ' +
      '100% del aforo para cada elemento; en campo puede repartirse por género según la ' +
      'composición real de bañistas.');
  }
}

function _planAccion(body, filas){
  body.appendPageBreak();
  _h1(body, '5. Plan de acción y matriz de cierre');
  var datos = [['ID','Ítem','Riesgo','Acción','Responsable','Compromiso','Avance']];
  var riesgos = [null]; // paralelo a `datos`, para colorear la celda de riesgo
  filas.forEach(function(f){
    var est = String(f[COL.estado-1]||'');
    if(est==='Cumple' || est==='No aplica' || est==='') return;
    var riesgo = String(f[COL.riesgo-1]||'—');
    datos.push([
      String(f[COL.id-1]),
      String(f[COL.item-1]).slice(0,70),
      riesgo,
      String(f[COL.accion-1]||'—').slice(0,70),
      String(f[COL.respCierre-1]||'—'),
      _fechaStr(f[COL.fechaCompromiso-1])||'—',
      (f[COL.avance-1]!==''? f[COL.avance-1]+'%':'—')
    ]);
    riesgos.push(riesgo);
  });
  if(datos.length===1){ body.appendParagraph('Sin acciones abiertas.'); return; }
  var t = body.appendTable(datos);
  _estiloTabla(t, true, [55, 110, 50, 110, 80, 60, 40]);
  for(var r=1;r<t.getNumRows();r++){
    t.getRow(r).getCell(2).getChild(0).asParagraph()
     .setForegroundColor(_colorRiesgo(riesgos[r])).setBold(true);
    t.getRow(r).getCell(0).getChild(0).asParagraph().setFontFamily(FUENTE_MONO).setFontSize(8);
  }
}

function _conclusion(body, m, sede, piscina){
  _h1(body, '6. Conclusión');
  var txt = 'El vaso ' + piscina + ' de la sede ' + sede + ' presenta un cumplimiento global del ' +
    m.pctCumplimiento + '% frente a los criterios de la Resolución 929 de 2026, con ' +
    m.noCumple + ' ítem(s) en estado No cumple y ' + (m.critico + m.alto) +
    ' hallazgo(s) clasificados en riesgo crítico o alto. ';
  if(m.vencidos>0) txt += 'Se registran ' + m.vencidos + ' hallazgo(s) con fecha compromiso vencida, ' +
    'lo que constituye la desviación de gestión más relevante del período. ';
  if(m.sinFecha>0) txt += m.sinFecha + ' hallazgo(s) permanecen sin fecha compromiso asignada. ';
  txt += 'La condición general se califica como: ' + _semaforo(m.pctCumplimiento) + '.';
  // Texto corrido en color normal: el rojo se reserva para estados y niveles
  // de riesgo puntuales, no para párrafos enteros.
  body.appendParagraph(txt).setForegroundColor(C_TITULO);
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
  _estiloTabla(t, true, [140, 365]);

  _rotulo(body, 'Firmas y aprobación');
  _tarjetaDatos(body, [
    ['Elaboró', String(responsable||'—')],
    ['Revisó',''],
    ['Aprobó','']
  ], {pctEtiqueta:0.3});
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
/* Tabla tipo "matriz" (encabezado + filas de datos, N columnas).
   `widths` es opcional: array de anchos en pt por columna — sin él Docs
   reparte el ancho solo, que en tablas de 6-7 columnas suele descuadrar. */
function _estiloTabla(tabla, conEncabezado, widths){
  tabla.setBorderColor(C_CELDA).setBorderWidth(0.75);
  for(var r=0;r<tabla.getNumRows();r++){
    var fila = tabla.getRow(r);
    var esEncabezado = conEncabezado && r===0;
    for(var c=0;c<fila.getNumCells();c++){
      var celda = fila.getCell(c);
      if(widths && widths[c]) celda.setWidth(widths[c]);
      celda.setBackgroundColor(esEncabezado ? C_HEADER_TABLA : (r%2===0 ? '#FFFFFF' : C_FONDO));
      celda.setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(8).setPaddingRight(8);
      var p = celda.getChild(0).asParagraph();
      p.setFontSize(esEncabezado ? 8.5 : 9).setForegroundColor(esEncabezado ? C_ENCABEZADO : C_TITULO);
      if(esEncabezado) p.setBold(true);
    }
  }
}

/* ---------- Gráficas (servicio Charts) ----------
   Genera imágenes de gráfica reales (PNG) para incrustar en el Doc — a
   diferencia del mockup HTML/SVG de Stitch, DocumentApp no soporta CSS ni
   SVG animado, así que el equivalente correcto en un informe de Google
   Docs es el servicio Charts (Charts.new...().build().getBlob()). */
function _chartDonutCumplimiento(pct){
  var colorOk = pct>=95 ? C_FILL_CUMPLE : (pct>=85 ? C_FILL_CUMPLE : (pct>=70 ? C_FILL_PROCESO : C_FILL_NOCUMPLE));
  var dt = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Estado')
    .addColumn(Charts.ColumnType.NUMBER, 'Valor')
    .addRow(['Cumplimiento', pct])
    .addRow(['Restante', Math.max(0, 100-pct)])
    .build();
  var chart = Charts.newPieChart()
    .setDataTable(dt)
    .setDimensions(280, 280)
    .setColors([colorOk, '#E2E8F0'])
    .setLegendPosition(Charts.Position.NONE)
    .setOption('pieHole', 0.68)
    .setOption('pieSliceText', 'none')
    .setOption('backgroundColor', 'transparent')
    .setOption('tooltip', {trigger:'none'})
    .setOption('chartArea', {left:6, top:6, width:'92%', height:'92%'})
    .build();
  return chart.getBlob();
}

function _chartDistribucionEstados(m){
  var dt = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Estado')
    .addColumn(Charts.ColumnType.NUMBER, 'Cantidad')
    .addRow(['Cumple', m.cumple])
    .addRow(['No cumple', m.noCumple])
    .addRow(['En proceso', m.enProceso])
    .addRow(['Pendiente', m.pendiente])
    .addRow(['No aplica', m.noAplica])
    .build();
  // Se renderiza al doble del tamaño al que se inserta (220 pt) para que el
  // PDF no la muestre pixelada; por eso las fuentes van en 15, que al
  // reducir quedan en ~7 pt reales.
  var chart = Charts.newPieChart()
    .setDataTable(dt)
    .setDimensions(440, 300)
    .setColors([C_FILL_CUMPLE, C_FILL_NOCUMPLE, C_FILL_PROCESO, C_FILL_PENDIENTE, C_FILL_NOAPLICA])
    .setOption('pieHole', 0.45)
    .setOption('backgroundColor', 'transparent')
    .setOption('pieSliceTextStyle', {fontSize:14, color:'#FFFFFF'})
    // Se define legend en un único setOption (posición + estilo juntos) —
    // combinarlo con setLegendPosition() pisaría este objeto o al revés,
    // según el orden de evaluación interno del builder.
    .setOption('legend', {position:'right', textStyle:{fontSize:15, color:C_TITULO}})
    .setOption('chartArea', {left:10, top:10, width:'62%', height:'88%'})
    .build();
  return chart.getBlob();
}

function _chartCapitulos(m){
  var caps = Object.keys(m.porCapitulo).map(function(cap){
    var c = m.porCapitulo[cap];
    var base = c.total - c.noAplica - c.fueraAlcance;
    var pct = base>0 ? Math.round(100*c.cumpleEnAlcance/base) : 0;
    // Los nombres de capítulo son largos ("1. Documentación técnica y
    // legal") — se recortan al número + primeras palabras para que la
    // gráfica quepa en el ancho A4 sin desbordar las etiquetas del eje Y.
    var etiqueta = cap.length>34 ? cap.slice(0,32)+'…' : cap;
    return {etiqueta:etiqueta, pct:pct};
  }).sort(function(a,b){ return a.pct - b.pct; }); // peor primero, arriba en la barra horizontal

  var dtb = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Capítulo')
    .addColumn(Charts.ColumnType.NUMBER, '% cumplimiento')
    .addColumn(Charts.ColumnType.NUMBER, '% restante');
  caps.forEach(function(c){ dtb.addRow([c.etiqueta, c.pct, 100-c.pct]); });
  var dt = dtb.build();

  // Igual que la torta: se renderiza a ~1.5x del ancho al que se inserta
  // (ANCHO_UTIL) y las fuentes se escalan en la misma proporción.
  var anchoRender = 760;
  var alturaPorFila = 26;
  var chart = Charts.newBarChart()
    .setDataTable(dt)
    .setDimensions(anchoRender, Math.max(180, caps.length*alturaPorFila + 60))
    .setColors([C_FILL_CUMPLE, '#E2E8F0'])
    .setStacked()
    .setLegendPosition(Charts.Position.NONE)
    .setOption('backgroundColor', 'transparent')
    .setOption('hAxis', {minValue:0, maxValue:100, textStyle:{fontSize:12, color:C_TITULO}})
    .setOption('vAxis', {textStyle:{fontSize:12, color:C_ENCABEZADO}})
    .setOption('bar', {groupWidth:'72%'})
    .setOption('chartArea', {left:'40%', top:10, width:'56%', height:'90%'})
    .build();
  return chart.getBlob();
}

/** Inserta una imagen centrada en el cuerpo del documento, con ancho fijo. */
function _insertarImagenCentrada(body, blob, width){
  var p = body.appendParagraph('');
  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  var img = p.appendInlineImage(blob);
  if(width){
    var esc = width / img.getWidth();
    img.setWidth(width);
    img.setHeight(Math.round(img.getHeight()*esc));
  }
  return img;
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
