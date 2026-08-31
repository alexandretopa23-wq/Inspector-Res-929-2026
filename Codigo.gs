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
   oscuro para texto/encabezados, azul para "conforme", ámbar para riesgo
   medio o en proceso, gris para pendiente/no aplica. El texto nunca va en
   rojo: ni en las tablas del checklist ni en las fichas técnicas ni en las
   notas, para que el informe se lea como un documento técnico y no como una
   alarma. La severidad se distingue por peso tipográfico (negrita + slate
   oscuro) y por la palabra misma ("Crítico", "No cumple"), no por el color.
   Los nombres de las constantes se mantienen (C_TITULO, C_CRITICO, etc.)
   porque los usan ~20 funciones de anexos más abajo — solo cambian los
   valores. */
var C_ENCABEZADO='#0F172A',      // slate-900 — títulos H1/portada
    C_TITULO='#475569',          // slate-600 — texto secundario/H2
    C_CRITICO='#1E293B',         // slate-800, negrita — texto de severidad alta (crítico/no cumple), sin rojo
    C_CELDA='#E2E8F0',           // slate-200 — bordes y filas alternas
    C_FONDO='#F8FAFC',           // slate-50 — filas pares
    C_OK='#2563EB';              // azul-600 — texto "conforme" (coherente con la app)

/* Colores de relleno para gráficas e insignias (más saturados que los de
   texto de arriba, pensados para áreas de color, no para letras. El rojo
   de aquí es de gráfica, no de texto, y por eso queda fuera de la regla
   de "sin rojo" que aplica a las tablas y fichas). */
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
  eq:'Pentair EQ Series comercial (curva de fábrica)',
  intelliflo_vs:'Pentair IntelliFlo VS+SVRS 3 HP, velocidad seleccionable (dato de fábrica)',
  whisperflo_vst:'Pentair WhisperFlo VST 2 HP, velocidad seleccionable (dato de fábrica)',
  whisperflo_xf:'Pentair WhisperFloXF 3 HP, velocidad única (dato de fábrica)',
  manual:'Curva real medida en campo (mínimo 4 puntos Q-H)'
};
/* Un mapa de velocidades por familia de bomba de velocidad variable: cada
   fabricante rotula sus preselecciones con RPM distintas, así que la ficha
   guarda el índice y el informe traduce con el mapa de SU familia. */
var BOMBA_VELOCIDAD_VS_LABEL = {
  '1':'Velocidad 1 (3000 RPM)', '2':'Velocidad 2 (2200 RPM)',
  '3':'Velocidad 3 (1400 RPM)', '4':'Velocidad 4 o máxima (3450 RPM)'
};
var BOMBA_VELOCIDAD_INTELLIFLO_LABEL = {
  '1':'Velocidad 1 (1100 RPM)', '2':'Velocidad 2 (1500 RPM)',
  '3':'Velocidad 3 (2350 RPM)', '4':'Velocidad 4 (3110 RPM)',
  '5':'Velocidad 5 o máxima (3450 RPM)'
};
var BOMBA_VELOCIDAD_WHISPER_VST_LABEL = {
  '1':'Velocidad 1 (1400 RPM)', '2':'Velocidad 2 (2200 RPM)',
  '3':'Velocidad 3 (3000 RPM)', '4':'Velocidad 4 o máxima (3450 RPM)'
};
/* Familias de velocidad variable → mapa de rótulos y campo de la ficha donde
   quedó guardada la selección. Evita repetir el mismo if/else en D.1. */
var BOMBA_VELOCIDAD_POR_FAMILIA = {
  superflo_vs:    {campo:'bombaVelocidadVS',           labels:BOMBA_VELOCIDAD_VS_LABEL},
  intelliflo_vs:  {campo:'bombaVelocidadIntelliflo',   labels:BOMBA_VELOCIDAD_INTELLIFLO_LABEL},
  whisperflo_vst: {campo:'bombaVelocidadWhisperVST',   labels:BOMBA_VELOCIDAD_WHISPER_VST_LABEL}
};
var FILTRO_TIPO_LABEL = {
  arena:'Arena / medio granular', cartucho:'Cartucho', de:'Tierra de diatomeas (D.E.)'
};
var FILTRO_VALVULA_LABEL = {
  bateria:'Batería de válvulas o colector (bocas de entrada y salida a distinta altura)',
  selectora:'Válvula selectora multipuerto (entrada y salida en la misma válvula, a la misma cota)'
};
/* Selects del modelo hidráulico v2 (circuito con tanque de compensación). */
var TANQUE_LABEL = {
  si:'Sí, la bomba succiona desde el tanque de compensación',
  no:'No, la bomba succiona directo del vaso'
};
var RETORNO_LABEL = {
  sumergido:'Boquillas sumergidas bajo la lámina',
  sobreLamina:'Descarga sobre la lámina (cascada, chorro o canal)'
};
var ESTADO_TUBERIA_LABEL = {
  nueva:'Nueva o en buen estado (C=150)',
  servicio:'Con años de servicio (C=130)',
  incrustada:'Con incrustación o estrechamiento visible (C=110)'
};
var CALENTADOR_TIPO_LABEL = {
  gas:'Caldera o calentador a gas con intercambiador',
  bombaCalor:'Bomba de calor',
  placas:'Intercambiador de placas'
};
var CALENTADOR_ARREGLO_LABEL = {
  serie:'En serie (el caudal pasa por todos, uno tras otro)',
  paralelo:'En paralelo (el caudal se reparte entre ellos)'
};
/* Acople del calentador al circuito principal (motor v2.1). Es el dato que
   decide si el bloque de calentadores suma ΔP a la curva del sistema: con
   bomba inyectora propia no lo hace, con bypass en línea lo hace con techo,
   y en línea directa lo hace sin techo. */
var CALENTADOR_ACOPLE_LABEL = {
  bypass_linea:'Derivación con válvula de estrangulamiento en la línea principal (tipo Pentair ETi 400)',
  inyector:'Derivación con bomba inyectora propia del calentador (tipo Raypak XTherm P), desacoplada de la línea principal',
  linea:'En línea directa, sin bypass (todo el caudal del circuito pasa por el intercambiador)'
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
    r.avisos.push('La pestaña usada no coincide con el GID configurado en el script (GID_DATOS='+GID_DATOS+'), así que se cayó a la primera pestaña como respaldo. Si esto es intencional (moviste/renombraste la pestaña), está bien; si no, revisa que la pestaña "'+sh.getName()+'" sea realmente la de datos.');
  }

  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  r.filas = Math.max(0, lastRow-1);
  r.columnas = lastCol;

  if(lastRow < 1){
    r.ok=false; r.errores.push('La pestaña está completamente vacía: ni siquiera tiene fila de encabezado.');
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
    r.errores.push('El encabezado no coincide con la estructura de 22 columnas en '+difs.length+' columna(s). Por eso el dashboard puede leer datos en el campo equivocado o simplemente no cargar.');
    r.diferenciasEncabezado = difs;
  }

  if(r.filas === 0){
    r.avisos.push('El encabezado está correcto pero todavía no hay filas de datos, algo normal si aún no se ha sincronizado ninguna inspección desde la app.');
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
      aviso:'La hoja anclada ('+id+') ya no es accesible para la cuenta del script, así que la app está cayendo al respaldo mientras tanto.'};
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
        // linkCarpeta es la única columna que la PWA nunca conoce: la escribe
        // el servidor en _marcarEvidenciaEnHoja() cuando sube la primera foto
        // del ítem, y el navegador no la vuelve a leer de vuelta. Si la fila
        // entrante la trae vacía pero la existente ya tenía un valor, se
        // conserva la existente — de lo contrario, resincronizar el capítulo
        // (p.ej. tras cambiar el estado o el hallazgo) borraría el link a la
        // carpeta de Drive aunque la carpeta y las fotos sigan ahí intactas.
        if(!row[COL.linkCarpeta-1] && datos[idx][COL.linkCarpeta-1]){
          row[COL.linkCarpeta-1] = datos[idx][COL.linkCarpeta-1];
        }
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
/* Genera el informe de un vaso. Si `area` viene con un valor de AREA_ORDEN,
   el informe sale segmentado: solo lleva los ítems que le corresponden a ese
   equipo (más CHK-003, el plano consolidado, que es compartido por las tres
   áreas técnicas), y el % de cumplimiento se recalcula sobre ese subconjunto.
   Sin `area` el comportamiento es el de siempre: informe completo del vaso. */
function generarInformeVaso(sede, piscina, fecha, area){
  var sh = _hojaDatos();
  var last = sh.getLastRow();
  if(last<2) return {ok:false, error:'La hoja no tiene datos aún'};
  var todo = sh.getRange(2,1,last-1,TOTAL_COLS).getValues();

  var filas = todo.filter(function(f){
    return String(f[COL.sede-1])===sede && String(f[COL.piscina-1])===piscina &&
           _fechaStr(f[COL.fecha-1])===String(fecha);
  });
  if(!filas.length) return {ok:false, error:'Sin registros para ese vaso y fecha'};

  if(area){
    filas = _filasDeArea(filas, area);
    if(!filas.length) return {ok:false, error:'Sin ítems del área '+area+' para ese vaso y fecha'};
  }

  var m = _metricas(filas);
  var fotos = _indiceFotos(sede, piscina, fecha);
  var ficha = _obtenerFicha(sede, piscina, fecha);

  var sufijo = area ? '_'+_slugArea(area) : '';
  var doc  = DocumentApp.create('Informe_Inspeccion_'+sede+'_'+piscina+'_'+fecha+sufijo);
  var body = doc.getBody();
  body.setPageWidth(595).setPageHeight(842); // A4 en puntos
  body.clear();
  _configurarEstilosDoc(body);

  _portada(body, sede, piscina, fecha, filas[0][COL.responsable-1], m, area);
  _objetivo(body, area);
  _tableroKPI(body, m);
  _tablaCapitulos(body, m);
  _hallazgosConFotos(body, filas, fotos);
  _planAccion(body, filas);
  _conclusion(body, m, sede, piscina, area);
  _responsabilidades(body, filas[0][COL.responsable-1]);
  _anexoFotografico(body, filas, fotos);
  // Los Anexos A-D solo salen si el inspector diligenció la ficha técnica
  // en la PWA para este vaso+fecha (botón "📐 Ficha") — un informe sin
  // ficha se genera igual, simplemente sin estas 4 secciones finales.
  // Los Anexos A-D son de dimensionamiento hidráulico: en los informes
  // segmentados solo tienen sentido para Térmica e Hidráulica. El informe
  // completo (sin área) los sigue trayendo como siempre.
  var conFichaTecnica = !!(ficha && (!area || area==='Térmica e Hidráulica'));
  if(conFichaTecnica){
    _anexoFichaEscenario(body, ficha);
    _anexoFichaHidraulica(body, ficha);
    _anexoDimensionamiento(body, ficha);
    _anexoMemoriaCalculo(body, ficha);
  }
  // Las referencias normativas y técnicas van siempre, incluso sin ficha:
  // la Resolución 929 (y la 234, cuando hay ítems fuera de su alcance) se
  // citan por nombre desde el Objetivo en adelante. Crane Co. y el
  // Hydraulic Institute solo entran si el modelo v2 llegó a imprimirse
  // (Anexo B, secciones B.6/B.7, y Anexo D.3), porque son la fuente de esos
  // dos criterios concretos y no de nada más en el documento.
  var conCriteriosV2 = conFichaTecnica && ficha.motorResultado && ficha.motorResultado.modelo==='v2';
  _referencias(body, conCriteriosV2, conFichaTecnica ? 13 : 9);

  doc.saveAndClose();

  // Los informes por área quedan en una subcarpeta propia por área, para que
  // cada interesado reciba un link de carpeta con solo lo suyo.
  var ruta = [RAIZ_DRIVE, sede, piscina, fecha];
  if(area) ruta.push('AREA_'+_slugArea(area));
  var carpeta = _carpetaRuta(ruta);
  var pdf = carpeta.createFile(DriveApp.getFileById(doc.getId()).getAs('application/pdf'))
                   .setName('Informe_'+sede+'_'+piscina+'_'+fecha+sufijo+'.pdf');
  return {ok:true, area:area||null, docUrl:doc.getUrl(), pdfUrl:pdf.getUrl(), metricas:m};
}

/* Genera de una vez los informes segmentados de un vaso: uno por cada área
   que tenga ítems diligenciados. Devuelve la lista de resultados (las áreas
   sin ítems simplemente no producen informe, no son un error). */
function generarInformesPorArea(sede, piscina, fecha){
  var res = [];
  // "Sin clasificar" va al final para que un CHK nuevo sin área asignada no
  // quede sin informe (se perdería para todos los interesados); si no hay
  // ítems así, simplemente no se genera ese documento.
  AREA_ORDEN.concat(['Sin clasificar']).forEach(function(area){
    // "Compartido (los 3)" no genera informe propio: sus ítems ya viajan
    // dentro de los tres informes técnicos.
    if(area==='Compartido (los 3)') return;
    var r = generarInformeVaso(sede, piscina, fecha, area);
    if(r.ok) res.push(r);
  });
  return res;
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

/* ---------- Reparto de ítems por área responsable ----------
   Clasificación acordada con el equipo (2026-08): cada CHK-xxx se asigna a
   quien debe cerrar ese hallazgo, no al capítulo normativo. La mayoría de
   capítulos cae completo en un área; los que mezclan oficios (Documentación,
   Vaso de piscina, Canaletas/cárcamos, Drenajes, Cuarto de equipos,
   Químicos) se resolvieron ítem por ítem. "Compartido (los 3)" es un único
   ítem (CHK-003, el plano consolidado hidráulico+eléctrico+gas) que no se
   puede partir entre áreas. Si se agregan ítems al checklist (nuevo CHK-xxx)
   y no aparecen aquí, _areaDe() los manda a "Sin clasificar" para que se
   note en el dashboard en vez de perderse en silencio dentro de otra área. */
var AREA_ORDEN = ['Eléctrico', 'Infraestructura', 'Térmica e Hidráulica', 'Gestión de sede', 'Compartido (los 3)'];
var ID_A_AREA = {
  'CHK-001':'Infraestructura', 'CHK-002':'Infraestructura', 'CHK-003':'Compartido (los 3)', 'CHK-004':'Eléctrico',
  'CHK-005':'Térmica e Hidráulica', 'CHK-006':'Térmica e Hidráulica', 'CHK-007':'Térmica e Hidráulica', 'CHK-008':'Infraestructura',
  'CHK-009':'Infraestructura', 'CHK-010':'Eléctrico', 'CHK-011':'Infraestructura', 'CHK-012':'Térmica e Hidráulica',
  'CHK-013':'Eléctrico', 'CHK-014':'Térmica e Hidráulica', 'CHK-015':'Eléctrico', 'CHK-016':'Térmica e Hidráulica',
  'CHK-017':'Gestión de sede', 'CHK-018':'Térmica e Hidráulica', 'CHK-019':'Térmica e Hidráulica', 'CHK-020':'Térmica e Hidráulica',
  'CHK-021':'Térmica e Hidráulica', 'CHK-022':'Térmica e Hidráulica', 'CHK-023':'Eléctrico', 'CHK-024':'Térmica e Hidráulica',
  'CHK-025':'Gestión de sede', 'CHK-026':'Infraestructura', 'CHK-027':'Infraestructura', 'CHK-028':'Infraestructura',
  'CHK-029':'Infraestructura', 'CHK-030':'Infraestructura', 'CHK-031':'Infraestructura', 'CHK-032':'Infraestructura',
  'CHK-033':'Infraestructura', 'CHK-034':'Infraestructura', 'CHK-035':'Infraestructura', 'CHK-036':'Infraestructura',
  'CHK-037':'Infraestructura', 'CHK-038':'Infraestructura', 'CHK-039':'Infraestructura', 'CHK-040':'Térmica e Hidráulica',
  'CHK-041':'Térmica e Hidráulica', 'CHK-042':'Infraestructura', 'CHK-043':'Infraestructura', 'CHK-044':'Infraestructura',
  'CHK-045':'Infraestructura', 'CHK-046':'Infraestructura', 'CHK-047':'Infraestructura', 'CHK-048':'Infraestructura',
  'CHK-049':'Infraestructura', 'CHK-050':'Infraestructura', 'CHK-051':'Infraestructura', 'CHK-052':'Infraestructura',
  'CHK-053':'Infraestructura', 'CHK-054':'Infraestructura', 'CHK-055':'Infraestructura', 'CHK-056':'Infraestructura',
  'CHK-057':'Infraestructura', 'CHK-058':'Térmica e Hidráulica', 'CHK-059':'Térmica e Hidráulica', 'CHK-060':'Térmica e Hidráulica',
  'CHK-061':'Infraestructura', 'CHK-062':'Infraestructura', 'CHK-063':'Térmica e Hidráulica', 'CHK-064':'Térmica e Hidráulica',
  'CHK-065':'Térmica e Hidráulica', 'CHK-066':'Térmica e Hidráulica', 'CHK-067':'Térmica e Hidráulica', 'CHK-068':'Térmica e Hidráulica',
  'CHK-069':'Térmica e Hidráulica', 'CHK-070':'Eléctrico', 'CHK-071':'Térmica e Hidráulica', 'CHK-072':'Térmica e Hidráulica',
  'CHK-073':'Térmica e Hidráulica', 'CHK-074':'Térmica e Hidráulica', 'CHK-075':'Térmica e Hidráulica', 'CHK-076':'Térmica e Hidráulica',
  'CHK-077':'Térmica e Hidráulica', 'CHK-078':'Térmica e Hidráulica', 'CHK-079':'Térmica e Hidráulica', 'CHK-080':'Térmica e Hidráulica',
  'CHK-081':'Térmica e Hidráulica', 'CHK-082':'Térmica e Hidráulica', 'CHK-083':'Térmica e Hidráulica', 'CHK-084':'Térmica e Hidráulica',
  'CHK-085':'Térmica e Hidráulica', 'CHK-086':'Térmica e Hidráulica', 'CHK-087':'Térmica e Hidráulica', 'CHK-088':'Térmica e Hidráulica',
  'CHK-089':'Térmica e Hidráulica', 'CHK-090':'Térmica e Hidráulica', 'CHK-091':'Térmica e Hidráulica', 'CHK-092':'Térmica e Hidráulica',
  'CHK-093':'Térmica e Hidráulica', 'CHK-094':'Térmica e Hidráulica', 'CHK-095':'Térmica e Hidráulica', 'CHK-096':'Térmica e Hidráulica',
  'CHK-097':'Térmica e Hidráulica', 'CHK-098':'Térmica e Hidráulica', 'CHK-099':'Térmica e Hidráulica', 'CHK-100':'Térmica e Hidráulica',
  'CHK-101':'Térmica e Hidráulica', 'CHK-102':'Térmica e Hidráulica', 'CHK-103':'Térmica e Hidráulica', 'CHK-104':'Térmica e Hidráulica',
  'CHK-105':'Térmica e Hidráulica', 'CHK-106':'Térmica e Hidráulica', 'CHK-107':'Térmica e Hidráulica', 'CHK-108':'Térmica e Hidráulica',
  'CHK-109':'Térmica e Hidráulica', 'CHK-110':'Térmica e Hidráulica', 'CHK-111':'Térmica e Hidráulica', 'CHK-112':'Térmica e Hidráulica',
  'CHK-113':'Eléctrico', 'CHK-114':'Eléctrico', 'CHK-115':'Térmica e Hidráulica', 'CHK-116':'Eléctrico',
  'CHK-117':'Eléctrico', 'CHK-118':'Eléctrico', 'CHK-119':'Eléctrico', 'CHK-120':'Eléctrico',
  'CHK-121':'Eléctrico', 'CHK-122':'Eléctrico', 'CHK-123':'Eléctrico', 'CHK-124':'Eléctrico',
  'CHK-125':'Térmica e Hidráulica', 'CHK-126':'Térmica e Hidráulica', 'CHK-127':'Térmica e Hidráulica', 'CHK-128':'Térmica e Hidráulica',
  'CHK-129':'Térmica e Hidráulica', 'CHK-130':'Térmica e Hidráulica', 'CHK-131':'Térmica e Hidráulica', 'CHK-132':'Térmica e Hidráulica',
  'CHK-133':'Eléctrico', 'CHK-134':'Eléctrico', 'CHK-135':'Eléctrico', 'CHK-136':'Eléctrico',
  'CHK-137':'Eléctrico', 'CHK-138':'Eléctrico', 'CHK-139':'Eléctrico', 'CHK-140':'Infraestructura',
  'CHK-141':'Térmica e Hidráulica', 'CHK-142':'Infraestructura', 'CHK-143':'Eléctrico', 'CHK-144':'Infraestructura',
  'CHK-145':'Térmica e Hidráulica', 'CHK-146':'Térmica e Hidráulica', 'CHK-147':'Térmica e Hidráulica', 'CHK-148':'Infraestructura',
  'CHK-149':'Infraestructura', 'CHK-150':'Térmica e Hidráulica', 'CHK-151':'Térmica e Hidráulica', 'CHK-152':'Térmica e Hidráulica',
  'CHK-153':'Térmica e Hidráulica', 'CHK-154':'Térmica e Hidráulica', 'CHK-155':'Térmica e Hidráulica', 'CHK-156':'Térmica e Hidráulica',
  'CHK-157':'Térmica e Hidráulica', 'CHK-158':'Térmica e Hidráulica', 'CHK-159':'Térmica e Hidráulica', 'CHK-160':'Térmica e Hidráulica',
  'CHK-161':'Térmica e Hidráulica', 'CHK-162':'Infraestructura', 'CHK-163':'Infraestructura', 'CHK-164':'Infraestructura',
  'CHK-165':'Infraestructura', 'CHK-166':'Infraestructura', 'CHK-167':'Infraestructura', 'CHK-168':'Infraestructura',
  'CHK-169':'Infraestructura', 'CHK-170':'Infraestructura', 'CHK-171':'Infraestructura', 'CHK-172':'Infraestructura',
  'CHK-173':'Infraestructura', 'CHK-174':'Infraestructura', 'CHK-175':'Gestión de sede', 'CHK-176':'Gestión de sede',
  'CHK-177':'Gestión de sede', 'CHK-178':'Infraestructura', 'CHK-179':'Gestión de sede', 'CHK-180':'Gestión de sede',
  'CHK-181':'Gestión de sede', 'CHK-182':'Gestión de sede', 'CHK-183':'Gestión de sede', 'CHK-184':'Gestión de sede',
  'CHK-185':'Gestión de sede', 'CHK-186':'Gestión de sede', 'CHK-187':'Gestión de sede', 'CHK-188':'Gestión de sede',
  'CHK-189':'Gestión de sede', 'CHK-190':'Gestión de sede', 'CHK-191':'Gestión de sede', 'CHK-192':'Gestión de sede',
  'CHK-193':'Gestión de sede', 'CHK-194':'Gestión de sede', 'CHK-195':'Gestión de sede', 'CHK-196':'Gestión de sede',
  'CHK-197':'Infraestructura', 'CHK-198':'Infraestructura', 'CHK-199':'Infraestructura', 'CHK-200':'Infraestructura',
  'CHK-201':'Infraestructura', 'CHK-202':'Infraestructura', 'CHK-203':'Infraestructura', 'CHK-204':'Gestión de sede',
  'CHK-205':'Gestión de sede', 'CHK-206':'Gestión de sede', 'CHK-207':'Gestión de sede', 'CHK-208':'Gestión de sede',
  'CHK-209':'Gestión de sede', 'CHK-210':'Térmica e Hidráulica', 'CHK-211':'Térmica e Hidráulica', 'CHK-212':'Térmica e Hidráulica',
  'CHK-213':'Térmica e Hidráulica', 'CHK-214':'Térmica e Hidráulica', 'CHK-215':'Térmica e Hidráulica', 'CHK-216':'Infraestructura',
  'CHK-217':'Infraestructura', 'CHK-218':'Infraestructura', 'CHK-219':'Infraestructura', 'CHK-220':'Infraestructura',
  'CHK-221':'Infraestructura', 'CHK-222':'Infraestructura', 'CHK-223':'Térmica e Hidráulica', 'CHK-224':'Térmica e Hidráulica',
  'CHK-225':'Térmica e Hidráulica', 'CHK-226':'Térmica e Hidráulica', 'CHK-227':'Térmica e Hidráulica', 'CHK-228':'Térmica e Hidráulica',
  'CHK-229':'Térmica e Hidráulica', 'CHK-230':'Térmica e Hidráulica', 'CHK-231':'Térmica e Hidráulica', 'CHK-232':'Térmica e Hidráulica',
  'CHK-233':'Térmica e Hidráulica', 'CHK-234':'Térmica e Hidráulica', 'CHK-235':'Térmica e Hidráulica', 'CHK-236':'Térmica e Hidráulica',
  'CHK-237':'Térmica e Hidráulica', 'CHK-238':'Térmica e Hidráulica', 'CHK-239':'Térmica e Hidráulica', 'CHK-240':'Térmica e Hidráulica',
  'CHK-241':'Térmica e Hidráulica', 'CHK-242':'Térmica e Hidráulica', 'CHK-243':'Térmica e Hidráulica', 'CHK-244':'Térmica e Hidráulica',
  'CHK-245':'Gestión de sede', 'CHK-246':'Gestión de sede', 'CHK-247':'Infraestructura', 'CHK-248':'Térmica e Hidráulica'
};
function _areaDe(id){ return ID_A_AREA[id] || 'Sin clasificar'; }

/* Áreas técnicas que además cargan con los ítems marcados "Compartido (los
   3)" (hoy solo CHK-003, el plano consolidado). Gestión de sede no entra:
   ese plano no es su responsabilidad. */
var AREAS_CON_COMPARTIDOS = {'Eléctrico':1, 'Infraestructura':1, 'Térmica e Hidráulica':1};

/* Filtra las filas de un vaso dejando solo las del área pedida. Un ítem
   compartido aparece en los tres informes técnicos a propósito: es un
   entregable que las tres áreas deben coordinar, y omitirlo en dos de los
   tres informes lo haría desaparecer de la conversación. */
function _filasDeArea(filas, area){
  return filas.filter(function(f){
    var a = _areaDe(String(f[COL.id-1]));
    if(a===area) return true;
    return a==='Compartido (los 3)' && !!AREAS_CON_COMPARTIDOS[area];
  });
}

/* Nombre de área apto para nombres de archivo y carpetas de Drive. */
function _slugArea(area){
  return String(area)
    .replace(/[áÁ]/g,'a').replace(/[éÉ]/g,'e').replace(/[íÍ]/g,'i')
    .replace(/[óÓ]/g,'o').replace(/[úÚ]/g,'u').replace(/[ñÑ]/g,'n')
    .replace(/[^A-Za-z0-9]+/g,'_').replace(/^_+|_+$/g,'').toUpperCase();
}

/* ---------- Métricas ejecutivas ---------- */
function _metricas(filas){
  var m = {total:filas.length, cumple:0, noCumple:0, enProceso:0, pendiente:0, noAplica:0, fueraAlcance:0,
           cumpleEnAlcance:0, noCumpleEnAlcance:0, enProcesoEnAlcance:0, pendienteEnAlcance:0, noAplicaEnAlcance:0,
           baseEnAlcance:0, fueraDelDenominador:0,
           critico:0, alto:0, medio:0, bajo:0, vencidos:0, sinFecha:0, avanceProm:0,
           porCapitulo:{}, porArea:{}, criticosAltos:[]};
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
    if(!m.porCapitulo[cap]) m.porCapitulo[cap] = {total:0, cumple:0, noCumple:0, noAplica:0, fueraAlcance:0, cumpleEnAlcance:0, otros:0, fueraDelDenominador:0};
    m.porCapitulo[cap].total++;
    if(esFueraAlcance){ m.fueraAlcance++; m.porCapitulo[cap].fueraAlcance++; }

    // Mismo acumulador que porCapitulo pero agrupado por área responsable
    // (Eléctrico/Infraestructura/Térmica e Hidráulica/Gestión de sede), para
    // poder repartir hallazgos entre los equipos sin reimplementar el cálculo
    // de cumplimiento en cada lugar que lo necesite.
    var area = _areaDe(id);
    if(!m.porArea[area]) m.porArea[area] = {total:0, cumple:0, noCumple:0, noAplica:0, fueraAlcance:0, cumpleEnAlcance:0, otros:0, fueraDelDenominador:0};
    m.porArea[area].total++;
    if(esFueraAlcance) m.porArea[area].fueraAlcance++;

    // Los conteos brutos (cumple/noCumple/...) incluyen TODO, fuera de
    // alcance o no — Hallazgos y Plan de acción filtran directamente sobre
    // `filas`, no sobre estos contadores, así que un ítem "No cumple" fuera
    // de alcance igual aparece ahí. Las variantes "...EnAlcance" son las que
    // excluyen fuera de alcance en cada estado — se necesitan las cinco (no
    // solo cumpleEnAlcance) para poder dibujar un desglose disjunto que sume
    // el total exacto, sin asumir que solo "Cumple" puede quedar fuera del
    // alcance de la Res. 929.
    if(est==='Cumple'){
      m.cumple++; m.porCapitulo[cap].cumple++; m.porArea[area].cumple++;
      if(!esFueraAlcance){ m.cumpleEnAlcance++; m.porCapitulo[cap].cumpleEnAlcance++; m.porArea[area].cumpleEnAlcance++; }
    }
    else if(est==='No cumple'){ m.noCumple++; m.porCapitulo[cap].noCumple++; m.porArea[area].noCumple++; if(!esFueraAlcance) m.noCumpleEnAlcance++; }
    else if(est==='En proceso'){ m.enProceso++; m.porCapitulo[cap].otros++; m.porArea[area].otros++; if(!esFueraAlcance) m.enProcesoEnAlcance++; }
    else if(est==='Pendiente'){ m.pendiente++; m.porCapitulo[cap].otros++; m.porArea[area].otros++; if(!esFueraAlcance) m.pendienteEnAlcance++; }
    else if(est==='No aplica'){ m.noAplica++; m.porCapitulo[cap].noAplica++; m.porArea[area].noAplica++; if(!esFueraAlcance) m.noAplicaEnAlcance++; }

    // Denominador del %: un ítem sale del denominador si es "No aplica" O
    // fuera de alcance (unión, no suma) — así uno que fuera ambas cosas a la
    // vez no se resta dos veces y el % no queda inflado. Mismo criterio por
    // capítulo y por área, para que _tablaCapitulos/_areasDashboard y sus
    // gráficos no repitan el bug de restar dos veces un ítem "No aplica" +
    // fuera de alcance.
    if(est==='No aplica' || esFueraAlcance){ m.fueraDelDenominador++; m.porCapitulo[cap].fueraDelDenominador++; m.porArea[area].fueraDelDenominador++; }

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

  m.baseEnAlcance = m.total - m.fueraDelDenominador;
  m.pctCumplimiento = m.baseEnAlcance>0 ? Math.round(100*m.cumpleEnAlcance/m.baseEnAlcance) : 0;
  m.avanceProm = conAvance? Math.round(sumAvance/conAvance) : 0;
  return m;
}

/* ---------- Bloque de KPIs + capítulos para un conjunto de filas ----------
   Se usa igual para el total de la hoja y para las filas de un solo vaso, de
   modo que cuando el dashboard filtra por la inspección que el usuario toca,
   los números salgan de esta misma función y no de un recuento paralelo en
   JavaScript que podría divergir de _metricas(). */
function _capitulosDashboard(m){
  return Object.keys(m.porCapitulo).map(function(cap){
    var c = m.porCapitulo[cap];
    var base = c.total - c.fueraDelDenominador;
    var pct = base>0 ? Math.round(100*c.cumpleEnAlcance/base) : 0;
    return {capitulo:cap, total:c.total, cumple:c.cumple, noCumple:c.noCumple,
             fueraAlcance:c.fueraAlcance||0, pctCumplimiento:pct};
  }).sort(function(a,b){ return a.pctCumplimiento - b.pctCumplimiento; });
}

/* ---------- Reparto de hallazgos por área responsable ----------
   Misma mecánica que _capitulosDashboard pero sobre m.porArea, para
   alimentar la tabla "Repartición por área" del dashboard (la que se envía
   a Eléctrico/Infraestructura/Térmica e Hidráulica/Gestión de sede). El
   orden es AREA_ORDEN fijo — no se ordena por % de cumplimiento como los
   capítulos — porque esta tabla es de enrutamiento (cada área siempre debe
   encontrarse en el mismo lugar), no un ranking de peores capítulos. */
function _areasDashboard(m){
  return AREA_ORDEN.filter(function(area){ return !!m.porArea[area]; }).map(function(area){
    var c = m.porArea[area];
    var base = c.total - c.fueraDelDenominador;
    var pct = base>0 ? Math.round(100*c.cumpleEnAlcance/base) : 0;
    return {area:area, total:c.total, cumple:c.cumple, noCumple:c.noCumple,
             fueraAlcance:c.fueraAlcance||0, pctCumplimiento:pct};
  });
}

/* Prioridad de riesgo para ordenar la lista de hallazgos abiertos: lo que el
   ingeniero necesita ver primero al presentar resultados es lo crítico, no
   el orden en que quedó la fila en la hoja. */
var _RIESGO_ORDEN = {'Critico':0,'Crítico':0,'Alto':1,'Medio':2,'Bajo':3};

/* ---------- Detalle de ítems "No cumple" para el dashboard ----------
   A diferencia de _metricas() (que solo cuenta), esto arma la lista
   completa de hallazgos abiertos con el texto que el ingeniero necesita
   para presentar resultados: qué ítem, por qué no cumple, qué acción y
   quién responde. Se calcula sobre el mismo `filas` que ya recibe
   _resumenDashboard, así que respeta el filtro por vaso sin duplicar lógica. */
function _hallazgosAbiertos(filas){
  var lista = [];
  filas.forEach(function(f){
    if(String(f[COL.estado-1]||'')!=='No cumple') return;
    var id = String(f[COL.id-1]);
    var fc = f[COL.fechaCompromiso-1], cierre = f[COL.fechaCierre-1];
    var hoy = new Date(); hoy.setHours(0,0,0,0);
    var vencido = !!fc && !cierre && new Date(fc) < hoy;
    lista.push({
      id:id,
      capitulo:String(f[COL.capitulo-1]||''),
      item:String(f[COL.item-1]||''),
      sede:String(f[COL.sede-1]||''),
      piscina:String(f[COL.piscina-1]||''),
      riesgo:String(f[COL.riesgo-1]||''),
      hallazgo:String(f[COL.hallazgo-1]||''),
      accion:String(f[COL.accion-1]||''),
      respCierre:String(f[COL.respCierre-1]||''),
      fechaCompromiso: fc ? _fechaStr(fc) : '',
      vencido:vencido,
      fueraAlcance: !!FUERA_ALCANCE_929[id],
      // Permite filtrar este mismo panel por área desde el dashboard (tabla
      // "Repartición por área") sin tener que duplicar ID_A_AREA en el
      // cliente ni pedir un endpoint nuevo.
      area: _areaDe(id)
    });
  });
  lista.sort(function(a,b){
    var ra = _RIESGO_ORDEN.hasOwnProperty(a.riesgo) ? _RIESGO_ORDEN[a.riesgo] : 9;
    var rb = _RIESGO_ORDEN.hasOwnProperty(b.riesgo) ? _RIESGO_ORDEN[b.riesgo] : 9;
    if(ra!==rb) return ra-rb;
    if(a.vencido!==b.vencido) return a.vencido ? -1 : 1;
    return a.capitulo.localeCompare(b.capitulo);
  });
  return lista;
}

function _resumenDashboard(filas){
  var m = _metricas(filas);
  // Evidencia faltante: hallazgos "No cumple" sin foto del estado actual —
  // el dato que más le importa a un auditor externo.
  var hallazgosNoCumple = 0, evidenciaFaltante = 0;
  filas.forEach(function(f){
    if(String(f[COL.estado-1]||'')==='No cumple'){
      hallazgosNoCumple++;
      if(!String(f[COL.evidActual-1]||'').trim()) evidenciaFaltante++;
    }
  });
  return {
    global: {
      total:m.total, pctCumplimiento:m.pctCumplimiento,
      // Numerador y denominador reales del %: el dashboard los muestra para
      // que no se lea "cumple" (conteo bruto) como si fuera el numerador.
      cumpleEnAlcance:m.cumpleEnAlcance, baseEnAlcance:m.baseEnAlcance,
      // Variantes "en alcance" de cada estado: permiten dibujar un desglose
      // disjunto (suma exacta = total) sin que un ítem fuera de alcance
      // pinte a la vez su segmento de estado y el de "Fuera de alcance".
      noCumpleEnAlcance:m.noCumpleEnAlcance, enProcesoEnAlcance:m.enProcesoEnAlcance,
      pendienteEnAlcance:m.pendienteEnAlcance, noAplicaEnAlcance:m.noAplicaEnAlcance,
      cumple:m.cumple, noCumple:m.noCumple, enProceso:m.enProceso,
      pendiente:m.pendiente, noAplica:m.noAplica, fueraAlcance:m.fueraAlcance,
      critico:m.critico, alto:m.alto, medio:m.medio, bajo:m.bajo,
      vencidos:m.vencidos, sinFecha:m.sinFecha, avanceProm:m.avanceProm,
      hallazgosNoCumple:hallazgosNoCumple, evidenciaFaltante:evidenciaFaltante,
      hallazgos:_hallazgosAbiertos(filas)
    },
    capitulos: _capitulosDashboard(m),
    areas: _areasDashboard(m)
  };
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

  var resumenGlobal = _resumenDashboard(todo);

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
    // Un solo recorrido de métricas por vaso: los campos de la fila de la
    // lista se leen del mismo resumen que después alimenta los paneles al
    // filtrar, así la fila y los KPIs filtrados nunca pueden discrepar.
    var resumen = _resumenDashboard(v.filas);
    var m = resumen.global;
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
      vencidos:m.vencidos, critico:m.critico, alto:m.alto,
      // Resumen propio del vaso: es lo que el dashboard pinta en los KPIs, el
      // desglose de estados y el gráfico de capítulos cuando el usuario toca
      // esta inspección en la lista. Viene pre-calculado en el servidor para
      // que filtrar sea instantáneo y funcione sin señal.
      resumen: resumen
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

  return {
    ok:true,
    actualizadoEn: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    spreadsheetNombre: ss.getName(),
    hojaNombre: sh.getName(),
    global: resumenGlobal.global,
    sedes: sedes,
    vasos: vasos,
    capitulos: resumenGlobal.capitulos,
    areas: resumenGlobal.areas,
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
function _portada(body, sede, piscina, fecha, responsable, m, area){
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
  // En los informes segmentados el área va bajo el título y también en la
  // ficha de metadatos, para que no haya duda de a quién le corresponde el
  // documento cuando circula suelto por correo.
  if(area){
    colTitulo.appendParagraph('Informe por área responsable · ' + area)
      .setFontSize(9.5).setBold(true).setForegroundColor(C_ENCABEZADO)
      .setSpacingBefore(3);
  }

  var metaFilas = [
    ['SEDE', String(sede)],
    ['VASO', String(piscina)],
    ['FECHA', String(fecha)],
    ['RESPONSABLE', String(responsable||'—')]
  ];
  if(area) metaFilas.splice(3, 0, ['ÁREA', String(area)]);
  var metaCard = _tarjetaEnCelda(colMeta, C_FONDO);
  var metaTabla = metaCard.appendTable(metaFilas);
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
  // Barra apilada al 100% en vez de la torta: con cinco estados, dos de
  // ellos casi siempre en 0, la torta gastaba media pagina para mostrar un
  // reparto que en una barra se lee de un vistazo y a ancho completo. Ademas
  // sale vectorial, alineada con la tabla de conteos que va justo debajo.
  var anchoTarjeta = ANCHO_UTIL - 24;
  _barraApilada(distCard, [
    {valor:m.cumple,    color:C_FILL_CUMPLE},
    {valor:m.noCumple,  color:C_FILL_NOCUMPLE},
    {valor:m.enProceso, color:C_FILL_PROCESO},
    {valor:m.pendiente, color:C_FILL_PENDIENTE},
    {valor:m.noAplica,  color:C_FILL_NOAPLICA}
  ], anchoTarjeta, 16);
  _chipsLeyenda(distCard, [
    {etiqueta:'Cumple',     valor:m.cumple,    color:C_FILL_CUMPLE},
    {etiqueta:'No cumple',  valor:m.noCumple,  color:C_FILL_NOCUMPLE},
    {etiqueta:'En proceso', valor:m.enProceso, color:C_FILL_PROCESO},
    {etiqueta:'Pendiente',  valor:m.pendiente, color:C_FILL_PENDIENTE},
    {etiqueta:'No aplica',  valor:m.noAplica,  color:C_FILL_NOAPLICA}
  ], anchoTarjeta);
  distCard.appendParagraph('').setFontSize(5);

  var distTabla = distCard.appendTable([
    ['Estado','Cantidad','%'],
    ['Cumple', String(m.cumple), (m.total>0?Math.round(100*m.cumple/m.total):0)+'%'],
    ['No cumple', String(m.noCumple), (m.total>0?Math.round(100*m.noCumple/m.total):0)+'%'],
    ['En proceso', String(m.enProceso), (m.total>0?Math.round(100*m.enProceso/m.total):0)+'%'],
    ['Pendiente', String(m.pendiente), (m.total>0?Math.round(100*m.pendiente/m.total):0)+'%'],
    ['No aplica', String(m.noAplica), '—']
  ]);
  _estiloTabla(distTabla, true, [230, 125, 120]);
  var coloresFila = [null, C_OK, C_CRITICO, C_AMBAR, C_TITULO, C_TITULO];
  for(var rr=1; rr<distTabla.getNumRows(); rr++){
    if(!coloresFila[rr]) continue;
    distTabla.getRow(rr).getCell(0).getChild(0).asParagraph().setForegroundColor(coloresFila[rr]).setBold(true);
    distTabla.getRow(rr).getCell(2).getChild(0).asParagraph().setForegroundColor(coloresFila[rr]);
  }
  // Aclaración obligada: esta tabla cuenta en bruto y la dona de arriba cuenta
  // solo lo que está en alcance, así que "Cumple" y el numerador del % nunca
  // coinciden cuando hay ítems fuera de alcance. Sin esta nota la diferencia
  // se lee como un error de cálculo.
  if(m.fueraAlcance > 0){
    _nota(distCard,
      'Los conteos de esta tabla son brutos e incluyen los ' + m.fueraAlcance + ' ítem(s) fuera del alcance de la ' +
      'Res. 929. Por eso "Cumple" (' + m.cumple + ') no coincide con el numerador del cumplimiento global: ese ' +
      'porcentaje se calcula únicamente sobre los ítems en alcance (' + m.cumpleEnAlcance + ' de ' +
      m.baseEnAlcance + ' = ' + m.pctCumplimiento + '%).');
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
  return pct>=85 ? C_OK : (pct>=70 ? C_AMBAR : C_CRITICO);
}
/* Color semántico de un nivel de riesgo (acepta "Critico" y "Crítico"). */
function _colorRiesgo(riesgo){
  var r = String(riesgo||'');
  if(r==='Critico'||r==='Crítico'||r==='Alto') return C_CRITICO;
  if(r==='Medio') return C_AMBAR;
  return C_TITULO;
}
/* Color semántico del estado de un ítem. Antes la columna "Estado" de la
   tarjeta de hallazgo se pintaba siempre igual sin mirar el valor real
   (todo quedaba en rojo, tanto "No cumple" como "En proceso"); ahora sigue
   la misma escala de severidad que el resto del informe. */
function _colorEstado(estado){
  var e = String(estado||'');
  if(e==='No cumple') return C_CRITICO;
  if(e==='En proceso' || e==='Pendiente') return C_AMBAR;
  if(e==='Cumple') return C_OK;
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
    // Alto de fila reducido en todo el informe: padding vertical mínimo y
    // párrafo sin espaciado extra ni interlineado. El texto largo sigue
    // creciendo lo necesario para caber completo, solo se quita el aire.
    etq.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(9).setPaddingRight(9);
    val.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(9).setPaddingRight(9);
    etq.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    val.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    var C = DocumentApp.HorizontalAlignment.CENTER;
    var pe = etq.getChild(0).asParagraph();
    pe.setFontSize(8.5).setBold(true).setForegroundColor(C_TITULO).setAlignment(C)
      .setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
    var pv = val.getChild(0).asParagraph();
    pv.setFontSize(9.5).setBold(false).setForegroundColor(C_ENCABEZADO).setAlignment(C)
      .setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
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

/* Párrafo de cuerpo del informe: Arial 11, peso normal, sin cursiva y
   justificado. Es el único estilo admitido para texto corrido (objetivo,
   hallazgos, conclusión, párrafos y observaciones de los anexos). La letra
   pequeña y en cursiva se reserva para los pies de imagen. */
function _cuerpo(contenedor, txt){
  return contenedor.appendParagraph(txt)
    .setFontFamily(FUENTE).setFontSize(11).setBold(false).setItalic(false)
    .setForegroundColor(C_TITULO)
    .setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY)
    .setSpacingAfter(8);
}

/* Aclaraciones, alcances y observaciones al pie de sección. Antes iban en
   cuerpo pequeño y cursiva; ahora comparten el mismo Arial 11 normal y
   justificado que el resto del texto corrido. */
function _nota(contenedor, txt){
  return _cuerpo(contenedor, txt);
}

function _objetivo(body, area){
  _h1(body, '1. Objetivo');
  _cuerpo(body,
    'Verificar el cumplimiento de los criterios técnicos constructivos y de seguridad ' +
    'establecidos en la Resolución 929 de 2026 y su Anexo Técnico para el vaso de piscina ' +
    'evaluado, documentar las condiciones observadas con evidencia fotográfica, clasificar ' +
    'los hallazgos por nivel de riesgo y establecer el plan de acción con responsables y ' +
    'fechas de cierre.');

  // Aclaración de alcance: sin ella, un lector que compara dos informes
  // segmentados del mismo vaso ve dos porcentajes distintos y asume error.
  if(area){
    _nota(body,
      'Alcance de este documento: únicamente los ítems del checklist cuya responsabilidad de ' +
      'cierre corresponde al área de ' + area + '. Los porcentajes, conteos y gráficas de este ' +
      'informe se calculan sobre ese subconjunto, por lo que no coinciden con el cumplimiento ' +
      'global del vaso ni con el de los informes de las demás áreas. Los ítems marcados como ' +
      'compartidos entre las tres áreas técnicas se incluyen también en los informes de las otras dos.');
  }
}

function _tableroKPI(body, m){
  // La distribución de hallazgos (gráfica + tabla) ya se muestra en el
  // "Resumen ejecutivo" de la portada (_portada) — esta sección cubre los
  // indicadores de gestión que no caben en esa tarjeta compacta, sin
  // repetir la misma gráfica dos veces en el informe.
  _h1(body, '2. Tablero de indicadores de gestión');

  // Las cuatro cifras que decide un gerente de un vistazo van arriba como
  // tarjetas; la tabla de abajo queda para el detalle y la columna de
  // lectura. Antes las ocho filas pesaban igual y ninguna resaltaba.
  var kpis = _filaColumnas(body, [126, 126, 126, 127]);
  _tarjetaKPI(kpis[0], 'Ítems evaluados', m.total,
    m.fueraAlcance>0 ? (m.baseEnAlcance+' en alcance 929') : 'Todos en alcance 929', C_ENCABEZADO);
  _tarjetaKPI(kpis[1], 'Riesgo crítico', m.critico,
    m.critico>0 ? 'Atención inmediata' : 'Sin críticos',
    m.critico>0 ? C_CRITICO : C_OK);
  _tarjetaKPI(kpis[2], 'Hallazgos vencidos', m.vencidos,
    m.vencidos>0 ? 'Fuera de plazo' : 'Al día',
    m.vencidos>0 ? C_CRITICO : C_OK);
  _tarjetaKPI(kpis[3], 'Avance de cierre', m.avanceProm+'%',
    'Promedio de los hallazgos abiertos', _colorPct(m.avanceProm));
  body.appendParagraph('').setFontSize(6);

  var t = body.appendTable([
    ['Indicador','Valor','Lectura'],
    ['Ítems evaluados', String(m.total), '—'],
    ['Fuera del alcance de la Res. 929', String(m.fueraAlcance), 'Excluido del cálculo: es dominio de la Res. 234'],
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
      t.getRow(i).getCell(2).getChild(0).asParagraph().setForegroundColor(C_CRITICO).setBold(true);
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

  _barrasCapitulos(body, m);
  _chipsLeyenda(body, [
    {etiqueta:'≥ 85% conforme',  valor:'', color:C_FILL_CUMPLE},
    {etiqueta:'70-84% deficiente', valor:'', color:C_FILL_PROCESO},
    {etiqueta:'< 70% crítico',    valor:'', color:C_FILL_NOCUMPLE},
    {etiqueta:'Faltante',          valor:'', color:C_CELDA}
  ], ANCHO_UTIL);
  body.appendParagraph('').setFontSize(6);

  var datos = [['Capítulo','Ítems','Cumple','No cumple','Fuera de alcance','% cumplimiento (Res. 929)']];
  var pcts = [null]; // paralelo a `datos`, guarda el % de cada fila para colorear después
  // Se ordena de menor a mayor % de cumplimiento para que la tabla siga la
  // misma secuencia que la gráfica de barras de arriba (crítico primero).
  var capsOrdenados = Object.keys(m.porCapitulo).map(function(cap){
    var c = m.porCapitulo[cap];
    var base = c.total - c.fueraDelDenominador;
    return {cap:cap, c:c, pct: base>0 ? Math.round(100*c.cumpleEnAlcance/base) : 0};
  }).sort(function(a,b){ return a.pct - b.pct; });
  capsOrdenados.forEach(function(o){
    var cap = o.cap, c = o.c;
    // FIX histórico: antes el % por capítulo se calculaba contra c.total
    // (incluyendo "No aplica"), mientras que el % global excluía "No
    // aplica" del denominador — no reconciliaban. Ahora ambos usan el mismo
    // criterio: base = total - fueraDelDenominador (unión de "No aplica" y
    // fuera de alcance, sin restar dos veces un ítem que sea ambos), y el
    // numerador es cumpleEnAlcance (no cumple, que sí incluiría ítems sin
    // numeral en la 929 y podría superar el 100% frente a la base reducida).
    var base = c.total - c.fueraDelDenominador;
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
    'de forma consolidada en el anexo de registro fotográfico (sección 8).');
  var conHallazgo = filas.filter(function(f){
    var est = String(f[COL.estado-1]||'');
    return est==='No cumple' || est==='En proceso' || est==='Pendiente';
  });

  if(!conHallazgo.length){
    _cuerpo(body, 'No se registraron hallazgos abiertos en esta inspección.')
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
    _colorValorFila(t, 2, _colorEstado(String(f[COL.estado-1])));  // Estado
    _colorValorFila(t, 3, _colorRiesgo(riesgo));                    // Nivel de riesgo
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
  _h1(body, '8. Anexo técnico: registro fotográfico');
  _nota(body,
    'Evidencia fotográfica de los ítems verificados, en orden de ítem, con la fecha de ' +
    'inspección y la observación o hallazgo correspondiente según el estado registrado.');

  var conFoto = filas.filter(function(f){
    var fi = fotos[String(f[COL.id-1])];
    return fi && (fi.actual.length || fi.cierre.length);
  });

  if(!conFoto.length){
    _cuerpo(body, 'No se registró evidencia fotográfica en esta inspección.');
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
    _colorValorFila(t, 2, esConforme ? C_OK : C_CRITICO);

    var fi = fotos[itemId];
    _insertarFotos(body, fi.actual, 'Fotografía del estado actual');
    _insertarFotos(body, fi.cierre, 'Fotografía de cierre');
    body.appendParagraph('').setFontSize(8);
  });
}

/** Galería de hasta 2 fotos en una fila de 2 columnas, con el nombre del
 *  archivo como pie en monoespaciada — la "photo card" del mockup. */
function _insertarFotos(body, archivos, rotulo){
  if(!archivos || !archivos.length) return;
  _rotulo(body, rotulo);
  var anchoCelda = Math.floor(ANCHO_UTIL/2);            // 252 pt por columna
  // Mismo ancho tope para todas las fotos —con una o con dos— para que el
  // anexo quede parejo y ocupe menos folios. La imagen y su pie van
  // centrados en la celda.
  var CEN = DocumentApp.HorizontalAlignment.CENTER;
  var maxW = 190;
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
    img.getParent().asParagraph().setAlignment(CEN);
    celda.appendParagraph(file.getName())
         .setFontSize(7).setFontFamily(FUENTE_MONO).setForegroundColor(C_TITULO)
         .setItalic(true).setAlignment(CEN)
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
  _h1(body, '9. Anexo técnico A: ficha del escenario');
  var area = ficha.areaManual || ((ficha.largo && ficha.ancho) ? (ficha.largo*ficha.ancho) : null);
  _h2(body, 'A.1 Geometría del vaso');
  _tarjetaDatos(body, [
    ['Largo del espejo de agua', ficha.largo!=null ? ficha.largo+' m' : '—'],
    ['Ancho del espejo de agua', ficha.ancho!=null ? ficha.ancho+' m' : '—'],
    ['Área directa (si forma irregular)', ficha.areaManual!=null ? ficha.areaManual+' m²' : '—'],
    ['Área usada en los cálculos', area!=null ? area.toFixed(1)+' m²' : 'Sin dato'],
    ['Profundidad máxima', ficha.profMax!=null ? ficha.profMax+' m' : '—'],
    ['Profundidad mínima', ficha.profMin!=null ? ficha.profMin+' m' : '—'],
    ['Profundidad intermedia', ficha.profIntermedia!=null ? ficha.profIntermedia+' m' : '—'],
    ['Área según ficha de diseño', ficha.areaDiseno!=null ? ficha.areaDiseno+' m²' : '—'],
    ['Volumen declarado en campo', ficha.volumenManual!=null
      ? ficha.volumenManual+' m³ (reemplaza el cálculo por geometría)' : 'Se calcula por geometría']
  ], {mono:true, pctEtiqueta:0.45});

  _h2(body, 'A.2 Clasificación de uso');
  _tarjetaDatos(body, [
    ['Tipo de uso del estanque', TIPO_USO_LABEL[ficha.tipoUso] || 'Sin dato'],
    ['¿Bajo cubierta / recinto cerrado?', ficha.cubierta==='si' ? 'Sí' : (ficha.cubierta==='no' ? 'No' : 'Sin dato')]
  ], {pctEtiqueta:0.45});
}

function _anexoFichaHidraulica(body, ficha){
  body.appendPageBreak();
  _h1(body, '10. Anexo técnico B: ficha hidráulica');
  var r = ficha.motorResultado;
  if(!r || r.error){
    _cuerpo(body, r && r.error ? ('Sin resultado válido: '+r.error) : 'No se calculó el caudal y las velocidades para este vaso en la PWA.');
    return;
  }
  var esV2 = (r.modelo === 'v2');

  _h2(body, 'B.1 Punto de operación');
  var filasB1 = [
    ['Modelo de cálculo', esV2
      ? 'v2, circuito con tanque de compensación, calentadores y método K por accesorio'
      : 'v1 (legado)'],
    ['Curva de bomba usada', String(r.origen||'—')],
    ['Fecha del cálculo', r.ts ? Utilities.formatDate(new Date(r.ts), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '—'],
    ['Caudal en el punto de operación', r.caudal!=null ? r.caudal.toFixed(2)+' m³/h' : 'Sin dato']
  ];
  if(esV2){
    filasB1.push(['Origen del caudal',
      r.fuenteCaudal==='medido' ? 'MEDIDO con caudalímetro en sitio (manda sobre la estimación)'
      : r.fuenteCaudal==='manometro' ? 'DESPEJADO DEL MANÓMETRO: reemplaza al modelo porque el caudal de la presión medida quedó muy por debajo del modelado — hay una restricción real (válvula estrangulada, válvula multipuerto del filtro, o bomba por debajo de su curva) que el modelo de tramos no captura'
      : 'Estimado por el modelo']);
    // Cuando manda el caudalímetro o el manómetro, el caudal del modelo no
    // desaparece: se reporta al lado para que se vea cuánto se apartó la
    // estimación, que es justamente lo que calibra la confianza en el modelo
    // en los vasos donde NO hay instrumento.
    // caudalModelo* (motor >= 2026-09) guarda lo que estimó el modelo aunque
    // mande otra fuente; fichas viejas con 'medido' lo tenían en caudalNominal.
    var qModeloRep = (r.caudalModelo!=null) ? r.caudalModelo
      : (r.fuenteCaudal==='medido' ? r.caudalNominal : null);
    if(r.fuenteCaudal==='medido' && qModeloRep!=null){
      var brecha = (qModeloRep - r.caudal)/r.caudal*100;
      filasB1.push(['Caudal que estimaba el modelo', qModeloRep.toFixed(2)+' m³/h ('+
        (brecha>=0?'+':'')+brecha.toFixed(1)+' % frente a la medición)']);
    }
    if(r.fuenteCaudal==='manometro' && r.caudalModelo!=null){
      var brechaM = (r.caudalModelo - r.caudal)/r.caudal*100;
      filasB1.push(['Caudal que estimaba el modelo de tramos', r.caudalModelo.toFixed(2)+' m³/h ('+
        (brechaM>=0?'+':'')+brechaM.toFixed(1)+' % frente al manómetro)']);
    }
    var bLo = r.caudalMin, bHi = r.caudalMax, bLbl = 'Banda de incertidumbre del modelo';
    if(r.fuenteCaudal==='manometro') bLbl = 'Banda del caudal (despejada del manómetro)';
    else if(r.fuenteCaudal==='medido'){ bLo = r.caudalModeloMin; bHi = r.caudalModeloMax; bLbl = 'Banda que estimaba el modelo'; }
    if(bLo!=null && bHi!=null){
      filasB1.push([bLbl, 'entre '+bLo.toFixed(1)+' y '+bHi.toFixed(1)+' m³/h']);
    }
  }
  filasB1.push(['Cabezal en el punto de operación', r.cabezal!=null ? r.cabezal.toFixed(2)+' m c.a.' : 'Sin dato']);
  _tarjetaDatos(body, filasB1, {pctEtiqueta:0.45});

  _h2(body, 'B.2 Velocidades y recirculación');
  var filasB2 = [
    ['Velocidad en succión', r.vSuccion!=null ? r.vSuccion.toFixed(2)+' m/s' : 'Sin dato'],
    ['Velocidad en descarga/retorno', r.vDescarga!=null ? r.vDescarga.toFixed(2)+' m/s' : 'Sin dato'],
    ['Velocidad de filtración', r.vFiltracion!=null ? r.vFiltracion.toFixed(1)+' m³/h/m²' : 'Sin dato'],
    ['Volumen estimado del vaso', r.volumen!=null ? r.volumen.toFixed(1)+' m³' : 'Sin dato']
  ];
  if(esV2 && r.volumenTanque!=null){
    filasB2.push(['Volumen del tanque de compensación', r.volumenTanque.toFixed(1)+' m³']);
    filasB2.push(['Volumen total en circulación', r.volumenSistema!=null ? r.volumenSistema.toFixed(1)+' m³' : 'Sin dato']);
  }
  filasB2.push(['Tiempo de recirculación normativo', r.tiempoRecirc!=null ? r.tiempoRecirc.toFixed(2)+' h' : 'Sin dato']);
  if(esV2 && r.tiempoRenovacionSistema!=null && r.volumenTanque!=null){
    filasB2.push(['Renovación real del sistema (vaso + tanque)', r.tiempoRenovacionSistema.toFixed(2)+' h']);
  }
  filasB2.push(['Rotaciones estimadas por día', r.rotacionesDia!=null ? r.rotacionesDia.toFixed(1) : 'Sin dato']);
  if(r.volumenNota) filasB2.push(['Origen del volumen', String(r.volumenNota)]);
  _tarjetaDatos(body, filasB2, {mono:true, pctEtiqueta:0.45});

  /* Alerta de caudal por rama del calentador (motor v2.1). Solo se dispara
     con acople 'linea' declarado: es un hallazgo de instalación, no un
     resultado de cálculo, y por eso va destacado en rojo y no en una tarjeta
     de datos más. */
  if(esV2 && r.alertaCalentador){
    body.appendParagraph('CAUDAL POR ENCIMA DEL MÁXIMO DE FÁBRICA DEL CALENTADOR')
        .setFontSize(9).setBold(true).setForegroundColor(C_CRITICO).setSpacingBefore(8).setSpacingAfter(2);
    body.appendParagraph(String(r.alertaCalentador))
        .setFontSize(8.5).setForegroundColor(C_CRITICO).setSpacingAfter(8);
  }

  if(esV2 && r.tiempoRenovacionSistema!=null && r.volumenTanque!=null){
    _nota(body,
      'El tiempo de recirculación normativo se calcula sobre el volumen del vaso, que es la base ' +
      'de la Tabla No. 1: el Numeral 10.5 dimensiona el tanque de compensación como un porcentaje ' +
      '"del volumen del agua", de modo que ese volumen de referencia no puede incluirse a sí mismo. ' +
      'La renovación del sistema se reporta aparte, como lectura sanitaria complementaria, ' +
      'porque el agua que retorna al vaso viene mezclada con el inventario del tanque. Por sí sola ' +
      'no es un incumplimiento.');
  }

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
        t.v > limite ? 'Supera los '+limite+' m/s' : 'Cumple'
      ]);
      excede.push(t.v > limite);
    });
    var tTramoV = body.appendTable(filasTramoV);
    _estiloTabla(tTramoV, true, [135, 60, 65, 50, 90, 105]);
    for(var i=1;i<tTramoV.getNumRows();i++){
      var col = excede[i] ? C_CRITICO : C_OK;
      tTramoV.getRow(i).getCell(4).getChild(0).asParagraph().setFontFamily(FUENTE_MONO).setBold(true).setForegroundColor(col);
      tTramoV.getRow(i).getCell(5).getChild(0).asParagraph().setBold(true).setForegroundColor(col);
    }
  }
  /* Secciones exclusivas del modelo v2. Un informe generado a partir de una
     inspección v1 no las incluye: no hay dato con qué llenarlas y no se
     inventa nada. */
  if(esV2){
    var VEREDICTO_TXT = {
      cumple:   'CUMPLE CON HOLGURA. Toda la banda de incertidumbre queda dentro de norma, así que la conclusión se sostiene sin medición directa.',
      duda:     'ZONA DE DUDA. La banda cruza el límite normativo, de modo que el modelo no alcanza para concluir. Este vaso requiere medición con caudalímetro.',
      incumple: 'INCUMPLE CON HOLGURA. Toda la banda queda fuera de norma, así que el hallazgo se sostiene sin medición directa.'
    };
    var v = r.veredictos || {};
    if(v.tiempoRecirc || v.vFiltracion){
      _h2(body, 'B.4 Veredicto sobre la banda de incertidumbre');
      var filasV = [];
      if(v.tiempoRecirc) filasV.push(['Tiempo de recirculación', VEREDICTO_TXT[v.tiempoRecirc.clase] || '—']);
      if(v.vFiltracion)  filasV.push(['Velocidad de filtración', VEREDICTO_TXT[v.vFiltracion.clase] || '—']);
      _tarjetaDatos(body, filasV, {pctEtiqueta:0.30});
    }

    if(r.cruceManometro || r.diseno || r.areaContraste || r.deltaPFiltro){
      _h2(body, 'B.5 Verificación cruzada y contraste con el diseño');
      var filasX = [];
      if(r.deltaPFiltro){
        var dpf = r.deltaPFiltro;
        filasX.push(['ΔP bruto del filtro (entrada − salida)',
          dpf.psi.toFixed(1)+' PSI ('+dpf.mca.toFixed(2)+' m c.a.)']);
        var DZ_ORIGEN = {
          medido:        'medida en sitio',
          selectora:     'el filtro tiene válvula selectora multipuerto, de modo que la entrada y la salida son dos bocas de la misma válvula, a la misma cota',
          tipicoDeclarado:'se declaró batería de válvulas o colector, así que se aplica el desnivel típico entre la boca de entrada y la de salida del filtro (entrada más alta)',
          tipicoAsumido: 'no se declaró el tipo de conexión del filtro (ítem CHK-083), así que se asume una batería de válvulas con el desnivel típico entre la boca de entrada y la de salida; si el filtro tiene válvula selectora multipuerto, declararlo para que esta corrección pase a cero'
        };
        if(dpf.dz){
          filasX.push(['Corrección por diferencia de cota entre los dos manómetros',
            '+'+dpf.dz.toFixed(2)+' m c.a. ('+(DZ_ORIGEN[dpf.dzFuente]||'')+'). El manómetro lee presión estática local, no energía: el que quede más abajo gana una columna de agua de '+
            'aproximadamente 1.4 PSI por metro que no es pérdida del filtro. Con el de entrada más alto que el de salida, la pérdida real del lecho es '+
            dpf.mca.toFixed(2)+' m c.a. (bruto) + '+dpf.dz.toFixed(2)+' m = '+dpf.mcaNeto.toFixed(2)+' m c.a.']);
        } else if(dpf.dzFuente==='selectora'){
          filasX.push(['Diferencia de cota entre los dos manómetros',
            'Nula: '+DZ_ORIGEN.selectora+'. El ΔP bruto se usa sin corrección hidrostática.']);
        }
        if(!dpf.valido){
          filasX.push(['Resultado',
            'La pérdida neta del lecho sale en '+dpf.mcaNeto.toFixed(2)+' m c.a., es decir cero o negativa. En un filtro con flujo hacia adelante el lecho solo disipa energía, nunca la añade, de modo que ese valor no es físico. '+
            'La causa está en la instrumentación (manómetros descalibrados, tomas de entrada y salida intercambiadas) o en que no había flujo por el lecho al momento de leer. Una válvula de bypass parcialmente cerrada aguas abajo no produce este efecto: sube ambas lecturas por igual y reduce el caudal. '+
            'La pérdida del filtro no se incorpora a la curva del sistema.']);
        } else {
          filasX.push(['Cómo entra al cálculo',
            'La pérdida real del lecho ('+dpf.mcaNeto.toFixed(2)+' m c.a. al caudal de referencia de '+
            dpf.qRef.toFixed(1)+' m³/h, tomado del '+(dpf.fuenteQRef==='caudalímetro'?'caudalímetro en sitio':'caudal despejado del propio manómetro')+
            ') se suma a la curva del sistema escalada de forma lineal con el caudal: h_L(Q) = h_L_ref · Q / Q_ref, con la misma banda de ±30 % que el resto de las pérdidas del modelo. '+
            'Se usa una relación lineal y no cuadrática porque a la velocidad de filtración de una piscina el lecho de arena trabaja en régimen viscoso —la ecuación de Ergun da un término inercial inferior al 15 %— y un lecho colmatado se comporta como una filtración de torta, que también es lineal. '+
            'Antes de esta lectura el modelo de tramos no incorporaba el filtro y un lecho colmatado no producía ningún efecto en el caudal calculado.'+
            (dpf.mcaNeto>=6 ? ' La pérdida es alta: el lecho probablemente requiere retrolavado.' : '')]);
          if(dpf.sospechaBaja){
            filasX.push(['Advertencia sobre la lectura',
              'El filtro es de arena y la pérdida corregida del lecho es de apenas '+dpf.mcaNeto.toFixed(2)+' m c.a. La ecuación de Ergun da entre 1 y 2 m de columna de agua para un lecho limpio a caudal de filtración normal, y el medio granular ofrece resistencia en todo momento. Un valor tan bajo apunta a un caudal real muy reducido o a instrumentación poco confiable (manómetros descalibrados, aire en la línea de toma, poco flujo por el lecho), y no debe interpretarse como que el lecho está limpio. El término se incorpora igual al cálculo, pero la lectura no sirve para concluir sobre el estado del filtro.']);
          }
          filasX.push(['Relación con el cruce del manómetro',
            'No hay doble conteo: el cruce del manómetro solo acumula la pérdida del circuito hasta la ENTRADA del filtro, y este término cubre exclusivamente el salto a través del lecho.']);
        }
      }
      if(r.cruceManometro){
        filasX.push(['Caudal despejado desde el manómetro', r.cruceManometro.caudal.toFixed(2)+' m³/h']);
        filasX.push(['Desviación frente al modelo',
          (r.cruceManometro.desviacionPct>=0?'+':'')+r.cruceManometro.desviacionPct.toFixed(1)+' %. '+
          (r.cruceManometro.coherente
            ? 'Son dos estimados independientes que concuerdan, así que la confianza en el resultado es alta.'
            : r.fuenteCaudal==='manometro'
              ? 'La divergencia es grande y por lo bajo: se adoptó este caudal como el de operación. Indica una restricción real (válvula estrangulada, válvula multipuerto del filtro o bomba fuera de curva) que el modelo de tramos no ve. Verificar en campo y medir con caudalímetro.'
              : 'La divergencia es alta. Conviene revisar diámetros, longitudes, accesorios de alta pérdida, la altura del manómetro o la curva de bomba seleccionada.')]);
      }
      if(r.diseno){
        var DIS = {
          conforme: 'El circuito responde a lo entregado en obra.',
          moderada: 'Hay una degradación moderada del circuito frente al diseño.',
          severa:   'La degradación es severa, o el circuito ya no corresponde al de la ficha de obra.',
          sobre:    'Está por encima del diseño, de modo que la ficha de obra no corresponde al circuito actual.'
        };
        filasX.push(['Caudal según ficha de entrega de obra', r.diseno.caudal.toFixed(2)+' m³/h']);
        filasX.push(['Desviación del caudal actual frente al diseño',
          (r.diseno.desviacionPct>=0?'+':'')+r.diseno.desviacionPct.toFixed(1)+' %. '+(DIS[r.diseno.clase]||'')]);
        if(r.diseno.volumenImplicito!=null){
          var CONV = {
            vaso:          'Cuadra con el volumen del vaso, así que el proveedor usó el criterio normativo.',
            vasoMasTanque: 'Cuadra con el vaso más el tanque, así que el proveedor incluyó el tanque de compensación. Sus tiempos no son directamente comparables con los normativos.',
            noCuadra:      'No cuadra ni con el vaso ni con el vaso más el tanque. O la piscina fue modificada, o la ficha no corresponde a este vaso.'
          };
          filasX.push(['Volumen implícito de la ficha de obra',
            r.diseno.volumenImplicito.toFixed(1)+' m³. '+(CONV[r.diseno.convencion]||'')]);
        }
      }
      /* Contraste del área medida contra la de la ficha de diseño: una
         diferencia grande explica por sí sola cualquier desviación de
         volumen, tiempo de recirculación y aforo — hay que verla antes de
         atribuirle la brecha al circuito hidráulico. */
      if(r.areaContraste){
        var ARE = {
          normal:   'Queda dentro de la tolerancia de medición en campo.',
          criterio: 'Corresponde a una diferencia de criterio de medición: bordes, canaleta perimetral o playa.',
          alerta:   'El vaso no corresponde al de la ficha de diseño, o fue modificado. Conviene revisarlo antes de usar el aforo y el volumen.'
        };
        filasX.push(['Área medida frente a la de diseño',
          r.areaContraste.medida.toFixed(1)+' m² frente a '+r.areaContraste.diseno.toFixed(1)+' m², una desviación del '+
          (r.areaContraste.desviacionPct>=0?'+':'')+r.areaContraste.desviacionPct.toFixed(1)+' %. '+
          (ARE[r.areaContraste.clase]||'')]);
      }
      _tarjetaDatos(body, filasX, {pctEtiqueta:0.38});
      _nota(body,
        'La ficha de entrega de obra dice lo que el circuito debía entregar al ser recibido, no cómo ' +
        'está hoy. La brecha entre ambos valores muestra cómo evolucionó la instalación (equipos ' +
        'sustituidos, desgaste de bomba, tubería con incrustación) y es, por sí misma, el contenido ' +
        'del ítem CHK-078.');
    }

    if(r.tanqueNorma){
      _h2(body, 'B.6 Capacidad del tanque de compensación (Numeral 10.5)');
      var TQ = {
        cumple:   'Cumple la capacidad mínima exigida.',
        parcial:  'Cumple uno de los dos criterios. La norma los ofrece con "o", de modo que basta con cumplir uno. Se deja como observación y no como incumplimiento.',
        incumple: 'No alcanza la capacidad mínima por ninguno de los dos criterios.'
      };
      var filasT = [
        ['Capacidad instalada', r.tanqueNorma.volumen.toFixed(2)+' m³'],
        ['Exigido por volumen (10% del vaso)', r.tanqueNorma.porVolumen.toFixed(2)+' m³']
      ];
      if(r.tanqueNorma.porArea!=null){
        filasT.push(['Exigido por lámina (60 L/m²)', r.tanqueNorma.porArea.toFixed(2)+' m³']);
      }
      filasT.push(['Resultado', TQ[r.tanqueNorma.clase] || '—']);
      _tarjetaDatos(body, filasT, {mono:true, pctEtiqueta:0.38});
    }

    if(r.npsh){
      _h2(body, 'B.7 Riesgo de cavitación (NPSH)');
      var filasN = [
        ['NPSH disponible', r.npsh.npsha.toFixed(2)+' m'],
        ['NPSH requerido (estimado)', r.npsh.npshr!=null ? r.npsh.npshr.toFixed(2)+' m' : 'Sin dato'],
        ['Margen', r.npsh.margen!=null ? r.npsh.margen.toFixed(2)+' m' : 'Sin dato'],
        ['Altitud de cálculo', r.npsh.altitud+' m.s.n.m.'+(r.npsh.altitudAsumida?' (asumida)':' (según sede)')],
        ['Temperatura del agua', r.npsh.temp.toFixed(1)+' °C'+(r.npsh.tempAsumida?' (asumida por vaso climatizado)':'')],
        ['Sumergencia disponible', r.npsh.sumergencia.toFixed(2)+' m']
      ];
      _tarjetaDatos(body, filasN, {mono:true, pctEtiqueta:0.38});
      _nota(body,
        'La succión inundada desde el tanque de compensación es la configuración favorable. El ' +
        'chequeo importa sobre todo en el caso degradado: si la canastilla del prefiltro se colmata, ' +
        'la pérdida en succión sube y el margen se consume. El NPSH requerido se estima por velocidad ' +
        'específica de succión y no es un dato de fábrica del equipo.');
    }
  }

  _nota(body,
    'Límites normativos de referencia: succión ≤ 1.8 m/s y descarga ≤ 2.4 m/s (Numeral 10.1), ' +
    'filtración entre 20 y 40 m³/h/m² (50 en uso restringido, Numeral 10.2), y tiempo de ' +
    'recirculación según la Tabla No. 1 del Anexo Técnico.');
}

function _anexoDimensionamiento(body, ficha){
  body.appendPageBreak();
  _h1(body, '11. Anexo técnico C: dimensionamiento normativo');
  var r = ficha.aforoResultado;
  if(!r || r.error){
    _cuerpo(body, r && r.error ? ('Sin resultado válido: '+r.error) : 'No se calculó el aforo ni la dotación sanitaria para este vaso en la PWA.');
    return;
  }
  _h2(body, 'C.1 Aforo máximo');
  _tarjetaDatos(body, [
    ['Área del espejo de agua', r.area.toFixed(1)+' m²'],
    ['Factor de uso aplicado (Tabla 4)', r.factorUso+' m²/bañista'+
      (r.criterioProf ? ' (criterio: '+String(r.criterioProf)+')' : '')],
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
    var encontrado = (d.encontrado==null) ? 'Sin dato' : String(d.encontrado);
    var falta = (d.encontrado!=null && d.encontrado < d.requerido);
    var faltantes = falta ? (d.requerido-d.encontrado) : 0;
    var estadoTxt = (d.encontrado==null) ? 'Pendiente de conteo'
      : (falta ? (faltantes===1 ? 'Falta 1' : 'Faltan '+faltantes) : 'Cumple');
    datos.push([etiquetas[k], d.ratio, String(d.requerido), encontrado, estadoTxt]);
    estados.push(d.encontrado==null ? 'sin dato' : (falta ? 'falta' : 'ok'));
  });
  var t = body.appendTable(datos);
  _estiloTabla(t, true, [125, 130, 80, 85, 85]);
  for(var i=1;i<t.getNumRows();i++){
    var col = estados[i]==='ok' ? C_OK : (estados[i]==='falta' ? C_CRITICO : C_AMBAR);
    t.getRow(i).getCell(4).getChild(0).asParagraph().setBold(true).setForegroundColor(col);
  }
}

function _anexoMemoriaCalculo(body, ficha){
  body.appendPageBreak();
  _h1(body, '12. Anexo técnico D: memoria de cálculo');
  _nota(body,
    'Estas estimaciones apoyan la decisión del inspector, pero no reemplazan la medición directa ' +
    'con caudalímetro ni el conteo físico de la dotación sanitaria. Con los supuestos que se ' +
    'declaran a continuación se puede reproducir o auditar cada resultado de los Anexos B y C.');

  _h2(body, 'D.1 Datos de la bomba');
  var filasBomba = [
    ['Potencia', ficha.bombaHP!=null ? ficha.bombaHP+' HP' : '—'],
    ['Frecuencia de operación', ficha.bombaHz!=null ? ficha.bombaHz+' Hz' : '60 Hz (valor por defecto)'],
    ['Origen de la curva', BOMBA_FAMILIA_LABEL[ficha.bombaFamilia] || 'Sin dato']
  ];
  if(ficha.bombasParalelo!=null && ficha.bombasParalelo>1){
    filasBomba.push(['Bombas iguales en paralelo',
      ficha.bombasParalelo+' — cada una trasiega ~1/'+ficha.bombasParalelo+' del caudal total; la curva del conjunto es H(Q)=H_una(Q/'+ficha.bombasParalelo+')']);
  }
  var mapaVS = BOMBA_VELOCIDAD_POR_FAMILIA[ficha.bombaFamilia];
  if(mapaVS){
    filasBomba.push(['Velocidad configurada', mapaVS.labels[ficha[mapaVS.campo]] || 'Sin dato']);
    if(ficha.bombaRPMObservada!=null && ficha.bombaRPMObservada>0){
      filasBomba.push(['RPM real leída en el variador',
        ficha.bombaRPMObservada+' RPM — la curva de fábrica se escaló por afinidad (Q∝N, H∝N²) a esta velocidad, no a la RPM nominal de la preselección']);
    }
  }
  if(ficha.bombaFamilia==='manual'){
    filasBomba.push(['Frecuencia de medición de la curva manual', ficha.curvaManualHz!=null ? ficha.curvaManualHz+' Hz' : 'Sin dato: se asume igual a la de operación']);
  }
  if(ficha.bombaFamilia==='manual' && ficha.curvaManual && ficha.curvaManual.length){
    var puntos = ficha.curvaManual
      .filter(function(p){ return p && p.q!=null && p.h!=null; })
      .map(function(p){ return 'Q='+p.q+' m³/h, H='+p.h+' m'; })
      .join(' / ');
    filasBomba.push(['Puntos de la curva manual', puntos || '—']);
  }
  _tarjetaDatos(body, filasBomba, {mono:true, pctEtiqueta:0.45});

  _h2(body, 'D.2 Datos de tubería y filtro');
  var tramosValidos = (ficha.tuberiaTramos||[]).filter(function(t){ return t && t.diametro>0; });
  if(tramosValidos.length){
    var filasTramos = [['Tramo','Ø (pulg)','Lado','Long. (m)','Acces. normales','Acces. alta pérdida','Líneas']];
    tramosValidos.forEach(function(t){
      filasTramos.push([
        t.nombre || '(sin nombre)',
        String(t.diametro),
        t.lado==='succion' ? 'Succión' : 'Presión',
        t.longitud!=null ? String(t.longitud) : '—',
        t.accesorios!=null ? String(t.accesorios) : '—',
        t.accesoriosAlta!=null ? String(t.accesoriosAlta) : '—',
        String(t.nLineas>0 ? t.nLineas : 1)
      ]);
    });
    var tTramos = body.appendTable(filasTramos);
    _estiloTabla(tTramos, true, [110, 52, 58, 60, 78, 85, 62]);
  } else {
    _tarjetaDatos(body, [
      ['Diámetro de succión', ficha.tuberiaSuccionDiam!=null ? ficha.tuberiaSuccionDiam+' pulg' : '—'],
      ['Diámetro de descarga/retorno', ficha.tuberiaDescargaDiam!=null ? ficha.tuberiaDescargaDiam+' pulg' : '—'],
      ['Longitud total de tubería', ficha.tuberiaLongitud!=null ? ficha.tuberiaLongitud+' m' : '—'],
      ['Número de accesorios', ficha.tuberiaAccesorios!=null ? String(ficha.tuberiaAccesorios) : '—'],
      ['Reparto succión y descarga', ficha.tuberiaPctSuccion!=null ? ficha.tuberiaPctSuccion+'% y '+(100-ficha.tuberiaPctSuccion)+'%' : '50% y 50% (por defecto)']
    ], {mono:true, pctEtiqueta:0.45});
  }
  body.appendParagraph('').setFontSize(4);
  var filasCircuito = [];
  if(ficha.tanqueEquilibrio){
    filasCircuito.push(['Tanque de compensación', TANQUE_LABEL[ficha.tanqueEquilibrio] || '—']);
  }
  if(ficha.tanqueEquilibrio==='si'){
    filasCircuito.push(['Dimensiones del tanque',
      (ficha.tanqueLargo!=null && ficha.tanqueAncho!=null && ficha.tanqueNivel!=null)
        ? ficha.tanqueLargo+' × '+ficha.tanqueAncho+' m, con lámina de '+ficha.tanqueNivel+' m'
        : 'Sin dato']);
    filasCircuito.push(['Desnivel entre la lámina de la piscina y la del tanque',
      ficha.desnivelPiscinaTanque!=null ? ficha.desnivelPiscinaTanque+' m' : 'Sin dato']);
  } else {
    filasCircuito.push(['Desnivel entre succión y descarga', ficha.desnivelSuccionDescarga!=null ? ficha.desnivelSuccionDescarga+' m' : '—']);
  }
  if(ficha.tipoRetorno)    filasCircuito.push(['Tipo de retorno al vaso', RETORNO_LABEL[ficha.tipoRetorno] || '—']);
  if(ficha.estadoTuberia)  filasCircuito.push(['Estado de la tubería', ESTADO_TUBERIA_LABEL[ficha.estadoTuberia] || '—']);
  if(ficha.calentadorN!=null){
    filasCircuito.push(['Número de calentadores', String(ficha.calentadorN)]);
    filasCircuito.push(['Tipo de calentador', CALENTADOR_TIPO_LABEL[ficha.calentadorTipo] || 'Sin declarar']);
    filasCircuito.push(['Arreglo de calentadores', CALENTADOR_ARREGLO_LABEL[ficha.calentadorArreglo] || 'Sin declarar']);
    // Fichas anteriores al motor v2.1 no traen `calentadorAcople`: el motor
    // asume 'linea' en ese caso, y el informe lo dice en vez de callarlo.
    filasCircuito.push(['Acople al circuito principal',
      CALENTADOR_ACOPLE_LABEL[ficha.calentadorAcople] ||
      (CALENTADOR_ACOPLE_LABEL.linea + '. No se declaró en la ficha, así que el modelo lo asume.')]);
    if(ficha.presionPostCalentador!=null){
      filasCircuito.push(['Presión aguas abajo del bloque de calentadores', ficha.presionPostCalentador+' PSI']);
    }
  }
  if(ficha.tempAgua!=null) filasCircuito.push(['Temperatura del agua', ficha.tempAgua+' °C']);
  filasCircuito.push(['Tipo de filtro', FILTRO_TIPO_LABEL[ficha.filtroTipo] || '—']);
  filasCircuito.push(['Área filtrante', ficha.filtroArea!=null ? ficha.filtroArea+' m²' : '—']);
  if(ficha.filtroValvula!=null){
    filasCircuito.push(['Conexión del filtro al circuito', FILTRO_VALVULA_LABEL[ficha.filtroValvula] || String(ficha.filtroValvula)]);
  }
  filasCircuito.push(['Presión de manómetro del filtro (entrada)', ficha.presionManometro!=null ? ficha.presionManometro+' PSI' : '—']);
  if(ficha.presionSalidaFiltro!=null){
    filasCircuito.push(['Presión de manómetro a la salida del filtro', ficha.presionSalidaFiltro+' PSI']);
  }
  if(ficha.desnivelManometrosFiltro!=null){
    filasCircuito.push(['Diferencia de cota entre los manómetros del filtro',
      ficha.desnivelManometrosFiltro+' m ('+(ficha.desnivelManometrosFiltro>0?'entrada más arriba que salida':'salida más arriba que entrada')+')']);
  }
  if(ficha.alturaManometro!=null){
    filasCircuito.push(['Altura del manómetro sobre la lámina', ficha.alturaManometro+' m']);
  }
  if(ficha.caudalMedido!=null) filasCircuito.push(['Caudal medido con caudalímetro', ficha.caudalMedido+' m³/h']);
  if(ficha.caudalDiseno!=null) filasCircuito.push(['Caudal según ficha de entrega de obra', ficha.caudalDiseno+' m³/h']);
  if(ficha.tiempoRecircDiseno!=null) filasCircuito.push(['Tiempo de recirculación de diseño', ficha.tiempoRecircDiseno+' h']);
  if(ficha.areaDiseno!=null) filasCircuito.push(['Área del espejo según ficha de diseño', ficha.areaDiseno+' m²']);
  _tarjetaDatos(body, filasCircuito, {mono:true, pctEtiqueta:0.45});

  var sup = (ficha.motorResultado && ficha.motorResultado.supuestos) || null;
  if(sup){
    _h2(body, 'D.3 Supuestos del cálculo hidráulico');
    var esV2mem = (ficha.motorResultado.modelo === 'v2');
    if(esV2mem){
      _tarjetaDatos(body, [
        ['Modelo de cálculo', 'v2'],
        ['Coeficiente de Hazen-Williams (C)', String(sup.C)+' (estado de la tubería: '+String(sup.estadoTuberia)+')'],
        ['Bombas en paralelo', (ficha.bombasParalelo>1)
          ? (sup.bombas ? String(sup.bombas) : String(ficha.bombasParalelo)+' bombas iguales en paralelo')+' (Karassik et al., 2008).'
          : 'Una sola bomba sobre el circuito'],
        ['Pérdida por accesorios', String(sup.kAccesorios)+' (Crane Co., 2022).'],
        ['Carga estática', String(sup.Hgeo)],
        ['Retorno al vaso', String(sup.retorno)],
        ['Calentadores', String(sup.calentadores)],
        ['Uso de la lectura del manómetro', String(sup.manometro)],
        ['Pérdida del filtro', sup.filtro ? String(sup.filtro) : 'Sin doble lectura del manómetro (entrada y salida): la pérdida del filtro no entra a la curva del sistema.'],
        ['Altitud para el cálculo de NPSH', String(sup.altitud)],
        ['NPSH requerido', String(sup.npshr)+' (Hydraulic Institute, 2017).'],
        ['Volumen del vaso', String(sup.volumen)],
        ['Banda de incertidumbre', String(sup.banda)]
      ], {mono:true, pctEtiqueta:0.38});
      _nota(body,
        'El modelo v2 corrige tres puntos del v1. Con tanque de compensación, la carga estática es ' +
        'solo la diferencia entre las dos láminas de agua, porque el recorrido de la tubería por los ' +
        'distintos niveles del cuarto de máquinas se cancela al tratarse de conducto cerrado y lleno. ' +
        'La lectura del manómetro ya no se suma a la curva del sistema, porque hacerlo contaba dos ' +
        'veces el circuito aguas abajo. Y los accesorios pasan de una longitud equivalente única a ' +
        'coeficientes K diferenciados por tipo. La revisión v2.1 agrega las curvas de fábrica de las ' +
        'familias EQ Series, IntelliFlo VS+SVRS, WhisperFlo VST y WhisperFloXF, corrige el punto de ' +
        'medida del manómetro con su cota respecto a la lámina, y modela el acople del calentador: ' +
        'una bomba inyectora propia no carga la línea principal, un bypass la carga con techo y un ' +
        'montaje en línea directa la carga sin techo. También admite varias bombas iguales ' +
        'operando en paralelo sobre un mismo circuito: a igual cabezal los caudales se suman, ' +
        'así que la curva del conjunto de N bombas es la de una sola evaluada en Q/N. Como la ' +
        'curva del sistema crece con el cuadrado del caudal, duplicar bombas no duplica el ' +
        'caudal: el punto de operación sube a más cabezal y menos caudal por bomba. ' +
        'Cuando el filtro tiene manómetro de entrada y de salida, el salto de presión medido ' +
        'entre ambos se suma a la curva del sistema escalado de forma lineal con el caudal ' +
        '(h_L(Q) = h_L_ref · Q / Q_ref), de modo que el estado real del lecho —limpio o ' +
        'colmatado— queda reflejado en el caudal calculado en lugar de quedar por fuera del ' +
        'modelo. Se usa una relación lineal y no cuadrática porque a la velocidad de filtración ' +
        'de una piscina el lecho de arena trabaja en régimen viscoso (la ecuación de Ergun deja ' +
        'el término inercial por debajo del 15 %) y un lecho colmatado se comporta como una ' +
        'filtración de torta, que también es lineal. Ese término no se solapa con la ' +
        'verificación por manómetro, que solo lleva la pérdida hasta la entrada del filtro. El ' +
        'manómetro mide presión estática local, no energía: si los dos instrumentos están a ' +
        'distinta altura, el que quede más abajo registra una columna hidrostática de cerca de ' +
        '1.4 PSI por metro que no corresponde al filtro, y por eso el ΔP bruto se corrige con la ' +
        'diferencia de cota antes de usarlo. Una vez hecha esa corrección, la pérdida a través del ' +
        'lecho tiene que ser positiva —el medio filtrante disipa energía, no la aporta—; si aun ' +
        'así sale cero o negativa, la causa está en manómetros descalibrados, en tomas de entrada ' +
        'y salida intercambiadas o en que no había flujo por el lecho al momento de leer, y el ' +
        'modelo descarta la lectura en lugar de introducir una pérdida negativa.');
    } else {
      _tarjetaDatos(body, [
        ['Modelo de cálculo', 'v1 (legado)'],
        ['Coeficiente de Hazen-Williams (C)', String(sup.C)],
        ['Longitud equivalente de accesorios', String(sup.LeqPorAccesorio)],
        ['Carga estática asumida', String(sup.Hgeo)],
        ['Reparto de tubería succión/descarga', String(sup.repartoLongitud)],
        ['Conversión de presión', String(sup.conversionPsi)]
      ], {mono:true, pctEtiqueta:0.45});
    }
  }

  if(ficha.aforoResultado && !ficha.aforoResultado.error){
    _h2(body, 'D.4 Supuesto del aforo y dotación');
    _nota(body,
      'La fila de la Tabla No. 4 (factor de uso) se determina con la profundidad MÁXIMA del vaso, ' +
      'que es el criterio conservador. La dotación sanitaria requerida (Tabla No. 5) se calcula ' +
      'sobre el 100% del aforo para cada elemento, aunque en campo puede repartirse por género ' +
      'según la composición real de bañistas.');
  }
}

/* ---------- Referencias (formato APA, 7.ª edición) ----------
   Toda fuente que el informe cita por nombre en el cuerpo del texto queda
   aquí con su ficha completa. La Resolución 929 se cita siempre, porque
   sostiene el objetivo y el % de cumplimiento de cualquier informe; la 234
   también, porque el tablero ejecutivo la nombra en la fila "Fuera del
   alcance de la Res. 929", que sale en todo informe, tenga o no ítems
   fuera de alcance. Crane Co. y el Hydraulic Institute solo entran cuando
   el Anexo B/D del modelo v2 llegó a citarlos (`incluirTecnicas`): son la
   fuente del método K por accesorio y del criterio de NPSH, y no aparecen
   en ningún otro lugar del informe. `numero` es el consecutivo de sección:
   13 si el informe trae Anexos A-D, 9 si no los trae (sin ficha técnica no
   hay secciones 9-12 que numerar). */
function _referencias(body, incluirTecnicas, numero){
  body.appendPageBreak();
  _h1(body, numero+'. Referencias');
  var refs = [
    'Ministerio de Salud y Protección Social. (2026, 12 de mayo). Resolución 0929 de 2026, ' +
      'por la cual se adoptan los criterios técnicos constructivos y de seguridad para los ' +
      'establecimientos e inmuebles con piscinas y estructuras similares, y se dictan otras ' +
      'disposiciones. Diario Oficial. https://www.minsalud.gov.co/sites/rid/Lists/' +
      'BibliotecaDigital/RIDE/DE/DIJ/resolucion-0929-de-2026.pdf',
    'Ministerio de Salud y Protección Social. (2026, febrero). Resolución 234 de 2026, ' +
      'por la cual se establecen los parámetros de calidad del agua y las condiciones ' +
      'sanitarias para piscinas y estructuras similares. Diario Oficial.'
  ];
  if(incluirTecnicas){
    refs.push('Crane Co. (2022). Flow of fluids through valves, fittings, and pipe (Technical Paper No. 410).');
    refs.push('Hydraulic Institute. (2017). Rotodynamic pumps: Guideline for NPSH margin (ANSI/HI 9.6.1-2017).');
    refs.push('Karassik, I. J., Messina, J. P., Cooper, P., & Heald, C. C. (2008). Pump handbook (4.ª ed.). McGraw-Hill.');
  }
  // Orden alfabético por autor, como pide APA, y con sangría francesa: la
  // segunda línea de cada referencia entra 0.5" respecto a la primera.
  refs.sort();
  refs.forEach(function(r){
    body.appendParagraph(r)
        .setFontFamily(FUENTE).setFontSize(11).setBold(false).setItalic(false)
        .setForegroundColor(C_TITULO)
        .setIndentFirstLine(0).setIndentStart(36)
        .setLineSpacing(1.5).setSpacingAfter(10);
  });
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

function _conclusion(body, m, sede, piscina, area){
  _h1(body, '6. Conclusión');
  var txt = area
    ? ('Los ítems a cargo del área de ' + area + ' en el vaso ' + piscina + ' de la sede ' + sede +
       ' presentan un cumplimiento del ' + m.pctCumplimiento + '% frente a los criterios de la ' +
       'Resolución 929 de 2026, con ')
    : ('El vaso ' + piscina + ' de la sede ' + sede + ' presenta un cumplimiento global del ' +
    m.pctCumplimiento + '% frente a los criterios de la Resolución 929 de 2026, con ');
  txt += m.noCumple + ' ítem(s) en estado No cumple y ' + (m.critico + m.alto) +
    ' hallazgo(s) clasificados en riesgo crítico o alto. ';
  if(m.vencidos>0) txt += 'Se registran ' + m.vencidos + ' hallazgo(s) con la fecha compromiso vencida, ' +
    'la desviación de gestión más relevante del período. ';
  if(m.sinFecha>0) txt += m.sinFecha + ' hallazgo(s) permanecen sin fecha compromiso asignada. ';
  txt += 'La condición general se califica como: ' + _semaforo(m.pctCumplimiento) + '.';
  // Texto corrido en color normal: el rojo se reserva para estados y niveles
  // de riesgo puntuales, no para párrafos enteros.
  _cuerpo(body, txt);
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
      celda.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(8).setPaddingRight(8);
      celda.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      var p = celda.getChild(0).asParagraph();
      p.setFontSize(esEncabezado ? 8.5 : 9).setForegroundColor(esEncabezado ? C_ENCABEZADO : C_TITULO)
       .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
       .setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
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
  var colorOk = _colorPctFill(pct);
  var dt = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Estado')
    .addColumn(Charts.ColumnType.NUMBER, 'Valor')
    .addRow(['Cumplimiento', pct])
    .addRow(['Restante', Math.max(0, 100-pct)])
    .build();
  var chart = Charts.newPieChart()
    .setDataTable(dt)
    // Se renderiza a 520 px para insertarse a 130 pt: 4x de sobremuestreo,
    // que es lo que hace falta para que el arco no salga dentado en el PDF
    // impreso. A 280 px (2.1x) el borde de la dona se veia escalonado.
    .setDimensions(520, 520)
    .setColors([colorOk, C_CELDA])
    .setLegendPosition(Charts.Position.NONE)
    .setOption('pieHole', 0.74)
    .setOption('pieSliceText', 'none')
    .setOption('backgroundColor', 'transparent')
    .setOption('tooltip', {trigger:'none'})
    .setOption('chartArea', {left:8, top:8, width:'94%', height:'94%'})
    .build();
  return chart.getBlob();
}

/* ---------- Graficas vectoriales compuestas con tablas ----------
   El servicio Charts devuelve un PNG: en el PDF se ve pixelado, trae la
   tipografia y los margenes de Google Charts -no los del informe- y no hay
   forma de alinearlo con las tablas de al lado. Una barra, sin embargo, es
   geometria pura: se compone con celdas de tabla de ancho proporcional y
   fondo de color, y asi sale VECTORIAL - nitida a cualquier zoom, en Arial
   como el resto del documento y encajada al punto con el ancho util A4.
   La dona del cumplimiento global sigue siendo imagen porque un arco no se
   puede componer con celdas rectangulares; para esa si se renderiza a 4x el
   tamano de insercion para que no se vea pixelada.

   `segmentos` es [{valor, color}]; los de valor 0 se omiten para que no
   dejen una hendidura de 3 pt de color en la barra. Lleva el mismo hairline
   gris de 0.75 pt que las tarjetas y tablas del resto del informe: sin
   borde, un bloque de color sólido sobre fondo blanco se ve plano y sin
   límite propio; con él, la barra queda enmarcada como una "pista" de
   medidor y cada segmento se separa del siguiente con una línea, en vez de
   fundirse a ojo por el color. */
function _barraApilada(contenedor, segmentos, ancho, alto){
  var vivos = segmentos.filter(function(sg){ return sg.valor > 0; });
  if(!vivos.length) vivos = [{valor:1, color:C_FILL_NOAPLICA}];
  var total = 0;
  vivos.forEach(function(sg){ total += sg.valor; });

  // El ultimo segmento absorbe el redondeo para que la suma de anchos de
  // exactamente `ancho` y la barra no quede corta ni desborde la celda.
  var anchos = [], acum = 0;
  vivos.forEach(function(sg, i){
    var w = (i === vivos.length-1)
      ? Math.max(3, ancho - acum)
      : Math.max(3, Math.round(ancho * sg.valor / total));
    anchos.push(w); acum += w;
  });

  var t = contenedor.appendTable([vivos.map(function(){ return ''; })]);
  t.setBorderWidth(0.75).setBorderColor(C_CELDA);
  var arriba = Math.round(alto/2);
  for(var i=0;i<vivos.length;i++){
    var celda = t.getCell(0,i);
    celda.setWidth(anchos[i]).setBackgroundColor(vivos[i].color);
    celda.setPaddingLeft(0).setPaddingRight(0)
         .setPaddingTop(arriba).setPaddingBottom(alto-arriba);
    celda.getChild(0).asParagraph().setFontSize(1).setSpacingBefore(0).setSpacingAfter(0);
  }
  return t;
}

/* Leyenda en una fila: cuadro de color + etiqueta + valor, uno por columna.
   Sustituye a la leyenda lateral de Google Charts, que forzaba a reservarle
   casi el 40% del ancho a la grafica y salia en otra tipografia. */
function _chipsLeyenda(contenedor, items, ancho){
  var t = contenedor.appendTable([items.map(function(){ return ''; })]);
  t.setBorderWidth(0).setBorderColor('#FFFFFF');
  var w = Math.floor(ancho/items.length);
  items.forEach(function(it, i){
    var celda = t.getCell(0,i);
    celda.setWidth(w);
    celda.setPaddingTop(5).setPaddingBottom(0).setPaddingLeft(i===0?0:6).setPaddingRight(0);
    var p = celda.getChild(0).asParagraph();
    // setText() no devuelve el párrafo para encadenar (a diferencia de
    // setFontSize/setBold/etc.) — hay que cortar la cadena aquí o el
    // siguiente método se llama sobre null y truena en tiempo real.
    p.setText('■ '+it.etiqueta+'  '+it.valor);
    p.setFontSize(7.5).setBold(true).setForegroundColor(C_TITULO)
     .setSpacingBefore(0).setSpacingAfter(0);
    // El cuadrito toma el color de la serie; el texto se queda en gris para
    // que la leyenda no compita visualmente con la barra.
    p.editAsText().setForegroundColor(0, 0, it.color);
  });
  return t;
}

/* Deja en tamano minimo los parrafos sueltos que Docs obliga a mantener
   alrededor de una tabla anidada dentro de una celda; sin esto cada barra
   arrastra dos renglones vacios de 11 pt y la grafica se estira al doble. */
function _comprimirParrafos(celda){
  for(var k=0;k<celda.getNumChildren();k++){
    var ch = celda.getChild(k);
    if(ch.getType()===DocumentApp.ElementType.PARAGRAPH){
      ch.asParagraph().setFontSize(1).setSpacingBefore(0).setSpacingAfter(0);
    }
  }
}

/* Color de RELLENO de un porcentaje de cumplimiento. Va aparte de
   _colorPct() porque aquel devuelve colores pensados para TEXTO (mas
   oscuros y desaturados); usados como area de color se ven sucios. */
function _colorPctFill(pct){
  return pct>=85 ? C_FILL_CUMPLE : (pct>=70 ? C_FILL_PROCESO : C_FILL_NOCUMPLE);
}

/* Barras horizontales de cumplimiento por capitulo, peor primero.
   Reemplaza al Charts.newBarChart() apilado: al ser vectorial caben
   etiquetas de capitulo mucho mas largas (el chart las recortaba a 32
   caracteres para que no desbordaran el eje), el % va en la misma linea de
   su barra en vez de tener que leerlo contra una escala, y cada barra toma
   el color de su semaforo en vez de ser todas azules. */
function _barrasCapitulos(body, m){
  var caps = Object.keys(m.porCapitulo).map(function(cap){
    var c = m.porCapitulo[cap];
    var base = c.total - c.fueraDelDenominador;
    var pct = base>0 ? Math.round(100*c.cumpleEnAlcance/base) : 0;
    return {etiqueta: cap.length>46 ? cap.slice(0,44)+'…' : cap, pct:pct};
  }).sort(function(a,b){ return a.pct - b.pct; });
  if(!caps.length) return null;

  var W_LBL=210, W_BAR=232, W_PCT=63;   // suman 505 pt = ANCHO_UTIL
  var t = body.appendTable(caps.map(function(){ return ['','','']; }));
  t.setBorderWidth(0).setBorderColor('#FFFFFF');

  caps.forEach(function(c, i){
    var fila = t.getRow(i);
    var cLbl = fila.getCell(0), cBar = fila.getCell(1), cPct = fila.getCell(2);
    cLbl.setWidth(W_LBL).setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(0).setPaddingRight(12);
    cBar.setWidth(W_BAR).setPaddingTop(5).setPaddingBottom(3).setPaddingLeft(0).setPaddingRight(0);
    cPct.setWidth(W_PCT).setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(8).setPaddingRight(0);

    var pLbl = cLbl.getChild(0).asParagraph();
    pLbl.setText(c.etiqueta);
    pLbl.setFontSize(8).setBold(false).setForegroundColor(C_ENCABEZADO)
      .setSpacingBefore(0).setSpacingAfter(0);

    _barraApilada(cBar, [
      {valor:c.pct,     color:_colorPctFill(c.pct)},
      {valor:100-c.pct, color:C_CELDA}
    ], W_BAR, 9);
    _comprimirParrafos(cBar);

    var pPctCap = cPct.getChild(0).asParagraph();
    pPctCap.setText(c.pct+'%');
    pPctCap.setFontFamily(FUENTE_MONO).setFontSize(8.5).setBold(true)
      .setForegroundColor(_colorPct(c.pct))
      .setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
      .setSpacingBefore(0).setSpacingAfter(0);
  });
  return t;
}

/* Tarjeta de indicador para el tablero ejecutivo: rotulo pequeno en gris,
   cifra grande en mono y una linea de lectura debajo. Es el "stat card" del
   mockup, que en la version anterior estaba aplanado como una fila mas de
   una tabla de 8 filas donde ninguna cifra resaltaba. */
function _tarjetaKPI(celda, etiqueta, valor, lectura, color){
  var c = _tarjetaEnCelda(celda);
  c.setPaddingTop(8).setPaddingBottom(8).setPaddingLeft(9).setPaddingRight(9);
  var pEtq = c.getChild(0).asParagraph();
  pEtq.setText(String(etiqueta).toUpperCase());
  pEtq.setFontSize(6.5).setBold(true).setForegroundColor(C_TITULO)
    .setSpacingBefore(0).setSpacingAfter(1);
  c.appendParagraph(String(valor))
    .setFontFamily(FUENTE_MONO).setFontSize(19).setBold(true)
    .setForegroundColor(color || C_ENCABEZADO)
    .setSpacingBefore(0).setSpacingAfter(0);
  c.appendParagraph(String(lectura))
    .setFontSize(7).setBold(false).setForegroundColor(C_TITULO)
    .setSpacingBefore(2).setSpacingAfter(0);
  return c;
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
  var ui = SpreadsheetApp.getUi();
  // Submenús con un ítem fijo por área: Apps Script no permite pasarle
  // argumentos a un addItem, así que cada área necesita su propia función
  // envoltorio (menuVasoArea_* / menuSedeArea_*).
  var vasoPorArea = ui.createMenu('Generar informe del vaso por ÁREA');
  vasoPorArea.addItem('Todas las áreas', 'menuGenerarInformesPorArea');
  AREA_ORDEN.concat(['Sin clasificar']).forEach(function(area){
    if(area==='Compartido (los 3)') return;  // sus ítems viajan dentro de las 3 áreas técnicas
    vasoPorArea.addItem(area, 'menuVasoArea_'+_slugArea(area));
  });

  // Se generan de a un área por corrida a propósito — hacer las cuatro
  // áreas de todos los vasos de una sede grande se pasa del límite de
  // 6 minutos de ejecución de Apps Script.
  var sedePorArea = ui.createMenu('Generar informe de sede por ÁREA');
  AREA_ORDEN.concat(['Sin clasificar']).forEach(function(area){
    if(area==='Compartido (los 3)') return;
    sedePorArea.addItem(area, 'menuSedeArea_'+_slugArea(area));
  });

  ui.createMenu('Inspección piscinas')
    .addItem('Generar informe del vaso seleccionado', 'menuGenerarInforme')
    .addItem('Generar informes de toda la sede', 'menuGenerarSede')
    .addSeparator()
    .addSubMenu(vasoPorArea)
    .addSubMenu(sedePorArea)
    .addToUi();
}

/* Para cada piscina de una sede, la fecha de inspección más reciente
   registrada en la hoja. "Toda la sede" no puede depender de que todos los
   vasos compartan la misma fecha de la fila seleccionada — cada vaso se
   inspecciona en su propio día — así que cada uno se reporta con su último
   dato disponible. _fechaStr ya normaliza a 'yyyy-MM-dd', por eso alcanza
   con comparar los strings. */
function _vasosSedeUltimaFecha(sede){
  var sh = _hojaDatos();
  var last = sh.getLastRow();
  var datos = last>=2 ? sh.getRange(2,1,last-1,TOTAL_COLS).getValues() : [];
  var ultima = {};
  datos.forEach(function(f){
    if(String(f[COL.sede-1])!==sede) return;
    var pi = String(f[COL.piscina-1]);
    var fecha = _fechaStr(f[COL.fecha-1]);
    if(!ultima[pi] || fecha > ultima[pi]) ultima[pi] = fecha;
  });
  return ultima; // {piscina: 'yyyy-MM-dd'}
}

/* Genera, para TODOS los vasos de una sede (cada uno en su fecha más
   reciente), el informe de una sola área. Es el flujo de entrega: se corre
   una vez por cada interesado y se le pasa el link de su carpeta. */
function _menuSedeArea(area){
  var ui = SpreadsheetApp.getUi();
  var sh = _hojaDatos();
  var fila = SpreadsheetApp.getActiveSheet().getActiveRange().getRow();
  if(fila<2){ ui.alert('Selecciona una fila de datos, para identificar la sede.'); return; }
  var v = sh.getRange(fila,1,1,TOTAL_COLS).getValues()[0];
  var sede = String(v[COL.sede-1]);

  var ultima = _vasosSedeUltimaFecha(sede);
  var lista = Object.keys(ultima);
  if(!lista.length){ ui.alert('No hay vasos registrados para la sede '+sede+'.'); return; }

  var ok = [], sinItems = [];
  lista.forEach(function(pi){
    var r = generarInformeVaso(sede, pi, ultima[pi], area);
    if(r.ok) ok.push(pi+' ('+ultima[pi]+'): '+r.metricas.pctCumplimiento+'%'+chr10()+r.pdfUrl);
    else sinItems.push(pi+' ('+ultima[pi]+')');
  });

  var msg = 'Área: '+area+chr10()+'Sede: '+sede+chr10()+chr10();
  msg += ok.length ? ('Informes generados ('+ok.length+'):'+chr10()+chr10()+ok.join(chr10()+chr10()))
                   : 'No se generó ningún informe: ningún vaso tiene ítems de esta área.';
  if(sinItems.length) msg += chr10()+chr10()+'Sin ítems de esta área: '+sinItems.join(', ');
  msg += chr10()+chr10()+'Carpeta: '+RAIZ_DRIVE+'/'+sede+'/<vaso>/<fecha>/AREA_'+_slugArea(area)+'/';
  ui.alert(msg);
}

/* Salto de línea para los alert(): más legible que escaparlo en cada
   concatenación. */
function chr10(){ return String.fromCharCode(10); }

/* Envoltorios del submenú de sede. Si se agrega un área a AREA_ORDEN hay
   que añadir aquí su función con el nombre que produce _slugArea(), o el
   ítem del menú quedará apuntando a una función inexistente. */
function menuSedeArea_ELECTRICO(){ _menuSedeArea('Eléctrico'); }
function menuSedeArea_INFRAESTRUCTURA(){ _menuSedeArea('Infraestructura'); }
function menuSedeArea_TERMICA_E_HIDRAULICA(){ _menuSedeArea('Térmica e Hidráulica'); }
function menuSedeArea_GESTION_DE_SEDE(){ _menuSedeArea('Gestión de sede'); }
function menuSedeArea_SIN_CLASIFICAR(){ _menuSedeArea('Sin clasificar'); }

/* Informe de una sola área para el vaso de la fila seleccionada (su propio
   sede+piscina+fecha, sin ambigüedad porque viene de una fila concreta). */
function _menuVasoArea(area){
  var ui = SpreadsheetApp.getUi();
  var sh = _hojaDatos();
  var fila = SpreadsheetApp.getActiveSheet().getActiveRange().getRow();
  if(fila<2){ ui.alert('Selecciona una fila de datos.'); return; }
  var v = sh.getRange(fila,1,1,TOTAL_COLS).getValues()[0];
  var r = generarInformeVaso(String(v[COL.sede-1]), String(v[COL.piscina-1]), _fechaStr(v[COL.fecha-1]), area);
  ui.alert(r.ok ? ('Informe generado ('+area+', '+r.metricas.pctCumplimiento+'%):'+chr10()+r.pdfUrl) : 'Error: '+r.error);
}

/* Envoltorios del submenú de vaso. Mismo criterio que menuSedeArea_*. */
function menuVasoArea_ELECTRICO(){ _menuVasoArea('Eléctrico'); }
function menuVasoArea_INFRAESTRUCTURA(){ _menuVasoArea('Infraestructura'); }
function menuVasoArea_TERMICA_E_HIDRAULICA(){ _menuVasoArea('Térmica e Hidráulica'); }
function menuVasoArea_GESTION_DE_SEDE(){ _menuVasoArea('Gestión de sede'); }
function menuVasoArea_SIN_CLASIFICAR(){ _menuVasoArea('Sin clasificar'); }

function menuGenerarInforme(){
  var sh = _hojaDatos();
  var fila = SpreadsheetApp.getActiveSheet().getActiveRange().getRow();
  if(fila<2){ SpreadsheetApp.getUi().alert('Selecciona una fila de datos.'); return; }
  var v = sh.getRange(fila,1,1,TOTAL_COLS).getValues()[0];
  var res = generarInformeVaso(String(v[COL.sede-1]), String(v[COL.piscina-1]), _fechaStr(v[COL.fecha-1]));
  SpreadsheetApp.getUi().alert(res.ok ? 'Informe generado:\n'+res.pdfUrl : 'Error: '+res.error);
}

/* Un informe por cada área responsable del vaso de la fila seleccionada
   ("Todas las áreas" del submenú "Generar informe del vaso por ÁREA").
   Es lo que se le entrega a cada equipo (eléctrico, T&H, infraestructura,
   gestión de sede) en vez del informe completo del vaso. */
function menuGenerarInformesPorArea(){
  var sh = _hojaDatos();
  var fila = SpreadsheetApp.getActiveSheet().getActiveRange().getRow();
  if(fila<2){ SpreadsheetApp.getUi().alert('Selecciona una fila de datos.'); return; }
  var v = sh.getRange(fila,1,1,TOTAL_COLS).getValues()[0];
  var res = generarInformesPorArea(String(v[COL.sede-1]), String(v[COL.piscina-1]), _fechaStr(v[COL.fecha-1]));
  if(!res.length){ SpreadsheetApp.getUi().alert('No se generó ningún informe: el vaso no tiene ítems clasificados por área.'); return; }
  var detalle = res.map(function(r){ return r.area+': '+r.metricas.pctCumplimiento+'%\n'+r.pdfUrl; }).join('\n\n');
  SpreadsheetApp.getUi().alert('Informes por área generados ('+res.length+'):\n\n'+detalle);
}

/* Genera el informe completo (sin segmentar por área) de todos los vasos de
   la sede, cada uno en su fecha de inspección más reciente. */
function menuGenerarSede(){
  var sh = _hojaDatos();
  var fila = SpreadsheetApp.getActiveSheet().getActiveRange().getRow();
  if(fila<2){ SpreadsheetApp.getUi().alert('Selecciona una fila de datos, para identificar la sede.'); return; }
  var v = sh.getRange(fila,1,1,TOTAL_COLS).getValues()[0];
  var sede = String(v[COL.sede-1]);
  var ultima = _vasosSedeUltimaFecha(sede);
  var lista = Object.keys(ultima);
  if(!lista.length){ SpreadsheetApp.getUi().alert('No hay vasos registrados para la sede '+sede+'.'); return; }
  var urls = [];
  lista.forEach(function(p){
    var r = generarInformeVaso(sede, p, ultima[p]);
    if(r.ok) urls.push(p+' ('+ultima[p]+'): '+r.pdfUrl);
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
      'Sin registro de pH en bitácora. Se deja consignado por control operativo interno, ya que no es un criterio exigido por la Res. 929 (ver nota del ítem).',
      'Retomar el registro diario de pH como buena práctica operativa.',
      'Operario de mantenimiento','2026-08-18','', '', '0', '', '', '', ''],

    ['CHK-233','17. Calidad del agua y operación sanitaria','Medidor de CO2','Calidad de agua','Alta',
      sede, piscina, fecha, 'Responsable Prueba', 'No cumple','Alto',
      'No hay medidor de CO2 instalado en el área cubierta.','Comprar e instalar medidor CO2.',
      'Coordinador Mantenimiento','2026-08-15','', '', '0', '', '', '', ''],

    ['CHK-078','8. Sistema de recirculación','Caudal real verificado frente al diseño.','Operación / Calidad de agua','Alta',
      sede, piscina, fecha, 'Responsable Prueba', 'No cumple','Alto',
      'Velocidad de filtración (42.8 m³/h/m²) y tiempo de recirculación (11.82 h) fuera de rango según el cálculo del motor hidráulico. Ver Anexo B.',
      'Evaluar aumento de área filtrante o bomba de mayor caudal.',
      'Coordinador Mantenimiento','2026-08-20','', '', '0', '', '', '', ''],

    ['CHK-204','24. Aforo y control de ingreso','Aforo máximo determinado por estanque.','Operación / Seguridad','Media',
      sede, piscina, fecha, 'Responsable Prueba', 'No cumple','Medio',
      'Faltan duchas frente a la dotación de la Tabla No. 5 (2 encontradas de 3 requeridas para 75 bañistas). Ver Anexo C.',
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
