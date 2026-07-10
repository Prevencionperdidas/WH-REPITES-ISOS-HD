/* ════════════════════════════════════════════════════════════════
   CONTROL DE ISOs/ASOs — Backend Google Apps Script
   IKEA Home Delivery & Logistics
   ════════════════════════════════════════════════════════════════
   INSTALACIÓN:
   1. Extensiones → Apps Script en tu planilla
   2. Pega este Code.gs completo
   3. Crea App.html    → "+" → HTML → nombre: App    → pega control-isos.html
   4. Crea Mobile.html → "+" → HTML → nombre: Mobile → pega mobile.html
   5. Implementar → Nueva implementación → Aplicación web
      · Ejecutar como: Yo  ·  Acceso: Cualquier usuario
   6. IMPORTANTE: en el Dashboard de la app, presiona
      "⚙ Configurar hojas en Sheets" para crear todas las pestañas
════════════════════════════════════════════════════════════════ */

/* ── Configuración de etapas ── */
var STAGES = [
  { key:'llegada',   label:'Llegada',          tab:'01-Llegada',        hex:'#1A5EA8' },
  { key:'recepcion', label:'Recepción L.Inv.',  tab:'02-Recepción LI',   hex:'#1269C0' },
  { key:'bodega',    label:'Bodega Rechazo',    tab:'03-Bodega Rechazo', hex:'#2978D4' },
  { key:'retiro',    label:'Retiro',            tab:'04-Retiro',         hex:'#3B87E0' },
  { key:'anden',     label:'Andén',             tab:'05-Andén',          hex:'#1F8A4C' },
  { key:'anulado',   label:'Anulados',          tab:'06-Anulados',       hex:'#D7263D' },
];
var MAIN_TAB  = 'ISOs';
var DASH_TAB  = '📊 Dashboard';

/* ── Columnas de la hoja ISOs ──────────────────────────────────
   A:ID  B:Código  C:Etapa  D:Creado  E:Actualizado  F:JSON
   G:Entrada a etapa  H:Responsable  I:Motivo  J:Transportista
──────────────────────────────────────────────────────────────── */

function getMainSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s  = ss.getSheetByName(MAIN_TAB);
  if (!s) {
    s = ss.insertSheet(MAIN_TAB);
    s.appendRow(['ID','Código','Etapa','Creado','Actualizado',
                 'JSON','Entrada a etapa','Responsable','Motivo','Transportista']);
    s.setFrozenRows(1);
    s.getRange('1:1').setFontWeight('bold')
      .setBackground('#0058A3').setFontColor('#FFFFFF');
    [1,2,3,4,5,7,8,9,10].forEach(function(c){ s.setColumnWidth(c,140); });
    s.setColumnWidth(6,500);
  }
  return s;
}

function readAll() {
  var s=getMainSheet(), last=s.getLastRow();
  if(last<2) return [];
  var rows=s.getRange(2,1,last-1,6).getValues(), out=[];
  rows.forEach(function(r){ try{ if(r[5]) out.push(JSON.parse(r[5])); }catch(e){} });
  return out;
}

function findRow(s,id) {
  var last=s.getLastRow(); if(last<2) return -1;
  var ids=s.getRange(2,1,last-1,1).getValues();
  for(var i=0;i<ids.length;i++){ if(ids[i][0]===id) return i+2; }
  return -1;
}

function saveOne(iso) {
  if(!iso||!iso.id) return;
  var lock=LockService.getScriptLock(); lock.waitLock(15000);
  try{
    var s=getMainSheet(), now=new Date().toISOString();
    var hist=iso.history||[], last=hist[hist.length-1]||{};
    var row=[
      iso.id,
      iso.code||'',
      iso.stage||'',
      iso.createdAt||now,
      iso.updatedAt||now,
      JSON.stringify(iso),
      iso.stageEnteredAt||now,
      last.responsable||'',
      (iso.data&&iso.data.llegada&&iso.data.llegada.motivoDevolucion)||'',
      (iso.data&&iso.data.llegada&&iso.data.llegada.empresaTransportista)||'',
    ];
    var idx=findRow(s,iso.id);
    if(idx===-1) s.appendRow(row);
    else s.getRange(idx,1,1,10).setValues([row]);
  }finally{ lock.releaseLock(); }
}

/* ════════════════════════════════════════════════════════════════
   FUNCIONES DE SERVIDOR — google.script.run + fetch GET
════════════════════════════════════════════════════════════════ */

function serverGetAll() {
  try{ return { ok:true, isos:readAll() }; }
  catch(e){ return { ok:false, error:e.message }; }
}

function serverSave(isoJson) {
  try{ saveOne(JSON.parse(isoJson)); return { ok:true }; }
  catch(e){ return { ok:false, error:e.message }; }
}

function serverBulk(isosJson) {
  try{
    var isos=JSON.parse(isosJson);
    isos.forEach(saveOne);
    return { ok:true, count:isos.length };
  }catch(e){ return { ok:false, error:e.message }; }
}

/* ════════════════════════════════════════════════════════════════
   CONFIGURACIÓN DE HOJAS — crea pestañas por etapa + Dashboard
════════════════════════════════════════════════════════════════ */

function serverSetup() {
  try{
    setupStageTabs();
    setupDashboard();
    return { ok:true, msg:'Hojas configuradas correctamente en Google Sheets.' };
  }catch(e){ return { ok:false, error:e.message }; }
}

function setupStageTabs() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  STAGES.forEach(function(st){
    var s=ss.getSheetByName(st.tab)||ss.insertSheet(st.tab);
    s.clearContents().clearFormats();
    // Título
    s.getRange(1,1,1,6).merge()
      .setValue('ISOs en etapa: '+st.label+' — CD El Sauce IKEA')
      .setBackground(st.hex).setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
    // Subtítulo contador dinámico
    s.getRange(2,1,1,6).merge()
      .setFormula('="Total: "&COUNTIF(ISOs!C:C,"'+st.key+'")&" ISO(s)"')
      .setBackground('#F4F8FD').setFontColor(st.hex)
      .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
    // Datos via QUERY
    var q='=IFERROR(QUERY(ISOs!A:J,"SELECT B,C,G,H,I,J WHERE C=\''+st.key+'\' AND A<>\'\' ORDER BY G ASC LABEL B \'Código\',C \'Etapa\',G \'Entrada a etapa\',H \'Responsable\',I \'Motivo\',J \'Transportista\'",0),"Sin ISOs en esta etapa")';
    s.getRange(3,1).setFormula(q);
    // Formato encabezados de columna (fila 3 tiene el LABEL)
    [150,120,170,150,200,180].forEach(function(w,i){ s.setColumnWidth(i+1,w); });
    s.setFrozenRows(3);
    // Color de pestaña
    s.setTabColor(st.hex);
  });
}

function setupDashboard() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var d=ss.getSheetByName(DASH_TAB)||ss.insertSheet(DASH_TAB,0);
  d.clearContents().clearFormats();
  d.setTabColor('#FFDB00');

  /* ── Título ── */
  d.getRange('A1:F1').merge()
    .setValue('📊  DASHBOARD — Control ISOs/ASOs · CD El Sauce IKEA')
    .setBackground('#002C52').setFontColor('#FFDB00')
    .setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  d.setRowHeight(1,48);

  d.getRange('A2:F2').merge()
    .setFormula('="Última actualización: "&TEXT(MAX(ISOs!E:E),"dd/mm/yyyy hh:mm")')
    .setBackground('#EAF1FB').setFontColor('#6B7C8F')
    .setFontSize(10).setHorizontalAlignment('center');

  /* ── Encabezados tabla resumen ── */
  var hdr=[['ETAPA','ISOs ACTIVOS','% DEL TOTAL','CRÍTICOS >48h','ADVERTENCIA >24h','ÚLTIMA ENTRADA']];
  d.getRange('A4:F4').setValues(hdr)
    .setBackground('#0058A3').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('center');
  d.setRowHeight(4,28);

  /* ── Filas por etapa ── */
  var stageKeys=[['llegada','#EAF1FB'],['recepcion','#F4F8FD'],
                  ['bodega','#EAF1FB'],['retiro','#F4F8FD'],['anden','#E7F5EC']];
  var stageLabels={
    llegada:'Llegada',recepcion:'Recepción L.Inv.',
    bodega:'Bodega Rechazo',retiro:'Retiro',anden:'Andén'
  };
  var colors={
    llegada:'#1A5EA8',recepcion:'#1269C0',
    bodega:'#2978D4',retiro:'#3B87E0',anden:'#1F8A4C'
  };

  stageKeys.forEach(function(st,i){
    var key=st[0], bg=st[1], row=5+i;
    d.getRange(row,1).setValue(stageLabels[key]).setFontColor(colors[key]).setFontWeight('bold');
    d.getRange(row,2).setFormula('=COUNTIF(ISOs!C:C,"'+key+'")').setHorizontalAlignment('center').setFontWeight('bold').setFontSize(14);
    d.getRange(row,3).setFormula('=IFERROR(TEXT(B'+row+'/SUM($B$5:$B$9),"0%"),"0%")').setHorizontalAlignment('center');
    // Críticos: en esa etapa Y entrada hace más de 48h
    d.getRange(row,4).setFormula('=IFERROR(COUNTIFS(ISOs!C:C,"'+key+'",ISOs!G:G,"<"&(NOW()-2)),"0")').setHorizontalAlignment('center').setFontColor('#D7263D').setFontWeight('bold');
    // Advertencia: en esa etapa Y entrada hace más de 24h pero menos de 48h
    d.getRange(row,5).setFormula('=IFERROR(COUNTIFS(ISOs!C:C,"'+key+'",ISOs!G:G,"<"&(NOW()-1),ISOs!G:G,">="&(NOW()-2)),"0")').setHorizontalAlignment('center').setFontColor('#B9790A').setFontWeight('bold');
    d.getRange(row,6).setFormula('=IFERROR(TEXT(MAXIFS(ISOs!G:G,ISOs!C:C,"'+key+'"),"dd/mm hh:mm"),"—")').setHorizontalAlignment('center');
    d.getRange(row,1,1,6).setBackground(bg);
    d.setRowHeight(row,30);
  });

  /* ── Totales ── */
  d.getRange('A10:F10').setBackground('#002C52').setFontColor('#FFDB00').setFontWeight('bold');
  d.getRange('A10').setValue('TOTAL ACTIVOS');
  d.getRange('B10').setFormula('=SUM(B5:B9)').setHorizontalAlignment('center').setFontSize(14);
  d.getRange('D10').setFormula('=SUM(D5:D9)').setHorizontalAlignment('center').setFontColor('#FF9FAE').setFontWeight('bold');
  d.getRange('E10').setFormula('=SUM(E5:E9)').setHorizontalAlignment('center').setFontColor('#FFDB00').setFontWeight('bold');
  d.setRowHeight(10,32);

  /* ── Sección: anulados ── */
  d.getRange('A12').setValue('Anulados total:').setFontWeight('bold').setFontColor('#D7263D');
  d.getRange('B12').setFormula('=COUNTIF(ISOs!C:C,"anulado")').setFontColor('#D7263D').setFontWeight('bold');

  d.getRange('A13').setValue('ISOs creados hoy:').setFontWeight('bold').setFontColor('#1F8A4C');
  d.getRange('B13').setFormula('=COUNTIFS(ISOs!D:D,">="&TODAY(),ISOs!D:D,"<"&TODAY()+1)').setFontColor('#1F8A4C').setFontWeight('bold');

  d.getRange('A14').setValue('Completados hoy (Andén):').setFontWeight('bold').setFontColor('#0058A3');
  d.getRange('B14').setFormula('=COUNTIFS(ISOs!C:C,"anden",ISOs!G:G,">="&TODAY())').setFontColor('#0058A3').setFontWeight('bold');

  /* ── Anchos de columna ── */
  [180,100,80,110,120,120].forEach(function(w,i){ d.setColumnWidth(i+1,w); });

  /* ── Gráfico de barras ── */
  try{
    var charts=d.getCharts(); charts.forEach(function(c){ d.removeChart(c); });
    var chart=d.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(d.getRange('A5:B9'))
      .setOption('title','ISOs por etapa')
      .setOption('colors',['#0058A3'])
      .setOption('legend',{position:'none'})
      .setOption('hAxis',{title:'Cantidad'})
      .setOption('backgroundColor','#F4F8FD')
      .setPosition(4,8,0,0)
      .setOption('width',400).setOption('height',220)
      .build();
    d.insertChart(chart);
  }catch(e){ /* gráfico opcional */ }
}

/* ════════════════════════════════════════════════════════════════
   PUNTO DE ENTRADA HTTP
════════════════════════════════════════════════════════════════ */
function doGet(e) {
  var p=(e&&e.parameter)||{}, fn=p.fn||'';
  if(!fn){
    var mob=p.m==='1';
    return HtmlService.createHtmlOutputFromFile(mob?'Mobile':'App')
      .setTitle(mob?'Scanner ISO/ASO — Bodega · Retiro · Andén':'Control de ISOs/ASOs — IKEA')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport','width=device-width,initial-scale=1');
  }
  var data=p.data;
  if(fn==='serverGetAll')  return jsonOut(serverGetAll());
  if(fn==='serverSave')    return jsonOut(serverSave(data||'{}'));
  if(fn==='serverBulk')    return jsonOut(serverBulk(data||'[]'));
  if(fn==='serverSetup')   return jsonOut(serverSetup());
  return jsonOut({ok:false,error:'Función desconocida: '+fn});
}
function doPost(e){ return doGet(e); }

function jsonOut(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
