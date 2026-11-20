/**
 * AGRÍCOLA 3P - Receptor de evaluaciones (Google Apps Script)
 * Pegar en: Google Sheet nuevo > Extensiones > Apps Script. Luego Implementar > Aplicación web
 * (Ejecutar como: Yo / Acceso: Cualquier persona) y copiar el enlace que termina en /exec.
 * Crea solo las hojas: Base, Resumen por día, Por lote, Tabla dinámica, Dashboard y Cierres.
 */
const HDR=['ID','Fecha','Hora','Tipo','Lote','Jaba','N°','Peso neto (Kg)','Prom. Brix','Mín','Máx','Peso prom. racimo','Desgrane %'];
const VERDE='#587f2e';

function salida(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function doGet(){return salida({ok:true,msg:'Agrícola 3P activo'})}
function doPost(e){
  const lock=LockService.getScriptLock();lock.waitLock(25000);
  try{
    const d=JSON.parse(e.postData.contents),ss=SpreadsheetApp.getActive();
    preparar(ss);
    if(d.accion==='sync')return salida(sync(ss,d));
    if(d.accion==='cierre')return salida(cierre(ss,d));
    return salida({ok:false,error:'Acción desconocida'});
  }catch(err){return salida({ok:false,error:String(err)})}
  finally{lock.releaseLock()}
}
function fecha(iso){const p=iso.split('-');return new Date(+p[0],+p[1]-1,+p[2],12)}
function hoja(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n)}
function cabecera(h,cols){h.getRange(1,1,1,cols.length).setValues([cols]).setFontWeight('bold').setBackground(VERDE).setFontColor('#ffffff');h.setFrozenRows(1)}

function preparar(ss){
  ss.setSpreadsheetTimeZone('America/Lima');
  const base=hoja(ss,'Base');
  if(base.getLastRow()===0){
    cabecera(base,HDR);
    base.getRange('A:A').setNumberFormat('0');
    base.getRange('B:B').setNumberFormat('dd/mm/yyyy');
    base.getRange('C:C').setNumberFormat('@');
    base.getRange('H:M').setNumberFormat('0.000');
    base.setColumnWidths(1,HDR.length,95);
  }
  const ci=hoja(ss,'Cierres');
  if(ci.getLastRow()===0){
    cabecera(ci,['Fecha','Hora de cierre','Evaluaciones','Brix prom.','Brix mín','Brix máx','Hora Brix máx','Hora Brix mín','Tendencia','Puntos clave','Clave']);
    ci.getRange('A:A').setNumberFormat('dd/mm/yyyy');
  }
}

function sync(ss,d){
  const base=ss.getSheetByName('Base'),n=base.getLastRow();
  const mapa={};
  if(n>1)base.getRange(2,1,n-1,1).getValues().forEach((r,i)=>mapa[String(r[0])]=i+2);
  const nuevas=[];let act=0;
  d.filas.forEach(f=>{
    const fila=[f.id,fecha(f.iso),f.hora,f.tipo,f.lote,f.jaba,f.n,f.neto,f.brix,f.min,f.max,f.promRac,f.des].map(v=>v===null||v===undefined?'':v);
    const r=mapa[String(f.id)];
    if(r){base.getRange(r,1,1,HDR.length).setValues([fila]);act++}else nuevas.push(fila);
  });
  if(nuevas.length)base.getRange(n+1,1,nuevas.length,HDR.length).setValues(nuevas);
  const ult=base.getLastRow();
  if(ult>2)base.getRange(2,1,ult-1,HDR.length).sort([{column:2,ascending:true},{column:3,ascending:true}]);
  tablero(ss);
  return {ok:true,nuevas:nuevas.length,actualizadas:act,total:ult-1};
}

// Hojas de análisis: se crean una sola vez, después se actualizan solas con las fórmulas
function tablero(ss){
  const base=ss.getSheetByName('Base');
  let res=ss.getSheetByName('Resumen por día');
  if(!res){
    res=ss.insertSheet('Resumen por día');
    res.getRange('A1').setFormula('=QUERY(Base!A:M,"select B,count(A),avg(I),min(I),max(I),sum(H),avg(M) where A is not null group by B order by B label B \'Fecha\',count(A) \'Evaluaciones\',avg(I) \'Brix prom.\',min(I) \'Brix mín\',max(I) \'Brix máx\',sum(H) \'Neto total (Kg)\',avg(M) \'Desgrane prom. %\'",1)');
    res.getRange('A:A').setNumberFormat('dd/mm/yyyy');res.getRange('C:G').setNumberFormat('0.00');
    res.getRange('1:1').setFontWeight('bold').setBackground(VERDE).setFontColor('#ffffff');res.setFrozenRows(1);
  }
  if(!ss.getSheetByName('Por lote')){
    const l=ss.insertSheet('Por lote');
    l.getRange('A1').setFormula('=QUERY(Base!A:M,"select E,count(A),avg(I),min(I),max(I),sum(H),avg(M) where E is not null and E <> \'\' group by E order by E label E \'Lote\',count(A) \'Evaluaciones\',avg(I) \'Brix prom.\',min(I) \'Brix mín\',max(I) \'Brix máx\',sum(H) \'Neto total (Kg)\',avg(M) \'Desgrane prom. %\'",1)');
    l.getRange('C:G').setNumberFormat('0.00');
    l.getRange('1:1').setFontWeight('bold').setBackground(VERDE).setFontColor('#ffffff');l.setFrozenRows(1);
  }
  if(!ss.getSheetByName('Tabla dinámica')){
    const pv=ss.insertSheet('Tabla dinámica');
    try{
      const pt=pv.getRange('A1').createPivotTable(base.getRange('A:M'));
      pt.addRowGroup(2);pt.addRowGroup(5);          // Fecha > Lote
      pt.addColumnGroup(4);                         // Tipo (Brix / Jaba / Racimos)
      pt.addPivotValue(1,SpreadsheetApp.PivotTableSummarizeFunction.COUNTA).setDisplayName('Evaluaciones');
    }catch(e){pv.getRange('A1').setValue('No se pudo crear la tabla dinámica: '+e)}
  }
  if(!ss.getSheetByName('Dashboard')){
    const d=ss.insertSheet('Dashboard');
    d.getRange('A1').setValue('AGRÍCOLA 3P - Dashboard de evaluaciones').setFontSize(16).setFontWeight('bold').setFontColor(VERDE);
    d.getRange('A3:B6').setFormulas([
      ['Evaluaciones','=COUNTA(Base!A2:A)'],
      ['Brix promedio','=IFERROR(AVERAGE(Base!I2:I),"")'],
      ['Último día con datos','=MAX(Base!B2:B)'],
      ['Brix prom. último día','=IFERROR(AVERAGEIFS(Base!I2:I,Base!B2:B,B5),"")']
    ]);
    d.getRange('A3:A6').setFontWeight('bold');d.getRange('B5').setNumberFormat('dd/mm/yyyy');d.getRange('B4').setNumberFormat('0.00');d.getRange('B6').setNumberFormat('0.00');
    const graf=(tipo,titulo,cols,fila,col)=>{
      const b=d.newChart().setChartType(tipo).setNumHeaders(1).setPosition(fila,col,0,0)
        .setOption('title',titulo).setOption('width',460).setOption('height',280).setOption('legend',{position:'none'});
      cols.forEach(c=>b.addRange(res.getRange(c)));
      d.insertChart(b.build());
    };
    graf(Charts.ChartType.COLUMN,'Evaluaciones por día',['A:A','B:B'],8,1);
    graf(Charts.ChartType.LINE,'Brix promedio por día',['A:A','C:C'],8,6);
    graf(Charts.ChartType.COLUMN,'Peso neto por día (Kg)',['A:A','F:F'],23,1);
    graf(Charts.ChartType.LINE,'Desgrane promedio por día (%)',['A:A','G:G'],23,6);
  }
}

function cierre(ss,d){
  const h=ss.getSheetByName('Cierres'),n=h.getLastRow();
  const fila=[fecha(d.iso),d.horaCierre,d.n,d.brixProm===null?'':d.brixProm,d.brixMin===null?'':d.brixMin,d.brixMax===null?'':d.brixMax,d.horaMax||'',d.horaMin||'',d.tendencia||'',(d.puntos||[]).join(' | '),d.iso];
  let r=0;
  if(n>1)h.getRange(2,11,n-1,1).getValues().forEach((x,i)=>{if(x[0]===d.iso)r=i+2});
  if(r)h.getRange(r,1,1,fila.length).setValues([fila]);else h.getRange(n+1,1,1,fila.length).setValues([fila]);
  return {ok:true};
}
