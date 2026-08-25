// ============================================================
// ALMACÉN · lógica de importar Excel + guardar en GitHub, copiada
// tal cual (sin reescribir) de inventario.html para poder usarse
// también desde el importador centralizado de plataforma.html.
//
// Namespaced en AlmacenShared para no chocar con Compras/Facturas
// cuando las 3 conviven en la misma página (el launcher).
//
// A diferencia de guardarTodo() (inventario.html), este módulo SOLO
// sube lo que puede salir de una importación de Excel fresca:
// Existencias, Requerimientos propios y Tránsito/Reservas (dentro de
// data.json). NO toca Picking ni Códigos generados — esas son
// decisiones manuales de UI que solo existen dentro de inventario.html
// y deben seguir guardándose únicamente desde ahí con su propio botón
// GUARDAR (por eso no se tocó ese archivo).
//
// setExistencias/setTransito/setRequerimientosAlm se simplificaron
// para no llamar funciones de refresco de UI (fillFilters, render,
// renderReq, etc.) que no existen en el launcher — el resto del
// parseo es idéntico al original.
// ============================================================
const AlmacenShared = (function(){
let DATA = [];
let TRANS = [], RESV = [];
let TRANS_META = null;
let TRANSITO_LIVE = false, TRANS_DIRTY = false;
let TRANS_BY_COD = {}, RESV_BY_COD = {};
let P5 = [], OC_RAW = [], SKUS_RAW = {}, SINOC = [], SALIDAS_ALM = {};
let REQ_ALM_LIVE = false, REQ_ALM_DIRTY = false;
let EXISTENCIAS_LIVE = false, EXISTENCIAS_DIRTY = false;
let REQCOD_OC = {}, SINOC_BY_REQCOD = {}, OC_BY_NUM = {};

const GH_REPO='Melissa2026714/dashboard-friopacking';
const GH_FILE='data.json';
const GH_API='https://api.github.com/repos/'+GH_REPO+'/contents/'+GH_FILE;
const GH_TOKEN_KEY='fp_gh_token_almacen';
const GH_FILE_EXIST='almacen_existencias.json';
const GH_API_EXIST='https://api.github.com/repos/'+GH_REPO+'/contents/'+GH_FILE_EXIST;
const GH_FILE_REQALM='almacen_requerimientos.json';
const GH_API_REQALM='https://api.github.com/repos/'+GH_REPO+'/contents/'+GH_FILE_REQALM;

/* ==== INICIO: copiado de inventario.html (ver cabecera del archivo) ==== */

// pickField, toNum, round2, fetchTO, getGHTokenAlmacen, setupGHToken
function getGHTokenAlmacen(){
  const raw=localStorage.getItem(GH_TOKEN_KEY)||'';
  const clean=raw.replace(/\s+/g,'');
  if(clean!==raw)localStorage.setItem(GH_TOKEN_KEY,clean);
  return clean;
}
// GUARDAR sube archivos de varios MB (Requerimientos propios, Existencias) y fetch() nunca
// tiene timeout propio: con la wifi lenta/inestable del almacén, un PUT que se cuelga deja el
// botón en "Guardando..." para siempre, sin error visible — parecía que "no guardaba nada"
// cuando en realidad nunca terminó de intentarlo. Con esto, a los 60s se corta y se avisa.
function fetchTO(url,opts,ms){
  const ms_=ms||60000;
  return Promise.race([
    fetch(url,opts),
    new Promise((_,rej)=>setTimeout(()=>rej(new Error('tiempo de espera agotado ('+Math.round(ms_/1000)+' s) — revisa tu conexión e intenta de nuevo')),ms_))
  ]);
}
function setupGHToken(){
  const cur=localStorage.getItem(GH_TOKEN_KEY)||'';
  const t=prompt('Ingresa el Token de GitHub de Almacén\n(propio de este módulo, no el de Compras — se guarda solo en este navegador):',cur);
  if(t!==null){
    const clean=t.replace(/\s+/g,'');
    localStorage.setItem(GH_TOKEN_KEY,clean);
    alert(clean?'✅ Token guardado. Ya puedes usar GUARDAR.':'⚠ Token eliminado.');
  }
}

/* ═══════════════ IMPORTAR EXCEL (Existencias/Kardex) — acepta varios archivos a la vez ═══════════════ */
// Busca un campo en la fila probando varios nombres de columna posibles (insensible a mayúsculas/espacios).
function pickField(row,candidates){
  const keys=Object.keys(row);
  for(const cand of candidates){
    const k=keys.find(kk=>kk.trim().toLowerCase()===cand.toLowerCase());
    if(k!==undefined && row[k]!=='' && row[k]!=null)return row[k];
  }
  return '';
}
function toNum(v){
  if(v==null||v==='')return 0;
  const s=String(v).trim();
  if(s==='#N/A'||s==='#N/D'||s==='')return 0;
  const n=Number(s.replace(/,/g,''));
  return isNaN(n)?0:n;
}
// Redondea a 2 decimales (evita ruido de punto flotante) SIN truncar a entero — hay materiales
// que se requieren en fracciones (ej. 0.44 MTS de cable/tubo), Math.round() los dejaba en 0.
function round2(n){return Math.round((n||0)*100)/100;}

// parseExistenciasWorkbook
function parseExistenciasWorkbook(wb){
  // Prioriza la hoja "Análisis Inventario" (el reporte real de ICA/Lima/Trujillo);
  // si no existe, cae a Existencias/Kardex/Stock; si tampoco, usa la primera hoja.
  let sheetName=wb.SheetNames.find(n=>/an[aá]lisis.*inventario|inventario.*an[aá]lisis/i.test(n))
    ||wb.SheetNames.find(n=>/existenc|kardex|stock/i.test(n))
    ||wb.SheetNames[0];
  const sheet=wb.Sheets[sheetName];
  const rows=XLSX.utils.sheet_to_json(sheet,{defval:''});
  const out=[];
  rows.forEach(row=>{
    const cod=String(pickField(row,['Código','Codigo','Idproducto','SKU','Cod'])).trim();
    if(!cod)return;
    const min=toNum(pickField(row,['Stock mínimo','Stock minimo','Mínimo','Minimo']));
    const max=toNum(pickField(row,['Stock máximo','Stock maximo','Máximo','Maximo']));
    out.push({
      cod,
      desc:String(pickField(row,['Descripción','Descripcion','Producto','Nombre'])).trim().toUpperCase(),
      um:String(pickField(row,['Medida','UM','Unidad','Unidad de Medida'])||'UND').trim(),
      cat:'—', ubi:'—', // no vienen en Análisis Inventario; se mantienen por compatibilidad con Códigos/Picking
      ica:toNum(pickField(row,['ICA'])),
      lima:toNum(pickField(row,['LIMA'])),
      trujillo:toNum(pickField(row,['TRUJILLO'])),
      transito:toNum(pickField(row,['Stock en tránsito','Stock en transito'])),
      stock:toNum(pickField(row,['Stock Actual','Stock actual'])),
      consumoM:toNum(pickField(row,['Consumo Mensual'])),
      consumoD:toNum(pickField(row,['Consumo Diario'])),
      leadTime:toNum(pickField(row,['Lead Time'])),
      desvest:toNum(pickField(row,['Desvest','Desviación','Desviacion'])),
      stockSeg:toNum(pickField(row,['Stock Seguridad'])),
      rop:toNum(pickField(row,['Punto de Reorden'])),
      min, max,
      estado:normEstado(pickField(row,['Estado'])),
      demandaAnual:toNum(pickField(row,['Demanda Anual'])),
      costoPedido:toNum(pickField(row,['Costo Pedido'])),
      pctMant:toNum(pickField(row,['% Mantenimiento'])),
      costo:toNum(pickField(row,['Costo Unitario','Costo','Precio','Precio Unitario'])),
      costoMant:toNum(pickField(row,['Costo Mantener'])),
      eoq:toNum(pickField(row,['EOQ'])),
      reservas:toNum(pickField(row,['Reservas'])),
      stockDisp:toNum(pickField(row,['Stock Disponible'])),
      cobertura:toNum(pickField(row,['Cobertura (meses)','Cobertura'])),
      deficitROP:toNum(pickField(row,['Deficit contra ROP','Déficit contra ROP'])),
      cantFaltante:toNum(pickField(row,['Cant. Faltante a comprar','Cant Faltante a comprar'])),
      prioridad:String(pickField(row,['Prioridad'])||'—').trim()
    });
  });
  return out;
}

// normEstado, fmt, fixEnc
function normEstado(e){
  const v=String(e||'').trim().toUpperCase();
  if(v==='CRITICO'||v==='CRÍTICO')return'CRITICO';
  if(v==='REABASTECER')return'REABASTECER';
  if(v==='SOBRESTOCK')return'SOBRESTOCK';
  if(v==='OK')return'OK';
  return'SIN DATOS';
}
function fmt(n){return (n||0).toLocaleString('es-PE',{minimumFractionDigits:2,maximumFractionDigits:2});}
function fixEnc(s){return (s||'').replace('revisi�n','revisión').replace('Almac�n','Almacén');}

// xfecha, dmyOrd, xtexto, xcodigo, fmtq, xfechaAny, hojaAObjetos
function xfecha(v){ // Date | serial Excel | texto → "dd/mm/yyyy"
  if(v==null||v==='')return '';
  if(v instanceof Date&&!isNaN(v))return String(v.getDate()).padStart(2,'0')+'/'+String(v.getMonth()+1).padStart(2,'0')+'/'+v.getFullYear();
  if(typeof v==='number'&&v>30000){const d=new Date(Math.round((v-25569)*86400000));return String(d.getUTCDate()).padStart(2,'0')+'/'+String(d.getUTCMonth()+1).padStart(2,'0')+'/'+d.getUTCFullYear();}
  return String(v).split(' ')[0];
}
function dmyOrd(s){const p=(s||'').split('/');return p.length===3?(+p[2])*10000+(+p[1])*100+(+p[0]):99999999;}
function xtexto(v){return v==null?'':String(v).trim();}
function xcodigo(v){const s=xtexto(v);return (s==='-'||s.toLowerCase()==='nan'||s.toLowerCase()==='none')?'':s;} // el Maestro usa "-" cuando no hay OC/pedido
function fmtq(n){return (n||0).toLocaleString('es-PE',{maximumFractionDigits:2});}
// El nombre exacto de la columna de fecha de aprobación de OC varía entre exports del Reporte
// Almacén (visto también en Compras/index.html) — se prueban los nombres conocidos en orden.
function xfechaAny(r,names){
  for(let i=0;i<names.length;i++){
    const v=xfecha(r[names[i]]);
    if(v)return v;
  }
  return '';
}

// Lee una hoja cuyos encabezados pueden no estar en la fila 1 (el Reporte Almacen los trae en la fila 3).
function hojaAObjetos(ws){
  const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:null});
  let hdr=-1;
  for(let i=0;i<Math.min(12,raw.length);i++){
    const r=raw[i]||[];
    if(r.some(c=>typeof c==='string'&&/^(CodRequerimiento|CodPedido)$/i.test(String(c).trim()))){hdr=i;break;}
  }
  if(hdr<0)return null;
  const heads=(raw[hdr]||[]).map(h=>h?String(h).trim():'');
  return raw.slice(hdr+1).filter(r=>r&&r.some(c=>c!=null&&c!=='')).map(r=>{
    const o={};heads.forEach((h,i)=>{if(h&&o[h]===undefined)o[h]=r[i];});return o;
  });
}

// Devuelve {t:[líneas tránsito], r:[líneas reserva]} o null si el workbook no trae esas hojas.

// parseTransitoWorkbook
function parseTransitoWorkbook(wb){
  const shM=wb.SheetNames.find(n=>/maestro/i.test(n));
  const shP=wb.SheetNames.find(n=>/ped\s*sin\s*req/i.test(n));
  if(!shM&&!shP)return null;
  const T=new Map(),R=new Map(); // dedup: el Maestro repite la línea por cada recepción/salida

  if(shM){
    (hojaAObjetos(wb.Sheets[shM])||[]).forEach(r=>{
      const eReq=xtexto(r['EstadoRequerimiento']).toLowerCase();
      const ePed=xtexto(r['EstadoPedido']).toLowerCase();
      // ── Fuente 1: Stock en Tránsito ── (la OC debe estar APROBADA)
      const eOrd=xtexto(r['EstadoOrden']).toLowerCase();
      const esTransito=(eReq==='enviado a pedido'||eReq==='enviado parcial a pedido')&&(ePed==='pendiente'||ePed==='con orden de compra')&&eOrd==='pendiente (aprobado)';
      if(esTransito){
        const key=[xcodigo(r['CodRequerimiento']),xcodigo(r['CodPedido']),xcodigo(r['CodOrden']),xtexto(r['IdProducto'])].join('|');
        let t=T.get(key);
        if(!t){
          t={fuente:'Reporte Maestro',oc:xcodigo(r['CodOrden']),ped:xcodigo(r['CodPedido']),ref:xcodigo(r['CodRequerimiento']),
             cod:xtexto(r['IdProducto']),prod:xtexto(r['Producto']),unid:xtexto(r['UnidadMedida'])||'UND',
             resp:xtexto(r['NombreResponsable']),proy:xtexto(r['NombreProyecto']),idproy:xtexto(r['IdProyecto']),
             cant:0,cantRec:0,fped:xfecha(r['FechaPedido']),foc:xfecha(r['FechaOrden']),fapro:xfecha(r['FechaAprobacionOC']),fent:xfecha(r['ParaFechaOrden']),
             respPed:xtexto(r['ResponsablePedido']),ucomp:xtexto(r['UsuarioCompras']),prov:xtexto(r['NombreProveedor']),
             estadoPed:xtexto(r['EstadoPedido']),estadoOC:xtexto(r['EstadoOrden'])};
          T.set(key,t);
        }
        // Cada fila del grupo es un ítem distinto del mismo producto (verificado: las filas no se repiten por recepción en estados pendientes) → se SUMA.
        t.cant+=toNum(r['Cantidad de orden'])||toNum(r['Cantidad de pedido'])||toNum(r['CantidadRequerimiento']);
        t.cantRec+=toNum(r['CantRecepcion']);
      }
      // ── Reservas (cálculo independiente) ──
      const esReserva=(eReq.includes('revisi')&&eReq.includes('almac'))||eReq==='enviado a pedido'||eReq==='enviado parcial a pedido';
      if(esReserva){
        const kr=xcodigo(r['CodRequerimiento'])+'|'+xtexto(r['IdProducto']);
        let rv=R.get(kr);
        if(!rv){
          rv={req:xcodigo(r['CodRequerimiento']),fecha:xfecha(r['FechaRequerimiento']),freq:xfecha(r['ParaFechaRequerimiento']),
            resp:xtexto(r['NombreResponsable']),idproy:xtexto(r['IdProyecto']),proy:xtexto(r['NombreProyecto']),
            cod:xtexto(r['IdProducto']),prod:xtexto(r['Producto']),cant:0,unid:xtexto(r['UnidadMedida'])||'UND',
            fam:xtexto(r['NombreFamilia']),estado:xtexto(r['EstadoRequerimiento'])};
          R.set(kr,rv);
        }
        // Se SUMA: las filas repetidas por recepción traen cantidad 0, y las filas con cantidad son ítems/partes distintas del requerimiento.
        rv.cant+=toNum(r['CantidadRequerimiento']);
      }
    });
  }

  if(shP){
    (hojaAObjetos(wb.Sheets[shP])||[]).forEach(r=>{
      const ePed=xtexto(r['EstadoPedido']).toLowerCase();
      if(ePed!=='pendiente'&&ePed!=='con orden de compra')return;
      if(xtexto(r['EstadoOrden']).toLowerCase()!=='pendiente (aprobado)')return; // solo OC aprobadas cuentan como tránsito
      const key=['P4',xcodigo(r['CodPedido']),xcodigo(r['CodOrden']),xtexto(r['IdProducto'])].join('|');
      let t=T.get(key);
      if(!t){
        t={fuente:'PedSinReq',oc:xcodigo(r['CodOrden']),ped:xcodigo(r['CodPedido']),ref:'',
           cod:xtexto(r['IdProducto']),prod:xtexto(r['Descripción']!=null?r['Descripción']:r['Descripcion']),unid:xtexto(r['Medida'])||'UND',
           resp:'',proy:'',idproy:'',
           cant:0,cantRec:0,
           fped:xfecha(r['FechaPedido']),foc:xfecha(r['FechaOrden']),fapro:xfecha(r['FechaAprobacionOC']),fent:xfecha(r['ParaFechaOrden']),
           respPed:xtexto(r['ResponsablePedido']),ucomp:xtexto(r['UsuarioCompras']),prov:xtexto(r['NombreProveedor']),
           estadoPed:xtexto(r['EstadoPedido']),estadoOC:xtexto(r['EstadoOrden'])};
        T.set(key,t);
      }
      // Cada fila del grupo es un ítem distinto del mismo producto → se SUMA (igual que en Reporte Maestro).
      t.cant+=toNum(r['Cantidad de orden'])||toNum(r['Cantidad de pedido']);
      t.cantRec+=toNum(r['CantRecepcion']);
    });
  }

  const trans=[];T.forEach(t=>{t.cantPend=Math.max(0,(t.cant||0)-(t.cantRec||0));trans.push(t);});
  trans.sort((a,b)=>dmyOrd(a.fent)-dmyOrd(b.fent));
  const resv=[];R.forEach(x=>resv.push(x));
  resv.sort((a,b)=>dmyOrd(a.fecha)-dmyOrd(b.fecha));
  return {t:trans,r:resv};
}

// Construye, desde la MISMA hoja "Reporte Maestro" de Reporte Almacen.xlsx, una copia propia e
// independiente de Requerimientos / OC / Sin OC — Cotización, para que Almacén ya no dependa de
// que Compras haga su propio import. Nunca toca D.p5reqs/D.oc/D.sinoc/SKUS (siguen siendo
// exclusivos de Compras) — esto se guarda aparte en almacen_requerimientos.json.

// parseRequerimientosWorkbook
function parseRequerimientosWorkbook(wb){
  const shM=wb.SheetNames.find(n=>/maestro/i.test(n));
  if(!shM)return null;
  const rows=hojaAObjetos(wb.Sheets[shM])||[];
  if(!rows.length)return null;

  // Requerimientos: mismo fix que en Compras — una fila con el mismo req+código+cantidad EXACTA
  // ya vista es el Maestro repitiendo la línea por cada recepción/salida (se ignora); una fila
  // con el mismo req+código pero OTRA cantidad es una línea real aparte (se suma, no se pierde).
  // OJO: ese dedup por cantidad exacta solo es seguro si la fila tiene evidencia de una OC o de
  // una recepción/salida real (solo entonces el Maestro puede repetir la MISMA línea) — tener
  // apenas un N° de Pedido asignado NO basta, porque eso no implica que ya se haya recepcionado
  // o despachado nada. Si el requerimiento sigue "Pendiente" sin OC ni movimiento todavía, no hay
  // repetición posible por recepción, así que dos filas con el mismo req+código+cantidad son dos
  // ítems reales distintos
  // (caso real: req 0001-0022016, 2 líneas de 2 CIL cada una = 4, no 2 — NISIRA las trae como
  // ítems 001 y 002 separados, y el Maestro no exporta ese N° de ítem para poder diferenciarlas).
  const p5Seen=new Set(),p5IdxByReqCod={},p5=[];
  rows.forEach(function(r){
    const estado=xtexto(r['EstadoRequerimiento']);
    if(!estado)return;
    const req=xcodigo(r['CodRequerimiento']),cod=xtexto(r['IdProducto']);
    if(!req||!cod)return;
    const cant=round2(toNum(r['CantidadRequerimiento']));
    const hayActividad=!!xcodigo(r['CodOrden'])||toNum(r['CantRecepcion'])>0||toNum(r['CantidadSalida'])>0;
    if(hayActividad){
      const key3=req+'|'+cod+'|'+cant;
      if(p5Seen.has(key3))return;
      p5Seen.add(key3);
    }
    const key2=req+'|'+cod;
    if(p5IdxByReqCod[key2]!=null){p5[p5IdxByReqCod[key2]].cant+=cant;return;}
    p5IdxByReqCod[key2]=p5.length;
    p5.push({
      req:req,fecha:xfecha(r['FechaRequerimiento']),
      resp:xtexto(r['NombreResponsable']),idproy:xtexto(r['IdProyecto']),
      proy:xtexto(r['NombreProyecto']),cod:cod,
      prod:xtexto(r['Producto']),cant:cant,
      unid:xtexto(r['UnidadMedida'])||xtexto(r['Medida'])||'',
      freq:xfecha(r['ParaFechaRequerimiento'])||xfecha(r['FechaRequerimiento']),
      estado:estado,fam:xtexto(r['NombreFamilia'])
    });
  });

  // OC + cruce por requerimiento (equivalente a D.oc / SKUS) — agrupado por N° de OC.
  const ocMap={},skus={};
  rows.forEach(function(r){
    const ocNum=xcodigo(r['CodOrden']);
    if(!ocNum)return;
    const req=xcodigo(r['CodRequerimiento']),cod=xtexto(r['IdProducto']);
    const ped=xcodigo(r['CodPedido']);
    const co=toNum(r['Cantidad de orden']),cr=toNum(r['CantRecepcion']);
    const fapro=xfechaAny(r,['FechaAprobacionOC','FechaAprobacion','Fecha Aprobacion','FechaAprobacionOrden']);
    if(!ocMap[ocNum]){
      ocMap[ocNum]={oc:ocNum,ped:ped,peds:[],estado:xtexto(r['EstadoOrden']),
        fped:xfecha(r['FechaPedido']),foc:xfecha(r['FechaOrden']),fapro:fapro,fent:xfecha(r['ParaFechaOrden']),prov:xtexto(r['NombreProveedor']),
        respPed:xtexto(r['ResponsablePedido'])};
    }
    if(ped&&ocMap[ocNum].peds.indexOf(ped)<0)ocMap[ocNum].peds.push(ped);
    if(!skus[ocNum])skus[ocNum]=[];
    skus[ocNum].push({
      cod:cod,prod:xtexto(r['Producto']),
      cantOrd:round2(co),cantRec:round2(cr),cantPend:Math.max(0,round2(co-cr)),
      foc:xfecha(r['FechaOrden']),fapro:fapro,fent:xfecha(r['ParaFechaOrden']),
      fechaRecep:xfecha(r['FechaRecepcion']),guiaProv:xcodigo(r['CodRecepcion']),
      estado:xtexto(r['EstadoOrden']),req:req,freq:xfecha(r['FechaRequerimiento']),
      unid:xtexto(r['UnidadMedida'])||xtexto(r['Medida'])||'',
      // Pedido EXACTO de esta línea — una OC puede consolidar varios pedidos distintos
      // (ver ocMap[ocNum].peds), así que el N° de Pedido "genérico" de la OC no siempre
      // corresponde al pedido real de este req+código puntual.
      ped:ped,fped:xfecha(r['FechaPedido']),respPed:xtexto(r['ResponsablePedido'])
    });
  });
  const oc=Object.values(ocMap);

  // Sin OC — Cotización (equivalente a D.sinoc): pedidos con N° de pedido pero sin OC todavía.
  // Se mantiene el registro aunque NISIRA lo haya marcado "Archivado" (para no perder N° de
  // Pedido/Fecha/Responsable en la tabla) — el campo "estado" ya guarda ese valor y estadoItem()
  // lo usa para mostrar la etiqueta correcta automáticamente, sin depender de que alguien lo
  // archive a mano en el dashboard.
  const sinocSeen=new Set(),sinoc=[];
  rows.forEach(function(r){
    if(xcodigo(r['CodOrden']))return;
    const ped=xcodigo(r['CodPedido']);
    if(!ped)return;
    const req=xcodigo(r['CodRequerimiento']),cod=xtexto(r['IdProducto']);
    if(!req||!cod)return;
    const key=ped+'|'+cod;
    if(sinocSeen.has(key))return;
    sinocSeen.add(key);
    sinoc.push({
      ped:ped,fped:xfecha(r['FechaPedido']),req:req,cod:cod,prod:xtexto(r['Producto']),
      cant:round2(toNum(r['CantidadRequerimiento'])),
      resp:xtexto(r['NombreResponsable']),proy:xtexto(r['NombreProyecto']),idproy:xtexto(r['IdProyecto']),
      estado:xtexto(r['EstadoPedido']),respPed:xtexto(r['ResponsablePedido'])
    });
  });

  // Salidas a obra (equivalente a "En Obra"/"N° Guía"): columnas DocumentoDeOrigen (guía),
  // CantidadSalida y FechaSalida, mismo join req+código. Una fila sin salida trae estas 3
  // columnas vacías, así que se ignora — solo se suma cuando hay un despacho real registrado.
  // Cuando ya existe esto, "En Obra"/"N° Guía" salen del Excel (solo lectura) en vez de digitarse
  // a mano; el campo manual queda solo para lo que todavía no aparece en el Excel.
  const salBy={};
  rows.forEach(function(r){
    const cantSal=toNum(r['CantidadSalida']);
    const guia=xtexto(r['DocumentoDeOrigen']);
    if(!cantSal&&!guia)return;
    const req=xcodigo(r['CodRequerimiento']),cod=xtexto(r['IdProducto']);
    if(!req||!cod)return;
    const key=req+'|'+cod;
    if(!salBy[key])salBy[key]={obra:0,guias:[],fsal:''};
    salBy[key].obra+=cantSal;
    if(guia&&salBy[key].guias.indexOf(guia)<0)salBy[key].guias.push(guia);
    const fsal=xfecha(r['FechaSalida']);
    if(fsal&&(!salBy[key].fsal||dmyOrd(fsal)>dmyOrd(salBy[key].fsal)))salBy[key].fsal=fsal;
  });

  return {p5:p5,oc:oc,skus:skus,sinoc:sinoc,salidas:salBy};
}

// setRequerimientosAlm, setTransito, applyTransitoADatos
function setRequerimientosAlm(p5,oc,skus,sinoc,salidas,marcarDirty){
  P5=(Array.isArray(p5)?p5:[]).map(function(r){return {...r,estado:fixEnc(r.estado)};});
  OC_RAW=Array.isArray(oc)?oc:[];
  SKUS_RAW=skus||{};
  SINOC=Array.isArray(sinoc)?sinoc:[];
  SALIDAS_ALM=salidas||{};
  P5_LIVE=true;
  REQ_ALM_LIVE=true;
  OC_BY_NUM={};
  OC_RAW.forEach(function(o){if(o.oc)OC_BY_NUM[o.oc]=o;});
  buildReqCodOC();
  buildSinocLookup();
  if(marcarDirty)REQ_ALM_DIRTY=true;
}

function setTransito(t,r,marcarDirty,ts){
  TRANS=Array.isArray(t)?t:[];
  RESV=Array.isArray(r)?r:[];
  TRANS_META={ts:ts||Date.now()};
  TRANSITO_LIVE=true;
  if(marcarDirty)TRANS_DIRTY=true;
  TRANS_BY_COD={};TRANS.forEach(x=>{if(x.cod)TRANS_BY_COD[x.cod]=(TRANS_BY_COD[x.cod]||0)+(x.cantPend!=null?x.cantPend:(x.cant||0));});
  RESV_BY_COD={};RESV.forEach(x=>{if(x.cod)RESV_BY_COD[x.cod]=(RESV_BY_COD[x.cod]||0)+(+x.cant||0);});
}

// Aplica los totales calculados a cada ítem de Existencias (columnas "En Tránsito" y "Reservas" de la hoja General).
// El Stock Disponible NO se recalcula: se respeta la columna "Stock Disponible" que ya trae el Análisis Inventario.
function applyTransitoADatos(){
  if(!TRANSITO_LIVE||!Array.isArray(DATA))return;
  DATA.forEach(d=>{
    d.transito=TRANS_BY_COD[d.cod]||0;
    d.reservas=RESV_BY_COD[d.cod]||0;
  });
}


// buildReqCodOC, buildSinocLookup
function buildReqCodOC(){
  REQCOD_OC={};
  Object.keys(SKUS_RAW).forEach(function(oc){
    (SKUS_RAW[oc]||[]).forEach(function(line){
      if(!line.req||!line.cod)return;
      const k=line.req+'|'+line.cod;
      if(!REQCOD_OC[k])REQCOD_OC[k]={ocs:[oc],foc:line.foc||'',fapro:line.fapro||'',fent:line.fent||'',compras:0,cantOrd:0,cantRec:0,guiasProv:[],frecep:'',ped:'',fped:'',respPed:''};
      else if(REQCOD_OC[k].ocs.indexOf(oc)<0)REQCOD_OC[k].ocs.push(oc);
      REQCOD_OC[k].compras+=(line.cantPend||0);
      REQCOD_OC[k].cantOrd+=(line.cantOrd||0);
      // Recepción real (hoja Reporte Maestro, columnas resaltadas): CantRecepcion (se suma, una
      // fila por cada recepción parcial), FechaRecepcion (se toma la más reciente) y CodRecepcion
      // (N° de guía del proveedor, se listan todos si hubo más de una recepción).
      REQCOD_OC[k].cantRec+=(line.cantRec||0);
      if(line.guiaProv&&REQCOD_OC[k].guiasProv.indexOf(line.guiaProv)<0)REQCOD_OC[k].guiasProv.push(line.guiaProv);
      if(line.fechaRecep&&(!REQCOD_OC[k].frecep||dmyOrd(line.fechaRecep)>dmyOrd(REQCOD_OC[k].frecep)))REQCOD_OC[k].frecep=line.fechaRecep;
      // Pedido exacto de ESTA línea (una OC puede consolidar varios pedidos distintos) — se
      // guarda el primero no vacío que aparezca para este req+código puntual.
      if(!REQCOD_OC[k].ped&&line.ped)REQCOD_OC[k].ped=line.ped;
      if(!REQCOD_OC[k].fped&&line.fped)REQCOD_OC[k].fped=line.fped;
      if(!REQCOD_OC[k].respPed&&line.respPed)REQCOD_OC[k].respPed=line.respPed;
    });
  });
}
// Lookup req+'|'+cod -> fila de D.sinoc, para saber si el ítem está en la ventana
// "Sin OC — Cotización" de Compras y en qué estado (en cotización / archivado / pendiente).
function buildSinocLookup(){
  SINOC_BY_REQCOD={};
  (SINOC||[]).forEach(function(r){
    if(!r.req||!r.cod)return;
    SINOC_BY_REQCOD[r.req+'|'+r.cod]=r;
  });
}

/* ==== FIN copiado de inventario.html ==== */

// Versión simplificada del setExistencias original: sin fillFilters()/render()
// (no existen en el launcher). El cálculo real (DATA=rows, marcar dirty) es igual.
function setExistencias(rows,marcarDirty){
  DATA=rows;
  EXISTENCIAS_LIVE=true;
  if(marcarDirty)EXISTENCIAS_DIRTY=true;
}

// Procesa un archivo "Reporte Almacen.xlsx" (puede traer Existencias +
// Reporte Maestro/PedSinReq para Tránsito y Requerimientos, todo en el mismo
// libro) — mismo comportamiento que handleImportAlmacen() de inventario.html,
// sin las llamadas de refresco de UI.
function procesarArchivo(wb){
  const partes=[];
  let existRows=[];
  try{existRows=parseExistenciasWorkbook(wb);}catch(_){existRows=[];}
  if(existRows.length) partes.push(existRows.length+' existencias');
  const tr=parseTransitoWorkbook(wb);
  if(tr){
    setTransito(tr.t,tr.r,true);
    partes.push('tránsito ('+tr.t.length+' líneas OC) y reservas ('+tr.r.length+' líneas req.)');
  }
  const req=parseRequerimientosWorkbook(wb);
  if(req){
    setRequerimientosAlm(req.p5,req.oc,req.skus,req.sinoc,req.salidas,true);
    partes.push(req.p5.length.toLocaleString('es-PE')+' requerimientos ('+req.oc.length+' OC · '+req.sinoc.length+' sin OC)');
  }
  if(existRows.length){ setExistencias(existRows,true); }
  if(!partes.length) throw new Error('no se encontraron filas con código válido (Existencias/Reporte Maestro/PedSinReq)');
  if(TRANS_DIRTY) applyTransitoADatos();
  return partes;
}

// Sube Existencias / Requerimientos / Tránsito (dentro de data.json) — cada uno
// solo si hubo import nuevo de ese tipo. NO toca Picking ni Códigos (eso vive
// solo en inventario.html, con sus propias decisiones manuales de UI).
async function guardarImportado(){
  const token=(localStorage.getItem(GH_TOKEN_KEY)||'').trim();
  if(!token) throw new Error('Falta configurar el Token de GitHub de Almacén.');
  const partes=[];
  if(EXISTENCIAS_DIRTY){
    const checkE=await fetchTO(GH_API_EXIST,{headers:{Authorization:'token '+token,Accept:'application/vnd.github.v3+json'}});
    const shaE=checkE.ok?(await checkE.json()).sha:undefined;
    const contentE=btoa(unescape(encodeURIComponent(JSON.stringify({existencias:DATA,ts:Date.now()}))));
    const bodyE={message:'almacen: guardar Existencias ('+DATA.length+' SKUs) — '+new Date().toISOString(),content:contentE};
    if(shaE)bodyE.sha=shaE;
    const resE=await fetchTO(GH_API_EXIST,{method:'PUT',headers:{Authorization:'token '+token,Accept:'application/vnd.github.v3+json','Content-Type':'application/json'},body:JSON.stringify(bodyE)},90000);
    if(!resE.ok)throw new Error('GitHub rechazó la subida de Existencias (HTTP '+resE.status+')');
    partes.push('Existencias ('+DATA.length+' SKUs)');
    EXISTENCIAS_DIRTY=false;
  }
  if(REQ_ALM_DIRTY){
    const checkR=await fetchTO(GH_API_REQALM,{headers:{Authorization:'token '+token,Accept:'application/vnd.github.v3+json'}});
    const shaR=checkR.ok?(await checkR.json()).sha:undefined;
    const contentR=btoa(unescape(encodeURIComponent(JSON.stringify({p5:P5,oc:OC_RAW,skus:SKUS_RAW,sinoc:SINOC,salidas:SALIDAS_ALM,ts:Date.now()}))));
    const bodyR={message:'almacen: guardar Requerimientos propios ('+P5.length+' líneas) — '+new Date().toISOString(),content:contentR};
    if(shaR)bodyR.sha=shaR;
    const resR=await fetchTO(GH_API_REQALM,{method:'PUT',headers:{Authorization:'token '+token,Accept:'application/vnd.github.v3+json','Content-Type':'application/json'},body:JSON.stringify(bodyR)},90000);
    if(!resR.ok)throw new Error('GitHub rechazó la subida de Requerimientos (HTTP '+resR.status+')');
    partes.push('Requerimientos ('+P5.length.toLocaleString('es-PE')+' líneas)');
    REQ_ALM_DIRTY=false;
  }
  if(TRANS_DIRTY){
    // OJO: el contenido de data.json se lee del propio meta.content de la Contents API,
    // NUNCA de raw.githubusercontent.com — ese CDN cachea varios minutos y si Compras
    // acababa de guardar data.json segundos antes, esto leía una copia vieja y la
    // reescribía encima, borrando el import recién hecho (bug detectado 2026-08-25).
    const check=await fetchTO(GH_API,{headers:{Authorization:'token '+token,Accept:'application/vnd.github.v3+json'}});
    if(!check.ok)throw new Error('No se pudo leer data.json actual (HTTP '+check.status+')');
    const meta=await check.json();
    const sha=meta.sha;
    const remote=JSON.parse(decodeURIComponent(escape(atob(meta.content.replace(/\n/g,'')))));
    if(!remote.D)throw new Error('data.json remoto no tiene el formato esperado');
    // Solo se toca almacenTransito — todo lo demás de remote.D (picking, códigos,
    // OCs de Compras, etc.) se sube exactamente como estaba en la nube.
    remote.D.almacenTransito={t:TRANS,r:RESV,ts:(TRANS_META&&TRANS_META.ts)||Date.now()};
    const content=btoa(unescape(encodeURIComponent(JSON.stringify(remote))));
    const body={message:'almacen: guardar Tránsito/Reservas ('+TRANS.length+' + '+RESV.length+' líneas) — '+new Date().toISOString(),content,sha};
    const res=await fetchTO(GH_API,{method:'PUT',headers:{Authorization:'token '+token,Accept:'application/vnd.github.v3+json','Content-Type':'application/json'},body:JSON.stringify(body)},90000);
    if(!res.ok)throw new Error('GitHub rechazó la subida (HTTP '+res.status+')');
    partes.push('Tránsito/Reservas ('+TRANS.length+' + '+RESV.length+' líneas)');
    TRANS_DIRTY=false;
  }
  return partes;
}

function hayCambios(){ return EXISTENCIAS_DIRTY||REQ_ALM_DIRTY||TRANS_DIRTY; }
function getToken(){ return (localStorage.getItem(GH_TOKEN_KEY)||'').trim(); }

return {
  procesarArchivo,
  guardarImportado,
  hayCambios,
  setupGHToken,
  getToken,
  GH_TOKEN_KEY
};
})();
