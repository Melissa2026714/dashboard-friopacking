// ============================================================
// COMPRAS · lógica de importar Excel + guardar en GitHub, copiada
// tal cual (sin reescribir) de compras.html para poder usarse también
// desde el importador centralizado de plataforma.html, sin duplicar
// la lógica de fusión anti-pisado que ya vive ahí.
//
// Namespaced en ComprasShared para no chocar con Almacen/Facturas
// cuando las 3 conviven en la misma página (el launcher).
//
// Dos únicas líneas removidas respecto al original (documentado):
// - refreshAllPages() dentro de importMaestro
// - refreshP4() dentro de importPedidoSinReq
// Ambas solo refrescan tablas del DOM propio de compras.html, que no
// existen en el launcher — el resto de la función (parseo + D.oc=...
// + saveLocal() + alert()) queda igual.
// ============================================================
const ComprasShared = (function(){
let D = {"p4":[],"oc":[],"p5reqs":[],"sinoc":[],"proj":[],"supervisores":[],"proveedores":[],"gerencial":{"summary":{"totalOC":0,"pendientes":0,"atendidoParcial":0,"atendidoCompleto":0,"totalReqs":0,"pendienteReqs":0},"byResponsable":[],"byProject":[]},"ocMeta":{},"cotMeta":{}};
let SKUS = {};
let SKUS_P4 = {}; // usado internamente por importPedidoSinReq, no se sube ni se usa fuera de este módulo
const LS_KEY='friopacking_2026_data';
const GH_REPO='Melissa2026714/dashboard-friopacking';
const GH_FILE='data.json';
const GH_RAW='https://raw.githubusercontent.com/'+GH_REPO+'/main/'+GH_FILE;
const GH_API='https://api.github.com/repos/'+GH_REPO+'/contents/'+GH_FILE;
const GH_TOKEN_KEY='fp_gh_token';
// fetch() no tiene timeout propio: con wifi lenta un GET/PUT colgado deja el Guardar
// esperando para siempre sin ningún error visible en el log del importador.
function fetchTO(url, opts, ms){
  const ms_ = ms || 45000;
  return Promise.race([
    fetch(url, opts),
    new Promise((_,rej)=>setTimeout(()=>rej(new Error('tiempo de espera agotado ('+Math.round(ms_/1000)+'s) al hablar con GitHub — revisa tu conexión e intenta de nuevo')), ms_))
  ]);
}

/* ==== INICIO: copiado de compras.html (ver cabecera del archivo) ==== */

// Reemplaza el confirm() nativo del navegador (ventana del sistema, aparte del recuadro
// de importación — fácil de perder de vista o dejar sin responder, y ahí es cuando el
// guardado se cancela en silencio) por un modal dentro de la misma página, imposible de
// no ver. Misma semántica: bloquea hasta que la persona responda, resuelve true/false.
function askConfirmVisual(mensaje){
  return new Promise(function(resolve){
    const overlay=document.createElement('div');
    overlay.style.cssText='position:fixed;inset:0;background:rgba(15,23,42,.6);z-index:99999;display:flex;align-items:center;justify-content:center;padding:20px';
    const box=document.createElement('div');
    box.style.cssText='background:#fff;border-radius:12px;padding:24px;max-width:480px;box-shadow:0 10px 40px rgba(0,0,0,.35);font-family:Arial,sans-serif';
    box.innerHTML='<div style="font-size:14px;line-height:1.6;white-space:pre-wrap;color:#1e293b;margin-bottom:18px">'+String(mensaje).replace(/&/g,'&amp;').replace(/</g,'&lt;')+'</div>'
      +'<div style="display:flex;gap:10px;justify-content:flex-end">'
      +'<button id="_cfNo" style="padding:9px 16px;border:1px solid #cbd5e1;border-radius:8px;background:#fff;cursor:pointer;font-size:13px;color:#1e293b">Cancelar</button>'
      +'<button id="_cfSi" style="padding:9px 16px;border:none;border-radius:8px;background:#dc2626;color:#fff;cursor:pointer;font-size:13px;font-weight:700">Sí, subir de todos modos</button>'
      +'</div>';
    overlay.appendChild(box);
    document.body.appendChild(overlay);
    function cerrar(v){ document.body.removeChild(overlay); resolve(v); }
    box.querySelector('#_cfSi').onclick=function(){cerrar(true);};
    box.querySelector('#_cfNo').onclick=function(){cerrar(false);};
  });
}

function saveLocal(){
  if(!D.oc.length&&!D.p4.length&&!D.proj.length)return false;
  // Excluir D.consultas (aparte, muy grande) y D.oc/p4/p5reqs (Fase 2: ya viven en
  // tablas propias de Supabase, no en el blob compartido).
  const {consultas:_omit,oc:_omitOc,p4:_omitP4,p5reqs:_omitP5,sinoc:_omitSinoc,proj:_omitProj,
    ocMeta:_omitOcMeta,cotMeta:_omitCotMeta,cotMetaFecha:_omitCotMetaFecha,
    supervisores:_omitSup,...Dsin}=D;
  const payload=JSON.stringify({D:Dsin,SKUS,ts:Date.now()});
  try{localStorage.setItem(LS_KEY,payload);}catch(e){console.warn('localStorage lleno, solo GitHub:',e);}
  return payload;
}
// loadLocal()/loadGitHub() de compras.html NO se copiaron aquí a propósito: el
// launcher no necesita leer el estado previo antes de importar — saveGitHub()
// ya se protege solo (lee la nube en el momento de guardar y fusiona ahí).

// Fusiona mapas de marcas manuales (ocMeta/cotMeta) campo por campo en vez de
// sobrescribir el objeto completo — evita que una pestaña con datos desactualizados
// borre marcas de REG./Cotización que otra persona guardó mientras tanto.

function mergeOcMeta(remoteMeta,localMeta){
  const merged=Object.assign({},remoteMeta||{});
  Object.keys(localMeta||{}).forEach(function(oc){
    const l=localMeta[oc],out=Object.assign({},merged[oc]||{});
    Object.keys(l||{}).forEach(function(f){if(l[f])out[f]=l[f];});
    merged[oc]=out;
  });
  return merged;
}
function mergeCotMeta(remoteMeta,localMeta){
  const merged=Object.assign({},remoteMeta||{});
  Object.keys(localMeta||{}).forEach(function(k){merged[k]=!!merged[k]||!!localMeta[k];});
  return merged;
}
// mergeSupervisores se eliminó: supervisores ya vive en Supabase (Fase 2), no en este
// blob — ver _MIGRADOS_SB en saveGitHub.
// Fusiona el directorio liviano de proveedores (nombre/ruc/correo, propio de
// importProveedores() en este módulo — NO la tabla Supabase "proveedores" de
// facturas.html/NEXO) por RUC o nombre normalizado, en vez de reemplazar la lista completa.
function mergeProveedores(remoteList,localList){
  const key=function(p){return(p.ruc&&String(p.ruc).trim())||_normNombre(p.nombre||'');};
  const merged={};
  (remoteList||[]).forEach(function(p){merged[key(p)]=Object.assign({},p);});
  (localList||[]).forEach(function(p){
    const k=key(p);
    const out=Object.assign({},merged[k]||{});
    ['nombre','ruc','correo'].forEach(function(f){if(p[f])out[f]=p[f];});
    merged[k]=out;
  });
  return Object.values(merged);
}

// ── oc / p4 / p5reqs: Fase 2, migrados de data.json a tablas propias en Supabase ──
// (mismo motivo que supervisores/proveedores). Este módulo solo escribe (el import
// centralizado de plataforma.html) — la lectura vive en compras.html/inventario.html.
function ocToSB(r){
  // ?? (no ||): cantRec/cantPend/ingreso/tc/punit/montoSoles suelen ser 0 legítimamente
  // (nada recibido aún, tipo de cambio raro, etc.) — con || esos ceros se volvían null.
  return {oc:r.oc,foc:r.foc??null,fapro:r.fapro??null,fent:r.fent??null,frec:r.frec??null,fsal:r.fsal??null,
    fingreso:r.fingreso??null,fpedido:r.fpedido??null,resp:r.resp??null,estado:r.estado??null,prov:r.prov??null,
    ruc:r.ruc??null,proy:r.proy??null,idproy:r.idproy??null,req:r.req??null,prod:r.prod??null,cod:r.cod??null,
    unid:r.unid??null,ucomp:r.ucomp??null,n_items:r.nItems??null,cant_ord:r.cantOrd??null,cant_rec:r.cantRec??null,
    cant_pend:r.cantPend??null,ingreso:r.ingreso??null,moneda:r.moneda??null,tc:r.tc??null,monto_soles:r.montoSoles??null,
    punit:r.punit??null,ped:r.ped??null,peds:r.peds??null,reqs:r.reqs??null,obs:r.obs??null};
}
function p4ToSB(r){
  return {oc:r.oc,foc:r.foc??null,fapro:r.fapro??null,fent:r.fent??null,frec:r.frec??null,fpedido:r.fpedido??null,
    resp_ped:r.respPed??null,estado:r.estado??null,prov:r.prov??null,ucomp:r.ucomp??null,prod:r.prod??null,
    cod:r.cod??null,unid:r.unid??null,n_items:r.nItems??null,cant_ord:r.cantOrd??null,cant_rec:r.cantRec??null,
    cant_pend:r.cantPend??null,ped:r.ped??null};
}
function p5ToSB(r){
  return {req:r.req,cod:r.cod,fecha:r.fecha??null,resp:r.resp??null,idproy:r.idproy??null,proy:r.proy??null,
    prod:r.prod??null,cant:r.cant??null,unid:r.unid??null,freq:r.freq??null,estado:r.estado??null};
}
// ── oc_items / p4_items: detalle de ítems por OC (2026-08-27 — antes SKUS/SKUS_P4
// solo vivía en data.json de GitHub, y ese blob solo se sube cuando el guardado con
// token de GitHub tiene éxito; en compras.html se migró a Supabase igual que oc/p4,
// pero esta copia del importador — la que usa el import centralizado de
// plataforma.html — no lo tenía, así que las OC importadas por acá seguían saliendo
// "0 item(s)" en el popup aunque el resumen de la OC ya se viera bien). Cada fila de
// SKUS[oc]/SKUS_P4[oc] es un movimiento del Maestro sin id propio en el Excel de
// origen — se usa (oc,cod,req) / (oc,cod,ped) como llave de upsert.
// seq = posición de cada ítem dentro del array de su OC — (oc,req)/(oc,cod) NO son
// únicos por sí solos (una OC puede tener varias recepciones parciales del mismo
// producto/requerimiento), así que seq es la única llave que nunca choca en el upsert.
function flattenSkus(map){
  const out=[];
  Object.keys(map||{}).forEach(function(oc){(map[oc]||[]).forEach(function(it,seq){out.push(Object.assign({oc:oc,seq:seq},it));});});
  return out;
}
function itemToSB(r){
  return {oc:r.oc,seq:r.seq,cod:r.cod??null,req:r.req??null,prod:r.prod??null,
    cant_ord:r.cantOrd??null,cant_rec:r.cantRec??null,cant_pend:r.cantPend??null,
    foc:r.foc??null,fent:r.fent??null,estado:r.estado??null,ucomp:r.ucomp??null,
    prov:r.prov??null,resp:r.resp??null,idproy:r.idproy??null,proy:r.proy??null,
    freq:r.freq??null,frec:r.frec??null,moneda:r.moneda??null,punit:r.punit??null,
    tc:r.tc??null,unid:r.unid??null};
}
function p4ItemToSB(r){
  return {oc:r.oc,seq:r.seq,cod:r.cod??null,ped:r.ped??null,prod:r.prod??null,
    cant_ord:r.cantOrd??null,cant_rec:r.cantRec??null,cant_pend:r.cantPend??null,
    foc:r.foc??null,fent:r.fent??null,estado:r.estado??null,ucomp:r.ucomp??null,
    prov:r.prov??null,resp:r.resp??null,frec:r.frec??null,unid:r.unid??null};
}
async function bulkUpsertSB(table,rows,onConflict,mapper,batchSize){
  const mapped=rows.map(mapper);
  batchSize=batchSize||500;
  for(let i=0;i<mapped.length;i+=batchSize){
    const chunk=mapped.slice(i,i+batchSize);
    const {error}=await sb.from(table).upsert(chunk,{onConflict});
    if(error){
      console.warn('No se pudo guardar en '+table+' (Supabase):',error);
      // Detalle visible sin abrir la consola — se acumula porque varias tablas se
      // suben en paralelo (Promise.all) y cualquiera puede fallar.
      window._sbErrors=(window._sbErrors||[]);
      window._sbErrors.push(table+': '+(error.message||error.details||JSON.stringify(error)));
      return false;
    }
  }
  return true;
}
// sinoc no tiene ninguna columna única por fila — se reemplaza la tabla entera en
// cada import en vez de hacer upsert (seguro: nadie edita sinoc a mano).
async function replaceAllSB(table,rows,mapper,batchSize,filterCol){
  // filterCol: columna NOT NULL cualquiera para poder borrar "todo" (PostgREST exige un
  // WHERE) — 'id' sirve si la tabla tiene bigserial; si no, hay que pasar una columna real.
  const {error:delErr}=await sb.from(table).delete().not(filterCol||'id','is',null);
  if(delErr){console.warn('No se pudo limpiar '+table+' (Supabase):',delErr);return false;}
  const mapped=rows.map(mapper);
  batchSize=batchSize||500;
  for(let i=0;i<mapped.length;i+=batchSize){
    const chunk=mapped.slice(i,i+batchSize);
    const {error}=await sb.from(table).insert(chunk);
    if(error){console.warn('No se pudo guardar en '+table+' (Supabase):',error);return false;}
  }
  return true;
}
function sinocToSB(r){
  return {ped:r.ped??null,fecha:r.fecha??null,resp_ped:r.respPed??null,resp:r.resp??null,
    idproy:r.idproy??null,proy:r.proy??null,cod:r.cod??null,prod:r.prod??null,cant:r.cant??null,
    unid:r.unid??null,req:r.req??null,freq:r.freq??null,estado:r.estado??null,
    parcial:r.parcial===true,obs:r.obs??null};
}
function projToSB(r){
  return {nombre:r.nombre,prys:r.prys??[],estado:r.estado??null,supervisor:r.supervisor??null,
    asistente:r.asistente??null,zona:r.zona??null,distrito:r.distrito??null,depto:r.depto??null,
    inicio:r.inicio??null,fin:r.fin??null,gps:r.gps??null,lat:r.lat??null,lon:r.lon??null,
    gps_exact:r.gps_exact===true};
}

// ── Subir a GitHub (solo quién tenga el token) ────────────────────────────────
async function saveGitHub(payload){
  const token=localStorage.getItem(GH_TOKEN_KEY);
  if(!token)return false;
  try{
    let sha='';
    // Protección: no sobrescribir la nube perdiendo información sin avisar. Si este chequeo
    // falla (red, GitHub caído, contenido corrupto) el guardado se CANCELA en vez de seguir
    // con el payload local tal cual — eso fue justo lo que pasó el 2026-08-25 (varias veces):
    // el chequeo falló en silencio y una pestaña con estado viejo en memoria subió una
    // versión incompleta, borrando proveedores/archivoMeta.
    const localData=JSON.parse(payload);
    // OJO: el contenido se lee del propio meta.content de la Contents API, NUNCA de
    // raw.githubusercontent.com — ese CDN cachea varios minutos y puede devolver una
    // copia vieja, resucitando claves ya migradas/purgadas.
    const rCheck=await fetchTO(GH_API,{headers:{Authorization:'token '+token,Accept:'application/vnd.github.v3+json'}});
    if(!rCheck.ok)throw new Error('No se pudo leer data.json actual antes de guardar (HTTP '+rCheck.status+')');
    const meta=await rCheck.json();
    sha=meta.sha||'';
    const remote=JSON.parse(decodeURIComponent(escape(atob(meta.content.replace(/\n/g,'')))));
    // oc/p4/p5reqs/sinoc/proj ya no viven aquí (Fase 2: Supabase) — ya no queda
    // ningún campo grande que valga la pena advertir por conteo antes de subir.
    // Anti-pisado: conservar claves que la nube tiene y esta pestaña no conoce
    // (ej. almacenValidado/almacenPicking creados por Almacén después de abrir esta pestaña).
    // oc/p4/p5reqs/sinoc/proj/ocMeta/cotMeta/cotMetaFecha/supervisores quedan excluidos
    // a propósito: ya viven en Supabase (Fase 2). OJO: D.proveedores aquí NO es la tabla
    // Supabase "proveedores" (esa es el directorio con datos bancarios de facturas.html/
    // NEXO) — es el directorio liviano nombre/ruc/correo propio de importProveedores()
    // en este módulo, nunca migrado, así que sigue fusionándose.
    const _MIGRADOS_SB=['oc','p4','p5reqs','sinoc','proj','ocMeta','cotMeta','cotMetaFecha','supervisores'];
    let cambiado=false;
    Object.keys(remote.D||{}).forEach(function(k){
      if(_MIGRADOS_SB.indexOf(k)!==-1)return;
      if(localData.D[k]===undefined){localData.D[k]=remote.D[k];cambiado=true;}
    });
    // Anti-pisado: fusionar proveedores (directorio liviano) en vez de reemplazar la
    // lista completa (ver bug "última escritura gana", 2026-07-31: quedaron en 0)
    if(Array.isArray(remote.D&&remote.D.proveedores)){
      localData.D.proveedores=mergeProveedores(remote.D.proveedores,localData.D.proveedores);
      cambiado=true;
    }
    // ocMeta/cotMeta ya no se fusionan aquí: Fase 2 los migró a Supabase (ver
    // _MIGRADOS_SB arriba). mergeOcMeta sigue existiendo solo para archivoMeta.
    // Igual anti-pisado para los motivos/fechas de Archivar en Sin OC — Cotización
    if(remote.D&&remote.D.archivoMeta){
      localData.D.archivoMeta=mergeOcMeta(remote.D.archivoMeta,localData.D.archivoMeta);
      cambiado=true;
    }
    if(cambiado){
      payload=JSON.stringify(localData);
      Object.assign(D,localData.D); // mantener también en memoria
    }
    // sha ya se obtuvo arriba (junto con el contenido, en la misma llamada a la Contents API)
    const content=btoa(unescape(encodeURIComponent(payload)));
    const body={message:'update '+new Date().toISOString(),content};
    if(sha)body.sha=sha;
    const res=await fetchTO(GH_API,{method:'PUT',headers:{Authorization:'token '+token,Accept:'application/vnd.github.v3+json','Content-Type':'application/json'},body:JSON.stringify(body)});
    return res.ok;
  }catch(e){
    console.warn('GitHub save failed:',e);
    alert('⚠️ No se pudo verificar la nube antes de guardar — se canceló el guardado para no subir una versión incompleta.\n\n'+e.message+'\n\nVuelve a intentar (revisa tu conexión).');
    return false;
  }
}

// ── Importar CONSULTAS.xlsx (copiado de compras.html, sin las llamadas al DOM
// propio de esa página: _st(id,...), renderP7(), alert(), markConsultasDirty()) ──
function importConsultas(wb){
  const sh=wb.Sheets[wb.SheetNames[0]];
  const raw=XLSX.utils.sheet_to_json(sh,{header:1,defval:''});
  let hdrIdx=0;
  for(let i=0;i<Math.min(10,raw.length);i++){
    const row=raw[i].map(c=>String(c||'').toUpperCase());
    if(row.some(c=>c.includes('OCOMPRA')||c.includes('IDPROD')||c.includes('RAZON_SOC')||c.includes('DESCRIPCION'))){
      hdrIdx=i; break;
    }
  }
  const hdrs=raw[hdrIdx].map(c=>String(c||'').trim().toUpperCase());
  const gi=(...cands)=>{for(const c of cands){const i=hdrs.findIndex(h=>h.includes(c.toUpperCase()));if(i>=0)return i;}return -1;};
  const iOC   =gi('DOCUM','N° OC','OCOMPRA','DOCUMENTO');
  const iFECH =gi('FECHA');
  const iRUC  =(()=>{const idx=hdrs.findIndex((h,i)=>h.includes('PROVEEDOR')&&!h.includes('RAZON')&&!h.includes('NOMBRE'));return idx>=0?idx:gi('IDCLIEPROV','IDCLIEROV');})();
  const iRAZON=(()=>{const idx=hdrs.findIndex(h=>h.includes('PROVEEDOR')&&(h.includes('RAZO')||h.includes('NOMBRE')||h.includes('SOC')));return idx>=0?idx:gi('RAZON_SOCIAL','RAZON_SOC','RAZON');})();
  const iMONE =gi('MONE','MONEDA','MON_DSC');
  const iPROD =(()=>{const idx=hdrs.findIndex(h=>h.includes('PRODUCTO')&&(h.includes('COD')||h.match(/PRODUCTO\s*C/)));return idx>=0?idx:gi('IDPROD','CODIGO')})();
  const iDESC =(()=>{const idx=hdrs.findIndex(h=>h.includes('PRODUCTO')&&(h.includes('DES')||h.includes('NOMBRE')));return idx>=0?idx:gi('DESCRIPCION','DESC')})();
  const iPRECMN=gi('PRECIO M.N. UNIT','PRECIO M.N UNIT','PRECIO MN UNIT');
  const iPRECME=gi('PRECIO M.E. UNIT','PRECIO M.E UNIT','PRECIO ME UNIT');
  const iPREC  =gi('PREC UNIT','PRECIO UNIT','VALOR UNIT','VALOR_UNIT','P.U.','PU ','PREC','VALOR','TOTAL','IMPORTE');
  const iCANT =gi('CANTIDAD','CANT');
  const iUM   =gi('U.M.','UM','UNIDAD','MEDIDA');
  const iCREA =gi('RESPON','RESP_DOC','CREADOO_POR','CREADO_POR','RESPONSABLE');
  const iEST  =gi('ESTAD','ESTADO','EST_DOC','STATUS');
  const gv=(r,i)=>i>=0?String(r[i]||'').trim():'';
  const fmtFecha=(v)=>{
    if(!v)return'';
    const s=String(v).trim();
    if(/^\d{2}\/\d{2}\/\d{4}$/.test(s))return s;
    const d=new Date(s);
    if(isNaN(d.getTime()))return s.slice(0,10);
    return d.getDate().toString().padStart(2,'0')+'/'+(d.getMonth()+1).toString().padStart(2,'0')+'/'+d.getFullYear();
  };
  const fmtMon=(v)=>{
    const s=String(v||'').trim();
    if(s==='01'||s.toLowerCase().includes('sol'))return'S/';
    if(s==='02'||s.toLowerCase().includes('dolar')||s.toLowerCase().includes('dollar'))return'US$';
    if(s==='03'||s.toLowerCase().includes('euro'))return'€';
    return s;
  };
  if(!D.consultas)D.consultas=[];
  D.consultas=raw.slice(hdrIdx+1)
    .filter(r=>r&&r.some(c=>c!=null&&c!==''))
    .map(r=>({
      oc:       gv(r,iOC),
      fecha:    fmtFecha(gv(r,iFECH)),
      fechaRaw: new Date(gv(r,iFECH)).getTime()||0,
      ruc:      gv(r,iRUC),
      proveedor:gv(r,iRAZON),
      moneda:   fmtMon(gv(r,iMONE)),
      codigo:   gv(r,iPROD),
      desc:     gv(r,iDESC),
      cant:     (()=>{const s=gv(r,iCANT);const n=parseFloat(s.replace(',','.'));return isNaN(n)?s:(n%1?n:Math.round(n));})(),
      um:       gv(r,iUM),
      valor:    (()=>{
        const mon=fmtMon(gv(r,iMONE));
        const esMN=(mon==='S/'||mon.toLowerCase().includes('sol'));
        const vMN=iPRECMN>=0?String(r[iPRECMN]||'').trim():'';
        const vME=iPRECME>=0?String(r[iPRECME]||'').trim():'';
        const vFB=iPREC>=0?String(r[iPREC]||'').trim():'';
        const noVal=(v)=>!v||v==='0'||v==='0.00'||v==='0,00';
        if(esMN){return(!noVal(vMN)?vMN:(!noVal(vME)?vME:vFB));}
        else{return(!noVal(vME)?vME:(!noVal(vMN)?vMN:vFB));}
      })(),
      resp:     gv(r,iCREA),
      estado:   gv(r,iEST),
    }));
  D.consultas.sort((a,b)=>b.fechaRaw-a.fechaRaw);

  // Hoja adicional "R.EOC" — monto total, forma de pago y comprador por OC, para Gerencial/KPIs.
  let nFin=0;
  const shEOC=wb.SheetNames.find(function(n){
    const r0=XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,defval:''}).slice(0,10);
    return r0.some(function(row){
      const up=row.map(function(c){return String(c||'').toUpperCase();});
      return up.some(function(c){return c.indexOf('IDESTADO')>=0;})&&up.some(function(c){return c.indexOf('FORMA_PAGO')>=0;});
    });
  });
  if(shEOC){
    const rawE=XLSX.utils.sheet_to_json(wb.Sheets[shEOC],{header:1,defval:''});
    let hdrE=0;
    for(let i=0;i<Math.min(10,rawE.length);i++){
      const up=rawE[i].map(function(c){return String(c||'').toUpperCase();});
      if(up.some(function(c){return c.indexOf('IDESTADO')>=0;})){hdrE=i;break;}
    }
    const hdrsE=rawE[hdrE].map(function(c){return String(c||'').trim();});
    const ciE={};hdrsE.forEach(function(h,i){if(h)ciE[h.toUpperCase()]=i;});
    const giE=function(){
      for(let a=0;a<arguments.length;a++){
        const cand=arguments[a].toUpperCase();
        const k=Object.keys(ciE).find(function(kk){return kk.indexOf(cand)>=0;});
        if(k!==undefined)return ciE[k];
      }
      return -1;
    };
    const iOC2=giE('CONCATENADO');
    const iTOTAL=giE('TOTAL');
    const iFECHA2=giE('FECHA');
    const iFAPRO2=giE('F_APROBADO','FAPROBADO');
    const iEST2=giE('EST_DSC');
    const iPROV2=giE('RAZON_SOCIAL');
    const iMON2=giE('MON_DSC');
    const iUSR2=giE('IDUSUARIOC');
    const iTC2=giE('TIPOCAMBIO');
    const iFPAGO2=giE('FORMA_PAGO');
    const gv2=function(r,i){return i>=0?r[i]:'';};
    const fmtFecha2=function(v){
      if(!v)return'';
      if(v instanceof Date)return v.getDate().toString().padStart(2,'0')+'/'+(v.getMonth()+1).toString().padStart(2,'0')+'/'+v.getFullYear();
      const s=String(v).trim();
      if(/^\d{2}\/\d{2}\/\d{4}$/.test(s))return s;
      const d=new Date(s);
      return isNaN(d.getTime())?s:(d.getDate().toString().padStart(2,'0')+'/'+(d.getMonth()+1).toString().padStart(2,'0')+'/'+d.getFullYear());
    };
    D.ocFin=D.ocFin||{};
    rawE.slice(hdrE+1).forEach(function(r){
      if(!r||!r.some(function(c){return c!=null&&c!=='';}))return;
      const oc=String(gv2(r,iOC2)||'').trim();
      if(!oc)return;
      const fecha=fmtFecha2(gv2(r,iFECHA2));
      if(!is2026(fecha))return;
      const total=parseFloat(String(gv2(r,iTOTAL)||'0').replace(',','.'))||0;
      const tc=parseFloat(String(gv2(r,iTC2)||'0').replace(',','.'))||0;
      const moneda=String(gv2(r,iMON2)||'').trim();
      const esDolares=moneda.toLowerCase().indexOf('dolar')>=0;
      const formaPago=String(gv2(r,iFPAGO2)||'').trim();
      nFin++;
      D.ocFin[oc]={
        total:total,
        moneda:esDolares?'US$':(moneda.toLowerCase().indexOf('sol')>=0?'S/':moneda),
        tc:tc,
        montoSoles:esDolares&&tc?Math.round(total*tc*100)/100:total,
        formaPago:formaPago,
        esCredito:formaPago.toUpperCase().indexOf('CREDITO')>=0,
        usuario:String(gv2(r,iUSR2)||'').trim().toUpperCase(),
        fecha:fecha,
        fechaAprob:fmtFecha2(gv2(r,iFAPRO2)),
        estado:String(gv2(r,iEST2)||'').trim(),
        proveedor:String(gv2(r,iPROV2)||'').trim(),
      };
    });
  }
  return {nConsultas:D.consultas.length, nFin};
}

// ── Guardar consultas.json en GitHub (copiado de compras.html) ───────────────
async function saveConsultasGitHub(token){
  if(!D.consultas||!D.consultas.length)return true; // nada que guardar
  try{
    const rows=D.consultas.map(function(r){return{oc:r.oc,fecha:r.fecha,ruc:r.ruc,proveedor:r.proveedor,moneda:r.moneda,codigo:r.codigo,desc:r.desc,cant:r.cant,um:r.um,valor:r.valor,resp:r.resp,estado:r.estado};});
    const cPayload=JSON.stringify({consultas:rows,ts:Date.now()});
    const GH_CONSULTAS_FILE='consultas.json';
    const GH_CONSULTAS_API='https://api.github.com/repos/'+GH_REPO+'/contents/'+GH_CONSULTAS_FILE;
    let sha='';
    const check=await fetchTO(GH_CONSULTAS_API,{headers:{Authorization:'token '+token,Accept:'application/vnd.github.v3+json'}});
    if(check.ok){const meta=await check.json();sha=meta.sha||'';}
    const content=btoa(unescape(encodeURIComponent(cPayload)));
    const body={message:'consultas '+new Date().toISOString(),content};
    if(sha)body.sha=sha;
    // consultas.json pesa ~20 MB — el timeout genérico de 45s (pensado para data.json,
    // mucho más chico) corta la subida antes de que termine con conexiones normales.
    const res=await fetchTO(GH_CONSULTAS_API,{method:'PUT',headers:{Authorization:'token '+token,Accept:'application/vnd.github.v3+json','Content-Type':'application/json'},body:JSON.stringify(body)},180000);
    if(!res.ok){
      let detalle='';
      try{ detalle=(await res.json()).message||''; }catch(e2){}
      throw new Error('GitHub respondió '+res.status+(detalle?': '+detalle:''));
    }
    return true;
  }catch(e){ console.warn('Consultas GitHub save failed:',e); throw e; }
}

function setupGHToken(){
  const cur=localStorage.getItem(GH_TOKEN_KEY)||'';
  const t=prompt('Ingresa tu GitHub Personal Access Token\n(se guarda solo en este navegador):',cur);
  if(t!==null){
    localStorage.setItem(GH_TOKEN_KEY,t.trim());
    alert(t.trim()?'✅ Token guardado. Ya puedes usar GUARDAR.':'⚠ Token eliminado.');
  }
}

// ── Helper etiqueta de fecha ──────────────────────────────────────────────────

function handleMultiUpload(input){
  if(!input.files||!input.files.length)return;
  const files=Array.from(input.files);
  let processed=0;
  const results=[];
  files.forEach(file=>{
    const name=file.name.toLowerCase();
    const reader=new FileReader();
    reader.onload=function(e){
      try{
        const wb=XLSX.read(e.target.result,{type:'array',cellDates:true});
        if(name.includes('consulta')){
          importConsultas(wb);
          results.push('✅ '+file.name+' (CONSULTAS)');
        } else if(name.includes('pedido')){
          importPedidoSinReq(wb);
          results.push('✅ '+file.name+' (Compras Directas)');
        } else if(name.includes('supervisor')){
          importSupervisores(wb);
          results.push('✅ '+file.name+' (Supervisores)');
        } else if(name.includes('proveedor')){
          importProveedores(wb);
          results.push('✅ '+file.name+' (Proveedores)');
        } else {
          importMaestro(wb);
          results.push('✅ '+file.name+' (MAESTRO)');
        }
      }catch(err){
        results.push('❌ '+file.name+': '+err.message);
      }
      processed++;
      if(processed===files.length){
        alert('Importación completada:\n\n'+results.join('\n'));
      }
    };
    reader.readAsArrayBuffer(file);
  });
  input.value='';
}


function xd(v){
  if(!v && v!==0)return '';
  // Date object
  if(v instanceof Date && !isNaN(v)){
    return v.getDate().toString().padStart(2,'0')+'/'+(v.getMonth()+1).toString().padStart(2,'0')+'/'+v.getFullYear();
  }
  // Serial number (Excel stores dates as numbers since 1900-01-01)
  if(typeof v==='number' && v>30000 && v<60000){
    const d=new Date((v-25569)*86400000);
    if(!isNaN(d))return d.getUTCDate().toString().padStart(2,'0')+'/'+(d.getUTCMonth()+1).toString().padStart(2,'0')+'/'+d.getUTCFullYear();
  }
  const s=String(v).trim();
  // dd/mm/yyyy
  if(/^\d{1,2}\/\d{1,2}\/\d{4}/.test(s))return s;
  // yyyy-mm-dd or yyyy-mm-ddThh:mm:ss
  if(/^\d{4}-\d{2}-\d{2}/.test(s)){
    const p=s.split(/[-T ]/);
    return p[2]+'/'+p[1]+'/'+p[0];
  }
  // m/d/yyyy (US format from SheetJS)
  if(/^\d{1,2}\/\d{1,2}\/\d{4}/.test(s)){
    const p=s.split('/');
    return p[1].padStart(2,'0')+'/'+p[0].padStart(2,'0')+'/'+p[2];
  }
  return '';
}
function xs(v){return(v!=null&&v!==''&&String(v)!=='NaN'&&String(v)!=='undefined'&&String(v)!=='null')?String(v).trim():'';}
function xn(v){const n=parseFloat(v);return isNaN(n)?0:n;}
function is2026(dateStr){return dateStr&&(dateStr.includes('/2026')||dateStr.includes('-2026'));}
function diffBizDays(d1s,d2s){
  const d1=dateObj(d1s),d2=dateObj(d2s);
  if(!d1||!d2||d2<d1)return null;
  let n=0,cur=new Date(d1);
  cur.setDate(cur.getDate()+1);
  while(cur<=d2){const w=cur.getDay();if(w!==0&&w!==6)n++;cur.setDate(cur.getDate()+1);}
  return n>365?null:n;
}
function signedBizDays(d1s,d2s){
  // Positivo = tardío, negativo/cero = a tiempo o adelantado
  const d1=dateObj(d1s),d2=dateObj(d2s);
  if(!d1||!d2)return null;
  if(d2>=d1)return diffBizDays(d1s,d2s);
  const inv=diffBizDays(d2s,d1s);
  return inv===null?null:-inv;
}
function dateObj(ds){
  if(!ds)return null;
  const[d,m,y]=ds.split('/');
  if(!d||!m||!y)return null;
  return new Date(+y,+m-1,+d);
}

// ══════════════════════════════════════════════════════════════════════════════
// IMPORT MAESTRO
// ══════════════════════════════════════════════════════════════════════════════

async function importMaestro(wb){
  const t0=performance.now();

  // ── 1. Read Reporte Maestro sheet ──────────────────────────────────────────
  const wsName=wb.SheetNames.find(n=>n.toLowerCase().includes('reporte'))|| wb.SheetNames[0];
  const ws=wb.Sheets[wsName];
  const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:null});
  
  // Auto-detect header row
  let hdr=0;
  for(let i=0;i<Math.min(10,raw.length);i++){
    const r=raw[i];
    if(r&&typeof r[0]==='string'&&r[0].match(/^Cod|^Id|^Fecha|^Nombre/i)){hdr=i;break;}
  }
  const headers=raw[hdr]||[];
  const ci={};headers.forEach((h,i)=>{if(h)ci[String(h).trim()]=i;});
  const g=(row,name)=>{const i=ci[name];return(i!==undefined)?xs(row[i]):'';};
  const gd=(row,name)=>{const i=ci[name];return(i!==undefined)?xd(row[i]):'';};
  const gn=(row,name)=>{const i=ci[name];return(i!==undefined)?xn(row[i]):0;};
  // Columna de Observación: el header exacto varía entre exportes ("Observación REQ",
  // "Observacion Req", etc.) — se busca por prefijo, sin distinguir mayúsculas/espacios.
  const _obsHdr=Object.keys(ci).find(h=>h.toLowerCase().replace(/\s+/g,'').indexOf('observaci')===0);
  const gObs=(row)=>_obsHdr?g(row,_obsHdr):'';

  const allRows=raw.slice(hdr+1).filter(r=>r&&r.some(c=>c!=null&&c!==''));
  
  // Filter 2026
  const rows=allRows.filter(r=>{
    const idx=ci['FechaRequerimiento'];
    if(idx===undefined)return false;
    const val=r[idx];
    if(val instanceof Date && !isNaN(val))return val.getFullYear()===2026;
    if(typeof val==='number' && val>30000){
      const dt=new Date((val-25569)*86400000);
      return dt.getUTCFullYear()===2026;
    }
    return is2026(xd(val));
  });
  
  console.log('MAESTRO: '+allRows.length+' total, '+rows.length+' rows 2026');
  if(rows.length===0 && allRows.length>0){
    // Debug: show what dates look like
    const sampleIdx=ci['FechaRequerimiento'];
    if(sampleIdx!==undefined){
      const samples=allRows.slice(0,3).map(r=>{const v=r[sampleIdx];return typeof v+': '+String(v);});
      console.log('Date samples: ',samples);
    }
    console.log('Headers found:',Object.keys(ci));
  }

  // ── 2. Build OC data (Page 2) ─────────────────────────────────────────────
  const ocRows=rows.filter(r=>{
    const cod=g(r,'CodOrden');
    return cod&&cod!=='-'&&cod!=='nan';
  });

  // Group by CodOrden — first row is representative
  const ocMap={};
  const ocReqsMap={};
  const skusMap={};
  
  ocRows.forEach(r=>{
    const cod=g(r,'CodOrden');
    if(!ocMap[cod]){
      ocMap[cod]={
        oc:cod,foc:gd(r,'FechaOrden'),fapro:gd(r,'FechaAprobacionOC'),fent:gd(r,'ParaFechaOrden'),
        resp:g(r,'NombreResponsable'),estado:'',_estados:[],
        prov:g(r,'NombreProveedor'),ruc:g(r,'RucProveedor'),
        proy:g(r,'NombreProyecto'),idproy:g(r,'IdProyecto'),
        req:g(r,'CodRequerimiento'),freq:gd(r,'FechaRequerimiento')||gd(r,'Fecha Requerimiento')||gd(r,'FechaReq')||gd(r,'Fecha de Requerimiento'),
        prod:g(r,'Producto'),cod:g(r,'IdProducto'),
        unid:g(r,'Medida')||g(r,'UnidadMedida')||g(r,'Unidad')||g(r,'UM')||g(r,'Um')||g(r,'Und')||g(r,'UndMed')||g(r,'Unidad de Medida')||g(r,'UnidMed')||'',
        nItems:0,cantOrd:0,cantRec:0,cantPend:0,
        frec:gd(r,'FechaRecepcion'),fsal:gd(r,'FechaSalida')||gd(r,'FechaSali'),ucomp:g(r,'UsuarioCompras'),
        ingreso:0,fingreso:'',
        moneda:'',tc:0,_pus:[],montoSoles:0,
        reqs:[],_obss:[],
        ped:g(r,'CodPedido'),peds:[],fpedido:gd(r,'FechaPedido')||gd(r,'Fecha Pedido')||gd(r,'FechaPed')
      };
    }
    const o=ocMap[cod];
    o.nItems++;
    const co=gn(r,'Cantidad de orden'), cr=gn(r,'CantRecepcion');
    o.cantOrd+=co; o.cantRec+=cr;
    o.ingreso+=gn(r,'Ingreso');
    const monR=g(r,'MONEDA');
    if(monR&&monR!=='-'){if(!o.moneda)o.moneda=monR;else if(o.moneda.indexOf(monR)<0)o.moneda+=' / '+monR;}
    o._pus.push(gn(r,'PRECIO UNITARIO SIN IGV'));
    const tcR=gn(r,'TIPO DE CAMBIO');if(tcR&&!o.tc)o.tc=tcR;
    // Monto de la línea en soles (para indicadores financieros): si la moneda de la fila
    // es dólares, convierte con el tipo de cambio de esa misma fila; si no, ya está en soles.
    const puRow=gn(r,'PRECIO UNITARIO SIN IGV');
    const esDolarRow=/DOLAR|DOLLAR|USD|US\$/.test((monR||'').toUpperCase());
    o.montoSoles+=puRow*co*((esDolarRow&&tcR)?tcR:1);
    const fing=gd(r,'FechaIngreso');
    if(fing){
      const dkey=s=>s.split('/').reverse().join('');
      if(!o.fingreso||dkey(fing)>dkey(o.fingreso))o.fingreso=fing;
    }
    const itemEst=g(r,'EstadoOrden');
    if(itemEst&&!o._estados.includes(itemEst))o._estados.push(itemEst);

    // Collect unique reqs and peds
    const req=g(r,'CodRequerimiento');
    if(req&&!o.reqs.includes(req))o.reqs.push(req);
    const ped=g(r,'CodPedido');
    if(ped&&!o.peds.includes(ped))o.peds.push(ped);
    const obsRow=gObs(r);
    if(obsRow&&!o._obss.includes(obsRow))o._obss.push(obsRow);

    // Build SKUS
    if(!skusMap[cod])skusMap[cod]=[];
    skusMap[cod].push({
      cod:g(r,'IdProducto'),prod:g(r,'Producto'),
      cantOrd:Math.round(co),cantRec:Math.round(cr),cantPend:Math.max(0,Math.round(co-cr)),
      foc:gd(r,'FechaOrden'),fent:gd(r,'ParaFechaOrden'),
      estado:g(r,'EstadoOrden'),ucomp:g(r,'UsuarioCompras'),
      prov:g(r,'NombreProveedor'),resp:g(r,'NombreResponsable'),
      idproy:g(r,'IdProyecto'),proy:g(r,'NombreProyecto'),
      req:g(r,'CodRequerimiento'),freq:gd(r,'FechaRequerimiento'),
      frec:gd(r,'FechaRecepcion'),
      moneda:g(r,'MONEDA'),punit:gn(r,'PRECIO UNITARIO SIN IGV'),tc:gn(r,'TIPO DE CAMBIO'),
      unid:g(r,'Medida')||g(r,'UnidadMedida')||g(r,'Unidad')||g(r,'UM')||g(r,'Um')||g(r,'Und')||g(r,'UndMed')||g(r,'Unidad de Medida')||g(r,'UnidMed')||''
    });
  });

  const ocData=Object.values(ocMap).map(o=>{
    o.cantOrd=Math.round(o.cantOrd);o.cantRec=Math.round(o.cantRec);
    o.ingreso=Math.round(o.ingreso);
    o.cantPend=Math.max(0,o.cantOrd-o.cantRec);
    // Derive estado from all item estados
    const es=o._estados;
    if(es.length===1){o.estado=es[0];}
    else if(es.every(e=>e==='Atendido Completo')){o.estado='Atendido Completo';}
    else if(es.every(e=>e==='Aprobado')){o.estado='Aprobado';}
    else if(es.every(e=>e==='Pendiente')){o.estado='Pendiente';}
    else if(es.includes('Atendido Completo')||es.includes('Atendido Parcial')){o.estado='Atendido Parcial';}
    else if(es.includes('Aprobado')&&es.includes('Pendiente')){o.estado='Aprobado';}
    else{o.estado=es[0]||'Pendiente';}
    delete o._estados;
    // Precio unitario sin IGV: un solo valor si todos los items lo comparten; null = "varios"
    const pus=(o._pus||[]).filter(v=>v>0);
    o.punit=pus.length&&pus.every(v=>v===pus[0])?pus[0]:(pus.length?null:0);
    delete o._pus;
    o.obs=(o._obss||[]).join(' | ');
    delete o._obss;
    o.montoSoles=Math.round(o.montoSoles*100)/100;
    return o;
  });
  // Sort desc by foc; si empatan en fecha, desempatar por número de OC descendente
  ocData.sort((a,b)=>{
    const pa=a.foc?a.foc.split('/').reverse().join(''):'';
    const pb=b.foc?b.foc.split('/').reverse().join(''):'';
    if(pa!==pb)return pb.localeCompare(pa);
    const na=parseInt((a.oc||'').replace(/\D/g,''),10)||0;
    const nb=parseInt((b.oc||'').replace(/\D/g,''),10)||0;
    return nb-na;
  });

  // ── 3. Build Sin OC data (Page 3) ─────────────────────────────────────────
  // Incluir filas Pendientes sin OC, y también filas Pendientes CON OC parcial
  // (cuando tiene CodOrden pero aún queda cantidad sin cubrir).
  const sinOcRows=rows.filter(r=>g(r,'EstadoPedido')==='Pendiente');

  const sinOcData=sinOcRows.map(r=>{
    const hasCod=g(r,'CodOrden')&&g(r,'CodOrden')!=='-'&&g(r,'CodOrden')!=='nan';
    const cantPed=Math.round(gn(r,'Cantidad de pedido'));
    const cantArch=Math.round(gn(r,'CantidadArchivadaPedido')||gn(r,'CantidadArchivada')||0);
    const cantOrd=hasCod?Math.round(gn(r,'Cantidad de orden')||gn(r,'CantidadOrden')||0):0;
    const cantFinal=hasCod?Math.max(cantPed-cantArch-cantOrd,0):cantPed;
    // Si tiene OC y el remanente es 0, no aparece en Sin OC
    if(hasCod&&cantFinal===0)return null;
    return {
      ped:g(r,'CodPedido'),
      fecha:gd(r,'FechaPedido'),
      respPed:g(r,'ResponsablePedido'),
      resp:g(r,'NombreResponsable'),
      idproy:g(r,'IdProyecto'),proy:g(r,'NombreProyecto'),
      cod:g(r,'IdProducto'),prod:g(r,'Producto'),
      cant:cantFinal,
      unid:g(r,'Medida')||g(r,'UnidadMedida')||g(r,'Unidad')||'',
      req:g(r,'CodRequerimiento'),
      freq:gd(r,'FechaRequerimiento'),
      estado:g(r,'EstadoPedido'),
      parcial:hasCod,  // tiene OC pero cubre solo una parte
      obs:gObs(r)
    };
  }).filter(Boolean);
  sinOcData.sort((a,b)=>{
    const pa=a.fecha?a.fecha.split('/').reverse().join(''):'';
    const pb=b.fecha?b.fecha.split('/').reverse().join(''):'';
    return pb.localeCompare(pa);
  });

  // ── 4. Build Req sin Pedido (Page 5) ───────────────────────────────────────
  const p5Rows=rows.filter(r=>!!g(r,'EstadoRequerimiento'));
  // El Maestro repite la MISMA línea (req+código+cantidad idéntica) una vez por cada evento de
  // recepción/salida que se registra — esas hay que ignorarlas o se duplicaría la cantidad. Pero
  // un mismo req+código con una cantidad DISTINTA es una línea de requerimiento real aparte (ver
  // caso 0001-0022802: código 001904900016 con cant 2 y cant 1 en filas separadas) y hay que
  // sumarla, no descartarla — antes se perdía silenciosamente quedándose solo con la primera.
  const p5Seen=new Set();     // req+código+cantidad exacta ya contado (repetición de recepción)
  const p5IdxByReqCod={};     // req+código -> índice en p5Data, para sumar cuando aparece otra cant
  const p5Data=[];
  p5Rows.forEach(r=>{
    const req=g(r,'CodRequerimiento'),cod=g(r,'IdProducto');
    const cant=Math.round(gn(r,'CantidadRequerimiento'));
    const key3=req+'|'+cod+'|'+cant;
    if(p5Seen.has(key3))return;
    p5Seen.add(key3);
    const key2=req+'|'+cod;
    if(p5IdxByReqCod[key2]!=null){
      p5Data[p5IdxByReqCod[key2]].cant+=cant;
      return;
    }
    p5IdxByReqCod[key2]=p5Data.length;
    p5Data.push({
      req:req,fecha:gd(r,'FechaRequerimiento'),
      resp:g(r,'NombreResponsable'),idproy:g(r,'IdProyecto'),
      proy:g(r,'NombreProyecto'),cod:cod,
      prod:g(r,'Producto'),cant:cant,
      unid:g(r,'Medida')||g(r,'UnidadMedida')||g(r,'Unidad')||g(r,'UM')||g(r,'Um')||g(r,'Und')||g(r,'UndMed')||g(r,'Unidad de Medida')||g(r,'UnidMed')||'',
      freq:gd(r,'FechaRequerida')||gd(r,'FechaEntrega')||gd(r,'ParaFechaOrden')||gd(r,'Fecha Requerida')||gd(r,'FechaRequerimiento')||'',
      estado:g(r,'EstadoRequerimiento')
    });
  });
  p5Data.sort((a,b)=>{
    const pa=a.fecha?a.fecha.split('/').reverse().join(''):'';
    const pb=b.fecha?b.fecha.split('/').reverse().join(''):'';
    return pb.localeCompare(pa);
  });

  // ── 5. Build Gerencial KPIs (Page 1) ───────────────────────────────────────
  const validReqOC=ocData.filter(o=>{
    const d1=dateObj(o.freq),d2=dateObj(o.foc);
    if(!d1||!d2)return false;
    const diff=Math.round((d2-d1)/(86400000));
    o._dReqOC=diff;
    return diff>=0&&diff<=365;
  });
  const validEntRec=ocData.filter(o=>{
    const d1=dateObj(o.fent),d2=dateObj(o.frec);
    if(!d1||!d2)return false;
    const diff=Math.round((d2-d1)/(86400000));
    o._dEntRec=diff;
    return diff>=-30&&diff<=365;
  });
  
  const avgReqOC=validReqOC.length?Math.round(validReqOC.reduce((s,o)=>s+o._dReqOC,0)/validReqOC.length*10)/10:0;
  const sortedReqOC=validReqOC.map(o=>o._dReqOC).sort((a,b)=>a-b);
  const medReqOC=sortedReqOC.length?sortedReqOC[Math.floor(sortedReqOC.length/2)]:0;
  const avgEntRec=validEntRec.length?Math.round(validEntRec.reduce((s,o)=>s+o._dEntRec,0)/validEntRec.length*10)/10:0;
  const pctOntime=validEntRec.length?Math.round(validEntRec.filter(o=>o._dEntRec<=0).length/validEntRec.length*1000)/10:0;

  // Distribution buckets
  const distReqOC={'0-3':0,'4-7':0,'8-15':0,'>15':0};
  validReqOC.forEach(o=>{const d=o._dReqOC;if(d<=3)distReqOC['0-3']++;else if(d<=7)distReqOC['4-7']++;else if(d<=15)distReqOC['8-15']++;else distReqOC['>15']++;});
  
  const distEntRec={'A tiempo':0,'1-7 días':0,'8-15 días':0,'>15 días':0};
  validEntRec.forEach(o=>{const d=o._dEntRec;if(d<=0)distEntRec['A tiempo']++;else if(d<=7)distEntRec['1-7 días']++;else if(d<=15)distEntRec['8-15 días']++;else distEntRec['>15 días']++;});

  // Monthly trend
  const monthlyReqOC={},monthlyEntRec={},monthlyCnt={},monthlyCnt2={};
  validReqOC.forEach(o=>{
    const m=o.foc?parseInt(o.foc.split('/')[1]):0;
    if(!m)return;
    monthlyReqOC[m]=(monthlyReqOC[m]||0)+o._dReqOC;
    monthlyCnt[m]=(monthlyCnt[m]||0)+1;
  });
  validEntRec.forEach(o=>{
    const m=o.frec?parseInt(o.frec.split('/')[1]):0;
    if(!m)return;
    monthlyEntRec[m]=(monthlyEntRec[m]||0)+o._dEntRec;
    monthlyCnt2[m]=(monthlyCnt2[m]||0)+1;
  });
  const mReqOC={},mEntRec={};
  Object.keys(monthlyReqOC).forEach(m=>{mReqOC[m]=Math.round(monthlyReqOC[m]/monthlyCnt[m]*10)/10;});
  Object.keys(monthlyEntRec).forEach(m=>{mEntRec[m]=Math.round(monthlyEntRec[m]/monthlyCnt2[m]*10)/10;});

  // By responsable
  const respMap={};
  validReqOC.forEach(o=>{
    if(!o.resp)return;
    if(!respMap[o.resp])respMap[o.resp]={sum:0,cnt:0};
    respMap[o.resp].sum+=o._dReqOC;respMap[o.resp].cnt++;
  });
  const byResp=Object.entries(respMap).filter(([,v])=>v.cnt>=3).map(([k,v])=>({resp:k,avg:Math.round(v.sum/v.cnt*10)/10,cnt:v.cnt})).sort((a,b)=>a.avg-b.avg);

  // Top proyectos
  const proyMap={};
  ocData.forEach(o=>{if(o.proy)proyMap[o.proy]=(proyMap[o.proy]||0)+1;});
  const topProy=Object.entries(proyMap).sort((a,b)=>b[1]-a[1]).slice(0,10).map(([k,v])=>({proy:k,cnt:v}));

  // Estado count
  const estadoCnt={};
  ocData.forEach(o=>{estadoCnt[o.estado]=(estadoCnt[o.estado]||0)+1;});

  const gerencial={
    summary:{
      total_oc:ocData.length,
      pendientes:ocData.filter(o=>o.estado==='Pendiente').length,
      parciales:ocData.filter(o=>o.estado==='Atendido Parcial').length,
      completadas:ocData.filter(o=>o.estado==='Atendido Completo').length,
      avg_req_oc:avgReqOC,med_req_oc:medReqOC,
      avg_ent_rec:avgEntRec,pct_ontime:pctOntime,
      total_responsables:new Set(ocData.map(o=>o.resp)).size,
      total_proyectos:new Set(ocData.map(o=>o.proy).filter(Boolean)).size,
    },
    monthly_req_oc:mReqOC,monthly_ent_rec:mEntRec,
    dist_req_oc:distReqOC,dist_ent_rec:distEntRec,
    by_resp:byResp,top_proy:topProy,estado_cnt:estadoCnt
  };

  // ── 6. Read Proyectos sheet ────────────────────────────────────────────────
  const projSheetName=wb.SheetNames.find(n=>n==='Proyectos');
  let projData=D.proj; // keep existing if no sheet
  if(projSheetName){
    const pws=wb.Sheets[projSheetName];
    const praw=XLSX.utils.sheet_to_json(pws,{header:1,defval:null,cellDates:true});
    // Auto-detect header row for Proyectos
    let phdrIdx=0;
    for(let i=0;i<Math.min(5,praw.length);i++){
      if(praw[i]&&typeof praw[i][0]==='string'&&praw[i][0].match(/Proyecto/i)){phdrIdx=i;break;}
    }
    const phdr=praw[phdrIdx]||[];
    const pci={};phdr.forEach((h,i)=>{if(h)pci[String(h).trim()]=i;});
    // case+accent-insensitive lookup helper
    const pciNorm={};Object.keys(pci).forEach(k=>{pciNorm[k.normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase()]=pci[k];});
    const pciGet=(name)=>{const exact=pci[name];if(exact!==undefined)return exact;return pciNorm[name.normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase()];};
    const pg=(row,name)=>{const i=pciGet(name);return(i!==undefined)?xs(row[i]):'';};
    const pgd=(row,name)=>{const i=pciGet(name);return(i!==undefined)?xd(row[i]):'';};

    // Read PRY sheet for ID mapping
    const prySheetName=wb.SheetNames.find(n=>n==='PRY');
    const pryMap={};
    if(prySheetName){
      const prws=wb.Sheets[prySheetName];
      const prraw=XLSX.utils.sheet_to_json(prws,{header:1,defval:null,raw:false});
      let prhdrIdx=0;
      for(let i=0;i<Math.min(5,prraw.length);i++){
        if(prraw[i]&&typeof prraw[i][0]==='string'&&prraw[i][0].match(/Proyecto/i)){prhdrIdx=i;break;}
      }
      const prhdr=prraw[prhdrIdx]||[];
      const prci={};prhdr.forEach((h,i)=>{if(h)prci[String(h).trim()]=i;});
      prraw.slice(prhdrIdx+1).forEach(r=>{
        const nombre=xs(r[prci['Proyectos']]).toUpperCase().trim();
        const pryId=xs(r[prci['PRY']]).trim();
        if(nombre&&pryId){
          if(!pryMap[nombre])pryMap[nombre]=[];
          pryMap[nombre].push(pryId);
        }
      });
    }
    
    const VALID_ESTADOS_PMO=new Set(['en progreso','por iniciar','con obs','pendiente valorización','pendiente valorizacion','cerrado','activo','en curso']);
    projData=praw.slice(phdrIdx+1).filter(r=>r&&r.some(c=>c!=null&&c!=='')).map(r=>{
      const nombre=pg(r,'Proyectos').trim();
      if(!nombre)return null;
      // Skip summary/financial rows by name
      if(/^(supervisor|valor|factur|total|subtotal|igv|monto)/i.test(nombre))return null;
      const estadoRaw=(pg(r,'Estado')||'').trim().toLowerCase();
      // Require a valid PMO state; rows with no state and no supervisor are also summary rows
      if(!estadoRaw&&!pg(r,'SUPERVISOR')&&!pg(r,'ZONA'))return null;
      if(estadoRaw&&!VALID_ESTADOS_PMO.has(estadoRaw))return null;
      // Merge with existing project to preserve dates/GPS already parsed
      const prev=D.proj.find(p=>p.nombre===nombre)||{};
      const newInicio=pgd(r,'Fecha de inicio')||pgd(r,'Fecha Inicio')||pgd(r,'FechaInicio')||pgd(r,'Inicio');
      const newFin=pgd(r,'Fecha de finalización')||pgd(r,'Fecha Fin')||pgd(r,'FechaFin')||pgd(r,'Fin');
      const newGps=pg(r,'UBICACIÓN GPS')||pg(r,'GPS')||pg(r,'Ubicacion GPS')||pg(r,'Ubicación GPS');
      return{
        nombre:nombre,
        prys:pryMap[nombre.toUpperCase()]||(prev.prys||[]),
        estado:pg(r,'Estado')||prev.estado||'',
        supervisor:pg(r,'SUPERVISOR')||prev.supervisor||'',
        asistente:pg(r,'ASITENTE')||pg(r,'ASISTENTE')||prev.asistente||'',
        zona:pg(r,'ZONA')||prev.zona||'',
        distrito:pg(r,'Distrito')||prev.distrito||'',
        depto:pg(r,'Departamento')||prev.depto||'',
        inicio:newInicio||prev.inicio||'',
        fin:newFin||prev.fin||'',
        gps:newGps||prev.gps||'',
        // Preserve parsed GPS coordinates
        lat:prev.lat,lon:prev.lon,gps_exact:prev.gps_exact
      };
    }).filter(Boolean);
  }

  // Para filas de OC Parcial, heredar el responsable de quien colocó la OC
  const reqRespMap={};
  ocData.forEach(o=>{if(o.req)reqRespMap[o.req]={ucomp:o.ucomp,resp:o.resp};});
  sinOcData.forEach(s=>{
    if(s.parcial&&reqRespMap[s.req]){
      s.respPed=reqRespMap[s.req].ucomp||s.respPed; // quien colocó la OC → RESP. PEDIDO
      s.resp=reqRespMap[s.req].resp||s.resp;         // responsable de la OC → RESPONSABLE/SOLICITANTE
    }
  });

  // ── 7. Apply to D and SKUS ─────────────────────────────────────────────────
  D.oc=ocData;
  D.sinoc=sinOcData;
  D.p5reqs=p5Data;

  D.gerencial=gerencial;
  D.proj=projData;
  SKUS=skusMap;

  // ── 8. Re-render everything ────────────────────────────────────────────────
  try{fixOcEstados();}catch(e){}
  saveLocal();

  // Fase 2: oc y p5reqs viven en Supabase (una fila por registro) en vez del blob compartido.
  window._sbErrors=[];
  const okSB=await Promise.all([
    bulkUpsertSB('oc',ocData,'oc',ocToSB),
    bulkUpsertSB('p5reqs',p5Data,'req,cod',p5ToSB),
    replaceAllSB('sinoc',sinOcData,sinocToSB),
    bulkUpsertSB('proj',projData,'nombre',projToSB),
    bulkUpsertSB('oc_items',flattenSkus(skusMap),'oc,seq',itemToSB)
  ]);
  if(okSB.some(function(ok){return !ok;}))alert('⚠️ Parte del MAESTRO no se pudo guardar en Supabase — revisa tu conexión e importa de nuevo.\n\n'+(window._sbErrors||[]).join('\n'));

  const t1=performance.now();
  const msg=`✅ MAESTRO importado (${Math.round(t1-t0)}ms)\n\n` +
    `• ${ocData.length} OC (Pág. OC Pendientes)\n` +
    `• ${sinOcData.length} Sin OC — Cotización\n` +
    `• ${p5Data.length} Req. pendientes (Almacén — Requerimientos/Picking)\n` +
    `• ${projData.length} Proyectos PMO\n` +
    `• ${Object.keys(skusMap).length} SKU detalle`;
  alert(msg);
}

async function importPedidoSinReq(wb){
  const t0=performance.now();
  const ws=wb.Sheets[wb.SheetNames[0]];
  const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:null});
  // Find header row: first row where first cell is a string that looks like a column name
  let hdrIdx=0;
  for(let i=0;i<Math.min(5,raw.length);i++){
    if(raw[i]&&typeof raw[i][0]==='string'&&raw[i][0].match(/^[A-Z]/i)&&!raw[i][0].match(/^\d{4}-/)){hdrIdx=i;break;}
  }
  const headers=raw[hdrIdx]||[];
  const ci={};headers.forEach((h,i)=>{if(h)ci[String(h).trim()]=i;});
  const g=(row,name)=>{const i=ci[name];return(i!==undefined)?xs(row[i]):'';};
  const gd=(row,name)=>{const i=ci[name];return(i!==undefined)?xd(row[i]):'';};
  const gn=(row,name)=>{const i=ci[name];return(i!==undefined)?xn(row[i]):0;};

  const allRows=raw.slice(hdrIdx+1).filter(r=>r&&r.some(c=>c!=null&&c!==''));
  const rows=allRows.filter(r=>{
    const idx=ci['FechaPedido'];
    if(idx===undefined)return false;
    const val=r[idx];
    if(val instanceof Date && !isNaN(val))return val.getFullYear()===2026;
    if(typeof val==='number' && val>30000){
      const dt=new Date((val-25569)*86400000);
      return dt.getUTCFullYear()===2026;
    }
    return is2026(xd(val));
  });
  
  // Filter rows with OC
  const ocRows=rows.filter(r=>{const c=g(r,'CodOrden');return c&&c!=='-'&&c!=='nan';});
  
  const ocMap={}, skusMap4={};
  ocRows.forEach(r=>{
    const cod=g(r,'CodOrden');
    if(!ocMap[cod]){
      ocMap[cod]={
        oc:cod,foc:gd(r,'FechaOrden'),fapro:gd(r,'FechaAprobacionOC')||gd(r,'FechaAprobacion')||gd(r,'Fecha Aprobacion')||gd(r,'FechaAprobacionOrden'),fent:gd(r,'ParaFechaOrden'),
        respPed:g(r,'ResponsablePedido'),estado:'',_estados:[],
        prov:g(r,'NombreProveedor').slice(0,35),ucomp:g(r,'UsuarioCompras'),
        ped:[],fpedido:gd(r,'FechaPedido')||gd(r,'Fecha Pedido')||gd(r,'FechaPed'),
        cod:g(r,'IdProducto')||g(r,'CodigoProducto')||'',
        prod:g(r,'Descripción')||g(r,'Producto')||g(r,'Descripcion')||'',
        unid:g(r,'Medida')||g(r,'Unidad')||'',
        cantOrd:0,cantRec:0,cantPend:0,nItems:0,frec:gd(r,'FechaRecepcion')
      };
    }
    const o=ocMap[cod];
    o.nItems++;
    const co=gn(r,'Cantidad de orden'), cr=gn(r,'CantRecepcion');
    o.cantOrd+=co; o.cantRec+=cr;
    const itemEst4=g(r,'EstadoOrden');
    if(itemEst4&&!o._estados.includes(itemEst4))o._estados.push(itemEst4);
    const ped=g(r,'CodPedido');
    if(ped&&!o.ped.includes(ped))o.ped.push(ped);
    if(!skusMap4[cod])skusMap4[cod]=[];
    skusMap4[cod].push({
      cod:g(r,'IdProducto')||g(r,'CodigoProducto')||'',
      prod:g(r,'Descripción')||g(r,'Producto')||g(r,'Descripcion')||'',
      cantOrd:Math.round(co),cantRec:Math.round(cr),cantPend:Math.max(0,Math.round(co-cr)),
      foc:gd(r,'FechaOrden'),fent:gd(r,'ParaFechaOrden'),
      estado:g(r,'EstadoOrden'),ucomp:g(r,'UsuarioCompras'),
      prov:g(r,'NombreProveedor').slice(0,35),resp:g(r,'ResponsablePedido'),
      ped:ped||'',frec:gd(r,'FechaRecepcion'),
      unid:g(r,'Medida')||g(r,'UnidadMedida')||g(r,'Unidad')||''
    });
  });
  SKUS_P4 = skusMap4;

  const p4Data=Object.values(ocMap).map(o=>{
    o.cantOrd=Math.round(o.cantOrd);o.cantRec=Math.round(o.cantRec);
    o.cantPend=Math.max(0,o.cantOrd-o.cantRec);
    // Derive estado from all item estados (una OC con ítems ya atendidos y otros
    // pendientes debe figurar como "Atendido Parcial", no quedarse con el estado
    // del primer ítem leído del Excel)
    const es4=o._estados;
    if(es4.length===1){o.estado=es4[0];}
    else if(es4.every(e=>e==='Atendido Completo')){o.estado='Atendido Completo';}
    else if(es4.every(e=>e==='Pendiente')){o.estado='Pendiente';}
    else{o.estado='Atendido Parcial';}
    delete o._estados;
    return o;
  });
  p4Data.sort((a,b)=>{
    const pa=a.foc?a.foc.split('/').reverse().join(''):'';
    const pb=b.foc?b.foc.split('/').reverse().join(''):'';
    if(pa!==pb)return pb.localeCompare(pa);
    const na=parseInt((a.oc||'').replace(/\D/g,''),10)||0;
    const nb=parseInt((b.oc||'').replace(/\D/g,''),10)||0;
    return nb-na;
  });

  D.p4=p4Data;
  try{fixOcEstados();}catch(e){}
  saveLocal();

  // Fase 2: p4 vive en Supabase (una fila por registro) en vez del blob compartido.
  window._sbErrors=[];
  const okSB=await Promise.all([
    bulkUpsertSB('p4',p4Data,'oc',p4ToSB),
    bulkUpsertSB('p4_items',flattenSkus(skusMap4),'oc,seq',p4ItemToSB)
  ]);
  if(okSB.some(function(ok){return !ok;}))alert('⚠️ PedidoSinReq no se pudo guardar en Supabase — revisa tu conexión e importa de nuevo.\n\n'+(window._sbErrors||[]).join('\n'));

  const t1=performance.now();
  alert(`✅ PedidoSinReq importado (${Math.round(t1-t0)}ms)\n\n• ${p4Data.length} OC Compras Directas`);
}

// ══════════════════════════════════════════════════════════════════════════════
// IMPORT SUPERVISORES (directorio Nombre → DNI / Correo)
// ══════════════════════════════════════════════════════════════════════════════

async function importSupervisores(wb){
  const t0=performance.now();
  const wsName=wb.SheetNames.find(n=>n.toLowerCase().includes('supervisor'))||wb.SheetNames[0];
  const ws=wb.Sheets[wsName];
  const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:null});
  let hdr=0;
  for(let i=0;i<Math.min(10,raw.length);i++){
    const r=raw[i];
    if(r&&r.some(c=>typeof c==='string'&&/^Nombres?\b/i.test(String(c).trim()))){hdr=i;break;}
  }
  const headers=raw[hdr]||[];
  const ci={};headers.forEach((h,i)=>{if(h)ci[String(h).trim()]=i;});
  const g=(row,name)=>{const i=ci[name];return(i!==undefined)?xs(row[i]):'';};
  // Encabezados flexibles: acepta "Nombre" o "Nombres Completos (…)", "Correo" o
  // "Correo electrónico", "Celular"/"Teléfono"/"Número de teléfono"/"WhatsApp", "Área", "Proyecto"
  const hKeys=Object.keys(ci);
  const hFind=re=>hKeys.find(k=>re.test(k))||'';
  const hNom=hFind(/^Nombres?\b/i);
  const hDni=hFind(/DNI/i);
  const hCor=hFind(/^Correo/i);
  const hCel=hFind(/tel[eé]fono|celular|whatsapp|^cel\b/i);
  const hArea=hFind(/^[aá]rea$/i);
  const hProy=hFind(/^Proyecto/i);
  const supervisores=raw.slice(hdr+1)
    .filter(r=>r&&g(r,hNom))
    .map(r=>{
      const s={nombre:g(r,hNom),dni:hDni?g(r,hDni):'',correo:(hCor?g(r,hCor):'').toLowerCase(),
        celular:(hCel?g(r,hCel):'').replace(/\D/g,'')};
      if(hArea)s.area=g(r,hArea);
      if(hProy)s.proyecto=g(r,hProy);
      return s;
    });
  // Fuente de verdad: tabla supervisores en Supabase (ya no el blob de data.json —
  // ver bug "última escritura gana" documentado arriba en mergeSupervisores).
  const {error:upErr}=await sb.from('supervisores').upsert(supervisores,{onConflict:'nombre'});
  if(upErr){alert('❌ No se pudo guardar el padrón de supervisores en Supabase:\n'+upErr.message);return;}
  D.supervisores=supervisores;
  const t1=performance.now();
  const conCorreo=supervisores.filter(s=>s.correo).length;
  const conCel=supervisores.filter(s=>s.celular&&s.celular!=='-').length;
  alert(`✅ Supervisores importado (${Math.round(t1-t0)}ms)\n\n• ${supervisores.length} personas\n• ${conCorreo} con correo registrado\n• ${conCel} con celular (para WhatsApp)`);
}

// Normaliza un nombre para comparar sin tildes, mayúsculas ni orden estricto
function _normNombre(s){
  var map={'Á':'A','É':'E','Í':'I','Ó':'O','Ú':'U','Ñ':'N','Ü':'U'};
  var up=(s||'').toUpperCase(),out='';
  for(var i=0;i<up.length;i++){out+=map[up[i]]||up[i];}
  return out.replace(/[^A-Z ]/g,' ').replace(/\s+/g,' ').trim();
}
// Busca en D.supervisores un nombre que comparta al menos 2 palabras y ≥50% de
// coincidencia (nombres pueden venir en distinto orden o con nombres de más)
function buscarSupervisor(nombre){
  const nws=new Set(_normNombre(nombre).split(' ').filter(Boolean));
  if(!nws.size||!D.supervisores||!D.supervisores.length)return null;
  let best=null,bestScore=0;
  D.supervisores.forEach(function(s){
    const sws=new Set(_normNombre(s.nombre).split(' ').filter(Boolean));
    if(!sws.size)return;
    let inter=0;nws.forEach(function(w){if(sws.has(w))inter++;});
    const score=inter/Math.max(nws.size,sws.size);
    if(inter>=2&&score>bestScore){bestScore=score;best=s;}
  });
  return (best&&bestScore>=0.5)?best:null;
}

// ══════════════════════════════════════════════════════════════════════════════
// IMPORT PROVEEDORES (directorio Nombre/RUC → Correo)
// ══════════════════════════════════════════════════════════════════════════════
function importProveedores(wb){
  const t0=performance.now();
  const wsName=wb.SheetNames.find(n=>n.toLowerCase().includes('proveedor')&&n.toLowerCase()!=='data')||wb.SheetNames[0];
  const ws=wb.Sheets[wsName];
  const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:null});
  let hdr=0;
  for(let i=0;i<Math.min(10,raw.length);i++){
    const r=raw[i];
    if(r&&r.some(c=>typeof c==='string'&&/^Nombre$/i.test(String(c).trim()))){hdr=i;break;}
  }
  const headers=raw[hdr]||[];
  const ci={};headers.forEach((h,i)=>{if(h)ci[String(h).trim()]=i;});
  const g=(row,name)=>{const i=ci[name];return(i!==undefined)?xs(row[i]):'';};
  const totalFilas=raw.slice(hdr+1).filter(r=>r&&g(r,'Nombre')).length;
  // Solo se guardan los que tienen correo: son los únicos útiles para autocompletar,
  // y guardar el padrón completo (11 mil+) inflaría data.json que ya está al límite.
  const proveedores=raw.slice(hdr+1)
    .filter(r=>r&&g(r,'Nombre')&&g(r,'Email'))
    .map(r=>({nombre:g(r,'Nombre'),ruc:g(r,'RUC'),correo:g(r,'Email').toLowerCase()}));
  D.proveedores=proveedores;
  saveLocal();
  const t1=performance.now();
  alert(`✅ Proveedores importado (${Math.round(t1-t0)}ms)\n\n• ${proveedores.length} proveedores con correo registrado\n• (${totalFilas-proveedores.length} sin correo se omitieron — agrégales correo en el Excel y reimporta)`);
}
// Razones sociales genéricas que no sirven para distinguir empresas entre sí
var _PROV_STOPWORDS={'S':1,'A':1,'C':1,'R':1,'L':1,'SA':1,'SAC':1,'SRL':1,'EIRL':1,'SOCIEDAD':1,'ANONIMA':1,'CERRADA':1,'ABIERTA':1,'EMPRESA':1,'INDIVIDUAL':1,'RESPONSABILIDAD':1,'LIMITADA':1,'COMERCIAL':1,'DEL':1,'DE':1,'LA':1,'EL':1,'Y':1,'PERU':1};
function _provWords(s){
  return _normNombre(s).split(' ').filter(function(w){return w&&!_PROV_STOPWORDS[w];});
}
// Busca un proveedor primero por RUC exacto (confiable) y, si no hay match o no
// hay RUC, por nombre — comparando solo palabras distintivas (sin S.A.C., etc.)
function buscarProveedor(ruc,nombre){
  if(!D.proveedores||!D.proveedores.length)return null;
  const rucLimpio=(ruc||'').trim();
  if(rucLimpio){
    const porRuc=D.proveedores.find(function(p){return p.ruc&&p.ruc.trim()===rucLimpio;});
    if(porRuc)return porRuc;
  }
  const nw=_provWords(nombre);
  if(!nw.length)return null;
  const nws=new Set(nw);
  let best=null,bestScore=0;
  D.proveedores.forEach(function(p){
    const pw=_provWords(p.nombre);
    if(!pw.length)return;
    const pws=new Set(pw);
    let inter=0;nws.forEach(function(w){if(pws.has(w))inter++;});
    // Con pocas palabras distintivas se exige coincidencia total para no
    // confundir empresas de nombre corto; con más, mayoría compartida.
    const minInter=Math.min(nws.size,pws.size)<2?Math.max(nws.size,pws.size):2;
    const score=inter/Math.max(nws.size,pws.size);
    if(inter>=minInter&&score>bestScore){bestScore=score;best=p;}
  });
  return (best&&bestScore>=0.5)?best:null;
}

// ══════════════════════════════════════════════════════════════════════════════
// REFRESH ALL PAGES
// ══════════════════════════════════════════════════════════════════════════════
// Repopulate a <select> preserving the current value if it still exists in new options

/* ==== FIN copiado de compras.html ==== */

function getD(){ return D; }
function getToken(){ return (localStorage.getItem(GH_TOKEN_KEY)||'').trim(); }

return {
  handleMultiUpload,
  importMaestro,
  importPedidoSinReq,
  importSupervisores,
  importProveedores,
  importConsultas,
  saveConsultasGitHub,
  setupGHToken,
  saveGitHub,
  saveLocal,
  getD,
  getToken,
  GH_TOKEN_KEY
};
})();
