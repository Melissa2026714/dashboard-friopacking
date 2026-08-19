// ============================================================
// FACTURAS · lógica de importar Excel + guardar en GitHub, copiada
// tal cual de facturas.html para poder usarse también desde el
// importador centralizado de plataforma.html.
//
// Namespaced en FacturasShared para no chocar con Compras/Almacén
// cuando las 3 conviven en la misma página (el launcher). Este
// bloque es 100% lógica de datos (sin DOM) — se copió sin cambios.
// ============================================================
const FacturasShared = (function(){
let PROVEEDORES_PILOTO = []; // el launcher no importa el Excel de Proveedores del piloto (Fase 1);
                              // aggregateOrders() usa esto solo para el respaldo de días de crédito
                              // — si está vacío, usa la Condición de la OC (que ya tiene prioridad).

/* ==== INICIO: copiado de facturas.html (ver cabecera del archivo) ==== */

const FIELD_ALIASES = {
  descripcion: ['Descripción','Producto'],
  unidad: ['Medida','UnidadMedida'],
  observacion: ['Observación Pedido','Observación REQ']
};

function findHeaderRowIndex(rows){
  for(let i=0;i<Math.min(6,rows.length);i++){
    if(rows[i] && rows[i].includes('CodOrden')) return i;
  }
  return -1;
}

function sheetToRecords(workbook, sheetName, origen){
  const ws = workbook.Sheets[sheetName];
  if(!ws) return [];
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:null});
  const hIdx = findHeaderRowIndex(rows);
  if(hIdx === -1) return [];
  const header = rows[hIdx];
  const colIndex = {};
  header.forEach((h,i)=>{ if(h) colIndex[String(h).trim()] = i; });

  function get(row, name){
    if(colIndex[name] !== undefined) return row[colIndex[name]];
    // try aliases
    for(const key in FIELD_ALIASES){
      if(FIELD_ALIASES[key].includes(name)){
        for(const alias of FIELD_ALIASES[key]){
          if(colIndex[alias] !== undefined && row[colIndex[alias]] != null) return row[colIndex[alias]];
        }
      }
    }
    return null;
  }

  const records = [];
  for(let r=hIdx+1;r<rows.length;r++){
    const row = rows[r];
    if(!row || row.every(c=>c===null||c==='')) continue;
    const codOrden = get(row,'CodOrden');
    if(!codOrden) continue; // sin OC asociada -> no aplica a seguimiento de facturas
    records.push({
      origen,
      codRequerimiento: get(row,'CodRequerimiento') || null,
      codPedido: get(row,'CodPedido'),
      codOrden: String(codOrden).trim(),
      cantidadOrden: Number(get(row,'Cantidad de orden')) || 0,
      fechaOrden: get(row,'FechaOrden'),
      paraFechaOrden: get(row,'ParaFechaOrden'),
      fechaAprobacionOC: get(row,'FechaAprobacionOC'),
      estadoOrden: get(row,'EstadoOrden') || '',
      rucProveedor: get(row,'RucProveedor') ? String(get(row,'RucProveedor')).trim() : '',
      nombreProveedor: get(row,'NombreProveedor') ? String(get(row,'NombreProveedor')).trim() : 'Sin nombre',
      usuarioCompras: get(row,'UsuarioCompras'),
      moneda: get(row,'MONEDA') || '',
      precioUnitario: Number(get(row,'PRECIO UNITARIO SIN IGV')) || 0,
      tipoCambio: Number(get(row,'TIPO DE CAMBIO')) || null,
      registroFactura: get(row,'REGISTRO FACTURA PROVISIÓN'),
      guiaAlmacen: get(row,'GUIA (ALMACÉN)'),
      condiciones: get(row,'Condiciones') || '',
    });
  }
  return records;
}

function excelDateToStr(v){
  if(v==null || v==='') return null;
  if(typeof v === 'number'){
    const d = XLSX.SSF.parse_date_code(v);
    if(!d) return null;
    return `${String(d.d).padStart(2,'0')}/${String(d.m).padStart(2,'0')}/${d.y}`;
  }
  // sometimes already a date string like "2026-02-03 08:39:51"
  const d = new Date(v);
  if(!isNaN(d)) return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
  return String(v);
}
function excelDateToDate(v){
  if(v==null || v==='') return null;
  if(typeof v === 'number'){
    const d = XLSX.SSF.parse_date_code(v);
    if(!d) return null;
    return new Date(d.y, d.m-1, d.d);
  }
  const d = new Date(v);
  return isNaN(d) ? null : d;
}

// Nisira exporta celdas sin dato como '-', '#N/A' o '#N/D' (versión en español del error de
// Excel) según la columna/formato — hay que reconocer las tres, si no una OC sin factura/guía
// real se clasifica como si tuviera documento cargado.
function esMarcadorVacio(v){
  if(v===null || v===undefined) return true;
  const s = String(v).trim().toUpperCase();
  return s==='' || s==='-' || s==='#N/A' || s==='#N/D';
}
function classifyFactura(v){
  if(esMarcadorVacio(v)) return 'sin';
  const s = String(v).trim();
  const prefix = s.split(/\s+/)[0].toUpperCase();
  if(prefix === 'FAC') return 'factura';
  if(['GPF','FND','FAN'].includes(prefix)) return 'provisional';
  return 'provisional';
}

function parseDiasCredito(text){
  if(!text) return null;
  const s = String(text).toUpperCase();
  if(s.includes('CONTADO')) return 0;
  const m = s.match(/(\d+)\s*D[IÍ]AS?/);
  if(m) return parseInt(m[1], 10);
  return null;
}

function sumarDias(fecha, dias){
  if(!fecha || dias==null) return null;
  const d = new Date(fecha.getTime());
  d.setDate(d.getDate() + dias);
  return d;
}

function aggregateOrders(records){
  const map = {};
  records.forEach(rec=>{
    if(!map[rec.codOrden]){
      map[rec.codOrden] = {
        codOrden: rec.codOrden,
        origen: rec.origen,
        codRequerimiento: rec.codRequerimiento,
        codPedido: rec.codPedido,
        rucProveedor: rec.rucProveedor,
        nombreProveedor: rec.nombreProveedor,
        usuarioCompras: rec.usuarioCompras,
        fechaOrden: rec.fechaOrden,
        paraFechaOrden: rec.paraFechaOrden,
        fechaAprobacionOC: rec.fechaAprobacionOC,
        estadoOrden: rec.estadoOrden,
        condiciones: rec.condiciones,
        items: 0,
        montos: {}, // moneda -> total
        facturaRefs: new Set(),
        guiaRefs: new Set(),
        classes: new Set()
      };
    }
    const o = map[rec.codOrden];
    o.items += 1;
    const monto = (rec.precioUnitario||0) * (rec.cantidadOrden||0);
    const moneda = rec.moneda || '—';
    o.montos[moneda] = (o.montos[moneda]||0) + monto;
    if(!esMarcadorVacio(rec.registroFactura)) o.facturaRefs.add(String(rec.registroFactura));
    if(!esMarcadorVacio(rec.guiaAlmacen)) o.guiaRefs.add(String(rec.guiaAlmacen));
    o.classes.add(classifyFactura(rec.registroFactura));
    // keep the most "advanced" estadoOrden text as-is (last wins is fine for display)
  });

  const today = new Date();
  return Object.values(map).map(o=>{
    let statusFactura, responsable, motivoAmarillo=null;
    const tieneGuia = o.guiaRefs.size>0;
    if(String(o.estadoOrden).toLowerCase().includes('anulado')){
      statusFactura = 'anulado';
      responsable = '—';
    } else if(!o.classes.has('sin')){
      // todas las líneas tienen algún registro de factura
      if(o.classes.has('provisional') && !o.classes.has('factura')){
        statusFactura = 'amarillo';
        motivoAmarillo = 'documento';
        responsable = 'Provisión';
      } else {
        statusFactura = 'verde';
        responsable = '—';
      }
    } else if(tieneGuia){
      // no hay factura, pero el proveedor ya entregó (guía registrada por almacén)
      statusFactura = 'amarillo';
      motivoAmarillo = 'entregado';
      responsable = 'Provisión';
    } else {
      // no hay ni guía ni factura: aún depende de Compras hacer seguimiento al proveedor
      statusFactura = 'rojo';
      responsable = 'Compras';
    }
    const fechaLimite = excelDateToDate(o.paraFechaOrden);
    const vencida = fechaLimite && fechaLimite < today;
    // el riesgo real de quedar mal con el proveedor aplica cuando ni siquiera hay guía: ahí depende de Compras
    const riesgoBloqueo = vencida && statusFactura==='rojo';

    // Fecha de vencimiento de pago = fecha de aprobación de la OC + días de crédito otorgado.
    // Prioridad: la condición YA APROBADA para esta OC específica en Nisira (ej. "CREDITO A 7
    // DIAS", tal como sale en el documento oficial de la orden) — es el dato contractual real de
    // esa compra puntual. Solo si esa OC no trae un número de días claro, se usa como respaldo el
    // dato general que el proveedor declaró en el piloto (a veces es un rango condicional, ej.
    // "7, 15 o 30 días según el monto", que no aplica igual a todas sus OCs). Decisión confirmada
    // con Melissa 2026-08-18 tras detectar un caso real donde el dato general (30 días) contradecía
    // la condición ya aprobada de la OC (7 días).
    const infoPiloto = pilotoInfo(o.rucProveedor);
    const diasCredito = (parseDiasCredito(o.condiciones)!=null)
      ? parseDiasCredito(o.condiciones)
      : (infoPiloto ? parseDiasCredito(infoPiloto.diasCredito) : null);
    const fechaAprobacion = excelDateToDate(o.fechaAprobacionOC);
    const fechaVencimientoCredito = (fechaAprobacion && diasCredito!=null) ? sumarDias(fechaAprobacion, diasCredito) : null;
    // Al contado (0 días) no hay vencimiento que rastrear — el pago se hace en el momento, no queda
    // pendiente por definición, así que "pago vencido" no aplica aunque la fecha ya haya pasado.
    const pagoVencido = diasCredito!==0 && fechaVencimientoCredito && fechaVencimientoCredito < today && statusFactura!=='verde' && statusFactura!=='anulado';
    // Contado (0 días) vs Crédito (N días otorgados) — parseDiasCredito ya devuelve 0 exacto para "CONTADO".
    const condicionPagoStr = diasCredito==null ? null : (diasCredito===0 ? 'Contado' : `Crédito ${diasCredito} días`);

    return {
      ...o,
      facturaRefs: Array.from(o.facturaRefs),
      guiaRefs: Array.from(o.guiaRefs),
      classes: Array.from(o.classes),
      statusFactura,
      motivoAmarillo,
      responsable,
      riesgoBloqueo,
      diasCredito,
      condicionPagoStr,
      fechaVencimientoCredito,
      fechaVencimientoCreditoStr: fechaVencimientoCredito ? excelDateToStr(fechaVencimientoCredito) : null,
      pagoVencido,
      fechaOrdenStr: excelDateToStr(o.fechaOrden),
      fechaAprobacionOCStr: excelDateToStr(o.fechaAprobacionOC),
      paraFechaOrdenStr: excelDateToStr(o.paraFechaOrden),
    };
  }).sort((a,b)=> (excelDateToDate(b.fechaOrden)||0) - (excelDateToDate(a.fechaOrden)||0));
}

/* ============================================================
   STORAGE — mismo patrón que Compras/Inventario: un archivo JSON
   propio en el repo de GitHub que ya usa la plataforma, con token
   de escritura guardado solo en el navegador de quien lo ingresa.
   IMPORTANTE: este archivo (facturas_piloto.json) NUNCA debe llevar
   cuenta bancaria ni contactos personales del proveedor (ver
   proveedoresFromWorkbook) — la lectura de GitHub Pages es pública,
   cualquiera con el link puede leerlo aunque no tenga el token.
   ============================================================ */
const GH_REPO = 'Melissa2026714/dashboard-friopacking';
const GH_FILE_FACTURAS = 'facturas_piloto.json';
const GH_RAW_FACTURAS = 'https://raw.githubusercontent.com/'+GH_REPO+'/main/'+GH_FILE_FACTURAS;
const GH_API_FACTURAS = 'https://api.github.com/repos/'+GH_REPO+'/contents/'+GH_FILE_FACTURAS;
const GH_TOKEN_KEY_FACTURAS = 'fp_gh_token_facturas'; // token propio de Facturas, independiente del de Compras/Almacén

// Un token de GitHub es una sola cadena alfanumérica: cualquier espacio/salto de línea colado
// por un copy-paste rompe el header Authorization. Se limpia al guardar y al leer.
function getGHTokenFacturas(){
  const raw = localStorage.getItem(GH_TOKEN_KEY_FACTURAS) || '';
  const clean = raw.replace(/\s+/g,'');
  if(clean !== raw) localStorage.setItem(GH_TOKEN_KEY_FACTURAS, clean);
  return clean;
}
function setupGHTokenFacturas(){
  const cur = localStorage.getItem(GH_TOKEN_KEY_FACTURAS) || '';
  const t = prompt('Ingresa el Token de GitHub de Facturas\n(propio de este módulo, no el de Compras/Almacén — se guarda solo en este navegador):', cur);
  if(t !== null){
    const clean = t.replace(/\s+/g,'');
    localStorage.setItem(GH_TOKEN_KEY_FACTURAS, clean);
    alert(clean ? '✅ Token guardado. Ya puedes usar Guardar Data.' : '⚠ Token eliminado.');
  }
}
// fetch() no tiene timeout propio: con wifi lenta un PUT/GET colgado deja el botón
// en "Guardando..." para siempre sin error visible. A los 45s se corta y se avisa.
function fetchTO(url, opts, ms){
  const ms_ = ms || 45000;
  return Promise.race([
    fetch(url, opts),
    new Promise((_,rej)=>setTimeout(()=>rej(new Error('tiempo de espera agotado ('+Math.round(ms_/1000)+' s) — revisa tu conexión e intenta de nuevo')), ms_))
  ]);
}
async function loadFacturasRemote(){
  try{
    const res = await fetchTO(GH_RAW_FACTURAS + '?t=' + Date.now());
    if(!res.ok) return null; // archivo aún no existe en el repo
    return await res.json();
  }catch(e){
    console.error('No se pudo leer facturas_piloto.json', e);
    return null;
  }
}
// Lee el remoto actual, aplica los cambios de una sección (sin tocar las demás) y sube de
// vuelta — así Maestro+Pedido, Pedido Sin Req y Proveedores se pueden importar por separado
// sin que uno pise lo que subió el otro.
async function saveFacturasSection(sectionKey, sectionValue){
  const token = getGHTokenFacturas();
  if(!token){
    if(confirm('Para guardar (compartido con todo el equipo de Compras y el piloto de proveedores) necesitas un Token de GitHub.\n¿Configurarlo ahora?')) setupGHTokenFacturas();
    throw new Error('Falta configurar el Token de GitHub de Facturas.');
  }
  const check = await fetchTO(GH_API_FACTURAS, {headers:{Authorization:'token '+token, Accept:'application/vnd.github.v3+json'}});
  let sha = '';
  let remote = {};
  if(check.ok){
    const meta = await check.json();
    sha = meta.sha || '';
    const remoteRaw = await fetchTO(GH_RAW_FACTURAS + '?t=' + Date.now());
    if(remoteRaw.ok) remote = await remoteRaw.json();
  }
  remote[sectionKey] = sectionValue;
  const content = btoa(unescape(encodeURIComponent(JSON.stringify(remote))));
  const body = {message:'facturas: actualizar '+sectionKey+' '+new Date().toISOString(), content};
  if(sha) body.sha = sha;
  const res = await fetchTO(GH_API_FACTURAS, {method:'PUT', headers:{Authorization:'token '+token, Accept:'application/vnd.github.v3+json', 'Content-Type':'application/json'}, body:JSON.stringify(body)}, 60000);
  if(!res.ok) throw new Error('GitHub rechazó la subida (HTTP '+res.status+')');
}
async function saveOrders(orders){
  await saveFacturasSection('ordenes', { updatedAt: new Date().toISOString(), orders });
}

async function classifyWorkbook(wb){
  // 1) por nombre de hoja
  const sheetMaestro = wb.SheetNames.find(n=>n.toLowerCase().includes('maestro'));
  if(sheetMaestro) return {tipo:'maestro', sheet:sheetMaestro};
  const sheetPedido = wb.SheetNames.find(n=>n.toLowerCase().includes('pedido'));
  if(sheetPedido) return {tipo:'pedido', sheet:sheetPedido};
  // 2) por encabezados
  for(const name of wb.SheetNames){
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], {header:1, raw:true, defval:null});
    const hIdx = findHeaderRowIndex(rows);
    if(hIdx === -1){
      // hoja de proveedores: encabezado siempre en la primera fila, sin 'CodOrden'
      const row0 = (rows[0]||[]).map(h=>h?String(h).trim():'');
      if(row0.includes('RUC') && row0.some(h=>h.toLowerCase().startsWith('razón social'))){
        return {tipo:'proveedores', sheet:name};
      }
      continue;
    }
    const header = rows[hIdx].map(h=>h?String(h):'');
    if(header.includes('CodRequerimiento')) return {tipo:'maestro', sheet:name};
    if(header.includes('CodOrden')) return {tipo:'pedido', sheet:name};
  }
  return {tipo:null, sheet:wb.SheetNames[0]};
}

/* ==== FIN copiado de facturas.html ==== */

function pilotoInfo(ruc){
  return PROVEEDORES_PILOTO.find(p=>p.ruc===ruc) || null;
}

function getToken(){ return getGHTokenFacturas(); }

return {
  classifyWorkbook,
  sheetToRecords,
  aggregateOrders,
  saveOrders,
  setupGHTokenFacturas,
  getToken,
  GH_TOKEN_KEY_FACTURAS
};
})();
