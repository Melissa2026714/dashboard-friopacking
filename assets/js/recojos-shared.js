// Recojos (OC/OPD que Friopacking recoge en el local del proveedor) — compartido entre
// compras.html (Compras llena la ficha) e inventario.html (Almacén/Katherine la ve en mapa y
// marca Recogido). Vive en la tabla public.recojos de Supabase (ver recojo_setup.sql), así que
// lo que guarda Compras aparece al instante en Almacén, sin pasar por GUARDAR/data.json.
// Requiere el cliente `sb` de Supabase ya creado por la página.

var RECOJO_ESTADOS = ['Falta completar', 'Listo para recoger', 'Recogido', 'Anulado'];
// Centro de Lima (vista inicial del mapa cuando no hay puntos).
var RECOJO_LIMA = [-12.0464, -77.0428];

function recojoEsc(v) {
  return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function recojoHoyISO() { return new Date().toLocaleDateString('en-CA'); }
function recojoFecha(d) {
  if (!d) return '—';
  var p = String(d).slice(0, 10).split('-');
  return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : d;
}
// Días que faltan hasta la fecha límite (negativo = vencida).
function recojoDiasRestantes(fLimite) {
  if (!fLimite) return null;
  return Math.round((new Date(String(fLimite).slice(0, 10)) - new Date(recojoHoyISO())) / 86400000);
}
// Semáforo de la fecha límite: rojo vencida, ámbar hoy o en ≤2 días, verde el resto.
function recojoSemaforo(r) {
  if (r.estado === 'Recogido' || r.estado === 'Anulado') return { txt: '—', col: '#94a3b8' };
  var d = recojoDiasRestantes(r.f_limite);
  if (d == null) return { txt: 'sin fecha límite', col: '#b45309' };
  if (d < 0) return { txt: 'vencida hace ' + (-d) + ' d', col: '#dc2626' };
  if (d === 0) return { txt: 'vence HOY', col: '#dc2626' };
  if (d <= 2) return { txt: 'vence en ' + d + ' d', col: '#d97706' };
  return { txt: 'quedan ' + d + ' d', col: '#16a34a' };
}
// Con dirección, contacto y fecha límite la ficha está lista para Almacén.
function recojoFaltantes(r) {
  var f = [];
  if (!r.direccion) f.push('dirección');
  if (!r.contacto && !r.celular) f.push('contacto o celular');
  if (!r.f_limite) f.push('fecha límite');
  return f;
}
function recojoEstadoSegunDatos(r) {
  if (r.estado === 'Recogido' || r.estado === 'Anulado') return r.estado;
  return recojoFaltantes(r).length ? 'Falta completar' : 'Listo para recoger';
}
// Coordenadas desde un link o texto de Google Maps. Los links cortos (maps.app.goo.gl) no
// traen coordenadas y no se pueden abrir desde el navegador por CORS: ahí se marca con clic.
function recojoParseMaps(txt) {
  txt = String(txt || '').trim();
  if (!txt) return null;
  var pats = [
    /!3d(-?\d{1,2}\.\d+)!4d(-?\d{1,3}\.\d+)/,
    /@(-?\d{1,2}\.\d+),\s*(-?\d{1,3}\.\d+)/,
    /[?&](?:q|query|ll|destination|daddr|center)=(-?\d{1,2}\.\d+)(?:,|%2C)\s*(-?\d{1,3}\.\d+)/i,
    /^\s*(-?\d{1,2}\.\d+)\s*,\s*(-?\d{1,3}\.\d+)\s*$/
  ];
  for (var i = 0; i < pats.length; i++) {
    var m = txt.match(pats[i]);
    if (m) {
      var lat = parseFloat(m[1]), lon = parseFloat(m[2]);
      if (Math.abs(lat) <= 90 && Math.abs(lon) <= 180) return { lat: lat, lon: lon };
    }
  }
  return null;
}
function recojoMapsHref(r) {
  if (r.maps_url && /^https?:\/\//i.test(r.maps_url)) return r.maps_url;
  if (r.lat != null && r.lon != null) return 'https://www.google.com/maps/search/?api=1&query=' + r.lat + ',' + r.lon;
  if (r.direccion) return 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent([r.direccion, r.distrito, 'Perú'].filter(Boolean).join(', '));
  return '';
}
function recojoWaHref(cel, texto) {
  var d = String(cel || '').replace(/\D/g, '');
  if (d.length === 9) d = '51' + d;
  if (d.length < 10) return '';
  return 'https://wa.me/' + d + (texto ? '?text=' + encodeURIComponent(texto) : '');
}
async function recojosCargarSB() {
  var r = await sb.from('recojos').select('*');
  if (r.error) {
    var msg = r.error.message || '';
    return {
      rows: [],
      error: /relation|does not exist|schema cache/i.test(msg)
        ? 'Falta crear la tabla <b>recojos</b> en Supabase: ejecuta <b>recojo_setup.sql</b> una vez en el SQL Editor.'
        : 'No se pudieron leer los recojos: ' + recojoEsc(msg)
    };
  }
  return { rows: r.data || [], error: '' };
}
