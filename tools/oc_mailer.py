# Servidor local FRIOPACKING — Envío OC/Cotizaciones por Outlook + subida automática a SharePoint
# Puerto 5100 — ejecutar con: python oc_mailer.py

from flask import Flask, request, jsonify
from flask_cors import CORS
import pythoncom
import win32com.client
import threading
import os
import re
import shutil
import time
from pathlib import Path

# Carpeta de descargas del usuario
DOWNLOADS = str(Path.home() / "Downloads")

# Auto-detectar la carpeta SharePoint sincronizada en cualquier PC
def _detectar_sp_local():
    HOME = Path.home()
    SUBCARPETA      = Path("Piero Linares Palomino - Trazabilidad Compras 2026") / "COMPARATIVO_COTIZACIONES_COMPRAS"
    SUBCARPETA_ARC  = Path("Archivos de Piero Linares Palomino - Trazabilidad Compras 2026") / "COMPARATIVO_COTIZACIONES_COMPRAS"
    candidatos = [
        HOME / "OneDrive - GRUPO FRIOPACKING" / SUBCARPETA_ARC,   # ✅ ruta correcta
        HOME / "OneDrive - GRUPO FRIOPACKING" / SUBCARPETA,
        HOME / "OneDrive"                    / "GRUPO FRIOPACKING" / SUBCARPETA_ARC,
        HOME / "OneDrive"                    / "GRUPO FRIOPACKING" / SUBCARPETA,
        HOME / "GRUPO FRIOPACKING"           / SUBCARPETA,
        HOME / "Friopacking"                 / SUBCARPETA,
    ]
    for c in candidatos:
        if c.exists():
            return str(c)
    return None

SP_LOCAL = _detectar_sp_local()

# Carpeta local de PDFs de Órdenes de Compra
def _detectar_oc_folder():
    Desktop = Path.home() / "Desktop"
    candidatos = [
        Desktop / "Melissa - FrioPacking 2026" / "Órdenes de Compra",
        Desktop / "Melissa - FrioPacking 2026" / "Ordenes de Compra",
        Desktop / "Órdenes de Compra",
        Desktop / "Ordenes de Compra",
    ]
    for c in candidatos:
        if c.exists():
            return str(c)
    return None

OC_FOLDER = _detectar_oc_folder()

# Carpetas donde Melissa guarda cotizaciones y facturas, nombrando el archivo con la OC
# a la que pertenecen (ej. "0001-0011557 Factura Proveedor X.pdf") para poder ubicarlas
# automáticamente al armar la solicitud de pago.
def _detectar_carpeta(nombre):
    Desktop = Path.home() / "Desktop"
    c = Desktop / "Melissa - FrioPacking 2026" / nombre
    return str(c) if c.exists() else None

COT_FOLDER  = _detectar_carpeta("COTIZACIONES")
FACT_FOLDER = _detectar_carpeta("FACTURA")
COMP_FOLDER = _detectar_carpeta("COMPROBANTES DE PAGO")
# OPD (Orden de Pago): la emite Provisión/Tesorería fuera de Nisira Compras. Se guarda en
# la carpeta OPD con su número en el nombre (ej. "ORDEN DE PAGO 0012413.pdf") y su cotización
# en COTIZACIONES con el mismo rótulo (ej. "COTIZACION 221-2026 ... ORDEN DE PAGO 0012413.pdf").
OPD_FOLDER  = _detectar_carpeta("OPD")

# "ORDEN DE PAGO 0012413", "OPD 12413", "OPD-0012413"…
_RE_OPD_NOMBRE = re.compile(r'(?:ORDEN\s*DE\s*PAGO|OPD)\s*(?:N\s*[°º.]?)?\s*[-_#:]?\s*0*(\d{1,9})(?!\d)', re.I)

def _normalizar_opd(numero):
    """'12413' / '0012413' / 'OPD 0012413' → '0012413' (7 dígitos, como viene en el PDF)."""
    dig = re.sub(r'\D', '', str(numero or ''))
    return (dig.lstrip('0') or '0').zfill(7) if dig else ''

def _es_archivo_opd(fname):
    return bool(_RE_OPD_NOMBRE.search(fname))

def _buscar_archivos_opd(carpeta, numero, max_files=5):
    """Archivos cuyo nombre dice ORDEN DE PAGO / OPD + número. En la carpeta OPD basta con el
    número; en COTIZACIONES se exige el rótulo para no confundirlo con una OC 0001-00XXXXX."""
    num = _normalizar_opd(numero)
    if not carpeta or not num:
        return []
    corto = num.lstrip('0') or '0'
    exts = ('.pdf', '.xlsx', '.xls', '.jpg', '.jpeg', '.png', '.docx')
    solo_carpeta_opd = (carpeta == OPD_FOLDER)
    encontrados = []
    for fname in sorted(os.listdir(carpeta)):
        if not fname.lower().endswith(exts):
            continue
        ok = any((m.group(1).lstrip('0') or '0') == corto for m in _RE_OPD_NOMBRE.finditer(fname))
        if not ok and solo_carpeta_opd:
            ok = bool(re.search(r'(?<!\d)0*' + corto + r'(?!\d)', fname))
        if ok:
            encontrados.append(os.path.join(carpeta, fname))
            if len(encontrados) >= max_files:
                break
    return encontrados

def _leer_datos_opd(fpath):
    """Lee proveedor, RUC, área, total, moneda y forma de pago del PDF de la OPD (para autollenar)."""
    try:
        from pypdf import PdfReader
    except ImportError:
        return {}
    try:
        txt = '\n'.join((pg.extract_text() or '') for pg in PdfReader(fpath).pages)
    except Exception as e:
        print(f'[OPD] No se pudo leer {fpath}: {e}', flush=True)
        return {}
    def _m(rx):
        m = re.search(rx, txt, re.I)
        return m.group(1).strip() if m else ''
    moneda = 'USD' if re.search(r'US\$|DOLARES', txt, re.I) else ('PEN' if re.search(r'S/|SOLES', txt, re.I) else '')
    # "Ica, 13 de Octubre del 2026" → "2026-10-13"
    fecha = ''
    mf = re.search(r'(\d{1,2})\s+de\s+([a-záéíóú]+)\s+del?\s+(\d{4})', txt, re.I)
    meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto',
             'setiembre', 'octubre', 'noviembre', 'diciembre']
    if mf:
        mes = mf.group(2).lower().replace('septiembre', 'setiembre')
        if mes in meses:
            fecha = f'{mf.group(3)}-{meses.index(mes) + 1:02d}-{int(mf.group(1)):02d}'
    return {
        'fecha':     fecha,
        'prov':      _m(r'Proveedor\s*/\s*Responsable\s*:?\s*([^\n]+)'),
        'ruc':       _m(r'DNI\s*/\s*RUC\s*:?\s*(\d{8,11})'),
        'area':      _m(r'rea\s+Solicitante\s*:?\s*([^\n]+)'),
        'formaPago': _m(r'Forma\s+de\s+pago\s*:?\s*([^:\n]+)'),
        'total':     _m(r'([\d,]+\.\d{2})\s*(?:US\$|S/)'),
        'moneda':    moneda,
    }

def _buscar_pdf_oc(numero_oc):
    """Busca el PDF de la OC en OC_FOLDER. numero_oc ej: '0001-0011557'"""
    if not OC_FOLDER or not numero_oc:
        return None
    partes     = numero_oc.split('-')
    num_limpio = numero_oc.replace('-', ' ').strip()   # "0001 0011557"
    solo_num   = partes[-1]                            # "0011557" (7 dígitos con ceros)
    ultimo     = solo_num.lstrip('0') or '0'           # "11557"  (sin ceros, para archivos formato antiguo)
    for fname in os.listdir(OC_FOLDER):
        fl = fname.lower()
        if not fl.endswith('.pdf') and not fl.endswith('.xlsx'):
            continue
        if num_limpio.lower() in fl or solo_num in fl or ultimo in fl:
            return os.path.join(OC_FOLDER, fname)
    return None

def _buscar_archivos_oc(carpeta, numero_oc, max_files=5):
    """Busca todos los archivos de una carpeta (cotización o factura) cuyo nombre
    contenga el número de OC. Acepta PDF, Excel e imágenes (facturas escaneadas)."""
    if not carpeta or not numero_oc:
        return []
    partes     = numero_oc.split('-')
    num_limpio = numero_oc.replace('-', ' ').strip()
    solo_num   = partes[-1]
    ultimo     = solo_num.lstrip('0') or '0'
    exts = ('.pdf', '.xlsx', '.xls', '.jpg', '.jpeg', '.png', '.docx')
    encontrados = []
    for fname in os.listdir(carpeta):
        fl = fname.lower()
        if not fl.endswith(exts):
            continue
        if _es_archivo_opd(fname):
            continue
        if num_limpio.lower() in fl or solo_num in fl or ultimo in fl:
            encontrados.append(os.path.join(carpeta, fname))
            if len(encontrados) >= max_files:
                break
    return encontrados

# Nombre SUNAT del comprobante electrónico: RUC-TIPO-SERIE-NUMERO (ej. "20607478440-01-F002-00000449").
# Tipo 01 = factura, 03 = boleta. Si no viene con RUC, se acepta una serie suelta tipo "F002-00000449".
_RE_CPE_SUNAT = re.compile(r'(?<!\d)\d{11}-(?:01|03)-([A-Z][A-Z0-9]{3})-(\d{1,8})(?!\d)', re.I)
_RE_CPE_SERIE = re.compile(r'(?<![A-Z0-9])([FBE][A-Z0-9]{3})-(\d{1,8})(?!\d)', re.I)
# Portal SUNAT (SOL): "PDF-DOC-E001-195120537220245" = serie E001, número 1951 pegado al RUC del emisor (11 díg.)
_RE_CPE_SOL   = re.compile(r'(?<![A-Z0-9])(E\d{3})-(\d{12,19})(?!\d)', re.I)

def _numero_factura_oc(numero_oc):
    """Devuelve el N° de factura (ej. 'F002-00000449') leído del nombre del archivo en FACTURA."""
    for fp in _buscar_archivos_oc(FACT_FOLDER, numero_oc):
        nombre = os.path.basename(fp)
        m = _RE_CPE_SUNAT.search(nombre) or _RE_CPE_SERIE.search(nombre)
        if m:
            return f'{m.group(1).upper()}-{m.group(2)}', nombre
        m = _RE_CPE_SOL.search(nombre)
        if m:
            return f'{m.group(1).upper()}-{m.group(2)[:-11]}', nombre
    return None, None

# ──────────────────────────────────────────────────────────────────────────────

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

_lock       = threading.Lock()
_seen_files = set()

# ── Headers requeridos por Chrome Private Network Access ──────────────────────
@app.after_request
def add_headers(response):
    response.headers['Access-Control-Allow-Origin']          = '*'
    response.headers['Access-Control-Allow-Private-Network'] = 'true'
    response.headers['Access-Control-Allow-Headers']         = 'Content-Type, Access-Control-Request-Private-Network'
    response.headers['Access-Control-Allow-Methods']         = 'GET, POST, OPTIONS'
    return response

@app.route('/ping',     methods=['OPTIONS'])
@app.route('/send-oc',  methods=['OPTIONS'])
@app.route('/send-cot', methods=['OPTIONS'])
@app.route('/factura-oc', methods=['OPTIONS'])
@app.route('/opd-info', methods=['OPTIONS'])
@app.route('/guardar-opd', methods=['OPTIONS'])
def handle_options():
    return ('', 204)

# ── N° de factura de una OC (para la solicitud de pago) ───────────────────────
@app.route('/factura-oc', methods=['GET'])
def factura_oc():
    numero_oc = (request.args.get('oc') or '').strip()
    factura, archivo = _numero_factura_oc(numero_oc)
    print(f'[FACTURA-OC] oc={numero_oc}  factura={factura or "—"}  archivo={archivo or "—"}', flush=True)
    return jsonify({'ok': bool(factura), 'factura': factura, 'archivo': archivo})

# ── Datos de una OPD (para autollenar el formulario del dashboard) ────────────
@app.route('/opd-info', methods=['GET'])
def opd_info():
    numero = _normalizar_opd(request.args.get('numero'))
    opds = _buscar_archivos_opd(OPD_FOLDER, numero, max_files=1)
    cots = _buscar_archivos_opd(COT_FOLDER, numero)
    datos = _leer_datos_opd(opds[0]) if opds and opds[0].lower().endswith('.pdf') else {}
    print(f'[OPD-INFO] opd={numero}  pdf={os.path.basename(opds[0]) if opds else "—"}  cotizaciones={len(cots)}', flush=True)
    return jsonify({'ok': bool(opds), 'numero': numero,
                    'archivo': os.path.basename(opds[0]) if opds else None,
                    'cotizaciones': [os.path.basename(c) for c in cots],
                    'carpetaOPD': bool(OPD_FOLDER), **datos})

# ── Guardar OPD/cotización adjuntadas en el dashboard ─────────────────────────
# El dashboard lee el PDF en el navegador y manda aquí los archivos para que queden en
# las carpetas de siempre con el N° de OPD en el nombre, y así /send-oc los adjunte.
# En la etapa de solicitud todavía no hay N° de OPD: la cotización se rotula con el código
# de la solicitud (ej. "SOL-OPD-2026-001") y, cuando llega la OPD, se renombra cambiando ese
# código por "ORDEN DE PAGO <n°>".
_RE_COD_SOL = re.compile(r'^SOL-OPD-\d{4}-\d{3,5}$', re.I)

def _buscar_archivos_codigo(carpeta, codigo, max_files=5):
    if not carpeta or not codigo:
        return []
    cod = codigo.upper()
    return [os.path.join(carpeta, f) for f in sorted(os.listdir(carpeta)) if cod in f.upper()][:max_files]

def _nombre_seguro(nombre):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', nombre).strip(' .') or 'archivo'

@app.route('/guardar-opd', methods=['POST'])
def guardar_opd():
    import base64
    data = request.json or {}
    numero = _normalizar_opd(data.get('numero'))
    codigo = str(data.get('codigo') or '').strip().upper()
    if codigo and not _RE_COD_SOL.match(codigo):
        return jsonify({'ok': False, 'error': f'Código de solicitud inválido: {codigo}'}), 400
    if not numero and not codigo:
        return jsonify({'ok': False, 'error': 'Falta el N° de OPD o el código de solicitud'}), 400
    global OPD_FOLDER
    if not OPD_FOLDER:
        base = Path.home() / "Desktop" / "Melissa - FrioPacking 2026"
        if base.exists():
            (base / "OPD").mkdir(exist_ok=True)
            OPD_FOLDER = str(base / "OPD")
    guardados, errores, renombrados = [], [], []
    # Llegó la OPD de una solicitud: la cotización rotulada con el código pasa a llevar el N°.
    if numero and codigo and COT_FOLDER:
        for fp in _buscar_archivos_codigo(COT_FOLDER, codigo):
            fname = os.path.basename(fp)
            nuevo = re.sub(re.escape(codigo), f'ORDEN DE PAGO {numero}', fname, flags=re.I)
            destino = os.path.join(COT_FOLDER, nuevo)
            try:
                if os.path.exists(destino):
                    raise FileExistsError('ya existe un archivo con ese nombre')
                os.rename(fp, destino)
                renombrados.append(nuevo)
            except Exception as e:
                errores.append(f'No se pudo renombrar {fname}: {e}')
    for clave, carpeta in (('opd', OPD_FOLDER), ('cot', COT_FOLDER)):
        f = data.get(clave)
        if not f:
            continue
        if clave == 'opd' and not numero:
            errores.append('Falta el N° de OPD para guardar el PDF de la OPD')
            continue
        if not carpeta:
            errores.append(f'Carpeta {"OPD" if clave == "opd" else "COTIZACIONES"} no detectada')
            continue
        stem, ext = os.path.splitext(_nombre_seguro(f.get('nombre') or ''))
        ext = ext.lower() or '.pdf'
        if clave == 'opd':
            nombre = f'ORDEN DE PAGO {numero}{ext}'
        elif not numero:
            nombre = f'{stem}{ext}' if codigo in stem.upper() else f'{stem} {codigo}{ext}'
        else:
            rotulada = any((m.group(1).lstrip('0') or '0') == numero.lstrip('0') for m in _RE_OPD_NOMBRE.finditer(stem))
            nombre = f'{stem}{ext}' if rotulada else f'{stem} ORDEN DE PAGO {numero}{ext}'
        try:
            contenido = base64.b64decode(f.get('b64') or '')
            if not contenido:
                raise ValueError('archivo vacío')
            with open(os.path.join(carpeta, nombre), 'wb') as fh:
                fh.write(contenido)
            guardados.append(nombre)
        except Exception as e:
            errores.append(f'{nombre}: {e}')
            continue
        try:
            print(f'[GUARDAR-OPD] guardado {nombre} en {carpeta}', flush=True)
        except Exception:
            pass
    return jsonify({'ok': not errores, 'guardados': guardados, 'renombrados': renombrados, 'errores': errores})

# ── Ping ───────────────────────────────────────────────────────────────────────
@app.route('/ping', methods=['GET'])
def ping():
    print(f'[PING] origen={request.headers.get("Origin","—")}', flush=True)
    return jsonify({'ok': True, 'servicio': 'OC Mailer — FRIOPACKING'})

# ── Envío OC por Outlook ───────────────────────────────────────────────────────
@app.route('/send-oc', methods=['POST'])
def send_oc():
    data      = request.json or {}
    to_addr   = data.get('to', '')
    cc_addr   = data.get('cc', '')
    subject   = data.get('subject', '')
    html_body = data.get('htmlBody', '')
    auto_send = data.get('autoSend', False)
    modality  = data.get('modality', '')
    incluir_doc = data.get('incluirFacturaCot', False)
    incluir_factura = data.get('incluirFactura', False)
    es_opd    = data.get('tipo') == 'opd'
    es_sol    = data.get('tipo') == 'sol-opd'
    codigo    = str(data.get('codigo') or '').strip().upper()

    # Extraer número de OC del asunto: "Orden de Compra Nº 0001-0011557 — ..."
    match_oc = re.search(r'(\d{4}-\d{7})', subject)
    numero_oc = _normalizar_opd(data.get('numero')) if es_opd else (match_oc.group(1) if match_oc else None)
    print(f'[SEND-OC] asunto={subject[:60] if subject else "—"}  oc={numero_oc}  modality={repr(modality)}  len={len(html_body)}', flush=True)
    if not subject:
        return jsonify({'ok': False, 'error': 'Falta el asunto'}), 400

    try:
        with _lock:
            pythoncom.CoInitialize()
            try:
                outlook = win32com.client.Dispatch('Outlook.Application')
                mail = outlook.CreateItem(0)
                mail.To      = to_addr
                mail.CC      = cc_addr
                mail.Subject = subject
                cuerpo = (
                    '<div style="font-family:Calibri,Arial,sans-serif;'
                    'font-size:16pt;line-height:1.9;color:#1e293b;max-width:760px;padding:8px 4px">'
                    + html_body.replace('<p>', '<p style="margin:12px 0;font-size:16pt;font-family:Calibri,Arial,sans-serif">')
                               .replace('<li>', '<li style="margin:5px 0;font-size:16pt;font-family:Calibri,Arial,sans-serif">')
                    + '</div>'
                )
                if auto_send:
                    mail.HTMLBody = cuerpo
                    mail.Send()
                    accion = 'enviado'
                else:
                    _ = mail.GetInspector
                    mail.HTMLBody = cuerpo
                    # OPD: adjuntar la OPD (carpeta OPD) y su cotización (COTIZACIONES, rotulada
                    # con el N° de OPD). A diferencia de la OC, aquí la cotización sí va al proveedor.
                    if es_sol:
                        archivos = _buscar_archivos_codigo(COT_FOLDER, codigo)
                        for fp in archivos:
                            mail.Attachments.Add(fp)
                            print(f'[SEND-OC] 📎 Cotización adjuntada (solicitud {codigo}): {os.path.basename(fp)}', flush=True)
                        if not archivos:
                            print(f'[SEND-OC] ⚠️  No se encontró cotización con "{codigo}" en {COT_FOLDER or "(carpeta no detectada)"}', flush=True)
                    if es_opd:
                        for carpeta, etiqueta in ((OPD_FOLDER, 'OPD'), (COT_FOLDER, 'Cotización')):
                            archivos = _buscar_archivos_opd(carpeta, numero_oc)
                            for fp in archivos:
                                mail.Attachments.Add(fp)
                                print(f'[SEND-OC] 📎 {etiqueta} adjuntada: {os.path.basename(fp)}', flush=True)
                            if not archivos:
                                print(f'[SEND-OC] ⚠️  No se encontró {etiqueta} para OPD {numero_oc} en {carpeta or "(carpeta no detectada)"}', flush=True)
                    # Adjuntar PDF de la OC si se encuentra en la carpeta
                    pdf_oc = None if (es_opd or es_sol) else _buscar_pdf_oc(numero_oc)
                    if es_opd or es_sol:
                        pass
                    elif pdf_oc:
                        mail.Attachments.Add(pdf_oc)
                        print(f'[SEND-OC] 📎 OC adjuntada: {os.path.basename(pdf_oc)}', flush=True)
                    else:
                        print(f'[SEND-OC] ⚠️  No se encontró PDF para OC {numero_oc}', flush=True)
                    # Adjuntar cotización, factura y/o comprobante de pago si el dashboard lo
                    # pidió, buscando por número de OC en las carpetas COTIZACIONES, FACTURA y
                    # COMPROBANTES DE PAGO. La solicitud de pago pide las tres (incluirFacturaCot);
                    # el correo de OC al proveedor solo pide factura y comprobante, por si ya
                    # existen (ej. regularización con el proveedor) — la cotización es información
                    # interna y no debe llegar al proveedor.
                    if es_opd or es_sol:
                        carpetas_chk = ()
                    elif incluir_doc:
                        carpetas_chk = ((COT_FOLDER, 'Cotización'), (FACT_FOLDER, 'Factura'), (COMP_FOLDER, 'Comprobante de pago'))
                    elif incluir_factura:
                        carpetas_chk = ((FACT_FOLDER, 'Factura'), (COMP_FOLDER, 'Comprobante de pago'))
                    else:
                        carpetas_chk = ()
                    for carpeta, etiqueta in carpetas_chk:
                        archivos = _buscar_archivos_oc(carpeta, numero_oc)
                        if archivos:
                            for fp in archivos:
                                mail.Attachments.Add(fp)
                                print(f'[SEND-OC] 📎 {etiqueta} adjuntada: {os.path.basename(fp)}', flush=True)
                        else:
                            print(f'[SEND-OC] ⚠️  No se encontró {etiqueta} para OC {numero_oc} en {carpeta or "(carpeta no detectada)"}', flush=True)
                    # Adjuntar CARTILLA SSOMA si es entrega en Lurín
                    # Detecta por modality (versión nueva del dashboard) O por contenido del cuerpo (versión online)
                    es_lurin = not es_sol and ((str(modality or '').lower() == 'planta') or ('lurín' in html_body.lower()) or ('lurin' in html_body.lower()))
                    if es_lurin:
                        cartilla = os.path.join(DOWNLOADS, 'CARTILLA SSOMA FP.pdf')
                        if os.path.exists(cartilla):
                            mail.Attachments.Add(cartilla)
                            print(f'[SEND-OC] 📎 Cartilla SSOMA adjuntada (entrega Lurín)', flush=True)
                        else:
                            print(f'[SEND-OC] ⚠️  No se encontró CARTILLA SSOMA FP.pdf en Descargas', flush=True)
                    mail.Display(False)
                    accion = 'abierto en Outlook'
                return jsonify({'ok': True, 'accion': accion})
            finally:
                pythoncom.CoUninitialize()
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

# ── Envío Cotización por Outlook ───────────────────────────────────────────────
@app.route('/send-cot', methods=['POST'])
def send_cot():
    data      = request.json or {}
    to_addr   = data.get('to', '')
    cc_addr   = data.get('cc', '')
    bcc_addr  = data.get('bcc', '')
    subject   = data.get('subject', '')
    html_body = data.get('htmlBody', '')

    print(f'[SEND-COT] asunto={subject[:60] if subject else "—"}  len={len(html_body)}', flush=True)
    if not subject:
        return jsonify({'ok': False, 'error': 'Falta el asunto'}), 400

    try:
        with _lock:
            pythoncom.CoInitialize()
            try:
                outlook = win32com.client.Dispatch('Outlook.Application')
                mail = outlook.CreateItem(0)
                mail.To      = to_addr
                mail.CC      = cc_addr
                mail.BCC     = bcc_addr
                mail.Subject = subject
                # Escalar font-size px → pt para que Outlook los respete
                def _escalar_px(m):
                    px = int(m.group(1))
                    # Mínimo 11pt para cuerpo, mantener proporciones
                    pt = max(11, round(px * 0.95))
                    return f'font-size:{pt}pt'
                body_escalado = re.sub(r'font-size:(\d+)px', _escalar_px, html_body)

                cuerpo = (
                    '<div style="font-family:Calibri,Arial,sans-serif;'
                    'font-size:12pt;line-height:1.8;color:#1e293b;max-width:820px;padding:8px 4px">'
                    + body_escalado
                        .replace('<p>', '<p style="margin:10px 0;font-size:12pt;font-family:Calibri,Arial,sans-serif;line-height:1.8">')
                        .replace('<li>', '<li style="margin:4px 0;font-size:12pt;font-family:Calibri,Arial,sans-serif;line-height:1.8">')
                        .replace('<strong>', '<strong style="font-size:12pt;font-family:Calibri,Arial,sans-serif">')
                    + '</div>'
                )
                _ = mail.GetInspector
                mail.HTMLBody = cuerpo
                mail.Display(False)
                return jsonify({'ok': True, 'accion': 'abierto en Outlook'})
            finally:
                pythoncom.CoUninitialize()
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

# ── Copia al SharePoint sincronizado localmente ───────────────────────────────
def _subir_sharepoint(fpath, fname):
    if not SP_LOCAL:
        print(f'[SHAREPOINT] ⚠️  Carpeta no detectada — sincroniza la carpeta de SharePoint primero.', flush=True)
        return
    try:
        destino = os.path.join(SP_LOCAL, fname)
        shutil.copy2(fpath, destino)
        print(f'[SHAREPOINT] ✅ Copiado a OneDrive: {fname}', flush=True)
        print(f'[SHAREPOINT]    → {destino}', flush=True)
    except Exception as e:
        print(f'[SHAREPOINT] ❌ Error: {e}', flush=True)

# ── Watcher: vigila Descargas ──────────────────────────────────────────────────
def _watcher():
    try:
        for f in os.listdir(DOWNLOADS):
            if f.startswith('Cotizacion_') and f.endswith('.xlsx'):
                _seen_files.add(os.path.join(DOWNLOADS, f))
        print(f'[WATCHER] Vigilando: {DOWNLOADS}', flush=True)
        print(f'[WATCHER] Archivos previos ignorados: {len(_seen_files)}', flush=True)
    except Exception as e:
        print(f'[WATCHER] Error al iniciar: {e}', flush=True)

    while True:
        try:
            for fname in os.listdir(DOWNLOADS):
                if fname.startswith('Cotizacion_') and fname.endswith('.xlsx'):
                    fpath = os.path.join(DOWNLOADS, fname)
                    if fpath not in _seen_files:
                        _seen_files.add(fpath)
                        time.sleep(1.5)
                        if os.path.exists(fpath):
                            print(f'[WATCHER] Nuevo archivo detectado: {fname}', flush=True)
                            threading.Thread(target=_subir_sharepoint, args=(fpath, fname), daemon=True).start()
        except Exception:
            pass
        time.sleep(4)

# ── Inicio ─────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    print('=' * 60)
    print('  FRIOPACKING — Servidor OC Mailer + SharePoint')
    print('=' * 60)
    print(f'  Vigilando descargas : {DOWNLOADS}')
    print(f'  Carpeta OCs         : {OC_FOLDER or "*** NO DETECTADA ***"}')
    print(f'  Carpeta Cotizaciones: {COT_FOLDER or "*** NO DETECTADA ***"}')
    print(f'  Carpeta Facturas    : {FACT_FOLDER or "*** NO DETECTADA ***"}')
    print(f'  Carpeta Comprobantes: {COMP_FOLDER or "*** NO DETECTADA ***"}')
    print(f'  Carpeta OPD         : {OPD_FOLDER or "*** NO DETECTADA ***"}')
    if SP_LOCAL:
        print(f'  Destino SharePoint  : {SP_LOCAL}')
    else:
        print('  Destino SharePoint  : *** NO DETECTADO ***')
        print()
        print('  ATENCION: La carpeta de SharePoint no esta sincronizada.')
        print('  Para activar la subida automatica:')
        print('  1. Abre SharePoint > Trazabilidad Compras 2026')
        print('  2. Entra a COMPARATIVO_COTIZACIONES_COMPRAS')
        print('  3. Clic en "Anadir acceso directo a Mis archivos"')
        print('  4. Espera que OneDrive sincronice y reinicia este servidor.')
    print(f'  Escuchando en       : http://127.0.0.1:5100')
    print('  Manten esta ventana abierta mientras usas el dashboard')
    print('=' * 60)

    t = threading.Thread(target=_watcher, daemon=True)
    t.start()

    app.run(host='127.0.0.1', port=5100, debug=False, threaded=True)
