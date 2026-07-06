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
def handle_options():
    return ('', 204)

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

    # Extraer número de OC del asunto: "Orden de Compra Nº 0001-0011557 — ..."
    match_oc = re.search(r'(\d{4}-\d{7})', subject)
    numero_oc = match_oc.group(1) if match_oc else None
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
                    # Adjuntar PDF de la OC si se encuentra en la carpeta
                    pdf_oc = _buscar_pdf_oc(numero_oc)
                    if pdf_oc:
                        mail.Attachments.Add(pdf_oc)
                        print(f'[SEND-OC] 📎 OC adjuntada: {os.path.basename(pdf_oc)}', flush=True)
                    else:
                        print(f'[SEND-OC] ⚠️  No se encontró PDF para OC {numero_oc}', flush=True)
                    # Adjuntar CARTILLA SSOMA si es entrega en Lurín
                    # Detecta por modality (versión nueva del dashboard) O por contenido del cuerpo (versión online)
                    es_lurin = (str(modality or '').lower() == 'planta') or ('lurín' in html_body.lower()) or ('lurin' in html_body.lower())
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
