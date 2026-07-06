# Plataforma Compras — FRIOPACKING S.A.

Dashboard de seguimiento de Órdenes de Compra, Requerimientos y Proyectos PMO 2026.

## Estructura

- `app/` — Aplicación web del dashboard (`index.html` + dependencias). Esta es la fuente local; la versión publicada en línea vive en el repositorio `dashboard-friopacking` (GitHub Pages) y se actualiza subiendo `app/index.html` ahí manualmente.
- `data/` — Excel de importación (Reporte Maestro, PedidoSinReq, Consulta). Se cargan desde el botón **IMPORTAR** del dashboard.
- `tools/` — Utilidades locales para Windows:
  - `iniciar_dashboard.bat` — levanta un servidor local en `http://localhost:8080` para probar `app/index.html` sin depender de GitHub Pages.
  - `iniciar_servidor_correos.bat` / `oc_mailer.py` — servidor local (puerto 5100) que envía OC/cotizaciones por Outlook y sube copias a SharePoint.
- `docs/` — Documentación de referencia (`generar_resumen_pdf.py` genera el PDF de análisis de gestión de compras).

## Uso local

1. Doble clic en `tools/iniciar_dashboard.bat` para abrir el dashboard en el navegador.
2. (Opcional) `tools/iniciar_servidor_correos.bat` si necesitas enviar correos de OC/cotizaciones desde el dashboard.

## Publicación

El dashboard en producción se sirve desde un repositorio separado vía GitHub Pages. Los cambios hechos en `app/index.html` aquí deben subirse manualmente a ese repositorio para reflejarse en el sitio en vivo.
