# Plataforma Compras — FRIOPACKING S.A.

Dashboard de seguimiento de Órdenes de Compra, Requerimientos y Proyectos PMO 2026.

Este repositorio también es el que publica el sitio en vivo vía GitHub Pages
(`melissa2026714.github.io/dashboard-friopacking`) — por eso `index.html` y
`msal-browser.min.js` se mantienen en la raíz (GitHub Pages solo puede servir
desde la raíz del repo o desde una carpeta llamada `docs`).

## Estructura

- `index.html` / `msal-browser.min.js` — Aplicación web del dashboard (raíz, para que GitHub Pages la sirva directamente).
- `data.json` / `consultas.json` — Datos compartidos que la app sincroniza con todos los usuarios (se actualizan solos desde el botón GUARDAR del dashboard, no se editan a mano).
- `data/` — Excel de importación (Reporte Maestro, PedidoSinReq, Consulta). Se cargan desde el botón **IMPORTAR** del dashboard.
- `tools/` — Utilidades locales para Windows:
  - `iniciar_dashboard.bat` — levanta un servidor local en `http://localhost:8080` para probar el dashboard sin depender de GitHub Pages.
  - `iniciar_servidor_correos.bat` / `oc_mailer.py` — servidor local (puerto 5100) que envía OC/cotizaciones por Outlook y sube copias a SharePoint.
- `referencias/` — Documentación de referencia (`generar_resumen_pdf.py` genera el PDF de análisis de gestión de compras).

## Uso local

1. Doble clic en `tools/iniciar_dashboard.bat` para abrir el dashboard en el navegador.
2. (Opcional) `tools/iniciar_servidor_correos.bat` si necesitas enviar correos de OC/cotizaciones desde el dashboard.

## Publicación

Los cambios hechos en `index.html` se publican automáticamente vía GitHub Pages al hacer push a `main`. No se necesita ningún paso adicional.
