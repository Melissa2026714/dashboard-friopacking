# -*- coding: utf-8 -*-
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    ListFlowable, ListItem, HRFlowable, PageBreak
)
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY

OUT = "Resumen_Cafe_AERCE_Analisis_Gestion_Control_Compras.pdf"

styles = getSampleStyleSheet()

styles.add(ParagraphStyle(
    name="TituloPortada", fontName="Helvetica-Bold", fontSize=20,
    leading=26, alignment=TA_CENTER, textColor=colors.HexColor("#1F3864"),
    spaceAfter=10
))
styles.add(ParagraphStyle(
    name="Subtitulo", fontName="Helvetica", fontSize=12,
    leading=16, alignment=TA_CENTER, textColor=colors.HexColor("#44546A"),
    spaceAfter=4
))
styles.add(ParagraphStyle(
    name="H1", fontName="Helvetica-Bold", fontSize=14,
    leading=18, textColor=colors.HexColor("#1F3864"),
    spaceBefore=16, spaceAfter=8
))
styles.add(ParagraphStyle(
    name="H2", fontName="Helvetica-Bold", fontSize=11.5,
    leading=15, textColor=colors.HexColor("#2E5395"),
    spaceBefore=10, spaceAfter=6
))
styles.add(ParagraphStyle(
    name="Cuerpo", fontName="Helvetica", fontSize=10, leading=14,
    alignment=TA_JUSTIFY, spaceAfter=6
))
styles.add(ParagraphStyle(
    name="CuerpoNegrita", parent=styles["Cuerpo"], fontName="Helvetica-Bold"
))
styles.add(ParagraphStyle(
    name="ViñetaCuerpo", parent=styles["Cuerpo"], leftIndent=10, spaceAfter=4
))
styles.add(ParagraphStyle(
    name="Nota", fontName="Helvetica-Oblique", fontSize=9, leading=12,
    textColor=colors.HexColor("#555555")
))

story = []

# ---------- PORTADA ----------
story.append(Spacer(1, 3*cm))
story.append(Paragraph("Resumen Ejecutivo", styles["TituloPortada"]))
story.append(Paragraph('"Un café con AERCE: Análisis, Gestión y Control de Compras"', styles["Subtitulo"]))
story.append(Spacer(1, 0.5*cm))
story.append(HRFlowable(width="60%", thickness=1, color=colors.HexColor("#1F3864"), hAlign="CENTER"))
story.append(Spacer(1, 0.8*cm))

meta_data = [
    ["Ponente:", "Rafael Castelló — Director Académico de AERCE"],
    ["Organizador:", "AERCE (Asociación Española de Profesionales de Compras)"],
    ["Formato:", "Webinar (~60 min), basado en el curso \"Análisis, Gestión y Control de Compras con Excel\""],
    ["Elaborado por:", "Análisis de Compras — FrioPacking"],
    ["Fecha del resumen:", "02 de julio de 2026"],
]
t = Table(meta_data, colWidths=[3.5*cm, 11*cm])
t.setStyle(TableStyle([
    ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
    ("FONTNAME", (1,0), (1,-1), "Helvetica"),
    ("FONTSIZE", (0,0), (-1,-1), 9.5),
    ("TEXTCOLOR", (0,0), (0,-1), colors.HexColor("#1F3864")),
    ("VALIGN", (0,0), (-1,-1), "TOP"),
    ("BOTTOMPADDING", (0,0), (-1,-1), 6),
]))
story.append(t)
story.append(PageBreak())

def h1(text):
    story.append(Paragraph(text, styles["H1"]))

def h2(text):
    story.append(Paragraph(text, styles["H2"]))

def p(text):
    story.append(Paragraph(text, styles["Cuerpo"]))

def bullets(items):
    story.append(ListFlowable(
        [ListItem(Paragraph(i, styles["ViñetaCuerpo"]), leftIndent=8) for i in items],
        bulletType="bullet", start="•"
    ))

# ---------- 1. Premisa ----------
h1("1. Premisa central")
p("Compras ya es reconocida como una función estratégica de la compañía, y como tal debe demostrar valor mediante objetivos medibles y mejora continua. Principio guía de la sesión:")
story.append(Paragraph(
    "&ldquo;Lo que no se define no se puede medir, lo que no se mide no se puede mejorar, "
    "y lo que no se mejora se degrada.&rdquo;",
    styles["CuerpoNegrita"]
))
p("Toda la sesión gira en torno a construir un sistema de análisis (ABC + mapa de compras) que permita fijar estrategias diferenciadas por proveedor, producto y categoría, en lugar de tratarlos a todos por igual.")
p("<b>Caso de estudio usado como ejemplo:</b> empresa con 124.000 líneas de pedido/año, €145M de gasto anual, 185 proveedores, 1.800 productos y 18 familias/categorías.")

# ---------- 2. ABC ----------
h1("2. Análisis ABC (Pareto 80/20) — el pilar metodológico")
p("Se aplica el mismo análisis a tres dimensiones, cruzando el % acumulado de proveedores/productos/familias frente al % acumulado del presupuesto:")

data = [
    ["Dimensión", "Tipo A (≈80%)", "Tipo B (hasta ≈95%)", "Tipo C (resto)"],
    ["Proveedores", "36 proveedores → €98M", "77 proveedores → €36M", "62 proveedores → €10M"],
    ["Productos", "398 productos → €115M", "—", "—"],
    ["Familias", "22% de familias → 76% del gasto", "50% de familias → 94%", "resto marginal"],
]
tbl = Table(data, colWidths=[3*cm, 4*cm, 4*cm, 3.5*cm])
tbl.setStyle(TableStyle([
    ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1F3864")),
    ("TEXTCOLOR", (0,0), (-1,0), colors.white),
    ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
    ("FONTNAME", (0,1), (-1,-1), "Helvetica"),
    ("FONTSIZE", (0,0), (-1,-1), 9),
    ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#AAAAAA")),
    ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F2F5FA")]),
    ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ("TOPPADDING", (0,0), (-1,-1), 5),
    ("BOTTOMPADDING", (0,0), (-1,-1), 5),
]))
story.append(tbl)
story.append(Spacer(1, 8))

p("<b>Hallazgo crítico del caso:</b> una sola familia concentra el <b>46% de toda la facturación</b> (€67M) → la empresa es &ldquo;mono-familia&rdquo;, con un riesgo de negocio muy alto si esa categoría falla (rotura de stock, problema de suministro, etc.).")
p("<b>Recomendación de presentación:</b> los informes de Compras no deben ser tablas con 14 columnas y colores — deben presentarse como gráficos ABC simples con una etiqueta explicativa (ej. &ldquo;36 proveedores tipo A = €98M&rdquo;) para que interlocutores no técnicos (dirección, finanzas) entiendan el mensaje sin necesidad de conocer la metodología.")

# ---------- 3. Mapa de compras ----------
h1("3. Mapa de Compras (adaptación de la matriz de Kraljic)")
p("Diagrama de burbujas por familia con:")
bullets([
    "<b>Eje X:</b> importancia/riesgo de la familia (escala 1–10)",
    "<b>Eje Y:</b> número de proveedores que la suministran",
    "<b>Tamaño de burbuja:</b> volumen de gasto",
    "<b>Color:</b> si está dentro o fuera del &ldquo;perímetro de Compras&rdquo; (gestionado o no por el departamento)",
])
p("<b>Insight clave:</b> detecta familias de alta importancia y alto volumen que <b>no</b> están gestionadas por Compras (ej. cartón, transporte, ferretería/EPIs) — con un número de proveedores anómalamente alto para el mercado real (30 proveedores de cartón cuando el mercado tiene 6-8 fabricantes relevantes; 40 de transporte sin ser empresa logística).")
p("Esto es la señal de alarma típica que un analista debe buscar: <b>exceso de proveedores relativo a la concentración natural del mercado proveedor.</b>")

# ---------- 4. Estrategias ----------
h1("4. Tres estrategias según segmento del mapa")
bullets([
    "<b>Familias importantes fuera del perímetro</b> → incorporarlas a Compras, luego reducir el número de proveedores y consolidar volumen para conseguir mejores condiciones.",
    "<b>Familias &ldquo;core&rdquo; (alto volumen + alta importancia, ya en Compras)</b> → homologar proveedores, desarrollarlos (innovación conjunta), construir relación estratégica bilateral.",
    "<b>Familias de bajo volumen/baja importancia (cola larga)</b> → minimizar el esfuerzo de gestión: contrato marco a largo plazo con proveedor único, subasta electrónica, para liberar tiempo del comprador hacia donde sí aporta valor.",
])

# ---------- 5. Organización ----------
h1("5. Organización: categoría → comprador → objetivos → KPI")
p("Cada comprador se asigna a un conjunto de familias con objetivos concretos y cuantificados (ej.: &ldquo;reducir proveedores de cartón un 50% en el primer año&rdquo;, &ldquo;revisar necesidad real de 35 proveedores de transporte y reducir 50% en 6 meses&rdquo;). Los KPI se derivan directamente de esos objetivos (% de reducción alcanzado por periodo), no son genéricos.")

# ---------- 6. Ahorro potencial ----------
h1("6. Estimación de ahorro potencial (ejercicio cuantitativo)")
p("Sobre el gasto total, aplicando cada estrategia a su segmento correspondiente:")

data2 = [
    ["Acción", "Volumen (perímetro)", "Palanca", "Ahorro estimado"],
    ["Acción 1", "≈ €46M", "Incorporar al perímetro de Compras", "≈ 4%"],
    ["Acción 2", "≈ €80M", "Análisis de estructura de costes del proveedor\n(fijos, variables, margen, financieros)", "≈ 10%"],
    ["Acción 3", "≈ €6M", "Subasta electrónica (cola larga)", "≈ 13%"],
]
tbl2 = Table(data2, colWidths=[2.2*cm, 3*cm, 6.3*cm, 3*cm])
tbl2.setStyle(TableStyle([
    ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1F3864")),
    ("TEXTCOLOR", (0,0), (-1,0), colors.white),
    ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
    ("FONTNAME", (0,1), (-1,-1), "Helvetica"),
    ("FONTSIZE", (0,0), (-1,-1), 8.7),
    ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#AAAAAA")),
    ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F2F5FA")]),
    ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ("TOPPADDING", (0,0), (-1,-1), 5),
    ("BOTTOMPADDING", (0,0), (-1,-1), 5),
]))
story.append(tbl2)
story.append(Spacer(1, 8))
p("<b>Total estimado: ≈ €2,5 millones de ahorro</b> con los recursos adecuados. Este es el tipo de &ldquo;business case&rdquo; que un analista de compras debería poder presentar a dirección: no ahorro genérico, sino desglosado por palanca y por segmento del gasto.")

# ---------- 7. Casos prácticos ----------
h1("7. Casos prácticos de \"mala gestión\" detectados en el drill-down por familia")
bullets([
    "Familia con 10 proveedores &ldquo;top&rdquo;: ninguno destaca en facturación → sin gestión clara. Se recomienda homologar/rankear y concentrar volumen en 2-3.",
    "Familia con 6.500 pedidos/año en una sola categoría → cuestionar si son necesarios; sustituir por contrato marco abierto con entregas periódicas.",
    "Familia &ldquo;compresores&rdquo;: 7 proveedores para solo el 0,97% del presupuesto e importancia 4 → gestión ineficiente (coste de gestionar &gt; valor obtenido); reducir a 1 proveedor salvo riesgo de suministro relevante.",
    "Familia &ldquo;engranajes&rdquo;: 17 proveedores, importancia 5, €4,5M → reducir drásticamente a 2.",
])
p("<b>Lección de analista:</b> el esfuerzo de gestión debe ser proporcional al valor/riesgo de la categoría — dedicar tiempo desproporcionado a categorías tipo C de bajo impacto es un error común y costoso en términos de oportunidad.")

# ---------- 8. Evaluación de ofertas ----------
h1("8. Evaluación de ofertas (más allá del precio)")
p("En productos estratégicos no debe ganar automáticamente el más barato. Ejemplo de ponderación usada en la sesión: 70% precio, 10% sistema de calidad, 10% planificación/JIT, 5% condiciones de pago.")
p("Se recomienda un gráfico comparativo por proveedor (real vs. &ldquo;óptimo&rdquo;) para retroalimentar a cada proveedor sobre en qué debe mejorar (precio, calidad, JIT, etc.) y mantenerlos en mejora continua.")

# ---------- 9. Q&A ----------
h1("9. Puntos clave del turno de preguntas")
bullets([
    "<b>Fuente de datos:</b> todo el análisis se puede construir a partir del plan general contable / pagos a terceros que provee Finanzas — aplicable tanto a sector industrial como a servicios (en servicios se añade el reto del control de calidad vía SLA).",
    "<b>Umbrales ABC:</b> 80% (A) / 95% (B) / resto (C) — estándar, aunque ajustable según el negocio.",
    "<b>Periodo de datos:</b> 1 año para el ABC; hasta 3 años si se analiza demanda o estacionalidad.",
    "<b>Frecuencia de revisión:</b> depende de la volatilidad del mercado y del objetivo fijado (ej. revisar a mitad de un plan de reducción de proveedores a 6 meses).",
    "<b>Software:</b> no se recomienda una herramienta específica — depende del volumen de la empresa; las herramientas potentes existen pero son costosas.",
    "<b>Contexto post-pandemia:</b> Compras no tiene un &ldquo;libro contable&rdquo; fijo — cada decisión depende del riesgo, sector, momento y visión de la empresa; es clave reevaluar la salud financiera de los proveedores.",
    "Próxima edición del curso Excel de AERCE: modalidad online, formato de 3 semanas con webinars.",
])

# ---------- 10. Valoración ----------
h1("10. Valoración como analista de compras")
p("Contenido muy práctico y accionable, centrado en tres herramientas replicables con datos que ya existen en la contabilidad de cualquier empresa (sin necesidad de software costoso):")
bullets([
    "ABC de proveedores/productos/familias (Pareto).",
    "Mapa de compras tipo Kraljic (importancia vs. número de proveedores vs. volumen).",
    "Ficha de análisis por familia (top-10 proveedores, concentración, número de pedidos) para decidir consolidación.",
])
p("<b>Debilidad de la sesión:</b> es de naturaleza promocional del curso de Excel de AERCE, por lo que el detalle metodológico completo (fórmulas, plantilla real) no se comparte en su totalidad — solo se muestra el resultado.")
p("<b>Recomendación de aplicación inmediata en FrioPacking</b> (sin necesidad del curso):")
bullets([
    "Solicitar a Finanzas el detalle de pagos a proveedores por cuenta contable (últimos 12 meses).",
    "Construir el ABC de proveedores, productos y familias en Excel con porcentajes acumulados.",
    "Cruzarlo con una matriz de importancia/riesgo por categoría para priorizar dónde enfocar el tiempo del equipo de compras.",
    "Identificar categorías de alto gasto/alto riesgo gestionadas fuera del perímetro formal de Compras.",
])

story.append(Spacer(1, 14))
story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#AAAAAA")))
story.append(Spacer(1, 6))
story.append(Paragraph(
    "Fuente: Webinar &ldquo;Un café con AERCE | Análisis, gestión y control de Compras&rdquo; — "
    "YouTube (video de referencia proporcionado por el usuario). Documento elaborado a partir de la "
    "transcripción completa de la sesión.",
    styles["Nota"]
))

doc = SimpleDocTemplate(
    OUT, pagesize=A4,
    topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm,
    title="Resumen Cafe AERCE - Analisis Gestion y Control de Compras"
)
doc.build(story)
print("OK ->", OUT)
