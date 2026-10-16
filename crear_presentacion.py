from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


OUT = Path(__file__).with_name("Presentacion_Dolce_Vita_Proceso_Unificado.pptx")

W, H = 13.333, 7.5
CREAM = RGBColor(250, 247, 240)
WHITE = RGBColor(255, 255, 255)
INK = RGBColor(43, 42, 38)
MUTED = RGBColor(100, 96, 88)
BROWN = RGBColor(102, 67, 48)
ROSE = RGBColor(192, 78, 92)
GREEN = RGBColor(46, 105, 88)
GOLD = RGBColor(218, 166, 78)
PALE_ROSE = RGBColor(245, 228, 229)
PALE_GREEN = RGBColor(226, 238, 233)
PALE_GOLD = RGBColor(247, 237, 213)
LINE = RGBColor(220, 214, 204)

prs = Presentation()
prs.slide_width = Inches(W)
prs.slide_height = Inches(H)
prs.core_properties.title = "Proceso Unificado aplicado a Dolce‑Vita"
prs.core_properties.subject = "Actividad sumativa, semana 3 — Diseño de Software"
prs.core_properties.author = "Equipo de trabajo"
prs.core_properties.keywords = "Proceso Unificado, SDLC, casos de uso, Dolce-Vita"


def set_bg(slide, color=CREAM):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_text(slide, text, x, y, w, h, size=18, color=INK, bold=False,
             font="Aptos", align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP,
             margin=0.06):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.margin_left = Inches(margin)
    tf.margin_right = Inches(margin)
    tf.margin_top = Inches(margin)
    tf.margin_bottom = Inches(margin)
    tf.word_wrap = True
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    return box


def add_rich_text(slide, runs, x, y, w, h, size=18, color=INK,
                  align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.05)
    tf.margin_bottom = Inches(0.05)
    tf.word_wrap = True
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    for text, bold, run_color in runs:
        r = p.add_run()
        r.text = text
        r.font.name = "Aptos"
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = run_color or color
    return box


def add_box(slide, x, y, w, h, fill=WHITE, line=LINE, radius=True):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = line
    shape.line.width = Pt(1)
    return shape


def add_bullets(slide, items, x, y, w, h, size=17, color=INK,
                bullet_color=None, spacing=6):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.03)
    tf.margin_top = Inches(0.03)
    tf.margin_bottom = Inches(0.03)
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if isinstance(item, tuple):
            lead, body = item
            r = p.add_run()
            r.text = lead
            r.font.bold = True
            r.font.color.rgb = bullet_color or BROWN
            r.font.name = "Aptos"
            r.font.size = Pt(size)
            r2 = p.add_run()
            r2.text = body
            r2.font.color.rgb = color
            r2.font.name = "Aptos"
            r2.font.size = Pt(size)
        else:
            p.text = item
            for r in p.runs:
                r.font.name = "Aptos"
                r.font.size = Pt(size)
                r.font.color.rgb = color
        p.text = "• " + p.text if not isinstance(item, tuple) else p.text
        p.space_after = Pt(spacing)
        p.level = 0
    return box


def add_header(slide, number, section, title, subtitle=None):
    add_text(slide, f"{number:02d}  {section.upper()}", 0.55, 0.27, 4.3, 0.32,
             10, ROSE, True)
    add_text(slide, title, 0.55, 0.70, 12.1, 0.60, 26, INK, True)
    if subtitle:
        add_text(slide, subtitle, 0.58, 1.30, 12.0, 0.45, 13, MUTED)
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.55), Inches(1.77),
                                  Inches(0.75), Inches(0.05))
    line.fill.solid()
    line.fill.fore_color.rgb = ROSE
    line.line.fill.background()


def add_footer(slide, number):
    add_text(slide, "Diseño de Software · Semana 3", 0.55, 7.14, 4.0, 0.18,
             9, MUTED)
    add_text(slide, str(number), 12.25, 7.10, 0.48, 0.23, 9, MUTED,
             align=PP_ALIGN.RIGHT)


def new_slide(number, section, title, subtitle=None, bg=CREAM):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, bg)
    add_header(slide, number, section, title, subtitle)
    add_footer(slide, number)
    return slide


def add_card(slide, title, body, x, y, w, h, accent=GREEN, fill=WHITE,
             title_size=16, body_size=13):
    add_box(slide, x, y, w, h, fill)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y),
                                 Inches(0.08), Inches(h))
    bar.fill.solid()
    bar.fill.fore_color.rgb = accent
    bar.line.fill.background()
    add_text(slide, title, x + 0.22, y + 0.17, w - 0.38, 0.35,
             title_size, accent, True)
    add_text(slide, body, x + 0.22, y + 0.62, w - 0.38, h - 0.76,
             body_size, INK)


def add_table(slide, data, x, y, w, h, col_widths=None, font_size=12,
              header_fill=BROWN):
    rows, cols = len(data), len(data[0])
    table = slide.shapes.add_table(rows, cols, Inches(x), Inches(y),
                                   Inches(w), Inches(h)).table
    if col_widths:
        for i, width in enumerate(col_widths):
            table.columns[i].width = Inches(width)
    for r in range(rows):
        for c in range(cols):
            cell = table.cell(r, c)
            cell.text = data[r][c]
            cell.margin_left = Inches(0.07)
            cell.margin_right = Inches(0.07)
            cell.margin_top = Inches(0.05)
            cell.margin_bottom = Inches(0.04)
            cell.fill.solid()
            cell.fill.fore_color.rgb = header_fill if r == 0 else (
                WHITE if r % 2 else RGBColor(244, 241, 235)
            )
            cell.border = None
            for p in cell.text_frame.paragraphs:
                p.space_after = Pt(0)
                p.vertical_anchor = MSO_ANCHOR.MIDDLE
                for run in p.runs:
                    run.font.name = "Aptos"
                    run.font.size = Pt(font_size if r else font_size - 0.5)
                    run.font.bold = r == 0 or c == 0
                    run.font.color.rgb = WHITE if r == 0 else INK
    return table


# 1. Portada
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, CREAM)
accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
                                Inches(0.22), Inches(H))
accent.fill.solid()
accent.fill.fore_color.rgb = ROSE
accent.line.fill.background()
add_text(slide, "ACTIVIDAD SUMATIVA · SEMANA 3", 0.72, 0.62, 5.5, 0.35,
         11, ROSE, True)
add_text(slide, "Proceso Unificado\naplicado a Dolce‑Vita", 0.72, 1.35, 7.0, 1.65,
         31, INK, True)
add_text(slide, "Propuesta metodológica para un sistema de ventas en línea",
         0.75, 3.20, 6.2, 0.65, 18, BROWN)
add_box(slide, 0.75, 4.30, 5.75, 1.18, PALE_GREEN, PALE_GREEN)
add_text(slide, "Integrantes: [Nombre 1] · [Nombre 2] · [Nombre 3]\n"
                "Asignatura: Diseño de Software  |  Docente: [Nombre]\n"
                "Sección: [Sección]  |  Fecha: [Fecha]",
         1.00, 4.57, 5.30, 0.78, 13, INK)

# Ilustración geométrica de carro y productos
add_box(slide, 8.20, 1.10, 3.95, 4.95, WHITE, LINE)
cart = slide.shapes.add_shape(MSO_SHAPE.TRAPEZOID, Inches(8.80), Inches(3.15),
                              Inches(2.70), Inches(1.35))
cart.fill.solid()
cart.fill.fore_color.rgb = PALE_ROSE
cart.line.color.rgb = ROSE
cart.line.width = Pt(2)
for cx in (9.25, 10.95):
    wheel = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx), Inches(4.58),
                                   Inches(0.40), Inches(0.40))
    wheel.fill.solid()
    wheel.fill.fore_color.rgb = BROWN
    wheel.line.fill.background()
handle = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(8.35), Inches(2.68),
                                    Inches(8.95), Inches(3.28))
handle.line.color.rgb = BROWN
handle.line.width = Pt(3)
for x, y, color in [(8.95, 2.05, GREEN), (9.85, 1.72, GOLD), (10.72, 2.15, ROSE)]:
    item = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y),
                                  Inches(0.72), Inches(0.82))
    item.fill.solid()
    item.fill.fore_color.rgb = color
    item.line.fill.background()
add_text(slide, "DOLCE\nVITA", 9.26, 5.28, 1.85, 0.70, 17, BROWN, True,
         align=PP_ALIGN.CENTER)
add_text(slide, "Presentación editable", 9.00, 6.40, 2.25, 0.28, 10, MUTED,
         align=PP_ALIGN.CENTER)
add_footer(slide, 1)


# 2. Caso y alcance
slide = new_slide(2, "Caso", "Problema y alcance del sistema",
                  "Una tienda en línea requiere compras confiables y control de inventario.")
add_card(slide, "Necesidad del cliente",
         "Implementar un canal digital que permita vender uno o más productos mediante "
         "un carro de compras, manteniendo información de usuarios, stock y ventas.",
         0.65, 2.15, 4.0, 2.02, ROSE, PALE_ROSE, 17, 14)
add_card(slide, "Valor esperado",
         "Reducir errores de inventario, facilitar la compra, alertar faltantes a tiempo "
         "y entregar una boleta trazable por cada venta.",
         0.65, 4.38, 4.0, 1.78, GREEN, PALE_GREEN, 17, 14)
add_text(slide, "Capacidades funcionales", 5.05, 2.12, 3.2, 0.38, 18, BROWN, True)
add_bullets(slide, [
    "Registro y autenticación de clientes y administradores.",
    "Compra de uno o más productos mediante carro.",
    "Consulta de productos y stock disponible.",
    "Ingreso y actualización de stock.",
    "Alertas automáticas ante stock crítico.",
    "Emisión de boleta de venta."
], 5.05, 2.66, 7.30, 3.55, 17, INK, GREEN, 8)
add_box(slide, 5.05, 6.12, 7.25, 0.66, PALE_GOLD, PALE_GOLD)
add_text(slide, "Alcance inicial: venta, inventario y comprobante; futuras integraciones "
                "se gestionan como cambios.", 5.28, 6.23, 6.82, 0.40, 11, BROWN, True)


# 3. Selección de metodología
slide = new_slide(3, "Metodología", "¿Por qué elegir el Proceso Unificado?",
                  "Es iterativo, incremental, dirigido por casos de uso y centrado en la arquitectura.")
add_box(slide, 0.70, 2.10, 3.75, 4.40, WHITE)
add_text(slide, "Decisión", 1.00, 2.42, 1.6, 0.35, 16, ROSE, True)
add_text(slide, "PROCESO\nUNIFICADO", 1.00, 3.05, 2.90, 1.10, 27, INK, True)
add_text(slide, "Adecuado para construir el sistema por incrementos verificables y "
                "mantener trazabilidad entre necesidad, diseño, código y prueba.",
         1.00, 4.50, 2.95, 1.18, 14, MUTED)
reasons = [
    ("Casos de uso", "Organizan requisitos y entregas alrededor del valor para cada actor.", ROSE),
    ("Arquitectura", "Estabiliza decisiones críticas antes de ampliar funcionalidades.", GREEN),
    ("Iteraciones", "Permiten retroalimentación temprana y ajuste controlado.", GOLD),
    ("Riesgos", "Los problemas técnicos y de negocio se abordan desde el inicio.", BROWN),
]
for i, (title, body, color) in enumerate(reasons):
    y = 2.10 + i * 1.12
    add_card(slide, title, body, 4.82, y, 7.65, 0.91, color, WHITE, 15, 12)
add_text(slide, "Principio de calidad", 4.90, 6.66, 1.70, 0.24, 10, ROSE, True)
add_text(slide, "Diagramas y arquitectura coherentes → implementación coherente.",
         6.45, 6.60, 5.65, 0.34, 13, INK, True)


# 4. Gestión del proyecto
slide = new_slide(4, "Gestión", "El desarrollo se gestiona como un proyecto",
                  "Planificación, recursos, programación, roles y control acompañan todas las iteraciones.")
steps = [
    ("1", "Planificar", "Definir alcance, casos de uso, iteraciones, entregables, costos y criterios de aceptación."),
    ("2", "Asignar", "Distribuir personas, herramientas y ambientes según competencias y carga de trabajo."),
    ("3", "Programar", "Priorizar por valor y riesgo; establecer hitos, dependencias y líneas base."),
    ("4", "Controlar", "Medir avance, defectos y riesgos; aprobar cambios y ajustar el plan."),
]
for i, (num, title, body) in enumerate(steps):
    x = 0.65 + i * 3.08
    add_box(slide, x, 2.17, 2.78, 2.46, WHITE)
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.18), Inches(2.38),
                                    Inches(0.48), Inches(0.48))
    circle.fill.solid()
    circle.fill.fore_color.rgb = [ROSE, GREEN, GOLD, BROWN][i]
    circle.line.fill.background()
    add_text(slide, num, x + 0.18, 2.43, 0.48, 0.22, 11, WHITE, True,
             align=PP_ALIGN.CENTER)
    add_text(slide, title, x + 0.82, 2.37, 1.70, 0.36, 16,
             [ROSE, GREEN, GOLD, BROWN][i], True)
    add_text(slide, body, x + 0.20, 3.05, 2.37, 1.24, 13, INK)
add_box(slide, 0.65, 4.95, 11.95, 1.35, PALE_GREEN, PALE_GREEN)
add_rich_text(slide, [
    ("Seguimiento propuesto: ", True, GREEN),
    ("reunión semanal, tablero Kanban, revisión al cierre de cada iteración y métricas de "
     "avance, defectos, cobertura de pruebas y cambios pendientes.", False, INK)
], 0.95, 5.20, 11.30, 0.72, 15)
add_text(slide, "Responsable: gestor de proyecto · Participan: equipo técnico y cliente",
         0.95, 5.88, 10.80, 0.25, 11, MUTED)


# 5. Roles
slide = new_slide(5, "Gestión", "Roles y responsabilidades",
                  "La responsabilidad es explícita, pero el trabajo se coordina alrededor de los casos de uso.")
roles = [
    ("Gestor/a", "Planifica iteraciones, recursos, riesgos y comunicación.", ROSE),
    ("Analista", "Levanta requisitos y mantiene el modelo de casos de uso.", BROWN),
    ("Arquitecto/a", "Define arquitectura, interfaces y decisiones técnicas.", GREEN),
    ("Diseñador/a", "Transforma casos de uso en modelos UML y componentes.", GOLD),
    ("Desarrollador/a", "Implementa, integra, revisa y documenta el código.", ROSE),
    ("Tester", "Diseña pruebas, verifica criterios y registra desviaciones.", GREEN),
    ("Config. y cambios", "Versiona artefactos, líneas base y solicitudes.", BROWN),
    ("Cliente", "Prioriza valor, aclara reglas y acepta cada incremento.", GOLD),
]
for i, (title, body, color) in enumerate(roles):
    col = i % 2
    row = i // 2
    x = 0.70 + col * 6.05
    y = 2.05 + row * 1.15
    add_card(slide, title, body, x, y, 5.72, 0.93, color, WHITE, 14, 12)
add_box(slide, 0.70, 6.67, 11.77, 0.28, PALE_ROSE, PALE_ROSE)
add_text(slide, "Regla: nadie modifica una línea base sin registrar, analizar y aprobar el cambio.",
         0.92, 6.69, 11.30, 0.20, 10, ROSE, True, align=PP_ALIGN.CENTER)


# 6. Fases
slide = new_slide(6, "Ciclo de vida", "Fases del Proceso Unificado",
                  "Cada iteración recorre las etapas para los casos de uso seleccionados y produce un incremento.")
phases = [
    ("Inicio", "Acordar visión, alcance, actores, casos de uso prioritarios y riesgos iniciales.",
     "Hito: objetivos aceptados", ROSE, PALE_ROSE),
    ("Elaboración", "Refinar requisitos y validar una arquitectura ejecutable mediante modelos UML.",
     "Hito: arquitectura estable", GREEN, PALE_GREEN),
    ("Construcción", "Implementar e integrar los casos de uso, con pruebas continuas y documentación.",
     "Hito: incremento operable", GOLD, PALE_GOLD),
    ("Transición", "Probar con usuarios, desplegar, capacitar y recoger desviaciones para la siguiente iteración.",
     "Hito: incremento aceptado", BROWN, RGBColor(238, 231, 226)),
]
for i, (title, body, milestone, color, fill) in enumerate(phases):
    x = 0.62 + i * 3.14
    add_box(slide, x, 2.12, 2.86, 3.80, fill, fill)
    add_text(slide, f"0{i + 1}", x + 0.20, 2.35, 0.50, 0.32, 12, color, True)
    add_text(slide, title, x + 0.20, 2.80, 2.35, 0.38, 18, INK, True)
    add_text(slide, body, x + 0.20, 3.42, 2.42, 1.35, 13, INK)
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x + 0.20), Inches(5.03),
                                  Inches(2.38), Inches(0.02))
    line.fill.solid()
    line.fill.fore_color.rgb = color
    line.line.fill.background()
    add_text(slide, milestone, x + 0.20, 5.25, 2.42, 0.35, 11, color, True)
for i in range(3):
    arrow = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(3.42 + i * 3.14),
                                   Inches(3.65), Inches(0.34), Inches(0.58))
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = WHITE
    arrow.line.fill.background()
add_text(slide, "Si en transición hay desviaciones, se actualizan casos de uso y arquitectura "
                "y se inicia una nueva iteración; si se acepta, se priorizan nuevos casos de uso.",
         0.85, 6.25, 11.65, 0.52, 13, ROSE, True, align=PP_ALIGN.CENTER)


# 7. Casos de uso
slide = new_slide(7, "Casos de uso", "Requisitos organizados por actores",
                  "El modelo de casos de uso conecta necesidades, diseño, implementación y pruebas.")
add_card(slide, "Cliente",
         "• Registrarse / autenticarse\n• Consultar productos y stock\n• Gestionar carro\n"
         "• Confirmar compra\n• Obtener boleta",
         0.65, 2.20, 3.20, 3.65, ROSE, PALE_ROSE, 18, 14)
add_card(slide, "Administrador",
         "• Autenticarse\n• Ingresar stock\n• Actualizar productos\n"
         "• Revisar alertas de stock crítico\n• Consultar ventas",
         9.48, 2.20, 3.20, 3.65, GREEN, PALE_GREEN, 18, 14)
system = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.45), Inches(2.13),
                                Inches(4.43), Inches(3.87))
system.fill.solid()
system.fill.fore_color.rgb = WHITE
system.line.color.rgb = BROWN
system.line.width = Pt(1.5)
add_text(slide, "SISTEMA DOLCE‑VITA", 5.00, 2.38, 3.35, 0.34, 15, BROWN, True,
         align=PP_ALIGN.CENTER)
for i, uc in enumerate(["Autenticar usuarios", "Procesar compra", "Controlar inventario",
                        "Emitir boleta", "Generar alerta crítica"]):
    y = 2.97 + i * 0.52
    oval = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.02), Inches(y),
                                  Inches(3.32), Inches(0.36))
    oval.fill.solid()
    oval.fill.fore_color.rgb = CREAM
    oval.line.color.rgb = LINE
    add_text(slide, uc, 5.14, y + 0.07, 3.05, 0.17, 10, INK, i in (1, 2),
             align=PP_ALIGN.CENTER)
for x1, x2, color in [(3.85, 4.45, ROSE), (8.88, 9.48, GREEN)]:
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(4.03),
                                      Inches(x2), Inches(4.03))
    conn.line.color.rgb = color
    conn.line.width = Pt(2)
add_box(slide, 1.82, 6.30, 9.72, 0.48, PALE_GOLD, PALE_GOLD)
add_text(slide, "Trazabilidad: caso de uso → escenario → diagrama → componente → prueba → aceptación",
         2.05, 6.42, 9.28, 0.20, 12, BROWN, True, align=PP_ALIGN.CENTER)


# 8. Iteraciones
slide = new_slide(8, "Plan", "Plan incremental propuesto",
                  "La prioridad combina valor para el negocio, dependencias y riesgo.")
iterations = [
    ("Iteración 1", "Acceso y base", "Registro, autenticación, roles y catálogo inicial.",
     "Riesgo: seguridad", ROSE),
    ("Iteración 2", "Compra", "Consulta de stock, carro, confirmación y boleta.",
     "Riesgo: consistencia", GREEN),
    ("Iteración 3", "Inventario", "Ingreso de stock, actualización y alertas críticas.",
     "Riesgo: concurrencia", GOLD),
    ("Iteración 4", "Entrega", "Pruebas integrales, piloto, capacitación y despliegue.",
     "Riesgo: adopción", BROWN),
]
for i, (it, title, body, risk, color) in enumerate(iterations):
    y = 2.05 + i * 1.15
    add_box(slide, 0.68, y, 11.95, 0.94, WHITE)
    add_text(slide, it, 0.94, y + 0.18, 1.15, 0.30, 13, color, True)
    add_text(slide, title, 2.14, y + 0.18, 1.55, 0.30, 15, INK, True)
    add_text(slide, body, 3.78, y + 0.16, 5.77, 0.42, 13, INK)
    add_box(slide, 9.84, y + 0.18, 2.30, 0.49,
            [PALE_ROSE, PALE_GREEN, PALE_GOLD, RGBColor(238, 231, 226)][i],
            [PALE_ROSE, PALE_GREEN, PALE_GOLD, RGBColor(238, 231, 226)][i])
    add_text(slide, risk, 9.98, y + 0.30, 2.02, 0.20, 10, color, True,
             align=PP_ALIGN.CENTER)
add_text(slide, "Resultado de cada iteración", 0.72, 6.70, 2.30, 0.22, 11, ROSE, True)
add_text(slide, "Incremento probado + documentación actualizada + retroalimentación + nueva línea base.",
         3.04, 6.63, 9.20, 0.31, 13, INK, True)


# 9. SDLC
slide = new_slide(9, "Ciclo de vida", "Relación con el SDLC clásico",
                  "El PU conserva las actividades del ciclo de vida, pero cambia su organización.")
data = [
    ["SDLC", "Aplicación en el PU", "Ejemplo Dolce‑Vita"],
    ["Planificación", "Inicio + gestión del proyecto", "Visión, alcance, riesgos y plan de iteraciones"],
    ["Análisis", "Requisitos en cada iteración", "Actores, reglas de stock y escenarios de compra"],
    ["Diseño", "Elaboración y refinamiento", "Arquitectura, clases, secuencias e interfaces"],
    ["Construcción", "Incrementos integrados", "Módulos de acceso, carro, inventario y boleta"],
    ["Pruebas", "Continuas + transición", "Unitarias, integración, sistema y aceptación"],
    ["Despliegue", "Transición", "Piloto, capacitación, versión y retroalimentación"],
    ["Mantenimiento", "Nuevas iteraciones", "Cambios priorizados y trazables"],
]
add_table(slide, data, 0.65, 2.05, 12.02, 4.33, [2.0, 3.25, 6.77], 11.3, GREEN)
add_box(slide, 0.65, 6.56, 12.02, 0.36, PALE_ROSE, PALE_ROSE)
add_text(slide, "Innovación frente al enfoque clásico: no se espera al final; análisis, diseño, "
                "código y prueba se revisan en ciclos con retroalimentación temprana.",
         0.88, 6.62, 11.58, 0.18, 10, ROSE, True, align=PP_ALIGN.CENTER)


# 10. Artefactos
slide = new_slide(10, "Artefactos", "Artefactos que permiten operar y gestionar el SDLC",
                  "Cada artefacto tiene propósito, responsable, versión y relación con casos de uso.")
artifact_groups = [
    ("Requisitos", "Visión y alcance\nModelo de casos de uso\nEspecificaciones suplementarias\nGlosario",
     ROSE, PALE_ROSE),
    ("Análisis y diseño", "Modelo de dominio\nDiagramas UML\nDocumento de arquitectura\nPrototipos de interfaz",
     GREEN, PALE_GREEN),
    ("Construcción", "Código fuente\nEstándares de código\nBuild e integración continua\nPaquetes versionados",
     GOLD, PALE_GOLD),
    ("Pruebas y entrega", "Plan y casos de prueba\nRegistro de defectos\nManual de usuario\nPlan de despliegue",
     BROWN, RGBColor(238, 231, 226)),
]
for i, (title, body, color, fill) in enumerate(artifact_groups):
    x = 0.62 + i * 3.14
    add_card(slide, title, body, x, 2.12, 2.85, 3.55, color, fill, 16, 13)
add_box(slide, 0.62, 5.95, 12.00, 0.78, WHITE)
add_rich_text(slide, [
    ("Artefactos transversales: ", True, BROWN),
    ("plan de iteración, lista de riesgos, tablero de seguimiento, solicitud de cambio, "
     "historial de versiones, auditoría y línea base.", False, INK)
], 0.90, 6.14, 11.48, 0.36, 13)


# 11. Soporte
slide = new_slide(11, "Soporte", "Procesos de soporte que mantienen el control",
                  "No entregan una función visible por sí solos, pero sostienen la calidad de cada incremento.")
support = [
    ("Configuración y cambios",
     "Identifica artefactos, controla versiones, audita líneas base y aplica el flujo: "
     "solicitud → impacto → aprobación → implementación.",
     "Git/GitHub · registro de cambios", ROSE, PALE_ROSE),
    ("Gestión del proyecto",
     "Planifica fases e iteraciones, asigna recursos, controla avance, riesgos, costos y comunicación.",
     "Jira/Trello · Kanban · métricas", GREEN, PALE_GREEN),
    ("Ambiente",
     "Provee estándares e infraestructura para modelar, integrar, probar y desplegar de forma reproducible.",
     "CASE · CI/CD · Docker · staging", GOLD, PALE_GOLD),
]
for i, (title, body, tools, color, fill) in enumerate(support):
    x = 0.68 + i * 4.05
    add_box(slide, x, 2.10, 3.72, 3.95, fill, fill)
    add_text(slide, title, x + 0.25, 2.43, 3.20, 0.58, 17, color, True)
    add_text(slide, body, x + 0.25, 3.32, 3.18, 1.35, 14, INK)
    add_text(slide, "Herramientas", x + 0.25, 5.00, 1.30, 0.22, 10, color, True)
    add_text(slide, tools, x + 0.25, 5.31, 3.15, 0.34, 12, BROWN, True)
add_text(slide, "Integración: la configuración aporta orden; la gestión, dirección; el ambiente, capacidad de ejecución.",
         0.85, 6.44, 11.65, 0.36, 13, INK, True, align=PP_ALIGN.CENTER)


# 12. Comparación
slide = new_slide(12, "Comparación", "Comparación de modelos de desarrollo",
                  "La elección depende de estabilidad, riesgo, regulación y necesidad de retroalimentación.")
data = [
    ["Criterio", "Cascada", "Iterativo/incremental", "Ágil", "Proceso Unificado"],
    ["Flujo", "Secuencial", "Ciclos + incrementos", "Iteraciones cortas", "Iterativo por casos de uso"],
    ["Cambio", "Costoso y tardío", "Se incorpora por ciclos", "Esperado y frecuente", "Controlado y trazable"],
    ["Arquitectura", "Diseño previo", "Evoluciona", "Emergente/intencional", "Central y temprana"],
    ["Documentación", "Amplia por fase", "Según incremento", "La necesaria", "Artefactos versionados"],
    ["Entrega", "Principalmente final", "Parcial y progresiva", "Frecuente", "Incrementos ejecutables"],
    ["Aplicación", "Requisitos estables", "Valor por partes", "Alta incertidumbre", "Complejidad y trazabilidad"],
]
add_table(slide, data, 0.42, 2.00, 12.50, 4.48,
          [1.62, 2.05, 2.66, 2.12, 4.05], 10.4, BROWN)
add_text(slide, "Para Dolce‑Vita, el PU equilibra adaptación con disciplina documental y arquitectónica.",
         1.15, 6.67, 11.10, 0.25, 12, ROSE, True, align=PP_ALIGN.CENTER)


# 13. Ventajas y desventajas
slide = new_slide(13, "Evaluación", "Ventajas y desventajas del Proceso Unificado",
                  "El beneficio depende de adaptar el nivel de formalidad al tamaño del proyecto.")
add_box(slide, 0.68, 2.10, 5.78, 4.45, PALE_GREEN, PALE_GREEN)
add_text(slide, "Ventajas", 1.00, 2.42, 2.00, 0.40, 20, GREEN, True)
add_bullets(slide, [
    ("Entrega: ", "incrementos reducen el riesgo de descubrir problemas al final."),
    ("Cambio: ", "retroalimentación se incorpora mediante nuevas iteraciones."),
    ("Calidad: ", "trazabilidad y pruebas continuas protegen el producto."),
    ("Riesgo: ", "arquitectura y decisiones críticas se validan temprano."),
    ("Comunicación: ", "casos de uso crean lenguaje común con el cliente.")
], 1.00, 3.07, 5.05, 2.92, 11, INK, GREEN, 6)
add_box(slide, 6.85, 2.10, 5.78, 4.45, PALE_ROSE, PALE_ROSE)
add_text(slide, "Desventajas", 7.17, 2.42, 2.30, 0.40, 20, ROSE, True)
add_bullets(slide, [
    ("Tiempo: ", "modelar y mantener artefactos exige esfuerzo sostenido."),
    ("Complejidad: ", "muchos roles y prácticas pueden sobredimensionar el proyecto."),
    ("Disciplina: ", "documentación desactualizada rompe la trazabilidad."),
    ("Especialización: ", "requiere experiencia en UML, arquitectura y gestión."),
    ("Adaptación: ", "si se aplica rígidamente, reduce velocidad y colaboración.")
], 7.17, 3.07, 5.05, 2.92, 11, INK, ROSE, 6)
add_text(slide, "Mitigación: seleccionar solo artefactos útiles, automatizar controles y revisar su vigencia en cada iteración.",
         0.86, 6.72, 11.68, 0.25, 11, BROWN, True, align=PP_ALIGN.CENTER)


# 14. Calidad y ciclo de cambio
slide = new_slide(14, "Calidad", "La desviación no se corrige solo en el código",
                  "Un cambio real debe conservar la coherencia entre requisito, arquitectura, implementación y prueba.")
cycle = [
    ("1", "Caso de uso", "Ajustar escenario y criterio", ROSE),
    ("2", "Arquitectura", "Revisar modelos y decisiones", GREEN),
    ("3", "Código", "Implementar el cambio aprobado", GOLD),
    ("4", "Prueba", "Verificar trazabilidad y aceptación", BROWN),
]
for i, (num, title, body, color) in enumerate(cycle):
    x = 0.70 + i * 3.08
    add_box(slide, x, 2.16, 2.74, 2.10, WHITE)
    add_text(slide, num, x + 0.20, 2.38, 0.35, 0.30, 12, color, True)
    add_text(slide, title, x + 0.68, 2.36, 1.80, 0.34, 16, color, True)
    add_text(slide, body, x + 0.20, 3.10, 2.30, 0.65, 13, INK)
    if i < 3:
        arrow = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x + 2.79),
                                       Inches(2.88), Inches(0.32), Inches(0.50))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = LINE
        arrow.line.fill.background()
add_box(slide, 0.70, 4.64, 5.78, 1.54, PALE_ROSE, PALE_ROSE)
add_text(slide, "Si hay desviaciones", 1.00, 4.90, 2.20, 0.62, 17, ROSE, True)
add_text(slide, "Se registra el hallazgo y se genera una nueva iteración para corregir "
                "los artefactos relacionados.", 1.00, 5.52, 5.10, 0.44, 12, INK)
add_box(slide, 6.85, 4.64, 5.78, 1.54, PALE_GREEN, PALE_GREEN)
add_text(slide, "Si el incremento cumple", 7.15, 4.90, 2.80, 0.62, 17, GREEN, True)
add_text(slide, "Se establece una línea base y comienza una iteración con nuevos casos de uso "
                "priorizados.", 7.15, 5.52, 5.10, 0.44, 12, INK)
add_text(slide, "Principio: documentación deficiente → decisiones deficientes → código deficiente.",
         1.10, 6.63, 11.00, 0.30, 12, BROWN, True, align=PP_ALIGN.CENTER)


# 15. Conclusión
slide = new_slide(15, "Reflexión", "Conclusión y reflexión del equipo",
                  "La metodología es valiosa cuando la documentación guía decisiones reales.")
add_box(slide, 0.70, 2.12, 7.34, 4.25, WHITE)
add_text(slide, "Conclusión", 1.03, 2.46, 2.00, 0.40, 19, ROSE, True)
add_text(slide,
         "El Proceso Unificado permite desarrollar Dolce‑Vita de forma incremental sin perder "
         "control sobre requisitos, arquitectura ni calidad. Los casos de uso orientan qué se "
         "construye; la arquitectura explica cómo se sostiene; y las iteraciones permiten comprobar "
         "ambas decisiones con usuarios reales.\n\n"
         "Nuestra reflexión es que documentar no significa acumular diagramas. Su valor aparece "
         "cuando cada artefacto se mantiene vigente, ayuda a decidir y puede rastrearse hasta una "
         "prueba. Así, el cambio deja de ser una corrección improvisada y se convierte en aprendizaje "
         "controlado para la siguiente iteración.",
         1.03, 3.15, 6.67, 2.63, 15, INK)
add_box(slide, 8.45, 2.12, 4.18, 4.25, BROWN, BROWN)
add_text(slide, "IDEA CENTRAL", 8.82, 2.60, 3.45, 0.30, 11, GOLD, True,
         align=PP_ALIGN.CENTER)
add_text(slide, "Calidad no es\ncorregir al final.", 8.78, 3.30, 3.55, 0.95,
         23, WHITE, True, align=PP_ALIGN.CENTER)
line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(9.53), Inches(4.49),
                              Inches(2.05), Inches(0.04))
line.fill.solid()
line.fill.fore_color.rgb = GOLD
line.line.fill.background()
add_text(slide, "Es conservar coherencia desde el caso de uso hasta el despliegue.",
         8.93, 4.90, 3.23, 0.82, 14, WHITE, False, align=PP_ALIGN.CENTER)


# 16. Referencias
slide = new_slide(16, "Referencias", "Bibliografía en formato APA 7",
                  "Fuentes utilizadas para fundamentar la propuesta.")
refs = [
    "IEEE. (2005). IEEE standard for software configuration management plans (IEEE Std 828-2005).",
    "Jacobson, I., Booch, G., & Rumbaugh, J. (2000). El proceso unificado de desarrollo de software. Pearson Educación.",
    "Jacobson, I., Booch, G., & Rumbaugh, J. (2006). El lenguaje unificado de modelado: Guía del usuario. Pearson.",
    "Project Management Institute. (2017). A guide to the project management body of knowledge (PMBOK® Guide) (6.ª ed.).",
    "Pressman, R. S. (2010). Ingeniería del software: Un enfoque práctico (7.ª ed.). McGraw-Hill.",
    "Sommerville, I. (2011). Ingeniería de software (9.ª ed.). Pearson Educación.",
    "Santo Tomás Online. (s. f.). Diseño de software: Apunte semanal, semana 3 [Material de asignatura].",
]
add_box(slide, 0.72, 2.08, 11.90, 4.58, WHITE)
box = slide.shapes.add_textbox(Inches(1.02), Inches(2.38), Inches(11.27), Inches(3.93))
tf = box.text_frame
tf.clear()
tf.margin_left = Inches(0.02)
tf.margin_right = Inches(0.02)
for i, ref in enumerate(refs):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.text = ref
    p.space_after = Pt(10)
    p.level = 0
    p.margin_left = Inches(0.30)
    p.first_line_indent = Inches(-0.30)
    for r in p.runs:
        r.font.name = "Aptos"
        r.font.size = Pt(13)
        r.font.color.rgb = INK
add_text(slide, "Revisar datos de edición y URL institucional antes de la entrega final.",
         1.02, 6.46, 11.20, 0.20, 10, MUTED)


prs.save(OUT)
print(f"Presentación creada: {OUT}")
