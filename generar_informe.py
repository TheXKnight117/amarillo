from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import (
    Flowable,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


OUTPUT = "Informe_S1_SDLC_EasyTrip.pdf"
PAGE_SIZE = landscape(A4)

NAVY = colors.HexColor("#12304A")
TEAL = colors.HexColor("#007C83")
CYAN = colors.HexColor("#2CA6A4")
SKY = colors.HexColor("#E8F5F5")
GOLD = colors.HexColor("#F2B544")
ORANGE = colors.HexColor("#E77728")
RED = colors.HexColor("#C94C4C")
INK = colors.HexColor("#243746")
MUTED = colors.HexColor("#5D7180")
PALE = colors.HexColor("#F4F7F8")
WHITE = colors.white


def register_fonts():
    regular = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    bold = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
    try:
        pdfmetrics.registerFont(TTFont("DejaVu", regular))
        pdfmetrics.registerFont(TTFont("DejaVu-Bold", bold))
        return "DejaVu", "DejaVu-Bold"
    except Exception:
        return "Helvetica", "Helvetica-Bold"


FONT, FONT_BOLD = register_fonts()


styles = getSampleStyleSheet()
styles.add(
    ParagraphStyle(
        name="CoverKicker",
        fontName=FONT_BOLD,
        fontSize=12,
        leading=15,
        textColor=TEAL,
        spaceAfter=12,
        alignment=TA_CENTER,
    )
)
styles.add(
    ParagraphStyle(
        name="CoverTitle",
        fontName=FONT_BOLD,
        fontSize=27,
        leading=32,
        textColor=NAVY,
        alignment=TA_CENTER,
        spaceAfter=15,
    )
)
styles.add(
    ParagraphStyle(
        name="CoverSub",
        fontName=FONT,
        fontSize=14,
        leading=19,
        textColor=MUTED,
        alignment=TA_CENTER,
    )
)
styles.add(
    ParagraphStyle(
        name="PageTitle",
        fontName=FONT_BOLD,
        fontSize=19,
        leading=23,
        textColor=NAVY,
        spaceAfter=5,
    )
)
styles.add(
    ParagraphStyle(
        name="PageLead",
        fontName=FONT,
        fontSize=9.5,
        leading=13,
        textColor=MUTED,
        spaceAfter=9,
    )
)
styles.add(
    ParagraphStyle(
        name="Body",
        fontName=FONT,
        fontSize=9,
        leading=12.5,
        textColor=INK,
    )
)
styles.add(
    ParagraphStyle(
        name="BodySmall",
        fontName=FONT,
        fontSize=7.7,
        leading=10.5,
        textColor=INK,
    )
)
styles.add(
    ParagraphStyle(
        name="BodySmallWhite",
        fontName=FONT,
        fontSize=7.7,
        leading=10.5,
        textColor=WHITE,
    )
)
styles.add(
    ParagraphStyle(
        name="CardTitle",
        fontName=FONT_BOLD,
        fontSize=11.2,
        leading=14,
        textColor=WHITE,
    )
)
styles.add(
    ParagraphStyle(
        name="ColHeader",
        fontName=FONT_BOLD,
        fontSize=7.5,
        leading=9,
        textColor=TEAL,
    )
)
styles.add(
    ParagraphStyle(
        name="CenterSmall",
        fontName=FONT,
        fontSize=7.4,
        leading=9.5,
        textColor=INK,
        alignment=TA_CENTER,
    )
)
styles.add(
    ParagraphStyle(
        name="CenterSmallWhite",
        fontName=FONT_BOLD,
        fontSize=8,
        leading=10,
        textColor=WHITE,
        alignment=TA_CENTER,
    )
)
styles.add(
    ParagraphStyle(
        name="Quote",
        fontName=FONT,
        fontSize=11,
        leading=16,
        textColor=NAVY,
        alignment=TA_CENTER,
    )
)


def bullet_lines(items, style="BodySmall"):
    return Paragraph("<br/>".join(f"• {item}" for item in items), styles[style])


class AccentRule(Flowable):
    def __init__(self, width=4.0 * cm):
        super().__init__()
        self.width = width
        self.height = 0.12 * cm

    def draw(self):
        self.canv.setFillColor(GOLD)
        self.canv.roundRect(0, 0, self.width, self.height, self.height / 2, fill=1, stroke=0)


class CycleDiagram(Flowable):
    labels = [
        ("1", "PLANIFICACIÓN"),
        ("2", "ANÁLISIS"),
        ("3", "DISEÑO"),
        ("4", "DESARROLLO"),
        ("5", "PRUEBAS"),
        ("6", "DESPLIEGUE Y\nMANTENIMIENTO"),
    ]

    def __init__(self, width, height=4.15 * cm):
        super().__init__()
        self.width = width
        self.height = height

    def draw(self):
        c = self.canv
        gap = 10
        box_w = (self.width - gap * 5) / 6
        box_h = 2.15 * cm
        y = 1.1 * cm
        for i, (num, label) in enumerate(self.labels):
            x = i * (box_w + gap)
            fill = NAVY if i % 2 == 0 else TEAL
            c.setFillColor(fill)
            c.roundRect(x, y, box_w, box_h, 7, fill=1, stroke=0)
            c.setFillColor(GOLD)
            c.circle(x + box_w / 2, y + box_h - 15, 10, fill=1, stroke=0)
            c.setFillColor(NAVY)
            c.setFont(FONT_BOLD, 8)
            c.drawCentredString(x + box_w / 2, y + box_h - 18, num)
            c.setFillColor(WHITE)
            c.setFont(FONT_BOLD, 7.4)
            lines = label.split("\n")
            baseline = y + 23 + (len(lines) - 1) * 5
            for line_index, line in enumerate(lines):
                c.drawCentredString(x + box_w / 2, baseline - line_index * 10, line)
            if i < 5:
                c.setStrokeColor(GOLD)
                c.setLineWidth(2)
                arrow_x = x + box_w + 2
                mid_y = y + box_h / 2
                c.line(arrow_x, mid_y, arrow_x + gap - 4, mid_y)
                c.line(arrow_x + gap - 7, mid_y + 3, arrow_x + gap - 4, mid_y)
                c.line(arrow_x + gap - 7, mid_y - 3, arrow_x + gap - 4, mid_y)
        c.setStrokeColor(CYAN)
        c.setLineWidth(1.5)
        c.line(self.width - box_w / 2, y - 8, box_w / 2, y - 8)
        c.line(box_w / 2, y - 8, box_w / 2, y - 1)
        c.line(self.width - box_w / 2, y - 8, self.width - box_w / 2, y - 1)
        c.setFillColor(TEAL)
        c.setFont(FONT, 7.2)
        c.drawCentredString(
            self.width / 2,
            y - 21,
            "Retroalimentación continua: los hallazgos permiten ajustar requisitos, diseño y prioridades.",
        )


phases = [
    {
        "number": "01",
        "name": "Planificación",
        "purpose": "Definir el problema, el alcance, la estrategia y la viabilidad de «Reserva tu viaje».",
        "activities": [
            "Acordar el alcance inicial: búsqueda de vuelos, reservas hoteleras y planificador inteligente.",
            "Identificar interesados: viajeros, EasyTrip, aerolíneas, hoteles y proveedores de mapas.",
            "Estimar tiempos, presupuesto, recursos y construir un backlog inicial priorizado.",
        ],
        "roles": ["Product Owner", "Jefe/a de proyecto", "Analista de negocio", "Representante de EasyTrip"],
        "deliverables": ["Acta de constitución", "Alcance y cronograma", "Backlog inicial", "Mapa de interesados"],
        "risks": [
            "Alcance demasiado amplio para la primera versión.",
            "Estimaciones poco realistas o dependencia de proveedores externos.",
        ],
    },
    {
        "number": "02",
        "name": "Análisis de requisitos",
        "purpose": "Comprender qué necesitan los usuarios y convertirlo en requisitos claros y verificables.",
        "activities": [
            "Entrevistar a viajeros, agencias, hoteles y responsables del negocio.",
            "Crear historias de usuario y criterios de aceptación para búsqueda, reserva e itinerarios.",
            "Definir requisitos no funcionales: seguridad, privacidad, rendimiento, accesibilidad y disponibilidad.",
        ],
        "roles": ["Analista de negocio", "Product Owner", "UX Researcher", "Especialista en seguridad"],
        "deliverables": ["Especificación de requisitos", "Historias de usuario", "Casos de uso", "Criterios de aceptación"],
        "risks": [
            "Requisitos ambiguos, contradictorios o cambiantes.",
            "Omitir necesidades de accesibilidad, protección de datos o cancelaciones.",
        ],
    },
    {
        "number": "03",
        "name": "Diseño",
        "purpose": "Definir cómo funcionará la solución antes de construirla.",
        "activities": [
            "Diseñar flujos y prototipos para buscar, comparar, reservar y consultar itinerarios.",
            "Definir arquitectura móvil, API, base de datos e integración con vuelos, hoteles, pagos y mapas.",
            "Modelar el planificador inteligente y establecer controles de seguridad y trazabilidad.",
        ],
        "roles": ["Diseñador/a UX/UI", "Arquitecto/a de software", "Diseñador/a de datos", "Especialista en seguridad"],
        "deliverables": ["Prototipo navegable", "Arquitectura técnica", "Modelo de datos", "Especificación de interfaces"],
        "risks": [
            "Experiencia de usuario compleja o poco accesible.",
            "Arquitectura que no escale o integraciones incompatibles.",
        ],
    },
    {
        "number": "04",
        "name": "Desarrollo",
        "purpose": "Construir incrementos funcionales y seguros según las prioridades del producto.",
        "activities": [
            "Programar la app móvil, servicios backend, autenticación y notificaciones.",
            "Integrar catálogos de vuelos y hoteles, pagos, mapas y motor de itinerarios.",
            "Aplicar revisión de código, integración continua y pruebas unitarias en cada sprint.",
        ],
        "roles": ["Desarrollador/a móvil", "Desarrollador/a backend", "Ingeniero/a de datos/IA", "DevOps"],
        "deliverables": ["Código versionado", "Incrementos ejecutables", "APIs integradas", "Documentación técnica"],
        "risks": [
            "Retrasos o cambios en APIs de terceros.",
            "Deuda técnica, vulnerabilidades o baja calidad del código.",
        ],
    },
    {
        "number": "05",
        "name": "Pruebas",
        "purpose": "Comprobar que el producto cumple los requisitos y funciona de forma confiable.",
        "activities": [
            "Ejecutar pruebas funcionales de búsqueda, reserva, pago, cancelación e itinerarios.",
            "Realizar pruebas de integración, usabilidad, rendimiento, seguridad y compatibilidad móvil.",
            "Registrar defectos, corregirlos y validar la aceptación con usuarios representativos.",
        ],
        "roles": ["QA / Tester", "Automatizador/a QA", "Especialista en seguridad", "Usuarios piloto y Product Owner"],
        "deliverables": ["Plan y casos de prueba", "Registro de defectos", "Informe de resultados", "Acta de aceptación"],
        "risks": [
            "Cobertura insuficiente o ambientes de prueba poco realistas.",
            "Fallos en pagos, sobreventa o exposición de datos personales.",
        ],
    },
    {
        "number": "06",
        "name": "Despliegue y mantenimiento",
        "purpose": "Publicar la aplicación, operar el servicio y mejorarlo con evidencia real.",
        "activities": [
            "Preparar ambientes, migraciones, monitoreo y publicación gradual en las tiendas.",
            "Capacitar a soporte y definir respuesta a incidentes, respaldo y recuperación.",
            "Medir uso, errores y satisfacción; priorizar mejoras y nuevos incrementos del backlog.",
        ],
        "roles": ["DevOps / SRE", "Soporte técnico", "Product Owner", "Equipo de desarrollo y QA"],
        "deliverables": ["Versión productiva", "Manual de operación", "Panel de monitoreo", "Plan de mantenimiento"],
        "risks": [
            "Caídas del servicio o errores durante la publicación.",
            "Baja adopción, costos operativos altos o incidentes sin respuesta oportuna.",
        ],
    },
]


def header_footer(canvas, doc):
    canvas.saveState()
    w, h = PAGE_SIZE
    canvas.setFillColor(NAVY)
    canvas.rect(0, h - 0.22 * cm, w, 0.22 * cm, fill=1, stroke=0)
    canvas.setStrokeColor(colors.HexColor("#D9E3E8"))
    canvas.line(doc.leftMargin, 0.72 * cm, w - doc.rightMargin, 0.72 * cm)
    canvas.setFont(FONT, 7.3)
    canvas.setFillColor(MUTED)
    canvas.drawString(doc.leftMargin, 0.4 * cm, "EasyTrip · Diseño del ciclo de vida del software")
    canvas.drawRightString(w - doc.rightMargin, 0.4 * cm, f"Página {doc.page}")
    canvas.restoreState()


def phase_card(phase):
    heading = Table(
        [
            [
                Paragraph(phase["number"], styles["CenterSmallWhite"]),
                Paragraph(phase["name"], styles["CardTitle"]),
                Paragraph(phase["purpose"], styles["BodySmallWhite"]),
            ]
        ],
        colWidths=[1.1 * cm, 4.5 * cm, 20.1 * cm],
    )
    heading.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, 0), GOLD),
                ("BACKGROUND", (1, 0), (-1, 0), NAVY),
                ("TEXTCOLOR", (0, 0), (0, 0), NAVY),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    detail = Table(
        [
            [
                Paragraph("ACTIVIDADES DEL CASO", styles["ColHeader"]),
                Paragraph("ROLES", styles["ColHeader"]),
                Paragraph("ENTREGABLES", styles["ColHeader"]),
                Paragraph("RIESGOS / CONTROLES", styles["ColHeader"]),
            ],
            [
                bullet_lines(phase["activities"]),
                bullet_lines(phase["roles"]),
                bullet_lines(phase["deliverables"]),
                bullet_lines(phase["risks"]),
            ],
        ],
        colWidths=[10.0 * cm, 5.0 * cm, 5.1 * cm, 5.6 * cm],
    )
    detail.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), SKY),
                ("BACKGROUND", (0, 1), (-1, 1), WHITE),
                ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#C9D8DE")),
                ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#DCE6EA")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    return KeepTogether([heading, detail, Spacer(1, 0.3 * cm)])


def build_story():
    story = []

    # Cover
    story.extend(
        [
            Spacer(1, 2.0 * cm),
            Paragraph("ACTIVIDAD FORMATIVA · SEMANA 1", styles["CoverKicker"]),
            Paragraph("Diseño del ciclo de vida<br/>del software", styles["CoverTitle"]),
            AccentRule(5.0 * cm),
            Spacer(1, 0.55 * cm),
            Paragraph("Caso aplicado: <b>EasyTrip — «Reserva tu viaje»</b>", styles["CoverSub"]),
            Spacer(1, 1.2 * cm),
        ]
    )
    info = Table(
        [
            [Paragraph("<b>Asignatura / código</b>", styles["Body"]), Paragraph("IGRV0132", styles["Body"])],
            [Paragraph("<b>Modalidad</b>", styles["Body"]), Paragraph("Trabajo individual", styles["Body"])],
            [Paragraph("<b>Estudiante</b>", styles["Body"]), Paragraph("[Escribe aquí tu nombre]", styles["Body"])],
            [Paragraph("<b>Fecha</b>", styles["Body"]), Paragraph("Octubre de 2026", styles["Body"])],
        ],
        colWidths=[4.6 * cm, 10.0 * cm],
        hAlign="CENTER",
    )
    info.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, -1), NAVY),
                ("TEXTCOLOR", (0, 0), (0, -1), WHITE),
                ("BACKGROUND", (1, 0), (1, -1), PALE),
                ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#CCD9DE")),
                ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D7E1E5")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    story.extend(
        [
            info,
            Spacer(1, 1.15 * cm),
            Paragraph(
                "Objetivo: representar visualmente las fases, actividades, roles, entregables y riesgos "
                "del SDLC para una aplicación móvil de búsqueda de vuelos, reservas hoteleras y "
                "planificación inteligente de itinerarios.",
                styles["Quote"],
            ),
            PageBreak(),
        ]
    )

    # Overview and model selection
    story.extend(
        [
            Paragraph("1. Visión general del ciclo de vida", styles["PageTitle"]),
            Paragraph(
                "El SDLC ordena el proyecto desde la definición del problema hasta la operación y mejora "
                "continua. En EasyTrip, cada fase produce evidencia verificable para disminuir la "
                "incertidumbre antes de avanzar.",
                styles["PageLead"],
            ),
            CycleDiagram(25.7 * cm),
            Spacer(1, 0.2 * cm),
            Paragraph("Modelo recomendado: ágil, iterativo e incremental", styles["PageTitle"]),
            Paragraph(
                "El mercado turístico cambia con rapidez y la solución depende de integraciones externas. "
                "Por ello se propone trabajar con Scrum en sprints cortos, entregando primero un producto "
                "mínimo viable y agregando capacidades según la retroalimentación de usuarios y negocio.",
                styles["PageLead"],
            ),
        ]
    )
    comparison_data = [
        [
            Paragraph("MODELO", styles["ColHeader"]),
            Paragraph("FORTALEZA", styles["ColHeader"]),
            Paragraph("LIMITACIÓN EN EASYTRIP", styles["ColHeader"]),
            Paragraph("APLICABILIDAD", styles["ColHeader"]),
        ],
        [
            Paragraph("<b>Cascada</b>", styles["BodySmall"]),
            Paragraph("Orden y documentación; útil con requisitos estables.", styles["BodySmall"]),
            Paragraph("Cambios tardíos resultan costosos y la validación llega al final.", styles["BodySmall"]),
            Paragraph("Baja para el producto completo; útil en hitos regulatorios.", styles["BodySmall"]),
        ],
        [
            Paragraph("<b>Iterativo e incremental</b>", styles["BodySmall"]),
            Paragraph("Permite construir, evaluar y ampliar la solución por versiones.", styles["BodySmall"]),
            Paragraph("Exige una arquitectura preparada para crecer.", styles["BodySmall"]),
            Paragraph("Alta: vuelos, hoteles e itinerarios pueden incorporarse gradualmente.", styles["BodySmall"]),
        ],
        [
            Paragraph("<b>Ágil (Scrum)</b>", styles["BodySmall"]),
            Paragraph("Entrega temprana, prioridades flexibles y retroalimentación frecuente.", styles["BodySmall"]),
            Paragraph("Requiere participación constante del Product Owner y disciplina del equipo.", styles["BodySmall"]),
            Paragraph("<b>Muy alta: modelo recomendado para EasyTrip.</b>", styles["BodySmall"]),
        ],
    ]
    comparison = Table(comparison_data, colWidths=[3.3 * cm, 7.0 * cm, 7.5 * cm, 7.9 * cm])
    comparison.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), SKY),
                ("BACKGROUND", (0, 3), (-1, 3), colors.HexColor("#FFF7E5")),
                ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#C9D8DE")),
                ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#DCE6EA")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.extend([comparison, PageBreak()])

    # Detailed phases
    story.extend(
        [
            Paragraph("2. Fases del SDLC: definición y diseño", styles["PageTitle"]),
            Paragraph(
                "Las primeras tres fases convierten una necesidad de negocio en una solución viable, "
                "comprensible y preparada para construirse.",
                styles["PageLead"],
            ),
        ]
    )
    for phase in phases[:3]:
        story.append(phase_card(phase))
    story.append(PageBreak())

    story.extend(
        [
            Paragraph("3. Fases del SDLC: construcción y operación", styles["PageTitle"]),
            Paragraph(
                "Las últimas tres fases producen el software, verifican su calidad y aseguran su continuidad "
                "en producción.",
                styles["PageLead"],
            ),
        ]
    )
    for phase in phases[3:]:
        story.append(phase_card(phase))
    story.append(PageBreak())

    # Governance, risks and close
    story.extend(
        [
            Paragraph("4. Gestión del proyecto y controles transversales", styles["PageTitle"]),
            Paragraph(
                "La coordinación de personas, decisiones y evidencia permite que cada incremento entregue "
                "valor sin comprometer calidad, seguridad ni continuidad operacional.",
                styles["PageLead"],
            ),
        ]
    )
    governance = Table(
        [
            [
                Paragraph("ASPECTO", styles["ColHeader"]),
                Paragraph("APLICACIÓN EN EASYTRIP", styles["ColHeader"]),
                Paragraph("RESPONSABLE PRINCIPAL", styles["ColHeader"]),
            ],
            [
                Paragraph("<b>Planificación</b>", styles["BodySmall"]),
                Paragraph(
                    "Backlog priorizado por valor y riesgo; sprints de duración fija; revisión de capacidad.",
                    styles["BodySmall"],
                ),
                Paragraph("Product Owner y Scrum Master / jefe de proyecto.", styles["BodySmall"]),
            ],
            [
                Paragraph("<b>Asignación de recursos</b>", styles["BodySmall"]),
                Paragraph(
                    "Equipo multidisciplinario con desarrollo móvil/backend, UX, QA, datos/IA, seguridad y DevOps.",
                    styles["BodySmall"],
                ),
                Paragraph("Jefe de proyecto y líderes técnicos.", styles["BodySmall"]),
            ],
            [
                Paragraph("<b>Seguimiento</b>", styles["BodySmall"]),
                Paragraph(
                    "Tablero de trabajo, revisión del sprint, demostración del incremento y retrospectiva.",
                    styles["BodySmall"],
                ),
                Paragraph("Scrum Master y equipo de desarrollo.", styles["BodySmall"]),
            ],
            [
                Paragraph("<b>Calidad y cambios</b>", styles["BodySmall"]),
                Paragraph(
                    "Definición de terminado, pruebas automatizadas, revisión de código y cambios re-priorizados.",
                    styles["BodySmall"],
                ),
                Paragraph("QA, líderes técnicos y Product Owner.", styles["BodySmall"]),
            ],
        ],
        colWidths=[4.0 * cm, 15.0 * cm, 6.7 * cm],
    )
    governance.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), SKY),
                ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#C9D8DE")),
                ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#DCE6EA")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.extend([governance, Spacer(1, 0.55 * cm)])

    risk_data = [
        [
            Paragraph("RIESGO CRÍTICO", styles["ColHeader"]),
            Paragraph("PROB.", styles["ColHeader"]),
            Paragraph("IMPACTO", styles["ColHeader"]),
            Paragraph("RESPUESTA PROPUESTA", styles["ColHeader"]),
        ],
        [
            Paragraph("Cambios en APIs de vuelos, hoteles, pagos o mapas.", styles["BodySmall"]),
            Paragraph("Alta", styles["CenterSmall"]),
            Paragraph("Alto", styles["CenterSmall"]),
            Paragraph("Adaptadores desacoplados, contratos de API, monitoreo y proveedor alternativo.", styles["BodySmall"]),
        ],
        [
            Paragraph("Exposición de datos personales o fraude en pagos.", styles["BodySmall"]),
            Paragraph("Media", styles["CenterSmall"]),
            Paragraph("Muy alto", styles["CenterSmall"]),
            Paragraph("Cifrado, mínimo privilegio, tokenización, pruebas de seguridad y respuesta a incidentes.", styles["BodySmall"]),
        ],
        [
            Paragraph("Sobreventa o información de disponibilidad desactualizada.", styles["BodySmall"]),
            Paragraph("Media", styles["CenterSmall"]),
            Paragraph("Alto", styles["CenterSmall"]),
            Paragraph("Confirmación en tiempo real, idempotencia, conciliación y mensajes claros al usuario.", styles["BodySmall"]),
        ],
        [
            Paragraph("Baja aceptación del planificador inteligente.", styles["BodySmall"]),
            Paragraph("Media", styles["CenterSmall"]),
            Paragraph("Medio", styles["CenterSmall"]),
            Paragraph("Prototipos tempranos, pruebas de usabilidad, métricas y ajuste por retroalimentación.", styles["BodySmall"]),
        ],
    ]
    risks = Table(risk_data, colWidths=[7.0 * cm, 2.2 * cm, 2.5 * cm, 14.0 * cm])
    risks.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#FFF0EA")),
                ("TEXTCOLOR", (0, 0), (-1, 0), RED),
                ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#DDCEC8")),
                ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#E8DDD8")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.extend(
        [
            Paragraph("Riesgos prioritarios", styles["PageTitle"]),
            risks,
            Spacer(1, 0.55 * cm),
            Table(
                [
                    [
                        Paragraph(
                            "<b>Conclusión.</b> El enfoque ágil e incremental permite que EasyTrip valide "
                            "tempranamente su propuesta de valor y reduzca el riesgo de construir funciones "
                            "que no respondan a las necesidades reales. Las seis fases permanecen presentes, "
                            "pero se recorren de manera repetida en cada incremento hasta obtener un producto "
                            "seguro, útil y sostenible.",
                            styles["Body"],
                        )
                    ]
                ],
                colWidths=[25.7 * cm],
                style=TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, -1), SKY),
                        ("BOX", (0, 0), (-1, -1), 0.8, TEAL),
                        ("LEFTPADDING", (0, 0), (-1, -1), 12),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
                        ("TOPPADDING", (0, 0), (-1, -1), 9),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
                    ]
                ),
            ),
        ]
    )
    return story


def main():
    doc = SimpleDocTemplate(
        OUTPUT,
        pagesize=PAGE_SIZE,
        rightMargin=1.4 * cm,
        leftMargin=1.4 * cm,
        topMargin=1.05 * cm,
        bottomMargin=1.05 * cm,
        title="Diseño del ciclo de vida del software — EasyTrip",
        author="[Nombre del estudiante]",
        subject="Actividad formativa Semana 1 — IGRV0132",
    )
    doc.build(build_story(), onFirstPage=header_footer, onLaterPages=header_footer)
    print(f"Documento generado: {OUTPUT}")


if __name__ == "__main__":
    main()
