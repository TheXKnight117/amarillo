from pathlib import Path
from textwrap import wrap

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    HRFlowable,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)
from pypdf import PdfReader, PdfWriter


OUT = Path("entrega")
OUT.mkdir(exist_ok=True)

REPORT_PATH = OUT / "Informe_conflicto_laboral_fabrica_dulces.pdf"
INFOGRAPHIC_PATH = OUT / "Infografia_plan_dulce_acuerdo.pdf"
COMBINED_PATH = OUT / "Entrega_completa_informe_e_infografia.pdf"

FONT_DIR = Path("/usr/share/fonts/truetype/croscore")
pdfmetrics.registerFont(TTFont("Arial", FONT_DIR / "Arimo-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Arial-Bold", FONT_DIR / "Arimo-Bold.ttf"))
pdfmetrics.registerFont(TTFont("Arial-Italic", FONT_DIR / "Arimo-Italic.ttf"))
pdfmetrics.registerFontFamily(
    "Arial", normal="Arial", bold="Arial-Bold", italic="Arial-Italic"
)

NAVY = colors.HexColor("#17324D")
TEAL = colors.HexColor("#1F7A78")
GOLD = colors.HexColor("#E5A93D")
CREAM = colors.HexColor("#FFF8EA")
PALE_TEAL = colors.HexColor("#E7F3F1")
PALE_BLUE = colors.HexColor("#EDF3F8")
INK = colors.HexColor("#24313B")
MUTED = colors.HexColor("#5E6A73")
LINE = colors.HexColor("#CBD4D9")
WHITE = colors.white


def report_header_footer(c: canvas.Canvas, doc):
    width, height = A4
    c.saveState()
    c.setStrokeColor(LINE)
    c.setLineWidth(0.6)
    c.line(2 * cm, height - 1.45 * cm, width - 2 * cm, height - 1.45 * cm)
    c.setFont("Arial", 8.5)
    c.setFillColor(MUTED)
    c.drawString(2 * cm, 1.15 * cm, "Comunicación profesional efectiva · Evaluación sumativa")
    c.drawRightString(width - 2 * cm, 1.15 * cm, f"Página {doc.page}")
    c.restoreState()


def build_report():
    doc = BaseDocTemplate(
        str(REPORT_PATH),
        pagesize=A4,
        rightMargin=2.2 * cm,
        leftMargin=2.2 * cm,
        topMargin=2.1 * cm,
        bottomMargin=1.8 * cm,
        title="Análisis de un caso sobre conflictos laborales",
        author="Estudiante",
        subject="Evaluación sumativa, semana 3",
    )
    frame = Frame(
        doc.leftMargin,
        doc.bottomMargin,
        doc.width,
        doc.height,
        leftPadding=0,
        rightPadding=0,
        topPadding=0,
        bottomPadding=0,
    )
    doc.addPageTemplates(
        [PageTemplate(id="content", frames=[frame], onPage=report_header_footer)]
    )

    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "Title",
        parent=styles["Title"],
        fontName="Arial-Bold",
        fontSize=22,
        leading=27,
        textColor=NAVY,
        alignment=TA_CENTER,
        spaceAfter=14,
    )
    subtitle = ParagraphStyle(
        "Subtitle",
        fontName="Arial",
        fontSize=12,
        leading=16,
        textColor=MUTED,
        alignment=TA_CENTER,
    )
    heading = ParagraphStyle(
        "Heading",
        fontName="Arial-Bold",
        fontSize=14,
        leading=18,
        textColor=NAVY,
        spaceBefore=10,
        spaceAfter=7,
        keepWithNext=True,
    )
    body = ParagraphStyle(
        "Body",
        fontName="Arial",
        fontSize=12,
        leading=14.4,
        textColor=INK,
        alignment=TA_JUSTIFY,
        spaceAfter=8,
    )
    bullet = ParagraphStyle(
        "Bullet",
        parent=body,
        leftIndent=14,
        firstLineIndent=-10,
        bulletIndent=2,
        spaceAfter=6,
    )
    small = ParagraphStyle(
        "Small",
        fontName="Arial",
        fontSize=9.5,
        leading=12,
        textColor=MUTED,
        alignment=TA_LEFT,
    )
    callout = ParagraphStyle(
        "Callout",
        fontName="Arial",
        fontSize=11.5,
        leading=15,
        textColor=INK,
        alignment=TA_LEFT,
    )

    story = [
        Spacer(1, 2.2 * cm),
        Paragraph("ANÁLISIS DE UN CASO<br/>SOBRE CONFLICTOS LABORALES", title),
        HRFlowable(
            width="38%",
            thickness=3,
            color=GOLD,
            spaceBefore=4,
            spaceAfter=18,
            hAlign="CENTER",
        ),
        Paragraph("Fábrica de dulces · Evaluación sumativa, semana 3", subtitle),
        Spacer(1, 1.7 * cm),
    ]

    student_data = [
        [Paragraph("<b>Nombre:</b>", body), Paragraph("________________________________________", body)],
        [Paragraph("<b>Carrera:</b>", body), Paragraph("________________________________________", body)],
        [Paragraph("<b>Docente:</b>", body), Paragraph("________________________________________", body)],
        [Paragraph("<b>Fecha:</b>", body), Paragraph("________________________________________", body)],
    ]
    data_table = Table(student_data, colWidths=[3.1 * cm, 10.4 * cm])
    data_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), PALE_BLUE),
                ("BOX", (0, 0), (-1, -1), 0.7, LINE),
                ("INNERGRID", (0, 0), (-1, -1), 0.4, LINE),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    story.extend(
        [
            data_table,
            Spacer(1, 1.4 * cm),
            Paragraph(
                "<b>Propósito del trabajo.</b> Analizar los problemas de comunicación, "
                "trabajo en equipo y liderazgo presentes en el caso, y proponer medidas "
                "realistas para recuperar la coordinación, la motivación y la confianza.",
                callout,
            ),
            PageBreak(),
            Paragraph("Introducción", heading),
            Paragraph(
                "El caso muestra una fábrica que no tiene un solo problema, sino varios "
                "que se alimentan entre sí. La información incompleta produce rumores; "
                "la falta de claridad en los roles genera errores; y la ausencia de un "
                "liderazgo cercano deja a las personas sin orientación. Como resultado, "
                "disminuyen la productividad y la calidad, empeora el ambiente laboral y "
                "también se debilita la relación con los clientes. A continuación, analizo "
                "cada aspecto y planteo una propuesta de mejora que se puede aplicar de "
                "manera gradual.",
                body,
            ),
            Paragraph(
                "1. Desafíos que afectan el funcionamiento de la empresa", heading
            ),
            Paragraph(
                "El primer desafío es la <b>comunicación deficiente</b>. La información "
                "oficial llega tarde o simplemente no llega, por lo que los rumores ocupan "
                "ese espacio. Esto aumenta la incertidumbre y hace que los trabajadores "
                "desconfíen tanto de sus jefaturas como de sus compañeros.",
                body,
            ),
            Paragraph(
                "El segundo desafío es la <b>falta de definición de funciones y "
                "responsabilidades</b>. En el caso, dos personas creen que la otra debía "
                "empacar los bombones. Esta situación demuestra que no existe una "
                "distribución clara de las tareas. Como consecuencia, algunas labores se "
                "duplican, otras quedan sin realizar y aparecen errores que afectan la "
                "calidad del producto.",
                body,
            ),
            Paragraph(
                "El tercer desafío es un <b>liderazgo débil y reactivo</b>. Los jefes se "
                "mantienen aislados, solo entregan órdenes y actúan cuando el problema ya "
                "ocurrió. No orientan, no motivan y tampoco abren espacios para conversar. "
                "A esto se suma un cuarto desafío: el deterioro del clima laboral, que se "
                "refleja en desmotivación, tensión y posible salida de trabajadores "
                "valiosos. Todos estos factores repercuten finalmente en la productividad, "
                "la calidad y la confianza de los clientes.",
                body,
            ),
            Paragraph("2. Dificultades de comunicación y posibles mejoras", heading),
            Paragraph(
                "Las principales dificultades son la comunicación en una sola dirección, "
                "la escasa transparencia y la falta de escucha. Las jefaturas comunican "
                "órdenes, pero no explican sus razones ni permiten que el equipo haga "
                "preguntas. Además, al no existir un canal oficial y frecuente, las "
                "personas completan la información con suposiciones. El rumor sobre el "
                "cierre de la línea de caramelos es un ejemplo claro. También falta "
                "retroalimentación: los problemas se conocen cuando ya provocaron atrasos "
                "o errores.",
                body,
            ),
            Paragraph(
                "Para lograr una comunicación más efectiva, propondría una reunión breve "
                "al inicio de cada turno, de no más de diez minutos, donde se informen las "
                "metas del día, cambios, responsables y riesgos. La misma información debe "
                "quedar escrita en un panel visible o canal interno para evitar versiones "
                "diferentes. También es importante aplicar escucha activa: dejar hablar, "
                "hacer preguntas para aclarar y resumir lo entendido antes de responder. "
                "Cuando se formule un juicio, este debe apoyarse en hechos observables. Por "
                "ejemplo, en vez de decir «el equipo es desordenado», corresponde señalar "
                "qué tarea quedó pendiente, cuándo ocurrió y qué efecto tuvo. Finalmente, "
                "un espacio semanal de retroalimentación permitiría resolver dudas antes "
                "de que se conviertan en conflictos.",
                body,
            ),
            PageBreak(),
            Paragraph("3. Análisis del trabajo en equipo", heading),
            Paragraph(
                "En la fábrica existen personas que trabajan en el mismo lugar, pero no "
                "se observa un equipo realmente coordinado. Un equipo necesita una meta "
                "común, funciones conocidas, comunicación constante y confianza para pedir "
                "ayuda. En este caso, cada trabajador parece concentrarse en resolver lo "
                "inmediato sin tener una visión compartida del proceso completo. Por eso, "
                "un error en el empaque no se entiende como un problema de todos, aunque "
                "termine afectando la producción y al cliente.",
                body,
            ),
            Paragraph(
                "La falta de claridad genera duplicación de tareas y vacíos de "
                "responsabilidad. Al mismo tiempo, los rumores y la tensión reducen la "
                "confianza, por lo que es probable que las personas eviten preguntar o "
                "reconocer un error por temor a ser culpadas. Esto impide la colaboración "
                "y hace que los conflictos se vuelvan personales en vez de centrarse en "
                "el proceso de trabajo.",
                body,
            ),
            Paragraph(
                "Para fortalecer el equipo, primero se debe acordar un objetivo concreto, "
                "por ejemplo: disminuir los errores de empaque y cumplir la producción "
                "diaria sin sacrificar la calidad. Después, conviene elaborar una matriz "
                "simple de responsabilidades para que cada tarea tenga una persona "
                "encargada, alguien que aprueba y compañeros que deben ser informados. "
                "También servirían reuniones cortas entre producción, empaque y control de "
                "calidad, ya que estas áreas dependen unas de otras. El avance debe medirse "
                "con indicadores compartidos, no para buscar culpables, sino para detectar "
                "qué parte del proceso necesita apoyo. Reconocer las mejoras del grupo "
                "ayudaría a recuperar el compromiso.",
                body,
            ),
            Paragraph("4. Tipo de liderazgo y propuesta de mejora", heading),
            Paragraph(
                "El liderazgo presente se puede describir como <b>autoritario y "
                "reactivo</b>. Es autoritario porque las jefaturas se limitan a dar "
                "instrucciones sin abrir diálogo, y es reactivo porque solo «apagan "
                "incendios» cuando los conflictos ya están instalados. También tiene "
                "rasgos de liderazgo ausente, debido a que los trabajadores no saben a "
                "quién acudir para recibir orientación. Este estilo no es efectivo en la "
                "situación descrita: puede lograr obediencia momentánea, pero no genera "
                "confianza, compromiso ni aprendizaje.",
                body,
            ),
            Paragraph(
                "La situación mejoraría con un <b>liderazgo participativo y "
                "transformacional</b>. Participativo, porque las decisiones relacionadas "
                "con el trabajo diario deben considerar la experiencia de quienes conocen "
                "la producción. Transformacional, porque se necesita recuperar un sentido "
                "común: fabricar productos de calidad, cuidar al equipo y volver a ganar la "
                "confianza de los clientes. Esto no significa que todas las decisiones se "
                "voten. En una emergencia, la jefatura debe actuar con rapidez; sin embargo, "
                "en la gestión normal debe explicar las decisiones, escuchar propuestas, "
                "dar retroalimentación y delegar con claridad.",
                body,
            ),
            Paragraph(
                "En la práctica, cada líder debería recorrer periódicamente la planta, "
                "realizar conversaciones breves con su equipo y comprometerse con acciones "
                "que luego pueda comprobar. Además, tendría que reconocer los buenos "
                "resultados, abordar los errores como oportunidades de mejora y mediar los "
                "conflictos usando hechos, escucha activa y acuerdos concretos. Así, el "
                "liderazgo dejaría de ser una figura distante y pasaría a funcionar como "
                "un apoyo para coordinar el trabajo.",
                body,
            ),
            Paragraph("Conclusión", heading),
            Paragraph(
                "La fábrica todavía puede revertir su situación si enfrenta los problemas "
                "de manera conectada. No basta con enviar más instrucciones: se requiere "
                "información clara, responsabilidades visibles y líderes presentes. La "
                "propuesta «Dulce Acuerdo», resumida en la infografía, comienza con medidas "
                "sencillas durante treinta días y usa indicadores para revisar si realmente "
                "hay avances. Si las personas saben qué hacer, pueden expresar sus dudas y "
                "ven que sus aportes son considerados, será más posible recuperar el buen "
                "clima, mejorar la calidad y fortalecer la relación con los clientes.",
                body,
            ),
            Spacer(1, 0.3 * cm),
            Paragraph("Fuentes consultadas", heading),
            Paragraph(
                "Material de evaluación sumativa, semana 3: <i>Análisis de un caso sobre "
                "conflictos laborales</i>.",
                small,
            ),
            Paragraph(
                "Vergara González, K. (s. f.). <i>Comunicación profesional efectiva: "
                "Apunte semana 3</i>. En especial, se aplicaron las etapas de lectura, "
                "inferencia y poslectura para organizar el análisis del caso.",
                small,
            ),
        ]
    )
    doc.build(story)


def rounded_box(c, x, y, w, h, fill, title_text, body_lines, accent=None):
    c.setFillColor(fill)
    c.setStrokeColor(LINE)
    c.setLineWidth(0.8)
    c.roundRect(x, y, w, h, 9, fill=1, stroke=1)
    if accent:
        c.setFillColor(accent)
        c.roundRect(x, y + h - 8, w, 8, 8, fill=1, stroke=0)
        c.rect(x, y + h - 8, w, 4, fill=1, stroke=0)
    c.setFillColor(NAVY)
    c.setFont("Arial-Bold", 13)
    c.drawString(x + 13, y + h - 25, title_text)
    c.setFont("Arial", 9.2)
    c.setFillColor(INK)
    line_y = y + h - 43
    for line in body_lines:
        if line.startswith("•"):
            c.setFillColor(accent or TEAL)
            c.circle(x + 16, line_y + 2, 2.2, fill=1, stroke=0)
            c.setFillColor(INK)
            text = line[1:].strip()
            for piece in wrap(text, width=43):
                c.drawString(x + 24, line_y, piece)
                line_y -= 11
        else:
            for piece in wrap(line, width=47):
                c.drawString(x + 14, line_y, piece)
                line_y -= 11
        line_y -= 3


def metric(c, x, y, value, label):
    c.setFillColor(WHITE)
    c.setStrokeColor(colors.Color(1, 1, 1, alpha=0.35))
    c.roundRect(x, y, 118, 42, 7, fill=1, stroke=1)
    c.setFillColor(NAVY)
    c.setFont("Arial-Bold", 14)
    c.drawCentredString(x + 59, y + 24, value)
    c.setFont("Arial", 7.8)
    c.setFillColor(MUTED)
    c.drawCentredString(x + 59, y + 10, label)


def build_infographic():
    page = landscape(A4)
    width, height = page
    c = canvas.Canvas(
        str(INFOGRAPHIC_PATH),
        pagesize=page,
        title="Infografía: plan Dulce Acuerdo",
        author="Estudiante",
    )
    c.setFillColor(colors.HexColor("#FAFBFC"))
    c.rect(0, 0, width, height, fill=1, stroke=0)

    c.setFillColor(NAVY)
    c.rect(0, height - 96, width, 96, fill=1, stroke=0)
    c.setFillColor(GOLD)
    c.rect(0, height - 101, width, 5, fill=1, stroke=0)
    c.setFillColor(WHITE)
    c.setFont("Arial-Bold", 24)
    c.drawString(34, height - 42, "DULCE ACUERDO")
    c.setFont("Arial", 12.5)
    c.drawString(
        34,
        height - 65,
        "Plan de 30 días para recuperar la comunicación, la coordinación y la confianza",
    )
    c.setFont("Arial-Bold", 9.5)
    c.setFillColor(GOLD)
    c.drawRightString(width - 34, height - 43, "FÁBRICA DE DULCES")
    c.setFont("Arial", 8.5)
    c.setFillColor(colors.HexColor("#DDE7EF"))
    c.drawRightString(width - 34, height - 62, "Propuesta individual · Semana 3")

    margin = 34
    gap = 13
    top_y = height - 238
    box_h = 121
    box_w = (width - 2 * margin - 2 * gap) / 3

    rounded_box(
        c,
        margin,
        top_y,
        box_w,
        box_h,
        PALE_BLUE,
        "1. COMUNICACIÓN CLARA",
        [
            "• Reunión de 10 minutos al iniciar cada turno.",
            "• Panel único con metas, cambios y responsables.",
            "• Escucha activa: preguntar, resumir y acordar.",
            "Responsable: jefatura de turno.",
        ],
        TEAL,
    )
    rounded_box(
        c,
        margin + box_w + gap,
        top_y,
        box_w,
        box_h,
        CREAM,
        "2. EQUIPO COORDINADO",
        [
            "• Matriz simple: quién hace, aprueba e informa.",
            "• Meta compartida entre producción, empaque y calidad.",
            "• Revisión semanal sin buscar culpables.",
            "Responsable: líderes y representantes de área.",
        ],
        GOLD,
    )
    rounded_box(
        c,
        margin + 2 * (box_w + gap),
        top_y,
        box_w,
        box_h,
        PALE_TEAL,
        "3. LIDERAZGO CERCANO",
        [
            "• Recorrido diario por la planta y canal de dudas.",
            "• Decisiones explicadas y tareas delegadas con claridad.",
            "• Reconocer avances y cumplir los acuerdos.",
            "Responsable: jefaturas y gerencia.",
        ],
        TEAL,
    )

    flow_y = 248
    c.setFillColor(NAVY)
    c.setFont("Arial-Bold", 12)
    c.drawString(margin, flow_y + 47, "RUTA DE IMPLEMENTACIÓN")
    steps = [
        ("DÍAS 1–3", "Escuchar al equipo\ny confirmar problemas"),
        ("DÍAS 4–7", "Definir funciones,\ncanales y metas"),
        ("DÍAS 8–21", "Aplicar reuniones\ny acompañamiento"),
        ("DÍAS 22–30", "Medir resultados\ny ajustar el plan"),
    ]
    total_w = width - 2 * margin
    step_w = (total_w - 3 * 24) / 4
    for i, (period, description) in enumerate(steps):
        x = margin + i * (step_w + 24)
        c.setFillColor(WHITE)
        c.setStrokeColor(LINE)
        c.roundRect(x, flow_y - 16, step_w, 48, 7, fill=1, stroke=1)
        c.setFillColor(TEAL if i < 3 else GOLD)
        c.setFont("Arial-Bold", 9)
        c.drawString(x + 10, flow_y + 15, period)
        c.setFillColor(INK)
        c.setFont("Arial", 8.4)
        lines = description.split("\n")
        c.drawString(x + 10, flow_y + 2, lines[0])
        c.drawString(x + 10, flow_y - 9, lines[1])
        if i < len(steps) - 1:
            arrow_x = x + step_w + 6
            c.setStrokeColor(MUTED)
            c.setLineWidth(1.4)
            c.line(arrow_x, flow_y + 8, arrow_x + 11, flow_y + 8)
            c.line(arrow_x + 11, flow_y + 8, arrow_x + 7, flow_y + 12)
            c.line(arrow_x + 11, flow_y + 8, arrow_x + 7, flow_y + 4)

    c.setFillColor(NAVY)
    c.roundRect(margin, 50, width - 2 * margin, 142, 10, fill=1, stroke=0)
    c.setFillColor(WHITE)
    c.setFont("Arial-Bold", 13)
    c.drawString(margin + 18, 168, "¿CÓMO SABREMOS SI FUNCIONÓ?")
    metric_y = 108
    metric_gap = 13
    metrics = [
        ("−30%", "errores de empaque"),
        ("100%", "turnos con reunión breve"),
        ("−50%", "tareas sin responsable"),
        ("≥ 80%", "acuerdos cumplidos"),
        ("+20%", "percepción de claridad"),
    ]
    available = width - 2 * margin - 36
    metric_w = (available - 4 * metric_gap) / 5
    for i, (value, label) in enumerate(metrics):
        x = margin + 18 + i * (metric_w + metric_gap)
        c.setFillColor(WHITE)
        c.roundRect(x, metric_y, metric_w, 45, 7, fill=1, stroke=0)
        c.setFillColor(GOLD if i in (0, 4) else TEAL)
        c.setFont("Arial-Bold", 14)
        c.drawCentredString(x + metric_w / 2, metric_y + 26, value)
        c.setFillColor(MUTED)
        c.setFont("Arial", 7.8)
        c.drawCentredString(x + metric_w / 2, metric_y + 11, label)

    c.setFillColor(colors.HexColor("#DDE7EF"))
    c.setFont("Arial-Italic", 8.5)
    c.drawString(
        margin + 18,
        72,
        "Idea central: informar con claridad, coordinar responsabilidades y liderar desde la presencia.",
    )
    c.setFillColor(MUTED)
    c.setFont("Arial", 7.5)
    c.drawRightString(
        width - margin,
        24,
        "Elaboración propia a partir del caso de la evaluación sumativa.",
    )
    c.showPage()
    c.save()


def combine_pdfs():
    writer = PdfWriter()
    for path in (REPORT_PATH, INFOGRAPHIC_PATH):
        reader = PdfReader(str(path))
        for page in reader.pages:
            writer.add_page(page)
    writer.add_metadata(
        {
            "/Title": "Entrega completa: informe e infografía",
            "/Author": "Estudiante",
            "/Subject": "Análisis de un caso sobre conflictos laborales",
        }
    )
    with COMBINED_PATH.open("wb") as file:
        writer.write(file)


if __name__ == "__main__":
    build_report()
    build_infographic()
    combine_pdfs()
    for path in (REPORT_PATH, INFOGRAPHIC_PATH, COMBINED_PATH):
        reader = PdfReader(str(path))
        print(f"{path}: {len(reader.pages)} página(s), {path.stat().st_size} bytes")
