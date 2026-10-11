from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


pdfmetrics.registerFont(
    TTFont("Arial", "/usr/share/fonts/truetype/croscore/Arimo-Regular.ttf")
)
pdfmetrics.registerFont(
    TTFont("Arial-Bold", "/usr/share/fonts/truetype/croscore/Arimo-Bold.ttf")
)

OUTPUT = "Semana1_Apellidos_del_grupo.pdf"
WORD_OUTPUT = "Semana1_Apellidos_del_grupo.docx"

doc = SimpleDocTemplate(
    OUTPUT,
    pagesize=A4,
    rightMargin=2.5 * cm,
    leftMargin=2.5 * cm,
    topMargin=2.5 * cm,
    bottomMargin=2.5 * cm,
    title="Análisis de liderazgo en personajes históricos",
    author="Estudiantes FGLV304",
)

normal = ParagraphStyle(
    "Normal",
    fontName="Arial",
    fontSize=12,
    leading=14.4,
    alignment=TA_JUSTIFY,
    spaceAfter=8,
)
title = ParagraphStyle(
    "Title",
    parent=normal,
    fontName="Arial-Bold",
    fontSize=16,
    leading=19,
    alignment=TA_CENTER,
    spaceAfter=18,
)
subtitle = ParagraphStyle(
    "Subtitle",
    parent=normal,
    fontName="Arial-Bold",
    fontSize=12,
    leading=14.4,
    spaceBefore=8,
    spaceAfter=8,
)
cover = ParagraphStyle(
    "Cover",
    parent=normal,
    alignment=TA_CENTER,
    leading=20,
)

sections = [
    (
        "1. Florence Nightingale: liderazgo transformacional",
        [
            """Elegimos a Florence Nightingale porque su historia nos pareció un buen ejemplo de cómo una persona puede cambiar una realidad muy difícil. Fue una enfermera británica del siglo XIX y se hizo conocida por su trabajo durante la Guerra de Crimea. En los hospitales militares había suciedad, poco orden y muchos soldados enfermos. Ella estuvo a cargo de un grupo de enfermeras y tuvo que trabajar en esas malas condiciones.""",
            """Nosotros relacionamos su forma de dirigir con el liderazgo transformacional. En el apunte se explica que este tipo de líder inspira a otras personas, propone ideas nuevas y busca cambios que sean importantes. Nightingale se dio cuenta de que varios soldados no morían solamente por sus heridas, sino también por infecciones que podían evitarse. Por eso organizó la limpieza de las salas, mejoró la ventilación y se preocupó por la comida y el orden del hospital. Incluso usó datos y gráficos para mostrarles a las autoridades lo que estaba pasando.""",
            """También motivó a las demás enfermeras para que cambiaran procedimientos que no daban buenos resultados. Esto se relaciona con lo planteado por Bass (1985), especialmente con la motivación y la estimulación de nuevas ideas. Después de la guerra siguió trabajando para que las enfermeras tuvieran una mejor preparación. Creemos que lo más importante fue que no se conformó con atender el problema del momento, sino que trató de mejorar todo el sistema. Por eso la consideramos una líder transformacional. Su trabajo dejó cambios que siguieron existiendo muchos años después y ayudó a que la enfermería fuera tomada más en serio como profesión.""",
        ],
    ),
    (
        "2. Nelson Mandela: liderazgo democrático",
        [
            """Como segundo personaje elegimos a Nelson Mandela. Fue un líder político sudafricano que luchó contra el apartheid, que era un sistema de discriminación racial que quitaba derechos a la población negra. Estuvo preso durante veintisiete años. Cuando recuperó su libertad participó en las negociaciones que hicieron posibles las elecciones democráticas y en 1994 se convirtió en el primer presidente negro de Sudáfrica.""",
            """Para nosotros, Mandela representa el liderazgo democrático. El apunte señala que este liderazgo se basa en la participación, el diálogo y la inclusión de las personas al momento de decidir. El líder sigue teniendo una responsabilidad, pero escucha opiniones y trata de llegar a acuerdos. Mandela hizo esto al conversar con el gobierno y con personas que antes habían sido sus adversarios. Algo que nos llamó la atención es que no buscó vengarse después de salir de la cárcel. En cambio, habló de reconciliación y apoyó un gobierno en el que participaron distintos grupos.""",
            """Un ejemplo concreto fue su apoyo a la Comisión de la Verdad y Reconciliación. Esta instancia permitió conocer violaciones a los derechos humanos ocurridas durante el apartheid y buscó que el país avanzara sin ocultar lo sucedido. Mandela entendió que Sudáfrica estaba profundamente dividida y que imponer una solución podía causar más violencia. Por eso prefirió el diálogo y la participación de distintos sectores.""",
            """Pensamos que su forma de liderar ayudó a generar confianza y a que muchas personas se sintieran parte del cambio del país. Su caso también muestra que un líder democrático igual debe tomar decisiones difíciles. La diferencia es que considera a los demás y no decide todo por su cuenta. Por el momento histórico que vivió y por su capacidad para escuchar, negociar e incluir a diferentes grupos, creemos que este estilo es el que mejor representa a Mandela.""",
        ],
    ),
    (
        "3. Ernest Shackleton: liderazgo situacional",
        [
            """El tercer personaje es Ernest Shackleton, un explorador británico-irlandés que dirigió una expedición a la Antártida en 1914. Al principio su meta era atravesar el continente, pero su barco, el Endurance, quedó atrapado y fue destruido por el hielo. En ese momento el plan tuvo que cambiar por completo. Ya no importaba terminar el viaje, sino lograr que todos regresaran vivos.""",
            """Escogimos este caso para representar el liderazgo situacional, porque Shackleton tuvo que cambiar su forma de dirigir varias veces. Hersey y Blanchard (1982) explican que un mismo estilo no sirve para todas las situaciones. A veces el líder debe dar instrucciones claras y otras veces debe apoyar o delegar. Por ejemplo, Shackleton actuó de manera muy directa cuando ordenó abandonar el barco y controlar los alimentos. Era una emergencia y demorarse demasiado en decidir podía ponerlos en mayor peligro.""",
            """Sin embargo, no actuó siempre de forma autoritaria. También escuchaba a sus compañeros, trataba de mantener su ánimo y organizaba actividades para evitar que el miedo y el aburrimiento dañaran la convivencia. Distribuyó las tareas considerando las capacidades de cada persona. Más adelante escogió a un grupo pequeño para realizar un peligroso viaje en bote hasta la isla Georgia del Sur en busca de ayuda, mientras otros quedaron a cargo del campamento.""",
            """Lo que más demuestra su capacidad de adaptación fue que aceptó abandonar su objetivo inicial. Seguir con la expedición habría sido demasiado arriesgado, así que se concentró en el rescate. Finalmente, todos sobrevivieron. Creemos que Shackleton representa bien el liderazgo situacional porque no reaccionó siempre de la misma manera: fue firme cuando era necesario, apoyó al grupo y también entregó responsabilidades. Su historia deja claro que cambiar un plan no siempre significa fracasar, especialmente cuando la situación y el bienestar del equipo exigen otra decisión.""",
        ],
    ),
]

story = [
    Spacer(1, 3.2 * cm),
    Paragraph("ANÁLISIS DE LIDERAZGO EN PERSONAJES HISTÓRICOS", title),
    Spacer(1, 1.2 * cm),
    Paragraph("<b>Asignatura:</b> Comunicación profesional efectiva", cover),
    Paragraph("<b>Semana:</b> 2", cover),
    Paragraph("<b>Integrantes:</b> ________________________________", cover),
    Paragraph("________________________________", cover),
    Paragraph("<b>Docente:</b> ________________________________", cover),
    Paragraph("<b>Fecha:</b> ________________________________", cover),
    PageBreak(),
    Paragraph("Análisis de liderazgo en personajes históricos", title),
]

for heading, paragraphs in sections:
    story.append(Paragraph(heading, subtitle))
    for text in paragraphs:
        story.append(Paragraph(text, normal))

story.append(
    KeepTogether(
    [
        Paragraph("Conclusión", subtitle),
        Paragraph(
            """Después de revisar estos tres casos, entendimos que no existe una sola manera correcta de liderar. Florence Nightingale cambió la enfermería con nuevas ideas, Nelson Mandela prefirió el diálogo para unir a un país dividido y Ernest Shackleton fue ajustando sus decisiones durante una emergencia. Sus situaciones eran muy distintas, pero los tres lograron guiar a otras personas hacia una meta. En nuestra opinión, un buen líder no solo necesita tener autoridad. También debe saber escuchar, organizar al equipo y darse cuenta de cuándo es necesario cambiar su forma de actuar.""",
            normal,
        ),
        Paragraph("Referencias", subtitle),
        Paragraph(
            """Bass, B. M. (1985). <i>Leadership and performance beyond expectations</i>. Free Press.""",
            normal,
        ),
        Paragraph(
            """Hersey, P. y Blanchard, K. H. (1982). <i>Management of Organizational Behavior: Utilizing Human Resources</i>. Prentice-Hall.""",
            normal,
        ),
        Paragraph(
            """Institución educativa. (2026). <i>Comunicación profesional efectiva: Apunte semana 2</i>.""",
            normal,
        ),
    ]
    )
)

doc.build(story)
print(OUTPUT)

word_doc = Document()
word_section = word_doc.sections[0]
word_section.top_margin = Cm(2.5)
word_section.bottom_margin = Cm(2.5)
word_section.left_margin = Cm(2.5)
word_section.right_margin = Cm(2.5)

word_normal = word_doc.styles["Normal"]
word_normal.font.name = "Arial"
word_normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
word_normal.font.size = Pt(12)
word_normal.paragraph_format.line_spacing = 1
word_normal.paragraph_format.space_after = Pt(6)


def add_centered(text="", bold=False, size=12):
    paragraph = word_doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.name = "Arial"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    run.font.size = Pt(size)
    return paragraph


def add_heading(text):
    paragraph = word_doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(6)
    paragraph.paragraph_format.space_after = Pt(6)
    run = paragraph.add_run(text)
    run.bold = True
    run.font.name = "Arial"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    run.font.size = Pt(12)


def add_body(text):
    paragraph = word_doc.add_paragraph(text)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.line_spacing = 1
    paragraph.paragraph_format.space_after = Pt(6)


for _ in range(5):
    word_doc.add_paragraph()
add_centered("ANÁLISIS DE LIDERAZGO EN PERSONAJES HISTÓRICOS", True, 16)
word_doc.add_paragraph()
add_centered("Asignatura: Comunicación profesional efectiva")
add_centered("Semana: 2")
add_centered("Integrantes: ________________________________")
add_centered("________________________________")
add_centered("Docente: ________________________________")
add_centered("Fecha: ________________________________")
word_doc.add_page_break()

add_centered("Análisis de liderazgo en personajes históricos", True, 16)
for heading, paragraphs in sections:
    add_heading(heading)
    for text in paragraphs:
        add_body(text)

add_heading("Conclusión")
add_body(
    """Después de revisar estos tres casos, entendimos que no existe una sola manera correcta de liderar. Florence Nightingale cambió la enfermería con nuevas ideas, Nelson Mandela prefirió el diálogo para unir a un país dividido y Ernest Shackleton fue ajustando sus decisiones durante una emergencia. Sus situaciones eran muy distintas, pero los tres lograron guiar a otras personas hacia una meta. En nuestra opinión, un buen líder no solo necesita tener autoridad. También debe saber escuchar, organizar al equipo y darse cuenta de cuándo es necesario cambiar su forma de actuar."""
)
add_heading("Referencias")
add_body("Bass, B. M. (1985). Leadership and performance beyond expectations. Free Press.")
add_body(
    "Hersey, P. y Blanchard, K. H. (1982). Management of Organizational Behavior: Utilizing Human Resources. Prentice-Hall."
)
add_body(
    "Institución educativa. (2026). Comunicación profesional efectiva: Apunte semana 2."
)

word_doc.save(WORD_OUTPUT)
print(WORD_OUTPUT)
