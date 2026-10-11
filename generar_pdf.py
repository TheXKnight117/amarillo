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
            """Florence Nightingale fue una enfermera británica que cambió la manera en que se atendía a los pacientes durante el siglo XIX. Se hizo conocida principalmente por su trabajo en la Guerra de Crimea, donde dirigió a un grupo de enfermeras en hospitales militares que tenían muy malas condiciones de higiene. Elegimos a Nightingale porque no solo ayudó a los soldados heridos, sino que también logró cambiar prácticas que en ese tiempo se consideraban normales.""",
            """Su forma de dirigir se relaciona con el liderazgo transformacional. Según el apunte de la semana, este liderazgo busca inspirar a otras personas, promover ideas nuevas y producir un cambio profundo. Nightingale observó que muchos soldados morían por infecciones y por la falta de limpieza, más que por sus heridas. Por eso organizó la limpieza de las salas, mejoró la ventilación y se preocupó por la alimentación y el orden del hospital. Además, utilizó datos y gráficos para demostrar a las autoridades la importancia de aplicar medidas sanitarias.""",
            """En su caso se pueden reconocer características mencionadas por Bass (1985), como la motivación inspiradora y la estimulación intelectual. Nightingale motivó a otras enfermeras y las llevó a cuestionar procedimientos que no estaban funcionando. También tuvo una visión de largo plazo, porque después de la guerra impulsó la formación profesional de las enfermeras. El contexto fue importante: en medio de una guerra había poco tiempo y muchos recursos eran escasos. A pesar de eso, pudo organizar al equipo y demostrar que era posible trabajar de otra manera. Por estas razones, creemos que su liderazgo fue transformacional, ya que su influencia produjo mejoras que continuaron después de la crisis y ayudaron a que la enfermería fuera reconocida como una profesión fundamental.""",
        ],
    ),
    (
        "2. Nelson Mandela: liderazgo democrático",
        [
            """Nelson Mandela fue un líder político sudafricano que luchó contra el apartheid, un sistema de discriminación racial que limitaba los derechos de la población negra. Estuvo encarcelado durante veintisiete años y, después de recuperar su libertad, participó en las negociaciones que permitieron realizar elecciones democráticas. En 1994 llegó a ser el primer presidente negro de Sudáfrica.""",
            """Consideramos que Mandela representa el liderazgo democrático. El apunte explica que este estilo promueve la participación, el diálogo y la inclusión de las personas en la toma de decisiones. El líder mantiene la responsabilidad final, pero escucha diferentes opiniones y trata de construir acuerdos. Mandela demostró estas características cuando decidió negociar con el gobierno y con grupos que habían sido sus adversarios. En vez de buscar venganza, promovió la reconciliación y apoyó la formación de un gobierno de unidad nacional.""",
            """Un ejemplo concreto fue su apoyo a la Comisión de la Verdad y Reconciliación. Esta instancia permitió conocer violaciones a los derechos humanos ocurridas durante el apartheid y buscó que el país avanzara sin ocultar lo sucedido. Mandela entendió que Sudáfrica estaba profundamente dividida y que imponer una solución podía causar más violencia. Por eso prefirió el diálogo y la participación de distintos sectores.""",
            """Su liderazgo generó confianza y ayudó a que muchas personas se sintieran parte de la construcción de una nueva sociedad. También demuestra que ser democrático no significa evitar decisiones difíciles, sino considerar a los demás y buscar resultados que tengan legitimidad. Por el contexto histórico que enfrentó y por su capacidad para escuchar, negociar e incluir, creemos que este tipo de liderazgo es el que mejor representa sus acciones.""",
        ],
    ),
    (
        "3. Ernest Shackleton: liderazgo situacional",
        [
            """Ernest Shackleton fue un explorador británico-irlandés que dirigió la Expedición Imperial Transantártica de 1914. Su objetivo era atravesar la Antártida, pero el barco Endurance quedó atrapado en el hielo y finalmente fue destruido. Desde ese momento, el objetivo principal dejó de ser completar la expedición y pasó a ser mantener con vida a todos los integrantes del grupo.""",
            """Su caso se relaciona con el liderazgo situacional, porque tuvo que cambiar su manera de dirigir de acuerdo con cada problema. De acuerdo con Hersey y Blanchard (1982), no existe un único estilo que funcione en todas las situaciones. El líder puede dar instrucciones, apoyar, persuadir o delegar según las capacidades del equipo y las circunstancias. Shackleton fue muy directivo cuando ordenó abandonar el barco y controlar los alimentos. En esos momentos necesitaba decidir con rapidez, ya que una discusión larga podía aumentar el peligro.""",
            """Sin embargo, no actuó siempre de forma autoritaria. También escuchaba a sus compañeros, trataba de mantener su ánimo y organizaba actividades para evitar que el miedo y el aburrimiento dañaran la convivencia. Distribuyó las tareas considerando las capacidades de cada persona. Más adelante escogió a un grupo pequeño para realizar un peligroso viaje en bote hasta la isla Georgia del Sur en busca de ayuda, mientras otros quedaron a cargo del campamento.""",
            """La mayor prueba de su capacidad de adaptación fue renunciar a la meta original. Continuar la expedición habría puesto en riesgo a todos, por lo que cambió el plan y concentró sus esfuerzos en el rescate. Finalmente, todos los miembros de la tripulación sobrevivieron. Pensamos que Shackleton es un buen ejemplo de liderazgo situacional porque combinó dirección, apoyo y delegación en distintas etapas. Su experiencia demuestra que un líder debe observar el contexto y cambiar su estrategia cuando las condiciones lo requieren.""",
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
            """Los tres personajes muestran que no hay una sola manera de liderar. Florence Nightingale produjo cambios mediante nuevas ideas y una visión de futuro; Nelson Mandela usó el diálogo para unir a una sociedad dividida; y Ernest Shackleton fue modificando sus decisiones frente a una emergencia. Aunque actuaron en áreas y contextos diferentes, los tres fueron capaces de orientar a otras personas hacia un objetivo común. Esto confirma lo señalado en el apunte: la efectividad del liderazgo depende tanto de las características del líder como de las necesidades del equipo y de la situación.""",
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
    """Los tres personajes muestran que no hay una sola manera de liderar. Florence Nightingale produjo cambios mediante nuevas ideas y una visión de futuro; Nelson Mandela usó el diálogo para unir a una sociedad dividida; y Ernest Shackleton fue modificando sus decisiones frente a una emergencia. Aunque actuaron en áreas y contextos diferentes, los tres fueron capaces de orientar a otras personas hacia un objetivo común. Esto confirma lo señalado en el apunte: la efectividad del liderazgo depende tanto de las características del líder como de las necesidades del equipo y de la situación."""
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
