from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import KeepTogether, Paragraph, SimpleDocTemplate, Spacer


OUTPUT_DIR = Path(__file__).parent
DOCX_PATH = OUTPUT_DIR / "Semana1_Apellido.docx"
PDF_PATH = OUTPUT_DIR / "Semana1_Apellido.pdf"
RTF_PATH = OUTPUT_DIR / "Semana1_Apellido.rtf"

TITLE = "Actividad formativa: El poder de la comunicación"
SUBTITLE = "Semana 1 – Análisis y reflexión"
STUDENT = "Nombre y apellido: ______________________________________________"

QUESTIONS_AND_ANSWERS = [
    (
        "1. Pensando en el concepto de comunicación efectiva, ¿qué importancia "
        "tiene el rol del emisor y del receptor en el proceso de comunicación?",
        (
            "Yo creo que el emisor y el receptor tienen la misma importancia, porque no "
            "basta con mandar un mensaje: también es necesario que la otra persona lo "
            "entienda bien. Como vimos en el contenido de la semana, el emisor codifica la "
            "idea y debe expresarla con claridad y precisión. También debe decir de dónde "
            "viene la información y pensar si el canal que está usando es el correcto. En el "
            "cortometraje esto no ocurre, ya que el aviso de bomba aparece en las redes "
            "sociales y no queda claro quién lo escribió ni si es verdadero. A pesar de "
            "eso, casi todos los estudiantes confían en el mensaje y deciden no tomar el "
            "metro para ir a clases. También se reconocen los demás elementos de la "
            "comunicación: el mensaje es el supuesto aviso de bomba, el canal son las redes "
            "sociales y el contexto es un grupo de estudiantes que debe trasladarse por "
            "Madrid para asistir a clases. El receptor también tiene una responsabilidad, porque "
            "es quien decodifica el mensaje. Debe leer con atención, hacer preguntas y "
            "comprobar la información antes de "
            "creerla o compartirla. En este caso, los jóvenes reaccionan por miedo y no "
            "buscan otra fuente que confirme la noticia. Solo dos estudiantes terminan "
            "asistiendo y realizan el examen sorpresa del profesor. Para mí, una buena "
            "comunicación se produce cuando el emisor entrega información clara y el "
            "receptor participa de manera activa. La retroalimentación permite responder "
            "y confirmar si el mensaje se entendió como esperaba el emisor. "
            "Si una de las dos partes falla, el mensaje puede provocar confusión y llevar "
            "a decisiones equivocadas."
        ),
    ),
    (
        "2. ¿Cómo influye el canal elegido en la percepción del mensaje?",
        (
            "El canal influye mucho, porque una misma información puede generar reacciones "
            "distintas según dónde y cómo se comunique. En el video, el aviso se comparte "
            "por redes sociales, un medio que los estudiantes usan todos los días y al que "
            "le prestan atención rápidamente. Como el mensaje se difunde entre compañeros, "
            "varios piensan que debe ser cierto, aunque nadie muestre una fuente oficial. "
            "También pasa que, cuando una publicación se comparte muchas veces, parece más "
            "confiable de lo que realmente es. Las redes permiten informar en pocos segundos, "
            "pero no siempre entregan contexto ni dan la posibilidad de preguntarle al autor "
            "qué quiso decir. En una conversación presencial se podrían observar los gestos, "
            "el tono de voz y la seguridad de la persona, pero aquí esos elementos no están. "
            "Además, hablar de una bomba provoca miedo y hace que los estudiantes reaccionen "
            "antes de pensar con calma. Por eso el canal ayuda a que el aviso llegue rápido, "
            "pero también facilita que el rumor se agrande. Si la misma información hubiera "
            "aparecido en un comunicado del Metro, de la policía o del centro educativo, "
            "los estudiantes habrían tenido más razones para confiar y sabrían qué hacer. "
            "Esto demuestra la importancia de adaptar el canal al contexto, a la audiencia "
            "y al tipo de mensaje que se necesita comunicar."
        ),
    ),
    (
        "3. Considerando el canal que se utiliza en la situación comunicativa, "
        "¿qué barreras pueden identificar? ¿Es el más adecuado para el mensaje y la audiencia?",
        (
            "En esta situación veo varios quiebres o barreras de comunicación. La primera "
            "es la falta de claridad, porque no "
            "se conoce al verdadero emisor, por lo que nadie sabe si el aviso viene de una "
            "autoridad o de otro estudiante. Otra barrera es que el mensaje parece tener "
            "pocos detalles: no se explica bien dónde está la bomba, a qué hora ocurrió el "
            "aviso ni quién confirmó la información. También existe una sobrecarga de "
            "información en las redes, porque allí se mezclan noticias, opiniones, bromas "
            "y rumores. El miedo es un problema emocional que también actúa como barrera. "
            "Al tratarse de una posible bomba, los jóvenes "
            "se preocupan y toman una decisión rápida sin revisar primero si la información "
            "es real. Esto muestra una falta de escucha activa y de clarificación, ya que "
            "nadie hace preguntas al supuesto emisor. Además, como casi todos piensan igual, "
            "se produce una especie de "
            "presión del grupo que hace más difícil cuestionar el mensaje. Considero que "
            "las redes sí sirven para avisar rápidamente a estudiantes, pero no son "
            "suficientes para una situación tan seria. Lo adecuado sería revisar primero "
            "los canales oficiales del Metro, la policía o la institución educativa. Las "
            "redes podrían usarse como apoyo, pero el mensaje tendría que incluir la fuente, "
            "la hora y un enlace oficial para comprobarlo."
        ),
    ),
    (
        "4. Basándose en lo observado, ¿qué podemos concluir sobre la comunicación "
        "a partir de este cortometraje? ¿Es efectiva? ¿Qué aspectos se podrían mejorar?",
        (
            "Después de ver el cortometraje, mi conclusión es que un mensaje puede influir "
            "mucho en las personas aunque no esté confirmado. El aviso funciona para "
            "convencer a la mayoría, porque los estudiantes deciden no usar el metro y "
            "tampoco van a clases. Sin embargo, yo no diría que la comunicación es realmente "
            "efectiva. Logra una reacción, pero la información no es clara, no se conoce su "
            "origen y los receptores no comprueban si es verdadera. El video muestra que las "
            "redes sociales tienen un gran poder para persuadir, especialmente cuando el "
            "mensaje provoca miedo o preocupación. También muestra que compartir algo muchas "
            "veces no lo convierte en verdad. Aunque el aviso se presenta como una afirmación, "
            "los estudiantes lo aceptan como un hecho sin tener pruebas. Podrían haber buscado un aviso "
            "oficial, preguntado al profesor o revisado la información del Metro de Madrid "
            "antes de tomar una decisión. El mensaje también debería haber indicado quién lo "
            "emitió, cuándo ocurrió el hecho y cuáles eran las instrucciones. Pienso que se "
            "podría mejorar la escucha activa, la verificación de fuentes y la comunicación "
            "entre el centro educativo y sus estudiantes. También sería útil pedir "
            "retroalimentación para comprobar que todos entendieron el aviso. De esa forma se evitarían rumores "
            "y cada persona podría decidir basándose en información confiable, no solamente "
            "en lo que otros publican."
        ),
    ),
]


def set_cell_borderless_font(run, bold=False):
    run.font.name = "Arial"
    run.font.size = Pt(12)
    run.bold = bold
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")


def set_docx_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    set_cell_borderless_font(run)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = "PAGE"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instruction, end])


def add_docx_paragraph(document, text, *, bold=False, centered=False, space_after=0):
    paragraph = document.add_paragraph()
    paragraph.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER if centered else WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    paragraph.paragraph_format.line_spacing = 1
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(space_after)
    run = paragraph.add_run(text)
    set_cell_borderless_font(run, bold=bold)
    return paragraph


def build_docx():
    document = Document()
    section = document.sections[0]
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

    normal = document.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(12)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")

    add_docx_paragraph(document, TITLE, bold=True, centered=True, space_after=3)
    add_docx_paragraph(document, SUBTITLE, bold=True, centered=True, space_after=14)
    add_docx_paragraph(document, STUDENT, space_after=12)

    for question, answer in QUESTIONS_AND_ANSWERS:
        add_docx_paragraph(document, question, bold=True, space_after=6)
        add_docx_paragraph(document, answer, space_after=12)

    set_docx_page_number(section.footer.paragraphs[0])
    document.core_properties.title = TITLE
    document.core_properties.subject = "Análisis de comunicación efectiva"
    document.save(DOCX_PATH)


def add_pdf_page_number(canvas, document):
    canvas.saveState()
    canvas.setFont("Arial", 9)
    canvas.drawCentredString(A4[0] / 2, 1.4 * cm, str(document.page))
    canvas.restoreState()


def build_pdf():
    pdfmetrics.registerFont(
        TTFont("Arial", "/usr/share/fonts/truetype/croscore/Arimo-Regular.ttf")
    )
    pdfmetrics.registerFont(
        TTFont("Arial-Bold", "/usr/share/fonts/truetype/croscore/Arimo-Bold.ttf")
    )

    document = SimpleDocTemplate(
        str(PDF_PATH),
        pagesize=A4,
        rightMargin=2.5 * cm,
        leftMargin=2.5 * cm,
        topMargin=2.5 * cm,
        bottomMargin=2.5 * cm,
        title=TITLE,
        subject="Análisis de comunicación efectiva",
    )

    title_style = ParagraphStyle(
        "TitleArial",
        fontName="Arial-Bold",
        fontSize=12,
        leading=12,
        alignment=TA_CENTER,
        spaceAfter=4,
    )
    subtitle_style = ParagraphStyle(
        "SubtitleArial",
        parent=title_style,
        spaceAfter=16,
    )
    question_style = ParagraphStyle(
        "QuestionArial",
        fontName="Arial-Bold",
        fontSize=12,
        leading=12,
        alignment=TA_JUSTIFY,
        spaceBefore=4,
        spaceAfter=6,
    )
    answer_style = ParagraphStyle(
        "AnswerArial",
        fontName="Arial",
        fontSize=12,
        leading=12,
        alignment=TA_JUSTIFY,
        spaceAfter=12,
    )

    story = [
        Paragraph(TITLE, title_style),
        Paragraph(SUBTITLE, subtitle_style),
        Paragraph(STUDENT, answer_style),
        Spacer(1, 6),
    ]
    for question, answer in QUESTIONS_AND_ANSWERS:
        story.append(
            KeepTogether(
                [
                    Paragraph(question, question_style),
                    Paragraph(answer, answer_style),
                ]
            )
        )

    document.build(
        story,
        onFirstPage=add_pdf_page_number,
        onLaterPages=add_pdf_page_number,
    )


def rtf_escape(text):
    escaped = []
    for character in text:
        if character in "\\{}":
            escaped.append("\\" + character)
        elif ord(character) > 127:
            value = ord(character)
            if value > 32767:
                value -= 65536
            escaped.append(f"\\u{value}?")
        else:
            escaped.append(character)
    return "".join(escaped)


def build_rtf():
    lines = [
        r"{\rtf1\ansi\ansicpg1252\deff0",
        r"{\fonttbl{\f0 Arial;}}",
        r"\paperw11907\paperh16840\margl1417\margr1417\margt1417\margb1417",
        r"\f0\fs24\sl240\slmult1",
        r"\qc\b " + rtf_escape(TITLE) + r"\b0\par",
        r"\qc\b " + rtf_escape(SUBTITLE) + r"\b0\par\par",
        r"\qj " + rtf_escape(STUDENT) + r"\par\par",
    ]
    for question, answer in QUESTIONS_AND_ANSWERS:
        lines.append(r"\qj\b " + rtf_escape(question) + r"\b0\par")
        lines.append(r"\qj " + rtf_escape(answer) + r"\par\par")
    lines.append("}")
    RTF_PATH.write_text("\n".join(lines), encoding="ascii")


if __name__ == "__main__":
    build_docx()
    build_pdf()
    build_rtf()
    print(DOCX_PATH)
    print(PDF_PATH)
    print(RTF_PATH)
