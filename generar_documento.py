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
            "El emisor y el receptor son fundamentales porque la comunicación efectiva "
            "no depende únicamente de enviar información, sino también de que esta sea "
            "comprendida correctamente. El emisor tiene la responsabilidad de formular "
            "un mensaje claro, coherente, verificable y apropiado para la situación. "
            "Además, debe escoger un canal adecuado y considerar las características de "
            "la audiencia antes de difundirlo. En el cortometraje, el supuesto aviso de "
            "bomba se propaga por redes sociales sin que los estudiantes conozcan con "
            "certeza su origen ni comprueben su autenticidad. Esto muestra un emisor poco "
            "responsable o difícil de identificar, cuyo mensaje se apoya en el temor para "
            "influir en las decisiones de otros. Por su parte, el receptor no es pasivo: "
            "debe interpretar, preguntar, contrastar fuentes y entregar retroalimentación. "
            "La mayoría de los jóvenes acepta el aviso como verdadero y decide no abordar "
            "el metro ni asistir a clases, mientras que solo dos estudiantes actúan de "
            "manera diferente y concurren al examen sorpresa. Por lo tanto, para que exista "
            "comunicación efectiva, el emisor debe transmitir con responsabilidad y el "
            "receptor debe escuchar o leer activamente, evaluar el contenido y verificarlo "
            "antes de actuar o reenviarlo."
        ),
    ),
    (
        "2. ¿Cómo influye el canal elegido en la percepción del mensaje?",
        (
            "El canal influye directamente en la forma en que las personas interpretan, "
            "valoran y reaccionan frente a un mensaje. En la situación observada, las "
            "redes sociales permiten que el supuesto aviso de bomba circule de manera "
            "inmediata entre los estudiantes. Esa rapidez, sumada a la posibilidad de "
            "reenviar el contenido muchas veces, puede producir una apariencia de verdad: "
            "si varias personas repiten la misma información, los receptores pueden creer "
            "que está confirmada, aunque todos provengan de una sola fuente no identificada. "
            "El formato digital también reduce elementos importantes de la comunicación "
            "cara a cara, como el tono de voz, los gestos, la oportunidad de formular "
            "preguntas y la respuesta inmediata del emisor. Además, palabras como “bomba” "
            "y “peligro” generan miedo y urgencia, emociones que dificultan una evaluación "
            "serena del contenido. El carácter informal de las redes favorece una reacción "
            "espontánea y la difusión impulsiva. Así, el canal vuelve el mensaje persuasivo "
            "por su velocidad y alcance, pero no garantiza su confiabilidad. Un aviso de "
            "seguridad comunicado por una autoridad, el Metro o la institución educativa "
            "habría sido percibido como más legítimo y habría permitido confirmar las "
            "medidas que correspondía adoptar."
        ),
    ),
    (
        "3. Considerando el canal que se utiliza en la situación comunicativa, "
        "¿qué barreras pueden identificar? ¿Es el más adecuado para el mensaje y la audiencia?",
        (
            "La principal barrera es la falta de verificación de la fuente, pues los "
            "estudiantes reciben el aviso mediante redes sociales y no interactúan con el "
            "supuesto remitente. También existe una barrera semántica: un mensaje breve, "
            "sin contexto ni detalles comprobables, puede interpretarse de distintas formas. "
            "A esto se suma el ruido informativo propio de las plataformas digitales, donde "
            "rumores, opiniones y datos reales aparecen mezclados. La repetición entre "
            "compañeros funciona como presión social y refuerza el contenido sin aportar "
            "evidencia nueva. El miedo provocado por una posible bomba constituye una "
            "barrera emocional, porque favorece decisiones apresuradas. Asimismo, la falta "
            "de retroalimentación impide aclarar quién emitió el aviso, cuándo ocurrió el "
            "hecho y qué instrucciones oficiales existen. Las redes sociales sí son útiles "
            "para comunicar con rapidez a una audiencia joven y numerosa, pero no son por "
            "sí solas el canal más adecuado para confirmar una emergencia. En un caso de "
            "seguridad, el mensaje debería provenir de canales oficiales del Metro, de las "
            "autoridades o de la institución educativa. Las redes podrían complementar esa "
            "información, siempre que indiquen la fuente, la hora, las instrucciones y un "
            "medio oficial donde verificarla."
        ),
    ),
    (
        "4. Basándose en lo observado, ¿qué podemos concluir sobre la comunicación "
        "a partir de este cortometraje? ¿Es efectiva? ¿Qué aspectos se podrían mejorar?",
        (
            "El cortometraje permite concluir que comunicar con rapidez no equivale a "
            "comunicar de manera efectiva. El aviso difundido por redes sociales sí logra "
            "persuadir a la mayoría de los estudiantes y modifica su conducta, porque "
            "deciden no usar el metro ni asistir a clases. Sin embargo, el proceso no es "
            "efectivo desde una perspectiva responsable: el origen del mensaje no está "
            "claro, la información no se verifica y no existe una interacción que permita "
            "resolver dudas. El resultado demuestra el poder del discurso digital y la "
            "facilidad con que una idea se amplifica cuando apela al miedo y circula dentro "
            "de un grupo. También evidencia una escucha poco activa, ya que los receptores "
            "aceptan y reproducen el contenido sin analizarlo críticamente. Para mejorar la "
            "comunicación, primero se debería identificar la fuente original y contrastar "
            "el aviso con medios oficiales. Después, sería necesario comunicar datos "
            "concretos, como el lugar, la hora, la autoridad responsable y las instrucciones "
            "de seguridad. Los estudiantes deberían evitar reenviar información no "
            "confirmada y consultar directamente al establecimiento educacional. Finalmente, "
            "la institución podría disponer de un canal oficial para emergencias y cambios "
            "académicos. Estas acciones aportarían claridad, retroalimentación, responsabilidad "
            "y confianza, elementos indispensables para una comunicación verdaderamente efectiva."
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
