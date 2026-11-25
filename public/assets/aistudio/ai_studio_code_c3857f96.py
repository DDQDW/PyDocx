from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn


def create_essay(filename="essay_plasmon_sensing.docx"):
    doc = Document()

    # Настройка страницы A4 и полей
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(1.5)

    # Настройка базового стиля: Times New Roman, 14 pt
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(14)
    r_fonts = style.element.rPr.rFonts
    r_fonts.set(qn("w:ascii"), "Times New Roman")
    r_fonts.set(qn("w:hAnsi"), "Times New Roman")
    r_fonts.set(qn("w:cs"), "Times New Roman")
    r_fonts.set(qn("w:eastAsia"), "Times New Roman")

    # Заглавие
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before = Pt(0)
    title.paragraph_format.space_after = Pt(4)
    title_run = title.add_run(
        "Применения плазмон-усиленной фотолюминесценции в сенсорике"
    )
    title_run.bold = True
    title_run.font.name = "Times New Roman"
    title_run.font.size = Pt(14)

    # Подзаголовок
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.paragraph_format.space_before = Pt(0)
    subtitle.paragraph_format.space_after = Pt(6)
    sub_run = subtitle.add_run("Эссе по нанохимии")
    sub_run.italic = True
    sub_run.font.name = "Times New Roman"
    sub_run.font.size = Pt(14)

    # Информация об авторе
    author = doc.add_paragraph()
    author.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    author.paragraph_format.space_before = Pt(0)
    author.paragraph_format.space_after = Pt(12)
    author_run = author.add_run(
        "Выполнил: студент 5 курса, 7 группы\nКузьменко Станислав"
    )
    author_run.italic = True
    author_run.font.name = "Times New Roman"
    author_run.font.size = Pt(14)

    # Текст эссе
    paragraphs = [
        (
            "Плазмон-усиленная фотолюминесценция (ПУФЛ) является одним из наиболее перспективных "
            "направлений современной нанохимии и нанофотоники. В основе явления лежит "
            "взаимодействие люминофоров (органических красителей, квантовых точек или "
            "ап-конверсионных наночастиц) с локализованным поверхностным плазмонным резонансом "
            "металлических наноструктур — преимущественно золота и серебра. При контролируемом "
            "нанометровом зазоре порядка 5–15 нм, который задается диэлектрической оболочкой из "
            "диоксида кремния или полимера, резко усиливается локальное электромагнитное поле и "
            "возрастает скорость излучательной рекомбинации за счет эффекта Парселла. Это приводит "
            "к многократному увеличению яркости и фотостабильности эмиттеров, предотвращая "
            "безызлучательное тушение люминесценции."
        ),
        (
            "Ключевой областью применения ПУФЛ стала биомедицинская сенсорика и высокочувствительная "
            "клиническая диагностика. Использование наноструктурированных золотых и серебряных "
            "подложек в твердофазном иммуноанализе и ДНК-микрочипах позволяет снизить предел "
            "обнаружения белковых биомаркеров, вирусных частиц и нуклеиновых кислот до "
            "фемтомолярных концентраций. Особое место занимают гибридные наносенсоры структуры "
            "«ядро–оболочка», объединяющие плазмонное ядро с люминесцентной меткой. Такие системы "
            "обеспечивают высокий оптический контраст при внутриклеточной визуализации и дают "
            "возможность детектировать единичные биомолекулы даже на фоне собственной "
            "автофлуоресценции биологических тканей."
        ),
        (
            "Широкое распространение плазмон-усиленная фотолюминесценция получила в химическом "
            "и экологическом мониторинге для выявления следовых количеств ионов тяжелых металлов "
            "(ртути, свинца, кадмия), пестицидов, токсичных газов и взрывчатых веществ. Работа "
            "многих сенсоров основана на изменении расстояния между металлом и флуорофором при "
            "связывании аналита с помощью аптамеров или молекулярно-импринтированных полимеров, "
            "что вызывает резкое переключение сигнала между тушением и усилением свечения. Кроме "
            "того, целенаправленный синтез наночастиц заданной формы (наностержней, нанозвезд) "
            "позволяет настраивать спектр плазмонного резонанса сразу под несколько красителей "
            "для проведения мультиплексного анализа."
        ),
        (
            "Таким образом, применение ПУФЛ в сенсорике объединяет достижения коллоидного "
            "наносинтеза и оптической спектроскопии, повышая чувствительность анализа на несколько "
            "порядков. Интеграция плазмон-усиленных платформ в микрофлюидные чипы и портативные "
            "оптические детекторы открывает широкие перспективы для быстрого и надежного контроля "
            "в медицине, фармацевтике и охране окружающей среды."
        ),
    ]

    for text in paragraphs:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(1.25)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)

        run = p.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(14)

    doc.save(filename)
    print(f"Документ успешно сохранен: {filename}")


if __name__ == "__main__":
    create_essay()