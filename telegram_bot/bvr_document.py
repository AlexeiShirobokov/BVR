# -*- coding: utf-8 -*-
"""
Генератор проекта массового взрыва: Excel (.xlsx) и PDF.
Формирует полный проект — обложка, акт, распорядок, таблицы, расчёты,
приказ, распоряжение, расчёт безопасных расстояний.
"""
import os
import datetime
from bvr_calc import MONTHS, depth_groups

# ---------- ФОРМАТИРОВАНИЕ ----------
def fmt(x, d=2):
    if x is None:
        return "0"
    s = f"{x:,.{d}f}".replace(",", " ").replace(".", ",")
    return s

def fmt0(x):
    return fmt(x, 0)

def date_parts(params):
    d = params.get("data") or datetime.date.today()
    return d.day, MONTHS[d.month - 1], d.year

def date_str(params):
    dd, mm, yy = date_parts(params)
    return f"« {dd} » {mm} {yy} г."


# ==================== EXCEL ====================
def build_excel(calc, path):
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    p = calc["params"]
    wb = openpyxl.Workbook()
    thin = Side(style="thin", color="B0B8C4")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill("solid", fgColor="1D4E89")
    sec_fill = PatternFill("solid", fgColor="EEF3F9")
    hdr_font = Font(bold=True, color="FFFFFF")
    bold = Font(bold=True)

    # --- Лист «Параметры» ---
    ws = wb.active
    ws.title = "Параметры"
    rows = [
        ("ПРОЕКТ МАССОВОГО ВЗРЫВА — расчёт БВР", "", ""),
        ("", "", ""),
        ("Предприятие", "", p["pred"]),
        ("Месторождение", "", p["mest"]),
        ("Блок", "", p["blok"]),
        ("Дата взрыва", "", date_str(p)),
        ("", "", ""),
        ("Параметр", "Ед. изм.", "Значение"),
        ("Расположение скважин", "", p["raspol"]),
        ("Плотность ВВ в заряде", "г/см³", p["plotvv"]),
        ("Условный диаметр скважин", "м", p["diam"]),
        ("КИС", "", p["kis"]),
        ("Расход селитры на 1 п.м", "кг/пог.м", calc["raskh_sel"]),
        ("ЛСПП", "м", calc["lspp"]),
        ("Коэффициент сближения зарядов", "", calc["ksbl"]),
        ("Расстояние между скважинами", "м", p["a"]),
        ("Расстояние между рядами", "м", calc["b"]),
        ("Количество рядов", "шт", p["ryad"]),
        ("Количество скважин", "шт", calc["n"]),
        ("Площадь массива", "м²", p["ploshad"]),
        ("Объём взрываемого массива", "м³", round(calc["objem_massiva"], 1)),
        ("Общий объём бурения", "п.м", round(calc["sumD"], 1)),
        ("Средняя глубина скважин", "м", round(calc["sr_glub"], 2)),
        ("Средняя высота уступа", "м", round(calc["sr_hust"], 2)),
        ("Средняя длина заряда", "м", round(calc["sr_zar"], 2)),
        ("Удельный расход ВВ", "кг/м³", round(calc["ud_vv"], 3)),
        ("", "", ""),
        ("РАСЧЁТ ВВ", "Ед.", "Значение"),
        ("Аммиачная селитра", "кг", round(calc["mas_as"], 2)),
        ("Дизтопливо", "кг", round(calc["mas_dt_kg"], 2)),
        ("Дизтопливо", "л", round(calc["mas_dt_l"], 2)),
        ("ПТ-П-2250", "кг", round(calc["mas_pt"], 2)),
        ("Общая масса заряда", "кг", round(calc["obsh_massa"], 2)),
        ("Средняя масса заряда в скважине", "кг", round(calc["sr_massa"], 2)),
        ("", "", ""),
        ("СРЕДСТВА ИНИЦИИРОВАНИЯ", "Ед.", "Значение"),
        ("Искра-С (500 мс)", "шт", calc["iskra_s"]),
        ("Искра-П (67 мс)", "шт", calc["iskra67"]),
        ("Искра-П (42 мс)", "шт", calc["iskra42"]),
        ("Искра-В (Старт-600м)", "шт", calc["iskraV"]),
        ("", "", ""),
        ("БЕЗОПАСНЫЕ РАССТОЯНИЯ", "Ед.", "Значение"),
        ("Опасная зона для людей", "м", calc["zona_ludi"]),
        ("Опасная зона для оборудования", "м", calc["zona_obor"]),
    ]
    for r in rows:
        ws.append(r)
    ws["A1"].font = Font(bold=True, size=13, color="1D4E89")
    for cell in ("A8", "B8", "C8", "A28", "B28", "C28", "A37", "B37", "C37",
                 "A43", "B43", "C43"):
        ws[cell].font = hdr_font
        ws[cell].fill = hdr_fill
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 22

    # --- Лист «Каталог скважин» ---
    ws = wb.create_sheet("Каталог скважин")
    head = ["№", "Ряд-№ скв", "Глубина, м", "Высота уступа, м", "Перебур, м",
            "Высота колонки, м", "Забойка, м", "АС, кг", "ДТ, кг", "ДТ, л",
            "Гранулит, кг", "Vгм, м³"]
    ws.append(head)
    for c in ws[1]:
        c.font = hdr_font
        c.fill = hdr_fill
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    for i, w in enumerate(calc["wells"], 1):
        ws.append([i, f"{w['ryad']}-{w['skv']}", round(w["d"], 2),
                   round(w["H"], 2), round(w["I"], 2), round(w["N"], 2),
                   round(w["zaboi"], 2), round(w["K"], 2), round(w["L"], 2),
                   round(w["M"], 2), round(w["E"], 2), round(w["J"], 2)])
    tot = ["", "ИТОГО", round(calc["sumD"], 2), round(calc["sumH"], 2),
           round(calc["sumI"], 2), round(calc["sumN"], 2),
           round(calc["n"] * p["zaboi"], 2), round(calc["mas_as"], 2),
           round(calc["mas_dt_kg"], 2), round(calc["mas_dt_l"], 2),
           round(calc["mas_as"] + calc["mas_dt_kg"], 2),
           round(calc["objem_massiva"], 2)]
    ws.append(tot)
    for c in ws[ws.max_row]:
        c.font = bold
        c.fill = PatternFill("solid", fgColor="FFF4E6")
    for col in "ABCDEFGHIJKL":
        ws.column_dimensions[col].width = 13

    # --- Лист «Зарядная карта» ---
    ws = wb.create_sheet("Зарядная карта")
    max_skv = max((w["skv"] for w in calc["wells"]), default=1)
    max_ryad = max((w["ryad"] for w in calc["wells"]), default=1)
    ws.append(["Ряд \\ Скв."] + list(range(1, max_skv + 1)))
    grid = {}
    for w in calc["wells"]:
        grid[(w["ryad"], w["skv"])] = w["d"]
    for r in range(1, max_ryad + 1):
        ws.append([r] + [grid.get((r, c), "") for c in range(1, max_skv + 1)])
    for c in ws[1]:
        c.font = hdr_font
        c.fill = hdr_fill

    # --- Лист «Таблица параметров» ---
    ws = wb.create_sheet("Таблица параметров")
    ws.append(["№", "Глубина, м", "Кол-во скв.", "Масса заряда, кг",
               "Длина заряда, м", "Длина забойки, м"])
    for c in ws[1]:
        c.font = hdr_font
        c.fill = hdr_fill
    for i, g in enumerate(depth_groups(calc), 1):
        mass = (g["sK"] + g["sL"]) / g["n"] + p["boevik"]
        ws.append([i, round(g["d"], 2), g["n"], round(mass, 2),
                   round(g["sN"] / g["n"], 2), round(g["sZ"] / g["n"], 2)])
    for col in "ABCDEF":
        ws.column_dimensions[col].width = 18

    wb.save(path)
    return path


# ==================== PDF ====================
def _register_font():
    """Подбирает TTF-шрифт с поддержкой кириллицы."""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    candidates = [
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf",
         "/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf"),
        ("/Library/Fonts/Arial Unicode.ttf",
         "/Library/Fonts/Arial Unicode.ttf"),
    ]
    for reg, bold in candidates:
        if os.path.exists(reg):
            pdfmetrics.registerFont(TTFont("BVR", reg))
            pdfmetrics.registerFont(TTFont("BVR-B", bold if os.path.exists(bold) else reg))
            return "BVR", "BVR-B"
    return "Helvetica", "Helvetica-Bold"


def build_pdf(calc, path, passport_path=None):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                    PageBreak, Table, TableStyle)

    FONT, FONT_B = _register_font()
    p = calc["params"]
    dd, _mon, yy = date_parts(p)
    D = date_str(p)

    body = ParagraphStyle("body", fontName=FONT, fontSize=10.5, leading=15,
                          alignment=4, spaceAfter=4)
    center = ParagraphStyle("center", parent=body, alignment=1)
    cbold = ParagraphStyle("cbold", parent=center, fontName=FONT_B)
    right = ParagraphStyle("right", parent=body, alignment=2)
    h1 = ParagraphStyle("h1", fontName=FONT_B, fontSize=14, leading=18,
                        alignment=1, spaceAfter=8)
    cover_big = ParagraphStyle("cb", fontName=FONT_B, fontSize=26, leading=34,
                               alignment=1)
    small = ParagraphStyle("small", parent=body, fontSize=8.5, leading=11)

    story = []
    P = lambda t, s=body: story.append(Paragraph(t, s))
    SP = lambda h=6: story.append(Spacer(1, h))

    pred = p["pred"]
    blok = p["blok"]
    mest = p["mest"]
    nachbvr = p["nachbvr"]
    gling = p["gling"]
    nachuch = p["nachuch"]
    vzr = p["vzryvnik"]

    # ---- ОБЛОЖКА ----
    SP(60)
    P("АКЦИОНЕРНОЕ ОБЩЕСТВО", cbold)
    P(pred.replace('АО', '').replace('"', '').strip().upper(), cbold)
    SP(120)
    P("П Р О Е К Т", cover_big)
    SP(6)
    P("МАССОВОГО ВЗРЫВА", ParagraphStyle("c2", parent=cover_big, fontSize=18))
    SP(40)
    P(f"месторождение: {mest}", center)
    P(f"Блок: {blok}", center)
    SP(150)
    P(p["np"], center)
    P(f"{yy} г.", center)
    story.append(PageBreak())

    # ---- АКТ ----
    P(pred, ParagraphStyle("b", parent=body, fontName=FONT_B))
    SP(10)
    P("А К Т", h1)
    P("о готовности блока к заряжанию", center)
    SP(10)
    P(f"Россыпное месторождение <b>{mest}</b>, блок <b>{blok}</b>")
    P(f"{D}&nbsp;&nbsp;&nbsp;{p['prisk']}")
    P(f"Мы, нижеподписавшиеся, ответственный руководитель ВР {nachbvr}, "
      f"маркшейдер ________________, составили настоящий акт о том, что блок "
      f"<b>{blok}</b> полностью забурен и подготовлен к заряжанию в следующих "
      f"параметрах:")
    P(f"1. Средняя глубина скважин — <b>{fmt(calc['sr_glub'])}</b> м")
    P(f"2. Площадь блока — <b>{fmt(p['ploshad']/1000,1)}</b> тыс.м²")
    P(f"3. Количество скважин — <b>{calc['n']}</b> шт.")
    P(f"4. Объём торфов на взрыв — <b>{fmt(calc['objem_massiva']/1000,2)}</b> тыс.м³")
    P("Скважины пробурены в соответствии с проектом и очищены. Блок очищен "
      "от посторонних предметов и подготовлен к производству взрывных работ.")
    SP(20)
    P(f"Ответственный руководитель ВР _____________ / {nachbvr} /")
    P("Маркшейдер _____________ / ________________ /")
    story.append(PageBreak())

    # ---- ЛИСТ 2: ОЗНАКОМЛЕНИЕ ----
    P("С проектом массового взрыва ознакомлены и проинструктированы:",
      ParagraphStyle("b", parent=body, fontName=FONT_B))
    SP(6)
    data = [["№", "Ф.И.О.", "Профессия", "Роспись"]]
    for i in range(1, 18):
        data.append([str(i), "", "", ""])
    t = Table(data, colWidths=[20, 200, 150, 90])
    t.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), FONT, 9),
        ("FONT", (0, 0), (-1, 0), FONT_B, 9),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EEF3F9")),
        ("ROWHEIGHT", (0, 0), (-1, -1), 18),
    ]))
    story.append(t)
    story.append(PageBreak())

    # ---- РАСПОРЯДОК ----
    P(f"УТВЕРЖДАЮ<br/>Главный инженер {pred}<br/>_______________ {gling}<br/>{D}", right)
    SP(6)
    P("Р А С П О Р Я Д О К", h1)
    P("проведения массового взрыва", center)
    SP(6)
    P(f"Россыпное месторождение <b>{mest}</b>, блок <b>{blok}</b>")
    items = [
        f"1. Дата взрыва: {D}",
        f"2. Место взрыва: месторождение {mest}, {p['prisk']}",
        f"3. Время взрыва: с {p['vz1']} до {p['vz2']} по местному времени",
        f"4. Зарядка производится: {D} с {p['zar1']} до {p['zar2']}",
        "5. Способы инициирования зарядов, взрывной сети: неэлектрический",
        f"6. Общее количество взрываемых скважин: <b>{calc['n']}</b> шт.",
        "7. Способ взрывания: инициирующий импульс от устройства Искра-С, "
        "с интервалом замедления 500 мс",
        "8. Тип замедлителей: ИСКРА-П (42; 67) мс",
        "9. Схема взрывания с указанием интервалов замедлений: согласно "
        "схеме (паспорт БВР)",
        "10. Порядок монтажа взрывной сети: согласно проекту и схеме монтажа",
        "11. Место расположения взрывной станции: согласно схеме",
        f"12. Опасная зона: для людей — {calc['zona_ludi']} м; для оборудования "
        f"— {calc['zona_obor']} м; для сооружений — {p['zoso']} м.",
        f"13. Объекты, находящиеся в опасной зоне: {p['obj']}",
        f"14. Мероприятия по предотвращению повреждений объектов: {p['meri']}",
        f"15. Оборудование отводится от ближайшей скважины: экскаваторы — "
        f"на {calc['zona_obor']} м, буровые станки — на {calc['zona_obor']} м.",
        "16. Схема расстановки постов охраны опасной зоны прилагается.",
        f"17. Ответственным руководителем массового взрыва назначен: "
        f"Начальник участка БВР {nachbvr}",
        "18. Подвозка взрывчатых материалов производится специализированным "
        "автомобилем.",
        "19. Для очистки скважин перед заряжанием используется буровой "
        "станок СБШ-250.",
    ]
    for it in items:
        P(it)
    story.append(PageBreak())

    # ---- ТАБЛИЦА ПАРАМЕТРОВ ----
    P("ТАБЛИЦА", h1)
    P("параметров взрывных работ", center)
    P(f"на россыпном месторождении {mest}, блок {blok}", small)
    SP(4)
    th = [["№", "Глубина\nскв., м", "Кол-во\nскв.", "Расст.\nм/у скв., м",
           "Расст.\nм/у ряд., м", "Диаметр\nскв., мм", "Масса\nзаряда, кг",
           "Длина\nзаряда, м", "Длина\nзабойки, м"]]
    for i, g in enumerate(depth_groups(calc), 1):
        mass = (g["sK"] + g["sL"]) / g["n"] + p["boevik"]
        th.append([str(i), fmt(g["d"], 1), str(g["n"]), fmt(p["a"], 1),
                   fmt(calc["b"], 1), str(round(p["diam"] * 1000)),
                   fmt(mass, 1), fmt(g["sN"] / g["n"], 1),
                   fmt(g["sZ"] / g["n"], 1)])
    t = Table(th, colWidths=[24, 56, 50, 58, 58, 56, 60, 56, 58])
    t.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), FONT, 8),
        ("FONT", (0, 0), (-1, 0), FONT_B, 8),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EEF3F9")),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(t)
    SP(8)
    P(f"Расчётный усреднённый вес заряда: боевик ПТ-П-2250 — {fmt(p['boevik'])} кг; "
      f"селитра/скв — {fmt(calc['mas_as']/calc['n'] if calc['n'] else 0)} кг; "
      f"дизтопливо — {fmt(calc['mas_dt_kg']/calc['n'] if calc['n'] else 0)} кг; "
      f"итого — {fmt(calc['sr_massa'])} кг.", small)
    P(f"Общая масса ВМ: Гранулит М (АС+ДТ) — {fmt(calc['mas_as']+calc['mas_dt_kg'],1)} кг; "
      f"ПТ-П-2250 — {fmt(calc['mas_pt'],1)} кг; итого ВМ — {fmt(calc['obsh_massa'],1)} кг.", small)
    SP(10)
    P(f"Расчёт составил: Начальник участка БВР {nachbvr}&nbsp;&nbsp;&nbsp;"
      f"Расчёт проверил: Главный инженер {gling}")
    story.append(PageBreak())

    # ---- ЛИСТ 5 ----
    P(f"Инструктаж работников провёл: {nachbvr}")
    for it in [
        f"1. Непосредственное руководство работой персонала при подготовке "
        f"и проведении массового взрыва осуществляет: Начальник участка БВР {nachbvr}.",
        f"2. Ответственным за вывод людей с территории запретной и опасной "
        f"зоны назначен: {nachuch}.",
        f"3. Ответственным за заряжание и монтаж взрывной (электровзрывной) "
        f"сети назначен: Старший взрывник {vzr}.",
        f"4. Ответственным за вывод внутрикарьерного транспорта из запретной "
        f"и опасной зоны назначен: {nachuch}.",
        f"5. Ответственным за охрану запретной и опасной зон назначен: {nachuch}.",
        f"6. Ответственным за отключение электроэнергии назначен: {nachuch}.",
        f"7. Ответственным за подачу звуковых и световых сигналов назначен: "
        f"Начальник участка БВР {nachbvr}.",
        f"8. Ответственным за оповещение соседних предприятий и подразделений "
        f"назначен: {nachbvr}.",
        "9. Подача сигналов проводится звуковой сиреной, установленной "
        "на зарядно-смесительной машине.",
        "10. После выставления постов подаётся предупредительный сигнал — "
        "один продолжительный звуковой.",
        "11. По указанию ответственного за вывод людей все люди, не занятые "
        "заряжанием, выводятся в безопасное место согласно схеме.",
        f"12. Заряжание скважин осуществляет: Старший взрывник {vzr}.",
    ]:
        P(it)
    story.append(PageBreak())

    # ---- ЛИСТ 7 ----
    P(f"1. Укладку в заряды боевиков с капсюлем, монтаж взрывной сети "
      f"осуществляет: Старший взрывник {vzr}, под руководством ответственного "
      f"руководителя ВР {nachbvr}.")
    P("2. Боевой сигнал: два длинных звуковых.")
    P("3. После подачи боевого сигнала производится взрыв (неэлектрический "
      "способ взрывания).")
    P(f"4. Сигнал «отбой»: три коротких после осмотра места взрыва, после "
      f"получения указания от ответственного руководителя ВР {nachbvr}.")
    P("5. Время проветривания и допуска людей к месту взрыва — не менее 30 мин.")
    P("6. С распорядком проведения массового взрыва ознакомлены:",
      ParagraphStyle("b", parent=body, fontName=FONT_B))
    for it in [
        f"Ответственный руководитель ВР: {nachbvr} _______________",
        f"Ответственный за вывод людей из опасной и запретной зон: {nachuch} _______________",
        f"Ответственный за заряжание и монтаж взрывной сети: {vzr} _______________",
        f"Ответственный за охрану запретной и опасной зон: {nachuch} _______________",
        f"Ответственный за подачу сигналов: {nachbvr} _______________",
    ]:
        P(it)
    SP(8)
    P(f"Составил: Начальник участка БВР {nachbvr}&nbsp;&nbsp;&nbsp;"
      f"Проверил: Главный инженер {gling}")
    P(D, right)
    story.append(PageBreak())

    # ---- ЛИСТ 8: ПРИКАЗ ----
    P(f"П Р И К А З № {p['prikaz']}", h1)
    P(f"от {D}&nbsp;&nbsp;&nbsp;{p['np']}", center)
    SP(4)
    P(f"О производстве массового взрыва {D}", cbold)
    P(f"на месторождении {mest}, блок № {blok}.", center)
    SP(4)
    P(f"Согласно графику производства массовых взрывов и проекту массового "
      f"взрыва на россыпном месторождении {mest}, блок {blok}, {D} "
      f"с {p['vz1']} до {p['vz2']} будет произведён массовый взрыв.")
    P("В целях обеспечения безопасности работ и сохранности ВМ,")
    P("ПРИКАЗЫВАЮ:", cbold)
    for it in [
        f"1. Проект производства массового взрыва на блок {blok}, разработанный "
        f"на основании Типового проекта, утвердить.",
        f"2. Ответственным руководителем массового взрыва назначить: "
        f"Начальник участка БВР {nachbvr}.",
        f"3. Ответственность за вывод людей за пределы опасной зоны взрыва "
        f"возложить на: {nachuch}.",
        f"4. Ответственность за вывод техники за пределы опасной зоны "
        f"возложить на: {nachuch}.",
        f"5. Непосредственное руководство работой персонала при подготовке "
        f"взрыва возложить на: Начальник участка БВР {nachbvr}.",
        f"6. Взрывником назначить: Старший взрывник {vzr}.",
        f"7. Зарядку скважин произвести: {D} с {p['zar1']} до {p['zar2']}.",
        f"8. Взрыв произвести: {D} с {p['vz1']} до {p['vz2']}.",
    ]:
        P(it)
    story.append(PageBreak())

    # ---- ЛИСТ 9 ----
    for it in [
        f"9. Ответственность за отключение электроэнергии возложить на: {nachuch}.",
        "10. Оповещение ответственного руководителя взрывных работ, "
        "диспетчерской службы Зонального центра и Якутского района провести "
        "в установленном порядке.",
        "11. Ответственному руководителю ВР обеспечить проведение инструктажа "
        "персонала, привлекаемого к массовому взрыву.",
        "12. Все назначенные настоящим приказом ответственные лица на период "
        "подготовки и проведения массового взрыва обязаны находиться "
        "на своих местах.",
        "13. Контроль за исполнением настоящего приказа оставляю за собой.",
    ]:
        P(it)
    SP(40)
    P(f"Главный инженер {pred}&nbsp;&nbsp;&nbsp;_______________ {gling}")
    story.append(PageBreak())

    # ---- ЛИСТ 11: РАСПОРЯЖЕНИЕ О ПОСТАХ ----
    P("РАСПОРЯЖЕНИЕ", h1)
    P("о расстановке постов оцепления по охране опасной зоны", center)
    P(f"ВР на месторождении {mest}, блок {blok}", center)
    SP(4)
    P(f"Взрыв {D}, время взрыва с {p['vz1']} до {p['vz2']} часов.")
    P(f"{nachuch}, в целях обеспечения безопасности при производстве "
      f"массового взрыва, выставляет следующие посты оцепления:")
    for i in range(1, 6):
        P(f"Пост №{i} — перекрывает подходы со стороны __________________.")
        P("Ответственный ________________ (подпись) ________________")
    P("ВСЕМ ПОСТОВЫМ!", ParagraphStyle("b", parent=body, fontName=FONT_B))
    P("1. Бдительно наблюдать за местностью, самовольно не покидать пост.")
    P("2. Внимательно следить за подаваемыми сигналами:")
    P("ОДИН ДЛИННЫЙ звуковой сигнал — предупредительный;")
    P("ДВА ДЛИННЫХ звуковых сигнала — боевой;")
    P("ТРИ КОРОТКИХ звуковых сигнала — отбой.")
    SP(8)
    P(f"Ответственный за выставление постов: {nachuch} _____________")
    P(f"Начальник участка БВР {pred}: {nachbvr} _____________")
    story.append(PageBreak())

    # ---- ЛИСТ 12: РАСЧЁТ ПАРАМЕТРОВ ----
    P("Расчёт параметров буровзрывных работ и потребности материалов", cbold)
    SP(4)
    P(f"Вместимость погонного метра скважины (Р) определяется по справочным "
      f"таблицам исходя из диаметра скважины и плотности ВВ в заряде "
      f"и составит: <b>{fmt(calc['raskh_sel'])}</b> кг/пог.м.")
    P("Линия сопротивления по подошве уступа (ЛСПП) для вертикального заряда "
      "определяется по формуле:")
    P("Wп = 53 · kт · dзар · √((ρвв·1000 − l) / ρпор)", center)
    P(f"где: kт — коэффициент местных геологических условий, kт = {fmt(p['ktresh'],1)}; "
      f"dзар — диаметр заряда, dзар = {fmt(p['diam'],3)} м; "
      f"ρвв — плотность ВВ, ρвв = {fmt(p['plotvv'])} г/см³; "
      f"l — коэффициент работоспособности ВВ, l = {fmt(p['krab'])}; "
      f"ρпор — плотность породы, ρпор = {fmt0(p['plotpor'])} кг/м³.")
    P(f"Wп = 53 · {fmt(p['ktresh'],1)} · {fmt(p['diam'],3)} · "
      f"√(({fmt0(p['plotvv']*1000)} − {fmt(p['krab'])}) / {fmt0(p['plotpor'])}) "
      f"= <b>{fmt(calc['lspp'],1)}</b> м.", center)
    P(f"Расстояние между скважинами в ряду: a = Wп · m = {fmt(calc['lspp'],1)} · "
      f"{fmt(calc['ksbl'])} ≈ <b>{fmt(p['a'],1)}</b> м; между рядами b = "
      f"<b>{fmt(calc['b'],1)}</b> м, где m — коэффициент сближения зарядов, "
      f"m = {fmt(calc['ksbl'])}.")
    P(f"Число рядов скважинных зарядов принято в количестве <b>{p['ryad']}</b> шт. "
      f"Расположение скважин — {p['raspol'].lower()}.")
    P(f"Величина перебура принимается равной <b>{fmt(calc['sr_pereb'],1)}</b> м "
      f"(расчётная Lпер = 14 · dзар · √(ρвв · 0,8) = {fmt(calc['l_per_form'])} м).")
    story.append(PageBreak())

    # ---- ЛИСТ 13 ----
    P("Глубина скважины устанавливается по формуле:")
    P(f"Lс = H + Lпер = {fmt(calc['sr_hust'])} + {fmt(calc['sr_pereb'])} = "
      f"<b>{fmt(calc['sr_glub'])}</b> м.", center)
    P("Коэффициент использования скважины (Кис):")
    P(f"Кис = H / Lс = {fmt(calc['sr_hust'])} / {fmt(calc['sr_glub'])} = "
      f"<b>{fmt(p['kis'])}</b>.", center)
    P("Вес заряда в скважине (Qскв) принимается по фактическим данным "
      "каталога скважин:")
    asn = calc['mas_as']/calc['n'] if calc['n'] else 0
    dtn = calc['mas_dt_kg']/calc['n'] if calc['n'] else 0
    P(f"Qскв = Qас/скв + Qдт/скв + Qбоевика = {fmt(asn)} + {fmt(dtn)} + "
      f"{fmt(p['boevik'])} = <b>{fmt(calc['sr_massa'])}</b> кг.", center)
    P(f"Длина заряда: Lзар = ΣLзар / nскв = <b>{fmt(calc['sr_zar'])}</b> м.")
    P(f"Длина забойки: Lзаб = ΣLзаб / nскв = <b>{fmt(calc['sr_zaboi'])}</b> м.")
    P(f"Площадь на одну скважину: Sскв = a · b = {fmt(p['a'],1)} · "
      f"{fmt(calc['b'],1)} = <b>{fmt(calc['s_per_skv'])}</b> м².")
    P(f"Выход горной массы: Vгм/пог.м = Vмаркш / ΣLс = "
      f"{fmt0(calc['objem_massiva'])} / {fmt0(calc['sumD'])} = "
      f"<b>{fmt(calc['vyhod_gm'])}</b> м³/пог.м.")
    P(f"Объём взрываемой горной массы: Vгм = <b>{fmt(calc['objem_massiva'],1)}</b> м³.")
    story.append(PageBreak())

    # ---- ЛИСТ 14: СУММАРНЫЙ РАСХОД + ОСНОВНЫЕ ПАРАМЕТРЫ ----
    P("Суммарный расход ВВ (игданита) на взрыв:",
      ParagraphStyle("b", parent=body, fontName=FONT_B))
    P(f"Qселитры = Nскв · Qскв = {calc['n']} · {fmt(asn)} = "
      f"{fmt(calc['mas_as']/1000)} т;")
    P(f"Qдт = Qселитры · {fmt(p['doldt'],3)} = {fmt(calc['mas_dt_kg']/1000)} т;")
    P(f"Mгранулита = {fmt(calc['mas_as']/1000)} + {fmt(calc['mas_dt_kg']/1000)} "
      f"= {fmt((calc['mas_as']+calc['mas_dt_kg'])/1000)} т.")
    P(f"Расход ПТ-П-2250 на боевики: {fmt(p['boevik'])} · {calc['n']} = "
      f"{fmt(calc['mas_pt'])} кг.")
    P(f"Расход СИ: Искра-С — {calc['iskra_s']} шт; Искра-П (67 мс) — "
      f"{calc['iskra67']} шт; Искра-П (42 мс) — {calc['iskra42']} шт; "
      f"Искра-В (Старт) — {calc['iskraV']} шт.")
    SP(4)
    P("Основные параметры буровзрывных работ", cbold)
    op = [
        ("1", "Площадь полигона", "тыс.кв.м", fmt(p["ploshad"]/1000, 1)),
        ("2", "Высота уступа", "м", fmt(calc["sr_hust"])),
        ("3", "Объём взрыва", "тыс.куб.м", fmt(calc["objem_massiva"]/1000)),
        ("4", "Средняя глубина скважин", "м", fmt(calc["sr_glub"])),
        ("5", "в т.ч. перебур", "м", fmt(calc["sr_pereb"])),
        ("6", "Диаметр скважины", "мм", str(round(p["diam"]*1000))),
        ("7", "КИС", "", fmt(p["kis"])),
        ("8", "ЛСПП", "м", fmt(calc["lspp"], 1)),
        ("9", "Расстояние между скважинами", "м", fmt(p["a"], 1)),
        ("10", "Расстояние между рядами", "м", fmt(calc["b"], 1)),
        ("11", "Расположение скважин", "", p["raspol"]),
        ("12", "Количество скважин", "шт", str(calc["n"])),
        ("13", "Объём бурения скважин", "пог.м", fmt0(calc["sumD"])),
        ("14", "Удельный расход Гранулита М", "кг/пог.м",
         fmt(calc["ud_vv"], 3)),
        ("15", "Длина заряда", "м", fmt(calc["sr_zar"])),
        ("16", "Длина забойки", "м", fmt(calc["sr_zaboi"])),
        ("17", "Величина заряда в скважине", "кг", fmt(calc["sr_massa"])),
        ("18", "Выход горной массы на 1 пог.м", "куб.м", fmt(calc["vyhod_gm"])),
        ("19", "Расход Гранулита М", "т",
         fmt((calc["mas_as"]+calc["mas_dt_kg"])/1000, 3)),
        ("20", "Удельный расход АС на 1 куб.м ГМ", "кг/куб.м", fmt(calc["ud_as"], 3)),
        ("21", "ПТ-П-2250", "кг", fmt(calc["mas_pt"])),
        ("22", "Расход ИСКРА-С", "шт", str(calc["iskra_s"])),
        ("23", "Расход Искра-П (67 мс)", "шт", str(calc["iskra67"])),
        ("24", "Расход Искра-П (42 мс)", "шт", str(calc["iskra42"])),
        ("25", "Расход Искра-В (Старт-600м)", "шт", str(calc["iskraV"])),
        ("26", "Расход аммиачной селитры на взрыв", "т", fmt(calc["mas_as"]/1000, 3)),
        ("27", "Удельный расход аммиачной селитры", "кг/м3", fmt(calc["ud_as"], 3)),
    ]
    tdata = [["№", "Параметр", "Ед. изм.", "Значение"]] + [list(r) for r in op]
    t = Table(tdata, colWidths=[26, 250, 80, 90])
    t.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), FONT, 8.5),
        ("FONT", (0, 0), (-1, 0), FONT_B, 8.5),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EEF3F9")),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("ALIGN", (2, 0), (-1, -1), "CENTER"),
    ]))
    story.append(t)
    story.append(PageBreak())

    # ---- ЛИСТ 15 ----
    P("Расчёт безопасных расстояний при взрывных работах", cbold)
    P("Расчёт производится согласно главе XII ФНП «Правила безопасности "
      "при взрывных работах».")
    P("Определение зон, опасных по разлёту отдельных кусков породы.",
      ParagraphStyle("b", parent=body, fontName=FONT_B))
    P("Расстояние, безопасное для людей при взрывании скважинных зарядов:")
    P("rраз = 1250 · nз · √(4d / ((1+1)·b)), м", center)
    P(f"где d — диаметр скважины, d = {fmt(p['diam'],3)} м; b — расстояние "
      f"между рядами, b = {fmt(calc['b'],1)} м; nз — коэффициент заполнения "
      f"скважин ВВ.")
    P(f"nз = Lз / Lс = {fmt(calc['sr_zar'])} / {fmt(calc['sr_glub'])} = "
      f"<b>{fmt(calc['nz'],3)}</b>.")
    P(f"rраз = 1250 · {fmt(calc['nz'],3)} · {fmt(calc['h9'])} = "
      f"<b>{fmt(calc['r_raz'],1)}</b> м.", center)
    P(f"Коэффициент рельефа Kр = {fmt(calc['kr'])} (превышение H = "
      f"{fmt(p['prev'],1)} м).")
    P(f"Rраз = rраз · Kр = {fmt(calc['r_raz'],1)} · {fmt(calc['kr'])} = "
      f"<b>{fmt(calc['R_raz'],1)}</b> м.", center)
    story.append(PageBreak())

    # ---- ЛИСТ 16 ----
    import math as _m
    P(f"Окончательное значение радиуса опасной зоны по разлёту кусков породы "
      f"принимается равным <b>{max(1000, _m.ceil(calc['R_raz']/100)*100)}</b> м.")
    P("Определение расстояний, безопасных по действию УВВ.",
      ParagraphStyle("b", parent=body, fontName=FONT_B))
    P("Rв = 10 · Q^(1/3), где Q — масса заряда ВВ блока.")
    P(f"Rв = 10 · {fmt0(calc['obsh_massa'])}^(1/3) = <b>{fmt(calc['r_uvv'],1)}</b> м.",
      center)
    P(f"Безопасное расстояние по действию УВВ для зданий и сооружений — "
      f"<b>{calc['r_uvv_zd']}</b> м.")
    P("Расчёт сейсмически безопасных расстояний.",
      ParagraphStyle("b", parent=body, fontName=FONT_B))
    P(f"Q = 2 · {fmt(calc['sr_massa'],1)} · 0,9 + 1 = {fmt0(calc['q_seism'])} кг.")
    P(f"rс = 15,14 · Q^(1/3) = 15,14 · {fmt0(calc['q_seism'])}^(1/3) = "
      f"<b>{fmt(calc['r_seism'],1)}</b> м.", center)
    P("Расчёт расстояний, безопасных по высоте разлёта кусков породы.",
      ParagraphStyle("b", parent=body, fontName=FONT_B))
    P(f"rразл.в = 1,4 · rраз = 1,4 · {fmt(calc['r_raz'],1)} = "
      f"<b>{fmt(calc['r_raz_v'],1)}</b> м.", center)
    story.append(PageBreak())

    # ---- ЛИСТ 17 ----
    P("Расчёт расстояния, исключающего передачу детонации при хранении ВМ.",
      ParagraphStyle("b", parent=body, fontName=FONT_B))
    P("Взрывчатые материалы при доставке к месту работ размещаются раздельно "
      "с соблюдением расчётных расстояний, исключающих передачу детонации "
      "между активным и пассивным зарядами.")
    P("По результатам расчётов радиус опасной зоны принимается:",
      ParagraphStyle("b", parent=body, fontName=FONT_B))
    P(f"для людей — <b>{calc['zona_ludi']}</b> м;")
    P(f"для оборудования — <b>{calc['zona_obor']}</b> м;")
    P(f"для сооружений — <b>{p['zoso']}</b> м.")
    P("Принятые расстояния не менее расчётных по разлёту кусков породы, "
      "действию УВВ и сейсмическому действию взрыва.")
    SP(20)
    P(f"Расчёт составил: Начальник участка БВР {nachbvr} _____________")
    SP(14)
    P("Проект массового взрыва со всеми графическими материалами хранится "
      "в делах взрывного участка (цеха) до полной отработки взорванного блока "
      "(ФНиП «Правила безопасности при взрывных работах»).", small)
    if passport_path:
        P(f"Приложение: паспорт буровзрывных работ — "
          f"{os.path.basename(passport_path)}.", small)

    doc = SimpleDocTemplate(path, pagesize=A4,
                            leftMargin=20*mm, rightMargin=15*mm,
                            topMargin=15*mm, bottomMargin=15*mm)
    doc.build(story)
    return path
