# -*- coding: utf-8 -*-
"""
Расчётный модуль буровзрывных работ (БВР).
Логика полностью идентична веб-форме «Форма_расчёта_БВР.html».
Сверено с эталоном «Проект БВР.xlsm».
"""
import math

MONTHS = ["января", "февраля", "марта", "апреля", "мая", "июня",
          "июля", "августа", "сентября", "октября", "ноября", "декабря"]

# Усреднённая температура воздуха по месяцам (Усть-Нера), °C
TEMP = [-46, -42, -31, -14, 3, 13, 15, 10, 2, -15, -35, -46]
# Плотность дизельного топлива по месяцам, кг/л
DT_ZIM = [0.88785, 0.88495, 0.876975, 0.86465, 0.852325, 0.845075,
          0.843625, 0.84725, 0.85305, 0.865375, 0.879875, 0.88785]
DT_ARK = [0.878708, 0.875756, 0.867638, 0.855092, 0.842546, 0.835166,
          0.83369, 0.83738, 0.843284, 0.85583, 0.87059, 0.878708]

# Значения параметров по умолчанию
DEFAULTS = {
    "pred": 'АО "Поиск Золото"',
    "mest": "Хаптагай-Хая",
    "blok": "П-14; С1-14в",
    "prisk": 'Прииск "Эрел"',
    "np": "п. Усть-Нера",
    "data": None,                # дата взрыва, datetime.date
    "zar1": "07-00", "zar2": "17-00",
    "vz1": "18-00", "vz2": "18-25",
    "prikaz": "37-ВМ",
    "raspol": "Шахматное",       # или "Прямоугольное"
    "plotvv": 0.82,              # плотность ВВ, г/см³
    "viddt": "Зимнее",           # или "Арктическое"
    "diam": 0.215,               # условный диаметр скважин, м
    "kis": 0.9,                  # коэффициент использования скважины
    "krab": 1.13,                # k работоспособности ВВ
    "plotpor": 2250,             # плотность пород, кг/м³
    "ktresh": 1.0,               # k трещиноватости
    "prev": 0.0,                 # превышение верхней отметки, м
    "a": 5.5,                    # расстояние между скважинами, м
    "zaboi": 2.0,                # длина забойки, м
    "ryad": 6,                   # количество рядов
    "ploshad": 8200,             # площадь взрываемого массива, м²
    "boevik": 2.25,              # масса боевика ПТ-П-2250, кг
    "doldt": 0.055,              # доля ДТ в игданите
    "iskra42": 20,               # Искра-П (42 мс), шт
    "iskraV": 1,                 # Искра-В (Старт-600м), шт
    "zoob": 600,                 # опасная зона для оборудования, м
    "zoso": 600,                 # опасная зона для сооружений, м
    "obj": "ЛЭП-6кв",
    "meri": "Отключение и демонтаж линии электропередач ЛЭП-6кв",
    "gendir": "Каздобин А.В.",
    "gling": "Парамонов В.И.",
    "nachbvr": "Сырбу Г.П.",
    "glgeo": "Помалейко А.А.",
    "glmark": "Ефимов С.Ф.",
    "nachuch": "Шмелёв С.В.",
    "vzryvnik": "Фурман А.С.",
}


def make_params(overrides=None):
    """Возвращает словарь параметров с применёнными переопределениями."""
    p = dict(DEFAULTS)
    if overrides:
        p.update(overrides)
    return p


def ceil_to(x, step):
    return math.ceil(x / step) * step


def density_dt(viddt, month_idx):
    """Плотность ДТ по виду топлива и индексу месяца (0..11)."""
    table = DT_ZIM if viddt == "Зимнее" else DT_ARK
    return table[month_idx]


def calculate(params, wells):
    """
    Выполняет расчёт БВР.
    params  — словарь параметров (см. DEFAULTS).
    wells   — список словарей {'ryad': int, 'skv': int, 'd': float}, d>0.
    Возвращает словарь с результатами расчёта.
    """
    p = params
    month_idx = (p["data"].month - 1) if p.get("data") else 4
    plotdt = density_dt(p["viddt"], month_idx)
    temp = TEMP[month_idx]

    # Расстояние между рядами
    b = p["a"] if p["raspol"] == "Прямоугольное" else p["a"] * 0.917
    b = round(b, 1)

    # Расход селитры на 1 п.м
    raskh_sel = round(math.pi * p["diam"] ** 2 / 4 * p["plotvv"] * 1000, 2)

    n = len(wells)
    sumD = sumH = sumN = sumK = sumL = sumM = sumI = 0.0
    rows = []
    for w in wells:
        d = float(w["d"])
        N = max(0.0, d - p["zaboi"])          # высота колонки
        zaboi = p["zaboi"] if d > 0 else 0.0
        H = d * p["kis"]                       # высота уступа
        I = d - H                              # перебур
        K = round(N * raskh_sel, 2)            # АС, кг
        L = K * p["doldt"] if K > 0 else 0.0   # ДТ, кг
        M = L / plotdt if plotdt > 0 else 0.0  # ДТ, л
        E = K + L                              # гранулит, кг
        sumD += d; sumH += H; sumN += N
        sumK += K; sumL += L; sumM += M; sumI += I
        rows.append(dict(ryad=w["ryad"], skv=w["skv"], d=d, N=N, zaboi=zaboi,
                         H=H, I=I, K=K, L=L, M=M, E=E))

    sr_glub = sumD / n if n else 0
    sr_hust = sumH / n if n else 0
    sr_pereb = sumI / n if n else 0
    sr_zar = sumN / n if n else 0
    sr_zaboi = p["zaboi"]
    objem_massiva = p["ploshad"] * sr_hust

    for r in rows:
        r["J"] = objem_massiva * r["H"] / sumH if sumH > 0 else 0

    mas_as = sumK
    mas_dt_kg = sumL
    mas_dt_l = sumM
    mas_pt = n * p["boevik"]
    obsh_massa = mas_as + mas_dt_kg + mas_pt
    sr_massa = (mas_as / n if n else 0) + (mas_dt_kg / n if n else 0) + p["boevik"]

    plotvv_kg = p["plotvv"] * 1000
    lspp = round(53 * p["ktresh"] * p["diam"] *
                 math.sqrt((plotvv_kg - p["krab"]) / p["plotpor"]), 1)
    ksbl = round(0.55 / (p["diam"] ** (1 / 3)), 2)
    s_per_skv = p["a"] * b
    vyhod_gm = objem_massiva / max(sumD, 1e-9)
    ud_vv = obsh_massa / objem_massiva if objem_massiva > 0 else 0
    ud_as = mas_as / objem_massiva if objem_massiva > 0 else 0

    iskra_s = n
    iskra42 = min(max(int(p["iskra42"]), 0), n)
    iskra67 = n - iskra42
    iskraV = p["iskraV"]

    # Безопасные расстояния
    nz = sr_zar / sr_glub if sr_glub > 0 else 0
    h9 = round(math.sqrt(4 * p["diam"] / (2 * b)), 2) if b > 0 else 0
    r_raz = 1250 * nz * h9
    kr = round(0.5 * (1 + math.sqrt(1 + 4 * p["prev"] / r_raz)), 2) if r_raz > 0 else 1
    R_raz = r_raz * kr
    r_uvv = 10 * (obsh_massa ** (1 / 3)) if obsh_massa > 0 else 0
    r_uvv_zd = ceil_to(r_uvv * 1.3, 100)
    f30 = 12 * 1.5 / (2 ** 0.25)
    q_seism = 2 * sr_massa * 0.9 + 1
    r_seism = f30 * (q_seism ** (1 / 3))
    r_raz_v = 1.4 * r_raz
    zona_ludi = max(1000, ceil_to(R_raz, 100))
    zona_obor = max(p["zoob"], ceil_to(r_raz_v, 100))
    l_per_form = 14 * p["diam"] * math.sqrt(p["plotvv"] * 0.8)

    return dict(
        params=p, wells=rows, n=n, b=b, plotdt=plotdt, temp=temp,
        raskh_sel=raskh_sel, sumD=sumD, sumH=sumH, sumN=sumN,
        sumK=sumK, sumL=sumL, sumM=sumM, sumI=sumI,
        sr_glub=sr_glub, sr_hust=sr_hust, sr_pereb=sr_pereb,
        sr_zar=sr_zar, sr_zaboi=sr_zaboi, objem_massiva=objem_massiva,
        mas_as=mas_as, mas_dt_kg=mas_dt_kg, mas_dt_l=mas_dt_l, mas_pt=mas_pt,
        obsh_massa=obsh_massa, sr_massa=sr_massa,
        lspp=lspp, ksbl=ksbl, s_per_skv=s_per_skv, vyhod_gm=vyhod_gm,
        ud_vv=ud_vv, ud_as=ud_as,
        iskra_s=iskra_s, iskra42=iskra42, iskra67=iskra67, iskraV=iskraV,
        nz=nz, h9=h9, r_raz=r_raz, kr=kr, R_raz=R_raz,
        r_uvv=r_uvv, r_uvv_zd=r_uvv_zd, r_seism=r_seism, q_seism=q_seism,
        r_raz_v=r_raz_v, zona_ludi=zona_ludi, zona_obor=zona_obor,
        l_per_form=l_per_form,
    )


def depth_groups(calc):
    """Группировка скважин по глубине для таблицы параметров."""
    g = {}
    for w in calc["wells"]:
        k = round(w["d"], 2)
        if k not in g:
            g[k] = dict(d=w["d"], n=0, sN=0.0, sK=0.0, sL=0.0, sZ=0.0)
        g[k]["n"] += 1
        g[k]["sN"] += w["N"]
        g[k]["sK"] += w["K"]
        g[k]["sL"] += w["L"]
        g[k]["sZ"] += w["zaboi"]
    return sorted(g.values(), key=lambda x: -x["d"])


def parse_charge_card(text):
    """
    Разбирает зарядную карту из текста: каждая строка — ряд,
    числа разделены пробелами/табуляцией/точкой с запятой. Десятичная
    запятая поддерживается. 0 и пустые — нет скважины.
    Возвращает список скважин [{'ryad','skv','d'}].
    """
    wells = []
    for ri, line in enumerate(text.strip().splitlines(), start=1):
        parts = []
        for raw in line.replace(";", " ").split():
            # Если пользователь ввёл 8,8,8 без пробелов, считаем запятые
            # разделителями. Одиночная запятая остаётся десятичной.
            if raw.count(",") > 1 and "." not in raw:
                parts.extend(raw.split(","))
            else:
                parts.append(raw)
        for ci, val in enumerate(parts, start=1):
            try:
                d = float(val.replace(",", "."))
            except ValueError:
                continue
            if d > 0:
                wells.append(dict(ryad=ri, skv=ci, d=d))
    return wells
