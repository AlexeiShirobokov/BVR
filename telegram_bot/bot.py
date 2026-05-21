# -*- coding: utf-8 -*-
"""
Telegram-бот расчёта буровзрывных работ (БВР).
Основной режим: Telegram Mini App с формой → загрузка паспорта БВР (PDF) →
расчёт → формирование проекта массового взрыва → подтверждение →
выдача Excel и PDF.

Запуск:
    export BVR_BOT_TOKEN="токен_от_BotFather"
    export BVR_WEBAPP_URL="https://.../telegram_webapp/index.html"
    python bot.py
"""
import os
import asyncio
import datetime
import json
import logging
import tempfile

from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (Message, CallbackQuery, FSInputFile,
                           InlineKeyboardButton, InlineKeyboardMarkup,
                           ReplyKeyboardMarkup, KeyboardButton,
                           ReplyKeyboardRemove, WebAppInfo)

from bvr_calc import make_params, calculate, parse_charge_card, DEFAULTS
from bvr_document import build_excel, build_pdf, fmt

logging.basicConfig(level=logging.INFO)

TOKEN = os.environ.get("BVR_BOT_TOKEN", "").strip()
WEBAPP_URL = os.environ.get("BVR_WEBAPP_URL", "").strip()

# Папка для временных файлов проекта
WORK_DIR = os.path.join(tempfile.gettempdir(), "bvr_bot")
os.makedirs(WORK_DIR, exist_ok=True)

bot = Bot(TOKEN) if TOKEN else None
dp = Dispatcher(storage=MemoryStorage())


WEBAPP_KEYS = {
    "p": "pred", "m": "mest", "b": "blok", "pr": "prisk", "np": "np",
    "d": "data", "z1": "zar1", "z2": "zar2", "v1": "vz1", "v2": "vz2",
    "pk": "prikaz", "rp": "raspol", "pv": "plotvv", "vd": "viddt",
    "dm": "diam", "ks": "kis", "kr": "krab", "pp": "plotpor",
    "kt": "ktresh", "pe": "prev", "a": "a", "zb": "zaboi", "rd": "ryad",
    "pl": "ploshad", "bo": "boevik", "dd": "doldt", "i42": "iskra42",
    "iv": "iskraV", "zo": "zoob", "zs": "zoso", "ob": "obj",
    "me": "meri", "gd": "gendir", "gi": "gling", "nb": "nachbvr",
    "gg": "glgeo", "gm": "glmark", "nu": "nachuch", "vr": "vzryvnik",
}
WEBAPP_NUMERIC = {
    "plotvv", "diam", "kis", "krab", "plotpor", "ktresh", "prev", "a",
    "zaboi", "ryad", "ploshad", "boevik", "doldt", "iskra42", "iskraV",
    "zoob", "zoso",
}
WEBAPP_INT = {"ryad", "iskra42", "iskraV", "zoob", "zoso"}


# ==================== СОСТОЯНИЯ ====================
class BVR(StatesGroup):
    mest = State()       # месторождение
    blok = State()       # блок
    data = State()       # дата взрыва
    diam = State()       # диаметр скважин
    kis = State()        # КИС
    a = State()          # расстояние между скважинами
    ryad = State()       # количество рядов
    ploshad = State()    # площадь массива
    raspol = State()     # расположение скважин
    card = State()       # зарядная карта
    passport = State()   # паспорт PDF
    confirm = State()    # подтверждение


# Описание шагов ввода: state, текст вопроса, ключ параметра, тип
STEPS = [
    (BVR.mest,    "Месторождение", "mest", "str"),
    (BVR.blok,    "Блок", "blok", "str"),
    (BVR.data,    "Дата проведения взрыва (ДД.ММ.ГГГГ)", "data", "date"),
    (BVR.diam,    "Условный диаметр скважин, м", "diam", "float"),
    (BVR.kis,     "Коэффициент использования скважины (КИС)", "kis", "float"),
    (BVR.a,       "Расстояние между скважинами, м", "a", "float"),
    (BVR.ryad,    "Количество рядов, шт", "ryad", "int"),
    (BVR.ploshad, "Площадь взрываемого массива, м²", "ploshad", "float"),
    (BVR.raspol,  "Расположение скважин", "raspol", "raspol"),
]


def step_index(state_name):
    for i, (st, *_) in enumerate(STEPS):
        if st.state == state_name:
            return i
    return -1


def def_text(key):
    d = DEFAULTS.get(key)
    if key == "data":
        return datetime.date.today().strftime("%d.%m.%Y")
    return str(d)


async def ask_step(message, idx):
    """Задаёт вопрос шага idx."""
    st, question, key, typ = STEPS[idx]
    if typ == "raspol":
        kb = ReplyKeyboardMarkup(
            keyboard=[[KeyboardButton(text="Шахматное"),
                       KeyboardButton(text="Прямоугольное")]],
            resize_keyboard=True, one_time_keyboard=True)
        await message.answer(
            f"Шаг {idx+1}/{len(STEPS)}. {question}:", reply_markup=kb)
    else:
        await message.answer(
            f"Шаг {idx+1}/{len(STEPS)}. {question}.\n"
            f"Текущее значение: <b>{def_text(key)}</b>\n"
            f"Отправьте новое значение или «-», чтобы оставить текущее.",
            parse_mode="HTML", reply_markup=ReplyKeyboardRemove())


def parse_value(typ, text):
    """Разбирает введённое значение по типу. Возвращает (значение, ошибка)."""
    text = text.strip()
    if typ == "str":
        return text, None
    if typ == "date":
        for f in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d"):
            try:
                return datetime.datetime.strptime(text, f).date(), None
            except ValueError:
                continue
        return None, "Неверный формат даты. Пример: 21.05.2026"
    if typ == "float":
        try:
            return float(text.replace(",", ".")), None
        except ValueError:
            return None, "Введите число, например 0.215"
    if typ == "int":
        try:
            return int(float(text.replace(",", "."))), None
        except ValueError:
            return None, "Введите целое число"
    if typ == "raspol":
        if text in ("Шахматное", "Прямоугольное"):
            return text, None
        return None, "Выберите «Шахматное» или «Прямоугольное»"
    return text, None


def decode_webapp_payload(raw):
    """Разбирает данные, пришедшие из Telegram Mini App."""
    try:
        payload = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValueError("данные формы повреждены") from exc
    if payload.get("type") != "bvr_form":
        raise ValueError("неизвестный тип данных формы")

    over = {}
    for short_key, value in (payload.get("over") or {}).items():
        key = WEBAPP_KEYS.get(short_key)
        if not key:
            continue
        if key == "data":
            try:
                over[key] = datetime.date.fromisoformat(str(value))
            except ValueError as exc:
                raise ValueError("неверная дата в форме") from exc
        elif key in WEBAPP_NUMERIC:
            try:
                num = float(str(value).replace(",", "."))
            except ValueError as exc:
                raise ValueError(f"неверное число в поле {key}") from exc
            over[key] = int(num) if key in WEBAPP_INT else num
        else:
            over[key] = str(value).strip()

    card = str(payload.get("card") or "").strip()
    wells = parse_charge_card(card)
    if not wells:
        raise ValueError("зарядная карта не содержит скважин")
    return over, wells


async def start_dialog(message: Message, state: FSMContext):
    """Запускает старый пошаговый чатовый ввод как запасной режим."""
    await state.clear()
    await state.update_data(over={})
    await message.answer(
        "🔷 <b>Расчёт буровзрывных работ (БВР)</b>\n\n"
        "Запасной чатовый режим: бот пошагово запросит параметры, затем "
        "паспорт БВР и сформирует проект в форматах Excel и PDF.\n\n"
        "На любом шаге можно отправить «-», чтобы оставить значение "
        "по умолчанию.",
        parse_mode="HTML", reply_markup=ReplyKeyboardRemove())
    await state.set_state(STEPS[0][0])
    await ask_step(message, 0)


async def send_webapp_entry(message: Message, state: FSMContext):
    """Показывает кнопку открытия Telegram Mini App."""
    await state.clear()
    if not WEBAPP_URL:
        await message.answer(
            "Адрес мини-приложения не задан. Установите переменную "
            "BVR_WEBAPP_URL с HTTPS-ссылкой на telegram_webapp/index.html.\n\n"
            "Пока могу запустить запасной чатовый ввод командой /dialog.")
        return
    kb = ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="Открыть форму БВР",
                            web_app=WebAppInfo(url=WEBAPP_URL))],
            [KeyboardButton(text="Чатовый ввод")],
        ],
        resize_keyboard=True)
    await message.answer(
        "Откройте форму БВР как мини-приложение Telegram. "
        "После отправки формы бот попросит загрузить паспорт PDF и "
        "сформирует Excel/PDF проекта.",
        reply_markup=kb)


# ==================== ОБРАБОТЧИКИ ====================
@dp.message(Command("start"))
async def cmd_start(message: Message, state: FSMContext):
    if WEBAPP_URL:
        await send_webapp_entry(message, state)
        return
    await start_dialog(message, state)


@dp.message(Command("form"))
async def cmd_form(message: Message, state: FSMContext):
    await send_webapp_entry(message, state)


@dp.message(Command("dialog"))
@dp.message(F.text == "Чатовый ввод")
async def cmd_dialog(message: Message, state: FSMContext):
    await start_dialog(message, state)


@dp.message(F.web_app_data)
async def handle_webapp_data(message: Message, state: FSMContext):
    try:
        over, wells = decode_webapp_payload(message.web_app_data.data)
    except ValueError as exc:
        await message.answer(
            f"Не удалось принять данные формы: {exc}. "
            "Откройте форму заново или используйте /dialog.")
        return

    await state.clear()
    await state.update_data(over=over, wells=wells)
    params = make_params(over)
    total = sum(w["d"] for w in wells)
    await message.answer(
        f"Данные из формы приняты.\n"
        f"Месторождение: <b>{params['mest']}</b>\n"
        f"Блок: <b>{params['blok']}</b>\n"
        f"Скважин: <b>{len(wells)}</b> · Объём бурения: "
        f"<b>{fmt(total,1)}</b> п.м\n\n"
        f"Теперь загрузите паспорт БВР в формате PDF либо отправьте «-», "
        f"чтобы пропустить.",
        parse_mode="HTML", reply_markup=ReplyKeyboardRemove())
    await state.set_state(BVR.passport)

@dp.message(Command("help"))
async def cmd_help(message: Message):
    await message.answer(
        "ℹ️ <b>Справка</b>\n\n"
        "Основной режим — форма Telegram Mini App: /start или /form.\n"
        "В форме заполняются параметры и зарядная карта, после отправки "
        "бот попросит паспорт БВР в PDF.\n\n"
        "Запасной режим — чатовый ввод: /dialog.\n\n"
        "Итог — проект массового взрыва в Excel и PDF. "
        "/cancel — отмена.",
        parse_mode="HTML")


@dp.message(Command("cancel"))
async def cmd_cancel(message: Message, state: FSMContext):
    await state.clear()
    await message.answer("Отменено. /start — начать заново.",
                         reply_markup=ReplyKeyboardRemove())


@dp.message(F.text, BVR.mest)
@dp.message(F.text, BVR.blok)
@dp.message(F.text, BVR.data)
@dp.message(F.text, BVR.diam)
@dp.message(F.text, BVR.kis)
@dp.message(F.text, BVR.a)
@dp.message(F.text, BVR.ryad)
@dp.message(F.text, BVR.ploshad)
@dp.message(F.text, BVR.raspol)
async def handle_step(message: Message, state: FSMContext):
    cur = await state.get_state()
    idx = step_index(cur)
    st, question, key, typ = STEPS[idx]
    data = await state.get_data()
    over = data.get("over", {})
    text = message.text.strip()

    if text != "-":
        val, err = parse_value(typ, text)
        if err:
            await message.answer("⚠️ " + err + ". Повторите ввод.")
            return
        over[key] = val
    await state.update_data(over=over)

    if idx + 1 < len(STEPS):
        await state.set_state(STEPS[idx + 1][0])
        await ask_step(message, idx + 1)
    else:
        await state.set_state(BVR.card)
        await message.answer(
            "📊 <b>Зарядная карта</b>\n\n"
            "Введите проектную глубину скважин. Каждая строка — ряд, "
            "глубины через пробел. Например:\n\n"
            "<code>0 4 4 4 4 0\n8 8 8 8 8 8\n8 8 8 8 8 0</code>\n\n"
            "0 или пропуск — скважина отсутствует.",
            parse_mode="HTML", reply_markup=ReplyKeyboardRemove())


@dp.message(F.text, BVR.card)
async def handle_card(message: Message, state: FSMContext):
    wells = parse_charge_card(message.text)
    if not wells:
        await message.answer("⚠️ Скважины не распознаны. Введите числа "
                             "построчно, например: 8 8 8 8")
        return
    await state.update_data(wells=wells)
    total = sum(w["d"] for w in wells)
    await message.answer(
        f"✅ Зарядная карта принята.\n"
        f"Скважин: <b>{len(wells)}</b> · Объём бурения: "
        f"<b>{fmt(total,1)}</b> п.м\n\n"
        f"📄 Загрузите паспорт БВР в формате PDF либо отправьте «-», "
        f"чтобы пропустить.",
        parse_mode="HTML")
    await state.set_state(BVR.passport)


@dp.message(BVR.passport, F.document)
async def handle_passport(message: Message, state: FSMContext):
    doc = message.document
    if not (doc.file_name or "").lower().endswith(".pdf"):
        await message.answer("⚠️ Нужен файл PDF. Повторите или отправьте «-».")
        return
    path = os.path.join(WORK_DIR, f"passport_{message.from_user.id}.pdf")
    await bot.download(doc, destination=path)
    await state.update_data(passport=path)
    await message.answer(f"✅ Паспорт БВР загружен: {doc.file_name}")
    await show_summary(message, state)


@dp.message(BVR.passport, F.text)
async def handle_passport_skip(message: Message, state: FSMContext):
    if message.text.strip() == "-":
        await state.update_data(passport=None)
        await message.answer("Паспорт пропущен.")
        await show_summary(message, state)
    else:
        await message.answer("Отправьте PDF-файл паспорта или «-».")


async def show_summary(message: Message, state: FSMContext):
    data = await state.get_data()
    params = make_params(data.get("over", {}))
    wells = data.get("wells", [])
    calc = calculate(params, wells)
    await state.update_data(calc_ready=True)

    txt = (
        f"📋 <b>Результаты расчёта БВР</b>\n\n"
        f"Месторождение: {params['mest']}\n"
        f"Блок: {params['blok']}\n"
        f"Количество скважин: <b>{calc['n']}</b> шт\n"
        f"Объём бурения: <b>{fmt(calc['sumD'],0)}</b> п.м\n"
        f"Объём взрыва: <b>{fmt(calc['objem_massiva']/1000,1)}</b> тыс.м³\n"
        f"ЛСПП: <b>{fmt(calc['lspp'],1)}</b> м\n"
        f"Сеть бурения: {fmt(params['a'],1)} × {fmt(calc['b'],1)} м\n\n"
        f"<b>Взрывчатые вещества:</b>\n"
        f"• Аммиачная селитра: {fmt(calc['mas_as'],0)} кг\n"
        f"• Дизтопливо: {fmt(calc['mas_dt_kg'],0)} кг\n"
        f"• ПТ-П-2250: {fmt(calc['mas_pt'],0)} кг\n"
        f"• Общая масса заряда: <b>{fmt(calc['obsh_massa']/1000,2)}</b> т\n"
        f"• Удельный расход ВВ: {fmt(calc['ud_vv'],3)} кг/м³\n\n"
        f"<b>Средства инициирования:</b>\n"
        f"• Искра-С: {calc['iskra_s']} · Искра-П(67): {calc['iskra67']} · "
        f"Искра-П(42): {calc['iskra42']} · Искра-В: {calc['iskraV']}\n\n"
        f"<b>Опасная зона:</b> люди — {calc['zona_ludi']} м, "
        f"оборудование — {calc['zona_obor']} м\n\n"
        f"Подтвердите формирование проекта массового взрыва."
    )
    kb = InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text="✅ Подтвердить и сформировать",
                             callback_data="confirm"),
        InlineKeyboardButton(text="✖️ Отмена", callback_data="cancel"),
    ]])
    await message.answer(txt, parse_mode="HTML", reply_markup=kb)
    await state.set_state(BVR.confirm)


@dp.callback_query(F.data == "cancel")
async def cb_cancel(cq: CallbackQuery, state: FSMContext):
    await state.clear()
    await cq.message.edit_reply_markup(reply_markup=None)
    await cq.message.answer("Отменено. /start — начать заново.")
    await cq.answer()


@dp.callback_query(F.data == "confirm", BVR.confirm)
async def cb_confirm(cq: CallbackQuery, state: FSMContext):
    await cq.answer("Формирую проект…")
    await cq.message.edit_reply_markup(reply_markup=None)
    data = await state.get_data()
    params = make_params(data.get("over", {}))
    wells = data.get("wells", [])
    passport = data.get("passport")
    calc = calculate(params, wells)

    uid = cq.from_user.id
    safe_blok = "".join(c if c.isalnum() else "_" for c in params["blok"])
    xlsx_path = os.path.join(WORK_DIR, f"Проект_БВР_{safe_blok}_{uid}.xlsx")
    pdf_path = os.path.join(WORK_DIR, f"Проект_БВР_{safe_blok}_{uid}.pdf")

    await cq.message.answer("⏳ Формирование проекта массового взрыва…")
    try:
        build_excel(calc, xlsx_path)
        build_pdf(calc, pdf_path, passport_path=passport)
    except Exception as e:
        logging.exception("Ошибка формирования")
        await cq.message.answer(f"❌ Ошибка формирования проекта: {e}")
        return

    await cq.message.answer_document(
        FSInputFile(xlsx_path, filename=f"Проект_БВР_{params['blok']}.xlsx"),
        caption="📊 Расчёт БВР и проект — Excel")
    await cq.message.answer_document(
        FSInputFile(pdf_path, filename=f"Проект_БВР_{params['blok']}.pdf"),
        caption="📄 Проект массового взрыва — PDF (готов к печати)")
    if passport:
        await cq.message.answer_document(
            FSInputFile(passport, filename="Паспорт_БВР.pdf"),
            caption="📎 Паспорт БВР (приложение)")

    await cq.message.answer(
        "✅ <b>Проект массового взрыва сформирован.</b>\n\n"
        "Документ готов к печати и подписанию. "
        "/start — новый расчёт.",
        parse_mode="HTML")
    await state.clear()


@dp.message()
async def fallback(message: Message, state: FSMContext):
    cur = await state.get_state()
    if cur is None:
        await message.answer("Отправьте /start, чтобы начать расчёт БВР.")
    else:
        await message.answer("Ожидается другой ввод. /help — справка, "
                             "/cancel — отмена.")


async def main():
    if not TOKEN:
        print("ОШИБКА: не задан токен бота.\n"
              "Установите переменную окружения BVR_BOT_TOKEN:\n"
              '  export BVR_BOT_TOKEN="токен_от_BotFather"')
        return
    print("Бот расчёта БВР запущен.")
    if WEBAPP_URL:
        print(f"Mini App: {WEBAPP_URL}")
    else:
        print("Mini App не настроен: задайте BVR_WEBAPP_URL для кнопки формы.")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
