import os
import asyncio
import logging
from datetime import datetime
from dotenv import load_dotenv, set_key
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

from aiohttp import web
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    FSInputFile,
    KeyboardButton,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
)

ENV_PATH = os.path.join(os.path.dirname(__file__), ".env")
load_dotenv(ENV_PATH)

BOT_TOKEN = os.getenv("BOT_TOKEN")
ADMIN_ID = os.getenv("ADMIN_ID")
PORT = int(os.getenv("PORT", 8080))

DATA_DIR = os.path.join(os.path.dirname(__file__), "excel_jadvallar")
os.makedirs(DATA_DIR, exist_ok=True)

GENERAL_EXCEL_FILE = os.path.join(DATA_DIR, "barcha_nomzodlar_umumiy.xlsx")

excel_lock = asyncio.Lock()
logging.basicConfig(level=logging.INFO)

# Holatlar (FSM)
class CandidateForm(StatesGroup):
    full_name = State()
    phone = State()
    age = State()
    region = State()
    direction = State()
    experience = State()
    resume = State()


def get_daily_filename():
    today = datetime.now().strftime("%Y-%m-%d")
    return os.path.join(DATA_DIR, f"nomzodlar_{today}.xlsx")


def create_styled_excel(file_path: str, title: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = title

    headers = [
        "№",
        "Sana va vaqt",
        "Telegram ID",
        "Username",
        "F.I.SH",
        "Telefon",
        "Yoshi",
        "Viloyat / Shahar",
        "Yo'nalish",
        "Ijodiy tajriba",
        "Rezyume yuklanganmi"
    ]
    ws.append(headers)

    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    col_widths = [6, 18, 14, 16, 26, 18, 10, 20, 24, 35, 20]
    for idx, width in enumerate(col_widths, 1):
        col_letter = openpyxl.utils.get_column_letter(idx)
        ws.column_dimensions[col_letter].width = width

    wb.save(file_path)


def write_entry_to_file(file_path: str, title: str, data: dict):
    if not os.path.exists(file_path):
        create_styled_excel(file_path, title)

    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    row_num = ws.max_row

    row = [
        row_num,
        datetime.now().strftime("%Y-%m-%d %H:%M"),
        data.get("tg_id", ""),
        data.get("username", ""),
        data.get("full_name", ""),
        data.get("phone", ""),
        data.get("age", ""),
        data.get("region", ""),
        data.get("direction", ""),
        data.get("experience", ""),
        "Ha" if data.get("has_resume") else "Yo'q"
    ]
    ws.append(row)

    current_row = ws.max_row
    ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="center")
    ws.cell(row=current_row, column=2).alignment = Alignment(horizontal="center")
    ws.cell(row=current_row, column=6).alignment = Alignment(horizontal="center")
    ws.cell(row=current_row, column=7).alignment = Alignment(horizontal="center")

    wb.save(file_path)
    return row_num


async def save_candidate_data(data: dict):
    async with excel_lock:
        today_file = get_daily_filename()
        today_str = datetime.now().strftime("%d-%m-%Y")
        today_count = write_entry_to_file(today_file, f"Kunlik_{today_str}", data)
        total_count = write_entry_to_file(GENERAL_EXCEL_FILE, "Barcha_nomzodlar", data)
        return today_count, total_count


def get_stats():
    today_file = get_daily_filename()
    today_count = 0
    total_count = 0

    if os.path.exists(today_file):
        try:
            wb = openpyxl.load_workbook(today_file, read_only=True)
            today_count = max(0, wb.active.max_row - 1)
            wb.close()
        except Exception:
            pass

    if os.path.exists(GENERAL_EXCEL_FILE):
        try:
            wb = openpyxl.load_workbook(GENERAL_EXCEL_FILE, read_only=True)
            total_count = max(0, wb.active.max_row - 1)
            wb.close()
        except Exception:
            pass

    return today_count, total_count


# Render uchun Web Server (Health Check)
async def handle_ping(request):
    return web.Response(text="“Folk Party — Yangi O‘zbekiston” bot is running 24/7!")

async def start_web_server():
    app = web.Application()
    app.router.add_get("/", handle_ping)
    app.router.add_get("/ping", handle_ping)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", PORT)
    await site.start()
    logging.info(f"Render Web Server 0.0.0.0:{PORT} portida ishga tushdi.")


# Klaviaturalar
def get_main_menu_kb(is_admin=False):
    buttons = [
        [KeyboardButton(text="📝 Ariza topshirish (Anketa to'ldirish)")],
        [KeyboardButton(text="ℹ️ Tanlov haqida ma'lumot")]
    ]
    if is_admin:
        buttons.append([KeyboardButton(text="🛠 Admin paneli")])
    return ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True)

def get_cancel_kb():
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="❌ Bekor qilish")]],
        resize_keyboard=True
    )

def get_phone_kb():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📱 Raqamni ulashish", request_contact=True)],
            [KeyboardButton(text="❌ Bekor qilish")]
        ],
        resize_keyboard=True
    )

def get_resume_kb():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="⏭ O'tkazib yuborish")],
            [KeyboardButton(text="❌ Bekor qilish")]
        ],
        resize_keyboard=True
    )

def get_admin_kb():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📥 Bugungi kunlik Excelni olish")],
            [KeyboardButton(text="📚 Barcha kunlar (Umumiy Excel)")],
            [KeyboardButton(text="📊 Statistika (Bugungi va Jami)")],
            [KeyboardButton(text="🔙 Asosiy menyu")]
        ],
        resize_keyboard=True
    )


dp = Dispatcher(storage=MemoryStorage())


@dp.message(F.text == "❌ Bekor qilish")
async def cancel_handler(message: types.Message, state: FSMContext):
    await state.clear()
    is_admin = bool(ADMIN_ID and str(message.from_user.id) == str(ADMIN_ID))
    await message.answer(
        "Jarayon bekor qilindi. Kerakli bo'limni tanlang:",
        reply_markup=get_main_menu_kb(is_admin)
    )


@dp.message(CommandStart())
@dp.message(F.text == "🔙 Asosiy menyu")
async def start_handler(message: types.Message, state: FSMContext):
    await state.clear()
    is_admin = bool(ADMIN_ID and str(message.from_user.id) == str(ADMIN_ID))

    welcome_text = (
        f"Assalomu alaykum, <b>{message.from_user.full_name}</b>!\n\n"
        "🎭 <b>“Folk Party — Yangi O‘zbekiston” Respublika tanlovining rasmiy botiga xush kelibsiz!</b>\n\n"
        "Ushbu bot orqali siz tanlovda ishtirok etish uchun o‘z arizangizni topshirishingiz, rezyume va ijodiy ma'lumotlaringizni yuborishingiz mumkin.\n\n"
        "📋 <b>Ariza topshirish bosqichlari:</b>\n"
        "1️⃣ To'liq ism-familiyangiz\n"
        "2️⃣ Telefon raqamingiz\n"
        "3️⃣ Yoshingiz\n"
        "4️⃣ Yashash viloyatingiz / shahringiz\n"
        "5️⃣ Tanlov yo'nalishi\n"
        "6️⃣ Ijodiy tajribangiz va faoliyatingiz\n"
        "7️⃣ Rezyume / CV yoki ijodiy namuna fayli\n\n"
        "Boshlash uchun pastdagi <b>«📝 Ariza topshirish (Anketa to'ldirish)»</b> tugmasini bosing:"
    )
    await message.answer(welcome_text, parse_mode="HTML", reply_markup=get_main_menu_kb(is_admin))


@dp.message(F.text == "ℹ️ Tanlov haqida ma'lumot")
async def info_handler(message: types.Message):
    is_admin = bool(ADMIN_ID and str(message.from_user.id) == str(ADMIN_ID))
    info_text = (
        "ℹ️ <b>“Folk Party — Yangi O‘zbekiston” Respublika tanlovi haqida:</b>\n\n"
        "Ushbu tanlov yurtimizdagi iqtidorli ijodkorlarni kashf etish, milliy va xalq og'zaki ijodini keng targ'ib qilish maqsadida o'tkazilmoqda.\n\n"
        "🏆 <b>Kimlar ishtirok etishi mumkin?</b>\n"
        "Folklor, xalq cholg'ulari, vokal, raqs va boshqa san'at yo'nalishlarida ijod qilayotgan barcha yurtdoshlarimiz qatnashishlari mumkin.\n\n"
        "Ariza topshirish uchun <b>«📝 Ariza topshirish»</b> tugmasini bosing!"
    )
    await message.answer(info_text, parse_mode="HTML", reply_markup=get_main_menu_kb(is_admin))


@dp.message(F.text == "📝 Ariza topshirish (Anketa to'ldirish)")
async def start_form(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "1️⃣ <b>Ism, familiya va otangizning ismini to'liq kiriting:</b>\n"
        "(Masalan: Abdullayev Dilshod Rustam o'g'li)",
        parse_mode="HTML",
        reply_markup=get_cancel_kb()
    )
    await state.set_state(CandidateForm.full_name)


@dp.message(CandidateForm.full_name)
async def process_full_name(message: types.Message, state: FSMContext):
    if len(message.text.strip()) < 3:
        await message.answer("Iltimos, ism va familiyangizni to'liq kiriting:")
        return

    await state.update_data(full_name=message.text.strip())
    await message.answer(
        "2️⃣ <b>Bog'lanish uchun telefon raqamingizni yuboring:</b>\n\n"
        "Pastdagi «📱 Raqamni ulashish» tugmasini bosing yoki yozing (+998901234567):",
        parse_mode="HTML",
        reply_markup=get_phone_kb()
    )
    await state.set_state(CandidateForm.phone)


@dp.message(CandidateForm.phone, F.contact | F.text)
async def process_phone(message: types.Message, state: FSMContext):
    phone_number = message.contact.phone_number if message.contact else message.text.strip()
    await state.update_data(phone=phone_number)

    await message.answer(
        "3️⃣ <b>Yoshingizni kiriting:</b>\n(Masalan: 22)",
        parse_mode="HTML",
        reply_markup=get_cancel_kb()
    )
    await state.set_state(CandidateForm.age)


@dp.message(CandidateForm.age)
async def process_age(message: types.Message, state: FSMContext):
    age_text = message.text.strip()
    if not age_text.isdigit() or int(age_text) < 5 or int(age_text) > 90:
        await message.answer("Iltimos, yoshingizni raqamda to'g'ri kiriting (masalan: 22):")
        return

    await state.update_data(age=age_text)
    await message.answer(
        "4️⃣ <b>Qaysi viloyat / shahardan ishtirok etmoqdasiz?</b>\n"
        "(Masalan: Toshkent shahri, Samarqand, Farg'ona, Xorazm...)",
        parse_mode="HTML",
        reply_markup=get_cancel_kb()
    )
    await state.set_state(CandidateForm.region)


@dp.message(CandidateForm.region)
async def process_region(message: types.Message, state: FSMContext):
    await state.update_data(region=message.text.strip())
    await message.answer(
        "5️⃣ <b>Tanlovning qaysi yo'nalishi bo'yicha qatnashyapsiz?</b>\n"
        "(Masalan: Folklor jamoasi, Vokal / Milliy xonandalik, Xalq cholg'ulari, Milliy raqs...)",
        parse_mode="HTML",
        reply_markup=get_cancel_kb()
    )
    await state.set_state(CandidateForm.direction)


@dp.message(CandidateForm.direction)
async def process_direction(message: types.Message, state: FSMContext):
    await state.update_data(direction=message.text.strip())
    await message.answer(
        "6️⃣ <b>Ijodiy faoliyatingiz, yutuqlaringiz yoki ish tajribangiz haqida qisqacha yozing:</b>\n"
        "(Masalan: Musiqa maktabi o'quvchisi / bitiruvchisi, Respublika tanlovlari sovrindori...)",
        parse_mode="HTML",
        reply_markup=get_cancel_kb()
    )
    await state.set_state(CandidateForm.experience)


@dp.message(CandidateForm.experience)
async def process_experience(message: types.Message, state: FSMContext):
    await state.update_data(experience=message.text.strip())
    await message.answer(
        "7️⃣ <b>Rezyume / CV yoki ijodiy ishingiz namunasi (PDF, Word, rasm) bo'lsa yuboring:</b>\n\n"
        "Agar hozircha faylingiz bo'lmasa, «⏭ O'tkazib yuborish» tugmasini bosing.",
        parse_mode="HTML",
        reply_markup=get_resume_kb()
    )
    await state.set_state(CandidateForm.resume)


@dp.message(CandidateForm.resume, F.document | F.photo | (F.text == "⏭ O'tkazib yuborish"))
async def process_resume(message: types.Message, state: FSMContext, bot: Bot):
    user_data = await state.get_data()
    has_resume = False
    file_id = None
    file_type = None

    if message.document:
        has_resume = True
        file_id = message.document.file_id
        file_type = "document"
    elif message.photo:
        has_resume = True
        file_id = message.photo[-1].file_id
        file_type = "photo"

    user_data["tg_id"] = message.from_user.id
    user_data["username"] = f"@{message.from_user.username}" if message.from_user.username else "Yo'q"
    user_data["has_resume"] = has_resume

    today_num = 1
    total_num = 1
    try:
        today_num, total_num = await save_candidate_data(user_data)
    except Exception as e:
        logging.error(f"Excelga saqlashda xatolik: {e}")

    is_admin = bool(ADMIN_ID and str(message.from_user.id) == str(ADMIN_ID))

    await message.answer(
        "🎉 <b>Tabriklaymiz! Sizning arizangiz muvaffaqiyatli qabul qilindi.</b>\n\n"
        "“Folk Party — Yangi O‘zbekiston” tanlovi tashkiliy qo'mitasi arizangizni ko'rib chiqib, ko'rsatilgan telefon raqamingiz orqali bog'lanadi. Omad tilaymiz!",
        parse_mode="HTML",
        reply_markup=get_main_menu_kb(is_admin)
    )
    await state.clear()

    if ADMIN_ID and str(ADMIN_ID).isdigit():
        admin_text = (
            f"🔔 <b>YANGI NOMZOD KELIB TUSHDI!</b>\n"
            f"📊 <i>(Bugungi: #{today_num} | Jami: #{total_num})</i>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"👤 <b>F.I.SH:</b> {user_data.get('full_name')}\n"
            f"📞 <b>Telefon:</b> {user_data.get('phone')}\n"
            f"🎂 <b>Yoshi:</b> {user_data.get('age')}\n"
            f"📍 <b>Viloyat/Shahar:</b> {user_data.get('region')}\n"
            f"🎭 <b>Yo'nalish:</b> {user_data.get('direction')}\n"
            f"📝 <b>Ijodiy tajriba:</b> {user_data.get('experience')}\n"
            f"💬 <b>Telegram:</b> {user_data.get('username')} (ID: {user_data.get('tg_id')})\n"
            f"📎 <b>Fayl:</b> {'Biriktirilgan (pastda)' if has_resume else 'Mavjud emas'}\n"
            f"🕒 <b>Vaqt:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        )

        try:
            await bot.send_message(chat_id=int(ADMIN_ID), text=admin_text, parse_mode="HTML")
            if file_id:
                if file_type == "document":
                    await bot.send_document(
                        chat_id=int(ADMIN_ID),
                        document=file_id,
                        caption=f"📎 {user_data.get('full_name')} ning fayli / rezyumesi"
                    )
                elif file_type == "photo":
                    await bot.send_photo(
                        chat_id=int(ADMIN_ID),
                        photo=file_id,
                        caption=f"📎 {user_data.get('full_name')} ning fotosurati"
                    )
        except Exception as e:
            logging.error(f"Adminga yuborishda xatolik: {e}")


# ================= ADMIN BO'LIMI =================

@dp.message(F.text == "🛠 Admin paneli")
@dp.message(Command("admin"))
async def admin_panel(message: types.Message):
    global ADMIN_ID
    if not ADMIN_ID:
        ADMIN_ID = str(message.from_user.id)
        try:
            set_key(ENV_PATH, "ADMIN_ID", ADMIN_ID)
        except Exception:
            pass

    if str(message.from_user.id) != str(ADMIN_ID):
        await message.answer("Kechirasiz, siz ushbu botning admini emassiz.")
        return

    today_count, total_count = get_stats()
    text = (
        "👋 <b>Admin paneliga xush kelibsiz!</b>\n\n"
        f"📅 <b>Bugun ro'yxatdan o'tganlar:</b> {today_count} ta\n"
        f"👥 <b>Jami ro'yxatdan o'tganlar:</b> {total_count} ta\n\n"
        "Quyidagi tugmalar orqali kerakli Excel jadvalini yuklab olishingiz mumkin:"
    )
    await message.answer(text, parse_mode="HTML", reply_markup=get_admin_kb())


@dp.message(F.text == "📥 Bugungi kunlik Excelni olish")
async def export_today_excel(message: types.Message):
    if not ADMIN_ID or str(message.from_user.id) != str(ADMIN_ID):
        return

    today_file = get_daily_filename()
    today_str = datetime.now().strftime("%Y-%m-%d")

    if os.path.exists(today_file):
        file = FSInputFile(today_file, filename=f"nomzodlar_{today_str}.xlsx")
        today_count, _ = get_stats()
        await message.answer_document(
            document=file,
            caption=f"📅 <b>Bugungi ({today_str}) nomzodlar ro'yxati</b>\nJami: <b>{today_count} ta</b> ishtirokchi",
            parse_mode="HTML"
        )
    else:
        await message.answer(f"Bugun ({today_str}) hali hech kim ro'yxatdan o'tmadi.")


@dp.message(F.text == "📚 Barcha kunlar (Umumiy Excel)")
async def export_general_excel(message: types.Message):
    if not ADMIN_ID or str(message.from_user.id) != str(ADMIN_ID):
        return

    if os.path.exists(GENERAL_EXCEL_FILE):
        file = FSInputFile(GENERAL_EXCEL_FILE, filename=f"barcha_nomzodlar_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
        _, total_count = get_stats()
        await message.answer_document(
            document=file,
            caption=f"📚 <b>Barcha kunlar davomidagi umumiy ro'yxat</b>\nJami: <b>{total_count} ta</b> ishtirokchi",
            parse_mode="HTML"
        )
    else:
        await message.answer("Hozircha hech qanday ma'lumot mavjud emas.")


@dp.message(F.text == "📊 Statistika (Bugungi va Jami)")
async def show_stats(message: types.Message):
    if not ADMIN_ID or str(message.from_user.id) != str(ADMIN_ID):
        return

    today_count, total_count = get_stats()
    today_str = datetime.now().strftime("%d.%m.%Y")
    text = (
        "📊 <b>Tanlov statistikasi:</b>\n"
        "━━━━━━━━━━━━━━━━━━━━\n"
        f"📅 <b>Bugun ({today_str}):</b> {today_count} ta nomzod\n"
        f"👥 <b>Barcha kunlar bo'yicha jami:</b> {total_count} ta nomzod\n"
    )
    await message.answer(text, parse_mode="HTML")


async def main():
    # Render uchun Web Serverni fon rejimida yoqish
    await start_web_server()

    bot = Bot(token=BOT_TOKEN)

    bot_description = (
        "“Folk Party — Yangi O‘zbekiston” Respublika tanlovining ro‘yxatdan o‘tish botiga xush kelibsiz!\n\n"
        "Bu yerda tanlov shartlari bilan tanishishingiz va ishtirok etish uchun ariza yuborishingiz mumkin.\n\n"
        "Boshlash uchun START tugmasini bosing."
    )
    bot_short_description = "“Folk Party — Yangi O‘zbekiston” tanloviga ro‘yxatdan o‘tish boti"

    try:
        await bot.set_my_description(description=bot_description)
        await bot.set_my_short_description(short_description=bot_short_description)
    except Exception as e:
        logging.warning(f"Tavsifni yangilashda xato: {e}")

    print("Render-tayyor bot ishga tushdi...")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
