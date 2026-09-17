"""
Talaba Yordamchi Bot (Groq + Pexels, Railway uchun tayyor)
-----------------------------------------------------------------
/slayd     - suhbat: mavzu -> slayd sahifalari soni (1-12) -> muallif -> .pptx
/mustaqil  - suhbat: mavzu -> ish sahifalari soni (1-12) -> muallif -> .docx
"""

import os
import json
import re
import random
import traceback
from io import BytesIO

import requests
from groq import Groq
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from docx import Document
from docx.shared import Pt as DocxPt
from docx.enum.text import WD_ALIGN_PARAGRAPH

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
GROQ_API_KEY = os.environ.get("GROQ_API_KEY")
PEXELS_API_KEY = os.environ.get("PEXELS_API_KEY")

if not TELEGRAM_TOKEN or not GROQ_API_KEY:
    raise ValueError(
        "TELEGRAM_TOKEN yoki GROQ_API_KEY topilmadi! "
        "Environment variables to'g'ri sozlanganini tekshiring."
    )

client = Groq(api_key=GROQ_API_KEY)
MODEL = "openai/gpt-oss-120b"

MAX_COUNT = 12

# Suhbat bosqichlari (ikkala buyruq uchun ham umumiy)
MAVZU, SONI, MUALLIF = range(3)

THEMES = [
    {"bg": RGBColor(0x1A, 0x23, 0x3A), "title": RGBColor(0xFF, 0xFF, 0xFF), "text": RGBColor(0xE0, 0xE0, 0xE0), "accent": RGBColor(0x4F, 0xA8, 0xE0)},
    {"bg": RGBColor(0xFF, 0xFF, 0xFF), "title": RGBColor(0x1A, 0x23, 0x3A), "text": RGBColor(0x33, 0x33, 0x33), "accent": RGBColor(0xE0, 0x6A, 0x4F)},
    {"bg": RGBColor(0x2D, 0x2A, 0x4A), "title": RGBColor(0xFF, 0xD9, 0x66), "text": RGBColor(0xF0, 0xF0, 0xF0), "accent": RGBColor(0xFF, 0xD9, 0x66)},
    {"bg": RGBColor(0xF4, 0xF1, 0xEA), "title": RGBColor(0x3A, 0x5A, 0x40), "text": RGBColor(0x2C, 0x2C, 0x2C), "accent": RGBColor(0x3A, 0x5A, 0x40)},
]


def ask_ai(prompt: str, max_tokens: int = 3000) -> str:
    response = client.chat.completions.create(
        model=MODEL,
        max_tokens=max_tokens,
        messages=[{"role": "user", "content": prompt}],
    )
    return response.choices[0].message.content


def extract_json(text: str) -> dict:
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if not match:
        raise ValueError("AI javobida JSON topilmadi")
    return json.loads(match.group(0))


def get_image(query: str):
    """Pexels'dan mavzuga mos rasm oladi. Muammo bo'lsa, sababini logga chiqarib None qaytaradi."""
    if not PEXELS_API_KEY:
        print("PEXELS_API_KEY topilmadi - rasm o'tkazib yuborildi")
        return None
    try:
        resp = requests.get(
            "https://api.pexels.com/v1/search",
            headers={"Authorization": PEXELS_API_KEY},
            params={"query": query, "per_page": 1, "orientation": "landscape"},
            timeout=15,
        )
        if resp.status_code != 200:
            print(f"Pexels xatosi: status={resp.status_code}, matn={resp.text[:200]}")
            return None
        data = resp.json()
        photos = data.get("photos", [])
        if not photos:
            print(f"Pexels'da '{query}' uchun rasm topilmadi")
            return None
        img_url = photos[0]["src"]["large"]
        img_resp = requests.get(img_url, timeout=15)
        if img_resp.status_code != 200:
            print(f"Rasmni yuklab bo'lmadi: status={img_resp.status_code}")
            return None
        return BytesIO(img_resp.content)
    except Exception:
        print("Pexels/rasm xatosi:")
        traceback.print_exc()
        return None


def add_background(slide, prs, color):
    """Slaydga fon rangi qo'shadi. Birinchi bo'lib qo'shilgani uchun avtomatik eng orqada turadi."""
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.line.fill.background()
    bg.shadow.inherit = False
    return bg


def create_pptx(mavzu: str, muallif: str, slides_data: list) -> BytesIO:
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    theme = random.choice(THEMES)
    blank = prs.slide_layouts[6]

    # --- Titul slayd ---
    slide = prs.slides.add_slide(blank)
    add_background(slide, prs, theme["bg"])

    title_box = slide.shapes.add_textbox(Inches(1), Inches(2.7), Inches(11.33), Inches(1.5))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = mavzu
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = theme["title"]

    author_box = slide.shapes.add_textbox(Inches(1), Inches(4.3), Inches(11.33), Inches(0.8))
    p2 = author_box.text_frame.paragraphs[0]
    p2.text = f"Tayyorladi: {muallif}"
    p2.alignment = PP_ALIGN.CENTER
    p2.font.size = Pt(22)
    p2.font.color.rgb = theme["accent"]

    total = len(slides_data)

    # --- Mazmun slaydlari ---
    for idx, item in enumerate(slides_data, start=1):
        slide = prs.slides.add_slide(blank)
        add_background(slide, prs, theme["bg"])

        image_on_left = idx % 2 == 1
        text_x = Inches(6.8) if image_on_left else Inches(0.7)
        img_x = Inches(0.5) if image_on_left else Inches(7.1)

        title_box = slide.shapes.add_textbox(text_x, Inches(0.5), Inches(6.0), Inches(1.0))
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = item.get("title", "")
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = theme["title"]

        body_box = slide.shapes.add_textbox(text_x, Inches(1.6), Inches(6.0), Inches(5.3))
        tf = body_box.text_frame
        tf.word_wrap = True
        bullets = item.get("bullets", [])
        for i, bullet in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = f"•  {strip_markdown(bullet)}"
            p.font.size = Pt(15)
            p.font.color.rgb = theme["text"]
            p.space_after = Pt(10)

        img_data = get_image(item.get("image_query", mavzu))
        if img_data:
            try:
                slide.shapes.add_picture(img_data, img_x, Inches(1.6), width=Inches(5.7), height=Inches(4.8))
            except Exception:
                print("Rasmni slaydga joylashtirishda xato:")
                traceback.print_exc()

        num_box = slide.shapes.add_textbox(Inches(12.5), Inches(7.05), Inches(0.7), Inches(0.4))
        p = num_box.text_frame.paragraphs[0]
        p.text = f"{idx}/{total}"
        p.font.size = Pt(12)
        p.font.color.rgb = theme["accent"]

    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer


def strip_markdown(text: str) -> str:
    """AI javobidagi **qalin**, *qiya*, # kabi Markdown belgilarini tozalaydi."""
    text = re.sub(r"\*\*(.*?)\*\*", r"\1", text)
    text = re.sub(r"\*(.*?)\*", r"\1", text)
    text = re.sub(r"^#{1,6}\s*", "", text, flags=re.MULTILINE)
    text = re.sub(r"^-\s+", "", text, flags=re.MULTILINE)
    return text


def create_docx(mavzu: str, muallif: str, matn: str) -> BytesIO:
    doc = Document()

    # Titul sahifasi
    for _ in range(6):
        doc.add_paragraph()
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run(mavzu)
    run.font.size = DocxPt(24)
    run.font.bold = True

    doc.add_paragraph()
    author_p = doc.add_paragraph()
    author_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = author_p.add_run(f"Tuzuvchi: {muallif}")
    run2.font.size = DocxPt(14)

    doc.add_page_break()

    # Asosiy matn
    matn = strip_markdown(matn)
    for paragraph in matn.split("\n"):
        paragraph = paragraph.strip()
        if not paragraph:
            continue
        if paragraph.endswith(":") and len(paragraph) < 40:
            doc.add_heading(paragraph, level=2)
        else:
            p = doc.add_paragraph(paragraph)
            for run in p.runs:
                run.font.size = DocxPt(12)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Salom! Men talabalarga yordam beruvchi botman.\n\n"
        "Buyruqlar:\n"
        "/slayd - rasmli, dizaynli PowerPoint (.pptx) tayyorlab beraman\n"
        "/mustaqil - Word (.docx) mustaqil ish tayyorlab beraman\n"
    )


# ===== Umumiy suhbat (slayd va mustaqil ish uchun) =====

async def slayd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["turi"] = "slayd"
    await update.message.reply_text("Taqdimot mavzusini kiriting:")
    return MAVZU


async def mustaqil_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["turi"] = "mustaqil"
    await update.message.reply_text("Mustaqil ish mavzusini kiriting:")
    return MAVZU


async def handle_mavzu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["mavzu"] = update.message.text.strip()
    if context.user_data["turi"] == "slayd":
        await update.message.reply_text(
            f"Slayd sahifalari sonini kiriting (1 dan {MAX_COUNT} gacha):"
        )
    else:
        await update.message.reply_text(
            f"Mustaqil ish sahifalari soni (1 dan {MAX_COUNT} gacha):"
        )
    return SONI


async def handle_soni(update: Update, context: ContextTypes.DEFAULT_TYPE):
    matn = update.message.text.strip()
    if not matn.isdigit() or not (1 <= int(matn) <= MAX_COUNT):
        await update.message.reply_text(f"Iltimos, 1 dan {MAX_COUNT} gacha butun son kiriting:")
        return SONI
    context.user_data["soni"] = int(matn)
    await update.message.reply_text("Tuzuvchi (muallif) ismi va familiyasini kiriting:")
    return MUALLIF


async def handle_muallif(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["muallif"] = update.message.text.strip()
    turi = context.user_data["turi"]
    mavzu = context.user_data["mavzu"]
    soni = context.user_data["soni"]
    muallif = context.user_data["muallif"]

    if turi == "slayd":
        await update.message.reply_text("Taqdimot tayyorlanmoqda, biroz kuting...")
        prompt = (
            f"'{mavzu}' mavzusida chuqur va batafsil taqdimot uchun {soni} ta "
            f"slaydlik reja tuz. FAQAT quyidagi JSON formatida javob ber, boshqa "
            f"hech qanday matn yozma:\n"
            f'{{"slides": [{{"title": "Slayd sarlavhasi", '
            f'"image_query": "2-3 word english keyword for photo search", '
            f'"bullets": ["to\'liq va batafsil band 1", "to\'liq va batafsil band 2"]}}]}}\n'
            f"O'zbek tilida yoz (faqat image_query maydoni inglizcha bo'lsin, chunki "
            f"u rasm qidirish uchun ishlatiladi). Aynan {soni} ta slayd bo'lsin, "
            f"har birida 3-5 ta band, har bir band kamida 10-12 so'zdan iborat, "
            f"aniq va ma'lumotga boy bo'lsin (shunchaki qisqa sarlavha emas, "
            f"to'liq fikr bildiruvchi gap bo'lsin). "
            f"MUHIM: hech qanday Markdown belgilaridan (**, *, #) foydalanma. "
            f"MUHIM: javobing FAQAT JSON bo'lsin, tushuntirish yoki boshqa matn yozma, "
            f"JSON to'liq va yopilgan bo'lishi shart."
        )
        try:
            javob = ask_ai(prompt, max_tokens=6000)
            try:
                data = extract_json(javob)
            except Exception:
                print("Birinchi urinish muvaffaqiyatsiz, qayta urinilmoqda...")
                javob = ask_ai(prompt, max_tokens=6000)
                data = extract_json(javob)
            pptx_file = create_pptx(mavzu, muallif, data["slides"])
            pptx_file.name = f"{mavzu[:40]}.pptx"
            await update.message.reply_document(document=pptx_file, filename=pptx_file.name)
        except Exception as e:
            traceback.print_exc()
            await update.message.reply_text(f"Xatolik yuz berdi, qayta urinib ko'ring. ({e})")
    else:
        await update.message.reply_text("Mustaqil ish yozilmoqda, biroz kuting...")
        soz_soni = soni * 350
        prompt = (
            f"'{mavzu}' mavzusida chuqur va batafsil mustaqil ish (referat) yoz, "
            f"taxminan {soz_soni} so'z (bu {soni} sahifaga teng). "
            f"Kirish:, Asosiy qism:, Xulosa:, Foydalanilgan adabiyotlar: kabi bo'lim "
            f"sarlavhalari bilan. Asosiy qismni 2-3 ta kichik mavzuga bo'lib, har "
            f"birini alohida sarlavha bilan (masalan 'Tarixiy jihatlari:', "
            f"'Hozirgi holati:' kabi) chuqur yorit. Har bir bo'limda aniq faktlar, "
            f"misollar va tushuntirishlar bo'lsin - umumiy va yuzaki gaplardan "
            f"qoching. O'zbek tilida, ilmiy uslubda, professional va ma'lumotga "
            f"boy qilib yoz. "
            f"MUHIM: hech qanday Markdown belgilaridan (**, *, #, -) foydalanma, "
            f"faqat oddiy toza matn yoz."
        )
        try:
            matn = ask_ai(prompt, max_tokens=4000)
            docx_file = create_docx(mavzu, muallif, matn)
            docx_file.name = f"{mavzu[:40]}.docx"
            await update.message.reply_document(document=docx_file, filename=docx_file.name)
        except Exception as e:
            traceback.print_exc()
            await update.message.reply_text(f"Xatolik yuz berdi, qayta urinib ko'ring. ({e})")

    context.user_data.clear()
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("Bekor qilindi.")
    return ConversationHandler.END


def main():
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[
            CommandHandler("slayd", slayd_start),
            CommandHandler("mustaqil", mustaqil_start),
        ],
        states={
            MAVZU: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_mavzu)],
            SONI: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_soni)],
            MUALLIF: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_muallif)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(conv)
    print("Bot ishga tushdi...")
    app.run_polling()


if __name__ == "__main__":
    main()
