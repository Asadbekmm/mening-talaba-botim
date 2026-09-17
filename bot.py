"""
Talaba Yordamchi Bot (Groq + Pexels, Railway uchun tayyor)
-----------------------------------------------------------------
/slayd  - suhbat orqali: mavzu -> slaydlar soni (1-12) -> muallif ismi
          -> rasmli, dizaynli .pptx fayl
/mustaqil <mavzu> - .docx (Word) fayl
"""

import os
import json
import re
import random
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
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from docx import Document
from docx.shared import Pt as DocxPt

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

MAX_SLIDES = 12

# Suhbat bosqichlari
MAVZU, SONI, MUALLIF = range(3)

# Turli dizaynlar (fon rangi, sarlavha rangi, matn rangi, urg'u rangi)
THEMES = [
    {"bg": RGBColor(0x1A, 0x23, 0x3A), "title": RGBColor(0xFF, 0xFF, 0xFF), "text": RGBColor(0xE0, 0xE0, 0xE0), "accent": RGBColor(0x4F, 0xA8, 0xE0)},
    {"bg": RGBColor(0xFF, 0xFF, 0xFF), "title": RGBColor(0x1A, 0x23, 0x3A), "text": RGBColor(0x33, 0x33, 0x33), "accent": RGBColor(0xE0, 0x6A, 0x4F)},
    {"bg": RGBColor(0x2D, 0x2A, 0x4A), "title": RGBColor(0xFF, 0xD9, 0x66), "text": RGBColor(0xF0, 0xF0, 0xF0), "accent": RGBColor(0xFF, 0xD9, 0x66)},
    {"bg": RGBColor(0xF4, 0xF1, 0xEA), "title": RGBColor(0x3A, 0x5A, 0x40), "text": RGBColor(0x2C, 0x2C, 0x2C), "accent": RGBColor(0x3A, 0x5A, 0x40)},
]


def ask_ai(prompt: str) -> str:
    response = client.chat.completions.create(
        model=MODEL,
        max_tokens=2500,
        messages=[{"role": "user", "content": prompt}],
    )
    return response.choices[0].message.content


def extract_json(text: str) -> dict:
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if not match:
        raise ValueError("AI javobida JSON topilmadi")
    return json.loads(match.group(0))


def get_image(query: str):
    """Pexels'dan mavzuga mos rasm oladi. Muvaffaqiyatsiz bo'lsa None qaytaradi."""
    if not PEXELS_API_KEY:
        return None
    try:
        resp = requests.get(
            "https://api.pexels.com/v1/search",
            headers={"Authorization": PEXELS_API_KEY},
            params={"query": query, "per_page": 1, "orientation": "landscape"},
            timeout=10,
        )
        data = resp.json()
        photos = data.get("photos", [])
        if not photos:
            return None
        img_url = photos[0]["src"]["large"]
        img_resp = requests.get(img_url, timeout=10)
        return BytesIO(img_resp.content)
    except Exception:
        return None


def create_pptx(mavzu: str, muallif: str, slides_data: list) -> BytesIO:
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    theme = random.choice(THEMES)
    blank = prs.slide_layouts[6]

    def add_background(slide):
        bg = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = theme["bg"]
        bg.line.fill.background()
        bg.shadow.inherit = False
        # Fonni orqaga qaytarish
        slide.shapes._spTree.remove(bg._element)
        slide.shapes._spTree.insert(2, bg._element)
        return bg

    # --- Titul slayd ---
    slide = prs.slides.add_slide(blank)
    add_background(slide)
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
    tf2 = author_box.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = f"Tayyorladi: {muallif}"
    p2.alignment = PP_ALIGN.CENTER
    p2.font.size = Pt(22)
    p2.font.color.rgb = theme["accent"]

    total = len(slides_data)

    # --- Mazmun slaydlari ---
    for idx, item in enumerate(slides_data, start=1):
        slide = prs.slides.add_slide(blank)
        add_background(slide)

        image_on_left = idx % 2 == 1  # slaydlar bo'yicha almashadi - dizayn xilma-xilligi uchun
        text_x = Inches(6.8) if image_on_left else Inches(0.7)
        img_x = Inches(0.5) if image_on_left else Inches(7.1)

        # Sarlavha
        title_box = slide.shapes.add_textbox(text_x, Inches(0.5), Inches(6.0), Inches(1.0))
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = item.get("title", "")
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = theme["title"]

        # Bandlar (bullets)
        body_box = slide.shapes.add_textbox(text_x, Inches(1.6), Inches(6.0), Inches(5.3))
        tf = body_box.text_frame
        tf.word_wrap = True
        bullets = item.get("bullets", [])
        for i, bullet in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = f"•  {bullet}"
            p.font.size = Pt(18)
            p.font.color.rgb = theme["text"]
            p.space_after = Pt(12)

        # Rasm
        img_data = get_image(item.get("title", mavzu))
        if img_data:
            try:
                slide.shapes.add_picture(img_data, img_x, Inches(1.6), width=Inches(5.7), height=Inches(4.8))
            except Exception:
                pass

        # Slayd raqami
        num_box = slide.shapes.add_textbox(Inches(12.5), Inches(7.05), Inches(0.7), Inches(0.4))
        p = num_box.text_frame.paragraphs[0]
        p.text = f"{idx}/{total}"
        p.font.size = Pt(12)
        p.font.color.rgb = theme["accent"]

    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer


def create_docx(mavzu: str, matn: str) -> BytesIO:
    doc = Document()
    doc.add_heading(mavzu, level=1)
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
        "/mustaqil <mavzu> - Word (.docx) fayl tayyorlab beraman\n"
    )


# ===== /slayd suhbati =====

async def slayd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Taqdimot mavzusini yozing:")
    return MAVZU


async def slayd_mavzu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["mavzu"] = update.message.text.strip()
    await update.message.reply_text(f"Nechta slayd kerak? (1 dan {MAX_SLIDES} gacha)")
    return SONI


async def slayd_soni(update: Update, context: ContextTypes.DEFAULT_TYPE):
    matn = update.message.text.strip()
    if not matn.isdigit() or not (1 <= int(matn) <= MAX_SLIDES):
        await update.message.reply_text(f"Iltimos, 1 dan {MAX_SLIDES} gacha son kiriting:")
        return SONI
    context.user_data["soni"] = int(matn)
    await update.message.reply_text("Tuzuvchi (muallif) ismi va familiyasini yozing:")
    return MUALLIF


async def slayd_muallif(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["muallif"] = update.message.text.strip()
    await update.message.reply_text("Taqdimot tayyorlanmoqda, biroz kuting...")

    mavzu = context.user_data["mavzu"]
    soni = context.user_data["soni"]
    muallif = context.user_data["muallif"]

    prompt = (
        f"'{mavzu}' mavzusida taqdimot uchun {soni} ta slaydlik reja tuz. "
        f"FAQAT quyidagi JSON formatida javob ber, boshqa hech qanday matn yozma:\n"
        f'{{"slides": [{{"title": "Slayd sarlavhasi", "bullets": '
        f'["band 1", "band 2", "band 3"]}}]}}\n'
        f"O'zbek tilida yoz. Aynan {soni} ta slayd bo'lsin, har birida 3-4 ta band."
    )

    try:
        javob = ask_ai(prompt)
        data = extract_json(javob)
        pptx_file = create_pptx(mavzu, muallif, data["slides"])
        pptx_file.name = f"{mavzu[:40]}.pptx"
        await update.message.reply_document(document=pptx_file, filename=pptx_file.name)
    except Exception as e:
        await update.message.reply_text(f"Xatolik yuz berdi, qayta urinib ko'ring. ({e})")

    context.user_data.clear()
    return ConversationHandler.END


async def slayd_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("Bekor qilindi.")
    return ConversationHandler.END


# ===== /mustaqil =====

async def mustaqil(update: Update, context: ContextTypes.DEFAULT_TYPE):
    mavzu = " ".join(context.args)
    if not mavzu:
        await update.message.reply_text("Iltimos, mavzuni ham yozing. Masalan:\n/mustaqil Bozor iqtisodiyoti")
        return

    await update.message.reply_text("Mustaqil ish yozilmoqda, biroz kuting...")

    prompt = (
        f"'{mavzu}' mavzusida mustaqil ish (referat) yoz. Kirish:, Asosiy qism:, "
        f"Xulosa:, Foydalanilgan adabiyotlar: kabi bo'lim sarlavhalari bilan. "
        f"O'zbek tilida, ilmiy uslubda, taxminan 600-800 so'z."
    )

    try:
        matn = ask_ai(prompt)
        docx_file = create_docx(mavzu, matn)
        docx_file.name = f"{mavzu[:40]}.docx"
        await update.message.reply_document(document=docx_file, filename=docx_file.name)
    except Exception as e:
        await update.message.reply_text(f"Xatolik yuz berdi, qayta urinib ko'ring. ({e})")


def main():
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()

    slayd_conv = ConversationHandler(
        entry_points=[CommandHandler("slayd", slayd_start)],
        states={
            MAVZU: [MessageHandler(filters.TEXT & ~filters.COMMAND, slayd_mavzu)],
            SONI: [MessageHandler(filters.TEXT & ~filters.COMMAND, slayd_soni)],
            MUALLIF: [MessageHandler(filters.TEXT & ~filters.COMMAND, slayd_muallif)],
        },
        fallbacks=[CommandHandler("cancel", slayd_cancel)],
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(slayd_conv)
    app.add_handler(CommandHandler("mustaqil", mustaqil))
    print("Bot ishga tushdi...")
    app.run_polling()


if __name__ == "__main__":
    main()
