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


def strip_markdown(text: str) -> str:
    text = re.sub(r"\*\*(.*?)\*\*", r"\1", text)
    text = re.sub(r"\*(.*?)\*", r"\1", text)
    text = re.sub(r"^#{1,6}\s*", "", text, flags=re.MULTILINE)
    text = re.sub(r"^-\s+", "", text, flags=re.MULTILINE)
    return text


def extract_json(text: str) -> dict:
    text = re.sub(r"```(?:json)?", "", text)
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if not match:
        raise ValueError("AI javobida JSON topilmadi")
    return json.loads(match.group(0))


def get_image(query: str):
    """
    Pexels'dan rasm oladi.
    Qaytaradi: (BytesIO yoki None, sabab_matni)
    """
    if not PEXELS_API_KEY:
        return None, "PEXELS_API_KEY sozlanmagan (Railway Variables'da yo'q)"
    try:
        resp = requests.get(
            "https://api.pexels.com/v1/search",
            headers={"Authorization": PEXELS_API_KEY.strip()},
            params={"query": query, "per_page": 1, "orientation": "landscape"},
            timeout=15,
        )
        if resp.status_code == 401:
            return None, "PEXELS_API_KEY noto'g'ri (401 - ruxsat berilmadi)"
        if resp.status_code != 200:
            return None, f"Pexels xatosi (status={resp.status_code})"
        data = resp.json()
        photos = data.get("photos", [])
        if not photos:
            return None, f"'{query}' uchun rasm topilmadi"
        img_url = photos[0]["src"]["large"]
        img_resp = requests.get(img_url, timeout=15)
        if img_resp.status_code != 200:
            return None, "Rasm faylini yuklab bo'lmadi"
        return BytesIO(img_resp.content), None
    except Exception as e:
        return None, f"Kutilmagan xato: {e}"


def add_background(slide, prs, color):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.line.fill.background()
    bg.shadow.inherit = False
    return bg


def create_pptx(mavzu: str, muallif: str, slides_data: list):
    """Qaytaradi: (BytesIO, rasm_xatolari_royxati)"""
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    theme = random.choice(THEMES)
    blank = prs.slide_layouts[6]
    image_errors = []

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
        p.text = strip_markdown(item.get("title", ""))
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
            p.font.size = Pt(13)
            p.font.color.rgb = theme["text"]
            p.space_after = Pt(8)

        query = item.get("image_query") or item.get("title") or mavzu
        img_data, err = get_image(query)
        if img_data:
            try:
                slide.shapes.add_picture(img_data, img_x, Inches(1.6), width=Inches(5.7), height=Inches(4.8))
            except Exception as e:
                image_errors.append(f"Slayd {idx}: rasmni joylashda xato - {e}")
        elif err:
            image_errors.append(f"Slayd {idx} ('{query}'): {err}")

        num_box = slide.shapes.add_textbox(Inches(12.5), Inches(7.05), Inches(0.7), Inches(0.4))
        p = num_box.text_frame.paragraphs[0]
        p.text = f"{idx}/{total}"
        p.font.size = Pt(12)
        p.font.color.rgb = theme["accent"]

    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer, image_errors


def create_docx(mavzu: str, muallif: str, matn: str) -> BytesIO:
    doc = Document()

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
        await update.message.reply_text(f"Slayd sahifalari soni (1 dan {MAX_COUNT} gacha):")
    else:
        await update.message.reply_text(f"Mustaqil ish sahifalari soni (1 dan {MAX_COUNT} gacha):")
    return SONI


async def handle_soni(update: Update, context: ContextTypes.DEFAULT_TYPE):
    matn = update.message.text.strip()
    if not matn.isdigit() or not (1 <= int(matn) <= MAX_COUNT):
        await update.message.reply_text(f"Iltimos, 1 dan {MAX_COUNT} gacha butun son kiriting:")
        return SONI
    context.user_data["soni"] = int(matn)
    await update.message.reply_text("Tuzuvchi (muallif) ismi va familiyasini kiriting:")
    return MUALLIF


def generate_outline(mavzu: str, soni: int) -> list:
    """1-bosqich: faqat slayd sarlavhalari ro'yxatini oladi (qisqa, ishonchli)."""
    prompt = (
        f"'{mavzu}' mavzusida {soni} ta slaydlik taqdimot uchun FAQAT slayd "
        f"sarlavhalari ro'yxatini tuz. FAQAT shu JSON formatida javob ber:\n"
        f'{{"titles": ["Sarlavha 1", "Sarlavha 2"]}}\n'
        f"Aynan {soni} ta sarlavha bo'lsin, mavzuni mantiqiy ketma-ketlikda "
        f"(kirishdan xulosagacha) yoritsin. O'zbek tilida yoz."
    )
    javob = ask_ai(prompt, max_tokens=1000)
    data = extract_json(javob)
    titles = data["titles"][:soni]
    while len(titles) < soni:
        titles.append(f"{mavzu} - qo'shimcha ma'lumot")
    return titles


def generate_slide_content(mavzu: str, title: str) -> dict:
    """2-bosqich: bitta slayd uchun batafsil bandlar va rasm kalit so'zini oladi."""
    prompt = (
        f"Mavzu: '{mavzu}'. Slayd sarlavhasi: '{title}'.\n"
        f"Shu slayd uchun FAQAT quyidagi JSON formatida javob ber:\n"
        f'{{"image_query": "2-3 word english keyword", "bullets": '
        f'["batafsil band 1", "batafsil band 2"]}}\n'
        f"4-5 ta band yoz, har biri kamida 20-25 so'zdan iborat, to'liq va "
        f"ma'lumotga boy fikr bo'lsin. Bandlar o'zbek tilida, image_query "
        f"inglizcha bo'lsin. Markdown belgilaridan (**, *, #) foydalanma."
    )
    javob = ask_ai(prompt, max_tokens=1200)
    return extract_json(javob)


def build_slides_data(mavzu: str, soni: int) -> list:
    titles = generate_outline(mavzu, soni)
    slides_data = []
    for title in titles:
        try:
            content = generate_slide_content(mavzu, title)
        except Exception:
            try:
                content = generate_slide_content(mavzu, title)
            except Exception:
                content = {"image_query": mavzu, "bullets": [f"{title} haqida ma'lumot."]}
        slides_data.append({
            "title": title,
            "image_query": content.get("image_query", mavzu),
            "bullets": content.get("bullets", []),
        })
    return slides_data


async def handle_muallif(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["muallif"] = update.message.text.strip()
    turi = context.user_data["turi"]
    mavzu = context.user_data["mavzu"]
    soni = context.user_data["soni"]
    muallif = context.user_data["muallif"]

    if turi == "slayd":
        await update.message.reply_text("Taqdimot tayyorlanmoqda, biroz kuting (bu bir necha o'n soniya davom etishi mumkin)...")
        try:
            slides_data = build_slides_data(mavzu, soni)

            pptx_file, image_errors = create_pptx(mavzu, muallif, slides_data)
            pptx_file.name = f"{mavzu[:40]}.pptx"
            await update.message.reply_document(document=pptx_file, filename=pptx_file.name)

            if image_errors:
                unique_reasons = list(dict.fromkeys(e.split(": ", 1)[-1] for e in image_errors))
                await update.message.reply_text(
                    "Eslatma: ba'zi slaydlarga rasm qo'shilmadi.\n"
                    "Sabab: " + "; ".join(unique_reasons[:3])
                )
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
            f"birini alohida sarlavha bilan chuqur yorit. Har bir bo'limda aniq "
            f"faktlar, misollar va tushuntirishlar bo'lsin. O'zbek tilida, ilmiy "
            f"uslubda, professional va ma'lumotga boy qilib yoz. "
            f"MUHIM: hech qanday Markdown belgilaridan (**, *, #, -) foydalanma."
        )
        try:
            matn = ask_ai(prompt, max_tokens=5000)
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
