"""
Talaba Yordamchi Bot (Groq bilan - BEPUL, Railway uchun tayyor)
-----------------------------------------------------------------
Bu bot Telegram orqali talabalarga:
  /slayd <mavzu>     - haqiqiy .pptx (PowerPoint) fayl tayyorlab beradi
  /mustaqil <mavzu>  - haqiqiy .docx (Word) fayl tayyorlab beradi
"""

import os
import json
import re
from io import BytesIO

from groq import Groq
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes

from pptx import Presentation
from pptx.util import Inches, Pt

from docx import Document
from docx.shared import Pt as DocxPt

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
GROQ_API_KEY = os.environ.get("GROQ_API_KEY")

if not TELEGRAM_TOKEN or not GROQ_API_KEY:
    raise ValueError(
        "TELEGRAM_TOKEN yoki GROQ_API_KEY topilmadi! "
        "Environment variables to'g'ri sozlanganini tekshiring."
    )

client = Groq(api_key=GROQ_API_KEY)
MODEL = "openai/gpt-oss-120b"


def ask_ai(prompt: str) -> str:
    response = client.chat.completions.create(
        model=MODEL,
        max_tokens=2000,
        messages=[{"role": "user", "content": prompt}],
    )
    return response.choices[0].message.content


def extract_json(text: str) -> dict:
    """AI javobidan JSON qismini ajratib oladi (agar qo'shimcha matn bo'lsa ham)."""
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if not match:
        raise ValueError("AI javobida JSON topilmadi")
    return json.loads(match.group(0))


def create_pptx(mavzu: str, slides_data: list) -> BytesIO:
    """Slaydlar ro'yxatidan haqiqiy .pptx fayl yaratadi (xotirada, diskka yozmasdan)."""
    prs = Presentation()

    # Sarlavha slaydi
    title_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_layout)
    slide.shapes.title.text = mavzu
    if len(slide.placeholders) > 1:
        slide.placeholders[1].text = "Taqdimot"

    # Mazmun slaydlari
    content_layout = prs.slide_layouts[1]
    for item in slides_data:
        slide = prs.slides.add_slide(content_layout)
        slide.shapes.title.text = item.get("title", "")
        body = slide.placeholders[1].text_frame
        body.clear()
        bullets = item.get("bullets", [])
        for i, bullet in enumerate(bullets):
            p = body.paragraphs[0] if i == 0 else body.add_paragraph()
            p.text = bullet
            p.font.size = Pt(20)

    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer


def create_docx(mavzu: str, matn: str) -> BytesIO:
    """Matndan haqiqiy .docx fayl yaratadi (xotirada, diskka yozmasdan)."""
    doc = Document()

    title = doc.add_heading(mavzu, level=1)

    for paragraph in matn.split("\n"):
        paragraph = paragraph.strip()
        if not paragraph:
            continue
        # Sarlavhalarni aniqlash (masalan "Kirish:", "Xulosa:")
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
        "/slayd <mavzu> - PowerPoint (.pptx) fayl tayyorlab beraman\n"
        "/mustaqil <mavzu> - Word (.docx) fayl tayyorlab beraman\n\n"
        "Masalan: /slayd Sun'iy intellekt tarixi"
    )


async def slayd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    mavzu = " ".join(context.args)
    if not mavzu:
        await update.message.reply_text("Iltimos, mavzuni ham yozing. Masalan:\n/slayd Ekologik muammolar")
        return

    await update.message.reply_text("Slayd tayyorlanmoqda, biroz kuting...")

    prompt = (
        f"'{mavzu}' mavzusida taqdimot uchun 8 ta slaydlik reja tuz. "
        f"FAQAT quyidagi JSON formatida javob ber, boshqa hech qanday matn yozma:\n"
        f'{{"slides": [{{"title": "Slayd sarlavhasi", "bullets": '
        f'["band 1", "band 2", "band 3"]}}]}}\n'
        f"O'zbek tilida yoz. 8 ta slayd bo'lsin, har birida 3-5 ta band."
    )

    try:
        javob = ask_ai(prompt)
        data = extract_json(javob)
        pptx_file = create_pptx(mavzu, data["slides"])
        pptx_file.name = f"{mavzu[:40]}.pptx"
        await update.message.reply_document(document=pptx_file, filename=pptx_file.name)
    except Exception as e:
        await update.message.reply_text(f"Xatolik yuz berdi, qayta urinib ko'ring. ({e})")


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
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("slayd", slayd))
    app.add_handler(CommandHandler("mustaqil", mustaqil))
    print("Bot ishga tushdi...")
    app.run_polling()


if __name__ == "__main__":
    main()
