"""
Talaba Yordamchi Bot (Groq bilan - BEPUL, Railway uchun tayyor)
-----------------------------------------------------------------
Bu bot Telegram orqali talabalarga:
  /slayd <mavzu>     - slayd mazmuni (rejasi) tayyorlab beradi
  /mustaqil <mavzu>  - mustaqil ish matnini yozib beradi
"""

import os
from groq import Groq
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
GROQ_API_KEY = os.environ.get("GROQ_API_KEY")

if not TELEGRAM_TOKEN or not GROQ_API_KEY:
    raise ValueError(
        "TELEGRAM_TOKEN yoki GROQ_API_KEY topilmadi! "
        "Environment variables to'g'ri sozlanganini tekshiring."
    )

client = Groq(api_key=GROQ_API_KEY)
MODEL = "llama-3.3-70b-versatile"


def ask_ai(prompt: str) -> str:
    response = client.chat.completions.create(
        model=MODEL,
        max_tokens=1500,
        messages=[{"role": "user", "content": prompt}],
    )
    return response.choices[0].message.content


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Salom! Men talabalarga yordam beruvchi botman.\n\n"
        "Buyruqlar:\n"
        "/slayd <mavzu> - slayd mazmuni tayyorlab beraman\n"
        "/mustaqil <mavzu> - mustaqil ish matnini yozib beraman\n\n"
        "Masalan: /slayd Sun'iy intellekt tarixi"
    )


async def slayd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    mavzu = " ".join(context.args)
    if not mavzu:
        await update.message.reply_text("Iltimos, mavzuni ham yozing. Masalan:\n/slayd Ekologik muammolar")
        return
    await update.message.reply_text("Slayd tayyorlanmoqda, biroz kuting...")
    prompt = (
        f"'{mavzu}' mavzusida taqdimot (prezentatsiya) uchun 8-10 ta slaydlik "
        f"rejani tuzib ber. Har bir slayd uchun: sarlavha va 3-5 ta qisqa band "
        f"(bullet point) yoz. O'zbek tilida, aniq va tushunarli qilib yoz."
    )
    natija = ask_ai(prompt)
    await update.message.reply_text(natija)


async def mustaqil(update: Update, context: ContextTypes.DEFAULT_TYPE):
    mavzu = " ".join(context.args)
    if not mavzu:
        await update.message.reply_text("Iltimos, mavzuni ham yozing. Masalan:\n/mustaqil Bozor iqtisodiyoti")
        return
    await update.message.reply_text("Mustaqil ish yozilmoqda, biroz kuting...")
    prompt = (
        f"'{mavzu}' mavzusida mustaqil ish (referat) yoz. Kirish, asosiy qism "
        f"(2-3 bo'lim), xulosa va foydalanilgan adabiyotlar ro'yxatidan iborat "
        f"bo'lsin. O'zbek tilida, ilmiy uslubda, taxminan 500-700 so'z."
    )
    natija = ask_ai(prompt)
    await update.message.reply_text(natija)


def main():
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("slayd", slayd))
    app.add_handler(CommandHandler("mustaqil", mustaqil))
    print("Bot ishga tushdi...")
    app.run_polling()


if __name__ == "__main__":
    main()
