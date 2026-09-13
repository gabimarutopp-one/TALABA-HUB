"""
PYTHONANYWHERE UCHUN WSGI FAYLI
================================
PythonAnywhere "Web" bo'limida ilova yaratganingizda avtomatik
/var/www/SIZNING_username_pythonanywhere_com_wsgi.py fayli hosil bo'ladi.
O'SHA FAYLNING ICHINI to'liq shu kod bilan almashtiring
(fayl nomini o'zi o'zgartirmang, PythonAnywhere shu nomni kutadi).

Ishga tushirishdan oldin quyidagi 2 ta joyni albatta o'zgartiring:
  1) project_home    -> bot.py joylashgan papka yo'li
  2) WEBHOOK_SECRET_PATH -> tasodifiy, taxmin qilib bo'lmaydigan uzun satr

HEMIS mini-app (talabalar login/parol kiritadigan forma) uchun:
  - hemis_login.html faylini ham xuddi shu papkaga (project_home) yuklang.
  - "Web" sahifasidagi "Environment variables"ga qo'shing:
      HEMIS_WEBAPP_URL = https://SIZNING_USERNAME.pythonanywhere.com/hemis-login
  - Ixtiyoriy, lekin tavsiya etiladi (parolni shifrlab saqlash uchun):
      HEMIS_ENC_KEY = (generatsiya: python3 -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())")
"""
import sys
import os
import asyncio
import logging

# ---------------------------------------------------------------
# 1) Loyiha papkasini Python yo'liga qo'shamiz
#    Masalan, agar bot.py papkasi /home/aliuz/bot bo'lsa:
project_home = "/home/SIZNING_USERNAME/bot"
if project_home not in sys.path:
    sys.path.insert(0, project_home)

# ---------------------------------------------------------------
# BOT_TOKEN'ni "Web" sahifasidagi "Environment variables" bo'limi orqali
# bering (tavsiya etiladi). Agar u yerda muammo bo'lsa, faqat vaqtinchalik
# sinov uchun quyidagi qatorni oching va o'z tokeningizni yozing:
# os.environ.setdefault("BOT_TOKEN", "SIZNING_TOKENINGIZ")

from flask import Flask, request
from aiogram.types import Update

import bot as botmodule  # bot.py ichidan tayyor 'bot' va 'dp' obyektlarini olamiz

application = Flask(__name__)

# ---------------------------------------------------------------
# 2) Xavfsizlik uchun maxfiy yo'l. Buni tasodifiy uzun satrga almashtiring
#    (masalan: python3 -c "import secrets; print(secrets.token_urlsafe(32))"
#    buyrug'i bilan Bash konsolda generatsiya qiling).
WEBHOOK_SECRET_PATH = "shu-yerga-oz-tasodifiy-maxfiy-satringizni-yozing"


@application.route(f"/webhook/{WEBHOOK_SECRET_PATH}", methods=["POST"])
def telegram_webhook():
    try:
        update_data = request.get_json(force=True)
        update = Update.model_validate(update_data)
        asyncio.run(botmodule.dp.feed_webhook_update(botmodule.bot, update))
    except Exception:
        logging.exception("Webhook so'rovini qayta ishlashda xatolik")
    return "OK"


@application.route("/", methods=["GET"])
def health_check():
    # Brauzerda saytga kirganda shu ko'rinadi - bot ishlab turganini bilish uchun
    return "Bot ishlayapti ✅"


@application.route("/hemis-login", methods=["GET"])
def hemis_login_page():
    # HEMIS mini-app (Telegram WebApp) sahifasi
    html_path = os.path.join(project_home, "hemis_login.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return f.read()
