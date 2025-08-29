"""
Production Telegram PDF bot with converters (aiogram 3 + FastAPI webhooks)
----------------------------------------------------------------------------

Features:
- PDF → JPG/PNG (PyMuPDF)
- PDF → DOCX/PPTX/XLSX (LibreOffice)
- PDF → Excel tables (Camelot)
- Split/Merge/Compress (pikepdf)
- Watermarks (reportlab + pikepdf)
- OCR (ocrmypdf)
"""

import asyncio
import logging
import os
import tempfile
from contextlib import asynccontextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv
from fastapi import FastAPI, Request, Response
from fastapi.responses import JSONResponse

from aiogram import Bot, Dispatcher, Router, F
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import CommandStart
from aiogram.fsm.storage.redis import RedisStorage
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import FSInputFile, Message, InlineKeyboardMarkup, InlineKeyboardButton, CallbackQuery

import redis.asyncio as aioredis
import fitz  # PyMuPDF
import subprocess
import camelot
import pikepdf
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import shutil
import zipfile

# --------------------------
# Config
# --------------------------
load_dotenv()

@dataclass
class Settings:
    token = os.getenv("BOT_TOKEN")
    webhook_base: str = os.getenv("WEBHOOK_BASE", "")
    webhook_path: str = os.getenv("WEBHOOK_PATH", "/tg/webhook")
    webhook_secret: str = os.getenv("WEBHOOK_SECRET", "")
    redis_url: str = os.getenv("REDIS_URL", "redis://localhost:6379/0")
    max_concurrent_jobs: int = int(os.getenv("MAX_CONCURRENT_JOBS", "6"))
    max_file_mb: int = int(os.getenv("MAX_FILE_MB", "50"))
    job_timeout_s: int = int(os.getenv("JOB_TIMEOUT_S", "300"))
    temp_dir: str = os.getenv("TEMP_DIR", tempfile.gettempdir())

SET = Settings()
logging.basicConfig(level=logging.INFO)
log = logging.getLogger("pdf-bot")

if not SET.token:
    raise SystemExit("BOT_TOKEN required")

router = Router()
redis_pool: Optional[aioredis.Redis] = None

storage = RedisStorage.from_url(SET.redis_url)
bot = Bot(token=SET.token, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher(storage=storage)
dp.include_router(router)

JOB_SEMAPHORE = asyncio.Semaphore(SET.max_concurrent_jobs)

# --------------------------
# Utils
# --------------------------
async def run_blocking(func, *args, **kwargs):
    return await asyncio.to_thread(func, *args, **kwargs)

@asynccontextmanager
async def temp_file(suffix: str = ""):
    fd, path = tempfile.mkstemp(suffix=suffix, dir=SET.temp_dir)
    os.close(fd)
    try:
        yield Path(path)
    finally:
        try:
            os.remove(path)
        except FileNotFoundError:
            pass

# --------------------------
# FSM
# --------------------------
class PDFStates(StatesGroup):
    waiting_file = State()
    waiting_action = State()

# --------------------------
# PDF Operations
# --------------------------
def pdf_to_jpg(src: Path, dst_dir: Path):
    doc = fitz.open(src)
    dst_dir.mkdir(parents=True, exist_ok=True)
    out_files = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=150)
        out_path = dst_dir / f"page_{i+1}.jpg"
        pix.save(out_path)
        out_files.append(out_path)
    return out_files

def pdf_to_docx(src: Path, dst: Path):
    subprocess.run(["libreoffice", "--headless", "--convert-to", "docx", str(src), "--outdir", str(dst.parent)], check=True)
    return dst

def pdf_to_pptx(src: Path, dst: Path):
    subprocess.run(["libreoffice", "--headless", "--convert-to", "pptx", str(src), "--outdir", str(dst.parent)], check=True)
    return dst

def pdf_to_xlsx(src: Path, dst: Path):
    subprocess.run(["libreoffice", "--headless", "--convert-to", "xlsx", str(src), "--outdir", str(dst.parent)], check=True)
    return dst

def pdf_tables_to_excel(src: Path, dst: Path):
    tables = camelot.read_pdf(str(src), pages="all")
    tables.export(str(dst), f="excel")
    return dst

def pdf_merge(files: list[Path], dst: Path):
    pdf = pikepdf.Pdf.new()
    for f in files:
        pdf.pages.extend(pikepdf.Pdf.open(f).pages)
    pdf.save(dst)
    return dst

def pdf_split(src: Path, dst_dir: Path):
    pdf = pikepdf.Pdf.open(src)
    outputs = []
    for i, page in enumerate(pdf.pages):
        new_pdf = pikepdf.Pdf.new()
        new_pdf.pages.append(page)
        out_path = dst_dir / f"page_{i+1}.pdf"
        new_pdf.save(out_path)
        outputs.append(out_path)
    return outputs

def pdf_compress(src: Path, dst: Path):
    pdf = pikepdf.open(src)
    pdf.save(dst, compress_streams=True)
    return dst

def pdf_watermark(src: Path, dst: Path, text: str = "WATERMARK"):
    wm_pdf = dst.parent / "wm.pdf"
    c = canvas.Canvas(str(wm_pdf), pagesize=letter)
    c.setFont("Helvetica", 40)
    c.setFillGray(0.5, 0.5)
    c.saveState()
    c.translate(300, 400)
    c.rotate(45)
    c.drawString(0, 0, text)
    c.restoreState()
    c.save()
    base = pikepdf.open(src)
    wm = pikepdf.open(wm_pdf)
    for page in base.pages:
        page_obj = page.as_form_xobject()
        page_obj.add_overlay(wm.pages[0])
    base.save(dst)
    return dst

def pdf_ocr(src: Path, dst: Path):
    subprocess.run(["ocrmypdf", str(src), str(dst)], check=True)
    return dst

# --------------------------
# Handlers
# --------------------------
@router.message(CommandStart())
async def on_start(message: Message, state: FSMContext):
    await message.answer("👋 Привет! Пришли PDF для обработки.")
    await state.set_state(PDFStates.waiting_file)

@router.message(PDFStates.waiting_file, F.document)
async def on_file(message: Message, state: FSMContext):
    if not message.document.file_name.endswith(".pdf"):
        await message.answer("⚠️ Пожалуйста, отправь PDF файл.")
        return

    file = await bot.get_file(message.document.file_id)
    async with temp_file(suffix=".pdf") as pdf_path:
        await bot.download_file(file.file_path, destination=str(pdf_path))
        await state.update_data(pdf_path=str(pdf_path))

        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="PDF→DOCX", callback_data="to_docx"),
             InlineKeyboardButton(text="PDF→PPTX", callback_data="to_pptx")],
            [InlineKeyboardButton(text="PDF→XLSX", callback_data="to_xlsx"),
             InlineKeyboardButton(text="PDF→JPG", callback_data="to_jpg")],
            [InlineKeyboardButton(text="Split", callback_data="split"),
             InlineKeyboardButton(text="Compress", callback_data="compress")],
            [InlineKeyboardButton(text="Watermark", callback_data="watermark"),
             InlineKeyboardButton(text="OCR", callback_data="ocr")],
        ])
        await message.answer("Выберите действие:", reply_markup=kb)
        await state.set_state(PDFStates.waiting_action)

@router.callback_query(PDFStates.waiting_action)
async def on_action(call: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    pdf_path = Path(data["pdf_path"])

    async with JOB_SEMAPHORE:
        if call.data == "to_docx":
            out = pdf_path.with_suffix(".docx")
            await run_blocking(pdf_to_docx, pdf_path, out)
            await call.message.answer_document(FSInputFile(out), caption="DOCX готов ✅")
        elif call.data == "to_pptx":
            out = pdf_path.with_suffix(".pptx")
            await run_blocking(pdf_to_pptx, pdf_path, out)
            await call.message.answer_document(FSInputFile(out), caption="PPTX готов ✅")
        elif call.data == "to_xlsx":
            out = pdf_path.with_suffix(".xlsx")
            await run_blocking(pdf_to_xlsx, pdf_path, out)
            await call.message.answer_document(FSInputFile(out), caption="XLSX готов ✅")
        elif call.data == "to_jpg":
            out_dir = Path(tempfile.mkdtemp(dir=SET.temp_dir))
            files = await run_blocking(pdf_to_jpg, pdf_path, out_dir)
            zip_path = pdf_path.with_suffix(".zip")
            with zipfile.ZipFile(zip_path, "w") as z:
                for f in files:
                    z.write(f, arcname=f.name)
            await call.message.answer_document(FSInputFile(zip_path), caption="JPG архив готов ✅")
        elif call.data == "split":
            out_dir = Path(tempfile.mkdtemp(dir=SET.temp_dir))
            files = await run_blocking(pdf_split, pdf_path, out_dir)
            zip_path = pdf_path.with_suffix("_split.zip")
            with zipfile.ZipFile(zip_path, "w") as z:
                for f in files:
                    z.write(f, arcname=f.name)
            await call.message.answer_document(FSInputFile(zip_path), caption="Разбито ✅")
        elif call.data == "compress":
            out = pdf_path.with_name(pdf_path.stem + "_compressed.pdf")
            await run_blocking(pdf_compress, pdf_path, out)
            await call.message.answer_document(FSInputFile(out), caption="Сжатый PDF ✅")
        elif call.data == "watermark":
            out = pdf_path.with_name(pdf_path.stem + "_wm.pdf")
            await run_blocking(pdf_watermark, pdf_path, out, "BOT")
            await call.message.answer_document(FSInputFile(out), caption="Водяной знак ✅")
        elif call.data == "ocr":
            out = pdf_path.with_name(pdf_path.stem + "_ocr.pdf")
            await run_blocking(pdf_ocr, pdf_path, out)
            await call.message.answer_document(FSInputFile(out), caption="OCR PDF ✅")

    await call.answer()
    await state.clear()

# --------------------------
# FastAPI app with lifespan
# --------------------------
@asynccontextmanager
async def lifespan(app: FastAPI):
    global redis_pool
    # startup
    redis_pool = aioredis.from_url(SET.redis_url, decode_responses=True)
    await bot.set_webhook(
        url=f"{SET.webhook_base.rstrip('/')}{SET.webhook_path}",
        secret_token=SET.webhook_secret or None,
        drop_pending_updates=True,
    )
    yield
    # shutdown
    if redis_pool:
        await redis_pool.close()

app = FastAPI(title="Telegram Bot", lifespan=lifespan)

@app.get("/healthz")
async def healthz():
    return {"ok": True}

@app.post(SET.webhook_path)
async def telegram_webhook(request: Request):
    secret = request.headers.get("X-Telegram-Bot-Api-Secret-Token")
    if SET.webhook_secret and secret != SET.webhook_secret:
        return JSONResponse(status_code=403, content={"detail": "bad secret"})
    data = await request.json()
    update = dp.update_factory(data)
    await dp.feed_update(bot, update)
    return Response(status_code=200)

