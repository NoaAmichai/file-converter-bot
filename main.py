from typing import Final
from telegram import ReplyKeyboardMarkup, Update
from telegram.ext import Application, CommandHandler, filters, ContextTypes, MessageHandler, ConversationHandler
import os
import aiohttp
from pdf2docx import Converter as PDFConverter
from docx2pdf import convert as docx_to_pdf
from PIL import Image
import sys
import comtypes.client

TOKEN: Final = "7059338358:AAEd_Uugl2F4Yxc7wfGqx42AlKcbS1kSoRI"
BOT_USERNAME: Final = "@convert_master_bot"

# States for the conversation handler
SELECT_FORMAT, RECEIVE_FILE = range(2)

class FileConverter:
    def __init__(self, input_path: str, output_path: str):
        self.input_path = input_path
        self.output_path = output_path

    def convert(self):
        raise NotImplementedError("Subclasses should implement this method")

class DocxToPdfConverter(FileConverter):
    def convert(self):
        docx_to_pdf(self.input_path, self.output_path)

class PdfToDocxConverter(FileConverter):
    def convert(self):
        pdf_converter = PDFConverter(self.input_path)
        pdf_converter.convert(self.output_path)
        pdf_converter.close()

class JpgToPngConverter(FileConverter):
    def convert(self):
        with Image.open(self.input_path) as img:
            # Convert to RGB mode if it's not already in RGB
            if img.mode != 'RGB':
                img = img.convert('RGB')
            img.save(self.output_path, "PNG")

class PngToJpgConverter(FileConverter):
    def convert(self):
        with Image.open(self.input_path) as img:
            img.save(self.output_path, "JPG")
            
class TiffToJpgConverter(FileConverter):
    def convert(self):
        with Image.open(self.input_path) as img:
            img.save(self.output_path, "JPG")

class JpgToTiffConverter(FileConverter):
    def convert(self):
        with Image.open(self.input_path) as img:
            img.save(self.output_path, "TIFF")

class PngToTiffConverter(FileConverter):
    def convert(self):
        with Image.open(self.input_path) as img:
            img.save(self.output_path, "TIFF")

class TiffToPngConverter(FileConverter):
    def convert(self):
        with Image.open(self.input_path) as img:
            img.save(self.output_path, "PNG")
            
class pptxToPdfConverter(FileConverter):
    def convert(self):
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        ppt = powerpoint.Presentations.Open(self.input_path)
        ppt.SaveAs(self.output_path, 32)
        ppt.Close()
        powerpoint.Quit()
    

def get_converter(original_format: str, target_format: str, input_path: str, output_path: str) -> FileConverter:
    converters = {
        ("docx", "pdf"): DocxToPdfConverter,
        ("pdf", "docx"): PdfToDocxConverter,
        ("jpg", "png"): JpgToPngConverter,
        ("png", "jpg"): PngToJpgConverter,
        ("tiff", "jpg"): TiffToJpgConverter,
        ("jpg", "tiff"): JpgToTiffConverter,
        ("png", "tiff"): PngToTiffConverter,
        ("tiff", "png"): TiffToPngConverter,
        ("pptx", "pdf"): pptxToPdfConverter
    }
    converter_class = converters.get((original_format, target_format))
    if converter_class:
        return converter_class(input_path, output_path)
    else:
        raise ValueError("Unsupported conversion format")

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Welcome to Convert Master Bot 🤖\n\n"
        "I can convert any file to any format you want.\n\n"
        "👉🏻 For more information type /help"
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "To use this bot, please enter /convert, select the format you want to convert the file to, "
        "and then send me the file you want to convert."
    )

async def convert_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    buttons = [
        ["PDF", "DOCX"],
        ["JPG", "PNG"],
        ["PPTX", "XLSX"]
    ]
    button_pictures = [
        ["📄", "📝"],
        ["🖼️", "🖼️"],
        ["📊", "📊"]
    ]
    buttons_with_pictures = [
        [f"{button_pictures[i][j]} {buttons[i][j]}" for j in range(len(buttons[i]))]
        for i in range(len(buttons))
    ]
    await update.message.reply_text(
        "Select the format you want to convert the file to.",
        reply_markup=ReplyKeyboardMarkup(buttons_with_pictures, one_time_keyboard=True)
    )
    return SELECT_FORMAT

async def select_format(update: Update, context: ContextTypes.DEFAULT_TYPE):
    format_selected = update.message.text.split()[1]  # Extract the format from the message text
    context.user_data['format'] = format_selected
    await update.message.reply_text(f"Format selected: {format_selected}\nNow, please send me the file you want to convert.")
    return RECEIVE_FILE

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file = None
    file_type = None
    file_name = None

    if update.message.document:
        file = await context.bot.get_file(update.message.document.file_id)
        file_type = 'document'
        file_name = update.message.document.file_name.split('.')[0]
    elif update.message.photo:
        file = await context.bot.get_file(update.message.photo[-1].file_id)
        file_type = 'photo'
        file_name = update.message.photo[-1].file_id

    if not file:
        await update.message.reply_text("Please send a valid document or photo.")
        return RECEIVE_FILE

    file_path = file.file_path
    print(f"Received {file_type}: {file_name}")
    local_file_path = f"downloads/{file_name}.{file_path.split('.')[-1]}"
    os.makedirs(os.path.dirname(local_file_path), exist_ok=True)

    async with aiohttp.ClientSession() as session:
        async with session.get(file_path) as response:
            if response.status == 200:
                with open(local_file_path, 'wb') as f:
                    f.write(await response.read())
                await update.message.reply_text(f"File received and saved as {file_name}")
            else:
                await update.message.reply_text("Failed to download the file.")
                return RECEIVE_FILE

    await perform_conversion(update, context, local_file_path)

    os.remove(local_file_path)
    await update.message.reply_text("Conversion complete. If you want to convert another file, type /convert.")
    return ConversationHandler.END

async def perform_conversion(update: Update, context: ContextTypes.DEFAULT_TYPE, local_file_path: str):
    target_format = context.user_data.get('format').lower()
    original_format = local_file_path.split('.')[-1].lower()
    output_path = local_file_path.replace(f".{original_format}", f".{target_format}")

    try:
        converter = get_converter(original_format, target_format, local_file_path, output_path)
        converter.convert()
        if target_format in ['jpg', 'png']:
            await update.message.reply_photo(open(output_path, 'rb'))
        else:
            await update.message.reply_document(open(output_path, 'rb'))
        os.remove(output_path)
    except ValueError as e:
        await update.message.reply_text(str(e))

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Conversion cancelled. To start again, type /convert.")
    return ConversationHandler.END

async def error(update: Update, context: ContextTypes.DEFAULT_TYPE):
    print(f"Update {update} caused error {context.error}")

if __name__ == "__main__":
    print("Bot is running...")
    app = Application.builder().token(TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler('convert', convert_command)],
        states={
            SELECT_FORMAT: [MessageHandler(filters.TEXT & ~filters.COMMAND, select_format)],
            RECEIVE_FILE: [MessageHandler(filters.Document.ALL | filters.PHOTO, handle_file)],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
    )

    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(conv_handler)
    app.add_error_handler(error)

    app.run_polling(poll_interval=3)