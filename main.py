from typing import Final
from telegram import ReplyKeyboardMarkup, Update
from telegram.ext import Application, CommandHandler, filters, ContextTypes, MessageHandler, ConversationHandler
import os
import aiohttp
from FileConverter import FileConverter, PptxToPdfConverter,DocxToPdfConverter, PdfToDocxConverter, JpgToPngConverter, PngToJpgConverter, TiffToJpgConverter, JpgToTiffConverter, PngToTiffConverter, TiffToPngConverter

TOKEN: Final = "7059338358:AAEd_Uugl2F4Yxc7wfGqx42AlKcbS1kSoRI"
BOT_USERNAME: Final = "@convert_master_bot"


# States for the conversation handler
SELECT_FORMAT, RECEIVE_FILE, SELECT_SLIDES_PER_PAGE = range(3)

def get_converter(original_format: str, target_format: str, input_path: str, output_path: str, slides_per_page: int = 1) -> FileConverter:
    converters = {
        ("docx", "pdf"): DocxToPdfConverter,
        ("pdf", "docx"): PdfToDocxConverter,
        ("jpg", "png"): JpgToPngConverter,
        ("png", "jpg"): PngToJpgConverter,
        ("tiff", "jpg"): TiffToJpgConverter,
        ("jpg", "tiff"): JpgToTiffConverter,
        ("png", "tiff"): PngToTiffConverter,
        ("tiff", "png"): TiffToPngConverter,
        ("pptx", "pdf"): lambda input_path, output_path: PptxToPdfConverter(input_path, output_path, slides_per_page),
        ("ppt", "pdf"): lambda input_path, output_path: PptxToPdfConverter(input_path, output_path, slides_per_page),
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

    context.user_data['local_file_path'] = local_file_path  # Store the file path for later use
    await update.message.reply_text("File saved. How many slides per page do you want?")
    return SELECT_SLIDES_PER_PAGE


async def handle_slides_per_page(update: Update, context: ContextTypes.DEFAULT_TYPE):
    slides_per_page = int(update.message.text)
    context.user_data['slides_per_page'] = slides_per_page
    local_file_path = context.user_data.get('local_file_path')
    
    await perform_conversion(update, context, local_file_path)

    os.remove(local_file_path)
    await update.message.reply_text("Conversion complete. If you want to convert another file, type /convert.")
    return ConversationHandler.END

async def perform_conversion(update: Update, context: ContextTypes.DEFAULT_TYPE, local_file_path: str):
    target_format = context.user_data.get('format').lower()
    original_format = local_file_path.split('.')[-1].lower()
    output_path = local_file_path.replace(f".{original_format}", f".{target_format}")
    slides_per_page = context.user_data.get('slides_per_page', 1)

    try:
        converter = get_converter(original_format, target_format, local_file_path, output_path, slides_per_page)
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
    try:
        print(f"Update {update} caused error {context.error}")
    except UnicodeEncodeError:
        print(f"Update caused error: UnicodeEncodeError")


if __name__ == "__main__":
    print("Bot is running...")
    app = Application.builder().token(TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler('convert', convert_command)],
        states={
            SELECT_FORMAT: [MessageHandler(filters.TEXT & ~filters.COMMAND, select_format)],
            RECEIVE_FILE: [MessageHandler(filters.Document.ALL | filters.PHOTO, handle_file)],
            SELECT_SLIDES_PER_PAGE: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_slides_per_page)],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
    )

    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(conv_handler)
    app.add_error_handler(error)

    app.run_polling(poll_interval=3)