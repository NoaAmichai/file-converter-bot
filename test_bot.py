import pytest
import asyncio
from unittest.mock import AsyncMock, patch, mock_open
from telegram import Update, Document, PhotoSize, File
from telegram.ext import ContextTypes

from main import (
    start_command, help_command, convert_command, select_format, handle_file,
    cancel, perform_conversion, get_converter, DocxToPdfConverter, PdfToDocxConverter
)

@pytest.fixture
def mock_update():
    return AsyncMock(spec=Update)

@pytest.fixture
def mock_context():
    return AsyncMock(spec=ContextTypes.DEFAULT_TYPE)

@pytest.mark.asyncio
async def test_start_command(mock_update, mock_context):
    await start_command(mock_update, mock_context)
    mock_update.message.reply_text.assert_called_once_with(
        "Welcome to Convert Master Bot 🤖\n\n"
        "I can convert any file to any format you want.\n\n"
        "👉🏻 For more information type /help"
    )

@pytest.mark.asyncio
async def test_help_command(mock_update, mock_context):
    await help_command(mock_update, mock_context)
    mock_update.message.reply_text.assert_called_once_with(
        "To use this bot, please enter /convert, select the format you want to convert the file to, "
        "and then send me the file you want to convert."
    )

@pytest.mark.asyncio
async def test_convert_command(mock_update, mock_context):
    await convert_command(mock_update, mock_context)
    mock_update.message.reply_text.assert_called_once()

@pytest.mark.asyncio
async def test_select_format(mock_update, mock_context):
    mock_update.message.text = "📄 PDF"
    await select_format(mock_update, mock_context)
    assert mock_context.user_data['format'] == "PDF"
    mock_update.message.reply_text.assert_called_once_with(
        "Format selected: PDF\nNow, please send me the file you want to convert."
    )

@pytest.mark.asyncio
@patch("bot.aiohttp.ClientSession")
@patch("bot.os.makedirs")
@patch("bot.open", new_callable=mock_open)
async def test_handle_file(mock_open, mock_makedirs, mock_ClientSession, mock_update, mock_context):
    # Mock the file download and telegram file object
    mock_file = AsyncMock(spec=File)
    mock_file.file_path = "path/to/file.docx"
    mock_update.message.document = AsyncMock(spec=Document)
    mock_update.message.document.file_id = "file_id"
    mock_update.message.photo = None
    mock_context.bot.get_file.return_value = mock_file

    mock_response = AsyncMock()
    mock_response.status = 200
    mock_response.read.return_value = b"file_content"
    mock_ClientSession.return_value.__aenter__.return_value.get.return_value = mock_response

    mock_context.user_data['format'] = 'PDF'

    await handle_file(mock_update, mock_context)
    mock_update.message.reply_text.assert_called_with(
        "File received and saved as downloads/file_id.docx"
    )

@pytest.mark.asyncio
@patch("bot.get_converter")
@patch("bot.os.remove")
async def test_perform_conversion(mock_os_remove, mock_get_converter, mock_update, mock_context):
    mock_context.user_data['format'] = "PDF"
    local_file_path = "downloads/file.docx"
    mock_converter = AsyncMock(spec=DocxToPdfConverter)
    mock_get_converter.return_value = mock_converter

    await perform_conversion(mock_update, mock_context, local_file_path)

    mock_get_converter.assert_called_once_with("docx", "pdf", local_file_path, "downloads/file.pdf")
    mock_converter.convert.assert_called_once()
    mock_update.message.reply_document.assert_called_once()
    mock_os_remove.assert_called_once_with("downloads/file.pdf")

def test_get_converter():
    input_path = "input.docx"
    output_path = "output.pdf"
    converter = get_converter("docx", "pdf", input_path, output_path)
    assert isinstance(converter, DocxToPdfConverter)

    converter = get_converter("pdf", "docx", input_path, output_path)
    assert isinstance(converter, PdfToDocxConverter)

    with pytest.raises(ValueError):
        get_converter("unsupported", "format", input_path, output_path)
