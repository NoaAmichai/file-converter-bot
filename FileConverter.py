from pdf2docx import Converter as PDFConverter
from docx2pdf import convert as docx_to_pdf
from PIL import Image
import os
import comtypes.client

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
            if img.mode != 'RGB':
                img = img.convert('RGB')
            img.save(self.output_path, "PNG")

class PngToJpgConverter(FileConverter):
    def convert(self):
        with Image.open(self.input_path) as img:
            if img.mode != 'RGB':
                img = img.convert('RGB')
            img.save(self.output_path, "JPG")
            
class TiffToJpgConverter(FileConverter):
    def convert(self):
        with Image.open(self.input_path) as img:
            if img.mode != 'RGB':
                img = img.convert('RGB')
            img.save(self.output_path, "JPG")

class JpgToTiffConverter(FileConverter):
    def convert(self):
        with Image.open(self.input_path) as img:
            if img.mode != 'RGB':
                img = img.convert('RGB')
            img.save(self.output_path, "TIFF")

class PngToTiffConverter(FileConverter):
    def convert(self):
        with Image.open(self.input_path) as img:
            if img.mode != 'RGB':
                img = img.convert('RGB')
            img.save(self.output_path, "TIFF")

class TiffToPngConverter(FileConverter):
    def convert(self):
        with Image.open(self.input_path) as img:
            if img.mode != 'RGB':
                img = img.convert('RGB')
            img.save(self.output_path, "PNG")
            
class PptxToPdfConverter(FileConverter):
    def __init__(self, input_path: str, output_path: str, slides_per_page: int = 1):
        super().__init__(input_path, output_path)
        self.slides_per_page = slides_per_page

def convert(self):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1  # Ensure the PowerPoint application is visible
    try:
        presentation = powerpoint.Presentations.Open(os.path.abspath(self.input_path))
        new_presentation = powerpoint.Presentations.Add()
        slides = presentation.Slides
        print(f"Total slides: {slides.Count}")
        print(f"Slides per page: {self.slides_per_page}")
        slide_per_page = self.slides_per_page
        for i in range(1, slides.Count + 1, self.slides_per_page):
            # Copy the slides one by one to the new presentation
            for j in range(self.slides_per_page):
                if i + j <= slides.Count:
                    slide = slides[i + j]
                    slide.Copy()
                    print(f"Copying slide {i + j}")                    
                    new_slide = new_presentation.Slides.Paste()[1]
                    new_slide.Design = slide.Design
        new_presentation.SaveAs(os.path.abspath(self.output_path), FileFormat=32)  # 32 for PDF format
        new_presentation.Close()
        presentation.Close()
    except Exception as e:
        print(f"An error occurred while converting PowerPoint: {e}")
    finally:
        powerpoint.Quit()

