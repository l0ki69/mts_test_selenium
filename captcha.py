from PIL import Image
import pytesseract
from pathlib import *

pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'


def read_captcha():

    value = Image.open(Path.cwd() / 'captcha' / 'img.png')
    text = pytesseract.image_to_string(value, lang='rus', config='')
    text = text.replace('\n', '').replace(' ', '').lower()

    return text[0:5]

