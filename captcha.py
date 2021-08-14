from PIL import Image
import pytesseract
from pathlib import *
from dotenv import load_dotenv
import os

# необходимо для корректной работы tesseract
load_dotenv()
pytesseract.pytesseract.tesseract_cmd = os.path.join(os.getenv('TESSERACT-OCR_DIR_LOCATION'), 'tesseract.exe')


def read_captcha():
    """
    Method solve captcha using tesseracr-ocr
    :return: Returns the line containing the resolved captcha
    :rtype: str
    """
    value = Image.open(Path.cwd() / 'img.png')
    text = pytesseract.image_to_string(value, lang='rus', config='')
    text = text.replace('\n', '').replace(' ', '').lower()

    return text[0:5]

