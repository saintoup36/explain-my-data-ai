import os
import pytesseract

print("STARTING TEST")

TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
print("CHECKING PATH:", TESSERACT_PATH)
print("PATH EXISTS:", os.path.exists(TESSERACT_PATH))

if os.path.exists(TESSERACT_PATH):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
    print("TESSERACT CMD SET")

print("ABOUT TO CHECK VERSION")
print(pytesseract.get_tesseract_version())

print("TEST FINISHED")
input("Press Enter to close...")