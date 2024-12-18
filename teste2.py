import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\tess\tesseract.exe'
print("Caminho do executável Tesseract configurado:")
print(pytesseract.pytesseract.tesseract_cmd)
import pytesseract
from PIL import Image

pytesseract.pytesseract.tesseract_cmd = r'C:\tess\tesseract.exe'
text = pytesseract.image_to_string(Image.open(r'C:\Users\00027336\OneDrive - PROGEN S.A\Área de Trabalho\Power BI\Imagens\avanco fisico.png'),lang="por")

print(text)

# Verificar se o executável está acessível
import os
if not os.path.exists(pytesseract.pytesseract.tesseract_cmd):
    print("Erro: O caminho configurado para o Tesseract não é válido!")
else:
    print("Tesseract está configurado corretamente!")


