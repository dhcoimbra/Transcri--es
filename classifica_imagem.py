from transformers import ViTImageProcessor, ViTForImageClassification
from PIL import Image
import torch
import pytesseract

# Carregar o modelo e o processador treinados
processor = ViTImageProcessor.from_pretrained(r'C:\modelos\imagens\vit_comprovante')
model = ViTForImageClassification.from_pretrained(r'C:\modelos\imagens\vit_comprovante')

# Função para realizar OCR na imagem
def extract_text_with_ocr(image):
    # Usando o pytesseract para extrair o texto
    text = pytesseract.image_to_string(image)
    return text

# Função para classificar novas imagens
def classify_document(image_path):
    # Carregar e processar a imagem
    image = Image.open(image_path).convert("RGB")
    encoding = processor(image, return_tensors="pt")

    # Realizar a inferência
    with torch.no_grad():
        outputs = model(**encoding)
        logits = outputs.logits
        predicted_class = torch.argmax(logits, dim=1).item()

    # Interpretação da previsão
    if predicted_class == 1:
        print("Documento identificado como: Comprovante Bancário")
        # Realizar OCR se for um comprovante bancário
        ocr_text = extract_text_with_ocr(image)
        print("Texto extraído via OCR:")
        print(ocr_text)
        return "Comprovante Bancário", ocr_text
    else:
        print("Documento identificado como: Outro Documento")
        return "Outro Documento", None

# Exemplo de uso
result, ocr_text = classify_document(r"C:\modelos\imagens\comprovante\3c711943-a7db-4875-8e7f-63c0e6c3ef06.jpg")
print("Resultado da classificação:", result)
if ocr_text:
    print("Texto OCR:", ocr_text)
