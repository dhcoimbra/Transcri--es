from transformers import ViTImageProcessor, ViTForImageClassification
from PIL import Image
import torch

# Carregar o modelo e o processador treinados
processor = ViTImageProcessor.from_pretrained(r'C:\modelos\imagens\vit_comprovante')
model = ViTForImageClassification.from_pretrained(r'C:\modelos\imagens\vit_comprovante')
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
        return "Comprovante Bancário"
    else:
        return "Outro Documento"

# Exemplo de uso
result = classify_document(r"C:\modelos\imagens\comprovante\401587d0-98f2-4360-8e28-b2afd6950582.jpg")
print("Resultado da classificação:", result)
