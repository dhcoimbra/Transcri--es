import streamlit as st
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
def classify_document(image):
    encoding = processor(image, return_tensors="pt")

    # Realizar a inferência
    with torch.no_grad():
        outputs = model(**encoding)
        logits = outputs.logits
        predicted_class = torch.argmax(logits, dim=1).item()

    # Interpretação da previsão
    if predicted_class == 1:
        return "Comprovante Bancário", extract_text_with_ocr(image)
    else:
        return "Outro Documento", None

# Estrutura de Interface no Streamlit
st.title("Classificação e OCR de Documentos")
st.write("Carregue uma imagem para classificar o tipo de documento e, se for um comprovante bancário, realizar OCR.")

# Upload de Imagem
uploaded_file = st.file_uploader("Escolha uma imagem...", type=["jpg", "jpeg", "png"])

if uploaded_file is not None:
    # Abrir a imagem usando PIL
    image = Image.open(uploaded_file).convert("RGB")
    st.image(image, caption="Imagem Carregada", use_column_width=True)
    st.write("Classificando...")

    # Classificar a imagem
    result, ocr_text = classify_document(image)

    # Mostrar o resultado da classificação
    st.write("**Resultado da Classificação:**", result)

    # Exibir o texto extraído via OCR, se for um comprovante bancário
    if ocr_text:
        st.write("**Texto Extraído via OCR:**")
        st.text_area("Texto do Comprovante", ocr_text, height=200)
    else:
        st.write("Nenhum texto extraído, pois não é um comprovante bancário.")
