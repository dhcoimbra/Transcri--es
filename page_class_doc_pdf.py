import streamlit as st
from transformers import ViTImageProcessor, ViTForImageClassification
from PIL import Image
import torch
import pytesseract
import fitz  # PyMuPDF
from io import BytesIO
from docx import Document

# Carregar o modelo e o processador treinados
processor = ViTImageProcessor.from_pretrained(r'C:\modelos\imagens\vit_comprovante')
model = ViTForImageClassification.from_pretrained(r'C:\modelos\imagens\vit_comprovante')

# Função para realizar OCR na imagem
def extract_text_with_ocr(image):
    text = pytesseract.image_to_string(image)
    return text

# Função para classificar novas imagens
def classify_document(image):
    encoding = processor(image, return_tensors="pt")
    with torch.no_grad():
        outputs = model(**encoding)
        logits = outputs.logits
        predicted_class = torch.argmax(logits, dim=1).item()

    if predicted_class == 1:
        return "Comprovante Bancário", extract_text_with_ocr(image)
    else:
        return "Outro Documento", None

# Função para extrair imagens de um documento PDF
def extract_images_from_pdf(pdf_file):
    pdf_document = fitz.open(stream=pdf_file.read(), filetype="pdf")
    images = []
    for page_num in range(len(pdf_document)):
        for img in pdf_document.get_page_images(page_num):
            xref = img[0]
            base_image = pdf_document.extract_image(xref)
            image_data = base_image["image"]
            images.append(Image.open(BytesIO(image_data)))
    return images

# Função para extrair imagens de um documento DOCX
def extract_images_from_docx(docx_file):
    doc = Document(docx_file)
    images = []
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            image_data = rel.target_part.blob
            images.append(Image.open(BytesIO(image_data)))
    return images

# Estrutura de Interface no Streamlit
st.title("Classificação e OCR de Documentos em Arquivos PDF e DOCX")
st.write("Carregue um documento PDF ou DOCX para identificar e realizar OCR em comprovantes bancários.")

# Upload de Documento
uploaded_file = st.file_uploader("Escolha um documento PDF ou DOCX...", type=["pdf", "docx"])

if uploaded_file is not None:
    file_type = uploaded_file.type
    images = []

    # Extrair imagens do documento
    if file_type == "application/pdf":
        images = extract_images_from_pdf(uploaded_file)
    elif file_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        images = extract_images_from_docx(uploaded_file)

    # Processar cada imagem extraída
    for i, image in enumerate(images):
        st.image(image, caption=f"Imagem {i + 1}", use_column_width=True)
        st.write("Classificando...")

        # Classificar a imagem
        result, ocr_text = classify_document(image)

        # Mostrar o resultado da classificação
        st.write("**Resultado da Classificação:**", result)

        # Exibir o texto extraído via OCR, se for um comprovante bancário
        if ocr_text:
            st.write("**Texto Extraído via OCR:**")
            st.text_area(f"Texto do Comprovante {i + 1}", ocr_text, height=200)
        else:
            st.write("Nenhum texto extraído, pois não é um comprovante bancário.")
