import streamlit as st
from transformers import DetrImageProcessor, DetrForObjectDetection
from PIL import Image
import torch
import fitz  # PyMuPDF
from io import BytesIO
from docx import Document

# Carregar o modelo e o processador DETR para detecção de objetos
processor = DetrImageProcessor.from_pretrained(r"C:\modelos\detr-resnet-50")
model = DetrForObjectDetection.from_pretrained(r"C:\modelos\detr-resnet-50")

# Função para detectar objetos na imagem
def detect_objects(image):
    # Processar a imagem para o modelo
    inputs = processor(images=image, return_tensors="pt")
    outputs = model(**inputs)

    # Definir o tamanho-alvo da imagem e a confiança mínima
    target_sizes = torch.tensor([image.size[::-1]])  # Converte para altura x largura
    results = processor.post_process_object_detection(outputs, target_sizes=target_sizes, threshold=0.9)[0]

    detected_objects = []

    # Iterar pelos resultados e extrair informações das detecções
    for score, label, box in zip(results["scores"], results["labels"], results["boxes"]):
        # Arredondar coordenadas da caixa delimitadora para duas casas decimais
        box = [round(i, 2) for i in box.tolist()]
        label_name = model.config.id2label.get(label.item(), "Unknown")
        
        # Adicionar informações do objeto detectado à lista
        detected_objects.append({
            "label": label_name,
            "confidence": round(score.item(), 3),
            "box": box
        })
    
    return detected_objects

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
st.title("Detecção de Objetos em Documentos PDF e DOCX")
st.write("Carregue um documento PDF ou DOCX para detectar objetos como carros, dinheiro, casas, animais, armas e drogas nas imagens contidas no arquivo.")

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
        st.image(image, caption=f"Imagem {i + 1}", width=150)#use_column_width=True,
        st.write("Detectando objetos...")

        # Detectar objetos na imagem
        detected_objects = detect_objects(image)

        # Mostrar os objetos detectados
        if detected_objects:
            st.write("**Objetos Detectados:**")
            for obj in detected_objects:
                st.write(f"- {obj['label'].capitalize()} (Confiança: {obj['confidence']:.2f}) em {obj['box']}")
        else:
            st.write("Nenhum objeto relevante detectado nesta imagem.")
