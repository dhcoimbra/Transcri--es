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
    
    # Realizar a detecção de objetos
    with torch.no_grad():
        outputs = model(**inputs)

    # Definir rótulos manualmente se necessário
    if not hasattr(model.config, "id2label"):
        model.config.id2label = {
            0: "N/A",           # Alguns modelos têm um rótulo 'N/A' para objetos não identificados
            1: "car",
            2: "money",
            3: "house",
            4: "animal",
            5: "weapon",
            6: "drug",
            7: "gun",
            8: "ammunition",
            9: "projectile",
            10: "cocaine",
            11: "marihuana",
            12: "telephone",
            13: "cell phone",
            14: "person",
            
            # Adicione ou ajuste rótulos conforme necessário para o modelo que está usando
        }

    target_labels = ["car", "money", "house", "animal", "weapon", "drug","gun","ammunition","projectile","cocaine","marihuana","telephone","cell phone","person"]
    detected_objects = []

    for logit, box in zip(outputs.logits[0], outputs.pred_boxes[0]):
        prob = logit.softmax(-1)
        label_index = prob.argmax()
        label = model.config.id2label.get(label_index.item(), "Unknown")
        score = prob[label_index].item()
        
        # Filtrar por objetos de interesse e pontuação mínima de confiança
        if label in target_labels and score > 0.7:
            detected_objects.append((label, score, box.tolist()))

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
                label, score, box = obj
                st.write(f"- {label.capitalize()} (Confiança: {score:.2f})")
        else:
            st.write("Nenhum objeto relevante detectado nesta imagem.")
