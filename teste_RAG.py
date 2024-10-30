import streamlit as st
from transformers import DetrImageProcessor, DetrForObjectDetection, pipeline
from sentence_transformers import SentenceTransformer, util
from transformers import AutoModelForSeq2SeqLM, AutoTokenizer
import fitz  # PyMuPDF
from docx import Document
from PIL import Image
import torch
from io import BytesIO

# Carregar os modelos necessários
detector_processor = DetrImageProcessor.from_pretrained("facebook/detr-resnet-50")
detector_model = DetrForObjectDetection.from_pretrained("facebook/detr-resnet-50")
retrieval_model = SentenceTransformer("all-MiniLM-L6-v2")
qa_tokenizer = AutoTokenizer.from_pretrained("t5-small")
qa_model = AutoModelForSeq2SeqLM.from_pretrained("t5-small")

# Função para extrair texto de PDF
def extract_text_from_pdf(pdf_file):
    pdf_document = fitz.open(stream=pdf_file.read(), filetype="pdf")
    text = ""
    for page_num in range(len(pdf_document)):
        page = pdf_document[page_num]
        text += page.get_text("text")
    return text

# Função para extrair texto de DOCX
def extract_text_from_docx(docx_file):
    doc = Document(docx_file)
    text = "\n".join([para.text for para in doc.paragraphs if para.text])
    return text

# Função para criar embeddings para recuperação de passagens
def create_embeddings(text):
    passages = text.split("\n")
    embeddings = retrieval_model.encode(passages, convert_to_tensor=True)
    return passages, embeddings

# Função para buscar a passagem mais relevante para uma pergunta
def retrieve_passage(question, passages, embeddings):
    question_embedding = retrieval_model.encode(question, convert_to_tensor=True)
    scores = util.pytorch_cos_sim(question_embedding, embeddings)[0]
    best_index = scores.argmax()
    return passages[best_index]

# Função para gerar uma resposta usando o modelo de QA
def generate_answer(question, passage):
    input_text = f"question: {question} context: {passage}"
    input_ids = qa_tokenizer.encode(input_text, return_tensors="pt")
    outputs = qa_model.generate(input_ids)
    answer = qa_tokenizer.decode(outputs[0], skip_special_tokens=True)
    return answer

# Função para a lógica de perguntas e respostas
def answer_question(question, text):
    passages, embeddings = create_embeddings(text)
    relevant_passage = retrieve_passage(question, passages, embeddings)
    answer = generate_answer(question, relevant_passage)
    return answer

# Estrutura de Interface no Streamlit
st.title("RAG: Perguntas e Respostas sobre Documentos PDF e DOCX")
st.write("Carregue um documento PDF ou DOCX e faça perguntas sobre seu conteúdo.")

# Upload de Documento
uploaded_file = st.file_uploader("Escolha um documento PDF ou DOCX...", type=["pdf", "docx"])

if uploaded_file is not None:
    file_type = uploaded_file.type
    text_content = ""

    # Extrair texto do documento
    if file_type == "application/pdf":
        text_content = extract_text_from_pdf(uploaded_file)
    elif file_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        text_content = extract_text_from_docx(uploaded_file)

    # Exibir o conteúdo do documento e permitir perguntas
    if text_content:
        st.write("**Conteúdo do Documento Carregado:**")
        st.text_area("Texto Extraído", text_content, height=250)

        # Entrada para pergunta do usuário
        question = st.text_input("Faça uma pergunta sobre o documento:")

        if question:
            answer = answer_question(question, text_content)
            st.write("**Resposta:**", answer)
