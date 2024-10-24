import os
import pandas as pd
import time
from docx import Document
from docx.shared import Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from assemblyai import upload_audio, transcrever_audio_assemblyai
from PIL import UnidentifiedImageError
from db import salvar_transcricao, buscar_transcricao
import streamlit as st

# Função para adicionar bordas a todas as células da tabela
def adicionar_bordas_a_tabela(table):
    for row in table.rows:
        for cell in row.cells:
            tc_pr = cell._element.get_or_add_tcPr()
            tc_borders = OxmlElement('w:tcBorders')
            
            for border_name in ['top', 'left', 'bottom', 'right']:
                border = OxmlElement(f'w:{border_name}')
                border.set(qn('w:val'), 'single')  # Define as bordas como simples
                border.set(qn('w:sz'), '4')        # Espessura da borda
                border.set(qn('w:space'), '0')
                border.set(qn('w:color'), '000000')  # Cor preta
                tc_borders.append(border)
            
            tc_pr.append(tc_borders)

# Função para criar o documento .docx com uma tabela no formato desejado e imagens
def criar_documento_docx(df, audio_dir):
    doc = Document()

    # Adicionar tabela com 3 colunas: "From", "Body", "Timestamp-Time"
    table = doc.add_table(rows=1, cols=3)

    # Remover travamento das colunas
    table.autofit = False

    # Adicionar cabeçalho da tabela
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'From'
    hdr_cells[1].text = 'Body'
    hdr_cells[2].text = 'Timestamp-Time'

    # Aplicar bordas a toda a tabela
    adicionar_bordas_a_tabela(table)

    total_linhas = len(df)
    start_time = time.time()

    # Criar barra de progresso
    progress_bar = st.progress(0)

    # Iterar por todas as linhas do DataFrame original
    for index, row in df.iterrows():
        from_field = row['From']
        to_field = row.get('To', '')
        timestamp = row.get('Timestamp-Time', '')
        body = row.get('Body', '')
        attachment = row.get('Attachment #1', None)

        # Criar nova linha na tabela
        row_cells = table.add_row().cells

        # Preencher a coluna "From"
        row_cells[0].text = from_field

        # Verificar se existe algum anexo (áudio ou imagem) para preencher a coluna "Body"
        #print(f"Processando linha {index}: Anexo={attachment}")

        if pd.notna(attachment):
            audio_filename = os.path.basename(attachment)
            #print(f"Anexo detectado: {audio_filename}")

            if attachment.endswith('.opus'):
                transcricao_existente = buscar_transcricao(audio_filename)  # Buscar no PostgreSQL
                if transcricao_existente:
                    #print(f"Transcrição já existe no banco de dados: {audio_filename}")
                    row_cells[1].text = f"ÁUDIO\nTranscrição: {transcricao_existente}"
                else:
                    audio_path = os.path.join(audio_dir, attachment)
                    #print(f"Processando áudio: {audio_path}")

                    if os.path.exists(audio_path):
                        audio_url = upload_audio(audio_path)
                        if audio_url:
                            #print(f"Áudio enviado para transcrição: {audio_url}")
                            transcription = transcrever_audio_assemblyai(audio_url)
                            if transcription:
                                #print(f"Salvando nova transcrição no banco: {audio_filename}")
                                row_cells[1].text = f"ÁUDIO\nTranscrição: {transcription}"
                                salvar_transcricao(audio_filename, transcription, from_field, to_field, timestamp)  # Salvar no PostgreSQL
                            else:
                                print(f"Falha ao transcrever o áudio: {audio_filename}")
                        else:
                            print(f"Falha ao enviar o áudio: {audio_path}")
            elif attachment.lower().endswith(('.jpg', '.jpeg', '.png', '.gif')):
                image_path = os.path.join(audio_dir, attachment)
                if os.path.exists(image_path):
                    try:
                        #print(f"Incluindo imagem: {image_path}")
                        row_cells[1].text = ""
                        paragraph = row_cells[1].paragraphs[0]
                        run = paragraph.add_run()
                        run.add_picture(image_path, width=Cm(7))  # Largura de 7 cm
                    except UnidentifiedImageError:
                        #print(f"IMAGEM NÃO SUPORTADA: {attachment}")
                        row_cells[1].text = f"IMAGEM NÃO SUPORTADA: {attachment}"
        else:
            row_cells[1].text = body

        # Preencher a coluna "Timestamp-Time"
        row_cells[2].text = timestamp

        # Atualizar a barra de progresso
        progress_bar.progress((index + 1) / total_linhas)

    return doc


# Função para processar o Excel e gerar o documento
def process_excel(file, audio_dir):
    # Carregar o arquivo Excel
    df = pd.read_excel(file, engine='openpyxl', header=1)

    # Criar o documento .docx
    return criar_documento_docx(df, audio_dir)
