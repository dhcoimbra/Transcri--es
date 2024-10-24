import streamlit as st
from utils import process_excel, verificar_arquivos_na_pasta # Certifique-se de que está importando de 'utils'
from db import criar_tabela_transcricoes

# Criar a tabela de transcrições no PostgreSQL
criar_tabela_transcricoes()

# Função para redefinir o estado da aplicação
@st.cache_data
def reset_state():
    for key in st.session_state.keys():
        del st.session_state[key]

# Configuração da interface com Streamlit
st.set_page_config(page_title="HawkData",)
st.title("Conversão do Excel do Cellebrite em Word (Com transcrições)")

# Solicitar ao usuário que insira o caminho completo da pasta de áudios e imagens
audio_folder = st.text_input("Digite o caminho completo da pasta onde os áudios e imagens estão localizados:")

# Upload do arquivo Excel
uploaded_file = st.file_uploader("Faça o upload do arquivo Excel", type=["xlsx"])

# Botão para iniciar a transcrição
if st.button("Iniciar transcrição"):
    # Garantir que ambos, arquivo Excel e pasta de áudio/imagem, estão disponíveis
    if uploaded_file is not None and audio_folder:
        st.write(f"Arquivo Excel e pasta de áudios/imagens carregados com sucesso!")

        # Verificar se a pasta contém os arquivos referenciados no Excel
        if verificar_arquivos_na_pasta(uploaded_file, audio_folder):
            # Processar o Excel
            doc = process_excel(uploaded_file, audio_folder)
            
            # Salvar o documento .docx
            output_file = "resultado_transcricoes_imagens_tabela_com_bordas.docx"
            doc.save(output_file)

            # Disponibilizar o download do documento
            with open(output_file, "rb") as f:
                st.download_button("Baixar documento com transcrições e imagens", data=f, file_name=output_file)
        else:
            st.warning("A pasta fornecida não contém os arquivos de mídia mencionados no arquivo Excel.")
    else:
        st.warning("Por favor, faça o upload do arquivo Excel e insira o caminho dos áudios.")

# Botão para limpar a tela
#if st.button("Limpar tela"):
#    reset_state()
