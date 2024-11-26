import streamlit as st
from utils import verificar_arquivos_na_pasta, processar_em_lotes, recriar_documento_final # Certifique-se de que está importando de 'utils'
from streamlit_option_menu import option_menu
from db import criar_tabela_transcricoes
import os
import utils

with st.sidebar:
    selected = option_menu(
        menu_title="GRE",
        options=["Conversão Cellebrite","Consulta números", ],
        icons=["house", "bi-chat-dots"],
        menu_icon="cast",
        default_index=0
    )

if selected == "Conversão Cellebrite":
    st.header('Conversão Cellebrite', divider='orange')

    # Criar a tabela de transcrições no PostgreSQL
    criar_tabela_transcricoes()

    # Função para redefinir o estado da aplicação
    @st.cache_data
    def reset_state():
        for key in st.session_state.keys():
            del st.session_state[key]

    # Configuração da interface com Streamlit
    st.subheader("Conversão do Excel do Cellebrite em Word (Com transcrições)")

    # Solicitar ao usuário que insira o caminho completo da pasta de áudios e imagens
    audio_folder = st.text_input("Digite o caminho completo da pasta onde os áudios e imagens estão localizados:")
    is_folder = st.checkbox("Conversa sem anexos.")

    # Upload do arquivo Excel
    uploaded_file = st.file_uploader("Faça o upload do arquivo Excel", type=["xlsx"])

    # Seleção das tags
    #is_tag = False
    is_tag = st.checkbox("Considerar apenas mensagens com tag.")

    # Botão para iniciar a transcrição
    if st.button("Iniciar transcrição"):
        # Garantir que ambos, arquivo Excel e pasta de áudio/imagem, estão disponíveis
        if uploaded_file is not None and (audio_folder or is_folder is True):
            st.write(f"Arquivo Excel e pasta de áudios/imagens carregados com sucesso!")

            # Barra de progresso para o processo completo
            progress_bar = st.progress(0)

            # Processar o Excel e gerar os documentos de lote
            documentos_gerados = utils.processar_em_lotes(uploaded_file, audio_folder, progress_bar, is_tag)

            # Verificar se a pasta contém os arquivos referenciados no Excel
            if verificar_arquivos_na_pasta(uploaded_file, audio_folder):
                # Unir todos os documentos de lote em um único documento final
                doc_final_path = utils.recriar_documento_final(documentos_gerados, audio_folder)

                # Exibir link de download para o documento final
                with open(doc_final_path, "rb") as f:
                    st.download_button(
                        label="Baixar documento final",
                        data=f,
                        file_name=doc_final_path,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

                # Limpar arquivos temporários dos lotes
                for doc_path in documentos_gerados:
                    os.remove(doc_path)
                st.session_state['doc_final_path'] = doc_final_path    
                
                    
                    
            else:
                st.warning("A pasta fornecida não contém os arquivos de mídia mencionados no arquivo Excel.")
            
          
        else:
            st.warning("Por favor, faça o upload do arquivo Excel e insira o caminho dos áudios ou marque a opção sem anexo.")
    if "doc_final_path" in st.session_state and st.button("Gerar anonimização"):
        
        anonimizado_path = os.path.join(os.getcwd(), "Anonimizado.docx")  # Cria um caminho absoluto para salvar o documento anonimizado
        utils.anonimizar_interlocutores(st.session_state['doc_final_path'], anonimizado_path)

        # Exibir link de download para o documento anonimizado
        with open(anonimizado_path, "rb") as f:
            st.download_button(
                label="Baixar documento anonimizado",
                data=f,
                file_name="Anonimizado.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    
if selected == "Consulta números":
    st.title("Números Qlik")
    st.write("Esta é a página de Números Qlik.")


# Botão para limpar a tela
#if st.button("Limpar tela"):
#    reset_state()