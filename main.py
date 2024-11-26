import streamlit as st
from utils import verificar_arquivos_na_pasta, processar_em_lotes, recriar_documento_final, selecionar_pasta, localizar_arquivo_excel, localizar_subpasta_com_arquivo, anonimizar_interlocutores
from db import criar_tabela_transcricoes
import os

st.header('Conversão Cellebrite', divider='orange')

# Criar a tabela de transcrições no PostgreSQL
criar_tabela_transcricoes()

# Redefinir o estado da aplicação
@st.cache_data
def reset_state():
    st.session_state.clear()

# Configuração da interface com Streamlit
st.subheader("Conversão do Excel do Cellebrite em Word (Com transcrições)")

# Checkbox de configurações
is_folder = st.checkbox("Conversa sem anexos.")
is_tag = st.checkbox("Considerar apenas mensagens com tag.")

# Botão para abrir o seletor de pastas
if "pasta_selecionada" not in st.session_state:
    st.session_state["pasta_selecionada"] = None

if st.button("Selecionar Pasta"):
    pasta = selecionar_pasta()
    if pasta:
        st.session_state["pasta_selecionada"] = pasta
        st.success(f"Pasta selecionada: {pasta}")
    else:
        st.warning("Nenhuma pasta foi selecionada.")

# Verificar se a pasta foi selecionada
if st.session_state["pasta_selecionada"]:
    st.success(f"Pasta selecionada: {st.session_state['pasta_selecionada']}")

    # Localizar o arquivo Excel
    if "arquivo_excel" not in st.session_state:
        st.session_state["arquivo_excel"] = localizar_arquivo_excel(st.session_state["pasta_selecionada"])

    if st.session_state["arquivo_excel"]:
        st.success(f"Arquivo Excel encontrado: {st.session_state['arquivo_excel']}")

        # Localizar a subpasta com os arquivos de mídia
        if "audio_folder" not in st.session_state:
            st.session_state["audio_folder"] = localizar_subpasta_com_arquivo(
                st.session_state["arquivo_excel"], st.session_state["pasta_selecionada"]
            )

        if st.session_state["audio_folder"]:
            st.success(f"Subpasta de áudio/imagens encontrada: {st.session_state['audio_folder']}")

            # Botão para iniciar a transcrição
            if st.button("Iniciar Transcrição"):
                st.write("Iniciando transcrição...")

                progress_bar = st.progress(0)

                # Processar o Excel e gerar os documentos de lote
                documentos_gerados = processar_em_lotes(
                    st.session_state["arquivo_excel"],
                    st.session_state["audio_folder"],
                    progress_bar,
                    is_tag
                )

                # Verificar se a pasta contém os arquivos referenciados no Excel
                if verificar_arquivos_na_pasta(
                    st.session_state["arquivo_excel"], st.session_state["audio_folder"]
                ):
                    # Unir todos os documentos de lote em um único documento final
                    doc_final_path = recriar_documento_final(
                        documentos_gerados, st.session_state["audio_folder"]
                    )
                    st.session_state["doc_final_path"] = doc_final_path


                    if os.path.exists(doc_final_path):
                        with open(doc_final_path, "rb") as f:
                            file_data = f.read()
                            st.download_button(
                                label="Baixar Documento Final",
                                data=file_data,
                                file_name="Documento_Final.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                    else:
                        st.error("Erro: O arquivo final não foi encontrado.")


                    # Limpar arquivos temporários dos lotes
                    for doc_path in documentos_gerados:
                        os.remove(doc_path)
                else:
                    st.warning("A pasta fornecida não contém os arquivos mencionados no Excel.")

            # Botão para gerar anonimização
            if "doc_final_path" in st.session_state and st.button("Gerar Anonimização"):
                anonimizado_path = os.path.join(os.getcwd(), "Anonimizado.docx")
                anonimizar_interlocutores(st.session_state["doc_final_path"], anonimizado_path)

                # Exibir link de download para o documento anonimizado
                with open(anonimizado_path, "rb") as f:
                    st.download_button(
                        label="Baixar Documento Anonimizado",
                        data=f,
                        file_name="Anonimizado.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
        else:
            st.warning("Nenhuma subpasta com arquivos foi encontrada.")
    else:
        st.warning("Nenhum arquivo Excel chamado 'Relatório' ou 'Report' foi encontrado.")
else:
    st.warning("Por favor, selecione uma pasta.")
