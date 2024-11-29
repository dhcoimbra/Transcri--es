import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from utils_tk import verificar_arquivos_na_pasta, processar_em_lotes, recriar_documento_final, anonimizar_interlocutores
from db import criar_tabela_transcricoes
import pandas as pd
import threading

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("HawkData 1.0")
        self.root.geometry("800x700")

        # Criação das abas
        self.tab_control = ttk.Notebook(root)
        self.main_tab = ttk.Frame(self.tab_control)
        self.second_tab = ttk.Frame(self.tab_control)

        self.tab_control.add(self.main_tab, text='Principal')
        self.tab_control.add(self.second_tab, text='Configurações')

        self.tab_control.pack(expand=1, fill='both')

        # Elementos da Aba Principal
        # Botão e campo de texto para selecionar a pasta de origem
        self.source_label = tk.Label(self.main_tab, text="Pasta exportada com relatórios:")
        self.source_label.grid(row=0, column=0, padx=10, pady=10, sticky='w')
        
        self.source_entry = tk.Entry(self.main_tab, width=50)
        self.source_entry.grid(row=0, column=1, padx=10, pady=10)
        
        self.source_button = tk.Button(self.main_tab, text="Selecionar", command=self.select_source_folder)
        self.source_button.grid(row=0, column=2, padx=10, pady=10)
        
        # Botão e campo de texto para selecionar a pasta de destino
        self.dest_label = tk.Label(self.main_tab, text="Pasta de destino:")
        self.dest_label.grid(row=1, column=0, padx=10, pady=10, sticky='w')
        
        self.dest_entry = tk.Entry(self.main_tab, width=50)
        self.dest_entry.grid(row=1, column=1, padx=10, pady=10)
        
        self.dest_button = tk.Button(self.main_tab, text="Selecionar", command=self.select_dest_folder)
        self.dest_button.grid(row=1, column=2, padx=10, pady=10)

        # Checkbox para conversa sem anexos
        self.checkbox_anexos = tk.BooleanVar()
        self.checkbox1 = tk.Checkbutton(self.main_tab, text="Conversa sem anexos.", variable=self.checkbox_anexos)
        self.checkbox1.grid(row=2, column=0, padx=10, pady=10, sticky='w')

        # Checkbox para considerar apenas mensagens com tag
        self.checkbox_tag = tk.BooleanVar()
        self.checkbox2 = tk.Checkbutton(self.main_tab, text="Considerar apenas mensagens com tag.", variable=self.checkbox_tag)
        self.checkbox2.grid(row=3, column=0, padx=10, pady=10, sticky='w')

        # Campo de monitoramento de atividades
        self.log_text = tk.Text(self.main_tab, width=80, height=15, state='disabled')
        self.log_text.grid(row=5, column=0, columnspan=3, padx=10, pady=10)

        # Barra de progresso
        self.progress_bar = ttk.Progressbar(self.main_tab, orient='horizontal', length=400, mode='determinate')
        self.progress_bar.grid(row=6, column=0, columnspan=3, padx=10, pady=10)

        # Botão Processar
        self.process_button = tk.Button(self.main_tab, text="Processar", command=self.start_thread_process)
        self.process_button.grid(row=4, column=1, pady=20)

        # Elementos da Aba Secundária
        # Checkbox na aba secundária
        self.checkbox_anonimizar = tk.BooleanVar()
        self.checkbox3 = tk.Checkbutton(self.second_tab, text="Anonimizar documento.", variable=self.checkbox_anonimizar)
        self.checkbox3.grid(row=0, column=0, padx=10, pady=10, sticky='w')

        # # Elementos da Aba Secundária (Consulta números)
        # self.info_label = tk.Label(self.second_tab, text="Esta é a página de Consulta de Números Qlik.")
        # self.info_label.grid(row=0, column=0, padx=10, pady=10, sticky='w')

    
    # # Botão para selecionar o arquivo Excel
    def select_source_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.source_entry.delete(0, tk.END)
            self.source_entry.insert(0, folder)
    
    def select_dest_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.dest_entry.delete(0, tk.END)
            self.dest_entry.insert(0, folder)

    def localizar_subpasta_com_arquivo(self,file, pasta_usuario):
        """
        Localiza a subpasta onde qualquer arquivo listado no Excel está presente.

        Args:
            file (str): Caminho do arquivo Excel.
            pasta_usuario (str): Caminho da pasta inicial para a busca.

        Returns:
            str: Caminho completo da subpasta onde o arquivo foi encontrado.
            None: Retorna None se nenhum arquivo listado foi encontrado.
        """
        # Carregar o arquivo Excel
        df = pd.read_excel(file, engine='openpyxl', header=1)

        # Filtrar linhas em branco
        df = df.dropna(how='all')

        # Iterar pelas linhas do Excel para verificar os arquivos listados
        for index, row in df.iterrows():
            attachment = row.get('Attachment #1', row.get('Anexo #1', None))
            if pd.notna(attachment):  # Se existe algum anexo
                # Percorrer subpastas e verificar se o arquivo está presente
                for root, _, files in os.walk(pasta_usuario):
                    if attachment in files:
                        return root  # Retorna a subpasta onde o arquivo foi encontrado
        return None  # Nenhum arquivo listado foi encontrado'

    def localizar_arquivo_excel(self, pasta_usuario):
        """
        Localiza um arquivo Excel chamado 'Relatório' ou 'Report' na pasta selecionada pelo usuário.

        Args:
            pasta_usuario (str): Caminho da pasta selecionada pelo usuário.

        Returns:
            str: Caminho completo do arquivo Excel encontrado, ou None se não for encontrado.
        """
        for root, _, files in os.walk(pasta_usuario):  # Percorre subpastas também
            for file in files:
                if file.lower().startswith(("relatório", "report")) and file.lower().endswith((".xls", ".xlsx")):
                    return os.path.join(root, file)  # Retorna o caminho completo
        return None  # Nenhum arquivo encontrado

 
    def log_activity(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.config(state='disabled')
        self.log_text.see(tk.END)

    
    def start_thread_process(self):
        self.thread = threading.Thread(target=self.process)
        self.thread.start()
    
    def process(self):
        source = self.source_entry.get()
        path_dest=self.dest_entry.get()
        print(source)
        excel_file = self.localizar_arquivo_excel(source)
        audio_folder = self.localizar_subpasta_com_arquivo(excel_file, source)
        is_folder = self.checkbox_anexos.get()
        is_tag = self.checkbox_tag.get()

        if not excel_file:
            messagebox.showerror("Erro", "Por favor, selecione o arquivo Excel.")
            return

        if not audio_folder and not is_folder:
            messagebox.showerror("Erro", "Por favor, selecione a pasta de áudios ou marque a opção 'Conversa sem anexos'.")
            return

        self.log_activity("Iniciando transcrição...")

        # Criar a tabela de transcrições no PostgreSQL
        os.makedirs('Relatórios Gerados', exist_ok=True)
        os.makedirs('files', exist_ok=True)
        criar_tabela_transcricoes()

        # Processar o Excel e gerar os documentos de lote
        try:
            documentos_gerados = processar_em_lotes(excel_file, audio_folder, self.update_progress, is_tag)
        except AttributeError as e:
            self.log_activity(f"Erro durante o processamento: {str(e)}")
            return

        # Verificar se a pasta contém os arquivos referenciados no Excel
        if verificar_arquivos_na_pasta(excel_file, audio_folder):
            # Unir todos os documentos de lote em um único documento final
            
            doc_final_path = recriar_documento_final(path_dest,documentos_gerados, audio_folder)

            self.log_activity(f"Transcrição concluída. \nRelatório criado em: {doc_final_path}")

            # Opção para anonimizar interlocutores
            if messagebox.askyesno("Anonimizar", "Deseja gerar um documento anonimizado?"):
                #anonimizado_path = os.path.join(os.getcwd(), "Anonimizado.docx")
                anonimizado_path = os.path.join(path_dest, "Anonimizado.docx")
                anonimizar_interlocutores(doc_final_path, anonimizado_path)
                self.log_activity(f"Documento anonimizado criado em: {anonimizado_path}")
        else:
            self.log_activity("A pasta fornecida não contém os arquivos de mídia mencionados no arquivo Excel.")

        self.log_activity("Processamento concluído!")
        any(os.remove(os.path.join('files', f)) for f in os.listdir('files') if os.path.isfile(os.path.join('files', f)))

    def update_progress(self, progress):
        self.progress_bar['value'] = progress
        self.root.update_idletasks()
        #self.log_activity(f"Progresso: {progress:.2f}%")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
