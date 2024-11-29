import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from utils_tk import verificar_arquivos_na_pasta, processar_em_lotes, recriar_documento_final, anonimizar_interlocutores
from db import criar_tabela_transcricoes

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
        self.tab_control.add(self.second_tab, text='Consulta números')

        self.tab_control.pack(expand=1, fill='both')

        # Elementos da Aba Principal (Conversão Cellebrite)
        self.source_label = tk.Label(self.main_tab, text="Pasta de áudios e imagens:")
        self.source_label.grid(row=0, column=0, padx=10, pady=10, sticky='w')

        self.source_entry = tk.Entry(self.main_tab, width=50)
        self.source_entry.grid(row=0, column=1, padx=10, pady=10)

        self.source_button = tk.Button(self.main_tab, text="Selecionar", command=self.select_source_folder)
        self.source_button.grid(row=0, column=2, padx=10, pady=10)

        # Upload do arquivo Excel
        self.excel_label = tk.Label(self.main_tab, text="Arquivo Excel:")
        self.excel_label.grid(row=1, column=0, padx=10, pady=10, sticky='w')

        self.excel_entry = tk.Entry(self.main_tab, width=50)
        self.excel_entry.grid(row=1, column=1, padx=10, pady=10)

        self.excel_button = tk.Button(self.main_tab, text="Selecionar", command=self.select_excel_file)
        self.excel_button.grid(row=1, column=2, padx=10, pady=10)

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
        self.process_button = tk.Button(self.main_tab, text="Iniciar transcrição", command=self.process)
        self.process_button.grid(row=4, column=1, pady=20)

        # Elementos da Aba Secundária (Consulta números)
        self.info_label = tk.Label(self.second_tab, text="Esta é a página de Consulta de Números Qlik.")
        self.info_label.grid(row=0, column=0, padx=10, pady=10, sticky='w')

    def select_source_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.source_entry.delete(0, tk.END)
            self.source_entry.insert(0, folder)

    def select_excel_file(self):
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file:
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, file)

    def log_activity(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.config(state='disabled')
        self.log_text.see(tk.END)

    def process(self):
        audio_folder = self.source_entry.get()
        excel_file = self.excel_entry.get()
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
            doc_final_path = recriar_documento_final(documentos_gerados, audio_folder)

            self.log_activity(f"Transcrição concluída. Documento final criado em: {doc_final_path}")

            # Opção para anonimizar interlocutores
            if messagebox.askyesno("Anonimizar", "Deseja gerar um documento anonimizado?"):
                anonimizado_path = os.path.join(os.getcwd(), "Anonimizado.docx")
                anonimizar_interlocutores(doc_final_path, anonimizado_path)
                self.log_activity(f"Documento anonimizado criado em: {anonimizado_path}")
        else:
            self.log_activity("A pasta fornecida não contém os arquivos de mídia mencionados no arquivo Excel.")

        self.log_activity("Processamento concluído!")

    def update_progress(self, progress):
        self.progress_bar['value'] = progress
        self.root.update_idletasks()
        self.log_activity(f"Progresso: {progress:.2f}%")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
