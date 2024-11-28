import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import pandas as pd
import os

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("HawkData 1.0")
        self.root.geometry("680x650")
        
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
        self.dest_label = tk.Label(self.main_tab, text="Pasta de Destino:")
        self.dest_label.grid(row=1, column=0, padx=10, pady=10, sticky='w')
        
        self.dest_entry = tk.Entry(self.main_tab, width=50)
        self.dest_entry.grid(row=1, column=1, padx=10, pady=10)
        
        self.dest_button = tk.Button(self.main_tab, text="Selecionar", command=self.select_dest_folder)
        self.dest_button.grid(row=1, column=2, padx=10, pady=10)
        
        # Checkbox 1
        self.checkbox_anexos = tk.BooleanVar()
        self.checkbox1 = tk.Checkbutton(self.main_tab, text="Conversa sem anexos.", variable=self.checkbox_anexos)
        self.checkbox1.grid(row=2, column=0, padx=10, pady=10, sticky='w')
        
        # Checkbox 2
        self.checkbox_tag = tk.BooleanVar()
        self.checkbox2 = tk.Checkbutton(self.main_tab, text="Considerar apenas mensagens com tag.", variable=self.checkbox_tag)
        self.checkbox2.grid(row=3, column=0, padx=10, pady=10, sticky='w')
        
        # Campo de monitoramento de atividades
        self.log_text = tk.Text(self.main_tab, width=80, height=20, state='disabled')
        self.log_text.grid(row=5, column=0, columnspan=3, padx=10, pady=10)
        
        # Botão Processar
        self.process_button = tk.Button(self.main_tab, text="Processar", command=self.process)
        self.process_button.grid(row=4, column=1, pady=20)
        
        # Elementos da Aba Secundária
        # Checkbox na aba secundária
        self.checkbox_anonimizar = tk.BooleanVar()
        self.checkbox3 = tk.Checkbutton(self.second_tab, text="Anonimizar documento.", variable=self.checkbox_anonimizar)
        self.checkbox3.grid(row=0, column=0, padx=10, pady=10, sticky='w')
        
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
    
    def localizar_subpasta_com_arquivo(self, file, pasta_usuario):
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
        return None  # Nenhum arquivo listado foi encontrado

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

    def process(self):
        source = self.source_entry.get()
        dest = self.dest_entry.get()
        
        if not source:
            messagebox.showerror("Erro", "O campo 'Pasta exportada com relatórios' não pode estar vazio.")
            return
        
        if not dest:
            messagebox.showerror("Erro", "O campo 'Pasta de Destino' não pode estar vazio.")
            return
        
        option1 = self.checkbox_anexos.get()
        option2 = self.checkbox_tag.get()
        
        # Exemplo de uso das funções adicionadas
        excel_file = self.localizar_arquivo_excel(source)
        if excel_file:
            self.log_activity(f"Arquivo .xlsx encontrado em: {excel_file}")
            subpasta = self.localizar_subpasta_com_arquivo(excel_file, source)
            if subpasta:
                self.log_activity(f"Anexos encontrados na subpasta: {subpasta}")
            else:
                self.log_activity("Nenhum arquivo listado no Excel foi encontrado nas subpastas.")
        else:
            self.log_activity("Nenhum arquivo Excel encontrado na pasta selecionada.")
        
        # Exemplo de saída no campo de monitoramento de atividades
        self.log_activity(f"\nPasta exportada com relatórios: {source}")
        self.log_activity(f"Pasta de Destino: {dest}")
        #self.log_activity(f"Opção 1: {'Selecionado' if option1 else 'Não selecionado'}")
        #self.log_activity(f"Opção 2: {'Selecionado' if option2 else 'Não selecionado'}")
        self.log_activity("Processamento concluído!")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
