import os
import re
import threading
from tkinter import filedialog, messagebox, Tk, Button, Label, IntVar, ttk
from PIL import Image
import pytesseract
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from PyPDF2 import PdfMerger
from ttkthemes import ThemedStyle
import pandas as pd
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

class ExtratorTexto:
    

    def __init__(self, root):
        self.todososoficios = []
        with open('numeros_oficio.txt', 'w'):
            pass
        self.root = root
        self.metadados = {
            'Assunto': '', 
            'Autor': '',
            'local da digitalizacao': '',
            'Identificador do documento digital': '',
            'Responsavel pela digitalizacao': '',
            'Titulo': '', 
            'Tipo documental': '', 
            'Classe': '',
            'Data de producao': '', 
            'Destinacao prevista': '', 
            'Genero': '',
            'Prazo de guarda': ''
        }
        self.caminho_da_pasta, self.caminho_da_destino = "", ""
        self.progresso_var = IntVar(value=0)
        self.processo_em_andamento = False
        self.df_oficios = pd.DataFrame(columns=['NumerosOficio'])
        self.create_widgets()

    def salvar_metadados(self):
        # Obtém o caminho completo para o arquivo "metadados.txt" dentro da pasta de destino
        caminho_metadados = os.path.join(self.caminho_da_destino, "metadados.txt")

        # Abre o arquivo para escrita e salva os metadados
        with open(caminho_metadados, "w") as arquivo:
            for metadado, valor in self.metadados.items():
                arquivo.write(f"{metadado}: {valor}\n")

        # Exibe uma mensagem indicando que os metadados foram salvos com sucesso
        messagebox.showinfo("Sucesso", "Metadados salvos com sucesso em metadados.txt na pasta de destino.")


    def create_widgets(self):
        # Criando o objeto de estilo usando TTKTheme
        style = ThemedStyle(self.root)
        style.set_theme("plastik")  # Escolha o tema desejado ("plastik", "aquativo", "breeze", etc.)

        # Metadados Frame
        metadados_frame = ttk.Frame(self.root)
        metadados_frame.pack(side="left", padx=10, pady=0)

        for metadado, label_text in self.metadados.items():
            ttk.Label(metadados_frame, text=metadado + ":").grid(row=len(metadados_frame.winfo_children()), column=0, sticky='w')
            ttk.Entry(metadados_frame, font=("Arial", 12)).grid(row=len(metadados_frame.winfo_children()) - 1, column=1, pady=5, padx=5, sticky='ew')

        # Botões Frame
        botoes_frame = ttk.Frame(self.root)
        botoes_frame.pack(side="right", padx=10, pady=10, fill='y')

        self.total_arquivos_label = ttk.Label(botoes_frame, text="Total de Arquivos: 0")
        self.total_arquivos_label.pack(pady=10)
        self.arquivos_processados_label = ttk.Label(botoes_frame, text="Arquivos Processados: 0")
        self.arquivos_processados_label.pack()

        ttk.Button(botoes_frame, text="Selecionar Pasta de Imagens", command=self.selecionar_pasta).pack(pady=10, fill='x')
        ttk.Label(botoes_frame, text="Pasta de Destino:").pack()
        ttk.Button(botoes_frame, text="Selecionar Destino", command=self.selecionar_destino).pack(pady=10, fill='x')
        ttk.Button(botoes_frame, text="Extrair Texto", command=self.iniciar_processo).pack(pady=10, fill='x')
        ttk.Button(botoes_frame, text="Unir PDFs", command=self.unir_pdfs).pack(pady=10, fill='x')
        self.progresso_label = ttk.Label(botoes_frame, text="Progresso: 0%")
        self.progresso_label.pack()
        self.arquivo_atual_label = ttk.Label(botoes_frame, text="")
        self.arquivo_atual_label.pack()
        ttk.Button(botoes_frame, text="Parar", command=self.parar_processo).pack(pady=10, fill='x')
        ttk.Button(botoes_frame, text="Salvar Metadados", command=self.salvar_metadados).pack(pady=10, fill='x') 
        self.entry_fields = {}  # Dicionário para armazenar os campos de entrada

        for idx, (metadado, label_text) in enumerate(self.metadados.items(), start=1):
            ttk.Label(metadados_frame, text=metadado + ":").grid(row=idx, column=0, sticky='w')
            entry = ttk.Entry(metadados_frame, font=("Arial", 12))
            entry.grid(row=idx, column=1, pady=5, padx=5, sticky='ew')
            self.entry_fields[metadado] = entry  # Armazena os campos de entrada no dicionário

       
    # Restante do código permanece inalterado


    def extrair_numeros_oficio(self, texto):
        padrao = r'of(?:[íi]cio|icio|ícios|ício)?\s*[nº°]*[.:]?\s*(\d+/\d+)'
        return re.findall(padrao, texto, re.IGNORECASE)

    def selecionar_pasta(self):
        self.caminho_da_pasta = filedialog.askdirectory()
        if self.caminho_da_pasta:
            arquivos_imagem = [arquivo for arquivo in os.listdir(self.caminho_da_pasta) if arquivo.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]
            self.total_arquivos_label.config(text=f"Total de Arquivos: {len(arquivos_imagem)}")

    def selecionar_destino(self):
        self.caminho_da_destino = filedialog.askdirectory()

    def verificar_arquivos_existentes(self, nome_arquivo):
        nome_pdf = os.path.join(self.caminho_da_destino, nome_arquivo + '.pdf')
        return os.path.exists(nome_pdf)

    def iniciar_processo(self):
        if not self.caminho_da_pasta or not self.caminho_da_destino:
            messagebox.showerror("Erro", "Por favor, selecione a pasta de origem e destino.")
            return
        self.processo_em_andamento = True
        threading.Thread(target=self.processar_imagens).start()

    def processar_imagens(self):
        try:
            arquivos = sorted([arquivo for arquivo in os.listdir(self.caminho_da_pasta) if arquivo.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))])
            total_arquivos = len(arquivos)
            self.total_arquivos_label.config(text=f"Total de Arquivos: {total_arquivos}")
            for i, arquivo in enumerate(arquivos, start=1):
                if not self.processo_em_andamento:
                    break
                nome_arquivo = os.path.splitext(arquivo)[0]
                if not self.verificar_arquivos_existentes(nome_arquivo):
                    self.arquivos_processados_label.config(text=f"Arquivos Processados: {i}/{total_arquivos}")
                    caminho_da_imagem = os.path.join(self.caminho_da_pasta, arquivo)
                    imagem = Image.open(caminho_da_imagem)
                    texto_extraido = pytesseract.image_to_string(imagem)
                    nome_pdf = os.path.join(self.caminho_da_destino, nome_arquivo + '.pdf')
                    self.criar_pdf(texto_extraido, caminho_da_imagem, nome_pdf)
                else:
                    self.arquivos_processados_label.config(text=f"Arquivos Processados: {i}/{total_arquivos} (Já Existem)")
                progresso = int(i / total_arquivos * 100)
                self.progresso_var.set(progresso)
                self.progresso_label.config(text=f"Progresso: {progresso}%")
                self.arquivo_atual_label.config(text=f"Arquivo Atual: {arquivo}")
            if self.processo_em_andamento:
                messagebox.showinfo("Concluído", "Texto extraído das imagens e salvos em arquivos PDF com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
        finally:
            self.processo_em_andamento = False

    def parar_processo(self):
        self.processo_em_andamento = False
        messagebox.showinfo("Parado", "Processo interrompido pelo usuário.")
    
    def criar_pdf(self, texto, arquivo_imagem, nome_pdf):
        try:
            # Restante do seu código...

            numeros_oficio = self.extrair_numeros_oficio(texto)

            if numeros_oficio:
                # Adicione apenas os novos números de ofício à lista contínua
                self.todososoficios.extend(numeros_oficio)
                # Converte os novos números do ofício em uma única string separada por ponto e vírgula
                numeros_oficio_string = '; '.join(numeros_oficio)
                # Escreve os novos números do ofício no arquivo de texto separados por ponto e vírgula
                with open(os.path.join(self.caminho_da_destino, 'oficios_encontrados.txt'), 'a') as arquivo_txt:
                    arquivo_txt.write(f'{numeros_oficio_string}\n')

                # Adicione os novos números de ofício ao DataFrame pandas
                df_novos_oficios = pd.DataFrame({'NumerosOficio': numeros_oficio})
                self.df_oficios = pd.concat([self.df_oficios, df_novos_oficios], ignore_index=True)

                # Salva o DataFrame em um arquivo CSV
                csv_path = os.path.join(self.caminho_da_destino, 'oficios_encontrados.csv')
                self.df_oficios.to_csv(csv_path, sep=';', index=False)
                print(f"Números do ofício encontrados: {numeros_oficio_string}")
            else:
                print("Nenhum número do ofício encontrado no texto.")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao criar o PDF: {str(e)}")



    def unir_pdfs(self):
        try:
            arquivos_pdf = [f for f in os.listdir(self.caminho_da_destino) if f.endswith(".pdf")]
            arquivos_pdf.sort()
            merger = PdfMerger()
            for arquivo_pdf in arquivos_pdf:
                caminho_pdf = os.path.join(self.caminho_da_destino, arquivo_pdf)
                merger.append(caminho_pdf)

            arquivo_final = os.path.join(self.caminho_da_destino, "Documento_Final.pdf")
            merger.write(arquivo_final)
            merger.close()
            messagebox.showinfo("Concluído", f"Arquivos PDF unidos em '{arquivo_final}' com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao unir os PDFs: {str(e)}")
    
    def salvar_metadados(self):
        # Obtém os valores dos campos de entrada
        for metadado, entry in self.entry_fields.items():
            self.metadados[metadado] = entry.get()

        # Obtém o caminho completo para o arquivo "metadados.txt" dentro da pasta de destino
        caminho_metadados = os.path.join(self.caminho_da_destino, "metadados.txt")

        # Abre o arquivo para escrita e salva os metadados
        with open(caminho_metadados, "w") as arquivo:
            for metadado, valor in self.metadados.items():
                arquivo.write(f"{metadado}: {valor}\n")

        # Exibe uma mensagem indicando que os metadados foram salvos com sucesso
        messagebox.showinfo("Sucesso", "Metadados salvos com sucesso em metadados.txt na pasta de destino.")
 
if __name__ == "__main__":
    root = Tk()
    root.title("Extrator inteligente")

    # Configurar o redimensionamento da janela
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)

    extrator = ExtratorTexto(root)
    root.mainloop()
