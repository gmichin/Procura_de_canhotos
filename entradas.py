import os
import PyPDF2
import pytesseract
from PIL import Image
import fitz  # PyMuPDF
import re
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import subprocess

# Configurar o caminho do Tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

class LocalizadorNotasFiscais:
    def __init__(self, caminho_base):
        self.caminho_base = caminho_base
        self.resultados = []
        self.debug_info = []
        
    def adicionar_debug(self, mensagem):
        """Adiciona mensagem de debug"""
        self.debug_info.append(mensagem)
        print(f"DEBUG: {mensagem}")
    
    def converter_pdf_para_imagem(self, pdf_path, pagina_num):
        """Converte uma página PDF em imagem para OCR com múltiplas tentativas"""
        try:
            doc = fitz.open(pdf_path)
            pagina = doc.load_page(pagina_num)
            
            # Tentar diferentes rotações para lidar com páginas horizontais
            rotacoes = [0, 90, 180, 270]
            
            for rotacao in rotacoes:
                try:
                    matriz = fitz.Matrix(3, 3).prerotate(rotacao)  # Aumentei a resolução
                    pix = pagina.get_pixmap(matrix=matriz)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    
                    # Verificar se a imagem tem conteúdo
                    if img.size[0] > 0 and img.size[1] > 0:
                        doc.close()
                        return img
                except Exception as e:
                    continue
            
            doc.close()
            return None
            
        except Exception as e:
            self.adicionar_debug(f"Erro ao converter PDF {pdf_path}, página {pagina_num}: {e}")
            return None
    
    def preprocessar_imagem(self, imagem):
        """Pré-processa a imagem para melhorar o OCR"""
        try:
            # Converter para escala de cinza
            if imagem.mode != 'L':
                imagem = imagem.convert('L')
            
            # Aumentar contraste
            from PIL import ImageEnhance
            enhancer = ImageEnhance.Contrast(imagem)
            imagem = enhancer.enhance(2.0)  # Aumenta contraste
            
            enhancer = ImageEnhance.Sharpness(imagem)
            imagem = enhancer.enhance(2.0)  # Aumenta nitidez
            
            return imagem
        except Exception as e:
            self.adicionar_debug(f"Erro no pré-processamento: {e}")
            return imagem
    
    def buscar_texto_ocr(self, imagem, numero_nota):
        """Realiza OCR na imagem com múltiplas estratégias"""
        try:
            imagem_processada = self.preprocessar_imagem(imagem)
            
            # Tentar diferentes configurações de PSM
            configs = [
                r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789',  # Bloco uniforme
                r'--oem 3 --psm 8 -c tessedit_char_whitelist=0123456789',  # Palavra única
                r'--oem 3 --psm 11 -c tessedit_char_whitelist=0123456789',  # Texto puro
                r'--oem 3 --psm 4 -c tessedit_char_whitelist=0123456789',   # Coluna única
            ]
            
            for config in configs:
                try:
                    texto = pytesseract.image_to_string(imagem_processada, config=config)
                    texto_limpo = re.sub(r'[^\d]', '', texto)  # Manter apenas números
                    
                    self.adicionar_debug(f"OCR config {config}: encontrou '{texto_limpo}'")
                    
                    # Verificar se o número da nota está no texto
                    if numero_nota in texto_limpo:
                        return True
                        
                    # Verificar combinações parciais (para casos onde o OCR pode ter errado alguns dígitos)
                    if len(numero_nota) > 5:
                        # Buscar pelos últimos 6 dígitos (mais prováveis de serem lidos corretamente)
                        ultimos_digitos = numero_nota[-6:]
                        if ultimos_digitos in texto_limpo:
                            return True
                            
                except Exception as e:
                    continue
            
            return False
            
        except Exception as e:
            self.adicionar_debug(f"Erro no OCR: {e}")
            return False
    
    def buscar_texto_direto_pdf(self, pdf_path, numero_nota):
        """Busca texto diretamente no PDF (sem OCR)"""
        try:
            with open(pdf_path, 'rb') as arquivo:
                leitor = PyPDF2.PdfReader(arquivo)
                
                for num_pagina in range(len(leitor.pages)):
                    pagina = leitor.pages[num_pagina]
                    texto = pagina.extract_text()
                    
                    if texto and numero_nota in texto:
                        self.adicionar_debug(f"Encontrado via texto direto em {pdf_path} página {num_pagina + 1}")
                        return [num_pagina + 1]
            
            return []
        except Exception as e:
            self.adicionar_debug(f"Erro na busca direta: {e}")
            return []
    
    def buscar_nome_arquivo(self, pdf_path, numero_nota):
        """Verifica se o número da nota está no nome do arquivo"""
        nome_arquivo = os.path.basename(pdf_path)
        if numero_nota in nome_arquivo:
            self.adicionar_debug(f"Nota encontrada no nome do arquivo: {nome_arquivo}")
            return True
        return False
    
    def buscar_nota_em_pdf(self, pdf_path, numero_nota):
        """Busca o número da nota em um arquivo PDF usando múltiplas estratégias"""
        try:
            resultados_paginas = []
            
            # Estratégia 1: Buscar no nome do arquivo
            if self.buscar_nome_arquivo(pdf_path, numero_nota):
                resultados_paginas.append(1)  # Assumir página 1 se encontrado no nome
            
            # Estratégia 2: Busca textual direta no PDF
            paginas_texto = self.buscar_texto_direto_pdf(pdf_path, numero_nota)
            resultados_paginas.extend(paginas_texto)
            
            # Estratégia 3: OCR (apenas se não encontrou pelas outras estratégias)
            if not resultados_paginas:
                self.adicionar_debug(f"Usando OCR para: {os.path.basename(pdf_path)}")
                
                with open(pdf_path, 'rb') as arquivo:
                    leitor = PyPDF2.PdfReader(arquivo)
                    total_paginas = len(leitor.pages)
                
                for num_pagina in range(total_paginas):
                    self.adicionar_debug(f"Processando página {num_pagina + 1} de {total_paginas}")
                    
                    imagem = self.converter_pdf_para_imagem(pdf_path, num_pagina)
                    if imagem:
                        encontrado = self.buscar_texto_ocr(imagem, numero_nota)
                        if encontrado:
                            resultados_paginas.append(num_pagina + 1)
                            self.adicionar_debug(f"Encontrado via OCR na página {num_pagina + 1}")
                            break  # Parar na primeira ocorrência
            
            return list(set(resultados_paginas)) if resultados_paginas else None
            
        except Exception as e:
            self.adicionar_debug(f"Erro ao processar {pdf_path}: {e}")
            return None
    
    def buscar_nota(self, mes, dia, numero_nota):
        """Busca a nota fiscal no mês e dia especificados"""
        self.resultados = []
        self.debug_info = []
        numero_nota = numero_nota.strip()
        
        self.adicionar_debug(f"Iniciando busca: {dia}/{mes} - Nota: {numero_nota}")
        
        # Mapear nome do mês para pasta
        meses = {
            'janeiro': 'JANEIRO', 'fevereiro': 'FEVEREIRO', 'março': 'MARÇO',
            'abril': 'ABRIL', 'maio': 'MAIO', 'junho': 'JUNHO',
            'julho': 'JULHO', 'agosto': 'AGOSTO', 'setembro': 'SETEMBRO',
            'outubro': 'OUTUBRO', 'novembro': 'NOVEMBRO', 'dezembro': 'DEZEMBRO'
        }
        
        mes_pasta = meses.get(mes.lower(), mes.upper())
        caminho_mes = os.path.join(self.caminho_base, mes_pasta)
        
        if not os.path.exists(caminho_mes):
            return f"Pasta do mês {mes} não encontrada em {caminho_mes}"
        
        self.adicionar_debug(f"Buscando em: {caminho_mes}")
        
        # Buscar nas pastas de dias
        dia_procurado = f"{int(dia):02d}"
        pastas_encontradas = []
        
        for pasta_dia in os.listdir(caminho_mes):
            caminho_dia = os.path.join(caminho_mes, pasta_dia)
            
            if os.path.isdir(caminho_dia) and dia_procurado in pasta_dia:
                pastas_encontradas.append(pasta_dia)
                self.adicionar_debug(f"Buscando na pasta: {pasta_dia}")
                
                # Buscar em todos os PDFs da pasta
                pdfs_encontrados = 0
                for arquivo in os.listdir(caminho_dia):
                    if arquivo.lower().endswith('.pdf'):
                        pdfs_encontrados += 1
                        caminho_pdf = os.path.join(caminho_dia, arquivo)
                        self.adicionar_debug(f"Processando: {arquivo}")
                        
                        paginas = self.buscar_nota_em_pdf(caminho_pdf, numero_nota)
                        
                        if paginas:
                            for pagina in paginas:
                                self.resultados.append({
                                    'arquivo': caminho_pdf,
                                    'pagina': pagina,
                                    'pasta_dia': pasta_dia,
                                    'nome_arquivo': arquivo
                                })
                
                self.adicionar_debug(f"Processados {pdfs_encontrados} PDFs na pasta {pasta_dia}")
        
        if not pastas_encontradas:
            return f"Nenhuma pasta encontrada para o dia {dia} em {mes}"
        
        self.adicionar_debug(f"Busca finalizada. {len(self.resultados)} resultado(s) encontrado(s)")
        return self.resultados

class InterfaceLocalizador:
    def __init__(self, root):
        self.root = root
        self.root.title("Localizador de Notas Fiscais - Melhorado")
        self.root.geometry("800x600")
        
        self.localizador = LocalizadorNotasFiscais(r"Z:\NOTAS-CANHOTOS-DEVOLUÇÕES\NOTAS ENTRADA")
        
        self.criar_interface()
    
    def criar_interface(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título
        titulo = ttk.Label(main_frame, text="Localizador de Notas Fiscais", font=('Arial', 14, 'bold'))
        titulo.grid(row=0, column=0, columnspan=2, pady=10)
        
        # Campos de entrada
        ttk.Label(main_frame, text="Mês:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.mes_entry = ttk.Combobox(main_frame, values=[
            'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
            'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
        ], state='readonly', width=20)
        self.mes_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        
        ttk.Label(main_frame, text="Dia:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.dia_entry = ttk.Spinbox(main_frame, from_=1, to=31, width=10)
        self.dia_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        
        ttk.Label(main_frame, text="Número da Nota:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.nota_entry = ttk.Entry(main_frame, width=25)
        self.nota_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        
        # Botão de busca
        self.buscar_btn = ttk.Button(main_frame, text="Buscar Nota", command=self.iniciar_busca)
        self.buscar_btn.grid(row=4, column=0, columnspan=2, pady=10)
        
        # Barra de progresso
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Aba para resultados e debug
        notebook = ttk.Notebook(main_frame)
        notebook.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        # Aba de resultados
        self.aba_resultados = ttk.Frame(notebook)
        notebook.add(self.aba_resultados, text="Resultados")
        
        # Aba de debug
        self.aba_debug = ttk.Frame(notebook)
        notebook.add(self.aba_debug, text="Log de Busca")
        
        # Configurar grid weights
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(6, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # Configurar áreas de texto com scroll
        self.configurar_areas_texto()
    
    def configurar_areas_texto(self):
        # Área de resultados
        self.resultados_texto = tk.Text(self.aba_resultados, height=15, wrap=tk.WORD)
        scrollbar_resultados = ttk.Scrollbar(self.aba_resultados, orient="vertical", command=self.resultados_texto.yview)
        self.resultados_texto.configure(yscrollcommand=scrollbar_resultados.set)
        
        self.resultados_texto.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        scrollbar_resultados.pack(side="right", fill="y")
        
        # Área de debug
        self.debug_texto = tk.Text(self.aba_debug, height=15, wrap=tk.WORD, bg='black', fg='white')
        scrollbar_debug = ttk.Scrollbar(self.aba_debug, orient="vertical", command=self.debug_texto.yview)
        self.debug_texto.configure(yscrollcommand=scrollbar_debug.set)
        
        self.debug_texto.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        scrollbar_debug.pack(side="right", fill="y")
    
    def iniciar_busca(self):
        mes = self.mes_entry.get()
        dia = self.dia_entry.get()
        nota = self.nota_entry.get()
        
        if not all([mes, dia, nota]):
            messagebox.showerror("Erro", "Preencha todos os campos")
            return
        
        # Limpar áreas de texto
        self.resultados_texto.delete(1.0, tk.END)
        self.debug_texto.delete(1.0, tk.END)
        
        self.buscar_btn.config(state='disabled')
        self.progress.start()
        
        # Executar busca em thread separada
        thread = threading.Thread(target=self.executar_busca, args=(mes, dia, nota))
        thread.daemon = True
        thread.start()
    
    def executar_busca(self, mes, dia, nota):
        try:
            resultados = self.localizador.buscar_nota(mes, dia, nota)
            
            # Atualizar interface na thread principal
            self.root.after(0, self.mostrar_resultados, resultados, mes, dia, nota)
            
        except Exception as e:
            self.root.after(0, self.mostrar_erro, str(e))
    
    def mostrar_resultados(self, resultados, mes, dia, nota):
        self.progress.stop()
        self.buscar_btn.config(state='normal')
        
        # Mostrar log de debug
        for mensagem in self.localizador.debug_info:
            self.debug_texto.insert(tk.END, mensagem + '\n')
        
        if isinstance(resultados, str):
            # Mensagem de erro
            self.resultados_texto.insert(tk.END, f"ERRO: {resultados}\n", 'erro')
            return
        
        if not resultados:
            self.resultados_texto.insert(tk.END, 
                f"Nenhuma nota {nota} encontrada para {dia} de {mes}\n\n"
                f"Sugestões:\n"
                f"1. Verifique se o número da nota está correto\n"
                f"2. Tente buscar por parte do número\n"
                f"3. Verifique o mês e dia corretos\n"
                f"4. Consulte o log de busca para detalhes", 'aviso')
            return
        
        self.resultados_texto.insert(tk.END, 
            f"Encontrada(s) {len(resultados)} ocorrência(s) da nota {nota}:\n\n", 'sucesso')
        
        for i, resultado in enumerate(resultados, 1):
            self.resultados_texto.insert(tk.END, 
                f"{i}. Arquivo: {resultado['nome_arquivo']}\n"
                f"   Pasta: {resultado['pasta_dia']}\n"
                f"   Página: {resultado['pagina']}\n"
                f"   Caminho: {resultado['arquivo']}\n\n")
        
        # Configurar tags para cores
        self.resultados_texto.tag_configure('erro', foreground='red')
        self.resultados_texto.tag_configure('aviso', foreground='orange')
        self.resultados_texto.tag_configure('sucesso', foreground='green')
    
    def mostrar_erro(self, erro):
        self.progress.stop()
        self.buscar_btn.config(state='normal')
        self.resultados_texto.insert(tk.END, f"ERRO: {erro}\n", 'erro')

def main():
    root = tk.Tk()
    app = InterfaceLocalizador(root)
    root.mainloop()

if __name__ == "__main__":
    main()