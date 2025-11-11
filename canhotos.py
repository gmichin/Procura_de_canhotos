import os
import fitz  # PyMuPDF
import re
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import subprocess
import platform
from PIL import Image, ImageEnhance
import pytesseract
import io

# Configurar o caminho do Tesseract (ajuste conforme sua instalação)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

class BuscadorCanhotosAvancado:
    def __init__(self, caminho_base):
        self.caminho_base = caminho_base
        self.resultados = []
        self.debug_info = []
        self.usar_ocr = True  # Ativar OCR como fallback
        
    def adicionar_debug(self, mensagem):
        """Adiciona mensagem de debug"""
        self.debug_info.append(mensagem)
        print(f"DEBUG: {mensagem}")
    
    def normalizar_mes(self, mes):
        """Normaliza o nome do mês para o formato das pastas"""
        meses = {
            'janeiro': 'JANEIRO', 'fevereiro': 'FEVEREIRO', 'março': 'MARÇO', 'marco': 'MARÇO',
            'abril': 'ABRIL', 'maio': 'MAIO', 'junho': 'JUNHO',
            'julho': 'JULHO', 'agosto': 'AGOSTO', 'setembro': 'SETEMBRO',
            'outubro': 'OUTUBRO', 'novembro': 'NOVEMBRO', 'dezembro': 'DEZEMBRO'
        }
        return meses.get(mes.lower(), mes.upper())
    
    def encontrar_pastas_dia(self, caminho_mes, dia):
        """Encontra pastas do dia em diferentes formatos"""
        dia_procurado = f"{int(dia):02d}"
        pastas_encontradas = []
        
        for pasta in os.listdir(caminho_mes):
            caminho_pasta = os.path.join(caminho_mes, pasta)
            
            if os.path.isdir(caminho_pasta):
                # Verificar diferentes formatos de data
                if (pasta.startswith(f"{dia_procurado}-") or 
                    pasta.startswith(f"{dia_procurado}_") or
                    f"-{dia_procurado}-" in pasta or
                    f"_{dia_procurado}_" in pasta):
                    pastas_encontradas.append({
                        'nome': pasta,
                        'caminho': caminho_pasta
                    })
        
        return pastas_encontradas
    
    def melhorar_imagem_ocr(self, imagem):
        """Melhora a imagem para OCR"""
        try:
            # Aumentar contraste
            enhancer = ImageEnhance.Contrast(imagem)
            imagem = enhancer.enhance(2.0)
            
            # Aumentar nitidez
            enhancer = ImageEnhance.Sharpness(imagem)
            imagem = enhancer.enhance(2.0)
            
            # Aumentar brilho se necessário
            enhancer = ImageEnhance.Brightness(imagem)
            imagem = enhancer.enhance(1.1)
            
            return imagem
        except Exception as e:
            self.adicionar_debug(f"Erro no pré-processamento de imagem: {e}")
            return imagem
    
    def buscar_com_ocr(self, pdf_path, numero_nota, num_pagina):
        """Busca o texto usando OCR"""
        try:
            doc = fitz.open(pdf_path)
            pagina = doc.load_page(num_pagina)
            
            # Converter a página em imagem com alta resolução
            mat = fitz.Matrix(3, 3)  # Aumentar a resolução
            pix = pagina.get_pixmap(matrix=mat)
            img_data = pix.tobytes("ppm")
            
            # Converter para imagem PIL
            imagem = Image.open(io.BytesIO(img_data))
            
            # Melhorar imagem para OCR
            imagem_melhorada = self.melhorar_imagem_ocr(imagem)
            
            # Configurações do Tesseract
            config_tesseract = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789'
            
            # Fazer OCR
            texto_ocr = pytesseract.image_to_string(imagem_melhorada, config=config_tesseract)
            
            doc.close()
            
            # Procurar o número da nota no texto do OCR
            if numero_nota in texto_ocr:
                self.adicionar_debug(f"OCR encontrou '{numero_nota}' na página {num_pagina + 1}")
                return True
                
            # Buscar por partes do número (caso o OCR tenha errado alguns dígitos)
            if len(numero_nota) > 4:
                # Buscar últimos 4-5 dígitos
                for i in range(4, len(numero_nota)):
                    parte = numero_nota[-i:]
                    if parte in texto_ocr and len(parte) >= 4:
                        self.adicionar_debug(f"OCR encontrou parte '{parte}' na página {num_pagina + 1}")
                        return True
            
            return False
            
        except Exception as e:
            self.adicionar_debug(f"Erro no OCR página {num_pagina + 1}: {e}")
            return False
    
    def buscar_texto_no_pdf(self, caminho_pdf, numero_nota):
        """Busca o número da nota no PDF usando múltiplas estratégias"""
        try:
            paginas_encontradas = []
            pdf_document = fitz.open(caminho_pdf)
            
            for num_pagina in range(len(pdf_document)):
                pagina = pdf_document.load_page(num_pagina)
                texto = pagina.get_text()
                
                # Estratégia 1: Busca direta
                if numero_nota in texto:
                    paginas_encontradas.append(num_pagina + 1)
                    self.adicionar_debug(f"Busca direta: encontrado na página {num_pagina + 1}")
                    continue
                
                # Estratégia 2: Busca por padrões com regex
                padroes = [
                    rf"\b{numero_nota}\b",
                    rf"\b{numero_nota[:6]}.*\b",
                    rf"\b.*{numero_nota[-6:]}\b",
                    rf"N[°º]?\s*{re.escape(numero_nota)}",
                    rf"NF-?E?\s*{re.escape(numero_nota)}",
                    rf"Nota\s*Fiscal\s*{re.escape(numero_nota)}",
                    rf"NUMERO\s*NF-?E?\s*{re.escape(numero_nota)}",
                ]
                
                texto_limpo = re.sub(r'\s+', ' ', texto)  # Normalizar espaços
                
                for padrao in padroes:
                    if re.search(padrao, texto_limpo, re.IGNORECASE):
                        if (num_pagina + 1) not in paginas_encontradas:
                            paginas_encontradas.append(num_pagina + 1)
                            self.adicionar_debug(f"Regex '{padrao}': encontrado na página {num_pagina + 1}")
                            break
                
                # Estratégia 3: OCR (se as anteriores não funcionaram)
                if not paginas_encontradas and self.usar_ocr:
                    self.adicionar_debug(f"Tentando OCR na página {num_pagina + 1}...")
                    if self.buscar_com_ocr(caminho_pdf, numero_nota, num_pagina):
                        paginas_encontradas.append(num_pagina + 1)
                        self.adicionar_debug(f"OCR: encontrado na página {num_pagina + 1}")
            
            pdf_document.close()
            return paginas_encontradas
            
        except Exception as e:
            self.adicionar_debug(f"Erro ao processar {caminho_pdf}: {e}")
            return []
    
    def abrir_pdf_pagina(self, caminho_pdf, numero_pagina):
        """Abre o PDF na página específica"""
        try:
            sistema = platform.system()
            
            if sistema == "Windows":
                # Tentar com SumatraPDF primeiro (suporta página específica)
                try:
                    subprocess.run(['SumatraPDF', f'-page={numero_pagina}', caminho_pdf], 
                                 timeout=10, check=True)
                    return True
                except:
                    # Fallback para visualizador padrão
                    os.startfile(caminho_pdf)
                    messagebox.showinfo("Abrir PDF", 
                                      f"Arquivo: {os.path.basename(caminho_pdf)}\n"
                                      f"Vá manualmente para a página: {numero_pagina}")
                    return True
                    
            elif sistema == "Darwin":  # macOS
                subprocess.run(['open', caminho_pdf])
                messagebox.showinfo("Abrir PDF", 
                                  f"Arquivo: {os.path.basename(caminho_pdf)}\n"
                                  f"Vá manualmente para a página: {numero_pagina}")
                return True
                
            elif sistema == "Linux":
                subprocess.run(['xdg-open', caminho_pdf])
                messagebox.showinfo("Abrir PDF", 
                                  f"Arquivo: {os.path.basename(caminho_pdf)}\n"
                                  f"Vá manualmente para a página: {numero_pagina}")
                return True
                
        except Exception as e:
            self.adicionar_debug(f"Erro ao abrir PDF: {e}")
            messagebox.showerror("Erro", f"Não foi possível abrir o PDF:\n{e}")
            return False
    
    def buscar_canhotos(self, mes, dia, numero_nota):
        """Busca principal pelos canhotos"""
        self.resultados = []
        self.debug_info = []
        
        self.adicionar_debug(f"Iniciando busca: {dia}/{mes} - Nota: {numero_nota}")
        
        # Normalizar mês
        mes_pasta = self.normalizar_mes(mes)
        caminho_mes = os.path.join(self.caminho_base, mes_pasta)
        
        if not os.path.exists(caminho_mes):
            return f"Pasta do mês {mes} não encontrada em {caminho_mes}"
        
        self.adicionar_debug(f"Buscando em: {caminho_mes}")
        
        # Encontrar pastas do dia
        pastas_dia = self.encontrar_pastas_dia(caminho_mes, dia)
        
        if not pastas_dia:
            return f"Nenhuma pasta encontrada para o dia {dia} em {mes}"
        
        self.adicionar_debug(f"Encontradas {len(pastas_dia)} pasta(s) do dia {dia}")
        
        # Buscar em todas as pastas do dia
        for pasta_info in pastas_dia:
            self.adicionar_debug(f"Buscando na pasta: {pasta_info['nome']}")
            
            pdfs_processados = 0
            for arquivo in os.listdir(pasta_info['caminho']):
                if arquivo.lower().endswith('.pdf'):
                    pdfs_processados += 1
                    caminho_pdf = os.path.join(pasta_info['caminho'], arquivo)
                    
                    self.adicionar_debug(f"Processando: {arquivo}")
                    paginas = self.buscar_texto_no_pdf(caminho_pdf, numero_nota)
                    
                    if paginas:
                        for pagina in paginas:
                            self.resultados.append({
                                'arquivo': caminho_pdf,
                                'pagina': pagina,
                                'pasta_dia': pasta_info['nome'],
                                'nome_arquivo': arquivo,
                                'mes': mes,
                                'dia': dia
                            })
            
            self.adicionar_debug(f"Processados {pdfs_processados} PDFs na pasta {pasta_info['nome']}")
        
        self.adicionar_debug(f"Busca finalizada. {len(self.resultados)} resultado(s) encontrado(s)")
        return self.resultados

class InterfaceBuscadorCanhotos:
    def __init__(self, root):
        self.root = root
        self.root.title("Buscador de Canhotos - Com OCR")
        self.root.geometry("950x750")
        
        # Caminho base
        self.caminho_base = r"Z:\NOTAS-CANHOTOS-DEVOLUÇÕES\CANHOTOS"
        self.buscador = BuscadorCanhotosAvancado(self.caminho_base)
        
        self.criar_interface()
        self.carregar_meses_disponiveis()
    
    def carregar_meses_disponiveis(self):
        """Carrega os meses disponíveis no caminho base"""
        try:
            if os.path.exists(self.caminho_base):
                pastas = [p for p in os.listdir(self.caminho_base) 
                         if os.path.isdir(os.path.join(self.caminho_base, p))]
                meses_ordenados = self.ordenar_meses(pastas)
                self.mes_combobox['values'] = meses_ordenados
        except Exception as e:
            print(f"Erro ao carregar meses: {e}")
    
    def ordenar_meses(self, meses):
        """Ordena os meses cronologicamente"""
        ordem_meses = [
            'JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO',
            'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO'
        ]
        return [mes for mes in ordem_meses if mes in meses]
    
    def criar_interface(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título
        titulo = ttk.Label(main_frame, text="🔍 BUSCADOR DE CANHOTOS - COM OCR", 
                          font=('Arial', 16, 'bold'))
        titulo.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Campos de entrada
        ttk.Label(main_frame, text="Mês:", font=('Arial', 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.mes_combobox = ttk.Combobox(main_frame, width=20, font=('Arial', 10))
        self.mes_combobox.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        
        ttk.Label(main_frame, text="Dia:", font=('Arial', 10)).grid(row=2, column=0, sticky=tk.W, pady=5)
        self.dia_combobox = ttk.Combobox(main_frame, width=10, font=('Arial', 10))
        self.dia_combobox['values'] = [f"{i:02d}" for i in range(1, 32)]
        self.dia_combobox.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        
        ttk.Label(main_frame, text="Número da Nota:", font=('Arial', 10)).grid(row=3, column=0, sticky=tk.W, pady=5)
        self.nota_entry = ttk.Entry(main_frame, width=25, font=('Arial', 10))
        self.nota_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        
        # Configurações
        settings_frame = ttk.Frame(main_frame)
        settings_frame.grid(row=4, column=0, columnspan=3, pady=10)
        
        self.ocr_var = tk.BooleanVar(value=True)
        self.ocr_check = ttk.Checkbutton(settings_frame, text="Usar OCR (recomendado)", 
                                       variable=self.ocr_var)
        self.ocr_check.pack(side=tk.LEFT, padx=10)
        
        # Botões
        botoes_frame = ttk.Frame(main_frame)
        botoes_frame.grid(row=5, column=0, columnspan=3, pady=15)
        
        self.buscar_btn = ttk.Button(botoes_frame, text="🔎 Buscar Canhoto", 
                                   command=self.iniciar_busca, width=20)
        self.buscar_btn.pack(side=tk.LEFT, padx=5)
        
        self.limpar_btn = ttk.Button(botoes_frame, text="🗑️ Limpar", 
                                   command=self.limpar_campos, width=15)
        self.limpar_btn.pack(side=tk.LEFT, padx=5)
        
        # Barra de progresso
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # Status
        self.status_label = ttk.Label(main_frame, text="Pronto para buscar", font=('Arial', 9))
        self.status_label.grid(row=7, column=0, columnspan=3, pady=5)
        
        # Abas
        notebook = ttk.Notebook(main_frame)
        notebook.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        # Aba de resultados
        self.aba_resultados = ttk.Frame(notebook)
        notebook.add(self.aba_resultados, text="📄 Resultados")
        
        # Aba de log
        self.aba_log = ttk.Frame(notebook)
        notebook.add(self.aba_log, text="📋 Log de Busca")
        
        # Configurar grid
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(8, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # Configurar áreas de texto
        self.configurar_areas_texto()
        
        # Bind events
        self.nota_entry.bind('<Return>', lambda e: self.iniciar_busca())
    
    def configurar_areas_texto(self):
        # Área de resultados
        resultados_frame = ttk.Frame(self.aba_resultados)
        resultados_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.resultados_texto = tk.Text(resultados_frame, height=20, wrap=tk.WORD, font=('Arial', 10))
        scrollbar_resultados = ttk.Scrollbar(resultados_frame, orient="vertical", command=self.resultados_texto.yview)
        self.resultados_texto.configure(yscrollcommand=scrollbar_resultados.set)
        
        self.resultados_texto.pack(side="left", fill="both", expand=True)
        scrollbar_resultados.pack(side="right", fill="y")
        
        # Área de log
        log_frame = ttk.Frame(self.aba_log)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.log_texto = tk.Text(log_frame, height=20, wrap=tk.WORD, font=('Arial', 9),
                                bg='black', fg='white')
        scrollbar_log = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_texto.yview)
        self.log_texto.configure(yscrollcommand=scrollbar_log.set)
        
        self.log_texto.pack(side="left", fill="both", expand=True)
        scrollbar_log.pack(side="right", fill="y")
    
    def iniciar_busca(self):
        mes = self.mes_combobox.get()
        dia = self.dia_combobox.get()
        nota = self.nota_entry.get().strip()
        
        if not all([mes, dia, nota]):
            messagebox.showerror("Erro", "Por favor, preencha todos os campos.")
            return
        
        if not nota.isdigit():
            messagebox.showerror("Erro", "O número da nota deve conter apenas dígitos.")
            return
        
        # Configurar OCR
        self.buscador.usar_ocr = self.ocr_var.get()
        
        # Limpar resultados anteriores
        self.resultados_texto.delete(1.0, tk.END)
        self.log_texto.delete(1.0, tk.END)
        
        self.buscar_btn.config(state='disabled')
        self.progress.start()
        self.status_label.config(text="Buscando...")
        
        # Executar em thread separada
        thread = threading.Thread(target=self.executar_busca, args=(mes, dia, nota))
        thread.daemon = True
        thread.start()
    
    def executar_busca(self, mes, dia, nota):
        try:
            resultados = self.buscador.buscar_canhotos(mes, dia, nota)
            self.root.after(0, self.mostrar_resultados, resultados, mes, dia, nota)
        except Exception as e:
            self.root.after(0, self.mostrar_erro, str(e))
    
    def mostrar_resultados(self, resultados, mes, dia, nota):
        self.progress.stop()
        self.buscar_btn.config(state='normal')
        self.status_label.config(text="Busca finalizada")
        
        # Mostrar log
        for mensagem in self.buscador.debug_info:
            self.log_texto.insert(tk.END, f"{mensagem}\n")
        
        if isinstance(resultados, str):
            # Mensagem de erro
            self.resultados_texto.insert(tk.END, f"❌ {resultados}\n\n", 'erro')
            self.resultados_texto.insert(tk.END, 
                "Sugestões:\n"
                "• Verifique se o mês e dia estão corretos\n"
                "• Confirme o número da nota\n"
                "• Verifique o log para detalhes", 'aviso')
            return
        
        if not resultados:
            self.resultados_texto.insert(tk.END, 
                f"❌ Nenhum canhoto encontrado para a nota {nota} em {dia}/{mes}\n\n", 'erro')
            self.resultados_texto.insert(tk.END,
                "Possíveis causas:\n"
                "• A nota pode estar em outro mês/dia\n"
                "• O número da nota pode estar incorreto\n"
                "• O arquivo PDF pode estar corrompido\n"
                "• O texto não está sendo reconhecido\n"
                "• Tente ativar/desativar o OCR\n", 'aviso')
            return
        
        # Mostrar resultados
        self.resultados_texto.insert(tk.END, 
            f"✅ Encontrado(s) {len(resultados)} canhoto(s) para a nota {nota}:\n\n", 'sucesso')
        
        for i, resultado in enumerate(resultados, 1):
            self.resultados_texto.insert(tk.END, 
                f"📄 Resultado {i}:\n"
                f"   Arquivo: {resultado['nome_arquivo']}\n"
                f"   Pasta: {resultado['pasta_dia']}\n"
                f"   Página: {resultado['pagina']}\n"
                f"   Caminho: {resultado['arquivo']}\n", 'normal')
            
            # Botão para abrir o PDF
            self.resultados_texto.insert(tk.END, "   ")
            self.resultados_texto.window_create(tk.END, window=ttk.Button(
                self.resultados_texto, text="Abrir PDF", 
                command=lambda r=resultado: self.abrir_resultado(r),
                width=10
            ))
            self.resultados_texto.insert(tk.END, "\n\n")
        
        # Configurar tags para cores
        self.resultados_texto.tag_configure('erro', foreground='red', font=('Arial', 10, 'bold'))
        self.resultados_texto.tag_configure('aviso', foreground='orange', font=('Arial', 9))
        self.resultados_texto.tag_configure('sucesso', foreground='green', font=('Arial', 11, 'bold'))
        self.resultados_texto.tag_configure('normal', font=('Arial', 9))
    
    def abrir_resultado(self, resultado):
        """Abre o PDF na página específica"""
        sucesso = self.buscador.abrir_pdf_pagina(resultado['arquivo'], resultado['pagina'])
        if not sucesso:
            messagebox.showerror("Erro", f"Não foi possível abrir o arquivo:\n{resultado['arquivo']}")
    
    def mostrar_erro(self, erro):
        self.progress.stop()
        self.buscar_btn.config(state='normal')
        self.status_label.config(text="Erro na busca")
        self.resultados_texto.insert(tk.END, f"❌ Erro durante a busca:\n{erro}\n", 'erro')
    
    def limpar_campos(self):
        self.mes_combobox.set('')
        self.dia_combobox.set('')
        self.nota_entry.delete(0, tk.END)
        self.resultados_texto.delete(1.0, tk.END)
        self.log_texto.delete(1.0, tk.END)
        self.status_label.config(text="Pronto para buscar")

def main():
    try:
        root = tk.Tk()
        app = InterfaceBuscadorCanhotos(root)
        root.mainloop()
    except Exception as e:
        print(f"Erro ao iniciar aplicação: {e}")
        messagebox.showerror("Erro", f"Erro ao iniciar aplicação:\n{e}")

if __name__ == "__main__":
    main()