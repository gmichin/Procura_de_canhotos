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
import platform
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

# Configurar o caminho do Tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

class LocalizadorNotasDevolucoes:
    def __init__(self, caminho_base):
        self.caminho_base = caminho_base
        self.resultados = []
        self.debug_info = []
        self._stop_event = threading.Event()
        
    def adicionar_debug(self, mensagem):
        """Adiciona mensagem de debug"""
        self.debug_info.append(mensagem)
        print(f"DEBUG: {mensagem}")
    
    def stop_search(self):
        """Para a busca em andamento"""
        self._stop_event.set()
    
    def reset_search(self):
        """Reinicia o estado da busca"""
        self._stop_event.clear()
    
    def converter_pdf_para_imagem_otimizado(self, pdf_path, pagina_num):
        """Converte uma página PDF em imagem com otimizações"""
        if self._stop_event.is_set():
            return None
            
        try:
            doc = fitz.open(pdf_path)
            pagina = doc.load_page(pagina_num)
            
            matriz = fitz.Matrix(2.0, 2.0)
            pix = pagina.get_pixmap(matrix=matriz)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            doc.close()
            return img
            
        except Exception as e:
            self.adicionar_debug(f"Erro ao converter PDF: {e}")
            return None
    
    def preprocessar_imagem_otimizado(self, imagem):
        """Pré-processa a imagem de forma otimizada"""
        try:
            if imagem.mode != 'L':
                imagem = imagem.convert('L')
            
            max_size = 2000
            if max(imagem.size) > max_size:
                ratio = max_size / max(imagem.size)
                new_size = (int(imagem.size[0] * ratio), int(imagem.size[1] * ratio))
                imagem = imagem.resize(new_size, Image.Resampling.LANCZOS)
            
            return imagem
        except Exception as e:
            self.adicionar_debug(f"Erro no pré-processamento: {e}")
            return imagem
    
    def buscar_texto_ocr_otimizado(self, imagem, numero_nota):
        """OCR otimizado"""
        if self._stop_event.is_set():
            return False
            
        try:
            imagem_processada = self.preprocessar_imagem_otimizado(imagem)
            
            config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789'
            texto = pytesseract.image_to_string(imagem_processada, config=config)
            texto_limpo = re.sub(r'\s+', '', texto)
            
            # Verificar número completo
            if numero_nota in texto_limpo:
                return True
            
            # Verificações parciais
            if len(numero_nota) >= 6:
                combinacoes = [
                    numero_nota[-10:],
                    numero_nota[-8:],  
                    numero_nota[-6:],
                    numero_nota[-4:],
                ]
                
                for combo in combinacoes:
                    if combo in texto_limpo:
                        return True
            
            return False
            
        except Exception as e:
            self.adicionar_debug(f"Erro no OCR: {e}")
            return False
    
    def buscar_texto_direto_pdf(self, pdf_path, numero_nota):
        """Busca textual no PDF"""
        if self._stop_event.is_set():
            return []
            
        try:
            paginas_encontradas = []
            
            with open(pdf_path, 'rb') as arquivo:
                leitor = PyPDF2.PdfReader(arquivo)
                total_paginas = len(leitor.pages)
                
                for num_pagina in range(total_paginas):
                    if self._stop_event.is_set():
                        break
                        
                    pagina = leitor.pages[num_pagina]
                    texto = pagina.extract_text()
                    
                    if texto:
                        texto_limpo = re.sub(r'\s+', '', texto)
                        
                        if numero_nota in texto_limpo:
                            paginas_encontradas.append(num_pagina + 1)
                            return paginas_encontradas
                        
                        if len(numero_nota) >= 6:
                            partes = [
                                numero_nota[-10:],
                                numero_nota[-8:], 
                                numero_nota[-6:],
                            ]
                            for parte in partes:
                                if parte in texto_limpo:
                                    paginas_encontradas.append(num_pagina + 1)
                                    return paginas_encontradas
            
            return paginas_encontradas
        except Exception as e:
            return []
    
    def buscar_nome_arquivo(self, pdf_path, numero_nota):
        """Verifica se o número da nota está no nome do arquivo"""
        nome_arquivo = os.path.basename(pdf_path)
        nome_sem_ext = os.path.splitext(nome_arquivo)[0]
        nome_limpo = re.sub(r'[^\d]', '', nome_sem_ext)
        
        if numero_nota in nome_limpo:
            return True
        
        if len(numero_nota) >= 6:
            partes = [
                numero_nota[-8:],
                numero_nota[-6:],
            ]
            for parte in partes:
                if parte in nome_limpo:
                    return True
        
        return False
    
    def processar_pdf(self, pdf_info):
        """Processa um PDF"""
        pdf_path, numero_nota = pdf_info
        
        if self._stop_event.is_set():
            return None
            
        try:
            resultados_paginas = []
            nome_arquivo = os.path.basename(pdf_path)
            
            # 1. Buscar no nome do arquivo
            if self.buscar_nome_arquivo(pdf_path, numero_nota):
                resultados_paginas.append(1)
                return pdf_path, resultados_paginas
            
            # 2. Busca textual direta
            paginas_texto = self.buscar_texto_direto_pdf(pdf_path, numero_nota)
            if paginas_texto:
                resultados_paginas.extend(paginas_texto)
                return pdf_path, resultados_paginas
            
            # 3. OCR
            with open(pdf_path, 'rb') as arquivo:
                leitor = PyPDF2.PdfReader(arquivo)
                total_paginas = len(leitor.pages)
            
            for num_pagina in range(total_paginas):
                if self._stop_event.is_set():
                    break
                    
                imagem = self.converter_pdf_para_imagem_otimizado(pdf_path, num_pagina)
                if imagem:
                    encontrado = self.buscar_texto_ocr_otimizado(imagem, numero_nota)
                    if encontrado:
                        resultados_paginas.append(num_pagina + 1)
                        break
            
            return pdf_path, resultados_paginas if resultados_paginas else None
            
        except Exception as e:
            return pdf_path, None
    
    def buscar_nota(self, mes, dia, numero_nota, max_workers=2):
        """Busca principal"""
        self.reset_search()
        self.resultados = []
        self.debug_info = []
        numero_nota = numero_nota.strip()
        
        self.adicionar_debug(f"Buscando nota {numero_nota} para {dia}/{mes}")
        
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
            return f"Pasta do mês {mes} não encontrada"
        
        # Buscar nas pastas de dias
        dia_procurado = f"{int(dia):02d}"
        pdfs_para_processar = []
        
        for pasta_dia in os.listdir(caminho_mes):
            if self._stop_event.is_set():
                break
                
            caminho_dia = os.path.join(caminho_mes, pasta_dia)
            
            if os.path.isdir(caminho_dia) and dia_procurado in pasta_dia:
                for arquivo in os.listdir(caminho_dia):
                    if arquivo.lower().endswith('.pdf'):
                        pdf_path = os.path.join(caminho_dia, arquivo)
                        pdfs_para_processar.append((
                            pdf_path,
                            numero_nota,
                            pasta_dia,
                            arquivo
                        ))
        
        if not pdfs_para_processar:
            return f"Nenhum PDF encontrado para o dia {dia}"
        
        # Processar PDFs em paralelo
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(self.processar_pdf, (pdf_path, numero_nota)): 
                (pdf_path, pasta_dia, nome_arquivo) 
                for pdf_path, numero_nota, pasta_dia, nome_arquivo in pdfs_para_processar
            }
            
            for future in as_completed(futures):
                if self._stop_event.is_set():
                    executor.shutdown(wait=False)
                    break
                    
                pdf_path, pasta_dia, nome_arquivo = futures[future]
                
                try:
                    resultado_path, paginas = future.result()
                    
                    if paginas:
                        for pagina in paginas:
                            self.resultados.append({
                                'arquivo': pdf_path,
                                'pagina': pagina,
                                'pasta_dia': pasta_dia,
                                'nome_arquivo': nome_arquivo
                            })
                            self.adicionar_debug(f"Encontrado: {nome_arquivo} página {pagina}")
                            
                except Exception:
                    pass
        
        return self.resultados

    def abrir_pdf_pagina(self, caminho_pdf, numero_pagina):
        """Abre o PDF na página específica"""
        try:
            sistema = platform.system()
            
            if sistema == "Windows":
                try:
                    subprocess.run(['SumatraPDF', f'-page={numero_pagina}', caminho_pdf], 
                                 timeout=10, check=True)
                    return True
                except:
                    os.startfile(caminho_pdf)
                    messagebox.showinfo("Abrir PDF", 
                                      f"Arquivo: {os.path.basename(caminho_pdf)}\n"
                                      f"Vá manualmente para a página: {numero_pagina}")
                    return True
                    
            elif sistema == "Darwin":
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
            messagebox.showerror("Erro", f"Não foi possível abrir o PDF:\n{e}")
            return False

class InterfaceLocalizadorDevolucoes:
    def __init__(self, root):
        self.root = root
        self.root.title("Localizador de Notas de Devoluções")
        self.root.geometry("800x600")
        
        self.localizador = LocalizadorNotasDevolucoes(r"Z:\NOTAS-CANHOTOS-DEVOLUÇÕES\DEVOLUÇÕES")
        self.busca_ativa = False
        
        self.criar_interface()
    
    def criar_interface(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        titulo = ttk.Label(main_frame, text="Localizador de Notas de Devoluções", 
                          font=('Arial', 14, 'bold'))
        titulo.grid(row=0, column=0, columnspan=3, pady=10)
        
        ttk.Label(main_frame, text="Mês:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.mes_entry = ttk.Combobox(main_frame, values=[
            'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
            'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
        ], state='readonly', width=15)
        self.mes_entry.grid(row=1, column=1, sticky=tk.W, pady=5, padx=5)
        
        ttk.Label(main_frame, text="Dia:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.dia_entry = ttk.Spinbox(main_frame, from_=1, to=31, width=8)
        self.dia_entry.grid(row=2, column=1, sticky=tk.W, pady=5, padx=5)
        
        ttk.Label(main_frame, text="Número da Nota:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.nota_entry = ttk.Entry(main_frame, width=20)
        self.nota_entry.grid(row=3, column=1, sticky=tk.W, pady=5, padx=5)
        
        botoes_frame = ttk.Frame(main_frame)
        botoes_frame.grid(row=4, column=0, columnspan=3, pady=10)
        
        self.buscar_btn = ttk.Button(botoes_frame, text="Buscar Nota", command=self.iniciar_busca)
        self.buscar_btn.pack(side=tk.LEFT, padx=5)
        
        self.parar_btn = ttk.Button(botoes_frame, text="Parar Busca", 
                                   command=self.parar_busca, state='disabled')
        self.parar_btn.pack(side=tk.LEFT, padx=5)
        
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.status_var = tk.StringVar(value="Pronto para buscar")
        status_label = ttk.Label(main_frame, textvariable=self.status_var)
        status_label.grid(row=6, column=0, columnspan=3, pady=5)
        
        notebook = ttk.Notebook(main_frame)
        notebook.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        self.aba_resultados = ttk.Frame(notebook)
        notebook.add(self.aba_resultados, text="Resultados")
        
        self.aba_debug = ttk.Frame(notebook)
        notebook.add(self.aba_debug, text="Log")
        
        main_frame.columnconfigure(2, weight=1)
        main_frame.rowconfigure(7, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        self.configurar_areas_texto()
        self.nota_entry.focus()
    
    def configurar_areas_texto(self):
        resultados_frame = ttk.Frame(self.aba_resultados)
        resultados_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.resultados_texto = tk.Text(resultados_frame, height=15, wrap=tk.WORD)
        scrollbar_resultados = ttk.Scrollbar(resultados_frame, orient="vertical", 
                                           command=self.resultados_texto.yview)
        self.resultados_texto.configure(yscrollcommand=scrollbar_resultados.set)
        
        self.resultados_texto.pack(side="left", fill="both", expand=True)
        scrollbar_resultados.pack(side="right", fill="y")
        
        debug_frame = ttk.Frame(self.aba_debug)
        debug_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.debug_texto = tk.Text(debug_frame, height=15, wrap=tk.WORD, 
                                 bg='black', fg='white')
        scrollbar_debug = ttk.Scrollbar(debug_frame, orient="vertical", 
                                      command=self.debug_texto.yview)
        self.debug_texto.configure(yscrollcommand=scrollbar_debug.set)
        
        self.debug_texto.pack(side="left", fill="both", expand=True)
        scrollbar_debug.pack(side="right", fill="y")
    
    def iniciar_busca(self):
        if self.busca_ativa:
            return
            
        mes = self.mes_entry.get()
        dia = self.dia_entry.get()
        nota = self.nota_entry.get()
        
        if not all([mes, dia, nota]):
            messagebox.showerror("Erro", "Preencha todos os campos")
            return
        
        self.resultados_texto.delete(1.0, tk.END)
        self.debug_texto.delete(1.0, tk.END)
        
        self.busca_ativa = True
        self.buscar_btn.config(state='disabled')
        self.parar_btn.config(state='normal')
        self.progress.start()
        self.status_var.set("Buscando...")
        
        thread = threading.Thread(target=self.executar_busca, args=(mes, dia, nota))
        thread.daemon = True
        thread.start()
    
    def parar_busca(self):
        if self.busca_ativa:
            self.localizador.stop_search()
            self.busca_ativa = False
            self.status_var.set("Busca interrompida")
    
    def executar_busca(self, mes, dia, nota):
        try:
            start_time = time.time()
            resultados = self.localizador.buscar_nota(mes, dia, nota)
            end_time = time.time()
            
            tempo_decorrido = end_time - start_time
            self.localizador.adicionar_debug(f"Tempo total: {tempo_decorrido:.2f} segundos")
            
            self.root.after(0, self.mostrar_resultados, resultados, mes, dia, nota, tempo_decorrido)
            
        except Exception as e:
            self.root.after(0, self.mostrar_erro, str(e))
    
    def mostrar_resultados(self, resultados, mes, dia, nota, tempo_decorrido):
        self.busca_ativa = False
        self.progress.stop()
        self.buscar_btn.config(state='normal')
        self.parar_btn.config(state='disabled')
        
        for mensagem in self.localizador.debug_info:
            self.debug_texto.insert(tk.END, mensagem + '\n')
        
        self.resultados_texto.insert(tk.END, 
            f"Busca concluída em {tempo_decorrido:.2f} segundos\n\n")
        
        if isinstance(resultados, str):
            self.resultados_texto.insert(tk.END, f"ERRO: {resultados}\n", 'erro')
            self.status_var.set("Erro na busca")
            return
        
        if not resultados:
            self.resultados_texto.insert(tk.END, 
                f"Nenhuma nota {nota} encontrada para {dia} de {mes}\n", 'aviso')
            self.status_var.set("Nenhum resultado")
            return
        
        self.resultados_texto.insert(tk.END, 
            f"Encontrada(s) {len(resultados)} ocorrência(s):\n\n", 'sucesso')
        
        for i, resultado in enumerate(resultados, 1):
            self.resultados_texto.insert(tk.END, 
                f"{i}. Arquivo: {resultado['nome_arquivo']}\n"
                f"   Pasta: {resultado['pasta_dia']}\n"
                f"   Página: {resultado['pagina']}\n", 'normal')
            
            self.resultados_texto.insert(tk.END, "   ")
            self.resultados_texto.window_create(tk.END, window=ttk.Button(
                self.resultados_texto, text="Abrir PDF", 
                command=lambda r=resultado: self.abrir_resultado(r),
                width=10
            ))
            self.resultados_texto.insert(tk.END, "\n\n")
        
        self.resultados_texto.tag_configure('erro', foreground='red')
        self.resultados_texto.tag_configure('aviso', foreground='orange')
        self.resultados_texto.tag_configure('sucesso', foreground='green')
        self.resultados_texto.tag_configure('normal', foreground='black')
        
        self.status_var.set(f"Encontrado: {len(resultados)} resultado(s)")
    
    def abrir_resultado(self, resultado):
        sucesso = self.localizador.abrir_pdf_pagina(resultado['arquivo'], resultado['pagina'])
        if not sucesso:
            messagebox.showerror("Erro", f"Não foi possível abrir o arquivo")
    
    def mostrar_erro(self, erro):
        self.busca_ativa = False
        self.progress.stop()
        self.buscar_btn.config(state='normal')
        self.parar_btn.config(state='disabled')
        self.resultados_texto.insert(tk.END, f"ERRO: {erro}\n", 'erro')
        self.status_var.set("Erro na busca")

def main():
    root = tk.Tk()
    app = InterfaceLocalizadorDevolucoes(root)
    root.mainloop()

if __name__ == "__main__":
    main()