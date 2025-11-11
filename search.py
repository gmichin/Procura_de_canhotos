import os
import PyPDF2
import pytesseract
from PIL import Image, ImageEnhance
import fitz  # PyMuPDF
import re
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import subprocess
import platform
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
import io
import numpy as np

# Configurar o caminho do Tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

class LocalizadorBase:
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
                    f"_{dia_procurado}_" in pasta or
                    pasta == dia_procurado):
                    pastas_encontradas.append({
                        'nome': pasta,
                        'caminho': caminho_pasta
                    })
        
        return pastas_encontradas

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
            self.adicionar_debug(f"Erro ao abrir PDF: {e}")
            messagebox.showerror("Erro", f"Não foi possível abrir o PDF:\n{e}")
            return False

class OCRMultiOrientacao:
    """Classe para lidar com OCR em múltiplas orientações"""
    
    @staticmethod
    def detectar_orientacao_texto(imagem):
        """Detecta a orientação do texto usando OCR do Tesseract"""
        try:
            # Configuração para detecção de orientação e script
            osd = pytesseract.image_to_osd(imagem, config='--psm 0')
            
            # Extrair ângulo de rotação do resultado OSD
            rotacao = 0
            linhas = osd.split('\n')
            for linha in linhas:
                if 'Rotate:' in linha:
                    rotacao = int(linha.split(':')[1].strip())
                    break
            
            return rotacao
        except:
            return 0  # Se falhar, assume orientação normal
    
    @staticmethod
    def rotacionar_imagem(imagem, angulo):
        """Rotaciona a imagem pelo ângulo especificado"""
        if angulo == 0:
            return imagem
        return imagem.rotate(angulo, expand=True, resample=Image.BICUBIC, fillcolor='white')
    
    @staticmethod
    def tentar_todas_orientacoes(imagem, numero_nota, config_ocr):
        """Tenta encontrar o texto em todas as orientações possíveis"""
        orientacoes = [0, 90, 180, 270]  # Todas as orientações possíveis
        
        for angulo in orientacoes:
            try:
                # Rotacionar imagem
                if angulo != 0:
                    imagem_rotacionada = imagem.rotate(angulo, expand=True, 
                                                     resample=Image.BICUBIC, 
                                                     fillcolor='white')
                else:
                    imagem_rotacionada = imagem
                
                # Fazer OCR
                texto = pytesseract.image_to_string(imagem_rotacionada, config=config_ocr)
                texto_limpo = re.sub(r'\s+', '', texto)
                
                # Verificar se encontrou
                if numero_nota in texto_limpo:
                    return True, angulo
                
                # Verificar partes do número
                if len(numero_nota) >= 6:
                    partes = [
                        numero_nota[-10:],
                        numero_nota[-8:],  
                        numero_nota[-6:],
                        numero_nota[-4:],
                    ]
                    
                    for parte in partes:
                        if parte in texto_limpo:
                            return True, angulo
                            
            except Exception as e:
                continue
        
        return False, 0
    
    @staticmethod
    def melhorar_imagem_para_ocr(imagem):
        """Melhora a imagem para OCR em qualquer orientação"""
        try:
            # Converter para escala de cinza se necessário
            if imagem.mode != 'L':
                imagem = imagem.convert('L')
            
            # Aumentar contraste
            enhancer = ImageEnhance.Contrast(imagem)
            imagem = enhancer.enhance(2.0)
            
            # Aumentar nitidez
            enhancer = ImageEnhance.Sharpness(imagem)
            imagem = enhancer.enhance(1.5)
            
            # Ajustar brilho se necessário
            enhancer = ImageEnhance.Brightness(imagem)
            imagem = enhancer.enhance(1.1)
            
            return imagem
        except Exception as e:
            return imagem

class LocalizadorNotasDevolucoes(LocalizadorBase):
    def __init__(self, caminho_base):
        super().__init__(caminho_base)
        self.ocr_multi = OCRMultiOrientacao()
        
    def converter_pdf_para_imagem_otimizado(self, pdf_path, pagina_num):
        """Converte uma página PDF em imagem com otimizações"""
        if self._stop_event.is_set():
            return None
            
        try:
            doc = fitz.open(pdf_path)
            pagina = doc.load_page(pagina_num)
            
            # Aumentar resolução para melhor detecção de orientação
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
    
    def buscar_texto_ocr_multiorientacao(self, imagem, numero_nota):
        """OCR otimizado para todas as orientações"""
        if self._stop_event.is_set():
            return False
            
        try:
            # Melhorar imagem para OCR
            imagem_melhorada = self.ocr_multi.melhorar_imagem_para_ocr(imagem)
            
            config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789'
            
            # Tentar todas as orientações
            encontrado, angulo = self.ocr_multi.tentar_todas_orientacoes(
                imagem_melhorada, numero_nota, config
            )
            
            if encontrado:
                self.adicionar_debug(f"Texto encontrado com rotação de {angulo} graus")
                return True
            
            return False
            
        except Exception as e:
            self.adicionar_debug(f"Erro no OCR multi-orientação: {e}")
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
            
            # 3. OCR multi-orientação
            with open(pdf_path, 'rb') as arquivo:
                leitor = PyPDF2.PdfReader(arquivo)
                total_paginas = len(leitor.pages)
            
            for num_pagina in range(total_paginas):
                if self._stop_event.is_set():
                    break
                    
                imagem = self.converter_pdf_para_imagem_otimizado(pdf_path, num_pagina)
                if imagem:
                    encontrado = self.buscar_texto_ocr_multiorientacao(imagem, numero_nota)
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
        
        mes_pasta = self.normalizar_mes(mes)
        caminho_mes = os.path.join(self.caminho_base, mes_pasta)
        
        if not os.path.exists(caminho_mes):
            return f"Pasta do mês {mes} não encontrada"
        
        # Buscar nas pastas de dias
        pastas_dia = self.encontrar_pastas_dia(caminho_mes, dia)
        pdfs_para_processar = []
        
        for pasta_info in pastas_dia:
            if self._stop_event.is_set():
                break
                
            caminho_dia = pasta_info['caminho']
            pasta_nome = pasta_info['nome']
            
            for arquivo in os.listdir(caminho_dia):
                if arquivo.lower().endswith('.pdf'):
                    pdf_path = os.path.join(caminho_dia, arquivo)
                    pdfs_para_processar.append((
                        pdf_path,
                        numero_nota,
                        pasta_nome,
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

class LocalizadorNotasFiscais(LocalizadorBase):
    def __init__(self, caminho_base):
        super().__init__(caminho_base)
        self.ocr_multi = OCRMultiOrientacao()
    
    def converter_pdf_para_imagem_otimizado(self, pdf_path, pagina_num):
        """Converte uma página PDF em imagem com otimizações"""
        if self._stop_event.is_set():
            return None
            
        try:
            doc = fitz.open(pdf_path)
            pagina = doc.load_page(pagina_num)
            
            # Usar resolução adequada para detecção de orientação
            matriz = fitz.Matrix(1.8, 1.8)
            pix = pagina.get_pixmap(matrix=matriz)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            doc.close()
            return img
            
        except Exception as e:
            self.adicionar_debug(f"Erro ao converter PDF {pdf_path}, página {pagina_num}: {e}")
            return None
    
    def preprocessar_imagem_otimizado(self, imagem):
        """Pré-processa a imagem de forma otimizada"""
        try:
            # Converter para escala de cinza
            if imagem.mode != 'L':
                imagem = imagem.convert('L')
            
            # Redimensionar imagem se for muito grande (mantém proporção)
            max_size = 1600
            if max(imagem.size) > max_size:
                ratio = max_size / max(imagem.size)
                new_size = (int(imagem.size[0] * ratio), int(imagem.size[1] * ratio))
                imagem = imagem.resize(new_size, Image.Resampling.LANCZOS)
            
            return imagem
        except Exception as e:
            self.adicionar_debug(f"Erro no pré-processamento: {e}")
            return imagem
    
    def buscar_texto_ocr_multiorientacao(self, imagem, numero_nota):
        """OCR otimizado para todas as orientações"""
        if self._stop_event.is_set():
            return False
            
        try:
            # Melhorar imagem primeiro
            imagem_melhorada = self.ocr_multi.melhorar_imagem_para_ocr(imagem)
            
            # Configuração única otimizada para números
            config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789'
            
            # Tentar todas as orientações
            encontrado, angulo = self.ocr_multi.tentar_todas_orientacoes(
                imagem_melhorada, numero_nota, config
            )
            
            if encontrado:
                self.adicionar_debug(f"Nota encontrada com rotação de {angulo}°")
                return True
            
            return False
            
        except Exception as e:
            self.adicionar_debug(f"Erro no OCR multi-orientação: {e}")
            return False
    
    def buscar_texto_direto_pdf_otimizado(self, pdf_path, numero_nota):
        """Busca textual otimizada com leitura em chunks"""
        if self._stop_event.is_set():
            return []
            
        try:
            paginas_encontradas = []
            
            with open(pdf_path, 'rb') as arquivo:
                leitor = PyPDF2.PdfReader(arquivo)
                
                # Verificar número total de páginas primeiro
                total_paginas = len(leitor.pages)
                
                for num_pagina in range(total_paginas):
                    if self._stop_event.is_set():
                        break
                        
                    pagina = leitor.pages[num_pagina]
                    texto = pagina.extract_text()
                    
                    if texto and numero_nota in texto:
                        paginas_encontradas.append(num_pagina + 1)
                        self.adicionar_debug(f"Encontrado via texto direto em {pdf_path} página {num_pagina + 1}")
                        break  # Parar na primeira ocorrência
            
            return paginas_encontradas
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
    
    def processar_pdf_paralelo(self, pdf_info):
        """Processa um PDF em paralelo"""
        pdf_path, numero_nota = pdf_info
        
        if self._stop_event.is_set():
            return None
            
        try:
            resultados_paginas = []
            
            # Estratégia 1: Buscar no nome do arquivo (mais rápido)
            if self.buscar_nome_arquivo(pdf_path, numero_nota):
                resultados_paginas.append(1)
                return pdf_path, resultados_paginas
            
            # Estratégia 2: Busca textual direta
            paginas_texto = self.buscar_texto_direto_pdf_otimizado(pdf_path, numero_nota)
            if paginas_texto:
                resultados_paginas.extend(paginas_texto)
                return pdf_path, resultados_paginas
            
            # Estratégia 3: OCR multi-orientação
            if not resultados_paginas:
                self.adicionar_debug(f"Usando OCR multi-orientação para: {os.path.basename(pdf_path)}")
                
                with open(pdf_path, 'rb') as arquivo:
                    leitor = PyPDF2.PdfReader(arquivo)
                    total_paginas = len(leitor.pages)
                
                # Limitar número de páginas para OCR
                max_paginas_ocr = min(total_paginas, 10)  # Máximo 10 páginas por PDF
                
                for num_pagina in range(max_paginas_ocr):
                    if self._stop_event.is_set():
                        break
                        
                    imagem = self.converter_pdf_para_imagem_otimizado(pdf_path, num_pagina)
                    if imagem:
                        encontrado = self.buscar_texto_ocr_multiorientacao(imagem, numero_nota)
                        if encontrado:
                            resultados_paginas.append(num_pagina + 1)
                            break
            
            return pdf_path, resultados_paginas if resultados_paginas else None
            
        except Exception as e:
            self.adicionar_debug(f"Erro ao processar {pdf_path}: {e}")
            return pdf_path, None
    
    def buscar_nota_otimizada(self, mes, dia, numero_nota, max_workers=3):
        """Busca otimizada com processamento paralelo"""
        self.reset_search()
        self.resultados = []
        self.debug_info = []
        numero_nota = numero_nota.strip()
        
        self.adicionar_debug(f"Iniciando busca otimizada: {dia}/{mes} - Nota: {numero_nota}")
        
        mes_pasta = self.normalizar_mes(mes)
        caminho_mes = os.path.join(self.caminho_base, mes_pasta)
        
        if not os.path.exists(caminho_mes):
            return f"Pasta do mês {mes} não encontrada em {caminho_mes}"
        
        # Buscar nas pastas de dias
        pastas_dia = self.encontrar_pastas_dia(caminho_mes, dia)
        pdfs_para_processar = []
        
        for pasta_info in pastas_dia:
            if self._stop_event.is_set():
                break
                
            caminho_dia = pasta_info['caminho']
            pasta_nome = pasta_info['nome']
            
            self.adicionar_debug(f"Buscando na pasta: {pasta_nome}")
            
            # Coletar todos os PDFs primeiro
            for arquivo in os.listdir(caminho_dia):
                if arquivo.lower().endswith('.pdf'):
                    pdfs_para_processar.append((
                        os.path.join(caminho_dia, arquivo),
                        numero_nota,
                        pasta_nome,
                        arquivo
                    ))
        
        if not pdfs_para_processar:
            return f"Nenhum PDF encontrado para o dia {dia} em {mes}"
        
        # Processar PDFs em paralelo
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(self.processar_pdf_paralelo, (pdf_path, numero_nota)): 
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
                            
                except Exception as e:
                    self.adicionar_debug(f"Erro no processamento paralelo: {e}")
        
        self.adicionar_debug(f"Busca finalizada. {len(self.resultados)} resultado(s) encontrado(s)")
        return self.resultados

class BuscadorCanhotosAvancado(LocalizadorBase):
    def __init__(self, caminho_base):
        super().__init__(caminho_base)
        self.usar_ocr = True  # Ativar OCR como fallback
        self.ocr_multi = OCRMultiOrientacao()
    
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
    
    def buscar_com_ocr_multiorientacao(self, pdf_path, numero_nota, num_pagina):
        """Busca o texto usando OCR em todas as orientações"""
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
            
            # Tentar todas as orientações
            encontrado, angulo = self.ocr_multi.tentar_todas_orientacoes(
                imagem_melhorada, numero_nota, config_tesseract
            )
            
            doc.close()
            
            if encontrado:
                self.adicionar_debug(f"OCR encontrou '{numero_nota}' na página {num_pagina + 1} (rotação: {angulo}°)")
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
                if self._stop_event.is_set():
                    break
                    
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
                
                # Estratégia 3: OCR multi-orientação (se as anteriores não funcionaram)
                if not paginas_encontradas and self.usar_ocr:
                    self.adicionar_debug(f"Tentando OCR multi-orientação na página {num_pagina + 1}...")
                    if self.buscar_com_ocr_multiorientacao(caminho_pdf, numero_nota, num_pagina):
                        paginas_encontradas.append(num_pagina + 1)
                        self.adicionar_debug(f"OCR: encontrado na página {num_pagina + 1}")
            
            pdf_document.close()
            return paginas_encontradas
            
        except Exception as e:
            self.adicionar_debug(f"Erro ao processar {caminho_pdf}: {e}")
            return []
    
    def buscar_canhotos(self, mes, dia, numero_nota):
        """Busca principal pelos canhotos"""
        self.reset_search()
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
                if self._stop_event.is_set():
                    break
                    
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

# A classe InterfaceLocalizadorUnificado permanece a mesma
class InterfaceLocalizadorUnificado:
    def __init__(self, root):
        self.root = root
        self.root.title("Localizador Unificado de Notas Fiscais")
        # Tamanho reduzido para caber em telas menores
        self.root.geometry("500x600")
        
        # Caminhos base para cada tipo
        self.caminhos_base = {
            "Canhoto": r"Z:\NOTAS-CANHOTOS-DEVOLUÇÕES\CANHOTOS",
            "Devoluções": r"Z:\NOTAS-CANHOTOS-DEVOLUÇÕES\DEVOLUÇÕES", 
            "Notas de Entrada": r"Z:\NOTAS-CANHOTOS-DEVOLUÇÕES\NOTAS ENTRADA"
        }
        
        self.localizador_atual = None
        self.tipo_busca_atual = None
        self.busca_ativa = False
        
        self.criar_interface()
        self.carregar_meses_disponiveis()
    
    def carregar_meses_disponiveis(self):
        """Carrega os meses disponíveis no caminho base atual"""
        try:
            if self.tipo_busca_atual and self.tipo_busca_atual in self.caminhos_base:
                caminho_base = self.caminhos_base[self.tipo_busca_atual]
                if os.path.exists(caminho_base):
                    pastas = [p for p in os.listdir(caminho_base) 
                             if os.path.isdir(os.path.join(caminho_base, p))]
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
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título
        titulo = ttk.Label(main_frame, text="🔍 LOCALIZADOR UNIFICADO DE NOTAS FISCAIS", 
                          font=('Arial', 14, 'bold'))
        titulo.grid(row=0, column=0, columnspan=3, pady=(0, 15))
        
        # Seletor de tipo de busca
        ttk.Label(main_frame, text="Tipo de Busca:", font=('Arial', 10, 'bold')).grid(
            row=1, column=0, sticky=tk.W, pady=8)
        
        self.tipo_var = tk.StringVar(value="Canhoto")
        tipo_frame = ttk.Frame(main_frame)
        tipo_frame.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=8)
        
        tipos = ["Canhoto", "Devoluções", "Notas de Entrada"]
        for i, tipo in enumerate(tipos):
            rb = ttk.Radiobutton(tipo_frame, text=tipo, variable=self.tipo_var, 
                               value=tipo, command=self.tipo_selecionado)
            rb.pack(side=tk.LEFT, padx=8)
        
        # Campos de entrada
        ttk.Label(main_frame, text="Mês:", font=('Arial', 9)).grid(
            row=2, column=0, sticky=tk.W, pady=4)
        self.mes_combobox = ttk.Combobox(main_frame, width=18, font=('Arial', 9))
        self.mes_combobox.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=4, padx=4)
        
        ttk.Label(main_frame, text="Dia:", font=('Arial', 9)).grid(
            row=3, column=0, sticky=tk.W, pady=4)
        self.dia_combobox = ttk.Combobox(main_frame, width=8, font=('Arial', 9))
        self.dia_combobox['values'] = [f"{i:02d}" for i in range(1, 32)]
        self.dia_combobox.grid(row=3, column=1, sticky=tk.W, pady=4, padx=4)
        
        ttk.Label(main_frame, text="Número da Nota:", font=('Arial', 9)).grid(
            row=4, column=0, sticky=tk.W, pady=4)
        self.nota_entry = ttk.Entry(main_frame, width=20, font=('Arial', 9))
        self.nota_entry.grid(row=4, column=1, sticky=(tk.W, tk.E), pady=4, padx=4)
        
        # Configurações
        settings_frame = ttk.Frame(main_frame)
        settings_frame.grid(row=5, column=0, columnspan=3, pady=8)
        
        self.ocr_var = tk.BooleanVar(value=True)
        self.ocr_check = ttk.Checkbutton(settings_frame, text="Usar OCR (recomendado)", 
                                       variable=self.ocr_var)
        self.ocr_check.pack(side=tk.LEFT, padx=8)
        
        # Botões
        botoes_frame = ttk.Frame(main_frame)
        botoes_frame.grid(row=6, column=0, columnspan=3, pady=12)
        
        self.buscar_btn = ttk.Button(botoes_frame, text="🔎 Iniciar Busca", 
                                   command=self.iniciar_busca, width=18)
        self.buscar_btn.pack(side=tk.LEFT, padx=4)
        
        self.parar_btn = ttk.Button(botoes_frame, text="⏹️ Parar", 
                                   command=self.parar_busca, state='disabled', width=12)
        self.parar_btn.pack(side=tk.LEFT, padx=4)
        
        self.limpar_btn = ttk.Button(botoes_frame, text="🗑️ Limpar", 
                                   command=self.limpar_campos, width=12)
        self.limpar_btn.pack(side=tk.LEFT, padx=4)
        
        # Barra de progresso
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=8)
        
        # Status
        self.status_label = ttk.Label(main_frame, text="Selecione o tipo de busca e preencha os campos", 
                                    font=('Arial', 8))
        self.status_label.grid(row=8, column=0, columnspan=3, pady=4)
        
        # Abas
        notebook = ttk.Notebook(main_frame)
        notebook.grid(row=9, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=8)
        
        # Aba de resultados
        self.aba_resultados = ttk.Frame(notebook)
        notebook.add(self.aba_resultados, text="📄 Resultados")
        
        # Aba de log
        self.aba_log = ttk.Frame(notebook)
        notebook.add(self.aba_log, text="📋 Log")
        
        # Configurar grid
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(9, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # Configurar áreas de texto
        self.configurar_areas_texto()
        
        # Bind events
        self.nota_entry.bind('<Return>', lambda e: self.iniciar_busca())
        
        # Inicializar com o primeiro tipo
        self.tipo_selecionado()
    
    def configurar_areas_texto(self):
        # Área de resultados
        resultados_frame = ttk.Frame(self.aba_resultados)
        resultados_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.resultados_texto = tk.Text(resultados_frame, height=15, wrap=tk.WORD, font=('Arial', 9))
        scrollbar_resultados = ttk.Scrollbar(resultados_frame, orient="vertical", command=self.resultados_texto.yview)
        self.resultados_texto.configure(yscrollcommand=scrollbar_resultados.set)
        
        self.resultados_texto.pack(side="left", fill="both", expand=True)
        scrollbar_resultados.pack(side="right", fill="y")
        
        # Área de log
        log_frame = ttk.Frame(self.aba_log)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.log_texto = tk.Text(log_frame, height=15, wrap=tk.WORD, font=('Arial', 8),
                                bg='black', fg='white')
        scrollbar_log = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_texto.yview)
        self.log_texto.configure(yscrollcommand=scrollbar_log.set)
        
        self.log_texto.pack(side="left", fill="both", expand=True)
        scrollbar_log.pack(side="right", fill="y")
    
    def tipo_selecionado(self):
        """Atualiza o localizador quando o tipo de busca é alterado"""
        self.tipo_busca_atual = self.tipo_var.get()
        caminho_base = self.caminhos_base.get(self.tipo_busca_atual)
        
        if self.tipo_busca_atual == "Canhoto":
            self.localizador_atual = BuscadorCanhotosAvancado(caminho_base)
        elif self.tipo_busca_atual == "Devoluções":
            self.localizador_atual = LocalizadorNotasDevolucoes(caminho_base)
        elif self.tipo_busca_atual == "Notas de Entrada":
            self.localizador_atual = LocalizadorNotasFiscais(caminho_base)
        
        self.status_label.config(text=f"Busca configurada para: {self.tipo_busca_atual}")
        self.carregar_meses_disponiveis()
    
    def iniciar_busca(self):
        if self.busca_ativa:
            return
            
        if not self.localizador_atual:
            messagebox.showerror("Erro", "Selecione um tipo de busca primeiro.")
            return
            
        mes = self.mes_combobox.get()
        dia = self.dia_combobox.get()
        nota = self.nota_entry.get().strip()
        
        if not all([mes, dia, nota]):
            messagebox.showerror("Erro", "Por favor, preencha todos os campos.")
            return
        
        if not nota.isdigit():
            messagebox.showerror("Erro", "O número da nota deve conter apenas dígitos.")
            return
        
        # Configurar OCR para Canhotos
        if self.tipo_busca_atual == "Canhoto":
            self.localizador_atual.usar_ocr = self.ocr_var.get()
        
        # Limpar resultados anteriores
        self.resultados_texto.delete(1.0, tk.END)
        self.log_texto.delete(1.0, tk.END)
        
        self.busca_ativa = True
        self.buscar_btn.config(state='disabled')
        self.parar_btn.config(state='normal')
        self.progress.start()
        self.status_label.config(text=f"Buscando {self.tipo_busca_atual}...")
        
        # Executar em thread separada
        thread = threading.Thread(target=self.executar_busca, args=(mes, dia, nota))
        thread.daemon = True
        thread.start()
    
    def parar_busca(self):
        if self.busca_ativa and self.localizador_atual:
            self.localizador_atual.stop_search()
            self.busca_ativa = False
            self.status_label.config(text="Busca interrompida")
    
    def executar_busca(self, mes, dia, nota):
        try:
            start_time = time.time()
            
            if self.tipo_busca_atual == "Canhoto":
                resultados = self.localizador_atual.buscar_canhotos(mes, dia, nota)
            elif self.tipo_busca_atual == "Devoluções":
                resultados = self.localizador_atual.buscar_nota(mes, dia, nota)
            elif self.tipo_busca_atual == "Notas de Entrada":
                resultados = self.localizador_atual.buscar_nota_otimizada(mes, dia, nota)
            
            end_time = time.time()
            tempo_decorrido = end_time - start_time
            
            self.localizador_atual.adicionar_debug(f"Tempo total da busca: {tempo_decorrido:.2f} segundos")
            
            # Atualizar interface na thread principal
            self.root.after(0, self.mostrar_resultados, resultados, mes, dia, nota, tempo_decorrido)
            
        except Exception as e:
            self.root.after(0, self.mostrar_erro, str(e))
    
    def mostrar_resultados(self, resultados, mes, dia, nota, tempo_decorrido):
        self.busca_ativa = False
        self.progress.stop()
        self.buscar_btn.config(state='normal')
        self.parar_btn.config(state='disabled')
        
        # Mostrar log de debug
        for mensagem in self.localizador_atual.debug_info:
            self.log_texto.insert(tk.END, mensagem + '\n')
        
        self.resultados_texto.insert(tk.END, 
            f"Busca de {self.tipo_busca_atual} concluída em {tempo_decorrido:.2f} segundos\n\n")
        
        if isinstance(resultados, str):
            self.resultados_texto.insert(tk.END, f"❌ {resultados}\n\n", 'erro')
            self.status_label.config(text="Erro na busca")
            return
        
        if not resultados:
            self.resultados_texto.insert(tk.END, 
                f"❌ Nenhum {self.tipo_busca_atual.lower()} encontrado para a nota {nota} em {dia}/{mes}\n\n", 'erro')
            self.resultados_texto.insert(tk.END,
                "Sugestões:\n"
                "• Verifique se o mês e dia estão corretos\n"
                "• Confirme o número da nota\n"
                "• Verifique o log para detalhes\n"
                "• Tente ativar/desativar o OCR", 'aviso')
            self.status_label.config(text="Nenhum resultado encontrado")
            return
        
        # Mostrar resultados
        self.resultados_texto.insert(tk.END, 
            f"✅ Encontrado(s) {len(resultados)} {self.tipo_busca_atual.lower()}(s) para a nota {nota}:\n\n", 'sucesso')
        
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
        self.resultados_texto.tag_configure('erro', foreground='red', font=('Arial', 9, 'bold'))
        self.resultados_texto.tag_configure('aviso', foreground='orange', font=('Arial', 8))
        self.resultados_texto.tag_configure('sucesso', foreground='green', font=('Arial', 10, 'bold'))
        self.resultados_texto.tag_configure('normal', font=('Arial', 8))
        
        self.status_label.config(text=f"Busca concluída - {len(resultados)} resultado(s) encontrado(s)")
    
    def abrir_resultado(self, resultado):
        """Abre o PDF na página específica"""
        if self.localizador_atual:
            sucesso = self.localizador_atual.abrir_pdf_pagina(resultado['arquivo'], resultado['pagina'])
            if not sucesso:
                messagebox.showerror("Erro", f"Não foi possível abrir o arquivo:\n{resultado['arquivo']}")
    
    def mostrar_erro(self, erro):
        self.busca_ativa = False
        self.progress.stop()
        self.buscar_btn.config(state='normal')
        self.parar_btn.config(state='disabled')
        self.resultados_texto.insert(tk.END, f"❌ Erro durante a busca:\n{erro}\n", 'erro')
        self.status_label.config(text="Erro na busca")
    
    def limpar_campos(self):
        self.mes_combobox.set('')
        self.dia_combobox.set('')
        self.nota_entry.delete(0, tk.END)
        self.resultados_texto.delete(1.0, tk.END)
        self.log_texto.delete(1.0, tk.END)
        self.status_label.config(text="Campos limpos - Pronto para buscar")

def main():
    try:
        root = tk.Tk()
        app = InterfaceLocalizadorUnificado(root)
        root.mainloop()
    except Exception as e:
        print(f"Erro ao iniciar aplicação: {e}")
        messagebox.showerror("Erro", f"Erro ao iniciar aplicação:\n{e}")

if __name__ == "__main__":
    main()