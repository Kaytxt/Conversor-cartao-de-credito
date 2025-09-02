import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, simpledialog, filedialog
import pandas as pd
import re
import os
import sys
from datetime import datetime
import openpyxl
import fitz  # PyMuPDF para leitura de PDF
from PIL import Image
import pytesseract
import io

def resource_path(relative_path):
    """Obtém o caminho absoluto do recurso, funcionando tanto no desenvolvimento quanto no executável."""
    try:
        # PyInstaller cria uma pasta temp e armazena o caminho em _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

class ConversorFaturas:
    def __init__(self, root):
        self.root = root
        self.root.title("💳 Conversor de Faturas para Excel")
        self.root.geometry("900x700")
        self.root.minsize(900, 700)
        
        # Detecta se está executando como executável
        self.is_executable = getattr(sys, 'frozen', False)
        
        # Cores modernas
        self.cores = {
            'fundo_principal': '#f0f2f5',
            'fundo_card': '#ffffff',
            'santander_red': '#ec0000',
            'sicoob_green': '#00a859',
            'texto_principal': '#1a1a1a',
            'texto_secundario': '#666666',
            'sucesso': '#28a745',
            'erro': '#dc3545',
            'processando': '#007bff',
            'borda': '#e1e5e9',
            'hover': '#f8f9fa'
        }
        
        # Configurar tema
        self.root.configure(bg=self.cores['fundo_principal'])
        
        # Define diretórios
        self.configurar_diretorios()
        
        # Verifica se as bibliotecas necessárias estão instaladas
        if not self.verificar_dependencias():
            return
        
        # Variáveis de estado
        self.texto_fatura = None
        self.formato_selecionado = None
        self.ano_fatura = None
        self.arquivo_pdf_selecionado = None
        self.colunas_modelo = self.obter_colunas_modelo()

        # Configurar estilo moderno
        self.configurar_estilos()
        
        self.criar_tela_inicial()

    def configurar_diretorios(self):
        """Configura os diretórios de trabalho para executável e desenvolvimento."""
        if self.is_executable:
            # Se for executável, usa a pasta onde o .exe está localizado
            self.diretorio_padrao = os.path.dirname(sys.executable)
            # Tenta encontrar o Tesseract incluído no executável
            tesseract_incluido = resource_path(os.path.join('tesseract', 'tesseract.exe'))
            if os.path.exists(tesseract_incluido):
                pytesseract.pytesseract.tesseract_cmd = tesseract_incluido
        else:
            # Se for desenvolvimento, usa o diretório do script
            self.diretorio_padrao = os.path.dirname(os.path.abspath(__file__))

    def configurar_estilos(self):
        """Configura estilos modernos para a interface."""
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Estilo para botões principais
        self.style.configure("Santander.TButton",
                            background=self.cores['santander_red'],
                            foreground="white",
                            font=("Segoe UI", 12, "bold"),
                            borderwidth=0,
                            focuscolor='none',
                            padding=(20, 15))
        
        self.style.configure("Sicoob.TButton",
                            background=self.cores['sicoob_green'],
                            foreground="white",
                            font=("Segoe UI", 12, "bold"),
                            borderwidth=0,
                            focuscolor='none',
                            padding=(20, 15))
        
        self.style.configure("Action.TButton",
                            background=self.cores['processando'],
                            foreground="white",
                            font=("Segoe UI", 10, "bold"),
                            borderwidth=0,
                            focuscolor='none',
                            padding=(15, 10))
        
        self.style.configure("Secondary.TButton",
                            background=self.cores['hover'],
                            foreground=self.cores['texto_principal'],
                            font=("Segoe UI", 10),
                            borderwidth=1,
                            focuscolor='none',
                            padding=(15, 10))
        
        # Estilo para labels
        self.style.configure("Title.TLabel",
                            background=self.cores['fundo_principal'],
                            foreground=self.cores['texto_principal'],
                            font=("Segoe UI", 24, "bold"))
        
        self.style.configure("Subtitle.TLabel",
                            background=self.cores['fundo_principal'],
                            foreground=self.cores['texto_secundario'],
                            font=("Segoe UI", 12))
        
        self.style.configure("Header.TLabel",
                            background=self.cores['fundo_card'],
                            foreground=self.cores['texto_principal'],
                            font=("Segoe UI", 18, "bold"))
        
        self.style.configure("Info.TLabel",
                            background=self.cores['fundo_card'],
                            foreground=self.cores['texto_secundario'],
                            font=("Segoe UI", 11))
        
        # Estilo para frames
        self.style.configure("Card.TFrame",
                            background=self.cores['fundo_card'],
                            borderwidth=1,
                            relief='solid',
                            bordercolor=self.cores['borda'])
        
        # Estilo para progressbar
        self.style.configure("Modern.Horizontal.TProgressbar",
                            background=self.cores['processando'],
                            troughcolor=self.cores['borda'],
                            borderwidth=0,
                            lightcolor=self.cores['processando'],
                            darkcolor=self.cores['processando'])

    def verificar_dependencias(self):
        """Verifica se todas as dependências estão instaladas."""
        bibliotecas_faltando = []
        
        try:
            import fitz
        except ImportError:
            bibliotecas_faltando.append("PyMuPDF")
        
        try:
            from PIL import Image
        except ImportError:
            bibliotecas_faltando.append("Pillow")
        
        try:
            import pytesseract
            # Configura automaticamente o Tesseract
            self.configurar_tesseract()
        except ImportError:
            bibliotecas_faltando.append("pytesseract")
        
        if bibliotecas_faltando:
            if self.is_executable:
                # Se for executável, mostra erro mais específico
                mensagem = f"""❌ Executável incompleto!

As seguintes bibliotecas não estão incluídas:
{chr(10).join(['• ' + lib for lib in bibliotecas_faltando])}

Este executável foi criado sem todas as dependências necessárias.
Por favor, contate o desenvolvedor ou use a versão Python completa."""
            else:
                mensagem = f"""As seguintes bibliotecas não estão instaladas:
{chr(10).join(['• ' + lib for lib in bibliotecas_faltando])}

Para instalar, execute no terminal:
pip install {' '.join(bibliotecas_faltando)}

📋 IMPORTANTE - Instalação do Tesseract OCR:
🪟 WINDOWS: https://github.com/UB-Mannheim/tesseract/wiki"""
            
            messagebox.showerror("Dependências Faltando", mensagem)
            self.root.quit()
            return False
        
        return True

    def obter_colunas_modelo(self):
        """Lê o arquivo modelo e retorna os nomes das colunas da aba 'Banco'."""
        # Primeiro tenta na pasta do executável/script
        caminho_modelo = os.path.join(self.diretorio_padrao, "Automação_Gransoft.xlsx")
        
        # Se não encontrar, tenta no resource_path (para executáveis)
        if not os.path.exists(caminho_modelo) and self.is_executable:
            caminho_modelo = resource_path("Automação_Gransoft.xlsx")
        
        try:
            df_modelo = pd.read_excel(caminho_modelo, sheet_name="Banco")
            print(f"Colunas do modelo: {df_modelo.columns.tolist()}")  # Debug
            return df_modelo.columns.tolist()
        except FileNotFoundError:
            erro_msg = f"""❌ Arquivo modelo não encontrado!

Procurado em:
• {caminho_modelo}

SOLUÇÃO:
1. Certifique-se de que o arquivo 'Automação_Gransoft.xlsx' está na mesma pasta que o executável
2. Ou execute o programa na pasta onde está o arquivo modelo

Para executáveis: O arquivo modelo deve estar junto com o .exe"""
            
            messagebox.showerror("Erro de Arquivo", erro_msg)
            return None
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler as colunas do arquivo modelo:\n{str(e)}")
            return None

    def configurar_tesseract(self):
        """Configura o caminho do Tesseract automaticamente."""
        import platform
        
        if platform.system() == 'Windows':
            # Se for executável, primeiro tenta o Tesseract incluído
            if self.is_executable:
                tesseract_incluido = resource_path(os.path.join('tesseract', 'tesseract.exe'))
                if os.path.exists(tesseract_incluido):
                    pytesseract.pytesseract.tesseract_cmd = tesseract_incluido
                    print(f"Tesseract incluído configurado: {tesseract_incluido}")
                    return True
            
            # Senão, tenta os caminhos padrão de instalação
            caminhos_possiveis = [
                r'C:\Program Files\Tesseract-OCR\tesseract.exe',
                r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
                r'C:\Users\%USERNAME%\AppData\Local\Tesseract-OCR\tesseract.exe',
            ]
            
            for caminho in caminhos_possiveis:
                caminho_expandido = os.path.expandvars(caminho)
                if os.path.exists(caminho_expandido):
                    pytesseract.pytesseract.tesseract_cmd = caminho_expandido
                    print(f"Tesseract sistema configurado: {caminho_expandido}")
                    return True
            
            print("⚠️ Tesseract não encontrado - OCR pode não funcionar")
            return False
        return True

    def extrair_texto_pdf_com_ocr(self, caminho_pdf):
        """Extrai texto de PDF usando OCR."""
        try:
            # Configura o Tesseract automaticamente
            if not self.configurar_tesseract():
                raise Exception("Tesseract OCR não encontrado. Funcionalidade de PDF limitada.")
            
            texto_completo = ""
            
            # Abre o PDF
            documento = fitz.open(caminho_pdf)
            
            for pagina_num in range(len(documento)):
                pagina = documento.load_page(pagina_num)
                
                # Converte página para imagem com maior resolução
                matriz = fitz.Matrix(3, 3)  # Aumenta ainda mais a resolução para melhor OCR
                pix = pagina.get_pixmap(matrix=matriz)
                img_data = pix.tobytes("png")
                
                # Converte para PIL Image
                imagem = Image.open(io.BytesIO(img_data))
                
                # Tenta OCR primeiro em português, depois em inglês se falhar
                try:
                    config = '--oem 3 --psm 6'
                    texto_pagina = pytesseract.image_to_string(imagem, lang='por', config=config)
                except Exception as e:
                    print(f"Erro com idioma português, tentando inglês: {e}")
                    try:
                        # Fallback para inglês se português não funcionar
                        texto_pagina = pytesseract.image_to_string(imagem, lang='eng', config=config)
                    except Exception as e2:
                        print(f"Erro com inglês também, tentando sem especificar idioma: {e2}")
                        # Último recurso: sem especificar idioma
                        texto_pagina = pytesseract.image_to_string(imagem, config=config)
                
                texto_completo += texto_pagina + "\n"
            
            documento.close()
            return texto_completo
            
        except Exception as e:
            raise Exception(f"Erro ao extrair texto do PDF: {str(e)}")

    def selecionar_arquivo_pdf(self):
        """Permite ao usuário selecionar um arquivo PDF."""
        arquivo = filedialog.askopenfilename(
            title="Selecionar Fatura PDF",
            filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")],
            initialdir=self.diretorio_padrao
        )
        
        if arquivo:
            self.arquivo_pdf_selecionado = arquivo
            self.processar_pdf_selecionado()
    
    def processar_pdf_selecionado(self):
        """Processa o PDF selecionado com OCR."""
        if not self.arquivo_pdf_selecionado:
            return
        
        self.iniciar_animacao_processamento()
        self.status_var.set("🔍 Extraindo texto do PDF...")
        
        try:
            # Extrai texto usando OCR
            texto_extraido = self.extrair_texto_pdf_com_ocr(self.arquivo_pdf_selecionado)
            
            # Insere o texto na área de texto
            self.texto_fatura.delete(1.0, tk.END)
            self.texto_fatura.insert(1.0, texto_extraido)
            
            self.finalizar_animacao_processamento(True, "✅ Texto extraído com sucesso! Revise e clique em 'Gerar Planilha Excel'.")
            
        except Exception as e:
            self.finalizar_animacao_processamento(False, f"❌ Erro ao processar PDF: {str(e)}")
            messagebox.showerror("Erro", f"Não foi possível processar o PDF:\n{str(e)}\n\nTente colar o texto manualmente.")

    def criar_frame_card(self, parent, padding=20):
        """Cria um frame com aparência de card moderno."""
        card = ttk.Frame(parent, style="Card.TFrame", padding=padding)
        return card

    def limpar_tela(self):
        """Remove todos os widgets da tela atual."""
        for widget in self.root.winfo_children():
            widget.destroy()
    
    def criar_tela_inicial(self):
        """Cria a tela inicial com as opções de fatura."""
        self.limpar_tela()
        if not self.colunas_modelo:
            return

        # Container principal
        main_container = tk.Frame(self.root, bg=self.cores['fundo_principal'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=40, pady=40)

        # Header
        header_frame = tk.Frame(main_container, bg=self.cores['fundo_principal'])
        header_frame.pack(fill=tk.X, pady=(0, 40))

        titulo = ttk.Label(header_frame, text="💳 Conversor de Faturas", style="Title.TLabel")
        titulo.pack()

        subtitulo = ttk.Label(header_frame, text="Transforme suas faturas de cartão em planilhas Excel organizadas", style="Subtitle.TLabel")
        subtitulo.pack(pady=(10, 0))
        
        # Indicador de versão para executável
        if self.is_executable:
            versao_label = ttk.Label(header_frame, text="📦 Versão Executável", style="Info.TLabel")
            versao_label.pack(pady=(5, 0))

        # Card principal
        card_principal = self.criar_frame_card(main_container, padding=40)
        card_principal.pack(fill=tk.BOTH, expand=True)

        # Título do card
        titulo_selecao = ttk.Label(card_principal, text="Selecione seu Banco", style="Header.TLabel")
        titulo_selecao.pack(pady=(0, 10))

        info_selecao = ttk.Label(card_principal, text="Escolha o banco emissor da sua fatura de cartão de crédito", style="Info.TLabel")
        info_selecao.pack(pady=(0, 40))

        # Frame para os botões dos bancos
        bancos_frame = tk.Frame(card_principal, bg=self.cores['fundo_card'])
        bancos_frame.pack(pady=20)

        # Botão Santander
        santander_frame = tk.Frame(bancos_frame, bg=self.cores['fundo_card'])
        santander_frame.grid(row=0, column=0, padx=30, pady=20)

        banco_santander_icon = ttk.Label(santander_frame, text="🏦", font=("Segoe UI", 40), background=self.cores['fundo_card'])
        banco_santander_icon.pack()

        btn_santander = ttk.Button(santander_frame, text="Santander", style="Santander.TButton", command=lambda: self.iniciar_processamento('santander'))
        btn_santander.pack(pady=(10, 5))

        info_santander = ttk.Label(santander_frame, text="Suporte a PDF\ncom OCR automático", style="Info.TLabel")
        info_santander.pack()

        # Botão Sicoob
        sicoob_frame = tk.Frame(bancos_frame, bg=self.cores['fundo_card'])
        sicoob_frame.grid(row=0, column=1, padx=30, pady=20)

        banco_sicoob_icon = ttk.Label(sicoob_frame, text="🏛️", font=("Segoe UI", 40), background=self.cores['fundo_card'])
        banco_sicoob_icon.pack()

        btn_sicoob = ttk.Button(sicoob_frame, text="Sicoob", style="Sicoob.TButton", command=lambda: self.iniciar_processamento('sicoob'))
        btn_sicoob.pack(pady=(10, 5))

        info_sicoob = ttk.Label(sicoob_frame, text="Cole o texto\nda fatura", style="Info.TLabel")
        info_sicoob.pack()

        # Footer com informações
        footer_frame = tk.Frame(card_principal, bg=self.cores['fundo_card'])
        footer_frame.pack(fill=tk.X, pady=(40, 0))

        features_text = "✨ Recursos: Preserva arquivo modelo • Cria planilhas organizadas • Interface amigável"
        features_label = ttk.Label(footer_frame, text=features_text, style="Info.TLabel")
        features_label.pack()

    def iniciar_processamento(self, formato):
        """
        Define o formato selecionado e avança para a tela de processamento.
        Se o formato for 'sicoob', pede o ano antes.
        """
        self.formato_selecionado = formato
        
        if self.formato_selecionado == 'sicoob':
            self.pedir_ano_sicoob()
        else:
            self.criar_tela_processamento()

    def pedir_ano_sicoob(self):
        """Pede o ano da fatura Sicoob via caixa de diálogo."""
        self.ano_fatura = simpledialog.askinteger("📅 Ano da Fatura", "Digite o ano da fatura (ex: 2024):",
                                                  parent=self.root,
                                                  minvalue=2000,
                                                  maxvalue=datetime.now().year + 1)
        if self.ano_fatura is not None:
            self.criar_tela_processamento()
        else:
            messagebox.showwarning("⚠️ Aviso", "Ano da fatura é obrigatório para o formato Sicoob.")

    def criar_tela_processamento(self):
        """Cria a tela com a área de texto para colar a fatura."""
        self.limpar_tela()
        
        # Container principal
        main_container = tk.Frame(self.root, bg=self.cores['fundo_principal'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)

        # Header
        header_frame = tk.Frame(main_container, bg=self.cores['fundo_principal'])
        header_frame.pack(fill=tk.X, pady=(0, 20))

        # Título com ícone do banco
        if self.formato_selecionado == 'santander':
            titulo_texto = f"🏦 Processar Fatura Santander"
            cor_titulo = self.cores['santander_red']
        else:
            titulo_texto = f"🏛️ Processar Fatura Sicoob - Ano {self.ano_fatura}"
            cor_titulo = self.cores['sicoob_green']

        titulo = tk.Label(header_frame, text=titulo_texto, font=("Segoe UI", 20, "bold"), 
                         fg=cor_titulo, bg=self.cores['fundo_principal'])
        titulo.pack()

        # Card principal
        card_principal = self.criar_frame_card(main_container, padding=30)
        card_principal.pack(fill=tk.BOTH, expand=True)

        # Instruções específicas para cada banco
        if self.formato_selecionado == 'santander':
            # Seção de opções para Santander
            opcoes_frame = tk.Frame(card_principal, bg=self.cores['fundo_card'])
            opcoes_frame.pack(fill=tk.X, pady=(0, 20))

            instrucoes_label = ttk.Label(opcoes_frame, text="📋 Escolha como fornecer sua fatura:", style="Info.TLabel")
            instrucoes_label.pack(anchor="w", pady=(0, 15))

            # Botões de opção em linha
            botoes_opcao_frame = tk.Frame(opcoes_frame, bg=self.cores['fundo_card'])
            botoes_opcao_frame.pack(fill=tk.X)

            upload_btn = ttk.Button(botoes_opcao_frame, text="📁 Fazer Upload de PDF", style="Action.TButton", command=self.selecionar_arquivo_pdf)
            upload_btn.pack(side=tk.LEFT, padx=(0, 15))

            ou_label = tk.Label(botoes_opcao_frame, text="OU", font=("Segoe UI", 10, "bold"), 
                              fg=self.cores['texto_secundario'], bg=self.cores['fundo_card'])
            ou_label.pack(side=tk.LEFT, padx=(0, 15))

            # Separador visual
            separador = tk.Frame(card_principal, height=2, bg=self.cores['borda'])
            separador.pack(fill=tk.X, pady=20)

            instrucoes2 = ttk.Label(card_principal, text="✏️ Ou cole o texto da fatura abaixo:", style="Info.TLabel")
            instrucoes2.pack(anchor="w", pady=(0, 10))
        else:
            instrucoes = ttk.Label(card_principal, text="✏️ Cole o texto da fatura do seu cartão de crédito:", style="Info.TLabel")
            instrucoes.pack(anchor="w", pady=(0, 10))
        
        # Área de texto com estilo moderno
        texto_frame = tk.Frame(card_principal, bg=self.cores['fundo_card'])
        texto_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.texto_fatura = scrolledtext.ScrolledText(
            texto_frame, 
            height=12, 
            width=80, 
            font=("Consolas", 11),
            bg='white',
            fg=self.cores['texto_principal'],
            insertbackground=self.cores['processando'],
            selectbackground=self.cores['processando'],
            selectforeground='white',
            borderwidth=2,
            relief='solid',
            highlightcolor=self.cores['processando']
        )
        self.texto_fatura.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Frame para botões de ação
        botoes_frame = tk.Frame(card_principal, bg=self.cores['fundo_card'])
        botoes_frame.pack(fill=tk.X, pady=(20, 0))
        
        # Botões com ícones
        processar_btn = ttk.Button(botoes_frame, text="📊 Gerar Planilha Excel", style="Action.TButton", command=self.processar_fatura)
        processar_btn.pack(side=tk.RIGHT, padx=(10, 0))
        
        limpar_btn = ttk.Button(botoes_frame, text="🗑️ Limpar", style="Secondary.TButton", command=self.limpar_texto)
        limpar_btn.pack(side=tk.RIGHT)
        
        voltar_btn = ttk.Button(botoes_frame, text="⬅️ Voltar", style="Secondary.TButton", command=self.criar_tela_inicial)
        voltar_btn.pack(side=tk.LEFT)

        # Área de status moderna
        status_frame = tk.Frame(main_container, bg=self.cores['fundo_principal'])
        status_frame.pack(fill=tk.X, pady=(20, 0))

        self.status_var = tk.StringVar()
        self.status_var.set("✨ Pronto para processar sua fatura!")
        self.status_label = tk.Label(status_frame, textvariable=self.status_var, 
                                   font=("Segoe UI", 11), fg=self.cores['texto_secundario'], 
                                   bg=self.cores['fundo_principal'])
        self.status_label.pack()
        
        self.progresso = ttk.Progressbar(status_frame, style="Modern.Horizontal.TProgressbar", 
                                       orient=tk.HORIZONTAL, length=400, mode='indeterminate')

    def limpar_texto(self):
        """Limpa o texto da área de fatura"""
        self.texto_fatura.delete(1.0, tk.END)
        self.status_var.set("✨ Área de texto limpa. Pronto para nova fatura!")
        self.root.update()

    def iniciar_animacao_processamento(self):
        """Configura a interface para indicar processamento"""
        self.status_var.set("⚙️ Processando...")
        self.status_label.configure(fg=self.cores['processando'])
        self.progresso.pack(pady=(10, 0))
        self.progresso.start(10)
        self.root.update()

    def finalizar_animacao_processamento(self, sucesso=True, mensagem=""):
        """Retorna a interface ao estado normal após processamento"""
        self.progresso.stop()
        self.progresso.pack_forget()
        
        if sucesso:
            self.status_var.set(mensagem if mensagem else "✅ Planilha gerada com sucesso!")
            self.status_label.configure(fg=self.cores['sucesso'])
        else:
            self.status_var.set(mensagem if mensagem else "❌ Erro ao processar a fatura.")
            self.status_label.configure(fg=self.cores['erro'])
        
        self.root.update()

    def identificar_coluna_data(self):
        """Identifica qual coluna representa a data no arquivo modelo."""
        possiveis_nomes = ['Data', 'Data Vencimento', 'Data_Vencimento', 'Dt_Vencimento', 'Date']
        
        for nome in possiveis_nomes:
            if nome in self.colunas_modelo:
                return nome
        
        # Se não encontrar, retorna a primeira coluna (assumindo que é a data)
        return self.colunas_modelo[0] if self.colunas_modelo else 'Data'

    def processar_fatura(self):
        """Processa o texto da fatura de acordo com o formato selecionado."""
        texto = self.texto_fatura.get(1.0, tk.END).strip()
        
        if not texto:
            messagebox.showwarning("⚠️ Atenção", "Por favor, forneça o texto da fatura primeiro.")
            return

        self.iniciar_animacao_processamento()
        
        try:
            if self.formato_selecionado == 'santander':
                dados = self.processar_formato_santander(texto)
            elif self.formato_selecionado == 'sicoob':
                dados = self.processar_formato_sicoob(texto, self.ano_fatura)
            
            if not dados:
                self.finalizar_animacao_processamento(False, "❌ Nenhuma transação de gasto encontrada.")
                return
            
            # Identifica o nome correto da coluna de data
            nome_coluna_data = self.identificar_coluna_data()
            print(f"Coluna de data identificada: {nome_coluna_data}")  # Debug
            
            # Cria o DataFrame com as colunas-modelo
            df = pd.DataFrame(columns=self.colunas_modelo)

            # Insere os dados processados no DataFrame usando nomes de colunas
            for item in dados:
                nova_linha = {}
                
                # Preenche todas as colunas com valores vazios primeiro
                for coluna in self.colunas_modelo:
                    nova_linha[coluna] = ''
                
                # Mapeia os dados para as colunas corretas
                nova_linha[nome_coluna_data] = item.get('Data', '')
                
                # Procura por colunas que podem conter descrição
                colunas_descricao = [col for col in self.colunas_modelo if any(palavra in col.lower() for palavra in ['descricao', 'description', 'desc'])]
                if colunas_descricao:
                    nova_linha[colunas_descricao[0]] = item.get('Descricao', '')
                
                # Procura por colunas que podem conter valor
                colunas_valor = [col for col in self.colunas_modelo if any(palavra in col.lower() for palavra in ['valor', 'value', 'amount'])]
                if colunas_valor:
                    nova_linha[colunas_valor[0]] = item.get('Valor', '')
                
                # Procura por colunas que podem conter observação
                colunas_observacao = [col for col in self.colunas_modelo if any(palavra in col.lower() for palavra in ['observacao', 'obs', 'observation'])]
                if colunas_observacao:
                    nova_linha[colunas_observacao[0]] = item.get('Observacao', '')
                
                df.loc[len(df)] = nova_linha
            
            # --- LÓGICA DE ORDENAÇÃO POR DATA CORRIGIDA ---
            if nome_coluna_data in df.columns and not df[nome_coluna_data].empty:
                try:
                    # Converte a coluna de data para o formato datetime
                    df[nome_coluna_data] = pd.to_datetime(df[nome_coluna_data], format='%d/%m/%Y', errors='coerce')
                    
                    # Ordena o DataFrame pela coluna de data e remove linhas com datas inválidas
                    df.sort_values(by=nome_coluna_data, inplace=True)
                    df.dropna(subset=[nome_coluna_data], inplace=True)
                    
                    # Converte a coluna de data de volta para o formato de string
                    df[nome_coluna_data] = df[nome_coluna_data].dt.strftime('%d/%m/%Y')
                except Exception as e:
                    print(f"Erro ao processar datas: {e}")
                    messagebox.showwarning("⚠️ Aviso", f"Não foi possível ordenar por data: {e}")
            # ----------------------------------------
            
            caminho_arquivo_gerado = self.criar_nova_planilha_com_estrutura_modelo(df)
            
            self.finalizar_animacao_processamento(True, f"🎉 {len(dados)} transações processadas com sucesso!")
            
            # Diálogo de sucesso mais elegante
            result = messagebox.askyesno("🎉 Sucesso!", 
                                       f"Nova planilha criada com {len(dados)} transações!\n\n"
                                       f"📁 Arquivo: {os.path.basename(caminho_arquivo_gerado)}\n\n"
                                       "Deseja abrir o arquivo agora?",
                                       icon='question')
            if result:
                self.abrir_arquivo(caminho_arquivo_gerado)
                
        except Exception as e:
            self.finalizar_animacao_processamento(False, f"❌ Erro: {str(e)}")
            messagebox.showerror("❌ Erro", f"Ocorreu um erro ao processar a fatura:\n{str(e)}")
            print(f"Erro detalhado: {e}")  # Debug

    def processar_formato_santander(self, texto):
        """Processa a fatura do Santander."""
        dados = []
        
        # Padrão atualizado para o formato Santander atual
        # Data + Descrição + Valor R$ + US$ + Cotação
        padrao = r'^(\d{2}\/\d{2}\/\d{4})\s+(.+?)\s+(\d{1,3}(?:\.\d{3})*,\d{2})\s+\d{1,3}(?:\.\d{3})*,\d{2}\s+\d{1,3}(?:[.,]\d{3})*'
        
        # Padrão alternativo para valores negativos (como estorno)
        padrao_negativo = r'^(\d{2}\/\d{2}\/\d{4})\s+(.+?)\s+(-\d{1,3}(?:\.\d{3})*,\d{2})'
        
        for linha in texto.split('\n'):
            linha = linha.strip()
            if not linha:
                continue
                
            # Tenta padrão principal
            match = re.search(padrao, linha)
            if match:
                data_completa, descricao, valor_str = match.groups()
                valor = float(valor_str.replace('.', '').replace(',', '.'))
                
                if valor > 0:  # Apenas valores positivos (gastos)
                    dados.append({
                        'Data': data_completa,
                        'Descricao': descricao.strip(),
                        'Valor': valor,
                        'Observacao': descricao.strip()
                    })
            else:
                # Tenta padrão para valores negativos (estornos)
                match_neg = re.search(padrao_negativo, linha)
                if match_neg:
                    data_completa, descricao, valor_str = match_neg.groups()
                    # Valores negativos são ignorados (estornos), mas você pode incluí-los se quiser
                    continue
        
        return dados

    def processar_formato_sicoob(self, texto, ano):
        """Processa a fatura do Sicoob."""
        dados = []
        
        padrao = r'^(\d{2}\/\d{2})\s+(.+?)\s+([-,]?\d{1,3}(?:\.\d{3})*?[.,]\d{2})$'
        
        linhas = [l.strip() for l in texto.split('\n') if l.strip()]
        
        for linha in linhas:
            if any(palavra in linha for palavra in ['SALDO ANTERIOR', 'TOTAL', 'GASTOS DE']):
                continue
            
            match = re.search(padrao, linha)
            if match:
                data_str, descricao, valor_str = match.groups()
                
                valor_str = valor_str.replace('.', '').replace(',', '.')
                valor = float(valor_str)
                
                if valor > 0:
                    data_completa = f"{data_str}/{ano}"
                    dados.append({
                        'Data': data_completa,
                        'Descricao': descricao.strip(),
                        'Valor': valor,
                        'Observacao': descricao.strip()
                    })
        return dados

    def criar_nova_planilha_com_estrutura_modelo(self, df):
        """Cria uma nova planilha copiando a estrutura do modelo e atualizando apenas a aba 'Banco'."""
        try:
            import shutil
            
            # Caminho do arquivo modelo
            caminho_modelo = os.path.join(self.diretorio_padrao, "Automação_Gransoft.xlsx")
            if not os.path.exists(caminho_modelo) and self.is_executable:
                caminho_modelo = resource_path("Automação_Gransoft.xlsx")
            
            nome_novo_arquivo = f"fatura_processada_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            caminho_novo_arquivo = os.path.join(self.diretorio_padrao, nome_novo_arquivo)
            
            # Verifica se o arquivo modelo existe
            if not os.path.exists(caminho_modelo):
                raise FileNotFoundError(f"Arquivo modelo '{os.path.basename(caminho_modelo)}' não encontrado")
            
            # Copia o arquivo modelo para criar a nova planilha
            shutil.copy2(caminho_modelo, caminho_novo_arquivo)
            print(f"Planilha modelo copiada para: {nome_novo_arquivo}")
            
            # Abre a nova planilha e substitui apenas a aba 'Banco'
            with pd.ExcelWriter(caminho_novo_arquivo, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                # Substitui apenas a aba 'Banco' com os novos dados
                df.to_excel(writer, sheet_name='Banco', index=False)
                
            print(f"Nova planilha criada com sucesso: {nome_novo_arquivo}")
            return caminho_novo_arquivo
            
        except Exception as e:
            raise Exception(f"Erro ao criar nova planilha: {str(e)}")

    def abrir_arquivo(self, caminho):
        """Abre o arquivo no aplicativo padrão."""
        try:
            if sys.platform.startswith('win'):
                os.startfile(caminho)
            elif sys.platform.startswith('darwin'):  # macOS
                os.system(f'open "{caminho}"')
            else:  # Linux
                os.system(f'xdg-open "{caminho}"')
        except Exception as e:
            messagebox.showwarning("⚠️ Aviso", f"Não foi possível abrir o arquivo automaticamente.\nArquivo salvo em: {caminho}")

def main():
    root = tk.Tk()
    app = ConversorFaturas(root)
    root.mainloop()

if __name__ == "__main__":
    main()