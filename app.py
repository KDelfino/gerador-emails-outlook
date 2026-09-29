import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import email_engine

class EmailSenderApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Gerador e Enviador de E-mails - Outlook Automation")
        self.geometry("1040x920")
        self.minsize(940, 740)

        self.COLOR_BG = "#12141A"
        self.COLOR_CARD = "#1A1D26"
        self.COLOR_CARD_BORDER = "#2E3444"
        self.COLOR_INPUT_BG = "#13151D"
        self.COLOR_INPUT_BORDER = "#374151"
        self.COLOR_TEXT_MAIN = "#F1F5F9"
        self.COLOR_TEXT_MUTED = "#94A3B8"
        self.COLOR_ACCENT = "#2563EB"
        self.COLOR_ACCENT_HOVER = "#1D4ED8"
        self.COLOR_TAG_BG = "#222736"
        self.COLOR_TAG_FG = "#93C5FD"
        self.COLOR_HEADER_BG = "#151722"

        self.caminho_planilha = tk.StringVar()
        self.aba_selecionada = tk.StringVar()
        self.linha_cabecalho_var = tk.StringVar()
        self.coluna_email = tk.StringVar()
        self.coluna_condicao = tk.StringVar(value="[Nenhum - Modelo unico para todos]")
        self.modo_condicao = tk.StringVar(value="unico")
        self.grupo_em_edicao = tk.StringVar()
        
        self.assunto_var = tk.StringVar()
        self.cc_var = tk.StringVar(value="")
        self.bcc_var = tk.StringVar(value="")
        self.tipo_acao = tk.StringVar(value="rascunho")
        
        self.colunas_disponiveis = []
        self.registros = []
        self.lista_anexos = []
        self.sugestoes_cabecalho = []
        
        self.grupos_disponiveis = []
        self.grupos_selecionados_vars = {}
        self.modelos_por_grupo = {}
        self.grupo_anterior_edicao = None

        self.usar_anexo_dinamico = tk.BooleanVar(value=False)
        self.pasta_anexos_dinamicos = tk.StringVar(value="")
        self.coluna_busca_anexo = tk.StringVar(value="")
        self.padrao_busca_anexo = tk.StringVar(value="#nome#")
        self.pular_se_sem_anexo_dinamico = tk.BooleanVar(value=True)

        self.ultimo_foco = None
        self.cancel_event = threading.Event()
        self.em_execucao = False

        self._configurar_estilos()
        self._construir_interface()

        self.ultimo_foco = self.entry_assunto

    def _configurar_estilos(self):
        self.style = ttk.Style(self)
        try:
            self.style.theme_use('clam')
        except Exception:
            pass

        self.configure(bg=self.COLOR_BG)

        self.style.configure("TFrame", background=self.COLOR_BG)
        self.style.configure("Card.TFrame", background=self.COLOR_CARD)
        
        self.style.configure(
            "TLabelframe", 
            background=self.COLOR_CARD, 
            bordercolor=self.COLOR_CARD_BORDER,
            darkcolor=self.COLOR_CARD_BORDER,
            lightcolor=self.COLOR_CARD_BORDER,
            padding=10
        )
        self.style.configure(
            "TLabelframe.Label", 
            font=("Segoe UI", 10, "bold"), 
            foreground=self.COLOR_TEXT_MAIN, 
            background=self.COLOR_CARD
        )
        
        self.style.configure("TLabel", background=self.COLOR_CARD, font=("Segoe UI", 9), foreground=self.COLOR_TEXT_MAIN)
        self.style.configure("Muted.TLabel", background=self.COLOR_CARD, font=("Segoe UI", 8), foreground=self.COLOR_TEXT_MUTED)

        self.style.configure(
            "TButton", 
            background="#262B3A", 
            foreground=self.COLOR_TEXT_MAIN, 
            bordercolor=self.COLOR_CARD_BORDER,
            darkcolor="#262B3A",
            lightcolor="#262B3A",
            font=("Segoe UI", 9),
            padding=5
        )
        self.style.map("TButton", background=[("active", "#32384D"), ("disabled", "#181A22")],
                                  foreground=[("disabled", "#52596D")])

        self.style.configure(
            "Primary.TButton", 
            background=self.COLOR_ACCENT, 
            foreground="#FFFFFF", 
            bordercolor=self.COLOR_ACCENT,
            darkcolor=self.COLOR_ACCENT,
            lightcolor=self.COLOR_ACCENT,
            font=("Segoe UI", 9, "bold"),
            padding=5
        )
        self.style.map("Primary.TButton", background=[("active", self.COLOR_ACCENT_HOVER)])

        self.style.configure(
            "Action.TButton", 
            background="#1E40AF", 
            foreground="#FFFFFF", 
            bordercolor="#1E40AF",
            darkcolor="#1E40AF",
            lightcolor="#1E40AF",
            font=("Segoe UI", 10, "bold"), 
            padding=7
        )
        self.style.map("Action.TButton", background=[("active", "#1D4ED8"), ("disabled", "#1E293B")],
                                         foreground=[("disabled", "#64748B")])

        self.style.configure(
            "Tag.TButton", 
            background=self.COLOR_TAG_BG, 
            foreground=self.COLOR_TAG_FG, 
            bordercolor="#2E374D",
            darkcolor=self.COLOR_TAG_BG,
            lightcolor=self.COLOR_TAG_BG,
            font=("Segoe UI", 8), 
            padding=3
        )
        self.style.map("Tag.TButton", background=[("active", "#2D3448")])

        self.style.configure(
            "Small.TButton", 
            background="#262B3A", 
            foreground=self.COLOR_TEXT_MAIN, 
            bordercolor=self.COLOR_CARD_BORDER,
            font=("Segoe UI", 8), 
            padding=2
        )
        self.style.map("Small.TButton", background=[("active", "#32384D")])

        self.style.configure(
            "TEntry", 
            fieldbackground=self.COLOR_INPUT_BG, 
            foreground=self.COLOR_TEXT_MAIN, 
            bordercolor=self.COLOR_INPUT_BORDER,
            insertcolor=self.COLOR_TEXT_MAIN,
            padding=4
        )

        self.style.configure(
            "TCombobox", 
            fieldbackground=self.COLOR_INPUT_BG, 
            background="#262B3A", 
            foreground=self.COLOR_TEXT_MAIN, 
            arrowcolor=self.COLOR_TEXT_MAIN,
            bordercolor=self.COLOR_INPUT_BORDER,
            darkcolor="#262B3A",
            lightcolor="#262B3A",
            padding=3
        )
        self.style.map("TCombobox", fieldbackground=[("readonly", self.COLOR_INPUT_BG)],
                                    selectbackground=[("readonly", self.COLOR_ACCENT)],
                                    selectforeground=[("readonly", "#FFFFFF")])

        self.style.configure(
            "TRadiobutton", 
            background=self.COLOR_CARD, 
            foreground=self.COLOR_TEXT_MAIN,
            font=("Segoe UI", 9),
            indicatorcolor=self.COLOR_INPUT_BG,
            bordercolor=self.COLOR_CARD_BORDER
        )
        self.style.map("TRadiobutton", indicatorcolor=[("selected", self.COLOR_ACCENT)])

        self.style.configure(
            "TCheckbutton", 
            background=self.COLOR_CARD, 
            foreground=self.COLOR_TEXT_MAIN,
            font=("Segoe UI", 9),
            indicatorcolor=self.COLOR_INPUT_BG,
            bordercolor=self.COLOR_CARD_BORDER
        )
        self.style.map("TCheckbutton", indicatorcolor=[("selected", self.COLOR_ACCENT)])

        self.style.configure(
            "Horizontal.TProgressbar", 
            background=self.COLOR_ACCENT, 
            troughcolor=self.COLOR_INPUT_BG,
            bordercolor=self.COLOR_CARD_BORDER
        )

        self.style.configure("TSeparator", background=self.COLOR_CARD_BORDER)

    def _construir_interface(self):
        header_frame = tk.Frame(self, bg=self.COLOR_HEADER_BG, height=65, padx=20, pady=12, 
                                highlightthickness=1, highlightbackground=self.COLOR_CARD_BORDER)
        header_frame.pack(fill="x", side="top")

        lbl_title = tk.Label(header_frame, text="Gerador e Enviador de E-mails Outlook", 
                             font=("Segoe UI", 13, "bold"), bg=self.COLOR_HEADER_BG, fg="#FFFFFF")
        lbl_title.pack(anchor="w")

        lbl_subtitle = tk.Label(header_frame, text="Carregue qualquer planilha Excel/CSV, crie modelos gerais ou por condicao/grupo e envie ou salve em rascunhos no Outlook.", 
                                font=("Segoe UI", 9), bg=self.COLOR_HEADER_BG, fg=self.COLOR_TEXT_MUTED)
        lbl_subtitle.pack(anchor="w")

        main_canvas = tk.Canvas(self, bg=self.COLOR_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=main_canvas.yview)
        
        self.scrollable_frame = tk.Frame(main_canvas, bg=self.COLOR_BG, padx=15, pady=10)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )

        main_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)

        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _on_mousewheel(event):
            main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        main_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        self._criar_secao_dados(self.scrollable_frame)
        self._criar_secao_tags(self.scrollable_frame)
        self._criar_secao_modelo(self.scrollable_frame)
        self._criar_secao_anexos(self.scrollable_frame)
        self._criar_secao_acoes(self.scrollable_frame)

    def _criar_secao_dados(self, parent):
        frame = ttk.LabelFrame(parent, text=" 1. Fonte de Dados (Planilha Excel / CSV) ")
        frame.pack(fill="x", pady=6)

        row1 = tk.Frame(frame, bg=self.COLOR_CARD)
        row1.pack(fill="x", pady=4)

        ttk.Label(row1, text="Arquivo:").pack(side="left", padx=5)
        self.entry_arquivo = ttk.Entry(row1, textvariable=self.caminho_planilha, width=50)
        self.entry_arquivo.pack(side="left", fill="x", expand=True, padx=5)

        btn_buscar = ttk.Button(row1, text="Selecionar Planilha...", command=self._selecionar_arquivo, style="Primary.TButton")
        btn_buscar.pack(side="left", padx=3)

        btn_trocar = ttk.Button(row1, text="Trocar / Limpar", command=self._limpar_dados_planilha)
        btn_trocar.pack(side="left", padx=3)

        row2 = tk.Frame(frame, bg=self.COLOR_CARD)
        row2.pack(fill="x", pady=4)

        ttk.Label(row2, text="Aba:").pack(side="left", padx=5)
        self.combo_abas = ttk.Combobox(row2, textvariable=self.aba_selecionada, state="readonly", width=18)
        self.combo_abas.pack(side="left", padx=5)
        self.combo_abas.bind("<<ComboboxSelected>>", lambda e: self._ao_mudar_aba())

        ttk.Label(row2, text="Cabecalho:").pack(side="left", padx=(10, 5))
        self.combo_cabecalho = ttk.Combobox(row2, textvariable=self.linha_cabecalho_var, state="readonly", width=34)
        self.combo_cabecalho.pack(side="left", padx=5)
        self.combo_cabecalho.bind("<<ComboboxSelected>>", lambda e: self._ao_mudar_cabecalho())

        ttk.Label(row2, text="Coluna E-mail:").pack(side="left", padx=(10, 5))
        self.combo_email = ttk.Combobox(row2, textvariable=self.coluna_email, state="readonly", width=20)
        self.combo_email.pack(side="left", padx=5)
        self.combo_email.bind("<<ComboboxSelected>>", lambda e: self._atualizar_contagem_destinatarios())

        row3 = tk.Frame(frame, bg=self.COLOR_CARD)
        row3.pack(fill="x", pady=(6, 4))

        ttk.Label(row3, text="Condicao / Grupo por Coluna:").pack(side="left", padx=5)
        self.combo_condicao = ttk.Combobox(row3, textvariable=self.coluna_condicao, state="readonly", width=32)
        self.combo_condicao.pack(side="left", padx=5)
        self.combo_condicao.bind("<<ComboboxSelected>>", lambda e: self._ao_mudar_coluna_condicao())

        self.frame_painel_grupos = tk.Frame(frame, bg="#161821", bd=1, relief="solid", 
                                            highlightthickness=1, highlightbackground=self.COLOR_CARD_BORDER, padx=10, pady=8)

        self.frame_grupos_header = tk.Frame(self.frame_painel_grupos, bg="#161821")
        self.frame_grupos_header.pack(fill="x", pady=(0, 4))

        tk.Label(self.frame_grupos_header, text="Grupos / Valores encontrados nesta coluna:", 
                 font=("Segoe UI", 9, "bold"), bg="#161821", fg=self.COLOR_TEXT_MAIN).pack(side="left")

        btn_marcar_todos = ttk.Button(self.frame_grupos_header, text="Marcar Todos", 
                                      command=lambda: self._marcar_desmarcar_grupos(True), style="Small.TButton")
        btn_marcar_todos.pack(side="right", padx=3)

        btn_desmarcar_todos = ttk.Button(self.frame_grupos_header, text="Desmarcar Todos", 
                                         command=lambda: self._marcar_desmarcar_grupos(False), style="Small.TButton")
        btn_desmarcar_todos.pack(side="right", padx=3)

        self.frame_grupos_checkboxes = tk.Frame(self.frame_painel_grupos, bg="#161821")
        self.frame_grupos_checkboxes.pack(fill="x", pady=4)

        self.frame_modo_modelo = tk.Frame(self.frame_painel_grupos, bg="#161821")
        self.frame_modo_modelo.pack(fill="x", pady=(6, 2))

        rb_mod_unico = tk.Radiobutton(
            self.frame_modo_modelo, 
            text="Usar o mesmo modelo de e-mail para todos os grupos selecionados", 
            value="unico", 
            variable=self.modo_condicao,
            command=self._ao_trocar_modo_condicao,
            bg="#161821",
            fg=self.COLOR_TEXT_MAIN,
            selectcolor=self.COLOR_INPUT_BG,
            activebackground="#161821",
            activeforeground="#FFFFFF",
            font=("Segoe UI", 9)
        )
        rb_mod_unico.pack(anchor="w")

        rb_mod_grupo = tk.Radiobutton(
            self.frame_modo_modelo, 
            text="Configurar modelo e anexos diferentes para cada grupo individualmente", 
            value="por_grupo", 
            variable=self.modo_condicao,
            command=self._ao_trocar_modo_condicao,
            bg="#161821",
            fg=self.COLOR_TEXT_MAIN,
            selectcolor=self.COLOR_INPUT_BG,
            activebackground="#161821",
            activeforeground="#FFFFFF",
            font=("Segoe UI", 9)
        )
        rb_mod_grupo.pack(anchor="w", pady=(2, 0))

        self.frame_seletor_grupo_edicao = tk.Frame(self.frame_painel_grupos, bg="#1E2333", bd=1, relief="solid", 
                                                   highlightthickness=1, highlightbackground=self.COLOR_ACCENT, padx=8, pady=6)

        tk.Label(self.frame_seletor_grupo_edicao, text="Editando Modelo e Anexos do Grupo:", 
                 font=("Segoe UI", 9, "bold"), bg="#1E2333", fg="#93C5FD").pack(side="left", padx=5)

        self.combo_grupo_edicao = ttk.Combobox(self.frame_seletor_grupo_edicao, textvariable=self.grupo_em_edicao, state="readonly", width=22)
        self.combo_grupo_edicao.pack(side="left", padx=5)
        self.combo_grupo_edicao.bind("<<ComboboxSelected>>", lambda e: self._ao_mudar_grupo_em_edicao())

        btn_copiar_modelo = ttk.Button(self.frame_seletor_grupo_edicao, text="Copiar este modelo para todos os grupos", 
                                       command=self._copiar_modelo_atual_para_todos, style="Small.TButton")
        btn_copiar_modelo.pack(side="right", padx=5)

        self.frame_resumo_dados = tk.Frame(frame, bg="#131F1A", bd=1, relief="solid", 
                                           highlightthickness=1, highlightbackground="#166534", padx=12, pady=8)
        self.frame_resumo_dados.pack(fill="x", padx=5, pady=(8, 4))

        self.lbl_status_dados = tk.Label(
            self.frame_resumo_dados, 
            text="Nenhuma planilha carregada.", 
            font=("Segoe UI", 9, "bold"), 
            bg="#131F1A", 
            fg="#4ADE80"
        )
        self.lbl_status_dados.pack(anchor="w")

    def _criar_secao_tags(self, parent):
        self.frame_tags_container = ttk.LabelFrame(parent, text=" 2. Tags Dinamicas da Planilha (Clique para inserir no texto) ")
        self.frame_tags_container.pack(fill="x", pady=6)

        self.lbl_dica_tags = tk.Label(self.frame_tags_container, 
                                      text="Dica: Posicione o cursor no Assunto ou Corpo e clique em qualquer tag abaixo para inseri-la:", 
                                      font=("Segoe UI", 8), bg=self.COLOR_CARD, fg=self.COLOR_TEXT_MUTED)
        self.lbl_dica_tags.pack(anchor="w", padx=5, pady=2)

        self.frame_botoes_tags = tk.Frame(self.frame_tags_container, bg=self.COLOR_CARD)
        self.frame_botoes_tags.pack(fill="x", padx=5, pady=4)

        self.lbl_sem_tags = tk.Label(self.frame_botoes_tags, text="Selecione uma planilha acima para exibir as tags das colunas.", 
                                     bg=self.COLOR_CARD, fg=self.COLOR_TEXT_MUTED, font=("Segoe UI", 9, "italic"))
        self.lbl_sem_tags.pack(anchor="w", pady=4)

    def _criar_secao_modelo(self, parent):
        self.frame_secao_modelo = ttk.LabelFrame(parent, text=" 3. Modelo do E-mail ")
        self.frame_secao_modelo.pack(fill="x", pady=6)

        self.lbl_info_modelo_grupo = tk.Label(
            self.frame_secao_modelo, 
            text="", 
            font=("Segoe UI", 9, "bold"), 
            bg=self.COLOR_CARD, 
            fg="#93C5FD"
        )

        row_assunto = tk.Frame(self.frame_secao_modelo, bg=self.COLOR_CARD)
        row_assunto.pack(fill="x", pady=3)

        ttk.Label(row_assunto, text="Assunto:").pack(side="left", padx=5)
        self.entry_assunto = ttk.Entry(row_assunto, textvariable=self.assunto_var, font=("Segoe UI", 10))
        self.entry_assunto.pack(side="left", fill="x", expand=True, padx=5)
        self.entry_assunto.bind("<FocusIn>", lambda e: self._definir_foco(self.entry_assunto))

        row_copia = tk.Frame(self.frame_secao_modelo, bg=self.COLOR_CARD)
        row_copia.pack(fill="x", pady=3)

        ttk.Label(row_copia, text="Cc (Com Copia):").pack(side="left", padx=5)
        self.entry_cc = ttk.Entry(row_copia, textvariable=self.cc_var, font=("Segoe UI", 9), width=28)
        self.entry_cc.pack(side="left", padx=5)
        self.entry_cc.bind("<FocusIn>", lambda e: self._definir_foco(self.entry_cc))

        ttk.Label(row_copia, text="Cco (Copia Oculta):").pack(side="left", padx=(10, 5))
        self.entry_bcc = ttk.Entry(row_copia, textvariable=self.bcc_var, font=("Segoe UI", 9), width=28)
        self.entry_bcc.pack(side="left", fill="x", expand=True, padx=5)
        self.entry_bcc.bind("<FocusIn>", lambda e: self._definir_foco(self.entry_bcc))

        row_corpo_lbl = tk.Frame(self.frame_secao_modelo, bg=self.COLOR_CARD)
        row_corpo_lbl.pack(fill="x", pady=(6, 2))

        ttk.Label(row_corpo_lbl, text="Corpo do E-mail (Texto / Tags / Links / Quebras de Linha):").pack(side="left", padx=5)
        
        btn_preview = ttk.Button(row_corpo_lbl, text="Pre-visualizar Exemplo", command=self._abrir_preview)
        btn_preview.pack(side="right", padx=5)

        self.txt_corpo = scrolledtext.ScrolledText(
            self.frame_secao_modelo, 
            wrap="word", 
            height=9, 
            font=("Segoe UI", 10),
            bg=self.COLOR_INPUT_BG,
            fg=self.COLOR_TEXT_MAIN,
            insertbackground="#FFFFFF",
            selectbackground=self.COLOR_ACCENT,
            selectforeground="#FFFFFF",
            relief="flat",
            bd=1,
            highlightthickness=1,
            highlightbackground=self.COLOR_INPUT_BORDER,
            highlightcolor=self.COLOR_ACCENT
        )
        self.txt_corpo.pack(fill="both", expand=True, padx=5, pady=4)
        self.txt_corpo.bind("<FocusIn>", lambda e: self._definir_foco(self.txt_corpo))

        self.assunto_var.set("")
        self.txt_corpo.delete("1.0", tk.END)

    def _criar_secao_anexos(self, parent):
        frame = ttk.LabelFrame(parent, text=" 4. Anexos do E-mail ")
        frame.pack(fill="x", pady=6)

        lbl_fixos = tk.Label(frame, text="Arquivos Fixos (Gerais ou do Grupo):", font=("Segoe UI", 9, "bold"), 
                             bg=self.COLOR_CARD, fg=self.COLOR_TEXT_MAIN)
        lbl_fixos.pack(anchor="w", padx=5, pady=(2, 2))

        botoes_anexos = tk.Frame(frame, bg=self.COLOR_CARD)
        botoes_anexos.pack(fill="x", pady=2)

        btn_add = ttk.Button(botoes_anexos, text="Adicionar Arquivo(s)...", command=self._adicionar_anexo)
        btn_add.pack(side="left", padx=5)

        btn_del = ttk.Button(botoes_anexos, text="Remover Selecionado", command=self._remover_anexo)
        btn_del.pack(side="left", padx=5)

        btn_clear = ttk.Button(botoes_anexos, text="Limpar Todos", command=self._limpar_anexos)
        btn_clear.pack(side="left", padx=5)

        self.listbox_anexos = tk.Listbox(
            frame, 
            height=3, 
            selectmode="single", 
            font=("Segoe UI", 9),
            bg=self.COLOR_INPUT_BG,
            fg=self.COLOR_TEXT_MAIN,
            selectbackground=self.COLOR_ACCENT,
            selectforeground="#FFFFFF",
            relief="flat",
            bd=1,
            highlightthickness=1,
            highlightbackground=self.COLOR_INPUT_BORDER,
            highlightcolor=self.COLOR_ACCENT
        )
        self.listbox_anexos.pack(fill="x", expand=True, padx=5, pady=(2, 6))

        ttk.Separator(frame, orient="horizontal").pack(fill="x", padx=5, pady=6)

        cb_dyn = ttk.Checkbutton(
            frame, 
            text="Anexar arquivo individual por destinatario a partir de uma pasta (Certificados, etc.)", 
            variable=self.usar_anexo_dinamico,
            command=self._ao_alternar_anexo_dinamico
        )
        cb_dyn.pack(anchor="w", padx=5, pady=4)

        self.frame_anexo_dinamico = tk.Frame(frame, bg="#161821", bd=1, relief="solid", 
                                             highlightthickness=1, highlightbackground=self.COLOR_CARD_BORDER, padx=10, pady=8)

        row_dyn1 = tk.Frame(self.frame_anexo_dinamico, bg="#161821")
        row_dyn1.pack(fill="x", pady=3)

        ttk.Label(row_dyn1, text="Pasta dos Arquivos:").pack(side="left", padx=5)
        self.entry_pasta_dyn = ttk.Entry(row_dyn1, textvariable=self.pasta_anexos_dinamicos, width=45)
        self.entry_pasta_dyn.pack(side="left", fill="x", expand=True, padx=5)

        btn_sel_pasta = ttk.Button(row_dyn1, text="Selecionar Pasta...", command=self._selecionar_pasta_anexos, style="Primary.TButton")
        btn_sel_pasta.pack(side="left", padx=3)

        row_dyn2 = tk.Frame(self.frame_anexo_dinamico, bg="#161821")
        row_dyn2.pack(fill="x", pady=4)

        ttk.Label(row_dyn2, text="Coluna com Nome/Chave:").pack(side="left", padx=5)
        self.combo_coluna_busca_anexo = ttk.Combobox(row_dyn2, textvariable=self.coluna_busca_anexo, state="readonly", width=18)
        self.combo_coluna_busca_anexo.pack(side="left", padx=5)
        self.combo_coluna_busca_anexo.bind("<<ComboboxSelected>>", lambda e: self._ao_mudar_coluna_busca_anexo())

        ttk.Label(row_dyn2, text="Padrao do Nome:").pack(side="left", padx=(10, 5))
        self.entry_padrao_anexo = ttk.Entry(row_dyn2, textvariable=self.padrao_busca_anexo, width=22)
        self.entry_padrao_anexo.pack(side="left", padx=5)

        lbl_hint_dyn = tk.Label(row_dyn2, text="(Ex: #nome# localiza 'NOME_certificado.pdf', 'NOME.pdf', etc.)",
                                font=("Segoe UI", 8), bg="#161821", fg=self.COLOR_TEXT_MUTED)
        lbl_hint_dyn.pack(side="left", padx=5)

        row_dyn3 = tk.Frame(self.frame_anexo_dinamico, bg="#161821")
        row_dyn3.pack(fill="x", pady=(4, 2))

        cb_skip = ttk.Checkbutton(
            row_dyn3,
            text="Pular destinatario se o arquivo nao for encontrado na pasta",
            variable=self.pular_se_sem_anexo_dinamico
        )
        cb_skip.pack(side="left", padx=5)

        btn_verificar_dyn = ttk.Button(row_dyn3, text="Verificar Arquivos na Pasta", command=self._verificar_arquivos_dinamicos, style="Small.TButton")
        btn_verificar_dyn.pack(side="left", padx=(15, 5))

        self.lbl_status_dyn = tk.Label(self.frame_anexo_dinamico, text="", font=("Segoe UI", 9, "bold"), bg="#161821", fg="#93C5FD")
        self.lbl_status_dyn.pack(anchor="w", padx=5, pady=(4, 0))

    def _criar_secao_acoes(self, parent):
        frame = ttk.LabelFrame(parent, text=" 5. Execucao e Status ")
        frame.pack(fill="x", pady=6)

        row_opts = tk.Frame(frame, bg=self.COLOR_CARD)
        row_opts.pack(fill="x", pady=4)

        rb_rascunho = tk.Radiobutton(
            row_opts, 
            text="Salvar como Rascunhos no Outlook (Recomendado)", 
            value="rascunho", 
            variable=self.tipo_acao, 
            command=self._atualizar_texto_botao_acao,
            bg=self.COLOR_CARD,
            fg=self.COLOR_TEXT_MAIN,
            selectcolor=self.COLOR_INPUT_BG,
            activebackground=self.COLOR_CARD,
            activeforeground="#FFFFFF",
            font=("Segoe UI", 9)
        )
        rb_rascunho.pack(side="left", padx=10)

        rb_enviar = tk.Radiobutton(
            row_opts, 
            text="Enviar Diretamente pelo Outlook", 
            value="enviar", 
            variable=self.tipo_acao, 
            command=self._atualizar_texto_botao_acao,
            bg=self.COLOR_CARD,
            fg=self.COLOR_TEXT_MAIN,
            selectcolor=self.COLOR_INPUT_BG,
            activebackground=self.COLOR_CARD,
            activeforeground="#FFFFFF",
            font=("Segoe UI", 9)
        )
        rb_enviar.pack(side="left", padx=10)

        row_btn = tk.Frame(frame, bg=self.COLOR_CARD)
        row_btn.pack(fill="x", pady=6)

        self.btn_iniciar = ttk.Button(row_btn, text="INICIAR CRIACAO DE RASCUNHOS", 
                                      command=self._iniciar_processamento, style="Action.TButton")
        self.btn_iniciar.pack(side="left", fill="x", expand=True, padx=5)

        self.btn_cancelar = ttk.Button(row_btn, text="Cancelar", state="disabled", command=self._cancelar_processamento)
        self.btn_cancelar.pack(side="right", padx=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(frame, variable=self.progress_var, maximum=100, style="Horizontal.TProgressbar")
        self.progress_bar.pack(fill="x", padx=5, pady=6)

        self.txt_log = scrolledtext.ScrolledText(
            frame, 
            wrap="word", 
            height=7, 
            font=("Consolas", 9), 
            bg="#0D0F14", 
            fg="#E2E8F0",
            insertbackground="#FFFFFF",
            selectbackground=self.COLOR_ACCENT,
            relief="flat",
            bd=1,
            highlightthickness=1,
            highlightbackground=self.COLOR_INPUT_BORDER
        )
        self.txt_log.pack(fill="both", expand=True, padx=5, pady=4)

        self.txt_log.tag_config("success", foreground="#4ADE80")
        self.txt_log.tag_config("warning", foreground="#FBBF24")
        self.txt_log.tag_config("error", foreground="#F87171")
        self.txt_log.tag_config("info", foreground="#60A5FA")

    def _definir_foco(self, widget):
        self.ultimo_foco = widget

    def _inserir_tag(self, tag_name):
        tag_formatada = f"#{tag_name}#"
        if self.ultimo_foco == self.entry_assunto:
            self.entry_assunto.insert(tk.INSERT, tag_formatada)
            self.entry_assunto.focus()
        else:
            self.txt_corpo.insert(tk.INSERT, tag_formatada)
            self.txt_corpo.focus()

    def _limpar_dados_planilha(self):
        self._salvar_modelo_grupo_atual()
        self.caminho_planilha.set("")
        self.aba_selecionada.set("")
        self.combo_abas['values'] = []
        self.linha_cabecalho_var.set("")
        self.combo_cabecalho['values'] = []
        self.coluna_email.set("")
        self.combo_email['values'] = []
        self.coluna_condicao.set("[Nenhum - Modelo unico para todos]")
        self.combo_condicao['values'] = ["[Nenhum - Modelo unico para todos]"]
        self.colunas_disponiveis = []
        self.registros = []
        self.sugestoes_cabecalho = []
        self.grupos_disponiveis = []
        self.grupos_selecionados_vars.clear()
        self.modelos_por_grupo.clear()
        self.grupo_anterior_edicao = None

        self.frame_painel_grupos.pack_forget()
        self.lbl_info_modelo_grupo.pack_forget()
        self.combo_coluna_busca_anexo['values'] = []
        self.coluna_busca_anexo.set("")
        self._atualizar_botoes_tags()

        self.lbl_status_dados.config(text="Nenhuma planilha carregada.", fg=self.COLOR_TEXT_MUTED, bg="#13151D")
        self.frame_resumo_dados.config(bg="#13151D", highlightbackground=self.COLOR_CARD_BORDER)
        self._atualizar_texto_botao_acao()
        self._log_msg("Planilha desmarcada/limpa.", "info")

    def _selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar Planilha",
            filetypes=[
                ("Planilhas Excel", "*.xlsx *.xls"),
                ("Arquivos CSV", "*.csv"),
                ("Todos os Arquivos", "*.*")
            ]
        )
        if caminho:
            self.caminho_planilha.set(caminho)
            self.colunas_disponiveis = []
            self.registros = []
            self.coluna_email.set("")
            self.modelos_por_grupo.clear()
            self.grupo_anterior_edicao = None

            abas = email_engine.get_sheet_names(caminho)
            if abas:
                self.combo_abas['values'] = abas
                self.aba_selecionada.set(abas[0])
                self._ao_mudar_aba()
            else:
                messagebox.showerror("Erro", "Nao foi possivel ler as abas do arquivo selecionado.")

    def _ao_mudar_aba(self):
        caminho = self.caminho_planilha.get().strip()
        aba = self.aba_selecionada.get().strip()
        if not caminho or not os.path.exists(caminho):
            return

        sugestoes, best_idx = email_engine.get_header_row_suggestions(caminho, sheet_name=aba)
        self.sugestoes_cabecalho = sugestoes
        
        labels = [s['label'] for s in sugestoes]
        self.combo_cabecalho['values'] = labels
        
        label_escolhida = labels[0]
        for s in sugestoes:
            if s['index'] == best_idx:
                label_escolhida = s['label']
                break
        self.linha_cabecalho_var.set(label_escolhida)

        self._carregar_dados_aba()

    def _ao_mudar_cabecalho(self):
        self._carregar_dados_aba()

    def _obter_indice_cabecalho_selecionado(self):
        texto_selecionado = self.linha_cabecalho_var.get().strip()
        for s in self.sugestoes_cabecalho:
            if s['label'] == texto_selecionado:
                return s['index']
        return 0

    def _carregar_dados_aba(self):
        caminho = self.caminho_planilha.get().strip()
        aba = self.aba_selecionada.get().strip()
        header_idx = self._obter_indice_cabecalho_selecionado()

        if not caminho or not os.path.exists(caminho):
            return

        try:
            colunas, registros = email_engine.load_sheet_data(caminho, sheet_name=aba, header_row_idx=header_idx)
            self.colunas_disponiveis = colunas
            self.registros = registros

            self.combo_email['values'] = colunas
            
            email_col_atual = self.coluna_email.get()
            if email_col_atual not in colunas:
                email_col = ""
                for c in colunas:
                    c_low = c.lower()
                    if any(k in c_low for k in ['email', 'e-mail', 'mail', 'endereço de e-mail', 'endereo de e-mail', 'correio']):
                        email_col = c
                        break
                if email_col:
                    self.coluna_email.set(email_col)
                elif colunas:
                    self.coluna_email.set(colunas[0])

            opcoes_condicao = ["[Nenhum - Modelo unico para todos]"] + colunas
            self.combo_condicao['values'] = opcoes_condicao
            if self.coluna_condicao.get() not in opcoes_condicao:
                self.coluna_condicao.set("[Nenhum - Modelo unico para todos]")

            self.combo_coluna_busca_anexo['values'] = colunas
            if colunas:
                if self.coluna_busca_anexo.get() not in colunas:
                    col_nome = ""
                    for c in colunas:
                        c_low = c.lower()
                        if any(k in c_low for k in ['nome', 'name', 'usuario', 'destinatario', 'pessoa', 'aluno']):
                            col_nome = c
                            break
                    if col_nome:
                        self.coluna_busca_anexo.set(col_nome)
                        self.padrao_busca_anexo.set(f"#{col_nome}#")
                    else:
                        self.coluna_busca_anexo.set(colunas[0])
                        self.padrao_busca_anexo.set(f"#{colunas[0]}#")

            self._atualizar_painel_condicao()
            self._atualizar_botoes_tags()
            self._atualizar_contagem_destinatarios()

            self._log_msg(f"Planilha carregada: {os.path.basename(caminho)} | Aba: '{aba}' | {len(registros)} registros de dados.", "info")

        except Exception as e:
            messagebox.showerror("Erro ao carregar dados", f"Erro: {e}")
            self._log_msg(f"Erro ao carregar planilha: {e}", "error")

    def _ao_mudar_coluna_condicao(self):
        self._atualizar_painel_condicao()
        self._atualizar_contagem_destinatarios()

    def _atualizar_painel_condicao(self):
        col_cond = self.coluna_condicao.get()
        if not col_cond or col_cond.startswith("[Nenhum"):
            self.frame_painel_grupos.pack_forget()
            self.lbl_info_modelo_grupo.pack_forget()
            self.grupos_disponiveis = []
            return

        contagem_grupos = {}
        for r in self.registros:
            val = str(r.get(col_cond, '')).strip()
            if not val:
                val = "(Vazio)"
            contagem_grupos[val] = contagem_grupos.get(val, 0) + 1

        self.grupos_disponiveis = sorted(contagem_grupos.keys())

        for w in self.frame_grupos_checkboxes.winfo_children():
            w.destroy()

        self.grupos_selecionados_vars.clear()

        MAX_COLS = 4
        for idx, grp in enumerate(self.grupos_disponiveis):
            var = tk.BooleanVar(value=True)
            self.grupos_selecionados_vars[grp] = var

            r = idx // MAX_COLS
            c = idx % MAX_COLS
            
            lbl_texto = f"{grp} ({contagem_grupos[grp]})"
            cb = tk.Checkbutton(
                self.frame_grupos_checkboxes, 
                text=lbl_texto, 
                variable=var,
                command=self._atualizar_contagem_destinatarios,
                bg="#161821",
                fg=self.COLOR_TEXT_MAIN,
                selectcolor=self.COLOR_INPUT_BG,
                activebackground="#161821",
                activeforeground="#FFFFFF",
                font=("Segoe UI", 9)
            )
            cb.grid(row=r, column=c, padx=6, pady=2, sticky="w")

        self.combo_grupo_edicao['values'] = self.grupos_disponiveis
        if self.grupos_disponiveis:
            if self.grupo_em_edicao.get() not in self.grupos_disponiveis:
                self.grupo_em_edicao.set(self.grupos_disponiveis[0])

        self.frame_painel_grupos.pack(fill="x", padx=5, pady=(6, 4), after=self.combo_condicao.master)
        self._ao_trocar_modo_condicao()

    def _marcar_desmarcar_grupos(self, marcar):
        for var in self.grupos_selecionados_vars.values():
            var.set(marcar)
        self._atualizar_contagem_destinatarios()

    def _ao_trocar_modo_condicao(self):
        modo = self.modo_condicao.get()
        if modo == "por_grupo" and self.grupos_disponiveis:
            self.frame_seletor_grupo_edicao.pack(fill="x", pady=(6, 2))
            self.lbl_info_modelo_grupo.pack(anchor="w", padx=5, pady=(0, 4), before=self.entry_assunto.master)
            self._carregar_modelo_grupo_atual()
        else:
            self.frame_seletor_grupo_edicao.pack_forget()
            self.lbl_info_modelo_grupo.pack_forget()

    def _salvar_modelo_grupo_atual(self):
        if self.modo_condicao.get() == "por_grupo" and self.grupo_anterior_edicao:
            self.modelos_por_grupo[self.grupo_anterior_edicao] = {
                'subject': self.assunto_var.get(),
                'cc': self.cc_var.get().strip(),
                'bcc': self.bcc_var.get().strip(),
                'body': self.txt_corpo.get("1.0", tk.END).strip(),
                'attachments': list(self.lista_anexos)
            }

    def _ao_mudar_grupo_em_edicao(self):
        self._salvar_modelo_grupo_atual()
        self._carregar_modelo_grupo_atual()

    def _carregar_modelo_grupo_atual(self):
        grp = self.grupo_em_edicao.get()
        if not grp:
            return

        self.grupo_anterior_edicao = grp
        self.lbl_info_modelo_grupo.config(text=f"Editando Modelo Exclusivo para o Grupo: {grp}")

        if grp in self.modelos_por_grupo:
            dados = self.modelos_por_grupo[grp]
            self.assunto_var.set(dados.get('subject', ''))
            self.cc_var.set(dados.get('cc', ''))
            self.bcc_var.set(dados.get('bcc', ''))
            self.txt_corpo.delete("1.0", tk.END)
            self.txt_corpo.insert("1.0", dados.get('body', ''))
            self.lista_anexos = list(dados.get('attachments', []))
        else:
            self.modelos_por_grupo[grp] = {
                'subject': self.assunto_var.get(),
                'cc': self.cc_var.get().strip(),
                'bcc': self.bcc_var.get().strip(),
                'body': self.txt_corpo.get("1.0", tk.END).strip(),
                'attachments': list(self.lista_anexos)
            }

        self._atualizar_listbox_anexos()

    def _copiar_modelo_atual_para_todos(self):
        assunto = self.assunto_var.get()
        cc = self.cc_var.get().strip()
        bcc = self.bcc_var.get().strip()
        corpo = self.txt_corpo.get("1.0", tk.END).strip()
        anexos = list(self.lista_anexos)
        
        for grp in self.grupos_disponiveis:
            self.modelos_por_grupo[grp] = {
                'subject': assunto,
                'cc': cc,
                'bcc': bcc,
                'body': corpo,
                'attachments': list(anexos)
            }
        messagebox.showinfo("Sucesso", "O modelo e anexos atuais foram copiados para todos os grupos!")

    def _obter_registros_filtrados(self):
        col_cond = self.coluna_condicao.get()
        if not col_cond or col_cond.startswith("[Nenhum"):
            return self.registros

        grupos_ativos = {grp for grp, var in self.grupos_selecionados_vars.items() if var.get()}
        filtrados = []
        for r in self.registros:
            val = str(r.get(col_cond, '')).strip()
            if not val:
                val = "(Vazio)"
            if val in grupos_ativos:
                filtrados.append(r)
        return filtrados

    def _atualizar_contagem_destinatarios(self):
        col_email = self.coluna_email.get()
        registros_filtrados = self._obter_registros_filtrados()
        total_filtrados = len(registros_filtrados)
        
        emails_validos = 0
        if col_email and registros_filtrados:
            for r in registros_filtrados:
                em = r.get(col_email, '').strip()
                if em and '@' in em:
                    emails_validos += 1

        aba = self.aba_selecionada.get() or "Padrao"
        header_idx = self._obter_indice_cabecalho_selecionado()
        col_cond = self.coluna_condicao.get()

        info_cond = ""
        if col_cond and not col_cond.startswith("[Nenhum"):
            grupos_ativos = [grp for grp, var in self.grupos_selecionados_vars.items() if var.get()]
            info_cond = f" | Filtro ({col_cond}): {len(grupos_ativos)} de {len(self.grupos_disponiveis)} grupos marcados"

        if emails_validos > 0:
            self.frame_resumo_dados.config(bg="#131F1A", highlightbackground="#166534")
            self.lbl_status_dados.config(
                bg="#131F1A",
                fg="#4ADE80",
                text=f"Aba: '{aba}' | {len(self.colunas_disponiveis)} colunas (Cabecalho Linha {header_idx+1}){info_cond}\n"
                     f"TOTAL A PROCESSAR: {emails_validos} destinatarios com e-mail valido (de {total_filtrados} linhas selecionadas)"
            )
        elif total_filtrados > 0:
            self.frame_resumo_dados.config(bg="#261414", highlightbackground="#991B1B")
            self.lbl_status_dados.config(
                bg="#261414",
                fg="#F87171",
                text=f"{total_filtrados} linhas selecionadas na aba '{aba}', mas nenhum e-mail valido com '@' na coluna '{col_email}'.\n"
                     f"Verifique se a 'Coluna E-mail' ou a 'Linha do Cabecalho' estao selecionadas corretamente."
            )
        else:
            self.frame_resumo_dados.config(bg="#13151D", highlightbackground=self.COLOR_CARD_BORDER)
            self.lbl_status_dados.config(
                bg="#13151D",
                fg=self.COLOR_TEXT_MUTED,
                text="Nenhum registro encontrado para os filtros selecionados."
            )

        self._atualizar_texto_botao_acao(emails_validos)

    def _atualizar_texto_botao_acao(self, qtd_emails=None):
        if qtd_emails is None:
            col_email = self.coluna_email.get()
            recs = self._obter_registros_filtrados()
            qtd_emails = sum(1 for r in recs if '@' in r.get(col_email, ''))

        is_draft = (self.tipo_acao.get() == "rascunho")
        if is_draft:
            if qtd_emails > 0:
                self.btn_iniciar.config(text=f"CRIAR {qtd_emails} RASCUNHOS NO OUTLOOK")
            else:
                self.btn_iniciar.config(text="INICIAR CRIACAO DE RASCUNHOS")
        else:
            if qtd_emails > 0:
                self.btn_iniciar.config(text=f"ENVIAR {qtd_emails} E-MAILS DIRETAMENTE")
            else:
                self.btn_iniciar.config(text="ENVIAR DIRETAMENTE PELO OUTLOOK")

    def _atualizar_botoes_tags(self):
        for widget in self.frame_botoes_tags.winfo_children():
            widget.destroy()

        if not self.colunas_disponiveis:
            lbl = tk.Label(self.frame_botoes_tags, text="Selecione uma planilha acima para exibir as tags das colunas.", 
                           bg=self.COLOR_CARD, fg=self.COLOR_TEXT_MUTED)
            lbl.pack(anchor="w")
            return

        MAX_COLS_PER_ROW = 5
        for idx, col in enumerate(self.colunas_disponiveis):
            r = idx // MAX_COLS_PER_ROW
            c = idx % MAX_COLS_PER_ROW
            
            texto_btn = f"#{col}#"
            btn = ttk.Button(self.frame_botoes_tags, text=texto_btn, 
                             command=lambda col_name=col: self._inserir_tag(col_name), style="Tag.TButton")
            btn.grid(row=r, column=c, padx=3, pady=2, sticky="ew")

        for c_idx in range(MAX_COLS_PER_ROW):
            self.frame_botoes_tags.columnconfigure(c_idx, weight=1)

    def _adicionar_anexo(self):
        arquivos = filedialog.askopenfilenames(
            title="Selecionar Anexo(s)",
            filetypes=[("Todos os Arquivos", "*.*")]
        )
        for arq in arquivos:
            if arq not in self.lista_anexos:
                self.lista_anexos.append(arq)
        self._atualizar_listbox_anexos()
        self._salvar_modelo_grupo_atual()

    def _remover_anexo(self):
        selecionados = self.listbox_anexos.curselection()
        if selecionados:
            idx = selecionados[0]
            del self.lista_anexos[idx]
            self._atualizar_listbox_anexos()
            self._salvar_modelo_grupo_atual()

    def _limpar_anexos(self):
        self.lista_anexos.clear()
        self._atualizar_listbox_anexos()
        self._salvar_modelo_grupo_atual()

    def _ao_alternar_anexo_dinamico(self):
        if self.usar_anexo_dinamico.get():
            self.frame_anexo_dinamico.pack(fill="x", padx=5, pady=(4, 4))
        else:
            self.frame_anexo_dinamico.pack_forget()

    def _selecionar_pasta_anexos(self):
        pasta = filedialog.askdirectory(title="Selecionar Pasta dos Certificados / Anexos Individuais")
        if pasta:
            self.pasta_anexos_dinamicos.set(pasta)
            self._verificar_arquivos_dinamicos()

    def _ao_mudar_coluna_busca_anexo(self):
        col = self.coluna_busca_anexo.get()
        if col:
            self.padrao_busca_anexo.set(f"#{col}#")

    def _verificar_arquivos_dinamicos(self):
        pasta = self.pasta_anexos_dinamicos.get().strip()
        if not pasta or not os.path.exists(pasta):
            messagebox.showwarning("Aviso", "Por favor, selecione uma pasta valida para os arquivos.")
            return

        recs = self._obter_registros_filtrados()
        if not recs:
            messagebox.showwarning("Aviso", "Nenhum registro carregado da planilha para verificar.")
            return

        col = self.coluna_busca_anexo.get().strip()
        padrao = self.padrao_busca_anexo.get().strip()

        resultado = email_engine.verify_dynamic_attachments(pasta, recs, column_name=col, pattern_template=padrao)

        total = resultado['total']
        found = resultado['found']
        missing = resultado['missing']

        if missing == 0:
            msg = f"Sucesso! Todos os {found} destinatarios possuem arquivo correspondente na pasta."
            self.lbl_status_dyn.config(text=msg, fg="#4ADE80")
            messagebox.showinfo("Verificacao Concluida", msg)
        else:
            faltantes_str = "\n- ".join(str(x) for x in resultado['missing_list'][:10])
            if len(resultado['missing_list']) > 10:
                faltantes_str += f"\n... e mais {len(resultado['missing_list']) - 10} registros."
            msg = f"{found} de {total} arquivos encontrados. {missing} destinatarios NAO possuem arquivo na pasta:\n\n- {faltantes_str}"
            self.lbl_status_dyn.config(text=f"{found} de {total} arquivos encontrados ({missing} sem arquivo)", fg="#F87171")
            messagebox.showwarning("Verificacao de Arquivos", msg)

    def _atualizar_listbox_anexos(self):
        self.listbox_anexos.delete(0, tk.END)
        for arq in self.lista_anexos:
            if os.path.exists(arq):
                tamanho_kb = os.path.getsize(arq) / 1024
                nome_formatado = f"{os.path.basename(arq)} ({tamanho_kb:.1f} KB) - {arq}"
            else:
                nome_formatado = f"{os.path.basename(arq)} (Arquivo nao encontrado) - {arq}"
            self.listbox_anexos.insert(tk.END, nome_formatado)

    def _abrir_preview(self):
        self._salvar_modelo_grupo_atual()
        registros_filtrados = self._obter_registros_filtrados()
        if not registros_filtrados:
            messagebox.showwarning("Aviso", "Nenhum registro encontrado para os filtros selecionados.")
            return

        col_cond = self.coluna_condicao.get()
        eh_por_grupo = (self.modo_condicao.get() == "por_grupo" and col_cond and not col_cond.startswith("[Nenhum"))
        
        registro_exemplo = registros_filtrados[0]
        grp_edicao = self.grupo_em_edicao.get()

        if eh_por_grupo and grp_edicao:
            for r in registros_filtrados:
                val = str(r.get(col_cond, '')).strip() or "(Vazio)"
                if val == grp_edicao:
                    registro_exemplo = r
                    break

        assunto_modelo = self.assunto_var.get()
        cc_modelo = self.cc_var.get().strip()
        bcc_modelo = self.bcc_var.get().strip()
        corpo_modelo = self.txt_corpo.get("1.0", tk.END).strip()
        anexos_modelo = list(self.lista_anexos)

        assunto_renderizado = email_engine.render_template(assunto_modelo, registro_exemplo)
        cc_renderizado = email_engine.render_template(cc_modelo, registro_exemplo) if cc_modelo else ""
        bcc_renderizado = email_engine.render_template(bcc_modelo, registro_exemplo) if bcc_modelo else ""
        corpo_renderizado = email_engine.render_template(corpo_modelo, registro_exemplo)

        janela = tk.Toplevel(self)
        janela.title("Pre-visualizacao do E-mail")
        janela.geometry("680x590")
        janela.configure(bg=self.COLOR_BG)

        card_preview = tk.Frame(janela, bg=self.COLOR_CARD, bd=1, relief="solid", 
                                highlightthickness=1, highlightbackground=self.COLOR_CARD_BORDER, padx=15, pady=12)
        card_preview.pack(fill="both", expand=True, padx=15, pady=15)

        info_grupo_txt = f" (Grupo: {grp_edicao})" if eh_por_grupo else ""
        lbl_info = tk.Label(card_preview, text=f"Exemplo gerado com os dados do 1o destinatario{info_grupo_txt}:", 
                            font=("Segoe UI", 9, "bold"), bg=self.COLOR_CARD, fg=self.COLOR_TEXT_MAIN)
        lbl_info.pack(anchor="w", pady=(0, 6))

        col_email = self.coluna_email.get()
        destinatario = registro_exemplo.get(col_email, 'E-mail nao encontrado')
        tk.Label(card_preview, text=f"Para: {destinatario}", font=("Segoe UI", 9), bg=self.COLOR_CARD, fg=self.COLOR_TEXT_MUTED).pack(anchor="w")

        if cc_renderizado:
            tk.Label(card_preview, text=f"Cc: {cc_renderizado}", font=("Segoe UI", 9), bg=self.COLOR_CARD, fg=self.COLOR_TEXT_MUTED).pack(anchor="w")

        if bcc_renderizado:
            tk.Label(card_preview, text=f"Cco: {bcc_renderizado}", font=("Segoe UI", 9), bg=self.COLOR_CARD, fg=self.COLOR_TEXT_MUTED).pack(anchor="w")

        tk.Label(card_preview, text=f"Assunto: {assunto_renderizado}", font=("Segoe UI", 10, "bold"), bg=self.COLOR_CARD, fg="#93C5FD").pack(anchor="w", pady=4)

        ttk.Separator(card_preview, orient="horizontal").pack(fill="x", pady=6)

        txt = scrolledtext.ScrolledText(
            card_preview, 
            wrap="word", 
            font=("Calibri", 11), 
            bg=self.COLOR_INPUT_BG,
            fg=self.COLOR_TEXT_MAIN,
            insertbackground="#FFFFFF",
            relief="flat",
            bd=1,
            highlightthickness=1,
            highlightbackground=self.COLOR_INPUT_BORDER
        )
        txt.pack(fill="both", expand=True, pady=5)
        txt.insert("1.0", corpo_renderizado)
        txt.config(state="disabled")

        anexos_info_partes = []
        if anexos_modelo:
            anexos_info_partes.append("Fixos: " + ", ".join(os.path.basename(a) for a in anexos_modelo))
        
        if self.usar_anexo_dinamico.get():
            pasta_dyn = self.pasta_anexos_dinamicos.get().strip()
            if pasta_dyn and os.path.exists(pasta_dyn):
                col_b = self.coluna_busca_anexo.get().strip()
                padr_b = self.padrao_busca_anexo.get().strip()
                dyn_match = email_engine.find_matching_attachment(pasta_dyn, registro_exemplo, col_b, padr_b)
                if dyn_match:
                    anexos_info_partes.append(f"Individual (Pasta): {os.path.basename(dyn_match)}")
                else:
                    anexos_info_partes.append("(Individual: Nenhum arquivo correspondente encontrado na pasta)")

        if anexos_info_partes:
            lbl_anexos = tk.Label(card_preview, text="Anexos: " + " | ".join(anexos_info_partes), 
                                  font=("Segoe UI", 8, "italic"), bg=self.COLOR_CARD, fg=self.COLOR_TEXT_MUTED)
            lbl_anexos.pack(anchor="w", pady=5)

        btn_fechar = ttk.Button(card_preview, text="Fechar", command=janela.destroy)
        btn_fechar.pack(pady=(6, 0))

    def _log_msg(self, mensagem, tag="info"):
        self.txt_log.insert(tk.END, f"{mensagem}\n", tag)
        self.txt_log.see(tk.END)

    def _iniciar_processamento(self):
        self._salvar_modelo_grupo_atual()
        
        registros_a_processar = self._obter_registros_filtrados()
        if not registros_a_processar:
            messagebox.showerror("Erro", "Nenhum registro selecionado para processamento.")
            return

        col_email = self.coluna_email.get()
        if not col_email:
            messagebox.showerror("Erro", "Por favor, selecione a coluna que contem o e-mail dos destinatarios.")
            return

        destinatarios_validos = sum(1 for r in registros_a_processar if '@' in r.get(col_email, ''))
        if destinatarios_validos == 0:
            messagebox.showerror(
                "Erro de E-mails",
                f"Nenhum e-mail valido com '@' foi encontrado na coluna '{col_email}'.\n"
                f"Por favor, verifique a selecao da coluna de e-mail ou do cabecalho."
            )
            return

        col_cond = self.coluna_condicao.get()
        eh_por_grupo = (self.modo_condicao.get() == "por_grupo" and col_cond and not col_cond.startswith("[Nenhum"))

        assunto_padrao = self.assunto_var.get().strip()
        cc_padrao = self.cc_var.get().strip()
        bcc_padrao = self.bcc_var.get().strip()
        corpo_padrao = self.txt_corpo.get("1.0", tk.END).strip()

        if not eh_por_grupo:
            if not assunto_padrao:
                messagebox.showerror("Erro", "O assunto do e-mail nao pode estar vazio.")
                return
            if not corpo_padrao:
                messagebox.showerror("Erro", "O corpo do e-mail nao pode estar vazio.")
                return
        else:
            grupos_ativos = [grp for grp, var in self.grupos_selecionados_vars.items() if var.get()]
            for grp in grupos_ativos:
                tpl = self.modelos_por_grupo.get(grp, {})
                sub = tpl.get('subject', '').strip()
                bod = tpl.get('body', '').strip()
                if not sub or not bod:
                    messagebox.showerror("Erro", f"O modelo do grupo '{grp}' possui assunto ou corpo vazio. Por favor, preencha o modelo do grupo.")
                    return

        pasta_dyn = None
        col_dyn = None
        padrao_dyn = None
        skip_dyn = False
        if self.usar_anexo_dinamico.get():
            pasta_dyn = self.pasta_anexos_dinamicos.get().strip()
            if not pasta_dyn or not os.path.exists(pasta_dyn):
                messagebox.showerror("Erro", "Voce ativou o anexo individual por pasta, mas o caminho informado nao existe ou nao foi selecionado.")
                return
            col_dyn = self.coluna_busca_anexo.get().strip()
            padrao_dyn = self.padrao_busca_anexo.get().strip()
            skip_dyn = self.pular_se_sem_anexo_dinamico.get()

        is_draft = (self.tipo_acao.get() == "rascunho")

        if not is_draft:
            modo_desc = f"Modelos exclusivos por grupo ({col_cond})" if eh_por_grupo else "Modelo unico para todos"
            msg_confirmacao = (
                "CONFIRMACAO DE ENVIO DIRETO\n\n"
                "ATENCAO: Voce selecionou a opcao de envio direto pelo Outlook.\n"
                "Os e-mails serao enviados imediatamente para as caixas de entrada dos destinatarios.\n\n"
                f"- Total de destinatarios a enviar: {destinatarios_validos}\n"
                f"- Coluna de e-mail: {col_email}\n"
                f"- Modo de envio: {modo_desc}\n\n"
                "Deseja realmente ENVIAR estes e-mails agora?"
            )
            if not messagebox.askyesno("Confirmar Envio Direto de E-mails", msg_confirmacao):
                return

        self.em_execucao = True
        self.cancel_event.clear()
        self.btn_iniciar.config(state="disabled")
        self.btn_cancelar.config(state="normal")
        self.progress_var.set(0)

        tipo_nome = "Rascunhos" if is_draft else "Envio Direto"
        self._log_msg(f"\n--- Iniciando processamento de {destinatarios_validos} e-mails ({tipo_nome}) ---", "info")

        group_column = col_cond if eh_por_grupo else None
        group_templates = self.modelos_por_grupo if eh_por_grupo else None

        thread = threading.Thread(
            target=self._executar_em_background,
            args=(registros_a_processar, col_email, assunto_padrao, corpo_padrao, self.lista_anexos, is_draft, group_column, group_templates, pasta_dyn, col_dyn, padrao_dyn, skip_dyn, cc_padrao, bcc_padrao),
            daemon=True
        )
        thread.start()

    def _executar_em_background(self, records, col_email, assunto, corpo, anexos, is_draft, group_column, group_templates, dynamic_folder, dynamic_column, dynamic_pattern, skip_if_missing_dynamic_file, cc_template, bcc_template):
        def callback(current, total, status, msg, item):
            self.after(0, self._atualizar_progresso, current, total, status, msg)

        res = email_engine.process_emails(
            records=records,
            email_column=col_email,
            subject_template=assunto,
            body_template=corpo,
            attachments=anexos,
            is_draft=is_draft,
            progress_callback=callback,
            cancel_event=self.cancel_event,
            group_column=group_column,
            group_templates=group_templates,
            dynamic_folder=dynamic_folder,
            dynamic_column=dynamic_column,
            dynamic_pattern=dynamic_pattern,
            skip_if_missing_dynamic_file=skip_if_missing_dynamic_file,
            cc_template=cc_template,
            bcc_template=bcc_template
        )

        self.after(0, self._finalizar_processamento, res, is_draft)

    def _atualizar_progresso(self, current, total, status, msg):
        if total > 0:
            porcentagem = (current / total) * 100
            self.progress_var.set(porcentagem)
        self._log_msg(msg, status)

    def _finalizar_processamento(self, res, is_draft):
        self.em_execucao = False
        self.btn_iniciar.config(state="normal")
        self.btn_cancelar.config(state="disabled")

        tipo_str = "Rascunhos criados" if is_draft else "E-mails enviados"
        resumo = f"\n=== PROCESSO CONCLUIDO ===\nTotal de Linhas Processadas: {res['total']} | Sucesso ({tipo_str}): {res['success']} | Falhas: {res['errors']} | Pulados: {res['skipped']}\n"
        self._log_msg(resumo, "success" if res['errors'] == 0 else "warning")

        if is_draft:
            messagebox.showinfo("Sucesso", f"{res['success']} rascunhos foram criados com sucesso na sua pasta Rascunhos do Outlook!")
        else:
            messagebox.showinfo("Sucesso", f"{res['success']} e-mails foram enviados com sucesso!")

    def _cancelar_processamento(self):
        if self.em_execucao:
            self.cancel_event.set()
            self._log_msg("Cancelamento solicitado...", "warning")

if __name__ == "__main__":
    app = EmailSenderApp()
    app.mainloop()




