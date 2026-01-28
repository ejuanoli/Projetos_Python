import os
import shutil
import re
import threading
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import win32api
import win32print
from pypdf import PdfReader, PdfWriter

# --- CORES DHL ---
DHL_YELLOW = "#FFCC00"
DHL_RED = "#D40511"
BG_WHITE = "#FFFFFF"
BG_GREY = "#F2F2F2"


class ImpressorDHL:
    def __init__(self, root):
        self.root = root
        self.root.title("DHL Fiscal Print Manager")

        # Inicia maximizado
        try:
            self.root.state('zoomed')
        except:
            self.root.attributes('-fullscreen', True)

        self.root.configure(bg=BG_WHITE)

        # Variáveis de Imagem (Para não serem deletadas da memória)
        self.img_logo_header = None
        self.img_icon_window = None

        # ------------------------------------------------------------------
        # CONFIGURAÇÃO DO ÍCONE DA JANELA (Barra de Título e Barra de Tarefas)
        # ------------------------------------------------------------------
        try:
            # Tenta carregar o PNG para usar como ícone
            if os.path.exists("dhl_logo.png"):
                self.img_icon_window = tk.PhotoImage(file="dhl_logo.png")
                # O True indica que esse ícone vale para todas as janelas futuras do app
                self.root.iconphoto(True, self.img_icon_window)
            elif os.path.exists("logo.ico"):
                self.root.iconbitmap("logo.ico")
        except Exception as e:
            print(f"Erro ao carregar icone: {e}")

        # Variáveis de Dados
        self.pasta_origem = tk.StringVar()
        self.dados_memoria = {}

        # --- ESTILOS ---
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview", background=BG_WHITE, foreground="black",
                        fieldbackground=BG_WHITE, rowheight=30, font=('Segoe UI', 10))
        style.configure("Treeview.Heading", background=DHL_RED, foreground="white",
                        font=('Segoe UI', 10, 'bold'), relief="flat")
        style.map("Treeview", background=[('selected', DHL_YELLOW)], foreground=[('selected', 'black')])

        # --- LAYOUT ---

        # 1. CABEÇALHO
        header_frame = tk.Frame(root, bg=DHL_YELLOW, height=110)
        header_frame.pack(fill=tk.X, side=tk.TOP)

        top_content = tk.Frame(header_frame, bg=DHL_YELLOW)
        top_content.pack(expand=True, fill=tk.BOTH, padx=20, pady=10)

        # LOGO GRANDE NO CABEÇALHO
        try:
            if os.path.exists("dhl_logo.png"):
                self.img_logo_header = tk.PhotoImage(file="dhl_logo.png")
                # Opcional: Reduzir se a imagem original for muito grande (ex: maior que 400px)
                # self.img_logo_header = self.img_logo_header.subsample(2, 2)

                lbl_logo = tk.Label(top_content, image=self.img_logo_header, bg=DHL_YELLOW)
                lbl_logo.pack(side=tk.LEFT, padx=(0, 25))
            else:
                lbl_logo = tk.Label(top_content, text="DHL", bg=DHL_YELLOW, fg=DHL_RED,
                                    font=('Arial Black', 30, 'italic'))
                lbl_logo.pack(side=tk.LEFT, padx=(0, 25))
        except:
            pass

        # TÍTULO
        lbl_title = tk.Label(top_content, text="GERENCIADOR DE IMPRESSÃO FISCAL",
                             bg=DHL_YELLOW, fg=DHL_RED, font=('Segoe UI', 20, 'bold', 'italic'))
        lbl_title.pack(side=tk.LEFT, anchor='center')

        # 2. CONTROLES
        ctrl_frame = tk.Frame(root, bg=BG_GREY, pady=15)
        ctrl_frame.pack(fill=tk.X)

        tk.Label(ctrl_frame, text="DIRETÓRIO:", bg=BG_GREY, font=('Segoe UI', 9, 'bold')).grid(row=0, column=0,
                                                                                               padx=(20, 5))
        self.entry_pasta = tk.Entry(ctrl_frame, textvariable=self.pasta_origem, width=50, font=('Segoe UI', 10),
                                    relief='solid')
        self.entry_pasta.grid(row=0, column=1, padx=5, ipady=3)
        btn_pasta = tk.Button(ctrl_frame, text="SELECIONAR", command=self.escolher_pasta, bg='white', relief='raised')
        btn_pasta.grid(row=0, column=2, padx=5)

        tk.Label(ctrl_frame, text="IMPRESSORA:", bg=BG_GREY, font=('Segoe UI', 9, 'bold')).grid(row=0, column=3,
                                                                                                padx=(40, 5))
        self.printers = [p[2] for p in
                         win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        self.combo_printer = ttk.Combobox(ctrl_frame, values=self.printers, width=30, state="readonly",
                                          font=('Segoe UI', 10))
        if self.printers: self.combo_printer.current(0)
        self.combo_printer.grid(row=0, column=4, padx=5, ipady=3)

        # 3. TABELA
        mid_frame = tk.Frame(root, bg=BG_WHITE)
        mid_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        cols = ("check", "id", "chave", "nf", "cte")
        self.tree = ttk.Treeview(mid_frame, columns=cols, show='headings', selectmode="browse")

        self.tree.heading("check", text="SEL");
        self.tree.column("check", width=50, stretch=False, anchor='center')
        self.tree.heading("id", text="#");
        self.tree.column("id", width=50, stretch=False, anchor='center')
        self.tree.heading("chave", text="CHAVE DE ACESSO (NF)");
        self.tree.column("chave", width=350, stretch=True)
        self.tree.heading("nf", text="ARQUIVO NF");
        self.tree.column("nf", width=250, stretch=True)
        self.tree.heading("cte", text="ARQUIVO CTE");
        self.tree.column("cte", width=250, stretch=True)

        scrollbar = ttk.Scrollbar(mid_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind('<Button-1>', self.on_click_tree)

        # 4. RODAPÉ
        bot_frame = tk.Frame(root, bg=DHL_YELLOW, height=80)
        bot_frame.pack(fill=tk.X, side=tk.BOTTOM)

        btn_scan = tk.Button(bot_frame, text="ATUALIZAR LISTA", command=self.iniciar_scan, bg=BG_WHITE,
                             font=('Segoe UI', 10, 'bold'), relief='flat', width=20)
        btn_scan.pack(side=tk.LEFT, padx=20, pady=15, ipady=5)

        btn_toggle = tk.Button(bot_frame, text="MARCAR TODOS", command=self.toggle_all, bg=BG_WHITE,
                               font=('Segoe UI', 10), relief='flat')
        btn_toggle.pack(side=tk.LEFT, padx=5, pady=15, ipady=5)

        self.lbl_resumo = tk.Label(bot_frame, text="SELECIONADOS: 0  |  TOTAL FOLHAS: 0", bg=DHL_YELLOW, fg=DHL_RED,
                                   font=('Segoe UI', 14, 'bold'))
        self.lbl_resumo.pack(side=tk.LEFT, padx=50, pady=15)

        btn_print = tk.Button(bot_frame, text="IMPRIMIR AGORA", command=self.iniciar_impressao, bg=DHL_RED, fg='white',
                              font=('Segoe UI', 12, 'bold'), relief='flat', width=25)
        btn_print.pack(side=tk.RIGHT, padx=20, pady=15, ipady=5)

        self.lbl_status = tk.Label(root, text=" Sistema pronto.", bg='#333', fg='white', anchor='w', font=('Arial', 8))
        self.lbl_status.pack(side=tk.BOTTOM, fill=tk.X)

    # --- FUNÇÕES LÓGICAS (Idênticas ao anterior) ---
    def escolher_pasta(self):
        folder = filedialog.askdirectory()
        if folder:
            self.pasta_origem.set(folder)
            self.iniciar_scan()

    def update_status(self, msg):
        self.lbl_status.config(text=f" {msg}")
        self.root.update_idletasks()

    def contar_paginas(self, caminho):
        try:
            reader = PdfReader(caminho)
            return len(reader.pages)
        except:
            return 0

    def calcular_resumo(self):
        total_pares = 0
        total_paginas = 0
        for item in self.tree.get_children():
            vals = self.tree.item(item, "values")
            if vals[0] == "[X]":
                total_pares += 1
                chave = vals[2]
                if chave in self.dados_memoria:
                    d = self.dados_memoria[chave]
                    if d['nf_path']: total_paginas += d['nf_pgs']
                    if d['cte_path']: total_paginas += d['cte_pgs']
        self.lbl_resumo.config(text=f"SELECIONADOS: {total_pares}  |  TOTAL FOLHAS: {total_paginas}")

    def on_click_tree(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region == "heading": return
        col = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if col == '#1' and row_id:
            curr = self.tree.item(row_id, "values")[0]
            new_val = "[X]" if curr == "[ ]" else "[ ]"
            vals = list(self.tree.item(row_id, "values"))
            vals[0] = new_val
            self.tree.item(row_id, values=vals)
            self.calcular_resumo()

    def toggle_all(self):
        items = self.tree.get_children()
        if not items: return
        primeiro = self.tree.item(items[0], "values")[0]
        novo = "[ ]" if primeiro == "[X]" else "[X]"
        for item in items:
            vals = list(self.tree.item(item, "values"))
            vals[0] = novo
            self.tree.item(item, values=vals)
        self.calcular_resumo()

    def iniciar_scan(self):
        pasta = self.pasta_origem.get()
        if not pasta or not os.path.exists(pasta):
            messagebox.showwarning("Aviso", "Selecione uma pasta.")
            return
        for i in self.tree.get_children(): self.tree.delete(i)
        self.dados_memoria = {}
        self.lbl_resumo.config(text="CALCULANDO...")
        t = threading.Thread(target=self.thread_scan, args=(pasta,))
        t.start()

    def extrair_chaves_pdf(self, caminho):
        try:
            reader = PdfReader(caminho)
            texto = "".join([p.extract_text() for p in reader.pages])
            return re.findall(r'\b[0-9]{44}\b', texto)
        except:
            return []

    def thread_scan(self, pasta):
        self.update_status("Lendo arquivos...")
        arquivos = [f for f in os.listdir(pasta) if f.lower().endswith('.pdf')]
        temp_dados = {}
        ctes_pendentes = []

        for i, f in enumerate(arquivos):
            if i % 10 == 0: self.update_status(f"Analisando {i}/{len(arquivos)}...")
            if "DACTE" not in f.upper():
                match = re.search(r'[0-9]{44}', f)
                if match:
                    chave = match.group(0)
                    pgs = self.contar_paginas(os.path.join(pasta, f))
                    temp_dados[chave] = {'nf_path': f, 'nf_pgs': pgs, 'cte_path': None, 'cte_pgs': 0}
            else:
                ctes_pendentes.append(f)

        self.update_status("Vinculando CTEs...")
        for f_cte in ctes_pendentes:
            full_path = os.path.join(pasta, f_cte)
            chaves = self.extrair_chaves_pdf(full_path)
            pgs = self.contar_paginas(full_path)
            for c in chaves:
                if c in temp_dados and temp_dados[c]['cte_path'] is None:
                    temp_dados[c]['cte_path'] = f_cte
                    temp_dados[c]['cte_pgs'] = pgs
                    break

        self.dados_memoria = temp_dados
        chaves_ord = sorted(temp_dados.keys())
        for idx, chave in enumerate(chaves_ord, start=1):
            d = temp_dados[chave]
            sel = "[X]" if (d['nf_path'] and d['cte_path']) else "[ ]"
            n_nf = f"{d['nf_path']} ({d['nf_pgs']}pgs)" if d['nf_path'] else "---"
            n_ct = f"{d['cte_path']} ({d['cte_pgs']}pgs)" if d['cte_path'] else "---"
            self.tree.insert("", "end", values=(sel, str(idx), chave, n_nf, n_ct))

        self.update_status("Concluído.")
        self.root.after(100, self.calcular_resumo)

    def iniciar_impressao(self):
        printer = self.combo_printer.get()
        if not printer: return
        lista = []
        for item in self.tree.get_children():
            if self.tree.item(item, "values")[0] == "[X]":
                chave = self.tree.item(item, "values")[2]
                if chave in self.dados_memoria: lista.append((chave, self.dados_memoria[chave]))

        if not lista: return
        if messagebox.askyesno("Imprimir", f"Imprimir {len(lista)} conjuntos?"):
            t = threading.Thread(target=self.thread_print, args=(lista, printer, self.pasta_origem.get()))
            t.start()

    def thread_print(self, lista, printer, base):
        dest = os.path.join(base, "Arquivos Impressos")
        if not os.path.exists(dest): os.makedirs(dest)
        try:
            win32print.SetDefaultPrinter(printer)
        except:
            pass

        for i, (chave, dados) in enumerate(lista, start=1):
            self.update_status(f"Imprimindo {i}/{len(lista)}...")
            writer = PdfWriter()
            mover = []

            if dados['nf_path']:
                p = os.path.join(base, dados['nf_path'])
                writer.append(p);
                mover.append(p)
            if dados['cte_path']:
                p = os.path.join(base, dados['cte_path'])
                writer.append(p);
                mover.append(p)

            if mover:
                tmp = os.path.join(base, f"TEMP_{chave}.pdf")
                with open(tmp, "wb") as f:
                    writer.write(f)
                try:
                    win32api.ShellExecute(0, "print", tmp, None, ".", 0)
                    time.sleep(3)
                except:
                    pass
                try:
                    os.remove(tmp)
                except:
                    pass
                for m in mover:
                    try:
                        shutil.move(m, os.path.join(dest, os.path.basename(m)))
                    except:
                        pass

        self.update_status("Finalizado.")
        messagebox.showinfo("Fim", "Processo Concluído")


if __name__ == "__main__":
    root = tk.Tk()
    app = ImpressorDHL(root)
    root.mainloop()