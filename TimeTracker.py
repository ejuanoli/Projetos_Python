import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import time
import os
from datetime import datetime, timedelta
import ctypes
import getpass

# Ícones PNG: crie a pasta "icons" junto ao TimeTracker.py e coloque:
#   theme_light.png, theme_dark.png (tema claro/escuro), edit.png, delete.png (tabela)
# Ver arquivo: icons/COMO_ADICIONAR_ICONES.txt
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ICONS_DIR = os.path.join(SCRIPT_DIR, "icons")


def load_icon(name, size=None):
    """Carrega PNG da pasta icons. size = (largura, altura) para redimensionar (subsample)."""
    path = os.path.join(ICONS_DIR, name)
    if not os.path.isfile(path):
        return None
    try:
        img = tk.PhotoImage(file=path)
        if size and img.width() > size[0]:
            w, h = size[0], size[1] if len(size) > 1 else size[0]
            img = img.subsample(max(1, img.width() // w), max(1, img.height() // h))
        return img
    except Exception:
        return None

# --- CONFIGURAÇÃO DE TEMAS (sem azul em textos) ---
THEMES = {
    "light": {
        "bg_main": "#E8ECF4",
        "bg_card": "#FFFFFF",
        "text_main": "#1A1D21",
        "text_sub": "#5C6370",
        "accent": "#2D6A4F",
        "accent_hover": "#1B4332",
        "success": "#198754",
        "success_hover": "#157347",
        "danger": "#DC3545",
        "danger_hover": "#BB2D3B",
        "input_bg": "#F0F2F5",
        "input_fg": "#1A1D21",
        "border": "#DEE2E6",
        "tree_head_bg": "#E9ECEF",
        "tree_head_fg": "#495057",
        "row_even": "#FFFFFF",
        "row_odd": "#F8F9FA",
        "cursor": "#1A1D21",
        "btn_neutral": "#E9ECEF",
        "btn_neutral_text": "#495057",
        "disabled_bg": "#CED4DA",
        "disabled_fg": "#6C757D",
        "table_border": "#E0E0E0",
    },
    "dark": {
        "bg_main": "#0D1117",
        "bg_card": "#161B22",
        "text_main": "#E6EDF3",
        "text_sub": "#8B949E",
        "accent": "#3FB950",
        "accent_hover": "#56D364",
        "success": "#3FB950",
        "success_hover": "#56D364",
        "danger": "#F85149",
        "danger_hover": "#FF7B72",
        "input_bg": "#21262D",
        "input_fg": "#E6EDF3",
        "border": "#30363D",
        "tree_head_bg": "#21262D",
        "tree_head_fg": "#8B949E",
        "row_even": "#161B22",
        "row_odd": "#0D1117",
        "cursor": "#E6EDF3",
        "btn_neutral": "#21262D",
        "btn_neutral_text": "#C9D1D9",
        "disabled_bg": "#30363D",
        "disabled_fg": "#6E7681",
        "table_border": "#586069",
    }
}

DB_FILE = "time_tracker_final.db"


# --- UTILITÁRIOS ---

def get_system_user():
    try:
        GetUserNameEx = ctypes.windll.secur32.GetUserNameExW
        NameDisplay = 3
        size = ctypes.pointer(ctypes.c_ulong(0))
        GetUserNameEx(NameDisplay, None, size)
        name_buffer = ctypes.create_unicode_buffer(size.contents.value)
        GetUserNameEx(NameDisplay, name_buffer, size)
        if name_buffer.value: return name_buffer.value
    except:
        pass
    return getpass.getuser().upper()


def center_window(win, w, h):
    win.update_idletasks()
    x = (win.winfo_screenwidth() // 2) - (w // 2)
    y = (win.winfo_screenheight() // 2) - (h // 2)
    win.geometry(f'{w}x{h}+{x}+{y}')


def format_duration(seconds):
    seconds = int(seconds)
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


def init_db():
    conn = sqlite3.connect(DB_FILE)
    conn.cursor().execute('''
        CREATE TABLE IF NOT EXISTS registros (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT, operacao TEXT, tipo TEXT, 
            atividade TEXT, inicio DATETIME, fim DATETIME, duracao TEXT
        )
    ''')
    conn.commit()
    conn.close()


# --- WIDGETS PERSONALIZADOS ---

class RoundedButton(tk.Canvas):
    """Botão com cantos arredondados, hover e proteção contra duplo clique. Suporta texto ou imagem PNG."""

    COOLDOWN_MS = 400  # Tempo mínimo entre cliques

    def __init__(self, parent, width, height, radius, color, text, text_color, command=None, hover_color=None, image=None):
        super().__init__(parent, width=width, height=height, bg=parent['bg'], highlightthickness=0)
        self.command = command
        self.base_color = color
        self.hover_color = hover_color if hover_color else self.adjust_brightness(color, 0.9)
        self.radius = radius
        self.text_color = text_color
        self.text_str = text
        self._state = "normal"
        self._cooldown_job = None
        self._photo = None  # Manter referência ao PhotoImage
        self.text_id = None
        self.img_id = None

        self.rect = self.round_rect(2, 2, width - 2, height - 2, radius, fill=color, outline=color)
        if image is not None:
            self._photo = image
            self.img_id = self.create_image(width / 2, height / 2, image=image)
        else:
            self.text_id = self.create_text(width / 2, height / 2, text=text, fill=text_color,
                                            font=('Segoe UI', 10, 'bold'))

        self.bind("<Button-1>", self.on_click)
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)

    def round_rect(self, x1, y1, x2, y2, radius=25, **kwargs):
        points = [x1 + radius, y1, x1 + radius, y1, x2 - radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius, x2,
                  y1 + radius, x2, y2 - radius, x2, y2 - radius, x2, y2, x2 - radius, y2, x2 - radius, y2, x1 + radius,
                  y2, x1 + radius, y2, x1, y2, x1, y2 - radius, x1, y2 - radius, x1, y1 + radius, x1, y1 + radius, x1,
                  y1]
        return self.create_polygon(points, **kwargs, smooth=True)

    def on_click(self, e):
        if self._state != "normal" or not self.command:
            return
        if self._cooldown_job is not None:
            return  # Ignora cliques durante o cooldown
        try:
            self.command()
        finally:
            self._cooldown_job = self.after(self.COOLDOWN_MS, self._clear_cooldown)

    def _clear_cooldown(self):
        self._cooldown_job = None

    def on_enter(self, e):
        if self._state == "normal": self.itemconfig(self.rect, fill=self.hover_color, outline=self.hover_color)

    def on_leave(self, e):
        if self._state == "normal": self.itemconfig(self.rect, fill=self.base_color, outline=self.base_color)

    def config_color(self, bg, fg):
        self.base_color = bg
        self.text_color = fg
        self.hover_color = self.adjust_brightness(bg, 0.9)
        self.itemconfig(self.rect, fill=bg, outline=bg)
        if self.text_id is not None:
            self.itemconfig(self.text_id, fill=fg)

    def set_image(self, photo):
        """Troca o conteúdo do botão para uma imagem PNG (para tema claro/escuro)."""
        self._photo = photo
        if photo is None:
            if self.img_id is not None:
                self.delete(self.img_id)
                self.img_id = None
            if self.text_id is None:
                self.text_id = self.create_text(self.winfo_reqwidth() / 2, self.winfo_reqheight() / 2,
                                                text=self.text_str, fill=self.text_color,
                                                font=('Segoe UI', 10, 'bold'))
            return
        if self.text_id is not None:
            self.delete(self.text_id)
            self.text_id = None
        if self.img_id is not None:
            self.itemconfig(self.img_id, image=photo)
        else:
            self.img_id = self.create_image(self.winfo_reqwidth() / 2, self.winfo_reqheight() / 2, image=photo)

    def set_state(self, state):
        self._state = state
        if state == "disabled":
            self.itemconfig(self.rect, fill="#CED4DA", outline="#CED4DA")
            if self.text_id is not None:
                self.itemconfig(self.text_id, fill="#6C757D")
        else:
            self.itemconfig(self.rect, fill=self.base_color, outline=self.base_color)
            if self.text_id is not None:
                self.itemconfig(self.text_id, fill=self.text_color)

    def adjust_brightness(self, color, amount):
        # CORREÇÃO: Lida com cores hex curtas (#EEE) e longas (#EEEEEE)
        if not color.startswith("#"): return color

        hex_val = color[1:]
        if len(hex_val) == 3:
            hex_val = "".join([c * 2 for c in hex_val])

        if len(hex_val) != 6: return color

        try:
            r = int(hex_val[0:2], 16)
            g = int(hex_val[2:4], 16)
            b = int(hex_val[4:6], 16)

            r = int(max(0, min(255, r * amount)))
            g = int(max(0, min(255, g * amount)))
            b = int(max(0, min(255, b * amount)))

            return f"#{r:02x}{g:02x}{b:02x}"
        except ValueError:
            return color

    def configure_bg(self, bg):
        # Atualiza o fundo do canvas (fora do botão redondo) para combinar com o container
        self.configure(bg=bg)


class DateTimePicker(tk.Frame):
    def __init__(self, parent, theme, initial_dt=None):
        super().__init__(parent, bg=theme["bg_card"])
        if initial_dt is None:
            dt = datetime.now()
        elif isinstance(initial_dt, str):
            try:
                dt = datetime.strptime(initial_dt.split(".")[0], "%Y-%m-%d %H:%M:%S")
            except (ValueError, TypeError):
                dt = datetime.now()
        else:
            dt = initial_dt

        sb_style = {"width": 3, "font": ('Segoe UI', 11), "relief": "flat", "justify": "center",
                    "bg": theme["input_bg"], "fg": theme["input_fg"], "bd": 0}

        self.lbl_icon = tk.Label(self, text="\u23F0", bg=theme["bg_card"], fg=theme["text_main"], font=('Segoe UI', 12))
        self.lbl_icon.pack(side=tk.LEFT)

        self.d = tk.Spinbox(self, from_=1, to=31, **sb_style);
        self.d.pack(side=tk.LEFT);
        self.d.delete(0, "end");
        self.d.insert(0, f"{dt.day:02d}")
        tk.Label(self, text="/", bg=theme["bg_card"], fg=theme["text_main"], font='bold').pack(side=tk.LEFT)
        self.m = tk.Spinbox(self, from_=1, to=12, **sb_style);
        self.m.pack(side=tk.LEFT);
        self.m.delete(0, "end");
        self.m.insert(0, f"{dt.month:02d}")
        tk.Label(self, text="/", bg=theme["bg_card"], fg=theme["text_main"], font='bold').pack(side=tk.LEFT)
        sb_style_y = {**sb_style, "width": 5}
        self.y = tk.Spinbox(self, from_=2020, to=2030, **sb_style_y);
        self.y.pack(side=tk.LEFT);
        self.y.delete(0, "end");
        self.y.insert(0, dt.year)

        self.lbl_clk = tk.Label(self, text="   \u23F1", bg=theme["bg_card"], fg=theme["text_main"], font=('Segoe UI', 12))
        self.lbl_clk.pack(side=tk.LEFT)
        self.hr = tk.Spinbox(self, from_=0, to=23, **sb_style);
        self.hr.pack(side=tk.LEFT);
        self.hr.delete(0, "end");
        self.hr.insert(0, f"{dt.hour:02d}")
        tk.Label(self, text=":", bg=theme["bg_card"], fg=theme["text_main"], font='bold').pack(side=tk.LEFT)
        self.mn = tk.Spinbox(self, from_=0, to=59, **sb_style);
        self.mn.pack(side=tk.LEFT);
        self.mn.delete(0, "end");
        self.mn.insert(0, f"{dt.minute:02d}")

    def get_dt(self):
        try:
            return datetime(int(self.y.get()), int(self.m.get()), int(self.d.get()), int(self.hr.get()),
                            int(self.mn.get()), 0)
        except:
            return None


# --- JANELA DE EDIÇÃO/CRIAÇÃO (MODAL) ---
class RecordDialog:
    def __init__(self, parent, theme_name, mode="add", record_data=None, username=None, callback=None, on_close_callback=None):
        self.top = tk.Toplevel(parent)
        t = THEMES[theme_name]
        self.top.configure(bg=t["bg_card"])
        self.callback, self.mode, self.rid = callback, mode, (record_data[0] if record_data else None)
        self.parent = parent
        self.on_close_callback = on_close_callback

        self.top.title("Registro Manual" if mode == "add" else "Editar Registro")
        center_window(self.top, 500, 420)
        self.top.transient(parent)
        self.top.grab_set()
        self.top.focus_set()
        self.top.protocol("WM_DELETE_WINDOW", self._on_close)

        op, tp, act = ("", "", "") if mode == "add" else (record_data[2], record_data[3], record_data[4])
        start = (datetime.now() - timedelta(hours=1)) if mode == "add" else record_data[5]
        end = datetime.now() if mode == "add" else record_data[6]

        c = tk.Frame(self.top, bg=t["bg_card"], padx=30, pady=20);
        c.pack(fill=tk.BOTH, expand=True)

        row1 = tk.Frame(c, bg=t["bg_card"]);
        row1.pack(fill=tk.X, pady=5)
        self.cb_op = self.combo(row1, t, "Operação", ["Suporte", "Infraestrutura", "Dev", "Admin"], op)
        self.cb_op.pack(side=tk.LEFT, padx=(0, 20))
        self.cb_tp = self.combo(row1, t, "Tipo", ["Incidente", "Melhoria", "Change", "Manutencao"], tp)
        self.cb_tp.pack(side=tk.LEFT)

        tk.Label(c, text="Atividade", bg=t["bg_card"], fg=t["text_sub"]).pack(anchor='w', pady=(15, 2))
        self.ent = tk.Entry(c, bg=t["input_bg"], fg=t["input_fg"], relief="flat", font=('Segoe UI', 11))
        self.ent.insert(0, act);
        self.ent.pack(fill=tk.X, ipady=5)

        tk.Label(c, text="Início", bg=t["bg_card"], fg=t["text_sub"]).pack(anchor='w', pady=(15, 2))
        self.dt_s = DateTimePicker(c, t, start);
        self.dt_s.pack(anchor='w')
        tk.Label(c, text="Fim", bg=t["bg_card"], fg=t["text_sub"]).pack(anchor='w', pady=(10, 2))
        self.dt_e = DateTimePicker(c, t, end);
        self.dt_e.pack(anchor='w')

        b_fr = tk.Frame(self.top, bg=t["bg_card"]);
        b_fr.pack(side=tk.BOTTOM, fill=tk.X, pady=20, padx=30)
        RoundedButton(b_fr, 120, 35, 10, t["success"], "SALVAR", "white", self.save).pack(side=tk.RIGHT)
        RoundedButton(b_fr, 100, 35, 10, t["btn_neutral"], "CANCELAR", t["btn_neutral_text"], self._on_cancel).pack(
            side=tk.RIGHT, padx=10)

    def combo(self, p, t, l, v, cur):
        f = tk.Frame(p, bg=t["bg_card"])
        tk.Label(f, text=l, bg=t["bg_card"], fg=t["text_sub"]).pack(anchor='w')
        cb = ttk.Combobox(f, values=v, width=18);
        cb.set(cur);
        cb.pack()
        return f

    def _on_close(self):
        try:
            self.top.grab_release()
        except tk.TclError:
            pass
        if self.on_close_callback:
            try:
                self.on_close_callback()
            except Exception:
                pass
        self.top.destroy()

    def _on_cancel(self):
        self._on_close()

    def save(self):
        op, tp, act = self.cb_op.children['!combobox'].get(), self.cb_tp.children[
            '!combobox'].get(), self.ent.get().strip()
        s, e = self.dt_s.get_dt(), self.dt_e.get_dt()
        if not op or not tp or not act:
            messagebox.showwarning("Aviso", "Preencha todos os campos.")
            return
        if not s or not e:
            messagebox.showwarning("Aviso", "Data/hora inválida.")
            return
        if e <= s:
            messagebox.showerror("Erro", "Fim deve ser maior que início.")
            return
        if self.callback:
            self.callback(self.mode, self.rid, op, tp, act, s, e,
                          format_duration((e - s).total_seconds()))
        self._on_close()


# --- APP PRINCIPAL ---
class TimeTrackerApp:
    def __init__(self, root, username):
        self.root = root
        self.username = username
        self.root.title("Time Tracker Enterprise")
        self.root.state('zoomed')

        init_db()
        self.theme_name = "light"
        self.running = False
        self.start_time = None
        self.job = None
        self.filter_mode = "hoje"
        self.view_mode = "detailed"

        self.var_op = tk.StringVar()
        self.var_tp = tk.StringVar()
        self.var_act = tk.StringVar()

        # Ícones PNG (theme_light.png, theme_dark.png, edit.png, delete.png em pasta icons/)
        self._icon_theme_light = load_icon("theme_light.png", (28, 28))
        self._icon_theme_dark = load_icon("theme_dark.png", (28, 28))
        self._icon_edit = load_icon("edit.png", (22, 22))
        self._icon_delete = load_icon("delete.png", (22, 22))

        self.setup_ui()
        self.apply_theme()
        self.refresh_table()

    def setup_ui(self):
        # 1. HEADER
        self.header = tk.Frame(self.root, height=70, padx=30)
        self.header.pack(fill=tk.X);
        self.header.pack_propagate(False)

        self.lbl_title = tk.Label(self.header, text="\u23F1 Time Tracker", font=('Segoe UI', 18, 'bold'))
        self.lbl_title.pack(side=tk.LEFT)

        # Botão tema: usa PNG se existir (em tema light mostra theme_dark.png para ir ao escuro)
        theme_img = self._icon_theme_dark
        self.btn_theme = RoundedButton(
            self.header, 40, 40, 20, "#E9ECEF", "\u263E", "#495057", self.toggle_theme, image=theme_img
        )
        self.btn_theme.pack(side=tk.RIGHT, padx=15)

        self.lbl_user = tk.Label(self.header, text=self.username, font=('Segoe UI', 11))
        self.lbl_user.pack(side=tk.RIGHT)

        # 2. MAIN CONTAINER
        self.main = tk.Frame(self.root, padx=30, pady=20)
        self.main.pack(fill=tk.BOTH, expand=True)

        # 3. TIMER CARD
        self.card = tk.Frame(self.main, padx=20, pady=20)
        self.card.pack(fill=tk.X, pady=(0, 20))

        self.lbl_timer = tk.Label(self.card, text="00:00:00", font=('Consolas', 72, 'bold'))
        self.lbl_timer.pack()

        # Inputs (guardamos referências para limpar focus ao clicar fora)
        self._input_widgets = []
        in_fr = tk.Frame(self.card);
        in_fr.pack(pady=20)
        self.lbls_in = []

        def mk_cb(l, v, var, c):
            lbl = tk.Label(in_fr, text=l, font=('Segoe UI', 10));
            lbl.grid(row=0, column=c, sticky='w', padx=10)
            self.lbls_in.append(lbl)
            cb = ttk.Combobox(in_fr, textvariable=var, values=v, state="readonly", width=20, font=('Segoe UI', 11))
            cb.grid(row=1, column=c, padx=10, ipady=5)
            cb.bind("<<ComboboxSelected>>", self.check_ready)
            self._input_widgets.append(cb)

        mk_cb("Operação", ["Suporte", "Infraestrutura", "Dev", "Admin"], self.var_op, 0)
        mk_cb("Tipo", ["Incidente", "Melhoria", "Change", "Manutencao"], self.var_tp, 1)

        l_act = tk.Label(in_fr, text="Atividade", font=('Segoe UI', 10));
        l_act.grid(row=0, column=2, sticky='w', padx=10)
        self.lbls_in.append(l_act)
        self.ent_act = tk.Entry(in_fr, textvariable=self.var_act, width=50, font=('Segoe UI', 11), relief="flat")
        self.ent_act.grid(row=1, column=2, padx=10, ipady=5)
        self.ent_act.bind("<KeyRelease>", self.check_ready)
        self._input_widgets.append(self.ent_act)

        # Control Buttons
        btns = tk.Frame(self.card);
        btns.pack(pady=10)
        self.btn_start = RoundedButton(btns, 160, 45, 10, "#E9ECEF", "\u25B6  INICIAR", "#6C757D", self.start)
        self.btn_start.pack(side=tk.LEFT, padx=10)
        self.btn_stop = RoundedButton(btns, 140, 45, 10, "#E9ECEF", "\u25A0  PARAR", "#6C757D", self.stop)
        self.btn_stop.pack(side=tk.LEFT, padx=10)
        self.btn_stop.set_state("disabled")

        # 4. TOOLBAR
        self.tool = tk.Frame(self.main);
        self.tool.pack(fill=tk.X, pady=(0, 10))

        # Filtros
        ff = tk.Frame(self.tool);
        ff.pack(side=tk.LEFT)
        self.f_btns = {}
        for k, l in [('hoje', 'Hoje'), ('ontem', 'Ontem'), ('semana', 'Semana'), ('mes', 'Mês'), ('geral', 'Tudo')]:
            b = RoundedButton(ff, 80, 30, 8, "#EEEEEE", l, "#333333", lambda m=k: self.set_filter(m))
            b.pack(side=tk.LEFT, padx=3)
            self.f_btns[k] = b

        # Ações
        af = tk.Frame(self.tool);
        af.pack(side=tk.RIGHT)
        self.btn_view = RoundedButton(af, 140, 35, 8, "#333333", "\u2630 Agrupar", "white", self.toggle_view)
        self.btn_view.pack(side=tk.LEFT, padx=5)
        self.btn_man = RoundedButton(af, 140, 35, 8, "#2D6A4F", "+ Manual", "white", self.manual)
        self.btn_man.pack(side=tk.LEFT, padx=5)
        self.btn_del = RoundedButton(af, 120, 35, 8, "#DC3545", "\u2715 Excluir", "white", self.del_sel)
        # pack condicional do btn_del

        # 5. TABELA (borda visível no detalhado e no agrupado, tema claro e escuro)
        self.table_border = tk.Frame(self.main, bg="#E8E8E8")
        self.table_border.pack(fill=tk.BOTH, expand=True)

        self.table_wrap = tk.Frame(self.table_border, bg="white")
        self.table_wrap.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        self.tree_header_frame = tk.Frame(self.table_wrap)
        self.tree_header_frame.pack(fill=tk.X)
        self._header_labels = []
        self._build_table_header()

        self.tree_separator = ttk.Separator(self.table_wrap, orient="horizontal")
        self.tree_separator.pack(fill=tk.X, pady=0)

        self.tree_container = tk.Frame(self.table_wrap, bg="white")
        self.tree_container.pack(fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(self.tree_container, show='headings', selectmode="extended", height=18)
        sb = ttk.Scrollbar(self.tree_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar = sb

        self.tree.bind("<Button-1>", self.on_click)
        self.tree.bind("<Double-1>", lambda e: "break")
        self.tree.bind("<<TreeviewSelect>>", self.on_sel)
        self._input_widgets.append(self.tree)
        self._table_action_lock = False

        # Clique fora de combos/entrada/tabela remove o foco (efeito de seleção some)
        self.root.bind("<Button-1>", self._on_click_clear_focus)

        # 6. FOOTER
        self.bar_total = tk.Frame(self.root, height=40);
        self.bar_total.pack(fill=tk.X, side=tk.BOTTOM)
        self.lbl_total = tk.Label(self.bar_total, text="TEMPO TOTAL: 00:00:00", font=('Segoe UI', 11, 'bold'))
        self.lbl_total.pack(pady=8)

    def _on_click_clear_focus(self, event):
        """Ao clicar fora de combos/entrada/tabela, tira o foco para o efeito de seleção sumir."""
        w = event.widget
        while w:
            if w in self._input_widgets:
                return
            try:
                w = w.master
            except (AttributeError, tk.TclError):
                break
        self.root.focus_set()

    # --- THEMING CORE ---
    def toggle_theme(self):
        self.theme_name = "dark" if self.theme_name == "light" else "light"
        self.apply_theme()

    def apply_theme(self):
        t = THEMES[self.theme_name]

        # Backgrounds Globais
        for w in [self.root, self.main, self.tool]: w.configure(bg=t["bg_main"])
        for w in [self.header, self.card, self.bar_total]: w.configure(bg=t["bg_card"])
        for f in [self.tool.winfo_children()[0], self.tool.winfo_children()[1],
                  self.card.winfo_children()[1], self.card.winfo_children()[2]]:
            f.configure(bg=t["bg_card"] if f.master == self.card else t["bg_main"])

        # Textos e Cores (sem azul: título e total usam text_main)
        self.lbl_title.configure(bg=t["bg_card"], fg=t["text_main"])
        self.lbl_user.configure(bg=t["bg_card"], fg=t["text_sub"])
        self.lbl_timer.configure(bg=t["bg_card"], fg=t["text_main"])
        self.lbl_total.configure(bg=t["bg_card"], fg=t["text_main"])

        for l in self.lbls_in: l.configure(bg=t["bg_card"], fg=t["text_sub"])
        self.ent_act.configure(bg=t["input_bg"], fg=t["input_fg"], insertbackground=t["cursor"])

        # Botão tema: ícone PNG ou fallback em texto
        self.btn_theme.config_color(t["input_bg"], t["text_main"])
        self.btn_theme.text_str = "\u2600" if self.theme_name == "dark" else "\u263E"
        if self.theme_name == "dark" and self._icon_theme_light:
            self.btn_theme.set_image(self._icon_theme_light)
        elif self.theme_name == "light" and self._icon_theme_dark:
            self.btn_theme.set_image(self._icon_theme_dark)
        else:
            self.btn_theme.set_image(None)
            if self.btn_theme.text_id is not None:
                self.btn_theme.itemconfig(self.btn_theme.text_id, text=self.btn_theme.text_str)
        self.btn_theme.configure_bg(t["bg_card"])

        self.btn_view.config_color(t["text_main"], t["bg_card"])
        self.btn_view.configure_bg(t["bg_main"])

        self.btn_man.config_color(t["accent"], "white")
        self.btn_man.configure_bg(t["bg_main"])

        self.btn_del.config_color(t["danger"], "white")
        self.btn_del.configure_bg(t["bg_main"])

        # Start/Stop Logic Colors
        if not self.running:
            self.btn_start.set_state("disabled")
            self.check_ready()
        else:
            self.btn_start.set_state("disabled")
            self.btn_stop.config_color(t["danger"], "white")

        # Importante: Atualizar fundo do canvas dos botões
        self.btn_start.configure_bg(t["bg_card"])
        self.btn_stop.configure_bg(t["bg_card"])

        # Filter Buttons
        self.update_filter_btns(t)

        # Style Treeview: cabeçalhos sempre visíveis, bordas e linhas
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview", background=t["bg_card"], foreground=t["text_main"],
                        fieldbackground=t["bg_card"], rowheight=36, font=('Segoe UI', 10))
        if self.view_mode == "detailed":
            style.configure("Treeview.Heading", background=t["tree_head_bg"], foreground=t["tree_head_bg"],
                            font=('Segoe UI', 1), relief="flat", padding=(0, 0))
        else:
            style.configure("Treeview.Heading", background=t["tree_head_bg"], foreground=t["tree_head_fg"],
                            font=('Segoe UI', 10, 'bold'), relief="flat", padding=(8, 6))
        style.map("Treeview", background=[('selected', t["accent"])], foreground=[('selected', 'white')])
        style.map("Treeview.Heading", background=[('active', t["tree_head_bg"])])

        # Borda da tabela (mesmo efeito no detalhado e no agrupado, claro e escuro)
        border_color = t["table_border"]
        self.table_border.configure(bg=border_color)
        self.table_wrap.configure(bg=t["bg_card"])
        self.tree_container.configure(bg=t["bg_card"])
        self.tree_header_frame.configure(bg=t["tree_head_bg"])
        for lbl in self._header_labels:
            lbl.configure(bg=t["tree_head_bg"], fg=t["tree_head_fg"])
        style.configure("TSeparator", background=border_color)

        self.refresh_table()  # Re-aplica tags zebradas

    def _build_table_header(self):
        """Cabeçalho da tabela: 9 colunas centralizadas; últimas duas usam PNG (edit/delete) se existirem."""
        widths = [36, 130, 100, 100, 340, 60, 80, 44, 44]
        texts = ["", "DATA / HORA", "OPERA\u00C7\u00C3O", "TIPO", "ATIVIDADE", "FIM", "DURA\u00C7\u00C3O", "", ""]
        for col, (w, txt) in enumerate(zip(widths, texts)):
            self.tree_header_frame.grid_columnconfigure(col, minsize=w, weight=(1 if col == 4 else 0))
            lbl = tk.Label(
                self.tree_header_frame, text=txt, font=('Segoe UI', 10, 'bold'),
                bg="#E9ECEF", fg="#495057", padx=6, pady=6, anchor="center",
                relief="solid", bd=1, highlightthickness=0
            )
            if col == 7 and self._icon_edit:
                lbl.configure(image=self._icon_edit, text="")
            elif col == 8 and self._icon_delete:
                lbl.configure(image=self._icon_delete, text="")
            elif col in (7, 8):
                lbl.configure(text="\u270E" if col == 7 else "\u2715")
            lbl.grid(row=0, column=col, sticky="nsew")
            self._header_labels.append(lbl)

    def update_filter_btns(self, t):
        for k, b in self.f_btns.items():
            b.configure_bg(t["bg_main"])
            if k == self.filter_mode:
                b.config_color(t["accent"], "white")
            else:
                b.config_color(t["btn_neutral"], t["btn_neutral_text"])

    # --- LÓGICA ---
    def check_ready(self, e=None):
        t = THEMES[self.theme_name]
        if self.running: return
        if self.var_op.get() and self.var_tp.get() and self.var_act.get().strip():
            self.btn_start.set_state("normal")
            self.btn_start.config_color(t["success"], "white")
        else:
            self.btn_start.set_state("disabled")

    def start(self):
        t = THEMES[self.theme_name]
        self.running = True
        self.start_time = datetime.now()
        self.btn_start.set_state("disabled")
        self.btn_stop.set_state("normal")
        self.btn_stop.config_color(t["danger"], "white")
        self.ent_act.config(state='disabled')
        self.tick()

    def tick(self):
        if self.running:
            d = datetime.now() - self.start_time
            self.lbl_timer.config(text=format_duration(d.total_seconds()), fg=THEMES[self.theme_name]["text_main"])
            self.job = self.root.after(1000, self.tick)

    def stop(self):
        if not self.running: return
        self.running = False
        if self.job: self.root.after_cancel(self.job)

        end = datetime.now()
        dur = format_duration((end - self.start_time).total_seconds())
        # Converter para str
        s_str = self.start_time.strftime("%Y-%m-%d %H:%M:%S")
        e_str = end.strftime("%Y-%m-%d %H:%M:%S")

        self.db_save("add", None, self.var_op.get(), self.var_tp.get(), self.var_act.get(), s_str, e_str, dur)
        self.reset_ui()

    def reset_ui(self):
        t = THEMES[self.theme_name]
        self.lbl_timer.config(text="00:00:00", fg=t["text_main"])
        self.ent_act.config(state='normal');
        self.var_act.set("")
        self.btn_stop.set_state("disabled")
        self.check_ready()

    def db_save(self, mode, rid, op, tp, act, s, e, dur):
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        s_str = s.strftime("%Y-%m-%d %H:%M:%S") if isinstance(s, datetime) else s
        e_str = e.strftime("%Y-%m-%d %H:%M:%S") if isinstance(e, datetime) else e

        if mode == "add":
            c.execute(
                "INSERT INTO registros (usuario, operacao, tipo, atividade, inicio, fim, duracao) VALUES (?,?,?,?,?,?,?)",
                (self.username, op, tp, act, s_str, e_str, dur))
        else:
            c.execute("UPDATE registros SET operacao=?, tipo=?, atividade=?, inicio=?, fim=?, duracao=? WHERE id=?",
                      (op, tp, act, s_str, e_str, dur, rid))
        conn.commit();
        conn.close()
        self.refresh_table()

    def set_filter(self, m):
        self.filter_mode = m
        self.update_filter_btns(THEMES[self.theme_name])
        self.refresh_table()

    def toggle_view(self):
        self.view_mode = "grouped" if self.view_mode == "detailed" else "detailed"
        self.btn_view.text_str = "\u270E Detalhada" if self.view_mode == "grouped" else "\u2630 Agrupar"
        self.btn_view.itemconfig(self.btn_view.text_id, text=self.btn_view.text_str)
        self.apply_theme()  # Reaplica estilo do cabeçalho (visível no agrupado, mínimo no detalhado)

    def refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        self.btn_del.pack_forget()

        t = THEMES[self.theme_name]
        conn = sqlite3.connect(DB_FILE)

        wh = "WHERE usuario = ?"
        if self.filter_mode == 'hoje':
            wh += " AND date(inicio, 'localtime') == date('now', 'localtime')"
        elif self.filter_mode == 'ontem':
            wh += " AND date(inicio, 'localtime') == date('now', '-1 day', 'localtime')"
        elif self.filter_mode == 'semana':
            wh += " AND date(inicio, 'localtime') >= date('now', '-6 days', 'localtime')"
        elif self.filter_mode == 'mes':
            wh += " AND strftime('%Y-%m', inicio, 'localtime') == strftime('%Y-%m', 'now', 'localtime')"

        total_secs = 0

        # Configure Tags para Zebrado
        self.tree.tag_configure('odd', background=t["row_odd"], foreground=t["text_main"])
        self.tree.tag_configure('even', background=t["row_even"], foreground=t["text_main"])

        if self.view_mode == "detailed":
            cols = ("id", "data", "op", "tipo", "ativ", "fim", "dur", "edit", "del")
            self.tree.config(columns=cols)
            hdrs = [
                ("id", 36), ("data", 130), ("op", 100), ("tipo", 100), ("ativ", 340), ("fim", 60), ("dur", 80),
                ("edit", 44), ("del", 44)
            ]
            for c, w in hdrs:
                self.tree.column(c, width=w, stretch=(c == "ativ"), anchor="center")
            for hc, ht in [("id", ""), ("data", ""), ("op", ""), ("tipo", ""), ("ativ", ""), ("fim", ""), ("dur", ""), ("edit", ""), ("del", "")]:
                self.tree.heading(hc, text=ht)
            # Reempacotar na ordem correta: cabeçalho em cima, depois separador, depois tree
            self.tree_container.pack_forget()
            self.tree_header_frame.pack(side=tk.TOP, fill=tk.X)
            self.tree_separator.pack(side=tk.TOP, fill=tk.X, pady=0)
            self.tree_container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

            rows = conn.cursor().execute(f"SELECT * FROM registros {wh} ORDER BY inicio DESC",
                                         (self.username,)).fetchall()
            for i, r in enumerate(rows):
                try:
                    dt_s = datetime.strptime(str(r[5]).split('.')[0], "%Y-%m-%d %H:%M:%S")
                    dt_e = datetime.strptime(str(r[6]).split('.')[0], "%Y-%m-%d %H:%M:%S")
                    d_col = f"{dt_s.strftime('%d/%m')} - {dt_s.strftime('%H:%M')}"
                    h_fim = dt_e.strftime("%H:%M")
                    secs = (dt_e - dt_s).total_seconds()
                    total_secs += secs
                    dur_str = format_duration(secs)
                except:
                    d_col, h_fim, dur_str = "Err", "-", "-"

                tag = 'even' if i % 2 == 0 else 'odd'
                self.tree.insert("", "end", values=(r[0], d_col, r[2], r[3], r[4], h_fim, dur_str, "\u270E", "\u2715"),
                                 tags=(tag,))
        else:
            self.tree_header_frame.pack_forget()
            self.tree_separator.pack_forget()
            # Container já está empacotado; ordem permanece correta
            cols = ("data", "op", "tipo", "ativ", "total")
            self.tree.config(columns=cols)
            self.tree.heading("data", text="DATA")
            self.tree.column("data", width=100, anchor="center")
            self.tree.heading("op", text="OPERA\u00C7\u00C3O")
            self.tree.column("op", width=120, anchor="center")
            self.tree.heading("tipo", text="TIPO")
            self.tree.column("tipo", width=120, anchor="center")
            self.tree.heading("ativ", text="ATIVIDADE")
            self.tree.column("ativ", width=450, anchor="w")
            self.tree.heading("total", text="TOTAL")
            self.tree.column("total", width=120, anchor="center")

            sql = f"""SELECT strftime('%d/%m', inicio, 'localtime'), operacao, tipo, atividade, 
                      SUM(strftime('%s', fim) - strftime('%s', inicio)) 
                      FROM registros {wh} GROUP BY 1, 2, 3, 4 ORDER BY inicio DESC"""
            rows = conn.cursor().execute(sql, (self.username,)).fetchall()
            for i, r in enumerate(rows):
                s = r[4] if r[4] else 0
                total_secs += s
                tag = 'even' if i % 2 == 0 else 'odd'
                self.tree.insert("", "end", values=(r[0], r[1], r[2], r[3], format_duration(s)), tags=(tag,))

        conn.close()
        self.lbl_total.config(text=f"TEMPO TOTAL: {format_duration(total_secs)}")
        if len(self.tree.get_children()) > 10:
            self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        else:
            self.scrollbar.pack_forget()

    def manual(self):
        RecordDialog(self.root, self.theme_name, "add", username=self.username, callback=self.db_save)

    def on_click(self, e):
        if self._table_action_lock:
            return "break"
        reg = self.tree.identify("region", e.x, e.y)
        if reg == "cell" and self.view_mode == "detailed":
            rid = self.tree.identify_row(e.y)
            col = self.tree.identify_column(e.x)

            if col in ['#8', '#9']:  # Editar / Excluir
                self._table_action_lock = True
                try:
                    vals = self.tree.item(rid, "values")
                    db_id = vals[0]
                    if col == '#8':  # Editar
                        conn = sqlite3.connect(DB_FILE)
                        d = conn.cursor().execute("SELECT * FROM registros WHERE id=?", (db_id,)).fetchone()
                        conn.close()
                        if d:
                            RecordDialog(
                                self.root, self.theme_name, "edit", d, callback=self.db_save,
                                on_close_callback=lambda: setattr(self, "_table_action_lock", False)
                            )
                    elif col == '#9':  # Excluir
                        if messagebox.askyesno("Excluir", "Apagar este registro?"):
                            conn = sqlite3.connect(DB_FILE)
                            conn.cursor().execute("DELETE FROM registros WHERE id=?", (db_id,))
                            conn.commit()
                            conn.close()
                            self.refresh_table()
                finally:
                    if col == '#9':
                        self._table_action_lock = False
                return "break"

            if rid in self.tree.selection():
                self.tree.selection_remove(rid)
                return "break"

    def on_sel(self, e):
        if self.view_mode == "detailed" and self.tree.selection():
            self.btn_del.pack(side=tk.LEFT, padx=5)
        else:
            self.btn_del.pack_forget()

    def del_sel(self):
        s = self.tree.selection()
        if not s or not messagebox.askyesno("Confirmar", f"Apagar {len(s)} itens?"): return
        ids = [self.tree.item(i, "values")[0] for i in s]
        conn = sqlite3.connect(DB_FILE)
        conn.cursor().executemany("DELETE FROM registros WHERE id=?", [(x,) for x in ids])
        conn.commit();
        conn.close()
        self.refresh_table()


if __name__ == "__main__":
    user = get_system_user()
    root = tk.Tk()
    app = TimeTrackerApp(root, user)
    root.mainloop()