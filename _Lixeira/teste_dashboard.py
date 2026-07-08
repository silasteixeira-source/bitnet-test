#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TESTE v3 — Dashboard EACE com Filtros Dinâmicos
Suporte completo para a estrutura de colunas lado a lado agrupadas por regional.
Enriquece com o nome da escola e município usando o eace_cache.json local.
"""

import sys
import os
import re
import json
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import pandas as pd
import threading
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIGURAÇÕES GLOBAIS ---
DASHBOARD_SPREADSHEET_ID = "1hoK7MT4ST6dQe5WsA4rfEG3Tk7NTYAovrL6K9cXlxWY"
DASHBOARD_CACHE_FILE = "eace_dashboard_cache.json"

THEMES = {
    "dark": {
        "bg": "#0f172a",
        "card_bg": "#1e293b",
        "fg": "#f8fafc",
        "fg_muted": "#94a3b8",
        "primary": "#6366f1",
        "primary_hover": "#4f46e5",
        "success": "#10b981",
        "info": "#0ea5e9",
        "warning": "#f59e0b",
        "danger": "#ef4444",
        "border": "#334155",
        "entry_bg": "#0f172a",
        "entry_fg": "#f8fafc",
        "tree_bg": "#1e293b",
        "tree_fg": "#f8fafc",
        "tree_selected": "#6366f1"
    }
}

class ThemeManager:
    current_theme = "dark"

    @classmethod
    def get_colors(cls):
        return THEMES[cls.current_theme]

    @classmethod
    def apply_theme_to_all(cls, root):
        colors = cls.get_colors()
        style = ttk.Style(root)
        style.theme_use("clam")

        style.configure("TNotebook", background=colors["bg"], borderwidth=0)
        style.configure("TNotebook.Tab", background=colors["card_bg"], foreground=colors["fg"], borderwidth=1, padding=8, font=("Segoe UI", 9, "bold"))
        style.map("TNotebook.Tab", background=[("selected", colors["primary"])], foreground=[("selected", "#ffffff")])

        style.configure("TFrame", background=colors["bg"])
        style.configure("TLabel", background=colors["bg"], foreground=colors["fg"], font=("Segoe UI", 10))
        style.configure("TButton", background=colors["card_bg"], foreground=colors["fg"], bordercolor=colors["border"], font=("Segoe UI", 9, "bold"), padding=6)
        style.map("TButton", background=[("active", colors["primary"])], foreground=[("active", "#ffffff")])

        style.configure("TEntry", fieldbackground=colors["entry_bg"], foreground=colors["entry_fg"], insertcolor=colors["fg"], bordercolor=colors["border"])
        style.configure("TCheckbutton", background=colors["bg"], foreground=colors["fg"])

        style.configure("Treeview", background=colors["tree_bg"], foreground=colors["tree_fg"], fieldbackground=colors["tree_bg"], rowheight=24, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", background=colors["card_bg"], foreground=colors["fg"], font=("Segoe UI", 9, "bold"), borderwidth=1)
        style.map("Treeview", background=[("selected", colors["tree_selected"])], foreground=[("selected", "#ffffff")])

        cls._apply_to_widget_recursive(root, colors)

    @classmethod
    def _apply_to_widget_recursive(cls, widget, colors):
        widget_type = widget.winfo_class()
        try:
            if widget_type == "Frame" or widget_type == "TFrame":
                widget.configure(bg=colors["bg"])
            elif widget_type == "Label" or widget_type == "TLabel":
                current_fg = widget.cget("fg") if hasattr(widget, "cget") else ""
                if current_fg not in ["red", "green", "blue", "orange", "#27ae60", "#e74c3c", "#17a2b8", "#2ecc71"]:
                    widget.configure(bg=colors["bg"], fg=colors["fg"])
                else:
                    widget.configure(bg=colors["bg"])
            elif widget_type == "Button" or widget_type == "TButton":
                current_bg = widget.cget("bg") if hasattr(widget, "cget") else ""
                if current_bg not in ["#6c757d", "#007bff", "#17a2b8", "#28a745", "#e67e22", "#3498db", "#e74c3c", "#27ae60", "#2ecc71", "#dc3545"]:
                    widget.configure(bg=colors["card_bg"], fg=colors["fg"], activebackground=colors["primary"], activeforeground="#ffffff", bd=1, relief="groove")
            elif widget_type == "Entry" or widget_type == "TEntry":
                widget.configure(bg=colors["entry_bg"], fg=colors["entry_fg"], insertbackground=colors["fg"], bd=1, relief="solid")
            elif widget_type == "Text":
                widget.configure(bg=colors["entry_bg"], fg=colors["entry_fg"], insertbackground=colors["fg"], bd=1, relief="solid")
            elif widget_type == "Listbox":
                widget.configure(bg=colors["entry_bg"], fg=colors["entry_fg"], selectbackground=colors["primary"], selectforeground="#ffffff", bd=1, relief="solid")
            elif widget_type == "LabelFrame" or widget_type == "TLabelframe":
                widget.configure(bg=colors["bg"], fg=colors["fg"], bd=1, relief="groove")
        except Exception:
            pass
        for child in widget.winfo_children():
            cls._apply_to_widget_recursive(child, colors)

def get_base_path():
    return os.path.dirname(os.path.abspath(__file__))

def get_dashboard_cache_path():
    return os.path.join(get_base_path(), DASHBOARD_CACHE_FILE)

def load_cache(cache_path=None):
    if cache_path is None:
        cache_path = os.path.join(get_base_path(), "eace_cache.json")
    if os.path.exists(cache_path):
        try:
            return pd.read_json(cache_path, dtype=str)
        except Exception:
            pass
    return pd.DataFrame()

def save_cache(df, cache_path):
    try:
        df.to_json(cache_path, force_ascii=False, indent=2)
    except Exception as e:
        print(f"Falha ao salvar cache: {e}")

def authenticate_google_sheets():
    try:
        creds_path = os.path.join(get_base_path(), "credentials.json")
        if not os.path.exists(creds_path):
            return None
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
        return gspread.authorize(creds)
    except Exception as e:
        print(f"Erro na autenticação: {e}")
        return None

def show_detailed_error(title, message, error_obj):
    messagebox.showerror(title, f"{message}\n\nDetalhes do erro: {error_obj}")

# --- CLASSE DASHBOARD ---
class DashboardEACE:
    def __init__(self, parent_frame, main_app_callback, gspread_client=None):
        self.frame = tk.Frame(parent_frame)
        self.frame.pack(fill=tk.BOTH, expand=True)
        self.main_app_callback = main_app_callback
        self.gspread_client = gspread_client
        self.root = parent_frame.winfo_toplevel()
        
        self.df = None
        
        self.colors = ThemeManager.get_colors()
        self._build()
        
        self.frame.after(100, self.carregar_so_local)
        ThemeManager.apply_theme_to_all(self.frame)
        
    def _build(self):
        colors = self.colors
        bg = colors["bg"]
        card = colors["card_bg"]
        border = colors["border"]
        fg = colors["fg"]
        muted = colors["fg_muted"]
        accent = colors["primary"]
        entry_bg = colors["entry_bg"]
        entry_fg = colors["entry_fg"]
        
        self.frame.configure(bg=bg)
        
        # ── TOP BAR ──────────────────────────────────────────
        top = tk.Frame(self.frame, bg=card, height=48)
        top.pack(fill="x")
        top.pack_propagate(False)
        
        tk.Label(top, text="  📊  DASHBOARD DE MONITORAMENTO — EACE (Validação)",
                 font=("Segoe UI", 12, "bold"), bg=card, fg="blue").pack(side="left", padx=8)
                 
        tk.Button(top, text="FECHAR TESTE", command=self.main_app_callback,
                  bg="#6c757d", fg="white", font=("Segoe UI", 8, "bold"),
                  bd=0, cursor="hand2", padx=10).pack(side="right", padx=10, pady=8)
                  
        # ── BARRA DE STATUS E ATUALIZAÇÃO ───────────────────
        bar = tk.Frame(self.frame, bg=bg, pady=10)
        bar.pack(fill="x", padx=16)
        
        self.lbl_status = tk.Label(bar, text="Verificando dados...", font=("Segoe UI", 9, "italic"), bg=bg, fg="orange")
        self.lbl_status.pack(side="left")
        
        tk.Button(bar, text="🔄 Sincronizar Nuvem (Sheets)", command=self.forcar_atualizacao,
                  bg="#17a2b8", fg="white", font=("Segoe UI", 9, "bold"),
                  bd=0, cursor="hand2", padx=12, pady=4).pack(side="right")
                  
        # Separador
        tk.Frame(self.frame, bg=border, height=1).pack(fill="x", pady=2)
        
        # Conteúdo Principal
        body = tk.Frame(self.frame, bg=bg)
        body.pack(fill="both", expand=True, padx=16, pady=10)
        
        # ── CARDS DE MÉTRICAS (Grid) ──────────────────────────
        self.cards_frame = tk.Frame(body, bg=bg)
        self.cards_frame.pack(fill="x", pady=(0, 10))
        self.cards_frame.columnconfigure((0, 1, 2, 3), weight=1, uniform="equal")

        self.cards_data = [
            ("Total de Escolas", "—", "lbl_card_total", "#3498db"),
            ("Instaladas / Concluídas", "—", "lbl_card_concluidas", "green"),
            ("Em Progresso", "—", "lbl_card_progresso", "orange"),
            ("Taxa de Conclusão", "—", "lbl_card_taxa", "#9b59b6")
        ]

        self.card_labels = {}
        for idx, (title, value, name, color) in enumerate(self.cards_data):
            f = tk.Frame(self.cards_frame, bg=card, bd=1, relief="solid", highlightthickness=0)
            f.grid(row=0, column=idx, padx=5, sticky="nsew")
            
            dec = tk.Frame(f, bg=color, height=3)
            dec.pack(fill="x")
            
            tk.Label(f, text=title.upper(), font=("Segoe UI", 8, "bold"), bg=card, fg=muted).pack(pady=(8, 2))
            lbl = tk.Label(f, text=value, font=("Segoe UI", 16, "bold"), bg=card, fg=color)
            lbl.pack(pady=(0, 8))
            self.card_labels[name] = lbl
            
        # ── FILTROS (Grid de 5 colunas) ──────────────────────
        filters_frame = tk.Frame(body, bg=card, bd=1, relief="solid", highlightthickness=0)
        filters_frame.pack(fill="x", pady=5, ipady=5)
        
        for i in range(7):
            filters_frame.columnconfigure(i, weight=1)
            
        # Regional
        f_reg = tk.Frame(filters_frame, bg=card)
        f_reg.grid(row=0, column=0, padx=8, pady=5, sticky="ew")
        tk.Label(f_reg, text="REGIONAL:", font=("Segoe UI", 8, "bold"), bg=card, fg=muted).pack(anchor="w")
        self.cb_reg = ttk.Combobox(f_reg, state="readonly", font=("Segoe UI", 9))
        self.cb_reg.pack(fill="x", pady=(2, 0))
        self.cb_reg.bind("<<ComboboxSelected>>", self.on_reg_changed)
            
        # UF
        f_uf = tk.Frame(filters_frame, bg=card)
        f_uf.grid(row=0, column=1, padx=8, pady=5, sticky="ew")
        tk.Label(f_uf, text="UF:", font=("Segoe UI", 8, "bold"), bg=card, fg=muted).pack(anchor="w")
        self.cb_uf = ttk.Combobox(f_uf, state="readonly", font=("Segoe UI", 9))
        self.cb_uf.pack(fill="x", pady=(2, 0))
        self.cb_uf.bind("<<ComboboxSelected>>", self.on_uf_changed)
        
        # Município
        f_mun = tk.Frame(filters_frame, bg=card)
        f_mun.grid(row=0, column=2, padx=8, pady=5, sticky="ew")
        tk.Label(f_mun, text="MUNICÍPIO:", font=("Segoe UI", 8, "bold"), bg=card, fg=muted).pack(anchor="w")
        self.cb_mun = ttk.Combobox(f_mun, state="readonly", font=("Segoe UI", 9))
        self.cb_mun.pack(fill="x", pady=(2, 0))
        self.cb_mun.bind("<<ComboboxSelected>>", self.on_filter_changed)
        
        # Status
        f_status = tk.Frame(filters_frame, bg=card)
        f_status.grid(row=0, column=3, padx=8, pady=5, sticky="ew")
        tk.Label(f_status, text="STATUS:", font=("Segoe UI", 8, "bold"), bg=card, fg=muted).pack(anchor="w")
        self.cb_status = ttk.Combobox(f_status, state="readonly", font=("Segoe UI", 9))
        self.cb_status.pack(fill="x", pady=(2, 0))
        self.cb_status.bind("<<ComboboxSelected>>", self.on_filter_changed)
        
        # Busca texto
        f_busca = tk.Frame(filters_frame, bg=card)
        f_busca.grid(row=0, column=4, padx=8, pady=5, sticky="ew")
        tk.Label(f_busca, text="BUSCA (INEP/NOME):", font=("Segoe UI", 8, "bold"), bg=card, fg=muted).pack(anchor="w")
        self.ent_search = tk.Entry(f_busca, bg=entry_bg, fg=entry_fg, insertbackground=fg, relief="solid", bd=1, font=("Segoe UI", 10))
        self.ent_search.pack(fill="x", ipady=2, pady=(2, 0))
        self.ent_search.bind("<KeyRelease>", self.on_filter_changed)
        
        # Ordem de Data
        f_ordem = tk.Frame(filters_frame, bg=card)
        f_ordem.grid(row=0, column=5, padx=8, pady=5, sticky="ew")
        tk.Label(f_ordem, text="ORDENAR DATA:", font=("Segoe UI", 8, "bold"), bg=card, fg=muted).pack(anchor="w")
        self.cb_ordem_data = ttk.Combobox(f_ordem, state="readonly", font=("Segoe UI", 9))
        self.cb_ordem_data["values"] = ["Padrão", "Mais Antigas Primeiro", "Mais Recentes Primeiro"]
        self.cb_ordem_data.set("Padrão")
        self.cb_ordem_data.pack(fill="x", pady=(2, 0))
        self.cb_ordem_data.bind("<<ComboboxSelected>>", self.on_filter_changed)

        # Mês do Filtro
        f_mes = tk.Frame(filters_frame, bg=card)
        f_mes.grid(row=0, column=6, padx=8, pady=5, sticky="ew")
        tk.Label(f_mes, text="MÊS/ANO:", font=("Segoe UI", 8, "bold"), bg=card, fg=muted).pack(anchor="w")
        self.cb_mes_data = ttk.Combobox(f_mes, state="readonly", font=("Segoe UI", 9))
        self.cb_mes_data["values"] = ["Todos"]
        self.cb_mes_data.set("Todos")
        self.cb_mes_data.pack(fill="x", pady=(2, 0))
        self.cb_mes_data.bind("<<ComboboxSelected>>", self.on_filter_changed)
        
        # ── TABELA (Treeview) ─────────────────────────────────
        table_frame = tk.Frame(body, bg=bg)
        table_frame.pack(fill="both", expand=True, pady=5)
        
        sc_y = ttk.Scrollbar(table_frame)
        sc_y.pack(side="right", fill="y")
        
        sc_x = ttk.Scrollbar(table_frame, orient="horizontal")
        sc_x.pack(side="bottom", fill="x")

        self.tree = ttk.Treeview(table_frame, columns=("inep", "escola", "municipio", "uf", "regional", "status", "data_inst", "observacao"),
                                 show="headings", yscrollcommand=sc_y.set, xscrollcommand=sc_x.set)
        
        self.tree.heading("inep", text="Código INEP")
        self.tree.heading("escola", text="Nome da Escola")
        self.tree.heading("municipio", text="Município")
        self.tree.heading("uf", text="UF")
        self.tree.heading("regional", text="Regional")
        self.tree.heading("status", text="Status")
        self.tree.heading("data_inst", text="Instalação")
        self.tree.heading("observacao", text="Observação")
        
        self.tree.column("inep", width=90, anchor="center")
        self.tree.column("escola", width=230, anchor="w")
        self.tree.column("municipio", width=140, anchor="w")
        self.tree.column("uf", width=50, anchor="center")
        self.tree.column("regional", width=90, anchor="center")
        self.tree.column("status", width=90, anchor="center")
        self.tree.column("data_inst", width=90, anchor="center")
        self.tree.column("observacao", width=180, anchor="w")
        
        self.tree.pack(fill="both", expand=True)
        sc_y.config(command=self.tree.yview)
        sc_x.config(command=self.tree.xview)
        
        # ── BOTÕES DE AÇÕES ───────────────────────────────────
        actions_frame = tk.Frame(body, bg=bg)
        actions_frame.pack(fill="x", pady=(5, 0))
        
        tk.Button(actions_frame, text="📋 Copiar Registro Selecionado", command=self.copiar_registro,
                  bg="#28a745", fg="white", font=("Segoe UI", 9, "bold"), bd=0, cursor="hand2", padx=12, pady=6).pack(side="left", padx=(0, 10))
                  
        tk.Button(actions_frame, text="📊 Copiar Tabela Filtrada (Excel)", command=self.copiar_tabela,
                  bg="#17a2b8", fg="white", font=("Segoe UI", 9, "bold"), bd=0, cursor="hand2", padx=12, pady=6).pack(side="left")

    def consolidar_matriz_raw(self, raw_data):
        if len(raw_data) < 2:
            return pd.DataFrame()
            
        reg_line = raw_data[0]
        header_line = raw_data[1]
        rows = raw_data[2:]
        
        regionais_map = []
        current_reg = "Desconhecido"
        for item in reg_line:
            if item.strip():
                current_reg = item.strip()
            regionais_map.append(current_reg)
            
        while len(regionais_map) < len(header_line):
            regionais_map.append(current_reg)
            
        records = []
        num_cols = len(header_line)
        
        for group_idx in range(0, num_cols, 5):
            if group_idx + 2 >= num_cols:
                break
                
            col_inep_name = header_line[group_idx].strip()
            if not col_inep_name or "inep" not in col_inep_name.lower():
                continue
                
            uf = col_inep_name.split()[-1] if len(col_inep_name.split()) > 1 else "N/A"
            regional = regionais_map[group_idx]
            regional_curta = regional.split()[0] if regional else "Desconhecido"
            
            for r in rows:
                if group_idx < len(r):
                    inep = str(r[group_idx]).strip()
                    inep = re.sub(r'\D', '', inep)
                    
                    if not inep:
                        continue
                        
                    data_inst = str(r[group_idx+1]).strip() if (group_idx+1 < len(r) and r[group_idx+1]) else ""
                    status = str(r[group_idx+2]).strip() if (group_idx+2 < len(r) and r[group_idx+2]) else ""
                    obs = str(r[group_idx+3]).strip() if (group_idx+3 < len(r) and r[group_idx+3]) else ""
                    links = str(r[group_idx+4]).strip() if (group_idx+4 < len(r) and r[group_idx+4]) else ""
                    
                    records.append({
                        "INEP": inep,
                        "UF": uf,
                        "REGIONAL": regional_curta,
                        "DATA_INSTALAÇÃO": data_inst,
                        "STATUS": status if status else "PENDENTE",
                        "OBSERVAÇÃO": obs,
                        "LINKS": links
                    })
                    
        return pd.DataFrame(records)

    def enriquecer_dados_escola(self, df_consolidado):
        if df_consolidado.empty:
            return df_consolidado
            
        df_cache = load_cache()
        if df_cache.empty:
            df_consolidado["ESCOLA"] = "—"
            df_consolidado["MUNICÍPIO"] = "—"
            return df_consolidado
            
        col_inep_c = None
        col_nome_c = None
        col_mun_c = None
        
        for c in df_cache.columns:
            c_low = c.lower()
            if "inep" in c_low and not col_inep_c:
                col_inep_c = c
            elif ("escola" in c_low or "nome" in c_low or "unidade" in c_low) and not col_nome_c:
                col_nome_c = c
            elif ("munic" in c_low or "cidade" in c_low) and not col_mun_c:
                col_mun_c = c
                
        if not col_inep_c or not col_nome_c:
            df_consolidado["ESCOLA"] = "—"
            df_consolidado["MUNICÍPIO"] = "—"
            return df_consolidado
            
        df_cache_clean = df_cache.copy()
        df_cache_clean[col_inep_c] = df_cache_clean[col_inep_c].astype(str).str.replace(r'\D', '', regex=True).str.strip()
        
        map_nome = dict(zip(df_cache_clean[col_inep_c], df_cache_clean[col_nome_c]))
        map_mun = dict(zip(df_cache_clean[col_inep_c], df_cache_clean[col_mun_c])) if col_mun_c else {}
        
        df_consolidado["ESCOLA"] = df_consolidado["INEP"].map(map_nome).fillna("—")
        df_consolidado["MUNICÍPIO"] = df_consolidado["INEP"].map(map_mun).fillna("—")
        
        return df_consolidado

    def importar_do_csv_local(self, path):
        if not os.path.exists(path):
            return pd.DataFrame()
        try:
            import csv
            raw_data = []
            with open(path, 'r', encoding='utf-8', errors='replace') as f:
                reader = csv.reader(f, delimiter=',')
                for row in reader:
                    raw_data.append(row)
            if len(raw_data) < 2:
                return pd.DataFrame()
            
            df_cons = self.consolidar_matriz_raw(raw_data)
            return self.enriquecer_dados_escola(df_cons)
        except Exception as e:
            print(f"Erro ao ler CSV local: {e}")
            return pd.DataFrame()

    def carregar_so_local(self):
        cache_p = get_dashboard_cache_path()
        
        # Tentar carregar cache consolidado primeiro
        self.df = load_cache(cache_p)
        
        # Fallback para o CSV se o cache estiver vazio
        if self.df.empty:
            csv_path = os.path.join(get_base_path(), "Validação Instação - Plan1.csv")
            if os.path.exists(csv_path):
                self.df = self.importar_do_csv_local(csv_path)
                if not self.df.empty:
                    save_cache(self.df, cache_p)
                    self.lbl_status.config(text="✅ Importado de Validação Instação - Plan1.csv local.", fg="green")
            
        if not self.df.empty:
            self.popular_filtros(reset_municipios=True)
            self.aplicar_filtros()
            self.lbl_status.config(text=f"✅ Dados locais carregados — {len(self.df)} escolas monitoradas.", fg="green")
        else:
            self.lbl_status.config(text="⚠️ Sem base de dados local. Clique em 'Sincronizar Nuvem' para carregar.", fg="orange")

    def forcar_atualizacao(self):
        self.lbl_status.config(text="⏳ Conectando à nuvem e baixando dados... Aguarde.", fg="blue")
        threading.Thread(target=self._sincronizar_nuvem, daemon=True).start()

    def _sincronizar_nuvem(self):
        try:
            if not self.gspread_client:
                self.gspread_client = authenticate_google_sheets()
            if not self.gspread_client:
                raise Exception("Sem credenciais válidas do Google (credentials.json ausente).")
            
            spreadsheet = self.gspread_client.open_by_key(DASHBOARD_SPREADSHEET_ID)
            worksheet = None
            for w in spreadsheet.worksheets():
                if str(w.id) == "1113059829":
                    worksheet = w
                    break
            if not worksheet:
                worksheet = spreadsheet.get_worksheet(0)
                
            raw_data = worksheet.get_all_values()
            if not raw_data:
                raise Exception("A planilha está vazia.")
                
            df_cons = self.consolidar_matriz_raw(raw_data)
            df_nuvem = self.enriquecer_dados_escola(df_cons)
            
            self.df = df_nuvem
            save_cache(df_nuvem, get_dashboard_cache_path())
            
            self.frame.after(0, lambda: self.popular_filtros(reset_municipios=True))
            self.frame.after(0, self.aplicar_filtros)
            self.frame.after(0, lambda: self.lbl_status.config(text=f"✅ Sincronizado com sucesso! {len(self.df)} registros.", fg="green"))
            
        except Exception as e:
            if self.df is not None and not self.df.empty:
                self.frame.after(0, lambda: self.lbl_status.config(text="⚠️ Falha na rede. Operando com dados locais salvos.", fg="orange"))
            else:
                self.frame.after(0, lambda: self.lbl_status.config(text="❌ ERRO DE CONEXÃO. Sem cache e sem rede.", fg="red"))
                self.frame.after(0, lambda: show_detailed_error("Erro de Conexão", "Não foi possível acessar a planilha em nuvem.", e))

    def popular_filtros(self, reset_municipios=True):
        if self.df is None or self.df.empty: return

        # Regional
        regs = sorted(self.df["REGIONAL"].dropna().unique())
        self.cb_reg["values"] = ["Todas"] + [str(r).strip() for r in regs if str(r).strip()]
        self.cb_reg.set("Todas")

        # UF
        ufs = sorted(self.df["UF"].dropna().unique())
        self.cb_uf["values"] = ["Todas"] + [str(u).strip() for u in ufs if str(u).strip()]
        self.cb_uf.set("Todas")

        # Status
        status = sorted(self.df["STATUS"].dropna().unique())
        self.cb_status["values"] = ["Todos"] + [str(s).strip() for s in status if str(s).strip()]
        self.cb_status.set("Todos")
        
        # Mês/Ano
        if hasattr(self, "cb_mes_data"):
            datas_validas = pd.to_datetime(self.df["DATA_INSTALAÇÃO"], format="%d/%m/%Y", errors="coerce").dropna()
            meses_unicos = sorted(datas_validas.dt.strftime("%m/%Y").unique())
            self.cb_mes_data["values"] = ["Todos"] + meses_unicos
            self.cb_mes_data.set("Todos")

        if reset_municipios:
            muns = sorted(self.df["MUNICÍPIO"].dropna().unique())
            self.cb_mun["values"] = ["Todos"] + [str(m).strip() for m in muns if str(m).strip()]
            self.cb_mun.set("Todos")

    def on_reg_changed(self, event=None):
        reg_sel = self.cb_reg.get()
        if self.df is not None and not self.df.empty:
            if reg_sel and reg_sel != "Todas":
                df_reg = self.df[self.df["REGIONAL"] == reg_sel]
                ufs = sorted(df_reg["UF"].dropna().unique())
            else:
                ufs = sorted(self.df["UF"].dropna().unique())
                
            self.cb_uf["values"] = ["Todas"] + [str(u).strip() for u in ufs if str(u).strip()]
            self.cb_uf.set("Todas")
        else:
            self.cb_uf["values"] = ["Todas"]
            self.cb_uf.set("Todas")
            
        self.on_uf_changed()

    def on_uf_changed(self, event=None):
        uf_sel = self.cb_uf.get()
        reg_sel = self.cb_reg.get()
        
        if self.df is not None and not self.df.empty:
            df_filtered = self.df.copy()
            if reg_sel and reg_sel != "Todas":
                df_filtered = df_filtered[df_filtered["REGIONAL"] == reg_sel]
            if uf_sel and uf_sel != "Todas":
                df_filtered = df_filtered[df_filtered["UF"] == uf_sel]
                
            muns = sorted(df_filtered["MUNICÍPIO"].dropna().unique())
            self.cb_mun["values"] = ["Todos"] + [str(m).strip() for m in muns if str(m).strip()]
            self.cb_mun.set("Todos")
        else:
            self.cb_mun["values"] = ["Todos"]
            self.cb_mun.set("Todos")
            
        self.aplicar_filtros()

    def on_filter_changed(self, event=None):
        self.aplicar_filtros()

    def aplicar_filtros(self):
        if self.df is None or self.df.empty: return

        df_filtrado = self.df.copy()

        # Regional
        reg_sel = self.cb_reg.get()
        if reg_sel and reg_sel != "Todas":
            df_filtrado = df_filtrado[df_filtrado["REGIONAL"] == reg_sel]

        # UF
        uf_sel = self.cb_uf.get()
        if uf_sel and uf_sel != "Todas":
            df_filtrado = df_filtrado[df_filtrado["UF"] == uf_sel]

        # Município
        mun_sel = self.cb_mun.get()
        if mun_sel and mun_sel != "Todos":
            df_filtrado = df_filtrado[df_filtrado["MUNICÍPIO"] == mun_sel]

        # Status
        status_sel = self.cb_status.get()
        if status_sel and status_sel != "Todos":
            df_filtrado = df_filtrado[df_filtrado["STATUS"] == status_sel]

        # Busca Texto
        busca = self.ent_search.get().strip().lower()
        if busca:
            condicoes = [
                df_filtrado["INEP"].astype(str).str.lower().str.contains(busca),
                df_filtrado["ESCOLA"].astype(str).str.lower().str.contains(busca),
                df_filtrado["MUNICÍPIO"].astype(str).str.lower().str.contains(busca)
            ]
            combined = condicoes[0]
            for cond in condicoes[1:]:
                combined = combined | cond
            df_filtrado = df_filtrado[combined]

        # Processamento de Data (Filtro e Ordenação)
        mes_sel = getattr(self, "cb_mes_data", None)
        ordem_sel = getattr(self, "cb_ordem_data", None)
        
        precisa_data = (mes_sel and mes_sel.get() != "Todos") or (ordem_sel and ordem_sel.get() != "Padrão")
        
        if precisa_data:
            df_filtrado["DATA_PARSED"] = pd.to_datetime(df_filtrado["DATA_INSTALAÇÃO"], format="%d/%m/%Y", errors="coerce")
            
            # 1. Filtrar por Mês
            if mes_sel and mes_sel.get() != "Todos":
                df_filtrado = df_filtrado[df_filtrado["DATA_PARSED"].dt.strftime("%m/%Y") == mes_sel.get()]
                
            # 2. Ordenar
            if ordem_sel and ordem_sel.get() != "Padrão":
                ordem = ordem_sel.get()
                if ordem == "Mais Antigas Primeiro":
                    df_filtrado = df_filtrado.sort_values(by="DATA_PARSED", ascending=True, na_position="last")
                elif ordem == "Mais Recentes Primeiro":
                    df_filtrado = df_filtrado.sort_values(by="DATA_PARSED", ascending=False, na_position="last")
                    
            df_filtrado = df_filtrado.drop(columns=["DATA_PARSED"])

        # Atualizar Treeview
        self.tree.delete(*self.tree.get_children())
        
        for idx, row in df_filtrado.iterrows():
            self.tree.insert("", "end", values=(
                row.get("INEP", ""),
                row.get("ESCOLA", ""),
                row.get("MUNICÍPIO", ""),
                row.get("UF", ""),
                row.get("REGIONAL", ""),
                row.get("STATUS", ""),
                row.get("DATA_INSTALAÇÃO", ""),
                row.get("OBSERVAÇÃO", "")
            ))

        self.atualizar_cards(df_filtrado)

    def atualizar_cards(self, df_filtrado):
        if df_filtrado.empty:
            self.card_labels["lbl_card_total"].config(text="0")
            self.card_labels["lbl_card_concluidas"].config(text="0")
            self.card_labels["lbl_card_progresso"].config(text="0")
            self.card_labels["lbl_card_taxa"].config(text="0.0%")
            return

        total = len(df_filtrado)
        concluidas = 0
        progresso = 0
        
        status_series = df_filtrado["STATUS"].astype(str).str.lower().str.strip()
        
        termos_concluidos = ["ok", "concluido", "concluída", "concluidos", "concluídas", "finalizado", "finalizada", "instalado", "instalada", "ativo", "ativa", "aprovado", "aprovada"]
        termos_progresso = ["andamento", "progresso", "instalando", "execucao", "execução", "iniciado", "iniciada", "vistoria", "agendado", "pendente"]
        
        for st in status_series:
            if any(t in st for t in termos_concluidos):
                concluidas += 1
            elif any(t in st for t in termos_progresso):
                progresso += 1
            else:
                if st and st != "nan" and st != "":
                    progresso += 1

        taxa = (concluidas / total) * 100 if total > 0 else 0.0
        
        self.card_labels["lbl_card_total"].config(text=str(total))
        self.card_labels["lbl_card_concluidas"].config(text=str(concluidas))
        self.card_labels["lbl_card_progresso"].config(text=str(progresso))
        self.card_labels["lbl_card_taxa"].config(text=f"{taxa:.1f}%")

    def copiar_registro(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione um registro na tabela para copiar.")
            return
            
        item = self.tree.item(sel[0])
        vals = item["values"]
        
        texto = (
            f"INEP:      {vals[0]}\n"
            f"ESCOLA:    {vals[1]}\n"
            f"MUNICÍPIO: {vals[2]}\n"
            f"UF:        {vals[3]}\n"
            f"REGIONAL:  {vals[4]}\n"
            f"STATUS:    {vals[5]}\n"
            f"INSTALAÇÃO:{vals[6]}\n"
            f"OBSERVAÇÃO:{vals[7]}"
        )
        self.frame.clipboard_clear()
        self.frame.clipboard_append(texto)
        messagebox.showinfo("Sucesso ✅", "Registro copiado para a área de transferência!")

    def copiar_tabela(self):
        children = self.tree.get_children()
        if not children:
            messagebox.showwarning("Aviso", "Tabela vazia. Nada para copiar.")
            return
            
        linhas = []
        linhas.append("Código INEP\tNome da Escola\tMunicípio\tUF\tRegional\tStatus\tData Instalação\tObservação")
        
        for child in children:
            vals = self.tree.item(child)["values"]
            linhas.append(f"{vals[0]}\t{vals[1]}\t{vals[2]}\t{vals[3]}\t{vals[4]}\t{vals[5]}\t{vals[6]}\t{vals[7]}")
            
        texto = "\n".join(linhas)
        self.frame.clipboard_clear()
        self.frame.clipboard_append(texto)
        messagebox.showinfo("Sucesso ✅", f"Copiadas {len(children)} linhas da tabela no formato Excel (Tabulado)!")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Monitoramento EACE - Dashboard (v3)")
    root.geometry("980x700")
    
    gs = authenticate_google_sheets()
    
    app = DashboardEACE(root, lambda: root.destroy(), gs)
    root.mainloop()
