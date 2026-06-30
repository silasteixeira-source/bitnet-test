#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TESTE v2 — Verificador de Velocidade EACE
Layout redesenhado: sem scroll, tudo visivel de uma vez.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import json, os, re
import pandas as pd

# ── Dados ─────────────────────────────────────────────────────
CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "eace_cache.json")
LIMIAR_PCT = 80
COL_DL_MIN = "Velocidade DL Mínima (Mbps)"
COL_DL_OFE = "Velocidade DL Ofertada (Mbps)"

# ── Paleta ────────────────────────────────────────────────────
BG      = "#0d1117"
CARD    = "#161b22"
CARD2   = "#1c2128"
FG      = "#e6edf3"
MUTED   = "#8b949e"
ACCENT  = "#58a6ff"
SUCCESS = "#3fb950"
DANGER  = "#f85149"
WARN    = "#d29922"
BORDER  = "#30363d"
BTN_BG  = "#21262d"


def carregar_cache():
    if not os.path.exists(CACHE_FILE):
        return pd.DataFrame()
    try:
        df = pd.read_json(CACHE_FILE, dtype=str)
        col = next((c for c in df.columns if "inep" in c.lower()), df.columns[0])
        df[col] = df[col].astype(str).str.split('.').str[0].str.strip()
        return df
    except Exception as e:
        print(f"Erro cache: {e}")
        return pd.DataFrame()


def extrair_mbps(v):
    try:
        nums = re.findall(r'[\d.,]+', str(v).replace(',', '.'))
        return float(nums[0]) if nums else None
    except Exception:
        return None


def estilo_entry(parent, **kw):
    e = tk.Entry(parent, bg=CARD2, fg=FG, insertbackground=FG,
                 relief="flat", bd=0, highlightthickness=1,
                 highlightbackground=BORDER, highlightcolor=ACCENT,
                 disabledbackground=CARD2, disabledforeground=FG, **kw)
    return e


def btn(parent, text, cmd, color=ACCENT, **kw):
    return tk.Button(parent, text=text, command=cmd,
                     bg=color, fg=FG if color != ACCENT else BG,
                     activebackground=color, activeforeground=FG,
                     font=("Segoe UI", 9, "bold"), bd=0,
                     cursor="hand2", relief="flat", **kw)


# ── App ───────────────────────────────────────────────────────
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Consulta EACE — Verificador de Velocidade")
        self.root.configure(bg=BG)
        self.root.geometry("980x640")
        self.root.minsize(900, 600)

        self.df = pd.DataFrame()
        self.row_atual = None
        self.dl_min_val = None
        self.dl_ofe_val = None

        self._build()
        self._load_cache()

    # ── Construção da UI ─────────────────────────────────────
    def _build(self):
        # ── TOP BAR ──────────────────────────────────────────
        top = tk.Frame(self.root, bg=CARD, height=48)
        top.pack(fill="x")
        top.pack_propagate(False)

        tk.Label(top, text="  🔍  CONSULTA RÁPIDA — EACE",
                 font=("Segoe UI", 13, "bold"), bg=CARD, fg=ACCENT).pack(side="left", padx=8)

        # Status (topo direito)
        self.lbl_status = tk.Label(top, text="Carregando...",
                                   font=("Segoe UI", 8, "italic"), bg=CARD, fg=WARN)
        self.lbl_status.pack(side="right", padx=16)

        # ── BARRA DE BUSCA ───────────────────────────────────
        bar = tk.Frame(self.root, bg=BG, pady=10)
        bar.pack(fill="x", padx=16)

        tk.Label(bar, text="Código INEP ou Nome:", bg=BG, fg=MUTED,
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 6))

        self.ent_busca = estilo_entry(bar, font=("Segoe UI", 11), width=32)
        self.ent_busca.pack(side="left", ipady=6, padx=(0, 8))
        self.ent_busca.bind("<Return>", lambda e: self._buscar())

        btn(bar, "  BUSCAR DADOS  ", self._buscar, color=ACCENT).pack(side="left", ipady=6, padx=(0, 8))
        btn(bar, "  ☁ Atualizar da Nuvem  ", self._placeholder_nuvem,
            color="#238636").pack(side="left", ipady=6)

        # ── SEPARADOR ────────────────────────────────────────
        tk.Frame(self.root, bg=BORDER, height=1).pack(fill="x")

        # ── CONTEÚDO PRINCIPAL (2 colunas) ───────────────────
        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=16, pady=10)
        body.columnconfigure(0, weight=6)
        body.columnconfigure(1, weight=4)
        body.rowconfigure(0, weight=1)

        # ── COLUNA ESQUERDA: Dados da Escola ─────────────────
        left = tk.Frame(body, bg=CARD, bd=0, relief="flat",
                        highlightthickness=1, highlightbackground=BORDER)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        tk.Label(left, text="DADOS DA ESCOLA",
                 font=("Segoe UI", 9, "bold"), bg=CARD, fg=ACCENT).pack(
                     anchor="w", padx=14, pady=(10, 6))
        tk.Frame(left, bg=BORDER, height=1).pack(fill="x", padx=14)

        # Grid de campos
        self.campos_frame = tk.Frame(left, bg=CARD)
        self.campos_frame.pack(fill="both", expand=True, padx=14, pady=8)
        self.campos_frame.columnconfigure(1, weight=1)

        # Define campos e variáveis
        self._campo_defs = [
            ("INEP",           "Código INEP",           False),
            ("Escola",         "Nome da Escola",         False),
            ("Município / UF", "__municipio_uf__",       False),
            ("Endereço",       "Endereço",               False),
            ("Latitude",       "Latitude",               False),
            ("Longitude",      "Longitude",              False),
            ("DL Mínima",      COL_DL_MIN,              True),   # destaque
            ("DL Ofertada",    COL_DL_OFE,              False),
            ("Status",         "Status",                 False),
        ]

        self._vars = {}
        for i, (label, col, destaque) in enumerate(self._campo_defs):
            tk.Label(self.campos_frame, text=label + ":",
                     font=("Segoe UI", 9), bg=CARD, fg=MUTED,
                     anchor="w", width=14).grid(row=i, column=0, sticky="w", pady=5)

            var = tk.StringVar(value="—")
            cor_val = WARN if destaque else FG
            peso = "bold" if destaque else "normal"
            lbl = tk.Label(self.campos_frame, textvariable=var,
                           font=("Segoe UI", 9, peso), bg=CARD, fg=cor_val,
                           anchor="w", wraplength=360, justify="left")
            lbl.grid(row=i, column=1, sticky="w", pady=5, padx=(8, 0))
            self._vars[col] = var

        # Botão copiar
        frame_btn_copy = tk.Frame(left, bg=CARD)
        frame_btn_copy.pack(fill="x", padx=14, pady=(0, 12))
        self.btn_copiar = btn(frame_btn_copy, "📋  Copiar Ficha", self._copiar,
                              color=BTN_BG)
        self.btn_copiar.pack(side="left", ipady=5, ipadx=8)
        self.btn_copiar.config(state="disabled")

        # ── COLUNA DIREITA: Verificador ───────────────────────
        right = tk.Frame(body, bg=CARD, bd=0, relief="flat",
                         highlightthickness=1, highlightbackground=BORDER)
        right.grid(row=0, column=1, sticky="nsew")

        tk.Label(right, text="⚡  VERIFICADOR DE VELOCIDADE",
                 font=("Segoe UI", 9, "bold"), bg=CARD, fg=WARN).pack(
                     anchor="w", padx=14, pady=(10, 6))
        tk.Frame(right, bg=BORDER, height=1).pack(fill="x", padx=14)

        inner = tk.Frame(right, bg=CARD)
        inner.pack(fill="both", expand=True, padx=14, pady=10)

        # Info de velocidade
        for label, attr in [("Velocidade Ofertada:", "lbl_contratada"),
                            ("Mínima Requerida (Planilha):", "lbl_threshold")]:
            row_f = tk.Frame(inner, bg=CARD)
            row_f.pack(fill="x", pady=3)
            tk.Label(row_f, text=label, font=("Segoe UI", 9), bg=CARD, fg=MUTED,
                     width=25, anchor="w").pack(side="left")
            lbl = tk.Label(row_f, text="—", font=("Segoe UI", 9, "bold"),
                           bg=CARD, fg=FG, anchor="w")
            lbl.pack(side="left")
            setattr(self, attr, lbl)

        tk.Frame(inner, bg=BORDER, height=1).pack(fill="x", pady=10)

        # Input speedtest
        tk.Label(inner, text="Resultado do Speedtest:", font=("Segoe UI", 9, "bold"),
                 bg=CARD, fg=FG, anchor="w").pack(fill="x")

        speed_row = tk.Frame(inner, bg=CARD)
        speed_row.pack(fill="x", pady=(6, 0))

        self.ent_speed = estilo_entry(speed_row, font=("Segoe UI", 22, "bold"),
                                      width=8, justify="center")
        self.ent_speed.pack(side="left", ipady=8)
        self.ent_speed.bind("<Return>", lambda e: self._verificar())

        tk.Label(speed_row, text=" Mbps", font=("Segoe UI", 11), bg=CARD,
                 fg=MUTED).pack(side="left")

        btn(inner, "  CALCULAR  ", self._verificar, color="#1f6feb").pack(
            fill="x", pady=(10, 0), ipady=8)

        # Área de resultado
        tk.Frame(inner, bg=BORDER, height=1).pack(fill="x", pady=(14, 8))

        self.frame_res = tk.Frame(inner, bg=CARD)
        self.frame_res.pack(fill="both", expand=True)

        self.lbl_res_icone = tk.Label(self.frame_res, text="",
                                      font=("Segoe UI", 28, "bold"), bg=CARD)
        self.lbl_res_icone.pack()

        self.lbl_res_status = tk.Label(self.frame_res, text="",
                                       font=("Segoe UI", 14, "bold"), bg=CARD)
        self.lbl_res_status.pack()

        self.lbl_res_pct = tk.Label(self.frame_res, text="",
                                    font=("Segoe UI", 10), bg=CARD, fg=MUTED)
        self.lbl_res_pct.pack(pady=(2, 8))

        # Mini barra de progresso
        self.canvas_barra = tk.Canvas(self.frame_res, height=18, bg=CARD2,
                                      bd=0, highlightthickness=0)
        self.canvas_barra.pack(fill="x", padx=2)

        # Detalhes resultado
        self.frame_det = tk.Frame(self.frame_res, bg=CARD)
        self.frame_det.pack(fill="x", pady=(10, 0))

        self._det_vars = {}
        for label in ["Medido", "Minimo", "Diferenca"]:
            r = tk.Frame(self.frame_det, bg=CARD)
            r.pack(fill="x", pady=1)
            tk.Label(r, text=label + ":", font=("Segoe UI", 9), bg=CARD,
                     fg=MUTED, width=12, anchor="w").pack(side="left")
            v = tk.StringVar(value="")
            lbl = tk.Label(r, textvariable=v, font=("Segoe UI", 9, "bold"),
                           bg=CARD, fg=FG, anchor="w")
            lbl.pack(side="left")
            self._det_vars[label] = (v, lbl)

        # Nota rodapé
        self.lbl_nota = tk.Label(right,
            text="Busque uma escola e insira o resultado do Speedtest.",
            font=("Segoe UI", 8, "italic"), bg=CARD, fg=MUTED, wraplength=280)
        self.lbl_nota.pack(padx=14, pady=(0, 10))

    # ── Lógica ───────────────────────────────────────────────
    def _load_cache(self):
        self.df = carregar_cache()
        if not self.df.empty:
            self.lbl_status.config(
                text=f"✅  Cache local carregado — {len(self.df)} escolas", fg=SUCCESS)
            self.ent_busca.focus()
        else:
            self.lbl_status.config(
                text="⚠  Cache vazio. Abra o app e sincronize com a nuvem.", fg=WARN)

    def _buscar(self):
        termo = self.ent_busca.get().strip()
        if not termo:
            return
        if self.df.empty:
            messagebox.showwarning("Aviso", "Sem base de dados. Sincronize a nuvem no app principal.")
            return

        self._limpar_escola()
        col_inep = next((c for c in self.df.columns if "inep" in c.lower()), self.df.columns[0])
        col_nome = next((c for c in self.df.columns
                         if "escola" in c.lower() or "nome" in c.lower()), None)

        if termo.isdigit() and len(termo) >= 8:
            res = self.df[self.df[col_inep] == termo]
        elif col_nome:
            res = self.df[self.df[col_nome].astype(str).str.contains(termo, case=False, na=False)]
        else:
            res = pd.DataFrame()

        if res.empty:
            messagebox.showinfo("Não encontrado", f"Nenhuma escola localizada para:\n'{termo}'")
            return

        if len(res) > 1:
            self._popup_escolha(res, col_nome)
        else:
            self._exibir(res.iloc[0])

    def _exibir(self, row):
        self.row_atual = row

        col_inep = next((c for c in row.index if "inep" in c.lower()), row.index[0])
        col_nome = next((c for c in row.index
                         if "escola" in c.lower() or "nome" in c.lower()), "")

        mun = row.get("Município", "N/A")
        uf  = row.get("UF", "N/A")

        # Preencher variáveis dos campos
        mapping = {
            "Código INEP":    row.get(col_inep, "N/A"),
            "Nome da Escola": row.get(col_nome, "N/A"),
            "__municipio_uf__": f"{mun} / {uf}",
            "Endereço":       row.get("Endereço", "N/A"),
            "Latitude":       row.get("Latitude", "N/A"),
            "Longitude":      row.get("Longitude", "N/A"),
            COL_DL_MIN:       str(row.get(COL_DL_MIN, "N/A")) + " Mbps",
            COL_DL_OFE:       str(row.get(COL_DL_OFE, "N/A")) + " Mbps",
            "Status":         row.get("Status", "Ativo"),
        }
        for col, var in self._vars.items():
            var.set(mapping.get(col, "N/A"))

        self.btn_copiar.config(state="normal")

        # Atualizar painel de velocidade
        dl_raw = row.get(COL_DL_MIN, "N/A")
        self.dl_min_val = extrair_mbps(dl_raw)
        self.dl_ofe_val = extrair_mbps(row.get(COL_DL_OFE, "N/A"))

        if self.dl_min_val is not None:
            self.lbl_contratada.config(text=f"{self.dl_ofe_val:.1f} Mbps" if self.dl_ofe_val is not None else "N/A", fg=FG)
            self.lbl_threshold.config(text=f"{self.dl_min_val:.1f} Mbps", fg=WARN)
            self.lbl_nota.config(text="Insira o resultado do Speedtest e clique em CALCULAR.")
            self.ent_speed.focus()
        else:
            self.lbl_contratada.config(text="N/A", fg=WARN)
            self.lbl_threshold.config(text=str(dl_raw) + " (não numérico)", fg=WARN)
            self.lbl_nota.config(text="Velocidade não numérica — verificação manual necessária.")

        self._limpar_resultado()

    def _verificar(self):
        if self.row_atual is None:
            messagebox.showwarning("Aviso", "Busque uma escola primeiro.")
            return
        if self.dl_min_val is None:
            messagebox.showwarning("Aviso", "Velocidade mínima não numérica. Verifique manualmente.")
            return

        raw = self.ent_speed.get().strip().replace(',', '.')
        if not raw:
            messagebox.showwarning("Aviso", "Digite o resultado do Speedtest.")
            return
        try:
            medido = float(raw)
        except ValueError:
            messagebox.showerror("Erro", "Valor inválido. Use números (ex: 45.3).")
            return

        threshold  = self.dl_min_val
        aprovado   = medido >= threshold
        diferenca  = medido - threshold
        
        if self.dl_ofe_val and self.dl_ofe_val > 0:
            pct = (medido / self.dl_ofe_val) * 100
        else:
            pct = (medido / self.dl_min_val) * 100

        self._mostrar_resultado(medido, threshold, pct, aprovado, diferenca)

    def _mostrar_resultado(self, medido, threshold, pct, aprovado, diferenca):
        cor    = SUCCESS if aprovado else DANGER
        status = "APROVADO" if aprovado else "REPROVADO"
        icone  = "✅" if aprovado else "❌"

        self.lbl_res_icone.config(text=icone, fg=cor)
        self.lbl_res_status.config(text=status, fg=cor)
        
        if self.dl_ofe_val and self.dl_ofe_val > 0:
            self.lbl_res_pct.config(text=f"{pct:.1f}% da velocidade ofertada")
        else:
            self.lbl_res_pct.config(text=f"{pct:.1f}% da velocidade mínima")

        # Barra de progresso
        self.canvas_barra.update_idletasks()
        w = self.canvas_barra.winfo_width() or 240
        self.canvas_barra.delete("all")
        self.canvas_barra.create_rectangle(0, 0, w, 18, fill=CARD2, outline="")
        
        fill_w = int(w * min(pct, 100) / 100)
        self.canvas_barra.create_rectangle(0, 0, fill_w, 18, fill=cor, outline="")
        
        if self.dl_ofe_val and self.dl_ofe_val > 0:
            threshold_pct = (self.dl_min_val / self.dl_ofe_val) * 100
        else:
            threshold_pct = 100
            
        x_thresh = int(w * threshold_pct / 100)
        self.canvas_barra.create_line(x_thresh, 0, x_thresh, 18, fill=WARN, width=2, dash=(3, 2))
        self.canvas_barra.create_text(w // 2, 9, text=f"{pct:.1f}%",
                                      fill="white", font=("Segoe UI", 7, "bold"))

        # Detalhes
        sinal = "+" if diferenca >= 0 else ""
        vals = {
            "Medido":    (f"{medido:.1f} Mbps", FG),
            "Minimo":    (f"{threshold:.1f} Mbps", WARN),
            "Diferenca": (f"{sinal}{diferenca:.1f} Mbps",
                          SUCCESS if diferenca >= 0 else DANGER),
        }
        for key, (texto, cor_det) in vals.items():
            var, lbl = self._det_vars[key]
            var.set(texto)
            lbl.config(fg=cor_det)

        self.lbl_nota.config(text="")

    def _limpar_escola(self):
        for var in self._vars.values():
            var.set("—")
        self.btn_copiar.config(state="disabled")
        self.lbl_contratada.config(text="—", fg=FG)
        self.lbl_threshold.config(text="—", fg=FG)
        self.dl_min_val = None
        self.row_atual = None
        self._limpar_resultado()

    def _limpar_resultado(self):
        self.lbl_res_icone.config(text="")
        self.lbl_res_status.config(text="")
        self.lbl_res_pct.config(text="")
        self.canvas_barra.delete("all")
        for var, lbl in self._det_vars.values():
            var.set("")
            lbl.config(fg=FG)

    def _copiar(self):
        if self.row_atual is None:
            return
        linhas = []
        for lbl, col, _ in self._campo_defs:
            val = self._vars.get(col, tk.StringVar()).get()
            linhas.append(f"{lbl+':':<20} {val}")
        ficha = "\n".join(linhas)
        self.root.clipboard_clear()
        self.root.clipboard_append(ficha)
        messagebox.showinfo("Copiado", "Ficha copiada para a área de transferência.")

    def _placeholder_nuvem(self):
        messagebox.showinfo("Info",
            "No app principal, use 'Atualizar Base da Nuvem' para sincronizar.\n"
            "Este teste opera apenas com o cache local.")

    def _popup_escolha(self, df_res, col_nome):
        pop = tk.Toplevel(self.root)
        pop.title("Múltiplos resultados")
        pop.geometry("500x320")
        pop.configure(bg=BG)
        pop.transient(self.root)
        pop.grab_set()

        tk.Label(pop, text=f"{len(df_res)} escolas encontradas — selecione:",
                 font=("Segoe UI", 10, "bold"), bg=BG, fg=FG).pack(pady=10)

        lb = tk.Listbox(pop, font=("Segoe UI", 9), bg=CARD, fg=FG,
                        selectbackground=ACCENT, selectforeground=BG,
                        bd=0, relief="flat", activestyle="none")
        lb.pack(fill="both", expand=True, padx=16)

        for _, r in df_res.iterrows():
            col_inep = next((c for c in r.index if "inep" in c.lower()), r.index[0])
            nome = r.get(col_nome, "—") if col_nome else "—"
            lb.insert(tk.END, f"  {r.get(col_inep, '?')}  —  {nome}")

        def ok():
            sel = lb.curselection()
            if not sel:
                return
            self._exibir(df_res.iloc[sel[0]])
            pop.destroy()

        btn(pop, "  Selecionar  ", ok, color=ACCENT).pack(pady=10, ipady=5, ipadx=10)


# ── Main ─────────────────────────────────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
