# -*- coding: utf-8 -*-
"""
Painel da Granja — controle de ovos, ração, gastos, lucros e prejuízos.
Abra o EXE e registre o dia. Feito para produtores de galinhas poedeiras.
"""
import os
import sys
import csv
import json
import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import urllib.request
import urllib.parse
from datetime import datetime, date
import calendar
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- paleta premium
VERDE_ESCURO = "#143d2b"
VERDE        = "#1f7a4d"
VERDE_CLARO  = "#e7f2ea"
AMARELO      = "#ffb300"
FUNDO        = "#f2f5f3"
CARD         = "#ffffff"
BORDA        = "#d7e0d9"
TEXTO        = "#1d2b22"
CINZA        = "#5f6f65"
VERMELHO     = "#c62828"

FONTE        = ("Segoe UI", 10)
FONTE_GRANDE = ("Segoe UI", 12)
FONTE_TITULO = ("Segoe UI", 16, "bold")

# ---------------------------------------------------------------- pasta de dados
def app_dir():
    if getattr(sys, "frozen", False):  # rodando como EXE
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(os.environ.get("APPDATA", app_dir()), "PainelGranja")
os.makedirs(DATA_DIR, exist_ok=True)
DB = os.path.join(DATA_DIR, "granja.db")

def conn():
    c = sqlite3.connect(DB)
    c.execute("""CREATE TABLE IF NOT EXISTS dias(
        data TEXT PRIMARY KEY,
        ovos INTEGER, vendidos REAL, preco_venda REAL,
        quebrados INTEGER, mortas INTEGER,
        racao_kg REAL, custo_racao REAL, outros REAL, obs TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS compras(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        data TEXT, kg REAL, custo REAL)""")
    c.execute("""CREATE TABLE IF NOT EXISTS config(
        chave TEXT PRIMARY KEY, valor TEXT)""")
    return c

def get_conf(c, chave, padrao=""):
    row = c.execute("SELECT valor FROM config WHERE chave=?", (chave,)).fetchone()
    return row[0] if row else padrao

def set_conf(c, chave, valor):
    c.execute("INSERT INTO config(chave,valor) VALUES(?,?) "
              "ON CONFLICT(chave) DO UPDATE SET valor=excluded.valor", (chave, str(valor)))
    c.commit()

def brl(v):
    return ("R$ %.2f" % (v or 0)).replace(".", ",", 2)

def num(txt, padrao=0.0):
    try:
        return float(str(txt).replace(",", "."))
    except (ValueError, TypeError):
        return padrao

# ---------------------------------------------------------------- clima (Open-Meteo, grátis, sem cadastro)
def buscar_clima(cidade):
    url_geo = ("https://geocoding-api.open-meteo.com/v1/search?name=" + urllib.parse.quote(cidade)
               + "&count=1&language=pt")
    with urllib.request.urlopen(url_geo, timeout=15) as r:
        geo = json.loads(r.read().decode())
    if not geo.get("results"):
        raise ValueError("Cidade não encontrada: " + cidade)
    g = geo["results"][0]
    url_fc = ("https://api.open-meteo.com/v1/forecast?latitude=%s&longitude=%s"
              "&daily=temperature_2m_max,temperature_2m_min,precipitation_sum&timezone=auto&forecast_days=6"
              % (g["latitude"], g["longitude"]))
    with urllib.request.urlopen(url_fc, timeout=15) as r:
        return json.loads(r.read().decode())

# ---------------------------------------------------------------- janela principal
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Painel da Granja")
        self.geometry("820x600")
        self.minsize(760, 560)
        self.configure(bg=FUNDO)

        self._estilos()

        # ---------- cabeçalho premium
        topo = tk.Frame(self, bg=VERDE_ESCURO)
        topo.pack(fill="x")
        tk.Label(topo, text="🐔 PAINEL DA GRANJA", bg=VERDE_ESCURO, fg="white",
                 font=FONTE_TITULO).pack(side="left", padx=16, pady=(10, 2))
        tk.Label(topo, text="Gestão profissional da sua produção de ovos",
                 bg=VERDE_ESCURO, fg="#a8cbb5",
                 font=("Segoe UI", 9)).pack(side="left", padx=(12, 0), pady=(12, 2))
        hoje = datetime.now()
        dias_sem = ["segunda", "terça", "quarta", "quinta", "sexta", "sábado", "domingo"]
        tk.Label(topo, text="%s, %d/%d/%d" % (dias_sem[hoje.weekday()].capitalize(),
                                              hoje.day, hoje.month, hoje.year),
                 bg=VERDE_ESCURO, fg=AMARELO,
                 font=("Segoe UI", 11, "bold")).pack(side="right", padx=16, pady=10)

        # ---------- abas
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=(10, 0))
        self.tab_dia = ttk.Frame(nb);   nb.add(self.tab_dia,   text="   📝  Registro do dia   ")
        self.tab_racao = ttk.Frame(nb); nb.add(self.tab_racao, text="   🌾  Ração   ")
        self.tab_rel = ttk.Frame(nb);   nb.add(self.tab_rel,   text="   💰  Lucro / Prejuízo   ")
        self.tab_clima = ttk.Frame(nb); nb.add(self.tab_clima, text="   ☁️  Clima   ")
        self.tab_conf = ttk.Frame(nb);  nb.add(self.tab_conf,  text="   ⚙️  Ajustes   ")

        self.montar_dia()
        self.montar_racao()
        self.montar_rel()
        self.montar_clima()
        self.montar_conf()

        # ---------- barra de status
        status = tk.Label(self, bg=VERDE_CLARO, fg=CINZA, font=("Segoe UI", 9),
                          text="  Dados salvos em: " + DATA_DIR + "   •   copie essa pasta de vez em quando para não perder nada")
        status.pack(fill="x", side="bottom")

        self.carregar_dia_hoje()
        self.atualizar_racao()
        self.gerar_rel()

    def _estilos(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", background=FUNDO, foreground=TEXTO, font=FONTE)
        style.configure("TFrame", background=FUNDO)
        style.configure("TNotebook", background=FUNDO, borderwidth=0, tabmargins=(2, 6, 2, 0))
        style.configure("TNotebook.Tab", background="#dfe8e1", foreground=CINZA,
                        padding=(18, 10), font=("Segoe UI", 10, "bold"), borderwidth=0)
        style.map("TNotebook.Tab",
                  background=[("selected", VERDE)], foreground=[("selected", "white")])
        # cartões
        style.configure("Card.TLabelframe", background=CARD, bordercolor=BORDA,
                        relief="solid", borderwidth=1)
        style.configure("Card.TLabelframe.Label", background=CARD, foreground=VERDE_ESCURO,
                        font=("Segoe UI", 10, "bold"))
        style.configure("Card.TFrame", background=CARD)
        style.configure("Card.TLabel", background=CARD, font=FONTE)
        style.configure("Card.TEntry", fieldbackground="white", foreground=TEXTO,
                        insertcolor=TEXTO, bordercolor=BORDA, lightcolor=BORDA, darkcolor=BORDA)
        style.map("Card.TEntry", bordercolor=[("focus", VERDE)], lightcolor=[("focus", VERDE)])
        style.configure("Card.TCheckbutton", background=CARD)
        # botões
        style.configure("Acao.TButton", background=VERDE, foreground="white",
                        font=("Segoe UI", 11, "bold"), padding=(14, 9), borderwidth=0)
        style.map("Acao.TButton",
                  background=[("pressed", VERDE_ESCURO), ("active", "#256b46")],
                  foreground=[("disabled", "#cccccc")])
        style.configure("Suave.TButton", background=CARD, foreground=VERDE,
                        font=("Segoe UI", 10, "bold"), padding=(12, 8), borderwidth=1,
                        bordercolor=BORDA)
        style.map("Suave.TButton", background=[("active", VERDE_CLARO)])
        style.configure("TLabel", background=FUNDO, font=FONTE)
        style.configure("Aviso.TLabel", font=("Segoe UI", 11, "bold"), foreground=VERMELHO, background=CARD)
        style.configure("Ok.TLabel", font=("Segoe UI", 11, "bold"), foreground=VERDE, background=CARD)
        style.configure("TituloCard.TLabel", background=CARD, foreground=VERDE_ESCURO,
                        font=("Segoe UI", 12, "bold"))
        style.configure("Treeview", background="white", fieldbackground="white",
                        foreground=TEXTO, rowheight=26, font=FONTE, bordercolor=BORDA)
        style.configure("Treeview.Heading", background=VERDE_ESCURO, foreground="white",
                        font=("Segoe UI", 10, "bold"), padding=(8, 6), borderwidth=0)
        style.map("Treeview", background=[("selected", VERDE)], foreground=[("selected", "white")])

    # ------------------------------------------------ registro do dia
    def montar_dia(self):
        f = self.tab_dia
        top = ttk.Frame(f); top.pack(fill="x", padx=12, pady=(8, 6))
        ttk.Label(top, text="Data:").pack(side="left")
        self.e_data = ttk.Entry(top, width=12, style="Card.TEntry", font=FONTE_GRANDE)
        self.e_data.insert(0, date.today().strftime("%d/%m/%Y"))
        self.e_data.pack(side="left", padx=6)
        ttk.Button(top, text="↻ Recarregar", style="Suave.TButton",
                   command=self.carregar_dia_hoje).pack(side="left", padx=4)

        body = ttk.Frame(f); body.pack(fill="both", expand=True, padx=12)
        self.campos = {}
        itens = [
            ("ovos", "🥚  Ovos coletados hoje"),
            ("vendidos", "💵  Dúzias vendidas"),
            ("preco_venda", "🏷️  Preço da dúzia (R$)"),
            ("quebrados", "💔  Ovos quebrados/perdidos"),
            ("mortas", "☠️  Galinhas mortas"),
            ("racao_kg", "🌾  Ração usada (kg)"),
            ("custo_racao", "💰  Custo dessa ração (R$)"),
            ("outros", "🧾  Outros gastos hoje (R$)"),
        ]
        for i, (chave, rot) in enumerate(itens):
            linha = i // 2; col = i % 2
            box = ttk.LabelFrame(body, text=rot, style="Card.TLabelframe")
            box.grid(row=linha, column=col, padx=7, pady=7, sticky="nsew")
            e = ttk.Entry(box, width=20, style="Card.TEntry", font=FONTE_GRANDE)
            e.pack(padx=12, pady=10, fill="x")
            self.campos[chave] = e
        body.columnconfigure(0, weight=1); body.columnconfigure(1, weight=1)

        receita = tk.Frame(body, bg=AMARELO)
        receita.grid(row=4, column=0, columnspan=2, padx=7, pady=(0, 7), sticky="nsew")
        self.lb_receita = tk.Label(receita, bg=AMARELO, fg="#4a3000",
                                   font=("Segoe UI", 12, "bold"),
                                   text="💵 Receita do dia: R$ 0,00   (dúzias × preço da dúzia)")
        self.lb_receita.pack(anchor="w", padx=14, pady=8)
        for e in self.campos["vendidos"], self.campos["preco_venda"]:
            e.bind("<KeyRelease>", self._calc_receita)

        rod = ttk.Frame(f); rod.pack(fill="x", padx=12, pady=(2, 6))
        ttk.Label(rod, text="Observações:", style="TLabel").pack(anchor="w")
        self.e_obs = tk.Text(rod, height=2, font=FONTE, bg="white",
                             highlightthickness=1, highlightbackground=BORDA)
        self.e_obs.pack(fill="x", pady=4)
        ttk.Button(rod, text="💾  SALVAR DIA", style="Acao.TButton",
                   command=self.salvar_dia).pack(pady=6, fill="x")

    def _calc_receita(self, _ev=None):
        receita = num(self.campos["vendidos"].get()) * num(self.campos["preco_venda"].get())
        self.lb_receita.config(text="💵 Receita do dia: %s   (dúzias × preço da dúzia)" % brl(receita))

    def _dado(self):
        s = self.e_data.get().strip().replace("-", "/").split("/")
        if len(s) == 3:
            dd, mm, aa = s
            if len(aa) == 4:
                try:
                    return "%04d-%02d-%02d" % (int(aa), int(mm), int(dd))
                except ValueError:
                    pass
        return date.today().isoformat()

    def carregar_dia_hoje(self):
        d = self._dado()
        with conn() as c:
            r = c.execute("SELECT * FROM dias WHERE data=?", (d,)).fetchone()
        ordem = ["ovos", "vendidos", "preco_venda", "quebrados", "mortas", "racao_kg", "custo_racao", "outros"]
        for e in self.campos.values():
            e.delete(0, "end")
        if r:
            for i, chave in enumerate(ordem):
                self.campos[chave].insert(0, str(r[i + 1] if r[i + 1] is not None else ""))
            self.e_obs.delete("1.0", "end")
            self.e_obs.insert("1.0", r[9] or "")
        else:
            self.e_obs.delete("1.0", "end")
        self._calc_receita()

    def salvar_dia(self):
        d = self._dado()
        vals = (int(num(self.campos["ovos"].get())),
                num(self.campos["vendidos"].get()),
                num(self.campos["preco_venda"].get()),
                int(num(self.campos["quebrados"].get())),
                int(num(self.campos["mortas"].get())),
                num(self.campos["racao_kg"].get()),
                num(self.campos["custo_racao"].get()),
                num(self.campos["outros"].get()),
                self.e_obs.get("1.0", "end").strip())
        with conn() as c:
            c.execute("INSERT INTO dias(data,ovos,vendidos,preco_venda,quebrados,mortas,"
                      "racao_kg,custo_racao,outros,obs) VALUES(?,?,?,?,?,?,?,?,?,?) "
                      "ON CONFLICT(data) DO UPDATE SET ovos=excluded.ovos,"
                      "vendidos=excluded.vendidos,preco_venda=excluded.preco_venda,"
                      "quebrados=excluded.quebrados,mortas=excluded.mortas,"
                      "racao_kg=excluded.racao_kg,custo_racao=excluded.custo_racao,"
                      "outros=excluded.outros,obs=excluded.obs", (d,) + vals)
        messagebox.showinfo("Salvo ✅", "Registro do dia salvo com sucesso!")
        self.atualizar_racao()
        self.gerar_rel()

    # ------------------------------------------------ ração
    def montar_racao(self):
        f = self.tab_racao
        box = ttk.LabelFrame(f, text="  Situação da ração  ", style="Card.TLabelframe")
        box.pack(fill="x", padx=12, pady=10)
        self.lb_estoque = tk.Label(box, bg=CARD, fg=VERDE_ESCURO,
                                   font=("Segoe UI", 14, "bold"), anchor="w")
        self.lb_estoque.pack(fill="x", padx=14, pady=(12, 2))
        self.lb_dias = tk.Label(box, bg=CARD, fg=CINZA, font=FONTE_GRANDE, anchor="w")
        self.lb_dias.pack(fill="x", padx=14, pady=(0, 12))

        compra = ttk.LabelFrame(f, text="  Comprei mais ração  ", style="Card.TLabelframe")
        compra.pack(fill="x", padx=12, pady=10)
        linha = ttk.Frame(compra, style="Card.TFrame"); linha.pack(padx=14, pady=10, fill="x")
        ttk.Label(linha, text="Quantidade (kg):", style="Card.TLabel").pack(side="left")
        self.e_kg = ttk.Entry(linha, width=10, style="Card.TEntry", font=FONTE_GRANDE)
        self.e_kg.pack(side="left", padx=6)
        ttk.Label(linha, text="Custo total (R$):", style="Card.TLabel").pack(side="left", padx=(14, 0))
        self.e_custo = ttk.Entry(linha, width=10, style="Card.TEntry", font=FONTE_GRANDE)
        self.e_custo.pack(side="left", padx=6)
        ttk.Button(compra, text="➕  Registrar compra", style="Acao.TButton",
                   command=self.registrar_compra).pack(padx=14, pady=(0, 14), fill="x")

        hist = ttk.LabelFrame(f, text="  Últimas compras  ", style="Card.TLabelframe")
        hist.pack(fill="both", expand=True, padx=12, pady=10)
        self.tv_compras = ttk.Treeview(hist, columns=("d", "kg", "custo"), show="headings", height=6)
        for cid, txt, w in (("d", "Data", 140), ("kg", "Kg", 120), ("custo", "Custo", 140)):
            self.tv_compras.heading(cid, text=txt); self.tv_compras.column(cid, width=w, anchor="center")
        self.tv_compras.pack(fill="both", expand=True, padx=10, pady=10)

    def estoque_atual(self):
        with conn() as c:
            kg_comprado = c.execute("SELECT COALESCE(SUM(kg),0) FROM compras").fetchone()[0]
            kg_usado = c.execute("SELECT COALESCE(SUM(racao_kg),0) FROM dias").fetchone()[0]
            inicial = num(get_conf(c, "estoque_inicial", "0"))
        return inicial + kg_comprado - kg_usado

    def atualizar_racao(self):
        est = self.estoque_atual()
        with conn() as c:
            linhas = c.execute("SELECT racao_kg FROM dias WHERE racao_kg>0 "
                               "ORDER BY data DESC LIMIT 7").fetchall()
        media = (sum(l[0] for l in linhas) / len(linhas)) if linhas else 0
        self.lb_estoque.config(text="Estoque atual: %.1f kg" % max(est, 0))
        if media <= 0:
            self.lb_dias.config(text="Registre o uso de ração no dia a dia para eu calcular os dias restantes.", fg=CINZA)
        else:
            dias = est / media
            self.lb_dias.config(text="Consumo médio: %.1f kg/dia  →  a ração dura ± %d dias" % (media, int(dias)),
                                fg=(VERMELHO if dias < 7 else VERDE))
        for i in self.tv_compras.get_children():
            self.tv_compras.delete(i)
        with conn() as c:
            for r in c.execute("SELECT data,kg,custo FROM compras ORDER BY data DESC LIMIT 30"):
                dt = datetime.fromisoformat(r[0]).strftime("%d/%m/%Y") if r[0] else ""
                self.tv_compras.insert("", "end", values=(dt, r[1], brl(r[2])))

    def registrar_compra(self):
        kg = num(self.e_kg.get())
        if kg <= 0:
            messagebox.showwarning("Atenção", "Digite a quantidade em kg."); return
        with conn() as c:
            c.execute("INSERT INTO compras(data,kg,custo) VALUES(?,?,?)",
                      (date.today().isoformat(), kg, num(self.e_custo.get())))
        self.e_kg.delete(0, "end"); self.e_custo.delete(0, "end")
        self.atualizar_racao()
        messagebox.showinfo("Salvo ✅", "Compra de ração registrada!")

    # ------------------------------------------------ relatório
    def montar_rel(self):
        f = self.tab_rel
        top = ttk.Frame(f); top.pack(fill="x", padx=12, pady=(8, 6))
        ttk.Label(top, text="Mês:").pack(side="left")
        self.cb_mes = ttk.Combobox(top, values=[calendar.month_name[m].capitalize() for m in range(1, 13)],
                                   state="readonly", width=14, font=FONTE)
        self.cb_mes.current(date.today().month - 1); self.cb_mes.pack(side="left", padx=6)
        self.cb_ano = ttk.Combobox(top, values=[str(a) for a in range(date.today().year - 4,
                                                                      date.today().year + 1)],
                                   state="readonly", width=7, font=FONTE)
        self.cb_ano.set(str(date.today().year)); self.cb_ano.pack(side="left", padx=6)
        ttk.Button(top, text="↻ Atualizar", style="Suave.TButton",
                   command=self.gerar_rel).pack(side="left", padx=6)
        ttk.Button(top, text="📊  Exportar planilha Excel", style="Acao.TButton",
                   command=self.exportar_xlsx).pack(side="right")

        self.lb_resumo = tk.Text(f, height=16, font=("Consolas", 11), bg=CARD, fg=TEXTO,
                                 highlightthickness=1, highlightbackground=BORDA,
                                 relief="flat", padx=16, pady=12)
        self.lb_resumo.tag_configure("verde", foreground=VERDE, font=("Consolas", 12, "bold"))
        self.lb_resumo.tag_configure("vermelho", foreground=VERMELHO, font=("Consolas", 12, "bold"))
        self.lb_resumo.tag_configure("titulo", foreground=VERDE_ESCURO, font=("Consolas", 12, "bold"))
        self.lb_resumo.pack(fill="both", expand=True, padx=12, pady=8)

    def _mes(self):
        mes = self.cb_mes.current() + 1
        return "%s-%02d" % (self.cb_ano.get(), mes)

    def _totais_mes(self, pref):
        with conn() as c:
            r = c.execute("""SELECT COUNT(*), COALESCE(SUM(ovos),0), COALESCE(SUM(vendidos),0),
                             COALESCE(SUM(vendidos*preco_venda),0), COALESCE(SUM(quebrados),0),
                             COALESCE(SUM(mortas),0), COALESCE(SUM(racao_kg),0),
                             COALESCE(SUM(custo_racao),0), COALESCE(SUM(outros),0)
                             FROM dias WHERE data LIKE ?""", (pref + "%",)).fetchone()
            galinhas = num(get_conf(c, "galinhas", "0"))
        (dias, ovos, vendidos, receita, quebrados, mortas, kg_racao,
         custo_racao, outros) = r
        return dict(dias=dias, ovos=ovos, vendidos=vendidos, receita=receita,
                    quebrados=quebrados, mortas=mortas, kg_racao=kg_racao,
                    custo_racao=custo_racao, outros=outros,
                    gastos=custo_racao + outros, galinhas=galinhas)

    def gerar_rel(self):
        pref = self._mes()
        t = self._totais_mes(pref)
        lucro = t["receita"] - t["gastos"]
        custo_ovo = (t["gastos"] / t["ovos"]) if t["ovos"] else 0
        estoque = self.estoque_atual()
        self.lb_resumo.delete("1.0", "end")
        txt = self.lb_resumo
        txt.insert("end", "        RESUMO DE %s/%s\n" % (pref[5:], pref[:4]), "titulo")
        txt.insert("end", "─" * 46 + "\n")
        txt.insert("end", "Dias registrados:        %d\n" % t["dias"])
        txt.insert("end", "Ovos coletados:          %d\n" % t["ovos"])
        txt.insert("end", "Dúzias vendidas:         %.1f (%d ovos)\n" % (t["vendidos"], round(t["vendidos"] * 12)))
        txt.insert("end", "Ovos quebrados:          %d\n" % t["quebrados"])
        txt.insert("end", "Galinhas mortas:         %d\n" % t["mortas"])
        txt.insert("end", "Ração usada:             %.1f kg\n" % t["kg_racao"])
        txt.insert("end", "─" * 46 + "\n")
        txt.insert("end", "Receita de vendas:       %s\n" % brl(t["receita"]))
        txt.insert("end", "Gastos (ração + outros): %s\n" % brl(t["gastos"]))
        txt.insert("end", "\n")
        sinal = "✅ LUCRO DO MÊS" if lucro >= 0 else "❌ PREJUÍZO DO MÊS"
        txt.insert("end", "%s: %s\n" % (sinal, brl(lucro)), "verde" if lucro >= 0 else "vermelho")
        txt.insert("end", "\n")
        txt.insert("end", "Custo por ovo produzido: %s\n" % brl(custo_ovo))
        txt.insert("end", "Estoque de ração:        %.1f kg\n" % max(estoque, 0))
        if t["galinhas"] and t["ovos"] and t["dias"]:
            txt.insert("end", "Média por galinha/dia:   %.2f ovos\n" % (t["ovos"] / (t["galinhas"] * t["dias"])))

    def exportar_xlsx(self):
        pref = self._mes()
        arq = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                           initialfile="Granja %s-%s.xlsx" % (pref[5:], pref[:4]),
                                           filetypes=[("Planilha Excel", "*.xlsx")])
        if not arq:
            return
        wb = openpyxl.Workbook()

        # ---------- folha de resumo
        res = wb.active; res.title = "Resumo"
        t = self._totais_mes(pref)
        lucro = t["receita"] - t["gastos"]
        verde_f = PatternFill("solid", fgColor="1F7A4D")
        branco_b = Font(bold=True, color="FFFFFF", size=12)
        res["A1"] = "🐔 PAINEL DA GRANJA — RESUMO DE %s/%s" % (pref[5:], pref[:4])
        res["A1"].font = Font(bold=True, size=14, color="143D2B")
        res.merge_cells("A1:C1")
        dados = [
            ("Dias registrados", t["dias"], None),
            ("Ovos coletados", t["ovos"], None),
            ("Dúzias vendidas", t["vendidos"], None),
            ("Ovos vendidos", round(t["vendidos"] * 12), None),
            ("Ovos quebrados", t["quebrados"], None),
            ("Galinhas mortas", t["mortas"], None),
            ("Ração usada (kg)", round(t["kg_racao"], 1), None),
            ("Receita de vendas", t["receita"], 'R$ #,##0.00'),
            ("Gastos com ração", t["custo_racao"], 'R$ #,##0.00'),
            ("Outros gastos", t["outros"], 'R$ #,##0.00'),
            ("Gastos totais", t["gastos"], 'R$ #,##0.00'),
            ("LUCRO / PREJUÍZO", lucro, 'R$ #,##0.00'),
            ("Custo por ovo", (t["gastos"] / t["ovos"]) if t["ovos"] else 0, 'R$ #,##0.00'),
            ("Estoque de ração (kg)", round(self.estoque_atual(), 1), None),
        ]
        linha = 3
        for rot, val, fmt in dados:
            res.cell(linha, 1, rot).font = Font(bold=(linha == 15))
            c2 = res.cell(linha, 2, val)
            if fmt: c2.number_format = fmt
            if linha == 15:
                c2.font = Font(bold=True, color=("1F7A4D" if lucro >= 0 else "C62828"))
                res.cell(linha, 1).font = Font(bold=True, color=("1F7A4D" if lucro >= 0 else "C62828"))
            linha += 1
        res.column_dimensions["A"].width = 24
        res.column_dimensions["B"].width = 16

        # ---------- folha do registro diário
        reg = wb.create_sheet("Registro do dia")
        cab = ["Data", "Ovos coletados", "Dúzias vendidas", "Preço da dúzia (R$)",
               "Receita (R$)", "Quebrados", "Mortas", "Ração (kg)", "Custo ração (R$)",
               "Outros gastos (R$)", "Observações"]
        for j, txt in enumerate(cab, 1):
            c = reg.cell(1, j, txt)
            c.fill = verde_f; c.font = branco_b
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        reg.freeze_panes = "A2"
        linha = 2
        with conn() as cdb:
            for r in cdb.execute("SELECT * FROM dias WHERE data LIKE ? ORDER BY data", (pref + "%",)):
                reg.cell(linha, 1, datetime.fromisoformat(r[0]).strftime("%d/%m/%Y"))
                reg.cell(linha, 2, r[1]); reg.cell(linha, 3, r[2])
                reg.cell(linha, 4, r[3]); reg.cell(linha, 5, round(r[2] * r[3], 2))
                reg.cell(linha, 6, r[4]); reg.cell(linha, 7, r[5])
                reg.cell(linha, 8, r[6]); reg.cell(linha, 9, r[7])
                reg.cell(linha, 10, r[8]); reg.cell(linha, 11, r[9] or "")
                for col in (4, 5, 9, 10):
                    reg.cell(linha, col).number_format = 'R$ #,##0.00'
                linha += 1
        larguras = [12, 14, 13, 16, 13, 11, 9, 11, 15, 16, 40]
        for j, w in enumerate(larguras, 1):
            reg.column_dimensions[get_column_letter(j)].width = w
        wb.save(arq)
        try:
            os.startfile(arq)  # já abre a planilha para o usuário
        except OSError:
            pass
        messagebox.showinfo("Exportado ✅", "Planilha salva e aberta:\n" + arq)

    # ------------------------------------------------ clima
    def montar_clima(self):
        f = self.tab_clima
        top = ttk.Frame(f); top.pack(fill="x", padx=12, pady=(8, 6))
        ttk.Label(top, text="Sua cidade:").pack(side="left")
        self.e_cidade = ttk.Entry(top, width=24, style="Card.TEntry", font=FONTE_GRANDE)
        with conn() as c:
            self.e_cidade.insert(0, get_conf(c, "cidade", "São Paulo"))
        self.e_cidade.pack(side="left", padx=6)
        ttk.Button(top, text="🔎  Ver previsão", style="Acao.TButton",
                   command=self.ver_clima).pack(side="left", padx=6)
        self.lb_clima = tk.Label(f, bg=CARD, fg=TEXTO, font=("Segoe UI", 11),
                                 justify="left", anchor="nw",
                                 highlightthickness=1, highlightbackground=BORDA)
        self.lb_clima.pack(fill="both", expand=True, padx=12, pady=8)

    def ver_clima(self):
        cidade = self.e_cidade.get().strip()
        with conn() as c:
            set_conf(c, "cidade", cidade)
        self.lb_clima.config(text="  Buscando... 🌐")
        self.update()
        try:
            fc = buscar_clima(cidade)
        except Exception as e:
            self.lb_clima.config(text="  Não consegui buscar o clima (sem internet?)\n  %s" % e)
            return
        d = fc["daily"]
        linhas = ["\n  Previsão para os próximos dias em %s:\n" % cidade.title()]
        for i in range(len(d["time"])):
            tmax, tmin = d["temperature_2m_max"][i], d["temperature_2m_min"][i]
            chuva = d["precipitation_sum"][i] or 0
            dia = datetime.fromisoformat(d["time"][i]).strftime("%a %d/%m")
            aviso = ""
            if tmin < 5:
                aviso += "  ⚠️ FRIO! Ligue aquecedor!"
            if tmax > 33:
                aviso += "  ⚠️ CALOR! Muita água e sombra!"
            if chuva > 15:
                aviso += "  🌧️ Chuva forte"
            linhas.append("  %s   %2.0f°C a %2.0f°C   chuva %.0fmm%s" % (dia, tmin, tmax, chuva, aviso))
        self.lb_clima.config(text="\n".join(linhas))

    # ------------------------------------------------ ajustes
    def montar_conf(self):
        f = self.tab_conf
        box = ttk.LabelFrame(f, text="  Minha granja  ", style="Card.TLabelframe")
        box.pack(fill="x", padx=12, pady=10)
        linha1 = ttk.Frame(box, style="Card.TFrame"); linha1.pack(fill="x", padx=14, pady=(12, 8))
        ttk.Label(linha1, text="Quantidade de galinhas:", style="Card.TLabel").pack(side="left")
        self.e_galinhas = ttk.Entry(linha1, width=8, style="Card.TEntry", font=FONTE_GRANDE)
        self.e_galinhas.insert(0, get_conf(conn(), "galinhas", "0"))
        self.e_galinhas.pack(side="left", padx=8)
        linha2 = ttk.Frame(box, style="Card.TFrame"); linha2.pack(fill="x", padx=14, pady=(0, 12))
        ttk.Label(linha2, text="Ração que já tinha em estoque (kg):", style="Card.TLabel").pack(side="left")
        self.e_inicial = ttk.Entry(linha2, width=8, style="Card.TEntry", font=FONTE_GRANDE)
        self.e_inicial.insert(0, get_conf(conn(), "estoque_inicial", "0"))
        self.e_inicial.pack(side="left", padx=8)
        ttk.Button(f, text="💾  Salvar ajustes", style="Acao.TButton",
                   command=self.salvar_conf).pack(padx=12, pady=6, fill="x")
        ttk.Button(f, text="📂  Abrir a pasta dos dados", style="Suave.TButton",
                   command=lambda: os.startfile(DATA_DIR)).pack(padx=12, pady=6, fill="x")
        ttk.Label(f, text="Ao desinstalar o programa, seus dados NÃO são apagados.",
                  foreground=CINZA).pack(padx=14, pady=10, anchor="w")

    def salvar_conf(self):
        with conn() as c:
            set_conf(c, "galinhas", num(self.e_galinhas.get()))
            set_conf(c, "estoque_inicial", num(self.e_inicial.get()))
        self.atualizar_racao()
        self.gerar_rel()
        messagebox.showinfo("Salvo ✅", "Ajustes salvos!")

# ----------------------------------------------------------------
if __name__ == "__main__":
    App().mainloop()
