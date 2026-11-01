#!/usr/bin/env python3
"""
Assistente de Compras com IA
Estoque do WMS + vendas da planilha + pedidos em aberto -> projeção de ruptura e recomendação de compra.

pip install -r requirements.txt
Configuração em .env (modelo em .env.example) e catálogo em itens.json (modelo em itens.exemplo.json).
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk
import threading, requests, json, csv, io, os, webbrowser, calendar, re
from datetime import datetime, date
from collections import defaultdict
from anthropic import Anthropic
try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side, GradientFill
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ──────────────────────────────────────────────────────────────
# CONFIGURAÇÕES
# ──────────────────────────────────────────────────────────────
def _carregar_env(caminho=".env"):
    """Lê um .env simples (CHAVE=valor) sem depender de bibliotecas extras."""
    if not os.path.exists(caminho):
        return
    with open(caminho, encoding="utf-8") as f:
        for linha in f:
            linha = linha.strip()
            if linha and not linha.startswith("#") and "=" in linha:
                chave, valor = linha.split("=", 1)
                os.environ.setdefault(chave.strip(), valor.strip())


_carregar_env()

# WMS (Stokki): endpoint de saldo por SKU
STOKKI_URL   = os.environ.get("STOKKI_URL", "https://SUA_EMPRESA.stokki.com.br/api/stokki/oms/inventory/show")
STOKKI_TOKEN = os.environ.get("STOKKI_TOKEN", "")
# Planilha de vendas (Google Sheets). Use uma planilha que NÃO tenha dados sensíveis
# ou acesse via conta de serviço: a leitura por CSV exige compartilhamento por link.
SHEET_ID     = os.environ.get("SHEET_ID", "")
SHEET_NAME   = os.environ.get("SHEET_NAME", "fVendas")
DATA_FILE    = os.environ.get("DATA_FILE", "compras_data.json")
SEGMENTO     = os.environ.get("SEGMENTO", "lâminas de aço silício para transformadores")


def _carregar_itens():
    """Catálogo de SKUs monitorados: itens.json (seu) ou itens.exemplo.json (fictício)."""
    for nome in ("itens.json", "itens.exemplo.json"):
        caminho = os.path.join(os.path.dirname(os.path.abspath(__file__)), nome)
        if os.path.exists(caminho):
            with open(caminho, encoding="utf-8") as f:
                return json.load(f)
    return []


ITENS = _carregar_itens()

def _sku_num_key(sku):
    """Extrai o primeiro número do SKU para ordenação numérica correta (EI-35 < EI-41 < EI-105)."""
    m = re.search(r'\d+', sku)
    return int(m.group()) if m else 0

SKU_LIST    = sorted([i["sku"] for i in ITENS], key=_sku_num_key)
CAT_ORDEM   = {"MONOFÁSICA":1,"TRIFÁSICA":2,"CARRETEL":3,"UI":4,"BORNE":5}
CAT_COLORS  = {
    "MONOFÁSICA":{"badge":"#FF6B35","row":"#FFF4F0","alt":"#FFE4D6"},
    "TRIFÁSICA": {"badge":"#F72585","row":"#FFF0F7","alt":"#FFD6E8"},
    "CARRETEL":  {"badge":"#3B82F6","row":"#F0F6FF","alt":"#D9E9FF"},
    "UI":        {"badge":"#8B5CF6","row":"#F7F3FF","alt":"#EEE5FF"},
    "BORNE":     {"badge":"#10B981","row":"#F0FDF7","alt":"#D6F6E8"},
}
SKU_CAT = {i["sku"]: i["cat"] for i in ITENS}

MESES_PT = ["","Janeiro","Fevereiro","Março","Abril","Maio","Junho",
            "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
MESES_PT_CURTO = ["","Jan","Fev","Mar","Abr","Mai","Jun",
                  "Jul","Ago","Set","Out","Nov","Dez"]

BG_APP     = "#F8FAFB"
BG_HEADER  = "#0F172A"
BG_CARD    = "#FFFFFF"
BG_SIDEBAR = "#F3F4F6"
CLR_ACCENT = "#3B82F6"
CLR_GREEN  = "#10B981"
CLR_PURPLE = "#8B5CF6"
CLR_RED    = "#EF4444"
CLR_AMBER  = "#F59E0B"
CLR_TEXT   = "#111827"
CLR_MUTED  = "#6B7280"
CLR_BORDER = "#E5E7EB"
CLR_TABLE_HEAD = "#1F2937"

# ──────────────────────────────────────────────────────────────
# FORMATAÇÃO NUMÉRICA BR
# ──────────────────────────────────────────────────────────────
def fmt_num(v, decimals=0):
    try:
        n = float(v)
    except (ValueError, TypeError):
        return "—" if v in (None, "") else str(v)
    if n == 0 and decimals == 0:
        return "0"
    if decimals == 0:
        inteiro = int(round(n))
        s = f"{inteiro:,}".replace(",", ".")
        return s
    else:
        s = f"{n:,.{decimals}f}"
        s = s.replace(",", "X").replace(".", ",").replace("X", ".")
        return s

def fmt_kg(v):
    return fmt_num(v, 1)

# ──────────────────────────────────────────────────────────────
# FUNÇÕES DE DADOS
# ──────────────────────────────────────────────────────────────
def consultar_sku(sku):
    try:
        r = requests.post(STOKKI_URL,
            json={"token": STOKKI_TOKEN, "sku": sku.strip()},
            headers={"X-HTTP-Method-Override":"GET"}, timeout=10)
        if r.status_code == 200:
            d = r.json()
            if d:
                p = d[0].get("product", {})
                return {"disp": p.get("quantity_available",0),
                        "reserva": p.get("quantity_committed",0)}
    except Exception:
        pass
    return {"disp":0,"reserva":0}

def buscar_vendas():
    url = (f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
           f"/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}")
    vendas = {}
    hoje = date.today()
    try:
        r = requests.get(url, timeout=15); r.raise_for_status()
        reader = csv.DictReader(io.StringIO(r.text))
        headers = [h.strip() for h in (reader.fieldnames or [])]

        # detecta coluna de data
        date_key = None
        for h in headers:
            if any(x in h.upper() for x in ["DATA","DATE","DT_","MÊS","MES","COMPETÊNCIA","COMPETENCIA","PERÍODO","PERIODO"]):
                date_key = h
                break

        for row in reader:
            # filtro de mês atual pela coluna de data (se existir)
            if date_key:
                dt = None
                for fmt in ("%d/%m/%Y","%d/%m/%y","%Y-%m-%d","%m/%Y","%Y-%m"):
                    try: dt = datetime.strptime(row.get(date_key,"").strip(), fmt).date(); break
                    except: pass
                if dt is None or dt.month != hoje.month or dt.year != hoje.year:
                    continue

            norm = {k.strip().upper(): v.strip() for k,v in row.items() if k and v}
            sku = None
            for key in norm.keys():
                if "PRODUTO" in key:
                    sku = norm[key].upper()
                    break
            if not sku: continue

            quantidade_str = None
            for key in norm.keys():
                if "QUANTIDADE" in key or "TOTAL" in key:
                    quantidade_str = norm[key]
                    break
            if not quantidade_str: continue

            try:
                quantidade = float(quantidade_str.replace(",", "."))
                vendas[sku] = vendas.get(sku, 0) + quantidade
            except: pass
    except Exception as e:
        print(f"[Sheets] Erro ao buscar vendas: {e}")
    return vendas

def vendas_mes_atual(sku, vendas):
    if sku not in vendas: return 0
    return float(vendas[sku])

def carregar_local():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE,"r",encoding="utf-8") as f:
            return json.load(f)
    return {"config":{}, "pedidos":[]}

def salvar_local(d):
    with open(DATA_FILE,"w",encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)

def necessidade_calc(disp, proj_venda):
    try:
        n = max(0, float(proj_venda) - float(disp))
        return int(n) if n > 0 else 0
    except: return 0

def parse_date(s):
    for fmt in ("%d/%m/%Y","%d/%m/%y","%Y-%m-%d"):
        try: return datetime.strptime(s.strip(), fmt).date()
        except: pass
    return None

def estoque_projetado(sku, disp, vendas_mes, pedidos, meses=12,
                      proj_venda="", proj_mensal=None):
    hoje = date.today()
    resultado = []
    saldo = float(disp) if disp not in ("—","") else 0
    venda_padrao = float(vendas_mes) if isinstance(vendas_mes, (int, float)) else 0
    try:
        venda_global = float(proj_venda) if str(proj_venda).strip() not in ("","0") else None
    except (ValueError, TypeError):
        venda_global = None

    for i in range(meses):
        m = (hoje.month - 1 + i) % 12 + 1
        a = hoje.year + (hoje.month - 1 + i) // 12
        key = f"{a}-{m:02d}"

        if proj_mensal and key in proj_mensal and str(proj_mensal[key]).strip():
            try:    venda_m = float(proj_mensal[key])
            except: venda_m = venda_global if venda_global is not None else 0
        elif venda_global is not None:
            venda_m = venda_global
        else:
            venda_m = 0

        # Mês atual (i == 0): o 'disp' do Stokki já reflete todas as vendas reais
        # e movimentações que ocorreram no mês. Não projetamos vendas adicionais
        # porque o estoque disponível já É o saldo real atual.
        if i == 0:
            venda_m = 0

        # Chegadas: Soma de pedidos pendentes para este mês. 
        # No mês atual (i == 0), varre e agrega os pedidos atrasados (meses anteriores).
        chegadas = 0.0
        for p in pedidos:
            if p.get("sku", "").upper() == sku.upper() and not p.get("chegado", False):
                dt = parse_date(p.get("data", "") or "")
                if dt:
                    qty = float(p.get("qty", 0))
                    # Pertence a este mês?
                    if dt.year == a and dt.month == m:
                        chegadas += qty
                    # Pertence ao passado, mas ainda não chegou? Acumula no mês atual (i==0)
                    elif i == 0 and (dt.year < a or (dt.year == a and dt.month < m)):
                        chegadas += qty

        saldo = saldo + chegadas - venda_m
        resultado.append((f"{MESES_PT[m]}/{a}", saldo, chegadas, m, a))
    return resultado

# ──────────────────────────────────────────────────────────────
# WIDGET KPI E INPUTS
# ──────────────────────────────────────────────────────────────
class KpiCard(ctk.CTkFrame):
    def __init__(self, parent, title, value, icon, color, **kw):
        super().__init__(parent, fg_color=BG_CARD, corner_radius=12, border_width=1, border_color=CLR_BORDER, **kw)
        bar = ctk.CTkFrame(self, fg_color=color, corner_radius=0, width=5)
        bar.pack(side="left", fill="y", padx=(0,0), pady=8)
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=(12,14), pady=10)
        top = ctk.CTkFrame(content, fg_color="transparent")
        top.pack(fill="x", anchor="w")
        ctk.CTkLabel(top, text=icon, font=ctk.CTkFont("Segoe UI",18)).pack(side="left", padx=(0,8))
        ctk.CTkLabel(top, text=title, font=ctk.CTkFont("Segoe UI",11, "bold"), text_color=CLR_MUTED).pack(side="left", fill="x")
        self.lbl = ctk.CTkLabel(content, text=value, font=ctk.CTkFont("Segoe UI",32,"bold"), text_color=color, anchor="w")
        self.lbl.pack(fill="x", pady=(6,0))

    def set(self, v):
        self.lbl.configure(text=str(v))

class ProjecaoVendaInput(ctk.CTkFrame):
    def __init__(self, parent, sku, valor_inicial="", callback=None, **kw):
        super().__init__(parent, fg_color=BG_CARD, corner_radius=14, border_width=2, border_color=CLR_ACCENT, **kw)
        self.sku = sku
        self.callback = callback
        self.valor = tk.StringVar(value=str(valor_inicial) if valor_inicial else "")
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=16, pady=14)
        header = ctk.CTkFrame(main, fg_color="transparent")
        header.pack(fill="x", pady=(0,12))
        ctk.CTkLabel(header, text="📊", font=ctk.CTkFont("Segoe UI",22)).pack(side="left", padx=(0,8))
        ctk.CTkLabel(header, text="Projeção de Venda (kg/mês)", font=ctk.CTkFont("Segoe UI",13,"bold"), text_color=CLR_TEXT).pack(side="left", fill="x", expand=True)
        input_frame = ctk.CTkFrame(main, fg_color="transparent")
        input_frame.pack(fill="x", pady=(0,4))
        self.entrada = ctk.CTkEntry(input_frame, textvariable=self.valor, placeholder_text="Digite o valor em kg", font=ctk.CTkFont("Segoe UI",14, "bold"), fg_color="#F9FAFB", border_color=CLR_ACCENT, border_width=2, corner_radius=10, height=48)
        self.entrada.pack(side="left", fill="both", expand=True, padx=(0,8))
        self.btn_save = ctk.CTkButton(input_frame, text="💾 Salvar", command=self._on_change, fg_color=CLR_GREEN, hover_color="#059669", corner_radius=10, height=48, width=100, font=ctk.CTkFont("Segoe UI",12,"bold"))
        self.btn_save.pack(side="left")
        info = ctk.CTkLabel(main, text="Pressione ENTER ou clique Salvar para atualizar a projeção", font=ctk.CTkFont("Segoe UI",10), text_color=CLR_MUTED)
        info.pack(anchor="w", pady=(6,0))
        self.entrada.bind("<Return>", self._on_change)
        self.entrada.bind("<FocusOut>", self._on_change)
    
    def _on_change(self, event=None):
        if self.callback: self.callback(self.sku, self.valor.get())
        self.btn_save.configure(text="✅ Salvo!")
        self.after(1500, lambda: self.btn_save.configure(text="💾 Salvar"))

# ──────────────────────────────────────────────────────────────
# GRADE VISUAL (Canvas) - DINÂMICA PARA ESTOQUE E CENÁRIO
# ──────────────────────────────────────────────────────────────
class GradeSimulacao(tk.Canvas):
    COL_W = 115
    ROW_H = 96
    SKU_W = 150
    HDR_H = 46
    CAT_H = 26
    GAP   = 2

    def __init__(self, parent, **kw):
        kw.setdefault("bg", "#FFFFFF")
        kw.setdefault("highlightthickness", 0)
        super().__init__(parent, **kw)

    def render(self, rows, meses, modo="cenario"):
        self.delete("all")
        nc     = len(meses)
        ROW_H  = 95 if modo == "estoque" else self.ROW_H
        tot_w  = self.SKU_W + self.GAP + nc * (self.COL_W + self.GAP)
        y      = 0
        HDR_BG = "#0B192C"
        CAT_BG = "#1E3A8A"

        self._rect(0, 0, self.SKU_W, self.HDR_H, HDR_BG)
        self.create_text(self.SKU_W // 2, self.HDR_H // 2, text="SKU", fill="white", font=("Segoe UI", 10, "bold"))
                         
        for ci, (_, _, label) in enumerate(meses):
            x = self.SKU_W + self.GAP + ci * (self.COL_W + self.GAP)
            self._rect(x, 0, self.COL_W, self.HDR_H, HDR_BG)
            self.create_text(x + self.COL_W // 2, self.HDR_H // 2, text=label, fill="white", font=("Segoe UI", 9, "bold"), justify="center")
        y += self.HDR_H + self.GAP

        for row in rows:
            if row["type"] == "cat":
                self._rect(0, y, tot_w, self.CAT_H, CAT_BG)
                self.create_text(10, y + self.CAT_H // 2, text=row['label'].upper(), fill="white", font=("Segoe UI", 10, "bold"), anchor="w")
                y += self.CAT_H + self.GAP
                continue

            self._rect(0, y, self.SKU_W, ROW_H, "#F8FAFB")
            self.create_text(8, y + ROW_H // 2, text=row["sku"], fill="#111827", font=("Segoe UI", 9), anchor="w")

            for ci, md in enumerate(row["months"]):
                x = self.SKU_W + self.GAP + ci * (self.COL_W + self.GAP)
                atual = md["atual"]
                cen   = md["cen"]
                cheg  = md["chegada"]

                if modo == "cenario":
                    pv         = md.get("proj_venda", 0)
                    cheg_real  = md.get("chegada_real", 0)
                    cheg_sim   = md.get("chegada_sim", 0)
                    bg = "#FCE4E4" if cen < 0 else "#D1F2D3"
                    self._rect(x, y, self.COL_W, ROW_H, bg)

                    cy = y + 7
                    # proj venda
                    pv_txt = f"Venda: {fmt_num(pv)}" if pv else "Venda: —"
                    self.create_text(x + self.COL_W // 2, cy, text=pv_txt, fill="#6B21A8", font=("Segoe UI", 8, "bold"), anchor="n")
                    cy += 14
                    # chegada real (pedido confirmado)
                    if cheg_real > 0:
                        self.create_text(x + self.COL_W // 2, cy, text=f"📦 {fmt_num(cheg_real)}", fill="#065F46", font=("Segoe UI", 8, "bold"), anchor="n")
                        cy += 13
                    # chegada simulada (cenário)
                    if cheg_sim > 0:
                        self.create_text(x + self.COL_W // 2, cy, text=f"▲ {fmt_num(cheg_sim)}", fill="#0369A1", font=("Segoe UI", 8, "bold"), anchor="n")
                        cy += 13
                    # saldo atual
                    fg_a = "#C62828" if atual < 0 else "#374151"
                    self.create_text(x + self.COL_W // 2, cy, text=f"Atual: {fmt_num(atual)}", fill=fg_a, font=("Segoe UI", 8), anchor="n")
                    cy += 14
                    # saldo cenário
                    fg_c = "#C62828" if cen < 0 else "#1D4ED8"
                    self.create_text(x + self.COL_W // 2, cy, text=f"Cen.: {fmt_num(cen)}", fill=fg_c, font=("Segoe UI", 9, "bold"), anchor="n")

                else:
                    bg = "#FCE4E4" if atual < 0 else "#D1F2D3"
                    self._rect(x, y, self.COL_W, ROW_H, bg)

                    cy = y + 6
                    # proj venda
                    pv = md.get("proj_venda", 0)
                    pv_txt = f"Proj: {fmt_num(pv)}" if pv else "Proj: —"
                    self.create_text(x + self.COL_W // 2, cy, text=pv_txt, fill="#6B21A8", font=("Segoe UI", 8, "bold"), anchor="n")
                    cy += 14
                    # chegadas (sempre visível)
                    cheg_txt = f"📦 {fmt_num(cheg)}" if cheg > 0 else "📦 —"
                    cheg_cor = "#065F46" if cheg > 0 else "#9CA3AF"
                    self.create_text(x + self.COL_W // 2, cy, text=cheg_txt, fill=cheg_cor, font=("Segoe UI", 8, "bold"), anchor="n")
                    cy += 14
                    # saldo
                    fg_a = "#C62828" if atual < 0 else "#111827"
                    self.create_text(x + self.COL_W // 2, cy, text=f"Saldo: {fmt_num(atual)}", fill=fg_a, font=("Segoe UI", 10, "bold"), anchor="n")

            y += ROW_H + self.GAP

        self.configure(scrollregion=(0, 0, tot_w, y))

    def _rect(self, x, y, w, h, fill):
        self.create_rectangle(x, y, x + w, y + h, fill=fill, outline=fill)

# ──────────────────────────────────────────────────────────────
# APP PRINCIPAL
# ──────────────────────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.title("Assistente de Compras")
        self.geometry("1640x980")
        self.minsize(1300,750)
        self.configure(fg_color=BG_APP)

        d = carregar_local()
        if "pedidos" not in d: d["pedidos"] = []
        if "config"  not in d: d["config"]  = {}
        self.local   = d
        self.estoque = {}
        self.vendas = {}
        self.cat_var = tk.StringVar(value="TODAS")
        self.busca_v = tk.StringVar()
        self.busca_v.trace_add("write", lambda *_: self._popular_estoque())
        self._loading = False
        self._ia_open = False
        self.cenario  = []
        # ajuste de demanda do cenário: {sku, mult, meses:[keys "YYYY-MM"]}
        self.cen_demanda = {"sku": "TODOS", "mult": 1.0, "meses": []}

        self._build_ui()
        self._popular_estoque()
        self._popular_pedidos()

        if not os.environ.get("ANTHROPIC_API_KEY"):
            self.btn_ia.configure(state="disabled", fg_color="#E5E7EB", text_color="#9CA3AF", text="🤖 IA (configure ANTHROPIC_API_KEY)")

    def _build_ui(self):
        hdr = ctk.CTkFrame(self, fg_color=BG_HEADER, corner_radius=0, height=64)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        lft = ctk.CTkFrame(hdr, fg_color="transparent")
        lft.pack(side="left", padx=28, pady=12)
        ctk.CTkLabel(lft, text="⚡", font=ctk.CTkFont("Segoe UI",28)).pack(side="left",padx=(0,10))
        txt_frame = ctk.CTkFrame(lft, fg_color="transparent")
        txt_frame.pack(side="left")
        ctk.CTkLabel(txt_frame, text="COMPRAS", font=ctk.CTkFont("Segoe UI",18,"bold"), text_color="#FFFFFF").pack(anchor="w")
        ctk.CTkLabel(txt_frame, text="Assistente de Compras Inteligente", font=ctk.CTkFont("Segoe UI",11), text_color="#A1A5B3").pack(anchor="w", pady=(2,0))
        
        rgt = ctk.CTkFrame(hdr, fg_color="transparent")
        rgt.pack(side="right", padx=28, pady=12)
        self.lbl_status = ctk.CTkLabel(rgt, text="", font=ctk.CTkFont("Segoe UI",11,"bold"), text_color=CLR_GREEN)
        self.lbl_status.pack()
        self.lbl_sync = ctk.CTkLabel(rgt, text="Nunca sincronizado", font=ctk.CTkFont("Segoe UI",9), text_color="#6B7280")
        self.lbl_sync.pack(pady=(2,0))

        self.prog = ctk.CTkProgressBar(self, mode="indeterminate", height=2, fg_color=BG_APP, progress_color=CLR_ACCENT)

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=16)

        self.kpi_pedidos = None
        self.kpi_chegada = None

        self._build_alerta_ruptura(body)

        self.tabs = ctk.CTkTabview(body, fg_color=BG_CARD, segmented_button_fg_color="#F3F4F6", segmented_button_selected_color=CLR_ACCENT, segmented_button_selected_hover_color="#2563EB", segmented_button_unselected_color="#FFFFFF", segmented_button_unselected_hover_color="#E5E7EB", text_color=CLR_TEXT, corner_radius=14, border_width=1, border_color=CLR_BORDER)
        self.tabs.pack(fill="both", expand=True)
        self.tabs.add("📊 Estoque & Necessidade")
        self.tabs.add("📦 Compras & Chegadas")
        self.tabs.add("🔮 Cenário")
        self.tabs.add("🗺️ Visual Atual")

        self._build_tab_estoque(self.tabs.tab("📊 Estoque & Necessidade"))
        self._build_tab_compras(self.tabs.tab("📦 Compras & Chegadas"))
        self._build_tab_cenario(self.tabs.tab("🔮 Cenário"))
        self._build_tab_visual_atual(self.tabs.tab("🗺️ Visual Atual"))

        self.tabs.configure(command=self._on_tab_change)

        self._build_ia_panel(body)

    # ══════════════════════════════════════════════════════════
    # BANNER DE ALERTAS PROATIVOS DE RUPTURA
    # ══════════════════════════════════════════════════════════
    def _build_alerta_ruptura(self, parent):
        self._alertas_expandido = False
        self.alerta_outer = ctk.CTkFrame(parent, fg_color="transparent")
        self.alerta_banner = ctk.CTkFrame(self.alerta_outer, corner_radius=10, border_width=1, fg_color="#FEE2E2", border_color="#EF4444")
        self.alerta_banner.pack(fill="x")

        hdr = ctk.CTkFrame(self.alerta_banner, fg_color="transparent")
        hdr.pack(fill="x", padx=12, pady=8)
        ctk.CTkLabel(hdr, text="⚠️", font=ctk.CTkFont("Segoe UI", 14)).pack(side="left", padx=(0, 8))
        self.alerta_titulo_lbl = ctk.CTkLabel(hdr, text="", font=ctk.CTkFont("Segoe UI", 12, "bold"), text_color="#991B1B")
        self.alerta_titulo_lbl.pack(side="left")
        self.alerta_toggle_btn = ctk.CTkButton(hdr, text="▼ Detalhes", width=90, height=26,
                                               corner_radius=6, font=ctk.CTkFont("Segoe UI", 10),
                                               fg_color="#EF4444", hover_color="#DC2626", text_color="white",
                                               command=self._toggle_alertas)
        self.alerta_toggle_btn.pack(side="right")

        self.alerta_pills = ctk.CTkFrame(self.alerta_banner, fg_color="transparent")
        self.alerta_pills.pack(fill="x", padx=12, pady=(0, 8))

        self.alerta_detalhe = ctk.CTkScrollableFrame(self.alerta_banner, fg_color="transparent", height=130)

    def _toggle_alertas(self):
        self._alertas_expandido = not self._alertas_expandido
        if self._alertas_expandido:
            self.alerta_detalhe.pack(fill="x", padx=8, pady=(0, 8))
            self.alerta_toggle_btn.configure(text="▲ Ocultar")
        else:
            self.alerta_detalhe.pack_forget()
            self.alerta_toggle_btn.configure(text="▼ Detalhes")

    def _popular_alertas(self):
        pedidos = self.local["pedidos"]
        alertas = []

        for it in ITENS:
            sku, cat = it["sku"], it["cat"]
            e    = self.estoque.get(sku, {})
            disp = e.get("disp", 0)
            vend = vendas_mes_atual(sku, self.vendas)
            er   = self.local["config"].get(sku, {}).get("proj_venda", "")
            pm   = self.local["config"].get(sku, {}).get("proj_mensal", {})

            if not er and not pm and vend == 0:
                continue

            proj = estoque_projetado(sku, disp, vend, pedidos, proj_venda=er, proj_mensal=pm)
            for i, (label, saldo, chegadas, m, a) in enumerate(proj):
                if saldo < 0:
                    alertas.append({"sku": sku, "cat": cat, "mes": label, "idx": i, "saldo": saldo})
                    break

        for w in self.alerta_detalhe.winfo_children(): w.destroy()
        for w in self.alerta_pills.winfo_children(): w.destroy()

        if not alertas:
            self.alerta_outer.pack_forget()
            return

        alertas.sort(key=lambda x: x["idx"])
        criticos = [a for a in alertas if a["idx"] == 0]
        urgentes = [a for a in alertas if a["idx"] == 1]
        atencao  = [a for a in alertas if a["idx"] >= 2]

        if criticos:
            bg, border, txt = "#FEE2E2", "#EF4444", "#991B1B"
        elif urgentes:
            bg, border, txt = "#FEF3C7", "#F59E0B", "#92400E"
        else:
            bg, border, txt = "#FEF9C3", "#EAB308", "#713F12"

        self.alerta_banner.configure(fg_color=bg, border_color=border)
        self.alerta_titulo_lbl.configure(text=f"{len(alertas)} SKU(s) em risco de ruptura", text_color=txt)
        self.alerta_toggle_btn.configure(fg_color=border, hover_color=border)

        for label, count, color in [("🔴 Crítico", len(criticos), "#EF4444"),
                                     ("🟠 Urgente", len(urgentes), "#F59E0B"),
                                     ("🟡 Atenção", len(atencao),  "#CA8A04")]:
            if count > 0:
                pill = ctk.CTkFrame(self.alerta_pills, fg_color=color, corner_radius=12)
                pill.pack(side="left", padx=(0, 6))
                ctk.CTkLabel(pill, text=f"{label}: {count}", font=ctk.CTkFont("Segoe UI", 10, "bold"),
                             text_color="white").pack(padx=10, pady=3)

        for a in alertas:
            if a["idx"] == 0:   urgencia, cor = "🔴 JÁ NEGATIVO",    "#EF4444"
            elif a["idx"] == 1: urgencia, cor = "🟠 PRÓXIMO MÊS",    "#F59E0B"
            else:               urgencia, cor = f"🟡 Ruptura em {a['mes']}", "#CA8A04"

            row = ctk.CTkFrame(self.alerta_detalhe, fg_color="white", corner_radius=8)
            row.pack(fill="x", pady=3, padx=4)
            ctk.CTkLabel(row, text=a["sku"],  font=ctk.CTkFont("Segoe UI", 11, "bold"), text_color="#111827", width=120, anchor="w").pack(side="left", padx=(12,4), pady=6)
            ctk.CTkLabel(row, text=a["cat"],  font=ctk.CTkFont("Segoe UI", 10), text_color="#6B7280").pack(side="left", padx=(0,12))
            ctk.CTkLabel(row, text=urgencia,  font=ctk.CTkFont("Segoe UI", 10, "bold"), text_color=cor).pack(side="left")
            ctk.CTkLabel(row, text=f"Saldo projetado: {fmt_num(a['saldo'])}",
                         font=ctk.CTkFont("Segoe UI", 10), text_color="#6B7280").pack(side="right", padx=12)

        self.alerta_outer.pack(fill="x", pady=(0, 8), before=self.tabs)

    # ══════════════════════════════════════════════════════════
    # ABA 1: ESTOQUE & NECESSIDADE E EXPORTS
    # ══════════════════════════════════════════════════════════
    def _build_tab_estoque(self, parent):
        tb = ctk.CTkFrame(parent, fg_color=BG_SIDEBAR, corner_radius=12, height=56)
        tb.pack(fill="x", padx=8, pady=(8,6)); tb.pack_propagate(False)

        ctk.CTkEntry(tb, textvariable=self.busca_v, placeholder_text="🔍 Buscar SKU…", width=180, height=40, corner_radius=8, fg_color="#FFFFFF", font=ctk.CTkFont("Segoe UI",12)).pack(side="left", padx=(10,6), pady=8)
        ctk.CTkLabel(tb, text="Categoria:", font=ctk.CTkFont("Segoe UI",11), text_color=CLR_MUTED).pack(side="left", padx=(6,4))
        ctk.CTkOptionMenu(tb, values=["TODAS","MONOFÁSICA","TRIFÁSICA","CARRETEL","UI","BORNE"], variable=self.cat_var, width=140, height=40, corner_radius=8, font=ctk.CTkFont("Segoe UI",11), fg_color="#FFFFFF", text_color=CLR_TEXT, button_color="#E5E7EB", button_hover_color="#D1D5DB", command=lambda _: self._popular_estoque()).pack(side="left", padx=6)
        ctk.CTkButton(tb, text="🔄 Atualizar", command=self._atualizar, fg_color=CLR_ACCENT, hover_color="#2563EB", corner_radius=8, height=40, width=110, font=ctk.CTkFont("Segoe UI",11,"bold")).pack(side="left", padx=6)
        ctk.CTkButton(tb, text="💾 Salvar", command=self._salvar, fg_color=CLR_GREEN, hover_color="#059669", corner_radius=8, height=40, width=90, font=ctk.CTkFont("Segoe UI",11,"bold")).pack(side="left", padx=4)

        ctk.CTkLabel(tb, text="Mês:", font=ctk.CTkFont("Segoe UI",11), text_color=CLR_MUTED).pack(side="left", padx=(6,2))
        hoje_tb = date.today()
        _meses_opt = ["Todos"] + [
            f"{MESES_PT_CURTO[(hoje_tb.month-1+i)%12+1]}/{str(hoje_tb.year+(hoje_tb.month-1+i)//12)[2:]}"
            for i in range(12)
        ]
        self.mes_view_var = tk.StringVar(value="Todos")
        ctk.CTkOptionMenu(tb, values=_meses_opt, variable=self.mes_view_var, width=110, height=40, corner_radius=8, font=ctk.CTkFont("Segoe UI",11), fg_color="#FFFFFF", text_color=CLR_TEXT, button_color="#E5E7EB", button_hover_color="#D1D5DB", command=lambda _: self._popular_estoque()).pack(side="left", padx=4)

        # BOTÃO GRADE VISUAL (TELA CHEIA)
        ctk.CTkButton(tb, text="🔲 Grade Visual", command=lambda: self._abrir_grade_tela_cheia(modo="estoque"), fg_color="#1E293B", hover_color="#334155", corner_radius=8, height=40, width=120, font=ctk.CTkFont("Segoe UI",11,"bold")).pack(side="left", padx=8)

        # NOVOS BOTÕES: EXPORTAR EXCEL E PDF
        ctk.CTkButton(tb, text="📊 Excel", command=self._exportar_excel, fg_color="#10B981", hover_color="#059669", corner_radius=8, height=40, width=80, font=ctk.CTkFont("Segoe UI",11,"bold")).pack(side="left", padx=4)
        ctk.CTkButton(tb, text="📄 PDF", command=self._exportar_pdf, fg_color="#EF4444", hover_color="#DC2626", corner_radius=8, height=40, width=80, font=ctk.CTkFont("Segoe UI",11,"bold")).pack(side="left", padx=2)
        ctk.CTkButton(tb, text="⚖️ Orçado x Real", command=self._exportar_orcado_realizado, fg_color="#F59E0B", hover_color="#D97706", corner_radius=8, height=40, width=120, font=ctk.CTkFont("Segoe UI",11,"bold")).pack(side="left", padx=4)

        self.btn_ia = ctk.CTkButton(tb, text="🤖 Recomendar (IA)", command=self._ia_recomendar, fg_color=CLR_PURPLE, hover_color="#7C3AED", corner_radius=8, height=40, width=180, font=ctk.CTkFont("Segoe UI",11,"bold"))
        self.btn_ia.pack(side="right", padx=10)

        leg = ctk.CTkFrame(parent, fg_color="transparent")
        leg.pack(fill="x", padx=10, pady=(2,6))
        for txt, cor in [("🔴 Ruptura prevista","#EF4444"),("🟠 Compra urgente","#F97316"),("🟡 Atenção","#EABB00"),("✏️ Duplo-clique para editar","#6B7280")]:
            ctk.CTkLabel(leg, text=txt, font=ctk.CTkFont("Segoe UI",10), text_color=cor).pack(side="left", padx=8)

        hoje = date.today()
        _meses_hdr = []
        for _i in range(12):
            _m = (hoje.month - 1 + _i) % 12 + 1
            _a = hoje.year + (hoje.month - 1 + _i) // 12
            _meses_hdr.append(f"{MESES_PT_CURTO[_m]}/{str(_a)[2:]}")

        COLS_E = ("cat","sku","disp","vendas_kg","proj_venda","necessidade","chegada",
                  "m_ini","m_ent","m_sai","m_fin",
                  "sf1","sf2","sf3","sf4","sf5","sf6","sf7","sf8","sf9","sf10","sf11","sf12")
        HDR_E  = ("Cat.","SKU","Disponível","Vendas/mês","Proj. Venda","Necessidade","Próx. Chegada",
                  "Saldo Inicial","Entradas","Saídas Proj.","Saldo Final",
                  *_meses_hdr)
        WID_E  = (105,175,95,100,110,105,120,
                  110,100,110,110,
                  68,68,68,68,68,68,68,68,68,68,68,68)
        COL_EDIT_E = {6:"chegada"}

        card = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=12, border_width=1, border_color=CLR_BORDER)
        card.pack(fill="both", expand=True, padx=8, pady=(0,8))

        s = ttk.Style(); s.theme_use("clam")
        s.configure("E.Treeview", background=BG_CARD, foreground=CLR_TEXT, rowheight=34, fieldbackground=BG_CARD, font=("Segoe UI",11), borderwidth=0)
        s.configure("E.Treeview.Heading", background=CLR_TABLE_HEAD, foreground="white", font=("Segoe UI",10,"bold"), relief="flat", padding=(8,8))
        s.map("E.Treeview", background=[("selected","#E0E7FF")], foreground=[("selected","#1E3A8A")])
        s.layout("E.Treeview",[("E.Treeview.treearea",{"sticky":"nswe"})])

        vsb = ttk.Scrollbar(card, orient="vertical")
        hsb = ttk.Scrollbar(card, orient="horizontal")
        self.tree_e = ttk.Treeview(card, columns=COLS_E, show="headings", style="E.Treeview", yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.config(command=self.tree_e.yview)
        hsb.config(command=self.tree_e.xview)

        for col, hdr, w in zip(COLS_E, HDR_E, WID_E):
            anch = "w" if col in ("sku","cat") else "center"
            self.tree_e.heading(col, text=f"  {hdr}", anchor="w", command=lambda c=col: self._sort_e(c))
            self.tree_e.column(col, width=w, minwidth=55, anchor=anch)

        for cat, c in CAT_COLORS.items():
            self.tree_e.tag_configure(cat, background=c["row"])
            self.tree_e.tag_configure(cat+"_alt", background=c["alt"])
        self.tree_e.tag_configure("rutura", foreground="#DC2626", font=("Segoe UI",11,"bold"))
        self.tree_e.tag_configure("urgente", foreground="#EA580C", font=("Segoe UI",11,"bold"))
        self.tree_e.tag_configure("atencao", foreground="#B45309")
        self.tree_e.tag_configure("sep", background="#F3F4F6", foreground="#4B5563", font=("Segoe UI",10,"bold"))

        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree_e.pack(fill="both", expand=True, padx=2, pady=2)
        self.tree_e.bind("<Double-1>", lambda ev: self._edit_cell(ev, self.tree_e, COL_EDIT_E))
        self._sort_col_e = None; self._sort_rev_e = False


    def _exportar_excel(self):
        if not HAS_OPENPYXL:
            messagebox.showerror("Dependência ausente", "Instale openpyxl:\n  pip install openpyxl")
            return
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], title="Exportar Projeção de Estoque — Excel")
        if not filepath: return

        rows, meses = self._preparar_dados_grade(modo="estoque")

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Projeção de Estoque"

        CLR_HDR_DARK = "0B192C"
        CLR_HDR_MES  = "1E3A8A"
        CLR_HDR_SUB  = "2D4BA0"
        CLR_CAT_BG   = "1E3A8A"
        CLR_POS      = "D1F2D3"
        CLR_NEG      = "FCE4E4"
        CLR_NEG_FNT  = "C62828"
        CLR_POS_FNT  = "065F46"
        CLR_COMPRA   = "065F46"
        CLR_BORDER   = "D1D5DB"

        def fill(hex6): return PatternFill("solid", fgColor=hex6)
        def font(bold=False, color="000000", sz=10): return Font(bold=bold, color=color, size=sz, name="Segoe UI")
        def border():
            s = Side(style="thin", color=CLR_BORDER)
            return Border(left=s, right=s, top=s, bottom=s)
        def center(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
        def left():   return Alignment(horizontal="left",   vertical="center")

        n_meses  = len(meses)
        COL_SKU  = 1
        COL_VEND = 2
        COL_M0   = 3

        ws.column_dimensions[get_column_letter(COL_SKU)].width  = 22
        ws.column_dimensions[get_column_letter(COL_VEND)].width = 10
        for mi in range(n_meses):
            for ci in range(3):
                ws.column_dimensions[get_column_letter(COL_M0 + mi*3 + ci)].width = 11

        last_col = get_column_letter(COL_M0 + n_meses*3 - 1)

        # linha 1 — título
        ws.row_dimensions[1].height = 24
        ws.merge_cells(f"A1:{last_col}1")
        c = ws["A1"]
        c.value = "Projeção de Estoque"
        c.font  = font(bold=True, color="FFFFFF", sz=13)
        c.fill  = fill(CLR_HDR_DARK)
        c.alignment = Alignment(horizontal="right", vertical="center")

        # linha 2 — data
        ws.row_dimensions[2].height = 16
        ws.merge_cells(f"A2:{last_col}2")
        c = ws["A2"]
        c.value = f"Gerado em: {date.today().strftime('%d/%m/%Y')}"
        c.font  = font(color="6B7280", sz=9)
        c.fill  = fill("F9FAFB")
        c.alignment = Alignment(horizontal="right", vertical="center")

        # linha 3 — cabeçalho meses
        ws.row_dimensions[3].height = 20
        ws["A3"].fill = fill(CLR_HDR_DARK)
        ws["B3"].fill = fill(CLR_HDR_DARK)
        for mi, (m, a, _) in enumerate(meses):
            c0 = COL_M0 + mi*3
            cl = get_column_letter(c0); cr = get_column_letter(c0+2)
            ws.merge_cells(f"{cl}3:{cr}3")
            c = ws[f"{cl}3"]
            c.value = f"{MESES_PT_CURTO[m]}/{a}"
            c.font  = font(bold=True, color="FFFFFF", sz=11)
            c.fill  = fill(CLR_HDR_MES)
            c.alignment = center()
            for ci in range(3):
                ws[f"{get_column_letter(c0+ci)}3"].fill = fill(CLR_HDR_MES)

        # linha 4 — sub-cabeçalhos
        ws.row_dimensions[4].height = 32
        ws["A4"].value = "SKU";        ws["A4"].font = font(bold=True, color="FFFFFF"); ws["A4"].fill = fill(CLR_HDR_DARK); ws["A4"].alignment = center()
        ws["B4"].value = "Venda\nMês"; ws["B4"].font = font(bold=True, color="FFFFFF"); ws["B4"].fill = fill(CLR_HDR_DARK); ws["B4"].alignment = center()
        for mi in range(n_meses):
            c0 = COL_M0 + mi*3
            for ci, sub in enumerate(["Proj.\nVenda", "Entradas", "Saldo"]):
                cx = ws[f"{get_column_letter(c0+ci)}4"]
                cx.value = sub; cx.font = font(bold=True, color="FFFFFF", sz=9)
                cx.fill = fill(CLR_HDR_SUB); cx.alignment = center()

        for col in range(1, COL_M0 + n_meses*3):
            for row in [3, 4]:
                ws.cell(row=row, column=col).border = border()

        # dados
        ROW = 5
        for r in rows:
            if r["type"] == "cat":
                ws.row_dimensions[ROW].height = 18
                ws.merge_cells(f"A{ROW}:{last_col}{ROW}")
                c = ws[f"A{ROW}"]
                c.value = r["label"]
                c.font  = font(bold=True, color="FFFFFF", sz=11)
                c.fill  = fill(CLR_CAT_BG)
                c.alignment = left()
                for col in range(1, COL_M0 + n_meses*3):
                    ws.cell(row=ROW, column=col).fill = fill(CLR_CAT_BG)
                    ws.cell(row=ROW, column=col).border = border()
                ROW += 1
                continue

            sku  = r["sku"]
            vend = vendas_mes_atual(sku, self.vendas)
            er   = self.local["config"].get(sku, {}).get("proj_venda", "")

            ws.row_dimensions[ROW].height = 18
            c_sku = ws.cell(row=ROW, column=COL_SKU, value=sku)
            c_sku.font = font(bold=True); c_sku.alignment = left(); c_sku.border = border()

            try:   vend_val = float(vend)
            except: vend_val = 0
            c_v = ws.cell(row=ROW, column=COL_VEND, value=vend_val if vend_val else None)
            c_v.alignment = center(); c_v.border = border(); c_v.number_format = "#,##0"

            for mi, md in enumerate(r["months"]):
                c0    = COL_M0 + mi*3
                saldo = md["atual"]
                cheg  = md["chegada"]

                try:   pv_val = float(er) if er else None
                except: pv_val = None
                cp = ws.cell(row=ROW, column=c0, value=pv_val)
                cp.alignment = center(); cp.border = border(); cp.number_format = "#,##0"

                cc = ws.cell(row=ROW, column=c0+1, value=cheg if cheg > 0 else None)
                cc.alignment = center(); cc.border = border(); cc.number_format = "#,##0"
                if cheg > 0: cc.font = font(bold=True, color=CLR_COMPRA)

                bg = CLR_NEG if saldo < 0 else (CLR_POS if saldo > 0 else "FFFFFF")
                cs = ws.cell(row=ROW, column=c0+2, value=round(saldo, 3))
                cs.fill = fill(bg); cs.alignment = center(); cs.border = border()
                cs.number_format = "#,##0.###"
                if saldo < 0:  cs.font = font(bold=True, color=CLR_NEG_FNT)
                elif saldo > 0: cs.font = font(color=CLR_POS_FNT)

            ROW += 1

        ws.freeze_panes = "C5"
        wb.save(filepath)
        messagebox.showinfo("Exportado", f"Excel gerado com sucesso:\n{filepath}")
        os.startfile(filepath)

    def _exportar_orcado_realizado(self):
        if not HAS_OPENPYXL:
            messagebox.showerror("Dependência ausente", "Instale openpyxl:\n  pip install openpyxl")
            return
        
        hoje = date.today()
        mes_str = f"{MESES_PT[hoje.month]}/{hoje.year}"
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=[("Excel", "*.xlsx")], 
            title=f"Exportar Orçado vs Realizado - {mes_str}",
            initialfile=f"orcado_realizado_{hoje.strftime('%Y_%m')}.xlsx"
        )
        if not filepath: return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"Orçado x Real {MESES_PT_CURTO[hoje.month]}"

        def fill(hex6): return PatternFill("solid", fgColor=hex6)
        def font(bold=False, color="000000", sz=10): return Font(bold=bold, color=color, size=sz, name="Segoe UI")
        def border():
            s = Side(style="thin", color="D1D5DB")
            return Border(left=s, right=s, top=s, bottom=s)
        def center(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
        def left():   return Alignment(horizontal="left",   vertical="center")

        ws.merge_cells("A1:F1")
        ws["A1"].value = f"Resumo Orçado vs Realizado: {mes_str}"
        ws["A1"].font = font(bold=True, color="FFFFFF", sz=13)
        ws["A1"].fill = fill("0B192C")
        ws["A1"].alignment = center()
        ws.row_dimensions[1].height = 26

        headers = ["Categoria", "SKU", "Orçado (kg)", "Realizado (kg)", "Diferença (kg)", "Atingimento"]
        col_widths = [16, 24, 16, 16, 16, 16]
        for i, (h, w) in enumerate(zip(headers, col_widths), 1):
            c = ws.cell(row=2, column=i, value=h)
            c.font = font(bold=True, color="FFFFFF")
            c.fill = fill("1E3A8A")
            c.alignment = center()
            c.border = border()
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.row_dimensions[2].height = 22

        _sel_key = f"{hoje.year}-{hoje.month:02d}"
        
        ROW = 3
        tot_orcado = 0.0
        tot_realizado = 0.0

        for it in sorted(ITENS, key=lambda x: (CAT_ORDEM.get(x["cat"],9), _sku_num_key(x["sku"]), x["sku"])):
            sku = it["sku"]
            cat = it["cat"]
            
            er = self.local["config"].get(sku, {}).get("proj_venda", "")
            pm = self.local["config"].get(sku, {}).get("proj_mensal", {})
            try:
                if pm and _sel_key in pm and str(pm[_sel_key]).strip():
                    orcado = float(pm[_sel_key])
                elif er:
                    orcado = float(er)
                else:
                    orcado = 0.0
            except:
                orcado = 0.0
                
            vend = vendas_mes_atual(sku, self.vendas)
            try:
                realizado = float(vend)
            except:
                realizado = 0.0
                
            if orcado == 0 and realizado == 0:
                continue
                
            tot_orcado += orcado
            tot_realizado += realizado
                
            diferenca = realizado - orcado
            atingimento = (realizado / orcado) if orcado > 0 else (1.0 if realizado > 0 else 0.0)

            ws.cell(row=ROW, column=1, value=cat).alignment = left()
            ws.cell(row=ROW, column=1).border = border()
            
            c_sku = ws.cell(row=ROW, column=2, value=sku)
            c_sku.font = font(bold=True)
            c_sku.border = border()
            c_sku.alignment = left()
            
            c_orcado = ws.cell(row=ROW, column=3, value=orcado)
            c_orcado.number_format = "#,##0"
            c_orcado.alignment = center()
            c_orcado.border = border()
            
            c_real = ws.cell(row=ROW, column=4, value=realizado)
            c_real.number_format = "#,##0"
            c_real.alignment = center()
            c_real.border = border()
            
            c_dif = ws.cell(row=ROW, column=5, value=diferenca)
            c_dif.number_format = "#,##0"
            c_dif.alignment = center()
            c_dif.border = border()
            c_dif.font = font(bold=True, color="065F46" if diferenca >= 0 else "C62828")
            
            c_ating = ws.cell(row=ROW, column=6, value=atingimento)
            c_ating.number_format = "0.0%"
            c_ating.alignment = center()
            c_ating.border = border()
            if atingimento >= 1:
                c_ating.font = font(bold=True, color="065F46")
            elif atingimento < 0.8:
                c_ating.font = font(bold=True, color="C62828")
            
            ROW += 1

        if ROW > 3:
            ws.merge_cells(f"A{ROW}:B{ROW}")
            c_tot_lbl = ws[f"A{ROW}"]
            c_tot_lbl.value = "TOTAL GERAL"
            c_tot_lbl.font = font(bold=True, color="FFFFFF", sz=11)
            c_tot_lbl.fill = fill("0F172A")
            c_tot_lbl.alignment = center()
            
            c_tot_orc = ws.cell(row=ROW, column=3, value=tot_orcado)
            c_tot_orc.font = font(bold=True, color="FFFFFF", sz=11)
            c_tot_orc.fill = fill("0F172A")
            c_tot_orc.number_format = "#,##0"
            c_tot_orc.alignment = center()
            
            c_tot_real = ws.cell(row=ROW, column=4, value=tot_realizado)
            c_tot_real.font = font(bold=True, color="FFFFFF", sz=11)
            c_tot_real.fill = fill("0F172A")
            c_tot_real.number_format = "#,##0"
            c_tot_real.alignment = center()
            
            tot_dif = tot_realizado - tot_orcado
            c_tot_dif = ws.cell(row=ROW, column=5, value=tot_dif)
            c_tot_dif.font = font(bold=True, color="FFFFFF", sz=11)
            c_tot_dif.fill = fill("0F172A")
            c_tot_dif.number_format = "#,##0"
            c_tot_dif.alignment = center()
            
            tot_ating = (tot_realizado / tot_orcado) if tot_orcado > 0 else (1.0 if tot_realizado > 0 else 0.0)
            c_tot_ating = ws.cell(row=ROW, column=6, value=tot_ating)
            c_tot_ating.font = font(bold=True, color="FFFFFF", sz=11)
            c_tot_ating.fill = fill("0F172A")
            c_tot_ating.number_format = "0.0%"
            c_tot_ating.alignment = center()

        ws.freeze_panes = "A3"
        wb.save(filepath)
        messagebox.showinfo("Exportado", f"Relatório Orçado vs Realizado salvo com sucesso:\n{filepath}")
        try: os.startfile(filepath)
        except: pass

    def _exportar_pedidos_excel(self):
        if not HAS_OPENPYXL:
            messagebox.showerror("Dependência ausente", "Instale openpyxl:\n  pip install openpyxl")
            return
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            title="Exportar Compras & Chegadas — Excel"
        )
        if not filepath: return

        pedidos = sorted(self.local["pedidos"], key=lambda p: (parse_date(p.get("data","") or "") or date(2099,1,1)))
        hoje = date.today()

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Compras & Chegadas"

        # estilos
        def font(bold=False, color="000000", sz=11):
            return Font(name="Segoe UI", bold=bold, color=color, size=sz)
        def fill(hex_color):
            return PatternFill("solid", fgColor=hex_color)
        def border():
            s = Side(style="thin", color="E5E7EB")
            return Border(left=s, right=s, top=s, bottom=s)
        def center():
            return Alignment(horizontal="center", vertical="center")
        def left():
            return Alignment(horizontal="left", vertical="center")

        HDR_BG  = "0B192C"
        PEND_BG = "FFFBEB"
        CHEG_BG = "F0FDF4"
        ATRA_BG = "FEE2E2"

        # cabeçalho título
        ws.merge_cells("A1:H1")
        c = ws["A1"]
        c.value = f"Compras & Chegadas — Exportado em {hoje.strftime('%d/%m/%Y')}"
        c.font = font(bold=True, color="FFFFFF", sz=13)
        c.fill = fill(HDR_BG)
        c.alignment = center()
        ws.row_dimensions[1].height = 26

        # cabeçalho colunas
        headers = ["Mês", "SKU", "Categoria", "Quantidade", "Data Chegada", "Observação", "Status", "Chegado?"]
        for ci, h in enumerate(headers, 1):
            c = ws.cell(row=2, column=ci, value=h)
            c.font = font(bold=True, color="FFFFFF", sz=10)
            c.fill = fill("1E3A8A")
            c.alignment = center()
            c.border = border()
        ws.row_dimensions[2].height = 22

        widths = [14, 18, 14, 13, 16, 28, 12, 10]
        for ci, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(ci)].width = w

        # dados
        ROW = 3
        for p in pedidos:
            sku    = p.get("sku", "")
            qty    = p.get("qty", "")
            data_s = p.get("data", "")
            obs    = p.get("obs", "")
            chegou = p.get("chegado", False)
            cat    = SKU_CAT.get(sku.upper(), "—")
            dt     = parse_date(data_s or "")

            if chegou:
                status = "Chegado"; row_bg = CHEG_BG; font_color = "16A34A"
            elif dt and dt < hoje:
                status = "Atrasado"; row_bg = ATRA_BG; font_color = "E11D48"
            else:
                status = "Pendente"; row_bg = PEND_BG; font_color = "92400E"

            mes_label = f"{MESES_PT[dt.month]}/{dt.year}" if dt else "—"

            valores = [mes_label, sku, cat, qty, data_s, obs, status, "Sim" if chegou else "Não"]
            for ci, v in enumerate(valores, 1):
                c = ws.cell(row=ROW, column=ci, value=v)
                c.fill = fill(row_bg)
                c.font = Font(name="Segoe UI", size=10, color=font_color if ci == 7 else "1F2937",
                              bold=(ci == 7))
                c.alignment = left() if ci in (2, 6) else center()
                c.border = border()
            ws.row_dimensions[ROW].height = 20
            ROW += 1

        wb.save(filepath)
        messagebox.showinfo("Exportado", f"Compras salvas em:\n{filepath}")
        try: os.startfile(filepath)
        except: pass

    def _importar_pedidos_excel(self):
        if not HAS_OPENPYXL:
            messagebox.showerror("Dependência ausente", "Instale openpyxl:\n  pip install openpyxl")
            return
        filepath = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx")],
            title="Importar Compras & Chegadas — Excel"
        )
        if not filepath: return

        try:
            wb = openpyxl.load_workbook(filepath)
            ws = wb.active
        except Exception as e:
            messagebox.showerror("Erro ao abrir", f"Não foi possível ler o arquivo:\n{e}")
            return

        # monta chave de identificação dos pedidos já existentes para evitar duplicatas
        def _chave(p):
            return (p.get("sku","").upper(), str(p.get("qty","")).strip(), str(p.get("data","")).strip())

        existentes = {_chave(p) for p in self.local["pedidos"]}

        adicionados = 0
        atualizados = 0
        erros = []

        # linha 1 = título, linha 2 = cabeçalho, dados a partir da linha 3
        for row in ws.iter_rows(min_row=3, values_only=True):
            # colunas: Mês, SKU, Categoria, Quantidade, Data Chegada, Observação, Status, Chegado?
            _, sku, _, qty, data_s, obs, _, chegado_s = (row + (None,)*8)[:8]

            if not sku or not qty or not data_s:
                continue

            sku    = str(sku).strip().upper()
            qty_s  = str(qty).strip()
            data_s = str(data_s).strip()
            obs    = str(obs).strip() if obs else ""
            chegou = str(chegado_s).strip().lower() in ("sim", "yes", "true", "1")

            # valida data
            dt = parse_date(data_s)
            if dt is None:
                erros.append(f"{sku} — data inválida: '{data_s}'")
                continue

            # valida quantidade
            try:
                float(qty_s.replace(",", "."))
            except ValueError:
                erros.append(f"{sku} — quantidade inválida: '{qty_s}'")
                continue

            chave = (sku, qty_s, data_s)
            if chave in existentes:
                # atualiza só o status chegado caso tenha mudado
                for p in self.local["pedidos"]:
                    if _chave(p) == chave and bool(p.get("chegado")) != chegou:
                        p["chegado"] = chegou
                        atualizados += 1
            else:
                self.local["pedidos"].append({
                    "sku": sku,
                    "qty": qty_s,
                    "data": data_s,
                    "obs": obs,
                    "chegado": chegou
                })
                existentes.add(chave)
                adicionados += 1

        salvar_local(self.local)
        self._popular_pedidos()
        self._popular_estoque()

        msg = f"Importação concluída!\n\n✅ Adicionados: {adicionados}\n🔄 Atualizados: {atualizados}"
        if erros:
            msg += f"\n\n⚠️ Erros ignorados ({len(erros)}):\n" + "\n".join(erros[:5])
            if len(erros) > 5:
                msg += f"\n... e mais {len(erros)-5}"
        messagebox.showinfo("Importar Excel", msg)

    def _exportar_pdf(self):
        filepath = filedialog.asksaveasfilename(defaultextension=".html", filetypes=[("Relatório PDF/HTML", "*.html")], title="Salvar Relatório Mês a Mês")
        if not filepath: return
        
        rows, meses = self._preparar_dados_grade(modo="estoque")
        
        html = ["<html><head><meta charset='utf-8'><style>"]
        html.append("body { font-family: 'Segoe UI', sans-serif; background-color: #F8FAFB; padding: 20px; }")
        html.append("table { width: 100%; border-collapse: collapse; margin-top: 20px; background: white; }")
        html.append("th, td { border: 1px solid #E5E7EB; padding: 8px; text-align: center; font-size: 12px; }")
        html.append("th { background-color: #0B192C; color: white; }")
        html.append(".cat { background-color: #1E3A8A; color: white; font-weight: bold; text-align: left; padding-left: 15px; }")
        html.append(".sku { font-weight: bold; text-align: left; }")
        html.append(".pos { background-color: #D1F2D3; }")
        html.append(".neg { background-color: #FCE4E4; }")
        html.append(".val-neg { color: #C62828; font-weight: bold; }")
        html.append(".chegada { color: #16A34A; font-weight: bold; display: block; margin-bottom: 4px; }")
        html.append("</style></head><body>")
        html.append(f"<h2>Relatório de Estoque e Necessidade Mês a Mês ({date.today().strftime('%d/%m/%Y')})</h2>")
        
        html.append("<table><thead><tr><th>SKU</th>")
        for m, a, lbl in meses:
            html.append(f"<th>{lbl.replace(chr(10), ' ')}</th>")
        html.append("</tr></thead><tbody>")
        
        for r in rows:
            if r["type"] == "cat":
                html.append(f"<tr><td colspan='{len(meses)+1}' class='cat'>{r['label']}</td></tr>")
            else:
                html.append(f"<tr><td class='sku'>{r['sku']}</td>")
                for md in r["months"]:
                    atual = md["atual"]
                    cheg = md["chegada"]
                    bg_class = "pos" if atual >= 0 else "neg"
                    txt_class = "val-neg" if atual < 0 else ""
                    
                    cell_html = f"<td class='{bg_class}'>"
                    if cheg > 0:
                        cell_html += f"<span class='chegada'>📦 {fmt_num(cheg)}</span>"
                    cell_html += f"<span class='{txt_class}'>Saldo: {fmt_num(atual)}</span></td>"
                    html.append(cell_html)
                html.append("</tr>")
        
        html.append("</tbody></table>")
        html.append("<script>window.print();</script>") # Chama a janela de PDF automática
        html.append("</body></html>")
        
        with open(filepath, "w", encoding="utf-8") as f:
            f.write("\n".join(html))
            
        webbrowser.open("file://" + os.path.realpath(filepath))

    def _sort_e(self, col):
        data = [(self.tree_e.set(k,col), k) for k in self.tree_e.get_children("") if not k.startswith("_sep_")]
        rev  = (self._sort_col_e == col) and not self._sort_rev_e
        def sort_key(t):
            v = t[0].replace(".","").replace(",",".")
            try: return float(v) if v not in ("—","") else -1
            except: return -1
        try:    data.sort(key=sort_key, reverse=rev)
        except: data.sort(key=lambda t: t[0], reverse=rev)
        for i,(_, k) in enumerate(data): self.tree_e.move(k,"",i)
        self._sort_col_e = col; self._sort_rev_e = rev


    # ══════════════════════════════════════════════════════════
    # DADOS DA GRADE E TELA CHEIA
    # ══════════════════════════════════════════════════════════

    def _preparar_dados_grade(self, modo="estoque"):
        hoje = date.today()
        meses = []
        for i in range(12):
            m = (hoje.month - 1 + i) % 12 + 1
            a = hoje.year + (hoje.month - 1 + i) // 12
            meses.append((m, a, f"{MESES_PT_CURTO[m]}\n{a}"))

        rows = []
        cur_cat = None
        
        pedidos_base = self.local["pedidos"]
        if modo == "cenario":
            pedidos_com_sim = pedidos_base + self.cenario
            skus_sim = {p["sku"].upper() for p in self.cenario}
            # proj_venda overrides definidos no cenário (primeiro encontrado por SKU)
            cen_proj_overrides = {}
            for p in self.cenario:
                sk = p.get("sku","").upper()
                pv = p.get("proj_venda","")
                if pv and sk not in cen_proj_overrides:
                    cen_proj_overrides[sk] = pv
        else:
            pedidos_com_sim = pedidos_base
            skus_sim = set()
            cen_proj_overrides = {}

        for it in sorted(ITENS, key=lambda x: (CAT_ORDEM.get(x["cat"],9), _sku_num_key(x["sku"]), x["sku"])):
            sku, cat = it["sku"], it["cat"]
            e    = self.estoque.get(sku, {})
            disp = e.get("disp", 0)
            vend = vendas_mes_atual(sku, self.vendas)
            er_cfg = self.local["config"].get(sku, {}).get("proj_venda", "")
            pm_cfg = self.local["config"].get(sku, {}).get("proj_mensal", {})

            proj_sem = estoque_projetado(sku, disp, vend, pedidos_base, proj_venda=er_cfg, proj_mensal=pm_cfg)
            if modo == "cenario":
                er_sim = cen_proj_overrides.get(sku.upper(), er_cfg)
                _, pm_sim = self._get_proj_mensal_sim(sku, er_sim, pm_cfg)
                proj_com = estoque_projetado(sku, disp, vend, pedidos_com_sim, proj_venda=er_sim, proj_mensal=pm_sim)
            else:
                er_sim = er_cfg
                pm_sim = pm_cfg
                proj_com = proj_sem

            if cat != cur_cat:
                cur_cat = cat
                rows.append({
                    "type":  "cat",
                    "label": cat,
                    "color": CAT_COLORS.get(cat, {}).get("badge", "#6B7280"),
                })

            month_data = []
            er_sim_grade = cen_proj_overrides.get(sku.upper(), er_cfg) if modo == "cenario" else er_cfg
            for ci, (m, a, _) in enumerate(meses):
                atual_s = proj_sem[ci][1] if ci < len(proj_sem) else 0
                cen_s   = proj_com[ci][1] if ci < len(proj_com) else 0

                # chegada real (pedidos confirmados)
                chegada_real = sum(float(p.get("qty", 0)) for p in pedidos_base
                                   if p.get("sku","").upper() == sku.upper()
                                   and not p.get("chegado")
                                   and parse_date(p.get("data","") or "")
                                   and parse_date(p["data"]).month == m
                                   and parse_date(p["data"]).year  == a)
                # chegada simulada (cenário)
                if modo == "cenario":
                    chegada_sim = sum(float(p.get("qty", 0)) for p in self.cenario
                                      if p.get("sku","").upper() == sku.upper()
                                      and parse_date(p.get("data","") or "")
                                      and parse_date(p["data"]).month == m
                                      and parse_date(p["data"]).year  == a)
                    chegada = chegada_sim  # compatibilidade
                else:
                    chegada_sim = 0
                    chegada = chegada_real

                # proj_venda efetiva para este mês
                m_key = f"{a}-{m:02d}"
                pm_use = pm_sim if modo == "cenario" else pm_cfg
                if pm_use and m_key in pm_use and str(pm_use[m_key]).strip():
                    try:    pv_mes = float(pm_use[m_key])
                    except: pv_mes = float(er_sim_grade) if er_sim_grade else 0
                elif er_sim_grade:
                    try:    pv_mes = float(er_sim_grade)
                    except: pv_mes = 0
                else:
                    pv_mes = 0

                month_data.append({"atual": atual_s, "cen": cen_s, "chegada": chegada,
                                   "chegada_real": chegada_real, "chegada_sim": chegada_sim,
                                   "proj_venda": pv_mes})

            rows.append({
                "type":    "data",
                "sku":     sku,
                "has_sim": sku in skus_sim,
                "months":  month_data,
            })
        return rows, meses

    def _abrir_grade_tela_cheia(self, modo="estoque"):
        dlg = ctk.CTkToplevel(self)
        dlg.title(f"Grade Visual — {'Cenário de Simulação' if modo == 'cenario' else 'Estoque & Necessidade'}")
        
        try: dlg.state('zoomed')
        except: dlg.attributes('-zoomed', True)
            
        dlg.configure(fg_color=BG_APP)
        
        tb = ctk.CTkFrame(dlg, fg_color=BG_HEADER, height=50, corner_radius=0)
        tb.pack(fill="x")
        tb.pack_propagate(False)
        ctk.CTkLabel(tb, text=f"📊 Visão Expandida — {'Cenário' if modo=='cenario' else 'Estoque'}", font=ctk.CTkFont("Segoe UI", 16, "bold"), text_color="white").pack(side="left", padx=20)
        ctk.CTkButton(tb, text="✕ Fechar", command=dlg.destroy, fg_color=CLR_RED, hover_color="#B91C1C", width=100).pack(side="right", padx=20)

        leg_row = ctk.CTkFrame(dlg, fg_color="transparent")
        leg_row.pack(fill="x", padx=10, pady=10)
        
        if modo == "cenario":
            legendas = [("🟩 Estoque Positivo", "#16A34A"), ("🟥 Ruptura", "#DC2626"), ("▲ Chegada Simulada", "#000000")]
        else:
            legendas = [("🟩 Estoque Positivo", "#16A34A"), ("🟥 Ruptura", "#DC2626"), ("📦 Chegada Prevista (Pedidos Reais)", "#16A34A")]
            
        for txt, cor in legendas:
            ctk.CTkLabel(leg_row, text=txt, font=ctk.CTkFont("Segoe UI", 11, "bold"), text_color=cor).pack(side="left", padx=10)

        grade_card = ctk.CTkFrame(dlg, fg_color=BG_CARD, corner_radius=12, border_width=1, border_color=CLR_BORDER)
        grade_card.pack(fill="both", expand=True, padx=10, pady=(0,10))

        vbar = ttk.Scrollbar(grade_card, orient="vertical")
        hbar = ttk.Scrollbar(grade_card, orient="horizontal")
        grade_fs = GradeSimulacao(grade_card)
        vbar.config(command=grade_fs.yview)
        hbar.config(command=grade_fs.xview)
        grade_fs.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)
        
        vbar.pack(side="right", fill="y")
        hbar.pack(side="bottom", fill="x")
        grade_fs.pack(fill="both", expand=True, padx=2, pady=2)

        grade_fs.bind("<MouseWheel>", lambda e: grade_fs.yview_scroll(-1*(e.delta//120), "units"))
        grade_fs.bind("<Shift-MouseWheel>", lambda e: grade_fs.xview_scroll(-1*(e.delta//120), "units"))

        rows, meses = self._preparar_dados_grade(modo=modo)
        grade_fs.render(rows, meses, modo=modo)

    # ══════════════════════════════════════════════════════════
    # ABA 2: COMPRAS & CHEGADAS
    # ══════════════════════════════════════════════════════════

    def _build_tab_compras(self, parent):
        form = ctk.CTkFrame(parent, fg_color=BG_SIDEBAR, corner_radius=12, border_width=1, border_color=CLR_BORDER)
        form.pack(fill="x", padx=8, pady=(10,8))
        ctk.CTkLabel(form, text="➕ Novo Pedido de Compra", font=ctk.CTkFont("Segoe UI",13,"bold"), text_color=CLR_ACCENT).grid(row=0,column=0,columnspan=9, sticky="w", padx=16, pady=(12,8))

        labels  = ["SKU *","Qtd *","Data Chegada *","Obs"]
        for i, lbl in enumerate(labels):
            ctk.CTkLabel(form, text=lbl, font=ctk.CTkFont("Segoe UI",10), text_color=CLR_MUTED).grid(row=1,column=i*2,padx=(16 if i==0 else 8,4), pady=10,sticky="w")

        self.form_sku = ctk.CTkComboBox(form, values=SKU_LIST, width=200, height=40, corner_radius=8, fg_color="#FFFFFF", font=ctk.CTkFont("Segoe UI",11))
        self.form_sku.set("")
        self.form_sku.grid(row=1,column=1, padx=4, pady=10)

        self.form_qty = ctk.CTkEntry(form, placeholder_text="Ex: 100", width=120, height=40, corner_radius=8, fg_color="#FFFFFF", font=ctk.CTkFont("Segoe UI",11))
        self.form_qty.grid(row=1,column=3, padx=4, pady=10)

        self.form_data = ctk.CTkEntry(form, placeholder_text="DD/MM/AAAA", width=130, height=40, corner_radius=8, fg_color="#FFFFFF", font=ctk.CTkFont("Segoe UI",11))
        self.form_data.grid(row=1,column=5, padx=4, pady=10)

        self.form_obs = ctk.CTkEntry(form, placeholder_text="Opcional", width=170, height=40, corner_radius=8, fg_color="#FFFFFF", font=ctk.CTkFont("Segoe UI",11))
        self.form_obs.grid(row=1,column=7, padx=4, pady=10)

        ctk.CTkButton(form, text="  ✚  Adicionar  ", command=self._add_pedido, fg_color=CLR_ACCENT, hover_color="#2563EB", corner_radius=8, height=40, font=ctk.CTkFont("Segoe UI",11,"bold")).grid(row=1,column=8, padx=16, pady=10)

        filt = ctk.CTkFrame(parent, fg_color="transparent")
        filt.pack(fill="x", padx=8, pady=(0,6))
        ctk.CTkLabel(filt, text="Filtrar por mês:", font=ctk.CTkFont("Segoe UI",11), text_color=CLR_MUTED).pack(side="left", padx=4)
        self.mes_var = tk.StringVar(value="TODOS")
        meses_opts = ["TODOS"] + [f"{MESES_PT[m]}/{date.today().year + (date.today().month-1+i)//12}" for i,m in enumerate([(date.today().month-1+j)%12+1 for j in range(12)])]
        ctk.CTkOptionMenu(filt, values=meses_opts, variable=self.mes_var, width=200, corner_radius=8, fg_color="#FFFFFF", command=lambda _: self._popular_pedidos(), button_color="#E5E7EB").pack(side="left", padx=6)

        COLS_P = ("mes","sku","cat","qty","data","obs","status")
        HDR_P  = ("Mês","SKU","Categoria","Quantidade","Data Chegada","Observação","Status")
        WID_P  = (130,180,120,110,130,200,120)

        card2 = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=12, border_width=1, border_color=CLR_BORDER)
        card2.pack(fill="both", expand=True, padx=8, pady=(0,8))

        s2 = ttk.Style()
        s2.configure("P.Treeview", background=BG_CARD, foreground=CLR_TEXT, rowheight=34, fieldbackground=BG_CARD, font=("Segoe UI",11), borderwidth=0)
        s2.configure("P.Treeview.Heading", background=CLR_TABLE_HEAD, foreground="white", font=("Segoe UI",10,"bold"), relief="flat", padding=(8,8))
        s2.map("P.Treeview", background=[("selected","#E0E7FF")], foreground=[("selected","#1E3A8A")])
        s2.layout("P.Treeview",[("P.Treeview.treearea",{"sticky":"nswe"})])

        vsb2 = ttk.Scrollbar(card2, orient="vertical")
        self.tree_p = ttk.Treeview(card2, columns=COLS_P, show="headings", style="P.Treeview", yscrollcommand=vsb2.set)
        vsb2.config(command=self.tree_p.yview)

        for col, hdr, w in zip(COLS_P, HDR_P, WID_P):
            self.tree_p.heading(col, text=f"  {hdr}", anchor="w")
            self.tree_p.column(col, width=w, minwidth=50, anchor="center" if col not in ("sku","obs") else "w")

        self.tree_p.tag_configure("chegou", background="#F0FDF4", foreground="#16A34A")
        self.tree_p.tag_configure("pendente", background="#FFFBEB")
        self.tree_p.tag_configure("atrasado", background="#FEE2E2", foreground="#E11D48", font=("Segoe UI",11,"bold"))
        self.tree_p.tag_configure("mes_hdr", background=CLR_TABLE_HEAD, foreground="white", font=("Segoe UI",11,"bold"))

        vsb2.pack(side="right", fill="y")
        self.tree_p.pack(fill="both", expand=True, padx=2, pady=2)

        rod = ctk.CTkFrame(parent, fg_color="transparent", height=44)
        rod.pack(fill="x", padx=8); rod.pack_propagate(False)
        ctk.CTkButton(rod, text="✏️ Editar Selecionado", command=self._edit_pedido, fg_color="#EFF6FF", text_color=CLR_ACCENT, hover_color="#DBEAFE", corner_radius=8, height=36, font=ctk.CTkFont("Segoe UI",11)).pack(side="left", pady=4, padx=(0,6))
        ctk.CTkButton(rod, text="🗑 Excluir Selecionado", command=self._del_pedido, fg_color="#FEE2E2", text_color=CLR_RED, hover_color="#FECACA", corner_radius=8, height=36, font=ctk.CTkFont("Segoe UI",11)).pack(side="left", pady=4)
        ctk.CTkButton(rod, text="📊 Exportar Excel", command=self._exportar_pedidos_excel, fg_color="#10B981", hover_color="#059669", corner_radius=8, height=36, font=ctk.CTkFont("Segoe UI",11,"bold")).pack(side="left", pady=4, padx=(12,0))
        ctk.CTkButton(rod, text="📥 Importar Excel", command=self._importar_pedidos_excel, fg_color="#6366F1", hover_color="#4F46E5", corner_radius=8, height=36, font=ctk.CTkFont("Segoe UI",11,"bold")).pack(side="left", pady=4, padx=(6,0))
        ctk.CTkLabel(rod, text="Duplo-clique para marcar como chegado", font=ctk.CTkFont("Segoe UI",10), text_color=CLR_MUTED).pack(side="right")
        self.tree_p.bind("<Double-1>", self._toggle_chegado)

    # ══════════════════════════════════════════════════════════
    # ABA 3: PREVISÃO POR SKU
    # ══════════════════════════════════════════════════════════

    def _build_tab_previsao(self, parent):
        top = ctk.CTkFrame(parent, fg_color=BG_SIDEBAR, corner_radius=12, border_width=1, border_color=CLR_BORDER)
        top.pack(fill="x", padx=10, pady=(10,10))
        top_inner = ctk.CTkFrame(top, fg_color="transparent")
        top_inner.pack(padx=14, pady=12, fill="x")
        ctk.CTkLabel(top_inner, text="SKU para análise:", font=ctk.CTkFont("Segoe UI",12,"bold"), text_color=CLR_TEXT).pack(side="left",padx=(0,8))
        self.prev_sku = ctk.CTkComboBox(top_inner, values=SKU_LIST, width=240, corner_radius=8, fg_color="#FFFFFF", font=ctk.CTkFont("Segoe UI",12))
        self.prev_sku.set(SKU_LIST[0] if SKU_LIST else "")
        self.prev_sku.pack(side="left", padx=8)
        ctk.CTkButton(top_inner, text="📈 Analisar", command=self._analisar_sku, fg_color=CLR_ACCENT, hover_color="#2563EB", corner_radius=8, height=40, font=ctk.CTkFont("Segoe UI",12,"bold")).pack(side="left", padx=8)

        self.prev_frame = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        self.prev_frame.pack(fill="both", expand=True, padx=10, pady=(0,10))

    # ══════════════════════════════════════════════════════════
    # ABA 4: RESUMO POR CATEGORIA
    # ══════════════════════════════════════════════════════════

    def _build_tab_resumo_cat(self, parent):
        self.resumo_cat_frame = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        self.resumo_cat_frame.pack(fill="both", expand=True, padx=10, pady=10)

    def _popular_resumo_cat(self):
        for w in self.resumo_cat_frame.winfo_children(): w.destroy()
        ctk.CTkLabel(self.resumo_cat_frame, text="📈 Resumo por Categoria", font=ctk.CTkFont("Segoe UI",16,"bold"), text_color=CLR_TEXT).pack(anchor="w", pady=(0,12))

        cats = defaultdict(lambda: {"total":0, "disp":0, "vendas":0, "ruptura":0, "urgente":0, "pedidos":0})
        hoje = date.today()

        for it in ITENS:
            sku, cat = it["sku"], it["cat"]
            c = cats[cat]
            c["total"] += 1
            e = self.estoque.get(sku, {})
            d = e.get("disp", 0)
            try: c["disp"] += float(d) if d not in ("—","") else 0
            except: pass
            v = vendas_mes_atual(sku, self.vendas)
            try: c["vendas"] += float(v) if v not in ("—","") else 0
            except: pass
            _pv = self.local["config"].get(sku, {}).get("proj_venda", "")
            _pm = self.local["config"].get(sku, {}).get("proj_mensal", {})
            proj = estoque_projetado(sku, d, v, self.local["pedidos"], proj_venda=_pv, proj_mensal=_pm)
            if any(p[1] <= 0 for p in proj): c["ruptura"] += 1
            er = _pv
            nec = necessidade_calc(d, er) if er else 0
            if nec >= 5: c["urgente"] += 1
            c["pedidos"] += sum(1 for p in self.local["pedidos"] if p.get("sku","").upper()==sku.upper() and not p.get("chegado"))

        for cat in sorted(cats, key=lambda x: CAT_ORDEM.get(x, 9)):
            c = cats[cat]
            cor = CAT_COLORS.get(cat, {}).get("badge", "#6B7280")
            card = ctk.CTkFrame(self.resumo_cat_frame, fg_color=BG_CARD, corner_radius=12, border_width=1, border_color=CLR_BORDER)
            card.pack(fill="x", pady=8)
            bar = ctk.CTkFrame(card, fg_color=cor, corner_radius=0, width=5)
            bar.pack(side="left", fill="y", padx=(0,0), pady=8)
            inner = ctk.CTkFrame(card, fg_color="transparent")
            inner.pack(fill="x", expand=True, padx=16, pady=14)
            ctk.CTkLabel(inner, text=f"  {cat}", font=ctk.CTkFont("Segoe UI",14,"bold"), text_color=cor).pack(anchor="w")

            metrics = ctk.CTkFrame(inner, fg_color="transparent")
            metrics.pack(fill="x", pady=(10,0))
            data_items = [
                ("SKUs", fmt_num(c["total"]), CLR_TEXT),
                ("Estoque", fmt_num(c["disp"]), CLR_ACCENT),
                ("Vendas/mês", fmt_num(c["vendas"]), CLR_TEXT), 
                ("Risco Ruptura", fmt_num(c["ruptura"]), CLR_RED if c["ruptura"]>0 else CLR_GREEN),
                ("Compra Urgente", fmt_num(c["urgente"]), CLR_AMBER if c["urgente"]>0 else CLR_GREEN),
                ("Pedidos Abertos", fmt_num(c["pedidos"]), CLR_PURPLE),
            ]
            for label, val, cor_val in data_items:
                col_frame = ctk.CTkFrame(metrics, fg_color="transparent")
                col_frame.pack(side="left", expand=True)
                ctk.CTkLabel(col_frame, text=val, font=ctk.CTkFont("Segoe UI",20,"bold"), text_color=cor_val).pack()
                ctk.CTkLabel(col_frame, text=label, font=ctk.CTkFont("Segoe UI",10), text_color=CLR_MUTED).pack()

    # ══════════════════════════════════════════════════════════
    # ABA 5: CENÁRIO DE SIMULAÇÃO
    # ══════════════════════════════════════════════════════════

    def _meses_opcoes(self, n=18):
        """Retorna lista de strings 'Abr/2026' para os próximos N meses."""
        hoje = date.today()
        opcoes = []
        for i in range(1, n + 1):
            m = (hoje.month - 1 + i) % 12 + 1
            a = hoje.year + (hoje.month - 1 + i) // 12
            opcoes.append(f"{MESES_PT_CURTO[m]}/{a}")
        return opcoes

    def _mes_label_para_data(self, mes_label):
        """Converte 'Abr/2026' → '01/04/2026'. Retorna None se inválido."""
        try:
            partes = mes_label.strip().split("/")
            if len(partes) != 2:
                return None
            nome, ano = partes[0].strip(), int(partes[1].strip())
            mes = next((i for i, n in enumerate(MESES_PT_CURTO) if n.lower() == nome.lower()), None)
            if mes is None or mes == 0:
                return None
            return f"01/{mes:02d}/{ano}"
        except Exception:
            return None

    def _build_tab_cenario(self, parent):
        # ══ FAIXA DE CONTROLE ═══════════════════════════════════
        ctrl = ctk.CTkFrame(parent, fg_color=BG_SIDEBAR, corner_radius=12, border_width=1, border_color=CLR_BORDER)
        ctrl.pack(fill="x", padx=8, pady=(8,4))

        # ── linha 1: compra manual + simular ────────────────────
        row1 = ctk.CTkFrame(ctrl, fg_color="transparent")
        row1.pack(fill="x", padx=10, pady=(10,6))

        ctk.CTkLabel(row1, text="📦", font=ctk.CTkFont("Segoe UI",14)).pack(side="left")
        ctk.CTkLabel(row1, text="Compra Manual", font=ctk.CTkFont("Segoe UI",11,"bold"), text_color=CLR_PURPLE).pack(side="left", padx=(2,14))

        ctk.CTkLabel(row1, text="SKU:", font=ctk.CTkFont("Segoe UI",10), text_color=CLR_MUTED).pack(side="left")
        self.cen_sku = ctk.CTkComboBox(row1, values=SKU_LIST, width=200, height=38, corner_radius=8, fg_color="#FFFFFF", font=ctk.CTkFont("Segoe UI",11))
        self.cen_sku.set("")
        self.cen_sku.pack(side="left", padx=4)

        ctk.CTkLabel(row1, text="Qtde:", font=ctk.CTkFont("Segoe UI",10), text_color=CLR_MUTED).pack(side="left", padx=(8,0))
        self.cen_qty = ctk.CTkEntry(row1, placeholder_text="Ex: 500", width=90, height=38, corner_radius=8, fg_color="#FFFFFF", font=ctk.CTkFont("Segoe UI",11))
        self.cen_qty.pack(side="left", padx=4)

        ctk.CTkButton(row1, text="＋ Adicionar", command=self._add_cenario_manual, fg_color=CLR_PURPLE, hover_color="#7C3AED", corner_radius=8, height=38, font=ctk.CTkFont("Segoe UI",11,"bold")).pack(side="left", padx=(10,0))

        # ── linha 1b: seletor de meses (pills) ──────────────────
        row1b = ctk.CTkFrame(ctrl, fg_color="transparent")
        row1b.pack(fill="x", padx=10, pady=(0,8))
        ctk.CTkLabel(row1b, text="📅 Chegada:", font=ctk.CTkFont("Segoe UI",10,"bold"), text_color=CLR_MUTED).pack(side="left", padx=(0,8))
        self._cen_meses_sel   = set()   # meses selecionados
        self._cen_meses_btns  = {}      # mes_label → CTkButton
        for _ml in self._meses_opcoes(12):
            _b = ctk.CTkButton(
                row1b, text=_ml, width=68, height=28, corner_radius=14,
                fg_color="#E5E7EB", text_color="#374151", hover_color="#D1D5DB",
                font=ctk.CTkFont("Segoe UI", 10),
                command=lambda ml=_ml: self._toggle_mes_cen(ml)
            )
            _b.pack(side="left", padx=2)
            self._cen_meses_btns[_ml] = _b

        ctk.CTkFrame(row1, fg_color=CLR_BORDER, width=2, height=34).pack(side="left", padx=18)

        ctk.CTkButton(row1, text="▶  Simular", command=self._executar_cenario, fg_color=CLR_GREEN, hover_color="#059669", corner_radius=8, height=38, width=110, font=ctk.CTkFont("Segoe UI",12,"bold")).pack(side="left", padx=4)
        ctk.CTkButton(row1, text="🔲 Grade", command=self._simular_e_abrir_grade, fg_color="#1E293B", hover_color="#334155", corner_radius=8, height=38, width=100, font=ctk.CTkFont("Segoe UI",11,"bold")).pack(side="left", padx=4)
        ctk.CTkButton(row1, text="🗑 Limpar", command=self._limpar_cenario, fg_color="#FEE2E2", text_color=CLR_RED, hover_color="#FECACA", corner_radius=8, height=38, width=90, font=ctk.CTkFont("Segoe UI",11)).pack(side="left", padx=4)

        # ── linha 2: gerar estoque ideal (destaque verde) ────────
        row2 = ctk.CTkFrame(ctrl, fg_color="#F0FDF4", corner_radius=8, border_width=1, border_color="#A7F3D0")
        row2.pack(fill="x", padx=10, pady=(0,10))

        ctk.CTkLabel(row2, text="🎯", font=ctk.CTkFont("Segoe UI",16)).pack(side="left", padx=(12,4), pady=8)
        ctk.CTkLabel(row2, text="Cenário Ideal", font=ctk.CTkFont("Segoe UI",12,"bold"), text_color="#065F46").pack(side="left", padx=(0,16))

        ctk.CTkLabel(row2, text="Chegada:", font=ctk.CTkFont("Segoe UI",10), text_color=CLR_MUTED).pack(side="left", padx=(0,4))
        _meses_ideal = self._meses_opcoes()
        self.ideal_data_var = tk.StringVar(value=_meses_ideal[0] if _meses_ideal else "")
        self.ideal_data = ctk.CTkComboBox(row2, variable=self.ideal_data_var, values=_meses_ideal,
                                          width=130, height=36, corner_radius=8,
                                          fg_color="#FFFFFF", font=ctk.CTkFont("Segoe UI",11))
        self.ideal_data.pack(side="left", padx=(0,14))

        ctk.CTkLabel(row2, text="SKU:", font=ctk.CTkFont("Segoe UI",10), text_color=CLR_MUTED).pack(side="left", padx=(0,4))
        self.ideal_sku_var = tk.StringVar(value="TODOS")
        ctk.CTkComboBox(row2, variable=self.ideal_sku_var, values=["TODOS"] + SKU_LIST, width=200, height=36, corner_radius=8, fg_color="#FFFFFF", font=ctk.CTkFont("Segoe UI",11)).pack(side="left", padx=(0,14))

        ctk.CTkButton(row2, text="🎯  Gerar Cenário Ideal", command=self._gerar_estoque_ideal, fg_color="#065F46", hover_color="#047857", corner_radius=8, height=36, font=ctk.CTkFont("Segoe UI",11,"bold")).pack(side="left", padx=(0,12))

        ctk.CTkLabel(row2, text="Proj > 15k → 2× mês  |  Proj ≤ 15k → cobre ruptura", font=ctk.CTkFont("Segoe UI",9), text_color="#6B7280").pack(side="left", padx=(0,8))

        self.ideal_status = ctk.CTkLabel(row2, text="", font=ctk.CTkFont("Segoe UI",10,"bold"), text_color="#065F46")
        self.ideal_status.pack(side="left", padx=4)

        # variáveis internas mantidas para compatibilidade com métodos existentes
        self.cob_status   = ctk.CTkLabel(ctrl, text="", font=ctk.CTkFont("Segoe UI",9))  # oculto
        self.cen_nome     = ctk.CTkEntry(ctrl, width=1); self.cen_nome.place_forget()
        self.cen_load_var = tk.StringVar(value="— Selecionar —")
        self.cen_load_cb  = ctk.CTkComboBox(ctrl, variable=self.cen_load_var, values=["— Selecionar —"], width=1)
        self.cen_load_cb.place_forget()
        self.dem_sku_var  = tk.StringVar(value="TODOS")
        self.dem_mult     = ctk.CTkEntry(ctrl, width=1); self.dem_mult.place_forget()
        self._dem_mes_vars = []
        self._dem_mes_btns = []
        self.dem_status   = ctk.CTkLabel(ctrl, text="")
        self._atualizar_lista_cenarios_salvos()

        # ══ LISTA DE COMPRAS NO CENÁRIO ═══════════════════════════
        lista_outer = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=10, border_width=1, border_color=CLR_BORDER)
        lista_outer.pack(fill="x", padx=8, pady=(0,4))

        lista_hdr = ctk.CTkFrame(lista_outer, fg_color="transparent")
        lista_hdr.pack(fill="x", padx=10, pady=(6,2))
        ctk.CTkLabel(lista_hdr, text="📋 Compras no Cenário", font=ctk.CTkFont("Segoe UI",11,"bold"), text_color=CLR_TEXT).pack(side="left")
        self.cen_status = ctk.CTkLabel(lista_hdr, text="Adicione compras acima e clique ▶ Simular.", font=ctk.CTkFont("Segoe UI",10), text_color=CLR_MUTED)
        self.cen_status.pack(side="left", padx=12)
        self.cob_status = ctk.CTkLabel(lista_hdr, text="", font=ctk.CTkFont("Segoe UI",10,"bold"), text_color=CLR_GREEN)
        self.cob_status.pack(side="left")
        ctk.CTkButton(lista_hdr, text="🗑 Remover", command=self._del_cenario, fg_color="#FEE2E2", text_color=CLR_RED, hover_color="#FECACA", corner_radius=6, height=26, font=ctk.CTkFont("Segoe UI",10)).pack(side="right", padx=4)
        ctk.CTkButton(lista_hdr, text="🗑 Excluir Todos", command=self._del_cenario_todos, fg_color=CLR_RED, text_color="white", hover_color="#B91C1C", corner_radius=6, height=26, font=ctk.CTkFont("Segoe UI",10,"bold")).pack(side="right", padx=4)
        ctk.CTkButton(lista_hdr, text="📋 Exportar Lista", command=self._exportar_lista_compras, fg_color="#10B981", hover_color="#059669", corner_radius=6, height=26, font=ctk.CTkFont("Segoe UI",10,"bold")).pack(side="right", padx=4)

        # salvar / carregar cenário
        ctk.CTkFrame(lista_hdr, fg_color=CLR_BORDER, width=2, height=22).pack(side="right", padx=8)
        ctk.CTkButton(lista_hdr, text="📂 Carregar", command=self._popup_carregar_cenario, fg_color="#E0E7FF", text_color="#3730A3", hover_color="#C7D2FE", corner_radius=6, height=26, font=ctk.CTkFont("Segoe UI",10,"bold")).pack(side="right", padx=4)
        ctk.CTkButton(lista_hdr, text="💾 Salvar", command=self._popup_salvar_cenario, fg_color="#3730A3", text_color="white", hover_color="#312E81", corner_radius=6, height=26, font=ctk.CTkFont("Segoe UI",10,"bold")).pack(side="right", padx=4)
        ctk.CTkButton(lista_hdr, text="📊 Excel", command=self._exportar_cenario_excel, fg_color="#1D4ED8", hover_color="#1E40AF", corner_radius=6, height=26, font=ctk.CTkFont("Segoe UI",10,"bold")).pack(side="right", padx=4)
        ctk.CTkButton(lista_hdr, text="🔲 Tela Cheia", command=lambda: self._abrir_grade_tela_cheia(modo="cenario"), fg_color="#1E293B", hover_color="#334155", corner_radius=6, height=26, font=ctk.CTkFont("Segoe UI",10)).pack(side="right", padx=4)

        sc = ttk.Style()
        sc.configure("C.Treeview", background=BG_CARD, foreground=CLR_TEXT, rowheight=26, fieldbackground=BG_CARD, font=("Segoe UI",10), borderwidth=0)
        sc.configure("C.Treeview.Heading", background="#5B21B6", foreground="white", font=("Segoe UI",10,"bold"), relief="flat", padding=(6,4))
        sc.map("C.Treeview", background=[("selected","#EDE9FE")], foreground=[("selected","#4C1D95")])
        sc.layout("C.Treeview", [("C.Treeview.treearea",{"sticky":"nswe"})])

        cen_vsb = ttk.Scrollbar(lista_outer, orient="vertical")
        self.tree_c = ttk.Treeview(lista_outer, columns=("sku","cat","qty","data","obs"), show="headings", style="C.Treeview", yscrollcommand=cen_vsb.set, height=3)
        cen_vsb.config(command=self.tree_c.yview)
        for col, hdr, w, anc in [("sku","SKU",180,"w"),("cat","Categoria",120,"w"),("qty","Quantidade",110,"center"),("data","Data Chegada",120,"center"),("obs","Observação",160,"w")]:
            self.tree_c.heading(col, text=f"  {hdr}", anchor="w")
            self.tree_c.column(col, width=w, minwidth=50, anchor=anc)
        self.tree_c.tag_configure("manual", background="#F0FDF4")
        self.tree_c.tag_configure("cobertura", background="#F5F3FF")
        cen_vsb.pack(side="right", fill="y")
        self.tree_c.pack(fill="x", padx=2, pady=(0,4))

        # ══ LEGENDA ══════════════════════════════════════════════
        leg_row = ctk.CTkFrame(parent, fg_color="transparent")
        leg_row.pack(fill="x", padx=10, pady=(0,2))
        for txt, cor in [("🟩 Estoque ≥ 0", "#16A34A"), ("🟥 Ruptura < 0", "#DC2626"), ("▲ Chegada Simulada", "#000000"), ("Atual = sem simulação  |  Cen. = com simulação", "#6B7280")]:
            ctk.CTkLabel(leg_row, text=txt, font=ctk.CTkFont("Segoe UI",9), text_color=cor).pack(side="left", padx=10)

        # ══ GRADE ════════════════════════════════════════════════
        grade_card = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=12, border_width=1, border_color=CLR_BORDER)
        grade_card.pack(fill="both", expand=True, padx=8, pady=(0,8))

        vbar = ttk.Scrollbar(grade_card, orient="vertical")
        hbar = ttk.Scrollbar(grade_card, orient="horizontal")
        self.cen_grade = GradeSimulacao(grade_card)
        vbar.config(command=self.cen_grade.yview)
        hbar.config(command=self.cen_grade.xview)
        self.cen_grade.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)

        vbar.pack(side="right",  fill="y")
        hbar.pack(side="bottom", fill="x")
        self.cen_grade.pack(fill="both", expand=True, padx=2, pady=2)

        self.cen_grade.bind("<MouseWheel>", lambda e: self.cen_grade.yview_scroll(-1*(e.delta//120), "units"))
        self.cen_grade.bind("<Shift-MouseWheel>", lambda e: self.cen_grade.xview_scroll(-1*(e.delta//120), "units"))

        self.cen_grade.create_text(20, 20, text="Adicione compras acima e clique  ▶ Simular  para visualizar o estoque projetado", fill="#9CA3AF", font=("Segoe UI", 12), anchor="nw")
        self.cen_grade.configure(scrollregion=(0, 0, 700, 100))

    # ══════════════════════════════════════════════════════════
    # ABA: VISUAL ATUAL
    # ══════════════════════════════════════════════════════════

    def _build_tab_visual_atual(self, parent):
        # barra de controle
        tb = ctk.CTkFrame(parent, fg_color=BG_SIDEBAR, corner_radius=12, border_width=1, border_color=CLR_BORDER)
        tb.pack(fill="x", padx=8, pady=(8,4))

        ctk.CTkLabel(tb, text="🗺️ Visual Atual — Estoque Real + Pedidos Confirmados",
                     font=ctk.CTkFont("Segoe UI",12,"bold"), text_color=CLR_TEXT).pack(side="left", padx=14, pady=10)

        ctk.CTkButton(tb, text="🔲 Tela Cheia", command=lambda: self._abrir_grade_tela_cheia(modo="estoque"),
                      fg_color="#1E293B", hover_color="#334155", corner_radius=8,
                      height=36, width=120, font=ctk.CTkFont("Segoe UI",11,"bold")).pack(side="right", padx=(0,10), pady=6)

        ctk.CTkButton(tb, text="🔄 Atualizar", command=self._popular_visual_atual,
                      fg_color=CLR_ACCENT, hover_color="#2563EB", corner_radius=8,
                      height=36, width=110, font=ctk.CTkFont("Segoe UI",11,"bold")).pack(side="right", padx=10, pady=6)

        # legenda
        leg = ctk.CTkFrame(parent, fg_color="transparent")
        leg.pack(fill="x", padx=10, pady=(2,4))
        for txt, cor in [("🟩 Estoque Positivo","#16A34A"), ("🟥 Ruptura < 0","#DC2626"),
                         ("📦 Chegada confirmada (pedidos reais)","#065F46")]:
            ctk.CTkLabel(leg, text=txt, font=ctk.CTkFont("Segoe UI",10,"bold"), text_color=cor).pack(side="left", padx=10)

        # grade
        grade_card = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=12,
                                   border_width=1, border_color=CLR_BORDER)
        grade_card.pack(fill="both", expand=True, padx=8, pady=(0,8))

        vbar = ttk.Scrollbar(grade_card, orient="vertical")
        hbar = ttk.Scrollbar(grade_card, orient="horizontal")
        self.visual_grade = GradeSimulacao(grade_card)
        vbar.config(command=self.visual_grade.yview)
        hbar.config(command=self.visual_grade.xview)
        self.visual_grade.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)

        vbar.pack(side="right", fill="y")
        hbar.pack(side="bottom", fill="x")
        self.visual_grade.pack(fill="both", expand=True, padx=2, pady=2)

        self.visual_grade.bind("<MouseWheel>",
            lambda e: self.visual_grade.yview_scroll(-1*(e.delta//120), "units"))
        self.visual_grade.bind("<Shift-MouseWheel>",
            lambda e: self.visual_grade.xview_scroll(-1*(e.delta//120), "units"))

        self.visual_grade.create_text(20, 20,
            text="Clique em 🔄 Atualizar para carregar o visual do estoque atual",
            fill="#9CA3AF", font=("Segoe UI", 12), anchor="nw")
        self.visual_grade.configure(scrollregion=(0, 0, 700, 60))

    def _popular_visual_atual(self):
        rows, meses = self._preparar_dados_grade(modo="estoque")
        self.visual_grade.render(rows, meses, modo="estoque")

    def _on_tab_change(self):
        tab = self.tabs.get()
        if tab == "🗺️ Visual Atual":
            self._popular_visual_atual()

    def _toggle_mes_cen(self, mes_label):
        """Liga/desliga seleção de um mês no seletor de pills."""
        btn = self._cen_meses_btns.get(mes_label)
        if mes_label in self._cen_meses_sel:
            self._cen_meses_sel.discard(mes_label)
            if btn:
                btn.configure(fg_color="#E5E7EB", text_color="#374151", hover_color="#D1D5DB",
                               font=ctk.CTkFont("Segoe UI", 10))
        else:
            self._cen_meses_sel.add(mes_label)
            if btn:
                btn.configure(fg_color=CLR_PURPLE, text_color="white", hover_color="#7C3AED",
                               font=ctk.CTkFont("Segoe UI", 10, "bold"))

    def _add_cenario_manual(self):
        sku  = self.cen_sku.get().strip().upper()
        qty  = self.cen_qty.get().strip()
        meses_sel = sorted(self._cen_meses_sel,
                           key=lambda ml: self._mes_label_para_data(ml) or "")

        if not sku or not qty:
            messagebox.showwarning("Campos obrigatórios", "Preencha SKU e Quantidade.")
            return
        if not meses_sel:
            messagebox.showwarning("Selecione o mês", "Clique em pelo menos um mês de chegada (pills roxas).")
            return
        if sku not in [i["sku"].upper() for i in ITENS]:
            messagebox.showwarning("SKU inválido", f"SKU '{sku}' não encontrado.")
            return
        try:
            qty_f = float(qty.replace(",","."))
            if qty_f <= 0: raise ValueError
        except:
            messagebox.showwarning("Quantidade inválida", "Digite um número positivo.")
            return

        cat = SKU_CAT.get(sku, "—")
        for mes_label in meses_sel:
            data = self._mes_label_para_data(mes_label)
            if not data or not parse_date(data):
                continue
            self.cenario.append({"sku": sku, "qty": str(qty_f), "data": data,
                                  "proj_venda": "", "obs": "MANUAL", "chegado": False})
            idx = len(self.cenario) - 1
            self.tree_c.insert("", "end", iid=f"_cen_{idx}",
                               values=(sku, cat, f"{fmt_num(qty_f)} un", mes_label, "Manual"),
                               tags=("manual",))

        # limpa seleção de meses após adicionar
        for ml in list(self._cen_meses_sel):
            self._toggle_mes_cen(ml)

        self.cen_sku.set("")
        self.cen_qty.delete(0, "end")
        n = len(self.cenario)
        self.cen_status.configure(text=f"{n} item(ns) no cenário. Clique ▶ Simular.", text_color=CLR_TEXT)

    def _del_cenario(self):
        sel = self.tree_c.selection()
        if not sel:
            messagebox.showinfo("Selecione", "Clique em um item da lista para selecionar antes de remover.")
            return
        iid = sel[0]
        # descobre índice pelo iid "_cen_N"
        try:
            idx = int(iid.split("_cen_")[1])
            if 0 <= idx < len(self.cenario):
                self.cenario.pop(idx)
                # re-indexa iids restantes
                filhos = list(self.tree_c.get_children())
                self.tree_c.delete(iid)
                for i, fid in enumerate(self.tree_c.get_children()):
                    self.tree_c.item(fid, tags=(self.tree_c.item(fid)["tags"],))
            else:
                self.tree_c.delete(iid)
        except Exception:
            self.tree_c.delete(iid)
        n = len(self.cenario)
        self.cen_status.configure(text=f"{n} item(ns) no cenário.", text_color=CLR_TEXT if n else CLR_MUTED)

    def _popup_salvar_cenario(self):
        if not self.cenario:
            messagebox.showinfo("Vazio", "Adicione itens ao cenário antes de salvar.")
            return
        dlg = ctk.CTkToplevel(self)
        dlg.title("Salvar Cenário")
        dlg.geometry("360x160")
        dlg.resizable(False, False)
        dlg.after(50, lambda: (dlg.lift(), dlg.focus_force(), dlg.grab_set()))
        ctk.CTkLabel(dlg, text="💾 Nome do cenário", font=ctk.CTkFont("Segoe UI",13,"bold"), text_color=CLR_TEXT).pack(pady=(20,8))
        e_nome = ctk.CTkEntry(dlg, placeholder_text="Ex: Cenário Junho 2026", width=280, height=38, corner_radius=8, fg_color="#FFFFFF", font=ctk.CTkFont("Segoe UI",12))
        e_nome.pack()
        def _salvar():
            nome = e_nome.get().strip()
            if not nome:
                messagebox.showwarning("Nome obrigatório", "Digite um nome para o cenário.", parent=dlg)
                return
            if "cenarios" not in self.local:
                self.local["cenarios"] = {}
            if nome in self.local["cenarios"]:
                if not messagebox.askyesno("Sobrescrever?", f"Já existe '{nome}'. Deseja sobrescrever?", parent=dlg):
                    return
            self.local["cenarios"][nome] = [dict(p) for p in self.cenario]
            salvar_local(self.local)
            dlg.destroy()
            messagebox.showinfo("Salvo", f"Cenário '{nome}' salvo com sucesso.")
        ctk.CTkButton(dlg, text="💾 Salvar", command=_salvar, fg_color="#3730A3", hover_color="#312E81", corner_radius=8, height=36, font=ctk.CTkFont("Segoe UI",12,"bold")).pack(pady=14)
        e_nome.bind("<Return>", lambda _: _salvar())

    def _popup_carregar_cenario(self):
        cenarios = self.local.get("cenarios", {})
        if not cenarios:
            messagebox.showinfo("Sem cenários", "Nenhum cenário salvo ainda.")
            return
        dlg = ctk.CTkToplevel(self)
        dlg.title("Carregar Cenário")
        dlg.geometry("420x360")
        dlg.resizable(False, True)
        dlg.after(50, lambda: (dlg.lift(), dlg.focus_force(), dlg.grab_set()))
        ctk.CTkLabel(dlg, text="📂 Selecione um cenário", font=ctk.CTkFont("Segoe UI",13,"bold"), text_color=CLR_TEXT).pack(pady=(16,8))

        frame_lista = ctk.CTkScrollableFrame(dlg, fg_color="transparent", height=220)
        frame_lista.pack(fill="both", expand=True, padx=16)

        sel_var = tk.StringVar()
        for nome in cenarios:
            n_itens = len(cenarios[nome])
            ctk.CTkRadioButton(frame_lista, text=f"{nome}  ({n_itens} itens)", variable=sel_var, value=nome,
                               font=ctk.CTkFont("Segoe UI",11), text_color=CLR_TEXT).pack(anchor="w", pady=3)

        btn_row = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_row.pack(pady=(8,16))

        def _carregar():
            nome = sel_var.get()
            if not nome:
                messagebox.showwarning("Selecione", "Escolha um cenário.", parent=dlg)
                return
            self._limpar_cenario()
            for p in cenarios[nome]:
                self.cenario.append(dict(p))
                cat = SKU_CAT.get(p.get("sku",""), "—")
                idx = len(self.cenario) - 1
                self.tree_c.insert("", "end", iid=f"_cen_{idx}",
                                   values=(p.get("sku",""), cat, f"{fmt_num(float(p.get('qty',0)))} un",
                                           p.get("data",""), p.get("obs","")),
                                   tags=("cobertura",))
            self.cen_status.configure(text=f"Cenário '{nome}' carregado — {len(self.cenario)} item(ns). Clique ▶ Simular.", text_color=CLR_TEXT)
            dlg.destroy()

        def _excluir():
            nome = sel_var.get()
            if not nome:
                messagebox.showwarning("Selecione", "Escolha um cenário.", parent=dlg)
                return
            if not messagebox.askyesno("Confirmar", f"Excluir o cenário '{nome}'?", parent=dlg):
                return
            self.local.get("cenarios", {}).pop(nome, None)
            salvar_local(self.local)
            dlg.destroy()

        ctk.CTkButton(btn_row, text="📂 Carregar", command=_carregar, fg_color="#3730A3", hover_color="#312E81", corner_radius=8, height=34, font=ctk.CTkFont("Segoe UI",11,"bold")).pack(side="left", padx=6)
        ctk.CTkButton(btn_row, text="🗑 Excluir", command=_excluir, fg_color="#FEE2E2", text_color=CLR_RED, hover_color="#FECACA", corner_radius=8, height=34, font=ctk.CTkFont("Segoe UI",11)).pack(side="left", padx=6)

    def _del_cenario_todos(self):
        if not self.cenario:
            messagebox.showinfo("Vazio", "Não há itens no cenário.")
            return
        if not messagebox.askyesno("Confirmar", f"Excluir todos os {len(self.cenario)} itens do cenário?"):
            return
        self.cenario.clear()
        for iid in self.tree_c.get_children():
            self.tree_c.delete(iid)
        self.cen_status.configure(text="Cenário limpo.", text_color=CLR_MUTED)
        self.ideal_status.configure(text="")

    def _aplicar_cobertura(self):
        data = self.cob_data.get().strip()
        if not data or not parse_date(data):
            messagebox.showwarning("Data inválida", "Informe a data de chegada no formato DD/MM/AAAA.")
            return
        txt_meses = self.cob_meses.get().strip().replace(",",".")
        try:
            meses_desejados = float(txt_meses)
            if meses_desejados <= 0: raise ValueError
        except:
            messagebox.showwarning("Meses inválido", "Digite um número positivo de meses de cobertura.")
            return

        sku_sel = self.cob_sku_var.get().strip().upper()
        skus = [i["sku"] for i in ITENS] if sku_sel == "TODOS" else [sku_sel]

        adicionados = 0
        ignorados   = 0
        for sku in skus:
            item = next((i for i in ITENS if i["sku"].upper() == sku.upper()), None)
            if not item:
                continue
            disp = item.get("disponivel","") or item.get("disp","") or "0"
            try:
                disp_f = float(str(disp).replace(",","."))
            except:
                disp_f = 0.0

            cfg     = self.local.get("config", {}).get(sku, {})
            er_cfg  = str(cfg.get("proj_venda","")).strip()
            pm_cfg  = cfg.get("proj_mensal", {})
            vend    = item.get("vendas_mes","") or "0"

            if er_cfg:
                try:    consumo = float(er_cfg)
                except: consumo = 0.0
            else:
                try:    consumo = float(str(vend).replace(",","."))
                except: consumo = 0.0

            if consumo <= 0:
                ignorados += 1
                continue

            cobertura_atual = disp_f / consumo
            qty = max(0.0, (meses_desejados - cobertura_atual) * consumo)
            if qty < 0.01:
                ignorados += 1
                continue

            self.cenario.append({
                "sku": sku,
                "qty": str(round(qty)),
                "data": data,
                "proj_venda": "",
                "obs": f"COBERTURA {meses_desejados}m",
                "chegado": False,
            })
            adicionados += 1

        if adicionados == 0:
            self.cob_status.configure(text=f"⚠ Nenhum SKU necessita reposição para {meses_desejados} meses ({ignorados} ignorado(s).", text_color=CLR_RED)
        else:
            self.cob_status.configure(text=f"✅ {adicionados} SKU(s) adicionado(s) | cobertura {meses_desejados} meses | chegada {data}", text_color=CLR_GREEN)
        self.cen_status.configure(text=f"{adicionados} item(ns) no cenário. Clique em ▶ Executar para visualizar.")

    def _limpar_cenario(self):
        self.cenario.clear()
        self.cen_grade.delete("all")
        self.cob_status.configure(text="")
        self.cen_status.configure(text="Cenário limpo. Configure a cobertura desejada e clique em Aplicar.")

    def _gerar_estoque_ideal(self):
        """Gera compras para estoque ideal baseado na projeção de venda.
        - Proj > 15.000 kg/mês → target = 2× proj (vende 1 mês + sobra 1 mês)
        - Proj ≤ 15.000 kg/mês → apenas cobre rupturas futuras (sem buffer extra)
        """
        mes_label = self.ideal_data_var.get().strip()
        data      = self._mes_label_para_data(mes_label)
        if not data or not parse_date(data):
            messagebox.showwarning("Mês inválido", "Selecione um mês de chegada válido (ex: Abr/2026).")
            return
        chegada_date = parse_date(data)

        sku_sel = self.ideal_sku_var.get().strip().upper()
        skus = [i["sku"] for i in ITENS] if sku_sel == "TODOS" else [sku_sel]

        adicionados = 0
        sem_necessidade = 0

        for sku in skus:
            e      = self.estoque.get(sku, {})
            disp   = e.get("disp", 0)
            vend   = vendas_mes_atual(sku, self.vendas)
            er_cfg = self.local["config"].get(sku, {}).get("proj_venda", "")
            pm_cfg = self.local["config"].get(sku, {}).get("proj_mensal", {})

            # projeção mensal efetiva (usa er_cfg se configurado, senão vendas reais do mês)
            try:
                proj_mensal = float(er_cfg) if str(er_cfg).strip() else float(vend) if vend else 0.0
            except (ValueError, TypeError):
                proj_mensal = 0.0

            # projeção dos 12 meses com pedidos reais + cenário já inserido
            pedidos_atuais = self.local["pedidos"] + self.cenario
            proj = estoque_projetado(sku, disp, vend, pedidos_atuais,
                                     proj_venda=er_cfg, proj_mensal=pm_cfg)

            # índice do mês de chegada na projeção
            chegada_idx = 0
            for i, (_, _, _, m, a) in enumerate(proj):
                if a > chegada_date.year or (a == chegada_date.year and m >= chegada_date.month):
                    chegada_idx = i
                    break

            saldo_no_mes_chegada = proj[chegada_idx][1] if chegada_idx < len(proj) else 0

            if proj_mensal > 15_000:
                # sobra 2 meses após vender o projetado
                target_end_saldo = proj_mensal * 2
            elif proj_mensal >= 5_000:
                # sobra 1 mês após vender o projetado
                target_end_saldo = proj_mensal
            else:
                # < 5k: só zera a ruptura
                target_end_saldo = 0

            qty_necessaria = max(0.0, target_end_saldo - saldo_no_mes_chegada)

            if qty_necessaria < 1:
                sem_necessidade += 1
                continue

            self.cenario.append({
                "sku":       sku,
                "qty":       str(int(round(qty_necessaria))),
                "data":      data,
                "proj_venda": "",
                "obs":       "IDEAL",
                "chegado":   False,
            })
            cat = SKU_CAT.get(sku, "—")
            idx = len(self.cenario) - 1
            self.tree_c.insert("", "end", iid=f"_cen_{idx}",
                               values=(sku, cat, f"{fmt_num(qty_necessaria)} un", data, "Ideal"),
                               tags=("cobertura",))
            adicionados += 1

        if adicionados == 0:
            self.ideal_status.configure(
                text=f"✅ Todos os {sem_necessidade} SKU(s) já estão em nível ideal a partir de {data}!")
        else:
            self.ideal_status.configure(
                text=f"✅ {adicionados} compra(s) gerada(s) — {sem_necessidade} SKU(s) já OK")
        self.cen_status.configure(
            text=f"{len(self.cenario)} item(ns) no cenário. Clique ▶ Simular para visualizar.",
            text_color=CLR_TEXT)

    # ── ajuste de demanda ─────────────────────────────────────
    def _dem_todos_meses(self):
        for var, _, _ in self._dem_mes_vars: var.set(True)

    def _dem_nenhum_mes(self):
        for var, _, _ in self._dem_mes_vars: var.set(False)

    def _aplicar_demanda(self):
        txt = self.dem_mult.get().strip().replace(",",".")
        try:
            mult = float(txt)
            if mult <= 0: raise ValueError
        except:
            messagebox.showwarning("Multiplicador inválido", "Digite um número positivo. Ex: 2  ou  0.5")
            return
        meses_sel = [key for var, key, _ in self._dem_mes_vars if var.get()]
        if not meses_sel:
            messagebox.showwarning("Nenhum mês selecionado", "Selecione ao menos um mês para aplicar o ajuste.")
            return
        sku = self.dem_sku_var.get().strip().upper()
        self.cen_demanda = {"sku": sku, "mult": mult, "meses": meses_sel}
        lbl_meses = ", ".join(lbl for var, _, lbl in self._dem_mes_vars if var.get())
        sku_txt = sku if sku != "TODOS" else "todos os SKUs"
        self.dem_status.configure(text=f"✅  Ajuste ativo: ×{mult} em {len(meses_sel)} mês(es) [{lbl_meses}] para {sku_txt}")

    def _limpar_demanda(self):
        self.cen_demanda = {"sku": "TODOS", "mult": 1.0, "meses": []}
        self.dem_mult.delete(0, "end")
        self._dem_nenhum_mes()
        self.dem_status.configure(text="")

    def _get_proj_mensal_sim(self, sku, er_base, pm_base):
        """Retorna proj_mensal efetivo aplicando o ajuste de demanda do cenário."""
        if not self.cen_demanda["meses"]:
            return er_base, pm_base
        sku_match = (self.cen_demanda["sku"] == "TODOS" or
                     self.cen_demanda["sku"].upper() == sku.upper())
        if not sku_match:
            return er_base, pm_base
        mult = self.cen_demanda["mult"]
        hoje = date.today()
        pm_sim = dict(pm_base) if pm_base else {}
        for i in range(12):
            m = (hoje.month - 1 + i) % 12 + 1
            a = hoje.year + (hoje.month - 1 + i) // 12
            key = f"{a}-{m:02d}"
            if key not in self.cen_demanda["meses"]:
                continue
            if key in pm_sim and str(pm_sim[key]).strip():
                try: base = float(pm_sim[key])
                except: base = float(er_base) if er_base else 0
            elif er_base:
                try: base = float(er_base)
                except: base = 0
            else:
                base = 0
            pm_sim[key] = str(round(base * mult, 4))
        return er_base, pm_sim

    # ── cenários nomeados ─────────────────────────────────────
    def _atualizar_lista_cenarios_salvos(self):
        salvos = list(self.local.get("cenarios", {}).keys())
        valores = ["— Selecionar —"] + salvos
        self.cen_load_cb.configure(values=valores)
        self.cen_load_var.set("— Selecionar —")

    def _salvar_cenario_nomeado(self):
        nome = self.cen_nome.get().strip()
        if not nome:
            messagebox.showwarning("Nome obrigatório", "Digite um nome para o cenário antes de salvar.")
            return
        if not self.cenario:
            messagebox.showwarning("Cenário vazio", "Adicione ao menos uma compra antes de salvar.")
            return
        if "cenarios" not in self.local:
            self.local["cenarios"] = {}
        if nome in self.local["cenarios"]:
            if not messagebox.askyesno("Sobrescrever?", f"Já existe um cenário '{nome}'. Deseja sobrescrever?"):
                return
        self.local["cenarios"][nome] = [dict(p) for p in self.cenario]
        salvar_local(self.local)
        self._atualizar_lista_cenarios_salvos()
        self.cen_nome.delete(0, "end")
        messagebox.showinfo("Salvo", f"Cenário '{nome}' salvo com sucesso.")

    def _carregar_cenario_nomeado(self, nome=None):
        if nome is None:
            nome = self.cen_load_var.get()
        if not nome or nome == "— Selecionar —":
            return
        cenarios = self.local.get("cenarios", {})
        if nome not in cenarios:
            return
        self._limpar_cenario()
        for p in cenarios[nome]:
            self.cenario.append(dict(p))
            cat = SKU_CAT.get(p.get("sku", ""), "—")
            idx = len(self.cenario) - 1
            data_s = p.get("data", "")
            # tenta converter data "01/04/2026" → "Abr/2026" para exibição
            try:
                from datetime import datetime as _dt
                _d = _dt.strptime(data_s, "%d/%m/%Y")
                data_disp = f"{MESES_PT_CURTO[_d.month]}/{_d.year}"
            except Exception:
                data_disp = data_s
            self.tree_c.insert("", "end", iid=f"_cen_{idx}",
                               values=(p.get("sku", ""), cat,
                                       f"{fmt_num(float(p.get('qty', 0)))} un",
                                       data_disp, p.get("obs", "")),
                               tags=("cobertura",))
        self.cen_status.configure(
            text=f"Cenário '{nome}' carregado — {len(self.cenario)} item(ns). Clique ▶ Simular.",
            text_color=CLR_TEXT)

    def _excluir_cenario_nomeado(self):
        nome = self.cen_load_var.get()
        if not nome or nome == "— Selecionar —":
            messagebox.showinfo("Selecione", "Selecione um cenário salvo para excluir.")
            return
        if not messagebox.askyesno("Confirmar", f"Excluir o cenário '{nome}'?"):
            return
        self.local.get("cenarios", {}).pop(nome, None)
        salvar_local(self.local)
        self._atualizar_lista_cenarios_salvos()

    def _exportar_lista_compras(self):
        if not HAS_OPENPYXL:
            messagebox.showerror("Dependência ausente", "Instale openpyxl:\n  pip install openpyxl")
            return
        if not self.cenario:
            messagebox.showinfo("Vazio", "Não há itens no cenário para exportar.")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            title="Exportar Lista de Compras Sugeridas",
            initialfile=f"compras_sugeridas_{date.today().strftime('%d%m%Y')}.xlsx"
        )
        if not filepath: return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Compras Sugeridas"

        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        def fill(hex6): return PatternFill("solid", fgColor=hex6)
        def bdr():
            s = Side(style="thin", color="D1D5DB")
            return Border(left=s, right=s, top=s, bottom=s)
        def ctr(): return Alignment(horizontal="center", vertical="center")
        def lft(): return Alignment(horizontal="left",   vertical="center")

        # título
        ws.merge_cells("A1:F1")
        c = ws["A1"]
        c.value = f"Lista de Compras Sugeridas — Gerado em {date.today().strftime('%d/%m/%Y')}"
        c.font = Font(bold=True, color="FFFFFF", size=12, name="Segoe UI")
        c.fill = fill("0F172A")
        c.alignment = ctr()
        ws.row_dimensions[1].height = 22

        # cabeçalho
        headers = ["SKU", "Categoria", "Quantidade (kg)", "Data Chegada", "Observação", "Proj. Venda/mês"]
        widths  = [22, 14, 18, 16, 16, 18]
        for ci, (h, w) in enumerate(zip(headers, widths), 1):
            cell = ws.cell(row=2, column=ci, value=h)
            cell.font      = Font(bold=True, color="FFFFFF", size=10, name="Segoe UI")
            cell.fill      = fill("1E3A8A")
            cell.alignment = ctr()
            cell.border    = bdr()
            ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = w
        ws.row_dimensions[2].height = 20

        # agrupa por data de chegada
        from collections import defaultdict as _dd
        grupos = _dd(list)
        for p in self.cenario:
            grupos[p.get("data","—")].append(p)

        ROW = 3
        total_geral = 0
        for data_ch in sorted(grupos.keys(), key=lambda d: parse_date(d) or date(2099,1,1)):
            ps = grupos[data_ch]
            # subtotal da data
            sub = sum(float(p.get("qty",0)) for p in ps)
            total_geral += sub

            # linha de agrupamento
            ws.merge_cells(f"A{ROW}:F{ROW}")
            gc = ws[f"A{ROW}"]
            gc.value = f"📅  Chegada: {data_ch}   —   {len(ps)} SKU(s)   |   Total: {fmt_num(sub)} kg"
            gc.font  = Font(bold=True, color="FFFFFF", size=10, name="Segoe UI")
            gc.fill  = fill("374151")
            gc.alignment = lft()
            ws.row_dimensions[ROW].height = 18
            ROW += 1

            for p in sorted(ps, key=lambda x: (SKU_CAT.get(x.get("sku",""),""), x.get("sku",""))):
                sku  = p.get("sku","")
                cat  = SKU_CAT.get(sku, "—")
                qty  = float(p.get("qty", 0))
                obs  = p.get("obs","")
                er   = self.local["config"].get(sku, {}).get("proj_venda","")
                pv   = fmt_num(er) if er else "—"

                bg_hex = "F0FDF4" if obs == "IDEAL" else "F5F3FF"
                row_fill = fill(bg_hex)

                vals = [sku, cat, qty, data_ch, obs, pv]
                for ci, v in enumerate(vals, 1):
                    cell = ws.cell(row=ROW, column=ci, value=v)
                    cell.fill      = row_fill
                    cell.border    = bdr()
                    cell.alignment = ctr() if ci != 1 else lft()
                    if ci == 1:
                        cell.font = Font(bold=True, name="Segoe UI", size=10)
                    if ci == 3:
                        cell.number_format = "#,##0"
                ws.row_dimensions[ROW].height = 17
                ROW += 1

        # linha de total geral
        ws.merge_cells(f"A{ROW}:B{ROW}")
        ws[f"A{ROW}"].value = "TOTAL GERAL"
        ws[f"A{ROW}"].font  = Font(bold=True, color="FFFFFF", name="Segoe UI", size=11)
        ws[f"A{ROW}"].fill  = fill("0F172A")
        ws[f"A{ROW}"].alignment = ctr()
        ws.cell(row=ROW, column=3, value=total_geral).number_format = "#,##0"
        ws.cell(row=ROW, column=3).font  = Font(bold=True, color="FFFFFF", name="Segoe UI", size=11)
        ws.cell(row=ROW, column=3).fill  = fill("0F172A")
        ws.cell(row=ROW, column=3).alignment = ctr()
        for ci in [4,5,6]:
            ws.cell(row=ROW, column=ci).fill = fill("0F172A")
        ws.row_dimensions[ROW].height = 20

        ws.freeze_panes = "A3"
        wb.save(filepath)
        messagebox.showinfo("Exportado", f"Lista exportada com sucesso:\n{filepath}")
        os.startfile(filepath)

    def _exportar_cenario_excel(self):
        if not HAS_OPENPYXL:
            messagebox.showerror("Dependência ausente", "Instale openpyxl:\n  pip install openpyxl")
            return
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], title="Exportar Projeção de Estoque — Excel")
        if not filepath: return

        rows, meses = self._preparar_dados_grade(modo="cenario" if self.cenario else "estoque")

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Projeção de Estoque"

        # ── paleta ────────────────────────────────────────────
        CLR_HDR_DARK  = "0B192C"
        CLR_HDR_MES   = "1E3A8A"
        CLR_HDR_SUB   = "2D4BA0"
        CLR_CAT_BG    = "1E3A8A"
        CLR_POS       = "D1F2D3"
        CLR_NEG       = "FCE4E4"
        CLR_NEG_FNT   = "C62828"
        CLR_POS_FNT   = "065F46"
        CLR_COMPRA    = "065F46"
        CLR_BORDER    = "D1D5DB"

        def fill(hex6): return PatternFill("solid", fgColor=hex6)
        def font(bold=False, color="000000", sz=10): return Font(bold=bold, color=color, size=sz, name="Segoe UI")
        def border():
            s = Side(style="thin", color=CLR_BORDER)
            return Border(left=s, right=s, top=s, bottom=s)
        def center(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
        def left():   return Alignment(horizontal="left",   vertical="center")

        n_meses = len(meses)
        # colunas: SKU(1), Venda Mês(2), depois 3 cols por mês
        COL_SKU  = 1
        COL_VEND = 2
        COL_M0   = 3  # primeira coluna do primeiro mês

        ws.column_dimensions[get_column_letter(COL_SKU)].width  = 22
        ws.column_dimensions[get_column_letter(COL_VEND)].width = 10
        for mi in range(n_meses):
            for ci in range(3):
                ws.column_dimensions[get_column_letter(COL_M0 + mi*3 + ci)].width = 11

        # ── linha 1: título ───────────────────────────────────
        ws.row_dimensions[1].height = 24
        last_col = get_column_letter(COL_M0 + n_meses*3 - 1)
        ws.merge_cells(f"A1:{last_col}1")
        c = ws["A1"]
        c.value = "Projeção de Estoque — Cenário de Simulação"
        c.font  = font(bold=True, color="FFFFFF", sz=13)
        c.fill  = fill(CLR_HDR_DARK)
        c.alignment = Alignment(horizontal="right", vertical="center")

        # ── linha 2: data gerado ──────────────────────────────
        ws.row_dimensions[2].height = 16
        ws.merge_cells(f"A2:{last_col}2")
        c = ws["A2"]
        c.value = f"Gerado em: {date.today().strftime('%d/%m/%Y')}"
        c.font  = font(color="6B7280", sz=9)
        c.fill  = fill("F9FAFB")
        c.alignment = Alignment(horizontal="right", vertical="center")

        # ── linha 3: cabeçalho de meses (merged 3 cols cada) ──
        ws.row_dimensions[3].height = 20
        ws["A3"].fill = fill(CLR_HDR_DARK); ws["A3"].font = font(bold=True, color="FFFFFF")
        ws["B3"].fill = fill(CLR_HDR_DARK); ws["B3"].font = font(bold=True, color="FFFFFF")
        for mi, (m, a, _) in enumerate(meses):
            c0 = COL_M0 + mi*3
            cl = get_column_letter(c0); cr = get_column_letter(c0+2)
            ws.merge_cells(f"{cl}3:{cr}3")
            c = ws[f"{cl}3"]
            c.value = f"{MESES_PT_CURTO[m]}/{a}"
            c.font  = font(bold=True, color="FFFFFF", sz=11)
            c.fill  = fill(CLR_HDR_MES)
            c.alignment = center()
            for ci in range(3):
                ws[f"{get_column_letter(c0+ci)}3"].fill = fill(CLR_HDR_MES)

        # ── linha 4: sub-cabeçalhos ───────────────────────────
        ws.row_dimensions[4].height = 32
        ws["A4"].value = "SKU";       ws["A4"].font = font(bold=True, color="FFFFFF"); ws["A4"].fill = fill(CLR_HDR_DARK); ws["A4"].alignment = center()
        ws["B4"].value = "Venda\nMês"; ws["B4"].font = font(bold=True, color="FFFFFF"); ws["B4"].fill = fill(CLR_HDR_DARK); ws["B4"].alignment = center()
        for mi in range(n_meses):
            c0 = COL_M0 + mi*3
            for ci, sub in enumerate(["Proj.\nVenda","Compra\n(Cen.)","Saldo"]):
                cx = ws[f"{get_column_letter(c0+ci)}4"]
                cx.value = sub; cx.font = font(bold=True, color="FFFFFF", sz=9)
                cx.fill = fill(CLR_HDR_SUB); cx.alignment = center()

        for col in range(1, COL_M0 + n_meses*3):
            for row in [3,4]:
                ws.cell(row=row, column=col).border = border()

        # ── dados ─────────────────────────────────────────────
        ws.row_dimensions[1].height = 22
        ROW = 5
        # proj_venda override por SKU do cenário
        cen_proj_ov = {}
        for p in self.cenario:
            sk = p.get("sku","").upper(); pv = p.get("proj_venda","")
            if pv and sk not in cen_proj_ov: cen_proj_ov[sk] = pv

        for r in rows:
            if r["type"] == "cat":
                ws.row_dimensions[ROW].height = 18
                ws.merge_cells(f"A{ROW}:{last_col}{ROW}")
                c = ws[f"A{ROW}"]
                c.value = r["label"]
                c.font  = font(bold=True, color="FFFFFF", sz=11)
                c.fill  = fill(CLR_CAT_BG)
                c.alignment = left()
                for col in range(1, COL_M0 + n_meses*3):
                    ws.cell(row=ROW, column=col).fill = fill(CLR_CAT_BG)
                    ws.cell(row=ROW, column=col).border = border()
                ROW += 1
                continue

            sku = r["sku"]
            er_cfg = self.local["config"].get(sku, {}).get("proj_venda","")
            er_sim = cen_proj_ov.get(sku.upper(), er_cfg)
            vend   = vendas_mes_atual(sku, self.vendas)

            ws.row_dimensions[ROW].height = 18
            c_sku = ws.cell(row=ROW, column=COL_SKU, value=sku)
            c_sku.font = font(bold=True); c_sku.alignment = left(); c_sku.border = border()

            try:   vend_val = float(vend)
            except: vend_val = 0
            c_v = ws.cell(row=ROW, column=COL_VEND, value=vend_val if vend_val else None)
            c_v.alignment = center(); c_v.border = border()
            c_v.number_format = "#,##0"

            for mi, md in enumerate(r["months"]):
                c0 = COL_M0 + mi*3
                saldo = md["cen"] if self.cenario else md["atual"]
                cheg  = md["chegada"]

                # Proj. Venda
                try:   pv_val = float(er_sim) if er_sim else None
                except: pv_val = None
                cp = ws.cell(row=ROW, column=c0, value=pv_val)
                cp.alignment = center(); cp.border = border(); cp.number_format = "#,##0"

                # Compra (Cen.)
                cc = ws.cell(row=ROW, column=c0+1, value=cheg if cheg > 0 else None)
                cc.alignment = center(); cc.border = border(); cc.number_format = "#,##0"
                if cheg > 0: cc.font = font(bold=True, color=CLR_COMPRA)

                # Saldo
                bg = CLR_NEG if saldo < 0 else (CLR_POS if saldo > 0 else "FFFFFF")
                cs = ws.cell(row=ROW, column=c0+2, value=round(saldo, 3))
                cs.fill = fill(bg); cs.alignment = center(); cs.border = border()
                cs.number_format = "#,##0.###"
                if saldo < 0: cs.font = font(bold=True, color=CLR_NEG_FNT)
                elif saldo > 0: cs.font = font(color=CLR_POS_FNT)

            ROW += 1

        ws.freeze_panes = "C5"
        wb.save(filepath)
        messagebox.showinfo("Exportado", f"Excel gerado com sucesso:\n{filepath}")
        os.startfile(filepath)

    def _simular_e_abrir_grade(self):
        """Simula o cenário e abre a grade em tela cheia diretamente."""
        if not self.cenario:
            messagebox.showinfo("Cenário vazio", "Adicione ao menos uma compra simulada antes de abrir a grade.")
            return
        self._executar_cenario()
        self._abrir_grade_tela_cheia(modo="cenario")

    def _executar_cenario(self):
        if not self.cenario:
            messagebox.showinfo("Cenário vazio", "Adicione ao menos uma compra simulada antes de executar.")
            return

        rows, meses = self._preparar_dados_grade(modo="cenario")
        self.cen_grade.render(rows, meses, modo="cenario")

        rutura_antes = rutura_depois = melhorados = 0
        pedidos_com_sim = self.local["pedidos"] + self.cenario
        cen_proj_ov = {}
        for p in self.cenario:
            sk = p.get("sku","").upper(); pv = p.get("proj_venda","")
            if pv and sk not in cen_proj_ov: cen_proj_ov[sk] = pv
        for it in ITENS:
            sku = it["sku"]
            e = self.estoque.get(sku, {})
            disp = e.get("disp", 0)
            vend = vendas_mes_atual(sku, self.vendas)
            er_cfg = self.local["config"].get(sku, {}).get("proj_venda", "")
            pm_cfg = self.local["config"].get(sku, {}).get("proj_mensal", {})
            proj_sem = estoque_projetado(sku, disp, vend, self.local["pedidos"], proj_venda=er_cfg, proj_mensal=pm_cfg)
            er_sim = cen_proj_ov.get(sku.upper(), er_cfg)
            _, pm_sim = self._get_proj_mensal_sim(sku, er_sim, pm_cfg)
            proj_com = estoque_projetado(sku, disp, vend, pedidos_com_sim, proj_venda=er_sim, proj_mensal=pm_sim)
            tinha_rutura = any(p[1] <= 0 for p in proj_sem)
            tem_rutura   = any(p[1] <= 0 for p in proj_com)
            if tinha_rutura: rutura_antes += 1
            if tem_rutura:   rutura_depois += 1
            if tinha_rutura and not tem_rutura: melhorados += 1

        icone = "✅" if melhorados > 0 else ("⚠️" if rutura_depois > 0 else "✅")
        self.cen_status.configure(text=(f"{icone}  Simulação: {rutura_antes} rupturas sem simulação → {rutura_depois} com simulação  ({melhorados} SKUs resolvidos)"))

    # ══════════════════════════════════════════════════════════
    # PAINEL IA
    # ══════════════════════════════════════════════════════════

    def _build_ia_panel(self, parent):
        self.ia_frame = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=12, border_width=1, border_color="#DDD6FE")
        top = ctk.CTkFrame(self.ia_frame, fg_color="transparent")
        top.pack(fill="x", padx=14, pady=(10,4))
        ctk.CTkLabel(top, text="🤖 Análise de IA — Recomendação de Compras", font=ctk.CTkFont("Segoe UI",13,"bold"), text_color=CLR_PURPLE).pack(side="left")
        ctk.CTkButton(top, text="✕ Fechar", command=self._ia_fechar, fg_color="#EDE9FE", text_color=CLR_PURPLE, hover_color="#DDD6FE", corner_radius=8, height=28, width=100, font=ctk.CTkFont("Segoe UI",11)).pack(side="right")
        self.ia_txt = ctk.CTkTextbox(self.ia_frame, font=ctk.CTkFont("Segoe UI",11), fg_color="#FAFAFA", wrap="word", corner_radius=8, border_width=0, height=180)
        self.ia_txt.pack(fill="both", expand=True, padx=10, pady=(0,10))

    # ══════════════════════════════════════════════════════════
    # POPULAR ABA ESTOQUE
    # ══════════════════════════════════════════════════════════

    def _popular_estoque(self, *_):
        for r in self.tree_e.get_children(): self.tree_e.delete(r)
        filtro = self.cat_var.get()
        busca  = self.busca_v.get().strip().upper()
        itens  = sorted(ITENS, key=lambda x: (CAT_ORDEM.get(x["cat"],9), _sku_num_key(x["sku"]), x["sku"]))
        rutura = critico = 0
        hoje   = date.today()
        cur_cat = None; cat_idx = {}

        # índice do mês selecionado (0-11) ou None = "Todos"
        mes_sel = self.mes_view_var.get()
        hoje_tb = date.today()
        _meses_keys = [
            f"{MESES_PT_CURTO[(hoje_tb.month-1+i)%12+1]}/{str(hoje_tb.year+(hoje_tb.month-1+i)//12)[2:]}"
            for i in range(12)
        ]
        mes_idx = _meses_keys.index(mes_sel) if mes_sel != "Todos" else None

        # chave AAAA-MM do mês selecionado (ou mês atual se "Todos")
        if mes_idx is not None:
            _sm = (hoje_tb.month - 1 + mes_idx) % 12 + 1
            _sa = hoje_tb.year + (hoje_tb.month - 1 + mes_idx) // 12
            _sel_key = f"{_sa}-{_sm:02d}"
        else:
            _sel_key = f"{hoje_tb.year}-{hoje_tb.month:02d}"

        # alternar colunas exibidas
        BASE = ("cat","sku","disp","vendas_kg","proj_venda","necessidade","chegada")
        if mes_idx is not None:
            self.tree_e.configure(displaycolumns=BASE + ("m_ini","m_ent","m_sai","m_fin"))
            # atualiza header das colunas de detalhe com o mês
            self.tree_e.heading("m_ini", text=f"  Saldo Ini. {mes_sel}")
            self.tree_e.heading("m_ent", text=f"  Entradas {mes_sel}")
            self.tree_e.heading("m_sai", text=f"  Saídas {mes_sel}")
            self.tree_e.heading("m_fin", text=f"  Saldo Final {mes_sel}")
        else:
            self.tree_e.configure(displaycolumns=BASE + ("sf1","sf2","sf3","sf4","sf5","sf6","sf7","sf8","sf9","sf10","sf11","sf12"))

        for it in itens:
            cat, sku = it["cat"], it["sku"]
            if filtro != "TODAS" and cat != filtro: continue
            if busca and busca not in sku.upper(): continue

            if cat != cur_cat:
                cur_cat = cat; cat_idx[cat] = 0
                self.tree_e.insert("","end", iid=f"_sep_{cat}", values=(f"── {cat} ──",*[""]*(len(self.tree_e["columns"])-1)), tags=("sep",))

            cat_idx[cat] = cat_idx.get(cat,0) + 1
            e    = self.estoque.get(sku, {})
            disp = e.get("disp","—")
            vend = vendas_mes_atual(sku, self.vendas)
            er   = self.local["config"].get(sku, {}).get("proj_venda", "")
            pm   = self.local["config"].get(sku, {}).get("proj_mensal", {})
            nec  = necessidade_calc(disp, er) if er else 0

            ped_sku = sorted(
                [p for p in self.local["pedidos"] if p.get("sku","").upper()==sku.upper() and not p.get("chegado") and parse_date(p.get("data","") or "") and parse_date(p["data"]) >= hoje],
                key=lambda p: parse_date(p["data"])
            )
            prox_ch = ped_sku[0]["data"] if ped_sku else ""

            proj = estoque_projetado(sku, disp, vend, self.local["pedidos"], proj_venda=er, proj_mensal=pm)
            saldos = [fmt_num(proj[i][1]) if len(proj) > i else "—" for i in range(12)]

            # valores do mês focado
            if mes_idx is not None and mes_idx < len(proj):
                disp_val = float(disp) if disp not in ("—","") else 0
                m_fin_val = proj[mes_idx][1]
                m_ent_val = proj[mes_idx][2]
                m_ini_val = proj[mes_idx-1][1] if mes_idx > 0 else disp_val
                m_sai_val = m_ini_val + m_ent_val - m_fin_val
                m_ini_fmt = fmt_num(m_ini_val)
                m_ent_fmt = fmt_num(m_ent_val)
                m_sai_fmt = fmt_num(m_sai_val)
                m_fin_fmt = fmt_num(m_fin_val)
            else:
                m_ini_fmt = m_ent_fmt = m_sai_fmt = m_fin_fmt = ""

            alt = (cat_idx[cat] % 2 == 0)
            tags = [cat+"_alt" if alt else cat]
            tem_rutura = any(v[1] <= 0 for v in proj)
            if tem_rutura: tags.append("rutura"); rutura += 1
            elif nec >= 5: tags.append("urgente"); critico += 1
            elif nec > 0:  tags.append("atencao")

            disp_fmt = fmt_num(disp) if disp != "—" else "—"
            vend_fmt = fmt_num(vend)
            nec_fmt  = fmt_num(nec) if nec > 0 else "—"
            if pm and _sel_key in pm and str(pm[_sel_key]).strip():
                er_fmt = fmt_num(pm[_sel_key]) + " ★"
            elif er:
                er_fmt = fmt_num(er) + (" ★" if pm else "")
            elif pm:
                _pm_val = pm.get(_sel_key, next(iter(pm.values()), ""))
                er_fmt = (fmt_num(_pm_val) if _pm_val else "—") + " ★"
            else:
                er_fmt = ""

            self.tree_e.insert("","end", iid=sku, values=(
                cat, sku, disp_fmt, vend_fmt, er_fmt, nec_fmt, prox_ch,
                m_ini_fmt, m_ent_fmt, m_sai_fmt, m_fin_fmt,
                *saldos
            ), tags=tuple(tags))



    # ══════════════════════════════════════════════════════════
    # POPULAR ABA PEDIDOS
    # ══════════════════════════════════════════════════════════

    def _popular_pedidos(self, *_):
        for r in self.tree_p.get_children(): self.tree_p.delete(r)
        hoje    = date.today()
        filtro  = self.mes_var.get()
        pedidos = sorted(self.local["pedidos"], key=lambda p: (parse_date(p.get("data","") or "") or date(2099,1,1)))

        meses_grp = defaultdict(list)
        for p in pedidos:
            dt = parse_date(p.get("data","") or "")
            k = f"{MESES_PT[dt.month]}/{dt.year}" if dt else "Sem data"
            if filtro == "TODOS" or filtro == k:
                meses_grp[k].append(p)

        total_meses = 0
        for mes_k, ps in meses_grp.items():
            total_qtd = sum(float(p.get("qty",0)) for p in ps)
            self.tree_p.insert("","end", iid=f"_mhdr_{mes_k}", values=(f"📅  {mes_k}","","", f"Total: {fmt_num(total_qtd)} un","","",""), tags=("mes_hdr",))
            for idx, p in enumerate(ps):
                dt   = parse_date(p.get("data","") or "")
                chegou = p.get("chegado", False)
                if chegou: status = "✅ Chegou"
                elif dt and dt < hoje: status = "⚠️ Atrasado"
                else: status = "⏳ Aguardando"
                tag = "chegou" if chegou else ("atrasado" if dt and dt<hoje else "pendente")
                sku = p.get("sku","")
                cat = SKU_CAT.get(sku.upper(), "—")
                mes = f"{MESES_PT[dt.month]}/{dt.year}" if dt else "—"
                qty_fmt = fmt_num(float(p.get('qty',0))) 
                self.tree_p.insert("","end", iid=f"_ped_{id(p)}_{idx}", values=(mes, sku, cat, f"{qty_fmt} un", p.get("data",""), p.get("obs",""), status), tags=(tag,))
            total_meses += 1

        if total_meses == 0:
            self.tree_p.insert("","end", values=("","Nenhum pedido cadastrado","","","","",""), tags=("pendente",))

    # ══════════════════════════════════════════════════════════
    # AÇÕES PEDIDOS
    # ══════════════════════════════════════════════════════════

    def _add_pedido(self):
        sku  = self.form_sku.get().strip().upper()
        qty  = self.form_qty.get().strip()
        data = self.form_data.get().strip()
        obs  = self.form_obs.get().strip()

        if not sku or not qty or not data:
            messagebox.showwarning("Campos obrigatórios","Preencha SKU, Quantidade e Data.")
            return
        if sku not in [i["sku"] for i in ITENS]:
            messagebox.showwarning("SKU inválido", f"SKU '{sku}' não encontrado na lista.")
            return
        if not parse_date(data):
            messagebox.showwarning("Data inválida","Use o formato DD/MM/AAAA.")
            return
        try: float(qty)
        except: messagebox.showwarning("Qtd inválida","Digite um número."); return

        pedido = {"sku": sku, "qty": qty, "data": data, "obs": obs, "chegado": False}
        self.local["pedidos"].append(pedido)
        salvar_local(self.local)

        self.form_sku.set("")
        self.form_qty.delete(0,"end")
        self.form_data.delete(0,"end")
        self.form_obs.delete(0,"end")
        self._popular_pedidos()
        self._popular_estoque()
        self._set_status("✅ Pedido adicionado!")

    def _del_pedido(self):
        sel = self.tree_p.selection()
        if not sel: messagebox.showinfo("Selecione","Clique em um pedido para selecionar."); return
        iid = sel[0]
        if iid.startswith("_mhdr_") or iid.startswith("_sep_"):
            messagebox.showinfo("Selecione","Selecione um pedido individual."); return
        if not messagebox.askyesno("Excluir","Excluir este pedido?"): return

        row_vals = self.tree_p.item(iid)["values"]
        sku_row  = str(row_vals[1]).strip().upper() if row_vals else ""
        data_row = str(row_vals[4]).strip() if row_vals else ""
        try:
            qty_row_f = float(str(row_vals[3]).replace(" un","").replace(".","").replace(",",".").strip())
        except (ValueError, IndexError):
            qty_row_f = None

        novos = []
        removido = False
        for p in self.local["pedidos"]:
            try:
                pqty_f = float(str(p.get('qty', 0)).replace(",", "."))
            except ValueError:
                pqty_f = None
            match = (
                not removido
                and p.get("sku","").upper() == sku_row
                and p.get("data","") == data_row
                and qty_row_f is not None
                and pqty_f is not None
                and abs(pqty_f - qty_row_f) < 0.01
            )
            if match:
                removido = True; continue
            novos.append(p)
        self.local["pedidos"] = novos
        salvar_local(self.local)
        self._popular_pedidos()
        self._popular_estoque()

    def _edit_pedido(self):
        sel = self.tree_p.selection()
        if not sel:
            messagebox.showinfo("Selecione", "Clique em um pedido para selecionar antes de editar.")
            return
        iid = sel[0]
        if iid.startswith("_mhdr_") or iid.startswith("_sep_"):
            messagebox.showinfo("Selecione", "Selecione um pedido individual.")
            return

        row_vals = self.tree_p.item(iid)["values"]
        sku_row  = str(row_vals[1]).strip().upper() if row_vals else ""
        data_row = str(row_vals[4]).strip() if row_vals else ""
        try:
            qty_row_f = float(str(row_vals[3]).replace(" un","").replace(".","").replace(",",".").strip())
        except (ValueError, IndexError):
            qty_row_f = None

        # acha o pedido correspondente
        pedido_ref = None
        for p in self.local["pedidos"]:
            try:
                pqty_f = float(str(p.get('qty', 0)).replace(",", "."))
            except ValueError:
                pqty_f = None
            if (p.get("sku","").upper() == sku_row
                    and p.get("data","") == data_row
                    and qty_row_f is not None and pqty_f is not None
                    and abs(pqty_f - qty_row_f) < 0.01):
                pedido_ref = p
                break

        if pedido_ref is None:
            messagebox.showerror("Erro", "Pedido não encontrado. Tente atualizar a lista.")
            return

        dlg = ctk.CTkToplevel(self)
        dlg.title("Editar Pedido")
        dlg.geometry("460x300")
        dlg.resizable(False, False)
        dlg.grab_set()

        ctk.CTkLabel(dlg, text="✏️ Editar Pedido", font=ctk.CTkFont("Segoe UI",14,"bold"), text_color=CLR_ACCENT).pack(pady=(18,10))

        grid = ctk.CTkFrame(dlg, fg_color="transparent")
        grid.pack(padx=24, fill="x")

        ctk.CTkLabel(grid, text="SKU", font=ctk.CTkFont("Segoe UI",11), text_color=CLR_MUTED).grid(row=0,column=0,sticky="w",pady=4)
        lbl_sku = ctk.CTkLabel(grid, text=sku_row, font=ctk.CTkFont("Segoe UI",12,"bold"), text_color=CLR_TEXT)
        lbl_sku.grid(row=0,column=1,sticky="w",padx=10)

        ctk.CTkLabel(grid, text="Quantidade *", font=ctk.CTkFont("Segoe UI",11), text_color=CLR_MUTED).grid(row=1,column=0,sticky="w",pady=4)
        e_qty = ctk.CTkEntry(grid, placeholder_text="Ex: 100", width=150, height=36, fg_color="#FFFFFF", font=ctk.CTkFont("Segoe UI",12))
        e_qty.insert(0, str(pedido_ref.get("qty","")))
        e_qty.grid(row=1,column=1,sticky="w",padx=10,pady=4)

        ctk.CTkLabel(grid, text="Data Chegada *", font=ctk.CTkFont("Segoe UI",11), text_color=CLR_MUTED).grid(row=2,column=0,sticky="w",pady=4)
        e_data = ctk.CTkEntry(grid, placeholder_text="DD/MM/AAAA", width=150, height=36, fg_color="#FFFFFF", font=ctk.CTkFont("Segoe UI",12))
        e_data.insert(0, pedido_ref.get("data",""))
        e_data.grid(row=2,column=1,sticky="w",padx=10,pady=4)

        ctk.CTkLabel(grid, text="Observação", font=ctk.CTkFont("Segoe UI",11), text_color=CLR_MUTED).grid(row=3,column=0,sticky="w",pady=4)
        e_obs = ctk.CTkEntry(grid, placeholder_text="Opcional", width=220, height=36, fg_color="#FFFFFF", font=ctk.CTkFont("Segoe UI",12))
        e_obs.insert(0, pedido_ref.get("obs",""))
        e_obs.grid(row=3,column=1,sticky="w",padx=10,pady=4)

        def _salvar():
            qty  = e_qty.get().strip()
            data = e_data.get().strip()
            obs  = e_obs.get().strip()
            if not qty or not data:
                messagebox.showwarning("Campos obrigatórios", "Quantidade e Data são obrigatórios.", parent=dlg)
                return
            if not parse_date(data):
                messagebox.showwarning("Data inválida", "Use o formato DD/MM/AAAA.", parent=dlg)
                return
            try:
                qty_f = float(qty.replace(",","."))
                if qty_f <= 0: raise ValueError
            except:
                messagebox.showwarning("Quantidade inválida", "Digite um número positivo.", parent=dlg)
                return
            pedido_ref["qty"]  = str(qty_f)
            pedido_ref["data"] = data
            pedido_ref["obs"]  = obs
            salvar_local(self.local)
            self._popular_pedidos()
            self._popular_estoque()
            dlg.destroy()

        ctk.CTkButton(dlg, text="💾 Salvar", command=_salvar, fg_color=CLR_GREEN, hover_color="#059669", corner_radius=8, height=38, font=ctk.CTkFont("Segoe UI",12,"bold")).pack(pady=18)

    def _toggle_chegado(self, ev):
        region = self.tree_p.identify("region", ev.x, ev.y)
        if region != "cell": return
        iid = self.tree_p.identify_row(ev.y)
        if not iid or iid.startswith("_mhdr_") or iid.startswith("_sep_"): return
        row_vals = self.tree_p.item(iid)["values"]
        sku_row  = str(row_vals[1]).strip().upper()
        data_row = str(row_vals[4]).strip()
        qty_row  = str(row_vals[3]).replace(" un","").replace(".","").strip()
        for p in self.local["pedidos"]:
            pqty = fmt_num(float(p.get('qty',0)))
            if (p.get("sku","").upper() == sku_row and p.get("data","") == data_row and pqty == qty_row):
                p["chegado"] = not p.get("chegado", False)
                break
        salvar_local(self.local)
        self._popular_pedidos()
        self._popular_estoque()

    # ══════════════════════════════════════════════════════════
    # PREVISÃO SKU
    # ══════════════════════════════════════════════════════════

    def _analisar_sku(self):
        for w in self.prev_frame.winfo_children(): w.destroy()
        sku = self.prev_sku.get().strip().upper()
        if not sku: return

        e    = self.estoque.get(sku, {})
        disp = e.get("disp", 0)
        vend = vendas_mes_atual(sku, self.vendas) 
        er   = self.local["config"].get(sku, {}).get("proj_venda","")

        cat = SKU_CAT.get(sku, "—")
        cor_cat = CAT_COLORS.get(cat, {}).get("badge", "#6B7280")

        info_card = ctk.CTkFrame(self.prev_frame, fg_color=BG_CARD, corner_radius=12, border_width=1, border_color=CLR_BORDER)
        info_card.pack(fill="x", pady=(0,12))
        
        info_inner = ctk.CTkFrame(info_card, fg_color="transparent")
        info_inner.pack(padx=16, pady=12, fill="x")
        
        ctk.CTkLabel(info_inner, text=f"📈 Projeção de Saldo Final — {sku}", font=ctk.CTkFont("Segoe UI",15,"bold"), text_color=CLR_TEXT).pack(anchor="w")
        
        info_txt = f"Categoria: {cat} | Disponível: {fmt_num(disp)} | Vendas/mês: {fmt_num(vend)}"
        if er: info_txt += f" | Proj. Venda: {fmt_num(er)}"
        ctk.CTkLabel(info_inner, text=info_txt, font=ctk.CTkFont("Segoe UI",11), text_color=CLR_MUTED).pack(anchor="w", pady=(6,0))

        proj_card = ProjecaoVendaInput(self.prev_frame, sku, er or "", callback=self._atualizar_proj_venda)
        proj_card.pack(fill="x", pady=(0,12))

        pm   = self.local["config"].get(sku, {}).get("proj_mensal", {})
        proj = estoque_projetado(sku, disp, vend, self.local["pedidos"], meses=12, proj_venda=er, proj_mensal=pm)

        for mes_str, saldo, chegadas, m, a in proj:
            row = ctk.CTkFrame(self.prev_frame, fg_color=BG_CARD, corner_radius=12, border_width=1, border_color=CLR_BORDER, height=60)
            row.pack(fill="x", pady=4); row.pack_propagate(False)

            if saldo <= 0: cor = CLR_RED; icone = "🚨"; txt_status = "RUPTURA"
            elif er and saldo < float(er)*0.3: cor = CLR_AMBER; icone = "⚠️"; txt_status = "CRÍTICO"
            elif er and saldo < float(er)*0.7: cor = "#F59E0B"; icone = "🟡"; txt_status = "BAIXO"
            else: cor = CLR_GREEN; icone = "✅"; txt_status = "OK"

            bar = ctk.CTkFrame(row, fg_color=cor, corner_radius=0, width=5)
            bar.pack(side="left", fill="y", padx=(0,0), pady=6)

            ctk.CTkLabel(row, text=icone, font=ctk.CTkFont("Segoe UI",18)).pack(side="left",padx=(14,10))
            ctk.CTkLabel(row, text=f"Saldo Final {mes_str}", font=ctk.CTkFont("Segoe UI",12,"bold"), width=200, text_color=CLR_TEXT).pack(side="left")
            ctk.CTkLabel(row, text=f"{fmt_num(saldo)} un", font=ctk.CTkFont("Segoe UI",13,"bold"), text_color=CLR_TEXT).pack(side="left",padx=20)
            if chegadas > 0:
                ctk.CTkLabel(row, text=f"📦 +{fmt_num(chegadas)} chegando", font=ctk.CTkFont("Segoe UI",11), text_color=CLR_GREEN).pack(side="left",padx=12)
            ctk.CTkLabel(row, text=f"● {txt_status}", font=ctk.CTkFont("Segoe UI",11,"bold"), text_color=cor).pack(side="right",padx=16)

        ped_sku = [p for p in self.local["pedidos"] if p.get("sku","").upper()==sku.upper() and not p.get("chegado")]
        if ped_sku:
            ctk.CTkLabel(self.prev_frame, text="📋 Pedidos Abertos", font=ctk.CTkFont("Segoe UI",13,"bold"), text_color=CLR_ACCENT).pack(anchor="w", pady=(16,8))
            for p in ped_sku:
                r2 = ctk.CTkFrame(self.prev_frame, fg_color="#F0F4FF", corner_radius=10, border_width=1, border_color=CLR_BORDER, height=42)
                r2.pack(fill="x", pady=3); r2.pack_propagate(False)
                qty_fmt = fmt_num(float(p.get('qty',0))) 
                ctk.CTkLabel(r2, text=f"  📦 {qty_fmt} un  |  Chegada: {p.get('data')}  |  {p.get('obs','')}", font=ctk.CTkFont("Segoe UI",11), text_color=CLR_TEXT).pack(side="left",padx=12)

    def _atualizar_proj_venda(self, sku, valor):
        try:
            v = float(valor) if valor else ""
            self.local["config"].setdefault(sku, {})["proj_venda"] = str(v) if v else ""
            salvar_local(self.local)
            self._analisar_sku()
        except: pass

    # ══════════════════════════════════════════════════════════
    # EDIÇÃO INLINE
    # ══════════════════════════════════════════════════════════

    def _edit_cell(self, ev, tree, col_edit_map):
        region = tree.identify("region", ev.x, ev.y)
        if region != "cell": return
        col = tree.identify_column(ev.x)
        row = tree.identify_row(ev.y)
        if not row or row.startswith("_sep_"): return
        cidx = int(col.replace("#","")) - 1

        if cidx == 4 and tree is self.tree_e:
            self._abrir_proj_mensal(row)
            return

        if cidx not in col_edit_map: return
        key  = col_edit_map[cidx]
        cur  = str(tree.item(row)["values"][cidx])
        cur_clean = cur.replace(".","").replace(",",".") if cur not in ("—","") else ""
        bbox = tree.bbox(row, col)
        if not bbox: return
        var  = tk.StringVar(value=cur_clean)
        e = ctk.CTkEntry(tree, textvariable=var, width=bbox[2], placeholder_text="DD/MM/AAAA", font=ctk.CTkFont("Segoe UI",11), fg_color="#EFF6FF", border_color=CLR_ACCENT, border_width=2, corner_radius=6)
        e.place(x=bbox[0], y=bbox[1], height=bbox[3])
        e.focus_set()
        def commit(_=None):
            val = var.get().strip()
            self.local["config"].setdefault(row, {})[key] = val
            e.destroy()
            self._popular_estoque()
            e.bind("<Return>", commit)
        e.bind("<FocusOut>", commit)
        e.bind("<Escape>", lambda _: e.destroy())

    def _abrir_proj_mensal(self, sku):
        dlg = ctk.CTkToplevel(self)
        dlg.title(f"Projeção de Venda — {sku}")
        dlg.geometry("600x500")
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.lift()

        hoje = date.today()
        cfg  = self.local["config"].get(sku, {})
        cat  = SKU_CAT.get(sku, "")
        cor  = CAT_COLORS.get(cat, {}).get("badge", CLR_ACCENT)

        hdr = ctk.CTkFrame(dlg, fg_color=CLR_TABLE_HEAD, corner_radius=0, height=54)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text=f"📊  Projeção de Venda — {sku}", font=ctk.CTkFont("Segoe UI", 14, "bold"), text_color="#FFFFFF").pack(side="left", padx=20, pady=14)
        if cat: ctk.CTkLabel(hdr, text=f"  {cat}  ", fg_color=cor, corner_radius=6, text_color="#FFFFFF", font=ctk.CTkFont("Segoe UI", 11, "bold")).pack(side="left", padx=4)

        body = ctk.CTkFrame(dlg, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=14)

        glob_card = ctk.CTkFrame(body, fg_color=BG_CARD, corner_radius=10, border_width=1, border_color=CLR_BORDER)
        glob_card.pack(fill="x", pady=(0, 12))
        gi = ctk.CTkFrame(glob_card, fg_color="transparent")
        gi.pack(fill="x", padx=14, pady=10)
        ctk.CTkLabel(gi, text="🌐 Global — padrão para meses sem valor específico", font=ctk.CTkFont("Segoe UI", 11, "bold"), text_color=CLR_TEXT).pack(anchor="w")
        gf = ctk.CTkFrame(gi, fg_color="transparent")
        gf.pack(fill="x", pady=(6, 0))
        global_var = tk.StringVar(value=cfg.get("proj_venda", ""))
        ctk.CTkEntry(gf, textvariable=global_var, width=140, height=38, placeholder_text="kg/mês", corner_radius=8, fg_color="#F9FAFB", border_color=CLR_ACCENT, font=ctk.CTkFont("Segoe UI", 13, "bold")).pack(side="left")
        ctk.CTkLabel(gf, text=" kg/mês", font=ctk.CTkFont("Segoe UI", 11), text_color=CLR_MUTED).pack(side="left", padx=6)

        ctk.CTkLabel(body, text="📅 Por mês  (deixe vazio para usar o Global acima)", font=ctk.CTkFont("Segoe UI", 11, "bold"), text_color=CLR_TEXT).pack(anchor="w", pady=(0, 6))

        grid_card = ctk.CTkFrame(body, fg_color=BG_CARD, corner_radius=10, border_width=1, border_color=CLR_BORDER)
        grid_card.pack(fill="x")

        proj_mensal = cfg.get("proj_mensal", {})
        month_vars  = {}

        for i in range(12):
            _m  = (hoje.month - 1 + i) % 12 + 1
            _a  = hoje.year + (hoje.month - 1 + i) // 12
            key = f"{_a}-{_m:02d}"
            lbl = f"{MESES_PT_CURTO[_m]}/{str(_a)[2:]}"
            val = proj_mensal.get(key, "")

            cell = ctk.CTkFrame(grid_card, fg_color="transparent")
            cell.grid(row=i // 3, column=i % 3, padx=10, pady=7, sticky="w")

            ctk.CTkLabel(cell, text=lbl, font=ctk.CTkFont("Segoe UI", 10, "bold"), text_color=CLR_MUTED, width=50).pack(side="left")
            var = tk.StringVar(value=str(val) if val != "" else "")
            ctk.CTkEntry(cell, textvariable=var, width=90, height=34, placeholder_text="—", corner_radius=6, fg_color="#F9FAFB", font=ctk.CTkFont("Segoe UI", 11)).pack(side="left", padx=(4, 0))
            month_vars[key] = var

        btn_row = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=(10, 16))

        def salvar():
            cs = self.local["config"].setdefault(sku, {})
            gv = global_var.get().strip()
            try:   cs["proj_venda"] = str(float(gv)) if gv else ""
            except: cs["proj_venda"] = ""
            pm = {}
            for k, v in month_vars.items():
                sv = v.get().strip()
                if sv:
                    try: pm[k] = float(sv)
                    except: pass
            cs["proj_mensal"] = pm
            salvar_local(self.local)
            dlg.destroy()
            self._popular_estoque()
    
        def limpar():
            global_var.set("")
            for v in month_vars.values(): v.set("")

        ctk.CTkButton(btn_row, text="💾 Salvar", command=salvar, fg_color=CLR_GREEN, hover_color="#059669", corner_radius=8, height=40, width=130, font=ctk.CTkFont("Segoe UI", 12, "bold")).pack(side="left", padx=(0,8))
        ctk.CTkButton(btn_row, text="🗑 Limpar tudo", command=limpar, fg_color="#FEE2E2", text_color=CLR_RED, hover_color="#FECACA", corner_radius=8, height=40, width=130, font=ctk.CTkFont("Segoe UI", 11)).pack(side="left", padx=(0,8))
        ctk.CTkButton(btn_row, text="✕ Cancelar", command=dlg.destroy, fg_color="#F3F4F6", text_color=CLR_TEXT, hover_color="#E5E7EB", corner_radius=8, height=40, width=100, font=ctk.CTkFont("Segoe UI", 11)).pack(side="left")

    # ══════════════════════════════════════════════════════════
    # ATUALIZAR DADOS E IA
    # ══════════════════════════════════════════════════════════

    def _atualizar(self):
        if self._loading: return
        self._loading = True
        self.prog.pack(fill="x")
        self.prog.start()
        self._set_status("⏳ Buscando dados…")
        threading.Thread(target=self._fetch_all, daemon=True).start()

    def _fetch_all(self):
        n = len(ITENS)
        for i, it in enumerate(ITENS):
            self.estoque[it["sku"]] = consultar_sku(it["sku"])
            if i % 5 == 0: self._set_status(f"Estoque {i+1}/{n}…")
        self._set_status("Buscando vendas…")
        self.vendas = buscar_vendas()
        self.after(0, self._pos_fetch)

    def _pos_fetch(self):
        self._loading = False
        self.prog.stop(); self.prog.pack_forget()
        self._set_status("✅ Atualizado")
        self.lbl_sync.configure(text=f"Sync: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        self._popular_estoque()
        self._popular_alertas()

    def _salvar(self):
        salvar_local(self.local)
        self._set_status("💾 Salvo!")
        messagebox.showinfo("Salvo","Dados salvos em compras_data.json ✅")

    def _ia_recomendar(self):
        if not self._ia_open:
            self.ia_frame.pack(fill="x", pady=(12,0))
            self._ia_open = True
        self.btn_ia.configure(state="disabled", text="⏳ Analisando…")
        self.ia_txt.delete("0.0","end")
        self.ia_txt.insert("0.0","🔍 Analisando estoque, vendas e pedidos…\n")
        threading.Thread(target=self._fetch_ia, daemon=True).start()

    def _ia_fechar(self):
        self.ia_frame.pack_forget(); self._ia_open = False

    def _fetch_ia(self):
        linhas = []
        for it in sorted(ITENS, key=lambda x:(CAT_ORDEM.get(x["cat"],9),x["sku"])):
            sku, cat = it["sku"], it["cat"]
            e    = self.estoque.get(sku, {})
            disp = e.get("disp",0); res = e.get("reserva",0)
            vend = vendas_mes_atual(sku, self.vendas) 
            er   = self.local["config"].get(sku, {}).get("proj_venda", "")
            _pm  = self.local["config"].get(sku, {}).get("proj_mensal", {})
            nec  = necessidade_calc(disp, er) if er else "?"
            proj = estoque_projetado(sku, disp, vend, self.local["pedidos"], proj_venda=er, proj_mensal=_pm)
            projs = " | ".join(f"SaldoFinal_{m}:{s:.0f}" for m,s,_,_,_ in proj)
            ped   = sum(float(p.get("qty",0)) for p in self.local["pedidos"] if p.get("sku","").upper()==sku.upper() and not p.get("chegado"))
            linhas.append(f"{cat}|{sku}|Disp:{disp}|Res:{res}|Vendas:{vend}kg|ProjVenda:{er or 'N/D'}|Nec:{nec}|PedAberto:{ped:.0f}|SaldoFinal:[{projs}]")
        prompt = f"""Você é um consultor sênior de supply chain especializado em {SEGMENTO}.
Analise e retorne em 4 seções:
🚨 RUPTURA PREVISTA — SKUs com saldo final ≤ 0 em algum mês
⚠️ URGENTE — necessidade > 0 sem pedido aberto
📋 ATENÇÃO — pedido aberto mas estoque ainda crítico
✅ ESTRATÉGIA GERAL — até 3 linhas de recomendação
Dados (Categoria|SKU|Disponível|Reserva|VendasKG|ProjVenda|Necessidade|PedidoAberto|SaldoFinal3meses):
{chr(10).join(linhas)}
Seja direto. Responda em português."""
        try:
            msg = Anthropic().messages.create(model="claude-sonnet-4-20250514", max_tokens=2000, messages=[{"role":"user","content":prompt}])
            resp = msg.content[0].text
        except Exception as ex:
            resp = f"❌ Erro: {ex}"
        self.after(0, lambda: self._exibir_ia(resp))

    def _exibir_ia(self, txt):
        self.ia_txt.delete("0.0","end")
        self.ia_txt.insert("0.0", txt)
        self.btn_ia.configure(state="normal", text="🤖 Recomendar (IA)")

    def _set_status(self, msg):
        self.after(0, lambda: self.lbl_status.configure(text=msg))

if __name__ == "__main__":
    app = App()
    app.mainloop()