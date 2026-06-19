import pandas as pd
import sys
import os
import re
import json
import shutil
import tempfile
from datetime import datetime

# ── Paths ─────────────────────────────────────────────────────────────────────
SAIDA_DIR   = os.path.dirname(os.path.abspath(__file__))
DADOS_DIR   = os.path.join(SAIDA_DIR, "Dados")
OUTPUT_DIR  = os.path.join(os.path.dirname(SAIDA_DIR), "output")
HOME_HTML   = os.path.join(OUTPUT_DIR, "index.html")

sys.path.insert(0, SAIDA_DIR)
from config import WMS_COLUNAS, ERP_COLUNAS, CHAVE_CONCILIACAO, ERP_SKU_PADRAO, ERP_FILTRO_TIPO_NF, SUBCLIENTES

GRUPO_CORES = {
    "LIV UP":    "#ec4899",  # rosa
    "DRUMATTOS": "#b91c1c",  # vermelho escuro
}
_COR_PADRAO = "#6366f1"
def _cor_grupo(g): return GRUPO_CORES.get(g, _COR_PADRAO)


# ── Carregamento ──────────────────────────────────────────────────────────────

def _read_excel_safe(caminho, **kwargs):
    """Copia o arquivo para temp antes de ler — evita PermissionError se aberto no Excel."""
    try:
        return pd.read_excel(caminho, **kwargs)
    except PermissionError:
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp.close()
        shutil.copy2(caminho, tmp.name)
        try:
            return pd.read_excel(tmp.name, **kwargs)
        finally:
            os.unlink(tmp.name)


def _strip_cols(df):
    df.columns = [c.strip() for c in df.columns]
    return df


def carregar_wms(caminho):
    df = _read_excel_safe(caminho, dtype=str)
    _strip_cols(df)
    renomear = {v: k for k, v in WMS_COLUNAS.items() if v in df.columns}
    df = df.rename(columns=renomear)
    faltando = [k for k in ["pedido", "sku", "quantidade", "data"] if k not in df.columns]
    if faltando:
        raise ValueError(f"WMS saida: colunas faltando {faltando}. Disponíveis: {list(df.columns)}")
    df["quantidade"] = pd.to_numeric(df["quantidade"].str.replace(",", "."), errors="coerce").fillna(0)
    if "peso" in df.columns:
        df["peso"] = pd.to_numeric(df["peso"].str.replace(",", "."), errors="coerce").fillna(0)
    df["data"] = pd.to_datetime(df["data"], errors="coerce")
    df["pedido"] = df["pedido"].str.strip()
    df["sku"]    = df["sku"].str.strip()
    return df


def carregar_erp(caminho):
    df = _read_excel_safe(caminho, dtype=str)
    _strip_cols(df)
    renomear = {v: k for k, v in ERP_COLUNAS.items() if v in df.columns}
    df = df.rename(columns=renomear)
    faltando = [k for k in ["pedido", "sku", "quantidade", "data"] if k not in df.columns]
    if faltando:
        raise ValueError(f"ERP saida: colunas faltando {faltando}. Disponíveis: {list(df.columns)}")
    df["quantidade"] = pd.to_numeric(df["quantidade"].str.replace(",", "."), errors="coerce").fillna(0)
    if "vlr_unit" in df.columns:
        df["vlr_unit"] = pd.to_numeric(df["vlr_unit"].str.replace(",", "."), errors="coerce").fillna(0)
    df["data"]   = pd.to_datetime(df["data"], errors="coerce")
    df["pedido"] = df["pedido"].str.strip()
    df["sku"]    = df["sku"].str.strip()
    return df


# ── Auto-detecção de pares ────────────────────────────────────────────────────

def detectar_pares(pasta):
    arquivos = sorted(os.listdir(pasta))
    pares    = []

    for f in arquivos:
        if not re.match(r'^SAIDA WMS', f, re.IGNORECASE):
            continue

        # Padrão com grupo e data: "SAIDA WMS LIV UP 15.06.xlsx"
        m = re.match(r'^SAIDA WMS (.+?)\s+(\d{1,2}[.\-]\d{2})\.xlsx$', f, re.IGNORECASE)
        if m:
            grupo    = m.group(1).strip()
            data_str = m.group(2).replace('-', '.')
            sufixo   = f[10:]  # remove "SAIDA WMS "
            erp_path = os.path.join(pasta, "SAIDA ERP " + sufixo)
            if os.path.exists(erp_path):
                pares.append((grupo, data_str, os.path.join(pasta, f), erp_path))
            continue

        # Padrão com grupo, sem data: "SAIDA WMS LIV UP.xlsx"
        m2 = re.match(r'^SAIDA WMS (.+?)\.xlsx$', f, re.IGNORECASE)
        if m2:
            grupo    = m2.group(1).strip()
            sufixo   = f[10:]
            erp_path = os.path.join(pasta, "SAIDA ERP " + sufixo)
            if os.path.exists(erp_path):
                pares.append((grupo, "", os.path.join(pasta, f), erp_path))
            continue

        # Padrão genérico: "SAIDA WMS.xlsx" — procura qualquer "SAIDA ERP *.xlsx"
        if re.match(r'^SAIDA WMS\.xlsx$', f, re.IGNORECASE):
            erp_candidatos = [
                x for x in arquivos
                if re.match(r'^SAIDA ERP', x, re.IGNORECASE) and x.endswith('.xlsx')
            ]
            for erp_f in erp_candidatos:
                # Extrai nome do grupo do ERP: "SAIDA ERP LIV UP 15.06.xlsx" → "LIV UP"
                mg = re.match(r'^SAIDA ERP (.+?)(?:\s+\d{1,2}[.\-]\d{2})?\.xlsx$', erp_f, re.IGNORECASE)
                grupo = mg.group(1).strip() if mg else "SAIDA"
                pares.append((grupo, "", os.path.join(pasta, f), os.path.join(pasta, erp_f)))

    return pares


# ── Conciliação ───────────────────────────────────────────────────────────────

def conciliar(df_wms, df_erp):
    # WMS: agrega múltiplos picks por pedido+sku
    wms_agg = df_wms.groupby(CHAVE_CONCILIACAO, as_index=False).agg(
        quantidade_wms=("quantidade", "sum"),
        unidade_wms=("unidade", "first"),
        data_wms=("data", "first"),
        peso_wms=("peso", "sum") if "peso" in df_wms.columns else ("quantidade", "first"),
    )

    # ERP: agrega por pedido+sku, consolida NFs
    def _join_nfs(s):
        vals = sorted(s.dropna().astype(str).str.strip().unique())
        return " / ".join(vals) if vals else ""

    erp_cols = dict(
        quantidade_erp=("quantidade", "sum"),
        unidade_erp=("unidade", "first"),
        data_erp=("data", "first"),
    )
    if "numero_nf" in df_erp.columns:
        erp_cols["numero_nf"] = ("numero_nf", _join_nfs)
    if "vlr_unit" in df_erp.columns:
        erp_cols["vlr_unit"] = ("vlr_unit", "mean")

    erp_agg = df_erp.groupby(CHAVE_CONCILIACAO, as_index=False).agg(**erp_cols)

    merged = pd.merge(wms_agg, erp_agg, on=CHAVE_CONCILIACAO,
                      how="outer", indicator=True)

    so_wms = merged[merged["_merge"] == "left_only"].copy()
    so_erp = merged[merged["_merge"] == "right_only"].copy()
    ambos  = merged[merged["_merge"] == "both"].copy()

    ambos["diff_qtd"] = (ambos["quantidade_wms"] - ambos["quantidade_erp"]).round(3)
    if "vlr_unit" in ambos.columns:
        ambos["valor_div"] = (ambos["diff_qtd"] * ambos["vlr_unit"]).round(2)

    ok      = ambos[ambos["diff_qtd"] == 0].copy()
    div_qtd = ambos[ambos["diff_qtd"] != 0].copy()

    return {"ok": ok, "so_wms": so_wms, "so_erp": so_erp, "div_qtd": div_qtd}


# ── Helpers de formato ────────────────────────────────────────────────────────

def _br_fmt(x, dec=3, strip_zeros=True):
    if not pd.notna(x):
        return ""
    s = f"{abs(x):,.{dec}f}"
    if strip_zeros:
        intg, frac = s.split(".")
        frac = frac.rstrip("0")
        s = intg if not frac else f"{intg}.{frac}"
    br = s.replace(",", "\x00").replace(".", ",").replace("\x00", ".")
    return ("-" + br) if x < 0 else br


def _slug(texto):
    return re.sub(r"[^a-z0-9]", "_", str(texto).lower())


# ── Geração HTML ──────────────────────────────────────────────────────────────

def _df_para_html(df, cols, rotulos):
    existentes = [c for c in cols if c in df.columns]
    if df.empty or not existentes:
        return '<p class="vazio">Nenhum registro nesta categoria.</p>'
    out = df[existentes].copy()
    for col in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[col]):
            out[col] = out[col].dt.strftime("%d/%m/%Y").fillna("")
        elif pd.api.types.is_float_dtype(out[col]):
            out[col] = out[col].apply(lambda x: _br_fmt(x, 3) if pd.notna(x) else "")
    out.rename(columns=rotulos, inplace=True)
    return out.to_html(index=False, classes="tabela", border=0, na_rep="")


def _tabela_div_html(df, cols, sg, data_str, grupo):
    rotulos = {
        "pedido": "Pedido", "numero_nf": "NF", "sku": "SKU",
        "unidade_erp": "Un.", "unidade_wms": "Un.",
        "quantidade_wms": "Qtd WMS", "quantidade_erp": "Qtd ERP",
        "diff_qtd": "Diferenca", "valor_div": "Valor Div. (R$)",
        "data_erp": "Data", "data_wms": "Data",
    }
    existentes = [c for c in cols if c in df.columns]
    if df.empty or not existentes:
        return '<p class="vazio">Nenhum registro nesta categoria.</p>'

    out = df[existentes].copy()
    raw_diff = out["diff_qtd"].to_dict() if "diff_qtd" in out.columns else {}
    raw_vdiv = out["valor_div"].to_dict() if "valor_div" in out.columns else {}

    for col in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[col]):
            out[col] = out[col].dt.strftime("%d/%m/%Y").fillna("")
        elif pd.api.types.is_float_dtype(out[col]):
            if col == "valor_div":
                out[col] = out[col].apply(lambda x: _br_fmt(x, 2, strip_zeros=False) if pd.notna(x) else "")
            else:
                out[col] = out[col].apply(lambda x: _br_fmt(x, 3) if pd.notna(x) else "")

    headers = "".join(f"<th>{rotulos.get(c, c)}</th>" for c in existentes) + "<th></th>"
    rows_html = ""
    for row_idx, row in out.iterrows():
        pedido = str(row.get("pedido", "")).strip()
        sku    = str(row.get("sku", "")).strip()
        key    = re.sub(r"[^a-zA-Z0-9]", "_", f"{pedido}__{sku}")
        qtd_wms = str(row.get("quantidade_wms", ""))
        qtd_erp = str(row.get("quantidade_erp", ""))

        rd = raw_diff.get(row_idx, 0.0)
        rd = 0.0 if (isinstance(rd, float) and rd != rd) else float(rd or 0)
        diff_str   = str(row.get("diff_qtd", ""))
        diff_modal = f"+{diff_str}" if rd > 0 else diff_str

        pedido_js = pedido.replace("'", "\\'")
        sku_js    = sku.replace("'", "\\'")

        cells = ""
        for c in existentes:
            v   = row[c]
            val = "" if (isinstance(v, float) and v != v) else ("" if v is None else str(v))
            if c == "diff_qtd":
                raw = raw_diff.get(row_idx, 0.0)
                raw = 0.0 if (isinstance(raw, float) and raw != raw) else float(raw or 0)
                display = f"+{val}" if raw > 0 else val
                css     = "td-diff td-diff-pos" if raw > 0 else "td-diff td-diff-neg"
                cells  += f'<td class="{css}" data-diff-original="{raw}">{display}</td>'
            elif c == "valor_div":
                raw = raw_vdiv.get(row_idx, 0.0)
                raw = 0.0 if (isinstance(raw, float) and raw != raw) else float(raw or 0)
                css   = "td-valordiv td-valordiv-pos" if raw > 0 else "td-valordiv td-valordiv-neg"
                cells += f'<td class="{css}">{val}</td>'
            else:
                cells += f"<td>{val}</td>"

        btn_cell = (
            f'<span id="aj-{key}" class="badge-ajuste" style="display:none;"></span>'
        )
        rows_html += (
            f'<tr data-ajuste-key="{key}" data-sg="{sg}" '
            f'data-data="{data_str}" data-grupo="{grupo}">'
            f'{cells}<td class="td-ajuste">{btn_cell}</td></tr>\n'
        )

    return f"""<div class="table-wrap">
<table class="tabela" border="0"><thead><tr>{headers}</tr></thead>
<tbody>{rows_html}</tbody></table></div>"""


# ── Seção de subgrupo (subclient dentro de NEXXA) ────────────────────────────

def _secao_subgrupo(nome, r, sg_parent, data_str, grupo, is_first=False):
    ok      = r["ok"]
    so_erp  = r["so_erp"]
    div_qtd = r["div_qtd"]

    sub_sg    = sg_parent + "_sub_" + _slug(nome)
    total_erp = len(ok) + len(so_erp) + len(div_qtd)
    pct_ok    = round(len(ok) / total_erp * 100, 1) if total_erp else 0

    def pct(n): return round(n / total_erp * 100, 1) if total_erp else 0

    rotulos_ok  = {"pedido": "Pedido", "numero_nf": "NF", "sku": "SKU",
                   "unidade_erp": "Un.", "quantidade_wms": "Qtd WMS",
                   "quantidade_erp": "Qtd ERP", "data_erp": "Data"}
    rotulos_erp = {"pedido": "Pedido", "numero_nf": "NF", "sku": "SKU",
                   "unidade_erp": "Un.", "quantidade_erp": "Qtd ERP", "data_erp": "Data"}

    cols_ok      = ["pedido", "numero_nf", "sku", "unidade_erp", "quantidade_wms", "quantidade_erp", "data_erp"]
    cols_so_erp  = ["pedido", "numero_nf", "sku", "unidade_erp", "quantidade_erp", "data_erp"]
    cols_div_qtd = ["pedido", "numero_nf", "sku", "unidade_erp", "quantidade_wms", "quantidade_erp", "diff_qtd", "valor_div", "data_erp"]

    tab_ok      = _df_para_html(ok,      cols_ok,     rotulos_ok)
    tab_so_erp  = _df_para_html(so_erp,  cols_so_erp, rotulos_erp)
    tab_div_qtd = _tabela_div_html(div_qtd, cols_div_qtd, sub_sg, data_str, f"{grupo} / {nome}")

    display = '' if is_first else 'display:none;'
    return f"""
<div id="subgrupo-{sub_sg}" class="subgrupo-section" style="{display}">
  <div class="painel painel-sub">
    <div class="donut-wrap">
      <div class="donut-container">
        <canvas id="donut-{sub_sg}" width="180" height="180"></canvas>
        <div class="donut-centro">
          <div class="pct">{pct_ok}%</div>
          <div class="sub">Conciliado</div>
        </div>
      </div>
      <div class="donut-titulo">{nome}</div>
    </div>
    <div class="painel-direita">
      <div class="painel-titulo">{nome} &nbsp;<span style="color:#555;font-size:12px;font-weight:400;">Base ERP: {total_erp} registros</span></div>
      <div class="metricas">
        <div class="metrica ok">
          <div class="m-numero">{len(ok)}</div>
          <div class="m-label">Conciliados OK</div>
          <div class="m-pct">{pct(len(ok))}%</div>
        </div>
        <div class="metrica so-erp">
          <div class="m-numero">{len(so_erp)}</div>
          <div class="m-label">So no ERP</div>
          <div class="m-pct">{pct(len(so_erp))}%</div>
        </div>
        <div class="metrica div-qtd">
          <div class="m-numero">{len(div_qtd)}</div>
          <div class="m-label">Div. Quantidade</div>
          <div class="m-pct">{pct(len(div_qtd))}%</div>
        </div>
      </div>
      <div class="barra-geral-wrap">
        <div class="barra-geral-label"><span>Distribuicao (base ERP)</span><span>{total_erp} registros</span></div>
        <div class="barra-geral">
          <div class="barra-seg ok"      style="width:{pct(len(ok))}%"></div>
          <div class="barra-seg so-erp"  style="width:{pct(len(so_erp))}%"></div>
          <div class="barra-seg div-qtd" style="width:{pct(len(div_qtd))}%"></div>
        </div>
        <div class="legenda">
          <div class="legenda-item"><div class="legenda-dot" style="background:#22c55e"></div> OK</div>
          <div class="legenda-item"><div class="legenda-dot" style="background:#60a5fa"></div> So ERP</div>
          <div class="legenda-item"><div class="legenda-dot" style="background:#f87171"></div> Div Qtd</div>
        </div>
      </div>
    </div>
  </div>

  <div class="tabs">
    <button class="tab-btn ativo" onclick="abrirAba('{sub_sg}','ok',this)">
      Conciliados <span class="badge ok">{len(ok)}</span>
    </button>
    <button class="tab-btn" onclick="abrirAba('{sub_sg}','so-erp',this)">
      So no ERP <span class="badge so-erp">{len(so_erp)}</span>
    </button>
    <button class="tab-btn" onclick="abrirAba('{sub_sg}','div-qtd',this)">
      Div. Quantidade <span class="badge div-qtd">{len(div_qtd)}</span>
    </button>
  </div>

  <div id="{sub_sg}-ok" class="tab-content ativo">
    <div class="acoes-wrap">
      <div class="filtro-wrap"><input type="text" placeholder="Filtrar..." onkeyup="filtrar(this,'{sub_sg}-ok')"></div>
    </div>
    <div class="table-wrap">{tab_ok}</div>
  </div>
  <div id="{sub_sg}-so-erp" class="tab-content" data-data="{data_str}" data-grupo="{grupo} / {nome}">
    <div class="acoes-wrap">
      <div class="filtro-wrap"><input type="text" placeholder="Filtrar..." onkeyup="filtrar(this,'{sub_sg}-so-erp')"></div>
    </div>
    <div class="table-wrap">{tab_so_erp}</div>
  </div>
  <div id="{sub_sg}-div-qtd" class="tab-content">
    <div class="acoes-wrap">
      <div class="filtro-wrap"><input type="text" placeholder="Filtrar..." onkeyup="filtrar(this,'{sub_sg}-div-qtd')"></div>
    </div>
    {tab_div_qtd}
  </div>
</div>"""


def _secao_grupo_com_subgrupos(grupo, resultados, data_slug="", data_str=""):
    ok        = resultados["ok"]
    so_erp    = resultados["so_erp"]
    div_qtd   = resultados["div_qtd"]
    subgrupos = resultados["subgrupos"]

    sg        = (data_slug + "_" if data_slug else "") + _slug(grupo)
    sg_geral  = sg + "_geral"
    total_erp = len(ok) + len(so_erp) + len(div_qtd)
    pct_ok    = round(len(ok) / total_erp * 100, 1) if total_erp else 0

    def pct(n): return round(n / total_erp * 100, 1) if total_erp else 0

    rotulos_ok  = {"pedido": "Pedido", "numero_nf": "NF", "sku": "SKU",
                   "unidade_erp": "Un.", "quantidade_wms": "Qtd WMS",
                   "quantidade_erp": "Qtd ERP", "data_erp": "Data"}
    rotulos_erp = {"pedido": "Pedido", "numero_nf": "NF", "sku": "SKU",
                   "unidade_erp": "Un.", "quantidade_erp": "Qtd ERP", "data_erp": "Data"}
    cols_ok      = ["pedido", "numero_nf", "sku", "unidade_erp", "quantidade_wms", "quantidade_erp", "data_erp"]
    cols_so_erp  = ["pedido", "numero_nf", "sku", "unidade_erp", "quantidade_erp", "data_erp"]
    cols_div_qtd = ["pedido", "numero_nf", "sku", "unidade_erp", "quantidade_wms", "quantidade_erp", "diff_qtd", "valor_div", "data_erp"]
    tab_ok      = _df_para_html(ok,      cols_ok,     rotulos_ok)
    tab_so_erp  = _df_para_html(so_erp,  cols_so_erp, rotulos_erp)
    tab_div_qtd = _tabela_div_html(div_qtd, cols_div_qtd, sg_geral, data_str, grupo)

    sub_btns = ""
    for i, (sub_nome, sub_r) in enumerate(subgrupos.items()):
        sub_total = len(sub_r["ok"]) + len(sub_r["so_erp"]) + len(sub_r["div_qtd"])
        ativo = ' ativo' if i == 0 else ''
        sub_btns += (
            f'<button class="subgrupo-btn{ativo}" '
            f'onclick="abrirSubgrupo(\'{sg}\',\'{_slug(sub_nome)}\')">'
            f'{sub_nome} <span style="font-size:11px;opacity:.6;">({sub_total})</span></button>\n'
        )

    sub_secoes = "".join(
        _secao_subgrupo(sub_nome, sub_r, sg, data_str, grupo, is_first=(i == 0))
        for i, (sub_nome, sub_r) in enumerate(subgrupos.items())
    )

    return f"""
<div id="grupo-{sg}" class="grupo-section" style="display:none;">

  <div class="painel">
    <div class="donut-wrap">
      <div class="donut-container">
        <canvas id="donut-{sg}" width="180" height="180"></canvas>
        <div class="donut-centro">
          <div class="pct">{pct_ok}%</div>
          <div class="sub">Conciliado</div>
        </div>
      </div>
      <div class="donut-titulo">{grupo} — Total</div>
    </div>
    <div class="painel-direita">
      <div class="painel-titulo">Resumo {grupo} &nbsp;<span style="color:#555;font-size:12px;font-weight:400;">Base ERP: {total_erp} registros</span></div>
      <div class="metricas">
        <div class="metrica ok">
          <div class="m-numero">{len(ok)}</div>
          <div class="m-label">Conciliados OK</div>
          <div class="m-pct">{pct(len(ok))}%</div>
        </div>
        <div class="metrica so-erp">
          <div class="m-numero">{len(so_erp)}</div>
          <div class="m-label">So no ERP</div>
          <div class="m-pct">{pct(len(so_erp))}%</div>
        </div>
        <div class="metrica div-qtd">
          <div class="m-numero">{len(div_qtd)}</div>
          <div class="m-label">Div. Quantidade</div>
          <div class="m-pct">{pct(len(div_qtd))}%</div>
        </div>
      </div>
      <div class="barra-geral-wrap">
        <div class="barra-geral-label"><span>Distribuicao (base ERP)</span><span>{total_erp} registros</span></div>
        <div class="barra-geral">
          <div class="barra-seg ok"      style="width:{pct(len(ok))}%"></div>
          <div class="barra-seg so-erp"  style="width:{pct(len(so_erp))}%"></div>
          <div class="barra-seg div-qtd" style="width:{pct(len(div_qtd))}%"></div>
        </div>
        <div class="legenda">
          <div class="legenda-item"><div class="legenda-dot" style="background:#22c55e"></div> OK</div>
          <div class="legenda-item"><div class="legenda-dot" style="background:#60a5fa"></div> So ERP</div>
          <div class="legenda-item"><div class="legenda-dot" style="background:#f87171"></div> Div Qtd</div>
        </div>
      </div>
    </div>
  </div>

  <!-- Abas gerais (todos os clientes combinados) -->
  <div id="geral-{sg_geral}">
    <div class="tabs">
      <button class="tab-btn ativo" onclick="abrirAba('{sg_geral}','ok',this)">
        Conciliados <span class="badge ok">{len(ok)}</span>
      </button>
      <button class="tab-btn" onclick="abrirAba('{sg_geral}','so-erp',this)">
        So no ERP <span class="badge so-erp">{len(so_erp)}</span>
      </button>
      <button class="tab-btn" onclick="abrirAba('{sg_geral}','div-qtd',this)">
        Div. Quantidade <span class="badge div-qtd">{len(div_qtd)}</span>
      </button>
    </div>
    <div id="{sg_geral}-ok" class="tab-content ativo">
      <div class="acoes-wrap">
        <div class="filtro-wrap"><input type="text" placeholder="Filtrar..." onkeyup="filtrar(this,'{sg_geral}-ok')"></div>
      </div>
      <div class="table-wrap">{tab_ok}</div>
    </div>
    <div id="{sg_geral}-so-erp" class="tab-content" data-data="{data_str}" data-grupo="{grupo}">
      <div class="acoes-wrap">
        <div class="filtro-wrap"><input type="text" placeholder="Filtrar..." onkeyup="filtrar(this,'{sg_geral}-so-erp')"></div>
      </div>
      <div class="table-wrap">{tab_so_erp}</div>
    </div>
    <div id="{sg_geral}-div-qtd" class="tab-content">
      <div class="acoes-wrap">
        <div class="filtro-wrap"><input type="text" placeholder="Filtrar..." onkeyup="filtrar(this,'{sg_geral}-div-qtd')"></div>
      </div>
      {tab_div_qtd}
    </div>
  </div>

  <div class="detalhe-label">Detalhe por Cliente</div>

  <div class="subgrupo-nav">
    <span class="subgrupo-nav-label">Clientes:</span>
    {sub_btns}
  </div>

  {sub_secoes}

</div>"""


# ── Seção de grupo (HTML) ─────────────────────────────────────────────────────

def _secao_grupo(grupo, resultados, data_slug="", data_str=""):
    if "subgrupos" in resultados:
        return _secao_grupo_com_subgrupos(grupo, resultados, data_slug, data_str)
    ok      = resultados["ok"]
    so_wms  = resultados["so_wms"]
    so_erp  = resultados["so_erp"]
    div_qtd = resultados["div_qtd"]

    sg = (data_slug + "_" if data_slug else "") + _slug(grupo)

    total_erp = len(ok) + len(so_erp) + len(div_qtd)
    pct_ok    = round(len(ok) / total_erp * 100, 1) if total_erp else 0

    def pct(n): return round(n / total_erp * 100, 1) if total_erp else 0

    rotulos_ok  = {"pedido": "Pedido", "numero_nf": "NF", "sku": "SKU",
                   "unidade_erp": "Un.", "quantidade_wms": "Qtd WMS",
                   "quantidade_erp": "Qtd ERP", "data_erp": "Data"}
    rotulos_erp = {"pedido": "Pedido", "numero_nf": "NF", "sku": "SKU",
                   "unidade_erp": "Un.", "quantidade_erp": "Qtd ERP", "data_erp": "Data"}
    rotulos_wms = {"pedido": "Pedido", "sku": "SKU", "unidade_wms": "Un.",
                   "quantidade_wms": "Qtd WMS", "data_wms": "Data"}

    cols_ok      = ["pedido", "numero_nf", "sku", "unidade_erp", "quantidade_wms", "quantidade_erp", "data_erp"]
    cols_so_erp  = ["pedido", "numero_nf", "sku", "unidade_erp", "quantidade_erp", "data_erp"]
    cols_div_qtd = ["pedido", "numero_nf", "sku", "unidade_erp", "quantidade_wms", "quantidade_erp", "diff_qtd", "valor_div", "data_erp"]

    tab_ok      = _df_para_html(ok,      cols_ok,     rotulos_ok)
    tab_so_erp  = _df_para_html(so_erp,  cols_so_erp, rotulos_erp)
    tab_div_qtd = _tabela_div_html(div_qtd, cols_div_qtd, sg, data_str, grupo)

    return f"""
<div id="grupo-{sg}" class="grupo-section" style="display:none;">

  <!-- Painel com donut + métricas -->
  <div class="painel">
    <div class="donut-wrap">
      <div class="donut-container">
        <canvas id="donut-{sg}" width="180" height="180"></canvas>
        <div class="donut-centro">
          <div class="pct">{pct_ok}%</div>
          <div class="sub">Conciliado</div>
        </div>
      </div>
      <div class="donut-titulo">{grupo}</div>
    </div>
    <div class="painel-direita">
      <div class="painel-titulo">Resumo &nbsp;<span style="color:#555;font-size:12px;font-weight:400;">Base ERP: {total_erp} registros</span></div>
      <div class="metricas">
        <div class="metrica ok">
          <div class="m-numero">{len(ok)}</div>
          <div class="m-label">Conciliados OK</div>
          <div class="m-pct">{pct(len(ok))}%</div>
        </div>
        <div class="metrica so-erp">
          <div class="m-numero">{len(so_erp)}</div>
          <div class="m-label">So no ERP</div>
          <div class="m-pct">{pct(len(so_erp))}%</div>
        </div>
        <div class="metrica div-qtd">
          <div class="m-numero">{len(div_qtd)}</div>
          <div class="m-label">Div. Quantidade</div>
          <div class="m-pct">{pct(len(div_qtd))}%</div>
        </div>
      </div>
      <div class="barra-geral-wrap">
        <div class="barra-geral-label"><span>Distribuicao (base ERP)</span><span>{total_erp} registros</span></div>
        <div class="barra-geral">
          <div class="barra-seg ok"      style="width:{pct(len(ok))}%"></div>
          <div class="barra-seg so-erp"  style="width:{pct(len(so_erp))}%"></div>
          <div class="barra-seg div-qtd" style="width:{pct(len(div_qtd))}%"></div>
        </div>
        <div class="legenda">
          <div class="legenda-item"><div class="legenda-dot" style="background:#22c55e"></div> OK</div>
          <div class="legenda-item"><div class="legenda-dot" style="background:#60a5fa"></div> So ERP</div>
          <div class="legenda-item"><div class="legenda-dot" style="background:#f87171"></div> Div Qtd</div>
        </div>
      </div>
    </div>
  </div>

  <!-- Abas -->
  <div class="tabs">
    <button class="tab-btn ativo" onclick="abrirAba('{sg}','ok',this)">
      Conciliados <span class="badge ok">{len(ok)}</span>
    </button>
    <button class="tab-btn" onclick="abrirAba('{sg}','so-erp',this)">
      So no ERP <span class="badge so-erp">{len(so_erp)}</span>
    </button>
    <button class="tab-btn" onclick="abrirAba('{sg}','div-qtd',this)">
      Div. Quantidade <span class="badge div-qtd">{len(div_qtd)}</span>
    </button>
  </div>

  <div id="{sg}-ok" class="tab-content ativo">
    <div class="acoes-wrap">
      <div class="filtro-wrap"><input type="text" placeholder="Filtrar..." onkeyup="filtrar(this,'{sg}-ok')"></div>
    </div>
    <div class="table-wrap">{tab_ok}</div>
  </div>
  <div id="{sg}-so-erp" class="tab-content" data-data="{data_str}" data-grupo="{grupo}">
    <div class="acoes-wrap">
      <div class="filtro-wrap"><input type="text" placeholder="Filtrar..." onkeyup="filtrar(this,'{sg}-so-erp')"></div>
    </div>
    <div class="table-wrap">{tab_so_erp}</div>
  </div>
  <div id="{sg}-div-qtd" class="tab-content">
    <div class="acoes-wrap">
      <div class="filtro-wrap"><input type="text" placeholder="Filtrar..." onkeyup="filtrar(this,'{sg}-div-qtd')"></div>
    </div>
    {tab_div_qtd}
  </div>

</div>"""


def _donut_js(grupo, resultados, data_slug=""):
    ok      = len(resultados["ok"])
    so_erp  = len(resultados["so_erp"])
    div_qtd = len(resultados["div_qtd"])
    sg = (data_slug + "_" if data_slug else "") + _slug(grupo)
    js = f"  desenharDonut('donut-{sg}', [{ok},{so_erp},{div_qtd}]);\n"
    if "subgrupos" in resultados:
        for sub_nome, sub_r in resultados["subgrupos"].items():
            sub_sg = sg + "_sub_" + _slug(sub_nome)
            s_ok  = len(sub_r["ok"])
            s_erp = len(sub_r["so_erp"])
            s_div = len(sub_r["div_qtd"])
            js += f"  desenharDonut('donut-{sub_sg}', [{s_ok},{s_erp},{s_div}]);\n"
    return js


# ── Dashboard principal ───────────────────────────────────────────────────────

def gerar_dashboard(todos, data_ref):
    datas = sorted(todos.keys())

    botoes_data = ""
    for d in datas:
        d_slug  = _slug(d) if d else "sem_data"
        d_label = d.replace('.', '/') if d else "–"
        ativo   = ' ativo' if d == datas[0] else ''
        botoes_data += f'<button class="data-btn{ativo}" onclick="abrirData(\'{d_slug}\')">{d_label}</button>\n'

    secoes_data = ""
    init_js     = ""
    donuts_js   = ""

    for d in datas:
        d_slug   = _slug(d) if d else "sem_data"
        grupos_d = todos[d]
        grupos   = list(grupos_d.keys())

        ok_g  = sum(len(r["ok"])       for r in grupos_d.values())
        erp_g = sum(len(r["so_erp"])  for r in grupos_d.values())
        div_g = sum(len(r["div_qtd"]) for r in grupos_d.values())
        tot_g = ok_g + erp_g + div_g   # base ERP apenas
        pct_g = round(ok_g / tot_g * 100, 1) if tot_g else 0

        cards = ""
        for g, r in grupos_d.items():
            sg  = _slug(g)
            t   = len(r["ok"]) + len(r["so_erp"]) + len(r["div_qtd"])
            pct = round(len(r["ok"]) / t * 100, 1) if t else 0
            cor_pct = "#22c55e" if pct == 100 else "#f59e0b" if pct >= 80 else "#f87171"
            cor_g   = _cor_grupo(g)
            cards += f"""
        <div class="card-grupo" data-color="{cor_g}" onclick="abrirGrupo('{d_slug}','{sg}')">
          <div class="cg-nome">{g}</div>
          <div class="cg-pct" style="color:{cor_pct};">{pct}%</div>
          <div class="cg-sub">OK</div>
          <div class="cg-detalhe">
            <span style="color:#22c55e;">&#10003; {len(r['ok'])}</span>
            <span style="color:#60a5fa;">E {len(r['so_erp'])}</span>
            <span style="color:#f87171;">D {len(r['div_qtd'])}</span>
          </div>
        </div>"""

        btns_grupo = ""
        for g in grupos:
            sg    = _slug(g)
            ativo = ' ativo' if g == grupos[0] else ''
            cor_g = _cor_grupo(g)
            btns_grupo += f'<button class="grupo-btn{ativo}" data-color="{cor_g}" onclick="abrirGrupo(\'{d_slug}\',\'{sg}\')">{g}</button>\n'

        secoes_grupos = "".join(_secao_grupo(g, r, d_slug, d) for g, r in grupos_d.items())
        donuts_js    += "".join(_donut_js(g, r, d_slug)       for g, r in grupos_d.items())

        primeiro_sg = _slug(grupos[0])
        init_js    += f"abrirGrupo('{d_slug}','{primeiro_sg}');\n"

        display = '' if d == datas[0] else 'display:none;'
        cor_pct = "#22c55e" if pct_g == 100 else "#f59e0b" if pct_g >= 80 else "#f87171"
        secoes_data += f"""
<div id="data-section-{d_slug}" class="data-section" style="{display}">
  <div class="resumo-geral">
    <div class="rg-total">
      <strong>{ok_g}</strong> de <strong>{tot_g}</strong> itens ERP conciliados &mdash;
      <strong style="color:{cor_pct};">{pct_g}%</strong> OK
    </div>
    <div style="flex:1;min-width:200px;">
      <div class="barra-geral">
        <div class="barra-seg ok"      style="width:{round(ok_g/tot_g*100,1) if tot_g else 0}%"></div>
        <div class="barra-seg so-erp"  style="width:{round(erp_g/tot_g*100,1) if tot_g else 0}%"></div>
        <div class="barra-seg div-qtd" style="width:{round(div_g/tot_g*100,1) if tot_g else 0}%"></div>
      </div>
    </div>
  </div>
  <div class="cards-grupos">{cards}</div>
  <div class="grupo-nav">{btns_grupo}</div>
  {secoes_grupos}
</div>"""

    # ── Dados para painel Dashboard ───────────────────────────────────────────
    grupos_uniq = sorted(set(g for grupos_d in todos.values() for g in grupos_d))
    dash_data   = []
    for d in datas:
        for g in grupos_uniq:
            r = todos[d].get(g)
            if not r:
                continue
            ok_n  = len(r["ok"]); erp_n = len(r["so_erp"]); div_n = len(r["div_qtd"])
            tot   = ok_n + erp_n + div_n
            dash_data.append({
                "data": d, "grupo": g,
                "ok": ok_n, "erp": erp_n, "div": div_n, "total": tot,
                "pct_ok": round(ok_n / tot * 100, 1) if tot else 0,
            })
    dash_json   = json.dumps(dash_data,  ensure_ascii=False)
    grupos_json = json.dumps(grupos_uniq)
    cores_json  = json.dumps({g: _cor_grupo(g) for g in grupos_uniq})
    todos_grupos_str = " | ".join(grupos_uniq)

    now_str = datetime.now().strftime("%d/%m/%Y %H:%M")

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Conciliacao de Saida</title>
<style>
  * {{ box-sizing:border-box; margin:0; padding:0; }}
  body {{ font-family:'Segoe UI',sans-serif; background:#111111; color:#e0e0e0; }}

  header {{
    display:flex; align-items:center; justify-content:space-between;
    padding:16px 28px; border-bottom:1px solid #2d2d2d;
    position:sticky; top:0; z-index:100; background:#111111; gap:12px; flex-wrap:wrap;
  }}
  header h1 {{ font-size:18px; font-weight:700; color:#f1f5f9; }}
  .meta {{ font-size:12px; color:#666; }}
  .container {{ max-width:1400px; margin:0 auto; padding:24px 28px; }}

  /* Data nav */
  .data-nav {{ display:flex; gap:8px; flex-wrap:wrap; margin-bottom:24px; padding-bottom:20px; border-bottom:1px solid #2d2d2d; }}
  .data-btn {{ padding:10px 28px; border:1px solid #2d2d2d; border-radius:8px; background:#1a1a1a; color:#666; font-size:15px; font-weight:700; cursor:pointer; transition:all .2s; letter-spacing:.5px; }}
  .data-btn:hover {{ border-color:#6366f1; color:#818cf8; }}
  .data-btn.ativo  {{ background:#6366f1; border-color:#6366f1; color:#fff; }}

  /* Resumo geral */
  .resumo-geral {{ display:flex; align-items:center; gap:20px; background:#1a1a1a; border:1px solid #2d2d2d; border-radius:10px; padding:14px 18px; margin-bottom:20px; flex-wrap:wrap; }}
  .rg-total {{ font-size:13px; color:#94a3b8; white-space:nowrap; }}

  /* Cards de grupo */
  .cards-grupos {{ display:flex; flex-wrap:wrap; gap:14px; margin-bottom:20px; }}
  .card-grupo {{
    background:#1e1e1e; border:2px solid #2d2d2d; border-radius:12px;
    padding:16px 20px; cursor:pointer; min-width:160px;
    transition:border-color .2s, box-shadow .2s;
  }}
  .card-grupo:hover {{ border-color:#555; box-shadow:0 4px 16px rgba(0,0,0,.4); }}
  .card-grupo.ativo {{ border-color:#6366f1; box-shadow:0 4px 20px rgba(99,102,241,.2); }}
  .cg-nome {{ font-size:11px; font-weight:700; color:#888; text-transform:uppercase; letter-spacing:.8px; margin-bottom:6px; }}
  .cg-pct  {{ font-size:28px; font-weight:800; line-height:1; margin-bottom:2px; }}
  .cg-sub  {{ font-size:10px; color:#555; margin-bottom:8px; }}
  .cg-detalhe {{ display:flex; gap:10px; font-size:11px; font-weight:700; }}

  /* Navegação por grupo */
  .grupo-nav {{ display:flex; gap:8px; flex-wrap:wrap; margin-bottom:20px; }}
  .grupo-btn {{
    padding:9px 22px; border:1px solid #2d2d2d; border-radius:8px;
    background:#1e1e1e; color:#888; font-size:13px; font-weight:600;
    cursor:pointer; transition:all .2s;
  }}
  .grupo-btn:hover {{ border-color:#6366f1; color:#818cf8; }}
  .grupo-btn.ativo  {{ background:#6366f1; border-color:#6366f1; color:#fff; }}

  /* Painel com donut */
  .grupo-section {{ display:none; }}
  .painel {{
    display:grid; grid-template-columns:240px 1fr; gap:24px;
    background:#1e1e1e; border:1px solid #2d2d2d; border-radius:14px;
    padding:24px; margin-bottom:20px; box-shadow:0 4px 20px rgba(0,0,0,.4);
  }}
  .donut-wrap {{ display:flex; flex-direction:column; align-items:center; justify-content:center; gap:10px; }}
  .donut-container {{ position:relative; width:180px; height:180px; }}
  .donut-container canvas {{ display:block; }}
  .donut-centro {{ position:absolute; top:50%; left:50%; transform:translate(-50%,-50%); text-align:center; }}
  .donut-centro .pct {{ font-size:30px; font-weight:700; color:#f1f5f9; line-height:1; }}
  .donut-centro .sub {{ font-size:11px; color:#666; margin-top:4px; }}
  .donut-titulo {{ font-size:12px; color:#666; text-align:center; }}
  .painel-direita {{ display:flex; flex-direction:column; justify-content:space-between; gap:14px; }}
  .painel-titulo {{ font-size:14px; font-weight:600; color:#ccc; border-bottom:1px solid #2d2d2d; padding-bottom:10px; }}

  /* Métricas */
  .metricas {{ display:grid; grid-template-columns:repeat(3,1fr); gap:12px; }}
  .metrica {{ background:#191919; border-radius:8px; padding:14px; border-left:3px solid #333; }}
  .metrica .m-numero {{ font-size:26px; font-weight:700; line-height:1; }}
  .metrica .m-label  {{ font-size:11px; color:#666; margin-top:4px; }}
  .metrica .m-pct    {{ font-size:11px; font-weight:600; margin-top:5px; }}
  .metrica.ok      {{ border-color:#22c55e; }} .metrica.ok .m-numero,.metrica.ok .m-pct           {{ color:#22c55e; }}
  .metrica.so-erp  {{ border-color:#60a5fa; }} .metrica.so-erp .m-numero,.metrica.so-erp .m-pct   {{ color:#60a5fa; }}
  .metrica.div-qtd {{ border-color:#f87171; }} .metrica.div-qtd .m-numero,.metrica.div-qtd .m-pct {{ color:#f87171; }}

  /* Barra geral */
  .barra-geral-wrap {{ }}
  .barra-geral-label {{ display:flex; justify-content:space-between; font-size:11px; color:#555; margin-bottom:6px; }}
  .barra-geral {{ display:flex; height:6px; border-radius:3px; overflow:hidden; background:#2a2a2a; margin-bottom:8px; }}
  .barra-seg {{ height:100%; }}
  .barra-seg.ok      {{ background:#22c55e; }}
  .barra-seg.so-erp  {{ background:#60a5fa; }}
  .barra-seg.div-qtd {{ background:#f87171; }}
  .legenda {{ display:flex; gap:14px; flex-wrap:wrap; margin-top:8px; }}
  .legenda-item {{ display:flex; align-items:center; gap:6px; font-size:11px; color:#888; }}
  .legenda-dot {{ width:9px; height:9px; border-radius:50%; flex-shrink:0; }}

  /* Abas */
  .tabs {{ display:flex; gap:4px; flex-wrap:wrap; }}
  .tab-btn {{ padding:9px 18px; border:none; border-radius:8px 8px 0 0; cursor:pointer; font-size:13px; font-weight:600; background:#191919; color:#888; transition:background .2s,color .2s; }}
  .tab-btn:hover {{ background:#222; color:#ccc; }}
  .tab-btn.ativo {{ background:#1c1c1c; color:#818cf8; border-bottom:2px solid #6366f1; }}
  .badge {{ display:inline-block; min-width:20px; padding:1px 5px; border-radius:10px; font-size:11px; font-weight:700; margin-left:5px; color:#fff; }}
  .badge.ok {{ background:#22c55e; }} .badge.so-erp {{ background:#60a5fa; }} .badge.div-qtd {{ background:#f87171; }}
  .tab-content {{ background:#1c1c1c; border-radius:0 8px 8px 8px; padding:20px; box-shadow:0 2px 12px rgba(0,0,0,.3); display:none; }}
  .tab-content.ativo {{ display:block; }}

  /* Filtro e tabela */
  .acoes-wrap {{ display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; gap:10px; flex-wrap:wrap; }}
  .filtro-wrap input {{ padding:7px 12px; border:1px solid #333; border-radius:6px; font-size:13px; outline:none; background:#141414; color:#e0e0e0; width:280px; }}
  .filtro-wrap input:focus {{ border-color:#6366f1; }}
  .filtro-wrap input::placeholder {{ color:#555; }}
  .table-wrap {{ overflow-x:auto; }}
  .tabela {{ width:100%; border-collapse:collapse; font-size:13px; }}
  .tabela th {{ background:#1a1a1a; color:#94a3b8; padding:9px 12px; text-align:left; white-space:nowrap; border-bottom:1px solid #333; font-weight:600; }}
  .tabela td {{ padding:8px 12px; border-bottom:1px solid #222; white-space:nowrap; color:#d0d0d0; }}
  .tabela tr:hover td {{ background:#222; }}
  .tabela tr:nth-child(even) td {{ background:#1a1a1a; }}
  .vazio {{ color:#555; font-style:italic; padding:14px 0; }}
  .td-diff {{ font-variant-numeric:tabular-nums; }}
  .td-diff-pos {{ color:#F59E0B; font-weight:700; }}
  .td-diff-neg {{ color:#F87171; font-weight:700; }}
  .td-valordiv {{ text-align:right; font-variant-numeric:tabular-nums; }}
  .td-valordiv-pos {{ color:#F59E0B; }}
  .td-valordiv-neg {{ color:#F87171; }}
  footer {{ text-align:center; color:#444; font-size:11px; padding:28px 0; }}

  /* Navegação de subgrupos (subclientes NEXXA) */
  .subgrupo-nav {{ display:flex; align-items:center; gap:8px; flex-wrap:wrap; margin:16px 0 12px; padding:12px 16px; background:#191919; border:1px solid #2d2d2d; border-radius:8px; }}
  .subgrupo-nav-label {{ font-size:11px; color:#555; font-weight:600; text-transform:uppercase; letter-spacing:.5px; margin-right:4px; white-space:nowrap; }}
  .subgrupo-btn {{ padding:6px 16px; border:1px solid #2d2d2d; border-radius:6px; background:#1a1a1a; color:#888; font-size:12px; font-weight:600; cursor:pointer; transition:all .2s; }}
  .subgrupo-btn:hover {{ border-color:#a78bfa; color:#c4b5fd; }}
  .subgrupo-btn.ativo {{ background:#7c3aed; border-color:#7c3aed; color:#fff; }}
  .subgrupo-section {{ display:none; }}
  .painel-sub {{ margin-top:0; }}
  .detalhe-label {{ font-size:11px; color:#555; font-weight:600; text-transform:uppercase; letter-spacing:.8px; margin:20px 0 8px; padding-bottom:6px; border-bottom:1px solid #222; }}

  /* View toggle */
  .view-toggle {{ display:flex; gap:4px; background:#141414; padding:3px; border-radius:8px; border:1px solid #2d2d2d; }}
  .view-btn {{ padding:6px 16px; border:none; border-radius:6px; font-size:12px; font-weight:600; cursor:pointer; color:#666; background:transparent; transition:all .2s; }}
  .view-btn.ativo {{ background:#6366f1; color:#fff; }}

  /* Botões exportar */
  .btn-export-cons {{ display:flex; align-items:center; gap:6px; padding:8px 16px; background:transparent; border:1px solid #22c55e; color:#22c55e; border-radius:6px; font-size:12px; font-weight:600; cursor:pointer; transition:all .2s; white-space:nowrap; }}
  .btn-export-cons:hover {{ background:#22c55e; color:#fff; }}
  .btn-atualizar {{ display:flex; align-items:center; gap:6px; padding:8px 16px; background:transparent; border:1px solid #6366f1; color:#818cf8; border-radius:6px; font-size:12px; font-weight:600; cursor:pointer; transition:all .2s; }}
  .btn-atualizar:hover {{ background:#6366f1; color:#fff; }}

  /* Painel Análise / Dashboard */
  .analise-panel {{ display:none; }}
  .analise-panel.ativo {{ display:block; }}
  .dash-panel {{ display:none; }}
  .dash-panel.ativo {{ display:block; }}
  .dash-cards {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(200px,1fr)); gap:14px; margin-bottom:20px; }}
  .dash-card {{ background:#1e1e1e; border:1px solid #2d2d2d; border-radius:12px; padding:20px 24px; }}
  .dash-card-label {{ font-size:10px; color:#666; text-transform:uppercase; letter-spacing:.6px; margin-bottom:8px; }}
  .dash-card-valor {{ font-size:32px; font-weight:700; color:#f1f5f9; line-height:1; }}
  .dash-card-sub {{ font-size:12px; color:#666; margin-top:6px; }}
  .dash-secao {{ background:#1e1e1e; border:1px solid #2d2d2d; border-radius:12px; padding:20px 24px; margin-bottom:16px; }}
  .dash-secao-titulo {{ font-size:11px; font-weight:600; color:#94a3b8; margin-bottom:16px; text-transform:uppercase; letter-spacing:.5px; }}
  .dash-tabela {{ width:100%; border-collapse:collapse; font-size:13px; }}
  .dash-tabela th {{ background:#141414; color:#94a3b8; padding:9px 14px; text-align:left; border-bottom:1px solid #333; font-weight:600; white-space:nowrap; }}
  .dash-tabela td {{ padding:9px 14px; border-bottom:1px solid #222; color:#d0d0d0; white-space:nowrap; }}
  .dash-tabela tr:hover td {{ background:#222; }}
  .dash-tabela .pct-bar {{ display:inline-block; height:6px; border-radius:3px; vertical-align:middle; margin-left:8px; background:#22c55e; }}
</style>
</head>
<body>

<header>
  <div style="display:flex;align-items:center;gap:16px;">
    <a href="index.html" style="display:flex;align-items:center;gap:6px;color:#818cf8;text-decoration:none;font-size:13px;font-weight:600;border:1px solid #6366f1;border-radius:6px;padding:6px 12px;"
       onmouseover="this.style.background='#6366f1';this.style.color='#fff'"
       onmouseout="this.style.background='transparent';this.style.color='#818cf8'">&#8592; Inicio</a>
    <div>
      <h1>Conciliacao de Saida</h1>
      <div class="meta">WMS vs ERP &mdash; {data_ref} &nbsp;&bull;&nbsp; Grupos: {todos_grupos_str} &nbsp;&bull;&nbsp; Gerado em {now_str}</div>
    </div>
  </div>
  <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
    <div class="view-toggle">
      <button class="view-btn ativo" id="vbtn-analise"   onclick="alternarView('analise')">&#128203; Analise</button>
      <button class="view-btn"       id="vbtn-dashboard" onclick="alternarView('dashboard')">&#128202; Dashboard</button>
    </div>
    <button class="btn-export-cons" id="btn-exportar-div"   onclick="exportarDivergencias()">&#11015; Exportar Divergencias</button>
    <button class="btn-export-cons" id="btn-exportar-soerp" onclick="exportarSoERP()" style="border-color:#3b82f6;color:#93c5fd;background:#1e3a5f20;">&#11015; Exportar So ERP</button>
    <button class="btn-atualizar"   id="btn-atualizar"      onclick="atualizarDados()">&#8635; Atualizar</button>
  </div>
</header>

<div class="container">

<div id="painel-analise" class="analise-panel ativo">
  {'<div class="data-nav">' + botoes_data + '</div>' if len(datas) > 1 else ''}
  {secoes_data}
</div>

<div id="painel-dashboard" class="dash-panel">
  <div class="dash-cards" id="dash-cards-sumario"></div>
  <div class="dash-secao">
    <div class="dash-secao-titulo">% Conciliacao OK por Grupo</div>
    <canvas id="dash-bar-canvas" style="width:100%;max-height:320px;"></canvas>
  </div>
  <div class="dash-secao">
    <div class="dash-secao-titulo">Detalhe por Grupo</div>
    <table class="dash-tabela">
      <thead><tr>
        <th>Data</th><th>Grupo</th><th>OK</th><th>So ERP</th><th>Divergencias</th><th>% OK</th>
      </tr></thead>
      <tbody id="dash-tbody"></tbody>
    </table>
  </div>
</div>

</div>

<footer>COMFRIO FOOD SERVICE &mdash; Conciliacao de Saida</footer>

<script>
const _DASH        = {dash_json};
const _GRUPOS_UNIQ = {grupos_json};
const _CORES_GRUPO = {cores_json};
let _dashIniciado  = false;

function desenharDonut(id, vals) {{
  const cores = ['#22c55e','#60a5fa','#f87171'];
  const canvas = document.getElementById(id);
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const cx=90,cy=90,r=72,inner=48;
  const total = vals.reduce((s,v)=>s+v,0);
  let ang = -Math.PI/2;
  ctx.clearRect(0,0,180,180);
  if (total===0) {{
    ctx.beginPath(); ctx.arc(cx,cy,r,0,Math.PI*2); ctx.fillStyle='#2a2a2a'; ctx.fill();
    ctx.beginPath(); ctx.arc(cx,cy,inner,0,Math.PI*2); ctx.fillStyle='#1e1e1e'; ctx.fill();
    return;
  }}
  vals.forEach((v,i) => {{
    if(v===0) return;
    const fatia=(v/total)*Math.PI*2;
    ctx.beginPath(); ctx.moveTo(cx,cy); ctx.arc(cx,cy,r,ang,ang+fatia); ctx.closePath();
    ctx.fillStyle=cores[i]; ctx.fill(); ang+=fatia;
  }});
  ctx.beginPath(); ctx.arc(cx,cy,inner,0,Math.PI*2); ctx.fillStyle='#1e1e1e'; ctx.fill();
}}

function alternarView(view) {{
  document.getElementById('painel-analise').classList.toggle('ativo',  view==='analise');
  document.getElementById('painel-dashboard').classList.toggle('ativo', view==='dashboard');
  document.getElementById('vbtn-analise').classList.toggle('ativo',   view==='analise');
  document.getElementById('vbtn-dashboard').classList.toggle('ativo', view==='dashboard');
  if (view==='dashboard') inicializarDashboard();
}}

function inicializarDashboard() {{
  if (_dashIniciado) return;
  _dashIniciado = true;
  let totalDiv=0, totalOk=0, totalErp=0;
  _DASH.forEach(d => {{ totalDiv+=d.div; totalOk+=d.ok; totalErp+=d.erp; }});
  const cardsEl = document.getElementById('dash-cards-sumario');
  if (cardsEl) cardsEl.innerHTML = `
    <div class="dash-card"><div class="dash-card-label">Conciliados OK</div><div class="dash-card-valor" style="color:#22c55e">${{totalOk}}</div><div class="dash-card-sub">WMS e ERP alinhados</div></div>
    <div class="dash-card"><div class="dash-card-label">So no ERP</div><div class="dash-card-valor" style="color:#60a5fa">${{totalErp}}</div><div class="dash-card-sub">pedidos sem separacao no WMS</div></div>
    <div class="dash-card"><div class="dash-card-label">Divergencias</div><div class="dash-card-valor" style="color:#f87171">${{totalDiv}}</div><div class="dash-card-sub">itens com diff de quantidade</div></div>
    <div class="dash-card"><div class="dash-card-label">Grupos</div><div class="dash-card-valor" style="color:#6366f1">${{_GRUPOS_UNIQ.length}}</div><div class="dash-card-sub">clientes analisados</div></div>
  `;
  const tbody = document.getElementById('dash-tbody');
  if (tbody) {{
    tbody.innerHTML = '';
    _DASH.forEach(d => {{
      const bar = `<div class="pct-bar" style="width:${{Math.round(d.pct_ok * 0.8)}}px;"></div>`;
      tbody.innerHTML += `<tr>
        <td>${{d.data || '—'}}</td><td>${{d.grupo}}</td>
        <td style="color:#22c55e">${{d.ok}}</td>
        <td style="color:#60a5fa">${{d.erp}}</td>
        <td style="color:#f87171">${{d.div}}</td>
        <td>${{d.pct_ok}}%${{bar}}</td>
      </tr>`;
    }});
  }}
  desenharBarChart();
}}

function desenharBarChart() {{
  const canvas = document.getElementById('dash-bar-canvas');
  if (!canvas || !_DASH.length) return;
  const dpr = window.devicePixelRatio || 1;
  const W = canvas.parentElement.clientWidth || 800;
  const H = 300;
  canvas.width  = W * dpr; canvas.height = H * dpr;
  canvas.style.width = W+'px'; canvas.style.height = H+'px';
  const ctx = canvas.getContext('2d');
  ctx.scale(dpr, dpr);
  const pad = {{top:20,right:20,bottom:60,left:50}};
  const cW = W-pad.left-pad.right, cH = H-pad.top-pad.bottom;
  const byData = {{}};
  _DASH.forEach(d => {{ if(!byData[d.data]) byData[d.data]={{}}; byData[d.data][d.grupo]=d.pct_ok; }});
  const datas = Object.keys(byData);
  const nD = datas.length, nG = _GRUPOS_UNIQ.length;
  const grpW = cW/nD, barW = Math.min(grpW/(nG+1),40);
  ctx.strokeStyle='#333'; ctx.lineWidth=1;
  ctx.beginPath(); ctx.moveTo(pad.left,pad.top); ctx.lineTo(pad.left,pad.top+cH); ctx.stroke();
  ctx.beginPath(); ctx.moveTo(pad.left,pad.top+cH); ctx.lineTo(pad.left+cW,pad.top+cH); ctx.stroke();
  ctx.fillStyle='#666'; ctx.font='10px sans-serif'; ctx.textAlign='right';
  for(let i=0;i<=4;i++) {{
    const y=pad.top+cH-(i/4)*cH, v=Math.round((i/4)*100);
    ctx.fillText(v+'%',pad.left-6,y+3);
    ctx.strokeStyle='#222'; ctx.beginPath(); ctx.moveTo(pad.left,y); ctx.lineTo(pad.left+cW,y); ctx.stroke();
  }}
  datas.forEach((data,di) => {{
    const baseX=pad.left+di*grpW+(grpW-nG*barW)/2;
    _GRUPOS_UNIQ.forEach((g,gi) => {{
      const val=byData[data][g]||0;
      const bH=val===0?0:Math.max(2,(val/100)*cH);
      const x=baseX+gi*barW, y=pad.top+cH-bH;
      ctx.fillStyle=_CORES_GRUPO[g]||'#6366f1'; ctx.fillRect(x,y,barW-2,bH);
      if(val>0){{ ctx.fillStyle='#ddd'; ctx.textAlign='center'; ctx.font='10px sans-serif'; ctx.fillText(val+'%',x+(barW-2)/2,y-4); }}
    }});
    ctx.fillStyle='#888'; ctx.textAlign='center'; ctx.font='11px sans-serif';
    ctx.fillText(data||'—',pad.left+di*grpW+grpW/2,pad.top+cH+18);
  }});
  _GRUPOS_UNIQ.forEach((g,gi) => {{
    const lx=pad.left+gi*120, ly=H-14;
    ctx.fillStyle=_CORES_GRUPO[g]||'#6366f1'; ctx.fillRect(lx,ly-8,10,10);
    ctx.fillStyle='#aaa'; ctx.textAlign='left'; ctx.font='11px sans-serif'; ctx.fillText(g,lx+14,ly);
  }});
}}

function abrirData(d) {{
  document.querySelectorAll('.data-section').forEach(el=>el.style.display='none');
  document.querySelectorAll('.data-btn').forEach(el=>el.classList.remove('ativo'));
  document.getElementById('data-section-'+d).style.display='block';
  document.querySelectorAll('.data-btn').forEach(btn=>{{
    if(btn.getAttribute('onclick').includes("'"+d+"'")) btn.classList.add('ativo');
  }});
}}

function abrirGrupo(d, sg) {{
  const fullSg = d ? d+'_'+sg : sg;
  const scope  = d ? document.getElementById('data-section-'+d) : document;
  scope.querySelectorAll('.grupo-section').forEach(el=>el.style.display='none');
  scope.querySelectorAll('.grupo-btn').forEach(el=>{{el.classList.remove('ativo');el.style.background='';el.style.borderColor='';}});
  scope.querySelectorAll('.card-grupo').forEach(el=>{{el.classList.remove('ativo');el.style.borderColor='';el.style.boxShadow='';}});
  const sec = document.getElementById('grupo-'+fullSg);
  if(sec) sec.style.display='block';
  scope.querySelectorAll('.grupo-btn').forEach(btn=>{{
    if(btn.getAttribute('onclick')&&btn.getAttribute('onclick').includes("'"+sg+"'")) {{
      btn.classList.add('ativo');
      const cor=btn.getAttribute('data-color');
      if(cor){{btn.style.background=cor;btn.style.borderColor=cor;}}
    }}
  }});
  scope.querySelectorAll('.card-grupo').forEach(c=>{{
    if(c.getAttribute('onclick')&&c.getAttribute('onclick').includes("'"+sg+"'")) {{
      c.classList.add('ativo');
      const cor=c.getAttribute('data-color');
      if(cor){{c.style.borderColor=cor;c.style.boxShadow='0 4px 20px '+cor+'40';}}
    }}
  }});
}}

function abrirAba(sg, id, btn) {{
  const container = document.getElementById('grupo-'+sg) || document.getElementById('subgrupo-'+sg) || document.getElementById('geral-'+sg);
  if (container) {{
    container.querySelectorAll('.tab-content').forEach(el=>el.classList.remove('ativo'));
    container.querySelectorAll('.tab-btn').forEach(el=>el.classList.remove('ativo'));
  }}
  document.getElementById(sg+'-'+id).classList.add('ativo');
  btn.classList.add('ativo');
}}

function abrirSubgrupo(sgParent, subSlug) {{
  const parentEl = document.getElementById('grupo-'+sgParent);
  if (!parentEl) return;
  parentEl.querySelectorAll('.subgrupo-section').forEach(el=>el.style.display='none');
  parentEl.querySelectorAll('.subgrupo-btn').forEach(el=>el.classList.remove('ativo'));
  const target = document.getElementById('subgrupo-'+sgParent+'_sub_'+subSlug);
  if (target) target.style.display='block';
  parentEl.querySelectorAll('.subgrupo-btn').forEach(btn=>{{
    if(btn.getAttribute('onclick')&&btn.getAttribute('onclick').includes("'"+subSlug+"'")) btn.classList.add('ativo');
  }});
}}

function filtrar(input, abaId) {{
  const t=input.value.toLowerCase();
  document.querySelectorAll('#'+abaId+' .tabela tbody tr').forEach(tr=>{{
    tr.style.display=tr.textContent.toLowerCase().includes(t)?'':'none';
  }});
}}

function _postXlsx(headers, linhas, filename, btnId) {{
  if(linhas.length===0){{ alert('Nenhum item encontrado.'); return; }}
  const btn = document.getElementById(btnId);
  if(btn){{ btn.disabled=true; btn.textContent='⏳ Gerando...'; }}
  fetch('/api/exportar_divergencias',{{
    method:'POST',
    headers:{{'Content-Type':'application/json'}},
    body:JSON.stringify({{headers,rows:linhas}})
  }})
  .then(r=>r.blob())
  .then(blob=>{{
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a'); a.href=url; a.download=filename; a.click();
    URL.revokeObjectURL(url);
  }})
  .catch(e=>alert('Erro ao gerar Excel: '+e))
  .finally(()=>{{ if(btn){{ btn.disabled=false; btn.textContent=btnId==='btn-exportar-div'?'↙ Exportar Divergencias':'↙ Exportar So ERP'; }} }});
}}

function exportarDivergencias() {{
  const headers = ['Data','Grupo','Pedido','NF','SKU','Unidade','Qtd WMS','Qtd ERP','Diferenca','Valor Div. (R$)'];
  const linhas  = [];
  document.querySelectorAll('tr[data-ajuste-key]').forEach(row => {{
    if(row.style.display==='none') return;
    const data  = row.getAttribute('data-data')  || '';
    const grupo = row.getAttribute('data-grupo') || '';
    const vals  = [...row.querySelectorAll('td')].map(td=>td.textContent.trim());
    // cols: Pedido=0,NF=1,SKU=2,Un=3,QWms=4,QErp=5,Diff=6,ValorDiv=7,Data=8
    linhas.push([data,grupo,vals[0],vals[1],vals[2],vals[3],vals[4],vals[5],vals[6],vals[7]]);
  }});
  _postXlsx(headers, linhas, 'saida_divergencias.xlsx', 'btn-exportar-div');
}}

function exportarSoERP() {{
  const headers = ['Data','Grupo','Pedido','NF','SKU','Unidade','Qtd ERP','Data ERP'];
  const linhas  = [];
  document.querySelectorAll('.tab-content[id$="-so-erp"]').forEach(tab => {{
    const data  = tab.getAttribute('data-data')  || '';
    const grupo = tab.getAttribute('data-grupo') || '';
    tab.querySelectorAll('.tabela tbody tr').forEach(row => {{
      if(row.style.display==='none') return;
      const vals=[...row.querySelectorAll('td')].map(td=>td.textContent.trim());
      if(!vals[0]) return;
      // cols: Pedido=0,NF=1,SKU=2,Un=3,QErp=4,Data=5
      linhas.push([data,grupo,vals[0],vals[1],vals[2],vals[3],vals[4],vals[5]]);
    }});
  }});
  _postXlsx(headers, linhas, 'saida_so_erp.xlsx', 'btn-exportar-soerp');
}}

function atualizarDados() {{
  const btn = document.getElementById('btn-atualizar');
  if(btn){{ btn.disabled=true; btn.textContent='⏳ Atualizando...'; }}
  fetch('/api/atualizar_saida',{{method:'POST'}})
    .then(r=>r.json())
    .then(d=>{{ if(d.ok) location.reload(); else alert('Erro: '+d.msg); }})
    .catch(e=>alert('Servidor não encontrado.\\nAbra o arquivo iniciar.bat primeiro.'))
    .finally(()=>{{ if(btn){{ btn.disabled=false; btn.textContent='↺ Atualizar'; }} }});
}}

{donuts_js}
{init_js}
</script>
</body></html>"""


# ── Rodar grupo ───────────────────────────────────────────────────────────────

def _rodar_grupo(grupo, wms_path, erp_path):
    print(f"  WMS: {os.path.basename(wms_path)}")
    df_wms = carregar_wms(wms_path)
    print(f"  {len(df_wms)} linhas WMS")

    print(f"  ERP: {os.path.basename(erp_path)}")
    df_erp = carregar_erp(erp_path)
    print(f"  {len(df_erp)} linhas ERP (bruto)")

    # Filtro tipo NF
    tipo_col = next((c for c in df_erp.columns if c.lower().startswith("tipo") and "nf" in c.lower()), None)
    if tipo_col:
        antes = len(df_erp)
        df_erp = df_erp[df_erp[tipo_col].str.strip().str.upper() == ERP_FILTRO_TIPO_NF].copy()
        print(f"  {antes - len(df_erp)} linhas ERP removidas ({tipo_col} != {ERP_FILTRO_TIPO_NF})")

    # Filtro SKU por padrão
    padrao = ERP_SKU_PADRAO.get(grupo)
    if padrao:
        mask      = df_erp["sku"].str.match(padrao, na=False)
        ignorados = (~mask).sum()
        df_erp    = df_erp[mask].copy()
        if ignorados:
            print(f"  {ignorados} SKUs ignorados (fora do padrão '{padrao}')")

    print(f"  {len(df_erp)} linhas ERP (válidas)")

    subclientes = SUBCLIENTES.get(grupo)  # dict {prefixo: nome} ou None

    def _split_subgrupos(wms_d, erp_d):
        prefixos_conhecidos = set(subclientes.keys())
        # Agrupa prefixos por nome de cliente (ex: 28 e 29 → ANTARIS)
        cliente_prefixos: dict = {}
        for prefixo, nome in subclientes.items():
            cliente_prefixos.setdefault(nome, []).append(prefixo)
        sub = {}
        for nome, prefixos in cliente_prefixos.items():
            erp_sub = erp_d[erp_d["sku"].str[:2].isin(prefixos)].copy()
            wms_sub = wms_d[wms_d["sku"].str[:2].isin(prefixos)].copy()
            if not erp_sub.empty or not wms_sub.empty:
                sub[nome] = conciliar(wms_sub, erp_sub)
        erp_outros = erp_d[~erp_d["sku"].str[:2].isin(prefixos_conhecidos)].copy()
        wms_outros = wms_d[~wms_d["sku"].str[:2].isin(prefixos_conhecidos)].copy()
        if not erp_outros.empty:
            sub["OUTROS"] = conciliar(wms_outros, erp_outros)
        return sub

    # Agrupa por data do ERP (Emissao)
    datas_erp = sorted(df_erp["data"].dropna().dt.date.unique())
    if not datas_erp:
        r = conciliar(df_wms, df_erp)
        if subclientes:
            r["subgrupos"] = _split_subgrupos(df_wms, df_erp)
        print(f"  OK:{len(r['ok'])}  WMS:{len(r['so_wms'])}  ERP:{len(r['so_erp'])}  Div:{len(r['div_qtd'])}")
        return {"": r}

    resultados = {}
    for dt in datas_erp:
        erp_d    = df_erp[df_erp["data"].dt.date == dt].copy()
        pedidos  = set(erp_d["pedido"].dropna())
        wms_d    = df_wms[df_wms["pedido"].isin(pedidos)].copy()
        r = conciliar(wms_d, erp_d)
        if subclientes:
            r["subgrupos"] = _split_subgrupos(wms_d, erp_d)
        data_str = f"{dt.day:02d}.{dt.month:02d}"
        print(f"  [{data_str}] OK:{len(r['ok'])}  WMS:{len(r['so_wms'])}  ERP:{len(r['so_erp'])}  Div:{len(r['div_qtd'])}")
        resultados[data_str] = r

    return resultados


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    pares = detectar_pares(DADOS_DIR)
    if not pares:
        print(f"Nenhum par SAIDA WMS/ERP encontrado em '{DADOS_DIR}/'.")
        sys.exit(1)
    print(f"Pares detectados: {[g for g, _, _, _ in pares]}")

    todos = {}
    for grupo, _data_arq, wms_path, erp_path in pares:
        print(f"\n[{grupo}]")
        por_data = _rodar_grupo(grupo, wms_path, erp_path)
        for data_str, r in por_data.items():
            todos.setdefault(data_str, {})[grupo] = r

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    datas_str = " | ".join(d for d in sorted(todos) if d) or datetime.now().strftime("%d/%m/%Y")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    html = gerar_dashboard(todos, datas_str)
    saida_html = os.path.join(OUTPUT_DIR, f"saida_conciliacao_{timestamp}.html")
    with open(saida_html, "w", encoding="utf-8") as f:
        f.write(html)
    latest = os.path.join(OUTPUT_DIR, "saida_conciliacao_latest.html")
    with open(latest, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\nDashboard: {saida_html}")
    print(f"Latest:    {latest}")


if __name__ == "__main__":
    main()
