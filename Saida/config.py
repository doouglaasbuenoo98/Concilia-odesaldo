# Mapeamento de colunas — Conciliação de SAÍDA

WMS_COLUNAS = {
    "pedido":     "DCSDOCUMENTO",
    "sku":        "PRDCODIGO",
    "quantidade": "DSIQUANTIDADESEPARADA",
    "unidade":    "UNICODIGO",
    "data":       "DATASEPARACAO",
    "peso":       "SEPPESOSEPARADO",
}

ERP_COLUNAS = {
    "pedido":     "No do Pedido",
    "numero_nf":  "Num. Docto.",
    "sku":        "Produto",
    "quantidade": "Quantidade",
    "unidade":    "Unidade",
    "data":       "Emissao",
    "vlr_unit":   "Vlr.Unitario",
}

# Chave de conciliação: pedido (ordem de separação) + SKU
CHAVE_CONCILIACAO = ["pedido", "sku"]

# Filtro de tipo de NF no ERP (coluna "Tipo de N.F." = "N" para NF normal)
ERP_FILTRO_TIPO_NF = "N"

# Padrão de SKU válido por grupo (vazio = sem filtro, usa pedido como filtro)
ERP_SKU_PADRAO = {
    "LIV UP": r"^[A-Za-z]\d{3,4}$",
}

# Subclientes por grupo: {grupo: {prefixo_sku: nome_cliente}}
# SKUs que não baterem nenhum prefixo são agrupados em "OUTROS"
SUBCLIENTES = {
    "NEXXA": {
        "22": "GNÓS",
        "28": "ANTARIS",
        "29": "ANTARIS",
        "31": "BOALI",
        "34": "MILKY MOO",
        "35": "HALIPAR",
        "36": "MANIA DE CHURRASCO",
        "37": "BURGUÊS",
    },
}
