# Mapeamento de colunas dos arquivos WMS e ERP
# Chave interna -> nome real da coluna no arquivo

WMS_COLUNAS = {
    "numero_nf":  "N.F.",
    "sku":        "PRODUTO",
    "quantidade": "RECEBIDO",
    "peso":       "PESO LIQ",
    "data":       "DATA INCLUS\xc3O",   # DATA INCLUSÃO (encoding do arquivo)
    "descricao":  "DESCRI\xc3\xc7\xc3O",
    "status":     "STATUS",
}

ERP_COLUNAS = {
    "numero_nf":  "Documento",
    "sku":        "Produto",
    "quantidade": "Quantidade",
    "unidade":    "Unidade",
    "data":       "DT Digitacao",
    "descricao":  "Descricao",
    "fornecedor": "Descricao.1",
    "vlr_unit":   "Vlr.Unitario",
}

# Tolerância de diferença de datas em dias (0 = datas devem ser iguais)
TOLERANCIA_DIAS = 1

# Colunas que formam a chave de conciliação
CHAVE_CONCILIACAO = ["numero_nf", "sku"]

# Filtrar WMS apenas por recebimentos concluídos
WMS_FILTRO_STATUS = ["Recebido"]

# Padrão de SKU válido por grupo (regex).
# SKUs do ERP que não baterem com o padrão do grupo serão ignorados no cálculo.
ERP_SKU_PADRAO = {
    "DRUMATTOS": r"^D\d{7}$",        # D + 7 dígitos  ex: D2700054
    "LIV UP":    r"^[A-Za-z]\d{3}$", # 1 letra + 3 dígitos  ex: P162, S015
    "NEXXA":     r"^\d{7}$",          # exatamente 7 dígitos  ex: 2200343
}
