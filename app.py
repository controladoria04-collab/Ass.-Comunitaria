import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import unicodedata
import re

# ============================
# CONFIG DO APP
# ============================

st.set_page_config(page_title="Conversão Planilha Setores", layout="centered")

# ============================
# FUNÇÕES AUXILIARES
# ============================

CENTROS_CUSTO_VALIDOS = {"FORTALEZA", "DIACONIA", "EXTERIOR", "BRASIL"}


def normalize_text(texto):
    texto = str(texto).lower().strip()
    texto = "".join(
        c for c in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(c)
    )
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def make_unique_columns(cols):
    """
    - strip
    - remove NBSP
    - torna nomes únicos: Col, Col__2, Col__3...
    """
    cleaned = []
    for c in cols:
        c = str(c).replace("\u00A0", " ").strip()
        cleaned.append(c)

    counts = {}
    out = []
    for c in cleaned:
        counts[c] = counts.get(c, 0) + 1
        if counts[c] == 1:
            out.append(c)
        else:
            out.append(f"{c}__{counts[c]}")
    return out


def extrair_centro(cliente):
    partes = [p.strip() for p in str(cliente).split(" - ")]
    if partes and partes[-1].upper() in CENTROS_CUSTO_VALIDOS:
        return partes[-1].upper()
    return ""


def preparar_categorias(df_cat):
    col = "Descrição da categoria financeira"
    df = df_cat.copy()

    def tirar_codigo(txt):
        txt = str(txt).strip()
        partes = txt.split(" ", 1)
        if len(partes) == 2 and any(ch.isdigit() for ch in partes[0]):
            return partes[1].strip()
        return txt

    df["nome_base"] = df[col].apply(tirar_codigo).apply(normalize_text)
    return df


def formatar_data_coluna(serie):
    datas = pd.to_datetime(serie, errors="coerce")
    return datas.dt.strftime("%d/%m/%Y")


def converter_valor(valor_str, is_despesa):
    if pd.isna(valor_str):
        return ""

    if isinstance(valor_str, (int, float)):
        base = f"{valor_str:.2f}".replace(".", ",")
    else:
        base = str(valor_str).strip()

    base_sem_sinal = base.lstrip("+- ").strip()
    return ("-" if is_despesa else "") + base_sem_sinal


# ============================
# LEITURA DE ARQUIVOS
# ============================

def carregar_arquivo_w4(arq):
    if arq.name.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(arq)
    else:
        df = pd.read_csv(arq, sep=";", encoding="latin1")

    df.columns = make_unique_columns(df.columns)
    return df


def carregar_mapeamento_upload(arq_map):
    """
    Arquivo Excel/CSV com colunas:
      - Cliente
      - Padrao
    """
    if arq_map.name.lower().endswith((".xlsx", ".xls")):
        dfm = pd.read_excel(arq_map)
    else:
        dfm = pd.read_csv(arq_map, sep=";", encoding="latin1")

    dfm.columns = make_unique_columns(dfm.columns)

    if "Cliente" not in dfm.columns or "Padrao" not in dfm.columns:
        raise ValueError("Arquivo de mapeamento precisa ter colunas: Cliente e Padrao")

    dfm["Cliente"] = dfm["Cliente"].astype(str).str.strip()
    dfm["Padrao"] = dfm["Padrao"].astype(str).str.strip()

    dfm = dfm[(dfm["Cliente"] != "") & (dfm["Padrao"] != "")]
    if dfm.empty:
        raise ValueError("Arquivo de mapeamento está vazio ou sem dados válidos.")

    # Prepara estruturas normalizadas (padrões maiores primeiro)
    regras = []
    for cliente, padrao in zip(dfm["Cliente"], dfm["Padrao"]):
        regras.append(
            (
                normalize_text(padrao),
                str(cliente).strip(),
                extrair_centro(cliente)
            )
        )

    regras.sort(key=lambda x: len(x[0]), reverse=True)
    return regras


# ============================
# FUNÇÃO PRINCIPAL
# ============================

def converter_w4(df_w4, df_categorias_prep, setor, regras_previdencia, debug=False):

    cols = set(df_w4.columns)

    # Colunas obrigatórias (ajuste aqui caso seu W4 tenha nomes diferentes)
    if "Detalhe Conta / Objeto" not in cols:
        raise ValueError(f"Coluna 'Detalhe Conta / Objeto' não existe no W4. Colunas: {list(df_w4.columns)}")
    if "Valor total" not in cols:
        raise ValueError(f"Coluna 'Valor total' não existe no W4. Colunas: {list(df_w4.columns)}")
    if "Data da Tesouraria" not in cols:
        raise ValueError(f"Coluna 'Data da Tesouraria' não existe no W4. Colunas: {list(df_w4.columns)}")

    col_cat = "Detalhe Conta / Objeto"

    # Remover transferências
    df = df_w4.loc[
        ~df_w4[col_cat].astype(str).str.contains("Transferência Entre Disponíveis", case=False, na=False)
    ].copy()

    # ============================
    # CATEGORIAS BASE
    # ============================

    col_desc_cat = "Descrição da categoria financeira"

    df["nome_base_w4"] = df[col_cat].astype(str).apply(normalize_text)

    df = df.merge(
        df_categorias_prep[["nome_base", col_desc_cat]],
        left_on="nome_base_w4",
        right_on="nome_base",
        how="left",
        suffixes=("", "__cat")
    )

    df["Categoria_final"] = df[col_desc_cat].where(df[col_desc_cat].notna(), df[col_cat].astype(str))

    # ============================
    # PREVIDÊNCIA: aplica regras do arquivo (somente para Previdência Brasil)
    # ============================

    df["ClienteFornecedor_final"] = ""
    df["CentroCusto_final"] = ""

    if str(setor).strip() == "Previdência Brasil":
        detalhe_norm = df[col_cat].astype(str).apply(normalize_text)

        def buscar(txt_norm):
            for padrao_norm, cliente, centro in regras_previdencia:
                if padrao_norm and padrao_norm in txt_norm:
                    return cliente, centro
            return "", ""

        pares = detalhe_norm.apply(buscar)
        df["ClienteFornecedor_final"] = pares.apply(lambda x: x[0])
        df["CentroCusto_final"] = pares.apply(lambda x: x[1])

        achou = df["ClienteFornecedor_final"].ne("")
        df["Categoria_final"] = np.where(
            achou.to_numpy(),
            "11318 - Repasse Recebido Fundo de Previdência",
            df["Categoria_final"].astype(str).to_numpy()
        )

        if debug:
            st.write("DEBUG: regras carregadas:", len(regras_previdencia))
            st.write("DEBUG: linhas com cliente encontrado:", int(achou.sum()))
            st.write(df.loc[achou, [col_cat, "ClienteFornecedor_final", "CentroCusto_final"]].head(30))

    # ============================
    # DESPESA / RECEITA
    # ============================

    fluxo = df.get("Fluxo", pd.Series("", index=df.index)).astype(str).str.lower()
    fluxo_vazio = fluxo.str.strip().isin(["", "nan", "none"])

    cond_fluxo_receita = fluxo.str.contains("receita", na=False)
    cond_fluxo_despesa = fluxo.str.contains("despesa", na=False)
    cond_imobilizado = fluxo.str.contains("imobilizado", na=False)

    detalhe_lower = df[col_cat].astype(str).str.lower()
    cond_palavra_despesa = fluxo_vazio & (
        detalhe_lower.str.contains("custo", na=False) |
        detalhe_lower.str.contains("despesa", na=False)
    )

    proc_original = df.get("Processo", pd.Series("", index=df.index)).astype(str)
    proc = proc_original.str.lower().apply(
        lambda t: unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode("ascii")
    )

    cond_pag_proc = fluxo_vazio & proc.str.contains("pagamento", na=False)
    cond_rec_proc = fluxo_vazio & proc.str.contains("recebimento", na=False)

    is_despesa = (cond_fluxo_despesa | cond_imobilizado | cond_palavra_despesa | cond_pag_proc).to_numpy()
    is_despesa = np.where((cond_fluxo_receita | cond_rec_proc).to_numpy(), False, is_despesa)
    df["is_despesa"] = is_despesa

    # ============================
    # VALORES
    # ============================

    df["Valor_str_final"] = [
        converter_valor(v, d) for v, d in zip(df["Valor total"], df["is_despesa"])
    ]

    # ============================
    # DATAS
    # ============================

    data_tes = formatar_data_coluna(df["Data da Tesouraria"])

    # ============================
    # SAÍDA
    # ============================

    if "Descrição" not in df.columns:
        df["Descrição"] = ""

    out = pd.DataFrame()
    out["Data de Competência"] = data_tes
    out["Data de Vencimento"] = data_tes
    out["Data de Pagamento"] = data_tes
    out["Valor"] = df["Valor_str_final"]
    out["Categoria"] = df["Categoria_final"]

    if "Id Item tesouraria" in df.columns:
        out["Descrição"] = df["Id Item tesouraria"].astype(str) + " " + df["Descrição"].astype(str)
    else:
        out["Descrição"] = df["Descrição"].astype(str)

    out["Cliente/Fornecedor"] = df["ClienteFornecedor_final"]
    out["CNPJ/CPF Cliente/Fornecedor"] = ""

    if str(setor).strip() == "Previdência Brasil":
        out["Centro de Custo"] = df["CentroCusto_final"]
    elif str(setor).strip() == "Sinodalidade" and "Lote" in df.columns:
        centro = df["Lote"].fillna("").astype(str).str.strip()
        centro = centro.replace(["", "nan", "NaN"], "Adm Financeiro")
        out["Centro de Custo"] = centro
    else:
        out["Centro de Custo"] = ""

    out["Observações"] = ""

    return out


# ============================
# CARREGAR CATEGORIAS
# ============================

@st.cache_data
def carregar_categorias():
    df_cat_raw = pd.read_excel("categorias_contabeis.xlsx")
    df_cat_raw.columns = make_unique_columns(df_cat_raw.columns)
    return preparar_categorias(df_cat_raw)

try:
    df_cat_prep = carregar_categorias()
except Exception as e:
    st.error(f"Erro ao carregar 'categorias_contabeis.xlsx': {e}")
    st.stop()


# ============================
# INTERFACE
# ============================

st.title("Conversão Planilha Setores")

setor = st.selectbox(
    "Selecione o setor",
    ["Ass. Comunitária", "Sinodalidade", "Previdência Brasil"]
)

st.markdown("### Mapeamento Previdência (OBRIGATÓRIO)")
arq_map = st.file_uploader(
    "Envie Excel/CSV com colunas: Cliente e Padrao",
    type=["csv", "xlsx", "xls"],
    key="map"
)

debug = st.checkbox("DEBUG", value=False)

st.markdown("### Arquivo W4")
arq_w4 = st.file_uploader(
    "Envie o arquivo W4 (CSV ou Excel)",
    type=["csv", "xlsx", "xls"],
    key="w4"
)

if arq_w4 and arq_map:
    if st.button("Converter planilha"):
        try:
            regras_previdencia = carregar_mapeamento_upload(arq_map)

            df_w4 = carregar_arquivo_w4(arq_w4)

            if debug:
                st.write("DEBUG: colunas do W4 (já saneadas):")
                st.write(df_w4.columns.tolist())

            df_final = converter_w4(df_w4, df_cat_prep, setor, regras_previdencia, debug=debug)

            st.success("Planilha convertida com sucesso!")

            buffer = BytesIO()
            df_final.to_excel(buffer, index=False, engine="openpyxl")
            buffer.seek(0)

            st.download_button(
                label="Baixar planilha convertida",
                data=buffer,
                file_name="planilha_convertida_setor.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erro: {e}")

elif arq_w4 and not arq_map:
    st.warning("Para Previdência, envie também o arquivo de mapeamento (Cliente e Padrao).")
else:
    st.info("Selecione um setor e envie o arquivo W4 e o mapeamento.")
