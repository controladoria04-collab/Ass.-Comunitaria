import streamlit as st
import pandas as pd
from io import BytesIO
import unicodedata
import re

# ============================
# CONFIG
# ============================

st.set_page_config(
    page_title="Conversor W4",
    layout="centered"
)

CENTROS_CUSTO_VALIDOS = {"DIACONIA", "FORTALEZA", "BRASIL", "EXTERIOR"}

# ============================
# FUNÇÕES AUXILIARES
# ============================

def clean_colname(name):
    return str(name).replace("\u00A0", " ").strip()


def make_unique_columns(cols):
    counts = {}
    out = []
    for c in [clean_colname(c) for c in cols]:
        counts[c] = counts.get(c, 0) + 1
        out.append(c if counts[c] == 1 else f"{c}__{counts[c]}")
    return out


def col(df, name):
    name = clean_colname(name)
    if name not in df.columns:
        raise ValueError(f"Coluna obrigatória não encontrada: {name}")
    s = df[name]
    return s.iloc[:, 0] if isinstance(s, pd.DataFrame) else s


def normalize_text(texto):
    texto = str(texto).lower().strip()
    texto = ''.join(
        c for c in unicodedata.normalize('NFKD', texto)
        if not unicodedata.combining(c)
    )
    texto = re.sub(r'[^a-z0-9]+', ' ', texto)
    return re.sub(r'\s+', ' ', texto).strip()


def formatar_data_coluna(serie):
    datas = pd.to_datetime(serie, errors="coerce")
    return datas.dt.strftime("%d/%m/%Y")


def converter_valor(valor, is_despesa):
    if pd.isna(valor):
        return ""
    base = f"{valor:.2f}".replace(".", ",") if isinstance(valor, (int, float)) else str(valor)
    base = base.lstrip("+- ").strip()
    return ("-" if is_despesa else "") + base


def extrair_centro(cliente):
    partes = [p.strip() for p in str(cliente).split(" - ")]
    return partes[-1] if partes and partes[-1].upper() in CENTROS_CUSTO_VALIDOS else ""


# ============================
# CATEGORIAS
# ============================

def preparar_categorias(df):
    col_desc = "Descrição da categoria financeira"

    def tirar_codigo(txt):
        p = str(txt).split(" ", 1)
        return p[1] if len(p) == 2 and p[0].isdigit() else txt

    df = df.copy()
    df["nome_base"] = df[col_desc].apply(tirar_codigo).apply(normalize_text)
    return df


# ============================
# MAPEAMENTO PREVIDÊNCIA
# ============================

def carregar_mapeamento(arq):
    if arq is None:
        raise ValueError("Envie o arquivo de mapeamento da Previdência")

    df = pd.read_excel(arq) if arq.name.endswith(("xlsx", "xls")) else pd.read_csv(arq, sep=";")
    df.columns = make_unique_columns(df.columns)

    regras = []
    for _, r in df.iterrows():
        regras.append((
            normalize_text(r["Padrao"]),
            r["Cliente"],
            extrair_centro(r["Cliente"])
        ))

    regras.sort(key=lambda x: len(x[0]), reverse=True)
    return regras


# ============================
# CONVERSÃO PRINCIPAL
# ============================

def converter_w4(df, df_cat, setor, regras_prev=None):

    df = df.copy()
    df.columns = make_unique_columns(df.columns)

    detalhe = col(df, "Detalhe Conta / Objeto")
    valor = col(df, "Valor total")
    data = col(df, "Data da Tesouraria")
    descricao = col(df, "Descrição")
    item_id = col(df, "Id Item tesouraria")

    df = df[~detalhe.astype(str).str.contains("Transferência Entre Disponíveis", case=False, na=False)]

    df["nome_base_w4"] = detalhe.apply(normalize_text)
    df = df.merge(df_cat[["nome_base", "Descrição da categoria financeira"]],
                  left_on="nome_base_w4",
                  right_on="nome_base",
                  how="left")

    df["Categoria_final"] = df["Descrição da categoria financeira"].fillna(detalhe)

    fluxo = df.get("Fluxo", "").astype(str).str.lower()
    processo = df.get("Processo", "").astype(str).str.lower()

    df["is_despesa"] = (
        fluxo.str.contains("despesa") |
        processo.str.contains("pagamento") |
        detalhe.str.lower().str.contains("despesa|custo")
    )

    df["Valor_final"] = [
        converter_valor(v, d) for v, d in zip(valor, df["is_despesa"])
    ]

    df["ClienteFornecedor"] = ""
    df["CentroCusto"] = ""

    if setor == "Previdência Brasil":
        detalhe_norm = detalhe.apply(normalize_text)

        def match(txt):
            for padrao, cliente, centro in regras_prev:
                if padrao in txt:
                    return cliente, centro
            return "", ""

        res = detalhe_norm.apply(match)
        df["ClienteFornecedor"] = res.apply(lambda x: x[0])
        df["CentroCusto"] = res.apply(lambda x: x[1])

        mask = df["ClienteFornecedor"] != ""
        df.loc[mask, "Categoria_final"] = "11318 - Repasse Recebido Fundo de Previdência"

    datas = formatar_data_coluna(data)

    out = pd.DataFrame({
        "Data de Competência": datas,
        "Data de Vencimento": datas,
        "Data de Pagamento": datas,
        "Valor": df["Valor_final"],
        "Categoria": df["Categoria_final"],
        "Descrição": item_id.astype(str) + " " + descricao.astype(str),
        "Cliente/Fornecedor": df["ClienteFornecedor"],
        "CNPJ/CPF Cliente/Fornecedor": "",
        "Centro de Custo": df["CentroCusto"],
        "Observações": ""
    })

    return out


# ============================
# INTERFACE
# ============================

st.title("Conversor W4")

setor = st.selectbox(
    "Selecione o setor",
    ["Ass. Comunitária", "Sinodalidade", "Previdência Brasil"]
)

arq_map = None
if setor == "Previdência Brasil":
    arq_map = st.file_uploader("Upload do mapeamento (Cliente / Padrao)", type=["xlsx", "xls", "csv"])

arq_w4 = st.file_uploader("Upload do arquivo W4", type=["xlsx", "xls", "csv"])

if arq_w4:
    if st.button("Converter"):
        try:
            df_w4 = pd.read_excel(arq_w4) if arq_w4.name.endswith(("xlsx", "xls")) else pd.read_csv(arq_w4, sep=";")
            df_w4.columns = make_unique_columns(df_w4.columns)

            df_cat = pd.read_excel("categorias_contabeis.xlsx")
            df_cat.columns = make_unique_columns(df_cat.columns)
            df_cat_prep = preparar_categorias(df_cat)

            regras = carregar_mapeamento(arq_map) if setor == "Previdência Brasil" else None

            df_final = converter_w4(df_w4, df_cat_prep, setor, regras)

            buffer = BytesIO()
            df_final.to_excel(buffer, index=False)
            buffer.seek(0)

            st.download_button(
                "Baixar arquivo convertido",
                buffer,
                "conta_azul_convertido.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.success("Conversão concluída!")

        except Exception as e:
            st.error(f"Erro: {e}")
