import streamlit as st
import pandas as pd
from io import BytesIO
import unicodedata
import re

# ============================
# CONFIG DO APP
# ============================
st.set_page_config(
    page_title="Conversão Planilha Setores",
    layout="centered"
)

# ============================
# FUNÇÕES AUXILIARES
# ============================
def normalize_text(texto):
    texto = str(texto).lower().strip()
    texto = "".join(
        c for c in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(c)
    )
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


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
# FUNÇÃO PRINCIPAL
# ============================
def converter_w4(df_w4, df_categorias_prep, setor, df_map_prev):
    if "Detalhe Conta / Objeto" not in df_w4.columns:
        raise ValueError("Coluna 'Detalhe Conta / Objeto' não existe no W4.")

    col_cat = "Detalhe Conta / Objeto"

    # >>> ALTERAÇÃO 2 — SUBSTITUIÇÃO FIXA <<<
    df_w4[col_cat] = df_w4[col_cat].replace(
        "Despesa com Repasse para Economato Geral - Encargos Folha de Pagamento",
        "13089 - Desp. com Rep. Eco. Geral - Encargos Folha"
    )

    # Remover transferências
    df = df_w4.loc[
        ~df_w4[col_cat].astype(str).str.contains(
            "Transferência Entre Disponíveis",
            case=False,
            na=False
        )
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
        how="left"
    )

    df["Categoria_final"] = df[col_desc_cat].where(
        df[col_desc_cat].notna(),
        df[col_cat]
    )

    df["Cliente/Fornecedor"] = ""
    df["Centro de Custo"] = ""

    # ============================
    # PREVIDÊNCIA BRASIL — REPASSES
    # ============================
    if setor == "Previdência Brasil":
        detalhe_norm = df[col_cat].astype(str).apply(normalize_text)

        for _, row in df_map_prev.iterrows():
            padrao = row["Padrao_norm"]
            cliente_raw = row["Cliente"]

            mask = detalhe_norm.str.contains(padrao, na=False)

            if mask.any():
                df.loc[mask, "Categoria_final"] = (
                    "11318 - Repasse Recebido Fundo de Previdência"
                )
                df.loc[mask, "Cliente/Fornecedor"] = cliente_raw

                if "-" in cliente_raw:
                    centro = cliente_raw.rsplit("-", 1)[1].strip().upper()
                    df.loc[mask, "Centro de Custo"] = centro

    # ============================
    # PROCESSO / EMPRÉSTIMOS
    # ============================
    fluxo = df.get("Fluxo", pd.Series("", index=df.index)).astype(str).str.lower()
    fluxo_vazio = fluxo.str.strip().isin(["", "nan", "none"])

    cond_fluxo_receita = fluxo.str.contains("receita", na=False)
    cond_fluxo_despesa = fluxo.str.contains("despesa", na=False)
    cond_imobilizado = fluxo.str.contains("imobilizado", na=False)

    proc_original = df.get("Processo", pd.Series("", index=df.index)).astype(str)
    proc = proc_original.str.lower()
    proc = proc.apply(
        lambda t: unicodedata.normalize("NFKD", t)
        .encode("ascii", "ignore")
        .decode("ascii")
    )

    pessoa = df.get("Pessoa", pd.Series("", index=df.index)).astype(str)

    cond_emprestimo = proc.str.contains("emprestimo", na=False)
    cond_pag_emp = cond_emprestimo & proc.str.contains("pagamento", na=False)
    cond_rec_emp = cond_emprestimo & proc.str.contains("recebimento", na=False)

    df.loc[cond_emprestimo, "Categoria_final"] = (
        proc_original[cond_emprestimo] + " " + pessoa[cond_emprestimo]
    )

    # ============================
    # CLASSIFICAÇÃO DESPESA / RECEITA
    # ============================
    detalhe_lower = df[col_cat].astype(str).str.lower()

    cond_palavra_despesa = (
        fluxo_vazio
        & (
            detalhe_lower.str.contains("custo", na=False)
            | detalhe_lower.str.contains("despesa", na=False)
        )
    )

    cond_pag_proc = fluxo_vazio & proc.str.contains("pagamento", na=False)
    cond_rec_proc = fluxo_vazio & proc.str.contains("recebimento", na=False)

    df["is_despesa"] = (
        cond_fluxo_despesa
        | cond_imobilizado
        | cond_palavra_despesa
        | cond_pag_proc
    )

    df.loc[cond_fluxo_receita | cond_rec_proc, "is_despesa"] = False

    # ============================
    # VALORES
    # ============================
    df["Valor_str_final"] = [
        converter_valor(v, d)
        for v, d in zip(df["Valor total"], df["is_despesa"])
    ]

    # ============================
    # DATAS
    # ============================
    data_tes = formatar_data_coluna(df["Data da Tesouraria"])

    # ============================
    # CENTRO DE CUSTO — SINODALIDADE
    # ============================
    if setor == "Sinodalidade" and "Lote" in df.columns:
        centro_padrao = (
            df["Lote"]
            .fillna("")
            .astype(str)
            .str.strip()
            .replace(["", "nan", "NaN"], "Adm Financeiro")
        )
    else:
        centro_padrao = ""

    # ============================
    # MONTAGEM FINAL
    # ============================
    out = pd.DataFrame()
    out["Data de Competência"] = data_tes
    out["Data de Vencimento"] = data_tes
    out["Data de Pagamento"] = data_tes
    out["Valor"] = df["Valor_str_final"]
    out["Categoria"] = df["Categoria_final"]

    if "Id Item tesouraria" in df.columns:
        out["Descrição"] = (
            df["Id Item tesouraria"].astype(str)
            + " "
            + df["Descrição"].astype(str)
        )
    else:
        out["Descrição"] = df["Descrição"]

    out["Cliente/Fornecedor"] = df["Cliente/Fornecedor"]
    out["CNPJ/CPF Cliente/Fornecedor"] = ""
    out["Centro de Custo"] = df["Centro de Custo"].replace("", centro_padrao)
    out["Observações"] = ""

    return out


# ============================
# NOME DO ARQUIVO
# ============================
MAPA_MESES = {
    "01": "jan", "02": "fev", "03": "marc", "04": "abr",
    "05": "mai", "06": "jun", "07": "jul", "08": "ago",
    "09": "set", "10": "out", "11": "nov", "12": "dez"
}

def gerar_nome_arquivo(df):
    empresa = str(df["Empresa"].iloc[0]).strip()
    disponivel = str(df["Disponível"].iloc[0]).strip()
    datas = pd.to_datetime(df["Data da Tesouraria"], errors="coerce")
    mes = datas.dt.strftime("%m").dropna().iloc[0]
    return f"{empresa} {disponivel} {MAPA_MESES.get(mes, mes)}.xlsx"


# ============================
# CARREGAMENTOS
# ============================
def carregar_arquivo_w4(arq):
    if arq.name.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(arq)
    return pd.read_csv(arq, sep=";", encoding="latin1")


df_cat_raw = pd.read_excel("categorias_contabeis.xlsx")
df_cat_prep = preparar_categorias(df_cat_raw)

df_map_prev = pd.read_excel("mapeamento_previdencia.xlsx")
df_map_prev["Padrao_norm"] = df_map_prev["Padrao"].apply(normalize_text)

# ============================
# INTERFACE
# ============================
st.title("Conversão Planilha Setores")

setor = st.selectbox(
    "Selecione o setor",
    ["Ass. Comunitária", "Sinodalidade", "Previdência Brasil"]
)

arq_w4 = st.file_uploader(
    "Envie o arquivo W4 (CSV ou Excel)",
    type=["csv", "xlsx", "xls"]
)

if arq_w4:
    if st.button("Converter planilha"):
        try:
            df_w4 = carregar_arquivo_w4(arq_w4)
            df_final = converter_w4(df_w4, df_cat_prep, setor, df_map_prev)

            buffer = BytesIO()
            df_final.to_excel(buffer, index=False, engine="openpyxl")
            buffer.seek(0)

            st.download_button(
                label="Baixar planilha convertida",
                data=buffer,
                file_name=gerar_nome_arquivo(df_w4),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erro: {e}")
else:
    st.info("Selecione um setor e envie o arquivo W4.")
