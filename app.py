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
    texto = ''.join(
        c for c in unicodedata.normalize('NFKD', texto)
        if not unicodedata.combining(c)
    )
    texto = re.sub(r'[^a-z0-9]+', ' ', texto)
    return re.sub(r'\s+', ' ', texto).strip()


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
# MAPEAMENTO CLIENTES - PREVIDÊNCIA
# ============================

CLIENTES_PREVIDENCIA = [
    "APARECIDA - BRASIL",
    "ARACAJU - BRASIL",
    "ARACATI - BRASIL",
    "ARAPIRACA - BRASIL",
    "ARARAQUARA - BRASIL",
    "ASSESSORIA DE DIFUSÃO DA OBRA - DIACONIA",
    "ASSESSORIA DE PROMOÇÃO HUMANA - DIACONIA",
    "ASSESSORIA JOVEM - DIACONIA",
    "ASSESSORIA LITURGICO - SACRAMENTAL - DIACONIA",
    "ASSESSORIA VOCACIONAL - DIACONIA",
    "ASSISTÊNCIA APOSTÓLICA - DIACONIA",
    "ASSISTÊNCIA COMUNITÁRIA - DIACONIA",
    "ASSISTÊNCIA DE COMUNICAÇÃO - DIACONIA",
    "ASSISTÊNCIA DE FORMAÇÃO - DIACONIA",
    "ASSISTÊNCIA LOCAL - FORTALEZA",
    "ASSISTÊNCIA MISSIONÁRIA - DIACONIA",
    "BEJAIA - EXTERIOR",
    "BELÉM - BRASIL",
    "BELO HORIZONTE - BRASIL",
    "BOGOTÁ - EXTERIOR",
    "BOSTON - EXTERIOR",
    "BRASÍLIA - BRASIL",
    "CABO VERDE - EXTERIOR",
    "CAMPINA GRANDE - BRASIL",
    "CAMPO GRANDE - BRASIL",
    "CASA DE RETIRO - DIACONIA",
    "CEST - CASA CONTEMPLATIVA - BRASIL",
    "CEST - FORTALEZA",
    "CEV - AQUIRAZ - FORTALEZA",
    "CEV - CARMO - FORTALEZA",
    "CEV - SHALOM CID. DOS FUNCIONÁRIOS - FORTALEZA",
    "CEV - SHALOM CRISTO REDENTOR - FORTALEZA",
    "CEV - SHALOM FÁTIMA - FORTALEZA",
    "CEV - SHALOM PARANGABA - FORTALEZA",
    "CEV - SHALOM PARQUELÂNDIA - FORTALEZA",
    "CEV - SHALOM PAZ - FORTALEZA",
    "CHAVES - BRASIL",
    "COLÉGIO SHALOM - FORTALEZA",
    "COORDENAÇÃO APOSTÓLICA - FORTALEZA",
    "CRATEÚS - BRASIL",
    "CRISMA - FORTALEZA",
    "CRUZEIRO DO SUL - BRASIL",
    "CUIABÁ - BRASIL",
    "CURITIBA - BRASIL",
    "DISCIPULADO EUSÉBIO - FORTALEZA",
    "DISCIPULADO PACAJUS - FORTALEZA",
    "DISCIPULADO QUIXADÁ - FORTALEZA",
    "ECONOMATO GERAL - DIACONIA",
    "EDIÇÕES - DIACONIA",
    "ESCOLA DE EVANGELIZAÇÃO - FORTALEZA",
    "ESCRITÓRIO GERAL - DIACONIA",
    "ESC. SECRETÁRIA COMUNITÁRIA - FORTALEZA",
    "FLORIANÓPOLIS - BRASIL",
    "GARANHUNS - BRASIL",
    "GOIÂNIA - BRASIL",
    "GUARULHOS - BRASIL",
    "GUIANA FRANCESA - EXTERIOR",
    "HAIFA - EXTERIOR",
    "IGREJA - DIACONIA",
    "IMPERATRIZ - BRASIL",
    "ITAPIPOCA - BRASIL",
    "JOÃO PESSOA - BRASIL",
    "JOINVILLE - BRASIL",
    "JUAZEIRO DA BAHIA - BRASIL",
    "JUAZEIRO DO NORTE - BRASIL",
    "JUIZ DE FORA - BRASIL",
    "LANÇAI AS REDES - DIACONIA",
    "LANCHONETE - CEV FÁTIMA - FORTALEZA",
    "LANCHONETE PARQUELÂNDIA - FORTALEZA",
    "LANCHONETE - SHALOM DA PAZ - FORTALEZA",
    "LIMA - EXTERIOR",
    "LIVRARIA - CEV FÁTIMA - FORTALEZA",
    "LIVRARIA - PARANGABA - FORTALEZA",
    "LUBANGO - EXTERIOR",
    "MACAPÁ - BRASIL",
    "MACEIÓ - BRASIL",
    "MADAGASCAR - EXTERIOR",
    "MANAUS - BRASIL",
    "MANILA - EXTERIOR",
    "MATRIZ - FORTALEZA",
    "MOÇAMBIQUE - EXTERIOR",
    "MOSSORÓ - BRASIL",
    "NATAL - BRASIL",
    "NAZARETH - EXTERIOR",
    "NITERÓI - BRASIL",
    "PALMAS - BRASIL",
    "PARNAÍBA - BRASIL",
    "PARRESIA - DIACONIA",
    "PATOS - BRASIL",
    "PATOS DE MINAS - BRASIL",
    "PH - CASA RENATA COURAS - FORTALEZA",
    "PH - CASA RONALDO PEREIRA - FORTALEZA",
    "PH - PROJETO AMIGO DOS POBRES - FORTALEZA",
    "PH - PROJETO MARIA MADALENA - FORTALEZA",
    "PH - SECRETARIA - FORTALEZA",
    "PIRACICABA - BRASIL",
    "PONTA GROSSA - BRASIL",
    "PREFEITURA - DIACONIA",
    "PROJETO ARTES - FORTALEZA",
    "PROJETO JUVENTUDE - FORTALEZA",
    "QUIXADÁ - BRASIL",
    "RÁDIO SHALOM - FORTALEZA",
    "RECIFE - BRASIL",
    "REG - CID. DOS FUNCIONÁRIOS - FORTALEZA",
    "REG.  PACAJUS - FORTALEZA",
    "REG - PARANGABA - FORTALEZA",
    "REG - PARQUELÂNDIA - FORTALEZA",
    "RIO DE JANEIRO - BRASIL",
    "ROMA - EXTERIOR",
    "SALVADOR - BRASIL",
    "SANTA CRUZ DE LA SIERRA - EXTERIOR",
    "SANTO AMARO - BRASIL",
    "SANTO ANDRÉ - BRASIL",
    "SÃO LEOPOLDO - BRASIL",
    "SÃO LUÍS - BRASIL",
    "SÃO PAULO - PERDIZES - BRASIL",
    "SÃO PAULO - TÁIPAS - BRASIL",
    "SEC. DE COMUNICAÇÃO - FORTALEZA",
    "SEC. DE SACERDOTES E SEMINARISTAS - DIACONIA",
    "SECRETARIA DE PLANEJAMENTO - DIACONIA",
    "SECRETARIA VOCACIONAL - FORTALEZA",
    "SENHOR DO BONFIM - BRASIL",
    "SETOR COLEGIADO - DIACONIA",
    "SETOR DE EVENTOS - FORTALEZA",
    "SETOR DOS CELIBATÁRIOS - DIACONIA",
    "SETOR FAMÍLIA - DIACONIA",
    "SOBRAL - BRASIL",
    "TAIWAN - EXTERIOR",
    "TECNOLOGIA DA INFORMAÇÃO - DIACONIA",
    "TERESINA - BRASIL",
    "TUNÍSIA - EXTERIOR",
    "UBERABA - BRASIL",
    "VITÓRIA - BRASIL",
    "VITÓRIA DA CONQUISTA - BRASIL",
]

# chave normalizada -> valor oficial
CLIENTE_MAP_PREV = {normalize_text(x): x for x in CLIENTES_PREVIDENCIA}


# ============================
# FUNÇÃO PRINCIPAL
# ============================

def converter_w4(df_w4, df_categorias_prep, setor):

    if "Detalhe Conta / Objeto" not in df_w4.columns:
        raise ValueError("Coluna 'Detalhe Conta / Objeto' não existe no W4.")

    col_cat = "Detalhe Conta / Objeto"

    # Remover transferências
    df = df_w4.loc[
        ~df_w4[col_cat].astype(str).str.contains(
            "Transferência Entre Disponíveis", case=False, na=False
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

    # ============================
    # REGRA PREVIDÊNCIA: REPASSE FUNDO DE PREVIDÊNCIA
    # ============================
    df["ClienteFornecedor_final"] = ""

    if setor in ["Previdência Brasil", "Previdência"]:
        base_txt = "Repasse Recebido Fundo de Previdência"

        mask_repasse = df[col_cat].astype(str).str.contains(
            base_txt, case=False, na=False
        )

        if mask_repasse.any():
            complemento = (
                df.loc[mask_repasse, col_cat]
                .astype(str)
                .str.replace(base_txt, "", case=False, regex=False)
                .str.strip()
            )

            complemento_norm = complemento.apply(normalize_text)
            cliente_oficial = complemento_norm.map(CLIENTE_MAP_PREV)

            # Só aplica quando achou cliente na lista
            mask_achou = mask_repasse & cliente_oficial.notna()
            df.loc[mask_achou, "Categoria_final"] = "11318 - Repasse Recebido Fundo de Previdência"
            df.loc[mask_achou, "ClienteFornecedor_final"] = cliente_oficial[mask_achou]

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

    # Categoria para empréstimos
    df.loc[cond_pag_emp, "Categoria_final"] = (
        proc_original[cond_pag_emp] + " " + pessoa[cond_pag_emp]
    )

    df.loc[cond_rec_emp, "Categoria_final"] = (
        proc_original[cond_rec_emp] + " " + pessoa[cond_rec_emp]
    )

    df.loc[
        cond_emprestimo & ~cond_pag_emp & ~cond_rec_emp,
        "Categoria_final"
    ] = (
        proc_original[cond_emprestimo & ~cond_pag_emp & ~cond_rec_emp]
        + " "
        + pessoa[cond_emprestimo & ~cond_pag_emp & ~cond_rec_emp]
    )

    # ============================
    # CLASSIFICAÇÃO DESPESA / RECEITA
    # ============================

    detalhe_lower = df[col_cat].astype(str).str.lower()

    cond_palavra_despesa = (
        fluxo_vazio &
        (
            detalhe_lower.str.contains("custo", na=False) |
            detalhe_lower.str.contains("despesa", na=False)
        )
    )

    cond_pag_proc = fluxo_vazio & proc.str.contains("pagamento", na=False)
    cond_rec_proc = fluxo_vazio & proc.str.contains("recebimento", na=False)

    df["is_despesa"] = (
        cond_fluxo_despesa |
        cond_imobilizado |
        cond_palavra_despesa |
        cond_pag_proc
    )

    df.loc[
        cond_fluxo_receita | cond_rec_proc,
        "is_despesa"
    ] = False

    # ============================
    # VALORES
    # ============================

    if "Valor total" not in df.columns:
        raise ValueError("Coluna 'Valor total' não existe no W4.")

    df["Valor_str_final"] = [
        converter_valor(v, d)
        for v, d in zip(df["Valor total"], df["is_despesa"])
    ]

    # ============================
    # DATAS
    # ============================

    if "Data da Tesouraria" not in df.columns:
        raise ValueError("Coluna 'Data da Tesouraria' não existe no W4.")

    data_tes = formatar_data_coluna(df["Data da Tesouraria"])

    # ============================
    # CENTRO DE CUSTO — SINODALIDADE
    # ============================

    if setor == "Sinodalidade" and "Lote" in df.columns:
        centro_custo = df["Lote"].fillna("").astype(str).str.strip()
        centro_custo = centro_custo.replace(
            ["", "nan", "NaN"], "Adm Financeiro"
        )
    else:
        centro_custo = ""

    # ============================
    # MONTAGEM FINAL (ORDEM CORRETA)
    # ============================

    out = pd.DataFrame()
    out["Data de Competência"] = data_tes
    out["Data de Vencimento"] = data_tes
    out["Data de Pagamento"] = data_tes
    out["Valor"] = df["Valor_str_final"]
    out["Categoria"] = df["Categoria_final"]

    # CONCATENAR ID + DESCRIÇÃO (se existir)
    if "Descrição" not in df.columns:
        df["Descrição"] = ""

    if "Id Item tesouraria" in df.columns:
        out["Descrição"] = (
            df["Id Item tesouraria"].astype(str) + " " + df["Descrição"].astype(str)
        )
    else:
        out["Descrição"] = df["Descrição"].astype(str)

    # Cliente/Fornecedor (preenchido pela regra da Previdência quando aplicável)
    out["Cliente/Fornecedor"] = df.get("ClienteFornecedor_final", "")

    out["CNPJ/CPF Cliente/Fornecedor"] = ""
    out["Centro de Custo"] = centro_custo
    out["Observações"] = ""

    return out


# ============================
# CARREGAR ARQUIVO W4
# ============================

def carregar_arquivo_w4(arq):
    if arq.name.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(arq)
    else:
        return pd.read_csv(arq, sep=";", encoding="latin1")


# ============================
# CARREGAR CATEGORIAS
# ============================

df_cat_raw = pd.read_excel("categorias_contabeis.xlsx")
df_cat_prep = preparar_categorias(df_cat_raw)

# ============================
# INTERFACE
# ============================

st.title("Conversão Planilha Setores")

setor = st.selectbox(
    "Selecione o setor",
    [
        "Ass. Comunitária",
        "Sinodalidade",
        "Previdência Brasil"
    ]
)

arq_w4 = st.file_uploader(
    "Envie o arquivo W4 (CSV ou Excel)",
    type=["csv", "xlsx", "xls"]
)

if arq_w4:
    if st.button("Converter planilha"):
        try:
            df_w4 = carregar_arquivo_w4(arq_w4)
            df_final = converter_w4(df_w4, df_cat_prep, setor)

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
else:
    st.info("Selecione um setor e envie o arquivo W4.")
