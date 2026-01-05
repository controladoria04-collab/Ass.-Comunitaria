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
    """
    Normaliza texto para comparação:
    - minúsculo
    - remove acentos
    - troca pontuação por espaço
    - remove espaços duplicados
    """
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
# PREVIDÊNCIA: CLIENTES + MATCH ROBUSTO + CENTRO DE CUSTO
# ============================

CENTROS_CUSTO_VALIDOS = {"DIACONIA", "FORTALEZA", "BRASIL", "EXTERIOR"}

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

def _extrair_base_e_centro(cliente_str: str):
    """
    Se terminar com ' - DIACONIA/FORTALEZA/BRASIL/EXTERIOR':
      - base = tudo antes do último ' - '
      - centro = último item
    """
    s = str(cliente_str).strip()
    partes = [p.strip() for p in s.split(" - ")]
    if len(partes) >= 2 and partes[-1].upper() in CENTROS_CUSTO_VALIDOS:
        centro = partes[-1].upper()
        base = " - ".join(partes[:-1]).strip()
        return base, centro
    return s, ""

# Pré-processamento: tokens da base + centro
CLIENTES_PREV_INFO = []
for c in CLIENTES_PREVIDENCIA:
    base, centro = _extrair_base_e_centro(c)
    base_norm = normalize_text(base)
    tokens_base = set(base_norm.split()) if base_norm else set()
    CLIENTES_PREV_INFO.append({
        "cliente_oficial": c,
        "base_norm": base_norm,
        "tokens_base": tokens_base,
        "centro": centro
    })

def _match_cliente_previdencia(complemento_norm: str):
    """
    Recebe complemento JÁ NORMALIZADO.
    Match robusto:
    - tokens em comum
    - bônus forte se existir palavra "forte" (>=5 letras) que bate
    - bônus substring para casos como "casa renata couras"
    Retorna (cliente_oficial, centro) ou ("", "").
    """
    comp_norm = str(complemento_norm).strip()
    if not comp_norm:
        return "", ""

    tokens_comp = set(comp_norm.split())
    if not tokens_comp:
        return "", ""

    melhor_cliente = ""
    melhor_centro = ""
    melhor_score = 0.0
    melhor_inter = 0

    for info in CLIENTES_PREV_INFO:
        tokens_base = info["tokens_base"]
        if not tokens_base:
            continue

        inter = tokens_comp & tokens_base
        inter_count = len(inter)

        # cobertura da base
        score = inter_count / max(1, len(tokens_base))

        # bônus: palavra forte bate exatamente (ex.: lubango, couras, parquelandia)
        if any((t in tokens_base) and (len(t) >= 5) for t in tokens_comp):
            score = max(score, 0.98)

        # bônus substring
        if info["base_norm"] and (info["base_norm"] in comp_norm or comp_norm in info["base_norm"]):
            score = max(score, 0.95)

        # desempate
        if (score > melhor_score) or (score == melhor_score and inter_count > melhor_inter):
            melhor_score = score
            melhor_cliente = info["cliente_oficial"]
            melhor_centro = info["centro"]
            melhor_inter = inter_count

    # Critérios mínimos (mais permissivos e práticos)
    if melhor_cliente and (melhor_score >= 0.55 or melhor_inter >= 1):
        return melhor_cliente, melhor_centro

    return "", ""


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
    # PREVIDÊNCIA: REPASSE FUNDO DE PREVIDÊNCIA (detecção normalizada)
    # ============================

    df["ClienteFornecedor_final"] = ""
    df["CentroCusto_final"] = ""

    if setor in ["Previdência Brasil", "Previdência"]:
        base_txt_norm = normalize_text("Repasse Recebido Fundo de Previdência")

        detalhe_norm = df[col_cat].astype(str).apply(normalize_text)
        mask_repasse = detalhe_norm.str.contains(base_txt_norm, na=False)

        if mask_repasse.any():
            complemento_norm = (
                detalhe_norm[mask_repasse]
                .str.replace(base_txt_norm, "", regex=False)
                .str.strip()
            )

            resultados = complemento_norm.apply(_match_cliente_previdencia)
            clientes = resultados.apply(lambda x: x[0])
            centros = resultados.apply(lambda x: x[1])

            # máscara final alinhada ao df
            mask_achou = mask_repasse.copy()
            mask_achou.loc[mask_repasse] = clientes.ne("").values

            df.loc[mask_achou, "Categoria_final"] = "11318 - Repasse Recebido Fundo de Previdência"
            df.loc[mask_achou, "ClienteFornecedor_final"] = clientes.values
            df.loc[mask_achou, "CentroCusto_final"] = centros.values

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
    # MONTAGEM FINAL (ORDEM CORRETA)
    # ============================

    out = pd.DataFrame()
    out["Data de Competência"] = data_tes
    out["Data de Vencimento"] = data_tes
    out["Data de Pagamento"] = data_tes
    out["Valor"] = df["Valor_str_final"]
    out["Categoria"] = df["Categoria_final"]

    # Descrição
    if "Descrição" not in df.columns:
        df["Descrição"] = ""

    if "Id Item tesouraria" in df.columns:
        out["Descrição"] = (
            df["Id Item tesouraria"].astype(str) + " " + df["Descrição"].astype(str)
        )
    else:
        out["Descrição"] = df["Descrição"].astype(str)

    # Cliente/Fornecedor (só Previdência preenche)
    out["Cliente/Fornecedor"] = df.get("ClienteFornecedor_final", "")

    out["CNPJ/CPF Cliente/Fornecedor"] = ""

    # Centro de custo:
    # - Previdência: centro extraído do sufixo do cliente
    # - Sinodalidade: Lote (como no seu código original)
    # - Outros: vazio
    if setor in ["Previdência Brasil", "Previdência"]:
        centro_custo_out = df.get("CentroCusto_final", "")
    elif setor == "Sinodalidade" and "Lote" in df.columns:
        centro_custo_out = df["Lote"].fillna("").astype(str).str.strip()
        centro_custo_out = centro_custo_out.replace(["", "nan", "NaN"], "Adm Financeiro")
    else:
        centro_custo_out = ""

    out["Centro de Custo"] = centro_custo_out
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
# CARREGAR CATEGORIAS (cache + erro amigável)
# ============================

@st.cache_data
def carregar_categorias():
    df_cat_raw = pd.read_excel("categorias_contabeis.xlsx")
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
