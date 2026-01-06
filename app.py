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
# PREVIDÊNCIA: MAPEAMENTO DIRETO
# ============================

CENTROS_CUSTO_VALIDOS = {"FORTALEZA", "DIACONIA", "EXTERIOR", "BRASIL"}

MAPEAMENTO_PREVIDENCIA = [
    ("APARECIDA - BRASIL", "Repasse Recebido Fundo de Previdência Missão Aparecida"),
    ("ARACAJU - BRASIL", "Repasse Recebido Fundo de Previdência Missão Aracaju"),
    ("ARACATI - BRASIL", "Repasse Recebido Fundo de Previdência Missão Aracati"),
    ("ARAPIRACA - BRASIL", "Repasse Recebido Fundo de Previdência Missão Arapiraca"),
    ("ARARAQUARA - BRASIL", "Repasse Recebido Fundo de Previdência Missão Araraquara"),
    ("ASSESSORIA DE DIFUSÃO DA OBRA - DIACONIA", "Repasse Recebido Fundo de Previdência Difusão da Obra"),
    ("ASSESSORIA DE PROMOÇÃO HUMANA - DIACONIA", "Repasse Recebido Fundo de Previdência Assessoria de Promoção Humana"),
    ("ASSESSORIA JOVEM - DIACONIA", "Repasse Recebido Fundo de Previdência Assessoria Jovem"),
    ("ASSESSORIA LITURGICO - SACRAMENTAL - DIACONIA", "Repasse Recebido Fundo de Previdência Litúrgico Sacramental"),
    ("ASSESSORIA VOCACIONAL - DIACONIA", "Repasse Recebido Fundo de Previdência Assessoria Vocacional"),
    ("ASSISTÊNCIA APOSTÓLICA - DIACONIA", "Repasse Recebido Fundo de Previdência Assistência Apostólica"),
    ("ASSISTÊNCIA COMUNITÁRIA - DIACONIA", "Repasse Recebido Fundo de Previdência Assessoria Comunitária"),
    ("ASSISTÊNCIA DE COMUNICAÇÃO - DIACONIA", "Repasse Recebido Fundo de Previdência Assistencia de Comunicação"),
    ("ASSISTÊNCIA DE FORMAÇÃO - DIACONIA", "Repasse Recebido Fundo de Previdência Assistência de Formação"),
    ("ASSISTÊNCIA LOCAL - FORTALEZA", "Repasse Recebido Fundo de Previdência Assistência Local"),
    ("ASSISTÊNCIA MISSIONÁRIA - DIACONIA", "Repasse Recebido Fundo de Previdência Assistência Missionária"),
    ("BEJAIA - EXTERIOR", "Repasse Recebido Fundo de Previdência Bejaia"),
    ("BELÉM - BRASIL", "Repasse Recebido Fundo de Previdência Missão Belém"),
    ("BELO HORIZONTE - BRASIL", "Repasse Recebido Fundo de Previdência Missão Belo Horizonte"),
    ("BOGOTÁ - EXTERIOR", "Repasse Recebido Fundo de Previdência Bogotá"),
    ("BOSTON - EXTERIOR", "Repasse Recebido Fundo de Previdência Missão Boston"),
    ("BRASÍLIA - BRASIL", "Repasse Recebido Fundo de Previdência Missão Brasília"),
    ("CABO VERDE - EXTERIOR", "Repasse Recebido Fundo de Previdência Missão Cabo Verde"),
    ("CAMPINA GRANDE - BRASIL", "Repasse Recebido Fundo de Previdência Missão Campina Grande"),
    ("CAMPO GRANDE - BRASIL", "Repasse Recebido Fundo de Previdência Missão Campo Grande"),
    ("CASA DE RETIRO - DIACONIA", "Repasse Recebido Fundo de Previdência Casa de Retiro São João Paulo II"),
    ("CEST - CASA CONTEMPLATIVA - BRASIL", "Repasse Recebido Fundo de Previdência Casa Contemplativa"),
    ("CEST - FORTALEZA", "Repasse Recebido Fundo de Previdência CEST Fortaleza"),
    ("CEV - AQUIRAZ - FORTALEZA", "Repasse Recebido Fundo de Previdência CEV Aquiraz"),
    ("CEV - CARMO - FORTALEZA", "Repasse Recebido Livraria CEV Nossa Sra do Carmo"),
    ("CEV - SHALOM CID. DOS FUNCIONÁRIOS - FORTALEZA", "Repasse Recebido Fundo de Previdência CEV Cidade dos Funcionários"),
    ("CEV - SHALOM CRISTO REDENTOR - FORTALEZA", "Repasse Recebido Fundo de Previdência CEV Cristo Redentor"),
    ("CEV - SHALOM FÁTIMA - FORTALEZA", "Repasse Recebido Fundo de Previdência CEV Fátima"),
    ("CEV - SHALOM PARANGABA - FORTALEZA", "Repasse Recebido Fundo de Previdência CEV Parangaba"),
    ("CEV - SHALOM PARQUELÂNDIA - FORTALEZA", "Repasse Recebido Fundo de Previdência CEV Parquelandia"),
    ("CEV - SHALOM PAZ - FORTALEZA", "Repasse Recebido Fundo de Previdência CEV Shalom da Paz"),
    ("CHAVES - BRASIL", "Repasse Recebido Fundo de Previdência Missão Chaves"),
    ("COLÉGIO SHALOM - FORTALEZA", "Repasse Recebido Fundo de Previdência Colégio Shalom"),
    ("COORDENAÇÃO APOSTÓLICA - FORTALEZA", "Repasse Recebido Fundo de Previdência Coordenação Apostólica"),
    ("CRATEÚS - BRASIL", "Repasse Recebido Fundo de Previdência Missão Crateús"),
    ("CRISMA - FORTALEZA", "Repasse Recebido Fundo de Previdência Crisma"),
    ("CRUZEIRO DO SUL - BRASIL", "Repasse Recebido Fundo de Previdência Missão Cruzeiro do Sul"),
    ("CUIABÁ - BRASIL", "Repasse Recebido Fundo de Previdência Missão Cuiabá"),
    ("CURITIBA - BRASIL", "Repasse Recebido Fundo de Previdência Missão Curitiba"),
    ("DISCIPULADO EUSÉBIO - FORTALEZA", "Repasse Recebido Fundo de Previdência Discipulado Eusébio"),
    ("DISCIPULADO PACAJUS - FORTALEZA", "Repasse Recebido Fundo de Previdência Discipulado Pacajus"),
    ("DISCIPULADO QUIXADÁ - FORTALEZA", "Repasse Recebido Fundo de Previdência Discipulado Quixadá"),
    ("ECONOMATO GERAL - DIACONIA", "Repasse Recebido Fundo de Previdência Economato Diaconia"),
    ("EDIÇÕES - DIACONIA", "Repasse Recebido Fundo de Previdência Edições Shalom"),
    ("ESCOLA DE EVANGELIZAÇÃO - FORTALEZA", "Repasse Recebido Fundo de Previdência Escola de Evangelização"),
    ("ESCRITÓRIO GERAL - DIACONIA", "Repasse Recebido Fundo de Previdência Escritório Geral"),
    ("ESC. SECRETÁRIA COMUNITÁRIA - FORTALEZA", "Repasse Recebido Fundo de Previdência Secretaria Comunitária Fortaleza"),
    ("FLORIANÓPOLIS - BRASIL", "Repasse Recebido Fundo de Previdência Missão Florianópolis"),
    ("GARANHUNS - BRASIL", "Repasse Recebido Fundo de Previdência Missão Garanhuns"),
    ("GOIÂNIA - BRASIL", "Repasse Recebido Fundo de Previdência Missão Goiânia"),
    ("GUARULHOS - BRASIL", "Repasse Recebido Fundo de Previdência Missão Guarulhos"),
    ("GUIANA FRANCESA - EXTERIOR", "Repasse Recebido Fundo de Previdência Missão Guiana Francesa"),
    ("HAIFA - EXTERIOR", "Repasse Recebido Fundo de Previdência Missão Haifa"),
    ("IGREJA - DIACONIA", "Repasse Recebido Fundo de Previdência Igreja do Ressuscitado"),
    ("IMPERATRIZ - BRASIL", "Repasse Recebido Fundo de Previdência Missão Imperatriz"),
    ("ITAPIPOCA - BRASIL", "Repasse Recebido Fundo de Previdência Missão Itapipoca"),
    ("JOÃO PESSOA - BRASIL", "Repasse Recebido Fundo de Previdência Missão João Pessoa"),
    ("JOINVILLE - BRASIL", "Repasse Recebido Fundo de Previdência Missão Joinville"),
    ("JUAZEIRO DA BAHIA - BRASIL", "Repasse Recebido Fundo de Previdência Missão Juazeiro da Bahia"),
    ("JUAZEIRO DO NORTE - BRASIL", "Repasse Recebido Fundo de Previdência Missão Juazeiro do Norte"),
    ("JUIZ DE FORA - BRASIL", "Repasse Recebido Fundo de Previdência Missão Juiz de Fora"),
    ("LANÇAI AS REDES - DIACONIA", "Repasse Recebido Fundo de Previdência Lançai as Redes"),
    ("LANCHONETE PARQUELÂNDIA - FORTALEZA", "Repasse Recebido Fundo de Previdência Lanchonete Parquelandia"),
    ("LANCHONETE - SHALOM DA PAZ - FORTALEZA", "Repasse Recebido Fundo de Previdência Lanchonete Shalom da Paz"),
    ("LIMA - EXTERIOR", "Repasse Recebido Fundo de Previdência Missão Lima"),
    ("LIVRARIA - PARANGABA - FORTALEZA", "Repasse Recebido Fundo de Previdência Livraria Parangaba"),
    ("LUBANGO - EXTERIOR", "Repasse Recebido Fundo de Previdência Missão Lubango"),
    ("MACAPÁ - BRASIL", "Repasse Recebido Fundo de Previdência Missão Macapá"),
    ("MACEIÓ - BRASIL", "Repasse Recebido Fundo de Previdência Missão Maceió"),
    ("MADAGASCAR - EXTERIOR", "Repasse Recebido Fundo de Previdência Missão Madagascar"),
    ("MANAUS - BRASIL", "Repasse Recebido Fundo de Previdência Missão Manaus"),
    ("MANILA - EXTERIOR", "Repasse Recebido Fundo de Previdência Missão Manila"),
    ("MATRIZ - FORTALEZA", "Repasse Recebido Fundo de Previdência Matriz"),
    ("MOÇAMBIQUE - EXTERIOR", "Repasse Recebido Fundo de Previdência Missão Moçambique"),
    ("MOSSORÓ - BRASIL", "Repasse Recebido Fundo de Previdência Missão Mossoró"),
    ("NATAL - BRASIL", "Repasse Recebido Fundo de Previdência Missão Natal"),
    ("NAZARETH - EXTERIOR", "Repasse Recebido Fundo de Previdência Missão Nazareth"),
    ("NITERÓI - BRASIL", "Repasse Recebido Fundo de Previdência Missão Niterói"),
    ("PALMAS - BRASIL", "Repasse Recebido Fundo de Previdência Missão Palmas"),
    ("PARNAÍBA - BRASIL", "Repasse Recebido Fundo de Previdência Missão Parnaíba"),
    ("PARRESIA - DIACONIA", "Repasse Recebido Fundo de Previdência Instituto Parresia"),
    ("PATOS - BRASIL", "Repasse Recebido Fundo de Previdência Missão Patos"),
    ("PATOS DE MINAS - BRASIL", "Repasse Recebido Fundo de Previdência Missão Patos de Minas"),
    ("PH - CASA RENATA COURAS - FORTALEZA", "Repasse Recebido Fundo de Previdência Casa Renata Couras"),
    ("PH - CASA RONALDO PEREIRA - FORTALEZA", "Repasse Recebido Fundo de Previdência Casa Ronaldo Pereira"),
    ("PH - PROJETO AMIGO DOS POBRES - FORTALEZA", "Repasse Recebido Fundo de Previdência Projeto Amigo dos Pobres"),
    ("PH - PROJETO MARIA MADALENA - FORTALEZA", "Repasse Recebido Fundo de Previdência Projeto Maria Madalena"),
    ("PH - SECRETARIA - FORTALEZA", "Repasse Recebido Fundo de Previdência Secretaria de PH Fortaleza"),

    # ---- NOVOS QUE VOCÊ PEDIU ----
    ("PIRACICABA - BRASIL", "Repasse Recebido Fundo de Previdência Missão Piracicaba"),
    ("PONTA GROSSA - BRASIL", "Repasse Recebido Fundo de Previdência Missão Ponta Grossa"),
    ("PREFEITURA - DIACONIA", "Repasse Recebido Fundo de Previdência Prefeitura"),
    ("PROJETO ARTES - FORTALEZA", "Repasse Recebido Fundo de Previdência Projeto Artes"),
    ("PROJETO JUVENTUDE - FORTALEZA", "Repasse Recebido Fundo de Previdência Projeto Juventude Fortaleza"),
    ("QUIXADÁ - BRASIL", "Repasse Recebido Fundo de Previdência Missão Quixadá"),
    ("RÁDIO SHALOM - FORTALEZA", "Repasse Recebido Fundo de Previdência Rádio Shalom AM 690"),
    ("RECIFE - BRASIL", "Repasse Recebido Fundo de Previdência Missão Recife"),
    ("REG - CID. DOS FUNCIONÁRIOS - FORTALEZA", "Repasse Recebido Fundo de Previdência Regional Shalom Cidade dos Funcionários"),
    ("REG.  PACAJUS - FORTALEZA", "Repasse Recebido Fundo de Previdência Pacajus"),
    ("REG - PARANGABA - FORTALEZA", "Repasse Recebido Fundo de Previdência Regional Shalom Parangaba"),
    ("REG - PARQUELÂNDIA - FORTALEZA", "Repasse Recebido Fundo de Previdência Regional Shalom Parquelandia"),
    ("RIO DE JANEIRO - BRASIL", "Repasse Recebido Fundo de Previdência Missão Rio de Janeiro"),
    ("ROMA - EXTERIOR", "Repasse Recebido Fundo de Previdência Missão Roma"),
    ("SALVADOR - BRASIL", "Repasse Recebido Fundo de Previdência Missão Salvador"),
    ("SANTA CRUZ DE LA SIERRA - EXTERIOR", "Repasse Recebido Fundo de Previdência Santa Cruz de La Sierra"),
    ("SANTO AMARO - BRASIL", "Repasse Recebido Fundo de Previdência Missão Santo Amaro"),
    ("SANTO ANDRÉ - BRASIL", "Repasse Recebido Fundo de Previdência Missão Santo André"),
    ("SÃO LEOPOLDO - BRASIL", "Repasse Recebido Fundo de Previdência Missão São Leopoldo"),
    ("SÃO LUÍS - BRASIL", "Repasse Recebido Fundo de Previdência Missão São Luis"),
    ("SÃO PAULO - PERDIZES - BRASIL", "Repasse Recebido Fundo de Previdência Missão São Paulo"),
    ("SÃO PAULO - TÁIPAS - BRASIL", "Repasse Recebido Fundo de Previdência Missão São Paulo"),
    ("SEC. DE COMUNICAÇÃO - FORTALEZA", "Repasse Recebido Fundo de Previdência Secretaria de Comunicação"),
    ("SEC. DE SACERDOTES E SEMINARISTAS - DIACONIA", "Repasse Recebido Fundo de Previdência Sacerdotes e Seminaristas"),
    ("SECRETARIA DE PLANEJAMENTO - DIACONIA", "Repasse Recebido Fundo de Previdência Secretaria de Planejamento"),
    ("SECRETARIA VOCACIONAL - FORTALEZA", "Repasse Recebido Fundo de Previdência Secretaria Vocacional Fortaleza"),
    ("SENHOR DO BONFIM - BRASIL", "Repasse Recebido Fundo de Previdência Missão Senhor do Bonfim"),
    ("SETOR COLEGIADO - DIACONIA", "Repasse Recebido Fundo de Previdência Setor Colegiado"),
    ("SETOR DE EVENTOS - FORTALEZA", "Repasse Recebido Fundo de Previdência Setor de Eventos Fortaleza"),
    ("SETOR DOS CELIBATÁRIOS - DIACONIA", "Repasse Recebido Fundo de Previdência Setor Celibatários"),
    ("SETOR FAMÍLLIA - DIACONIA", "Repasse Recebido Fundo de Previdência Setor Familias"),
    ("SOBRAL - BRASIL", "Repasse Recebido Fundo de Previdência Missão Sobral"),
    ("TAIWAN - EXTERIOR", "Repasse Recebido Fundo de Previdência Taiwan"),
    ("TECNOLOGIA DA INFORMAÇÃO - DIACONIA", "Repasse Recebido Tecnologia da Informação"),
    ("TERESINA - BRASIL", "Repasse Recebido Fundo de Previdência Missão Teresina"),
    ("TUNÍSIA - EXTERIOR", "Repasse Recebido Fundo de Previdência Missão Tunísia"),
    ("UBERABA - BRASIL", "Repasse Recebido Fundo de Previdência Missão Uberaba"),
    ("VITÓRIA - BRASIL", "Repasse Recebido Fundo de Previdência Missão Vitória ES"),
    ("VITÓRIA DA CONQUISTA - BRASIL", "Repasse Recebido Fundo de Previdência Missão Vitória da Conquista - Difusão"),
]

# Pré-normalizado + centro de custo calculado
MAPEAMENTO_PREVIDENCIA_PREP = []
for cliente, texto_w4 in MAPEAMENTO_PREVIDENCIA:
    cliente_str = str(cliente).strip()
    texto_norm = normalize_text(texto_w4)
    partes = [p.strip() for p in cliente_str.split(" - ")]
    centro = partes[-1].upper() if partes and partes[-1].upper() in CENTROS_CUSTO_VALIDOS else ""
    MAPEAMENTO_PREVIDENCIA_PREP.append((texto_norm, cliente_str, centro))

# ============================
# FUNÇÃO PRINCIPAL
# ============================

def converter_w4(df_w4, df_categorias_prep, setor):

    df_w4 = df_w4.copy()
    df_w4.columns = df_w4.columns.astype(str).str.strip()

    if "Detalhe Conta / Objeto" not in df_w4.columns:
        raise ValueError(
            f"Coluna 'Detalhe Conta / Objeto' não existe no W4. Colunas: {list(df_w4.columns)}"
        )

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
    # PREVIDÊNCIA (sem erro de assign): mapeamento por linha
    # ============================

    df["ClienteFornecedor_final"] = ""
    df["CentroCusto_final"] = ""

    if str(setor).strip() == "Previdência Brasil":
        detalhe_norm = df[col_cat].astype(str).apply(normalize_text)

        def buscar_mapeamento(txt_norm: str):
            for padrao_norm, cliente, centro in MAPEAMENTO_PREVIDENCIA_PREP:
                if padrao_norm and padrao_norm in txt_norm:
                    return cliente, centro
            return "", ""

        resultados = detalhe_norm.apply(buscar_mapeamento)
        df["ClienteFornecedor_final"] = resultados.apply(lambda x: x[0])
        df["CentroCusto_final"] = resultados.apply(lambda x: x[1])

        # Só troca categoria onde encontrou cliente
        achou = df["ClienteFornecedor_final"].ne("")
        df.loc[achou, "Categoria_final"] = "11318 - Repasse Recebido Fundo de Previdência"

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
    # MONTAGEM FINAL
    # ============================

    out = pd.DataFrame()
    out["Data de Competência"] = data_tes
    out["Data de Vencimento"] = data_tes
    out["Data de Pagamento"] = data_tes
    out["Valor"] = df["Valor_str_final"]
    out["Categoria"] = df["Categoria_final"]

    if "Descrição" not in df.columns:
        df["Descrição"] = ""

    if "Id Item tesouraria" in df.columns:
        out["Descrição"] = (
            df["Id Item tesouraria"].astype(str) + " " + df["Descrição"].astype(str)
        )
    else:
        out["Descrição"] = df["Descrição"].astype(str)

    out["Cliente/Fornecedor"] = df.get("ClienteFornecedor_final", "")
    out["CNPJ/CPF Cliente/Fornecedor"] = ""

    if str(setor).strip() == "Previdência Brasil":
        out["Centro de Custo"] = df.get("CentroCusto_final", "")
    elif str(setor).strip() == "Sinodalidade" and "Lote" in df.columns:
        centro_custo_out = df["Lote"].fillna("").astype(str).str.strip()
        centro_custo_out = centro_custo_out.replace(["", "nan", "NaN"], "Adm Financeiro")
        out["Centro de Custo"] = centro_custo_out
    else:
        out["Centro de Custo"] = ""

    out["Observações"] = ""

    return out


# ============================
# CARREGAR ARQUIVO W4
# ============================

def carregar_arquivo_w4(arq):
    if arq.name.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(arq)
    else:
        df = pd.read_csv(arq, sep=";", encoding="latin1")
    df.columns = df.columns.astype(str).str.strip()
    return df


# ============================
# CARREGAR CATEGORIAS
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
