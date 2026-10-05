import streamlit as st
import os
import pandas as pd
from datetime import date, datetime
from io import BytesIO

# ─── Configuração da página ───────────────────────────────────────────────────
st.set_page_config(
    page_title="DOit - Atualização de Custos",
    page_icon="💰",
    layout="wide",
)

# ─── Estilo customizado ──────────────────────────────────────────────────────
st.markdown("""
<style>
    .main .block-container { padding-top: 1.2rem; max-width: 1080px; }
    h1 { color: #1f2933; font-size: 1.7rem; }

    /* Cabeçalho de seção: sóbrio, sem cores fortes */
    .section-header {
        font-size: 1.05rem;
        font-weight: 600;
        color: #1f2933;
        margin: 0.2rem 0 0.6rem 0;
        display: flex;
        align-items: center;
        gap: 0.45rem;
    }
    .step-badge {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 1.5rem; height: 1.5rem;
        border-radius: 50%;
        background: #e5e7eb;
        color: #374151;
        font-size: 0.85rem;
        font-weight: 700;
    }

    /* Caixa informativa neutra (cinza claro), usada com parcimônia */
    .info-box {
        background: #f4f5f7;
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        padding: 0.7rem 0.9rem;
        font-size: 0.85rem;
        color: #4b5563;
        margin-bottom: 0.8rem;
    }

    /* Abas com respiro */
    .stTabs [data-baseweb="tab-list"] { gap: 1.2rem; }

    /* Botão primário com azul discreto */
    .stButton > button[kind="primary"] {
        background: #2563eb;
        border: none;
    }
</style>
""", unsafe_allow_html=True)

# ─── Carregar planilha DOit ───────────────────────────────────────────────────
CAMINHO_DOIT = "ListagemdeProdutos DOit.xlsx"

# Planilha de controle (Google Sheets) onde o registro é colado
URL_PLANILHA_CONTROLE = "https://docs.google.com/spreadsheets/d/11N3sHupQLe4FjITXUvhvGdx6Sf_ZosB_PG7xGuP1PUo/edit?pli=1&gid=0#gid=0"


CAMINHO_DOIT_PARQUET = "ListagemdeProdutos DOit.parquet"


@st.cache_data
def carregar_doit():
    # Parquet é ~20x mais rápido de ler que Excel; usamos quando disponível para
    # acelerar o "acordar" do app. Cai para o Excel se o Parquet não existir.
    if os.path.exists(CAMINHO_DOIT_PARQUET):
        df = pd.read_parquet(CAMINHO_DOIT_PARQUET, engine="pyarrow")
    else:
        df = pd.read_excel(CAMINHO_DOIT)
    df["# Referência"] = df["# Referência"].astype(str).str.strip()
    # Muitas referências do DOit vêm como "CÓDIGO | 394" (código + id interno).
    # Guardamos a parte antes do " | " para cruzar com o código puro do fornecedor.
    df["_ref_base"] = (
        df["# Referência"].str.split(r"\s*\|\s*", n=1, regex=True).str[0].str.strip()
    )
    return df


def data_atualizacao_doit():
    """Data da última modificação do arquivo da base DOit (dd/mm/aaaa) ou None."""
    caminho = CAMINHO_DOIT_PARQUET if os.path.exists(CAMINHO_DOIT_PARQUET) else CAMINHO_DOIT
    try:
        ts = os.path.getmtime(caminho)
        return datetime.fromtimestamp(ts).strftime("%d/%m/%Y")
    except OSError:
        return None


try:
    df_doit = carregar_doit()
except FileNotFoundError:
    st.error("Arquivo 'ListagemdeProdutos DOit.xlsx' não encontrado na pasta do projeto.")
    st.stop()

# ─── Receitas por fornecedor ──────────────────────────────────────────────────
# Cada receita pré-configura o app para um fornecedor conhecido: qual aba usar,
# em que linha está o cabeçalho, quais colunas são código/preço e quais opções
# ativar. O IPI é sempre preenchido à mão pelo usuário.
#   aba: nome exato da aba a manter selecionada (None = todas)
#   header: linha do cabeçalho (0 = primeira linha)
#   codigo/preco: nome da coluna após aplicar o cabeçalho
#   opcoes: flags de compatibilização a marcar
#   segunda_col / valor_col / ipi_col: colunas auxiliares de cada opção
#   passos: instruções mostradas ao usuário
RECEITAS = {
    "Stella": {
        "aba": "Planilha",
        "header": 0,
        "codigo": "Referência",
        "preco": "Prç.ven",
        "ipi_por_produto": True,
        "ipi_col": "% IPI",
        "fab_id": 7586,
        "passos": [
            "Cabeçalho na linha 0 (primeira linha).",
            "Código = coluna **Referência** · Preço = coluna **Prç.ven**.",
            "IPI vem por produto na coluna **% IPI** (já marcado).",
            "Deixe o campo IPI (%) em 0 — o IPI de cada item é usado automaticamente.",
        ],
    },
    "Rosa Maria": {
        "aba": "Tabela 2026",
        "header": 1,
        "codigo": "REFERÊNCIA",
        "preco": "PREÇO",
        "agregar_acabamentos": True,
        "fab_id": 7603,
        "passos": [
            "Use apenas a aba **Tabela 2026** (as outras são markup/descritivo).",
            "Cabeçalho na linha 1.",
            "Código = coluna **REFERÊNCIA** · Preço = coluna **PREÇO**.",
            "Cada código tem 2 acabamentos: o app usa o **maior preço** (já marcado).",
            "Preencha o IPI (%) à mão, se houver.",
        ],
    },
    "Revoluz": {
        "aba": "Plan1",
        "header": 1,
        "codigo": "Produto",
        "preco": "Valor\nUnitário",
        "normalizar": True,
        "fab_id": 7684,
        "passos": [
            "Cabeçalho na linha 1.",
            "Código = coluna **Produto** · Preço = coluna **Valor Unitário**.",
            "Códigos têm acabamentos (ex: -BFM OU PTO): normalização já marcada.",
            "Preencha o IPI (%) à mão, se houver.",
        ],
    },
    "Revolux": {
        "aba": "Plan1",
        "header": 1,
        "codigo": "Produto",
        "preco": "Valor\nUnitário",
        "normalizar": True,
        "fab_id": 7684,
        "passos": [
            "Cabeçalho na linha 1.",
            "Código = coluna **Produto** · Preço = coluna **Valor Unitário**.",
            "Códigos têm acabamentos: normalização já marcada.",
            "Preencha o IPI (%) à mão, se houver.",
        ],
    },
    "Spotline": {
        "aba": "TABELA SPOLINE",
        "header": 0,
        "codigo": "ID",
        "preco": "PREÇO",
        "chaves_multiplas": True,
        "segunda_col": "DESCRIÇÃO",
        "fab_id": 7629,
        "passos": [
            "Cabeçalho na linha 0.",
            "Código = coluna **ID** · Preço = coluna **PREÇO**.",
            "No DOit o produto pode estar pelo **ID** (ex: 9766) ou pelo **modelo** "
            "(início da **DESCRIÇÃO**, ex: 1362/1). O app tenta casar pelos dois.",
            "Preencha o IPI (%) à mão, se houver.",
        ],
    },
    "Accord": {
        # 6 abas de produto; as demais (Opções/Informações/Política) são
        # informativas e podem ficar selecionadas — suas linhas não têm preço
        # válido e são descartadas no processamento.
        "aba": None,
        "header": 1,
        # Na aba Pendente (1ª) o cabeçalho "Referência"/"Descrição" não está na
        # linha 1, então as colunas saem como "Unnamed: 2" (código) e
        # "Geral 2025" (preço Lâmina Natural). As posições são as mesmas em todas
        # as abas de produto: código = col 2, Natural = col 4, Tingida = col 5.
        # Apenas as abas de produto são processadas; as informativas (Opções de
        # Acabamento, Informações LED, Política de Compra e Venda) são deixadas
        # fora da seleção — não têm preço e quebravam o preview.
        "abas": ["Pendente", "Plafon", "Arandela", "Abajur", "Coluna", "Mobiliário"],
        "codigo": "Unnamed: 2",
        "preco": "Geral 2025",       # col 4 → Lâmina Natural
        "preco_tingida": "Unnamed: 5",  # col 5 → Lâmina Tingida
        # Dois acabamentos: cada referência tem 2 preços (Natural e Tingida) e no
        # DOit vira vários SKUs. Os SKUs com "NATURAL" no nome recebem o preço
        # Natural; os de cor/tingido recebem o preço Tingida.
        "dois_acabamentos": True,
        # A Accord NÃO tem IPI. Os 10% são MARGEM aplicada sobre o preço:
        # CUSTO LÍQUIDO = preço do acabamento · CUSTO = líquido × 1,10.
        "margem_fixa": 10.0,
        "fab_id": 7623,
        "passos": [
            "A receita já seleciona só as abas de produto (Pendente, Plafon, "
            "Arandela, Abajur, Coluna e Mobiliário). As abas informativas "
            "(acabamento, LED, política) são ignoradas automaticamente.",
            "Cabeçalho na linha 1.",
            "Código = coluna **Referência** (col. C). A planilha tem 2 preços: "
            "**Lâmina Natural** (col. E) e **Lâmina Tingida** (col. F).",
            "O app diferencia o acabamento: SKUs com **NATURAL** no nome recebem "
            "o preço da Lâmina Natural; os de **cor/tingido** recebem o da Tingida.",
            "**Sem IPI.** Os 10% são **margem**: CUSTO = preço × 1,10 e "
            "CUSTO LÍQUIDO = preço (antes da margem).",
            "Se a Lâmina Tingida vier **vazia**, os SKUs de cor daquela "
            "referência ficam **sem atualizar**.",
            "Aba **Coluna**: ignore os valores de IPI (7,5% / 10%) que aparecem "
            "no topo — a Accord não usa IPI.",
        ],
        "observacao": "Layout irregular por aba: confira o preview antes de processar.",
    },
    "Golden Art": {
        "aba": None,  # tem 8 abas (Table 1..8) — deixe todas selecionadas
        "header": 0,
        "usar_valor": True,
        "fab_id": 7583,
        "passos": [
            "Golden Art tem várias abas (Table 1..8) — deixe **todas** selecionadas.",
            "O preço fica separado: 'R$' numa coluna e o número em outra.",
            "Marque a **coluna com o VALOR numérico** (opção já ativada).",
            "Confira as colunas de Código e Preço — o layout varia entre abas.",
            "Preencha o IPI (%) à mão, se houver.",
        ],
        "observacao": "Layout irregular: confirme código/preço manualmente antes de processar.",
    },
    "Outro / configurar manualmente": {
        "manual": True,
        "passos": [
            "Fornecedor sem receita pronta: configure os campos manualmente.",
            "Escolha a aba, a linha do cabeçalho, as colunas de código e preço.",
            "Marque as opções de compatibilização conforme a planilha.",
        ],
    },
}

# ─── Header ───────────────────────────────────────────────────────────────────
st.title("💰 Atualização de Custos")
_data_doit = data_atualizacao_doit()
_txt_data = f" · atualizada em {_data_doit}" if _data_doit else ""
st.caption(
    f"Base DOit: {len(df_doit):,} produtos · {df_doit['Fabricante'].nunique()} fabricantes{_txt_data}"
)


def aplicar_receita(rec, abas_disponiveis):
    """Grava no session_state os valores dos widgets conforme a receita."""
    if rec.get("abas"):
        # Lista de abas de produto (ex: Accord): seleciona só as que existem no
        # arquivo, ignorando abas informativas (acabamento, política, etc.).
        # Match tolerante a espaços no nome (ex: "Plafon " na planilha).
        _alvos = {str(a).strip().lower() for a in rec["abas"]}
        selecao = [a for a in abas_disponiveis if str(a).strip().lower() in _alvos]
        st.session_state["k_abas"] = selecao or list(abas_disponiveis)
    elif rec.get("aba") and rec["aba"] in abas_disponiveis:
        st.session_state["k_abas"] = [rec["aba"]]
    else:
        st.session_state["k_abas"] = list(abas_disponiveis)
    if "header" in rec:
        st.session_state["k_header"] = int(rec["header"])
    st.session_state["k_normalizar"] = bool(rec.get("normalizar", False))
    st.session_state["k_concatenar"] = bool(rec.get("concatenar", False))
    st.session_state["k_usar_valor"] = bool(rec.get("usar_valor", False))
    st.session_state["k_ipi_prod"] = bool(rec.get("ipi_por_produto", False))
    st.session_state["k_agregar"] = bool(rec.get("agregar_acabamentos", False))
    st.session_state["k_multichave"] = bool(rec.get("chaves_multiplas", False))
    st.session_state["k_preco_liquido"] = bool(rec.get("preco_e_liquido", False))
    st.session_state["k_dois_acab"] = bool(rec.get("dois_acabamentos", False))
    if "ipi_fixo" in rec:
        st.session_state["k_ipi"] = float(rec["ipi_fixo"])
    else:
        # Receitas sem IPI (ex: Accord) zeram o campo para não aplicar nada
        st.session_state["k_ipi"] = 0.0
    st.session_state["_receita_cols"] = {
        "codigo": rec.get("codigo"),
        "preco": rec.get("preco"),
        "preco_tingida": rec.get("preco_tingida"),
        "segunda_col": rec.get("segunda_col"),
        "ipi_col": rec.get("ipi_col"),
    }
    st.session_state["_receita_margem"] = rec.get("margem_fixa")
    st.session_state["_receita_fab_id"] = rec.get("fab_id")
    st.session_state["_receita_aplicada"] = True


# ─── Passo 1 · Fornecedor e planilha ──────────────────────────────────────────
st.divider()
st.markdown(
    '<div class="section-header"><span class="step-badge">1</span> Fornecedor e planilha</div>',
    unsafe_allow_html=True,
)

nomes_receitas = list(RECEITAS.keys())
col_forn, col_up = st.columns([1, 1.3])

with col_forn:
    fornecedor_receita = st.selectbox(
        "Fornecedor",
        options=nomes_receitas,
        index=len(nomes_receitas) - 1,  # padrão: "Outro / configurar manualmente"
        key="k_fornecedor",
        help="Ao escolher um fornecedor conhecido, o app já preenche aba, colunas, "
             "opções e o fabricante no DOit automaticamente.",
    )

with col_up:
    arquivo_fornecedor = st.file_uploader(
        "Planilha do fornecedor (.xlsx, .xls)",
        type=["xlsx", "xls"],
    )

receita = RECEITAS[fornecedor_receita]


def _limpar_estado_planilha():
    """Remove do session_state tudo que depende da planilha enviada.

    Chamado quando o usuário tira a planilha do uploader, para que a próxima
    comece do zero (sem abas, colunas, opções ou resultado da anterior)."""
    chaves = [
        # Widgets de configuração (abas, cabeçalho, colunas e opções)
        "k_abas", "k_header", "k_ipi",
        "k_normalizar", "k_concatenar", "k_usar_valor", "k_ipi_prod",
        "k_agregar", "k_multichave", "k_preco_liquido", "k_dois_acab", "k_soltos",
        # Estado da receita aplicada
        "_receita_cols", "_receita_margem", "_receita_fab_id",
        "_receita_aplicada", "_receita_aplicada_para",
        # Resultado do processamento
        "df_conferir", "df_merge_raw", "df_forn_valido", "df_precisam_criar",
        "ipi_fixo", "usar_ipi_por_produto", "preco_e_liquido", "dois_acabamentos",
        "margem_fixa", "processado",
    ]
    for k in chaves:
        st.session_state.pop(k, None)


if arquivo_fornecedor is None:
    # Se havia uma planilha carregada antes, limpar todo o estado para a próxima
    # começar zerada (sem herdar abas/colunas/opções/resultado da anterior).
    if st.session_state.get("_tinha_arquivo"):
        _limpar_estado_planilha()
        st.session_state["_tinha_arquivo"] = False
        st.session_state["_nome_arquivo_atual"] = None
        st.rerun()
    st.info("Escolha o fornecedor e envie a planilha para começar.")
    st.stop()

# Marcar que há planilha carregada (habilita a limpeza quando ela for removida)
st.session_state["_tinha_arquivo"] = True

# Troca direta de planilha (sem passar pelo estado vazio): se o nome do arquivo
# mudou, limpar o estado da anterior para a nova começar zerada.
if st.session_state.get("_nome_arquivo_atual") != arquivo_fornecedor.name:
    if st.session_state.get("_nome_arquivo_atual") is not None:
        _limpar_estado_planilha()
        st.session_state["_tinha_arquivo"] = True
    st.session_state["_nome_arquivo_atual"] = arquivo_fornecedor.name
    st.rerun()

# ─── Leitura de abas ─────────────────────────────────────────────────────────
todas_abas = pd.read_excel(arquivo_fornecedor, header=None, sheet_name=None)
nomes_abas = list(todas_abas.keys())

# Aplicar a receita automaticamente quando o fornecedor selecionado ainda não
# foi aplicado (funciona em qualquer ordem: escolher fornecedor antes ou depois
# de enviar a planilha). Só aqui temos as abas disponíveis.
if (
    not receita.get("manual")
    and st.session_state.get("_receita_aplicada_para") != fornecedor_receita
):
    aplicar_receita(receita, nomes_abas)
    st.session_state["_receita_aplicada_para"] = fornecedor_receita
    st.rerun()

# Instruções do fornecedor num expander compacto (não polui a tela)
if receita.get("passos"):
    with st.expander(f"📋 Como preparar a {fornecedor_receita}", expanded=False):
        for p in receita["passos"]:
            st.markdown(f"- {p}")
        if receita.get("observacao"):
            st.warning(receita["observacao"])
        if not receita.get("manual"):
            if st.button("Aplicar configuração deste fornecedor", key="btn_reaplicar"):
                aplicar_receita(receita, nomes_abas)
                st.rerun()

st.caption(
    f"Arquivo: {arquivo_fornecedor.name} · "
    f"{len(nomes_abas)} aba{'s' if len(nomes_abas) > 1 else ''} ({', '.join(nomes_abas)})"
)

# ─── Passo 2 · Conferir e processar ───────────────────────────────────────────
st.divider()
st.markdown(
    '<div class="section-header"><span class="step-badge">2</span> Confira e processe</div>',
    unsafe_allow_html=True,
)

# ─── Seleção das abas a processar ─────────────────────────────────────────────
# Alguns fornecedores (ex: Rosa Maria) têm várias abas, e só uma contém os
# produtos (as outras são markup, descritivo, etc.). Por isso o usuário escolhe
# quais abas processar. Por padrão, todas ficam selecionadas (mantém o
# comportamento de fornecedores que concatenam múltiplas abas de produtos).
if "k_abas" not in st.session_state:
    st.session_state["k_abas"] = list(nomes_abas)
# Remover abas que não existem neste arquivo (ex: ao trocar de planilha)
st.session_state["k_abas"] = [a for a in st.session_state["k_abas"] if a in nomes_abas] or list(nomes_abas)

abas_escolhidas = st.multiselect(
    "Abas a processar (a 1ª selecionada define as colunas)",
    options=nomes_abas,
    key="k_abas",
    help="Se a planilha tiver abas que não são de produtos (ex: Rosa Maria: 'Markup', "
         "'DESCRITIVO'), deixe marcada apenas a aba com a tabela de produtos "
         "(ex: 'Tabela 2026').",
)

if not abas_escolhidas:
    st.warning("Selecione ao menos uma aba para processar.")
    st.stop()

df_raw = todas_abas[abas_escolhidas[0]]

col1, col2 = st.columns(2)

with col1:
    if "k_header" not in st.session_state:
        st.session_state["k_header"] = 1
    linha_header = st.number_input(
        "Linha do cabeçalho (0 = primeira linha)",
        min_value=0,
        max_value=max(len(df_raw) - 1, 0),
        key="k_header",
        help="Indique em qual linha estão os nomes das colunas",
    )

# Recarregar as abas selecionadas com header correto e concatenar
dfs_abas = []
df_primeira = pd.read_excel(arquivo_fornecedor, header=int(linha_header), sheet_name=abas_escolhidas[0])
df_primeira.columns = [str(c).strip() for c in df_primeira.columns]
df_primeira["_aba_origem"] = abas_escolhidas[0]
dfs_abas.append(df_primeira)

colunas_base = df_primeira.columns.drop("_aba_origem")

for nome_aba in abas_escolhidas[1:]:
    df_aba = pd.read_excel(arquivo_fornecedor, header=None, sheet_name=nome_aba)
    if len(df_aba.columns) >= len(colunas_base):
        df_aba = df_aba.iloc[:, :len(colunas_base)]
        df_aba.columns = colunas_base
    else:
        df_aba.columns = colunas_base[:len(df_aba.columns)]
    df_aba["_aba_origem"] = nome_aba
    dfs_abas.append(df_aba)

df_forn = pd.concat(dfs_abas, ignore_index=True)
colunas_forn = [c for c in df_forn.columns if c != "_aba_origem"]

with col2:
    nome_fornecedor = st.text_input(
        "Nome do fornecedor (para o relatório)",
        value=arquivo_fornecedor.name.split(".")[0],
    )

# Colunas sugeridas pela receita (aplicadas após o cabeçalho ser lido)
_rc = st.session_state.get("_receita_cols", {}) if st.session_state.get("_receita_aplicada") else {}


def _indice_col(nome_alvo, colunas, padrao=0):
    """Índice de 'nome_alvo' em 'colunas' (comparando sem quebras/espaços). Senão, padrao."""
    if not nome_alvo:
        return padrao
    alvo = str(nome_alvo).strip().replace("\n", " ").lower()
    for i, c in enumerate(colunas):
        if str(c).strip().replace("\n", " ").lower() == alvo:
            return i
    return padrao


# Lista de fabricantes do DOit (para o seletor abaixo)
fabricantes_doit = (
    df_doit[["Id do Fabricante", "Fabricante"]]
    .dropna(subset=["Id do Fabricante"])
    .drop_duplicates()
    .sort_values("Fabricante")
    .reset_index(drop=True)
)
fabricantes_doit["_label"] = (
    fabricantes_doit["Fabricante"].astype(str)
    + " (ID: "
    + fabricantes_doit["Id do Fabricante"].astype(int).astype(str)
    + ")"
)
_labels_fab = fabricantes_doit["_label"].tolist()
_idx_fab = 0
_fab_id_receita = st.session_state.get("_receita_fab_id")
if _fab_id_receita is not None:
    _match = fabricantes_doit.index[
        fabricantes_doit["Id do Fabricante"].astype("Int64") == int(_fab_id_receita)
    ].tolist()
    if _match:
        _idx_fab = int(_match[0])

# Campos essenciais numa linha só: Código · Preço · IPI · Fabricante
col3, col4, col5 = st.columns([1, 1, 0.7])
with col3:
    col_codigo = st.selectbox(
        "Coluna do CÓDIGO",
        options=colunas_forn,
        index=_indice_col(_rc.get("codigo"), colunas_forn, 0),
    )
with col4:
    col_preco = st.selectbox(
        "Coluna do PREÇO",
        options=colunas_forn,
        index=_indice_col(_rc.get("preco"), colunas_forn, min(1, len(colunas_forn) - 1)),
    )
with col5:
    if "k_ipi" not in st.session_state:
        st.session_state["k_ipi"] = 0.0
    ipi = st.number_input("IPI (%)", min_value=0.0, max_value=100.0, step=0.25, key="k_ipi")

fabricante_escolhido = st.selectbox(
    "Fabricante no DOit",
    options=_labels_fab,
    index=_idx_fab,
    help="Fabricante cadastrado no DOit correspondente a esta planilha. "
         "Ao escolher um fornecedor conhecido, já vem pré-selecionado.",
)
idx_selecionado = fabricantes_doit["_label"].tolist().index(fabricante_escolhido)
id_fabricante = fabricantes_doit.iloc[idx_selecionado]["Id do Fabricante"]

# ─── Opções avançadas (recolhidas) ────────────────────────────────────────────
# A receita já marca as opções certas; o expander fica fechado por padrão e só
# quem configura manualmente precisa abrir.
for _k in ["k_normalizar", "k_concatenar", "k_usar_valor", "k_ipi_prod", "k_agregar", "k_multichave", "k_preco_liquido", "k_dois_acab"]:
    if _k not in st.session_state:
        st.session_state[_k] = False
# Match de códigos soltos vem ligado por padrão
if "k_soltos" not in st.session_state:
    st.session_state["k_soltos"] = True

col_valor_separado = None
col_ipi_produto = None
col_concat_segunda = None
col_preco_tingida = None

with st.expander("⚙️ Opções avançadas (normalmente já configuradas pelo fornecedor)", expanded=False):
    normalizar_codigos = st.checkbox(
        "Normalizar códigos (remover acabamentos e hifens)",
        key="k_normalizar",
        help="Ex: Revoluz envia 'RI-H54414-1-BFM OU PTO' e no DOit é 'RI-H54414-1'. "
             "Usar com: Revoluz, Revolux e similares.",
    )
    concatenar_colunas = st.checkbox(
        "Concatenar colunas para formar código",
        key="k_concatenar",
        help="Junta o código + a 1ª palavra de outra coluna (ex: ID + modelo → 84-385-2). "
             "Use quando o DOit guarda o código já concatenado.",
    )
    chaves_multiplas = st.checkbox(
        "Casar por 2 chaves (código OU modelo)",
        key="k_multichave",
        help="Ex: Spotline — o produto às vezes está no DOit pelo ID (9766) e às vezes "
             "pelo modelo (1362/1, início da descrição). Tenta casar por qualquer um dos "
             "dois. Usar com: Spotline.",
    )
    usar_col_valor = st.checkbox(
        "Valor em coluna separada",
        key="k_usar_valor",
        help="Ex: Golden Art tem 'R$' numa coluna e o número em outra. Usar com: Golden Art.",
    )
    usar_ipi_por_produto = st.checkbox(
        "IPI por produto (coluna da planilha)",
        key="k_ipi_prod",
        help="Ex: Stella traz o IPI de cada item na coluna '% IPI'. Usar com: Stella.",
    )
    agregar_acabamentos = st.checkbox(
        "Código com acabamentos em linhas (usar maior preço)",
        key="k_agregar",
        help="Ex: Rosa Maria repete a referência e deixa a linha seguinte em branco com "
             "o preço do outro acabamento. Usar com: Rosa Maria.",
    )
    preco_e_liquido = st.checkbox(
        "Preço já é o custo líquido (não aplicar o ×1,10 do passo 1)",
        key="k_preco_liquido",
        help="Quando o fornecedor já manda o custo líquido na planilha. Com esta "
             "opção o CUSTO LÍQUIDO = preço (sem o ×1,10) e o CUSTO = preço × (1 + IPI).",
    )
    dois_acabamentos = st.checkbox(
        "Dois acabamentos: Natural × Tingida (margem 10%, sem IPI)",
        key="k_dois_acab",
        help="Ex: Accord — cada referência tem 2 preços (Lâmina Natural e Lâmina "
             "Tingida). Os SKUs com 'NATURAL' no nome recebem o preço Natural; os de "
             "cor/tingido recebem o da Tingida. CUSTO = preço × 1,10 (margem, sem IPI). "
             "Se a Tingida estiver vazia, os SKUs de cor ficam sem atualizar.",
    )
    casar_codigos_soltos = st.checkbox(
        "Casar códigos soltos / sem padrão (recomendado)",
        key="k_soltos",
        help="Alguns produtos foram cadastrados no DOit sem o padrão do fornecedor "
             "(ex: '0001' em vez de 'SL-0001-BR'). Esta opção casa esses casos quando "
             "há um único produto compatível. Quando houver mais de um candidato, o "
             "item vai para a aba 'Conferir manualmente' (não é atualizado sozinho).",
    )

    col_valor_separado = None
    if usar_col_valor:
        col_valor_separado = st.selectbox(
            "Coluna com o VALOR numérico",
            options=colunas_forn,
            index=_indice_col(_rc.get("preco"), colunas_forn, 0),
        )

    col_ipi_produto = None
    if usar_ipi_por_produto:
        col_ipi_produto = st.selectbox(
            "Coluna com o % de IPI por produto",
            options=colunas_forn,
            index=_indice_col(_rc.get("ipi_col"), colunas_forn, 0),
            help="Ex: Stella → coluna '% IPI'. Aceita 9.75, '9,75%' ou 0,0975.",
        )

    col_concat_segunda = None
    if concatenar_colunas or chaves_multiplas:
        col_concat_segunda = st.selectbox(
            "Coluna do modelo / 2ª parte (usa a 1ª palavra)",
            options=colunas_forn,
            index=_indice_col(_rc.get("segunda_col"), colunas_forn, 0),
            help="Ex: Descrição='1362/1 PENDENTE...' → usa '1362/1'.",
        )

    col_preco_tingida = None
    if dois_acabamentos:
        col_preco_tingida = st.selectbox(
            "Coluna do preço da Lâmina TINGIDA",
            options=colunas_forn,
            index=_indice_col(_rc.get("preco_tingida"), colunas_forn, min(2, len(colunas_forn) - 1)),
            help="Ex: Accord → coluna 'Lâmina Tingida' (col. F). A coluna do PREÇO "
                 "acima deve apontar para a Lâmina Natural (col. E).",
        )

# A receita já foi aplicada aos widgets deste rerun; consumir a flag.
st.session_state["_receita_aplicada"] = False

with st.expander("👁️ Pré-visualizar a planilha (com cabeçalho aplicado)", expanded=False):
    # Converter para texto só na exibição: evita o erro de serialização do
    # Arrow quando uma coluna mistura tipos (ex: abas informativas trazem
    # texto e números na mesma coluna). Não afeta o processamento.
    _preview = df_forn[colunas_forn].head(10).astype(str).replace({"nan": "", "None": ""})
    st.dataframe(_preview, use_container_width=True)


# ─── Funções auxiliares ───────────────────────────────────────────────────────
def parse_preco(valor):
    if pd.isna(valor):
        return None
    # Quando a célula já é um número (int/float), o ponto é sempre decimal — não
    # aplicamos a heurística de milhar (que destruiria valores como 1394.275,
    # comuns na Accord por serem resultado de cálculo de markup).
    if isinstance(valor, (int, float)):
        v = float(valor)
        return v if v >= 0 else None
    s = str(valor).strip()
    s = s.replace("R$", "").replace("r$", "").strip()
    s = s.replace(" ", "")
    if not s:
        return None

    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif "." in s:
        if s.count(".") > 1:
            s = s.replace(".", "")
        else:
            partes = s.split(".")
            if len(partes[1]) == 3:
                s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return None


def parse_ipi(valor):
    """Converte o IPI de uma célula em percentual (ex: 9.75 -> 9.75).

    Aceita formatos como 9.75, '9,75', '9,75%' e frações como 0,0975 (→ 9.75).
    Retorna None quando não há valor válido.
    """
    if pd.isna(valor):
        return None
    s = str(valor).strip().replace("%", "").replace(" ", "")
    if not s or s.lower() == "nan":
        return None
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        v = float(s)
    except ValueError:
        return None
    if v < 0:
        return None
    # Se veio como fração (ex: 0,0975), converter para percentual
    if 0 < v < 1:
        v *= 100
    return v


# ─── Processamento ────────────────────────────────────────────────────────────
st.write("")
if st.button("▶️ Processar atualização", type="primary", use_container_width=True):
    # Limpar código do fornecedor
    df_forn["_codigo_limpo"] = df_forn[col_codigo].astype(str).str.strip()
    # Remover .0 de números inteiros lidos como float (ex: 9766.0 -> 9766)
    df_forn["_codigo_limpo"] = df_forn["_codigo_limpo"].str.replace(r"\.0$", "", regex=True)

    # Código com acabamentos em linhas (ex: Rosa Maria): a referência aparece só na
    # 1ª linha e a linha seguinte fica em branco com o preço do outro acabamento.
    # Preenche o código para baixo (ffill) para que ambas as linhas apontem para a
    # mesma referência. O "maior preço por referência" é resolvido no dedupe abaixo.
    if agregar_acabamentos:
        df_forn["_codigo_limpo"] = df_forn["_codigo_limpo"].replace(
            {"nan": pd.NA, "": pd.NA, "None": pd.NA}
        ).ffill()
        df_forn["_codigo_limpo"] = df_forn["_codigo_limpo"].fillna("").astype(str).str.strip()

    # IPI por produto (ex: Stella): guardar o IPI de cada linha
    if usar_ipi_por_produto and col_ipi_produto:
        df_forn["_ipi_produto"] = df_forn[col_ipi_produto].apply(parse_ipi)

    def _extrair_primeira_palavra(val):
        s = str(val).strip()
        return s.split()[0] if s and s != "nan" else ""

    # Concatenar colunas se ativado (ex: base antiga: ID + primeira palavra da descrição)
    if concatenar_colunas and col_concat_segunda:
        df_forn["_segunda_parte"] = df_forn[col_concat_segunda].apply(_extrair_primeira_palavra)
        # Concatenar: código + "-" + primeira palavra (com / trocado por -)
        df_forn["_codigo_limpo"] = (
            df_forn["_codigo_limpo"] + "-" + df_forn["_segunda_parte"].str.replace("/", "-", regex=False)
        )
        # Limpar casos onde uma das partes é vazia
        df_forn["_codigo_limpo"] = df_forn["_codigo_limpo"].str.strip("-")

    # Chave alternativa (ex: Spotline hoje): além do código principal (ID), guardar
    # o "modelo" = primeira palavra da 2ª coluna. O cruzamento tenta casar por
    # qualquer uma das duas chaves. Resolve bases onde o mesmo produto às vezes
    # está cadastrado pelo ID e às vezes pelo modelo.
    if chaves_multiplas and col_concat_segunda:
        df_forn["_codigo_alt"] = df_forn[col_concat_segunda].apply(_extrair_primeira_palavra)

    df_forn["_preco_limpo"] = df_forn[col_preco].apply(parse_preco)

    # Dois acabamentos (ex: Accord): guardar também o preço da Lâmina Tingida.
    # "_preco_limpo" = Natural (coluna do PREÇO) e "_preco_tingida" = Tingida.
    if dois_acabamentos and col_preco_tingida:
        df_forn["_preco_tingida"] = df_forn[col_preco_tingida].apply(parse_preco)

    # Remover linhas sem código ou preço válido. No modo dois acabamentos basta
    # ter o Natural; linhas sem Natural não têm referência de produto utilizável.
    df_forn_valido = df_forn.dropna(subset=["_preco_limpo"]).copy()
    df_forn_valido = df_forn_valido[
        (df_forn_valido["_codigo_limpo"] != "nan")
        & (df_forn_valido["_codigo_limpo"] != "")
    ]

    # Colunas extras a carregar no merge (além de código e preço)
    colunas_extra_merge = []
    if usar_ipi_por_produto and col_ipi_produto and "_ipi_produto" in df_forn_valido.columns:
        colunas_extra_merge.append("_ipi_produto")
    if dois_acabamentos and "_preco_tingida" in df_forn_valido.columns:
        colunas_extra_merge.append("_preco_tingida")

    # Rosa Maria: com o código preenchido para baixo, a mesma referência tem 2 preços
    # (um por acabamento). Ordenar por preço decrescente faz o dedupe (keep="first")
    # manter o MAIOR preço por referência — e o IPI da linha correspondente.
    if agregar_acabamentos:
        df_forn_valido = df_forn_valido.sort_values("_preco_limpo", ascending=False)

    # ─── Dois acabamentos: Natural × Tingida (ex: Accord) ─────────────────────
    # Cada referência tem 2 preços (Natural e Tingida). No DOit a referência vira
    # vários SKUs; decidimos por SKU: nome com "NATURAL" → preço Natural, senão
    # (cor/tingido) → preço Tingida. Se o preço do acabamento for nulo (ex:
    # Tingida vazia), o SKU não é atualizado (fica em "não atualizados").
    if dois_acabamentos:
        _cols_forn = ["_codigo_limpo", "_preco_limpo"]
        if "_preco_tingida" in df_forn_valido.columns:
            _cols_forn.append("_preco_tingida")
        # 1 linha por referência do fornecedor
        df_forn_unico = df_forn_valido.drop_duplicates(subset=["_codigo_limpo"], keep="first")

        # Cruzar com o DOit pela referência base (código puro, sem sufixo " | 394")
        df_merge = df_doit.merge(
            df_forn_unico[_cols_forn],
            left_on="_ref_base",
            right_on="_codigo_limpo",
            how="inner",
        )

        # Escolher o preço por SKU conforme o acabamento no nome do produto
        _nome_upper = df_merge["Nome"].astype(str).str.upper()
        _eh_natural = _nome_upper.str.contains("NATURAL", na=False)
        _preco_tin = (
            df_merge["_preco_tingida"] if "_preco_tingida" in df_merge.columns
            else pd.Series([None] * len(df_merge), index=df_merge.index)
        )
        # Natural → preço Natural; cor/tingido → preço Tingida
        df_merge["_preco_limpo"] = _preco_tin.where(~_eh_natural, df_merge["_preco_limpo"])

        # Remover SKUs sem preço para o seu acabamento (ex: cor sem Tingida):
        # ficam fora do merge e aparecem em "não atualizados".
        df_merge = df_merge[df_merge["_preco_limpo"].notna()].copy()

        # Produtos do fornecedor que NÃO existem no DOit (precisam ser criados)
        refs_doit = set(df_doit["_ref_base"].astype(str).str.strip())
        mask_nao_encontrado = ~df_forn_valido["_codigo_limpo"].isin(refs_doit)
        df_precisam_criar = df_forn_valido[mask_nao_encontrado].copy()

    # Normalização de códigos (se ativada)
    elif normalizar_codigos:
        import re

        # Acabamentos conhecidos (Revoluz e similares)
        _acabamentos = [
            "BCO", "BCX", "BFM", "CRT", "DSB", "CRG", "GRM", "MCX",
            "PTB", "PTO", "PTX", "VBE", "VCR", "VOC", "CORES",
            "BR", "PT", "DO", "CR", "CZ", "VD", "AM", "AZ", "LR", "MR",
        ]

        def remover_acabamento(codigo):
            """Remove sufixos de acabamento do código do fornecedor."""
            s = str(codigo).strip().upper()
            partes = s.split("-")
            while len(partes) > 1 and any(acab in partes[-1] for acab in _acabamentos + ["OU"]):
                partes.pop()
            return "-".join(partes)

        def normalizar(codigo):
            """Remove prefixos de fabricante, hifens, traços, barras e espaços para comparação."""
            s = str(codigo).strip().upper()
            # Remover prefixos comuns de fabricante (ex: SL-, SP-)
            s = re.sub(r"^[A-Z]{2,3}-", "", s)
            # Remover hifens, barras e espaços
            s = re.sub(r"[-/\s]", "", s)
            return s

        # Normalizar código do fornecedor: remover acabamento + remover hifens
        df_forn_valido["_codigo_norm"] = df_forn_valido["_codigo_limpo"].apply(
            lambda x: normalizar(remover_acabamento(x))
        )

        # Normalizar referências do DOit a partir da base (sem o sufixo " | 394")
        df_doit_norm = df_doit.copy()
        df_doit_norm["_ref_norm"] = df_doit_norm["_ref_base"].apply(normalizar)

        # Cruzar usando códigos normalizados
        df_forn_unico = df_forn_valido.drop_duplicates(subset=["_codigo_norm"], keep="first")

        df_merge = df_doit_norm.merge(
            df_forn_unico[["_codigo_norm", "_preco_limpo", "_codigo_limpo"] + colunas_extra_merge],
            left_on="_ref_norm",
            right_on="_codigo_norm",
            how="inner",
        )

        # Produtos do fornecedor que NÃO estão no DOit
        refs_doit_norm = set(df_doit_norm["_ref_norm"])
        mask_nao_encontrado = ~df_forn_valido["_codigo_norm"].isin(refs_doit_norm)
        df_precisam_criar = df_forn_valido[mask_nao_encontrado].copy()

    else:
        # Cruzar com Doit pela referência (usar apenas 1 preço por código do fornecedor)
        df_forn_unico = df_forn_valido.drop_duplicates(subset=["_codigo_limpo"], keep="first")

        df_merge = df_doit.merge(
            df_forn_unico[["_codigo_limpo", "_preco_limpo"] + colunas_extra_merge],
            left_on="_ref_base",
            right_on="_codigo_limpo",
            how="inner",
        )

        # Produtos do fornecedor que NÃO estão no DOit (precisam ser criados)
        refs_doit = set(df_doit["_ref_base"].astype(str).str.strip())
        mask_nao_encontrado = ~df_forn_valido["_codigo_limpo"].isin(refs_doit)
        df_precisam_criar = df_forn_valido[mask_nao_encontrado].copy()

    # ─── Match por núcleo (códigos soltos / sem padrão) ───────────────────────
    # Alguns produtos do DOit foram cadastrados sem o padrão do fornecedor
    # (ex: "0001" em vez de "SL-0001-BR"). Aqui tentamos casar o que sobrou em
    # "precisam ser criados" reduzindo os códigos ao "núcleo" (sem prefixo de
    # fabricante, sem hífens/barras/espaços). Só casa quando há UM único produto
    # compatível no fabricante selecionado; se houver 2+, vai para conferência.
    df_conferir = pd.DataFrame()
    # No modo dois acabamentos (Accord) o match por núcleo é pulado: ele aplicaria
    # só o preço Natural, sem a lógica de acabamento por SKU.
    if casar_codigos_soltos and not dois_acabamentos and not df_precisam_criar.empty:
        import re as _re

        def _nucleo(codigo):
            s = str(codigo).strip().upper()
            s = _re.sub(r"^[A-Z]{1,4}-", "", s)   # remove prefixo tipo SL-
            s = _re.sub(r"[-/\s]", "", s)          # remove separadores
            return s

        # Referências do DOit apenas do fabricante selecionado, com núcleo
        doit_fab = df_doit[df_doit["Id do Fabricante"] == id_fabricante].copy()
        doit_fab["_nucleo"] = doit_fab["_ref_base"].apply(_nucleo)

        # Núcleos que já foram usados no match principal (evita recasar)
        if not df_merge.empty:
            refs_ja = set(df_merge["_ref_base"].astype(str)) if "_ref_base" in df_merge.columns else set()
        else:
            refs_ja = set()

        # Agrupar candidatos por núcleo (excluindo os já casados)
        cand = doit_fab[~doit_fab["_ref_base"].astype(str).isin(refs_ja)]
        candidatos_por_nucleo = cand.groupby("_nucleo")

        linhas_novo_match = []   # viram atualizações (match único)
        linhas_conferir = []     # ambíguos (2+ candidatos)
        idx_resolvidos = []      # índices de df_precisam_criar que saíram

        tem_alt = "_codigo_alt" in df_precisam_criar.columns
        for idx, row in df_precisam_criar.iterrows():
            # Tenta casar pelo núcleo do código principal; se não achar e houver
            # chave alternativa (ex: modelo da Spotline), tenta por ela.
            grupo = None
            for chave in ([row["_codigo_limpo"]] + ([row.get("_codigo_alt")] if tem_alt else [])):
                nuc = _nucleo(chave) if chave is not None else ""
                if nuc and nuc in candidatos_por_nucleo.groups:
                    grupo = candidatos_por_nucleo.get_group(nuc)
                    break
            if grupo is None:
                continue
            refs_distintas = grupo["_ref_base"].nunique()
            if refs_distintas == 1:
                # Match único e seguro: gerar uma linha de merge para cada SKU dessa referência
                for _, prod in grupo.iterrows():
                    linha = prod.to_dict()
                    linha["_codigo_limpo"] = row["_codigo_limpo"]
                    linha["_preco_limpo"] = row["_preco_limpo"]
                    if "_ipi_produto" in df_precisam_criar.columns:
                        linha["_ipi_produto"] = row.get("_ipi_produto")
                    linha["_match_tipo"] = "núcleo"
                    linhas_novo_match.append(linha)
                idx_resolvidos.append(idx)
            else:
                # Ambíguo: registrar candidatos para conferência (NÃO atualiza)
                for _, prod in grupo.iterrows():
                    linhas_conferir.append({
                        "Código fornecedor": row["_codigo_limpo"],
                        "Preço fornecedor": row["_preco_limpo"],
                        "SKU (DOit)": prod.get("SKU"),
                        "# Referência (DOit)": prod.get("# Referência"),
                        "Nome (DOit)": prod.get("Nome"),
                        "Preço atual (DOit)": prod.get("Preço"),
                    })
                idx_resolvidos.append(idx)

        # Incorporar os matches por núcleo ao df_merge
        if linhas_novo_match:
            df_extra = pd.DataFrame(linhas_novo_match)
            df_merge = pd.concat([df_merge, df_extra], ignore_index=True)

        # Remover de "precisam ser criados" o que foi resolvido (casou ou foi p/ conferir)
        if idx_resolvidos:
            df_precisam_criar = df_precisam_criar.drop(index=idx_resolvidos)

        if linhas_conferir:
            df_conferir = pd.DataFrame(linhas_conferir)

    # Guardar no session_state
    st.session_state["df_conferir"] = df_conferir
    st.session_state["df_merge_raw"] = df_merge
    st.session_state["df_forn_valido"] = df_forn_valido
    st.session_state["df_precisam_criar"] = df_precisam_criar
    st.session_state["ipi_fixo"] = ipi
    st.session_state["usar_ipi_por_produto"] = bool(
        usar_ipi_por_produto and col_ipi_produto and "_ipi_produto" in df_merge.columns
    )
    st.session_state["preco_e_liquido"] = bool(preco_e_liquido)
    st.session_state["dois_acabamentos"] = bool(dois_acabamentos)
    st.session_state["margem_fixa"] = float(st.session_state.get("_receita_margem") or 10.0)
    st.session_state["processado"] = True

# ─── Resultados ───────────────────────────────────────────────────────────────
if st.session_state.get("processado", False):
    df_merge = st.session_state["df_merge_raw"]
    df_precisam_criar = st.session_state["df_precisam_criar"]
    df_conferir = st.session_state.get("df_conferir", pd.DataFrame())

    # Filtrar o merge para manter apenas produtos do fabricante selecionado
    if not df_merge.empty:
        df_merge = df_merge[df_merge["Id do Fabricante"] == id_fabricante].copy()

    # Todos os produtos desse fabricante no Doit
    df_fabricante_doit = df_doit[df_doit["Id do Fabricante"] == id_fabricante].copy()
    refs_atualizadas = set(df_merge["# Referência"].astype(str).str.strip()) if not df_merge.empty else set()
    df_nao_atualizados = df_fabricante_doit[
        ~df_fabricante_doit["# Referência"].isin(refs_atualizadas)
    ].copy()

    # ─── Cálculos ─────────────────────────────────────────────────────────────
    ipi_fixo = st.session_state.get("ipi_fixo", ipi)
    usar_ipi_col = st.session_state.get("usar_ipi_por_produto", False)

    preco_e_liquido_calc = st.session_state.get("preco_e_liquido", False)
    dois_acab_calc = st.session_state.get("dois_acabamentos", False)
    margem_calc = st.session_state.get("margem_fixa", 10.0)

    if not df_merge.empty:
        if dois_acab_calc:
            # Accord: SEM IPI. O preço do acabamento já escolhido é o custo líquido
            # e a margem (10%) gera o custo bruto. CUSTO = líquido × 1,10.
            df_merge["_custo_liquido"] = df_merge["_preco_limpo"].round(2)
            df_merge["_ipi_aplicado"] = 0.0
            df_merge["_custo_bruto"] = (
                df_merge["_custo_liquido"] * (1 + margem_calc / 100)
            ).round(2)
        else:
            if preco_e_liquido_calc:
                # Preço do fornecedor JÁ é o custo líquido (não aplica os 10%)
                df_merge["_custo_liquido"] = df_merge["_preco_limpo"].round(2)
            else:
                df_merge["_custo_liquido"] = (df_merge["_preco_limpo"] * 1.10).round(2)

            if usar_ipi_col and "_ipi_produto" in df_merge.columns:
                # IPI por produto: usa o IPI de cada linha; linhas sem IPI caem no IPI fixo
                ipi_linha = df_merge["_ipi_produto"].fillna(ipi_fixo)
                df_merge["_ipi_aplicado"] = ipi_linha
                df_merge["_custo_bruto"] = (
                    df_merge["_custo_liquido"] * (1 + ipi_linha / 100)
                ).round(2)
            else:
                df_merge["_ipi_aplicado"] = ipi_fixo
                df_merge["_custo_bruto"] = (
                    df_merge["_custo_liquido"] * (1 + ipi_fixo / 100)
                ).round(2)

    # ─── Métricas ─────────────────────────────────────────────────────────────
    st.divider()
    st.markdown('<div class="section-header">📊 Resultado</div>', unsafe_allow_html=True)

    n_conferir = 0 if df_conferir is None or df_conferir.empty else df_conferir["Código fornecedor"].nunique()
    col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
    col_m1.metric("✅ Atualizados", f"{len(df_merge):,}")
    col_m2.metric("🆕 Precisam ser criados", f"{len(df_precisam_criar):,}")
    col_m3.metric("⚠️ Não atualizados", f"{len(df_nao_atualizados):,}")
    col_m4.metric("🔎 Conferir manual", f"{n_conferir:,}")
    if dois_acab_calc:
        ipi_label = f"Sem IPI · margem {margem_calc:g}%"
    elif usar_ipi_col:
        ipi_label = "Por produto"
    else:
        ipi_label = f"{ipi_fixo}%"
    col_m5.metric("📦 IPI", ipi_label)

    if n_conferir:
        st.warning(
            f"{n_conferir} código(s) do fornecedor bateram com mais de um produto no DOit "
            "e **não** foram atualizados automaticamente. Confira na aba "
            "'🔎 Conferir manualmente' e ajuste esses no DOit."
        )

    # ─── Modelo Custo ─────────────────────────────────────────────────────────
    # Data como datetime nativo (não string) para o Excel reconhecer como data.
    hoje_dt = pd.Timestamp(date.today())

    def _texto_id(serie):
        """SKU/FORNECEDOR como texto, sem '.0' e preservando zeros à esquerda."""
        return (
            serie.astype(str)
            .str.replace(r"\.0$", "", regex=True)
            .str.strip()
            .replace({"nan": "", "None": ""})
        )

    if not df_merge.empty:
        df_modelo_custo = pd.DataFrame(
            {
                # Texto para não perder zeros à esquerda (a coluna vai formatada como texto no Excel)
                "SKU": _texto_id(df_merge["SKU"]),
                "FORNECEDOR": _texto_id(df_merge["Id do Fabricante"]),
                "NOME ORIGINAL": "",
                "PART NUMBER": "",
                "CONDIÇÃO": "DDP",
                # Números nativos (float) para o Excel tratar como número
                "CUSTO": df_merge["_custo_bruto"].astype(float).round(2),
                "MOEDA": "BRL",
                "CUSTO FINAL?": "",
                "CUSTO LÍQUIDO": df_merge["_custo_liquido"].astype(float).round(2),
                "MODIFICADO EM": hoje_dt,
            }
        )
    else:
        df_modelo_custo = pd.DataFrame()

    # ─── Relatório ────────────────────────────────────────────────────────────
    if not df_merge.empty:
        df_produtos_doit = df_merge.drop(
            columns=["_codigo_limpo", "_preco_limpo", "_preco_tingida", "_custo_liquido", "_custo_bruto", "_codigo_norm", "_ref_norm", "_ref_base", "_segunda_parte", "_codigo_alt", "_match_tipo", "_ipi_produto", "_ipi_aplicado"],
            errors="ignore",
        )
    else:
        df_produtos_doit = pd.DataFrame()

    df_forn_valido = st.session_state["df_forn_valido"]
    df_produtos_forn = df_forn_valido.drop(
        columns=["_codigo_limpo", "_preco_limpo", "_preco_tingida", "_aba_origem", "_codigo_norm", "_segunda_parte", "_codigo_alt", "_match_tipo", "_ipi_produto", "_ipi_aplicado"], errors="ignore"
    )

    df_criar_saida = df_precisam_criar.drop(
        columns=["_codigo_limpo", "_preco_limpo", "_preco_tingida", "_aba_origem", "_codigo_norm", "_segunda_parte", "_codigo_alt", "_match_tipo", "_ipi_produto", "_ipi_aplicado"], errors="ignore"
    )

    # ─── Tabs de visualização ─────────────────────────────────────────────────
    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        ["📋 Modelo Custo", "✅ Atualizados", "🆕 Precisam ser criados",
         "⚠️ Não atualizados", "🔎 Conferir manualmente"]
    )

    with tab1:
        if not df_modelo_custo.empty:
            # Versão só para exibição: data como dd/mm/aaaa (sem horário)
            df_modelo_view = df_modelo_custo.copy()
            if "MODIFICADO EM" in df_modelo_view.columns:
                df_modelo_view["MODIFICADO EM"] = pd.to_datetime(
                    df_modelo_view["MODIFICADO EM"]
                ).dt.strftime("%d/%m/%Y")
            st.dataframe(df_modelo_view, use_container_width=True, height=400)
        else:
            st.info("Nenhum produto para atualizar.")

    with tab2:
        if not df_produtos_doit.empty:
            st.dataframe(
                df_produtos_doit[["SKU", "# Referência", "Nome", "Preço"]],
                use_container_width=True,
                height=400,
            )
        else:
            st.info("Nenhum produto atualizado.")

    with tab3:
        if not df_criar_saida.empty:
            st.dataframe(df_criar_saida, use_container_width=True, height=400)
        else:
            st.info("Todos os produtos do fornecedor já existem no DOit.")

    with tab4:
        if not df_nao_atualizados.empty:
            st.dataframe(
                df_nao_atualizados[["SKU", "# Referência", "Nome", "Preço"]],
                use_container_width=True,
                height=400,
            )
        else:
            st.info("Todos os produtos do fabricante foram atualizados.")

    with tab5:
        if df_conferir is not None and not df_conferir.empty:
            st.caption(
                "Estes códigos do fornecedor bateram com mais de um produto no DOit. "
                "Como o preço pode variar por acabamento, o app não escolhe sozinho — "
                "confira e atualize manualmente no DOit o SKU correto."
            )
            st.dataframe(df_conferir, use_container_width=True, height=400)
        else:
            st.info("Nenhum caso ambíguo para conferir. 👍")

    # ─── Texto para o cliente ─────────────────────────────────────────────────
    st.divider()
    st.markdown('<div class="section-header">💬 Texto para o cliente</div>', unsafe_allow_html=True)

    texto_padrao = (
        f"Referente a {nome_fornecedor}, os custos foram atualizados:\n\n"
        f"Produtos que foram atualizados: {len(df_merge)}\n"
        f"Produtos que precisam ser criados: {len(df_precisam_criar)}\n"
        f"Produtos que não foram atualizados: {len(df_nao_atualizados)}\n"
        f"Atualizado na Luminata 1 e 2."
    )

    st.text_area("Copie e envie ao cliente:", value=texto_padrao, height=140, label_visibility="collapsed")

    # ─── Downloads ────────────────────────────────────────────────────────────
    st.divider()
    st.markdown('<div class="section-header">📥 Downloads</div>', unsafe_allow_html=True)

    col_dl1, col_dl2 = st.columns(2)

    def escrever_modelo_custo(writer, df_modelo, sheet_name="Modelo Custo"):
        """Escreve a aba Modelo Custo aplicando os formatos que o DOit espera:
        SKU/FORNECEDOR como texto (preserva zero à esquerda), CUSTO/CUSTO LÍQUIDO
        como número e MODIFICADO EM como data."""
        df_modelo.to_excel(writer, index=False, sheet_name=sheet_name)
        if df_modelo.empty:
            return
        wb = writer.book
        ws = writer.sheets[sheet_name]
        fmt_texto = wb.add_format({"num_format": "@"})           # texto
        fmt_num = wb.add_format({"num_format": "0.00"})           # número 2 casas
        fmt_data = wb.add_format({"num_format": "dd/mm/yyyy"})    # data dd/mm/aaaa
        cols = list(df_modelo.columns)

        def _ci(nome):
            return cols.index(nome) if nome in cols else None

        for nome in ["SKU", "FORNECEDOR"]:
            i = _ci(nome)
            if i is not None:
                ws.set_column(i, i, 14, fmt_texto)
        for nome in ["CUSTO", "CUSTO LÍQUIDO"]:
            i = _ci(nome)
            if i is not None:
                ws.set_column(i, i, 12, fmt_num)
        i = _ci("MODIFICADO EM")
        if i is not None:
            ws.set_column(i, i, 14, fmt_data)

    with col_dl1:
        buffer1 = BytesIO()
        with pd.ExcelWriter(buffer1, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as writer:
            escrever_modelo_custo(writer, df_modelo_custo)
        buffer1.seek(0)

        st.download_button(
            label="📥 Importação DOit",
            data=buffer1,
            file_name=f"custo_doit_{nome_fornecedor}_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col_dl2:
        buffer2 = BytesIO()
        with pd.ExcelWriter(buffer2, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as writer:
            if not df_produtos_doit.empty:
                df_produtos_doit.to_excel(writer, index=False, sheet_name="Produtos DOit")
            df_produtos_forn.to_excel(writer, index=False, sheet_name="Produtos")
            escrever_modelo_custo(writer, df_modelo_custo)
            if not df_criar_saida.empty:
                df_criar_saida.to_excel(writer, index=False, sheet_name="Precisam ser criados")
            if not df_nao_atualizados.empty:
                df_nao_atualizados.to_excel(
                    writer, index=False, sheet_name="Não foram atualizados"
                )
            if df_conferir is not None and not df_conferir.empty:
                df_conferir.to_excel(writer, index=False, sheet_name="Conferir manualmente")
        buffer2.seek(0)

        st.download_button(
            label="📥 Relatório Cliente",
            data=buffer2,
            file_name=f"relatorio_{nome_fornecedor}_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # ─── Registro para planilha de controle ───────────────────────────────────
    st.divider()
    st.markdown('<div class="section-header">📝 Registro (planilha de controle)</div>', unsafe_allow_html=True)

    st.link_button(
        "🔗 Abrir planilha de controle (Google Sheets)",
        URL_PLANILHA_CONTROLE,
        use_container_width=True,
    )

    col_reg1, col_reg2 = st.columns(2)

    with col_reg1:
        responsavel = st.text_input("Responsável", value="Isabela Silva")
        caracteristicas = st.text_input("Características do fornecedor", value="", placeholder="Ex: Muitos acabamentos, cód diferente...")
        dificuldade = st.selectbox("Dificuldade", options=["Fácil", "Médio", "Fácil/Médio", "Médio/Difícil", "Difícil"])

    with col_reg2:
        if usar_ipi_col:
            ipi_texto = "Por produto"
        else:
            ipi_texto = f"{ipi_fixo}" if ipi_fixo > 0 else "Não tem"
        muitas_abas = "TRUE" if len(nomes_abas) > 1 else "FALSE"
        formulas_usadas = st.text_input("Fórmulas/Observações", value="", placeholder="Ex: Procv e ses junto")
        colunas_usar = st.text_input("Colunas usadas", value="", placeholder="Ex: Referência, Prç.ven, ICMS e IPI")

    # Linha formatada para colar no Sheets (tab-separated)
    linha_sheets = (
        f"{nome_fornecedor}\t"
        f"{len(df_merge)}\t"
        f"{len(df_precisam_criar)}\t"
        f"{len(df_nao_atualizados)}\t"
        f"{date.today().strftime('%d/%m/%Y')}\t"
        f"{responsavel}\t"
        f"{caracteristicas}\t"
        f"{ipi_texto}\t"
        f"{dificuldade}\t"
        f"{formulas_usadas}\t"
        f"{colunas_usar}\t"
        f"{muitas_abas}"
    )

    st.text_area(
        "Copie e cole na planilha de controle (Sheets):",
        value=linha_sheets,
        height=80,
        help="Selecione tudo, copie (Ctrl+C), abra a planilha pelo botão acima e "
             "cole na próxima linha vazia do Google Sheets.",
    )

    st.caption(f"📋 Planilha de controle: {URL_PLANILHA_CONTROLE}")
