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
    .main .block-container { padding-top: 1rem; max-width: 1200px; }
    h1 { color: #2c3e50; }
    .stTabs [data-baseweb="tab-list"] { gap: 1.5rem; }
    .section-header {
        font-size: 1.1rem;
        font-weight: 600;
        color: #2c3e50;
        margin-bottom: 0.5rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    .info-box {
        background: #eef6ff;
        border-left: 4px solid #3b82f6;
        border-radius: 6px;
        padding: 0.8rem 1rem;
        font-size: 0.85rem;
        color: #1e40af;
        margin-bottom: 1rem;
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
        "concatenar": True,
        "segunda_col": "DESCRIÇÃO",
        "fab_id": 7629,
        "passos": [
            "Cabeçalho na linha 0.",
            "Código = coluna **ID** · Preço = coluna **PREÇO**.",
            "O código do DOit junta ID + 1ª palavra da **DESCRIÇÃO**: concatenação já marcada.",
            "Preencha o IPI (%) à mão, se houver.",
        ],
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
_txt_data = f' | 🗓️ atualizada em <strong>{_data_doit}</strong>' if _data_doit else ""
st.markdown(f'<div class="info-box">📦 Base DOit carregada com <strong>{len(df_doit):,}</strong> produtos | {df_doit["Fabricante"].nunique()} fabricantes{_txt_data}</div>', unsafe_allow_html=True)

# ─── Upload ───────────────────────────────────────────────────────────────────
st.divider()
st.markdown('<div class="section-header">📁 Upload da planilha do fornecedor</div>', unsafe_allow_html=True)

arquivo_fornecedor = st.file_uploader(
    "Arraste ou selecione a planilha (.xlsx, .xls)",
    type=["xlsx", "xls"],
)

if arquivo_fornecedor is None:
    st.stop()

# ─── Leitura de abas ─────────────────────────────────────────────────────────
todas_abas = pd.read_excel(arquivo_fornecedor, header=None, sheet_name=None)
nomes_abas = list(todas_abas.keys())

st.markdown(
    f'<div class="info-box">📑 Arquivo: <strong>{arquivo_fornecedor.name}</strong> — '
    f'{len(nomes_abas)} aba{"s" if len(nomes_abas) > 1 else ""} encontrada{"s" if len(nomes_abas) > 1 else ""} '
    f'({", ".join(nomes_abas)})</div>',
    unsafe_allow_html=True,
)

# ─── Seletor de fornecedor (receita) ──────────────────────────────────────────
# A pessoa escolhe o fornecedor e o app pré-configura aba, cabeçalho, colunas e
# opções. Isso deixa claro para outras equipes o que usar em cada fornecedor.
st.markdown('<div class="section-header">🏷️ Qual é o fornecedor?</div>', unsafe_allow_html=True)

nomes_receitas = list(RECEITAS.keys())
fornecedor_receita = st.selectbox(
    "Escolha o fornecedor para pré-configurar as opções",
    options=nomes_receitas,
    index=len(nomes_receitas) - 1,  # padrão: "Outro / configurar manualmente"
    help="Selecionar um fornecedor conhecido preenche automaticamente a aba, o "
         "cabeçalho, as colunas e as opções corretas.",
)

receita = RECEITAS[fornecedor_receita]

# Passo a passo da receita
passos_html = "".join(f"<li>{p}</li>" for p in receita.get("passos", []))
obs = receita.get("observacao")
obs_html = f'<div style="margin-top:0.4rem;color:#b45309;">⚠️ {obs}</div>' if obs else ""
st.markdown(
    f'<div class="info-box"><strong>Passo a passo — {fornecedor_receita}</strong>'
    f'<ol style="margin:0.4rem 0 0 1rem;padding:0;">{passos_html}</ol>{obs_html}</div>',
    unsafe_allow_html=True,
)


def aplicar_receita(rec, abas_disponiveis):
    """Grava no session_state os valores dos widgets conforme a receita."""
    # Abas
    if rec.get("aba") and rec["aba"] in abas_disponiveis:
        st.session_state["k_abas"] = [rec["aba"]]
    else:
        st.session_state["k_abas"] = list(abas_disponiveis)
    # Linha do cabeçalho
    if "header" in rec:
        st.session_state["k_header"] = int(rec["header"])
    # Opções (marcar só as da receita; desmarcar as demais)
    st.session_state["k_normalizar"] = bool(rec.get("normalizar", False))
    st.session_state["k_concatenar"] = bool(rec.get("concatenar", False))
    st.session_state["k_usar_valor"] = bool(rec.get("usar_valor", False))
    st.session_state["k_ipi_prod"] = bool(rec.get("ipi_por_produto", False))
    st.session_state["k_agregar"] = bool(rec.get("agregar_acabamentos", False))
    # Colunas alvo (guardadas para aplicar após ler o cabeçalho)
    st.session_state["_receita_cols"] = {
        "codigo": rec.get("codigo"),
        "preco": rec.get("preco"),
        "segunda_col": rec.get("segunda_col"),
        "ipi_col": rec.get("ipi_col"),
    }
    # Fabricante do DOit sugerido pela receita (casado por ID)
    st.session_state["_receita_fab_id"] = rec.get("fab_id")
    st.session_state["_receita_aplicada"] = True


if not receita.get("manual"):
    if st.button(f"✨ Aplicar configuração de {fornecedor_receita}", use_container_width=True):
        aplicar_receita(receita, nomes_abas)
        st.rerun()

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
    "📑 Abas a processar (a 1ª selecionada define as colunas)",
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

with st.expander("👁️ Pré-visualização (primeiras 15 linhas da 1ª aba selecionada)", expanded=False):
    st.dataframe(df_raw.head(15), use_container_width=True)

# ─── Configuração ────────────────────────────────────────────────────────────
st.divider()
st.markdown('<div class="section-header">⚙️ Configuração</div>', unsafe_allow_html=True)

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
        "Nome do fornecedor",
        value=arquivo_fornecedor.name.split(".")[0],
    )

col3, col4, col5, col6 = st.columns(4)

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
    ipi = st.number_input("IPI (%)", min_value=0.0, max_value=100.0, value=0.0, step=0.25)

# ─── Opções avançadas ─────────────────────────────────────────────────────────
st.markdown('<div class="section-header">🔧 Opções de compatibilização</div>', unsafe_allow_html=True)

# Inicializa as flags de opção no session_state (padrão desmarcado)
for _k in ["k_normalizar", "k_concatenar", "k_usar_valor", "k_ipi_prod", "k_agregar"]:
    if _k not in st.session_state:
        st.session_state[_k] = False

col_opt1, col_opt2, col_opt3 = st.columns(3)

with col_opt1:
    normalizar_codigos = st.checkbox(
        "🔄 Normalizar códigos (remover acabamentos e hifens)",
        key="k_normalizar",
        help="Ex: Revoluz envia 'RI-H54414-1-BFM OU PTO' e no DOit é 'RI-H54414-1'. "
             "Remove sufixos de acabamento e hifens para compatibilizar. "
             "Usar com: Revoluz, Revolux e similares.",
    )

with col_opt2:
    concatenar_colunas = st.checkbox(
        "🔗 Concatenar colunas para formar código",
        key="k_concatenar",
        help="Ex: Spotline tem ID=84 e Descrição='385/2 PLAFON SMART...', no DOit é 'SL-84-385-2'. "
             "Junta o ID + primeira palavra da descrição para formar o código completo. "
             "Usar com: Spotline.",
    )

with col_opt3:
    usar_col_valor = st.checkbox(
        "💲 Valor em coluna separada",
        key="k_usar_valor",
        help="Ex: Golden Art tem 'R$' em uma coluna e o valor numérico em outra. "
             "Selecione a coluna com o número após ativar. "
             "Usar com: Golden Art.",
    )

col_opt4, col_opt5, _col_opt6 = st.columns(3)

with col_opt4:
    usar_ipi_por_produto = st.checkbox(
        "📊 IPI por produto (coluna da planilha)",
        key="k_ipi_prod",
        help="Ex: Stella traz o IPI de cada item na coluna '% IPI'. "
             "Ativa o uso do IPI de cada linha no lugar do IPI fixo acima. "
             "Linhas sem IPI usam o IPI fixo como fallback. "
             "Usar com: Stella.",
    )

with col_opt5:
    agregar_acabamentos = st.checkbox(
        "🧩 Código com acabamentos em linhas (usar maior preço)",
        key="k_agregar",
        help="Ex: Rosa Maria repete a referência numa linha e deixa a linha seguinte "
             "em branco com o preço do outro acabamento. Preenche o código para baixo "
             "e usa o MAIOR preço por referência. Usar com: Rosa Maria.",
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
        help="Ex: Stella → coluna '% IPI'. Aceita valores como 9.75, '9,75%' ou 0,0975.",
    )

col_concat_segunda = None
if concatenar_colunas:
    col_concat_segunda = st.selectbox(
        "Segunda coluna (será extraída a primeira palavra e concatenada ao código)",
        options=colunas_forn,
        index=_indice_col(_rc.get("segunda_col"), colunas_forn, 0),
        help="A primeira palavra desta coluna será unida ao código com hífen. Ex: ID=84, Descrição='385/2 PLAFON...' → 84-385-2",
    )

# A receita já foi aplicada aos widgets deste rerun; consumir a flag para que o
# usuário possa ajustar os campos manualmente sem que voltem ao valor da receita.
st.session_state["_receita_aplicada"] = False

# ─── Seleção do fabricante ────────────────────────────────────────────────────
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

# Pré-selecionar o fabricante sugerido pela receita (casando pelo ID do DOit)
_labels_fab = fabricantes_doit["_label"].tolist()
_idx_fab = 0
_fab_id_receita = st.session_state.get("_receita_fab_id")
if _fab_id_receita is not None:
    _match = fabricantes_doit.index[
        fabricantes_doit["Id do Fabricante"].astype("Int64") == int(_fab_id_receita)
    ].tolist()
    if _match:
        _idx_fab = int(_match[0])

fabricante_escolhido = st.selectbox(
    "🏭 Fabricante no DOit",
    options=_labels_fab,
    index=_idx_fab,
    help="Selecione o fabricante cadastrado no DOit correspondente a esta planilha. "
         "Ao aplicar a configuração de um fornecedor, ele já vem pré-selecionado.",
)

idx_selecionado = fabricantes_doit["_label"].tolist().index(fabricante_escolhido)
id_fabricante = fabricantes_doit.iloc[idx_selecionado]["Id do Fabricante"]

with st.expander("📋 Planilha do fornecedor (com cabeçalho aplicado)", expanded=False):
    st.dataframe(df_forn[colunas_forn].head(10), use_container_width=True)


# ─── Funções auxiliares ───────────────────────────────────────────────────────
def parse_preco(valor):
    if pd.isna(valor):
        return None
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
st.divider()
st.markdown('<div class="section-header">🔄 Processamento</div>', unsafe_allow_html=True)

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

    # Concatenar colunas se ativado (ex: Spotline: ID + primeira palavra da descrição)
    if concatenar_colunas and col_concat_segunda:
        def _extrair_primeira_palavra(val):
            s = str(val).strip()
            return s.split()[0] if s and s != "nan" else ""

        df_forn["_segunda_parte"] = df_forn[col_concat_segunda].apply(_extrair_primeira_palavra)
        # Concatenar: código + "-" + primeira palavra (com / trocado por -)
        df_forn["_codigo_limpo"] = (
            df_forn["_codigo_limpo"] + "-" + df_forn["_segunda_parte"].str.replace("/", "-", regex=False)
        )
        # Limpar casos onde uma das partes é vazia
        df_forn["_codigo_limpo"] = df_forn["_codigo_limpo"].str.strip("-")

    df_forn["_preco_limpo"] = df_forn[col_preco].apply(parse_preco)

    # Remover linhas sem código ou preço válido
    df_forn_valido = df_forn.dropna(subset=["_preco_limpo"]).copy()
    df_forn_valido = df_forn_valido[
        (df_forn_valido["_codigo_limpo"] != "nan")
        & (df_forn_valido["_codigo_limpo"] != "")
    ]

    # Colunas extras a carregar no merge (além de código e preço)
    colunas_extra_merge = []
    if usar_ipi_por_produto and col_ipi_produto and "_ipi_produto" in df_forn_valido.columns:
        colunas_extra_merge.append("_ipi_produto")

    # Rosa Maria: com o código preenchido para baixo, a mesma referência tem 2 preços
    # (um por acabamento). Ordenar por preço decrescente faz o dedupe (keep="first")
    # manter o MAIOR preço por referência — e o IPI da linha correspondente.
    if agregar_acabamentos:
        df_forn_valido = df_forn_valido.sort_values("_preco_limpo", ascending=False)

    # Normalização de códigos (se ativada)
    if normalizar_codigos:
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

    # Guardar no session_state
    st.session_state["df_merge_raw"] = df_merge
    st.session_state["df_forn_valido"] = df_forn_valido
    st.session_state["df_precisam_criar"] = df_precisam_criar
    st.session_state["ipi_fixo"] = ipi
    st.session_state["usar_ipi_por_produto"] = bool(
        usar_ipi_por_produto and col_ipi_produto and "_ipi_produto" in df_merge.columns
    )
    st.session_state["processado"] = True

# ─── Resultados ───────────────────────────────────────────────────────────────
if st.session_state.get("processado", False):
    df_merge = st.session_state["df_merge_raw"]
    df_precisam_criar = st.session_state["df_precisam_criar"]

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

    if not df_merge.empty:
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

    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    col_m1.metric("✅ Atualizados", f"{len(df_merge):,}")
    col_m2.metric("🆕 Precisam ser criados", f"{len(df_precisam_criar):,}")
    col_m3.metric("⚠️ Não atualizados", f"{len(df_nao_atualizados):,}")
    ipi_label = "Por produto" if usar_ipi_col else f"{ipi_fixo}%"
    col_m4.metric("📦 IPI", ipi_label)

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
            columns=["_codigo_limpo", "_preco_limpo", "_custo_liquido", "_custo_bruto", "_codigo_norm", "_ref_norm", "_ref_base", "_segunda_parte", "_ipi_produto", "_ipi_aplicado"],
            errors="ignore",
        )
    else:
        df_produtos_doit = pd.DataFrame()

    df_forn_valido = st.session_state["df_forn_valido"]
    df_produtos_forn = df_forn_valido.drop(
        columns=["_codigo_limpo", "_preco_limpo", "_aba_origem", "_codigo_norm", "_segunda_parte", "_ipi_produto", "_ipi_aplicado"], errors="ignore"
    )

    df_criar_saida = df_precisam_criar.drop(
        columns=["_codigo_limpo", "_preco_limpo", "_aba_origem", "_codigo_norm", "_segunda_parte", "_ipi_produto", "_ipi_aplicado"], errors="ignore"
    )

    # ─── Tabs de visualização ─────────────────────────────────────────────────
    tab1, tab2, tab3, tab4 = st.tabs(
        ["📋 Modelo Custo", "✅ Atualizados", "🆕 Precisam ser criados", "⚠️ Não atualizados"]
    )

    with tab1:
        if not df_modelo_custo.empty:
            st.dataframe(df_modelo_custo, use_container_width=True, height=400)
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
        fmt_data = wb.add_format({"num_format": "yyyy-mm-dd"})    # data ISO
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
        with pd.ExcelWriter(buffer1, engine="xlsxwriter") as writer:
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
        with pd.ExcelWriter(buffer2, engine="xlsxwriter") as writer:
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
