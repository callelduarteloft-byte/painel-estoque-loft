import streamlit as st
from supabase import create_client, Client
import pandas as pd
import numpy as np
import plotly.express as px
from openai import OpenAI
import re
import os
import glob
from pathlib import Path
from datetime import datetime

# ==========================================
# 1. CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(
    page_title="Painel Estratégico de Estoque - LOFT",
    layout="wide",
    page_icon="📦",
    initial_sidebar_state="expanded",
)

CATEGORIAS_IGNORADAS = ["SERVIÇOS", "USO PRÓPRIO", "VARIADOS"]


# ==========================================
# 2. FUNÇÕES DE FORMATAÇÃO E UTILITÁRIOS
# ==========================================
def formatar_inteiro(valor) -> str:
    try:
        return f"{int(round(float(valor))):,}".replace(",", ".")
    except Exception:
        return "0"


def formatar_moeda(valor) -> str:
    try:
        return "R$ " + f"{float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"


def normalizar_sku(serie: pd.Series) -> pd.Series:
    return serie.astype(str).str.strip().str.upper()


def gerar_arquivo_csv(df_exportacao: pd.DataFrame) -> bytes:
    return df_exportacao.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")


def ordenar_meses_cronologicamente(lista_meses: list[str]) -> list[str]:
    """Ordena meses no formato MM-AA (ex: '09-25', '10-25', '01-26') na ordem cronológica exata."""
    def chave_mes(m_str: str):
        try:
            partes = str(m_str).strip().split("-")
            mes = int(partes[0])
            ano = int(partes[1]) + 2000
            return (ano, mes)
        except Exception:
            return (0, 0)

    unicos = list(dict.fromkeys([str(m).strip() for m in lista_meses if pd.notna(m) and str(m).strip() != ""]))
    return sorted(unicos, key=chave_mes)


def encontrar_planilha_local() -> str | None:
    """
    Localiza automaticamente a planilha '5 - Análise Estoque.xlsx' tanto na pasta OneDrive oficial
    ('OneDrive - Gustavo - Pessoal / One Drive LOFT') quanto na pasta local do projeto,
    priorizando sempre o arquivo com a data de modificação mais recente.
    """
    candidatos_encontrados = []

    # 1. Procura nas pastas OneDrive do usuário (incluindo OneDrive LOFT)
    try:
        home = Path.home()
        for pasta_od in home.glob("OneDrive*"):
            if pasta_od.is_dir():
                for arq in pasta_od.rglob("*An*lise*Estoque*.xls*"):
                    if not arq.name.startswith("~$"):
                        candidatos_encontrados.append(arq)
    except Exception:
        pass

    # 2. Procura na pasta atual do projeto
    nomes_diretos = [
        "5 - Análise Estoque.xlsx",
        "5 - Analise Estoque.xlsx",
        "5 - Análise Estoque.xls",
    ]
    for nome in nomes_diretos:
        if os.path.exists(nome):
            candidatos_encontrados.append(Path(nome).resolve())

    for arq_str in glob.glob("*.xls*"):
        if not os.path.basename(arq_str).startswith("~$"):
            nome_low = arq_str.lower()
            if "estoque" in nome_low or "analise" in nome_low or "análise" in nome_low or "5" in nome_low:
                candidatos_encontrados.append(Path(arq_str).resolve())

    if candidatos_encontrados:
        mais_recente = max(candidatos_encontrados, key=lambda p: p.stat().st_mtime)
        return str(mais_recente)

    return None


def obter_cliente_supabase() -> Client | None:
    """Conecta ao Supabase usando st.secrets (Nuvem ou Local)."""
    try:
        url = st.secrets.get("SUPABASE_URL", "")
        key = st.secrets.get("SUPABASE_KEY", "")
        if url and key:
            return create_client(url, key)
    except Exception:
        pass
    return None


def buscar_tabela_completa_supabase(_client: Client, nome_tabela: str, tamanho_lote: int = 1000) -> pd.DataFrame:
    """
    Busca 100% dos registros de uma tabela do Supabase paginando em blocos de 1.000 linhas
    para contornar o limite padrão da API e trazer todos os +26.000 registros sem perdas.
    """
    todos_registros = []
    inicio = 0
    while True:
        res = (
            _client.table(nome_tabela)
            .select("*")
            .range(inicio, inicio + tamanho_lote - 1)
            .execute()
        )
        dados = res.data
        if not dados:
            break
        todos_registros.extend(dados)
        if len(dados) < tamanho_lote:
            break
        inicio += tamanho_lote
    return pd.DataFrame(todos_registros)


# ==========================================
# 3. EXTRAÇÃO PROFUNDA DA '5 - Análise Estoque.xlsx' E SUPABASE
# ==========================================
def ler_aba_curva_completa(xls: pd.ExcelFile, nome_aba: str) -> pd.DataFrame:
    df_raw = pd.read_excel(xls, sheet_name=nome_aba)
    if "Código" not in df_raw.columns:
        for idx, row in df_raw.head(5).iterrows():
            if "Código" in row.values:
                df_raw.columns = row.values
                df_raw = df_raw.iloc[idx + 1 :].reset_index(drop=True)
                break
    cols_validas = [c for c in df_raw.columns if pd.notna(c)]
    df_raw = df_raw[cols_validas]
    df_raw = df_raw.loc[:, ~df_raw.columns.duplicated()]

    mapa_colunas = {
        "Código": "codigo",
        "Produto": "produto",
        "Dt 1° Rec": "dt_1_rec",
        "Qt.Est.": "qt_est_rede",
        "Qt.Venda": "qt_venda",
        "Qt. SELL IN": "qt_sell_in",
        "Cap.Supri": "cap_sup_rede",
        "Val.custo": "val_custo_venda",
        "Val.Total": "val_total_venda",
        "Estoque Cd": "estoque_cd",
        "Cap. Sup. Cd": "cap_sup_cd",
        "Tabela atacado": "tabela_atacado",
        "Valor total estoque ": "val_total_estoque_curva",
        "Markup": "markup",
        "ABC": "curva_abc",
        "Rank": "rank_geral",
        "Grupo": "grupo",
        "Família": "familia",
        "ABC SHOP": "abc_shop",
        "Qt.Est. Shop": "qt_est_shop",
        "Qt.Venda Shop": "qt_venda_shop",
        "Cap.Supri Shop": "cap_sup_shop",
        "Val.Total Shop": "val_total_shop",
        "ABC Aero": "abc_aero",
        "Qt.Est. Aero": "qt_est_aero",
        "Qt.Venda Aero": "qt_venda_aero",
        "Cap.Supri Aero": "cap_sup_aero",
        "Val.Total Aero": "val_total_aero",
    }
    colunas_presentes = {k: v for k, v in mapa_colunas.items() if k in df_raw.columns}
    df_limpo = df_raw[list(colunas_presentes.keys())].rename(columns=colunas_presentes).copy()
    df_limpo["codigo"] = normalizar_sku(df_limpo["codigo"])

    # Remove linhas de rodapé ("TOTAL GERAL", vazias ou cabeçalhos repetidos)
    df_limpo = df_limpo[
        df_limpo["codigo"].notna()
        & (~df_limpo["codigo"].isin(["NAN", "CÓDIGO", "CODIGO", "TOTAL", "TOTAL GERAL"]))
        & df_limpo["produto"].notna()
    ].copy()

    cols_numericas = [
        "qt_est_rede", "qt_venda", "qt_sell_in", "cap_sup_rede", "val_custo_venda",
        "val_total_venda", "estoque_cd", "cap_sup_cd", "tabela_atacado", "markup",
        "qt_est_shop", "qt_venda_shop", "cap_sup_shop", "val_total_shop",
        "qt_est_aero", "qt_venda_aero", "cap_sup_aero", "val_total_aero",
    ]
    for col_num in cols_numericas:
        if col_num in df_limpo.columns:
            df_limpo[col_num] = pd.to_numeric(df_limpo[col_num], errors="coerce").fillna(0.0)
        else:
            df_limpo[col_num] = 0.0

    for col_txt in ["curva_abc", "abc_shop", "abc_aero", "grupo", "familia"]:
        if col_txt in df_limpo.columns:
            df_limpo[col_txt] = df_limpo[col_txt].fillna("S/ Curva").astype(str).str.strip()
        else:
            df_limpo[col_txt] = "S/ Curva"

    df_limpo["cap_sup_total"] = (df_limpo["cap_sup_rede"] + df_limpo["cap_sup_cd"]).round(2)
    df_limpo["dt_1_rec"] = pd.to_datetime(df_limpo.get("dt_1_rec"), errors="coerce")
    return df_limpo.drop_duplicates("codigo")


def ler_aba_estoque_completa(xls: pd.ExcelFile, nome_aba: str) -> pd.DataFrame:
    df_est = pd.read_excel(xls, sheet_name=nome_aba)
    renomear = {}
    for c in df_est.columns:
        c_low = str(c).lower().replace(" ", "")
        if "código" in c_low or "codigo" in c_low:
            renomear[c] = "codigo"
        elif "nome" in c_low:
            renomear[c] = "produto_est"
        elif "atacado" in c_low:
            renomear[c] = "atacado_padrao"
        elif "franquia" in c_low:
            renomear[c] = "tabela_franquia"
        elif "preçolojas" in c_low or "precolojas" in c_low:
            renomear[c] = "preco_lojas"
        elif "quantidade" in c_low:
            renomear[c] = "qtd_cd_estoque"
        elif "grupo" in c_low:
            renomear[c] = "grupo_est"
        elif "família" in c_low or "familia" in c_low:
            renomear[c] = "familia_est"
    df_est = df_est.rename(columns=renomear)
    cols_manter = [
        c for c in ["codigo", "produto_est", "atacado_padrao", "tabela_franquia", "preco_lojas", "qtd_cd_estoque", "grupo_est", "familia_est"]
        if c in df_est.columns
    ]
    df_est = df_est[cols_manter].copy()
    df_est["codigo"] = normalizar_sku(df_est["codigo"])
    df_est = df_est[df_est["codigo"].notna() & (~df_est["codigo"].isin(["NAN", "CÓDIGO", "TOTAL"]))]

    for c_num in ["atacado_padrao", "tabela_franquia", "preco_lojas", "qtd_cd_estoque"]:
        if c_num in df_est.columns:
            df_est[c_num] = pd.to_numeric(df_est[c_num], errors="coerce").fillna(0.0)
        else:
            df_est[c_num] = 0.0
    return df_est.drop_duplicates("codigo")


def construir_master_a_partir_de_bases(df_est_atual: pd.DataFrame, df_curva_atual: pd.DataFrame) -> pd.DataFrame:
    """Consolida o Catálogo Mestre do último mês unindo Estoque Completo + Curva ABC com todas as métricas."""
    df_master = pd.merge(df_est_atual, df_curva_atual, on="codigo", how="outer", suffixes=("_est_dup", ""))

    if "produto_est" in df_master.columns:
        df_master["produto"] = df_master.get("produto", pd.Series(dtype=object)).combine_first(df_master["produto_est"]).fillna("Sem Nome")
    else:
        df_master["produto"] = df_master.get("produto", pd.Series("Sem Nome", index=df_master.index)).fillna("Sem Nome")

    if "grupo_est" in df_master.columns:
        df_master["grupo"] = df_master.get("grupo", pd.Series(dtype=object)).combine_first(df_master["grupo_est"]).fillna("GERAL").astype(str).str.strip().str.upper()
    else:
        df_master["grupo"] = df_master.get("grupo", pd.Series("GERAL", index=df_master.index)).fillna("GERAL").astype(str).str.strip().str.upper()

    if "familia_est" in df_master.columns:
        df_master["familia"] = df_master.get("familia", pd.Series(dtype=object)).combine_first(df_master["familia_est"]).fillna("GERAL").astype(str).str.strip().str.upper()
    else:
        df_master["familia"] = df_master.get("familia", pd.Series("GERAL", index=df_master.index)).fillna("GERAL").astype(str).str.strip().str.upper()

    df_master["curva_abc"] = df_master.get("curva_abc", pd.Series("S/ Curva", index=df_master.index)).fillna("S/ Curva")
    df_master["abc_shop"] = df_master.get("abc_shop", pd.Series("S/ Curva", index=df_master.index)).fillna("S/ Curva")
    df_master["abc_aero"] = df_master.get("abc_aero", pd.Series("S/ Curva", index=df_master.index)).fillna("S/ Curva")

    # Garante colunas de preço e estoque vindas de ambas as pontas
    if "atacado_padrao_est_dup" in df_master.columns:
        df_master["atacado_padrao"] = df_master["atacado_padrao_est_dup"].combine_first(df_master.get("atacado_padrao"))
    if "preco_lojas_est_dup" in df_master.columns:
        df_master["preco_lojas"] = df_master["preco_lojas_est_dup"].combine_first(df_master.get("preco_lojas"))

    df_master["qt_estoque"] = pd.to_numeric(
        df_master.get("qtd_cd_estoque", pd.Series(dtype=float)).combine_first(df_master.get("estoque_cd", pd.Series(dtype=float))),
        errors="coerce",
    ).fillna(0.0)

    atacado_pad = pd.to_numeric(df_master.get("atacado_padrao", 0.0), errors="coerce").fillna(0.0)
    tab_atacado = pd.to_numeric(df_master.get("tabela_atacado", 0.0), errors="coerce").fillna(0.0)
    df_master["val_custo"] = np.where(atacado_pad > 0, atacado_pad, tab_atacado)
    df_master["preco_lojas"] = pd.to_numeric(df_master.get("preco_lojas", 0.0), errors="coerce").fillna(0.0)

    for col_z in [
        "qt_est_rede", "qt_venda", "qt_sell_in", "cap_sup_rede", "estoque_cd",
        "cap_sup_cd", "cap_sup_total", "val_total_venda", "markup",
        "qt_est_shop", "qt_venda_shop", "cap_sup_shop", "val_total_shop",
        "qt_est_aero", "qt_venda_aero", "cap_sup_aero", "val_total_aero",
    ]:
        if col_z in df_master.columns:
            df_master[col_z] = pd.to_numeric(df_master[col_z], errors="coerce").fillna(0.0)
        else:
            df_master[col_z] = 0.0

    df_master["estoque_total_empresa"] = df_master["qt_estoque"] + df_master["qt_est_rede"]
    df_master["capital_imobilizado"] = (df_master["qt_estoque"] * df_master["val_custo"]).round(2)
    df_master["capital_lojas"] = (df_master["qt_est_rede"] * df_master["val_custo"]).round(2)
    df_master["capital_total_rede_cd"] = (df_master["capital_imobilizado"] + df_master["capital_lojas"]).round(2)

    df_master["quantidade_a_receber"] = 0.0
    df_master["rma_pct_display"] = 0.0

    return df_master.drop_duplicates("codigo")


@st.cache_data(ttl=600, show_spinner=False)
def carregar_ecossistema_analise_estoque(ultimos_n_meses: int = 12):
    """
    Constrói todo o ecossistema da LOFT:
    1º Tenta ler diretamente da planilha '5 - Análise Estoque.xlsx' (OneDrive LOFT ou pasta local).
    2º Se estiver rodando na Nuvem (Streamlit Cloud) ou sem a planilha local, lê 100% do Supabase!
    """
    # ROTA 1: Planilha Local / OneDrive (Quando rodando na máquina local)
    caminho_excel = encontrar_planilha_local()
    if caminho_excel:
        try:
            xls = pd.ExcelFile(caminho_excel)
            abas_curva = [s for s in xls.sheet_names if s.startswith("Curva ")]
            meses_disponiveis = [s.replace("Curva ", "").strip() for s in abas_curva]
            meses_recorte = meses_disponiveis[-ultimos_n_meses:]
            ultimo_mes = meses_recorte[-1]

            lista_historico = []
            for m in meses_recorte:
                aba_c = f"Curva {m}"
                aba_e = f"Estoque {m}"
                df_c = ler_aba_curva_completa(xls, aba_c)
                if aba_e in xls.sheet_names:
                    df_e = ler_aba_estoque_completa(xls, aba_e)
                    df_c = df_c.merge(
                        df_e[["codigo", "atacado_padrao", "tabela_franquia", "preco_lojas", "qtd_cd_estoque"]],
                        on="codigo",
                        how="left",
                    )
                else:
                    df_c["atacado_padrao"] = df_c["tabela_atacado"]
                    df_c["tabela_franquia"] = 0.0
                    df_c["preco_lojas"] = 0.0
                    df_c["qtd_cd_estoque"] = df_c["estoque_cd"]
                df_c["mes_ref"] = m
                lista_historico.append(df_c)

            df_hist = pd.concat(lista_historico, ignore_index=True)

            df_est_atual = ler_aba_estoque_completa(xls, f"Estoque {ultimo_mes}")
            df_curva_atual = ler_aba_curva_completa(xls, f"Curva {ultimo_mes}")
            df_master = construir_master_a_partir_de_bases(df_est_atual, df_curva_atual)

            return df_master, df_hist, meses_recorte, f"Planilha Oficial ({os.path.basename(caminho_excel)} | Mês {ultimo_mes})"
        except Exception as e_excel:
            st.warning(f"Aviso ao ler planilha local ({e_excel}). Acionando banco de dados em nuvem (Supabase)...")

    # ROTA 2: Nuvem Supabase (Quando rodando no Streamlit Cloud ou sem Excel local)
    client_supa = obter_cliente_supabase()
    if client_supa is not None:
        try:
            df_hist = buscar_tabela_completa_supabase(client_supa, "loft_historico_curva")
            df_cat = buscar_tabela_completa_supabase(client_supa, "loft_catalogo_estoque")

            if not df_hist.empty:
                df_hist["codigo"] = normalizar_sku(df_hist["codigo"])
                cols_num_h = [
                    "qt_est_rede", "qt_venda", "qt_sell_in", "cap_sup_rede", "val_custo_venda",
                    "val_total_venda", "estoque_cd", "cap_sup_cd", "cap_sup_total", "tabela_atacado",
                    "atacado_padrao", "preco_lojas", "markup", "qt_est_shop", "qt_venda_shop",
                    "cap_sup_shop", "val_total_shop", "qt_est_aero", "qt_venda_aero", "cap_sup_aero", "val_total_aero",
                ]
                for c_num in cols_num_h:
                    if c_num in df_hist.columns:
                        df_hist[c_num] = pd.to_numeric(df_hist[c_num], errors="coerce").fillna(0.0)
                    else:
                        df_hist[c_num] = 0.0

                for col_txt in ["curva_abc", "abc_shop", "abc_aero", "grupo", "familia"]:
                    if col_txt in df_hist.columns:
                        df_hist[col_txt] = df_hist[col_txt].fillna("S/ Curva").astype(str).str.strip()
                    else:
                        df_hist[col_txt] = "S/ Curva"

                df_hist["dt_1_rec"] = pd.to_datetime(df_hist.get("dt_1_rec"), errors="coerce")

                meses_ordenados = ordenar_meses_cronologicamente(df_hist["mes_ref"].dropna().unique().tolist())
                meses_recorte = meses_ordenados[-ultimos_n_meses:]
                ultimo_mes = meses_recorte[-1]

                df_curva_atual = df_hist[df_hist["mes_ref"] == ultimo_mes].copy().drop_duplicates("codigo")

                if not df_cat.empty:
                    df_est_atual = df_cat.rename(
                        columns={
                            "produto": "produto_est",
                            "grupo": "grupo_est",
                            "familia": "familia_est",
                        }
                    ).copy()
                    df_est_atual["codigo"] = normalizar_sku(df_est_atual["codigo"])
                    for c_n in ["atacado_padrao", "tabela_franquia", "preco_lojas", "qtd_cd_estoque"]:
                        if c_n in df_est_atual.columns:
                            df_est_atual[c_n] = pd.to_numeric(df_est_atual[c_n], errors="coerce").fillna(0.0)
                        else:
                            df_est_atual[c_n] = 0.0
                    df_est_atual = df_est_atual.drop_duplicates("codigo")
                else:
                    df_est_atual = pd.DataFrame(columns=["codigo", "produto_est", "atacado_padrao", "tabela_franquia", "preco_lojas", "qtd_cd_estoque", "grupo_est", "familia_est"])

                df_master = construir_master_a_partir_de_bases(df_est_atual, df_curva_atual)
                hora_leitura = datetime.now().strftime("%d/%m %H:%M")
                return df_master, df_hist, meses_recorte, f"Supabase Nuvem (Mês {ultimo_mes} | Sinc: {hora_leitura})"
        except Exception as e_supa:
            st.error(f"Erro ao carregar dados do Supabase: {e_supa}")

    return pd.DataFrame(), pd.DataFrame(), [], "Base não encontrada"


with st.spinner("Carregando inteligência completa de estoque LOFT (OneDrive / Supabase)..."):
    df_completo, df_historico_mensal, meses_planilha, fonte_ativa = carregar_ecossistema_analise_estoque(ultimos_n_meses=12)

if df_completo.empty:
    st.error(
        "Não foi possível carregar a base oficial via Planilha Local nem via Supabase. "
        "No Streamlit Cloud, confira em 'Manage app -> Settings -> Secrets' se SUPABASE_URL e SUPABASE_KEY estão preenchidas corretamente."
    )
    st.stop()


# ==========================================
# 4. BARRA LATERAL (FILTROS + UPLOAD DE RECOMPRA / RMA)
# ==========================================
with st.sidebar:
    st.markdown("### 📦 LOFT DESIGN")
    st.caption("Intelligence & Supply Chain")
    st.divider()

    modo_escuro = st.toggle("🌙 Ativar Modo Escuro", value=False)

    st.markdown("#### 🚢 Simulador de Recompra & RMA")
    arquivo_recompra = st.file_uploader(
        "Subir Planilha de Recompra (.xlsx / .csv)",
        type=["xlsx", "xls", "csv"],
        help="Suba aqui a planilha de Recompra (feita via PROCV) para ativar os dados de 'A Receber', 'Taxa de RMA' e simular como ficará a Capacidade de Suprimento!",
    )

    st.divider()
    st.markdown("#### Filtros de Visualização")
    ocultar_uso_interno = st.checkbox("Excluir Serviços, Uso Próprio e Variados", value=True)
    apenas_estoque_positivo = st.checkbox("Somente itens com estoque > 0 (CD ou Lojas)", value=True)
    ocultar_sem_curva = st.checkbox("Ocultar 'S/ Curva' (Sem venda no mês)", value=False)
    grupos_escolhidos = st.multiselect("Filtrar por Grupo", sorted(df_completo["grupo"].unique()))

    st.divider()
    if st.button("🔄 Atualizar Base Agora", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.caption(f"🟢 Fonte Ativa: **{fonte_ativa}**")
    st.caption(f"Catálogo Total: **{formatar_inteiro(len(df_completo))} SKUs**")


# ==========================================
# 5. PROCESSAMENTO DINÂMICO DA PLANILHA DE RECOMPRA + RMA (QUANDO SUBIDA)
# ==========================================
recompra_ativa = False
df_recompra_detalhe = pd.DataFrame()
resumo_recompra_ia = "Nenhuma planilha de recompra enviada nesta sessão."

if arquivo_recompra is not None:
    try:
        if arquivo_recompra.name.lower().endswith(".csv"):
            df_rec_raw = pd.read_csv(arquivo_recompra, sep=None, engine="python")
        else:
            xls_rec = pd.ExcelFile(arquivo_recompra)
            aba_rec_sel = xls_rec.sheet_names[0]
            if len(xls_rec.sheet_names) > 1:
                aba_rec_sel = st.sidebar.selectbox("Selecionar Aba da Recompra:", xls_rec.sheet_names)
            df_rec_raw = pd.read_excel(xls_rec, sheet_name=aba_rec_sel)

        # Identifica automaticamente colunas de Código, Quantidade a Receber/Comprar e RMA/Defeito
        col_cod, col_qtd_rec, col_rma, col_data = None, None, None, None
        for c in df_rec_raw.columns:
            c_clean = str(c).strip().lower()
            if col_cod is None and any(k in c_clean for k in ["código", "codigo", "sku", "cod", "item"]):
                col_cod = c
            if col_qtd_rec is None and any(k in c_clean for k in ["receber", "recompra", "pedido", "compra", "sugest", "container", "contêiner", "qtd", "quant"]):
                if c != col_cod:
                    col_qtd_rec = c
            if col_rma is None and any(k in c_clean for k in ["rma", "defeito", "garantia", "avaria", "devolu"]):
                col_rma = c
            if col_data is None and any(k in c_clean for k in ["data", "chegada", "previs", "eta"]):
                col_data = c

        if col_cod is not None:
            recompra_ativa = True
            df_rec_proc = pd.DataFrame()
            df_rec_proc["codigo"] = normalizar_sku(df_rec_raw[col_cod])

            if col_qtd_rec is not None:
                df_rec_proc["qtd_recompra_input"] = pd.to_numeric(df_rec_raw[col_qtd_rec], errors="coerce").fillna(0.0)
            else:
                df_rec_proc["qtd_recompra_input"] = 0.0

            if col_rma is not None:
                rma_vals = pd.to_numeric(df_rec_raw[col_rma], errors="coerce").fillna(0.0)
                if rma_vals.max() <= 1.0 and rma_vals.max() > 0:
                    rma_vals = rma_vals * 100.0
                df_rec_proc["rma_input"] = rma_vals.round(2)
            else:
                df_rec_proc["rma_input"] = 0.0

            if col_data is not None:
                df_rec_proc["previsao_chegada"] = df_rec_raw[col_data]

            df_rec_agg = (
                df_rec_proc.groupby("codigo")
                .agg(
                    quantidade_a_receber=("qtd_recompra_input", "sum"),
                    rma_pct_display=("rma_input", "max"),
                )
                .reset_index()
            )

            # Atualiza df_completo com os dados da planilha de Recompra subida!
            df_completo = df_completo.drop(columns=["quantidade_a_receber", "rma_pct_display"], errors="ignore")
            df_completo = df_completo.merge(df_rec_agg, on="codigo", how="left")
            df_completo["quantidade_a_receber"] = df_completo["quantidade_a_receber"].fillna(0.0)
            df_completo["rma_pct_display"] = df_completo["rma_pct_display"].fillna(0.0)

            # Calcula a Nova Capacidade de Suprimento Projetada pós-Recompra
            df_completo["cap_sup_pos_recompra"] = np.where(
                df_completo["qt_venda"] > 0,
                ((df_completo["estoque_total_empresa"] + df_completo["quantidade_a_receber"]) / df_completo["qt_venda"]).round(2),
                df_completo["cap_sup_total"],
            )

            df_recompra_detalhe = df_completo[
                (df_completo["quantidade_a_receber"] > 0) | (df_completo["rma_pct_display"] > 0)
            ].copy()

            resumo_recompra_ia = (
                f"🚢 PLANILHA DE RECOMPRA E RMA ATIVA ({arquivo_recompra.name}):\n"
                f"Total de Unidades na Recompra: {formatar_inteiro(df_completo['quantidade_a_receber'].sum())} un. | "
                f"SKUs com Recompra > 0: {int((df_completo['quantidade_a_receber'] > 0).sum())} | "
                f"SKUs com RMA registrado: {int((df_completo['rma_pct_display'] > 0).sum())}\n"
                f"Amostra dos Itens em Recompra/RMA:\n"
                f"{df_recompra_detalhe[['codigo', 'produto', 'curva_abc', 'qt_estoque', 'qt_est_rede', 'qt_venda', 'cap_sup_total', 'quantidade_a_receber', 'cap_sup_pos_recompra', 'rma_pct_display']].head(150).to_csv(index=False)}\n"
            )
            st.sidebar.success(f"Recompra aplicada! ({formatar_inteiro(df_completo['quantidade_a_receber'].sum())} un. a receber)")
        else:
            st.sidebar.warning("Não encontrei uma coluna chamada 'Código' ou 'SKU' na planilha enviada.")
    except Exception as e_rec:
        st.sidebar.error(f"Erro ao processar planilha de recompra: {e_rec}")


# ==========================================
# 6. ESTILIZAÇÃO CSS (BRANCO PADRÃO + LARANJA / DARK SUAVE)
# ==========================================
COR_LARANJA = "#FF6B00"
MAPA_CORES_ABC = {"A": "#1f77b4", "B": "#FF6B00", "C": "#2ca02c", "S/ Curva": "#7f7f7f"}

if modo_escuro:
    ESTILO = {
        "bg_app": "#181B22",
        "bg_sidebar": "#12151B",
        "bg_card": "#222731",
        "texto": "#F1F5F9",
        "texto_sec": "#94A3B8",
        "borda": "#333A48",
        "btn_bg": "#2A303C",
        "btn_texto": "#F8FAFC",
        "plotly_tpl": "plotly_dark",
        "grade": "rgba(255, 255, 255, 0.07)",
    }
else:
    ESTILO = {
        "bg_app": "#FFFFFF",
        "bg_sidebar": "#F8FAFC",
        "bg_card": "#FFFFFF",
        "texto": "#1E293B",
        "texto_sec": "#64748B",
        "borda": "#E2E8F0",
        "btn_bg": "#F8FAFC",
        "btn_texto": "#1E293B",
        "plotly_tpl": "plotly_white",
        "grade": "rgba(0, 0, 0, 0.06)",
    }

st.markdown(
    f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    html, body, .stApp {{
        font-family: 'Inter', sans-serif;
        background-color: {ESTILO["bg_app"]};
        color: {ESTILO["texto"]};
    }}

    section[data-testid="stSidebar"] {{
        background-color: {ESTILO["bg_sidebar"]};
        border-right: 1px solid {ESTILO["borda"]};
    }}

    .painel-header {{
        background-color: {ESTILO["bg_card"]};
        border: 1px solid {ESTILO["borda"]};
        border-left: 5px solid {COR_LARANJA};
        border-radius: 10px;
        padding: 18px 24px;
        margin-bottom: 20px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.03);
    }}
    .painel-titulo {{
        font-size: 1.55rem;
        font-weight: 700;
        color: {ESTILO["texto"]};
        margin: 0;
    }}
    .painel-subtitulo {{
        font-size: 0.88rem;
        color: {ESTILO["texto_sec"]};
        margin-top: 4px;
    }}
    .painel-badge {{
        background-color: rgba(255, 107, 0, 0.12);
        color: {COR_LARANJA};
        border: 1px solid {COR_LARANJA};
        font-weight: 600;
        font-size: 0.82rem;
        padding: 6px 14px;
        border-radius: 999px;
    }}

    .kpi-container {{
        display: grid;
        grid-template-columns: repeat(6, 1fr);
        gap: 12px;
        margin-bottom: 18px;
    }}
    @media (max-width: 1200px) {{
        .kpi-container {{ grid-template-columns: repeat(3, 1fr); }}
    }}
    .kpi-box {{
        background-color: {ESTILO["bg_card"]};
        border: 1px solid {ESTILO["borda"]};
        border-top: 3px solid {COR_LARANJA};
        border-radius: 10px;
        padding: 14px 16px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.02);
    }}
    .kpi-titulo {{
        font-size: 0.74rem;
        font-weight: 600;
        text-transform: uppercase;
        color: {ESTILO["texto_sec"]};
        margin-bottom: 6px;
    }}
    .kpi-numero {{
        font-size: 1.28rem;
        font-weight: 700;
        color: {ESTILO["texto"]};
    }}
    .kpi-detalhe {{
        font-size: 0.72rem;
        color: {ESTILO["texto_sec"]};
        margin-top: 4px;
    }}

    .stButton > button, .stDownloadButton > button {{
        background-color: {ESTILO["btn_bg"]} !important;
        color: {ESTILO["btn_texto"]} !important;
        border: 1px solid {ESTILO["borda"]} !important;
        border-radius: 8px !important;
        font-weight: 500 !important;
        transition: all 0.2s ease !important;
    }}
    .stButton > button:hover, .stDownloadButton > button:hover {{
        border-color: {COR_LARANJA} !important;
        color: {COR_LARANJA} !important;
    }}

    .stTabs [data-baseweb="tab"] {{
        color: {ESTILO["texto_sec"]} !important;
        font-weight: 600;
    }}
    .stTabs [aria-selected="true"] {{
        color: {COR_LARANJA} !important;
        border-bottom-color: {COR_LARANJA} !important;
    }}

    .secao-tag {{
        font-size: 0.75rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: {COR_LARANJA};
        margin-top: 8px;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


def padronizar_grafico(fig, altura: int = 380):
    fig.update_layout(
        template=ESTILO["plotly_tpl"],
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, sans-serif", color=ESTILO["texto"], size=12),
        title_font=dict(family="Inter, sans-serif", size=14, color=ESTILO["texto"]),
        margin=dict(l=14, r=14, t=46, b=20),
        height=altura,
    )
    fig.update_xaxes(showgrid=True, gridcolor=ESTILO["grade"], zeroline=False)
    fig.update_yaxes(showgrid=False, gridcolor=ESTILO["grade"], zeroline=False)
    return fig


# ==========================================
# 7. APLICAÇÃO DOS FILTROS E KPIs DO TOPO
# ==========================================
df = df_completo.copy()
if ocultar_uso_interno:
    df = df[~df["grupo"].isin(CATEGORIAS_IGNORADAS)]
if apenas_estoque_positivo:
    df = df[(df["qt_estoque"] > 0) | (df["qt_est_rede"] > 0)]
if ocultar_sem_curva:
    df = df[df["curva_abc"].isin(["A", "B", "C"])]
if grupos_escolhidos:
    df = df[df["grupo"].isin(grupos_escolhidos)]

if df.empty:
    st.warning("Nenhum produto encontrado para a combinação de filtros selecionada.")
    st.stop()

st.markdown(
    f"""
    <div class="painel-header">
        <div>
            <h1 class="painel-titulo">📦 Painel Estratégico de Estoque - LOFT</h1>
            <div class="painel-subtitulo">Base 100% Oficial ('5 - Análise Estoque.xlsx' / Supabase Cloud): CD vs. Lojas, Shopping vs. Aeroporto, Elasticidade de Preço e Simulador de Recompra</div>
        </div>
        <div class="painel-badge">
            {formatar_inteiro(len(df))} SKUs na Visão
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

vol_cd = df["qt_estoque"].sum()
vol_lojas = df["qt_est_rede"].sum()
val_cd = df["capital_imobilizado"].sum()
faturamento_mes = df["val_total_venda"].sum()
vendas_un_mes = df["qt_venda"].sum()
markup_med = df.loc[df["markup"] > 0, "markup"].mean()
if pd.isna(markup_med):
    markup_med = 0.0

if recompra_ativa:
    kpi6_titulo = "Recompra / RMA Ativo"
    kpi6_num = f"+{formatar_inteiro(df['quantidade_a_receber'].sum())} un."
    rma_m = df.loc[df["rma_pct_display"] > 0, "rma_pct_display"].mean()
    kpi6_det = f"RMA Médio: {0.0 if pd.isna(rma_m) else rma_m:.1f}%"
else:
    kpi6_titulo = "Simulador Recompra"
    kpi6_num = "Aguardando"
    kpi6_det = "Suba a planilha na barra lateral"

st.markdown(
    f"""
    <div class="kpi-container">
        <div class="kpi-box">
            <div class="kpi-titulo">Estoque CD (Galpão)</div>
            <div class="kpi-numero">{formatar_inteiro(vol_cd)} un.</div>
            <div class="kpi-detalhe">Capital CD: {formatar_moeda(val_cd)}</div>
        </div>
        <div class="kpi-box">
            <div class="kpi-titulo">Estoque Rede (Lojas)</div>
            <div class="kpi-numero">{formatar_inteiro(vol_lojas)} un.</div>
            <div class="kpi-detalhe">Shop: {formatar_inteiro(df['qt_est_shop'].sum())} | Aero: {formatar_inteiro(df['qt_est_aero'].sum())}</div>
        </div>
        <div class="kpi-box">
            <div class="kpi-titulo">Faturamento Curva (Mês)</div>
            <div class="kpi-numero">{formatar_moeda(faturamento_mes)}</div>
            <div class="kpi-detalhe">Giro: {formatar_inteiro(vendas_un_mes)} un. vendidas</div>
        </div>
        <div class="kpi-box">
            <div class="kpi-titulo">Shoppings vs. Aeroportos</div>
            <div class="kpi-numero">{formatar_moeda(df['val_total_shop'].sum())}</div>
            <div class="kpi-detalhe">Aero: {formatar_moeda(df['val_total_aero'].sum())}</div>
        </div>
        <div class="kpi-box">
            <div class="kpi-titulo">Markup Médio / Curva A</div>
            <div class="kpi-numero">{markup_med:.2f}x</div>
            <div class="kpi-detalhe">{int((df['curva_abc'] == 'A').sum())} SKUs na Curva A</div>
        </div>
        <div class="kpi-box">
            <div class="kpi-titulo">{kpi6_titulo}</div>
            <div class="kpi-numero">{kpi6_num}</div>
            <div class="kpi-detalhe">{kpi6_det}</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ==========================================
# 8. NAVEGAÇÃO PRINCIPAL EM 5 ABAS ESTRATÉGICAS
# ==========================================
aba_diretoria, aba_canais, aba_recompra, aba_ia, aba_tabela = st.tabs(
    [
        "🎯 1. Métricas da Diretoria (Automático)",
        "🏬 2. Canais (Shop vs. Aero) & CD",
        "🚢 3. Simulador de Recompra & RMA",
        "🧠 4. Assistente Executivo IA",
        "🗂️ 5. Explorador de Dados",
    ]
)

# ------------------------------------------
# ABA 1: MÉTRICAS DA DIRETORIA (GUSTAVO & DIRETOR)
# ------------------------------------------
with aba_diretoria:
    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        mes_atual_sel = st.selectbox("Mês de Fechamento (Atual):", meses_planilha, index=len(meses_planilha) - 1)
    idx_ant = max(0, meses_planilha.index(mes_atual_sel) - 1)
    with col_m2:
        mes_ant_sel = st.selectbox("Mês de Comparação (Anterior):", meses_planilha, index=idx_ant)
    with col_m3:
        criterio_preco = st.selectbox(
            "Referência para Redução de Preço:",
            ["Ambos (Preço Lojas ou Tabela Atacado)", "Apenas Preço Lojas (Varejo)", "Apenas Tabela Atacado"],
        )

    df_mes_atu = df_historico_mensal[df_historico_mensal["mes_ref"] == mes_atual_sel].copy()
    df_mes_ant = df_historico_mensal[df_historico_mensal["mes_ref"] == mes_ant_sel].copy()

    # ANÁLISE 1: REDUÇÕES DE PREÇO VS. AUMENTO OU QUEDA DE VENDA
    st.markdown('<div class="secao-tag">1. ELASTICIDADE COMERCIAL • REDUÇÕES DE PREÇO VS. GIRO</div>', unsafe_allow_html=True)
    st.markdown(f"#### Impacto das Reduções de Preço nas Vendas (`{mes_ant_sel}` → `{mes_atual_sel}`)")

    comp_preco = pd.merge(
        df_mes_atu[["codigo", "produto", "grupo", "curva_abc", "qt_venda", "qt_est_rede", "estoque_cd", "tabela_atacado", "preco_lojas"]],
        df_mes_ant[["codigo", "qt_venda", "tabela_atacado", "preco_lojas"]],
        on="codigo",
        how="inner",
        suffixes=("_atu", "_ant"),
    )
    comp_preco["diff_loja_rs"] = (comp_preco["preco_lojas_atu"] - comp_preco["preco_lojas_ant"]).round(2)
    comp_preco["diff_atacado_rs"] = (comp_preco["tabela_atacado_atu"] - comp_preco["tabela_atacado_ant"]).round(2)
    comp_preco["var_venda_un"] = comp_preco["qt_venda_atu"] - comp_preco["qt_venda_ant"]
    comp_preco["var_venda_pct"] = np.where(
        comp_preco["qt_venda_ant"] > 0,
        ((comp_preco["var_venda_un"] / comp_preco["qt_venda_ant"]) * 100).round(1),
        0.0,
    )

    if criterio_preco.startswith("Apenas Preço Lojas"):
        df_reduzidos = comp_preco[(comp_preco["preco_lojas_ant"] > 0) & (comp_preco["diff_loja_rs"] < -0.50)].copy()
    elif criterio_preco.startswith("Apenas Tabela Atacado"):
        df_reduzidos = comp_preco[(comp_preco["tabela_atacado_ant"] > 0) & (comp_preco["diff_atacado_rs"] < -0.50)].copy()
    else:
        df_reduzidos = comp_preco[
            ((comp_preco["preco_lojas_ant"] > 0) & (comp_preco["diff_loja_rs"] < -0.50))
            | ((comp_preco["tabela_atacado_ant"] > 0) & (comp_preco["diff_atacado_rs"] < -0.50))
        ].copy()

    if df_reduzidos.empty:
        st.info(f"Nenhuma redução de preço detectada entre {mes_ant_sel} e {mes_atual_sel}.")
    else:
        qtd_subiu = int((df_reduzidos["var_venda_un"] > 0).sum())
        qtd_caiu = int((df_reduzidos["var_venda_un"] < 0).sum())

        kp1, kp2, kp3, kp4 = st.columns(4)
        kp1.metric("SKUs na Curva c/ Redução de Preço", formatar_inteiro(len(df_reduzidos)))
        kp2.metric("Aumentaram Vendas 📈", f"{qtd_subiu} SKUs", f"{(qtd_subiu/len(df_reduzidos)*100):.1f}% do total")
        kp3.metric("Caíram Vendas 📉", f"{qtd_caiu} SKUs", f"-{(qtd_caiu/len(df_reduzidos)*100):.1f}% do total")
        kp4.metric("Saldo Líquido de Unidades", f"{int(df_reduzidos['var_venda_un'].sum()):+d} un.")

        col_p1, col_p2 = st.columns(2)
        with col_p1:
            top_aumento = df_reduzidos[df_reduzidos["var_venda_un"] > 0].nlargest(10, "var_venda_un")
            if not top_aumento.empty:
                fig_up = px.bar(
                    top_aumento,
                    x="var_venda_un",
                    y="produto",
                    orientation="h",
                    title="Top 10: Reduziram Preço e MAIS AUMENTARAM Vendas (+ un.)",
                    text_auto=True,
                    color="var_venda_un",
                    color_continuous_scale="Emrld",
                    hover_data=["codigo", "preco_lojas_ant", "preco_lojas_atu", "tabela_atacado_ant", "tabela_atacado_atu", "qt_venda_ant", "qt_venda_atu"],
                )
                fig_up.update_layout(yaxis={"categoryorder": "total ascending"}, yaxis_title="", xaxis_title="Ganho (+ Unidades)", coloraxis_showscale=False)
                st.plotly_chart(padronizar_grafico(fig_up, altura=380), width="stretch")

        with col_p2:
            top_queda = df_reduzidos[df_reduzidos["var_venda_un"] < 0].nsmallest(10, "var_venda_un")
            if not top_queda.empty:
                fig_down = px.bar(
                    top_queda,
                    x="var_venda_un",
                    y="produto",
                    orientation="h",
                    title="Top 10: Reduziram Preço e MAIS CAÍRAM Vendas (- un.)",
                    text_auto=True,
                    color="var_venda_un",
                    color_continuous_scale="Reds_r",
                    hover_data=["codigo", "preco_lojas_ant", "preco_lojas_atu", "tabela_atacado_ant", "tabela_atacado_atu", "qt_venda_ant", "qt_venda_atu"],
                )
                fig_down.update_layout(yaxis={"categoryorder": "total descending"}, yaxis_title="", xaxis_title="Queda (- Unidades)", coloraxis_showscale=False)
                st.plotly_chart(padronizar_grafico(fig_down, altura=380), width="stretch")

        with st.expander("📋 Ver e Exportar Tabela Completa de Produtos com Redução de Preço"):
            st.download_button(
                "📥 Baixar Relatório de Reduções de Preço (CSV)",
                data=gerar_arquivo_csv(df_reduzidos.sort_values("var_venda_un", ascending=False)),
                file_name=f"reducoes_preco_{mes_ant_sel}_para_{mes_atual_sel}.csv",
                mime="text/csv",
            )
            st.dataframe(df_reduzidos.sort_values("var_venda_un", ascending=False), hide_index=True, width="stretch")

    st.divider()

    # ANÁLISE 2: ITENS SUPERESTOCADOS (1 A 3 MESES DA CHEGADA)
    st.markdown('<div class="secao-tag">2. AUDITORIA DE SUPERESTOCADOS • JANELA DE CHEGADA (1 A 3 MESES)</div>', unsafe_allow_html=True)
    st.markdown(f"#### Produtos com 1 a 3 Meses desde a Chegada (`Dt 1° Rec`) e Maior Capacidade de Suprimento (`{mes_atual_sel}`)")

    try:
        mes_num, ano_curto = [int(x) for x in mes_atual_sel.split("-")]
        data_ref_aba = pd.Timestamp(year=2000 + ano_curto, month=mes_num, day=1) + pd.offsets.MonthEnd(1)
    except Exception:
        data_ref_aba = df_mes_atu["dt_1_rec"].max()

    df_mes_atu["dias_chegada"] = (data_ref_aba - df_mes_atu["dt_1_rec"]).dt.days

    col_s1, col_s2 = st.columns(2)
    with col_s1:
        intervalo_dias = st.slider(
            "Janela de Dias desde o 1º Recebimento (Padrão: 30 a 95 dias = 1 a 3 meses):",
            min_value=15,
            max_value=180,
            value=(30, 95),
        )
    with col_s2:
        tipo_cap_sup = st.radio(
            "Métrica de Capacidade de Suprimento para Ordenar:",
            ["Cap. Sup. Total (Rede + CD)", "Apenas Cap. Sup. CD", "Apenas Cap. Sup. Rede (Lojas)"],
            horizontal=True,
        )

    col_ord_cap = (
        "cap_sup_total"
        if tipo_cap_sup.startswith("Cap. Sup. Total")
        else ("cap_sup_cd" if "CD" in tipo_cap_sup else "cap_sup_rede")
    )

    df_janela_1_3m = df_mes_atu[
        (df_mes_atu["dias_chegada"] >= intervalo_dias[0])
        & (df_mes_atu["dias_chegada"] <= intervalo_dias[1])
        & (~df_mes_atu["grupo"].isin(CATEGORIAS_IGNORADAS))
    ].sort_values(col_ord_cap, ascending=False)

    c_j1, c_j2 = st.columns([3, 2])
    with c_j1:
        if not df_janela_1_3m.empty:
            top12_janela = df_janela_1_3m.head(12).copy()
            df_melt_cap = top12_janela.melt(
                id_vars=["produto", "codigo"],
                value_vars=["cap_sup_rede", "cap_sup_cd"],
                var_name="Local",
                value_name="Meses de Cobertura",
            )
            df_melt_cap["Local"] = df_melt_cap["Local"].map(
                {"cap_sup_rede": "Cap. Sup. Rede (Lojas)", "cap_sup_cd": "Cap. Sup. CD"}
            )
            fig_cap_1_3 = px.bar(
                df_melt_cap,
                x="Meses de Cobertura",
                y="produto",
                color="Local",
                orientation="h",
                title=f"Top 12 Maior Cap. Suprimento (Chegada entre {intervalo_dias[0]} e {intervalo_dias[1]} dias)",
                color_discrete_map={"Cap. Sup. Rede (Lojas)": COR_LARANJA, "Cap. Sup. CD": "#1f77b4"},
            )
            fig_cap_1_3.update_layout(yaxis={"categoryorder": "total ascending"}, yaxis_title="", legend_title="")
            st.plotly_chart(padronizar_grafico(fig_cap_1_3, altura=400), width="stretch")

    with c_j2:
        st.markdown("**Lista Completa (1 a 3 Meses da Chegada):**")
        st.dataframe(
            df_janela_1_3m[
                ["codigo", "produto", "dt_1_rec", "dias_chegada", "qt_venda", "qt_est_rede", "estoque_cd", "cap_sup_rede", "cap_sup_cd", "cap_sup_total"]
            ],
            hide_index=True,
            height=360,
            width="stretch",
        )

    st.divider()

    # ANÁLISE 3: OSCILAÇÃO DE CAP. SUP. NO TEMPO (MAIORES DA CURVA)
    st.markdown('<div class="secao-tag">3. MONITORAMENTO TEMPORAL • OSCILAÇÃO DE CAPACIDADE DE SUPRIMENTO</div>', unsafe_allow_html=True)
    st.markdown("#### Evolução Mensal de `Cap. Sup.` dos Produtos Mais Superestocados da Curva")

    df_candidatos_osc = df_mes_atu[~df_mes_atu["grupo"].isin(CATEGORIAS_IGNORADAS)].sort_values(col_ord_cap, ascending=False)
    opcoes_skus_osc = [
        f"{row['codigo']} | {row['produto']} (Cap. Total: {row['cap_sup_total']:.1f}m)"
        for _, row in df_candidatos_osc.head(50).iterrows()
    ]
    mapa_label_sku = {
        f"{row['codigo']} | {row['produto']} (Cap. Total: {row['cap_sup_total']:.1f}m)": row["codigo"]
        for _, row in df_candidatos_osc.head(50).iterrows()
    }

    selecionados_osc = st.multiselect(
        "Selecione os produtos com maior Cap. Sup. para ver a oscilação no tempo (Padrão: Top 5):",
        options=opcoes_skus_osc,
        default=opcoes_skus_osc[:5],
    )

    skus_osc_ativos = [mapa_label_sku[s] for s in selecionados_osc]
    if skus_osc_ativos:
        df_osc_hist = df_historico_mensal[df_historico_mensal["codigo"].isin(skus_osc_ativos)].copy()
        df_osc_hist["legenda_sku"] = df_osc_hist["codigo"] + " - " + df_osc_hist["produto"].astype(str).str[:28]

        fig_osc = px.line(
            df_osc_hist,
            x="mes_ref",
            y=col_ord_cap,
            color="legenda_sku",
            markers=True,
            title=f"Oscilação Mensal de {tipo_cap_sup} ({meses_planilha[0]} a {meses_planilha[-1]})",
            hover_data=["qt_venda", "qt_est_rede", "estoque_cd", "cap_sup_rede", "cap_sup_cd"],
            labels={"mes_ref": "Mês de Referência", col_ord_cap: "Meses de Cobertura (Cap. Sup.)", "legenda_sku": "Produto"},
        )
        fig_osc.update_xaxes(categoryorder="array", categoryarray=meses_planilha)
        st.plotly_chart(padronizar_grafico(fig_osc, altura=420), width="stretch")


# ------------------------------------------
# ABA 2: INTELIGÊNCIA DE CANAIS (SHOPPINGS VS. AEROPORTOS) & CD VS. LOJAS
# ------------------------------------------
with aba_canais:
    st.markdown('<div class="secao-tag">1. GUERRA DE CANAIS • SHOPPINGS VS. AEROPORTOS</div>', unsafe_allow_html=True)
    st.markdown("#### Comparativo de Desempenho entre Lojas de Shopping e Lojas de Aeroporto")

    kc1, kc2, kc3, kc4 = st.columns(4)
    kc1.metric("Faturamento Shoppings", formatar_moeda(df["val_total_shop"].sum()), f"{formatar_inteiro(df['qt_venda_shop'].sum())} un. vendidas")
    kc2.metric("Faturamento Aeroportos", formatar_moeda(df["val_total_aero"].sum()), f"{formatar_inteiro(df['qt_venda_aero'].sum())} un. vendidas")
    kc3.metric("Estoque Físico Shoppings", f"{formatar_inteiro(df['qt_est_shop'].sum())} un.")
    kc4.metric("Estoque Físico Aeroportos", f"{formatar_inteiro(df['qt_est_aero'].sum())} un.")

    col_can1, col_can2 = st.columns(2)
    with col_can1:
        top_shop_aero = df[df["val_total_venda"] > 0].nlargest(12, "val_total_venda").copy()
        df_melt_canais = top_shop_aero.melt(
            id_vars=["produto", "codigo"],
            value_vars=["val_total_shop", "val_total_aero"],
            var_name="Canal",
            value_name="Faturamento (R$)",
        )
        df_melt_canais["Canal"] = df_melt_canais["Canal"].map(
            {"val_total_shop": "Lojas Shopping (R$)", "val_total_aero": "Lojas Aeroporto (R$)"}
        )
        fig_canais = px.bar(
            df_melt_canais,
            x="Faturamento (R$)",
            y="produto",
            color="Canal",
            orientation="h",
            title="Top 12 Campeões de Faturamento: Shopping vs. Aeroporto (R$)",
            color_discrete_map={"Lojas Shopping (R$)": COR_LARANJA, "Lojas Aeroporto (R$)": "#1f77b4"},
        )
        fig_canais.update_layout(yaxis={"categoryorder": "total ascending"}, yaxis_title="", legend_title="")
        st.plotly_chart(padronizar_grafico(fig_canais, altura=400), width="stretch")

    with col_can2:
        # Alerta Operacional: Faltando na Loja (Cap.Supri < 0.8m) mas Sobrando no CD (Estoque CD >= 100 un.)
        df_transferencia = df[
            (df["cap_sup_rede"] < 0.85) & (df["qt_venda"] >= 15) & (df["qt_estoque"] >= 80)
        ].sort_values("qt_venda", ascending=False)

        if df_transferencia.empty:
            st.info("Nenhum produto em alerta de expedição CD -> Lojas com os filtros atuais.")
        else:
            top_transf = df_transferencia.head(12).copy()
            df_melt_tr = top_transf.melt(
                id_vars=["produto", "codigo"],
                value_vars=["qt_est_rede", "qt_estoque"],
                var_name="Posicao",
                value_name="Unidades",
            )
            df_melt_tr["Posicao"] = df_melt_tr["Posicao"].map(
                {"qt_est_rede": "1. Estoque nas Lojas (Baixo)", "qt_estoque": "2. Parado no CD (Disponível p/ Envio)"}
            )
            fig_tr = px.bar(
                df_melt_tr,
                x="Unidades",
                y="produto",
                color="Posicao",
                barmode="group",
                orientation="h",
                title="Alerta de Expedição: Baixa Cobertura nas Lojas vs. Alto Estoque no CD",
                color_discrete_map={"1. Estoque nas Lojas (Baixo)": "#EF4444", "2. Parado no CD (Disponível p/ Envio)": "#10B981"},
            )
            fig_tr.update_layout(yaxis={"categoryorder": "total ascending"}, yaxis_title="", legend_title="")
            st.plotly_chart(padronizar_grafico(fig_tr, altura=400), width="stretch")

    st.divider()
    st.markdown('<div class="secao-tag">2. PORTFÓLIO, CAPITAL IMOBILIZADO E ESTAGNAÇÃO</div>', unsafe_allow_html=True)
    cg1, cg2, cg3 = st.columns(3)

    with cg1:
        valor_curva = df.groupby("curva_abc")["capital_total_rede_cd"].sum().reset_index()
        fig_valor = px.bar(
            valor_curva, x="curva_abc", y="capital_total_rede_cd", title="Capital Total (CD + Lojas) por Curva (R$)",
            text_auto=".2s", color="curva_abc", color_discrete_map=MAPA_CORES_ABC,
        )
        fig_valor.update_layout(xaxis_title="Curva ABC", yaxis_title="R$ (Atacado)", showlegend=False)
        st.plotly_chart(padronizar_grafico(fig_valor), width="stretch")

    with cg2:
        maiores_grupos = df.groupby("grupo")["capital_total_rede_cd"].sum().nlargest(10).index
        df_grupos_top = df[df["grupo"].isin(maiores_grupos)]
        resumo_grupo = df_grupos_top.groupby(["grupo", "curva_abc"])["capital_total_rede_cd"].sum().reset_index()
        fig_kat = px.bar(
            resumo_grupo, x="capital_total_rede_cd", y="grupo", color="curva_abc", orientation="h",
            title="Top 10 Grupos por Capital Total (CD + Lojas)",
            text_auto=".2s", color_discrete_map=MAPA_CORES_ABC,
        )
        fig_kat.update_layout(yaxis={"categoryorder": "total ascending"}, yaxis_title="", xaxis_title="R$ Total", legend_title="Curva")
        st.plotly_chart(padronizar_grafico(fig_kat), width="stretch")

    with cg3:
        df_estagnados = df[(df["curva_abc"] == "C") & (df["estoque_total_empresa"] > 0) & (df["qt_venda"] <= 5)].nlargest(10, "capital_total_rede_cd")
        fig_estag = px.bar(
            df_estagnados, x="capital_total_rede_cd", y="produto", orientation="h",
            title="Top 10 Curva C Estagnados (Maior Capital Empatado)",
            text_auto=".2s", color="capital_total_rede_cd", color_continuous_scale="Oranges",
        )
        fig_estag.update_layout(yaxis={"categoryorder": "total ascending"}, yaxis_title="", xaxis_title="R$ Parado", coloraxis_showscale=False)
        st.plotly_chart(padronizar_grafico(fig_estag), width="stretch")


# ------------------------------------------
# ABA 3: SIMULADOR DE RECOMPRA & AUDITORIA DE RMA (ON-DEMAND)
# ------------------------------------------
with aba_recompra:
    st.markdown('<div class="secao-tag">LABORATÓRIO WHAT-IF • SIMULAÇÃO DE RECOMPRA E RMA</div>', unsafe_allow_html=True)
    st.markdown("#### 🚢 Simulador de Impacto da Planilha de Recompra (A Receber + RMA)")

    if not recompra_ativa:
        st.info(
            "👈 **Como usar:** Suba a planilha mensal de **Recompra** (feita via PROCV) na caixa **'🚢 Simulador de Recompra & RMA'** da barra lateral esquerda.\n\n"
            "Assim que você subir o arquivo, esta aba mostrará automaticamente:\n"
            "1. **Antes vs. Depois da Capacidade de Suprimento (`Cap. Sup.`):** Como fica o estoque de cada produto somando o que está para chegar;\n"
            "2. **Auditoria de RMA (Taxa de Defeito):** Ranking dos produtos com maior índice de defeito para travar recompras ruins;\n"
            "3. **Alerta de Compra Desnecessária:** Produtos que estão na planilha de recompra mas já possuem mais de 6 meses de cobertura em estoque."
        )
    else:
        kr1, kr2, kr3, kr4 = st.columns(4)
        kr1.metric("Total Unidades a Receber", formatar_inteiro(df["quantidade_a_receber"].sum()))
        kr2.metric("SKUs com Pedido de Recompra", formatar_inteiro((df["quantidade_a_receber"] > 0).sum()))
        kr3.metric("Alertas de Superestoque na Recompra", formatar_inteiro(((df["quantidade_a_receber"] > 0) & (df["cap_sup_total"] >= 6)).sum()))
        kr4.metric("SKUs com RMA Crítico (> 5%)", formatar_inteiro((df["rma_pct_display"] > 5).sum()))

        col_sim1, col_sim2 = st.columns(2)

        with col_sim1:
            df_com_rec = df[df["quantidade_a_receber"] > 0].nlargest(12, "quantidade_a_receber").copy()
            if not df_com_rec.empty:
                df_melt_sim = df_com_rec.melt(
                    id_vars=["produto", "codigo"],
                    value_vars=["cap_sup_total", "cap_sup_pos_recompra"],
                    var_name="Cenario",
                    value_name="Meses de Cobertura",
                )
                df_melt_sim["Cenario"] = df_melt_sim["Cenario"].map(
                    {"cap_sup_total": "1. Cap. Sup. Atual (Sem Recompra)", "cap_sup_pos_recompra": "2. Cap. Sup. Projetada (Com Recompra)"}
                )
                fig_sim_cap = px.bar(
                    df_melt_sim,
                    x="Meses de Cobertura",
                    y="produto",
                    color="Cenario",
                    barmode="group",
                    orientation="h",
                    title="Impacto da Recompra na Capacidade de Suprimento (Top 12 Pedidos)",
                    color_discrete_map={"1. Cap. Sup. Atual (Sem Recompra)": "#64748B", "2. Cap. Sup. Projetada (Com Recompra)": COR_LARANJA},
                )
                fig_sim_cap.update_layout(yaxis={"categoryorder": "total ascending"}, yaxis_title="", legend_title="")
                st.plotly_chart(padronizar_grafico(fig_sim_cap, altura=400), width="stretch")

        with col_sim2:
            df_rma_sim = df[df["rma_pct_display"] > 0].nlargest(12, "rma_pct_display").copy()
            if df_rma_sim.empty:
                st.info("Nenhuma coluna de RMA/Defeito com valores > 0 foi detectada na planilha de recompra enviada.")
            else:
                fig_rma_s = px.bar(
                    df_rma_sim,
                    x="rma_pct_display",
                    y="produto",
                    orientation="h",
                    title="Top 12 Maiores Taxas de RMA Detectadas na Planilha (%)",
                    text_auto=".1f",
                    color="rma_pct_display",
                    color_continuous_scale="OrRd",
                )
                fig_rma_s.update_layout(yaxis={"categoryorder": "total ascending"}, yaxis_title="", xaxis_title="Taxa RMA (%)", coloraxis_showscale=False)
                st.plotly_chart(padronizar_grafico(fig_rma_s, altura=400), width="stretch")

        st.markdown("#### 🚨 Auditoria da Planilha de Recompra Cruzada com o Estoque Oficial")
        st.download_button(
            "📥 Exportar Auditoria Completa da Recompra (CSV)",
            data=gerar_arquivo_csv(df_recompra_detalhe),
            file_name=f"auditoria_recompra_loft_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
        )
        st.dataframe(
            df_recompra_detalhe[
                [
                    "codigo", "produto", "grupo", "curva_abc", "qt_Venda" if "qt_Venda" in df_recompra_detalhe.columns else "qt_venda",
                    "qt_est_rede", "qt_estoque", "cap_sup_total", "quantidade_a_receber", "cap_sup_pos_recompra", "rma_pct_display",
                ]
            ].sort_values("quantidade_a_receber", ascending=False),
            hide_index=True,
            width="stretch",
        )


# ==========================================
# 9. MOTOR DE INTELIGÊNCIA E SEPARAÇÃO EXATA DE VARIANTES (IA)
# ==========================================
def extrair_variantes_modelo(texto: str, modelo_base: str) -> set[str]:
    txt = texto.lower().replace("+", " plus ").replace("/", " / ")
    txt = re.sub(rf"\b({re.escape(modelo_base)})\s+(?:padrao|padrão|normal|base)\b", r"\1", txt)

    padrao = rf"\b{re.escape(modelo_base)}(?:\s+(?:pro\s+max|pro|max|air|plus|ultra|mini|fe|lite|e))*\b"
    achados = re.findall(padrao, txt)
    variantes = {" ".join(m.split()) for m in achados}

    if variantes and re.search(r"\bpro\s+max\b", txt):
        variantes.add(f"{modelo_base} pro max")

    return variantes


def rotular_variante_produto(nome_produto: str, modelos_pesquisados: list[str]) -> str:
    if not modelos_pesquisados:
        return "Geral"
    etiquetas = []
    for mod in modelos_pesquisados:
        vars_prod = extrair_variantes_modelo(str(nome_produto), mod)
        for v in sorted(vars_prod):
            if v == mod:
                etiquetas.append(f"{mod.upper()} (PADRÃO)")
            else:
                etiquetas.append(v.upper())
    return " | ".join(etiquetas) if etiquetas else "Outros"


def construir_inteligencia_ia(
    base_total: pd.DataFrame,
    df_hist: pd.DataFrame,
    mensagens_chat: list,
    forcar_estoque_positivo: bool = True,
) -> tuple[str, str, str]:
    cols_ia = [
        c for c in [
            "codigo", "produto", "grupo", "curva_abc", "qt_estoque", "qt_est_rede",
            "val_custo", "preco_lojas", "capital_total_rede_cd", "qt_venda",
            "val_total_venda", "val_total_shop", "val_total_aero", "cap_sup_total",
            "quantidade_a_receber", "rma_pct_display",
        ] if c in base_total.columns
    ]

    df_limpo = base_total[~base_total["grupo"].isin(["SERVIÇOS", "USO PRÓPRIO"])].copy()

    if forcar_estoque_positivo:
        df_ativo = df_limpo[
            (df_limpo["estoque_total_empresa"] > 0)
            | (df_limpo["quantidade_a_receber"] > 0)
            | (df_limpo["qt_venda"] > 0)
        ].copy()
    else:
        df_ativo = df_limpo.copy()

    df_com_estoque = df_limpo[df_limpo["estoque_total_empresa"] > 0].copy()

    resumo_curva = (
        df_com_estoque.groupby("curva_abc")
        .agg(
            skus=("codigo", "count"),
            estoque_cd=("qt_estoque", "sum"),
            estoque_lojas=("qt_est_rede", "sum"),
            capital_total_rs=("capital_total_rede_cd", "sum"),
            unidades_vendidas=("qt_venda", "sum"),
            faturamento_rs=("val_total_venda", "sum"),
        )
        .reset_index()
        .to_csv(index=False)
    )

    top_melhores = (
        df_limpo[df_limpo["qt_venda"] > 0]
        .nlargest(15, "val_total_venda")[cols_ia]
        .to_csv(index=False)
    )

    piores_capital_c = (
        df_com_estoque[(df_com_estoque["curva_abc"] == "C") & (df_com_estoque["qt_venda"] <= 5)]
        .nlargest(15, "capital_total_rede_cd")[cols_ia]
        .to_csv(index=False)
    )

    bloco_diretoria_ia = ""
    if not df_hist.empty and len(df_hist["mes_ref"].unique()) >= 2:
        meses_disp = ordenar_meses_cronologicamente(list(df_hist["mes_ref"].unique()))
        m_atu, m_ant = meses_disp[-1], meses_disp[-2]
        d_atu = df_hist[df_hist["mes_ref"] == m_atu].copy()
        d_ant = df_hist[df_hist["mes_ref"] == m_ant].copy()

        m_p = pd.merge(
            d_atu[["codigo", "produto", "qt_venda", "tabela_atacado", "preco_lojas", "cap_sup_rede", "cap_sup_cd", "cap_sup_total"]],
            d_ant[["codigo", "qt_venda", "tabela_atacado", "preco_lojas"]],
            on="codigo",
            suffixes=("_atu", "_ant"),
        )
        m_p["diff_atacado"] = (m_p["tabela_atacado_atu"] - m_p["tabela_atacado_ant"]).round(2)
        m_p["diff_loja"] = (m_p["preco_lojas_atu"] - m_p["preco_lojas_ant"]).round(2)
        m_p["var_venda"] = m_p["qt_venda_atu"] - m_p["qt_venda_ant"]

        red = m_p[(m_p["diff_atacado"] < -0.5) | (m_p["diff_loja"] < -0.5)]
        top_red_subiu = red.nlargest(10, "var_venda")[["codigo", "produto", "diff_loja", "diff_atacado", "qt_venda_ant", "qt_venda_atu", "var_venda"]].to_csv(index=False)
        top_red_caiu = red.nsmallest(10, "var_venda")[["codigo", "produto", "diff_loja", "diff_atacado", "qt_venda_ant", "qt_venda_atu", "var_venda"]].to_csv(index=False)

        try:
            mn, an = [int(x) for x in m_atu.split("-")]
            dt_ref = pd.Timestamp(year=2000 + an, month=mn, day=1) + pd.offsets.MonthEnd(1)
        except Exception:
            dt_ref = d_atu["dt_1_rec"].max()
        d_atu["dias_chegada"] = (dt_ref - d_atu["dt_1_rec"]).dt.days
        rec_1_3m = (
            d_atu[(d_atu["dias_chegada"] >= 30) & (d_atu["dias_chegada"] <= 95)]
            .nlargest(15, "cap_sup_total")[["codigo", "produto", "dt_1_rec", "dias_chegada", "qt_venda", "qt_est_rede", "estoque_cd", "cap_sup_rede", "cap_sup_cd", "cap_sup_total"]]
            .to_csv(index=False)
        )

        bloco_diretoria_ia = (
            f"4. REDUÇÃO DE PREÇO ({m_ant} -> {m_atu}) - TOP 10 QUE MAIS AUMENTARAM VENDAS:\n{top_red_subiu}\n"
            f"5. REDUÇÃO DE PREÇO ({m_ant} -> {m_atu}) - TOP 10 QUE MAIS CAÍRAM VENDAS:\n{top_red_caiu}\n"
            f"6. SUPERESTOCADOS COM 1 A 3 MESES DE CHEGADA (MAIOR CAP. SUPRIMENTO EM {m_atu}):\n{rec_1_3m}\n"
        )

    dossie_python = (
        "=== DOSSIÊ MATEMÁTICO PRÉ-CALCULADO PELO SISTEMA (100% PRECISO) ===\n"
        f"1. RESUMO POR CURVA ABC (ESTOQUE CD + LOJAS):\n{resumo_curva}\n"
        f"2. TOP 15 CAMPEÕES DE FATURAMENTO E GIRO (COM SHOPPING VS AEROPORTO):\n{top_melhores}\n"
        f"3. TOP 15 PIORES PRODUTOS FINANCEIROS (CURVA C COM ALTO CAPITAL PARADO):\n{piores_capital_c}\n"
        f"{bloco_diretoria_ia}"
    )

    msgs_user = [m["content"].lower() for m in mensagens_chat if m["role"] == "user"]
    pergunta_atual = msgs_user[-1] if msgs_user else ""

    palavras_neutras = {
        "me", "de", "do", "da", "dos", "das", "um", "uma", "comparativo", "vendas", "venda",
        "estoque", "para", "com", "sem", "no", "na", "nos", "nas", "banco", "dados", "tem",
        "so", "só", "somente", "apenas", "essas", "esses", "quais", "qual", "como", "esta",
        "estão", "sao", "são", "todos", "todas", "lista", "mostre", "analise", "entre", "e",
        "ou", "que", "mais", "menos", "piores", "pior", "melhores", "melhor", "produtos",
        "produto", "itens", "item", "fale", "sobre", "resumo", "geral", "empresa", "loft",
        "situacao", "situação", "atual", "hoje", "quero", "ver", "traga", "puxa", "puxe",
        "padrao", "padrão", "normal", "base", "modelo", "modelos", "linha", "deles", "delas",
        "pro", "max", "air", "plus", "ultra", "mini", "fe", "lite",
    }

    termos_atuais = [
        p for p in re.findall(r"[a-z0-9]+", pergunta_atual)
        if p not in palavras_neutras and len(p) >= 2
    ]

    if not termos_atuais and len(msgs_user) >= 2:
        texto_referencia = msgs_user[-2]
        termos_busca = [
            p for p in re.findall(r"[a-z0-9]+", texto_referencia)
            if p not in palavras_neutras and len(p) >= 2
        ]
    else:
        texto_referencia = pergunta_atual
        termos_busca = termos_atuais

    texto_combinado = (
        df_ativo["codigo"].astype(str) + " " +
        df_ativo["produto"].astype(str) + " " +
        df_ativo["grupo"].astype(str)
    ).str.lower()

    bloco_busca_especifica = ""
    if termos_busca:
        pontuacao = pd.Series(0, index=df_ativo.index)
        for termo in termos_busca:
            pontuacao += texto_combinado.str.contains(re.escape(termo), na=False).astype(int)

        maior_nota = pontuacao.max()
        if maior_nota > 0:
            nota_corte = max(1, maior_nota - 1) if maior_nota >= 2 else 1
            encontrados = df_ativo[pontuacao >= nota_corte].copy()
            encontrados["relevancia"] = pontuacao[pontuacao >= nota_corte]

            modelos_num = [t for t in termos_busca if re.match(r"^(?:[a-z]{1,2}\d{1,3}|\d{1,3})$", t)]
            pediu_familia_toda = any(
                termo_fam in texto_referencia
                for termo_fam in ["toda a linha", "todos os modelos", "todas as vers", "familia", "família"]
            )

            descricao_filtro_variante = "Todas as variações da linha"
            if modelos_num and not pediu_familia_toda:
                mascara_variante_exata = pd.Series(True, index=encontrados.index)
                variantes_alvo_resumo = []

                for mod in modelos_num:
                    variantes_pedidas = extrair_variantes_modelo(texto_referencia, mod)
                    if variantes_pedidas:
                        variantes_alvo_resumo.extend(sorted(variantes_pedidas))

                        def produto_bate_variante(nome_prod: str) -> bool:
                            vars_no_produto = extrair_variantes_modelo(str(nome_prod), mod)
                            if not vars_no_produto:
                                return False
                            return bool(vars_no_produto.intersection(variantes_pedidas))

                        mascara_mod = encontrados["produto"].apply(produto_bate_variante)
                        if mascara_mod.any():
                            mascara_variante_exata &= mascara_mod

                encontrados = encontrados[mascara_variante_exata].copy()
                if variantes_alvo_resumo:
                    descricao_filtro_variante = ", ".join([v.upper() for v in variantes_alvo_resumo])

            encontrados["variante_modelo"] = encontrados["produto"].apply(
                lambda nome: rotular_variante_produto(nome, modelos_num)
            )

            encontrados = encontrados.sort_values(
                ["relevancia", "variante_modelo", "qt_venda", "qt_estoque", "capital_total_rede_cd"],
                ascending=[False, True, False, False, False],
            )

            cols_recorte = ["codigo", "variante_modelo"] + [c for c in cols_ia if c != "codigo"]

            sub_skus = len(encontrados)
            sub_est_cd = encontrados["qt_estoque"].sum()
            sub_est_lj = encontrados["qt_est_rede"].sum()
            sub_ven = encontrados["qt_venda"].sum()
            sub_fat = encontrados["val_total_venda"].sum()

            recorte_csv = encontrados.head(250)[cols_recorte].to_csv(index=False)
            bloco_busca_especifica = (
                f"🎯 RECORTE EXATO E FILTRADO PARA OS TERMOS ({', '.join(termos_busca)}) | "
                f"FILTRO DE VARIANTE APLICADO PELO PYTHON: [{descricao_filtro_variante}]\n"
                f"Subtotais exatos deste recorte: {sub_skus} SKUs | "
                f"Estoque CD: {formatar_inteiro(sub_est_cd)} un. | Estoque Lojas: {formatar_inteiro(sub_est_lj)} un. | "
                f"Vendas: {formatar_inteiro(sub_ven)} un. | Faturamento: {formatar_moeda(sub_fat)}\n"
                f"Lista Exata de Produtos:\n{recorte_csv}\n"
            )

    df_contexto_geral = (
        df_com_estoque.sort_values(["capital_total_rede_cd", "qt_venda"], ascending=[False, False])
        .head(350)[cols_ia]
        .to_csv(index=False)
    )

    return dossie_python, bloco_busca_especifica, df_contexto_geral


# ------------------------------------------
# ABA 4: ASSISTENTE EXECUTIVO DE IA
# ------------------------------------------
with aba_ia:
    col_cab_ia, col_filtro_ia, col_reset_ia = st.columns([3, 2, 1])
    with col_cab_ia:
        st.markdown("#### 🧠 Assistente Executivo de Inventário LOFT")
        st.caption("Conectado 100% à base oficial ('5 - Análise Estoque.xlsx' / Supabase Cloud) + Simulador de Recompra/RMA.")
    with col_filtro_ia:
        ignorar_zerados_ia = st.checkbox("Ignorar cadastro morto (Focar no Estoque Ativo)", value=True)
    with col_reset_ia:
        if st.button("🗑️ Reiniciar Conversa", use_container_width=True):
            st.session_state.chat = []
            st.rerun()

    if "chat" not in st.session_state:
        st.session_state.chat = []

    st.markdown("**Sugestões de Análise Rápida:**")
    bt1, bt2, bt3, bt4 = st.columns(4)
    atalho_selecionado = None
    if bt1.button("🏷️ Redução de Preço vs Vendas", use_container_width=True):
        atalho_selecionado = "Quais produtos tiveram redução de preço no último mês e mais aumentaram vendas? E quais mais caíram vendas mesmo reduzindo preço?"
    if bt2.button("📦 Superestocados (1 a 3 meses)", use_container_width=True):
        atalho_selecionado = "Liste os itens com chegada entre 1 e 3 meses que estão com a maior Capacidade de Suprimento (Cap. Sup. Rede e CD)."
    if bt3.button("🏬 Shopping vs Aeroporto", use_container_width=True):
        atalho_selecionado = "Compare o desempenho dos nossos principais produtos entre as lojas de Shopping e as lojas de Aeroporto."
    if bt4.button("📱 Comparar iPhone 17 vs 17 Pro", use_container_width=True):
        atalho_selecionado = "Compare as capas e películas do iPhone 17 e 17 Pro, separando em duas tabelas distintas o modelo iPhone 17 (Padrão) e o modelo iPhone 17 Pro."

    janela_chat = st.container(height=460)

    with janela_chat:
        if not st.session_state.chat:
            st.info("Clique em um dos atalhos acima ou digite sua pergunta abaixo.")
        for mensagem in st.session_state.chat:
            with st.chat_message(mensagem["role"]):
                st.markdown(mensagem["content"])

    entrada_usuario = st.chat_input("Pergunte sobre reduções de preço, cap. suprimento, shopping vs aeroporto, iPhone 17 ou recompra...")
    pergunta_ativa = atalho_selecionado or entrada_usuario

    if pergunta_ativa:
        st.session_state.chat.append({"role": "user", "content": pergunta_ativa})
        with janela_chat:
            with st.chat_message("user"):
                st.markdown(pergunta_ativa)

            with st.chat_message("assistant"):
                with st.spinner("Analisando base oficial de estoque..."):
                    try:
                        cliente_ia = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

                        dossie_matematico, recorte_focado, csv_top_ativos = construir_inteligencia_ia(
                            df_completo,
                            df_historico_mensal,
                            st.session_state.chat,
                            forcar_estoque_positivo=ignorar_zerados_ia,
                        )

                        prompt_sistema = (
                            "Você é o Diretor de Inteligência de Supply Chain da LOFT Design.\n"
                            "DIRETRIZES INEGOCIÁVEIS DE NEGÓCIO:\n"
                            "1. ISOLAMENTO RIGOROSO DE VARIANTES DE MODELO: Quando houver um 'RECORTE EXATO E FILTRADO', utilize APENAS os produtos listados nele. Se o usuário pediu apenas 'iPhone 17', mostre exclusivamente `17 (PADRÃO)`.\n"
                            "2. MÉTRICAS DA DIRETORIA: Use as seções 4, 5 e 6 do DOSSIÊ para responder sobre elasticidade de redução de preço e itens superestocados entre 1 e 3 meses da chegada.\n"
                            "3. RECOMPRA E RMA SOB DEMANDA: Os dados de 'A Receber' e 'RMA' vêm da Planilha de Recompra quando enviada na sessão (veja a seção 'PLANILHA DE RECOMPRA E RMA ATIVA').\n"
                            "4. PROIBIDO CITAR PRODUTOS ZERADOS COMO PROBLEMA: Nunca aponte produtos sem estoque como 'piores produtos'.\n\n"
                            f"{resumo_recompra_ia}\n\n"
                            f"{recorte_focado}\n\n"
                            f"{dossie_matematico}\n\n"
                            f"AMOSTRA DOS MAIORES PRODUTOS EM ESTOQUE (CSV):\n{csv_top_ativos}"
                        )

                        ultimas_interacoes = st.session_state.chat[-6:]
                        resposta_ia = cliente_ia.chat.completions.create(
                            model="gpt-4o-mini",
                            messages=[{"role": "system", "content": prompt_sistema}] + ultimas_interacoes,
                            temperature=0.0,
                        )
                        conteudo_resposta = resposta_ia.choices[0].message.content
                        st.markdown(conteudo_resposta)
                        st.session_state.chat.append({"role": "assistant", "content": conteudo_resposta})
                    except Exception as erro_ia:
                        st.error(f"Erro na consulta OpenAI: {erro_ia}")


# ------------------------------------------
# ABA 5: EXPLORADOR DA BASE DE DADOS COMPLETA
# ------------------------------------------
with aba_tabela:
    col_pesq, col_exp = st.columns([3, 1])
    with col_pesq:
        busca_tabela = st.text_input(
            "🔎 Pesquisar por Código, Nome do Produto, Grupo ou Família:",
            placeholder="Ex: LEI364GR, Magsafe, iPhone 17, Power Bank...",
        )

    df_mostrado = df.copy()
    if busca_tabela.strip():
        termo_b = busca_tabela.strip().lower()
        df_mostrado = df_mostrado[
            df_mostrado["codigo"].astype(str).str.lower().str.contains(re.escape(termo_b), na=False)
            | df_mostrado["produto"].astype(str).str.lower().str.contains(re.escape(termo_b), na=False)
            | df_mostrado["grupo"].astype(str).str.lower().str.contains(re.escape(termo_b), na=False)
        ]

    with col_exp:
        st.write("")
        st.download_button(
            label="📥 Exportar Tabela (CSV)",
            data=gerar_arquivo_csv(df_mostrado),
            file_name=f"estoque_oficial_loft_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    cols_exibir_tab = [
        "codigo", "produto", "grupo", "familia", "curva_abc", "qt_estoque", "qt_est_rede",
        "qt_venda", "cap_sup_rede", "cap_sup_cd", "cap_sup_total", "val_custo", "preco_lojas",
        "markup", "val_total_venda", "val_total_shop", "val_total_aero", "quantidade_a_receber", "rma_pct_display",
    ]
    cols_exibir_tab = [c for c in cols_exibir_tab if c in df_mostrado.columns]

    st.dataframe(df_mostrado[cols_exibir_tab], hide_index=True, width="stretch")
