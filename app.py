"""
Bônus de Desempenho

Funcionalidades principais:
- Upload de múltiplos arquivos Excel da planilha de viagens
- Filtro por intervalo de datas
- Limpeza e padronização dos dados
- Cálculo de bônus por centro de custo e por motorista
- Tabelas de resumo:
    * Bônus por motorista e centro de custo
    * Bônus total por motorista
    * Bônus total por centro de custo
    * Dias trabalhados por motorista
- Gráficos:
    * Bônus dos motoristas (barras horizontais empilhadas)
    * Evolução do faturamento ao longo do tempo
    * Heatmap de dias trabalhados por motorista
- Exportação para Excel com botão de download
"""

from __future__ import annotations

import io
from datetime import date
from typing import Dict, Tuple, Optional

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
import streamlit as st

# ------------------------------------------------------------
# CONFIGURAÇÃO VISUAL GLOBAL
# ------------------------------------------------------------
plt.style.use("ggplot")
sns.set_palette("husl")

# Paleta de cores usada no notebook original
AZUL1, AZUL2, AZUL3, AZUL4, AZUL5 = "#03045e", "#0077b6", "#00b4d8", "#90e0ef", "#CDDBF3"
CINZA1, CINZA2, CINZA3, CINZA4, CINZA5 = "#212529", "#495057", "#adb5bd", "#dee2e6", "#f8f9fa"
VERMELHO1, LARANJA1, AMARELO1, VERDE1, VERDE2 = "#e76f51", "#f4a261", "#e9c46a", "#4c956c", "#2a9d8f"

# ------------------------------------------------------------
# CONFIG STREAMLIT
# ------------------------------------------------------------
st.set_page_config(
    page_title="Bônus Desempenho LogArte",
    layout="wide",
)

st.title("📊 Cálculo de Bônus")

st.markdown(
    """
Este painel lê as planilhas de **Controle de viagens**, aplica as regras de
**bônus por centro de custo** e gera relatórios e gráficos de desempenho
dos motoristas.
"""
)


# ------------------------------------------------------------
# FUNÇÕES AUXILIARES
# ------------------------------------------------------------
def carregar_arquivos(files) -> pd.DataFrame:
    """Lê todos os arquivos enviados e concatena em um único DataFrame."""
    dfs = []
    for f in files:
        try:
            df = pd.read_excel(
                f,
                sheet_name="Controle de viagens",
                skiprows=2,
                usecols="A:O",
                parse_dates=["DATA"],
            )
            dfs.append(df)
        except Exception as e:
            st.error(f"Erro ao ler o arquivo '{f.name}': {e}")
            raise
    if not dfs:
        raise ValueError("Nenhum dado foi carregado dos arquivos enviados.")
    return pd.concat(dfs, ignore_index=True)


def limpar_e_filtrar_dados(
    dados: pd.DataFrame,
    data_ini: date,
    data_fim: date,
) -> pd.DataFrame:
    """Replica a etapa de limpeza e filtro do notebook original."""

    # Garantir tipo datetime
    dados["DATA"] = pd.to_datetime(dados["DATA"], errors="coerce")

    # Padronizar MOTORISTA
    dados["MOTORISTA"] = dados["MOTORISTA"].astype(str).str.upper().str.strip()
    dados["MOTORISTA"] = dados["MOTORISTA"].str.replace(
        r"\bVINICIUS\b", "VINÍCIUS", regex=True
    )
    dados["MOTORISTA"] = dados["MOTORISTA"].str.replace(
        r"\bMARCOS NASCIMENTO\b", "MARCOS", regex=True
    )

    # Remover linhas inválidas
    dados = dados.dropna(subset=["MOTORISTA", "DATA"])

    # Filtro temporal
    data_ini_ts = pd.to_datetime(data_ini)
    data_fim_ts = pd.to_datetime(data_fim)
    mask = (dados["DATA"] >= data_ini_ts) & (dados["DATA"] <= data_fim_ts)
    dados_filtrados = dados.loc[mask].copy()

    if dados_filtrados.empty:
        raise ValueError("Nenhum registro no intervalo de datas selecionado.")

    return dados_filtrados


def calcular_bonus(dados_filtrados: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Aplica as regras de bônus e gera tabelas de resumo."""

    # Regras de bônus:
    # - FRETE BRITA / AREIA: 1,00 por tonelada (QUANT.)
    # - FRETE CIMENTO / FRETE ADITIVO: 2% do TOTAL (R$)
    conditions = [
        dados_filtrados["CENTRO DE CUSTO"].isin(["FRETE BRITA", "AREIA"]),
        dados_filtrados["CENTRO DE CUSTO"].isin(["FRETE CIMENTO", "FRETE ADITIVO"]),
    ]
    calculos = [
        dados_filtrados["QUANT."] * 1.0,
        dados_filtrados["TOTAL (R$)"] * 0.02,
    ]

    dados_filtrados["BÔNUS"] = np.select(conditions, calculos, default=0).round(2)

    # Tabela por motorista + centro de custo
    tabela = dados_filtrados.groupby(["MOTORISTA", "CENTRO DE CUSTO"]).agg(
        VIAGENS=("QUANT.", "count"),
        QUANT=("QUANT.", "sum"),
        FATURAMENTO=("TOTAL (R$)", "sum"),
        BÔNUS=("BÔNUS", "sum"),
    ).round(2)

    # Resumo por motorista
    bonus_motorista = dados_filtrados.groupby("MOTORISTA").agg(
        VIAGENS=("QUANT.", "count"),
        BÔNUS=("BÔNUS", "sum"),
    ).round(2).sort_values("BÔNUS", ascending=False)

    # Adiciona BÔNUS TOTAL na tabela detalhada (apenas na primeira linha de cada motorista)
    tabela["BÔNUS TOTAL"] = tabela.index.get_level_values("MOTORISTA").map(
        bonus_motorista["BÔNUS"]
    )
    tabela["BÔNUS TOTAL"] = tabela.groupby("MOTORISTA")["BÔNUS TOTAL"].transform(
        lambda x: x.where(x.index == x.index[0], None)
    )

    # Resumo por centro de custo
    resumo_centro_custo = (
        dados_filtrados.groupby("CENTRO DE CUSTO")
        .agg(BÔNUS_TOTAL=("BÔNUS", "sum"))
        .round(2)
        .sort_values("BÔNUS_TOTAL", ascending=False)
    )

    return tabela, bonus_motorista, resumo_centro_custo


def gerar_nome_periodo(data_ini: date, data_fim: date) -> str:
    """Gera string amigável de período (ex.: '1 a 15 de Janeiro')."""

    meses_pt = {
        1: "Janeiro",
        2: "Fevereiro",
        3: "Março",
        4: "Abril",
        5: "Maio",
        6: "Junho",
        7: "Julho",
        8: "Agosto",
        9: "Setembro",
        10: "Outubro",
        11: "Novembro",
        12: "Dezembro",
    }

    dia_i = data_ini.day
    mes_i = meses_pt[data_ini.month]
    ano_i = data_ini.year

    dia_f = data_fim.day
    mes_f = meses_pt[data_fim.month]
    ano_f = data_fim.year

    if ano_i == ano_f:
        if mes_i == mes_f:
            return f"{dia_i} a {dia_f} de {mes_i} de {ano_i}"
        else:
            return f"{dia_i} de {mes_i} a {dia_f} de {mes_f} de {ano_i}"
    else:
        return f"{dia_i} de {mes_i} de {ano_i} a {dia_f} de {mes_f} de {ano_f}"


def grafico_bonus_motoristas(bonus_motorista: pd.DataFrame, nome_data: str):
    """Gera o gráfico de barras horizontais empilhadas de bônus por motorista."""

    cores = [VERDE2, VERMELHO1]  # VIAGENS, BÔNUS

    fig, ax = plt.subplots(figsize=(10, 6))

    bonus_motorista.plot(kind="barh", stacked=True, ax=ax, color=cores)

    ax.set_title("Bônus dos motoristas\n", loc="left", fontsize=22, color=CINZA1)
    ax.text(
        0,
        1.0,
        nome_data,
        transform=ax.transAxes,
        ha="left",
        va="bottom",
        fontsize=18,
        color=CINZA2,
    )
    ax.set_xlabel("")
    ax.set_ylabel("")
    ax.xaxis.set_tick_params(labelsize=14, labelcolor=CINZA2, rotation=0)
    ax.set_frame_on(False)

    # remover todos os ticks do eixo x e y
    ax.tick_params(axis="both", which="both", length=0)

    # Adicionar valores totais no fim das barras
    totais = bonus_motorista.sum(axis=1)
    offset = 0.01 * totais.max() if len(totais) else 0

    for i, total in enumerate(totais):
        ax.text(
            total + offset,
            i,
            f"R$ {total:,.2f}".replace(",", "X")
            .replace(".", ",")
            .replace("X", "."),
            ha="left",
            va="center",
            fontsize=12,
            color="#2f2f2f",
            clip_on=False,
            zorder=3,
        )

    ax.margins(x=0.05)

    ax.legend(
        title="",
        loc="upper left",
        bbox_to_anchor=(0.8, 0.4),
        frameon=False,
        ncol=1,
        prop={"size": 12},
    )

    plt.tight_layout()
    return fig


def grafico_faturamento(dados_filtrados: pd.DataFrame):
    """Evolução do faturamento ao longo do tempo (linha)."""
    sns.set_theme(style="white")
    plt.rcParams["font.family"] = "DejaVu Sans"

    df = dados_filtrados.copy()
    df["DATA"] = pd.to_datetime(df["DATA"])

    faturamento_por_dia = (
        df.groupby("DATA")["TOTAL (R$)"].sum().reset_index()
    )

    fig, ax = plt.subplots(figsize=(12, 6))
    sns.lineplot(
        data=faturamento_por_dia,
        x="DATA",
        y="TOTAL (R$)",
        marker="o",
        color="green",
        linewidth=2,
        ax=ax,
    )

    ax.set_title(
        "Evolução do Faturamento ao Longo do Tempo",
        fontsize=16,
        pad=20,
        fontweight="bold",
        color="#2f2f2f",
    )
    ax.set_xlabel("Data", fontsize=12, labelpad=15)
    ax.set_ylabel("Faturamento Total (R$)", fontsize=12)

    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d/%m"))
    ax.xaxis.set_major_locator(
        mdates.WeekdayLocator(byweekday=[mdates.MO, mdates.FR])
    )
    ax.yaxis.set_major_locator(plt.MaxNLocator(nbins=6))

    ax.grid(True, linestyle="--", alpha=0.6, axis="both")
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    return fig


def grafico_heatmap_trabalho(dados_filtrados: pd.DataFrame):
    """Heatmap dos dias em que os motoristas trabalharam."""

    sns.set_theme(style="whitegrid")
    plt.rcParams["font.family"] = "DejaVu Sans"

    df = dados_filtrados.copy()
    df["DATA"] = pd.to_datetime(df["DATA"])
    df["DIA"] = df["DATA"].dt.date

    matriz_trabalho = pd.pivot_table(
        df,
        index="MOTORISTA",
        columns="DIA",
        aggfunc="size",
        fill_value=0,
    )
    matriz_trabalho = matriz_trabalho.applymap(lambda x: 1 if x > 0 else 0)

    fig, ax = plt.subplots(figsize=(14, 4))
    sns.heatmap(
        matriz_trabalho,
        cmap="Blues",
        cbar=False,
        linewidths=0.5,
        linecolor="gray",
        annot=False,
        ax=ax,
    )

    ax.set_title(
        "Dias em que os Motoristas Trabalharam",
        fontsize=16,
        fontweight="bold",
    )
    ax.set_xlabel("Dia", fontsize=12)
    ax.set_ylabel("Motorista", fontsize=12)

    ax.set_xticks(range(len(matriz_trabalho.columns)))
    ax.set_xticklabels(
        [d.strftime("%d/%m") for d in matriz_trabalho.columns],
        rotation=45,
        fontsize=10,
    )
    ax.set_yticklabels(ax.get_yticklabels(), rotation=0, fontsize=10)

    plt.tight_layout()
    return fig, matriz_trabalho


def tabela_dias_trabalhados(dados_filtrados: pd.DataFrame) -> pd.DataFrame:
    df = dados_filtrados.copy()
    df["DATA"] = pd.to_datetime(df["DATA"])
    df["DIA"] = df["DATA"].dt.date

    dias_trabalhados = df.groupby("MOTORISTA")["DIA"].nunique()
    tabela = dias_trabalhados.reset_index()
    tabela.columns = ["Motorista", "Dias Trabalhados"]
    tabela.set_index("Motorista", inplace=True)
    return tabela


def gerar_excel_para_download(
    nome_arquivo: str,
    tabela: pd.DataFrame,
    bonus_motorista: pd.DataFrame,
    resumo_centro_custo: pd.DataFrame,
    tabela_dias: pd.DataFrame,
) -> bytes:
    """Gera um arquivo Excel com múltiplas abas em memória."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        tabela.to_excel(writer, sheet_name="BÔNUS DETALHADO")
        bonus_motorista.to_excel(writer, sheet_name="BÔNUS POR MOTORISTA")
        resumo_centro_custo.to_excel(writer, sheet_name="BÔNUS POR C.CUSTO")
        tabela_dias.to_excel(writer, sheet_name="DIAS TRABALHADOS")
    buffer.seek(0)
    return buffer.getvalue()


# ------------------------------------------------------------
# INTERFACE STREAMLIT
# ------------------------------------------------------------

st.sidebar.header("⚙️ Configurações")

uploaded_files = st.sidebar.file_uploader(
    "1) Envie os arquivos xlsx de viagens",
    type=["xlsx"],
    accept_multiple_files=True,
    help="Planilhas com a aba 'Controle de viagens'. Você pode enviar mais de um arquivo.",
)

if uploaded_files:
    # Carrega dados brutos
    try:
        dados_raw = carregar_arquivos(uploaded_files)
    except Exception:
        st.stop()

    st.sidebar.success(
        f"{len(uploaded_files)} arquivo(s) carregado(s). Registros totais: {len(dados_raw)}"
    )

    # Sugere intervalo de datas com base nos dados
    if "DATA" in dados_raw.columns:
        datas_validas = pd.to_datetime(dados_raw["DATA"], errors="coerce").dropna()
        if not datas_validas.empty:
            min_data = datas_validas.min().date()
            max_data = datas_validas.max().date()
        else:
            min_data = max_data = date.today()
    else:
        st.error("A coluna 'DATA' não foi encontrada na planilha.")
        st.stop()

    data_ini = st.sidebar.date_input("2) Data inicial", value=min_data)
    data_fim = st.sidebar.date_input("3) Data final", value=max_data)

    if data_ini > data_fim:
        st.sidebar.error("A data inicial não pode ser maior que a data final.")
        st.stop()

    if st.sidebar.button("Calcular bônus"):
        with st.spinner("Processando dados, calculando bônus e gerando gráficos..."):
            try:
                dados_filtrados = limpar_e_filtrar_dados(dados_raw, data_ini, data_fim)
                tabela, bonus_motorista, resumo_centro_custo = calcular_bonus(
                    dados_filtrados
                )
                nome_periodo = gerar_nome_periodo(data_ini, data_fim)
                fig_bonus = grafico_bonus_motoristas(bonus_motorista, nome_periodo)
                fig_fat = grafico_faturamento(dados_filtrados)
                fig_heat, matriz_trabalho = grafico_heatmap_trabalho(dados_filtrados)
                tabela_dias = tabela_dias_trabalhados(dados_filtrados)
            except ValueError as e:
                st.error(str(e))
                st.stop()
            except Exception as e:
                st.error(f"Erro inesperado ao processar os dados: {e}")
                st.stop()

        st.success(
            f"Análise concluída para o período **{nome_periodo}**. "
            f"Registros considerados: **{len(dados_filtrados)}**."
        )

        # --------------------------------------------------------------------
        # VISÃO GERAL / TABELAS
        # --------------------------------------------------------------------
        st.subheader("📋 Amostra da tabela filtrada")
        st.dataframe(dados_filtrados.head(20))

        col1, col2 = st.columns(2)

        with col1:
            st.subheader("Bônus por Motorista e Centro de Custo")
            st.dataframe(tabela)

        with col2:
            st.subheader("Resumo de Bônus por Motorista")
            st.dataframe(bonus_motorista)

        st.subheader("Bônus por Centro de Custo")
        st.dataframe(resumo_centro_custo)

        # --------------------------------------------------------------------
        # GRÁFICOS
        # --------------------------------------------------------------------
        st.subheader("📈 Gráficos")

        st.markdown("### Bônus dos motoristas")
        st.pyplot(fig_bonus)

        st.markdown("### Evolução do faturamento no período")
        st.pyplot(fig_fat)

        st.markdown("### Heatmap de dias trabalhados por motorista")
        st.pyplot(fig_heat)

        # --------------------------------------------------------------------
        # DIAS TRABALHADOS
        # --------------------------------------------------------------------
        st.subheader("📆 Dias que houveram entrega por motorista")
        st.dataframe(tabela_dias)

        # --------------------------------------------------------------------
        # EXPORTAÇÃO EXCEL
        # --------------------------------------------------------------------
        st.subheader("⬇️ Exportar resultados")

        nome_arquivo_excel = f"Bonus_Motoristas_Logarte_{nome_periodo}.xlsx".replace(
            " ", "_"
        )
        excel_bytes = gerar_excel_para_download(
            nome_arquivo_excel,
            tabela,
            bonus_motorista,
            resumo_centro_custo,
            tabela_dias,
        )

        st.download_button(
            label="📥 Baixar planilha de resultados (Excel)",
            data=excel_bytes,
            file_name=nome_arquivo_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

else:
    st.info("Envie os arquivos Excel na barra lateral para iniciar a análise.")
