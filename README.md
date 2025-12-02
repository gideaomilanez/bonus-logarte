# 📊 Bônus de Desempenho – LogArte

Aplicativo em **Python + Streamlit** para cálculo e análise do **bônus de desempenho dos motoristas** da LogArte, a partir das planilhas de **Controle de viagens**.

## ✨ Funcionalidades

- Upload de múltiplos arquivos Excel com a aba `Controle de viagens`
- Filtro por intervalo de datas
- Cálculo automático de bônus por:
  - Motorista
  - Centro de custo
- Tabelas de resumo:
  - Bônus por motorista e centro de custo
  - Bônus total por motorista
  - Bônus por centro de custo
  - Dias trabalhados por motorista
- Gráficos:
  - Bônus por motorista (barras horizontais)
  - Evolução do faturamento
  - Heatmap de dias trabalhados
- Download de um Excel com todos os resultados consolidados

## 🛠️ Tecnologias

- Python
- Streamlit
- Pandas
- Matplotlib
- Seaborn
- OpenPyXL / XlsxWriter

## 📦 Instalação

```bash
pip install -r requirements.txt
