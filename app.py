import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import re
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Dashboard Seguradora", layout="wide")

# Funções de Mascaramento de CPF
def mascarar_cpf_st(val):
    if pd.isna(val):
        return "N/A"
    digits = re.sub(r'\D', '', str(val)).zfill(11)
    if len(digits) >= 11:
        return f"\*\*\*.\*\*\*.{digits[-5:-2]}-{digits[-2:]}"
    return "\*\*\*.\*\*\*.\*\*\*-\*\*"

def mascarar_cpf_excel(val):
    if pd.isna(val):
        return "N/A"
    digits = re.sub(r'\D', '', str(val)).zfill(11)
    if len(digits) >= 11:
        return f"***.***.{digits[-5:-2]}-{digits[-2:]}"
    return "***.***.***-**"

# Função para classificar barras em relação à média
def classificar_desempenho(val, media, tol=0.05):
    if val > media * (1 + tol):
        return 'Acima da Média'
    elif val < media * (1 - tol):
        return 'Abaixo da Média'
    else:
        return 'Na Média'

# Gerador de Excel Formatado em Memória
def gerar_excel_formatado(df_original):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    # Aba 1: Detalhamento por Participante
    ws = wb.active
    ws.title = "DETALHAMENTO POR PARTICIPANTE"
    
    df_exp = df_original.copy()
    
    # Identifica coluna de CPF para mascarar no Excel
    col_cpf_name = next((c for c in df_exp.columns if 'CPF' in c.upper()), None)
    if col_cpf_name:
        df_exp[col_cpf_name] = df_exp[col_cpf_name].apply(mascarar_cpf_excel)
        
    # Remove colunas auxiliares
    cols_drop = ['DATA_REF', 'ANO', 'MES_ANO', 'TRIMESTRE', 'CPF_MASC']
    df_exp.drop(columns=[c for c in cols_drop if c in df_exp.columns], inplace=True, errors='ignore')
    
    headers = list(df_exp.columns)
    ws.append(headers)
    
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    for row in df_exp.itertuples(index=False):
        ws.append(list(row))
        
    # Formatação de 2 casas decimais e largura
    for row_idx in range(2, ws.max_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                cell.number_format = '#,##0.00'
                
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)
        
    wb.save(output)
    return output.getvalue()

# INTERFACE STREAMLIT
st.title("📊 Dashboard e Relatório - Carteira de Seguros")

st.sidebar.header("1. Upload de Dados")
file = st.sidebar.file_uploader("Envie a planilha Excel (.xlsx)", type=["xlsx"])

if file is not None:
    df = pd.read_excel(file)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Identificação Dinâmica de Colunas
    col_nome = next((c for c in df.columns if 'NOME' in c.upper() or 'PROPONENTE' in c.upper()), df.columns[5])
    col_cpf = next((c for c in df.columns if 'CPF' in c.upper()), df.columns[6])
    col_corretor = next((c for c in df.columns if 'AGENTE' in c.upper() or 'CORRETOR' in c.upper()), df.columns[7])
    col_data = next((c for c in df.columns if 'PARTIR' in c.upper() or 'DATA' in c.upper()), df.columns[8])
    cols_coberturas = df.columns[9:22]
    
    # Tratamento de datas
    df['DATA_REF'] = pd.to_datetime(df[col_data], errors='coerce')
    df['ANO'] = df['DATA_REF'].dt.year.astype(str)
    df['MES_ANO'] = df['DATA_REF'].dt.strftime('%m/%Y')
    df['TRIMESTRE'] = df['DATA_REF'].apply(lambda d: f"{d.year}-Q{(d.month-1)//3 + 1}" if pd.notna(d) else 'N/A')
    
    # CPF Mascarado para exibição
    df['CPF_MASC'] = df[col_cpf].apply(mascarar_cpf_st)
    
    # Mapeamento de Cores para o Gráfico
    color_map = {
        'Acima da Média': '#10B981',  # Verde
        'Na Média': '#2563EB',        # Azul
        'Abaixo da Média': '#EF4444'   # Vermelho
    }
    
    # REQUISITO 1: VENDAS POR PERÍODO
    st.subheader("1. Indicadores de Vendas")
    vis_opcao = st.radio("Selecione a visão temporal:", ["Mês", "Trimestre", "Ano"], horizontal=True)
    
    col_periodo = 'MES_ANO' if vis_opcao == "Mês" else ('TRIMESTRE' if vis_opcao == "Trimestre" else 'ANO')
    df_vendas = df.groupby(col_periodo).size().reset_index(name='Total Vendas')
    df_vendas = df_vendas[df_vendas[col_periodo] != 'N/A']
    
    media_vendas = df_vendas['Total Vendas'].mean()
    df_vendas['Desempenho'] = df_vendas['Total Vendas'].apply(lambda v: classificar_desempenho(v, media_vendas))
    
    fig_vendas = px.bar(
        df_vendas, x=col_periodo, y='Total Vendas', color='Desempenho',
        color_discrete_map=color_map, text='Total Vendas',
        title=f"Vendas por {vis_opcao} (Média no Período: {media_vendas:.1f})"
    )
    st.plotly_chart(fig_vendas, use_container_width=True)
    
    # REQUISITO 2: TOP CORRETORES
    st.subheader("2. Ranking de Corretores")
    c1, c2, c3 = st.columns(3)
    
    def get_top_corretores(grupo):
        top = df.groupby([grupo, col_corretor]).size().reset_index(name='Vendas')
        top = top.sort_values(by=[grupo, 'Vendas'], ascending=[True, False])
        return top.groupby(grupo).first().reset_index()

    with c1:
        st.markdown("**Top Corretor por Mês**")
        st.dataframe(get_top_corretores('MES_ANO'), use_container_width=True)
    with c2:
        st.markdown("**Top Corretor por Trimestre**")
        st.dataframe(get_top_corretores('TRIMESTRE'), use_container_width=True)
    with c3:
        st.markdown("**Top Corretor por Ano**")
        st.dataframe(get_top_corretores('ANO'), use_container_width=True)
        
    # REQUISITO 3: COBERTURAS POR PARTICIPANTE (Colunas J até V)
    st.subheader("3. Coberturas Contratadas por Participante")
    coberturas_resumo = []
    for _, row in df.iterrows():
        cob_ativas = []
        for c in cols_coberturas:
            val = row[c]
            if pd.notna(val) and val != 0 and str(val).strip() != "" and str(val).upper() != "NÃO":
                cob_ativas.append(c)
        coberturas_resumo.append({
            'CPF': row['CPF_MASC'],
            'Proponente': row[col_nome],
            'Qtd Coberturas': len(cob_ativas),
            'Coberturas': ", ".join(cob_ativas) if cob_ativas else "Nenhuma"
        })
    st.dataframe(pd.DataFrame(coberturas_resumo), use_container_width=True)
    
    # REQUISITO 5: DOWNLOAD DO EXCEL COMPLETO FORMATADO
    st.subheader("5. Exportar Relatório Formatado")
    excel_bytes = gerar_excel_formatado(df)
    st.download_button(
        label="📥 Baixar Relatório Completo (.xlsx)",
        data=excel_bytes,
        file_name="Relatorio_Detalhamento_Participantes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Aguardando o envio da planilha na barra lateral para carregar os indicadores.")
