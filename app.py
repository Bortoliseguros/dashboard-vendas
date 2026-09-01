import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import re
import io
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Dashboard e Gestão de Riscos", layout="wide")

# Adicione esta linha para exibir a logo na barra lateral:
st.sidebar.image("logo.png", width=150)
# Mascaramento de CPF
def mascarar_cpf(val):
    if pd.isna(val):
        return "N/A"
    digits = re.sub(r'\D', '', str(val)).zfill(11)
    if len(digits) >= 11:
        return f"***.***.{digits[-5:-2]}-{digits[-2:]}"
    return "***.***.***-**"

# Classificação de Desempenho
def classificar_desempenho(val, media, tol=0.05):
    if val > media * (1 + tol):
        return 'Acima da Média'
    elif val < media * (1 - tol):
        return 'Abaixo da Média'
    else:
        return 'Na Média'

# Regra oficial para a Coluna W (Status no 104)
def obter_status_104(val):
    if pd.isna(val) or str(val).strip() == '' or str(val).strip().upper() == 'NAN':
        return 'A IMPLANTAR'
    v = str(val).strip().upper()
    if 'IMPLANTAD' in v:
        return 'IMPLANTADA'
    return 'A IMPLANTAR'

# Gerador de Relatório Excel Ajustado
def gerar_excel_formatado(df_original):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DETALHAMENTO POR PARTICIPANTE"
    
    df_exp = df_original.copy()
    
    col_d_name = df_exp.columns[3] if len(df_exp.columns) > 3 else None
    if col_d_name:
        df_exp[col_d_name] = df_exp[col_d_name].fillna('A IMPLANTAR')
        df_exp[col_d_name] = df_exp[col_d_name].apply(lambda x: 'A IMPLANTAR' if pd.isna(x) or str(x).strip() == '' else str(x))

    col_w_name = df_exp.columns[22] if len(df_exp.columns) > 22 else None
    if col_w_name:
        df_exp[col_w_name] = df_exp[col_w_name].apply(obter_status_104)
        
    col_cpf_name = next((c for c in df_exp.columns if 'CPF' in c.upper()), None)
    if col_cpf_name:
        df_exp[col_cpf_name] = df_exp[col_cpf_name].apply(mascarar_cpf)
        
    col_i = df_exp.columns[8] if len(df_exp.columns) > 8 else None
    col_x = df_exp.columns[23] if len(df_exp.columns) > 23 else None
    
    if col_i:
        df_exp[col_i] = pd.to_datetime(df_exp[col_i], errors='coerce').dt.strftime('%d/%m/%Y').fillna('')
    if col_x:
        df_exp[col_x] = pd.to_datetime(df_exp[col_x], errors='coerce').dt.strftime('%d/%m/%Y').fillna('')

    cols_drop = ['DATA_REF', 'ANO', 'MES_ANO', 'TRIMESTRE', 'SEMESTRE', 'STATUS_DASHBOARD']
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
        
    for row_idx in range(2, ws.max_row + 1):
        cell_a = ws.cell(row=row_idx, column=1)
        if cell_a.value is not None and str(cell_a.value).strip() != '':
            try:
                cell_a.value = int(float(cell_a.value))
                cell_a.number_format = '0'
            except: pass

        cell_c = ws.cell(row=row_idx, column=3)
        if cell_c.value is not None and str(cell_c.value).strip() != '':
            try:
                cell_c.value = int(float(cell_c.value))
                cell_c.number_format = '0'
            except: pass

        for col_idx in range(1, ws.max_column + 1):
            if col_idx in [1, 3]: 
                continue
            cell = ws.cell(row=row_idx, column=col_idx)
            if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                cell.number_format = '#,##0.00'
                
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)
        
    wb.save(output)
    return output.getvalue()

st.title("📊 Painel de Vendas e Gestão de Riscos")

# CARREGAMENTO AUTOMÁTICO DO ARQUIVO EXATO NO GITHUB
ARQUIVO_PADRAO = "RELACAO DE MALOTE RECEBIDO SEGURADORA.xlsx"

df = None
if os.path.exists(ARQUIVO_PADRAO):
    try:
        df = pd.read_excel(ARQUIVO_PADRAO)
        st.sidebar.success("✅ Dados carregados automaticamente!")
    except Exception as e:
        st.sidebar.error(f"Erro ao ler o arquivo: {e}")
else:
    # Fallback caso o arquivo ainda não tenha sido enviado para o repositório
    st.sidebar.warning(f"Arquivo '{ARQUIVO_PADRAO}' não encontrado no GitHub. Envie abaixo:")
    file_up = st.sidebar.file_uploader("Enviar planilha (.xlsx)", type=["xlsx"])
    if file_up is not None:
        df = pd.read_excel(file_up)

if df is not None:
    df.columns = [str(c).strip() for c in df.columns]
    
    col_nome = next((c for c in df.columns if 'NOME' in c.upper() or 'PROPONENTE' in c.upper()), df.columns[5])
    col_cpf = next((c for c in df.columns if 'CPF' in c.upper()), df.columns[6])
    col_corretor = df.columns[7] # Coluna H (AGENTE MAG)
    col_data = df.columns[8]    # Coluna I (A PARTIR DE)
    col_w_status = df.columns[22] if len(df.columns) > 22 else None # Coluna W
    cols_coberturas = df.columns[9:22] # Colunas J até V
    
    if col_w_status:
        df['STATUS_DASHBOARD'] = df[col_w_status].apply(obter_status_104)
    else:
        df['STATUS_DASHBOARD'] = 'A IMPLANTAR'
    
    # Tratamento de datas e períodos
    df['DATA_REF'] = pd.to_datetime(df[col_data], errors='coerce')
    df['ANO'] = df['DATA_REF'].dt.year.astype(str)
    df['MES_ANO'] = df['DATA_REF'].dt.strftime('%m/%Y')
    df['TRIMESTRE'] = df['DATA_REF'].apply(lambda d: f"{d.year}-Q{(d.month-1)//3 + 1}" if pd.notna(d) else 'N/A')
    df['SEMESTRE'] = df['DATA_REF'].apply(lambda d: f"{d.year} - {1 if d.month <= 6 else 2}º Semestre" if pd.notna(d) else 'N/A')
    
    # Estrutura das Guias
    tab1, tab2, tab3, tab4 = st.tabs([
        "📈 Vendas no Mês, Trimestre e Ano",
        "🤝 Vendas por Corretores (Mensal, Trimestral e Anual)",
        "🏆 Ranking de Produtores (Coluna H)",
        "📋 Relatorio dos riscos e peculios contratados"
    ])
    
    # -------------------------------------------------------------------------
    # GUIA 1: VENDAS GERAIS
    # -------------------------------------------------------------------------
    with tab1:
        st.subheader("Vendas Gerais por Mês, Trimestre e Ano")
        vis_opcao = st.radio("Selecione a escala temporal:", ["Mês", "Trimestre", "Ano"], key="vis_g1", horizontal=True)
        col_p = 'MES_ANO' if vis_opcao == "Mês" else ('TRIMESTRE' if vis_opcao == "Trimestre" else 'ANO')
        
        df_vendas = df.groupby(col_p).size().reset_index(name='Total Vendas')
        df_vendas = df_vendas[df_vendas[col_p] != 'N/A']
        
        media_vendas = df_vendas['Total Vendas'].mean() if not df_vendas.empty else 0
        df_vendas['Desempenho'] = df_vendas['Total Vendas'].apply(lambda v: classificar_desempenho(v, media_vendas))
        
        color_map = {'Acima da Média': '#10B981', 'Na Média': '#2563EB', 'Abaixo da Média': '#EF4444'}
        
        fig = px.bar(
            df_vendas, x=col_p, y='Total Vendas', color='Desempenho',
            color_discrete_map=color_map, text='Total Vendas',
            title=f"Volume Geral de Vendas por {vis_opcao} (Média do Período: {media_vendas:.1f})"
        )
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(df_vendas, use_container_width=True)

    # -------------------------------------------------------------------------
    # GUIA 2: VENDAS POR CORRETORES
    # -------------------------------------------------------------------------
    with tab2:
        st.subheader("Vendas por Corretores")
        
        c1, c2 = st.columns([1, 2])
        with c1:
            per_corretor = st.radio("Visão Temporal:", ["Mensal", "Trimestral", "Anual"], key="vis_corr", horizontal=True)
        with c2:
            lista_corretores = ["TODOS"] + sorted(df[col_corretor].dropna().unique().tolist())
            corretor_sel = st.selectbox("Filtrar por Corretor Especifico:", options=lista_corretores)

        col_c = 'MES_ANO' if per_corretor == "Mensal" else ('TRIMESTRE' if per_corretor == "Trimestral" else 'ANO')
        
        df_corr = df.copy()
        if corretor_sel != "TODOS":
            df_corr = df_corr[df_corr[col_corretor] == corretor_sel]
            
        df_grp = df_corr.groupby([col_c, col_corretor]).size().reset_index(name='Vendas')
        df_grp = df_grp[df_grp[col_c] != 'N/A']
        
        df_grp['TEMP_DATE'] = pd.to_datetime(df_grp[col_c], format='%m/%Y', errors='coerce')
        df_grp = df_grp.sort_values(by=['TEMP_DATE', 'Vendas'], ascending=[False, False]).drop(columns=['TEMP_DATE']).reset_index(drop=True)
        
        fig_corr = px.bar(
            df_grp, x=col_c, y='Vendas', color=col_corretor, barmode='group',
            text='Vendas', title=f"Vendas por Corretor ({per_corretor})"
        )
        st.plotly_chart(fig_corr, use_container_width=True)
        
        st.markdown("**Tabela Detalhada de Vendas por Corretor (Ordenado do Maior para o Menor por Período)**")
        st.dataframe(df_grp, use_container_width=True)

    # -------------------------------------------------------------------------
    # GUIA 3: RANKING DE PRODUTORES (Com Seletor de Período)
    # -------------------------------------------------------------------------
    with tab3:
        st.subheader("🏆 Ranking de Produtores (Coluna H)")
        
        vis_ranking = st.radio(
            "Selecione a base temporal para o Ranking:",
            ["Mensal", "Trimestral", "Semestral", "Anual"],
            key="vis_ranking_sel",
            horizontal=True
        )
        
        if vis_ranking == "Mensal":
            col_rank = 'MES_ANO'
            titulo_rank = "Ranking de Vendas por Mês (Do Maior para o Menor)"
        elif vis_ranking == "Trimestral":
            col_rank = 'TRIMESTRE'
            titulo_rank = "Ranking de Vendas por Trimestre (Do Maior para o Menor)"
        elif vis_ranking == "Semestral":
            col_rank = 'SEMESTRE'
            titulo_rank = "Ranking de Vendas por Semestre (Do Maior para o Menor)"
        else:
            col_rank = 'ANO'
            titulo_rank = "Ranking de Vendas por Ano (Do Maior para o Menor)"
            
        st.markdown(f"#### {titulo_rank}")
        
        df_ranking_res = df.groupby([col_rank, col_corretor]).size().reset_index(name='Qtd Vendas')
        df_ranking_res = df_ranking_res[df_ranking_res[col_rank] != 'N/A']
        
        df_ranking_res['TEMP_DATE'] = pd.to_datetime(df_ranking_res[col_rank], format='%m/%Y', errors='coerce')
        df_ranking_res = df_ranking_res.sort_values(by=['TEMP_DATE', 'Qtd Vendas'], ascending=[False, False]).drop(columns=['TEMP_DATE']).reset_index(drop=True)
        
        st.dataframe(df_ranking_res, use_container_width=True)

    # -------------------------------------------------------------------------
    # GUIA 4: RELATÓRIO DOS RISCOS E PECÚLIOS CONTRATADOS
    # -------------------------------------------------------------------------
    with tab4:
        st.subheader("Relatorio dos riscos e peculios contratados")
        
        lista_participantes = sorted(df[col_nome].dropna().unique().tolist())
        
        if lista_participantes:
            participante_sel = st.selectbox(
                "Busque ou selecione o participante:",
                options=lista_participantes,
                index=0
            )
            
            df_part = df[df[col_nome] == participante_sel].iloc[0]
            status_final = df_part['STATUS_DASHBOARD']
            
            if status_final == 'IMPLANTADA':
                status_html = "<span style='background-color: #D1FAE5; color: #059669; padding: 4px 10px; border-radius: 6px; font-weight: bold;'>IMPLANTADA</span>"
            else:
                status_html = "<span style='background-color: #FEE2E2; color: #DC2626; padding: 4px 10px; border-radius: 6px; font-weight: bold;'>A IMPLANTAR</span>"
            
            st.markdown("---")
            i1, i2, i3 = st.columns(3)
            with i1:
                st.markdown(f"**Proponente:** {df_part[col_nome]}")
                st.markdown(f"**CPF:** `{mascarar_cpf(df_part[col_cpf])}`")
            with i2:
                st.markdown(f"**Proposta:** {df_part.get('PROPOSTA', 'N/A')}")
                st.markdown(f"**Agente/Corretor:** {df_part[col_corretor]}")
            with i3:
                st.markdown(f"**Status:** {status_html}", unsafe_allow_html=True)
                st.markdown(f"**Data da Proposta:** {df_part['MES_ANO']}")
                
            st.markdown("#### Detalhamento de Pecúlios e Riscos (Colunas J a V)")
            
            detalhes_cob = []
            for col_c in cols_coberturas:
                val = df_part[col_c]
                val_num = pd.to_numeric(val, errors='coerce')
                
                if pd.notna(val_num) and val_num != 0:
                    val_fmt = f"R$ {val_num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    detalhes_cob.append({"Item / Cobertura": col_c, "Valor Registrado": val_fmt})
                elif pd.notna(val) and str(val).strip() != "" and str(val).upper() != "NÃO":
                    detalhes_cob.append({"Item / Cobertura": col_c, "Valor Registrado": str(val)})
                    
            if detalhes_cob:
                st.table(pd.DataFrame(detalhes_cob))
            else:
                st.info("Nenhuma cobertura ativa para o participante selecionado.")
                
            st.markdown("---")
            excel_bytes = gerar_excel_formatado(df)
            st.download_button(
                label="📥 Exportar Relatório Completo Ajustado em Excel",
                data=excel_bytes,
                file_name="Relatorio_Riscos_Peculios_Completo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nenhum participante encontrado na base de dados.")
else:
    st.error("O arquivo 'RELACAO DE MALOTE RECEBIDO SEGURADORA.xlsx' não foi encontrado no repositório do GitHub. Certifique-se de enviá-lo para a mesma pasta do app.py.")
