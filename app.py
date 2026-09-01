import streamlit as st
import pandas as pd
import plotly.express as px
import re
from datetime import datetime

# Configuração da Página
st.set_page_config(
    page_title="Dashboard Comercial & Coberturas",
    page_icon="📊",
    layout="wide"
)

# Estilização CSS personalizada
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        color: #1E3A8A;
        font-weight: 700;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        color: #4B5563;
        font-size: 1.1rem;
        margin-bottom: 2rem;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">📊 Dashboard Comercial e Relatório de Coberturas</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Faça o upload do seu arquivo Excel para processar os dados com conformidade de privacidade (LGPD).</div>', unsafe_allow_html=True)

# Sidebar - Upload
st.sidebar.header("📁 Importar Dados")
uploaded_file = st.sidebar.file_uploader("Selecione a planilha Excel (.xlsx)", type=["xlsx", "xls"])

def mascarar_cpf(cpf):
    """Mascara o CPF exibindo asteriscos no Streamlit sem conflito de Markdown"""
    if pd.isna(cpf):
        return "N/A"
    digits = re.sub(r'\D', '', str(cpf))
    if len(digits) == 11:
        return f"\*\*\*.\*\*\*.{digits[6:9]}-{digits[9:]}"
    elif len(digits) > 5:
        return f"\*\*\*.\*\*\*.{digits[-5:-2]}-{digits[-2:]}"
    return "\*\*\*.\*\*\*.\*\*\*-\*\*"

def parse_data_partir(val):
    if pd.isna(val):
        return None
    
    if isinstance(val, (pd.Timestamp, datetime)):
        return val

    val_str = str(val).strip().lower()
    meses_pt = {
        'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4, 'mai': 5, 'jun': 6,
        'jul': 7, 'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12
    }
    
    match = re.match(r'([a-z]{3})[/-](\d{2,4})', val_str)
    if match:
        mes_str, ano_str = match.groups()
        if mes_str in meses_pt:
            mes = meses_pt[mes_str]
            ano = int('20' + ano_str if len(ano_str) == 2 else ano_str)
            return datetime(ano, mes, 1)
            
    try:
        dt = pd.to_datetime(val_str, dayfirst=True, errors='coerce')
        if pd.notna(dt):
            return dt.to_pydatetime()
    except:
        pass
    return None

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        
        # Normalizar nomes de colunas
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        # Mapeamento de colunas
        col_id = next((c for c in df.columns if 'ID' in c), None)
        col_nome = next((c for c in df.columns if 'NOME' in c), None)
        col_cpf = next((c for c in df.columns if 'CPF' in c), None)
        col_proposta = next((c for c in df.columns if 'PROPOSTA' in c), None)
        col_corretor = next((c for c in df.columns if 'CORRETOR' in c), None)
        col_a_partir = next((c for c in df.columns if 'PARTIR' in c or 'DATA' in c), None)
        
        # Aplicação de Mascaramento do CPF (LGPD)
        if col_cpf:
            df['CPF_MASCARADO'] = df[col_cpf].apply(mascarar_cpf)
        else:
            df['CPF_MASCARADO'] = "N/A"

        # Processar Data
        if col_a_partir:
            df['DATA_REF'] = df[col_a_partir].apply(parse_data_partir)
            df['ANO'] = df['DATA_REF'].apply(lambda d: str(d.year) if pd.notna(d) and d else 'N/A')
            df['MES_ANO'] = df['DATA_REF'].apply(lambda d: d.strftime('%m/%Y') if pd.notna(d) and d else 'N/A')
            df['TRIMESTRE'] = df['DATA_REF'].apply(lambda d: f"{d.year}-Q{(d.month-1)//3 + 1}" if pd.notna(d) and d else 'N/A')
        else:
            df['ANO'], df['MES_ANO'], df['TRIMESTRE'] = 'N/A', 'N/A', 'N/A'
            
        # Converter colunas numéricas de coberturas
        cols_coberturas = [
            'RISCO MORTE', 'PECULIO MORTE', 'RISCO INVALIDEZ', 'PECULIO INVALIDEZ',
            'ADICIONAL MORTE', 'PECULIO ADICIONAL MORTE', 'ADICIONAL INVALIDEZ',
            'PECULIO ADICIONAL INVALIDEZ', 'TOTAL'
        ]
        
        for c in cols_coberturas:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce').fillna(0)
            else:
                df[c] = 0.0

        # Abas da Aplicação
        tab1, tab2, tab3 = st.tabs([
            "📈 1. Vendas por Período",
            "🏆 2. Ranking de Corretores",
            "👤 3. Detalhamento por Participante"
        ])
        
        # ----------------------------------------------------
        # ABA 1: VENDAS POR MÊS, TRIMESTRE E ANO
        # ----------------------------------------------------
        with tab1:
            st.subheader("Análise Temporal de Vendas")
            
            col_kpi1, col_kpi2, col_kpi3 = st.columns(3)
            total_vendas = len(df)
            valor_total = df['TOTAL'].sum()
            ticket_medio = df['TOTAL'].mean() if total_vendas > 0 else 0
            
            col_kpi1.metric("Total de Propostas", f"{total_vendas:,}")
            col_kpi2.metric("Volume Total (R$)", f"R$ {valor_total:,.2f}")
            col_kpi3.metric("Ticket Médio (R$)", f"R$ {ticket_medio:,.2f}")
            
            st.divider()
            
            visao = st.radio("Selecione a visão temporal:", ["Mês", "Trimestre", "Ano"], horizontal=True)
            col_agrupador = 'MES_ANO' if visao == "Mês" else ('TRIMESTRE' if visao == "Trimestre" else 'ANO')
            
            df_vendas = df.groupby(col_agrupador).agg(
                Qtd_Vendas=(col_proposta if col_proposta else df.columns[0], 'count'),
                Valor_Total=('TOTAL', 'sum')
            ).reset_index()
            
            fig_vendas = px.bar(
                df_vendas, 
                x=col_agrupador, 
                y='Qtd_Vendas',
                text='Qtd_Vendas',
                title=f'Quantidade de Vendas por {visao}',
                labels={col_agrupador: visao, 'Qtd_Vendas': 'Nº de Vendas'},
                color_discrete_sequence=['#2563EB']
            )
            fig_vendas.update_traces(textposition='outside')
            st.plotly_chart(fig_vendas, use_container_width=True)
            
            fig_valor = px.line(
                df_vendas, 
                x=col_agrupador, 
                y='Valor_Total',
                markers=True,
                title=f'Faturamento Total por {visao} (R$)',
                labels={col_agrupador: visao, 'Valor_Total': 'Valor Total (R$)'},
                color_discrete_sequence=['#10B981']
            )
            st.plotly_chart(fig_valor, use_container_width=True)
            
            st.dataframe(df_vendas, use_container_width=True)

        # ----------------------------------------------------
        # ABA 2: RANKING DE CORRETORES
        # ----------------------------------------------------
        with tab2:
            st.subheader("Desempenho dos Corretores")
            
            c_f1, c_f2 = st.columns(2)
            with c_f1:
                tipo_periodo = st.selectbox("Filtrar por Granularidade:", ["Mensal", "Trimestral", "Anual"])
            
            col_p = 'MES_ANO' if tipo_periodo == "Mensal" else ('TRIMESTRE' if tipo_periodo == "Trimestral" else 'ANO')
            
            with c_f2:
                opcoes_p = sorted(df[col_p].unique().tolist())
                periodo_sel = st.selectbox("Selecione o Período:", opcoes_p)
                
            df_corretor_filt = df[df[col_p] == periodo_sel]
            
            if col_corretor:
                ranking = df_corretor_filt.groupby(col_corretor).agg(
                    Qtd_Vendas=(col_proposta if col_proposta else df.columns[0], 'count'),
                    Valor_Total=('TOTAL', 'sum')
                ).reset_index().sort_values(by='Valor_Total', ascending=False)
                
                if not ranking.empty:
                    campeao = ranking.iloc[0]
                    st.success(f"🏆 **Corretor Campeão em {periodo_sel}:** {campeao[col_corretor]} com R$ {campeao['Valor_Total']:,.2f} em vendas ({campeao['Qtd_Vendas']} propostas).")
                    
                fig_rank = px.bar(
                    ranking,
                    x=col_corretor,
                    y='Valor_Total',
                    text='Valor_Total',
                    color='Valor_Total',
                    title=f'Ranking de Vendas por Corretor ({periodo_sel})',
                    labels={'Valor_Total': 'Valor Total (R$)', col_corretor: 'Corretor'},
                    color_continuous_scale='Blues'
                )
                fig_rank.update_traces(texttemplate='R$ %{text:,.2f}', textposition='outside')
                st.plotly_chart(fig_rank, use_container_width=True)
                
                st.dataframe(ranking, use_container_width=True)

        # ----------------------------------------------------
        # ABA 3: CONSULTA DE PARTICIPANTE E COBERTURAS
        # ----------------------------------------------------
        with tab3:
            st.subheader("Consulta de Coberturas Contratadas")
            
            busca = st.text_input("🔍 Digite o Nome ou CPF do Participante para buscar:")
            
            df_part = df.copy()
            if busca and col_nome:
                cond_nome = df_part[col_nome].astype(str).str.contains(busca, case=False, na=False)
                cond_cpf = df_part[col_cpf].astype(str).str.contains(busca, case=False, na=False) if col_cpf else False
                df_part = df_part[cond_nome | cond_cpf]
                
            if not df_part.empty and col_nome:
                nome_sel = st.selectbox("Selecione o Participante:", df_part[col_nome].unique())
                row = df_part[df_part[col_nome] == nome_sel].iloc[0]
                
                c_p1, c_p2, c_p3 = st.columns(3)
                with c_p1:
                    st.markdown(f"**Nome:** {row[col_nome] if col_nome else 'N/A'}")
                    st.markdown(f"**CPF (Protegido):** {row['CPF_MASCARADO']}")
                with c_p2:
                    st.markdown(f"**Proposta:** {row[col_proposta] if col_proposta else 'N/A'}")
                    st.markdown(f"**Corretor:** {row[col_corretor] if col_corretor else 'N/A'}")
                with c_p3:
                    st.markdown(f"**Início:** {row[col_a_partir] if col_a_partir else 'N/A'}")
                    st.markdown(f"**Valor Total:** R$ {row['TOTAL']:,.2f}")
                    
                st.write("### 🛡️ Coberturas (Colunas G a O)")
                
                coberturas_data = []
                for cob in cols_coberturas[:-1]:
                    val = row.get(cob, 0.0)
                    coberturas_data.append({
                        "Cobertura / Risco": cob,
                        "Valor Contratado": f"R$ {val:,.2f}"
                    })
                    
                df_cob_view = pd.DataFrame(coberturas_data)
                st.table(df_cob_view)

    except Exception as e:
        st.error(f"Ocorreu um erro ao ler a planilha: {e}")
else:
    st.info("👈 Por favor, faça o upload da sua planilha Excel na barra lateral para começar.")
