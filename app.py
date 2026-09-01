import pandas as pd
import numpy as np

# 1. CARREGAMENTO DOS DADOS
caminho_arquivo = "1 RELACAO DE MALOTE RECEBIDO SEGURADORA_2.xlsx"

# Lê a primeira aba (ou ajuste 'sheet_name' para o nome exato se necessário)
df = pd.read_excel(caminho_arquivo, sheet_name=0)

# Mapeamento dinâmico das colunas por índice base zero:
# Col F (5), Col G (6), Col H (7), Col I (8), Col J a V (9 a 21)
col_cpf = df.columns[5]          # Coluna F
col_nome = df.columns[6]         # Coluna G
col_corretor = df.columns[7]     # Coluna H
col_data = df.columns[8]         # Coluna I
cols_coberturas = df.columns[9:22] # Colunas J a V

# Tratamento da Data (Coluna I)
df['Data_Tratada'] = pd.to_datetime(df[col_data], errors='coerce')
df['Ano'] = df['Data_Tratada'].dt.year
df['Trimestre'] = df['Data_Tratada'].dt.to_period('Q').astype(str)
df['Mes_Ano'] = df['Data_Tratada'].dt.to_period('M').astype(str)

# -------------------------------------------------------------------------
# REQUISITO 4: MASCARAMENTO DE CPF (Formato: ***.***.117-99)
# -------------------------------------------------------------------------
def mascarar_cpf(val):
    if pd.isna(val):
        return ""
    digits = ''.join(filter(str.isdigit, str(val))).zfill(11)
    if len(digits) >= 11:
        return f"***.***.{digits[-5:-2]}-{digits[-2:]}"
    return str(val)

df['CPF_Mascarado'] = df[col_cpf].apply(mascarar_cpf)

# -------------------------------------------------------------------------
# REQUISITO 1: VENDAS POR MÊS, TRIMESTRE E ANO
# -------------------------------------------------------------------------
vendas_ano = df.groupby('Ano').size().reset_index(name='Total Vendas')
vendas_trimestre = df.groupby('Trimestre').size().reset_index(name='Total Vendas')
vendas_mes = df.groupby('Mes_Ano').size().reset_index(name='Total Vendas')

# -------------------------------------------------------------------------
# REQUISITO 2: MAIOR CORRETOR (MÊS, TRIMESTRE, ANO)
# -------------------------------------------------------------------------
def get_top_corretor(grupo):
    top = df.groupby([grupo, col_corretor]).size().reset_index(name='Vendas')
    top = top.sort_values(by=[grupo, 'Vendas'], ascending=[True, False])
    return top.groupby(grupo).first().reset_index()

top_corretor_ano = get_top_corretor('Ano')
top_corretor_trimestre = get_top_corretor('Trimestre')
top_corretor_mes = get_top_corretor('Mes_Ano')

# -------------------------------------------------------------------------
# REQUISITO 3: COBERTURAS CONTRATADAS POR PARTICIPANTE (Colunas J a V)
# -------------------------------------------------------------------------
# Mapeia quais coberturas possuem valor ativo/preenchido para cada participante
coberturas_list = []
for idx, row in df.iterrows():
    cob_ativas = []
    for col in cols_coberturas:
        val = row[col]
        if pd.notna(val) and val != 0 and str(val).strip() != "" and str(val).upper() != "NÂO":
            cob_ativas.append(str(col))
    
    coberturas_list.append({
        'CPF': row['CPF_Mascarado'],
        'Nome': row[col_nome],
        'Qtd Coberturas': len(cob_ativas),
        'Coberturas Contratadas': ", ".join(cob_ativas) if cob_ativas else "Nenhuma"
    })

df_coberturas_resumo = pd.DataFrame(coberturas_list)

# -------------------------------------------------------------------------
# REQUISITO 5 & EXPORTAÇÃO COMPLETA EM EXCEL FORMATADO
# -------------------------------------------------------------------------
arquivo_saida = "Relatorio_e_Dashboard_Seguradora.xlsx"

with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
    # 1. Detalhamento Formatado
    df_export = df.copy()
    df_export[col_cpf] = df_export['CPF_Mascarado']
    df_export.drop(columns=['Data_Tratada', 'Ano', 'Trimestre', 'Mes_Ano', 'CPF_Mascarado'], inplace=True)
    df_export.to_excel(writer, sheet_name='DETALHAMENTO POR PARTICIPANTE', index=False)
    
    # 2. Resumo de Coberturas
    df_coberturas_resumo.to_excel(writer, sheet_name='COBERTURAS POR PARTICIPANTE', index=False)
    
    # 3. Resumo por Período
    vendas_mes.to_excel(writer, sheet_name='VENDAS POR MES', index=False)
    vendas_trimestre.to_excel(writer, sheet_name='VENDAS POR TRIMESTRE', index=False)
    vendas_ano.to_excel(writer, sheet_name='VENDAS POR ANO', index=False)
    
    # 4. Ranking de Corretores
    top_corretor_mes.to_excel(writer, sheet_name='TOP CORRETOR MES', index=False)
    top_corretor_trimestre.to_excel(writer, sheet_name='TOP CORRETOR TRIMESTRE', index=False)
    top_corretor_ano.to_excel(writer, sheet_name='TOP CORRETOR ANO', index=False)

# Formatação visual (casas decimais e larguras de coluna) via openpyxl
import openpyxl

wb = openpyxl.load_workbook(arquivo_saida)

for sheetname in wb.sheetnames:
    ws = wb[sheetname]
    
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            # Aplica formatação de duas casas decimais em números flutuantes
            if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                cell.number_format = '#,##0.00'
                
    # Auto-ajuste de largura das colunas
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 50)

wb.save(arquivo_saida)
print(f"Relatório e indicadores gerados com sucesso no arquivo: {arquivo_saida}")
