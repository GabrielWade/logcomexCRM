import pandas as pd

# Carregar os arquivos
crm_bidroute_path = 'CRM_BIDROUTE.xlsx'
tabela_bi_agrupada_path = 'tabelaBI_agrupada.xlsx'

# Carregar dados com pandas
crm_bidroute = pd.read_excel(crm_bidroute_path)
tabela_bi_agrupada = pd.read_excel(tabela_bi_agrupada_path)

# Colunas de comparação
comparison_columns = ['mode', 'profile_id', 'origin_id', 'destination_id']

# Colunas para extrair de CRM_BIDROUTE
crm_columns_to_take = [
    'id', 'deleted_at', 'created_at', 'updated_at', 'kgs_teu', 'volumes', 'opportunity_type', 'logcomex',
    'created_by_id', 'updated_by_id', 'awarded_route', 'target', 'project_maya',
    'desirable_percentage', 'focus_route', 'high_value_goods', 'high_value_goods_answer',
    'incoterm', 'operational_supplier_answer', 'operational_supplier_id', 'service',
    'shipper_problem', 'shipper_problem_answer', 'shipper_problem_justification'
]

# Colunas para extrair de tabelaBI_agrupada
bi_columns_to_take = ['player', 'shipment', 'total_volumes', 'total_weight_transported']

# Realizar a mesclagem com base nas colunas de comparação
merged_data = pd.merge(crm_bidroute, tabela_bi_agrupada,
                       on=comparison_columns,
                       how='outer', indicator=True)

# Ajustar nomes das colunas adicionando sufixos após a mesclagem
crm_columns_to_take_adjusted = [col + '_x' for col in crm_columns_to_take if col + '_x' in merged_data.columns]
bi_columns_to_take_adjusted = [col + '_y' for col in bi_columns_to_take if col + '_y' in merged_data.columns]

# Incluir as colunas de comparação na tabela final
comparison_columns_adjusted = [col for col in comparison_columns]

# Criar tabela final combinando os dados
new_table = pd.DataFrame()

# Para linhas que existem em ambas as tabelas
matched_rows = merged_data[merged_data['_merge'] == 'both']
new_table = pd.concat([new_table, matched_rows[comparison_columns_adjusted + crm_columns_to_take_adjusted + bi_columns_to_take_adjusted]])

# Para linhas que existem apenas na tabelaBI_agrupada
unmatched_rows = merged_data[merged_data['_merge'] == 'right_only']
new_table = pd.concat([new_table, unmatched_rows[comparison_columns_adjusted + bi_columns_to_take_adjusted]])

# Substituir valores NaN por 'NULL'
new_table.fillna('NULL', inplace=True)

# Salvar o resultado em um novo arquivo Excel
new_table.to_excel('merged_table_output.xlsx', index=False)

print("Tabela final gerada com sucesso!")
