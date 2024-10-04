import pandas as pd

# Carregar o arquivo Excel
file_path = 'tabelaBI.xlsx'  # Coloque o caminho correto do seu arquivo aqui
df = pd.read_excel(file_path)

# Função auxiliar para combinar os valores da coluna 'player' em uma lista única
def combine_players(players):
    return list(set(players))

# Agrupar os dados pelas colunas específicas e somar as colunas necessárias
grouped_df = df.groupby(
    ['mode', 'profile_id', 'origin_id', 'destination_id'], as_index=False
).agg({
    'total_volumes': 'sum',
    'total_weight_transported': 'sum',
    'shipment': 'sum',
    'player': combine_players,  # Combinar os jogadores em uma lista
    # Manter as outras colunas sem agregação
    'id': 'first',
    'deleted_at': 'first',
    'created_at': 'first',
    'updated_at': 'first',
    'kgs_teu': 'first',
    'volumes': 'first',
    'opportunity_type': 'first',
    'logcomex': 'first',
    'created_by_id': 'first',
    'updated_by_id': 'first',
    'awarded_route': 'first',
    'target': 'first',
    'project_maya': 'first',
    'desirable_percentage': 'first',
    'focus_route': 'first',
    'high_value_goods': 'first',
    'high_value_goods_answer': 'first',
    'incoterm': 'first',
    'operational_supplier_answer': 'first',
    'operational_supplier_id': 'first',
    'service': 'first',
    'shipper_problem': 'first',
    'shipper_problem_answer': 'first',
    'shipper_problem_justification': 'first',
})

# Reorganizar as colunas na ordem desejada
final_columns_order = [
    'id', 'deleted_at', 'created_at', 'updated_at', 'mode', 'player', 'kgs_teu',
    'volumes', 'opportunity_type', 'logcomex', 'created_by_id', 'origin_id',
    'destination_id', 'profile_id', 'updated_by_id', 'awarded_route', 'target',
    'project_maya', 'shipment', 'desirable_percentage', 'focus_route',
    'high_value_goods', 'high_value_goods_answer', 'incoterm',
    'operational_supplier_answer', 'operational_supplier_id', 'service',
    'shipper_problem', 'shipper_problem_answer', 'shipper_problem_justification',
    'total_volumes', 'total_weight_transported'
]

# Selecionar e ordenar as colunas
grouped_df = grouped_df[final_columns_order]

# Salvar o resultado em um novo arquivo Excel
output_path = 'tabelaBI_agrupada.xlsx'
grouped_df.to_excel(output_path, index=False)

print(f"Dados agrupados salvos em {output_path}")
