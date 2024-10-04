import pandas as pd
import re

# Leitura da planilha
file_path = 'Log_ComexXCRM_SYS.xlsx'
df = pd.read_excel(file_path)

# Adicionando a coluna ID com valores sequenciais
df.insert(0, 'id', range(1, len(df) + 1))

# Função para formatar corretamente os valores das colunas player e opportunity_type
def format_array(value):
    if pd.isna(value):  # Verifica se o valor é NaN
        return 'NULL'
    if isinstance(value, str):
        # Remove as chaves externas e quebra o conteúdo por vírgulas
        items = re.split(r',\s*(?![^{}]*\})', value.strip('{}'))
        # Adiciona aspas duplas a cada item e junta novamente com vírgula
        formatted_items = ', '.join([f'"{item.strip()}"' for item in items])
        return f'{{{formatted_items}}}'
    return value

# Aplicando a formatação para as colunas 'player' e 'opportunity_type'
df['player'] = df['player'].apply(format_array)
df['opportunity_type'] = df['opportunity_type'].apply(format_array)

# Gerando os comandos INSERT INTO
insert_statements = []
for _, row in df.iterrows():
    insert_statements.append(
        f"INSERT INTO your_table_name "
        f"(id, mode, priority, origin_id, destination_id, player, volumes, kgs_teu, profile_id, shipment, project_maya, opportunity_type, awarded_route, focus_route, logcomex, created_at, created_by_id, updated_at, updated_by_id) "
        f"VALUES "
        f"({row['id']}, '{row['mode']}', {row['priority']}, {row['origin_id']}, {row['destination_id']}, {row['player']}, {row['volumes'] if pd.notna(row['volumes']) else 'NULL'}, {row['kgs_teu'] if pd.notna(row['kgs_teu']) else 'NULL'}, {row['profile_id']}, {row['shipment']}, '{row['project_maya']}', {row['opportunity_type']}, {row['awarded_route']}, {row['focus_route']}, {row['logcomex']}, '{row['created_at']}', {row['created_by_id']}, '{row['updated_at']}', {row['updated_by_id']});"
    )

# Salvando os comandos em um arquivo SQL
with open('output.sql', 'w') as f:
    for statement in insert_statements:
        f.write(f"{statement}\n")

print("Arquivo SQL gerado com sucesso!")
