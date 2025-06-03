import os
import pandas as pd
import sqlite3

def import_excel_to_sqlite(excel_file_path, sheet_name, db_file, table_name):
    """
    Importa dados de um arquivo Excel para uma tabela SQLite.
    """
    
    # Garantir que o diretório do banco de dados exista
    db_directory = os.path.dirname(db_file)
    if db_directory and not os.path.exists(db_directory):
        os.makedirs(db_directory)

    # Verificar se o arquivo existe
    if not os.path.exists(excel_file_path):
        print(f"Erro: O arquivo {excel_file_path} não foi encontrado.")
        return

    # Ler o arquivo Excel em um DataFrame do Pandas sem cabeçalho para encontrar a linha correta
    df = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=None, engine='openpyxl')

    # Encontrar a linha onde a primeira coluna contém 'Id Paciente'
    header_row = None
    for index, row in df.iterrows():
        if '#Id Paciente' in row.values:
            header_row = index
            break

    if header_row is None:
        print(f"Erro: Cabeçalho 'Id Paciente' não encontrado no arquivo: {excel_file_path}")
        return

    # Recarregar o DataFrame usando a linha correta como cabeçalho
    df = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=header_row, engine='openpyxl')

    # Conectar ao banco de dados SQLite (ou criar se não existir)
    conn = sqlite3.connect(db_file)

    # Verificar se a tabela já existe
    query = f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}';"
    table_exists = pd.read_sql(query, conn).shape[0] > 0

    if table_exists:
        # Se a tabela existir, importe apenas as atualizações
        existing_data = pd.read_sql(f"SELECT * FROM {table_name}", conn)
        
        # Encontre as novas linhas que não estão no banco de dados
        new_data = df[~df['#Id Paciente'].isin(existing_data['#Id Paciente'])]
        
        # Insira as novas linhas
        if not new_data.empty:
            new_data.to_sql(table_name, conn, if_exists='append', index=False)
            print(f"Novos dados adicionados à tabela {table_name}.")
        else:
            print("Nenhum dado novo para adicionar.")
    else:
        # Se a tabela não existir, crie-a e insira todos os dados
        df.to_sql(table_name, conn, if_exists='replace', index=False)
        print(f"Dados importados com sucesso para a tabela {table_name}.")

    # Fechar a conexão
    conn.close()

    print(f"Dados importados com sucesso para a tabela {table_name} no banco de dados {db_file}!")

if __name__ == "__main__":
    # Exemplo de uso
    import_excel_to_sqlite(
        excel_file_path='caminho/para/seu/arquivo.xlsx',
        sheet_name='NomeDaAba',
        db_file='caminho/para/seu/banco_de_dados.db',
        table_name='sua_tabela'
    )