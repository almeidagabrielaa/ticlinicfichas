import pandas as pd
from pandas_gbq import to_gbq

def import_excel_to_bq(excel_file_path, sheet_name, project_id, dataset_id, table_id):
    """
    Importa dados de um arquivo Excel para uma tabela do Google BigQuery.

    Parameters:
    - excel_file_path: Caminho para o arquivo Excel.
    - sheet_name: Nome da aba do Excel a ser lida.
    - project_id: ID do projeto do Google Cloud.
    - dataset_id: ID do dataset no BigQuery.
    - table_id: Nome da tabela no BigQuery.
    """
    
    # Ler o arquivo Excel em um DataFrame do Pandas
    df = pd.read_excel(excel_file_path, sheet_name=sheet_name, engine='openpyxl')

    # Visualizar os dados (opcional)
    print("Primeiros registros do DataFrame:")
    print(df.head())

    # Nome completo da tabela no formato dataset.tabela
    table_full_id = f'{dataset_id}.{table_id}'

    # Importar o DataFrame para o BigQuery
    to_gbq(df, table_full_id, project_id=project_id, if_exists='replace')

    print("Dados importados com sucesso para o BigQuery!")

if __name__ == "__main__":
    # Exemplo de uso
    import_excel_to_bq(
        excel_file_path='caminho/para/seu/arquivo.xlsx',
        sheet_name='NomeDaAba',
        project_id='seu-project-id',
        dataset_id='seu_dataset',
        table_id='sua_tabela'
    )