import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime
from query import query_mdcs, query_prd, query_ve_versao
from variaveis_conexao import engine_prd, engine_mdcs_sync
from link import link

def gera_relatorio():
    
    # Caminho para salvar o relatório
    save_path = 'C:/Temp/'
    # Formatação da data para nome do arquivo
    data_formatada = datetime.now().strftime("%Y%m%d")

    # Execução das queries
    print('Executando Querys...')
    result_mdcs = pd.read_sql(query_mdcs, engine_mdcs_sync)
    result_prd = pd.read_sql(query_prd, engine_prd)
    versao_prd = pd.read_sql(query_ve_versao, engine_prd)
    
    # Merge dos dataframes
    print('Realizando Merge dos Registros...')
    result_version = pd.merge(left=result_mdcs, right=result_prd, how='inner', left_on="cLogin", right_on="Usuário")
    
    # Formatação dos campos de data
    print('Formatando os Campos...')
    versao_prd_data = pd.to_datetime(versao_prd['cParameter'].iloc[0])
    result_version['Versão'] = pd.to_datetime(result_version['Versão'])
    
    # Formatação do dataframe
    result_version.drop(columns=['cLogin_y', 'Usuário'], inplace=True)
    result_version.rename(columns={'cLogin_x': 'Veículo'}, inplace=True)
    result_version.drop_duplicates(subset='Serial', inplace=True)
    result_version = result_version[['Veículo', 'Placa', 'Centro', 'Serial', 'Versão']]
    
    # Divisão dos resultados
    print('Criando as Planilhas...')
    result_prd = result_version[result_version['Versão'] == versao_prd_data]
    result_dif = result_version[result_version['Versão'] != versao_prd_data]

    # Verifica se possui Versão Piloto
    print('Validação Piloto...')
    
    p = input('Possui Versão em Piloto? (S ou N): ')
    
    if p.upper() == 'S':

        piloto = input(f"""Copie o Número da Versão Piloto do Site de Download e Cole Abaixo...
        {link}
        : """)
        print('Carregando Informações...')
        piloto_version = piloto.split('.') # recebe o numero da versão copiada do site conforme input
        piloto_version = piloto_version[0]+piloto_version[1]
        piloto_version = pd.to_datetime(piloto_version)
        print('Criando Planilha Piloto...')
        result_pilot = result_dif[result_dif['Versão'] == piloto_version]
        result_dif = result_dif.drop(result_pilot.index)

        # Exportação dos dataframes para Excel
        print('Gerando Relatório de Versão...')
        with pd.ExcelWriter(f'{save_path}RelatorioVersao_{data_formatada}_V1.xlsx', engine='openpyxl') as writer:
            result_version.to_excel(writer, sheet_name='Geral', index=False)
            result_prd.to_excel(writer, sheet_name='Versão Prod', index=False)
            result_dif.to_excel(writer, sheet_name='Versão Incorreta', index=False)
            result_pilot.to_excel(writer, sheet_name='Versão Piloto', index=False)
        print(f"""Relatório salvo no caminho: {save_path}
            Nome do arquivo: RelatorioVersao_{data_formatada}_V1.xlsx""")
        print(f'\nAté o momento foram sincronizados {result_version.shape[0]} dispositivos, '
        f'{result_prd.shape[0]} com a versão de produção, '
        f'{result_dif.shape[0]} dispositivos com a versão incorreta e '
        f'{result_pilot.shape[0]} dispositivos em piloto.\n')
    else:

        # Exportação dos dataframes para Excel
        print('Gerando Relatório de Versão...')
        with pd.ExcelWriter(f'{save_path}RelatorioVersao_{data_formatada}_V1.xlsx', engine='openpyxl') as writer:
            result_version.to_excel(writer, sheet_name='Geral', index=False)
            result_prd.to_excel(writer, sheet_name='Versão Prod', index=False)
            result_dif.to_excel(writer, sheet_name='Versão Incorreta', index=False)
            # Mensagem de sucesso
        print(f"""Relatório salvo no caminho: {save_path}
            Nome do arquivo: RelatorioVersao_{data_formatada}_V1.xlsx""")
        print(f'Até o momento foram sincronizados {result_version.shape[0]} dispositivos, '
        f'{result_prd.shape[0]} com a versão de produção e '
        f'{result_dif.shape[0]} dispositivos com a versão incorreta.')



def main():
    print('Iniciando Programa para Gerar Relatório de Versão...')
    gera_relatorio()
    print('Pressione Qualquer Tecla para Sair')
    input('')

if __name__ == "__main__":
    main()
