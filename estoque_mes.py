import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime

# Configurações iniciais
caminho_pasta = r'C:\Users\Win10\OneDrive\Documentos\Custos médios - Junho 2025'
downloads_path = os.path.join(os.path.expanduser('~'), 'Downloads')
arquivo_saida = os.path.join(downloads_path, 'Custos_Consolidados.xlsx')

# Criar um writer para o arquivo Excel de saída
writer = pd.ExcelWriter(arquivo_saida, engine='openpyxl')

# Dicionário para armazenar os dados consolidados
dados_consolidados = {}

# Dicionário para armazenar os DataFrames de cada aba individual
abas_individuais = {}

# Lista para armazenar todas as datas encontradas
datas_encontradas = []

# Processar cada arquivo na pasta
for arquivo in os.listdir(caminho_pasta):
    if arquivo.startswith('ev') and (arquivo.endswith('.xlsx') or arquivo.endswith('.csv')):
        # Extrair data do nome do arquivo (assumindo formato evddmmyy)
        data_str = arquivo[2:8]  # Pega os 6 dígitos após 'ev'
        
        try:
            # Converter para objeto de data
            data = datetime.strptime(data_str, '%d%m%y').date()
            data_formatada = data.strftime('%d/%m/%Y')
            datas_encontradas.append(data_formatada)
            
            # Determinar o caminho completo do arquivo
            caminho_arquivo = os.path.join(caminho_pasta, arquivo)
            
            # Ler o arquivo (Excel ou CSV)
            if arquivo.endswith('.xlsx'):
                df = pd.read_excel(caminho_arquivo, skiprows=2)  # Pular 2 linhas de cabeçalho
            else:
                # Para CSV, primeiro detectar o delimitador
                with open(caminho_arquivo, 'r', encoding='latin1') as f:
                    first_line = f.readline()
                    second_line = f.readline()
                    third_line = f.readline()
                
                # Verificar se o cabeçalho está na terceira linha
                if 'PRODUTO' in third_line:
                    # Se o delimitador for tabulação
                    if '\t' in third_line:
                        df = pd.read_csv(caminho_arquivo, skiprows=2, delimiter='\t', encoding='latin1')
                    # Se for ponto e vírgula
                    elif ';' in third_line:
                        df = pd.read_csv(caminho_arquivo, skiprows=2, delimiter=';', encoding='latin1')
                    # Se for vírgula
                    elif ',' in third_line:
                        df = pd.read_csv(caminho_arquivo, skiprows=2, delimiter=',', encoding='latin1')
                    else:
                        # Tentar ler sem especificar delimitador
                        df = pd.read_csv(caminho_arquivo, skiprows=2, encoding='latin1')
                else:
                    # Se 'PRODUTO' não estiver na terceira linha, tentar encontrar
                    df = pd.read_csv(caminho_arquivo, encoding='latin1')
                    # Procurar a linha que contém 'PRODUTO'
                    for i, row in df.iterrows():
                        if 'PRODUTO' in str(row.values):
                            df = pd.read_csv(caminho_arquivo, skiprows=i, encoding='latin1')
                            break
            
            # Renomear colunas para padrão (remover espaços extras)
            df.columns = df.columns.str.strip()
            
            # Armazenar o DataFrame para escrever depois
            abas_individuais[data_str] = df
            
            # Processar dados para a planilha consolidada
            for _, row in df.iterrows():
                # Verificar nomes alternativos das colunas
                produto = row.get('PRODUTO', row.get('Produto', row.get('produto', '')))
                descricao = row.get('DESCRICAO', row.get('Descricao', row.get('descricao', '')))
                grupo = row.get('GRUPO', row.get('Grupo', row.get('grupo', '')))
                custo = row.get('CUSTO', row.get('Custo', row.get('custo', '')))
                
                if produto and pd.notna(produto):
                    if produto not in dados_consolidados:
                        dados_consolidados[produto] = {
                            'DESCRICAO': descricao,
                            'GRUPO': grupo,
                            'CUSTOS': {}
                        }
                    
                    dados_consolidados[produto]['CUSTOS'][data_formatada] = custo
        
        except Exception as e:
            print(f"Erro ao processar o arquivo {arquivo}: {str(e)}")

# Ordenar as datas
datas_encontradas.sort(key=lambda x: datetime.strptime(x, '%d/%m/%Y'))

# Criar DataFrame consolidado
linhas_consolidadas = []
for produto, dados in dados_consolidados.items():
    linha = {
        'PRODUTO': produto,
        'DESCRICAO': dados['DESCRICAO'],
        'GRUPO': dados['GRUPO']
    }
    
    for data in datas_encontradas:
        custo = dados['CUSTOS'].get(data, '')
        # Formatar como moeda se houver valor
        if custo != '' and pd.notna(custo):
            try:
                linha[data] = f'R$ {float(custo):,.2f}'.replace('.', ',').replace(',', '.', 1)
            except:
                linha[data] = custo
        else:
            linha[data] = ''
    
    linhas_consolidadas.append(linha)

df_consolidado = pd.DataFrame(linhas_consolidadas)

# Escrever primeiro a aba consolidada
df_consolidado.to_excel(writer, sheet_name='Consolidado', index=False)

# Depois escrever as abas individuais
for data_str, df in abas_individuais.items():
    df.to_excel(writer, sheet_name=data_str, index=False)

# Ajustar a largura das colunas para todas as abas
workbook = writer.book
for sheet_name in workbook.sheetnames:
    worksheet = workbook[sheet_name]
    
    # Ajustar a largura de todas as colunas
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        
        # Adicionar um pequeno buffer
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column_letter].width = adjusted_width

# Salvar o arquivo Excel
writer.close()

print(f"Processo concluído. Arquivo gerado em: {arquivo_saida}")