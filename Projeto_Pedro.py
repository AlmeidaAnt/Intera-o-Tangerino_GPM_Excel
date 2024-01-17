import time
from selenium import webdriver
from datetime import datetime
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import openpyxl
import os
from shutil import move
import pandas as pd

navegador = webdriver.Chrome()
navegador.maximize_window()

# Abrir Página do Sistema
navegador.get('https://app.tangerino.com.br/Tangerino/?wicket:bookmarkablePage=wicket-4:com.frw.tangerino.web.pages.web.cadastro.LoginFuncionarioPage&wicket:interface=wicket-4:19::INewBrowserWindowListener::')

#Campo Login Empresa
campo_login = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div/div/div[1]/form/fieldset/div[1]/input')
campo_login.send_keys('LJMAP')

#Campo PIN
pin = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div/div/div[1]/form/fieldset/div[2]/input')
pin.send_keys('5735')

#Botão Login
botao = navegador.find_element(By.XPATH,'/html/body/div[1]/div[1]/div/div/div[1]/form/fieldset/div[3]/input')
botao.click()

time.sleep(10)

navegador.refresh()
time.sleep(3)

#Acessar Relatorio Analitico
navegador.get('https://app.tangerino.com.br/Tangerino/pages/relatorio/banco-horas?funcionalidade=16&wicket:pageMapName=wicket-0')

#Horas Extras
he = navegador.find_element(By.XPATH,'/html/body/div[1]/main/div/div/div[2]/span/div/div[1]/ul/li[2]/a')
he.click()

time.sleep(5)

#Tipo de Documento Para Baixar
tipo_doc = navegador.find_element(By.XPATH,'/html/body/div[1]/main/div/div/div[2]/span/div/div[2]/form/fieldset/div[16]/span/span[1]/span/span[1]')
tipo_doc.click()

doc = navegador.find_element(By.XPATH,'/html/body/span[2]/span/span[1]/input')
doc.send_keys('Excel')
time.sleep(2)
doc.send_keys(Keys.ENTER)
time.sleep(5)

#Baixar Documento
baixar = navegador.find_element(By.XPATH,'/html/body/div[1]/main/div/div/div[2]/span/div/div[2]/form/fieldset/div[18]/input[1]')
baixar.click()
time.sleep(30)

# MOVER DOCUMENTO #

# Obtém o diretório de downloads do usuário
pasta_downloads = os.path.expanduser('~/Downloads')

# Lista todos os arquivos na pasta de downloads
arquivos_downloads = [os.path.join(pasta_downloads, arquivo) for arquivo in os.listdir(pasta_downloads)]

# Imprime a lista de arquivos para ver se está correto
print("Lista de arquivos na pasta de downloads:")
print(arquivos_downloads)

# Obtém o arquivo mais recente com base na data de modificação
arquivo_mais_recente = max(arquivos_downloads, key=os.path.getmtime)

# Imprime o arquivo mais recente antes de mover
print("Arquivo mais recente antes de mover:")
print(arquivo_mais_recente)

# Novo diretório
novo_diretorio = 'C:\\temp'

# Obtém a data atual no formato desejado
data_atual = datetime.now().strftime('%d%m%Y')

# Sufixo "TANGERINO"
sufixo = 'TANGERINO'

# Obtém o nome do arquivo mais recente
nome_original = os.path.basename(arquivo_mais_recente)
# Novo nome do arquivo
novo_nome = f'{data_atual}_{sufixo}_{nome_original}'

# Caminho completo do novo arquivo
novo_caminho = os.path.join(novo_diretorio, novo_nome)

# Move e renomeia o arquivo mais recente
move(arquivo_mais_recente, novo_caminho)

print(f'O arquivo mais recente foi movido e renomeado para: {novo_caminho}')

#GPM


# 1. Abrir o GPM
navegador.get('https://endiconpa.gpm.srv.br/index.php')

# 1.1 Inserir credenciais de acesso - LOGIN
usuario = navegador.find_element(By.XPATH, '//*[@id="idLogin"]')
usuario.send_keys('PEDRO.MORAES')
time.sleep(1)

# 1.2 Inserir credenciais de acesso - SENHA
senha = navegador.find_element(By.XPATH,'//*[@id="idSenha"]')
senha.send_keys('E1234567@')


# 1.3 Inserir credenciais de acesso - LOGAR
botao_login = navegador.find_element(By.XPATH,'//*[@id="form_login"]/input[5]')
botao_login.click()
time.sleep(1)

# 2. Navegar até a aba de cadastro de Multas
consulta_indv = 'https://endiconpa.gpm.srv.br/gpm/geral/checklist_individual_consulta.php'
navegador.get(consulta_indv)

#Data Inicial
dt_inicial = navegador.find_element(By.XPATH,'//*[@id="id_data_in"]')
dt_inicial.send_keys('01/01/2024')
dt_inicial.send_keys(Keys.TAB)

#Data Final
dt_final = navegador.find_element(By.XPATH,'//*[@id="id_data_out"]')
dt_final.send_keys(data_atual)

#Pesquisar
pesquisar = navegador.find_element(By.XPATH,'/html/body/form[2]/input')
pesquisar.click()
time.sleep(5)
#Baixar Excel
excel = navegador.find_element(By.XPATH,'//*[@id="tab_resultados_wrapper"]/div[1]/button[2]')
excel.click()
time.sleep(15)

#MUDAR LOCAL DO ARQUIVO

# Obtém o diretório de downloads do usuário
pasta_downloads2 = os.path.expanduser('~/Downloads')

# Lista todos os arquivos na pasta de downloads
arquivos_downloads2 = [os.path.join(pasta_downloads2, arquivo) for arquivo in os.listdir(pasta_downloads)]

# Imprime a lista de arquivos para ver se está correto
print("Lista de arquivos na pasta de downloads:")
print(arquivos_downloads2)

# Obtém o arquivo mais recente com base na data de modificação
arquivo_mais_recente2 = max(arquivos_downloads2, key=os.path.getmtime)

# Imprime o arquivo mais recente antes de mover
print("Arquivo mais recente antes de mover:")
print(arquivo_mais_recente2)

# Novo diretório
novo_diretorio2 = 'C:\\temp'

# Sufixo "GPM"
sufixo2 = 'GPM'

# Obtém o nome do arquivo mais recente
nome_original2 = os.path.basename(arquivo_mais_recente2)
# Novo nome do arquivo
novo_nome2 = f'{data_atual}_{sufixo2}_{nome_original2}'

# Caminho completo do novo arquivo
novo_caminho2 = os.path.join(novo_diretorio2, novo_nome2)

# Move e renomeia o arquivo mais recente
move(arquivo_mais_recente2, novo_caminho2)

print(f'O arquivo mais recente foi movido e renomeado para: {novo_caminho2}')


# Definir caminhos dos arquivos
caminho_arquivo2 = os.path.join(novo_diretorio, novo_nome)
caminho_arquivo1 = os.path.join(novo_diretorio2, novo_nome2)
caminho_arquivo_combinado = os.path.join('C:\\temp', f'{data_atual}_Relatorio_combinado.xls')  # Substitua pelo caminho desejado

# Ler os arquivos Excel
df_arquivo1 = pd.read_excel(caminho_arquivo1, header=1)
df_arquivo2 = pd.read_excel(caminho_arquivo2, header=3)

print(df_arquivo1.columns)


# Dividir a coluna 'Matricula' pelo delimitador "-" e renomear as novas colunas
df_arquivo1[['Matricula', 'Nome Funcionario']] = df_arquivo1['Matricula'].str.split('-', n=1, expand=True)

# Adicionar uma nova coluna antes da coluna A na aba GPM com a fórmula "=B2+G2"
df_arquivo1.insert(0, 'Chave', df_arquivo1['Matricula'] + df_arquivo1['Data de cadastro'].astype(str))

# Corrigir a formatação da coluna 'Data' na aba 'Tangerino'
df_arquivo2['Data'] = pd.to_datetime(df_arquivo2['Data'], errors='coerce')
df_arquivo2['Data'] = df_arquivo2['Data'].dt.strftime('%d/%m/%Y')

# Excluir linhas na aba 'Tangerino' que contenham as informações específicas
palavras_chave = ["Centro de Custo", "Relatório de Horas Extras e Faltas", "Período:", "Empregador:"]
df_arquivo2 = df_arquivo2[~df_arquivo2.iloc[:, 0].astype(str).str.contains('|'.join(palavras_chave))]

# Garantir que a coluna 'Chave' existe no DataFrame 'df_arquivo2'
if 'Chave' not in df_arquivo2.columns:
    df_arquivo2.insert(0, 'Chave', df_arquivo2['Cod.Externo'] + ' ' + df_arquivo2['Data'])

# Mesclar com a tabela GPM para obter os valores correspondentes usando o método 'merge'
df_arquivo2 = pd.merge(df_arquivo2, df_arquivo1[['Chave', 'Equipe']], how='left', on='Chave')

# Adicionar uma nova coluna antes da coluna 'Chave' na aba 'Tangerino'
df_arquivo2.insert(df_arquivo2.columns.get_loc('Chave'), 'cruzamento', df_arquivo2['Equipe'])  # Substitua 'Nome Funcionario' conforme necessário

# Criar um escritor Excel
with pd.ExcelWriter(caminho_arquivo_combinado, engine='xlsxwriter') as writer:
    # Salvar o primeiro DataFrame na aba 'GPM'
    df_arquivo1.to_excel(writer, sheet_name='GPM', na_rep='', header=True, index=False)

    # Salvar o segundo DataFrame na aba 'Tangerino'
    df_arquivo2.to_excel(writer, sheet_name='Tangerino', na_rep='', header=True, index=False)