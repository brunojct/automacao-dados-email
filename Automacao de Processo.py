### Passo 1 - Importar Arquivos e Bibliotecas


#importando as tabelas
import pandas as pd
import win32com.client as win32
from pathlib import Path

emails = pd.read_excel(r'Bases de Dados/Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados/Lojas.csv', encoding='latin', sep = ';')
vendas = pd.read_excel(r'Bases de Dados/Vendas.xlsx')

#adicionando o meu e-mail para testar o funcionamento dos envios

for email in emails['E-mail']:
    emails.loc[emails['E-mail']==email, 'E-mail'] = 'testepythonbruno'+ email[20:]
    

### Passo 2 - Criar uma Tabela para cada Loja e Definir o dia do Indicador


#acrescentando o nome da loja na tabela de vendas
tabela_vendas = vendas.merge(lojas, on='ID Loja')

dicionario_lojas = {}

for loja in lojas['Loja']:
    dicionario_lojas[loja] =  (tabela_vendas.loc[tabela_vendas['Loja']==loja, :])

#definindo dia, mês e ano indicador

dia_indicador = max(tabela_vendas['Data']).day
mes_indicador = max(tabela_vendas['Data']).month
ano_indicador = max(tabela_vendas['Data']).year
data_completa = max(tabela_vendas['Data'])
    


### Passo 3 - Salvar a planilha na pasta de backup


caminho_backup = Path.cwd () / (r'Backup Arquivos Lojas')

#criando as pastas
pastas_backup = caminho_backup.iterdir()
lista_pastas_backup = []

for pasta in pastas_backup:
    lista_pastas_backup.append(pasta.name)

for loja in lojas['Loja']:
    if loja not in lista_pastas_backup:
        (caminho_backup / loja) .mkdir()

#salvando os arquivos

for loja in dicionario_lojas:
    nome_arquivo = f'{mes_indicador}_{dia_indicador}_{loja}.xlsx'
    dicionario_lojas[loja].to_excel(caminho_backup / loja / f'{nome_arquivo}')


### Passo 4 - Calcular o indicador para as lojas e enviar email para os gerentes


#definindo as metas

meta_faturamento_dia = 1000
meta_quantidade_produtos_dia = 4
meta_ticket_dia = 500

meta_faturamento_ano = 1650000
meta_quantidade_produtos_ano = 120
meta_ticket_ano = 500


for loja in dicionario_lojas:

    #calcular faturamento dia

    tabela_vendas_loja_ano = dicionario_lojas[loja]
    tabela_vendas_loja_dia = tabela_vendas_loja_ano.loc[tabela_vendas_loja_ano['Data']==data_completa, :]

    faturamento_dia = tabela_vendas_loja_dia['Valor Final'].sum()

    #calcular faturamento ano

    faturamento_ano = tabela_vendas_loja_ano['Valor Final'].sum()


    #calcular diversidade de produtos por dia

    tabela_produtos_loja_dia = tabela_vendas_loja_dia['Produto'].unique()
    quantidade_produtos_dia = (len(tabela_produtos_loja_dia))

    #calcular diversidade de produtos por ano

    tabela_produtos_loja_ano = tabela_vendas_loja_ano['Produto'].unique()
    quantidade_produtos_ano = (len(tabela_produtos_loja_ano))


    #calcular ticket médio dia

    tabela_ticket_medio_dia = tabela_vendas_loja_dia.groupby(by='Código Venda').sum()
    ticket_medio_dia = tabela_ticket_medio_dia['Valor Final'].mean()

    #calcular ticket médio ano

    tabela_ticket_medio_ano = tabela_vendas_loja_ano.groupby(by='Código Venda').sum()
    ticket_medio_ano = tabela_ticket_medio_ano['Valor Final'].mean()
    
     #definindo a cor do botão

    if faturamento_dia>=meta_faturamento_dia:
        cor_faturamento_dia = 'green'
    else:
        cor_faturamento_dia = 'red'

    if faturamento_ano>= meta_faturamento_ano:
        cor_faturamento_ano = 'green'
    else:
        cor_faturamento_ano = 'red'

    if quantidade_produtos_dia>=meta_quantidade_produtos_dia:
        cor_quantidade_produtos_dia = 'green'
    else:
        cor_quantidade_produtos_dia = 'red'

    if quantidade_produtos_ano>=meta_quantidade_produtos_ano:
        cor_quantidade_produtos_ano = 'green'
    else:
        cor_quantidade_produtos_ano = 'red'

    if ticket_medio_dia>=meta_ticket_dia:
        cor_ticket_medio_dia = 'green'
    else:
        cor_ticket_medio_dia = 'red'

    if ticket_medio_ano>=meta_ticket_ano:
        cor_ticket_medio_ano = 'green'
    else:
        cor_ticket_medio_ano = 'red'

    #encontrando o email e o nome do gerente

    email_gerente = emails.loc[emails['Loja']==loja, 'E-mail'].values[0]
    gerente = emails.loc[emails['Loja']==loja, 'Gerente'].values[0]

    #enviando o email

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = email_gerente
    mail.Subject = f'OnePage {dia_indicador}/{mes_indicador}/{ano_indicador} {loja}'
    mail.HTMLBody = f'''
    <p>Bom dia, <strong>{gerente}</strong></p> 

    <p>O resultado da loja <strong>{loja}</strong> no dia <strong>{dia_indicador}/{mes_indicador}/{ano_indicador}</strong> foi:</p>

    <table style = 'width: 50%'>
      <tr>
        <th>Indicador</th>
        <th>Valor Atual</th>
        <th>Meta Diária</th>
        <th>Cenário</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td align = 'center'>R${faturamento_dia:.2f}</td>
        <td align = 'center'>R${meta_faturamento_dia:.2f}</td>
        <td style = 'color: {cor_faturamento_dia}' align = 'center'>◙</td>
      </tr>
      <tr>
        <td>Qtd Produtos</td>
        <td align = 'center'>{quantidade_produtos_dia}</td>
        <td align = 'center'>{meta_quantidade_produtos_dia}</td>
        <td style = 'color: {cor_quantidade_produtos_dia}' align = 'center'>◙</td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td align = 'center'>R${ticket_medio_dia:.2f}</td>
        <td align = 'center'>R${meta_ticket_dia:.2f}</td>
        <td style = 'color: {cor_ticket_medio_dia}' align = 'center'>◙</td>
      </tr>
    </table>

    <p></p>
    <p></p>
    <p></p>

    <table style = 'width: 50%'>
      <tr>
        <th>Indicador</th>
        <th>Valor Atual</th>
        <th>Meta Anual</th>
        <th>Cenário</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td align = 'center'>R${faturamento_ano:.2f}</td>
        <td align = 'center'>R${meta_faturamento_ano:.2f}</td>
        <td style = 'color: {cor_faturamento_ano}' align = 'center'>◙</td>
      </tr>
      <tr>
        <td>Qtd Produtos</td>
        <td align = 'center'>{quantidade_produtos_ano}</td>
        <td align = 'center'>{meta_quantidade_produtos_ano}</td>
        <td style = 'color: {cor_quantidade_produtos_ano}' align = 'center'>◙</td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td align = 'center'>R${ticket_medio_ano:.2f}</td>
        <td align = 'center'>R${meta_ticket_ano:.2f}</td>
        <td style = 'color: {cor_ticket_medio_ano}' align = 'center'>◙</td>
      </tr>
    </table>
    '''

    anexo = Path.cwd() / caminho_backup / loja / f'{mes_indicador}_{dia_indicador}_{loja}.xlsx'
    mail.Attachments.Add(str(anexo))

    mail.Send()


# ### Passo 5 - Criar ranking para diretoria


#ranking anual
ranking_lojas_anual = tabela_vendas[['Loja', 'Valor Final']].groupby(by='Loja').sum()
ranking_lojas_anual = ranking_lojas_anual.sort_values(by='Valor Final', ascending=False)
melhor_loja_ano = ranking_lojas_anual.index[0]
faturamento_melhor_loja_ano = ranking_lojas_anual.iloc[0,0]
pior_loja_ano = ranking_lojas_anual.index[-1]
faturamento_pior_loja_ano = ranking_lojas_anual.iloc[-1, 0]

#ranking diario
ranking_lojas_diario = tabela_vendas.loc[tabela_vendas['Data']==data_completa, ['Loja', 'Valor Final']]
ranking_lojas_diario = ranking_lojas_diario.groupby(by='Loja').sum()
ranking_lojas_diario = ranking_lojas_diario.sort_values(by='Valor Final', ascending=False)
melhor_loja_dia = ranking_lojas_diario.index[0]
faturamento_melhor_loja_dia = ranking_lojas_diario.iloc[0,0]
pior_loja_dia = ranking_lojas_diario.index[-1]
faturamento_pior_loja_dia = ranking_lojas_diario.iloc[-1, 0]

#salvando os arquivos no computador

caminho_diretoria_ano = Path(caminho_backup / f'{mes_indicador}_{dia_indicador}_Ranking Anual.xlsx')
caminho_diretoria_dia = Path(caminho_backup / f'{mes_indicador}_{dia_indicador}_Ranking Diário.xlsx')

ranking_lojas_anual.to_excel(caminho_diretoria_ano)
ranking_lojas_diario.to_excel(caminho_diretoria_dia)



# ### Passo 6 - Enviar e-mail para diretoria



outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Gerente']=='Diretoria', 'E-mail'].values[0]
mail.Subject = f'Ranking das Lojas {dia_indicador}/{mes_indicador}/{ano_indicador}'
mail.Body= f'''
Bom dia! 

A melhor loja de ontem foi {melhor_loja_dia} com faturamento de R${faturamento_melhor_loja_dia:.2f}
A pior loja de ontem foi {pior_loja_dia} com faturamento de R${faturamento_pior_loja_dia:.2f}

A melhor loja do ano está sendo {melhor_loja_ano} com faturamento de R${faturamento_melhor_loja_ano:.2f}
A pior loja do ano está sendo {pior_loja_ano} com faturamento de R${faturamento_pior_loja_ano:.2f}

Segue em anexo o ranking completo
'''

anexo1 = (str(caminho_diretoria_ano))
anexo2 = (str(caminho_diretoria_dia))
mail.Attachments.Add(anexo1)
mail.Attachments.Add(anexo2)

mail.Send()

