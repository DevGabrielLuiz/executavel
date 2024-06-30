import pandas as pd
import win32com.client as win32

# importar a base de dados python
tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados 
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# faturamento por loja 
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-'*50)
# ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# enviar um email com relatorio
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'chris1406813@gmail.com'
mail.Subject = 'Relatóro de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:.2f}'.format})}

<p>Qualquer dúvida estou à disposição</p>

<p>Fala Maicon.,</p>
<p>Gabriel</p>
'''
mail.Send()

print('Email Enviado')

# se o arquvio tiver que abrir uma janela eu devo passar o paramentro pyinstaller --onefile -w [nome do arquivo]