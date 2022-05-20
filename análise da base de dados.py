import pandas as pd
import win32com.client as win32


# importar a base de dados


tabela_vendas = pd.read_excel(r"C:\Users\Igor Mello\OneDrive\Área de Trabalho\vscode project\outros\AnBaDa\Vendas.xlsx")


# visualizar a base de dados

pd.set_option('display.max_columns', None)

print(tabela_vendas)


# faturamento por loja

faturamento = tabela_vendas [['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

print(faturamento)



# quantidade de produtos vendidos por loja

quantidade = tabela_vendas [['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

print(quantidade)

print('-'*50)
#ticket médio por produto em cada loja

ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns= {0:  'Ticket Médio'})
print(ticket_medio)

# enviar um email com o relatório

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'example@eq.ufrj.br'
mail.Subject = 'Relatório de vendas por loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>segue o relatório de vendas por cada loja.</p>
<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}
 

<p>Quantidade:</p>
{quantidade.to_html()}

<p>Ticket Médio por loja:</p>
{ticket_medio.to_html(formatters = {'Ticket Médio': 'R${:,.2f}'.format})}
<p>Att.</p>
<p>Igor Mello</p>


'''

mail.Send()
print('Email enviado')