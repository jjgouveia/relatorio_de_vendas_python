import pandas as pd
import win32com.client as win32

tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)

faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

produtos_vendidos = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(produtos_vendidos)

ticket_medio = (faturamento['Valor Final'] / produtos_vendidos['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# enviar um email com o relatório

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'gouvik.dev@gmail.com'
mail.Subject = 'Relatório de vendas por Loja'
mail.HTMLBody = f'''
   <p>Prezados,</p>
   
    <p>Segue o relatório de vendas por loja.</p>
    
    <h3>Faturamento:</h3>
    {faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}
    
    <h3>Quantidade vendida:</h3>
    {produtos_vendidos.to_html()}
    
    <h3>Ticket Médio Por Produto:</h3>
    {ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}
    
    <p>Em caso de dúvidas, estou à disposição.</p>
    <h4>Jr Gouveia - CEO na Doodly.</h4>
'''

mail.Send()
print('Relatório enviado com sucesso')
