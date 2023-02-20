import pandas as pd
import win32com.client as win32

tabela_vendas = pd.read_excel('Cópia de vendas.xlsx')
# Display de colunas ilimitada
pd.set_option('display.max_columns', None)
tabela_vendas = tabela_vendas[['ID Loja','Valor Unitário','Valor Final','Quantidade']]
# Faturamento por loja
Faturamento_por_loja = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum(numeric_only = True)
# Quantidade de vendas por loja
qt_produtos_vd_por_loja = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum(numeric_only = True)
# Ticket médio
Ticket_Médio = (Faturamento_por_loja['Valor Final']/qt_produtos_vd_por_loja['Quantidade']).to_frame()
Ticket_Médio= Ticket_Médio.rename(columns={0: 'Ticket Médio'})


# Envio de email automático com relatório

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.to = 'Email para qual o relatório deve ser enviado'
mail.Subject = 'Relatório'
mail.HTMLBody = '''
<p>Segue o Relatório de Vendas por Cada Loja<p/>


<p>Faturamento: <p/>
{}



<p>Quantidade de Produtos Vendidos: <p/>
{}



<p>Ticket Médio dos Produtos em Cada Loja: <p/>
{}



'''.format(Faturamento_por_loja.to_html(),qt_produtos_vd_por_loja.to_html(),Ticket_Médio.to_html())
mail.send