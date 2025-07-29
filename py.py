
from IPython.display import display

import pandas as pd

import pathlib

import smtplib
import email.message

emails = pd.read_excel(r'Bases de Dados/Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados/Lojas.csv', encoding = 'latin1', sep=';')
vendas = pd.read_excel(r'Bases de Dados/Vendas.xlsx')

display(emails)
display(lojas)
display(vendas)

#incluir nome da loja em vendas
vendas = vendas.merge(lojas, on='ID Loja')
display(vendas)

dicionario_lojas = {}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja'] == loja, :]

display(dicionario_lojas['Rio Mar Recife'])
display(dicionario_lojas['Shopping Vila Velha'])

dia_indicador = vendas['Data'].max()
print(dia_indicador)
print('{}/{}'.format(dia_indicador.day, dia_indicador.month))
    
#identificar se a pasta já existe

caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')

arquivos_pasta_backup = caminho_backup.iterdir()
lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]

for loja in dicionario_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()
    
    #salvar dentro da pasta
    #local_arquivo = "C:/User...12_26.loja.xlsx"

    nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.month, dia_indicador.day, loja)    
    local_arquivo = caminho_backup / loja / nome_arquivo

    dicionario_lojas[loja].to_excel(local_arquivo)

#definição de metas

meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500


#Calcular o indicador para 1 loja

emailcont = 0

for loja in dicionario_lojas:   

    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador ,:]

    #faturamento
    faturamento_ano = vendas_loja['Valor Final'].sum()
    #print(faturamento_ano)

    faturamento_dia = vendas_loja_dia['Valor Final'].sum()
    #print(faturamento_dia)

    #diversidade de produtos

    qtde_produtos_ano = len(vendas_loja['Produto'].unique())
    qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())

    #ticket medio
    valor_venda = vendas_loja.groupby('Código Venda').sum(numeric_only=True)
    ticket_media_ano = valor_venda['Valor Final'].mean()
    #print(ticket_media_ano)

    #ticket_medio_dia
    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()

    #enviar e-mail   

    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.application import MIMEApplication
    
    def enviar_email():
        
        msg = MIMEMultipart()
        #msg = email.message.Message()
        nome    = emails.loc[emails['Loja']==loja, 'Gerente'].values[0]
        destino = emails.loc[emails['Loja']==loja, 'E-mail'].values[0]

        # Assunto personalizado
        msg["Subject"] = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'
        msg["From"] = "macedo.anderson@gmail.com"
        msg["To"] = "macedo.anderson@gmail.com"

        corpo_email = """
        <html>
            <body>
                <h2 style="color:blue;">Boa tarde, tudo bem?</h2>
                <p>Este é um <strong>exemplo de e-mail com HTML</strong> enviado usando Python e Gmail.</p>
                <p style="color:green;">Funciona parecido com <code>mail.HTMLBody</code> do Outlook.</p>
                <hr>
                <p>Atenciosamente,<br>Anderson Macedo</p>
            </body>
        </html>
        """

        if faturamento_dia >= meta_faturamento_dia:
            cor_fat_dia = 'green'
        else:
            cor_fat_dia = 'red'

        if faturamento_ano >= meta_faturamento_ano:
            cor_fat_ano = 'green'
        else:
            cor_fat_ano = 'red'

        if qtde_produtos_dia >= meta_qtdeprodutos_dia:
            cor_qtde_dia = 'green'
        else:
            cor_qtde_dia = 'red'

        if qtde_produtos_ano >= meta_qtdeprodutos_ano:
            cor_qtde_ano = 'green'
        else:
            cor_qtde_ano = 'red'

        if ticket_medio_dia > meta_ticketmedio_dia:
            cor_ticket_dia = 'green'
        else:
            cor_ticket_dia = 'red'

        if ticket_media_ano >= meta_ticketmedio_ano:
            cor_ticket_ano = 'green'
        else:
            cor_ticket_ano = 'red'
        
        
        corpo_email = f'''
            <p>Bom dia, {nome}</p>

            <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong> foi:</p>

            <table>
            <tr>
                <th>Indicador</th>
                <th>Valor Dia</th>
                <th>Meta Dia</th>
                <th>Cenário Dia</th>
            </tr>
            <tr>
                <td>Faturamento</td>
                <td style="text-align: center">R${faturamento_dia:.2f}</td>
                <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
                <td style="text-align: center"><font color = "{cor_fat_dia}">◙</font></td>
            </tr>
            <tr>
                <td>Diversidade de Produtos</td>
                <td style="text-align: center">{qtde_produtos_dia}</td>
                <td style="text-align: center">{meta_qtdeprodutos_dia}</td>
                <td style="text-align: center"><font color = "{cor_qtde_dia}">◙</font></td>
            </tr>
            <tr>
                <td>Ticket Médio</td>
                <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
                <td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
                <td style="text-align: center"><font color = "{cor_ticket_dia}">◙</font></td>
            </tr>

            </tr>
            </table>

            <br>

            <table>
            <tr>
                <th>Indicador</th>
                <th>Valor Ano</th>
                <th>Meta Ano</th>
                <th>Cenário Ano</th>
            </tr>
            <tr>
                <td>Faturamento</td>
                <td style="text-align: center">R${faturamento_ano:.2f}</td>
                <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
                <td style="text-align: center"><font color = "{cor_fat_ano}">◙</font></td>
            </tr>
            <tr>
                <td>Diversidade de Produtos</td>
                <td style="text-align: center">{qtde_produtos_ano}</td>
                <td style="text-align: center">{meta_qtdeprodutos_ano}</td>
                <td style="text-align: center"><font color = "{cor_qtde_ano}">◙</font></td>
            </tr>
            <tr>
                <td>Ticket Médio</td>
                <td style="text-align: center">R${ticket_media_ano:.2f}</td>
                <td style="text-align: center">R${meta_ticketmedio_ano:.2f}</td>
                <td style="text-align: center"><font color = "{cor_ticket_ano}">◙</font></td>
            </tr>

            </tr>
            </table>

            <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

            <p>Qualquer dúvida estou à disposição.</p>

            <p><strong>Att.</strong> </p>
            <p>Anderson Macedo</p>

        
        '''

        msg.attach(MIMEText(corpo_email, "html", _charset="utf-8"))

        #colocando o anexo
        caminho_arquivo = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
        
        # Abrir e anexar o arquivo
        with open(caminho_arquivo, "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="xlsx")
            anexo.add_header("Content-Disposition", "attachment", filename=caminho_arquivo.name)
            msg.attach(anexo)


        servidor = smtplib.SMTP("smtp.gmail.com",587)
        servidor.starttls()
        servidor.login(msg["From"],"mbqt wwem ujdt jvsy")
        servidor.send_message(msg)
        servidor.quit()    
    
    enviar_email()

    emailcont += 1 
    print(f'{emailcont} - Email da loja: {loja} enviado com sucesso')
    
print('Envio finalizado com sucesso...')

