import pandas as pd
import smtplib
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# importar a base de dados
print("Importando base de dados...")
table_sales = pd.read_excel('Vendas.xlsx')

# faturamento por loja
earnings = table_sales[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

# quantidade de produtos vendidos por loja
amount_products_per_store = table_sales[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

# ticket medio da loja
ticket_medio = (earnings['Valor Final'] / amount_products_per_store['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'})
# enviar um email com relatorio

# variables
sender_email = str(input("Insira o email REMETENTE: "))
password = str(input("Digite a senha: "))
receiver_email = str(input("Insira o email DESTINATARIO: "))
subject = str(input("Insira o ASSUNTO do email: "))

# setup the parameters of the message
msg = MIMEMultipart()
msg['From'] = sender_email
msg['To'] = receiver_email
msg['Subject'] = subject
html = f"""
<html>
    <body>
        <h2>Segundo Relatorio de Vendas dos Shoppings de São Paulo</h2>
        <h3>Faturamento de cada loja:</h3>
        <p>
        {earnings.to_html()}
        </p>
        <br>
        <h3>Produtos vendidos por loja:</h3>
        <p>
        {amount_products_per_store.to_html()}
        </p>
        <br>
        <h3>Ticket medio por loja:</h3>
        <p>
        {ticket_medio.to_html()}
        </p>
        <br>
        Dúvidas me contate, Obrigado.
        </p>
    </body>
</html>

"""

# add in the message body and send email
part = MIMEText(html, "html")
msg.attach(part)
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.login(sender_email, password)
server.sendmail(sender_email, receiver_email, msg.as_string())
server.quit()