import pandas as pd

# Carrega o Excel com os dados
df = pd.read_excel('dados3.xlsx')

# Abre o template HTML usando a codificação correta (utf-8)
with open('OFERTAS.html', 'r', encoding='utf-8') as file:
    html_template = file.read()

# Lista de variáveis que precisam ser substituídas
variaveis = ['BANNER', 'DESCRICAO','CTA','FOTOPROD', 'MARCA','NOMEPROD','DE','POR','CONDICAO','LIMITE','OBSERVACAO','EXCETOLOJAS','DATAINICIO','DATAFIM']

# Substitui as variáveis no HTML conforme as linhas do Excel
email_content = html_template

for var in variaveis:
    for index, row in df.iterrows():
        placeholder = f'<#{var.upper()}{str(index + 1).zfill(2)}>'
        
        # Verifica se a coluna existe no DataFrame antes de tentar substituí-la
        if var in df.columns:
            email_content = email_content.replace(placeholder, str(row[var]))
        else:
            print(f"A coluna '{var}' não foi encontrada no Excel. Verifique o arquivo.")

# Aqui você pode salvar ou enviar o email_content para o cliente
with open('email_final.html', 'w', encoding='utf-8') as output_file:
    output_file.write(email_content)
