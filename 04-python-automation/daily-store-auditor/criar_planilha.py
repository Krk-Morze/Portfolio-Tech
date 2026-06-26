import pathlib
import pandas as pd

# 1. Pega o caminho absoluto exato de onde ESTE script está salvo no seu HD
PASTA_DO_SCRIPT = pathlib.Path(__file__).parent.resolve()

# 2. Aponta para a subpasta e CRIA ELA na marra caso ela não exista
PASTA_BASE_DADOS = PASTA_DO_SCRIPT / 'Bases de Dados'
PASTA_BASE_DADOS.mkdir(parents=True, exist_ok=True)

# 3. Monta o caminho absoluto final do arquivo Excel
caminho_absoluto = PASTA_BASE_DADOS / 'Emails.xlsx'

# Coloque o seu e-mail real aqui entre as aspas:
MEU_EMAIL = "xxxxxxxxxxxxxx@outlook.com" # pode ser qualquer um E-mail

dados_exatos = {
    'Loja': ['Astoria', 'Lower Manhattan', "Hell's Kitchen", 'Diretoria'],
    'Gerente': [
        'Marcos Silva',
        'Beatriz Costa',
        'Thiago Reis',
        'Conselho Executivo',
    ],
    'E-mail': [MEU_EMAIL, MEU_EMAIL, MEU_EMAIL, MEU_EMAIL],
}

df_novo = pd.DataFrame(dados_exatos)

df_novo.to_excel(caminho_absoluto, index=False)

print("-" * 50)
print("✅ SUCESSO! O arquivo foi criado no seguinte caminho absoluto:")
print(caminho_absoluto)
print("-" * 50)


