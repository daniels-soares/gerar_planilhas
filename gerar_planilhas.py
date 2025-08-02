import pandas as pd
import random
from datetime import datetime, timedelta

# Listas base para geração de dados
nomes = ["Ana", "Bruno", "Carlos", "Daniela", "Eduardo", "Fernanda", "Gustavo", "Helena", "Igor", "Juliana"]
produtos = ["Celular", "Notebook", "Fone de Ouvido", "Teclado", "Mouse", "Monitor", "Tablet", "Webcam"]
status_falta = ["Presente", "Faltou", "Justificado"]

# Data base para geração de datas
data_inicio = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)

# Gera dados de transações
df_transacoes = pd.DataFrame({
    "ID": range(1, 101),
    "Cliente": [random.choice(nomes) for _ in range(100)],
    "Valor (R$)": [round(random.uniform(10, 2000), 2) for _ in range(100)],
    "Data": [(data_inicio - timedelta(days=random.randint(0, 30))).strftime('%Y-%m-%d') for _ in range(100)],
    "Fraude": [random.choices([0, 1], weights=[95, 5])[0] for _ in range(100)]
})

# Gera dados de vendas
df_vendas = pd.DataFrame({
    "VendaID": range(1, 101),
    "Produto": [random.choice(produtos) for _ in range(100)],
    "Cliente": [random.choice(nomes) for _ in range(100)],
    "Quantidade": [random.randint(1, 5) for _ in range(100)],
    "ValorUnitario (R$)": [round(random.uniform(100, 1500), 2) for _ in range(100)],
    "DataCompra": [(data_inicio - timedelta(days=random.randint(0, 60))).strftime('%Y-%m-%d') for _ in range(100)],
})

# Gera dados de frequência
df_frequencia = pd.DataFrame({
    "Professor": [random.choice(nomes) for _ in range(100)],
    "Data": [(data_inicio - timedelta(days=random.randint(0, 60))).strftime('%Y-%m-%d') for _ in range(100)],
    "Status": [random.choice(status_falta) for _ in range(100)]
})

# Gera e salva o arquivo Excel com múltiplas abas
nome_arquivo = "planilha_treinamento.xlsx"

with pd.ExcelWriter(nome_arquivo, engine='xlsxwriter') as writer:
    df_transacoes.to_excel(writer, sheet_name="Transacoes", index=False)
    df_vendas.to_excel(writer, sheet_name="Vendas", index=False)
    df_frequencia.to_excel(writer, sheet_name="Frequencia", index=False)

print(f"Arquivo '{nome_arquivo}' criado com sucesso!")
