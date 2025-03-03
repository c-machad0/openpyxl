from openpyxl import load_workbook
import random

def cria_nome_produto(produtos_disponiveis):
    return random.choice(produtos_disponiveis)

wb = load_workbook('dados.xlsx')

sheets = ['Janeiro', 'Fevereiro', 'Março', 'Resumo']
headers = ['Produto', 'Quantidade', 'Preço Unitário', 'Total']

for sheet in sheets:
    ws = wb.create_sheet(f'{sheet}')
    for idx, header in enumerate(headers, start=1):
        ws.cell(row=1, column=idx, value=header)

    if sheet != 'Resumo':
        produtos = ['Uva', 'Morango', 'Maçã', 'Manga', 'Kiwi']

        for row in range(2, 2 + len(produtos)):
            nome_produto = cria_nome_produto(produtos)
            produtos.remove(nome_produto)
            quantidade_produto = random.randint(1, 20)
            preco_produto = round(random.uniform(1, 10), 2)

            ws.cell(row=row, column=1, value=nome_produto)
            ws.cell(row=row, column=2, value=quantidade_produto)
            ws.cell(row=row, column=3, value=preco_produto)
 
wb.save('dados.xlsx')