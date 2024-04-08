from openpyxl import load_workbook
import csv

try:
    planilhas = load_workbook(filename='credito.xlsx')
    planilha = planilhas.active

    cabecalho = next(planilha.values)

    id = []
    sexo = []
    idade = []
    default = []
    estado_civil = []

    indice_id = cabecalho.index('id')
    indice_sexo = cabecalho.index('sexo')
    indice_idade = cabecalho.index('idade')
    indice_default = cabecalho.index('default')
    indice_estado_civil = cabecalho.index('estado_civil')

    for linha in planilha.values:
        if linha[indice_estado_civil] == 'solteiro' and linha[indice_default] == 1:
            id.append(linha[indice_id])
            sexo.append(linha[indice_sexo])
            idade.append(linha[indice_idade])

    with open(file='./credito.csv', mode='w', newline='') as arquivo_csv:
        escritor_csv = csv.writer(arquivo_csv, delimiter=';')
        escritor_csv.writerow(['id', 'sexo', 'idade'])
        dados = map(lambda id, sexo, idade: [id, sexo, idade], id, sexo, idade)
        escritor_csv.writerows(dados)

except Exception as exc:
    print('Erro, parando a execução.')
    raise exc
else:
    print('Arquivo csv criado com sucesso!')
