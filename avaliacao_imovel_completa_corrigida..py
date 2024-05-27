import openpyxl

# Cria um novo arquivo de planilha
workbook = openpyxl.Workbook()
sheet = workbook.active

# Dados a serem escritos na planilha
dados = '''
AVALIADOR E CORRETOR DE IMÓVEIS
EDSON VIANA ESMECELATO
CNAI      47536       CRECI 44628
Terra Roxa / PR

Imóvel Descrito

Tipo de construção
Alvenaria
Posição  DA CONSTRUÇÃO EM RELAÇÃO LOTE
Frente para rua
Dormitórios
3
Suíte
1
Banheiros
2
Cozinha
1
Salas
1
Lavabos
2
Edícula, Garagens, Áreas nos Fundo e Frente
60
Ano da Construção
2010
Financiamento  SIM OU NÃO
SIM
Área Total Terreno
300
Área Privativa
300
Área Construída
142
Localização
Bairro
CENTRO
Endereço
Rua Azauri Guedes Pereira, 85
CEP
85990-000
INÍCIO DE CÁLCULOS REDUÇÕES E ACRÉSCIMOS
CUB Paraná Maio 2024
R$ 1.962,69
Valor Total da Construção
R$ 278.701,98
Depreciação pelo ano de construção 5 % a 10 %
R$ 13.935,10
Redução pelas áreas abertas Garagens e Áreas Edícula 50% OU 70%
R$ 35.328,42
Reformas e Pintura Necessária
R$ 8.361,06
ATÉ AQUI valor da construção sem terreno
R$ 221.077,40
Valor do terreno restante em aberto
R$ 30.000,00
Avaliação Construção em Laje
R$ 33.444,24
Avaliação em Porcelanato
R$ 16.722,12
Cobertura em Telha
R$ 8.361,06
Valorização ou Desv. Pela Localização +10%  ou -10%
-R$ 27.870,20
FINALIZANDO A AVALIAÇÃO
R$ 221.734,62
OSCILAÇÃO PARA CIMA
R$ 232.821,35
OSCILAÇÃO PARA BAIXO
R$ 210.647,89
'''

# Divide os dados em linhas
linhas = dados.strip().split('\n')

# Escreve cada linha na planilha na coluna A
for i, linha in enumerate(linhas):
    celula = sheet.cell(row=i+1, column=1)
    celula.value = linha

# Salva a planilha
workbook.save('imovel.xlsx')
