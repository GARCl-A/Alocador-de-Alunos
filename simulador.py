import pandas as pd #Biblioteca pra lidar com planilhas

#Glossário - Corpo: uma das três possíveis áreas; Divisão: subdivisão de um corpo

#Objeto do aluno com os atributos necessários
class Aluno():
    def __init__(self,nome,primeira_op,segunda_op,terceira_op,primeira_opcao_corpo1,segunda_opcao_corpo1,terceira_opcao_corpo1,primeira_opcao_corpo2,segunda_opcao_corpo2,terceira_opcao_corpo2, classificacao):
        self._nome = nome.strip().capitalize()
        self._primeira_op = primeira_op.strip().capitalize()
        self._segunda_op = segunda_op.strip().capitalize()
        self._terceira_op = terceira_op.strip().capitalize()
        self._primeira_opcao_corpo1 = primeira_opcao_corpo1.strip().capitalize()
        self._segunda_opcao_corpo1 = segunda_opcao_corpo1.strip().capitalize()
        self._terceira_opcao_corpo1 = terceira_opcao_corpo1.strip().capitalize()
        self._primeira_opcao_corpo2 = primeira_opcao_corpo2.strip().capitalize()
        self._segunda_opcao_corpo2 = segunda_opcao_corpo2.strip().capitalize()
        self._terceira_opcao_corpo2 = terceira_opcao_corpo2.strip().capitalize()
        self._classificacao = classificacao

#Criando os objetos iterando sobre as colunas do formulário
#As colunas já estão na ordem correta        
def cria_objetos(formulario, objetos_alunos):
    for linha in formulario.iterrows():
        nome = linha[1][1]
        primeira_op = linha[1][2]
        segunda_op = linha[1][3]
        terceira_op = linha[1][4]
        primeira_opcao_corpo1 = linha[1][5]
        segunda_opcao_corpo1 = linha[1][6]
        terceira_opcao_corpo1 = linha[1][7]
        primeira_opcao_corpo2 = linha[1][8]
        segunda_opcao_corpo2 = linha[1][9]
        terceira_opcao_corpo2 = linha[1][10]
        classificacao = linha[1][11]
        aluno = Aluno(nome,primeira_op,segunda_op,terceira_op,primeira_opcao_corpo1,segunda_opcao_corpo1,terceira_opcao_corpo1,primeira_opcao_corpo2,segunda_opcao_corpo2,terceira_opcao_corpo2, classificacao)
        objetos_alunos.append(aluno)

#Criando uma lista de objetos
formulario = pd.read_excel("respostas.xlsx")
objetos_alunos = []
cria_objetos(formulario,objetos_alunos)

#ordenando os alunos de acordo com suas classificações
objetos_alunos.sort(key = lambda x:x._classificacao)

#Definindo a quantidade de vagas por corpo
aprovados_corpo1 = []
vagas_corpo1 = 124
aprovados_corpo2 = []
vagas_corpo2 = 31
aprovados_corpo3 = []
vagas_corpo3 = 30

excedentes = []

#Adicionando os alunos aos corpos de acordo com suas preferências
for aluno in objetos_alunos:
    #ALOCANDO NA PRIMEIRA OPÇÃO
    if aluno._primeira_op == 'Corpo 1' and len(aprovados_corpo1) < vagas_corpo1:
        aprovados_corpo1.append(aluno)
        continue
    elif aluno._primeira_op == 'Corpo 2' and len(aprovados_corpo2) < vagas_corpo2:
        aprovados_corpo2.append(aluno)
        continue
    elif aluno._primeira_op == 'Corpo 3' and len(aprovados_corpo3) < vagas_corpo3:
        aprovados_corpo3.append(aluno)
        continue

    #ALOCANDO NA SEGUNDA OPÇÃO
    elif aluno._segunda_op == 'Corpo 1' and len(aprovados_corpo1) < vagas_corpo1:
        aprovados_corpo1.append(aluno)
        continue
    elif aluno._segunda_op == 'Corpo 2' and len(aprovados_corpo2) < vagas_corpo2:
        aprovados_corpo2.append(aluno)
        continue
    elif aluno._segunda_op == 'Corpo 3' and len(aprovados_corpo3) < vagas_corpo3:
        aprovados_corpo3.append(aluno)
        continue

    #ALOCANDO NA TERCEIRA OPÇÃO
    elif aluno._terceira_op == 'Corpo 1' and len(aprovados_corpo1) < vagas_corpo1:
        aprovados_corpo1.append(aluno)
        continue
    elif aluno._terceira_op == 'Corpo 2' and len(aprovados_corpo2) < vagas_corpo2:
        aprovados_corpo2.append(aluno)
        continue
    elif aluno._terceira_op == 'Corpo 3' and len(aprovados_corpo3) < vagas_corpo3:
        aprovados_corpo3.append(aluno)
        continue

    else:   
        excedentes.append(aluno)

#Gerando os arquivos com os classificados para cada um dos corpos
classificacao_parcial_corpo1 = open("classificacao_parcial_corpo1.txt", "a")
index = 1
for aluno in aprovados_corpo1:
    classificacao_parcial_corpo1.writelines(f"Aluno:{aluno._nome} - Pos.:{index}\n")
    index += 1
classificacao_parcial_corpo1.close()

classificacao_parcial_corpo2 = open("classificacao_parcial_corpo2.txt", "a")
index = 1
for aluno in aprovados_corpo2:
    classificacao_parcial_corpo2.writelines(f"Aluno:{aluno._nome} - Pos.:{index}\n")
    index += 1
classificacao_parcial_corpo2.close()

#O corpo 3 não possui subdivisões, portanto essa é classificação final dele
classificacao_final_corpo3 = open("classificacao_final_corpo3.txt", "a")
index = 1
for aluno in aprovados_corpo3:
    classificacao_final_corpo3.writelines(f"Aluno:{aluno._nome} - Pos.:{index}\n")
    index += 1
classificacao_final_corpo3.close()

#Defininindo as vagas das divisões do corpo
aprovados_corpo1_divisao1 = []
vagas_divisao1 = 59
aprovados_corpo1_divisao2 = []
vagas_divisao2 = 41
aprovados_corpo1_divisao3 = []
vagas_divisao3 = 24

#Adicionando os alunos as divisões de acordo com suas preferências
for aluno in aprovados_corpo1:
    #alocando primeira opção
    if aluno._primeira_opcao_corpo1 == 'Divisao 1' and len(aprovados_corpo1_divisao1) < 59:
        aprovados_corpo1_divisao1.append(aluno)
    elif aluno._primeira_opcao_corpo1 == 'Divisao 2' and len(aprovados_corpo1_divisao2) < 41:
        aprovados_corpo1_divisao2.append(aluno)
    elif aluno._primeira_opcao_corpo1 == 'Divisao 3' and len(aprovados_corpo1_divisao3) < 24:
        aprovados_corpo1_divisao3.append(aluno)

    #alocando segunda opção
    elif aluno._segunda_opcao_corpo1 == 'Divisao 1' and len(aprovados_corpo1_divisao1) < 59:
        aprovados_corpo1_divisao1.append(aluno)
    elif aluno._segunda_opcao_corpo1 == 'Divisao 2' and len(aprovados_corpo1_divisao2) < 41:
        aprovados_corpo1_divisao2.append(aluno)
    elif aluno._segunda_opcao_corpo1 == 'Divisao 3' and len(aprovados_corpo1_divisao3) < 24:
        aprovados_corpo1_divisao3.append(aluno)

    #alocando terceira opção
    elif aluno._terceira_opcao_corpo1 == 'Divisao 1' and len(aprovados_corpo1_divisao1) < 59:
        aprovados_corpo1_divisao1.append(aluno)
    elif aluno._terceira_opcao_corpo1 == 'Divisao 2' and len(aprovados_corpo1_divisao2) < 41:
        aprovados_corpo1_divisao2.append(aluno)
    elif aluno._terceira_opcao_corpo1 == 'Divisao 3' and len(aprovados_corpo1_divisao3) < 24:
        aprovados_corpo1_divisao3.append(aluno)

#Defininindo as vagas das divisões do corpo
aprovados_corpo2_divisao1 = []
vagas_divisao1 = 18
aprovados_corpo2_divisao2 = []
vagas_divisao2 = 6
aprovados_corpo2_divisao3 = []
vagas_divisao3 = 9

#Adicionando os alunos as divisões de acordo com suas preferências
for aluno in aprovados_corpo2:
    #alocando primeira opção
    if aluno._primeira_opcao_corpo2 == 'Divisao 1' and len(aprovados_corpo2_divisao1) < 59:
        aprovados_corpo2_divisao1.append(aluno)
    elif aluno._primeira_opcao_corpo2 == 'Divisao 2' and len(aprovados_corpo2_divisao2) < 41:
        aprovados_corpo2_divisao2.append(aluno)
    elif aluno._primeira_opcao_corpo2 == 'Divisao 3' and len(aprovados_corpo2_divisao3) < 24:
        aprovados_corpo2_divisao3.append(aluno)

    #alocando segunda opção
    elif aluno._segunda_opcao_corpo2 == 'Divisao 1' and len(aprovados_corpo2_divisao1) < 59:
        aprovados_corpo2_divisao1.append(aluno)
    elif aluno._segunda_opcao_corpo2 == 'Divisao 2' and len(aprovados_corpo2_divisao2) < 41:
        aprovados_corpo2_divisao2.append(aluno)
    elif aluno._segunda_opcao_corpo2 == 'Divisao 3' and len(aprovados_corpo2_divisao3) < 24:
        aprovados_corpo2_divisao3.append(aluno)

    #alocando terceira opção
    elif aluno._terceira_opcao_corpo2 == 'Divisao 1' and len(aprovados_corpo2_divisao1) < 59:
        aprovados_corpo2_divisao1.append(aluno)
    elif aluno._terceira_opcao_corpo2 == 'Divisao 2' and len(aprovados_corpo2_divisao2) < 41:
        aprovados_corpo2_divisao2.append(aluno)
    elif aluno._terceira_opcao_corpo2 == 'Divisao 3' and len(aprovados_corpo2_divisao3) < 24:
        aprovados_corpo2_divisao3.append(aluno)

#Gerando os arquivos com as classificações finais
classificacao_final_corpo1 = open("classificacao_final_corpo1.txt", "a")
index = 1
for aluno in aprovados_corpo1:
    if aluno in aprovados_corpo1_divisao1:
        divisao = "divisao1"
    elif aluno in aprovados_corpo1_divisao2:
        divisao = "divisao2"
    elif aluno in aprovados_corpo1_divisao3:
        divisao = "divisao3"
    else:
        divisao = "Erro na simulação"
    classificacao_final_corpo1.writelines(f"Aluno:{aluno._nome} - Pos no corpo.:{index} - Divis.: {divisao}\n")
    index += 1
classificacao_final_corpo1.close()


classificacao_final_corpo2 = open("classificacao_final_corpo2.txt", "a")
index = 1
for aluno in aprovados_corpo2:
    if aluno in aprovados_corpo2_divisao1:
        divisao = "divisao1"
    elif aluno in aprovados_corpo2_divisao2:
        divisao = "divisao2"
    elif aluno in aprovados_corpo2_divisao3:
        divisao = "divisao3"
    else:
        divisao = "Erro na simulação"
    classificacao_final_corpo2.writelines(f"Aluno:{aluno._nome} - Pos no corpo.:{index} - Divis.: {divisao}\n")
    index += 1
classificacao_final_corpo2.close()




    
