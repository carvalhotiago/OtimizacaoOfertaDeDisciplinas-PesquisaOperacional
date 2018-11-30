from gurobipy import *
import xlrd
import numpy as np
import anexo
import datetime

model = Model('projeto')

workbook = xlrd.open_workbook(r'C:\Users\Tiago Carvalho\Documents\Faculdade\PO\ProjetoVagas\PedidosVagasSI_20143.xlsx')
pedidos = workbook.sheet_by_index(0)
grade = workbook.sheet_by_index(2)

d = 38
a = 0
h = 5

for i in range(0,114): #Conta o numero de alunos
    if (pedidos.cell_type(i, 0) == 2):
        a += 1

disciplinas = []
for i in range(1, 39):
    disciplinas.append(grade.cell_value(i, 0))
        

peso = np.matlib.zeros((a, d))
#Atrivui pesos maiores para disciplinas dos primeiros periodos
for i in range(6,108):
    pesoinicial = 10
    for j in range(1,39):
        if (pedidos.cell_value(i, j) == 'H' or pedidos.cell_value(i, j) == 'R'):
            if(grade.cell_value(j, 0)[0] == 'D'):
                peso[i-6, j-1] = pesoinicial
                if(pesoinicial > 1):        #peso nunca pode ser zero
                    pesoinicial = pesoinicial - 1
print(peso)
                

x = model.addVars(a, d, h, vtype=GRB.BINARY, name='Xadh')
y = model.addVars(d, h, vtype=GRB.BINARY, name='Ydh')


FO = 0
for k in range(h):
    for j in range (d):
        for i in range (a):
            FO+= x[i, j, k]*peso[i, j]

model.setObjective(FO, GRB.MAXIMIZE)

model.addConstrs((x.sum(i, '*', k) <= 1 for i in range(a) for k in range(h)), 'alunoMaxIDiscH')
model.addConstrs((y.sum(j, '*') <= 1 for j in range(d)), 'DisciplinaIH')
model.addConstrs((x[i, j, k] <= y[j, k] for i in range(a) for j in range(d) for k in range(h)), 'alunoDisciplinaOferecida')
model.addConstr(y.sum('', '') <= 20, 'MaxDisciplina')

print(datetime.datetime.now())

model.write("model.lp")
model.optimize()
data = datetime.datetime.now()
tempo = model.Runtime
count = 0
if model.status == GRB.Status.OPTIMAL:
    for c in model.getVars():
        if(c.varName[0] == 'Y'):
            if(c.x == 1.0):
                count += 1
                print(c.varName, c.x)          
print(count)
#CORREÇÃO MATRIZ
ValorXadh = []
for c in model.getVars():
    if(c.varName[0] == 'X'):
        ValorXadh.append(c.x)

c = 0    
for i in range(a):
    for j in range(d):
        for k in range(h):
            x[i, j, k] = ValorXadh[c]
            c += 1  
            
ValorYdh = []
for c in model.getVars():
    if(c.varName[0] == 'Y'):
        ValorYdh.append(c.x)
c = 0    
for j in range(d):
    for k in range(h):
        y[j, k] = ValorYdh[c]
        c += 1

#LISTAGEM ALUNO - DISCIPLINA - HORARIO
for i in range(a):
    print('Aluno', i)
    for k in range(h):
        for j in range(d):
            if(x[i, j, k] == 1.0):
                print('\t {} -- H{}'.format(Disciplinas_CodigoNome[disciplinas[j]], k))
                
with open('aluno-disciplina-horario.csv', 'w') as new_file:
    csv_writer = csv.writer(new_file, delimiter='\t')
    csv_writer.writerow(('Aluno', '   Disciplina', '    Horario'))
    for i in range(a):
        for k in range(h):
            for j in range(d):
                if(x[i, j, k] == 1.0):
                    csv_writer.writerow((i, Disciplinas_CodigoNome[disciplinas[j]], "h"+str(k)))

#LISTAGEM DICSPLINA - HORARIO
for j in range(d):
    for k in range(h):
        if(y[j, k] == 1.0):
            print('{} --- h{}'.format(disciplinas[j], k))
            
with open('disciplina-horario.csv', 'w') as new_file:
    csv_writer = csv.writer(new_file, delimiter='\t')
    csv_writer.writerow(('Valor otimo: ', model.ObjVal))
    csv_writer.writerow(('Data', '        Hora', '                      Tempo de Execucao (s)'))
    csv_writer.writerow((data, tempo))
    csv_writer.writerow(('Disciplina', 'Horario'))
    for j in range(d):
        for k in range(h):
            if(y[j, k] == 1.0):
                csv_writer.writerow((Disciplinas_CodigoNome[disciplinas[j]], '  ','h'+str(k)))            
            

contador = []
ct = 0
for i in range(a):
    for k in range(h):
        for j in range(d):
            if(x[i, j, k] == 1.0):
                ct += 1
            if(k == h-1 and j == d-1):
                contador.append(ct)
                ct = 0
print("Total de alunos: ", len(contador))

alunoCom5Disciplinas = 0
alunoCom4Disciplinas = 0
for i in range(a):
    if(contador[i] == 5):
        alunoCom5Disciplinas += 1
    elif(contador[i] == 4):
        alunoCom4Disciplinas += 1
print("Alunos matriculados em 5 disciplinas: {}".format(alunoCom5Disciplinas))
print("Alunos matriculados em 4 disciplinas: {}".format(alunoCom4Disciplinas))
