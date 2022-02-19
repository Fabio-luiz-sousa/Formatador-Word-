from turtle import title
from funcoes_capa.funcoes_capa import*

print(80*'#')
print('\t\t\tFormatador de Documentos .docx')
print(80*'#')


op = input('[1]-Capa\nEscolha sua opção: ')
t_2 = []
if op == '1':
    t1 = input('Digite a instituição que vc estuda: ')
    t1 = t1.upper()
    titulo(t1)
    t4 = input('Digite o nome do seu curso: ')
    t4 = t4.upper()
    tipo_curso(t4)
    quant = input('Quanntos autores vc quer adicionar: ')
    quant = int(quant)
    while quant:
        t2 = input('Digite o nome do(s)(as) integrante(s): ')
        t2 = t2.title()
        t_2.append(t2)
        quant -= 1
        if quant == 0:
            break
    nome_autores(t_2)

    t3 = input('Dfigite o titulo do trabalho: ')
    t3 = t3.upper()
    titulo_trabalho(t3)
    t5 = input('Digite o subitulo do trabalho: ')
    t5 = t5.title()
    subtitulo_trabalho(t5)
    t6 = input('Digite a cidade: ')
    t6 = t6.upper()
    cidade(t6)
    t7 = input('Digite o ano: ')
    ano(t7)
    ori = input('Digite o nome do orientador: ')
    ori = ori.title()
    tsave = input('digite o nome do documento para ser salvo: ')
    t4=t4.title()
    t1=t1.title()

    nome_autores_contracapa(t_2)
    titulo_trabalho_contracapa(t3)
    subtitulo_trabalho_contracapa(t5)
    descricao_contracapa(t4, t1, ori)  
    ano_contracapa(t7)

    salvar_doc(tsave)
