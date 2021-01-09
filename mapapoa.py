from flask import Flask, request, render_template
import idh
import regioes_poa
import consulta_correios
import pycep_correios  # ViaCEP
from collections import Counter
import xlsxwriter
from datetime import datetime
from random import randrange
import csv

RegioesAlunos = []
RegioesAlunos.clear()

outPath = ''
app = Flask(__name__)
app.config["DEBUG"] = True
CorrecaoMonetaria = 1   # 1.8257 --> fator de correção monetária, 82,57% , IGPM- 2010/2020 #  http://drcalc.net/correcao2.asp?descricao=&valor=515%2C67&diainiSelect=1&mesiniSelect=7&anoiniSelect=2010&diafimSelect=1&mesfimSelect=6&anofimSelect=2020&prorata=s&indice=6&juro=0%2C00&periodojuro=m&capitalizacao=c&inicialjuros=&finaljuros=&multa=0%2C00&honorario=0%2C00&Executar2=Executar+o+c%E1lculo&ml=&it=

def RetiraAcentos(string):
    A = [ 'á', 'ã', 'à', 'â']
    E = [ 'é', 'ê' ]
    I = [ 'í', 'î' ]
    O = [ 'ô', 'ó', 'õ' ]
    U = [ 'ú', 'û' ]

    for acento in A:
        string = string.replace(acento, 'a')
    for acento in E:
        string = string.replace(acento, 'e')
    for acento in I:
        string = string.replace(acento, 'i')
    for acento in O:
        string = string.replace(acento, 'o')
    for acento in U:
        string = string.replace(acento, 'u')
    return string

class Aluno:
    def __init__(self, cep, end, bairro, cidade, uf, idx, cartao):
        self.cep = cep
        self.end = end
        self.bairro = bairro
        self.cidade = cidade
        self.regiao = 'Não categorizado'
        self.renda = 0
        self.vulnerabilidade = 0
        Reg=False
        for regiao in regioes_poa.Regioes:  # bairro de poa
            for bar in regiao.bairros:
                if RetiraAcentos(self.bairro.lower() ) == bar:
                    self.regiao = regiao.nome
                    RegioesAlunos.append(self.regiao)
                    Reg=True
        if Reg==False:
            for bar in regiao.bairros:  # busca região metropolitana pela cidade
                if RetiraAcentos(self.cidade.lower() ) == bar:
                    self.regiao = regiao.nome
                    RegioesAlunos.append(self.regiao)
                    Reg=True
        if Reg==False:  # outrar região
            RegioesAlunos.append('Não categorizado')
        self.uf = uf
        self.idh = '0'
        try:
            for bar in idh.BairrosPOA:
                if RetiraAcentos(self.bairro.lower() ) == bar.nome:
                    self.idh = str(bar.idh)
                    self.renda = bar.renda * CorrecaoMonetaria
                    self.vulnerabilidade = bar.vulnerabilidade
        except:
            pass
        if self.renda == 0:  # se não encontra dados do bairro, busca dados da cidade (região metropolitana)
            try:
                for bar in idh.BairrosPOA:
                    if RetiraAcentos(self.cidade.lower() ) == bar.nome:
                        self.idh = str(bar.idh)
                        self.renda = bar.renda * CorrecaoMonetaria
                        self.vulnerabilidade = bar.vulnerabilidade
            except:
                pass
        self.lat = 0
        self.long = 0
        self.indice = idx
        self.cartao = cartao # cartão UFRGS

def getIDH(obj):
    return float(obj.idh.replace(',','.'))

def getRenda(obj):
    return float(obj.renda)

def getIDS(obj):
    return float(obj.vulnerabilidade)

def process_data(input_file, formato):
    #Listas
    Bairro = []
    CEPS = []
    IDENTIFICACAO = []
    Alunos = []
    ListaIDH = []
    # UFRGS = (-30.033440, -51.218720) #  Reitoria UFRGS

    i=0
    j=0

    print(formato)
    input_file.save('input/temp.input')

    if 'xlsx' in formato:
        print('xlsx')
    elif 'csx' in formato:
        print('csx')

    #chatice: lidar com encodings bizarros do excel
    try:
        with open('input/temp.input', encoding='utf-8') as f:
            old = f.read()
            old = old.replace(";",",")
            with open('input/temp.input2', 'w', encoding='utf-8') as f2:
                f2.write(old)
    except:
        with open('input/temp.input', encoding='cp1252') as f:
            old = f.read()
            old = old.replace(";",",")
            with open('input/temp.input2', 'w', encoding='utf-8') as f2:
                f2.write(old)

    # Lê arquivo CSV ou XLSX
    with open('input/temp.input2', newline='', encoding='utf-8') as csvfile:
        i=0
        spamreader = csv.reader(csvfile, delimiter=',', quotechar='|')
        for row in spamreader:
            print(row)
            try:
                i=row[0]
                c=row[1]
                c=c.replace("-","")
                if c.isdigit():
                    CEPS.append(c)
                    IDENTIFICACAO.append(i)
                    i+1
            except:
                pass

    #input_data = input_file.stream.read()
    #input_data = convert_encoding(input_data)

    '''
    input_data = "eu,90035-030\n"
    print(input_data)
    linhas = input_data.splitlines()

    for linha in linhas:
        print(linha)
        # ToDo: flexibilizar essa rotina para permitir maior variedade de formatos
        if ',' in linha:
            c = linha.split(',')[1]
            c = c.replace("-","")
            if c.isdigit(): # cep válido
                try:
                    CEPS.append(linha.split(',')[1])
                    IDENTIFICACAO.append(linha.split(',')[0])
                except:
                    pass
        elif ';' in linha:
            c = linha.split(';')[1]
            c = c.replace("-","")
            if c.isdigit(): # cep válido
                try:
                    CEPS.append(linha.split(';')[1])
                    IDENTIFICACAO.append(linha.split(';')[0])
                except:
                    pass
    '''

    for CEP in CEPS:
        print("Consultando CEP: " + str(CEP))
        try:
            x2 = consulta_correios.busca_cep(str(CEP).replace("-","") )
            kValida=False
            for key in x2:
                if key != 'error':
                    citystate = key['city/state'].split('/')
                    cidade = citystate[0]
                    uf = citystate[1]
                    Alunos.append( Aluno(key['zipcode'], key['address'],key['neighborhood'],cidade,uf, i, IDENTIFICACAO[j]) )
                    Bairro.append( key['neighborhood'] )
                    i+=1
                    kValida=True
                elif key== 'error' and kValida==False:
                    print("Erro, tentando consulta alternativa:" + str(CEP) )
                    x = pycep_correios.consultar_cep(CEP)
                    Alunos.append( Aluno(x['cep'], x['end'],x['bairro'],x['cidade'],x['uf'], i, IDENTIFICACAO[j]) )
                    Bairro.append( x['bairro'] )
                    i+=1
        except:
            print("erro")
        j+=1

    Bairros = Counter(Bairro)   # conta ocorrências de bairros
    now = datetime.now()
    dt = now.strftime("%d-%m-%Y-%H_%M_%S") + str(randrange(1000))
    global outPath
    outPath = '/output/alunos_excel-' + dt + '.xlsx'
    workbook = xlsxwriter.Workbook('./output/alunos_excel-' + dt + '.xlsx')

    # Ordena abas
    wIDH = workbook.add_worksheet('IDH')
    wRENDA = workbook.add_worksheet('Renda')
    wIDS = workbook.add_worksheet('Vulnerabilidade')
    wGRAFICO = workbook.add_worksheet('Gráficos')
    wGRUPOS2 = workbook.add_worksheet('Grupos(2)')
    wGRUPOS3 = workbook.add_worksheet('Grupos(3)')
    wGRUPOS4 = workbook.add_worksheet('Grupos(4)')

    # Estilo
    bold = workbook.add_format({'bold': True})
    vermelho = workbook.add_format({'bold': True, 'font_color': 'red'})

    # IDH da turma
    IDHContagem = 0
    RendaContagem =0
    IDHTotal = 0
    RendaTotal = 0
    alunoCount = 0
    ListaRenda = []

    #print(Alunos)
    #print(Bairro)

    for aluno in Alunos:
        #print(aluno)
        alunoCount += 1
        IDHTotal += float(aluno.idh.replace(',','.'))
        RendaTotal += aluno.renda
        ListaIDH.append(float(aluno.idh.replace(',','.')))

        if (aluno.renda != 0):
            ListaRenda.append(aluno.renda)

        if float(aluno.idh.replace(',','.')) != 0:
            IDHContagem += 1  # ignora ausência de dados na média
        if aluno.renda != 0:
            RendaContagem += 1  # ignora ausência de dados na média

    print(alunoCount)
    IDHMedio = IDHTotal / IDHContagem
    RendaMedia = RendaTotal / IDHContagem

    def gini(list_of_values):
        sorted_list = sorted(list_of_values)
        height, area = 0, 0
        for value in sorted_list:
            height += value
            area += height - value / 2.
        fair_area = height * len(list_of_values) / 2.
        return (fair_area - area) / fair_area

    RendaGini = gini(ListaRenda)

    # Organiza listas
    AlunosIDH = sorted(Alunos, key=getIDH, reverse=False)
    AlunosRenda = sorted(Alunos, key=getRenda, reverse=False)
    AlunosIDS = sorted(Alunos, key=getIDS, reverse=True)

    #Gera worksheets
    def WorksheetAlunos(worksheet, lista):
        worksheet.set_column('A:A', 10)
        worksheet.set_column('B:B', 26)
        worksheet.set_column('D:C', 16)
        worksheet.set_column('E:E', 6)
        worksheet.set_column('F:F', 9)
        worksheet.set_column('I:I', 8)
        worksheet.set_column('J:J', 24)

        headings = ['ID', 'Endereço', 'Bairro', 'Cidade', 'Estado', 'CEP', 'IDH' , 'Renda', 'Vuln', 'Região' ]
        worksheet.write_row('A1', headings, bold)

        linha=2 # novo índice (linha)
        for aluno in lista:
            worksheet.write( 'A' + str(linha), int(aluno.cartao) )
            worksheet.write( 'B' + str(linha), str(aluno.end) )
            worksheet.write( 'C' + str(linha), str(aluno.bairro) )
            worksheet.write( 'D' + str(linha), str(aluno.cidade) )
            worksheet.write( 'E' + str(linha), str(aluno.uf) )
            worksheet.write( 'F' + str(linha), str(aluno.cep) )

            if float(aluno.idh) < 0.76 :
                worksheet.write( 'G' + str(linha), float(aluno.idh.replace(',','.')), vermelho  )
            else:
                worksheet.write( 'G' + str(linha), float(aluno.idh.replace(',','.')) )
            #wIDH.write( 'G' + str(linha), int(Bairros[aluno.bairro]) ) # total por bairro
            worksheet.write( 'H' + str(linha), str(round(aluno.renda,2)) )

            if float(aluno.vulnerabilidade) > 0.24:
                worksheet.write( 'I' + str(linha), str(aluno.vulnerabilidade), vermelho )
            else:
                worksheet.write( 'I' + str(linha), str(aluno.vulnerabilidade) )
            worksheet.write( 'J' + str(linha), str(aluno.regiao) )
            linha+=1

    # Gera worksheets
    WorksheetAlunos(wIDH, AlunosIDH)
    WorksheetAlunos(wRENDA, AlunosRenda)
    WorksheetAlunos(wIDS, AlunosIDS)

    # Grupos - IDH
    melhorIDH = sorted(Alunos, key=getIDH, reverse=True) # matriz referência
    gIDH = []
    gIDH3 = []
    gIDH4 = []

    par=False
    if alunoCount % 2 == 0:
        par= True
    mid=int(alunoCount/2)
    um_terco=int(alunoCount/3)
    um_quarto=int(alunoCount/4)
    sobra3=alunoCount%3  #módulo 3
    sobra4=alunoCount%4  #módulo 3

    if par==True:
        for i in range(0,mid):
            gIDH.append(melhorIDH[i])
            gIDH.append(melhorIDH[mid+i])

        for i in range(0,um_quarto):
            a = i
            b = um_quarto*1 + i
            c = um_quarto*2 + i
            d = um_quarto*3 + i
            gIDH4.append(melhorIDH[a])
            gIDH4.append(melhorIDH[b])
            gIDH4.append(melhorIDH[c])
            gIDH4.append(melhorIDH[d])

        for i in range(0,sobra4):
            gIDH4.append(melhorIDH[um_quarto*4+i])

    else:
        for i in range(0,mid):
            gIDH.append(melhorIDH[i])
            gIDH.append(melhorIDH[mid+1+i])
        # agrupa o aluno restante( meio da tabela) com o último aluno
        gIDH.append(melhorIDH[mid])

        for i in range(0,um_quarto):
            a = i
            b = um_quarto*1 + i
            c = um_quarto*2 + i
            d = um_quarto*3 + i
            gIDH4.append(melhorIDH[a])
            gIDH4.append(melhorIDH[b])
            gIDH4.append(melhorIDH[c])
            gIDH4.append(melhorIDH[d])

        for i in range(0,sobra4):
            gIDH4.append(melhorIDH[(um_quarto*4)+i])

    for i in range(0,um_terco):
        a = i
        b = um_terco*1 + i
        c = um_terco*2 + i
        gIDH3.append(melhorIDH[a])
        gIDH3.append(melhorIDH[b])
        gIDH3.append(melhorIDH[c])

    for i in range(0,sobra3):
        gIDH3.append(melhorIDH[(um_terco*3)+i])

    def worksheetGrupo(worksheet):
        worksheet.set_column('A:A', 6)
        worksheet.set_column('B:B', 12)
        worksheet.set_column('C:C', 28)
        worksheet.set_column('D:D', 16)
        worksheet.set_column('E:E', 16)
        worksheet.set_column('F:F', 6)
        worksheet.set_column('G:G', 10)
        worksheet.set_column('H:H', 10)
        worksheet.set_column('I:I', 28)
                    #  a       B     C             D        E        F       G             H            I
        headings = ['Grupo', 'ID', 'Endereço', 'Bairro', 'Cidade', 'IDH' , 'IDH Grupo', 'IDH Turma', 'Região' ]
        worksheet.write_row('A1', headings, bold)

    worksheetGrupo(wGRUPOS2)  # grupo2
    worksheetGrupo(wGRUPOS3)  # grupo3
    worksheetGrupo(wGRUPOS4)  # grupo4

    # Grupos (2)
    linha=2
    grupo=0
    for aluno in gIDH:
        if (linha%2==0): # par
            grupo+=1
            estilo = workbook.add_format({'italic': True, 'bold': True})
            if (grupo%2==0):
                estilo.set_bg_color('#D3D3D3')

            wGRUPOS2.write( 'A' + str(linha), grupo, estilo)
            wGRUPOS2.write( 'G' + str(linha), '=AVERAGE(F' + str(linha) + ':F' + str(linha+1) + ')', estilo)
            if linha-1 != alunoCount:  # último aluno
                wGRUPOS2.write( 'G' + str(linha+1), '', estilo)

            if linha-1 == alunoCount:  # último aluno não forma grupo
                wGRUPOS2.write( 'A' + str(linha), '', estilo)
        else: # impar
            estilo = workbook.add_format({'italic': False})
            if (grupo%2==0):
                estilo.set_bg_color('#D3D3D3')
            wGRUPOS2.write( 'A' + str(linha), '', estilo)

        wGRUPOS2.write( 'B' + str(linha), int(aluno.cartao), estilo)
        wGRUPOS2.write( 'C' + str(linha), str(aluno.end), estilo)
        wGRUPOS2.write( 'D' + str(linha), str(aluno.bairro), estilo)
        wGRUPOS2.write( 'I' + str(linha), str(aluno.regiao), estilo)
        wGRUPOS2.write( 'E' + str(linha), str(aluno.cidade), estilo)

        if float(aluno.idh.replace(',','.')) < 30:
            wGRUPOS2.write( 'F' + str(linha), float(aluno.idh.replace(',','.')), estilo)
        else:
            wGRUPOS2.write( 'F' + str(linha), float(aluno.idh.replace(',','.')), estilo )
        wGRUPOS2.write( 'H' + str(linha), float(round(IDHMedio,3)), estilo )
        linha+=1

    # Grupos (3)
    linha=3
    grupo=0
    for aluno in gIDH3:
        if (linha%3==0): # novo grupo
            grupo+=1
            estilo = workbook.add_format({'italic': True, 'bold': True})
            if (grupo%2==0):
                estilo.set_bg_color('#D3D3D3')

            wGRUPOS3.write( 'A' + str(linha-1), grupo, estilo)
            wGRUPOS3.write( 'G' + str(linha-1), '=AVERAGE(F' + str(linha-1) + ':F' + str(linha+1) + ')', estilo)
            if linha-2 != alunoCount:  # último aluno
                wGRUPOS3.write( 'G' + str(linha), '', estilo)

            if linha-2 == alunoCount:  # último aluno não forma grupo
                wGRUPOS3.write( 'A' + str(linha), '', estilo)
        else: # impar
            estilo = workbook.add_format({'italic': False})
            if (grupo%2==0):
                estilo.set_bg_color('#D3D3D3')
            wGRUPOS3.write( 'A' + str(linha-1), '', estilo)

            if linha-1 != alunoCount:  # último aluno
                wGRUPOS3.write( 'G' + str(linha-1), '', estilo)

        wGRUPOS3.write( 'B' + str(linha-1), int(aluno.cartao), estilo)
        wGRUPOS3.write( 'C' + str(linha-1), str(aluno.end), estilo)
        wGRUPOS3.write( 'D' + str(linha-1), str(aluno.bairro), estilo)
        wGRUPOS3.write( 'I' + str(linha-1), str(aluno.regiao), estilo)
        wGRUPOS3.write( 'E' + str(linha-1), str(aluno.cidade), estilo)

        if float(aluno.idh.replace(',','.')) < 30:
            wGRUPOS3.write( 'F' + str(linha-1), float(aluno.idh.replace(',','.')), estilo)
        else:
            wGRUPOS3.write( 'F' + str(linha-1), float(aluno.idh.replace(',','.')), estilo )
        wGRUPOS3.write( 'H' + str(linha-1), float(round(IDHMedio,3)), estilo )
        linha+=1

    # Grupos (4)
    linha=4
    grupo=0
    for aluno in gIDH4:
        if (linha%4==0):
            grupo+=1
            estilo = workbook.add_format({'italic': True, 'bold': True})
            if (grupo%2==0):
                estilo.set_bg_color('#D3D3D3')

            wGRUPOS4.write( 'A' + str(linha-2), grupo, estilo)
            wGRUPOS4.write( 'G' + str(linha-2), '=AVERAGE(F' + str(linha-2) + ':F' + str(linha+1) + ')', estilo)
            if linha-1 != alunoCount:  # último aluno
                wGRUPOS4.write( 'G' + str(linha-1), '', estilo)

            if linha-1 == alunoCount:  # último aluno não forma grupo
                wGRUPOS4.write( 'A' + str(linha-1), '', estilo)
        else: # impar
            estilo = workbook.add_format({'italic': False})
            if (grupo%2==0):
                estilo.set_bg_color('#D3D3D3')
            wGRUPOS4.write( 'A' + str(linha-2), '', estilo)

            if linha-1 != alunoCount:  # último aluno
                wGRUPOS4.write( 'G' + str(linha-1), '', estilo)


        wGRUPOS4.write( 'B' + str(linha-2), int(aluno.cartao), estilo)
        wGRUPOS4.write( 'C' + str(linha-2), str(aluno.end), estilo)
        wGRUPOS4.write( 'D' + str(linha-2), str(aluno.bairro), estilo)
        wGRUPOS4.write( 'I' + str(linha-2), str(aluno.regiao), estilo)
        wGRUPOS4.write( 'E' + str(linha-2), str(aluno.cidade), estilo)

        if float(aluno.idh.replace(',','.')) < 30:
            wGRUPOS4.write( 'F' + str(linha-2), float(aluno.idh.replace(',','.')), estilo)
        else:
            wGRUPOS4.write( 'F' + str(linha-2), float(aluno.idh.replace(',','.')), estilo )
        wGRUPOS4.write( 'H' + str(linha-2), float(round(IDHMedio,3)), estilo )
        linha+=1

    #Gráfico - IDH
    headings = ['Aluno', 'IDH', 'Faixa', 'Freq', 'Bairro', 'qtde', 'Região' , 'qtde', 'Renda', 'faixa', 'Freq' ]
    ContagemAlunos = []
    for i in range(0,alunoCount):
        ContagemAlunos.append(i)

    ListaIDHint=[]
    for i in ListaIDH:
        if i > 0:
            ListaIDHint.append(float(i))

    data = [ ContagemAlunos,
            ListaIDHint,
            [ 0.50,0.55,0.60,0.65,0.70,0.75,0.80,0.85,0.90,0.95,1],
            ListaRenda,
            [ 400, 800, 1200, 1600, 2000, 2400, 2800, 3200, 3600, 4000, 4400, 4800 ],
            ]

    wGRAFICO.write_row('A1', headings, bold)
    wGRAFICO.write_column('A2', data[0])
    wGRAFICO.write_column('B2', data[1])
    wGRAFICO.write_column('C2', data[2])
    wGRAFICO.write_column('I2', data[3])
    wGRAFICO.write_column('J2', data[4])
    linha=2

    for i in range(0,10):
        wGRAFICO.write('D' + str(i+2), '=COUNTIFS(B2:B' + str(alunoCount+1) + ',">=' + str(0.5+i*0.05).replace('.',',') + '",$B2:$B' + str(alunoCount+1) + ',"<=' + str( (i*0.05)+0.55).replace('.',',') + '")' )

    for i in range(0,13):
        wGRAFICO.write('K' + str(i+2), '=COUNTIFS(I2:I' + str(alunoCount+1) + ',">=' + str(400+i*400).replace('.',',') + '",$I2:$I' + str(alunoCount+1) + ',"<=' + str( (400+i*400)+400).replace('.',',') + '")' )

    #GRÁFICO - IDH
    chart1 = workbook.add_chart({'type': 'column'})
    chart1.add_series({ 'name': '=Gráficos!$D$1',
                        'categories': '=Gráficos!$C$2:$C$11',
                        'values': '=Gráficos!$D$2:$D$11' })
    chart1.set_title ({'name': 'IDH: distribuição (Média=' + str(round(IDHMedio,3)) + ', POA = 0.805) ' })
    chart1.set_style(11)
    chart1.set_size({'x_scale': 1.4, 'y_scale': 1.4})
    wGRAFICO.insert_chart('A1', chart1, {'x_offset': 5, 'y_offset': 5})

    #GRÁFICO - RENDA
    chart1 = workbook.add_chart({'type': 'column'})
    chart1.add_series({ 'name': '=Gráficos!$K$1',
                        'categories': '=Gráficos!$J$2:$J$14',
                        'values': '=Gráficos!$K$2:$K$14' })
    chart1.set_title ({'name': 'Renda: distribuição (Média=' + str(round(RendaMedia,2)) + ', POA = 1600, Índice de Gini = ' + str(round(RendaGini,3)) + ')' })
    chart1.set_style(11)
    chart1.set_size({'x_scale': 1.4, 'y_scale': 1.4})
    wGRAFICO.insert_chart('L1', chart1, {'x_offset': 5, 'y_offset': 5})

    #Bairros
    chart1 = workbook.add_chart({'type': 'pie'})

    i=0
    for k in Bairros.keys():
        wGRAFICO.write('E' + str(2+i), k)
        wGRAFICO.write('F' + str(2+i), Bairros[k])
        i+=1

    chart1.add_series({
        'name':       'Bairros',
        'categories': 'Gráficos!$E$2:$E$' + str(2+len(Bairros)-1),
        'values':     'Gráficos!$F$2:$F$' + str(2+len(Bairros)-1),
        'data_labels': {'value': True},
    })
    chart1.set_title({'name': 'Bairros'})
    chart1.set_style(10)
    chart1.set_size({'x_scale': 2, 'y_scale': 1.8})
    wGRAFICO.insert_chart('I22', chart1, {'x_offset': 5, 'y_offset': 5})

    #Regioes
    chart1 = workbook.add_chart({'type': 'pie'})
    RegioesLista = Counter(RegioesAlunos)   # conta ocorrências de bairros

    i=0
    for k in RegioesLista.keys():
        wGRAFICO.write('G' + str(2+i), k)
        wGRAFICO.write('H' + str(2+i), RegioesLista[k])
        i+=1

    chart1.add_series({
        'name':       'Regiões',
        'categories': 'Gráficos!$G$2:$G$' + str(2+len(RegioesLista)-1 ),
        'values':     'Gráficos!$H$2:$H$' + str(2+len(RegioesLista)-1 ),
        'data_labels': {'value': True},
    })

    chart1.set_title({'name': 'Regiões'})
    chart1.set_style(10)
    chart1.set_size({'x_scale': 1.4, 'y_scale': 1.4})
    wGRAFICO.insert_chart('A22', chart1, {'x_offset': 25, 'y_offset': 10})

    # escreve arquivo
    workbook.close()
    with open('./output/alunos_excel-' + dt + '.xlsx', 'rb') as xls:
        result=xls.read()
        return result


@app.route("/", methods=["GET", "POST"])
def file_summer_page():
    if request.method == "POST":
        input_file = request.files["input_file"]
        formato = request.files["input_file"].name
        process_data(input_file, formato)
        return render_template('resultados.html', outPath=outPath)
    return render_template('index.html')