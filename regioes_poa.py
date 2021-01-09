# http://www.observapoa.com.br/default.php?reg=1&p_secao=46
class Regiao:
    def __init__(self, numero, nome, bairros, idh):
        
        self.numero = numero
        self.nome = nome
        self.bairros = bairros
        self.idh = idh

Regioes = []
bairros = []

bairros = [ 'anchieta', 'arquipelago', 'farrapos', 'humaita', 'navegantes', 'sao geraldo' ]
Regioes.append( Regiao(1, "Ilhas e Humaita/Navegantes", bairros, 0) )

bairros = [ 'boa vista', 'cristo redentor', 'higienopolis', 'jardim floresta', 'jardim itu', 'jardim lindoia', 'jardim sao pedro', "passo d'areia", 'sarandi', 'santa maria goretti', 'sao joao', 'sao sebastiao', 'vila ipiranga'  ]
Regioes.append( Regiao(2, "Norte e Nordeste", bairros, 0) )

bairros = [ 'bom jesus', 'chacara das pedras', 'jardim carvalho', 'jardim sabara', 'jardim do salso', 'morro santana', 'tres figueiras', 'vila jardim' ]
Regioes.append( Regiao(3, "Leste", bairros, 0) )

bairros = [ 'coronel aparicio borges', 'partenon', 'santo antonio', 'sao jose', 'vila joao pessoa' ]
Regioes.append( Regiao(4, "Paternon", bairros, 0) )

bairros = [ 'belem velho', 'cascata', 'cristal', 'gloria', 'medianeira', 'santa teresa', 'santa tereza']
Regioes.append( Regiao(5, "Gloria, Cruzeiro e Cristal", bairros, 0) )

bairros = [ 'camaqua', 'campo novo', 'cavalhada', 'espirito santo', 'guaruja', 'hipica', 'ipanema', 'jardim isabel', 'nonoai', 'serraria', 'teresopolis', 'tristeza', 'vila assunçao', 'vila conceiçao', 'vila nova']
Regioes.append( Regiao(6, "Sul e Centro-sul", bairros, 0) )

bairros = [ 'belem novo', 'chapeu do sol', 'lageado', 'lami', 'ponta grossa', 'restinga' ]
Regioes.append( Regiao(7, "Restinga e Extremo Sul", bairros, 0) )

bairros = [ 'auxiliadora', 'azenha', 'bela vista', 'bom fim', 'bonfim', 'centro', 'centro historico', 'cidade baixa', 'farroupilha', 'floresta', 'independencia', 'jardim botânico', 'menino deus', 'moinhos de vento', 'mont serrat', 'petropolis', 'praia de belas', 'rio branco', 'santa cecilia', 'santana' ]
Regioes.append( Regiao(8, "Centro", bairros, 0) )

bairros = [ 'agronomia', 'lomba do pinheiro' ]
Regioes.append( Regiao(9, "Lomba do Pinheiro", bairros, 0) )

bairros = [ 'mario quintana', 'passo das pedras', 'rubem berta' ]
Regioes.append( Regiao(10, "Eixo Baltazar e Nordeste", bairros, 0) )

cidades = [ 'alvorada', 'ararica', 'arroio dos ratos', 'cachoeirinha', 'campo bom', 'canoas', 'capela de santana', 'charqueadas', 'dois irmaos', 'eldorado do sul', 'estancia velha', 'esteio', 'glorinha', 'gravatai', 'guaiba', 'igrejinha', 'ivoti', 'montenegro', 'nova hartz', 'nova santa rita', 'novo hamburgo', 'parobe', 'portao', 'rolante', 'santo antonio da patrulha', 'sao jeronimo', 'sao leopoldo', 'sao sebastiao do cai', 'sapiranga', 'sapucaia do sul', 'taquara', 'triunfo', 'viamao'  ]
Regioes.append( Regiao(11, "Região Metropolitana", cidades, 0) )