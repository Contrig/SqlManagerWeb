<%
'------------------------------------------------------------'
'### VARIAVEIS
'------------------------------------------------------------'		
'### Rodapé
'------------------------------------------------------------'		
	SHD_Version		=	"1.0"
	SHD_DtRelease	=	"31.01.2011"
'------------------------------------------------------------'
'### ENVIO DE EMAILS 
'------------------------------------------------------------'
	'Componente	= "AspMail"
	'Componente	= "AspEmail"
	'Componente	= "AspQmail"
	Componente	= "CDONTS"
	'Componente	= "JMail"
	'Componente	= "CdoSys" '### W2K 2003
'------------------------------------------------------------'		
	eMailGeral	= "atendimento@"& Replace(Request.ServerVariables("server_name"),"www.","")
	nomeGeral	= "Automático"
'------------------------------------------------------------'		
Host 		= "localhost" '## SMTP DO SERVER
strHost		= Lower(Replace(Request.ServerVariables("server_name"),"www.",""))
AssuntoGeral	= "Newsletter " '## O sistema complementa com a data
'------------------------------------------------------------'		
'** Vetor de prioridade
'------------------------------------------------------------'		
Redim	Prioridade(3,1)
		Prioridade(0,0)	=	"imagens/icon/priority_low.png"
		Prioridade(0,1)	=	"Baixa"
		Prioridade(1,0)	=	"imagens/icon/priority_normal.png"
		Prioridade(1,1)	=	"Normal"
		Prioridade(2,0)	=	"imagens/icon/priority_high.png"
		Prioridade(2,1)	=	"Alta"
		Prioridade(3,0)	=	"imagens/icon/priority_urgente.png"
		Prioridade(3,1)	=	"Urgente"
'------------------------------------------------------------'		
'### VETOR DO CADASTRO
'------------------------------------------------------------'		
MeuVetorCadastro = "Rua,Avenida,Alameda,Praça,Quadra,Residencial,Rodovia,Setor,Travessa,Viaduto"
MeuVetorCadastro = Split(MeuVetorCadastro,",")
MeuVetorCadastro = ordena(MeuVetorCadastro,0)	
'------------------------------------------------------------'
MeuVetorEstados = "AC,AL,AM,AP,BA,CE,DF,ES,GO,MA,MG,MS,MT,PA,PB,PE,PI,PR,RJ,RN,RO,RR,RS,SC,SE,SP,TO"
MeuVetorEstados = Split(MeuVetorEstados,",")
MeuVetorEstados = ordena(MeuVetorEstados,0)	
'------------------------------------------------------------'
'### EFEITO DO MENU 
'------------------------------------------------------------'
strMenu = "onMouseOver=""foco('busca',this,'#EFEFEF','#9F9F9F')"" onMouseOut=""foco('busca',this,'#ADAEAD','#E7E7DE')"" style='cursor:hand'"
'------------------------------------------------------------'
'### A
	Acao			= Pegar("Acao",0)
	Agora			= Now()
	Assunto			= Pegar("Assunto",0)
	adm				= Pegar("adm",0)		: if adm = "" then adm = 0
	atendente		= Pegar("atendente",0)	: if atendente = "" then atendente = 0
	
'### B
	bairro		= Pegar("bairro",0)
	bway		= Pegar("bway",0)
	bilheteria	= Pegar("bilheteria",0)
	bd			= Pegar("bd",0)
	busca		= Pegar("busca",0)
	boxImg		= Pegar("imagem",0)
	
'### C
	canal		= Pegar("canal",0)
	cor			= "sim"
	Cont		= 0
	cpf			= Pegar("camponumero",0)
	cep			= Pegar("cep",0)
	cidade		= Pegar("cidade",0)
	complemento	= Pegar("complemento",0)
	como		= Pegar("como",0)
	conteudo	= Pegar("conteudo",0)	
	contato		= Pegar("contato",0)
	CodProduto	= Pegar("Product",0)
		
'### D
	de			= Pegar("de",0)	
	dtData		= Pegar("data",0)
	dtnasc		= Pegar("dtnasc",0)
	departamento= Pegar("departamento",0)
	
'### E
	email		= Pegar("email",0)
	endereco	= Pegar("endereco",0)
	esq			= Pegar("esq",0)
	Erro		= Pegar("Error",0)
	evento		= Pegar("evento",0)
	estudante	= Pegar("estudante",0)
	
'### F
	flash		= Pegar("Flash",0)	
	foto		= Pegar("foto",0)
	fabricante	= Pegar("fabricante",0)

'### G
	Grupo			= Pegar("grupo",0)
	Genero			= Pegar("Genero",0)
	grupo_locacao	= Pegar("grupo_locacao",0)

'### H
	Header		= Request.ServerVariables("ALL_HTTP")
	Horario		= Pegar("Horario",0)
	

'### I
	ItensPorPagina	= 20
	Import			= Pegar("Total",0)
	info			= Pegar("Info",0)
	itens			= Pegar("itens",0)
	ip				= Pegar("ip",0) : if ip = "" then IP = Request.ServerVariables("REMOTE_ADDR")
	IdCliente		= Session.SessionID	
	
	IdAssunto		= Pegar("assunto",0)

'### J

'### K

'### L
	logo		= Pegar("logo",0)
	legenda		= Pegar("legenda",0)

'### M
	mensagem	= Pegar("mensagem",0)
	media		= Pegar("media",0)
	modulo		= Pegar("modulo",0)
	modelo		= Pegar("modelo",0)
	
	
'### N
	nome		= Pegar("nome",0)
	numero		= Pegar("numero",0)
	Nota		= Pegar("Nota",0)		
	NewsID		= Pegar("NewsID",0)

'### O
	Origem		= Pegar("Origem",0)
	IF Instr(Request.QueryString,"Origem=") > 0 THEN
		Origem = Replace(Request.QueryString,"Origem=","")
	END IF
	OperID		= Pegar("OperID",0)
	Opcao		= Pegar("Opcao",0)
	olho		= Pegar("olho",0)
	
'### P
	Pagina		= Pegar("pagina",0)
	If Pagina = "" Then Pagina = 1 End If
	patrimonio	= Pegar("patrimonio",0)

'### Q 

'### R
	ramal		= Pegar("ramal",0)
	
'### S
	senha		= Pegar("senha",0)
	sla			= Pegar("sla",0)
	status		= Pegar("status",0)	
	IF status = "" THEN status = 0
	sessao		= Pegar("sessao",0)
	strTemp		= Pegar("temp",0)
	strSelect	= Pegar("itens",0)	
	strLocal	= Pegar("local",0)
	
	solicitante	= Pegar("solicitante",0)	: if solicitante = "" then solicitante = 0
	supervisor	= Pegar("supervisor",0)		: if supervisor = "" then supervisor = 0
	stt_atual	= Pegar("stt_atual",0)		: if stt_atual = "" then stt_atual = 0
	stt_novo	= Pegar("stt_novo",0)		: if stt_novo = "" then stt_novo = 0

'### T
	telefone	= Pegar("telefone",0)
	tipo		= Pegar("tipo",0)
	titulo		= Pegar("titulo",0)
	
'### U
	user		= Pegar("user",0)
	uf			= Pegar("uf",0)

'### V	
	voltar		= "<table border=0 cellpadding=1 cellspacing=1 class=texto><tr><td> <a href='javascript:history.go(-1)'><img src='images/voltar.png' border=0></a></td><td valign=middle><a href='javascript:history.go(-1)'>Voltar</a></td></tr></table>"

'### X

'### W

'### Y

'### Z
	
'--------------------------------------------------------------'%>