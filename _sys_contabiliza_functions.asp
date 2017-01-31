<%
'-------------------------------------------------
response.Buffer = true
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2 
Response.AddHeader "pragma","no-cache" 
Response.AddHeader "cache-control","private" 
Response.CacheControl = "no-store"
'-------------------------------------------------
	Pagina = request("pagina")
	total = 30 '** por pagina
'-------------------------------------------------
'** FUNÇÃO DE PAGINAÇÃO
'-------------------------------------------------
Function QuebraPagina(Atual,NPagina,total,Etapa)
'-------------------------------------------------
		IF Etapa = 1 THEN
'-------------------------------------------------
			If pagina = "" Then				
				RS.AbsolutePage = 1
				pagina = 1
			Else
				RS.AbsolutePage = pagina
			End If
'-------------------------------------------------
		ELSE
'-------------------------------------------------
Pagina = NPagina
'-------------------------------------------------
			saida =		"<script>"&VBCrlf &_
						"   function ir(travaPag)"&VBCrlf &_
						"{"&VBCrlf &_
						"vamos_la = travaPag.options[travaPag.selectedIndex].value"&VBCrlf &_
						"	if(vamos_la != """")"&VBCrlf &_
						"	{"&VBCrlf &_
						"	window.location.href = vamos_la"&VBCrlf &_
						"	}"&VBCrlf &_
						"}"&VBCrlf &_
						"</script>"&VBCrlf &_
						"    <table width='150' border='0' cellspacing='1' cellpadding='1' class=texto	>"&VBCrlf&_
						"    <form>"&VBCrlf&_
						"      <tr>"&VBCrlf&_
						"        <td class=nlink>"&VBCrlf
'------------------------------------------------------------------------------ 
						If CInt(pagina) > 1 Then
							saida = saida & "<a href='?"&Atual&"&Pagina=1' class='link' title='Primeira Pagina'>"&CHR(171)&"</a>"
						Else
							saida = saida &  CHR(171)
						End If
'------------------------------------------------------------------------------					
					saida = saida & "</td>"&VBCrlf&_
							"        <td class='nlink'>"&VBCrlf
'------------------------------------------------------------------------------ 
						If CInt(pagina) > 1 Then
							saida = saida & "<a href='?"&Atual&"&pagina=" & pagina-1 &"' class='link'>Anterior</a>"
						Else
							saida = saida &  "Anterior"
						End If
'------------------------------------------------------------------------------					
					saida = saida & "</td>"&VBCrlf&_
							"       <td><select name=pagina onChange=""ir(this.form.pagina)"">"&VBCrlf
'------------------------------------------------------------------------------ 
						For i = 1 To Total
							IF CInt(Pagina) = i then 
'------------------------------------------------------------------------------ 
									saida = saida & "<option value='?"& Atual &"&pagina=" & i & "' selected class='scl'>" & i & "</option>"&VBCrlf
'------------------------------------------------------------------------------ 
							ELSE
'------------------------------------------------------------------------------ 
									saida = saida & "<option value='?"& Atual &"&pagina=" & i & "' class='scl1'>" & i & "</option>"&VBCrlf
'------------------------------------------------------------------------------ 					
							END IF
						Next
'------------------------------------------------------------------------------ 
					saida = saida &	"</select></td>"&VBCrlf&_
							"        <td class='nlink'>"&VBCrlf
'------------------------------------------------------------------------------ 
						If CInt(pagina) < total Then
							saida = saida & "<a href='?"& Atual &"&pagina=" & pagina+1 & "' class='link'>Próxima</a>"
						Else
							saida = saida & "Próxima"
						End If
'------------------------------------------------------------------------------					
					saida = saida & "</td>"&VBCrlf&_
							"        <td class='nlink'>"&VBCrlf
'------------------------------------------------------------------------------ 
						If CInt(pagina) < Total Then
							saida = saida & "<a href='?"& Atual &"&pagina="& Total &"' class='link' title='Ultima Pagina'>"&CHR(187)&"</a>"
						Else
							saida = saida &  CHR(187)
						End If
'------------------------------------------------------------------------------
					saida = saida & "</td>"&VBCrlf&_
							"      </tr>"&VBCrlf&_
							"    </form>"&VBCrlf&_
							"      <tr>"&VBCrlf&_
							"      <td colspan=5 align=center>Página "& pagina &" de "& Total &"</td>"&VBCrlf&_
							"      </tr>"&VBCrlf&_
							"    </table>"&VBCrlf
'------------------------------------------------------------------------------
		QuebraPagina =  saida
'------------------------------------------------------------------------------
		END IF
'------------------------------------------------------------------------------
END Function
'-------------------------------------------------
'** Funcao para criacao de string de conexao
'-------------------------------------------------
Function ConnectionString(intStr,intType,strServer,strBase,strUser,strSenha,boMaped)
'-------------------------------------------------
	Select Case intStr
	'-----------------------------------------
		Case 1 '### SQL SERVER
		'---------------------------------
			Select Case intType
			'-------------------------
				case 1 '// Standard Security
				'------------------
				  strConn = "Driver={SQL Server};Server="& strServer &";Database="& strBase &";Uid="& strUser &";Pwd="& strSenha &";" 
				'------------------
				case 2 '// Trusted connection:
				'------------------
				  strConn = "Driver={SQL Server};Server="& strServer &";Database="& strBase &";Uid="& strUser &";Pwd="& strSenha &";Trusted_Connection=yes;" 
				'------------------
				case 3 '// DSN connection:
				'------------------
				  strConn = "DSN="& strServer &";Uid="& strUser &";Pwd="& strSenha &";" 
			'-------------------------
			End Select
		'---------------------------------
		Case 2 '### ACCESS
		'---------------------------------
		  IF boMaped = False Then
		  '-------------------------
		    strBase = server.MapPath(strBase)
		  '-------------------------
		  END IF
		'---------------------------------
			Select Case intType
			'-------------------------
				Case 1 '// Standard Security
				'------------------
				  strConn = "Driver={Microsoft Access Driver (*.mdb)};Dbq="& strBase &";Uid="& strUser &";Pwd="& strSenha &";" 
				'------------------
				Case 2 '// Exclusive
				'------------------
				  strConn = "Driver={Microsoft Access Driver (*.mdb)};Dbq="& strBase &";Exclusive=1;Uid="& strUser &";Pwd="& strSenha &";" 
			'-------------------------
			End Select
		'---------------------------------
		Case 3 '### ORACLE
		'---------------------------------
			Select Case intType
			'-------------------------
				Case 1 '// New version
				'------------------
				  strConn = "Driver={Microsoft ODBC for Oracle};Server="& strServer &";Uid="& strUser &";Pwd="& strSenha &";" 
				'------------------
				Case 2 '// Old version
				'------------------
				  strConn = "Driver={Microsoft ODBC Driver for Oracle};ConnectString="& strServer &";Uid="& strUser &";Pwd="& strSenha &";" 
			'-------------------------
			End Select
		'---------------------------------
		Case 4 '### MYSQL
		'---------------------------------
			Select Case intType
			'-------------------------
				Case 1 '// ODBC 2.50 Local database
				'------------------
				  strConn = "Driver={mySQL};Server=localhost;Option=16834;Database="& strBase &";" 
				'------------------
				Case 2 '// ODBC 2.50 Remote database
				'------------------
				  strConn = "Driver={mySQL};Server="& strServer &";Port=3306;Option=131072;Stmt=;Database="& strBase &";Uid="& strUser &";Pwd="& strSenha &";"
				'------------------
				Case 3 '// ODBC 3.51 Local database
				'------------------
				  strConn = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE="& strBase &";USER="& strUser &";PASSWORD="& strSenha &";OPTION=3;" 
				'------------------
				Case 4 '// ODBC 3.51 Remote database
				'------------------
				  strConn = "DRIVER={MySQL ODBC 3.51 Driver};SERVER="& strSenha &";PORT=3306;DATABASE="& strBase &";USER="& strUser &";PASSWORD="& strSenha &";OPTION=3;"
 			'-------------------------
			End Select
	'-----------------------------------------
	End Select
	'-----------------------------------------
	ConnectionString = strConn
'-------------------------------------------------
End Function
'-------------------------------------------------------'
'## Função responsável por formatar o texto cxalta_cxbaixa
'-------------------------------------------------------'
Function Lower(nome)
'-------------------------------------------------------'
		saida = split(nome," ")
		'-----------------------------------------------'
		for each item in saida
		'-----------------------------------------------'
			IF nm = "" then
			'-------------------------------------------'
				nm = Ucase(Left(item,1)) & Lcase(Mid(item,2,Len(item)))
			'-------------------------------------------'
			ELSE
			'-------------------------------------------'
				nm = nm &" "& Ucase(Left(item,1)) & Lcase(Mid(item,2,Len(item)))
			'-------------------------------------------'
			End IF
		'-----------------------------------------------'
		next
		'-----------------------------------------------'
		Lower = nm
'-------------------------------------------------------'
End Function 
'---------------------------------
'** Calculo para exibicação
'---------------------------------
Function Calculo(strValor)
'---------------------------------
	strValor = FormatNumber((strValor / 1024),2)
	tipo = Split(strValor,".")
'---------------------------------
	Select Case Ubound(tipo)	
'---------------------------------
			Case 1 : tipo = " <font color=blue>Mb</font>"
			Case 2 : tipo = " <font color=red>Gb</font>"
			Case else : tipo = " Kb"
'---------------------------------
	End Select
'---------------------------------
	Calculo = "<b>"& strValor & "</b>" & tipo	
'---------------------------------
End Function
'---------------------------------
'** Propriedades do banco / Tabelas / Views / Procedures
'---------------------------------
Function Propriedades(strFilter,strProc)
'---------------------------------
	IF strFilter <> "" THEN strFilter =  "'"& strFilter &"'" : strBanco = False Else  strBanco = True
'---------------------------------
	Set Conx	=	Server.CreateObject("ADODB.Connection")
	Conx.Open strConn	
'---------------------------------
	Set objRS = Server.CreateObject("ADODB.Recordset")
	objRs.Open "EXEC "& strProc & strFilter, Conx, 1, 3
	If Not objRs.EOF Then
'---------------------------------
		IF strBanco Then
'---------------------------------
			strS =	 " <b>"& Lower(objRs("database_name"))  &"</b> <font color=red>| Espaço Utilizado ("& objRs("database_size") &") | Não Alocado ("& objRs("unallocated space") &") </font>" 
'---------------------------------
		ELSE
'---------------------------------
			strS =	" Registros: <a href=""javascript:RunQuery("& strFilter &",'block')"">"& objRs("rows") &"</a>"& VBCrlf &_
					"| Esp. Reservado: "& objRs("reserved") & VBCrlf &_
					"<br> Esp. Utilizado: "& objRs("data") & VBCrlf &_
					"| Esp. Índices: "& objRs("index_size") & VBCrlf &_
					"<br> Esp. não usado: "& objRs("unused") 
'---------------------------------
		END IF
'---------------------------------
	    'For Each objCol In objRs.Fields
	    '    Response.Write(objCol.Name & ": " & objCol.Value & "<br>")
	    'Next
'---------------------------------
	End If
'---------------------------------
	Propriedades = strS
'---------------------------------
	objRs.Close : Conx.Close
	set objRs = Nothing 
	set Conx = Nothing
'---------------------------------
End Function
'---------------------------------
'** Colunas da Tabela / Tipo de dados
'---------------------------------
Function ColunasTabela(strFilter,strProc)
'---------------------------------
	'IF strFilter <> "" THEN strFilter =  "'"& strFilter &"'" 
'---------------------------------
	Set Conx	=	Server.CreateObject("ADODB.Connection")
	Conx.Open strConn	
'---------------------------------
	Set objRS = Server.CreateObject("ADODB.Recordset")
	objRs.Open "EXEC "& strProc &" "& strFilter, Conx, 1, 3
	If Not objRs.EOF Then
'---------------------------------
		strS =	"<table border=0 cellpadding=0 cellspacing=0 width='100%' class=texto>"&_
				"<tr class=titulo>"&_
				"<td>Coluna</td>"&_
				"<td>Tipo</td>"&_
				"<td>Tamanho</td>"&_
				"<td>Precisão</td>"&_
				"<td>Valor Padrão</td>"&_
				"<td>Is Null?</td>"&_
				"</tr>"
'---------------------------------
		Do While Not objRs.EOF 
'---------------------------------
			strS =	strS & 	"<tr>"&_
							"<td>"& objRs("COLUMN_NAME") &"</td>"&_
							"<td>"& objRs("TYPE_NAME") &"</td>"&_
							"<td>"& objRs("LENGTH") &"</td>"&_
							"<td>"& objRs("PRECISION") &"</td>"&_
							"<td>"& objRs("COLUMN_DEF") &"</td>"&_
							"<td>"& objRs("IS_NULLABLE") &"</td>"&_
							"</r>"
'---------------------------------
		objRs.MoveNext
		Loop
'---------------------------------
		strS =	strS & 	"</table>"
'---------------------------------
	End If
'---------------------------------
	ColunasTabela = strS
'---------------------------------
	objRs.Close : Conx.Close
	set objRs = Nothing 
	set Conx = Nothing
'---------------------------------
End Function
'---------------------------------
'** Colunas da Tabela / Tipo de dados * editor online
'---------------------------------
Function ColunasTabelaQueryAnalyser(strFilter,strProc)
'---------------------------------
	'IF strFilter <> "" THEN strFilter =  "'"& strFilter &"'" 
'---------------------------------
	Set Conx	=	Server.CreateObject("ADODB.Connection")
	Conx.Open strConn	
'---------------------------------
	Set objRS = Server.CreateObject("ADODB.Recordset")
	objRs.Open "EXEC "& strProc &" "& strFilter, Conx, 1, 3
	If Not objRs.EOF Then
'---------------------------------
		strS =	"<table border=0 align=center cellpadding=0 cellspacing=0 width='95%' class=texto>"&_
				"<tr class=titulo>"&_
				"<td>Coluna</td>"&_
				"<td>Tipo</td>"&_
				"</tr>"
'---------------------------------
		Do While Not objRs.EOF 
'---------------------------------
			strS =	strS & 	"<tr>"&_
							"<td><a href=""javascript:AddToSql(document.frm.strSQL,'" & objRs("COLUMN_NAME") &"')"">"& objRs("COLUMN_NAME") &"</a></td>"&_
							"<td>"& objRs("TYPE_NAME") &"</td>"&_
							"</r>"
'---------------------------------
		objRs.MoveNext
		Loop
'---------------------------------
		strS =	strS & 	"</table>"
'---------------------------------
	End If
'---------------------------------
	ColunasTabelaQueryAnalyser = strS
'---------------------------------
	objRs.Close : Conx.Close
	set objRs = Nothing 
	set Conx = Nothing
'---------------------------------
End Function
'---------------------------------
'** Tipo de dado das colunas da tabela
'---------------------------------
Function ColunasTabelaDescr(strFilter,strProc,strColumn)
'---------------------------------
	Set Conx	=	Server.CreateObject("ADODB.Connection")
	Conx.Open strConn	
'---------------------------------
	Set objRS = Server.CreateObject("ADODB.Recordset")
	objRs.Open "EXEC "& strProc &" "& strFilter, Conx, 1, 3
	If Not objRs.EOF Then
'---------------------------------
		Do While Not objRs.EOF 
'---------------------------------
			strX =	strX & 	objRs(strColumn) &","
'---------------------------------
		objRs.MoveNext
		Loop
'---------------------------------
	End If
'---------------------------------
	ColunasTabelaDescr = strX
'---------------------------------
	objRs.Close 
	set objRs = Nothing 
'---------------------------------
End Function
'---------------------------------
'** trata a saida para o SQL 
'---------------------------------
Function PrepairSql(sType,sValor)
'---------------------------------
		Select Case cStr(sType)
'---------------------------------
			Case "char","datetime","nchar","ntext","nvarchar","smalldatetime","text","timestamp","varchar"  
'---------------------------------
				sType = "'"& sValor &"'"
'---------------------------------
			Case else 
'---------------------------------
				sType =  sValor
'---------------------------------		
		End Select
'---------------------------------		
		PrepairSql = sType
'---------------------------------
End Function
'---------------------------------
'** Listagem das linhas da tabela escolhida
'---------------------------------
Function ViewLines(strFilter,OrderBy)
'---------------------------------
	Set Conx	=	Server.CreateObject("ADODB.Connection")
	Conx.Open strConn	
	Session("Colunas") = ""
'---------------------------------
	Set objRS = Server.CreateObject("ADODB.Recordset")
	objRs.Open "EXEC sp_columns "& strFilter, Conx, 1, 3
	If Not objRs.EOF Then
'---------------------------------
		strS =	"<font face=tahoma size=2><b>Dados da Tabela '<font color=red>"& strFilter &"</font>'</b></span><br><br>"& VBCrlf &_
				"<a href=""javascript:sqlInsPop('?"& atual &"&Metodo=ins')"" style='text-decoration:none'>[Adicionar Registro]</a><br><br>"&VBCrlf &_
				"<script language=JavaScript src='sys_contabiliza.js'></script>"& VBCrlf &_
				"<body scroll=auto>"& VBCrlf &_
				"<table border=0 cellpadding=1 cellspacing=1 width='100%' class=texto>"& VBCrlf &_
				"<form name='editSql' method='post' action='?"& Atual &"&Metodo=UP'>"& VBCrlf &_
				"<tr class=titulo bgcolor='#FFEBCD'>"& VBCrlf 
'---------------------------------
		Do While Not objRs.EOF 
'---------------------------------
			'## PRIMEIRA VEZ IGUALA AO ID DA TABLE
			'---------------------------------
			IF Trim(sColumn) = "" THEN sColumn = cStr(objRs("COLUMN_NAME")) 
'---------------------------------
			IF cStr(sColumn) = cStr(objRs("COLUMN_NAME")) THEN
'---------------------------------
				cnt = 1
'---------------------------------
			ELSE
'---------------------------------
				cnt = 0
'---------------------------------
			END IF
'---------------------------------				
			strS =	strS & 	"<td><a href='?"& Atual  &"&Pagina="& Pagina &"&sColumn="& objRs("COLUMN_NAME") &"&OrderBy="& OrderBy &"&Change=1'>"& objRs("COLUMN_NAME") &"</a> <font color=red><b>"& VBCrlf 
'---------------------------------	
			IF cnt = 1 THEN strS = strS & OrderBy
'---------------------------------	
			strS =	strS & 	"</b></font> </td>"& VBCrlf 
'---------------------------------
			sColunas  = sColunas  & objRs("COLUMN_NAME") &","
'---------------------------------
		objRs.MoveNext
		Loop
'---------------------------------		
		Session("Colunas") = sColunas 
'---------------------------------
	End If
'---------------------------------
strS =	strS & 	"<td align=center>&nbsp;</td>"& VBCrlf &_
				"<td align=center>&nbsp;</td>"& VBCrlf 	
'---------------------------------
	ViewLines = strS
'---------------------------------
	objRs.Close : Conx.Close
	set objRs = Nothing 
	set Conx = Nothing
'---------------------------------
End Function%>