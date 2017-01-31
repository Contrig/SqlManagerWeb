<%response.buffer = true%>
<!--#include file="functions.asp"-->
<!--#include file="_sys_contabiliza_functions.asp"-->
<html>
<head>
<title>Análise de banco de dados SQL</title>
<link href="_sys_contabiliza.css" rel="stylesheet" type="text/css">
<style>
table{
font-family : verdana;
font-size : 12px; 
}
.sim {background-color: #EFEFEF;}
.nao {background-color: #E2E2AD;}
.item { font-size : 9px; text-align:center; background-color: #DFF9CC; }
#box {width: 100%; }
</style>
</head>
<body scroll=auto>
<%
Registros	= Request("registros")
sTable		= Request("sTable")
sColumn		= Request("sColumn")
sMetodo		= Request("Metodo")
sId			= Request("sId")
OrderBy		= Request("OrderBy")
Change		= Request("Change")
Cor			= "sim" 
'---------------------------------
'** Compactação do Banco de Dados
'--------------------------------
IF Registros = "compact_sql" THEN
'---------------------------------
	Set Conx	=	Server.CreateObject("ADODB.Connection")
	Conx.Open strConn
'-----------------------------------------'	
	strSQL = "SP_HELPFILE"
'-----------------------------------------'
	'On Error Resume Next
'-----------------------------------------'
	Set Rs = server.CreateObject("adodb.recordset")
		Rs.Open strSQL, Conx, 1, 3
'---------------------------------
	IF Err.number <> 0 THEN
'---------------------------------
		sAlert = "<img src='alert.gif' hspace=5 vspace=5>"
'---------------------------------	
	END IF
'-----------------------------------------'%>
	<body onload=this.focus()>
	<table width="100%" border="0" cellspacing="0" cellpadding="0" ID="Table2">
	<tr>
		<td colspan="2" align=center><font color=#3300ff><b>Resultado da compactação do banco de dados</b></font> </td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor='ffffe0'>
		<a name='resultado'>
		<%IF sAlert <> "" THEN%>
			<table border=0 width='70%' cellpadding=0 cellspacing=0 class=texto align=center ID="Table3">
			<tr align=center>
				<td><%=sAlert%></td>
				<td><font size=3><b>Ocorreram erros durante o processo de execução do comando sql</b></font></td>
			</tr>
			<tr>
			<td colspan=2>&nbsp;</td>
			</tr>
			<tr>
			<td colspan=2><b>Linha de comando solicitada</b><pre><%=strSQL%></pre></td>
			</tr>
			<tr>
			<td colspan=2><br><b>Código do Erro</b> <font color=red><%=Err.number%></font><br><br></td>
			</tr>
			<tr>
			<td colspan=2><b>Descrição do erro</b><pre><%=Err.Description%></pre> </td>
			</tr>
			<tr>
			<td colspan=2 align=center><b><font color=red>A operação não pode ser concluída.</font></b></td>
			</tr>		  
			</table>
		<%ELSE
		'----------------------------
			while not Rs.EOF
		'----------------------------			
				nome	= Lcase(Rs.Fields(0))
				arquivo = Rs.Fields(2)
				tamanho = Rs.Fields(4)
				maxtam	= Rs.Fields(5)
				taxcres	= Rs.Fields(6)
		'----------------------------		
				if instr(nome,"_log") > 0 then
		'----------------------------
					nm1	= nome
					ar1	= arquivo
					tm1 = cInt(replace(tamanho," KB","")) *1024
					mx1	= maxtam
					tx1	= taxcres
		'----------------------------
				end if		
		'----------------------------
			rs.MoveNext : wend
		'----------------------------
		'** Compactando o banco de dados 
		'----------------------------
			Log_Bd = StrReverse(ar1) : Log_Bd = mid(Log_Bd,1,Instr(Log_Bd,"\")-1) : Log_Bd	= StrReverse(Log_Bd) 
			Log_Bd	= mid(Log_Bd,1,Instr(Log_Bd,".")-1)
		'----------------------------
			strSQL1 = "dbcc shrinkfile('"& Log_Bd &"',TRUNCATEONLY)" 
			Set objRs = server.CreateObject("adodb.recordset")
				objRs.Open strSQL1, Conx, 1, 3
		'----------------------------
		'** Pegando os novos valores do banco de dados
		'----------------------------	
		Set Rs = server.CreateObject("adodb.recordset")
			Rs.Open strSQL, Conx, 1, 3
			while not Rs.EOF
		'----------------------------			
				nome	= Lcase(Rs.Fields(0))
				arquivo = Rs.Fields(2)
				tamanho = Rs.Fields(4)
				maxtam	= Rs.Fields(5)
				taxcres	= Rs.Fields(6)
		'----------------------------		
				if instr(nome,"_log") > 0 then
		'----------------------------
					nm2	= nome
					ar2	= arquivo
					tm2 = cInt(replace(tamanho," KB","")) *1024
					mx2	= maxtam
					tx2	= taxcres
		'----------------------------
				end if		
		'----------------------------
			rs.MoveNext : wend
		'----------------------------%>
			<table border=0 width='98%' cellpadding=1 cellspacing=1 class=texto align=center>
			 <tr>
			   <td colspan=4 align=center><b>Antes da compactação</b></td>
			 </tr>
			 <tr>
			   <td colspan=4 align=center>&nbsp;</td>
			 </tr>			 
			 <tr bgcolor=#669900 style='color:#FFFFFF'>
				<td><b>Arquivo</b></td>
				<td><b>Tamanho Atual</b></td>
				<td><b>Tamanho Permitido</b></td>
				<td><b>Taxa de crescimento</b></td>
			 </tr>
			 <tr>
				<td><b><%=ar1%></b></td>
				<td><b><%=Calculo(tm1)%></b></td>
				<td><b><%=mx1%></b></td>
				<td><b><%=tx1%></b></td>
			 </tr>
			 <tr>
			   <td colspan=4 align=center>&nbsp;</td>
			 </tr>			 
			 <tr>
			   <td colspan=4 align=center><b>Depois da compactação</b></td>
			 </tr>
			 <tr>
			   <td colspan=4 align=center>&nbsp;</td>
			 </tr>			 
			 <tr bgcolor=#ffcc00 style='color:#000000'>
				<td><b>Arquivo</b></td>
				<td><b>Tamanho Atual</b></td>
				<td><b>Tamanho Permitido</b></td>
				<td><b>Taxa de crescimento</b></td>
			 </tr>
			 <tr>
				<td><b><%=ar2%></b></td>
				<td><b><%=Calculo(tm2)%></b></td>
				<td><b><%=mx2%></b></td>
				<td><b><%=tx2%></b></td>
			 </tr>
			</table>
		<%END IF%>
	</tr>
	</table>
<%'---------------------------------
'** Compactação do Banco de Dados
'--------------------------------
ELSEIF Registros = "property_sql" THEN
'---------------------------------
	Set Conx	=	Server.CreateObject("ADODB.Connection")
	Conx.Open strConn
'-----------------------------------------'	
	strSQL = "SP_HELPFILE"
'-----------------------------------------'
	On Error Resume Next
'-----------------------------------------'
	Set Rs = server.CreateObject("adodb.recordset")
		Rs.Open strSQL, Conx, 1, 3
'---------------------------------
	IF Err.number <> 0 THEN
'---------------------------------
		sAlert = "<img src='alert.gif' hspace=5 vspace=5>"
'---------------------------------	
	END IF
'-----------------------------------------'%>
	<body onload=this.focus()>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td colspan="2" align=center><font color=#3300ff><b>Propriedades do Banco de dados </b></font> </td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor='ffffe0'>
		<a name='resultado'>
		<%IF sAlert <> "" THEN%>
			<table border=0 width='70%' cellpadding=0 cellspacing=0 class=texto align=center ID="Table1">
			<tr align=center>
				<td><%=sAlert%></td>
				<td><font size=3><b>Ocorreram erros durante o processo de execução do comando sql</b></font></td>
			</tr>
			<tr>
			<td colspan=2>&nbsp;</td>
			</tr>
			<tr>
			<td colspan=2><b>Linha de comando solicitada</b><pre><%=strSQL%></pre></td>
			</tr>
			<tr>
			<td colspan=2><br><b>Código do Erro</b> <font color=red><%=Err.number%></font><br><br></td>
			</tr>
			<tr>
			<td colspan=2><b>Descrição do erro</b><pre><%=Err.Description%></pre> </td>
			</tr>
			<tr>
			<td colspan=2 align=center><b><font color=red>A operação não pode ser concluída.</font></b></td>
			</tr>		  
			</table>
		<%ELSE%>
			<table border=0 width='98%' cellpadding=1 cellspacing=1 class=texto align=center>
				<tr>			
					<%for i = 0 to Rs.Fields.Count -1
						response.Write "<td>"& Rs.Fields(i).Name & "</td>"& VBCrlf 
					next%>
				</tr>
				<tr>
					<td bgcolor=000000 style='line-height:1px;' colspan=<%=Rs.Fields.Count %>>&nbsp;</td>
				</tr>
				<tr>
					<td style='line-height:10px;' colspan=<%=Rs.Fields.Count %>>&nbsp;</td>
				</tr>
				<%while not Rs.EOF%>
				<tr>			
					<%for i = 0 to Rs.Fields.Count -1
						response.Write "<td class='"& cor &"'>"& Rs.Fields(i) & "</td>"& VBCrlf 						
					next%>
				</tr>
				<%
				cont = cont + 1
				if cor="sim" then cor="nao" else cor="sim"
				rs.MoveNext
				wend%>
			</table>
		<%END IF%>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td colspan=2 align=center> <a href='javascript:window.close()'>:: Fechar ::</a> </td>
	</tr>
	</table>
<%'---------------------------------
'** Painel de execução SQL
'---------------------------------
ELSEIF Registros = "query_analyser" THEN
'---------------------------------
IF sMetodo <> "" THEN strSQL =  Request("strSQL")
'---------------------------------%>
	<script language=javascript src='sys_contabiliza.js'></script>
	<script>
	function AddToSql(objForm,valor)
	{
		valor			= valor.toUpperCase();
		objForm.value  += valor.toUpperCase() +' ';
		var range		= objForm.createTextRange();
		range.move("textedit");
		range.select();
		//objForm.focus();
	}
	function textSql(objForm)
	{
		objForm.value = objForm.value.toUpperCase()
	}
	function limpar()
	{
		var cpo	=	document.frm.strSQL;
					cpo.value="";
					cpo.focus();
	}
	function envio()
	{
		document.frm.xEnvio.value = 'Processando...';
		document.frm.xEnvio.disabled = true;	
	}
	setTimeout("document.frm.strSQL.focus()",1000)
	</script>
	<body onload=this.focus()>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<form name='frm' method='post' action='<%=request.ServerVariables("script_name")%>?Registros=query_analyser&Metodo=exec' onsubmit="envio()">
	<tr>
		<td colspan="2" align=center><font color=red><b>ADVERTÊNCIA:</b></font> Qualquer comando executado neste módulo, <b>não</b> poderá ser revertido.</td>
	</tr>
	<tr>
		<td colspan="2" align=center>&nbsp;</td>
	</tr>
	<tr>
		<td width="25%" height=30><b>Objetos</b></td>
		<td width="75%"><b>Comandos Sql</b> 
		</td>
	</tr>
	<tr>
		<td>
			<span style='display:none;' id="B0"></span>
			<a href='javascript:banco(0)'>(Hide)</a> / 
			<a href='javascript:banco(1)'>Tabelas</a> /
			<a href='javascript:banco(2)'>View</a> /
			<a href='javascript:banco(3)'>Procedure</a>
		</td>
		<td>
			<a href="javascript:AddToSql(document.frm.strSQL,'select')">Select</a> /
			<a href="javascript:AddToSql(document.frm.strSQL,'delete from')">Delete</a> /
			<a href="javascript:AddToSql(document.frm.strSQL,'update')">Update</a> /
			<a href="javascript:AddToSql(document.frm.strSQL,'insert into')">Insert</a> /
			<a href="javascript:AddToSql(document.frm.strSQL,'truncate')">Truncate</a> /
			<a href="javascript:AddToSql(document.frm.strSQL,'alter <objeto:table/view/procedure>')">Alter</a> /
			<a href="javascript:AddToSql(document.frm.strSQL,'drop <objeto:table/view/procedure>')">Drop</a> /
			<a href="javascript:AddToSql(document.frm.strSQL,'exec <stored procedure>')">Exec</a> <b>|</b>
			<a href="javascript:AddToSql(document.frm.strSQL,'Distinct')">Distinct</a> / 
			<a href="javascript:AddToSql(document.frm.strSQL,' from')">from</a> / 
			<a href="javascript:AddToSql(document.frm.strSQL,'top <quantidade>')">Top</a> / 
			<a href="javascript:AddToSql(document.frm.strSQL,'where')">Where</a> / 
			<a href="javascript:AddToSql(document.frm.strSQL,'like')">Like</a> / 
			<a href="javascript:AddToSql(document.frm.strSQL,'group by')">Group By</a> 
		</td>
	</tr>
	<tr>
		<td height="23"><br>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
		   <td colspan=2 valign=top>
			<table border=0 width='100%' cellpadding=0 cellspacing=0 class=texto style='display:none;' id='B1'>
			<%On Error Resume next
			'---------------------------------	
			Tabela	= Session("Tabela")
			View	= Session("View")
			Proced	= Session("Proced")
			'---------------------------------	
			for i=Lbound(Tabela) to Ubound(Tabela,2)
			'---------------------------------	
				Response.Write	"<tr class="& cor &">"& VBCrlf &_
								"<td align=center  onClick=""javascript:coluna('"&Tabela(0,i)&"')"" style='cursor:hand' title='Clique para ver as colunas desta tabela'><img src='base_table.gif'></td>"& VBCrlf &_
								"<td><a href=""javascript:AddToSql(document.frm.strSQL,'" & Tabela(0,i) &"')"">" & Tabela(0,i) &"</a></td>"& VBCrlf &_
								"</tr>"& VBCrlf &_
								"<tr class="& cor &" id="&Tabela(0,i)&" style='display:none'>"& VBCrlf &_
								"<td colspan='4' align='center'>"& ColunasTabelaQueryAnalyser(Tabela(0,i),"sp_columns") &"</td>"& VBCrlf &_
								"</tr>"& VBCrlf &_
								"<tr>"& VBCrlf &_
								"<td colspan='4' style='line-height: 2px;'>&nbsp;</td>"& VBCrlf &_
								"</tr>"
			'---------------------------------
				IF Cor = "sim" THEN Cor = "nao"  ELSE Cor = "sim" 
			'---------------------------------
				next
			'---------------------------------
			Response.Write "</table>"& VBCrlf &_
							"<table border=0 width='100%' cellpadding=1 cellspacing=1 class=texto style='display:none;' id='B2'>"
			'---------------------------------
			for i=Lbound(View) to Ubound(View,2)
			'---------------------------------
				Response.Write	"<tr class="& cor &">"& VBCrlf &_
								"<td align=center  onClick=""javascript:coluna('"&View(0,i)&"')"" style='cursor:hand' title='Clique para ver as colunas desta View (Visão)'><img src='base_view.gif'></td>"& VBCrlf &_
								"<td><a href=""javascript:AddToSql(document.frm.strSQL,'" & View(0,i) &"')"">"& View(0,i) &"</a></td>"& VBCrlf &_
								"</tr>"& VBCrlf &_
								"<tr class="& cor &" id="&View(0,i)&" style='display:none'>"& VBCrlf &_
								"<td colspan=4>"& ColunasTabelaQueryAnalyser(View(0,i),"sp_columns") &" </td>"& VBCrlf &_
								"</tr>"& VBCrlf &_
								"<tr>"& VBCrlf &_
								"<td colspan=4 style='line-height: 2px;'>&nbsp;</td>"& VBCrlf &_
								"</tr>"
			'---------------------------------
				IF Cor = "sim" THEN Cor = "nao"  ELSE Cor = "sim" 
			'---------------------------------
			next
			'---------------------------------
				Response.Write "</table><table border=0 width='100%' cellpadding=1 cellspacing=1 class=texto style='display:none;' id='B3'>"
			'---------------------------------
			for i=Lbound(Proced) to Ubound(Proced,2)
			'---------------------------------
				Response.Write	"<tr class="& cor &" align=center>"& VBCrlf &_
								"<td><img src='base_procedure.gif'></td>"& VBCrlf &_
								"<td><a href=""javascript:AddToSql(document.frm.strSQL,'" & Proced(0,i) &"')"">"& Proced(0,i) &"</a></td>"& VBCrlf &_
								"<td colspan=2>" & Proced(1,i) &"</td>"& VBCrlf &_
								"</tr>"
			'---------------------------------
				IF Cor = "sim" THEN Cor = "nao"  ELSE Cor = "sim" 
			'---------------------------------
			next
			'---------------------------------
			Response.Write "</table>"
			%>
			</td>
			</tr>
		</table>
		
		</td>
		<td valign=top align=center>
			<textarea cols=50 rows=10 name='strSQL' style='width:100%;' onBlur='textSql(this)' tabindex=1><%=strSQL%></textarea>
			<input type="submit" name="Submit" value="Executar" style='background-color:#8FE36B;width:150px;font: 12px; height: 22px;'>
			<input type="button" onclick='limpar()' name="Limpar" value="Limpar" style='background-color:#EfEfEf;width:50px;font: 12px; height: 22px;'>
			<br><a href='javascript:window.close()'>:: Fechar ::</a>			
		</td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
<%'---------------------------------------'
'** Executando o script SQL
'-----------------------------------------'
IF Lcase(sMetodo) = "exec" THEN
'-----------------------------------------'
	strSQL		= Request("Q")
	Err.Clear
'-----------------------------------------'	
	Set Conx	=	Server.CreateObject("ADODB.Connection")
	Conx.Open strConn
'-----------------------------------------'
	On Error Resume Next
'-----------------------------------------'
	Set Rs = server.CreateObject("adodb.recordset")
		Rs.Open strSQL, Conx, 1, 3
'---------------------------------
	IF Err.number <> 0 THEN
'---------------------------------
		sAlert = "<img src='alert.gif' hspace=5 vspace=5>"
'---------------------------------	
	END IF
'-----------------------------------------'%>
	<tr>
		<td colspan="2" align=center><font color=#3300ff><b>Resultado do comando sql</b></font> </td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor='ffffe0'>
		<a name='resultado'>
		<%IF sAlert <> "" THEN%>
		<table border=0 width='70%' cellpadding=0 cellspacing=0 class=texto align=center>
		  <tr align=center>
		    <td><%=sAlert%></td>
		    <td><font size=3><b>Ocorreram erros durante o processo de execução do comando sql</b></font></td>
		  </tr>
		  <tr>
		   <td colspan=2>&nbsp;</td>
		  </tr>
		  <tr>
		   <td colspan=2><b>Linha de comando solicitada</b><pre><%=strSQL%></pre></td>
		  </tr>
		  <tr>
		   <td colspan=2><br><b>Código do Erro</b> <font color=red><%=Err.number%></font><br><br></td>
		  </tr>
		  <tr>
		   <td colspan=2><b>Descrição do erro</b><pre><%=Err.Description%></pre> </td>
		  </tr>
		  <tr>
		   <td colspan=2 align=center><b><font color=red>A operação não pode ser concluída.</font></b></td>
		  </tr>		  
		</table>
		<%ELSE
		'---------------------------------	
			if UCase(Mid(strSQL,1,6)) = "SELECT" then
		'---------------------------------	
				If Not Rs.EOF Then
		'---------------------------------
					Rs.PageSize =  20  '### QTD DE ITENS POR PAGINA
					Call QuebraPagina(0,0,0,1)
					cont = 0
		'---------------------------------%>
					<table border=0 width='98%' cellpadding=1 cellspacing=1 class=texto align=center>
					<tr>			
						<%for i = 0 to Rs.Fields.Count -1
							response.Write "<td>"& Rs.Fields(i).Name & "</td>"& VBCrlf 
						next%>
					</tr>
					<tr>
						<td bgcolor=000000 style='line-height:1px;' colspan=<%=Rs.Fields.Count %>>&nbsp;</td>
					</tr>
					<tr>
						<td style='line-height:10px;' colspan=<%=Rs.Fields.Count %>>&nbsp;</td>
					</tr>
					<%while not Rs.EOF AND cont < Rs.PageSize%>
					<tr>			
						<%for i = 0 to Rs.Fields.Count -1
							response.Write "<td class='"& cor &"'>"& Rs.Fields(i) & "</td>"& VBCrlf 
							
						next%>
					</tr>
					<%
					cont = cont + 1
					if cor="sim" then cor="nao" else cor="sim"
					rs.MoveNext
					wend
	'---------------------------------
				total = int(Rs.PageCount)
				Atual = "Registros=query_analyser&Metodo=exec&strSQL="& strSQL
	'---------------------------------
			response.Write 	"<tr>"& VBCrlf &_
							"<td colspan='"& Rs.Fields.Count -1 &"' align=center>"& VBCrlf &_
							QuebraPagina(Atual,Pagina,total,0) & VBCrlf &_
							"</td>"& VBCrlf &_
							"</tr>"& VBCrlf &_
							"</table>"& VBCrlf 
		'---------------------------------
				else
		'---------------------------------
					response.Write "<font color=Red><b>Não foram localizados registros para o comando executado</b>.</font>"
		'---------------------------------
				end if
		'---------------------------------
			else
		'---------------------------------	
				if UCase(Mid(strSQL,1,6)) = "DELETE" OR UCase(Mid(strSQL,1,6)) = "UPDATE" then
		'---------------------------------
					response.Write "Comando executado com sucesso"
		'---------------------------------
				else
		'---------------------------------
					do while not Rs.EOF
						for i = 0 to Rs.Fields.Count -1
							response.Write Rs.Fields(i) &" "'"<font color=blue>Comando executado com sucesso.</font>"
						next
							response.Write " <br>" 
						rs.movenext
					loop
		'---------------------------------
				end if
		'---------------------------------
			end if
		'---------------------------------	
		END IF
		'---------------------------------	%>
		</td>
	</tr>
	<script>setTimeout("location='#resultado'",1000)</script>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
<%'-----------------------------------------'
END IF
'-------------------------------------------'%>
	</form>
	</table>
<%'---------------------------------
	response.End
'---------------------------------
END IF
'---------------------------------
'** MOSTRANDO OS REGISTROS DA TABELA ESCOLHIDA
'---------------------------------
IF Registros <> "" THEN
'---------------------------------
	'** Metodos usados junto ao banco
'---------------------------------
	Select Case Ucase(sMetodo)
'---------------------------------
			Case "INS" '** pop de visualizacao
'---------------------------------
					MyColumns		= split(sColumn,",")									'* Colunas passadas pelo metodo GET
					sTypeColumns	= ColunasTabelaDescr(sTable,"sp_columns","TYPE_NAME")	'* Tipo de dados das colunas
					sSizeColumns	= ColunasTabelaDescr(sTable,"sp_columns","PRECISION")	'* Tamanho dos campos das colunas
					sNullColumns	= ColunasTabelaDescr(sTable,"sp_columns","IS_NULLABLE")	'* A coluna pode ser vazia
'---------------------------------
					vetTypeColumns = split(sTypeColumns,",",-1)
					vetSizeColumns = split(sSizeColumns,",",-1)
					vetNullColumns = split(sNullColumns,",",-1)
'---------------------------------
						strS =	"<font face=tahoma size=2><b>Inclusão de registro na tabela '<font color=red>"& sTable &"</font>'</b></span><br><br>"& VBCrlf &_
								"<script language=JavaScript src='sys_contabiliza.js'></script>"& VBCrlf &_
								"<body scroll=auto onLoad='this.focus()'>"& VBCrlf &_
								"<script>resizeTo(660,(28 * "& Ubound(MyColumns) + 1 &")+120)</script>"& VBCrlf &_
								"<table border=0 cellpadding=1 cellspacing=1 width='100%' class=texto>"& VBCrlf &_
								"<form name='editSql' method='post' action='?Registros=True&Metodo=exec'>"& VBCrlf &_
								"<tr class=item>"& VBCrlf &_
								"<td>&nbsp;</td>"& VBCrlf &_
								"<td>&nbsp;</td>"& VBCrlf &_
								"<td><b>Tipo</b></td>"& VBCrlf &_
								"<td><b>Tamanho</b></td>"& VBCrlf &_
								"<td><b>Nulo</b></td>"& VBCrlf &_
								"</tr>"& VBCrlf 
'---------------------------------
						for i=LBound(MyColumns) to Ubound(MyColumns)
'---------------------------------
							if cStr(vetTypeColumns(i)) <> "int identity"  then 	
'---------------------------------
							strS =	strS &  "<tr>"& VBCrlf &_
												"<td width='25%' bgcolor='CFCFCF'>"& MyColumns(i) &"</td>"& VBCrlf &_
												"<td bgcolor='CFCFCF'><input id='box' type='text' name='"& MyColumns(i) &"' value='' size='30' maxlength='"& vetSizeColumns(i) &"' style='font-size: 11px;font-family:tahoma'></td>"& VBCrlf &_
												"<td class=item width='8%'>"& vetTypeColumns(i) &"</td>"& VBCrlf &_
												"<td class=item width='8%'>"& vetSizeColumns(i) &"</td>"& VBCrlf &_
												"<td class=item width='8%'>"& vetNullColumns(i) &"</td>"& VBCrlf &_
											"</tr>"& VBCrlf 
'---------------------------------
							else
'---------------------------------
							strS =	strS &  "<tr>"& VBCrlf &_
												"<td width='25%' bgcolor='CFCFCF'>"& MyColumns(i) &" <img src='chave.gif'></td>"& VBCrlf &_
												"<td width='30%' bgcolor='CFCFCF'><b>Auto Incremento</b></td>"& VBCrlf &_
												"<td class=item width='8%'>"& vetTypeColumns(i) &"</td>"& VBCrlf &_
												"<td class=item width='8%'>"& vetSizeColumns(i) &"</td>"& VBCrlf &_
												"<td class=item width='8%'>"& vetNullColumns(i) &"</td>"& VBCrlf &_
											"</tr>"& VBCrlf 
'---------------------------------
							end if
'---------------------------------
						next
'---------------------------------
				strS =	strS &	"<tr>"& VBCrlf &_
									"<td colspan=5>&nbsp;</td>"& VBCrlf &_
								"</tr>"& VBCrlf &_
								"<tr>"& VBCrlf &_
									"<td colspan=5 align=center><input type='submit' value='Incluir' style='font-size: 11px;font-family:tahoma'> &nbsp; &nbsp; <input type='button' onClick='javascript:window.close()' value='Fechar' style='font-size: 11px;font-family:tahoma'></td>"& VBCrlf &_
								"</tr>"& VBCrlf &_
								"<input type='hidden' name='sTable' value='"& sTable &"'>"& VBCrlf &_
								"<input type='hidden' name='sColumn' value='"& sColumn &"'>"& VBCrlf &_
								"<input type='hidden' name='sTypeColumns' value='"& sTypeColumns &"'>"& VBCrlf &_											
								"<input type='hidden' name='sId' value='"& sId &"'>"& VBCrlf &_
								"</table>"& VBCrlf 
'---------------------------------							
				response.Write strS
				Response.End
'---------------------------------
			Case "EDIT" '** pop de visualizacao
'---------------------------------
					Set Conx	=	Server.CreateObject("ADODB.Connection")
					Conx.Open strConn	
'---------------------------------
					MyColumns		= split(sColumn,",")						'* Colunas passadas pelo metodo GET
					sTypeColumns	= ColunasTabelaDescr(sTable,"sp_columns","TYPE_NAME")	'* Tipo de dados das colunas
					sSizeColumns	= ColunasTabelaDescr(sTable,"sp_columns","PRECISION")	'* Tamanho dos campos das colunas
					sNullColumns	= ColunasTabelaDescr(sTable,"sp_columns","IS_NULLABLE")	'* A coluna pode ser vazia
'---------------------------------
					vetTypeColumns = split(sTypeColumns,",",-1)
					vetSizeColumns = split(sSizeColumns,",",-1)
					vetNullColumns = split(sNullColumns,",",-1)
					fstColumns		=  vetTypeColumns(0)
'---------------------------------
					strSQL = "SELECT "& sColumn &" FROM "& sTable &" WHERE "& MyColumns(0) &" = "& PrepairSql(fstColumns,sId)					
					Set Rs = Conx.Execute(strSQL)
'---------------------------------
					IF not Rs.EOF then
'---------------------------------
						strS =	"<font face=tahoma size=2><b>Edição de registro da tabela '<font color=red>"& sTable &"</font>'</b></span><br><br>"& VBCrlf &_
								"<script language=JavaScript src='sys_contabiliza.js'></script>"& VBCrlf &_
								"<body scroll=auto onLoad='this.focus()'>"& VBCrlf &_
								"<script>resizeTo(660,(28 * "& Ubound(MyColumns) + 1 &")+120)</script>"& VBCrlf &_
								"<table border=0 cellpadding=1 cellspacing=1 width='100%' class=texto>"& VBCrlf &_
								"<form name='editSql' method='post' action='"& request.ServerVariables("script_name") &"?Registros=True&Metodo=UP'>"& VBCrlf  &_
								"<tr class=item>"& VBCrlf &_
								"<td>&nbsp;</td>"& VBCrlf &_
								"<td>&nbsp;</td>"& VBCrlf &_
							 	"<td><b>Tipo</b></td>"& VBCrlf &_
								"<td><b>Tamanho</b></td>"& VBCrlf &_
								"<td><b>Nulo</b></td>"& VBCrlf &_
								"</tr>"& VBCrlf 
'---------------------------------
						for i=LBound(MyColumns) to Ubound(MyColumns)
'---------------------------------
							if cStr(vetTypeColumns(i)) <> "int identity"  then 	
'---------------------------------
							strS =	strS &  "<tr>"& VBCrlf &_
												"<td width='25%' bgcolor='CFCFCF'>"& MyColumns(i) &"</td>"& VBCrlf &_
												"<td bgcolor='CFCFCF'><input id='box' type='text' name='"& MyColumns(i) &"' value='"& Rs(i) &"' size='30' maxlength='"& vetSizeColumns(i) &"' style='font-size: 11px;font-family:tahoma'></td>"& VBCrlf &_
												"<td class=item width='8%'>"& vetTypeColumns(i) &"</td>"& VBCrlf &_
												"<td class=item width='8%'>"& vetSizeColumns(i) &"</td>"& VBCrlf &_
												"<td class=item width='8%'>"& vetNullColumns(i) &"</td>"& VBCrlf &_
											"</tr>"& VBCrlf 
'---------------------------------
							else
'---------------------------------
							strS =	strS &  "<tr>"& VBCrlf &_
												"<td width='25%' bgcolor='CFCFCF'>"& MyColumns(i) &" <img src='chave.gif'></td>"& VBCrlf &_
												"<td width='30%' bgcolor='CFCFCF'><b>Auto Incremento</b></td>"& VBCrlf &_
												"<td class=item width='8%'>"& vetTypeColumns(i) &"</td>"& VBCrlf &_
												"<td class=item width='8%'>"& vetSizeColumns(i) &"</td>"& VBCrlf &_
												"<td class=item width='8%'>"& vetNullColumns(i) &_
												"<input id='box' type='hidden' name='"& MyColumns(i) &"' value='"& Rs(i) &"' size='30' maxlength='"& vetSizeColumns(i) &"' style='font-size: 11px;font-family:tahoma'>"&_
												"</td>"& VBCrlf &_
											"</tr>"& VBCrlf 
'---------------------------------
							end if
'---------------------------------
						next
'---------------------------------
							strS =	strS &	"<tr>"& VBCrlf &_
												"<td colspan=5>&nbsp;</td>"& VBCrlf &_
											"</tr>"& VBCrlf &_
											"<tr>"& VBCrlf &_
												"<td colspan=5 align=center><input type='submit' value='Atualizar' style='font-size: 11px;font-family:tahoma'> &nbsp; &nbsp; <input type='button' onClick='javascript:window.close()' value='Fechar' style='font-size: 11px;font-family:tahoma'></td>"& VBCrlf &_
											"</tr>"& VBCrlf &_
											"<input type='hidden' name='sTable' value='"& sTable &"'>"& VBCrlf &_
											"<input type='hidden' name='sColumn' value='"& sColumn &"'>"& VBCrlf &_
											"<input type='hidden' name='sTypeColumns' value='"& sTypeColumns &"'>"& VBCrlf &_											
											"<input type='hidden' name='sId' value='"& sId &"'>"& VBCrlf &_
											"</table>"& VBCrlf 
'---------------------------------
					ELSE
'---------------------------------
							strS =	"<font face=tahoma size=2><b><font color=red>Registro não encontrado</font></b></span><br><br><a href='javascript:window.close()'>:: Fechar ::</a>"& VBCrlf 
'---------------------------------
					END IF
'---------------------------------							
					response.Write strS
'---------------------------------
				Conx.Close
				Set Conx = nothing
'---------------------------------
					Response.End
'---------------------------------
'** Delete no banco de dados
'---------------------------------
			Case "DROP" '** exclusao do banco de dados
'---------------------------------
					MyColumns		= split(sColumn,",")						'* Colunas passadas pelo metodo GET
					sTypeColumns	= ColunasTabelaDescr(sTable,"sp_columns","TYPE_NAME")	'* Tipo de dados das colunas
					sTypeColumns	= Left(sTypeColumns,Len(sTypeColumns)-1)
					vetTypeColumns = split(sTypeColumns,",")					
'---------------------------------
					strSQL = "DELETE FROM  "& sTable &" WHERE "& MyColumns(0) &" = "& PrepairSql(vetTypeColumns(0),sId)
		'---------------------------------
				On Error Resume Next
		'---------------------------------
				Set Conx	=	Server.CreateObject("ADODB.Connection")
				Conx.Open strConn	
				Conx.Execute(strSQL)
				'---------------------------------
					IF Err.number <> 0 THEN
				'---------------------------------
						response.Write	"<script>"& VBCrlf &_
										"alert(""Ocorreram erros durante o processo de exclusão.\n\n "&_
										"Código do Erro: "& Err.number &"\n\n"&_
										"Descrição do Erro \n\n"&_
										Err.Description &_
										"\n\nA operação não pode ser concluída."");"&_
										"history.go(-1)"&_
										"</script>"
						response.End
				'---------------------------------	
					END IF
				'---------------------------------
				Conx.Close
				Set Conx = nothing
				'---------------------------------
				response.Write "<script>alert('Registro excluído com sucesso!');opener.location.reload();window.close()</script>"
				Response.End
'---------------------------------
'** Update no banco de dados
'---------------------------------
			Case "UP"
'---------------------------------
				vetCpo	= split(sColumn,",")
				sTypeColumns = request("sTypeColumns")
				vetTypeColumns = split(sTypeColumns,",") ' * colunas da tabelas
		'---------------------------------
		'** COLUNAS
		'---------------------------------
				strSQL = "UPDATE "& sTable &" SET "
		'---------------------------------
				for i=Lbound(vetCpo) to Ubound(vetCpo)
		'---------------------------------
					valor = Request(vetCpo(i))
		'---------------------------------
					IF valor <> "" THEN
		'---------------------------------
						if cStr(vetTypeColumns(i)) <> "int identity" then 	
		'---------------------------------
							strSQL = strSQL &" "& vetCpo(i) &" = "&  PrepairSql(vetTypeColumns(i),valor)
							IF i < Ubound(vetCpo) THEN strSQL = strSQL & ","
		'---------------------------------
						end if
		'---------------------------------
					END IF
		'---------------------------------
				next
		'---------------------------------
				IF Right(strSQL,1) = "," THEN strSQL = Left(strSQL,Len(strSQL)-1)
		'---------------------------------
				valor		= Request(vetCpo(0))
				fstColumn	= vetTypeColumns(0)
		'---------------------------------
				strSQL = strSQL & " WHERE "& vetCpo(0) &" = "& PrepairSql(fstColumn,valor)
		'---------------------------------
			On Error Resume Next
		'---------------------------------
				Set Conx	=	Server.CreateObject("ADODB.Connection")
				Conx.Open strConn	
				Conx.Execute(strSQL)
		'---------------------------------
			IF Err.number <> 0 THEN
		'---------------------------------
				response.Write	"<script>"& VBCrlf &_
								"alert(""Ocorreram erros durante o processo de atualização.\n\n "&_
								"Código do Erro: "& Err.number &"\n\n"&_
								"Descrição do Erro \n\n"&_
								Err.Description &_
								"\n\nA operação não pode ser concluída."");"&_
								"history.go(-1)"&_
								"</script>"
				response.End
		'---------------------------------	
			END IF
		'---------------------------------
				Conx.Close
				Set Conx = nothing
		'---------------------------------		
				response.Write "<script>alert('Registro atualizado com sucesso!');opener.location.reload();window.close();</script>"
				' parent.frames['registros'].location.reload();
				Response.End
'---------------------------------
'** Insert no banco de dados
'---------------------------------
			Case "EXEC"
'---------------------------------
				vetCpo	= split(sColumn,",")
				sTypeColumns = request("sTypeColumns")
				vetTypeColumns = split(sTypeColumns,",") ' * colunas da tabelas
		'---------------------------------
		'** COLUNAS
		'---------------------------------
				strSQL = "INSERT INTO "& sTable &" ("
		'---------------------------------
				for i=Lbound(vetCpo) to Ubound(vetCpo)
		'---------------------------------
					valor = Request(vetCpo(i))
		'---------------------------------
					IF valor <> "" THEN
		'---------------------------------
						if cStr(vetTypeColumns(i)) <> "int identity" then 	
		'---------------------------------
							strSQL = strSQL & vetCpo(i) &","
							strValue = strValue & PrepairSql(vetTypeColumns(i),valor) &","
		'---------------------------------
						end if
		'---------------------------------
					END IF
		'---------------------------------
				next
		'---------------------------------
				strSQL = Left(strSQL,Len(strSQL)-1)
				strValue = Left(strValue,Len(strValue)-1)
		'---------------------------------
				valor = Request(vetCpo(0))
		'---------------------------------
				strSQL = strSQL &	")"&_
									" VALUES "&_
									"("& strValue &")"
		'---------------------------------
			On Error Resume Next
		'---------------------------------
				Set Conx	=	Server.CreateObject("ADODB.Connection")
				Conx.Open strConn	
				Conx.Execute(strSQL)
		'---------------------------------
			IF Err.number <> 0 THEN
		'---------------------------------
				response.Write	"<script>"& VBCrlf &_
								"alert(""Ocorreram erros durante o processo de inclusão.\n\n "&_
								"Código do Erro: "& Err.number &"\n\n"&_
								"Descrição do Erro \n\n"&_
								Err.Description &_
								"\n\nA operação não pode ser concluída."");"&_
								"history.go(-1)"&_
								"</script>"
				response.End
		'---------------------------------	
			END IF
		'---------------------------------
				Conx.Close
				Set Conx = nothing
		'---------------------------------		
				response.Write "<script>alert('Registro incluido com sucesso!');window.close()</script>"
				Response.End
'---------------------------------
	End Select
'---------------------------------
	'## PAGINA ATUAL
	'---------------------------------
	Atual = "Registros="& Registros&"&sTable="& sTable &"&Sessao="& Sessao &"&Acao="& Acao 
'---------------------------------	
	IF sTable <> "" THEN
'---------------------------------
	'** COLUNAS DAS TABELAS
	'---------------------------------
		if Change = 1 then 
			IF OrderBy = "Asc" THEN OrderBy = "Desc" ELSE OrderBy = "Asc" 
		end if	
	'---------------------------------
		strS = ViewLines(sTable,OrderBy)
		vetColuna = split(Session("Colunas"),",")			
	'---------------------------------
		On Error Resume Next
	'---------------------------------
		Set Conx	=	Server.CreateObject("ADODB.Connection")
		Conx.Open strConn	
	'---------------------------------
		Set Rs = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM "& sTable &" ORDER BY "& sColumn &" "& OrderBy 		
	'---------------------------------
		Rs.Open strSQL, Conx, 1, 3
	'---------------------------------
	'** Caso ocorra um erro o sistema irá dar um alerta	
	'---------------------------------
		IF Err.number <> 0 THEN
	'---------------------------------
			response.Write	"<script>"& VBCrlf &_
							"alert(""Ocorreram erros durante a sua solicitação.\n\n "&_
							"Código do Erro: "& Err.number &"\n\n"&_
							"Descrição do Erro \n\n"&_
							Err.Description &_
							"\n\nA operação não pode ser concluída."");"&_
							"window.close()"&_
							"</script>"
			response.End
	'---------------------------------	
		END IF
	'---------------------------------		
		If Not Rs.EOF Then
	'---------------------------------
			Rs.PageSize =  15  '### QTD DE ITENS POR PAGINA
			Call QuebraPagina(0,0,0,1)
	'---------------------------------
			while not Rs.EOF AND cont < Rs.PageSize
	'---------------------------------
				strS =	strS & 	"<tr class="& cor &">"
	'---------------------------------
				for i=Lbound(vetColuna) to (Ubound(vetColuna)-1)
	'---------------------------------					
					IF i = 0 THEN idLinha = Replace(Rs(vetColuna(i)),VBCrlf,"") : sId = sId & Rs(vetColuna(i)) &","
					'---------------------------------
					IF IsNull(Rs(vetColuna(i))) = True THEN
						strS =	strS & 	"<td>&nbsp;</td>"& VBCrlf 
					Else
						strS =	strS & 	"<td>"& Replace(Rs(vetColuna(i)),VBCrlf,"<br>") &"</td>"& VBCrlf 
					eNd If
	'---------------------------------
				next				
	'---------------------------------
				strS =	strS & 	"<td align=center><a href=""javascript:sqlEditPop('"& idLinha &"','?"& atual &"&Metodo=edit')""><img src='editar.gif' border=0 alt='Editar o Registro "& idLinha &"'></a></td>"& VBCrlf &_
								"<td align=center><a href=""javascript:sqlDropPop('"& idLinha &"','?"& atual &"&Metodo=drop','"& sTable &"')""><img src='xis.gif' border=0 alt='Excluir o Registro "& idLinha &"'></td>"& VBCrlf &_
								"</tr>"
	'---------------------------------
				IF Cor = "sim" THEN Cor = "nao"  ELSE Cor = "sim" 
				cont = cont + 1
	'---------------------------------
			Rs.Movenext
			wend
	'---------------------------------
	'### VARS UTILIZADAS PARA A PAGINACAO
	'---------------------------------
		total = int(Rs.PageCount)
	'---------------------------------
			strS =	strS & 	"<input type='hidden' name='sColumn' value='"& Left(Session("Colunas"),Len(Session("Colunas"))-1) &"'>"& VBCrlf &_
							"<input type='hidden' name='sTable' value='"& sTable &"'>"& VBCrlf &_
							"<input type='hidden' name='idDisp' value='"& Left(sId,Len(sId)-1)  &"'>"& VBCrlf &_
							"</form>"& VBCrlf &_
							"<tr>"& VBCrlf &_
							"<td colspan='"& (Ubound(vetColuna)-1)+2 &"' align=center>"& VBCrlf &_
							QuebraPagina(Atual &"&sColumn="& sColumn &"&OrderBy="& OrderBy  ,Pagina,total,0) & VBCrlf &_
							"</td>"& VBCrlf &_
							"</tr>"& VBCrlf &_
							"</table>"& VBCrlf &_
							"<body onLoad='this.focus()'>"
	'---------------------------------
		ELSE
	'---------------------------------
			sColumn	= ColunasTabelaDescr(sTable,"sp_columns","COLUMN_NAME")	'* Tipo de dados das colunas
			sColumn	= Left(sColumn,Len(sColumn)-1)
	'---------------------------------
			strS =	strS & 	"<tr>"&_
							"<td colspan='"& (Ubound(vetColuna)-1) &"'>"&_
							"<br><br><Font color=red><b>Nenhum registro encontrado</b></font>"&_
							"</td>"&_
							"</tr>"&_
							"<input type='hidden' name='sTable' value='"& sTable &"'>"& VBCrlf &_
							"<input type='hidden' name='sColumn' value='"& sColumn &"'>"& VBCrlf &_
							"<input type='hidden' name='sTypeColumns' value='"& sTypeColumns &"'>"& VBCrlf &_											
							"</form>"&_
							"</table>"
	'---------------------------------
		END IF
	'---------------------------------	
		strS =	strS & 	"<br><center><a href='javascript:window.close()' style='text-decoration:none'>:: fechar ::</a></center>"
	'---------------------------------
		Response.Write strS
'---------------------------------
	END IF
'---------------------------------
Response.End
'---------------------------------
END IF
'---------------------------------%>
<script language=javascript src='sys_contabiliza.js'></script>
<script language=javascript>
function RunQuery(stTable,sStatus)
{
  var strLocation = '?Registros=True&sTable='+stTable
  var w = screen.width - 20;
  var h = 480;
  window.open(strLocation,'registros','width='+ w +',height='+ h +',top=0,left=0')
}
function QueryAnalyser(param)
{
	var strLocation = '?Registros=query_analyser'
	var w = screen.width - 20;
	var h = 480;
	
	if(param!=""){strLocation += '&metodo=exec&q=sp_helptext '+ param}
	var query_analyser = window.open(strLocation,'registros','width='+ w +',height='+ h +',top=0,left=0');
		query_analyser.focus()
}
function doCompact()
{
  var strLocation = '?Registros=compact_sql';
  var w = screen.width - 20;
  var h = 480;
  var query_analyser = window.open(strLocation,'registros','width='+ w +',height='+ h +',top=0,left=0')
  query_analyser.focus()
}
function doProperty()
{
  var strLocation = '?Registros=property_sql';
  var w = screen.width - 20;
  var h = 200;
  var query_analyser = window.open(strLocation,'registros','width='+ w +',height='+ h +',top=0,left=0')
  query_analyser.focus()
}
property_sql
</script>
<%
On Error Resume Next
'---------------------------------
'** Armazenando a coleção de objetos nos vetores
'---------------------------------
Set Conn	=	Server.CreateObject("ADODB.Connection")
Conn.Open strConn	
'---------------------------------
'** Pegando os objetos do banco de dados
'---------------------------------
	Tabela	= Conn.Execute("select name,crdate,type from sysobjects where xtype in ('U') and name not like '%dt%' and name not like '%sys%'  order by 3 desc,1").GetRows
	Session("tabela") = Tabela
'---------------------------------
	View	= Conn.Execute("select name,crdate,type from sysobjects where xtype in ('V') and name not like '%sys%' order by 3 desc,1").GetRows
	Session("View") = View
'---------------------------------
	Proced	= Conn.Execute("select name,crdate,type from sysobjects where xtype in ('P') and name not like '%dt_%' order by 3 desc,1").GetRows 
	Session("Proced") = Proced
'---------------------------------
Conn.Close
Set Conn = Nothing
'---------------------------------%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class=texto>
<tr> 
  <td class="titPagina">Análise (Status) do sistema - <a href='?logout=x'>SAIR</a></td>
</tr>
<tr> 
  <td class="linha10">&nbsp;</td>
</tr>
<tr>
  <td><a href='javascript:doProperty()'>
	<img src='base_sql.gif' border=0><%=Propriedades("","sp_spaceused")%> </a>
	- <a href='javascript:doCompact()'><b>Otimizar banco de dados</b></a> 
	- <a href="javascript:QueryAnalyser(0)"><b>Executar Query</b></a> 
	<br><span style='display:none;' id="B0"></span>
	<br>
		 Visualizar <a href='javascript:banco(0)'>(Off)</a>: <a href='javascript:banco(1)'>Tabelas</a> - <a href='javascript:banco(2)'>View</a> - <a href='javascript:banco(3)'>Procedure </a>
	<br>
	<br>
	<table border=0 width='100%' cellpadding=1 cellspacing=1 class=texto>
	<tr>
	 <td width='5%'>Tipo</td>	 
	 <td class=titulo width='25%'>Nome Objeto</td>
	 <td width='25%'>Data de Criação</td>
	 <td width='25%'>Propriedades</td>
	</tr>
	<tr>
	 <td colspan=4 class=linha5>&nbsp;</td>	 
	</tr>
	<tr>
	 <td colspan=4>	
	 <table border=0 width='100%' cellpadding=1 cellspacing=1 class=texto style='display:none;' id="B1">
	<%for i=Lbound(Tabela) to Ubound(Tabela,2)
	'---------------------------------	
		Response.Write	"<tr class="& cor &"><td align=center  onClick=""javascript:coluna('"&Tabela(0,i)&"')"" style='cursor:hand' title='Clique para ver as colunas desta tabela'><img src='base_table.gif'></td><td>" & Tabela(0,i) &"</td><td>" & Tabela(1,i) &"</a></td><td>"& Propriedades(Tabela(0,i),"sp_spaceused") &"</td> </tr>"
		'---------------------------------
		Response.Write	"<tr class="& cor &" id="&Tabela(0,i)&" style='display:none'><td colspan=4>"& ColunasTabela(Tabela(0,i),"sp_columns") &"</td></tr>"
		'---------------------------------
		Response.Write	"<tr><td colspan=4 style='line-height: 2px;'>&nbsp;</td></tr>"
	'---------------------------------
		IF Cor = "sim" THEN Cor = "nao"  ELSE Cor = "sim" 
	'---------------------------------
	  next
	'---------------------------------
	  Response.Write "</table><table border=0 width='100%' cellpadding=1 cellspacing=1 class=texto style='display:none;' id='B2'>"
	'---------------------------------
	for i=Lbound(View) to Ubound(View,2)
	'---------------------------------
		Response.Write	"<tr class="& cor &"><td align=center  onClick=""javascript:coluna('"&View(0,i)&"')"" style='cursor:hand' title='Clique para ver as colunas desta View (Visão)'><img src='base_view.gif'></td><td>"& View(0,i) &"</td><td>" & View(1,i) &"</td><td><a href=""javascript:RunQuery('"& View(0,i) &"','block')""> Visualizar registros</a></td></tr>"
		'---------------------------------
		Response.Write	"<tr class="& cor &" id="&View(0,i)&" style='display:none'><td colspan=4>"& ColunasTabela(View(0,i),"sp_columns") &" </td></tr>"
		'---------------------------------
		Response.Write	"<tr><td colspan=4 style='line-height: 2px;'>&nbsp;</td></tr>"
	'---------------------------------
		IF Cor = "sim" THEN Cor = "nao"  ELSE Cor = "sim" 
	'---------------------------------
	next
	'---------------------------------
		Response.Write "</table><table border=0 width='100%' cellpadding=1 cellspacing=1 class=texto style='display:none;' id='B3'>"
	'---------------------------------
	for i=Lbound(Proced) to Ubound(Proced,2)
	'---------------------------------
		Response.Write	"<tr class="& cor &" align=center><td><a href=""javascript:QueryAnalyser('"& Proced(0,i) &"')""><img src='base_procedure.gif' border='0'></a></td><td>"& Proced(0,i) &"</td><td colspan=2>" & Proced(1,i) &"</td></tr>"
	'---------------------------------
		IF Cor = "sim" THEN Cor = "nao"  ELSE Cor = "sim" 
	'---------------------------------
	next
	'---------------------------------
	Response.Write "</table>"
	'---------------------------------%>
	<hr size=1 color='#000000'>
	<div style='text-align:center;font-size: 10px;'>Análise do Banco de Dados | Desenvolvido por: <a href='#' onclick="parent.location='mailto:contrig@pop.com.br?Subject=Gerenciador SQL&body=Insira aqui o seu comentário&bcc=rogerio_silva@estadao.com.br'">Rogério Silva </a>  <br>&copy; 2001 - <%=year(date)%></div>
			</td>	 
		</tr>
	</table>
</td>
</tr>
</table> 
</body>
</html>