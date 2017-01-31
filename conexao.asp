<%
On Error Resume Next
'------------------------------------------------------------------------------------'
Set conexao= Server.CreateObject("ADODB.Connection")
'------------------------------------------------------------------------------------'
'** Prototipo
'------------------------------------------------------------------------------------'
	strServer	=	Trim(Request.Form("server"))
	strBanco	=	Trim(Request.Form("banco"))
	strUser		=	Trim(Request.Form("user"))
	strPass		=	Trim(Request.Form("senha"))
	if request("logout") <> "" then 
		response.write "<script>if(confirm('Logout executado com sucesso.\n\nDeseja realizar nova conexão?')==true){parent.location=top.location}else{alert('processo cancelado')}</script>"	
		Session.Abandon()
		response.end
	end if
	if trim(strServer) <> "" and trim(strBanco) <> "" and trim(strUser) <> "" then
		strConn	=  "Provider=SQLOLEDB.1;Data Source="& strServer &";Initial Catalog="& strBanco &";User ID="& strUser &";pwd="& strPass
	'	strConn	=  "Driver={SQL SERVER};Server="& strServer &";DataBase="& strBanco &";uid="& strUser &";pwd="& strPass	
		Session("strConn")	= strConn	
	end if
'------------------------------------------------------------------------------------'
	if trim(Session("strConn")) = "" then 
%>
<script>
	var server	=	prompt("Informe o nome do servidor","");
	var banco	=	prompt("Informe o nome do banco de dados","");
	var user	=	prompt("Informe o nome do usuário","");
	var senha	=	prompt("Informe a senha do usuário","");
	
	document.write("<form method='post' action='?' name='frm'><input type='hidden' name='server' value='"+ server +"'><input type='hidden' name='banco' value='"+ banco +"'><input type='hidden' name='user' value='"+ user +"'><input type='hidden' name='senha' value='"+ senha +"'></form><scri"+"pt>document.frm.submit()</scr"+"ipt>");
</script>
<% 
			response.end
	else
'------------------------------------------------------------------------------------'
	conexao.Open Session("strConn")
		if conexao.state = 0 then 
			response.write "<script>if(confirm('Não foi possível autenticação com as informações fornecidas.\n\nDeseja tentar novamente?')==true){parent.location=top.location}else{alert('processo cancelado')}</script>"
			response.wrte Err.Description &" --> "& Err.Number
			Session.Abandon()
			response.end
		end if
	end if
'------------------------------------------------------------------------------------'		
%>

