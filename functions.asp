<!--#include file="conexao.asp" -->
<%
	idPai 		= Session("CODPAI") : if idPai = 0 then idPai = Session("CODADM")		
	CODADM    	= idPai
'###################################################################################################################################	

	function Login(user,password)
			Session("ATEND")  = false  
			Session("ADMS")   = false
  		    Session("ADMA")   = false
		    Session("CODADM") = false
			Session("USU")    = false
			
			STRSQL = "SELECT * FROM USUARIOS WHERE LOGINUSER = '"& user &"' AND PASSUSER = '" & password & "'" 	
			Set RSLOGIN = Conexao.Execute(strSQL)
		
			if not RSLOGIN.EOF then
			   Session("LOGIN")		= true
			   Session("NAME")		= RSLOGIN("nameUser")
			   Session("LOGIN_USER")= user
			   Session("EMAIL")		= RSLOGIN("emailUser")
			   Session("CODUSER")	= RSLOGIN("coduser") : Session("ID") = Session("CODUSER")
			   Session("USER")		= RSLOGIN("loginUser")
			   Session("FONEUSER")	= RSLOGIN("foneUser")
			   Session("V_LOGIN")	= true			   
			   '**********************************************************'
			   Session("ADMS")		=	RSLOGIN("flagAdmS")
			   if RSLOGIN("flagAdmA") then  Session("ADMA")	= 	RSLOGIN("flagAdmA")  :  Session("CODADM") = RSLOGIN("CODADM")
			   if RSLOGIN("flagAtend") then  Session("ATEND")	= RSLOGIN("flagAtend")  else  Session("USU")	= true
			   '**********************************************************'
			   Session("CODADM")	= RSLOGIN("CODADM")	
			   Session("CODPAI")	= 0
			   '**********************************************************'
			   strSQL	=	" SELECT IDPAI FROM ADMINISTRADORA WHERE CODADM = " & RSLOGIN("CODADM") 
			   set Rx	=	Conexao.execute(strSQL)
			   if not Rx.eof then
			   '**********************************************************'
					if Rx("IDPAI") <> 0 then Session("CODPAI") = Rx("IDPAI")				
			   '**********************************************************'
			   end if
			   '**********************************************************'
			else
			   Session("LOGIN")		= false	  			   
			end if 

			set RSLOGIN = nothing
			'Se o usuário for um administrador é verificado as datas das ocorrências
			if Session("ADMS") = true or  Session("ADMA") = true then	 MensagemDiferencaDatas 5
            'set conexao=nothing
	end function
	
'###################################################################################################################################		
	
	function Limpa
	
			Session("LOGIN") = false	
			Session("ADMS") = false
			Session("ADMA") = false
    		Session("ATEND") = false
            Session("USU") = false
  		    Session("ANO")= false
	        Session("MES")= false
	        Session("DIA")= false
	        Session("SEGUNDOS")= false
	        Session("MINUTO")= false
	        Session("HORA")= false
			
	end function
	
'###################################################################################################################################		
	
	function AtualizaDataPrev(data,occur)            
						
		    STRSQL = "UPDATE OCORRENCIA SET DATAPREV = '"& data &"' WHERE CODOCCUR = '"& occur &"'" 	
			Set RSDATAPREV = Server.CreateObject("Adodb.Recordset")
			Set RSDATAPREV.ActiveConnection = Conexao
			Set RSDATAPREV = Conexao.Execute(strSQL)
            'Altero variável de Sessao para exibir mensagem de atualização
			Session("DATAPREV") = true
			
			set RSDATAPREV = nothing
			
	end function
	
'###################################################################################################################################		
	
	function InsereSequencia(codOccur,codPro,descri,data,hora)
		    
		'----------------------------------------------------------
		'** Modificado por: Rogério Silva (26/12/2006)
		'----------------------------------------------------------
		strSQL = "SELECT CODSTATUS FROM OCORRENCIA WHERE CODOCCUR = '" & codOccur & "'"
		Set TrySt = Conexao.Execute(strSQL)
		strStatus = TrySt(0)
		
		IF	strStatus = "1" OR  strStatus = "2" then 
			Session("SEQUENCIA") = false
			response.Write "<script>alert('O chamado foi finalizado, não será possivel enviar sua mensagem');window.close()</script>"
			exit function
			response.End
		END IF
	    
			STRSQL = "INSERT INTO SEQUENCIA VALUES "
			STRSQL = STRSQL & "("
			STRSQL = STRSQL & "'" & codOccur & "',"
			STRSQL = STRSQL & "'" & replace(descri,"'","") & "'," 
			STRSQL = STRSQL & "'" & data & "'," 
			STRSQL = STRSQL & "'" & hora & "'," 
			STRSQL = STRSQL & "" & codPro & ""
			STRSQL = STRSQL & ")"

			Set RSINSSEQUENCIA = Server.CreateObject("Adodb.Recordset")
			Set RSINSSEQUENCIA.ActiveConnection = Conexao
			Set RSINSSEQUENCIA = Conexao.Execute(strSQL)
            'Altero variável de Sessao para exibir mensagem de nova sequência e fechar a janela
			Session("SEQUENCIA") = true
			Session("FECHA") = true

			set RSINSSEQUENCIA = nothing
	
	end function
	
'###################################################################################################################################		
	
	function AtualizaStatus(codOccur,codStatus,data)
	
		if codStatus = 2 then
			hora = right("00"&hour(time),2) &":"& right("00"&minute(time),2)
			STRSQL = "UPDATE OCORRENCIA SET CODSTATUS = "& codStatus &", DATAFECHADO = '"& data &" "& hora &"' WHERE CODOCCUR = '"& codOccur &"'" 	
		else
			STRSQL = "UPDATE OCORRENCIA SET CODSTATUS = "& codStatus &" WHERE CODOCCUR = '"& codOccur &"'" 	
		end if	
			'Response.Write(strsql)
			'Response.End()
			Set RSATUASTATUS = Server.CreateObject("Adodb.Recordset")
			Set RSATUASTATUS.ActiveConnection = Conexao
			Set RSATUASTATUS = Conexao.Execute(strSQL)
			
			set RSATUASTATUS = nothing
			
	end function

'###################################################################################################################################		
function DefineAtendente(modulo)
		
				if strDescri <> "" then
								
					STRSQL="SELECT CODUSER FROM MODATEND WHERE CODMOD = " & modulo & ""
					Set RSATENDMOD = Server.CreateObject("Adodb.Recordset")
					Set RSATENDMOD.ActiveConnection = Conexao
					Set RSATENDMOD = Conexao.Execute(strSQL)
					 
					v1=0 'Valor
					v2=0  
					pass = true
					while not RSATENDMOD.eof 
					
						STRSQL="SELECT COUNT(CODUSER) AS QTD, CODPRO FROM OCORRENCIA WHERE CODPRO = " & RSATENDMOD("CODUSER") & " AND FLAGCLOSE = 0 AND FLAGCANCELA = 0 GROUP BY CODPRO"
						'Response.Write(strsql)
						Set RSATENDOCOR = Server.CreateObject("Adodb.Recordset")
						Set RSATENDOCOR.ActiveConnection = Conexao
						Set RSATENDOCOR = Conexao.Execute(strSQL)
						
						'Verifico se existe alguma ocorrencia para o usuário
						STRSQL="SELECT CODPRO FROM OCORRENCIA WHERE CODPRO = " & RSATENDMOD("CODUSER") & " AND FLAGCLOSE = 0 AND FLAGCANCELA = 0"
						'Response.Write(strsql)
						'Response.end
						Set RSATENDOCOR2 = Server.CreateObject("Adodb.Recordset")
						Set RSATENDOCOR2.ActiveConnection = Conexao
						Set RSATENDOCOR2 = Conexao.Execute(strSQL)
						
						
						'A primeira entrada é livre para guardar o valor do primeiro registro
						if pass = true then
							if not RSATENDOCOR2.eof then
							   v2 = RSATENDOCOR("QTD")
							   pass = false
							   coduser = RSATENDOCOR("CODPRO")
							else
							   v2 = 0
							   pass = false
							   coduser = RSATENDMOD("CODUSER")
							end if   
						else
							if not RSATENDOCOR2.eof then 'Se nenhuma ocorrencia é encontrada para o atendente a quantidade de registros é zero
							   v1 = RSATENDOCOR("QTD")
							else
							   v1 = 0
							end if      
							if v2 >= v1 then
							   v2 = v1
								if not RSATENDOCOR2.eof then 'Se nenhuma ocorrencia é encontrada para o atendente é registrado seu código de usuário
								   coduser = RSATENDOCOR("CODPRO")
								else
								   coduser = RSATENDMOD("CODUSER")
								end if     
							end if					
						end if	
											
						RSATENDMOD.movenext
					wend
						
						'Response.End()
			            set RSATENDMOD = nothing
						
						set RSATENDOCOR = nothing
			
						set RSATENDOCOR2 = nothing
				
					DefineAtendente = coduser
					
			end if
		
		end function
		
	function DefineAtendente2(modulo)
		
			if strDescri <> "" then
				'-----------------------------------------------------------------------'
				'** Pesquisa os atendentes que não tem chamado aberto
				'-----------------------------------------------------------------------'
					strSQL =		"	SELECT CODUSER  "&_
									"	FROM modAtend   "&_
									"	WHERE CODUSER NOT IN (  "&_
									"		SELECT CODPRO  "&_
									"		FROM OCORRENCIA   "&_
									"		WHERE   "&_
									"			CODPRO IN   "&_
									"				(SELECT CODUSER  FROM modAtend WHERE codMod = " & modulo & ")   "&_
									"			AND FLAGCLOSE = 0 AND FLAGCANCELA = 0   "&_
									"	)  "&_
									"	AND codMod = " & modulo & "  "&_
									"	GROUP BY CODUSER  "
					Set RsUsers	=	Conexao.Execute(strSQL)
				'-----------------------------------------------------------------------'
					if not RsUsers.eof then
				'-----------------------------------------------------------------------'
						Total	=	Conexao.Execute(strSQL).getRows()
				'-----------------------------------------------------------------------'
					else
				'-----------------------------------------------------------------------'
				'** Caso o módulo não tenha algum chamado em aberto, pega todos atend.
				'-----------------------------------------------------------------------'
						strSQL	=	"	SELECT CODUSER  FROM modAtend WHERE codMod = " & modulo 
						set At = Conexao.Execute(strSQL)
						if At.eof then
							DefineAtendente2 = DefineAtendente(modulo)
						else
							Total	=	Conexao.Execute(strSQL).getRows()
							Randomize  : NTemp = Int( Ubound(Total) * RND) + 1
							if NTemp > Ubound(Total,2) then NTemp = Ubound(Total,2)
							coduser = Total(0,NTemp)
							DefineAtendente2 = coduser							
						end if
				'-----------------------------------------------------------------------'
					end if
				'-----------------------------------------------------------------------'
			end if
		
		end function

'###################################################################################################################################	
	
	function InsereChamado(codOccur,codUser,codMod,data,hora,assunto,pathFile,descri,codPro,msg_erro,modulo_sistema,prioridade,CODADM)

	    strData = Cdate(date+5) 
		strAno = year(strData)
		strMes = month(strData)
		strDia = day(strData)
		strData = strAno & "/" & strMes & "/" & strDia
		'strData = strDia & "/" & strMes & "/" & strAno
		'Response.Write(data & " - " & strData)
		'Response.End()
		
		Replace msg_erro,Chr(39),Chr(34)
		
        STRSQL="INSERT INTO OCORRENCIA "
		STRSQL=STRSQL & "("
		STRSQL=STRSQL & "codOccur,"
		STRSQL=STRSQL & "codMod,"
		STRSQL=STRSQL & "codUser,"
		STRSQL=STRSQL & "codStatus,"
		STRSQL=STRSQL & "dataOccur,"
		STRSQL=STRSQL & "horaOccur,"
		STRSQL=STRSQL & "assuntoOccur,"
		STRSQL=STRSQL & "pathFile,"
		STRSQL=STRSQL & "descrOccur,"
		STRSQL=STRSQL & "codPro,"
		STRSQL=STRSQL & "msg_erro,"
		STRSQL=STRSQL & "modulo_sistema,"
		STRSQL=STRSQL & "datared,"		
		STRSQL=STRSQL & "prioridade,"		
		STRSQL=STRSQL & "CODADM"		
		STRSQL=STRSQL & ") "
		STRSQL=STRSQL & "values "
		STRSQL=STRSQL & "("
		STRSQL=STRSQL & "'" & codOccur & "',"
		STRSQL=STRSQL & "" & codMod & ","
		STRSQL=STRSQL & "" & codUser & ","
		STRSQL=STRSQL & "3,"
		STRSQL=STRSQL & "'" & data &" "& hora & "',"		
		STRSQL=STRSQL & "'" & hora & "',"
		STRSQL=STRSQL & "'" & replace(assunto,"'","") & "',"		
		STRSQL=STRSQL & "'" & pathFile & "',"
		STRSQL=STRSQL & "'" & replace(descri,"'","") & "',"		
		STRSQL=STRSQL & "" & codPro & ","
		STRSQL=STRSQL & "'" & replace(msg_erro,"'","") & "',"
		STRSQL=STRSQL & "'" & replace(modulo_sistema,"'","") & "',"
		STRSQL=STRSQL & "'" & replace(strData,"'","") & "',"	
		STRSQL=STRSQL & "'" & prioridade & "',"		
		STRSQL=STRSQL & CODADM 		
		STRSQL=STRSQL & ") "
'		Response.Write(strsql)
'		Response.End()
'------------------------------------------------------------------------------------------'
		Set RSINSCHAM = Conexao.Execute(strSQL) 

		Session("INSCHAMADO") = true     		
	
		set RSINSCHAM = nothing

		Host				= "localhost"
		Componente			= "CDONTS"
		Email				= "novochamado.SHD@unimeds.com.br"
		NomeEmail			= "SHD"
		Assunto_Email		= "SHD - Sistema HelpDesk: Chamado Nº "&codOccur&" - "& assunto
		
		STRSQL = "SELECT NAMEUSER,EMAILUSER,FLAGATEND FROM USUARIOS WHERE CODUSER = "&codUser		: 	Set RSEMAIL_USUARIO = Conexao.Execute(strSQL)		
		STRSQL = "SELECT NAMEUSER,EMAILUSER,FLAGATEND FROM USUARIOS WHERE CODUSER = "&codPro		:	Set RSEMAIL_ATENDENTE = Conexao.Execute(strSQL)		
		if not RSEMAIL_ATENDENTE.eof then
		'MENSAGEM PARA O USUÁRIO
		'==============================================================================================================================
			Mensagem	=	"<div style='font: 12px; font-family:Verdana, Arial, Helvetica, sans-serif;width:100%;border:1px #FFF solid;'>"& _
							"Prezado(a) "&  RSEMAIL_USUARIO(0) &", <br /><br />"&_
							"Você abriu um novo chamado com o assunto <b>"& assunto &"</b> e este chamado será  "&_
							"analisado pelo atendente <b>"&RSEMAIL_ATENDENTE(0)&"</b>.<br /><br /> "&_
							"A descrição inicial inserida foi: <br /><br /> "&_
							"<div style='border:1px solid #000;width:80%;padding: 8px; margin:10px;'>"&_
							Replace(descri,VBCrlf,"<br>")&_
							"</div>"&_
							"<br /><br />Atenciosamente,<br / ><br/ >"&_
							"Sistema de HelpDesk - SHD<br / ><br/ >"&_
							"<a href='http://"& Request.ServerVariables("Server_Name") &"'>Acesse o SHD agora.</a><br / ><br/ >"&_
							"<b>E-mail automático, favor não responder este e-mail.</b>"&_
							"</div>"
						
		EnviaEmail Host,Componente,Email,NomeEmail,RSEMAIL_USUARIO(1),Assunto_Email,Mensagem
		'==============================================================================================================================
		'MENSAGEM PARA O ATENDENTE
		'==============================================================================================================================	
			Mensagem	=	"<div style='font: 12px; font-family:Verdana, Arial, Helvetica, sans-serif;width:100%;border:1px #FFF solid;'>"& _
							"Prezado(a) "&  RSEMAIL_ATENDENTE(0) &", <br /><br />"&_
							"Você recebeu um novo chamado com o assunto <b>"& assunto &"</b> e este chamado foi  "&_
							"aberto pelo solicitante <b>"& RSEMAIL_USUARIO(0) &"</b>. <br /><br />"&_
							"A descrição inicial inserida foi: "&_
							"<div style='border:1px solid #000;width:80%;padding: 8px; margin:10px;'>"&_
							Replace(descri,VBCrlf,"<br>")&_
							"</div>"&_							
							"<br /><br />Atenciosamente,<br / ><br/ >"&_
							"Sistema de HelpDesk - SHD<br / ><br/ >"&_
							"<a href='http://"& Request.ServerVariables("Server_Name") &"'>Acesse o SHD agora.</a><br / ><br/ >"&_
							"<b>E-mail automático, favor não responder este e-mail.</b>"&_
							"</div>"									
		EnviaEmail Host,Componente,Email,NomeEmail,RSEMAIL_ATENDENTE(1),Assunto_Email,Mensagem						
		'==============================================================================================================================	
		end if
		
		set RSEMAIL_USUARIO	= nothing
		
		set RSEMAIL_ATENDENTE = nothing
		
	end function
	
'###################################################################################################################################	

    function MensagemConcluido(codOcorr,tipo_finalizado,descri_finalizado)
	    
		'SELECIONO O CODIGO DO USUÁRIO
	    STRSQL = "SELECT CODUSER,ASSUNTOOCCUR FROM OCORRENCIA WHERE CODOCCUR = '" & codOcorr & "'"
		Set RSMENSUSER = Conexao.Execute(strSQL)
		'=========================================================================================================================
		'MONTA SEQUÊNCIA
		'=========================================================================================================================
		STRSQL = "SELECT * FROM SEQUENCIA WHERE CODOCCUR = '" & codOcorr & "' ORDER BY CODSEQ ASC" 
		Set RSSEQUENCIA = Conexao.Execute(strSQL)	
		'=========================================================================================================================
		'MONTA SEQUÊNCIA
		'=========================================================================================================================
		STRSQL = "SELECT * FROM SEQUENCIA WHERE CODOCCUR = '" & strcodOccur & "' ORDER BY CODSEQ" 
		Set RSSEQUENCIA = Conexao.Execute(strSQL)	
		
		ix	= 1	:	max = Len(RSSEQUENCIA.RecordCount) + 1	:	cor = "sim" 
		
		while not RSSEQUENCIA.eof 
	
			 strMDescri = Replace(RSSEQUENCIA("descrSeq"),VBCrlf,"<br>")
			 strDataSeq = RSSEQUENCIA("dataSeq")
			 strHoraSeq = RSSEQUENCIA("horaSeq")
			 strcodUsuSeq = RSSEQUENCIA("CODUSER")
			 
			 STRSQL = "SELECT nameUser FROM USUARIOS WHERE CODUSER = '" & strcodUsuSeq & "'" 	
			 Set RSUSUSEQ = Server.CreateObject("Adodb.Recordset")
			 Set RSUSUSEQ.ActiveConnection = Conexao
			 Set RSUSUSEQ = Conexao.Execute(strSQL)
					
		strMontaSequencia = strMontaSequencia & "<div class='"& cor &"'><div align=center><i>Seq: "& MaxZeros(ix,max) &" - Data: "& RSSEQUENCIA("dataSeq") & " - Hora: " & Replace(RSSEQUENCIA("horaSeq"),"1/1/1900","") & " - Usuário: " & RSUSUSEQ("nameUser") &"</i></div><br>"& vbcrlf &_
												strMDescri &"</div><hr size=1>"& vbcrlf
		ix = ix + 1
		if cor = "sim" then cor = "nao" else cor = "sim" 
												
		RSSEQUENCIA.movenext
			
		wend		
		
		'=========================================================================================================================		

		'SELECIONO O CODIGO DA ADMINISTRADORA A QUAL PERTENCE O USUÁRIO
	    STRSQL = "SELECT CODADM FROM USUARIOS WHERE CODUSER = " & RSMENSUSER("codUser")  & ""
		Set RSMENSADM = Conexao.Execute(strSQL)

		'SELECIONO O ADMINISTRADOR DA ADMINISTRADORA
	    STRSQL = "SELECT CODUSER,NAMEUSER,EMAILUSER FROM USUARIOS WHERE FLAGADMA = 1 AND CODADM = " & RSMENSADM("CODADM")  & ""
		Set RSMENSADMIN_ADM = Conexao.Execute(strSQL)

		'SELECIONO OS ADMINISTRADORES DO SUPORTE
	    STRSQL = "SELECT CODUSER,NAMEUSER,EMAILUSER FROM USUARIOS WHERE FLAGADMS = 1"
		Set RSMENSADMIN_SUPORTE = Conexao.Execute(strSQL)
		
		'INSERE
		strMens="Chamado concluído"
		STRSQL = "INSERT INTO ALERT (codOccur,codUsu,mens) VALUES ('" & codOcorr & "'," & RSMENSUSER("codUser") & ",'" & strMens & "')"
		Set RSINSMENSUSER = Conexao.Execute(strSQL)
		
		'PODE EXISTIR MAIS DE UM ADMINISTRADOR
		while not RSMENSADMIN_ADM.eof 
			STRSQL = "INSERT INTO ALERT (codOccur,codAdm,mens) VALUES ('" & codOcorr & "'," & RSMENSADMIN_ADM("codUser") & ",'" & strMens & "')"
			Set RSINSMENSADMIN_ADM = Conexao.Execute(strSQL)
	
			'======================================================================================================================
				
			Host				= "localhost"
			Componente			= "CDONTS"
			Email				= "novochamado.SHD@unimeds.com.br"
			NomeEmail			= "SHD"
			Assunto_Email		= "SHD - Sistema HelpDesk: Chamado Nº "&codOcorr&" - "& RSMENSUSER("ASSUNTOOCCUR") & " chamado concluído."
			
			 Mensagem	= "<font size='2' face='Verdana, Arial, Helvetica, sans-serif'>"& _
						  "Prezado (a) "&RSMENSADMIN_ADM("NAMEUSER")&", <br><br> O chamado com o código <b>"&codOcorr&"</b> e com o assunto:<b>"& RSMENSUSER("ASSUNTOOCCUR") &"</b> , foi concluído por <b>"&Session("NAME")&"</b>" & _
						  "</b> e possui a seguinte descrição:<br><br>"& _
						  Replace(strMontaSequencia,VBCrlf,"<br>")& _
						  "<br><br>Atenciosamente,<br><br>Sistema de HelpDesk - SHD<br><br><b><div align=''>E-mail automático, favor não responder este e-mail.</div></b></font>"
										
			 EnviaEmail Host,Componente,Email,NomeEmail,RSMENSADMIN_ADM("EMAILUSER"),Assunto_Email,Mensagem
			 
 			'======================================================================================================================			
			
            RSMENSADMIN_ADM.movenext
		wend

		'PODE EXISTIR MAIS DE UM ADMINISTRADOR				
	    while not RSMENSADMIN_SUPORTE.eof
			STRSQL = "INSERT INTO ALERT (codOccur,codAdm,mens) VALUES ('" & codOcorr & "'," & RSMENSADMIN_SUPORTE("codUser") & ",'" & strMens & "')"
			Set RSINSMENSADMIN_SUPORTE = Server.CreateObject("Adodb.Recordset")
			Set RSINSMENSADMIN_SUPORTE.ActiveConnection = Conexao
			Set RSINSMENSADMIN_SUPORTE = Conexao.Execute(strSQL)
			'======================================================================================================================
			Host				= "localhost"
			Componente			= "CDONTS"
			Email				= "novochamado.SHD@unimeds.com.br"
			NomeEmail			= "SHD"
			Assunto_Email		= "SHD - Sistema HelpDesk: Chamado Nº "&codOcorr&" - "& RSMENSUSER("ASSUNTOOCCUR") & " chamado concluído."
			
			Mensagem	=	"<div style='font: 12px; font-family:Verdana, Arial, Helvetica, sans-serif;width:100%;border:1px #FFF solid;'>"& _
							"Prezado (a) "& RSMENSADMIN_SUPORTE("NAMEUSER")&", <br /><br />"&_
							"O chamado com o código <b>"& codOcorr &"</b> que tem o "&_
							"assunto <b>"&  RSMENSUSER("ASSUNTOOCCUR") &"</b>, foi concluído pelo atendente <b>"& Session("NAME") &"</b><br /><br />"&_
							"A descrição de fechamento foi esta: "&_
							"<div style='border:1px solid #000;width:80%;padding: 8px; margin:10px;'>"&_
							Replace(strMontaSequencia,VBCrlf,"<br>")&_
							"</div>"&_							
							"<br /><br />Atenciosamente,<br / ><br/ >"&_
							"Sistema de HelpDesk - SHD<br / ><br/ >"&_
							"<a href='http://"& Request.ServerVariables("Server_Name") &"'>Acesse o SHD agora.</a><br / ><br/ >"&_
							"<b>E-mail automático, favor não responder este e-mail.</b>"&_
							"</div>"			
			
			 EnviaEmail Host,Componente,Email,NomeEmail,RSMENSADMIN_SUPORTE("EMAILUSER"),Assunto_Email,Mensagem			 
 			'======================================================================================================================			
            RSMENSADMIN_SUPORTE.movenext
		wend
			if tipo_finalizado = "" then tipo_finalizado = 0
			data = year(date) &"/"& month(date) &"-"& day(date)  &" "& hour(time) &":"& minute(time)
		'----------------------------------------------------------
		'** Modificado por: Rogério Silva (21/02/2007)
		'----------------------------------------------------------
		on error resume next
		'------------------------------------'
			STRSQL = "UPDATE OCORRENCIA SET FLAGCLOSE = 1, DATAFECHADO = '"& data &"', TIPO_FINALIZADO = "&tipo_finalizado&", DESCRI_FINALIZADO = '"&descri_finalizado&"' WHERE CODOCCUR = '" & codOcorr & "'" 
			Set RSATUOCORR = Server.CreateObject("Adodb.Recordset")
			Set RSATUOCORR.ActiveConnection = Conexao
			Set RSATUOCORR = Conexao.Execute(strSQL)
		'------------------------------------'
			if err.number <> 0 then
		'------------------------------------'
				response.Write	"<font face=tahoma size=2>Erro durante ao executar o comando:<br><br>"&_
								"<pre>"& strSQL &"</pre></font>"
				response.End
		'------------------------------------'
			end if
		'------------------------------------'
		set RSATUOCORR = nothing									
		
		set RSMENSADMIN_SUPORTE = nothing
		
		set RSINSMENSADMIN_SUPORTE = nothing
		
		set RSINSMENSADMIN_ADM = nothing
		
		set RSINSMENSUSER = nothing
		
		set RSMENSADMIN_SUPORTE = nothing
		
		set RSMENSADMIN_ADM = nothing												
		
		set RSMENSADM = nothing		
		
		set RSMENSUSER = nothing
		
		set RSSEQUENCIA = nothing		
		
	end function
	
'####################################################################################################################################	
	
	function VoltaLogin
	
	    if Session("CODUSER") = "" then
		
			str="<script language=JavaScript>"
			str=str & "window.parent.location='default.asp';"
			str=str & "</script>"
			
			VoltaLogin = str
			
		end if
	
	end function
'###################################################################################################################################	
	
	function MensagemNovoChamado(codOcorr,assuntoOccur,descrOccur)
	
	    strMens	=	"Novo chamado"
   	    STRSQL	=	"SELECT U.CODADM FROM USUARIOS U, ADMINISTRADORA A, OCORRENCIA O WHERE U.CODUSER = " & Session("CODUSER") & " AND O.CODOCCUR = '" & Session("CODOCCUR") & "' GROUP BY U.CODADM"
		Set RSADMNOVO_CHAMADO = Conexao.Execute(strSQL)
		
   	    STRSQL = "SELECT codUser,EMAILUSER,NAMEUSER FROM USUARIOS WHERE CODADM = " & RSADMNOVO_CHAMADO("CODADM") & " AND FLAGADMA = 1 OR FLAGADMS = 1"
		Set RSADMIN_NOVO_CHAMADO = Conexao.Execute(strSQL)
		
   	    STRSQL = "SELECT U.NAMEUSER FROM OCORRENCIA O,USUARIOS U WHERE CODOCCUR = '"&codOcorr&"' AND O.CODPRO = U.CODUSER"
		Set RSCHAMADO = Conexao.Execute(strSQL)
		
		Host				= "localhost"
		Componente			= "CDONTS"
		Email				= "novochamado.SHD@unimeds.com.br"
		NomeEmail			= "SHD"
		Assunto_Email		= "SHD - Sistema HelpDesk: Chamado Nº "&codOcorr&" - "& assuntoOccur
		
		while not RSADMIN_NOVO_CHAMADO.eof		
			STRSQL = "INSERT INTO ALERT (CODOCCUR,CODADM,MENS) VALUES ('" & codOcorr & "'," & RSADMIN_NOVO_CHAMADO("codUser") & ",'" & strMens & "')"
			
			Set RSINSMENSADMIN = Conexao.Execute(strSQL)
			
			
			Mensagem	=	"<div style='font: 12px; font-family:Verdana, Arial, Helvetica, sans-serif;width:100%;border:1px #FFF solid;'>"& _
							"Prezado (a) "&RSADMIN_NOVO_CHAMADO(2)&", <br /><br />"&_
							"Um novo chamado foi aberto com o código <b>"& codOcorr &"</b>.<br /><br />"&_
							"O assunto deste chamado é <b>"& assuntoOccur &"</b>, a solicitação foi gerada pelo usuário(a) "&_
							"<b>"& Session("NAME")&"</b> e será atendido pelo atendente <b>"&RSCHAMADO(0)&"</b>.<br /><br />"&_
							"A descrição inicial foi esta: "&_
							"<div style='border:1px solid #000;width:80%;padding: 8px; margin:10px;'>"&_
							replace(descrOccur&" ",VBCrlf,"<br>") &"</div>"& _
							"<br /><br />Atenciosamente,<br / ><br/ >"&_
							"Sistema de HelpDesk - SHD<br / ><br/ >"&_
							"<b>E-mail automático, favor não responder este e-mail.</b>"&_
							"</div>"
							
			EnviaEmail Host,Componente,Email,NomeEmail,RSADMIN_NOVO_CHAMADO(1),Assunto_Email,Mensagem
		RSADMIN_NOVO_CHAMADO.movenext
		wend
		
		set RSINSMENSADMIN = nothing		
		set RSADMIN_NOVO_CHAMADO = nothing		
		set RSADMNOVO_CHAMADO = nothing		
		set RSCHAMADO = nothing	
			    	
	end function
	
'###################################################################################################################################
	
	function MensagemRespostaChamado(codOcorr,assuntoOccur,descrOccur,nomeUsuario)
	
		'----------------------------------------------------------
		'** Modificado por: Rogério Silva (26/12/2006)
		'----------------------------------------------------------
		strSQL = "SELECT CODSTATUS FROM OCORRENCIA WHERE CODOCCUR = '" & codOcorr & "'"
		Set TrySt = Conexao.Execute(strSQL)
		strStatus = TrySt(0)
		
		IF	strStatus = "1" OR  strStatus = "2" then 
			Session("SEQUENCIA") = false
			response.Write "<script>alert('O chamado foi finalizado, não será possivel enviar sua mensagem');window.close()</script>"
			exit function
			response.End
		ELSE
		
			strMens="Nova resposta para o chamado"
			
   			STRSQL = "SELECT U.CODADM FROM USUARIOS U, ADMINISTRADORA A, OCORRENCIA O WHERE U.CODUSER = " & Session("CODUSER") & " AND O.CODOCCUR = '" & codOcorr & "' GROUP BY U.CODADM"
			Set RSADMNOVO_CHAMADO = Server.CreateObject("Adodb.Recordset")
			Set RSADMNOVO_CHAMADO.ActiveConnection = Conexao
			Set RSADMNOVO_CHAMADO = Conexao.Execute(strSQL)
			
   			STRSQL = "SELECT codUser,EMAILUSER,NAMEUSER FROM USUARIOS WHERE CODADM = " & RSADMNOVO_CHAMADO("CODADM") & " AND FLAGADMA = 1 OR FLAGADMS = 1"
			Set RSADMIN_NOVO_CHAMADO = Server.CreateObject("Adodb.Recordset")
			Set RSADMIN_NOVO_CHAMADO.ActiveConnection = Conexao
			Set RSADMIN_NOVO_CHAMADO = Conexao.Execute(strSQL)
			
   			STRSQL = "SELECT U.NAMEUSER FROM OCORRENCIA O,USUARIOS U WHERE CODOCCUR = '"&codOcorr&"' AND O.CODPRO = U.CODUSER"
			Set RSCHAMADO = Server.CreateObject("Adodb.Recordset")
			Set RSCHAMADO.ActiveConnection = Conexao
			Set RSCHAMADO = Conexao.Execute(strSQL)
			
			Host				= "localhost"
			Componente			= "CDONTS"
			Email				= "novochamado.SHD@unimeds.com.br"
			NomeEmail			= "SHD"
			Assunto_Email		= "SHD - Sistema HelpDesk: Chamado Nº "&codOcorr&" - "& assuntoOccur
			
			while not RSADMIN_NOVO_CHAMADO.eof
			
			STRSQL = "INSERT INTO ALERT (codOccur,codAdm,mens) VALUES ('" & codOcorr & "'," & RSADMIN_NOVO_CHAMADO("codUser") & ",'" & strMens & "')"
			Set RSINSMENSADMIN = Server.CreateObject("Adodb.Recordset")
			Set RSINSMENSADMIN.ActiveConnection = Conexao
			Set RSINSMENSADMIN = Conexao.Execute(strSQL)

			
			Mensagem	=	"<div style='font: 12px; font-family:Verdana, Arial, Helvetica, sans-serif;width:100%;border:1px #FFF solid;'>"& _
							"Prezado (a) "& RSADMIN_NOVO_CHAMADO(2) &", <br /><br />"&_
							"Uma nova resposta foi inserida no chamado de código <b>"& codOcorr &"</b> que possui o "&_
							"assunto <b>"&  assuntoOccur  &"</b>, esta solicitação foi gerada pelo usuário(a)  <b>"& nomeUsuario &"</b> <br /><br />"&_
							" e está sendo atendido pelo atendente <b>"& RSCHAMADO(0) &"</b>."&_
							"A descrição de fechamento foi esta: "&_
							"<div style='border:1px solid #000;width:80%;padding: 8px; margin:10px;'>"&_
							replace(descrOccur&" ",VBCrlf,"<br>") &"</div>"& _
							"<br /><br />Atenciosamente,<br / ><br/ >"&_
							"Sistema de HelpDesk - SHD<br / ><br/ >"&_
							"<a href='http://"& Request.ServerVariables("Server_Name") &"'>Acesse o SHD agora.</a><br / ><br/ >"&_
							"<b>E-mail automático, favor não responder este e-mail.</b>"&_
							"</div>"
							
			EnviaEmail Host,Componente,Email,NomeEmail,RSADMIN_NOVO_CHAMADO(1),Assunto_Email,Mensagem		 
			RSADMIN_NOVO_CHAMADO.movenext
			 
			wend
			
			set RSINSMENSADMIN = nothing		
			set RSADMIN_NOVO_CHAMADO = nothing		
			set RSADMNOVO_CHAMADO = nothing		
			set RSCHAMADO = nothing	
		END IF
	end function	

'###################################################################################################################################
	
	function MensagemDiferencaDatas(Dias)
		
		'Response.Write("Atraso")
		
	    strData = Cdate(date) 
		strAno = year(strData)
		strMes = month(strData)
		strDia = day(strData)
		strData = strAno & "/" & strMes & "/" & strDia
		
		strMens = "Chamado não concluído há mais de 5 dias"
   	    STRSQL = "SELECT CODUSER,CODPRO,CODOCCUR FROM OCORRENCIA WHERE DATARED <= '" & strData & "' AND FLAGCLOSE = 0 AND CODOCCUR NOT IN (SELECT CODOCCUR FROM ALERT)"
		'Response.Write(strsql)
		'Response.End()
		Set RSUSU_ATEND = Conexao.Execute(strSQL)
	    
		if not RSUSU_ATEND.eof then
		
			while not RSUSU_ATEND.eof
			
				STRSQL = "SELECT * FROM ALERT WHERE CODOCCUR = '" & RSUSU_ATEND("CODOCCUR") & "' AND FLAGALERTDATE <> 0"
				'Response.Write(strsql)
				Set RSALERT = Conexao.Execute(strSQL)
					
			 'Se encontrar alguma ocorrencia com atraso é enviado o alerta.	
			 if (not RSUSU_ATEND.eof) and (RSALERT.eof) then
		
					STRSQL = "INSERT INTO ALERT (codOccur,codUsu,mens,flagAtraso) VALUES ('" & RSUSU_ATEND("CODOCCUR") & "'," & RSUSU_ATEND("CODUSER") & ",'" & strMens & "',1)"
					Conexao.Execute(strSQL)

					STRSQL = "SELECT U.CODADM FROM USUARIOS U WHERE U.CODUSER = " & RSUSU_ATEND("CODUSER") & ""
					Set RSADM = Conexao.Execute(strSQL)
				   
					STRSQL = "SELECT CODUSER,NAMEUSER,EMAILUSER FROM USUARIOS WHERE CODADM = " & RSADM("CODADM") & " AND FLAGADMA = 1 OR FLAGADMS = 1"
					Set RSADMIN_ATRASO_CHAMADO = Conexao.Execute(strSQL)
					set RSADM = nothing
					
					'=========================================================================================================================
					'MONTA SEQUÊNCIA
					'=========================================================================================================================
					
					STRSQL = "SELECT * FROM SEQUENCIA WHERE CODOCCUR = '" & RSUSU_ATEND("CODOCCUR") & "' ORDER BY CODSEQ DESC" 
					Set RSSEQUENCIA = Server.CreateObject("Adodb.Recordset")
					Set RSSEQUENCIA.ActiveConnection = Conexao
					Set RSSEQUENCIA = Conexao.Execute(strSQL)	
					
					while not RSSEQUENCIA.eof 
				
						 strMDescri = RSSEQUENCIA("descrSeq")
						 strDataSeq = RSSEQUENCIA("dataSeq")
						 strHoraSeq = RSSEQUENCIA("horaSeq")
						 strcodUsuSeq = RSSEQUENCIA("CODUSER")
						 
						 STRSQL = "SELECT nameUser FROM USUARIOS WHERE CODUSER = '" & strcodUsuSeq & "'" 	
						 Set RSUSUSEQ = Server.CreateObject("Adodb.Recordset")
						 Set RSUSUSEQ.ActiveConnection = Conexao
						 Set RSUSUSEQ = Conexao.Execute(strSQL)
								
						strMontaSequencia = strMontaSequencia & "=====================================================================================" & vbcrlf 
						strMontaSequencia = strMontaSequencia & RSSEQUENCIA("dataSeq") & " - " & Replace(RSSEQUENCIA("horaSeq"),"1/1/1900","") & " - " & RSUSUSEQ("nameUser") & vbcrlf
						strMontaSequencia = strMontaSequencia & strMDescri & vbcrlf
						RSSEQUENCIA.movenext
						
					wend							

					'=========================================================================================================================
					
					while not RSADMIN_ATRASO_CHAMADO.eof
			
						 STRSQL = "INSERT INTO ALERT (codOccur,codAdm,mens,flagAtraso) VALUES ('" & RSUSU_ATEND("CODOCCUR") & "'," & RSADMIN_ATRASO_CHAMADO("codUser") & ",'" & strMens & "',1)"
						 Set RSINSMENSADMIN = Server.CreateObject("Adodb.Recordset")
						 Set RSINSMENSADMIN.ActiveConnection = Conexao
						 Set RSINSMENSADMIN = Conexao.Execute(strSQL)
						 
				   	     STRSQL = "SELECT NAMEUSER FROM USUARIOS WHERE CODUSER = " & RSUSU_ATEND("CODUSER") 
						 Set RSNOME_USU = Server.CreateObject("Adodb.Recordset")
						 Set RSNOME_USU.ActiveConnection = Conexao
						 Set RSNOME_USU = Conexao.Execute(strSQL)
						 
				   	     STRSQL = "SELECT NAMEUSER FROM USUARIOS WHERE CODUSER = " & RSUSU_ATEND("CODPRO") 
						 Set RSNOME_ATEND = Server.CreateObject("Adodb.Recordset")
						 Set RSNOME_ATEND.ActiveConnection = Conexao
						 Set RSNOME_ATEND = Conexao.Execute(strSQL)						 
						 						 
						 Host				= "localhost"
						 Componente			= "CDONTS"
						 Email				= "novochamado.SHD@unimeds.com.br"
						 NomeEmail			= "SHD"
						 Assunto_Email		= "SHD - Sistema HelpDesk: Chamado Nº "&RSUSU_ATEND("CODOCCUR")&" - "& assuntoOccur &" não concluído há 5 dias"						 
						 
						 
						Mensagem	=	"<div style='font: 12px; font-family:Verdana, Arial, Helvetica, sans-serif;width:100%;border:1px #FFF solid;'>"& _
										"Prezado (a) "& RSADMIN_ATRASO_CHAMADO("NAMEUSER") &", <br /><br />"&_
										"O chamado de código <b>"& RSUSU_ATEND("CODOCCUR") &"</b> que possui o "&_
										"assunto <b>"&  assuntoOccur  &"</b>, esta solicitação foi gerada pelo usuário(a)  <b>"& RSNOME_USU("NAMEUSER") &"</b> <br /><br />"&_
										" e está sendo atendido pelo atendente <b>"& RSNOME_ATEND("NAMEUSER") &"</b>."&_
										"A descrição de fechamento foi esta: "&_
										"<div style='border:1px solid #000;width:80%;padding: 8px; margin:10px;'>"&_
										replace(strMontaSequencia&" ",VBCrlf,"<br>") &"</div>"& _										
										"<br /><br />Atenciosamente,<br / ><br/ >"&_
										"Sistema de HelpDesk - SHD<br / ><br/ >"&_
										"<a href='http://"& Request.ServerVariables("Server_Name") &"'>Acesse o SHD agora.</a><br / ><br/ >"&_
										"<b>E-mail automático, favor não responder este e-mail.</b>"&_
										"</div>"						 
										
						EnviaEmail Host,Componente,Email,NomeEmail,RSADMIN_ATRASO_CHAMADO("EMAILUSER"),Assunto_Email,Mensagem						 
						
					 	RSADMIN_ATRASO_CHAMADO.movenext
						
					wend
					
					set RSADMIN_ATRASO_CHAMADO = nothing
					
					set RSINSMENSADMIN = nothing					
					
					STRSQL = "UPDATE ALERT SET flagAlertDate = 1 WHERE CODOCCUR = '" & RSUSU_ATEND("CODOCCUR") & "' AND FLAGALERTDATE = 0"
					Set RSATU_ALERT = Server.CreateObject("Adodb.Recordset")
					Set RSATU_ALERT.ActiveConnection = Conexao
					Set RSATU_ALERT = Conexao.Execute(strSQL)

			 end if
		
				RSUSU_ATEND.movenext
				
				strMontaSequencia = ""
				
			wend	 

        end if

		set RSUSU_ATEND = nothing
		
		set RSATU_ALERT = nothing
		
		set RSALERT = nothing					

		set RSNOME_USU = nothing			  		   
		
		set RSNOME_ATEND = nothing		
		
	end function
	
'###################################################################################################################################	
	function SelOcorrencias(entrada, arquivo)
	'On error resume next
'-----------------------------------------------------------------------------------------'
	strSQL	=		" SELECT DISTINCT  o.assuntoOccur AS assuntoOccur, o.codOccur, o.codMod, o.codUser, o.codStatus, o.codSequence, o.dataOccur, "&_
					" o.horaOccur, o.pathFile, o.dataPrev, o.codPro, o.flagClose, "&_
					" o.flagUsuWeb, o.flagCancela, o.tipo_finalizado, o.modulo_sistema, o.dataFechado, o.dataRed, o.prioridade, "&_
					" ME.NOME AS Area, M.NAMEMOD AS Modulo	"&_
					" FROM  ocorrencia as o	 "&_
					" INNER JOIN MODULO as M ON o.CODMOD = M.CODMOD "&_
					" INNER JOIN MODULO_ESPECIFICO as ME ON M.IDMODULOESPEC = ME.IDMODULOESPEC "&_														
					" WHERE  (o.flagClose = 0) AND (o.flagCancela = 0) "
'-----------------------------------------------------------------------------------------'
       'Verifico se é um atendente ou usuário
 		if arquivo = "trans" then
'-----------------------------------------------------------------------------------------'
			arquivo = "sys.asp"
'-----------------------------------------------------------------------------------------'			
			if Session("ADMS") = true  then
'-----------------------------------------------------------------------------------------'
				strSQL	=	strSQL	& " and o.CODADM IN  (SELECT CODADM FROM ADMINISTRADORA WHERE CODADM = "&  IDPAI &" OR IDPAI ="&  IDPAI &"  ) " &_
							" ORDER BY DATAOCCUR DESC "
				'Set RSCHAMADOS = Conexao.Execute(strSQL)
				F_Chamado = True
'-----------------------------------------------------------------------------------------'
			else
'-----------------------------------------------------------------------------------------'
				strSQL	=	strSQL	&	" and o.CODADM IN  (SELECT CODADM FROM ADMINISTRADORA WHERE CODADM = "&  IDPAI &" OR IDPAI ="&  IDPAI &" ) " &_
										" AND CODPRO = " & Session("CODUSER") &_
										" ORDER BY DATAOCCUR DESC "
				'Set RSCHAMADOS = Conexao.Execute(strSQL)
				F_Chamado = True
'-----------------------------------------------------------------------------------------'
			end if
'-----------------------------------------------------------------------------------------'
		else
'-----------------------------------------------------------------------------------------'		
			if (Session("ATEND") = false and Session("ADMS") = false and Session("ADMA") = false) then
'				response.write "1"		
				'Seleciono as ocorrências	
'-----------------------------------------------------------------------------------------'	
				strSQL	=	strSQL	& 	" and  o.CODADM IN  ( (SELECT CODADM FROM ADMINISTRADORA WHERE (CODADM = "&  Session("CODADM") &" OR IDPAI = "&  IDPAI &")) ) "&_
										" AND CODUSER = " & Session("CODUSER") &_
										" OR codOccur IN (SELECT codOccur FROM  ocorrencia WHERE CODPRO = " & Session("CODUSER") &" AND flagClose = 0 AND flagCancela = 0)" &_
										" ORDER BY DATAOCCUR DESC "
'-----------------------------------------------------------------------------------------'
'				Response.Write(strsql)
'				response.End				
				'Set RSCHAMADOS = Conexao.Execute(strSQL)
				F_Chamado = True
			elseif Session("ATEND") = true then
'				response.Write "2"
				'Seleciono as ocorrências	
'-----------------------------------------------------------------------------------------'
				strSQL	=	strSQL	&	" and o.CODADM IN  (SELECT CODADM FROM ADMINISTRADORA WHERE CODADM = "&  IDPAI &" OR IDPAI = "&  IDPAI &") AND CODPRO = " & Session("CODUSER") &_
										" OR codOccur IN (SELECT codOccur FROM  ocorrencia WHERE CODUSER = " & Session("CODUSER") &" AND flagClose = 0 AND flagCancela = 0)" &_
										" ORDER BY DATAOCCUR DESC "
'-----------------------------------------------------------------------------------------'
'				Response.Write(strsql)
'				response.End
				'Set RSCHAMADOS = Conexao.Execute(strSQL)
				F_Chamado = True
			elseif Session("ADMS") = true  then
'				response.write "3"
				'Seleciono as ocorrências	
'-----------------------------------------------------------------------------------------'
				strSQL	=	strSQL	&	" and o.CODADM IN  (  (SELECT CODADM FROM ADMINISTRADORA WHERE (CODADM = "&  Session("CODADM") &" OR IDPAI = "&  IDPAI &")) ) " &_
										" ORDER BY DATAOCCUR DESC "
'-----------------------------------------------------------------------------------------'
'				Response.Write(strsql)
'				response.End
				'Set RSCHAMADOS = Conexao.Execute(strSQL)
				F_Chamado = True
			elseif Session("ADMA") = true  then
'				response.write "4"
				'Seleciono as ocorrências	
'-----------------------------------------------------------------------------------------'
				strSQL	=	strSQL	& 	" and o.CODADM IN  ("&  Session("CODADM") &") " &_
										" ORDER BY DATAOCCUR DESC "
'-----------------------------------------------------------------------------------------'
				'Set RSCHAMADOS = Conexao.Execute(strSQL)
				F_Chamado = True				
			end if
'-----------------------------------------------------------------------------------------'
		end if
'---------------------------------------------------------------------------------------------'		
'		response.write strSQL
		Set RsChamados = Conexao.Execute(strSQL)
		if not RsChamados.eof then
			SelOcorrencias = Conexao.Execute(strSQL).GetRows()		
		else
			SelOcorrencias = true
		end if
'---------------------------------------------------------------------------------------------'		
	end function

'###################################################################################################################################
	
	function SelAlerts
	
			if Session("ADMS") = true or Session("ADMA") = true then
		
				STRSQL = "SELECT * FROM ALERT WHERE FLAGREAD = 0 AND CODADM = " & Session("CODUSER") 
				'Response.Write(strsql)
				'Response.End()
				
				Set RSALERTS = Server.CreateObject("Adodb.Recordset")
				Set RSALERTS.ActiveConnection = Conexao
				Set RSALERTS = Conexao.Execute(strSQL)
		
			elseif (Session("ADMS") = true or Session("ADMA") = true) or (Session("ATEND") = false and Session("ADMS") = false and Session("ADMA") = false) then
			
				STRSQL = "SELECT * FROM ALERT  WHERE CODUSU = " & Session("CODUSER") & " AND FLAGREAD = 0 ORDER BY CODALERT DESC"
				'Response.Write(strsql)
				'Response.End()				
				
				Set RSALERTS = Server.CreateObject("Adodb.Recordset")
				Set RSALERTS.ActiveConnection = Conexao
				Set RSALERTS = Conexao.Execute(strSQL)
		    
			elseif  Session("ATEND") = true then
			 
				STRSQL = "SELECT * FROM ALERT  WHERE CODUSU = " & Session("CODUSER") & " AND FLAGREAD = 0 ORDER BY CODALERT DESC"
				'Response.Write(strsql)
				'Response.End()
				
				Set RSALERTS = Server.CreateObject("Adodb.Recordset")
				Set RSALERTS.ActiveConnection = Conexao
				Set RSALERTS = Conexao.Execute(strSQL)	
				
			end if
			 
	 if Not RSALERTS.eof THEN 			 
			 
			 if (Session("ADMS") = true or Session("ADMA") = true) or (Session("ATEND") = false and Session("ADMS") = false and Session("ADMA") = false) or Session("ATEND") = true then
			
			while not RSALERTS.eof
			
			  if sTempOccur <> RSALERTS("codOccur") then
			
				'Seleciono as informações do chamado do alerta
				STRSQL = "SELECT * FROM OCORRENCIA WHERE CODOCCUR = '" & RSALERTS("codOccur") & "'" 
				Set RSCHAMADOSALERTS = Server.CreateObject("Adodb.Recordset")
				Set RSCHAMADOSALERTS.ActiveConnection = Conexao
				Set RSCHAMADOSALERTS = Conexao.Execute(strSQL)
				
				if not RSCHAMADOSALERTS.eof then
				'Seleciono o atendente do alerta
				STRSQL = "SELECT loginUser FROM USUARIOS WHERE CODUSER = " & RSCHAMADOSALERTS("codPro") & "" 
				Set RSATENDENTEALERTS = Server.CreateObject("Adodb.Recordset")
				Set RSATENDENTEALERTS.ActiveConnection = Conexao
				Set RSATENDENTEALERTS = Conexao.Execute(strSQL)
			  
                Response.Write("<tr>") 
				
				if RSALERTS("flagJust") = 0 and RSALERTS("flagAtraso") = 0 then
				   Response.Write("<td class ='grid' width=4% background=images/fundocaixa1.gif><div align=center><img src=images/mensagens.gif width=22 height=22></div></td>")
				   strImg = "mensagens.gif"
				elseif RSALERTS("flagJust") = 1 then
				   Response.Write("<td class ='grid' width=4% background=images/fundocaixa1.gif><div align=center><img src=images/transchamado.gif width=22 height=22></div></td>")				
				   strImg = "transchamado.gif"
				elseif RSALERTS("flagAtraso") = 1 then
				   Response.Write("<td class ='grid' width=4% background=images/fundocaixa1.gif><div align=center><img src=images/chamadoatrasado.gif width=22 height=22></div></td>")				
				   strImg = "chamadoatrasado.gif"
				end if
				
				'Caso o alerta seja de uma justificativa
				if RSALERTS("flagJust") = 0 then 
				   Response.Write("<td class ='grid' width=45% background=images/fundocaixa1.gif><a href=javascript:abre('asp/chamadomensagem.asp?codPro=" & RSCHAMADOSALERTS("codPro") & "&codOccur=" & RSCHAMADOSALERTS("codOccur") & "&img=images/"& strImg & "&fjust=0" & "','Forum','width=720,height=480,%20scrollbars=yes');>" & RSCHAMADOSALERTS("assuntoOccur") & " - " & RSALERTS("mens") & "</a></td>")
				else
                   Response.Write("<td class ='grid' width=45% background=images/fundocaixa1.gif><a href=javascript:abre('asp/chamadomensagem.asp?codPro=" & RSCHAMADOSALERTS("codPro") & "&codOccur=" & RSCHAMADOSALERTS("codOccur") & "&img=images/"& strImg & "&fjust=1" & "','Forum','width=720,height=480,%20scrollbars=yes');>" & RSCHAMADOSALERTS("assuntoOccur") & " - " & RSALERTS("mens") & "</a></td>")				   
                end if 				   
				
				Response.Write("<td class ='grid' background=images/fundocaixa1.gif><div align=center>" & RSATENDENTEALERTS("loginUser") & "</div></td>")

				'SELECIONO O USUÁRIO DONO DO CHAMADO

				STRSQL = "SELECT CODUSER FROM OCORRENCIA WHERE CODOCCUR = '" & RSALERTS("codOccur") & "'" 
				Set RSCODUSER = Server.CreateObject("Adodb.Recordset")
				Set RSCODUSER.ActiveConnection = Conexao
				Set RSCODUSER = Conexao.Execute(strSQL)
								
				STRSQL = "SELECT loginUser FROM USUARIOS WHERE CODUSER = " & RSCODUSER("CODUSER") & "" 	
				Set RSUSEROCCUR = Server.CreateObject("Adodb.Recordset")
				Set RSUSEROCCUR.ActiveConnection = Conexao
				Set RSUSEROCCUR = Conexao.Execute(strSQL)	
					
				Response.Write("<td class ='grid' background=images/fundocaixa1.gif> <div align=center>" & RSUSEROCCUR("loginUser") & "</div></td>")									

				'Seleciono a foto do status
				
				STRSQL = "SELECT pathImg,nameStatus FROM STATUS WHERE CODSTATUS = " & RSCHAMADOSALERTS("codStatus") & "" 	
				Set RSSTATUSALERTS = Server.CreateObject("Adodb.Recordset")
				Set RSSTATUSALERTS.ActiveConnection = Conexao
				Set RSSTATUSALERTS = Conexao.Execute(strSQL)
				
				set RSCHAMADOSALERTS = nothing				
				
				if (Session("ADMS") = true or Session("ADMA") = true) or (Session("ATEND") = false and Session("ADMS") = false and Session("ADMA") = false) or Session("ATEND") = true then
					
					slk = Ucase(RSSTATUSALERTS("pathImg"))
					slk = replace(slk,"../","")
					slk = replace(slk,"/","_")
					slk = Mid(slk,1,Instr(slk,".")-1)
					slk = replace(slk,"_",".")
					ext = slk
					slk = Mid(slk,Instr(slk,".")+1,Len(slk))
					
					slk = " ""javascript:CheckOnlyThis('"& slk &"')"" " 
					
					
				end if 
'** CHECK LIXEIRA				
				'Response.Write("<td class ='grid' background=images/fundocaixa1.gif align=center><a href="& slk &"><img src=images/" & RSSTATUSALERTS("pathImg") & " width=22 height=22 border=0 alt='"& RSSTATUSALERTS("nameStatus") &"' title='"& RSSTATUSALERTS("nameStatus") &"'><br><font size=1>"& RSSTATUSALERTS("nameStatus") &"</font></a></td>")
				
'** CHECK LIXEIRA
				if (Session("ADMS") = true or Session("ADMA") = true) or (Session("ATEND") = false and Session("ADMS") = false and Session("ADMA") = false) or Session("ATEND") = true then
					'Response.Write("<td class ='grid' align=center background=images/fundocaixa1.gif> <input type=checkbox name=apagar value='" & RSALERTS("codAlert") & "|img."& ext &"'></td>")
				end if
					  
			    Response.Write("</tr>")
				
				set RSSTATUSALERTS = nothing
				
				set RSATENDENTEALERTS = nothing							
					  
				end if 
				
	
				intContador = intContador + 1			
				sTempOccur = RSALERTS("codOccur") 
			   end if	
			   
				RSALERTS.movenext

			  wend   
					
			  if not RSALERTS.eof then 
				 Response.Write("<tr>")
				 Response.Write("<td class=azul><div align=center></div></td>")
				 Response.Write("<td class=azul><div align=center></div></td>")
				 Response.Write("<td class=azul><div align=center></div></td>")
				 Response.Write("<td class=azul><div align=center></div></td>")
				 Response.Write("<td class=azul><div align=center><img src=images/nsequencia.gif width=105 height=30 onClick=submit();return true></div></td>")
				 Response.Write("</tr>")
			  end if 
			  
			  set RSALERTS = nothing			  
			  
			  if intContador = 0 then 
                   
				   ZeroChamadosAlerts
				   
              end if
			  
		end if    
		
		 end if    
	
	end function

'###################################################################################################################################
	
	function ZeroChamadosAlerts
             
			 Response.Write("<tr>") 
             'Response.Write("<td background=images/fundocaixa.gif></td>")
             Response.Write("<td colspan=2 background=images/fundocaixa.gif class=preto>Nenhuma mensagem disponível.</td>")
             Response.Write("<td background=images/fundocaixa.gif>&nbsp;</td>")
             Response.Write("<td background=images/fundocaixa.gif>&nbsp;</td>")
             Response.Write("<td background=images/fundocaixa.gif>&nbsp;</td>")
             Response.Write("</tr>")
			 
	end function

'###################################################################################################################################	

    function QtdOcorrencia(codUser)
       
		STRSQL="SELECT COUNT(CODUSER) AS QTD, CODPRO FROM OCORRENCIA WHERE CODPRO = " & codUser & " AND FLAGCLOSE = 0 AND FLAGCANCELA <> 1 GROUP BY CODPRO"
		Set RSQTDOCOR = Server.CreateObject("Adodb.Recordset")
	    Set RSQTDOCOR.ActiveConnection = Conexao
		Set RSQTDOCOR = Conexao.Execute(strSQL)
	    
		if not RSQTDOCOR.eof then QtdOcorrencia = RSQTDOCOR("QTD")

		set RSQTDOCOR = nothing		
		
    end function

'###################################################################################################################################

    function CadastraJustificativa(CODOCCUR,JUSTIFICATIVA,CODPRO,CODNEWUSER)
	    
		'Cadastro a justificativa para trânsferência da ocorrência
	  	STRSQL="INSERT INTO JUSTIFICATIVA (CODOCCUR, JUSTIFICATIVA,CODOLDUSER,CODNEWUSER) VALUES ('" & CODOCCUR & "','" & JUSTIFICATIVA & "',"& CODPRO &","& CODNEWUSER &")"
		Set RSINSJUST = Server.CreateObject("Adodb.Recordset")
	    Set RSINSJUST.ActiveConnection = Conexao
		Set RSINSJUST = Conexao.Execute(strSQL)

		STRSQL = "SELECT U.CODADM FROM USUARIOS U WHERE U.CODUSER = " & CODPRO & ""
		Set RSADM = Server.CreateObject("Adodb.Recordset")
		Set RSADM.ActiveConnection = Conexao
		Set RSADM = Conexao.Execute(strSQL)
	   
		STRSQL = "SELECT CODUSER,NAMEUSER,EMAILUSER FROM USUARIOS WHERE CODADM = " & RSADM("CODADM") & " AND FLAGADMA = 1 OR FLAGADMS = 1"
		Set RSSELADMIN = Server.CreateObject("Adodb.Recordset")
		Set RSSELADMIN.ActiveConnection = Conexao
		Set RSSELADMIN = Conexao.Execute(strSQL)
		
		strMens="Chamado trânsferido - Alteração de atendente"
		
		'Seleciono o usuário da ocorrencia
		STRSQL = "SELECT CODUSER FROM OCORRENCIA WHERE CODOCCUR = '" & CODOCCUR & "'"
		Set RSSELUSER = Server.CreateObject("Adodb.Recordset")
		Set RSSELUSER.ActiveConnection = Conexao
		Set RSSELUSER = Conexao.Execute(strSQL)
		
		'Cadastro um alerta da trânsferencia para o usuário
		STRSQL="INSERT INTO ALERT (CODOCCUR, MENS, CODUSU, FLAGJUST) VALUES ('" & CODOCCUR & "','" & STRMENS & "'," & RSSELUSER("CODUSER") & ",1)"
		Set RSINSALERTUSER = Server.CreateObject("Adodb.Recordset")
		Set RSINSALERTUSER.ActiveConnection = Conexao
		Set RSINSALERTUSER = Conexao.Execute(strSQL)
		
		STRSQL = "SELECT NAMEUSER FROM USUARIOS WHERE CODUSER = " & CODNEWUSER & ""
		Set RSNOMENEWATEND = Server.CreateObject("Adodb.Recordset")
		Set RSNOMENEWATEND.ActiveConnection = Conexao
		Set RSNOMENEWATEND = Conexao.Execute(strSQL)
		
		STRSQL = "SELECT NAMEUSER FROM USUARIOS WHERE CODUSER = " & CODPRO & ""
		Set RSNOMEATEND = Server.CreateObject("Adodb.Recordset")
		Set RSNOMEATEND.ActiveConnection = Conexao
		Set RSNOMEATEND = Conexao.Execute(strSQL)
		
		STRSQL = "SELECT ASSUNTOOCCUR FROM OCORRENCIA WHERE CODOCCUR = '" & CODOCCUR & "'"
		Set RSASSUNTO = Server.CreateObject("Adodb.Recordset")
		Set RSASSUNTO.ActiveConnection = Conexao
		Set RSASSUNTO = Conexao.Execute(strSQL)								
		
		assuntoOccur = RSASSUNTO("ASSUNTOOCCUR")
		
		while not RSSELADMIN.eof 
		
		    'Cadastro um alerta para o(s) administrador(es) do sistema
			STRSQL="INSERT INTO ALERT (CODOCCUR, MENS, CODADM, FLAGJUST) VALUES ('" & CODOCCUR & "','" & STRMENS & "'," & RSSELADMIN("CODUSER") & ",1)"
			Set RSINSALERTADM = Server.CreateObject("Adodb.Recordset")
			Set RSINSALERTADM.ActiveConnection = Conexao
			Set RSINSALERTADM = Conexao.Execute(strSQL)
			
			Host				= "localhost"
			Componente			= "CDONTS"
			Email				= "novochamado.SHD@unimeds.com.br"
			NomeEmail			= "SHD"
			Assunto_Email		= "SHD - Sistema HelpDesk: Chamado Nº "&CODOCCUR&" - "& assuntoOccur & " trânsferido"
		
			Mensagem	=	"<div style='font: 12px; font-family:Verdana, Arial, Helvetica, sans-serif;width:100%;border:1px #FFF solid;'>"& _
							"Prezado (a) "& RSSELADMIN("NAMEUSER") &", <br /><br />"&_
							"O chamado de código <b>"& CODOCCUR &"</b> que possui o "&_
							"assunto <b>"&  assuntoOccur  &"</b> foi transferido. <br /><br />"&_
							"O atendente <b> "&RSNOMEATEND("NAMEUSER")&"</b> "&_
							"transferiu este chamado para o atendente <b>"&RSNOMENEWATEND("NAMEUSER")&"</b> <br /><br />"&_
							"A justificativa para a transferência foi esta: "&_
							"<div style='border:1px solid #000;width:80%;padding: 8px; margin:10px;'>"&_
							Replace(JUSTIFICATIVA,VBCrlf,"<br>")&_
							"</div>"&_ 							
							"<br /><br />Atenciosamente,<br / ><br/ >"&_
							"Sistema de HelpDesk - SHD<br / ><br/ >"&_
							"<a href='http://"& Request.ServerVariables("Server_Name") &"'>Acesse o SHD agora.</a><br / ><br/ >"&_
							"<b>E-mail automático, favor não responder este e-mail.</b>"&_
							"</div>"						 
									
			EnviaEmail Host,Componente,Email,NomeEmail,RSSELADMIN("EMAILUSER"),Assunto_Email,Mensagem
		
			RSSELADMIN.movenext
				
		wend

		'========================================================================================================================
		'SELECIONO O E-MAIL DO USUARIO E ENVIO E-MAIL NOTIFICANDO A TRANSFERÊNCIA		
		'========================================================================================================================
		
		STRSQL = "SELECT U.EMAILUSER,U.NAMEUSER FROM USUARIOS U,OCORRENCIA O WHERE O.CODOCCUR = '" & CODOCCUR & "' AND U.CODUSER = O.CODUSER"
		Set RSSEL_EMAILUSER = Server.CreateObject("Adodb.Recordset")
		Set RSSEL_EMAILUSER.ActiveConnection = Conexao
		Set RSSEL_EMAILUSER = Conexao.Execute(strSQL)
		
		Host				= "localhost"
		Componente			= "CDONTS"
		Email				= "novochamado.SHD@unimeds.com.br"
		NomeEmail			= "SHD"
		Assunto_Email		= "SHD - Sistema HelpDesk: Chamado Nº "&CODOCCUR&" - "& assuntoOccur & " trânsferido"
	

		Mensagem	=	"<div style='font: 12px; font-family:Verdana, Arial, Helvetica, sans-serif;width:100%;border:1px #FFF solid;'>"& _
						"Prezado (a) "& RSSEL_EMAILUSER("NAMEUSER") &", <br /><br />"&_
						"O chamado de código <b>"& CODOCCUR &"</b> que possui o "&_
						"assunto <b>"&  assuntoOccur  &"</b> foi transferido. <br /><br />"&_
						"O atendente <b> "&  RSNOMEATEND("NAMEUSER") &"</b> "&_
						"transferiu este chamado para o atendente <b>"&  RSNOMENEWATEND("NAMEUSER") &"</b> "&_
						"A justificativa para a transferência foi esta: "&_
						"<div style='border:1px solid #000;width:80%;padding: 8px; margin:10px;'>"&_
						Replace(JUSTIFICATIVA,VBCrlf,"<br>")&_
						"</div>"&_ 													
						"<br /><br />Atenciosamente,<br / ><br/ >"&_
						"Sistema de HelpDesk - SHD<br / ><br/ >"&_
						"<a href='http://"& Request.ServerVariables("Server_Name") &"'>Acesse o SHD agora.</a><br / ><br/ >"&_
						"<b>E-mail automático, favor não responder este e-mail.</b>"&_
						"</div>"						 	
								
		 EnviaEmail Host,Componente,Email,NomeEmail,RSSEL_EMAILUSER("EMAILUSER"),Assunto_Email,Mensagem				
		
		'========================================================================================================================
		
		'========================================================================================================================
		'SELECIONO O E-MAIL DO USUARIO E ENVIO E-MAIL NOTIFICANDO A TRANSFERÊNCIA		
		'========================================================================================================================
		
		STRSQL = "SELECT U.EMAILUSER,U.NAMEUSER FROM USUARIOS U,OCORRENCIA O WHERE O.CODOCCUR = '" & CODOCCUR & "' AND U.CODUSER = O.CODPRO"
		Set RSSEL_EMAILATEND = Server.CreateObject("Adodb.Recordset")
		Set RSSEL_EMAILATEND.ActiveConnection = Conexao
		Set RSSEL_EMAILATEND = Conexao.Execute(strSQL)
		
		Host				= "localhost"
		Componente			= "CDONTS"
		Email				= "novochamado.SHD@unimeds.com.br"
		NomeEmail			= "SHD"
		Assunto_Email		= "SHD - Sistema HelpDesk: Chamado Nº "&CODOCCUR&" - "& assuntoOccur & " trânsferido"
	
		Mensagem	=	"<div style='font: 12px; font-family:Verdana, Arial, Helvetica, sans-serif;width:100%;border:1px #FFF solid;'>"& _
						"Prezado (a) "& RSSEL_EMAILUSER("NAMEUSER") &", <br /><br />"&_
						"O chamado de código <b>"& CODOCCUR &"</b> que possui o "&_
						"assunto <b>"&  assuntoOccur  &"</b> foi transferido. <br /><br />"&_
						"O atendente <b> "&  RSNOMEATEND("NAMEUSER") &"</b> "&_
						"transferiu este chamado para o você."&_
						"A justificativa para a transferência foi esta: "&_
						"<div style='border:1px solid #000;width:80%;padding: 8px; margin:10px;'>"&_
						Replace(JUSTIFICATIVA,VBCrlf,"<br>")&_
						"</div>"&_ 							
						"<br /><br />Atenciosamente,<br / ><br/ >"&_
						"Sistema de HelpDesk - SHD<br / ><br/ >"&_
						"<a href='http://"& Request.ServerVariables("Server_Name") &"'>Acesse o SHD agora.</a><br / ><br/ >"&_
						"<b>E-mail automático, favor não responder este e-mail.</b>"&_
						"</div>"						 	
								
		 EnviaEmail Host,Componente,Email,NomeEmail,RSSEL_EMAILATEND("EMAILUSER"),Assunto_Email,Mensagem				
		
		'========================================================================================================================		
		
		'Exibo mensagem de sucesso para o usuário
        strExibeMens = "<script language=JavaScript>"
		strExibeMens = strExibeMens & " alert('Chamado trânsferido com sucesso!');"
	    strExibeMens = strExibeMens & " </script>"
	    CadastraJustificativa = strExibeMens
	    
		set RSINSALERTADM = nothing		
		
		set RSINSALERTUSER = nothing
		
		set RSSELUSER = nothing
		
		set RSSELADMIN = nothing						
		
		set RSINSJUST = nothing
		
		set RSADM = nothing		
		
		
	end function
	
'###################################################################################################################################			

    function CadastraJustificativaModulo(CODOCCUR,JUSTIFICATIVA,CODPRO,CODNEWUSER,CODMODULO)
	    
		'Cadastro a justificativa para trânsferência da ocorrência
	  	STRSQL="INSERT INTO JUSTIFICATIVA (CODOCCUR, JUSTIFICATIVA,CODOLDUSER,CODNEWUSER) VALUES ('" & CODOCCUR & "','" & JUSTIFICATIVA & "',"& CODPRO &","& CODNEWUSER &")"
		Set RSINSJUST = Server.CreateObject("Adodb.Recordset")
	    Set RSINSJUST.ActiveConnection = Conexao
		Set RSINSJUST = Conexao.Execute(strSQL)
		
		STRSQL = "SELECT U.CODADM FROM USUARIOS U WHERE U.CODUSER = " & CODPRO & ""
		Set RSADM = Server.CreateObject("Adodb.Recordset")
		Set RSADM.ActiveConnection = Conexao
		Set RSADM = Conexao.Execute(strSQL)
	   
		STRSQL = "SELECT CODUSER,NAMEUSER,EMAILUSER FROM USUARIOS WHERE CODADM = " & RSADM("CODADM") & " AND FLAGADMA = 1 OR FLAGADMS = 1"
		Set RSSELADMIN = Server.CreateObject("Adodb.Recordset")
		Set RSSELADMIN.ActiveConnection = Conexao
		Set RSSELADMIN = Conexao.Execute(strSQL)
		
		strMens="Chamado trânsferido - Alteração de módulo"
		
		'Seleciono o usuário da ocorrencia
		STRSQL = "SELECT CODUSER FROM OCORRENCIA WHERE CODOCCUR = '" & CODOCCUR & "'"
		Set RSSELUSER = Server.CreateObject("Adodb.Recordset")
		Set RSSELUSER.ActiveConnection = Conexao
		Set RSSELUSER = Conexao.Execute(strSQL)
		
		'Cadastro um alerta da trânsferencia para o usuário
		STRSQL="INSERT INTO ALERT (CODOCCUR, MENS, CODUSU, FLAGJUST) VALUES ('" & CODOCCUR & "','" & STRMENS & "'," & RSSELUSER("CODUSER") & ",1)"
		Set RSINSALERTUSER = Server.CreateObject("Adodb.Recordset")
		Set RSINSALERTUSER.ActiveConnection = Conexao
		Set RSINSALERTUSER = Conexao.Execute(strSQL)
		
		STRSQL = "SELECT NAMEUSER FROM USUARIOS WHERE CODUSER = " & CODNEWUSER & ""
		Set RSNOMENEWATEND = Server.CreateObject("Adodb.Recordset")
		Set RSNOMENEWATEND.ActiveConnection = Conexao
		Set RSNOMENEWATEND = Conexao.Execute(strSQL)
		
		STRSQL = "SELECT NAMEUSER FROM USUARIOS WHERE CODUSER = " & CODPRO & ""
		Set RSNOMEATEND = Server.CreateObject("Adodb.Recordset")
		Set RSNOMEATEND.ActiveConnection = Conexao
		Set RSNOMEATEND = Conexao.Execute(strSQL)
		
		STRSQL = "SELECT ASSUNTOOCCUR FROM OCORRENCIA WHERE CODOCCUR = '" & CODOCCUR & "'"
		Set RSASSUNTO = Server.CreateObject("Adodb.Recordset")
		Set RSASSUNTO.ActiveConnection = Conexao
		Set RSASSUNTO = Conexao.Execute(strSQL)								
		
		assuntoOccur = RSASSUNTO("ASSUNTOOCCUR")
		
		STRSQL = "SELECT NAMEMOD FROM MODULO WHERE CODMOD = " & CODMODULO 
		Set RSSEL_MODULO = Server.CreateObject("Adodb.Recordset")
		Set RSSEL_MODULO.ActiveConnection = Conexao
		Set RSSEL_MODULO = Conexao.Execute(strSQL)		
		
		while not RSSELADMIN.eof 
		
		    'Cadastro um alerta para o(s) administrador(es) do sistema
			STRSQL="INSERT INTO ALERT (CODOCCUR, MENS, CODADM, FLAGJUST) VALUES ('" & CODOCCUR & "','" & STRMENS & "'," & RSSELADMIN("CODUSER") & ",1)"
			Set RSINSALERTADM = Server.CreateObject("Adodb.Recordset")
			Set RSINSALERTADM.ActiveConnection = Conexao
			Set RSINSALERTADM = Conexao.Execute(strSQL)
			
			Host				= "localhost"
			Componente			= "CDONTS"
			Email				= "novochamado.SHD@unimeds.com.br"
			NomeEmail			= "SHD"
			Assunto_Email		= "SHD - Sistema HelpDesk: Chamado Nº "&CODOCCUR&" - "& assuntoOccur & " trânsferido"
		
			Mensagem	=	"<div style='font: 12px; font-family:Verdana, Arial, Helvetica, sans-serif;width:100%;border:1px #FFF solid;'>"& _
							"Prezado (a) "& RSSELADMIN("NAMEUSER") &", <br /><br />"&_
							"O chamado de código <b>"& CODOCCUR &"</b> que possui o "&_
							"assunto <b>"&  assuntoOccur  &"</b> foi transferido. <br /><br />"&_
							"O atendente <b> "&  RSNOMEATEND("NAMEUSER") &"</b> "&_
							"alterou o módulo para <b> "&RSSEL_MODULO("NAMEMOD")&"</b> ."&_
							"A justificativa para a transferência foi esta: "&_
							"<div style='border:1px solid #000;width:80%;padding: 8px; margin:10px;'>"&_
							Replace(JUSTIFICATIVA,VBCrlf,"<br>")&_
							"</div>"&_ 							
							"<br /><br />Atenciosamente,<br / ><br/ >"&_
							"Sistema de HelpDesk - SHD<br / ><br/ >"&_
							"<a href='http://"& Request.ServerVariables("Server_Name") &"'>Acesse o SHD agora.</a><br / ><br/ >"&_
							"<b>E-mail automático, favor não responder este e-mail.</b>"&_
							"</div>"			
									
			EnviaEmail Host,Componente,Email,NomeEmail,RSSELADMIN("EMAILUSER"),Assunto_Email,Mensagem
		
			RSSELADMIN.movenext
				
		wend

		'========================================================================================================================
		'SELECIONO O E-MAIL DO USUARIO E ENVIO E-MAIL NOTIFICANDO A TRANSFERÊNCIA		
		'========================================================================================================================
		
		STRSQL = "SELECT U.EMAILUSER,U.NAMEUSER FROM USUARIOS U,OCORRENCIA O WHERE O.CODOCCUR = '" & CODOCCUR & "' AND U.CODUSER = O.CODUSER"
		Set RSSEL_EMAILUSER = Server.CreateObject("Adodb.Recordset")
		Set RSSEL_EMAILUSER.ActiveConnection = Conexao
		Set RSSEL_EMAILUSER = Conexao.Execute(strSQL)
		
		Host				= "localhost"
		Componente			= "CDONTS"
		Email				= "novochamado.SHD@unimeds.com.br"
		NomeEmail			= "SHD"
		Assunto_Email		= "SHD - Sistema HelpDesk: Chamado Nº "&CODOCCUR&" - "& assuntoOccur & " trânsferido"


		Mensagem	=	"<div style='font: 12px; font-family:Verdana, Arial, Helvetica, sans-serif;width:100%;border:1px #FFF solid;'>"& _
						"Prezado (a) "& RSSEL_EMAILUSER("NAMEUSER") &", <br /><br />"&_
						"O chamado de código <b>"& CODOCCUR &"</b> que possui o "&_
						"assunto <b>"&  assuntoOccur  &"</b> foi transferido. <br /><br />"&_
						"O atendente <b> "&  RSNOMEATEND("NAMEUSER") &"</b> "&_
						"alterou o módulo para <b> "&RSSEL_MODULO("NAMEMOD")&"</b> ."&_
						"A justificativa para a transferência foi esta: "&_
						"<div style='border:1px solid #000;width:80%;padding: 8px; margin:10px;'>"&_
						Replace(JUSTIFICATIVA,VBCrlf,"<br>")&_
						"</div>"&_ 							
						"<br /><br />Atenciosamente,<br / ><br/ >"&_
						"Sistema de HelpDesk - SHD<br / ><br/ >"&_
						"<a href='http://"& Request.ServerVariables("Server_Name") &"'>Acesse o SHD agora.</a><br / ><br/ >"&_
						"<b>E-mail automático, favor não responder este e-mail.</b>"&_
						"</div>"	
								
		 EnviaEmail Host,Componente,Email,NomeEmail,RSSEL_EMAILUSER("EMAILUSER"),Assunto_Email,Mensagem				
		
		'========================================================================================================================
		
		'========================================================================================================================
		'SELECIONO O E-MAIL DO NOVO ATENDENTE E ENVIO E-MAIL NOTIFICANDO A TRANSFERÊNCIA		
		'========================================================================================================================
		
		STRSQL = "SELECT U.EMAILUSER,U.NAMEUSER FROM USUARIOS U,OCORRENCIA O WHERE O.CODOCCUR = '" & CODOCCUR & "' AND U.CODUSER = O.CODPRO"
		Set RSSEL_EMAILATEND = Server.CreateObject("Adodb.Recordset")
		Set RSSEL_EMAILATEND.ActiveConnection = Conexao
		Set RSSEL_EMAILATEND = Conexao.Execute(strSQL)
		
		Host				= "localhost"
		Componente			= "CDONTS"
		Email				= "novochamado.SHD@unimeds.com.br"
		NomeEmail			= "SHD"
		Assunto_Email		= "SHD - Sistema HelpDesk: Chamado Nº "&CODOCCUR&" - "& assuntoOccur & " foi trânsferido para você"
	
		Mensagem	=	"<div style='font: 12px; font-family:Verdana, Arial, Helvetica, sans-serif;width:100%;border:1px #FFF solid;'>"& _
						"Prezado (a) "& RSSEL_EMAILATEND("NAMEUSER") &", <br /><br />"&_
						"O chamado de código <b>"& CODOCCUR &"</b> que possui o "&_
						"assunto <b>"&  assuntoOccur  &"</b> foi transferido. <br /><br />"&_
						"O atendente <b> "&  RSNOMEATEND("NAMEUSER") &"</b> "&_
						"alterou o módulo para <b> "&RSSEL_MODULO("NAMEMOD")&"</b> ."&_
						"A justificativa para a transferência foi esta: "&_
						"<div style='border:1px solid #000;width:80%;padding: 8px; margin:10px;'>"&_
						Replace(JUSTIFICATIVA,VBCrlf,"<br>")&_
						"</div>"&_ 							
						"<br /><br />Atenciosamente,<br / ><br/ >"&_
						"Sistema de HelpDesk - SHD<br / ><br/ >"&_
						"<a href='http://"& Request.ServerVariables("Server_Name") &"'>Acesse o SHD agora.</a><br / ><br/ >"&_
						"<b>E-mail automático, favor não responder este e-mail.</b>"&_
						"</div>"		
								
		 EnviaEmail Host,Componente,Email,NomeEmail,RSSEL_EMAILATEND("EMAILUSER"),Assunto_Email,Mensagem				
		
		'========================================================================================================================		
		
		'Exibo mensagem de sucesso para o usuário
        strExibeMens = "<script language=JavaScript>"
		strExibeMens = strExibeMens & " alert('Chamado trânsferido com sucesso!');"
	    strExibeMens = strExibeMens & " </script>"
	    CadastraJustificativaModulo = strExibeMens
	    
		set RSINSALERTADM = nothing		
		
		set RSINSALERTUSER = nothing
		
		set RSSELUSER = nothing
		
		set RSSELADMIN = nothing						
		
		set RSINSJUST = nothing
		
		set RSADM = nothing		
		
	end function
	
'###################################################################################################################################			


	function AlteraOccurAtendente(CODPRO,CODOCCUR)
	
			STRSQL="UPDATE OCORRENCIA SET CODPRO = " & CODPRO & " WHERE CODOCCUR = '" & CODOCCUR & "'"
			Set RSALTERAATEND = Server.CreateObject("Adodb.Recordset")
			Set RSALTERAATEND.ActiveConnection = Conexao
			Set RSALTERAATEND = Conexao.Execute(strSQL)
	
		set RSALTERAATEND = nothing		
	
	end function
	
'###################################################################################################################################

	function AlteraOccurModulo(CODPRO,CODOCCUR,CODMODULO)
	
			STRSQL="UPDATE OCORRENCIA SET CODPRO = " & CODPRO & ", CODMOD = "&CODMODULO&" WHERE CODOCCUR = '" & CODOCCUR & "'"
			Set RSALTERAMOD = Server.CreateObject("Adodb.Recordset")
			Set RSALTERAMOD.ActiveConnection = Conexao
			Set RSALTERAMOD = Conexao.Execute(strSQL)
	
		set RSALTERAMOD = nothing		
	
	end function
	
'###################################################################################################################################

	function SendMail(EMAIL_REM, NOME_REM, EMAIL_DEST, ASSUNTO, CORPO)	'-------------------------------------------------
%><!-- 
    METADATA 
    TYPE="typelib" 
    UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
    NAME="CDO for Windows 2000 Library" 
--> <%
	Set cdoConfig = CreateObject("CDO.Configuration")  
	With cdoConfig.Fields  
        .Item(cdoSendUsingMethod) = cdoSendUsingPort
        .Item(cdoSMTPServer) = "192.168.10.38"  
        .Item(cdoSMTPAuthenticate) = 0  
        .Update  
    End With 
 
    Set cdoMessage = CreateObject("CDO.Message")  
 
    With cdoMessage 
        Set .Configuration = cdoConfig 
        .From = Email
        .To = ParaEmail
        .Subject = ASSUNTO 
        .HTMLBody = Mensagem
        .Send 
    End With 
 
    Set cdoMessage = Nothing  
    Set cdoConfig = Nothing  	
	'-------------------------------------------------

	end function	
	
'###################################################################################################################################

	function FlagConclusaoChamado(FLAGCONCLUSAO,CODOCCUR)	
		
		data = day(date)&"/"& month(date) &"/"& year(date) &" "& hour(time) &":"& minute(time)
		
		STRSQL="UPDATE OCORRENCIA SET FLAGCLOSE = " & FLAGCONCLUSAO & ", DATAFECHADO = '"& data &"' WHERE CODOCCUR = '" & CODOCCUR & "'"
		Set RSCONCHAMADO = Server.CreateObject("Adodb.Recordset")
		Set RSCONCHAMADO.ActiveConnection = Conexao
		Set RSCONCHAMADO = Conexao.Execute(strSQL)
		
		set RSCONCHAMADO = nothing		
		
	end function

'###################################################################################################################################

	function FlagReadAlert(CODALERT,CODUSER)
		'-----------------------------------------------------------------------'
		if Instr(CODALERT,",") > 0 THEN
		'-----------------------------------------------------------------------'
			Linhas = split(CODALERT,",")
		'-----------------------------------------------------------------------'
			CODALERT = ""
		'-----------------------------------------------------------------------'
			for i=Lbound(linhas) to Ubound(linhas)
		'-----------------------------------------------------------------------'
				dados		= split(Linhas(i),"|")
				CODALERT	= CODALERT & Trim(dados(0)) &","
		'-----------------------------------------------------------------------'	
			next
		'-----------------------------------------------------------------------'		
		else
		'-----------------------------------------------------------------------'
			dados		= split(CODALERT,"|")
			CODALERT = ""
		'-----------------------------------------------------------------------'
			CODALERT	= Trim(dados(0)) &","
		'-----------------------------------------------------------------------'
		end if
		'-----------------------------------------------------------------------'
		CODALERT = Left(CODALERT,Len(CODALERT)-1)
		'-----------------------------------------------------------------------'
		if (Session("USU") = true) OR (Session("ATEND") = true) then
			STRSQL="DELETE FROM ALERT WHERE CODALERT in ("& CODALERT & ") AND CODUSU = " & CODUSER & ""
			Conexao.Execute(strSQL)	
		end if
		'-----------------------------------------------------------------------'
		if (Session("ADMA") = true) or (Session("ADMS") = true) then
			STRSQL="DELETE FROM ALERT WHERE CODALERT in (" & CODALERT & ") AND CODADM = " & CODUSER & ""
			Conexao.Execute(strSQL)
		end if
		'-----------------------------------------------------------------------'
		' response.Write "<!--[strSQL=["& strSQL &"]-->"
		' response.Write strSQL
		' response.End	
		
	end function

'###################################################################################################################################	
	
	function SelLixeira
'----------------------------------------------------------------------------------------------------'		
			if Session("ADMS") = true or Session("ADMA") = true then
'----------------------------------------------------------------------------------------------------'
				STRSQL = "SELECT * FROM ALERT WHERE CODADM = " & Session("CODUSER") & " AND FLAGREAD = 1 ORDER BY CODALERT DESC" 	
				Set RSALERTS = Server.CreateObject("Adodb.Recordset")
				Set RSALERTS.ActiveConnection = Conexao
				Set RSALERTS = Conexao.Execute(strSQL)
'----------------------------------------------------------------------------------------------------'
			elseif (Session("ADMS") = true or Session("ADMA") = true) or (Session("ATEND") = false and Session("ADMS") = false and Session("ADMA") = false) then
'----------------------------------------------------------------------------------------------------'
				STRSQL = "SELECT * FROM ALERT WHERE CODUSU = " & Session("CODUSER") & " AND FLAGREAD = 1 ORDER BY CODALERT DESC"
				Set RSALERTS = Server.CreateObject("Adodb.Recordset")
				Set RSALERTS.ActiveConnection = Conexao
				Set RSALERTS = Conexao.Execute(strSQL)
'----------------------------------------------------------------------------------------------------'
			elseif  Session("ATEND") = true then
'----------------------------------------------------------------------------------------------------'
				STRSQL = "SELECT * FROM ALERT WHERE CODUSU = " & Session("CODUSER") & " AND FLAGREAD = 1 ORDER BY CODALERT DESC"
				Set RSALERTS = Server.CreateObject("Adodb.Recordset")
				Set RSALERTS.ActiveConnection = Conexao
				Set RSALERTS = Conexao.Execute(strSQL)
'----------------------------------------------------------------------------------------------------'
			end if
'----------------------------------------------------------------------------------------------------'
			 if (Session("ADMS") = true or Session("ADMA") = true) or (Session("ATEND") = false and Session("ADMS") = false and Session("ADMA") = false) or Session("ATEND") = true then
'----------------------------------------------------------------------------------------------------'
			while not RSALERTS.eof
				'Seleciono as informações do chamado do alerta
				STRSQL = "SELECT * FROM OCORRENCIA WHERE CODOCCUR = '" & RSALERTS("codOccur") & "'" 
				Set RSCHAMADOSALERTS = Server.CreateObject("Adodb.Recordset")
				Set RSCHAMADOSALERTS.ActiveConnection = Conexao
				Set RSCHAMADOSALERTS = Conexao.Execute(strSQL)
'----------------------------------------------------------------------------------------------------'
				if not RSCHAMADOSALERTS.eof then
					'Seleciono o atendente do alerta
					STRSQL = "SELECT loginUser FROM USUARIOS WHERE CODUSER = " & RSCHAMADOSALERTS("codPro") & "" 
					Set RSATENDENTEALERTS = Server.CreateObject("Adodb.Recordset")
					Set RSATENDENTEALERTS.ActiveConnection = Conexao
					Set RSATENDENTEALERTS = Conexao.Execute(strSQL)
'----------------------------------------------------------------------------------------------------'
					Response.Write("<tr>") 
'----------------------------------------------------------------------------------------------------'
					if RSALERTS("flagJust") = 0 and RSALERTS("flagAtraso") = 0 then
					   Response.Write("<td class ='grid' width=4% background=images/fundocaixa1.gif><div align=center><img src=images/mensagens.gif width=22 height=22></div></td>")
					   strImg = "mensagens.gif"
					elseif RSALERTS("flagJust") = 1 then
					   Response.Write("<td class ='grid' width=4% background=images/fundocaixa1.gif><div align=center><img src=images/transchamado.gif width=22 height=22></div></td>")				
					   strImg = "transchamado.gif"
					elseif RSALERTS("flagAtraso") = 1 then
					   Response.Write("<td class ='grid' width=4% background=images/fundocaixa1.gif><div align=center><img src=images/chamadoatrasado.gif width=22 height=22></div></td>")				
					   strImg = "chamadoatrasado.gif"
					end if
'----------------------------------------------------------------------------------------------------'
					'Caso o alerta seja de uma justificativa
					if RSALERTS("flagJust") = 0 then 
					   Response.Write("<td class ='grid' width=45% background=images/fundocaixa1.gif><a href=javascript:abre('asp/chamadomensagem.asp?codPro=" & RSCHAMADOSALERTS("codPro") & "&codOccur=" & RSCHAMADOSALERTS("codOccur") & "&img=images/"& strImg & "','Forum','width=720,height=480,%20scrollbars=yes');>" & RSCHAMADOSALERTS("assuntoOccur") & " - " & RSALERTS("mens") & "</a></td>")
					else
					   Response.Write("<td class ='grid' width=45% background=images/fundocaixa1.gif><a href=javascript:abre('asp/chamadomensagem.asp?codPro=" & RSCHAMADOSALERTS("codPro") & "&codOccur=" & RSCHAMADOSALERTS("codOccur") & "&img=images/"& strImg & "','Forum','width=720,height=480,%20scrollbars=yes');>" & RSCHAMADOSALERTS("assuntoOccur") & " - " & RSALERTS("mens") & "</a></td>")				   
					end if 				   
'----------------------------------------------------------------------------------------------------'
					Response.Write("<td class ='grid' background=images/fundocaixa1.gif> <div align=center>" & RSATENDENTEALERTS("loginUser") & "</div></td>")
'----------------------------------------------------------------------------------------------------'
				'SELECIONO O USUÁRIO DONO DO CHAMADO
'----------------------------------------------------------------------------------------------------'
				STRSQL = "SELECT CODUSER FROM OCORRENCIA WHERE CODOCCUR = '" & RSALERTS("codOccur") & "'" 
				Set RSCODUSER = Server.CreateObject("Adodb.Recordset")
				Set RSCODUSER.ActiveConnection = Conexao
				Set RSCODUSER = Conexao.Execute(strSQL)
'----------------------------------------------------------------------------------------------------'
				STRSQL = "SELECT loginUser FROM USUARIOS WHERE CODUSER = " & RSCODUSER("CODUSER") & "" 	
				Set RSUSEROCCUR = Server.CreateObject("Adodb.Recordset")
				Set RSUSEROCCUR.ActiveConnection = Conexao
				Set RSUSEROCCUR = Conexao.Execute(strSQL)	
'----------------------------------------------------------------------------------------------------'
				Response.Write("<td class ='grid' background=images/fundocaixa1.gif> <div align=center>" & RSUSEROCCUR("loginUser") & "</div></td>")									
'----------------------------------------------------------------------------------------------------'
					'Seleciono a foto do status
'----------------------------------------------------------------------------------------------------'
					STRSQL = "SELECT pathImg FROM STATUS WHERE CODSTATUS = " & RSCHAMADOSALERTS("codStatus") & "" 	
					Set RSSTATUSALERTS = Server.CreateObject("Adodb.Recordset")
					Set RSSTATUSALERTS.ActiveConnection = Conexao
					Set RSSTATUSALERTS = Conexao.Execute(strSQL)
'----------------------------------------------------------------------------------------------------'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
					if (Session("ADMS") = true or Session("ADMA") = true) or (Session("ATEND") = false and Session("ADMS") = false and Session("ADMA") = false) or Session("ATEND") = true then
'----------------------------------------------------------------------------------------------------'				
						slk = Ucase(RSSTATUSALERTS("pathImg"))
						slk = replace(slk,"../","")
						slk = replace(slk,"/","_")
						slk = Mid(slk,1,Instr(slk,".")-1)
						slk = replace(slk,"_",".")
						ext = slk
						slk = Mid(slk,Instr(slk,".")+1,Len(slk))						
						slk = " ""javascript:CheckOnlyThis('"& slk &"')"" " 
'----------------------------------------------------------------------------------------------------'
					end if 
'----------------------------------------------------------------------------------------------------'			
						Response.Write("<td class ='grid' background=images/fundocaixa1.gif align=center><a href="& slk &"><img src=images/" & RSSTATUSALERTS("pathImg") & " width=22 height=22 border=0></a></td>")
'----------------------------------------------------------------------------------------------------'
					if (Session("ADMS") = true or Session("ADMA") = true) or (Session("ATEND") = false and Session("ADMS") = false and Session("ADMA") = false) or Session("ATEND") = true then
'----------------------------------------------------------------------------------------------------'
						Response.Write("<td class ='grid' align=center background=images/fundocaixa1.gif> <input type=checkbox name=apagar value='" & RSALERTS("codAlert") & "|"& ext &"'>") 
						Response.Write("</td>")
'----------------------------------------------------------------------------------------------------'
					end if
'----------------------------------------------------------------------------------------------------'
					Response.Write("</tr>")
'----------------------------------------------------------------------------------------------------'
				  set RSSTATUSALERTS = nothing
'----------------------------------------------------------------------------------------------------'
				  set RSATENDENTEALERTS = nothing				  					
'----------------------------------------------------------------------------------------------------'
				  set RSCHAMADOSALERTS = nothing
'----------------------------------------------------------------------------------------------------'
				end if 
'----------------------------------------------------------------------------------------------------'
				'Seleciono os IDS dos Alerts para apaga-los de uma única vez na lixeira
				IDS_ALERTS = IDS_ALERTS & "," & RSALERTS("CODALERT")				
				RSALERTS.movenext
				intContador = intContador + 1
'----------------------------------------------------------------------------------------------------'
			  wend 
'----------------------------------------------------------------------------------------------------'
			  if not RSALERTS.eof then 
				 Response.Write("<tr>")
				 Response.Write("<td class=azul><div align=center></div></td>")
				 Response.Write("<td class=azul><div align=center></div></td>")
				 Response.Write("<td class=azul><div align=center></div></td>")
				 Response.Write("<td class=azul><div align=center></div></td>")
				 Response.Write("<td class=azul><div align=center><img src=images/images/nsequencia.gif width=105 height=30 onClick=submit();return true></div></td>")
				 Response.Write("</tr>")
			  end if 
'----------------------------------------------------------------------------------------------------'			  
			  if intContador = 0 then 
'----------------------------------------------------------------------------------------------------'
				   ZeroChamadosAlerts
'----------------------------------------------------------------------------------------------------'
              end if
'----------------------------------------------------------------------------------------------------'
			  set RSALERTS = nothing
'----------------------------------------------------------------------------------------------------'
		 end if
'----------------------------------------------------------------------------------------------------'
		SESSION("IDS_ALERTS") = MID(IDS_ALERTS,2,LEN(IDS_ALERTS))
'----------------------------------------------------------------------------------------------------'
	end function	
'###################################################################################################################################

   function FlagApagaAlert(CODALERT)
		'-----------------------------------------------------------------------'
		if Instr(CODALERT,",") > 0 THEN
		'-----------------------------------------------------------------------'
			Linhas = split(CODALERT,",")
		'-----------------------------------------------------------------------'
			CODALERT = ""
			for i=Lbound(linhas) to Ubound(linhas)
		'-----------------------------------------------------------------------'
				dados		= split(Linhas(i),"|")
				CODALERT	= CODALERT & Trim(dados(0)) &","
		'-----------------------------------------------------------------------'	
			next
		'-----------------------------------------------------------------------'		
		else
		'-----------------------------------------------------------------------'
			dados		= split(CODALERT,"|")
			CODALERT = ""
		'-----------------------------------------------------------------------'
			CODALERT	= Trim(dados(0)) &","
		'-----------------------------------------------------------------------'
		end if
		'-----------------------------------------------------------------------'
		CODALERT = Left(CODALERT,Len(CODALERT)-1)
'-----------------------------------------------------------------------'
	 STRSQL="DELETE ALERT WHERE CODALERT In(" & CODALERT & ")"
	 Set RSAPAGACHAMADO = Server.CreateObject("Adodb.Recordset")
	 Set RSAPAGACHAMADO.ActiveConnection = Conexao
	 Set RSAPAGACHAMADO = Conexao.Execute(strSQL)	
'----------------------------------------------------------------------------------------------------'
	'Exibo mensagem de sucesso para o usuário
     strExibeMens = "<script language=JavaScript>"
     strExibeMens = strExibeMens & " alert('Mensagem apagada!');"
     strExibeMens = strExibeMens & " </script>"
     FlagApagaAlert = strExibeMens
'----------------------------------------------------------------------------------------------------'
	set RSAPAGACHAMADO = nothing
'----------------------------------------------------------------------------------------------------'
  end function
		
'###################################################################################################################################

	function CadastraAdministradora(CODADM,NOMEADM,DESCADM,FONEADM,CADUSER,LOGINUSER)
		
		strAdmin= VerificaAdministradora(CODADM,NOMEADM)
		
		if CADUSER = true then strLogin=VerificaLogin(LOGINUSER,CODADM)
		
		if (CADUSER = true) and (strAdmin = false) then
		
			if strLogin = false then
				'Cadastro administradora
				idPai = Session("CODPAI") : if idPai = 0 then idPai = Session("CODADM")
				
				STRSQL="INSERT INTO ADMINISTRADORA (CODADM, NAMEADM, DESCADM, FONEADM, IDPAI) VALUES (" & CODADM & ",'" & NOMEADM & "','" & DESCADM & "','" & FONEADM & "',"& idPai &")"
				Set RSCADADM = Server.CreateObject("Adodb.Recordset")
				Set RSCADADM.ActiveConnection = Conexao
				Set RSCADADM = Conexao.Execute(strSQL)
				
				'Exibo mensagem de sucesso para o usuário
				strExibeMens = "<script language=JavaScript>"
				strExibeMens = strExibeMens & " alert('Administradora cadastrada!');"
				strExibeMens = strExibeMens & " </script>"
				CadastraAdministradora = strExibeMens
				Session("strEXIBE3") = true
    		        else
				if (strLogin = false) and (strAdmin = true) then
				   Session("strEXIBE4") = true
				elseif (strLogin = true) and (strAdmin = true) then
                   Session("strEXIBE4") = true
                   Session("strEXIBE2") = true
				end if	
			end if
			
		elseif (CADUSER = "") and (strAdmin = false) then
						
				'Cadastro administradora
				STRSQL="INSERT INTO ADMINISTRADORA (CODADM, NAMEADM, DESCADM, FONEADM) VALUES (" & CODADM & ",'" & NOMEADM & "','" & DESCADM & "'," & FONEADM & ")"
				Set RSCADADM = Server.CreateObject("Adodb.Recordset")
				Set RSCADADM.ActiveConnection = Conexao
				Set RSCADADM = Conexao.Execute(strSQL)
				Session("strEXIBE3") = true
	   else
				if (strLogin = false) and (strAdmin = true) then
				   Session("strEXIBE4") = true
				elseif (strLogin = true) and (strAdmin = true) then
                   Session("strEXIBE4") = true
                   Session("strEXIBE2") = true
				elseif (strLogin = true) and (strAdmin = false) then
                   Session("strEXIBE4") = true
               	end if					
       end if
	   
	set RSCADADM = nothing
				
	end function		
	
'###################################################################################################################################

	function CadastraUsuario(CODADM, NAMEUSER, EMAILUSER, FLAGADMS, FLAGADMA, FLAGATEND, FONEUSER, LOGINUSER, PASSUSER,CADADMIN)
		
		strLogin = VerificaLogin(LOGINUSER,CODADM)
		
		if CADADMIN = "" then
		
			if strLogin = false then
				'Cadastro usuário
				STRSQL="INSERT INTO USUARIOS (CODADM, NAMEUSER, EMAILUSER," & _
					   "FLAGADMS, FLAGADMA, FLAGATEND, FONEUSER, LOGINUSER," & _
					   "PASSUSER) VALUES " & _
					   "(" & CODADM & ",'" & NAMEUSER & "','" & EMAILUSER & "'," & FLAGADMS & "," & FLAGADMA & "," & FLAGATEND & ",'" & FONEUSER & "','" & LOGINUSER & "','" & PASSUSER & "'); SELECT @@IDENTITY AS COD;"
		  			   Set RSCADUSU = Server.CreateObject("Adodb.Recordset")
					   Set RSCADUSU =  Conexao.Execute(strSQL)
					   Set RSCADUSU = RSCADUSU.NextRecordSet()	
							Conexao.Execute("INSERT INTO GRUPO_USUARIO (IDGRUPO,IDATENDENTE,STATUS) VALUES (3,"&  RSCADUSU.Fields("COD").value	&",1)")
				       Session("strEXIBE") = true
			else
					   Session("strEXIBE2")=true
			end if	
		elseif CADADMIN = true then		    
			if (strLogin = false) and (Session("strEXIBE3") = true) then
				'Cadastro usuário
				STRSQL="INSERT INTO USUARIOS (CODADM, NAMEUSER, EMAILUSER, FLAGADMS, FLAGADMA, FLAGATEND, FONEUSER, LOGINUSER, PASSUSER) VALUES (" & CODADM & ",'" & NAMEUSER & "','" & EMAILUSER & "'," & FLAGADMS & "," & FLAGADMA & "," & FLAGATEND & ",'" & FONEUSER & "','" & LOGINUSER & "','" & PASSUSER & "'); SELECT @@IDENTITY AS COD;"
				Set RSCADUSU = Server.CreateObject("Adodb.Recordset")
				Set RSCADUSU = Conexao.Execute(strSQL)				
			   Set RSCADUSU = RSCADUSU.NextRecordSet()	
					Conexao.Execute("INSERT INTO GRUPO_USUARIO (IDGRUPO,IDATENDENTE,STATUS) VALUES (4,"&  RSCADUSU.Fields("COD").value	&",1)")
response.end					
				set RSCADUSU = nothing
				'Exibo mensagem de sucesso para o usuário
				strExibeMens = "<script language=JavaScript>"
				strExibeMens = strExibeMens & " alert('Usuário cadastrado!');"
				strExibeMens = strExibeMens & " </script>"
				CadastraUsuario = strExibeMens
				Session("strEXIBE") = true
			else
			   Session("strEXIBE2")=true
			end if
		end if 
	end function
	
'###################################################################################################################################

	Const EncC1 = 109
	Const EncC2 = 191
	Const EncKey = 161
	
	Public Function EncriptaStr(Texto)

		Dim TempStr, TempResult, TempNum, TempChar
		Dim TempKey
		Dim i
		
		TempStr = Texto
		TempResult = ""
		TempKey = ((EncKey * EncC1) + EncC2) Mod 65536
		
		For i = 1 To Len(TempStr)
		TempNum = (Asc(Mid(TempStr, i, 1)) Xor (AuxShr(TempKey, 8))) Mod 256
		TempChar = Chr(TempNum)
		TempKey = (((Asc(TempChar) + TempKey) * EncC1) + EncC2) Mod 65536
		TempResult = TempResult & TempChar
		Next
		
		EncriptaStr = TempResult

	End Function

'###################################################################################################################################	
	
	'descriptografando o texto
	Public Function DecriptaStr(Texto)
		
		Dim TempStr, TempResult, TempNum, TempChar
		Dim TempKey
		Dim i
		
		TempStr = Texto
		TempResult = ""
		TempKey = ((EncKey * EncC1) + EncC2) Mod 65536
		
		For i = 1 To Len(TempStr)
		TempNum = (Asc(Mid(TempStr, i, 1)) Xor (AuxShr(TempKey, 8))) Mod 256
		TempChar = Chr(TempNum)
		TempKey = (((Asc(Mid(TempStr, i, 1)) + TempKey) * EncC1) + EncC2) Mod 65536
		TempResult = TempResult & TempChar
		Next
		
		DecriptaStr = TempResult
		
	End Function

'###################################################################################################################################	
		
	Private Function AuxShr(Numero, BShr)
	
		AuxShr = Int(Numero / (2 ^ BShr))
		
	End Function

'###################################################################################################################################

	function VerificaLogin(STRLOGIN,CODADM)
	
		STRSQL = "SELECT * FROM USUARIOS WHERE LOGINUSER = '" & STRLOGIN & "' AND CODADM = " & CODADM
		Set RSVLOGIN = Server.CreateObject("Adodb.Recordset")
		Set RSVLOGIN.ActiveConnection = Conexao
		Set RSVLOGIN = Conexao.Execute(strSQL)
		
		if RSVLOGIN.eof then
		   VerificaLogin=false
		else
		   VerificaLogin=true
		end if
		
	set RSVLOGIN = nothing
		
	end function

'###################################################################################################################################

	function VerificaAdministradora(STRCOD,STRNAME)
	
		if STRCOD <> "" and STRNAME <> "" then
		   STRSQL = "SELECT * FROM ADMINISTRADORA WHERE CODADM = " & STRCOD & " AND NAMEADM = '" & STRNAME & "'"
		elseif STRCOD = "" and STRNAME <> "" then   
		   STRSQL = "SELECT * FROM ADMINISTRADORA WHERE NAMEADM = '" & STRNAME & "'"
		elseif STRCOD <> "" and STRNAME = "" then      
		   STRSQL = "SELECT * FROM ADMINISTRADORA WHERE CODADM = " & STRCOD & ""
		end if
		
		Set RSVADM = Server.CreateObject("Adodb.Recordset")
		Set RSVADM.ActiveConnection = Conexao
		Set RSVADM = Conexao.Execute(strSQL)
		
		if RSVADM.eof then
		   VerificaAdministradora=false
		else
		   VerificaAdministradora=true
		end if
		
	set RSVADM = nothing
		
	end function
	
'###################################################################################################################################

	function AtualizaArquivo(PATH,CODOCCUR)
		
		STRSQL="UPDATE OCORRENCIA SET PATHFILE = '" & PATH & "' WHERE CODOCCUR = '" & CODOCCUR & "'"
		Set RSATUCHAMADO = Server.CreateObject("Adodb.Recordset")
		Set RSATUCHAMADO.ActiveConnection = Conexao
		Set RSATUCHAMADO = Conexao.Execute(strSQL)

		set RSATUCHAMADO = nothing
	
	end function
	
'###################################################################################################################################

	function CadastraAtendente(NAMEUSER,EMAILUSER,FONEUSER,LOGINUSER,PASSUSER,ATEND,CODADM)
				
		STRSQL="INSERT INTO USUARIOS (NAMEUSER,EMAILUSER,FONEUSER,LOGINUSER,PASSUSER,FLAGATEND,CODADM) VALUES ('" & NAMEUSER & "','" & EMAILUSER & "','" & FONEUSER & "','" & LOGINUSER & "','" & PASSUSER & "',1,"& CODADM &"); SELECT @@IDENTITY AS COD;"
		Set RSINSATENDE = Server.CreateObject("Adodb.Recordset")
'		Set RSINSATENDE.ActiveConnection = Conexao
		Set RSINSATENDE = Conexao.Execute(strSQL)
		Set RSINDATENDE = RSINSATENDE.NextRecordSet()		
			DivModAtend ATEND, RSINDATENDE.Fields("COD").value		
			CadastraAtendente = true		
			Conexao.Execute("INSERT INTO GRUPO_USUARIO (IDGRUPO,IDATENDENTE,STATUS) VALUES (5,"&  RSINDATENDE.Fields("COD").value  &",1)")
	set RSINSATENDE = nothing
	
	end function
	
'###################################################################################################################################

	function CadastraModulo(NAMEMOD,DESCRMOD,CODATEND,ESPECIFICA)
				
		STRSQL="INSERT INTO MODULO (NAMEMOD,DESCRMOD,IDMODULOESPEC,CODADM) VALUES ('" & NAMEMOD & "','" & DESCRMOD & "',"& ESPECIFICA &","& Session("CODADM") &"); SELECT @@IDENTITY AS COD;"
	
		Set RSINSMOD = Server.CreateObject("Adodb.Recordset")
		Set RSINSMOD.ActiveConnection = Conexao
		Set RSINSMOD = Conexao.Execute(strSQL)
		
		Set RSNCODMOD = RSINSMOD.NextRecordSet()
		CODMOD = RSNCODMOD.Fields("COD").value

		AtualizaAtendente CODATEND,CODMOD
		
		CadastraModulo = true
		
	set RSINSMOD = nothing
	

	end function
	
'###################################################################################################################################

	function ApagaModulo(CODMOD)
				
		STRSQL="SELECT * FROM MODATEND WHERE CODMOD = " & CODMOD
		Set RSSELMODATEND = Server.CreateObject("Adodb.Recordset")
		Set RSSELMODATEND.ActiveConnection = Conexao
		Set RSSELMODATEND = Conexao.Execute(strSQL)
		
		while not RSSELMODATEND.eof
			ApagaModAtend RSSELMODATEND("CODUSER"),CODMOD
			RSSELMODATEND.movenext()
		wend

		STRSQL="DELETE FROM MODULO WHERE CODMOD = " & CODMOD
		Set RSSELMODATEND = Server.CreateObject("Adodb.Recordset")
		Set RSSELMODATEND.ActiveConnection = Conexao
		Set RSSELMODATEND = Conexao.Execute(strSQL)		
		
		ApagaModulo = true
		
	set RSSELMODATEND = nothing

	end function
	
'###################################################################################################################################

	function AtualizaModuloSistema(CODMOD,APAGA)
				
		if APAGA <> "1" THEN
			STRSQL="UPDATE MODULO SET FLAG_SISTEMA = 1 WHERE CODMOD = " & CODMOD &" AND CODADM = "& Session("CODADM")
		else
			STRSQL="UPDATE MODULO SET FLAG_SISTEMA = 0 WHERE CODMOD = " & CODMOD &" AND CODADM = "& Session("CODADM")	
		end if	
		
		Set RSUPDMOD = Server.CreateObject("Adodb.Recordset")
		Set RSUPDMOD.ActiveConnection = Conexao
		Set RSUPDMOD = Conexao.Execute(strSQL)
		
		set RSUPDMOD = nothing

	end function		
	
'###################################################################################################################################

	function CadastraModAtend(CODUSER,CODMOD)
		
		   STRSQL="INSERT INTO MODATEND (CODUSER,CODMOD) VALUES (" & CODUSER & "," & CODMOD & ")"
		   Set RSINSMOD = Server.CreateObject("Adodb.Recordset")
		   Set RSINSMOD.ActiveConnection = Conexao
		   Set RSINSMOD = Conexao.Execute(strSQL)
	 	   CadastraModAtend = true
		   
	set RSINSMOD = nothing
	    
	end function
	
'###################################################################################################################################

	function DivModAtend(ATENDE,CODUSER)
         
		strTamanho = len(ATENDE)
 				
		for i = 1 to strTamanho
		
		  if mid(ATENDE,i,1) <> "," and mid(ATENDE,i,1) <> " " then strCarc = strCarc & mid(ATENDE,i,1)
		  if mid(ATENDE,i,1) = "," or i = strTamanho then
		     CadastraModAtend CODUSER,	trim(strCarc)
			 strCarc = ""
		  end if
		  
		next
	
	end function
	
'###################################################################################################################################
	
	function AtualizaAtendente(CODATEND,CODMOD)

		strTamanho = len(CODATEND)
	
		for i = 1 to strTamanho
				  
		  if mid(CODATEND,i,1) <> "," and mid(CODATEND,i,1) <> " " then strCarc = strCarc & mid(CODATEND,i,1)
		  if mid(CODATEND,i,1) = "," or i = strTamanho then
		     CadastraModAtend trim(strCarc),CODMOD
			 strCarc = ""
		  end if	 	
		  
		next
		
	end function
	
'###################################################################################################################################

	function MensagemCancelada(codOcorr)
		
		strMens="Chamado cancelado"
   	    STRSQL = "SELECT U.CODADM FROM USUARIOS U, ADMINISTRADORA A, OCORRENCIA O WHERE U.CODUSER = " & Session("CODUSER") & " AND O.CODOCCUR = '" & codOcorr & "' GROUP BY U.CODADM"
		Set RSADMCHAMADO_CANCEL = Server.CreateObject("Adodb.Recordset")
		Set RSADMCHAMADO_CANCEL.ActiveConnection = Conexao
		Set RSADMCHAMADO_CANCEL = Conexao.Execute(strSQL)
		
	    STRSQL = "SELECT CODUSER,ASSUNTOOCCUR FROM OCORRENCIA WHERE CODOCCUR = '" & codOcorr & "'"
		Set RSMENSUSER = Server.CreateObject("Adodb.Recordset")
		Set RSMENSUSER.ActiveConnection = Conexao
		Set RSMENSUSER = Conexao.Execute(strSQL)		
		
		'=========================================================================================================================
		'MONTA SEQUÊNCIA
		'=========================================================================================================================
		
		STRSQL = "SELECT * FROM SEQUENCIA WHERE CODOCCUR = '" & codOcorr & "' ORDER BY CODSEQ DESC" 
		Set RSSEQUENCIA = Server.CreateObject("Adodb.Recordset")
		Set RSSEQUENCIA.ActiveConnection = Conexao
		Set RSSEQUENCIA = Conexao.Execute(strSQL)	
		
		while not RSSEQUENCIA.eof 
	
			 strMDescri = RSSEQUENCIA("descrSeq")
			 strDataSeq = RSSEQUENCIA("dataSeq")
			 strHoraSeq = RSSEQUENCIA("horaSeq")
			 strcodUsuSeq = RSSEQUENCIA("CODUSER")
			 
			 STRSQL = "SELECT nameUser FROM USUARIOS WHERE CODUSER = '" & strcodUsuSeq & "'" 	
			 Set RSUSUSEQ = Server.CreateObject("Adodb.Recordset")
			 Set RSUSUSEQ.ActiveConnection = Conexao
			 Set RSUSUSEQ = Conexao.Execute(strSQL)
					
			strMontaSequencia_cancelado = strMontaSequencia_cancelado & "=====================================================================================" & vbcrlf 
			strMontaSequencia_cancelado = strMontaSequencia_cancelado & RSSEQUENCIA("dataSeq") & " - " & Replace(RSSEQUENCIA("horaSeq"),"1/1/1900","") & " - " & RSUSUSEQ("nameUser") & vbcrlf
			strMontaSequencia_cancelado = strMontaSequencia_cancelado & strMDescri & vbcrlf
			RSSEQUENCIA.movenext
			
		wend							

		'=========================================================================================================================		
		
   	    STRSQL = "SELECT codUser,NAMEUSER,EMAILUSER FROM USUARIOS WHERE CODADM = " & RSADMCHAMADO_CANCEL("CODADM") & " AND FLAGADMA = 1 OR FLAGADMS = 1"
		Set RSADMIN_CANCEL_CHAMADO = Server.CreateObject("Adodb.Recordset")
		Set RSADMIN_CANCEL_CHAMADO.ActiveConnection = Conexao
		Set RSADMIN_CANCEL_CHAMADO = Conexao.Execute(strSQL)
		
   	    STRSQL = "SELECT CODPRO FROM OCORRENCIA WHERE CODOCCUR = '" & codOcorr & "'"
		Set RSATEND_CANCEL_CHAMADO = Server.CreateObject("Adodb.Recordset")
		Set RSATEND_CANCEL_CHAMADO.ActiveConnection = Conexao
		Set RSATEND_CANCEL_CHAMADO = Conexao.Execute(strSQL)
		
		 STRSQL = "INSERT INTO ALERT (codOccur,codUsu,mens) VALUES ('" & codOcorr & "'," & RSATEND_CANCEL_CHAMADO("CODPRO") & ",'" & strMens & "')"
		 Set RSINSMENSATEND = Server.CreateObject("Adodb.Recordset")
		 Set RSINSMENSATEND.ActiveConnection = Conexao
		 Set RSINSMENSATEND = Conexao.Execute(strSQL)
		
		while not RSADMIN_CANCEL_CHAMADO.eof
		
		 STRSQL = "INSERT INTO ALERT (codOccur,codAdm,mens) VALUES ('" & codOcorr & "'," & RSADMIN_CANCEL_CHAMADO("codUser") & ",'" & strMens & "')"
		 Set RSINSMENSADMIN = Server.CreateObject("Adodb.Recordset")
		 Set RSINSMENSADMIN.ActiveConnection = Conexao
		 Set RSINSMENSADMIN = Conexao.Execute(strSQL)
		 
		'======================================================================================================================
			
		Host				= "localhost"
		Componente			= "CDONTS"
		Email				= "novochamado.SHD@unimeds.com.br"
		NomeEmail			= "SHD"
		Assunto_Email		= "SHD - Sistema HelpDesk: Chamado Nº "&codOcorr&" - "& RSMENSUSER("ASSUNTOOCCUR") & " chamado cancelado."
		
		Mensagem	=	"<div style='font: 12px; font-family:Verdana, Arial, Helvetica, sans-serif;width:100%;border:1px #FFF solid;'>"& _
						"Prezado (a) "& RSADMIN_CANCEL_CHAMADO("NAMEUSER") &", <br /><br />"&_
						"O chamado de código <b>"& CODOCCUR &"</b> que possui o "&_
						"assunto <b>"&  RSMENSUSER("ASSUNTOOCCUR")  &"</b> foi cancelado por <b>"&Session("NAME")&"</b>. <br /><br />"&_
						"A justificativa para o cancelamento foi esta: "&_
						"<div style='border:1px solid #000;width:80%;padding: 8px; margin:10px;'>"&_
						Replace(strMontaSequencia_cancelado,VBCrlf,"<br>")&_
						"</div>"&_ 													
						"<br /><br />Atenciosamente,<br / ><br/ >"&_
						"Sistema de HelpDesk - SHD<br / ><br/ >"&_
						"<a href='http://"& Request.ServerVariables("Server_Name") &"'>Acesse o SHD agora.</a><br / ><br/ >"&_
						"<b>E-mail automático, favor não responder este e-mail.</b>"&_
						"</div>"			
									
		 EnviaEmail Host,Componente,Email,NomeEmail,RSADMIN_CANCEL_CHAMADO("EMAILUSER"),Assunto_Email,Mensagem
		 
		'======================================================================================================================					 
		 
		 RSADMIN_CANCEL_CHAMADO.movenext
		 
		wend

	STRSQL = "INSERT INTO ALERT (codOccur,codUsu,mens) VALUES ('" & codOcorr & "'," & RSMENSUSER("codUser") & ",'" & strMens & "')"
	Set RSINSMENSUSER = Server.CreateObject("Adodb.Recordset")
	Set RSINSMENSUSER.ActiveConnection = Conexao
	Set RSINSMENSUSER = Conexao.Execute(strSQL)		
		
	set RSINSMENSADMIN = nothing
	
	set RSADMIN_CANCEL_CHAMADO = nothing	
	
	set RSINSMENSATEND = nothing	
	
	set RSATEND_CANCEL_CHAMADO = nothing	
	
	set RSADMIN_CANCEL_CHAMADO = nothing
	
	set RSADMCHAMADO_CANCEL = nothing		
	
		
	end function

'###################################################################################################################################
	
	function FlagCancelChamado(FLAGCANCELA,CODOCCUR)	
		data = year(date)&"-"& month(date) &"-"& day(date) &" "& hour(time) &":"& minute(time)
		STRSQL="UPDATE OCORRENCIA SET FLAGCANCELA = " & FLAGCANCELA & ", DATAFECHADO = '"& data &"' WHERE CODOCCUR = '" & CODOCCUR & "'"
		Set RSCANCELCHAMADO = Server.CreateObject("Adodb.Recordset")
		Set RSCANCELCHAMADO.ActiveConnection = Conexao
		Set RSCANCELCHAMADO = Conexao.Execute(strSQL)
		
	set RSCANCELCHAMADO = nothing		
		
	end function
	
'###################################################################################################################################

	function AlteraDados(STRNOME,STREMAIL,STRFONE,STRSENHA,STRLOGIN,STRCODUSER,DEPTO)
	

		if DEPTO <> "" then
			strSQL =	" UPDATE USUARIOS SET "&_
						" CODADM = "& DEPTO &_
						" WHERE CODUSER = "& STRCODUSER
			Conexao.Execute(strSQL)
		end if				
	
	    if STRNOME <> "" then
			STRSQL = "UPDATE USUARIOS SET "
			STRSQL=STRSQL & " NAMEUSER = '" & STRNOME & "'"
			STRSQL=STRSQL & " WHERE CODUSER = " & STRCODUSER & ";"
			Conexao.Execute(strSQL)
		end if
		
		if STREMAIL <> "" then		
			STRSQL=			"UPDATE USUARIOS SET "
			STRSQL=STRSQL & " EMAILUSER = '" & STREMAIL & "'"
			STRSQL=STRSQL & " WHERE CODUSER = " & STRCODUSER & ";"
			Conexao.Execute(strSQL)
		end if	
		
		if STRFONE <> "" then
			STRSQL=			"UPDATE USUARIOS SET "
			STRSQL=STRSQL & " FONEUSER = " & STRFONE & ""
			STRSQL=STRSQL & " WHERE CODUSER = " & STRCODUSER & ";"
			Conexao.Execute(strSQL)
		end if		
		
		if STRSENHA <> "" then
		    'ENCRIPTA_SENHA = EncriptaStr(STRSENHA) 
			ENCRIPTA_SENHA = Cripto(STRSENHA,true)
			'ENCRIPTA_SENHA = Replace(ENCRIPTA_SENHA,"'","")
			STRSQL=			"UPDATE USUARIOS SET "
			STRSQL=STRSQL & " PASSUSER = '" & ENCRIPTA_SENHA & "'"
			STRSQL=STRSQL & " WHERE CODUSER = " & STRCODUSER & ";"
			Conexao.Execute(strSQL)
			'Response.Write(STRSQL)
			'Response.End()					
		end if
		
		if STRLOGIN <> "" then
		    STRSQL=			"UPDATE USUARIOS SET "
			STRSQL=STRSQL & " LOGINUSER = '" & STRLOGIN & "'"
			STRSQL=STRSQL & " WHERE CODUSER = " & STRCODUSER & ";"		
			Conexao.Execute(strSQL)
		end if		



		AlteraDados = true
	
	end function
	
'###################################################################################################################################

	function AlteraAdministradora(STRCOD,STRNOME,STRDESC,STRFONE,STRCODADM)
	
	
		   STRSQL=  " UPDATE ADMINISTRADORA SET "&_
		   	    " DESCADM = [DESCADM] "
	
	    if STRCOD <> "" then  STRSQL=STRSQL & ", CODADM = " & STRCOD 		
	    if STRNOME <> "" then STRSQL=STRSQL & ", NAMEADM = '" & STRNOME & "'"
  	    if STRDESC <> "" then STRSQL=STRSQL & ", DESCADM = '" & STRDESC & "'"
	    if STRFONE <> "" then STRSQL=STRSQL & ", FONEADM = '" & STRFONE & "'"
		STRSQL=STRSQL & " WHERE CODADM = " & STRCODADM & ";"		
		
		Conexao.Execute(strSQL)
	
	end function
	
'###################################################################################################################################

	function AlteraAtendente(STRNOME,STRFONE,STREMAIL,STRLOGIN,STRSENHA,STRCODUSER,ATENDE)
	
		STRSQL = ""	
	
	    if STRNOME <> "" then
			STRSQL="UPDATE USUARIOS SET "
			STRSQL=STRSQL & " NAMEUSER = '" & STRNOME & "'"
			STRSQL=STRSQL & " WHERE CODUSER = " & STRCODUSER & ";"
			Conexao.Execute(strSQL)
		'	response.Write "NOME"
		end if
		
		if STRFONE <> "" then		
			STRSQL = ""
			STRSQL=STRSQL & "UPDATE USUARIOS SET "
			STRSQL=STRSQL & " FONEUSER = '" & STRFONE & "'"
			STRSQL=STRSQL & " WHERE CODUSER = " & STRCODUSER 
			Conexao.Execute(strSQL)
		'	response.Write "FONE"
		end if	
		
		if STREMAIL <> "" then
			STRSQL = ""
			STRSQL=STRSQL & "UPDATE USUARIOS SET "
			STRSQL=STRSQL & " EMAILUSER = '" & STREMAIL & "'"
			STRSQL=STRSQL & " WHERE CODUSER = " & STRCODUSER & ";"
			Conexao.Execute(strSQL)
		'	response.Write "EMAIL"
		end if		
		
		if STRLOGIN <> "" then
			STRSQL = ""
		    STRSQL=STRSQL & "UPDATE USUARIOS SET "
			STRSQL=STRSQL & " LOGINUSER = '" & STRLOGIN & "'"
			STRSQL=STRSQL & " WHERE CODUSER = " & STRCODUSER & ";"		
			Conexao.Execute(strSQL)
		'	response.Write "LOGIN"
		end if
		
		if STRSENHA <> "" then
			STRSQL = ""
		    ENCRIPTA_SENHA = Cripto(STRSENHA,true) 
		    STRSQL=STRSQL & "UPDATE USUARIOS SET "
			STRSQL=STRSQL & " PASSUSER = '" & ENCRIPTA_SENHA & "'"
			STRSQL=STRSQL & " WHERE CODUSER = " & STRCODUSER & ";"		
			Conexao.Execute(strSQL)
		'	response.Write "SENHA"
		end if				
		AlteraAtendente = true
		
	end function
	
'###################################################################################################################################

	function ApagaModAtend(CODATEND,ATEND)
			 
			 STRSQL="DELETE FROM MODATEND WHERE CODUSER = " & CODATEND & " AND CODMOD = (" & ATEND & ")" 
			 Set RSDELMOD = Server.CreateObject("Adodb.Recordset")
			 Set RSDELMOD.ActiveConnection = Conexao
			 Set RSDELMOD = Conexao.Execute(strSQL)
			 
			set RSDELMOD = nothing		
			 
	end function
	
'###################################################################################################################################

	function InsereModAtend(CODATEND,ATEND)
			 
			 STRSQL="INSERT MODATEND VALUES (" & CODATEND & "," & ATEND & ")" 
			 Set RSINSMOD = Server.CreateObject("Adodb.Recordset")
			 Set RSINSMOD.ActiveConnection = Conexao
			 Set RSINSMOD = Conexao.Execute(strSQL)
			 
			set RSINSMOD = nothing		
			 
	end function
	
'###################################################################################################################################

	function AlteraModulo(STRNOME,STRDESCRI,STRCODMOD,ESPECIFICACAO)
	
	    if STRNOME <> "" then
			STRSQL="UPDATE MODULO SET "
			STRSQL=STRSQL & " NAMEMOD = '" & STRNOME & "' "
			STRSQL=STRSQL & " WHERE CODMOD = " & STRCODMOD & ";"
		end if
		
		if STRDESCRI <> "" then		
			STRSQL=STRSQL & "UPDATE MODULO SET "
			STRSQL=STRSQL & " DESCRMOD = '" & STRDESCRI & "'"
			STRSQL=STRSQL & " WHERE CODMOD = " & STRCODMOD & ";"
		end if	

		Set RSATUAMOD = Server.CreateObject("Adodb.Recordset")
		Set RSATUAMOD.ActiveConnection = Conexao
		Set RSATUAMOD = Conexao.Execute(strSQL)
		
		
	    if ESPECIFICACAO <> "" then
			STRSQL="UPDATE MODULO SET "
			STRSQL=STRSQL & " IDMODULOESPEC = " & ESPECIFICACAO  
			STRSQL=STRSQL & " WHERE CODMOD = " & STRCODMOD & ";"
			Conexao.Execute(strSQL)
		end if
				
		AlteraModulo = true
		
		set RSATUAMOD = nothing		

	
	end function
	
'###################################################################################################################################	

	function Atualiza_CodAdministradora_Usuario(NEW_COD,OLD_COD)
		STRSQL="UPDATE USUARIOS SET CODADM = " & NEW_COD & " WHERE CODADM = " & OLD_COD & ""
		Set RSATUCODADM_USER = Server.CreateObject("Adodb.Recordset")
		Set RSATUCODADM_USER.ActiveConnection = Conexao
		Set RSATUCODADM_USER = Conexao.Execute(strSQL) 
		set RSATUCODADM_USER = nothing	
	end function

'###################################################################################################################################

	function SelModulo(nome)
		
		STRSQL="SELECT * FROM MODULO WHERE NAMEMOD = '" & TRIM(nome) & "'"
		Set RSMOD = Server.CreateObject("Adodb.Recordset")
		Set RSMOD.ActiveConnection = Conexao
		Set RSMOD = Conexao.Execute(strSQL)
		
		if RSMOD.EOF then
			SelModulo = false
		else
			SelModulo = true		
		end if

		set RSMOD = nothing		
		
	end function

'###################################################################################################################################

	function GeraCodOcorrencia(CODADM,CODUSER,ANO,MES,DIA,HORA,MINUTO,SEGUNDOS)
		
		GeraCodOcorrencia = CODADM & CODUSER & ANO & MES & DIA & HORA & MINUTO & SEGUNDOS				
		
	end function

'###################################################################################################################################

	'-------------------------------------------------
	'## ENVIA EMAIL
	'-------------------------------------------------
	Function EnviaEmail(Host,Componente,Email,NomeEmail,ParaEmail,Assunto,Mensagem)	'-------------------------------------------------
%><!-- 
    METADATA 
    TYPE="typelib" 
    UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
    NAME="CDO for Windows 2000 Library" 
--> <%
	Set cdoConfig = CreateObject("CDO.Configuration")  
	With cdoConfig.Fields  
        .Item(cdoSendUsingMethod) = cdoSendUsingPort
        .Item(cdoSMTPServer) = "192.168.10.38"  
        .Item(cdoSMTPAuthenticate) = 0  
        .Update  
    End With 
 
    Set cdoMessage = CreateObject("CDO.Message")  
 
    With cdoMessage 
        Set .Configuration = cdoConfig 
        .From = Email
        .To = ParaEmail
        .Subject = ASSUNTO 
        .HTMLBody = Mensagem
        .Send 
    End With 
 
    Set cdoMessage = Nothing  
    Set cdoConfig = Nothing  	
	'-------------------------------------------------
	End Function

'###################################################################################################################################
	
	Function Cripto(StrCripto, BolAcao) 'Início de da função de criptografia. Aonde o parâmetro String é o valor que será criptografado ou descriptografado. E o parâmetro BolAcao é um valor booleano (True ou False) para indicar se deve ser criptografado (True) ou descriptografado (False).
		If BolAcao Then
		Cripto = EncodeBase64(StrCripto)
			   Else
		Cripto = DecodeBase64(StrCripto)
		End If
	End Function

'###################################################################################################################################

	Function EncodeBase64(inData)
		On Error Resume Next
		  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
		  Dim cOut
		  Dim sOut
		  Dim I
		  For I = 1 To Len(inData) Step 3
			Dim nGroup, pOut, sGroup
			nGroup = &H10000 * Asc(Mid(inData, I, 1)) + &H100 * MyASC(Mid(inData, I + 1, 1)) + MyASC(Mid(inData, I + 2, 1))
			nGroup = Oct(nGroup)
			nGroup = String(8 - Len(nGroup), "0") & nGroup
			pOut = Mid(Base64, CLng("&o" & Mid(nGroup, 1, 2)) + 1, 1) + _
			  Mid(Base64, CLng("&o" & Mid(nGroup, 3, 2)) + 1, 1) + _
			  Mid(Base64, CLng("&o" & Mid(nGroup, 5, 2)) + 1, 1) + _
			  Mid(Base64, CLng("&o" & Mid(nGroup, 7, 2)) + 1, 1)
			sOut = sOut + pOut
			If (I + 2) Mod 57 = 0 Then sOut = sOut + vbCrLf
		  Next
		  Select Case Len(inData) Mod 3
			Case 1:
			  sOut = Left(sOut, Len(sOut) - 2) + "=="
			Case 2:
			  sOut = Left(sOut, Len(sOut) - 1) + "="
		  End Select
		  EncodeBase64 = sOut
	End Function

'###################################################################################################################################	
	
	Function MyASC(OneChar)
	  If OneChar = "" Then MyASC = 0 Else MyASC = Asc(OneChar)
	End Function
	
'###################################################################################################################################

	Function DecodeBase64(ByVal base64String)
		On Error Resume Next
		  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
		  Dim dataLength
		  Dim sOut
		  Dim groupBegin
		  base64String = Replace(base64String, vbCrLf, "")
		  base64String = Replace(base64String, vbTab, "")
		  base64String = Replace(base64String, " ", "")
		  dataLength = Len(base64String)
		  If dataLength Mod 4 <> 0 Then
			Err.Raise 1, "WbSolutions", "String de criptografia com problemas. " & VBNewline & "Contate nosso suporte técnico pelo telefone (0xx51) 475-7545."
			Exit Function
		  End If
		
		  For groupBegin = 1 To dataLength Step 4
			Dim numDataBytes
			Dim CharCounter
			Dim thisChar
			Dim thisData
			Dim nGroup
			Dim pOut
			numDataBytes = 3
			nGroup = 0
		
			For CharCounter = 0 To 3
			  thisChar = Mid(base64String, groupBegin + CharCounter, 1)
			  If thisChar = "=" Then
				numDataBytes = numDataBytes - 1
				thisData = 0
			  Else
				thisData = InStr(Base64, thisChar) - 1
			  End If
			  If thisData = -1 Then
				Err.Raise 2, "WbSolutions", "String de criptografia com problemas. " & VBNewline & "Contate nosso suporte técnico pelo telefone (0xx51) 475-7545."
				Exit Function
			  End If
			  nGroup = 64 * nGroup + thisData
			Next
			nGroup = Hex(nGroup)
			nGroup = String(6 - Len(nGroup), "0") & nGroup
			pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
			  Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
			  Chr(CByte("&H" & Mid(nGroup, 5, 2)))
			sOut = sOut & Left(pOut, numDataBytes)
		  Next
		  DecodeBase64 = sOut
	End Function	

'###################################################################################################################################
	Function InsereArquivo(path,CODOCCUR)
	
		STRSQL = "SELECT * FROM FILES WHERE CODOCCUR = '"&CODOCCUR&"' AND PATH_FILE = '"&path&"'"
		Set RSSEL_ARQUIVO = Server.CreateObject("Adodb.Recordset")
		Set RSSEL_ARQUIVO.ActiveConnection = Conexao
		Set RSSEL_ARQUIVO = Conexao.Execute(strSQL)	
	
		if RSSEL_ARQUIVO.eof then
	
			STRSQL = "INSERT INTO FILES (CODOCCUR,path_file) VALUES ('" & CODOCCUR & "','" & path & "')"
			Set RSINSERE_ARQUIVO = Server.CreateObject("Adodb.Recordset")
			Set RSINSERE_ARQUIVO.ActiveConnection = Conexao
			Set RSINSERE_ARQUIVO = Conexao.Execute(strSQL)		
		
		end if
				
	End Function
'###################################################################################################################################

	Function ApagaLixeiraFull(IDS)
	
		STRSQL = "DELETE FROM ALERT WHERE CODALERT IN ("&IDS&")"
		Set RSSEL_ARQUIVO = Server.CreateObject("Adodb.Recordset")
		Set RSSEL_ARQUIVO.ActiveConnection = Conexao
		Set RSSEL_ARQUIVO = Conexao.Execute(strSQL)

		'Exibo mensagem de sucesso para o usuário
		 strExibeMens = "<script language=JavaScript>"
		 strExibeMens = strExibeMens & " alert('Mensagens apagadas!');"
		 strExibeMens = strExibeMens & " </script>"
		 ApagaLixeiraFull = strExibeMens
		
	End Function
'###################################################################################################################################

'-------------------------------------------------
'//## FUNÇÃO DE SELECT
'-------------------------------------------------
Function MontaSelect(SQL,Nome,Valor,Titulo,Multiplo,Selecionado,AtivaJs)
		IF Instr(Titulo,",") > 0 THEN	MeuArray = Split(Titulo,",")
		IF AtivaJs = "" THEN			AtivaJs = ""	
		IF Multiplo = "1" THEN			Multiplo = " multiple size='5' "	ELSE	Multiplo = " size=1 "
		Set Lnk = Conexao.Execute(SQL)
		IF Not Lnk.Eof Then
			lnk1 = "<select name='"&Nome&"' class=input "& multiplo & AtivaJs &">"&VBcrlf 
			IF Multiplo = " size=1 " THEN	lnk1 = lnk1 & "<option value=''> ---------------------</option>"&VBcrlf
				Do While Not Lnk.eof
						IF Instr(Titulo,",") > 0 THEN
							lnk1 = lnk1 & "<option value='"& Lnk(Valor)&"'"
							IF Selecionado <> "0" THEN	IF cStr(Lnk(Valor)) = cStr(Selecionado) THEN	lnk1 = lnk1 & " SELECTED "
							lnk1 = lnk1 & ">"
								for i=Lbound(MeuArray) to Ubound(MeuArray)
									lnk1 = lnk1 & Lnk(MeuArray(i)) 
									IF i <> Ubound(MeuArray) Then	lnk1 = lnk1 & " - "
								next
							lnk1 = lnk1 & "</option>"& VBCrlf
						ELSE
							lnk1 = lnk1 & "<option value='"& Lnk(Valor)&"'"
								IF Selecionado <> "0" THEN	IF cStr(Lnk(Valor)) = cStr(Selecionado) THEN	lnk1 = lnk1 & " SELECTED "
							lnk1 = lnk1 & ">"& Lnk(TITULO) &"</option>"& VBCrlf	
						END IF
				Lnk.MoveNext
				Loop
			lnk1 = lnk1 & "</select>"
		ELSE
			lnk1 = "Não foi possível localizar informações."
		END IF
	MontaSelect	= Lnk1
	Lnk.close
	Set Lnk = Nothing
End Function 	
'-------------------------------------------------
'//## FUNÇÃO DO MENU DINAMICO
'-------------------------------------------------
Function montaMenu(strMsg)
	strMsg = Replace(strMsg," ","_")
	strMsg = Replace(strMsg,"/","//")
	strMsg = Replace(strMsg,"(","_")
	strMsg = Replace(strMsg,")","_")
	strMsg = Replace(strMsg,"'","_")
	strMsg = Replace(strMsg,chr(34),"_")
	montaMenu = strMsg
End Function
'-------------------------------------------------
'//## FUNÇÃO DE CAMPOS NO HTML
'-------------------------------------------------
Function Campos(tipo,tamanho,coluna,linhas,maximo,nome,valor,js)
'-------------------------------------------------
On Error resume next

		Select Case Tipo
'-------------------------------------------------		
				Case 1 '### TEXT
'-------------------------------------------------
					IF maximo = "" OR maximo = 0 THEN
						maximo = tamanho
					END IF
'-------------------------------------------------
					saida = "<input type='text' size='"&tamanho&"' maxlength='"& maximo &"' name='"& nome &"' value='"& valor &"' id='"& nome &"' "& js &">" & VBCRlf 
'-------------------------------------------------
				Case 2 '### TEXTAREA
'-------------------------------------------------
					saida = "<textarea name='"& nome &"' id='"& nome &"' cols='"& coluna &"' rows='"& linhas &"' "& js &">"& valor &"</textarea>" & VBCRlf 				
'-------------------------------------------------		
				Case 3 '### HIDDEN
'-------------------------------------------------
					IF maximo = "" OR maximo = 0 THEN
						maximo = tamanho
					END IF
'-------------------------------------------------
					saida = "<input type='hidden' size='"&tamanho&"' maxlength='"& maximo &"' name='"& nome &"' value='"& valor &"' "& js &" id='"& nome &"'>" & VBCRlf
'-------------------------------------------------		
				Case 4 '### RESET
'-------------------------------------------------
					saida = "<input type='reset' style='width:"&tamanho&" px' name='"& nome &"' value='"& valor &"' id='"& nome &"' "& js &">" & VBCRlf
'-------------------------------------------------		
				Case 5 '### SUBMIT
'-------------------------------------------------
					saida = "<input type='submit' style='width:"&tamanho&" px' name='"& nome &"' value='"& valor &"' id='"& nome &"' "& js &">" & VBCRlf    
'-------------------------------------------------		
				Case 6 '### CHECKBOX
'-------------------------------------------------
					strT = valor
					IF cstr(Flag(strT)) = cstr(tamanho) OR tamanho = "checked"  THEN						
						tamanho = " CHECKED "
					ELSE
						tamanho = " "
					END IF
					saida = "<input type='checkbox' name='"& nome &"' value='"& valor &"' "&tamanho&"  id='"& nome &"' "& js &">" & VBCRlf
'-------------------------------------------------		
				Case 7 '### RADIO
'-------------------------------------------------
					IF cstr(Flag(valor)) = cstr(tamanho) THEN
						tamanho = " CHECKED "
					ELSE
						tamanho = " "
					END IF										
					saida = "<input type='radio' name='"& nome &"' value='"& valor &"' "&tamanho&" id='"& nome &"' "& js &">" & VBCRlf 
'-------------------------------------------------		
				Case 8 '### PASSWORD
'-------------------------------------------------
					IF maximo = "" OR maximo = 0 THEN
						maximo = tamanho
					END IF
'-------------------------------------------------
					saida = "<input type='password' size='"&tamanho&"' maxlength='"& maximo &"' name='"& nome &"' value='"& valor &"' id='"& nome &"' "& js &">" & VBCRlf 
'-------------------------------------------------		
				Case 9 '### READONLY
'-------------------------------------------------
					saida = "<input type='text' size='"&tamanho&"' maxlength='"& maximo &"' name='"& nome &"' value='"& valor &"' id='"& nome &"' readonly "& js &">" & VBCRlf
'-------------------------------------------------		
				Case 10 '### DISABLED
'-------------------------------------------------
					saida = "<input type='text' size='"&tamanho&"' maxlength='"& maximo &"' name='"& nome &"' value='"& valor &"' id='"& nome &"' disabled "& js &">" & VBCRlf 					 
'-------------------------------------------------		
				Case 11 '### FILE
'-------------------------------------------------
					saida = "<input type='file' size='"&tamanho&"' name='"& nome &"' id='"& nome &"'>" & VBCRlf 					 
'-------------------------------------------------		
				Case 12 '### BUTTON
'-------------------------------------------------
					saida = "<input type='button' style='width:"&tamanho&" px' name='"& nome &"' value='"& valor &"' id='"& nome &"' "& js &">" & VBCRlf
'-------------------------------------------------
		End select
'-------------------------------------------------
		Campos = saida		
'-------------------------------------------------
End Function

'-------------------------------------------------
'//## FUNÇÃO PARA ALTERAR VALOR DE UM RADIO BUTTON OU CHECKBOX
'-------------------------------------------------
Function FLAG(valor)
	if valor = "s" OR valor = true OR valor = "1" then
		valor = 1
	else
		valor = 0
	end if
	FLAG = valor
End Function

'-------------------------------------------------
'//## FUNÇÃO DE PAGINAÇÃO
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

'-------------------------------------------------------'
'## Função responsável por formatar o texto
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

'## Função de Ordenação da matriz
'-------------------------------------------------
Function ordena(sequencia,pOrder)
	numero=sequencia 
	if pOrder = 0 then iniCont = UBound(numero) : fimCount = LBound(numero) : pTipo = - 1
	if pOrder = 1 then iniCont = LBound(numero) : fimCount = UBound(numero) : pTipo = 1

	'Atribuo o qtde de 1 para cada item do array 
	Dim total(100) 
	For cont = iniCont To fimCount step pTipo
		total(cont)=1 
	Next 

	'Aqui eu agrupo os numeros iguais e somo as quantidades 
	For contador=0 To 100 
		For count=contador+1 To UBound(numero)
			IF numero(contador)<>"anulado" And numero(contador)=numero(count) Then 
				Total(contador)=Total(contador)+1 
				numero(count)="anulado" 
				total(count)=0 
			End IF 
		Next 
	Next 

	Dim maior(100), qtde(100) 

	For cont=0 To UBound(numero) 
		maior(cont)=numero(0) 
		For count=1 To UBound(numero)
			IF maior(cont)< numero(count) Then 
				maior(cont)=numero(count) 
			End IF 
		Next 
	
		For contador=0 To UBound(numero)
			IF maior(cont)=numero(contador) Then 
				numero(contador)=0 
				qtde(cont)=total(contador) 
			End IF
		Next 
	Next 

	'Gerando o resultado que vai ser impresso
	nConta = "1"
	For cont = iniCont To fimCount step pTipo
		IF qtde(cont)<>0 And maior(cont)<>"0" Then 
			If nConta <> 1 Then sVirgula = ", "
			varResult = varResult & sVirgula & maior(cont)
			End IF
		nConta = nConta + 1
	Next 
	ordena=varResult 
End Function 

'-------------------------------------------------
'## FUNCAO DE REQUEST 0 = GET, 1 = POST
'-------------------------------------------------
Function Pegar(Campo,Tipo)
'-------------------------------------------------
		IF cInt(Tipo) = 0 THEN
	'-------------------------------------------------
			xis = Request.QueryString(Campo)
	'-------------------------------------------------
		ELSE
	'-------------------------------------------------
			xis = Request.Form(Campo)
	'-------------------------------------------------
		END IF
'-------------------------------------------------
		IF xis = "" Then
'-------------------------------------------------
			Select Case Tipo
	'-------------------------------------------------
					Case 0
	'-------------------------------------------------
						xis = Request.Form(Campo)
	'-------------------------------------------------
					Case Else
	'-------------------------------------------------
						xis = Request.QueryString(Campo)
	'-------------------------------------------------
			End Select
'-------------------------------------------------
		END IF
'-------------------------------------------------
		Pegar = xis
'-------------------------------------------------
End Function 

'-------------------------------------------------------'
'## Função responsável pela gravação do texto no banco
'-------------------------------------------------------'
Function gravar(MsgIn)
	'------------------------------------------------------------------------------
	Dim Codificado
	Codificado= MsgIn
	'------------------------------------------------------------------------------
	Codificado = server.HTMLEncode(Codificado)
	'------------------------------------------------------------------------------
	Codificado = replace(Codificado,CHR(34),"")
	Codificado = replace(Codificado,CHR(39),"")
	Codificado = replace(Codificado,VBCrlf,"<br>")
	Codificado = Trim(Codificado)
	'------------------------------------------------------------------------------	
	gravar = Codificado
	'------------------------------------------------------------------------------
End Function

'-------------------------------------------------------'
'## Função responsável pela gravação do texto no banco
'-------------------------------------------------------'
Function exibir(MsgIn)
	'------------------------------------------------------------------------------
	Codificado= MsgIn
	'------------------------------------------------------------------------------
	Codificado = replace(Codificado,CHR(34),"&quot;")
	Codificado = replace(Codificado,CHR(39),"'")
	Codificado = replace(Codificado,"<br>",VBCrlf)
	Codificado = replace(Codificado,"&lt;","<")
	Codificado = replace(Codificado,"&gt;",">")
	'Codificado = replace(Codificado,VBCrlf,"<br>")
	Codificado = ConvertToHtml(Codificado)
	'------------------------------------------------------------------------------	
	exibir = Codificado	
	'------------------------------------------------------------------------------
End Function

'----------------------------------------------------'
'** Funcao que inseri zeros a esquerda
'----------------------------------------------------'
Function MaxZeros(valor,max)
	MaxZeros = right(string(max,"0") & valor,max)
End Function 

'-------------------------------------------------
'*** FUNCAO DE SAUDAÇÃO
'-------------------------------------------------
Function Saudacao()
'-------------------------------------------------
Dim s
	Select case Hour(Time)
			case 0,1,2,3,4,5,6,7,8,9,10,11
			s = "Bom dia"
			case 12,13,14,15,16,18
			s = "Boa tarde"
			case else
			s = "Boa noite"
	End Select
'-------------------------------------------------
	s = s & ", <b>"&Session("NAME")&"</b>.<br>"&_
			" Login:<font color=red> "& Session("LOGIN_USER") &"</font>.<br />"	&_
			" Seu IP: <b>"& Request.ServerVariables("REMOTE_ADDR") &"</b>" 
'-------------------------------------------------		
	Saudacao = s
'-------------------------------------------------
End Function
'------------------------------------------------------
'** Traduz semana para o português
'------------------------------------------------------
Function SemanaPT(i)
	Select Case cInt(i)
			Case 1 : i = "Domingo"
			Case 2 : i = "Segunda"
			Case 3 : i = "Terça"
			Case 4 : i = "Quarta"
			Case 5 : i = "Quinta"
			Case 6 : i = "Sexta"
			Case 7 : i = "Sábado"					
	End Select
	SemanaPT  =i
End Function


Function CharInvalid(x)
	for i=1 to len(x)
		y = mid(x,i,1)
		select case y
				case "\","/",":","*","?","""","<",">","|","," : k = false
				case else k = true
		end select
		if k = true then z = z & y
	next
	CharInvalid = z
End Function


'-------------------------------------------------
'*** Calculo de horas somente para os dias úteis
'-------------------------------------------------
Function CalcHora(dtini,dtfim,tipo)
'-------------------------------------------------
'## NOTE Q DTINI DEVE SER A MENOR DATA
'## E POR SUA VEZ O DTFIM DEVE SER A MAIOR DATA
'## ISSO É OBRIGATORIO!!!!!!!!!!!!!!!!!!!!!!!!!!!
'-------------------------------------------------
on error resume next
'-------------------------------------------------
TotalD		= 86400 '### TOTAL DE SEG EM UM DIA
'-------------------------------------------------
	IF Tipo = 0 THEN
	'---------------------------------------------'
			HoraI		= Mid(dtini,Instr(dtini," ")+1,5)
			HoraF		= Mid(dtfim,Instr(dtfim," ")+1,5)
	'---------------------------------------------'
	'### TOTAL DE SEGUNDOS
	'---------------------------------------------'
		HoraIseg		= ((Left(HoraI,2))*3600) + ((Right(HoraI,2))*60) 
		HoraFseg		= ((Left(HoraF,2))*3600) + ((Right(HoraF,2))*60)
		Diferenca		=  (TotalD - HoraIseg) + HoraFseg
	'---------------------------------------------'
		dtini		= cdate(dtini)
		dtfim		= cdate(dtfim)
		dias		= datediff("d",dtini,dtFim)
		dias		= CInt(dias)		
	'---------------------------------------------'	
		IF dias > 0 THEN
	'---------------------------------------------'
			for i=1 to cint(dias)
	'---------------------------------------------'
				dtemp = dateAdd("d",i,dtini)
	'---------------------------------------------'
				Select Case WeekDay(dtemp)
						Case 1
							diaLimpo = diaLimpo
						Case 7
							diaLimpo = diaLimpo
						Case Else
							diaLimpo = diaLimpo + 1
				END Select 
'-------------------------------------------------'
			next	
'-------------------------------------------------'
		ELSE
'-------------------------------------------------'
			DiaLimpo = 0
'-------------------------------------------------'
		END IF	
'-------------------------------------------------'
		IF diaLimpo > 0 then
			diaLimpo	= diaLimpo - 1
			diaLimpos	= ((diaLimpo * TotalD)+ Diferenca)/ TotalD
		ELSE
			diaLimpos	= (Diferenca - TotalD) / TotalD
		end if
'-------------------------------------------------'
		diaLimpo	= round(diaLimpos,2)
'---------------------------------------------'	
	ELSEIF Tipo = 1 THEN
'---------------------------------------------'
		dtini		= cdate(dtini)
		dias		= dtFim
'---------------------------------------------'
		for i=1 to cint(dias)
	'---------------------------------------------'
			dtemp = dateAdd("d",i,dtini)
	'---------------------------------------------'
			Select Case WeekDay(dtemp)
					Case 1
						diaLimpo = diaLimpo + 2
					Case 7
						diaLimpo = diaLimpo + 2
					Case Else
						diaLimpo = diaLimpo + 1
			END Select 
'-------------------------------------------------'
		next
		diaLimpo = dateAdd("d",diaLimpo,dtini)
'---------------------------------------------'
	END IF
'-------------------------------------------------'
	CalcHora = 	diaLimpo
'-------------------------------------------------
End Function
'-------------------------------------------------
Function ConvertHoraToDate(valor)
'-------------------------------------------------
	xDados	=	Split(valor,",")
'------------------------------------------------------
	HrCh	=	xDados(0)
'------------------------------------------------------
	HrTp	=	(xDados(1) * 24) / 100
'------------------------------------------------------
	if instr(HrTp,",") > 0 then HrFr	=	Left(HrTp,Instr(HrTp,",")-1 ) 
'------------------------------------------------------
	if instr(HrMi,",") > 0 then
'------------------------------------------------------
		HrMi	=	Right(HrTp,Instr(HrTp,",")) 
		HrMi	=	round((HrMi * 60) / 100,0)
'------------------------------------------------------
	end if
'------------------------------------------------------
	'** Saida dos dados
'------------------------------------------------------
	if HrCh > 0 then xOqcdao = HrCh &" dia(s)"
'------------------------------------------------------
	if HrFr	> 0 then 
'------------------------------------------------------
		if xOqcdao <> "" then xOqcdao = xOqcdao & ", "
		xOqcdao = xOqcdao & HrFr &" hora(s)"
'------------------------------------------------------
	end if
'------------------------------------------------------
	if HrMi	> 0 then 
'------------------------------------------------------
		if xOqcdao <> "" then xOqcdao = xOqcdao & ", "
		xOqcdao = xOqcdao &  HrMi &" minuto(s)"
'------------------------------------------------------
	end if
'------------------------------------------------------
	ConvertHoraToDate = xOqcdao 
'-------------------------------------------------
End Function 
'-------------------------------------------------
Function MostraCampos(Get_Post)
'-------------------------------------------------
	sRet = ""
	sRet =  "<font face=tahoma size=2><h3>Listagem de variaveis postadas da página "& request.ServerVariables("HTTP_REFERER") &"</h3><hr size=1>"
	IF Get_Post = "Post" or Get_Post = "0" THEN
		for each item in request.form
			sRet = sRet & item &" = "& request.form (item) & "<br>"&VBCrlf
		next
	ELSE
		for each item in request.QueryString
			sRet = sRet & item &" = "& request.QueryString(item) & "<br>"&VBCrlf
		next
	END IF
	sRet = sRet &  "</font><hr size=1>"
	MostraCampos = sRet
'-------------------------------------------------
End Function
'-------------------------------------------------
'** Aplicando SLA
'-------------------------------------------------
Function ApplySla(codOccur)
'-------------------------------------------------
	strSQL =	" SELECT codMod, a.codUser, a.codStatus, dataOccur, codPro, DATEDIFF(hh, dataOccur, getdate()) AS tempo,  b.nameUser solicitante, c.nameUser atendente, d.nameStatus AS status, b.emailUser emailSolic, c.emailUser emailAtend, c.codAdm, e.nameadm depto, DATEDIFF(hh, f.dataseq, getdate()) tempo2   "&_
				" FROM ocorrencia a "&_
				" INNER JOIN usuarios b ON a.codUser = b.codUser "&_
				" INNER JOIN usuarios c ON a.codPro = c.codUser "&_
				" INNER JOIN status d ON a.codStatus = d.codStatus "&_
				" INNER JOIN ADMINISTRADORA E ON B.CODADM = E.CODADM "&_
				" LEFT JOIN sequencia f ON f.codOccur = a.codOccur  "&_
				" WHERE a.codOccur = '"& codOccur &"' "&_
				" AND  flagCancela = 0 AND flagClose = 0 "
	Set Rs = Conexao.execute(strSQL)
	if not Rs.eof Then
'-------------------------------------------------
		'** Variaveis de email 
'-------------------------------------------------
		solic	= Rs("solicitante")
		atend	= Rs("atendente")
		status	= Rs("status")
		depto	= Rs("depto")
		emailS	= Rs("emailSolic")
		emailA	= Rs("emailAtend")
'-------------------------------------------------
		Host				= "localhost"
		Componente			= "CDONTS"
		Email				= "sla.SHD@unimeds.com.br"
		NomeEmail			= "SHD"
'-------------------------------------------------
		strSQL	= "SELECT EMAILUSER FROM USUARIOS WHERE FLAGADMA = 1 AND CODADM = "& Rs("codAdm")
		Set Ry	= Conexao.execute(strSQL)
		If Not ry.eof then
'-------------------------------------------------
			do while not ry.eof
				emailC	= emailC & Ry(0) &";"
			ry.movenext
			loop
'-------------------------------------------------		
		End if	
'-------------------------------------------------
		strSQL	= "SELECT EMAILUSER FROM USUARIOS WHERE FLAGADMS = 1 AND CODADM = "& Session("CODADM")
		Set Ry	= Conexao.execute(strSQL)
		If Not ry.eof then
'-------------------------------------------------
			do while not ry.eof
				emailAd	= emailAd & Ry(0) &";"
			ry.movenext
			loop
'-------------------------------------------------		
		End if
'-------------------------------------------------
		Last_Date = cInt(Rs("Tempo")) : if IsNull(Rs("Tempo2")) = false then Last_Date = cInt(Rs("Tempo2"))
'-------------------------------------------------
'$$ Debug
'-------------------------------------------------
'		response.Write codOccur &" =  " & Rs("Tempo") &" : "& Rs("Tempo2") &"<br>"
'-------------------------------------------------
		strSQL =	" SELECT * "&_
					" FROM SLA_VS_REGRA "&_
					" WHERE IDMODULO = "& Rs("CodMod") &" AND SLA >= "& Last_Date &_
					" AND IDREGRA NOT IN ("&_
					"	SELECT IDREGRA FROM SLA_NOTIFICACAO WHERE CODOCCUR = '"& codOccur &"')"
		Set Rx = Conexao.execute(strSQL)
		if not Rx.eof then
'-------------------------------------------------
			do while not rx.eof
'-------------------------------------------------
'$$ Debug
'-------------------------------------------------
'				response.Write last_date &" >= "& cInt(Rx("SLA")) &" ( "
'				response.Write last_date >= cInt(Rx("SLA")) 
'				Response.Write ")<br>"				
'-------------------------------------------------
			  if Last_Date >= cInt(Rx("SLA")) then '@@@ remover o item AND XYQ = TRUE
'-------------------------------------------------
				'** Flag´s para notificação de email
'-------------------------------------------------
				F_SOL		= Rx("F_SOL")			'** notificar solicitante
				F_ATE		= Rx("F_ATE")			'** notificar atendente
				F_SUP		= Rx("F_SUP")			'** notificar supervisor
				F_ADM		= Rx("F_ADM")			'** notificar administrador SHD
'-------------------------------------------------
				'** Status para modificações
'-------------------------------------------------
				STT_AT		= Rx("CODSTATUS")		'** caso o Rs("codStatus") = STT_AT, verificar se o STT_MOD tem valor
				STT_MOD		= Rx("CODMODSTATUS")	'** caso tem valor o STT_MOD deverá ser atualizado na TAB OCORRENCIA
				MENSAGEM	= Rx("MENSAGEM")		'** e enviar o email com essa informação para 
'-------------------------------------------------
				'** Inicio das Regras
'-------------------------------------------------
					if Rs("codStatus") = STT_AT then
'-------------------------------------------------
						if STT_MOD <> 0 then
'-------------------------------------------------
							if STT_MOD = 1 then Y_sql = ", flagCancela = 1 "
							if STT_MOD = 2 then Y_sql = ", flagClose = 1 "
'-------------------------------------------------
							strSQL =	" UPDATE OCORRENCIA SET "&_
										" codStatus = "& STT_MOD & Y_sql &_
										" WHERE CODOCCUR = '"& codOccur &"'"
							 Conexao.execute(strSQL)
'-------------------------------------------------
							strSQL	=	" SELECT nameStatus FROM STATUS WHERE CODSTATUS = "& STT_MOD
							Set Tmp	=	Conexao.execute(strSQL)
							status	=	Tmp(0)
'-------------------------------------------------
						end if
'-------------------------------------------------
					end if
'-------------------------------------------------
				'** Preparando a mensagem para o envio de emails
'-------------------------------------------------					
				MENSAGEM	=	Replace(mensagem,"{occur}",codOccur)
				MENSAGEM	=	Replace(mensagem,"{solic}",solic)
				MENSAGEM	=	Replace(mensagem,"{atend}",atend)
				MENSAGEM	=	Replace(mensagem,"{status}",status)
				MENSAGEM	=	Replace(mensagem,"{depto}",depto)				
				Mensagem	=	"<div style='font: 12px; font-family:Verdana, Arial, Helvetica, sans-serif;width:100%;border:1px #FFF solid;'>"& _
								MENSAGEM &_
								"<br /><br />Atenciosamente,<br / ><br/ >"&_
								"Sistema de HelpDesk - SHD<br / ><br/ >"&_
								"<a href='http://"& Request.ServerVariables("Server_Name") &"'>Acesse o SHD agora.</a><br / ><br/ >"&_
								"<b>E-mail automático, favor não responder este e-mail.</b>"&_
								"</div>"	
'-------------------------------------------------
				'** Enviando email com as informações
'-------------------------------------------------
				if F_SOL = 1 and emailS <> "" then EnviaEmail Host,Componente,Email,NomeEmail,emailS,"SHD - Alerta do chamado "& codOccur,Mensagem
				if F_ATE = 1 and emailA <> "" then EnviaEmail Host,Componente,Email,NomeEmail,emailA,"SHD - Alerta de SLA ["& codOccur &"]",Mensagem
				if F_SUP = 1 and emailC <> "" then EnviaEmail Host,Componente,Email,NomeEmail,emailC,"SHD - Alerta de SLA ["& codOccur &"]",Mensagem
				if F_ADM = 1 and emailAd <> "" then EnviaEmail Host,Componente,Email,NomeEmail,emailAd,"SHD - Alerta de SLA ["& codOccur &"]",Mensagem
'-------------------------------------------------
				'** Gravando o log na tabela
'-------------------------------------------------
				strSQL	=	" INSERT SLA_NOTIFICACAO "&_
							" (IDREGRA,CODOCCUR) "&_
							" VALUES "&_
							" ("& Rx(2) &",'"& codOccur &"') "
				Conexao.execute(strSQL)
'-------------------------------------------------
			  end if
'-------------------------------------------------
				'** Fim das Regras				
'-------------------------------------------------
			rx.movenext
			loop
'-------------------------------------------------
		end if
'-------------------------------------------------	
	END IF
'-------------------------------------------------
End Function 
'-------------------------------------------------
Function ConvertToHtml(valor)
		valor = replace(valor,"&#193;","á")
		valor = replace(valor,"&#192;","à")
		valor = replace(valor,"&#195;","ã")
		valor = replace(valor,"&#196;","ä")
		valor = replace(valor,"&#225;","á")
		valor = replace(valor,"&#224;","à")
		valor = replace(valor,"&#227;","ã")
		valor = replace(valor,"&#228;","ä")
		valor = replace(valor,"&#233;","é")
		valor = replace(valor,"&#232;","è")
		valor = replace(valor,"&#235;","ë")
		valor = replace(valor,"&#201;","É")
		valor = replace(valor,"&#200;","È")
		valor = replace(valor,"&#203;","Ë")
		valor = replace(valor,"&#237;","í")
		valor = replace(valor,"&#236;","ì")
		valor = replace(valor,"&#239;","ï")
		valor = replace(valor,"&#205;","Í")
		valor = replace(valor,"&#204;","Ì")
		valor = replace(valor,"&#207;","Ï")
		valor = replace(valor,"&#243;","ó")
		valor = replace(valor,"&#242;","ò")
		valor = replace(valor,"&#245;","õ")
		valor = replace(valor,"&#246;","ö")
		valor = replace(valor,"&#211;","Ó")
		valor = replace(valor,"&#210;","Ò")
		valor = replace(valor,"&#213;","Õ")
		valor = replace(valor,"&#214;","Ö")
		valor = replace(valor,"&#250;","ú")
		valor = replace(valor,"&#249;","ù")
		valor = replace(valor,"&#252;","ü")
		valor = replace(valor,"&#218;","Ú")
		valor = replace(valor,"&#217;","Ù")
		valor = replace(valor,"&#220;","Ü")
	ConvertToHtml = valor
End Function

Function IconePrioridade(sPrior)
 
	for i=Lbound(Prioridade) to Ubound(Prioridade)
		if sPrior = Prioridade(i,0) then 
			sPriorLabel = Prioridade(i,1)
			PriorIx		= i
			exit for
		end if
	next
	IconePrioridade = "&nbsp;<img src='"& sPrior &"' border='0' alt='"& sPriorLabel &"' name='wicon' width=22 height=22   title='"& sPriorLabel &"'>"
End Function

'-------------------------------------------------
if Session("DisabledLoadVars") = False then %>
<!--#include file="bib_vars.asp"-->
<%'-------------------------------------------------
end if%>