<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%option explicit%>

<!--#include file="../inc/inc_conexao.inc"-->
<!--#include file="../inc/inc_formato_data.inc"-->

<%

Dim login 
Dim senha 
Dim rs01
Dim sql01
Dim errorMessage

login = lcase(Request("login"))
senha = lcase(Request("senha"))

call abreConexao()


set rs01 = Server.CreateObject("ADODB.Recordset")

sql01 = "SELECT usuID, usuLogin, usuSenha, nivelID, usuAtivo FROM tb_usuario WHERE usuLogin='"&login&"' AND usuSenha='"&senha&"' AND usuAtivo = 'S'"

set rs01 = conn.Execute(sql01)

if (not rs01.EOF OR not rs01.BOF) Then

	Dim acesso

	Session("acesso") = "confirmado"

	Session("usuLogin") = rs01.fields.item("usuLogin").value

	Session("nivelID") = rs01.fields.item("nivelID").value

	Session("usuID") = rs01.fields.item("usuID").value

	Session("isMobile") = "true"

	rs01.Close

	set rs01 = Nothing


	Dim usuLogin
	Dim aceData
	Dim aceHora
	Dim usuIP

	usuLogin = Session("usuLogin")

	aceData = data

	aceHora = time()

	usuIP = Request.ServerVariables("REMOTE_ADDR")

	Dim RS02
	Dim SQL02
	
	SET RS02 = Server.CreateObject("ADODB.Recordset")		

	SQL02	=	"INSERT INTO"
	SQL02	=	SQL02	&	" tb_acesso "
	SQL02	=	SQL02	&	"("
	SQL02	=	SQL02	&	"aceData, "
	SQL02	=	SQL02	&	"aceHora, "
	SQL02	=	SQL02	&	"usuLogin, "
	SQL02	=	SQL02	&	"usuIP"
	SQL02	=	SQL02	&	")"
	SQL02	=	SQL02	&	" VALUES "
	SQL02	=	SQL02	&	"("
	SQL02	=	SQL02	&	"'"&aceData&"', "
	SQL02	=	SQL02	&	"'"&aceHora&"', "
	SQL02	=	SQL02	&	"'"&usuLogin&"', "
	SQL02	=	SQL02	&	"'"&usuIP&"'"
	SQL02	=	SQL02	&	")"

	SET RS02 = conn.Execute(SQL02)
	SET RS02 = Nothing

	response.write "{success:true, error:false, message:""}"


else
	
	errorMessage = "Login ou Senha Inválido"
	response.write "success:false, error:true, message:" + errorMessage

end if


%>

