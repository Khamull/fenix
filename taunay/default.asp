<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>



<%option explicit%>



<!--#include file="inc/inc_conexao.inc"-->



<!--#include file="inc/inc_formato_data.inc"-->



<%

Dim erro

Dim msg

erro = request.querystring("erro")

if (erro = 1) Then

msg = "<font color='red'>Login ou Senha Inválido!</font>"

else

msg = "Entre com seus dados"

end if

%>





<%

call abreConexao()

%>



<%

'---- CONSULTA PERMISSÃO DE ACESSO

Dim login

Dim senha

Dim usuLogin

Dim usuID



login = Trim(Request.Form("login"))

senha = Trim(Request.Form("senha"))



if (not isEmpty(Request.Form("ACESSAR"))) Then



Dim rs01

Dim sql01



	set rs01 = Server.CreateObject("ADODB.Recordset")

	sql01 = "SELECT usuID, usuLogin, usuSenha, nivelID, usuAtivo FROM tb_usuario WHERE usuLogin='"&login&"' AND usuSenha='"&senha&"' AND usuAtivo = 'S'"

	set rs01 = conn.Execute(sql01)

	

	if (not rs01.EOF OR not rs01.BOF) Then

		

		Dim acesso





		Session("acesso") = "confirmado"

		Session("usuLogin") = rs01.fields.item("usuLogin").value

		Session("nivelID") = rs01.fields.item("nivelID").value

		Session("usuID") = rs01.fields.item("usuID").value



		

		rs01.Close

		set rs01 = Nothing

		

		'//********************************//

		

		Dim aceData

		Dim aceHora

		Dim usuIP

		

		usuLogin 	= Session("usuLogin")

		aceData 	= data

		aceHora 	= time()

		usuIP	= Request.ServerVariables("REMOTE_ADDR")

		

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

		

		'//********************************//

			

		

		Response.redirect("pg_menu.asp")

		Response.end()

		

	else 

	

		rs01.Close

		set rs01 = Nothing

		

		Response.redirect("Default.asp?erro=1")

		Response.end()	

			

	end if

	

end if



%>



<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">

<head>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

<title>SISTEM FORTE EM MÍDIA</title>

<link href="css/css1.css" rel="stylesheet" type="text/css" />



<script language="javascript" type="text/javascript">



function verForm(form1){

	

var login

var senha



Login = document.form1.Login.value;

Senha = document.form1.Senha.value;



if (Login.length < 3) {

	alert("Login ou Senha Inválido!");

	document.form1.Login.focus();

	return false;

	}

	

if (Senha.length < 3) {

	alert("Login ou Senha Inválido!");

	document.form1.Senha.focus();

	return false;



	}	



}

</script>



</head>

<body>

<!--LAYOUT-->

<div id="container">

<!-- -->

<div id="topo"></div>

<div id="tituloBar"><img src="img/img_titulo_mp.gif" width="200" height="30" /></div>

<div id="corpo">

<!-- -->

<div id="areaConteudo">



	<div id="areaMenuVerfical">

	<div style="height:25px; line-height:25px; background:#ccc">Menu</div>    

	  <ul>

	    <li><a href="default.asp">Login</a></li>

	    <li><a href="#">Sair</a></li>

	  </ul>

	</div>

	<div id="areaPrincipal">

    <div style="height:25px; line-height:25px; background:#ccc">Acesso Restríto</div>

    <form id="form1" name="form1" method="post" action="default.asp" onsubmit="return verForm(this)">

      <table width="234" border="0" align="center" cellpadding="2" cellspacing="2">

    <tr>

      <td height="22"><img src="ico/ico_cadeado.gif" alt="Acesso Restrito" width="60" height="60" class="icone" /></td>

      <td height="22" align="center"><%=msg%></td>

    </tr>

    <tr>

      <td width="50" height="22" align="right">&nbsp;Login:</td>

      <td width="170" height="22"><label for="Login"></label>

        <input type="text" name="Login" id="Login" /></td>

    </tr>

    <tr>

      <td height="22" align="right">&nbsp;Senha:</td>

      <td height="22"><input type="password" name="Senha" id="Senha" /></td>

    </tr>

    <tr>

      <td height="22" colspan="2" align="center"><input name="aceData" type="hidden" id="aceData" value="<%=DATA%>" />

        <input name="ACESSAR" type="submit" class="botao" id="ACESSAR" value="Acessar" /></td>

    </tr>

  </table>

</form>

    </div>

</div>

<!-- -->

</div>

<div id="rodape"><br /><!--#include file="inc/inc_status.inc"--><br /></div>

</div>

<!--FIM DO LAYOUT-->



</body>

</html>



<%

call fechaConexao()

%>

