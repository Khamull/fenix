<%@LANGUAGE="VBSCRIPT" CODEPAGE="28592"%>
<%option explicit%>

<!--#include file="inc/inc_conexao.inc"-->

<%
 Call abreConexao()
%>

<%
'VARIAVEL DE MENSAGEM
Dim msg
%>

<%
'RECUPERA O ID DA VENDA (venID)
Dim venID
venID = Request.QueryString("venID")
%>

<%
'RECUPERA O TIPO DE VENDA
Dim tipo
if(Not isEmpty(Request.QueryString("tipo")))Then
	tipo = Request.QueryString("tipo")
else
	tipo = Request.Form("tipo")
end if
%>


<%
'VERIFICA SE FOI CLICADO NO BOTAO << Logar >>
if(Not isEmpty(Request.Form("ACESSAR")))Then


'RECUPERA DADOS DO FORMULÁRIO
Dim login
Dim senha

login = Request.Form("login")
senha = Request.Form("senha")


'VERIFICA USUÁRIO E SENHA DO ADMINISTRADOR
Dim rs01
Dim sql01

set rs01 = Server.CreateObject("ADODB.Recordset")
sql01 = "SELECT * FROM tb_usuario WHERE usuLogin = '"&login&"' AND usuSenha = '"&senha&"' AND nivelID = '1'"
set rs01 = conn.execute(sql01)


if(Not rs01.EoF)Then

	'CASO EXISTA ESSE USUARIO
	'O PEDIDO É EXCLUIDO PELO ADMINISTARDOR
	
	'------------ Registra a Desistencia -----------------
	  Dim rsx1
	  Dim sqlx1
	  Dim rsy1
	  Dim sqly1
	  Dim caixaID
	  
	  set rsx1 = Server.CreateObject("ADODB.Recordset")
	  sqlx1 = "SELECT * FROM tb_caixa WHERE status = 'A'"
	  set rsx1 = conn.execute(sqlx1)
	  
	   caixaID = rsx1.fields.item("caixaID").value
	  
	  set rsy1 = Server.CreateObject("ADODB.Recordset")
	   '********************* VERIFICA O TIPO DE VENDA *******************************
	   if(tipo = "telefone")Then
	   	sqly1 = "INSERT INTO tb_cancelados (caixaID, telefone) VALUES ('"&caixaID&"', '1')"
	   elseif(tipo = "balcao")Then
	   	sqly1 = "INSERT INTO tb_cancelados (caixaID, balcao) VALUES ('"&caixaID&"', '1')"
	   elseif(tipo = "mesa")Then
	    sqly1 = "INSERT INTO tb_cancelados (caixaID, mesa) VALUES ('"&caixaID&"', '1')"
	   end if
	   '******************************************************************************
	  
	  set rsy1 = conn.execute(sqly1)
	 '------------------------------------------------------
	
	Dim rs08
	Dim sql08
	
	set rs08 = server.createObject("adodb.recordset")
	sql08 = "DELETE FROM tb_venda WHERE venID = '"&Request.Form("venID")&"'"
	set rs08 = conn.execute(sql08)
	
	response.redirect("pg_excluir_pedidos.asp?excluido=ok")

else

 Response.Redirect("pg_excluir_pedidos.asp?erro=1&venID="&Request.Form("venID")&"&tipo="&tipo)

end if

end if
%>



<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-2" />
<title>Excluir Pedido</title>

<style type="text/css">

body, html{
	font-family:Arial, Helvetica, sans-serif;
	background-color:#EEEEEE;
	padding:0px;
	margin:0px;
}


#contorno{
	width:280px;
	height:240px;
	border:1px dashed #333;
	background-color:#FFF;
}

</style>


<script type="text/javascript">  
// funçao usada para carregar o código  
function fecha() {  
// fechando a janela atual ( popup )  
window.close();  
// dando um refresh na página principal  
//opener.location.href=opener.location.href;  
  opener.location.href="pg_menu.asp";
/* ou assim:  
* window.opener.location.reload(); 
*/  
//document.location="pg_menu.asp"  
// fim da funçao  
}  
</script>

</head>

<center />

<body>

<%if(Request.QueryString("excluido") <> "")Then%>

<br />
<font color="#FF0000">PEDIDO EXCLUIDO COM SUCESSO!</font>
<br />
<br />
<b><a href="javascript:void(0)" onclick="fecha()">FECHAR</a></b>

<%else%>


<div id="contorno">

<form method="post" name="form1" action="pg_excluir_pedidos.asp">

 <br />
 <table border="0" cellpadding="1" cellspacing="1" align="center" width="246" >
  <tr>
   <td height="79" colspan="2" align="center">
     <p>Para excluir um pedido que j&aacute; foi fechado &eacute; necess&aacute;rio fornecer o login e senha do Administrador </p></td>
  </tr>
  <tr>
   <td colspan="2" height="10" align="center">
   <font color="#FF0000">
   <%
    if(Request.QueryString("erro") = "1")Then
    	Response.Write("Login ou Senha Incorretos!")
	end if
   %>
   </font>
   </td>
  </tr>
  <tr>
   <td colspan="2" align="center" bgcolor="#CCCCCC"><strong>ADMINISTRADOR</strong></td>
  </tr>
  <tr>
   <td width="68" align="left">Login</td>
   <td width="166" align="left"><input type="text" name="login" maxlength="20" /></td>
  </tr>
  <tr>
   <td align="left">Senha</td>
   <td align="left"><input type="password" name="senha" maxlength="20" /></td>
  </tr>
  <tr>
   <td colspan="2" align="center">
    <input type="hidden" name="venID" value="<%=Request.QueryString("venID")%>" />
   	<input type="hidden" name="tipo" value="<%=Request.QueryString("tipo")%>" />
    <input type="submit" name="ACESSAR" value="Logar" />
   </td>
  </tr>
 </table>

</form>

</div>

<%end if%>

</body>
</html>

<%
 Call fechaConexao
%>