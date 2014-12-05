<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%option explicit%>

<!--#include file="Connections/conn.asp" -->

<!--#include file="inc/inc_conexao.inc"-->

<!--#include file="inc/inc_formato_data.inc"-->

<!--#include file="inc/inc_acesso.inc" -->

<%
call abreConexao()
%>

<%
Dim msg
Dim usuID
Dim usuLogin
Dim usuSenha
Dim nivelID
Dim rs02
Dim sql02
Dim rs01
Dim sql01

usuID = request.QueryString("usuID")
usuLogin = UCase(request.Form("usuLogin"))
usuSenha = request.Form("usuSenha")
nivelID = request.Form("nivelID")

set rs01 = Server.CreateObject("ADODB.Recordset")
sql01 = "Select * FROM tb_usuario WHERE usuID = '"&usuID&"'"
set rs01 = conn.execute(sql01)

if (not IsEmpty(request.Form("atualizar"))) Then

set rs02 = Server.CreateObject("ADODB.Recordset")
sql02 = "UPDATE tb_usuario SET usuLogin = '"&usuLogin&"', usuSenha = '"&usuSenha&"', nivelID = '"&nivelID&"' WHERE usuID = '"&usuID&"'"
set rs02 = conn.Execute(sql02)

response.Write("<script>alert('Usuario Alterado com Sucesso!');")
response.write("window.location.href = 'pg_select_usuario.asp'")
response.Write("</script>")

end if

%>



<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>SISTEM FORTE EM MÍDIA</title>
<link href="css/css1.css" rel="stylesheet" type="text/css" />

<script language="javascript" type="text/javascript">
/*
function verForm(form1){
	
var cliNome		= document.form1.cliNome.value;
var cliEndereco	= document.form1.cliEndereco.value;
var cidID		= document.form1.cidID.value;
var baiID		= document.form1.baiID.value;
var cliCpf		= document.form1.cliCpf.value;
var cliTelefone	= document.form1.cliTelefone.value;

if (cliNome.length < 3) {
	alert("Favor informar o nome completo!");
	document.form1.cliNome.focus();
	return false;
	}
	
if (cliEndereco.length < 3) {
	alert("Favor informar o endereço completo!");
	document.form1.cliEndereco.focus();
	return false;
	}	
	
if (cidID == "") {
	alert("Favor selecionar a cidade!");
	document.form1.cidID.focus();
	return false;
	}		
if (baiID == "") {
	alert("Favor selecionar o bairro!");
	document.form1.baiID.focus();
	return false;
	}		
if (cliCpf.length < 10) {
	document.form1.cliCpf.value = document.form1.cliCod.value;
	}			
if (cliTelefone.length < 8) {
	alert("Favor informar o Telefone corretamente!");
	document.form1.cliTelefone.focus();
	return false;
	}		
}
function carregaBairro()
{


var cidNome 	= document.form1.cidID.value;
cidNome 		= document.getElementById('cidID');
cidNome 		= cidNome.options[cidNome.selectedIndex].text;
cidID 			= document.getElementById('cidID').value;

var cliNome 	= document.form1.cliNome.value;
var cliEndereco	= document.form1.cliEndereco.value;
var baiID		= document.form1.baiID.value;
var cliCpf		= document.form1.cliCpf.value;
var cliTelefone	= document.form1.cliTelefone.value;
var cliEndRef 	= document.form1.cliEndRef.value;
var cliEmail	= document.form1.cliEmail.value;
var cliSenha 	= document.form1.cliSenha.value;
var cliCod 		= document.form1.cliCod.value;

window.location.href =  "pg_update_usuario.asp?cliNome="+cliNome+"&cliEndereco="+cliEndereco+"&cidID="+cidID+"&cidNome="+cidNome+"&baiID="+baiID+"&cliCpf="+cliCpf+"&cliTelefone="+cliTelefone+"&cliEmail="+cliEmail+"&cliEndRef="+cliEndRef+"&cliSenha="+cliSenha+"&cliCod="+cliCod+"&cliID="+cliID

}
*/
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
	    <li><a href="pg_menu.asp">Menu Principal</a></li>
	    <li><a href="pg_select_usuario.asp">Listar Usuários</a></li>        
             
	  </ul>
	</div>
	<div id="areaPrincipal">
    <div style="height:25px; line-height:25px; background:#ccc"></div>
    <table width="96%" border="0" align="left" cellpadding="3" cellspacing="3">
  <tr>
    <td width="50" align="center"><img src="ico/ico_usuario.gif" width="60" height="60" class="icone" /></td>
    <td width="751" align="center" class="titulo">ATUALIZAR USUÁRIO</td>
    </tr>
  <tr>
    <td colspan="2" align="center"><form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?usuID=<%=usuID%>"
    onsubmit="return verForm(this)">
      <table width="777" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td height="25" colspan="4" align="center"><%=msg%></td>
        </tr>
        <tr>
          <td width="109" height="25" align="right">Login:</td>
          <td width="267" height="25" align="left"><label>
            <input name="usuLogin" type="text" value="<%=(rs01.Fields.Item("usuLogin").Value)%>" size="40" maxlength="30" />
          *</label></td>
          <td width="110" height="25" align="right">Nível</td>
          <td width="263" height="25" align="left">
           <%
			Dim nivelDescricao
			if(rs01.Fields.Item("nivelID").Value = "1")Then
			 nivelDescricao = "Administrador"
			elseif(rs01.Fields.Item("nivelID").Value = "2")Then
			 nivelDescricao = "Vendedor"
			elseif(rs01.Fields.Item("nivelID").Value = "3")Then
			 nivelDescricao = "Caixa"
			end if
		   %>
           <select name="nivelID">           
            <option value="<%=(rs01.Fields.Item("nivelID").Value)%>"><%=nivelDescricao%></option>
           <%if(Session("nivelID") = "1")Then%>
            <option value="1">Administrador</option>
            <option value="2">Vendedor</option>
            <option value="3">Caixa</option>
           <%end if%>
           </select>
          </td>
        </tr>
        <tr>
          <td height="25" align="right">Endereço:</td>
          <td height="25" align="left"><input name="usuSenha" type="text" value="<%=(rs01.Fields.Item("usuSenha").Value)%>" size="40" maxlength="50" />
            *</td>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="left">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="left"><input name="atualizar" type="submit" class="botao" id="atualizar" value="Atualizar" /></td>
          <td height="25" align="right"><input name="usuLogin2" type="hidden" id="usuLogin2" value="<%=Session("usuLogin")%>" /></td>
          <td height="25" align="left">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="left">&nbsp;</td>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="left">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="left">&nbsp;</td>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="left">&nbsp;</td>
        </tr>
        <tr>
          <td height="22" align="right">&nbsp;</td>
          <td height="22">&nbsp;</td>
          <td height="22" align="right">&nbsp;</td>
          <td height="22">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" align="right">&nbsp;</td>
          <td height="25">&nbsp;</td>
          <td height="25" align="right">&nbsp;</td>
          <td height="25">&nbsp;</td>
        </tr>
      </table>
    </form></td>
  </tr>
  </table>
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
call FechaConexao()
%>
