<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%option explicit%>

<!--#include file="inc/inc_conexao.inc"-->

<!--#include file="inc/inc_formato_data.inc"-->

<!--#include file="inc/inc_acesso.inc" -->

<%
call abreConexao()
%>

<%
Dim msg
Dim usuLogin
Dim tipDescr

usuLogin 		=	Request.Form("usuLogin")
tipDescr		=	Ucase(Trim(Replace(Request.Form("tipDescr"),"'","")))

%>

<%
if (not isEmpty(Request.Form("cadastrar"))) then

	Dim rs00
	Dim sql00
	set rs00 = server.CreateObject("adodb.recordset")
		sql00 = "SELECT * FROM tb_tipo WHERE tipDescr = '"&tipDescr&"'"
	set rs00 = conn.Execute(sql00)
	
	if (not rs00.eof) then
	
	'Mensagem
	rs00.close
	set rs00 = nothing
	msg = "Já existe um registro com os dados informados!"	
	
	else
	
	'Grava
		Dim rs01
		Dim sql01
		set rs01 = server.CreateObject("adodb.recordset")
			sql01 = "INSERT INTO tb_tipo (usuLogin, tipDescr) VALUES ('"&usuLogin&"','"&tipDescr&"')"
		set rs01 = conn.execute(sql01)
		set rs01 = nothing
		
		'****************************
		
		rs00.close
		set rs00 = nothing
		msg = "Os Dados foram cadastrados com sucesso!"
		
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
	
var tipDescr

tipDescr = document.form1.tipDescr.value;

if (tipDescr.length < 3) {
	alert("Favor informar o tipo de produto!");
	document.form1.tipDescr.focus();
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
	    <li><a href="pg_menu.asp">Menu Principal</a></li>
	    <li><a href="pg_select_tipoProduto.asp">Listar Tipos</a></li>  
	    <li><a href="pg_insert_produto.asp">Produto</a></li>                   
	  </ul>
	</div>
	<div id="areaPrincipal">
    <div style="height:25px; line-height:25px; background:#ccc"></div>
    <table width="96%" border="0" align="left" cellpadding="3" cellspacing="3">
  <tr>
    <td width="50" align="center"><img src="ico/ico_pizza.gif" width="60" height="60" class="icone" /></td>
    <td width="751" align="center" class="titulo">CADASTRO DE TIPO DE PRODUTO</td>
    </tr>
  <tr>
    <td colspan="2" align="center"><form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>" onsubmit="return verForm(this)">
      <table width="777" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td height="25" colspan="4" align="center"><%=msg%></td>
        </tr>
        <tr>
          <td width="109" height="25" align="right">Tipo de Produto:</td>
          <td width="267" height="25" align="left"><label>
            <input name="tipDescr" type="text" id="tipDescr" size="40" maxlength="30" />
          *</label></td>
          <td width="110" height="25" align="right">&nbsp;</td>
          <td width="263" height="25" align="left">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" align="right">&nbsp;</td>
          <td height="25"><input name="cadastrar" type="submit" class="botao" id="cadastrar" value="Cadastrar" /></td>
          <td height="25" align="right"><input name="usuLogin" type="hidden" id="usuLogin" value="<%=Session("usuLogin")%>" /></td>
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
call fechaConexao()
%>