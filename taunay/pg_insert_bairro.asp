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
Dim baiNome
Dim cidID
Dim baiFrete

usuLogin 	=	Request.Form("usuLogin")
baiNome		=	UCase(Trim(Replace(Request.Form("baiNome"),"'","")))
cidID		=	Trim(Replace(Request.Form("cidID"),"'",""))
baiFrete	=	Trim(Replace(Request.Form("baiFrete"),"'",""))
baiFrete 	= 	Replace(baiFrete,",",".")
%>

<%
if (not isEmpty(Request.Form("cadastrar"))) then

	Dim rs00
	Dim sql00
	set rs00 = server.CreateObject("adodb.recordset")
		sql00 = "SELECT * FROM tb_bairro WHERE baiNome = '"&baiNome&"' AND cidID = '"&cidID&"'"
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
			sql01 = "INSERT INTO tb_bairro (usuLogin, baiNome, cidID, baiFrete) VALUES ('"&usuLogin&"','"&baiNome&"','"&cidID&"','"&baiFrete&"')"
		set rs01 = conn.execute(sql01)
		set rs01 = nothing
		
		'****************************
		
		rs00.close
		set rs00 = nothing
		msg = "Os Dados foram cadastrados com sucesso!"
		
	end if
	
end if

%>

<%
Dim rs02
Dim sql02
set rs02 = server.CreateObject("adodb.recordset")
	sql02 = "SELECT * FROM tb_cidade where cidAtiva = 'S' ORDER BY cidNome"
set rs02 = conn.execute(sql02)	
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>SISTEM FORTE EM MÍDIA</title>
<link href="css/css1.css" rel="stylesheet" type="text/css" />

<script language="javascript" type="text/javascript">

function verForm(form1){
	
var baiNome
var cidID
var baiFrete

baiNome = document.form1.baiNome.value;
cidID = document.form1.cidID.value;
baiFrete = document.form1.baiFrete.value;

if (baiNome.length < 3) {
	alert("Favor Informar o Bairro!");
	document.form1.baiNome.focus();
	return false;
	}
	
if (cidID.length < 1) {
	alert("Favor Selecionar a Cidade!");
	document.form1.cidID.focus();
	return false;
	}	
if (baiFrete.length < 4) {
	alert("Favor Informar o valor do frete!");
	document.form1.baiFrete.focus();
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
	    <li><a href="pg_select_bairro.asp">Listar Bairros</a></li>         
	    <li><a href="pg_insert_cidade.asp">Cidade</a></li>
	    <li><a href="pg_insert_cliente.asp">Cliente</a></li>  
	    <li><a href="<%=Request.ServerVariables("script_name")%>?funcao=sair">Sair</a></li>                        
	  </ul>
	</div>
	<div id="areaPrincipal">
    <div style="height:25px; line-height:25px; background:#ccc"></div>
    <table width="96%" border="0" align="left" cellpadding="3" cellspacing="3">
  <tr>
    <td width="50" align="center"><img src="ico/ico_casabairro.gif" width="60" height="60" class="icone" /></td>
    <td width="751" align="center" class="titulo">CADASTRO DE BAIRRO</td>
    </tr>
  <tr>
    <td colspan="2" align="center"><form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>" onsubmit="return verForm(this)">
      <table width="777" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td height="25" colspan="4" align="center"><%=msg%></td>
        </tr>
        <tr>
          <td width="109" height="25" align="right">Bairro:</td>
          <td width="267" height="25" align="left"><label>
            <input name="baiNome" type="text" id="baiNome" size="40" maxlength="30" />
          *</label></td>
          <td width="110" height="25" align="right">&nbsp;</td>
          <td width="263" height="25" align="left">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" align="right">Cidade:</td>
          <td height="25" align="left">
          
          <select name="cidID" id="cidID">
          <option value=""></option>
          <%
		  if (not rs02.eof) then
          do while not rs02.eof
		  %>
           <option value="<%=rs02.fields.item("cidID").value%>"><%=rs02.fields.item("cidNome").value%></option>
          <%
		  rs02.moveNext 
		  Loop
		  end if
		  %>
          </select>
*</td>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="left">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" align="right">Valor do Frete:</td>
          <td height="25" align="left"><input name="baiFrete" type="text" id="baiFrete" value="0.00" size="10" maxlength="10" />
*</td>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="left">&nbsp;</td>
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
rs02.close
set rs02 = nothing
%>
<%
call FechaConexao()
%>