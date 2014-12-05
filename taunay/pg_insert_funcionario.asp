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
Dim funNome
Dim funTelefone
Dim carID
Dim funCpf

usuLogin 	=	Request.Form("usuLogin")
funNome		=	UCase(Trim(Replace(Request.Form("funNome"),"'","")))
funCpf		=	UCase(Trim(Replace(Request.Form("funCpf"),"'","")))
carID		=	Trim(Replace(Request.Form("carID"),"'",""))
funTelefone	=	Trim(Replace(Request.Form("funTelefone"),"'",""))

%>

<%
if (not isEmpty(Request.Form("cadastrar"))) then

	Dim rs00
	Dim sql00
	set rs00 = server.CreateObject("adodb.recordset")
		sql00 = "SELECT * FROM tb_funcionario WHERE (funNome = '"&funNome&"' AND funCpf = '"&funCpf&"') OR (funNome = '"&funNome&"' AND carID = '"&carID&"') OR funCpf = '"&funCpf&"'"
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
			sql01 = "INSERT INTO tb_funcionario (usuLogin, funNome, funCpf, funTelefone, carID) VALUES ('"&usuLogin&"','"&funNome&"','"&funCpf&"','"&funTelefone&"','"&carID&"')"
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
	sql02 = "SELECT * FROM tb_cargo where carAtivo = 'S' ORDER BY carDescr"
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
	
var funNome
var funCpf
var funTelefone
var carID

funNome = document.form1.funNome.value;
funCpf = document.form1.funCpf.value;
funTelefone = document.form1.funTelefone.value;
carID = document.form1.carID.value;

if (funNome.length < 3) {
	alert("Favor informar o nome do funcionário!");
	document.form1.funNome.focus();
	return false;
	}
if (funCpf.length < 11) {
	alert("Favor informar o Cpf do funcionário corretamente!");
	document.form1.funCpf.focus();
	return false;
	}	
if (funTelefone.length < 8) {
	alert("Favor informar o Telefone do funcionário!");
	document.form1.funTelefone.focus();
	return false;
	}	
if (carID == "") {
	alert("Favor selecionar o cargo!");
	document.form1.carID.focus();
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
	    <li><a href="pg_select_funcionario.asp">Lista Funcionarios</a></li>         
	    <li><a href="pg_insert_cargo.asp">Cargo</a></li>              
	  </ul>
	</div>
	<div id="areaPrincipal">
    <div style="height:25px; line-height:25px; background:#ccc"></div>
    <table width="96%" border="0" align="left" cellpadding="3" cellspacing="3">
  <tr>
    <td width="50" align="center"><img src="ico/ico_garcon.gif" width="60" height="60" class="icone" /></td>
    <td width="751" align="center" class="titulo">CADASTRO DE FUNCIONÁRIO</td>
    </tr>
  <tr>
    <td colspan="2" align="center"><form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>" onsubmit="return verForm(this)">
      <table width="777" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td height="25" colspan="4" align="center"><%=msg%></td>
        </tr>
        <tr>
          <td width="109" height="25" align="right">Funcionário:</td>
          <td width="267" height="25" align="left"><label>
            <input name="funNome" type="text" id="funNome" size="40" maxlength="30" />
          *</label></td>
          <td width="110" height="25" align="right">CPF:</td>
          <td width="263" height="25" align="left"><input name="funCpf" type="text" id="funCpf" size="11" maxlength="11" />
*</td>
        </tr>
        <tr>
          <td height="25" align="right">Cargo:</td>
          <td height="25" align="left"><label>
            <select name="carID" id="carID">
              <option value=""></option>
              <%
		  if (not rs02.eof) then
          do while not rs02.eof
		  %>
              <option value="<%=rs02.fields.item("carID").value%>"><%=rs02.fields.item("carDescr").value%></option>
              <%
		  rs02.moveNext 
		  Loop
		  end if
		  %>
            </select>
*</label></td>
          <td height="25" align="right">Telefone:</td>
          <td height="25" align="left"><input name="funTelefone" type="text" id="funTelefone" size="8" maxlength="9" />
*</td>
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
call fechaConexao()
%>