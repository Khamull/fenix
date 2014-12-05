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
Dim forNome
Dim forEndereco
Dim forBairro
Dim forCidade
Dim forUf
Dim forCep
Dim forCnpj
Dim forIE
Dim forEmail
Dim forTelefone
Dim forContato

usuLogin 		=	Request.Form("usuLogin")
forNome 		=	UCase(Trim(Replace(Request.Form("forNome"),"'","")))
forEndereco		=	UCase(Trim(Replace(Request.Form("forEndereco"),"'","")))
forBairro		=	UCase(Trim(Replace(Request.Form("forBairro"),"'","")))
forCidade		=	UCase(Trim(Replace(Request.Form("forCidade"),"'","")))
forUf			=	UCase(Trim(Replace(Request.Form("forUf"),"'","")))
forCep			=	UCase(Trim(Replace(Request.Form("forCep"),"'","")))
forCnpj			=	UCase(Trim(Replace(Request.Form("forCnpj"),"'","")))
forIE			=	UCase(Trim(Replace(Request.Form("forIE"),"'","")))
forEmail		=	LCase(Trim(Replace(Request.Form("forEmail"),"'","")))
forTelefone		=	UCase(Trim(Replace(Request.Form("forTelefone"),"'","")))
forContato		=	UCase(Trim(Replace(Request.Form("forContato"),"'","")))

%>

<%
if (not isEmpty(Request.Form("cadastrar"))) then

	Dim rs00
	Dim sql00
	set rs00 = server.CreateObject("adodb.recordset")
		sql00 = "SELECT * FROM tb_fornecedor WHERE forCnpj = '"&forCnpj&"' OR forNome = '"&forNome&"'"
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
			sql01 = "INSERT INTO tb_fornecedor (usuLogin, forNome,forEndereco,forBairro,forCidade, forUf, forCep,forCnpj,forIE,forEmail,forTelefone,forContato) VALUES ('"&usuLogin&"','"&forNome&"','"&forEndereco&"','"&forBairro&"','"&forCidade&"','"&forUf&"','"&forCep&"','"&forCnpj&"','"&forIE&"','"&forEmail&"','"&forTelefone&"','"&forContato&"')"
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
	
var forNome
var forCnpj
var forIE

forNome = document.form1.forNome.value;
forCnpj = document.form1.forCnpj.value;
forIE = document.form1.forIE.value;

if (forNome.length < 3) {
	alert("Favor informar o nome do fornecedor");
	document.form1.forNome.focus();
	return false;
	}
	
if (forCnpj.length < 11) {
	alert("Favor informar o CNPJ/CPF do fornecedor corretamente!");
	document.form1.forCnpj.focus();
	return false;
	}	

if (forIE.length < 6) {
	alert("Favor informar a IE/RG do fornecedor!");
	document.form1.forIE.value = "ISENTO";
	document.form1.forIE.focus();
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
	    <li><a href="pg_insert_fornecedor.asp">Novo Fornecedor</a></li>          
	    <li><a href="pg_select_fornecedor.asp">Listar Fornecedores</a></li>         
	    <li><a href="pg_insert_produto.asp">Produto</a></li>               
	  </ul>
	</div>
	<div id="areaPrincipal">
    <div style="height:25px; line-height:25px; background:#ccc"></div>
    <table width="96%" border="0" align="left" cellpadding="3" cellspacing="3">
  <tr>
    <td width="50" align="center"><img src="ico/ico_pessoajuridica.gif" width="60" height="60" border="0" class="icone" /></td>
    <td width="751" align="center" class="titulo">CADASTRO DE FORNECEDOR</td>
    </tr>
  <tr>
    <td colspan="2" align="center"><form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>" onsubmit="return verForm(this)">
      <table width="777" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td height="25" colspan="4" align="center"><%=msg%></td>
        </tr>
        <tr>
          <td width="109" height="25" align="right">Fornecedor:</td>
          <td width="267" height="25" align="left"><label>
            <input name="forNome" type="text" id="forNome" size="40" maxlength="30" />
          *</label></td>
          <td width="110" height="25" align="right">CNPJ/CPF:</td>
          <td width="263" height="25" align="left"><input name="forCnpj" type="text" id="forCnpj" size="20" maxlength="20" />
            *</td>
        </tr>
        <tr>
          <td height="25" align="right">Endereço:</td>
          <td height="25" align="left"><input name="forEndereco" type="text" id="forEndereco" size="40" maxlength="50" /></td>
          <td height="25" align="right">IE/RG:</td>
          <td height="25" align="left"><input name="forIE" type="text" id="forIE" value="Isento" size="20" maxlength="20" />
            *</td>
        </tr>
        <tr>
          <td height="25" align="right">Bairro:</td>
          <td height="25" align="left"><input name="forBairro" type="text" id="forBairro" size="40" maxlength="50" /></td>
          <td height="25" align="right">E-mail:</td>
          <td height="25" align="left"><input name="forEmail" type="text" id="forEmail" size="40" maxlength="50" style="text-transform:lowercase"/></td>
        </tr>
        <tr>
          <td height="25" align="right">Cidade:</td>
          <td height="25" align="left"><input name="forCidade" type="text" id="forCidade" size="30" maxlength="50" />
            Uf:
            <input name="forUf" type="text" id="forUf" value="SP" size="4" maxlength="4" /></td>
          <td height="25" align="right">Telefone:</td>
          <td height="25" align="left"><input name="forTelefone" type="text" id="forTelefone" size="10" maxlength="12" /></td>
        </tr>
        <tr>
          <td height="25" align="right"> Cep:</td>
          <td height="25" align="left"><input name="forCep" type="text" id="forCep" size="10" maxlength="10" /></td>
          <td height="25" align="right">Contato:</td>
          <td height="25" align="left"><input name="forContato" type="text" id="forContato" size="40" maxlength="50" /></td>
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
