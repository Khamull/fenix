<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%option explicit%>

<!--#include file="inc/inc_conexao.inc"-->

<!--#include file="inc/inc_formato_data.inc"-->

<!--#include file="inc/inc_acesso.inc" -->

<%
call abreConexao()
%>

<%
'DECLARA A VARIÁVEL DE CONSULTA
Dim forID
'VERIFICA O PARAMETRO
forID = REQUEST.QUERYSTRING("forID")
if(forID = "") THEN
'REDIRECIONADA PARA O MENU PRINCIPAL
RESPONSE.REDIRECT("pg_menu.asp")
END IF

'VARIÁVEL PADRÃO
Dim msg
Dim usuLogin

'VERIFICA A MENSAGEM
msg = REQUEST.QUERYSTRING("msg")
IF(msg = "1") THEN
msg = "Os Dados foram atualizados com sucesso!"
ELSE
msg = ""
END IF

'-**********

'DECLARA AS VARIÁVEIS
Dim forNome
Dim forAtivo
Dim forCnpj
Dim forIE
Dim forEndereco
Dim forBairro
Dim forEmail
Dim forCidade
Dim forTelefone
Dim forUf
Dim forContato

'RECUPERA OS VALORES
usuLogin 		=	Request.Form("usuLogin")
forAtivo		=	Request.Form("forAtivo")
forNome			=	UCase(Trim(Replace(Request.Form("forNome"),"'","")))
forCnpj			= 	Request.Form("forCnpj")
forIE			= 	Request.Form("forIE")
forEndereco		= 	Request.Form("forEndereco")
forBairro		= 	Request.Form("forbairro")
forEmail		= 	Request.Form("forEmail")
forCidade		= 	Request.Form("forCidade")
forTelefone		= 	Request.Form("forTelefone")
forUf			= 	Request.Form("forUf")
forContato		= 	Request.Form("forContato")

'SELECIONA OS REGISTROS

		Dim rs00
		Dim sql00
		set rs00 = server.CreateObject("adodb.recordset")
		sql00="SELECT "		
		sql00=sql00&" *"
		sql00=sql00&" FROM"
		sql00=sql00&" tb_fornecedor" 
		sql00=sql00&" WHERE"
		sql00=sql00&" tb_fornecedor.forID='"&forID&"'"
		set rs00=conn.Execute(sql00)
%>

<%
'VERIFICA SE O FORME FOI DISPARADO
if (not isEmpty(Request.Form("atualizar"))) then

		Dim rs01
		Dim sql01
		set rs01 = server.CreateObject("adodb.recordset")
		sql01="UPDATE tb_fornecedor SET"
		sql01=sql01&" forNome='"&forNome&"',"
		sql01=sql01&" usuLogin='"&usuLogin&"'," 
		sql01=sql01&" forAtivo='"&forAtivo&"'," 
		sql01=sql01&" forCnpj='"&forCnpj&"',"
		sql01=sql01&" forIE='"&forIE&"',"
		sql01=sql01&" forEndereco='"&forEndereco&"',"
		sql01=sql01&" forBairro='"&forBairro&"',"
		sql01=sql01&" forEmail='"&forEmail&"',"
		sql01=sql01&" forCidade='"&forCidade&"',"
		sql01=sql01&" forTelefone='"&forTelefone&"',"
		sql01=sql01&" forUf='"&forUf&"',"
		sql01=sql01&" forContato='"&forContato&"'"	
		sql01=sql01&" WHERE" 
		sql01=sql01&" forID='"&forID&"'"
		set rs01 = conn.execute(sql01)
		
		'RETORNA UMA MENSAGEM DO RESULTADO DA OPERAÇÃO
		response.redirect("pg_select_fornecedor.asp")
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
    <td width="751" align="center" class="titulo">ATUALIZAR  FORNECEDOR</td>
    </tr>
  <tr>
    <td colspan="2" align="center"><form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?forID=<%=forID%>" onsubmit="return verForm(this)">
      <table width="777" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td height="25" colspan="4" align="center"><%=msg%></td>
        </tr>
        <tr>
          <td width="109" height="25" align="right">Fornecedor:</td>
          <td width="267" height="25" align="left"><label>
            <input name="forNome" type="text" id="forNome" value="<%=rs00.fields.item("forNome").value%>" size="40" maxlength="30" />
          *</label></td>
          <td width="110" height="25" align="right">CNPJ/CPF:</td>
          <td width="263" height="25" align="left"><input name="forCnpj" type="text" id="forCnpj" value="<%=rs00.fields.item("forCnpj").value%>" size="20" maxlength="20" />
            *</td>
        </tr>
        <tr>
          <td height="25" align="right">Endereço:</td>
          <td height="25" align="left"><input name="forEndereco" type="text" id="forEndereco" value="<%=rs00.fields.item("forEndereco").value%>" size="40" maxlength="50" /></td>
          <td height="25" align="right">IE/RG:</td>
          <td height="25" align="left"><input name="forIE" type="text" id="forIE" value="Isento" size="20" maxlength="20" />
            *</td>
        </tr>
        <tr>
          <td height="25" align="right">Bairro:</td>
          <td height="25" align="left"><input name="forBairro" type="text" id="forBairro" value="<%=rs00.fields.item("forBairro").value%>" size="40" maxlength="50" /></td>
          <td height="25" align="right">E-mail:</td>
          <td height="25" align="left"><input name="forEmail" type="text" id="forEmail" style="text-transform:lowercase" value="<%=rs00.fields.item("forEmail").value%>" size="40" maxlength="50"/></td>
        </tr>
        <tr>
          <td height="25" align="right">Cidade:</td>
          <td height="25" align="left"><input name="forCidade" type="text" id="forCidade" value="<%=rs00.fields.item("forCidade").value%>" size="30" maxlength="50" />
            Uf:
            <input name="forUf" type="text" id="forUf" value="SP" size="4" maxlength="4" /></td>
          <td height="25" align="right">Telefone:</td>
          <td height="25" align="left"><input name="forTelefone" type="text" id="forTelefone" value="<%=rs00.fields.item("forTelefone").value%>" size="10" maxlength="12" /></td>
        </tr>
        <tr>
          <td height="25" align="right"> Cep:</td>
          <td height="25" align="left"><input name="forCep" type="text" id="forCep" value="<%=rs00.fields.item("forUf").value%>" size="10" maxlength="10" /></td>
          <td height="25" align="right">Contato:</td>
          <td height="25" align="left"><input name="forContato" type="text" id="forContato" value="<%=rs00.fields.item("forContato").value%>" size="40" maxlength="50" /></td>
        </tr>
        <tr>
          <td height="25" align="right">Ativo:</td>
          <td height="25" align="left">
          <select name="forAtivo" id="forAtivo">
            <option value="S" selected="selected">SIM</option>
            <option value="N">NÃO</option>
          </select>
          </td>
          <td height="25" align="right">&nbsp;</td>
          <td height="25">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="left"><input name="atualizar" type="submit" class="botao" id="atualizar" value="Atualizar" /></td>
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
