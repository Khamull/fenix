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
Dim funID
'VERIFICA O PARAMETRO
funID = REQUEST.QUERYSTRING("funID")
if(funID = "") THEN
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
Dim funNome
Dim funAtivo
Dim funTelefone
Dim funCpf
Dim carID
Dim carDescr

'RECUPERA OS VALORES
usuLogin 	=	Request.Form("usuLogin")
funAtivo	=	Request.Form("funAtivo")
funTelefone =	Request.Form("funTelefone")
funCpf		 =	Request.Form("funCpf")
funNome		=	UCase(Trim(Replace(Request.Form("funNome"),"'","")))
carID		= 	Request.Form("carID")

%>

<%
'VERIFICA SE O FORME FOI DISPARADO

if (not isEmpty(Request.Form("atualizar"))) then

		Dim rs01
		Dim sql01
		set rs01 = server.CreateObject("adodb.recordset")
		sql01="UPDATE tb_funcionario SET"
		sql01=sql01&" funNome='"&funNome&"',"
		sql01=sql01&" usuLogin='"&usuLogin&"'," 
		sql01=sql01&" funAtivo='"&funAtivo&"'," 
		sql01=sql01&" funTelefone='"&funTelefone&"'," 
		sql01=sql01&" funCpf='"&funCpf&"',"
		sql01=sql01&" carID='"&carID&"'" 	 		
		sql01=sql01&" WHERE" 
		sql01=sql01&" funID='"&funID&"'"
		set rs01 = conn.execute(sql01)
		
		'RETORNA UMA MENSAGEM DO RESULTADO DA OPERAÇÃO
		response.redirect("pg_select_funcionario.asp")
end if

'CARREGA CARGO

Dim rs02
Dim sql02
set rs02 = Server.CreateObject("ADODB.Recordset")
sql02 = "SELECT carID, carDescr FROM  tb_cargo WHERE carAtivo = 'S' ORDER BY carDescr"
set rs02 = conn.Execute(sql02)

%>	
<%
'SELECIONA OS REGISTROS

		Dim rs00
		Dim sql00
		set rs00 = server.CreateObject("adodb.recordset")
		sql00="SELECT "		
		sql00=sql00&" tb_funcionario.funID,"
		sql00=sql00&" tb_funcionario.funNome,"
		sql00=sql00&" tb_funcionario.funCpf,"
		sql00=sql00&" tb_funcionario.funTelefone,"
		sql00=sql00&" tb_funcionario.carID,"						
		sql00=sql00&" tb_funcionario.funAtivo,"
		sql00=sql00&" tb_cargo.carID,"
		sql00=sql00&" tb_cargo.carDescr"
		sql00=sql00&" FROM"
		sql00=sql00&" tb_funcionario" 
		sql00=sql00&" INNER JOIN"
		sql00=sql00&" tb_cargo"
		sql00=sql00&" ON tb_cargo.carID = tb_funcionario.carID"
		sql00=sql00&" WHERE"
		sql00=sql00&" tb_funcionario.funID='"&funID&"'"
		set rs00=conn.Execute(sql00)
		
		carID = rs00.fields.item("carID").value
		carDescr = rs00.fields.item("carDescr").value
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
	    <li><a href="pg_select_funcionario.asp">Lista Funcionários</a></li>
      	<li><a href="pg_insert_funcionario.asp">Novo Funcionario</a></li>           
	    <li><a href="pg_insert_cargo.asp">Cargo</a></li>
	    <li><a href="pg_insert_cliente.asp">Cliente</a></li>                
	  </ul>
	</div>
	<div id="areaPrincipal">
    <div style="height:25px; line-height:25px; background:#ccc"></div>
    <table width="96%" border="0" align="left" cellpadding="3" cellspacing="3">
  <tr>
    <td width="50" align="center"><img src="ico/ico_casabairro.gif" width="60" height="60" class="icone" /></td>
    <td width="751" align="center" class="titulo">ATUALIZAR FUNCIONÁRIO</td>
    </tr>
  <tr>
    <td colspan="2" align="center"><form id="form1" name="form1" method="post" action="pg_update_funcionario.asp?funID=<%=funID%>" onsubmit="return verForm(this)">
      <table width="777" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td height="25" colspan="4" align="center"><%=msg%></td>
        </tr>
        <tr>
          <td width="109" height="25" align="right">Funcionario:</td>
          <td width="267" height="25" align="left"><input name="funNome" type="text" id="funNome" value="<%=rs00.fields.item("funNome").value%>" size="40" maxlength="30" /></td>
          <td width="110" height="25" align="right">CPF:</td>
          <td width="263" height="25" align="left"><label>
            <input name="funCpf" type="text" id="funCpf" value="<%=rs00.fields.item("funCpf").value%>" />
          </label></td>
        </tr>
        <tr>
          <td height="25" align="right">Cargo:&nbsp;</td>
          <td height="25" align="left">
          
          <select name="carID" id="carID">

		  <option value="<%=carID%>"><%=carDescr%></option>
<%
if not rs02.eof then
on error resume next
rs02.moveFirst
do while not rs02.eof
%>		 

 <option value="<%=rs02.fields.item("carID").value%>"><%=rs02.fields.item("carDescr").value%></option>
 
<%
rs02.moveNext
loop
end if
%>  
      
          </select>

			</td>
          <td height="25" align="right">Telefone:</td>
          <td height="25" align="left"><input name="funTelefone" type="text" id="funTelefone" value="<%=rs00.fields.item("funTelefone").value%>" /></td>
        </tr>
        <tr>
          <td height="25" align="right">Ativo:</td>
          <td height="25" align="left"><select name="funAtivo" id="funAtivo">
            <option value="S" selected="selected">SIM</option>
            <option value="N">NÃO</option>
          </select></td>
          <td height="25" align="right">&nbsp;</td>
          <td height="25">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="left"><input name="atualizar" type="submit" class="botao" id="atualizar" value="Atualizar" />
            <input name="usuLogin" type="hidden" id="usuLogin" value="<%=Session("usuLogin")%>" /></td>
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