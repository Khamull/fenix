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
Dim cliNome
Dim cliEndereco
Dim cidID
Dim baiID
Dim cliEndRef
Dim cliCpf
Dim cliTelefone
Dim cliTelefone2
Dim cliTelefone3
Dim cliTelefone4
Dim cliTelefone5
Dim cliEmail
Dim cliSenha
Dim cliPreferencias
Dim cliCod
Dim cliID
Dim cliAtivo

cliID =	Request.QueryString("cliID")

if(cliID = "") Then
Response.Redirect("pg_menu.asp")
end if


usuLogin 		=	Request.Form("usuLogin")
cliNome 		=	UCase(Trim(Replace(Request.Form("cliNome"),"'","")))
cliEndereco		=	UCase(Trim(Replace(Request.Form("cliEndereco"),"'","")))
cidID			=	Trim(Replace(Request.Form("cidID"),"'",""))
baiID			=	Trim(Replace(Request.Form("baiID"),"'",""))
cliEndRef		=	UCase(Trim(Replace(Request.Form("cliEndRef"),"'","")))
cliCpf			=	UCase(Trim(Replace(Request.Form("cliCpf"),"'","")))
cliTelefone		=	UCase(Trim(Replace(Request.Form("cliTelefone"),"'","")))
cliTelefone2	=	UCase(Trim(Replace(Request.Form("cliTelefone2"),"'","")))
cliTelefone3	=	UCase(Trim(Replace(Request.Form("cliTelefone3"),"'","")))
cliTelefone4	=	UCase(Trim(Replace(Request.Form("cliTelefone4"),"'","")))
cliTelefone5	=	UCase(Trim(Replace(Request.Form("cliTelefone5"),"'","")))
cliEmail		=	LCase(Trim(Replace(Request.Form("cliEmail"),"'","")))
cliSenha		=	UCase(Trim(Replace(Request.Form("cliSenha"),"'","")))
cliPreferencias = 	Request.Form("cliPreferencias")
cliAtivo 		= 	Request.Form("cliAtivo")
%>


<%
Dim rs00
Dim sql00
set rs00 = server.CreateObject("adodb.recordset")
	sql00 = "SELECT *, tb_bairro.baiID, tb_bairro.baiNome, tb_cidade.cidID, tb_cidade.cidNome FROM tb_cliente INNER JOIN tb_bairro ON tb_cliente.baiID = tb_bairro.baiID INNER JOIN tb_cidade ON tb_cidade.cidID = tb_bairro.cidID where tb_cliente.cliID = '"&cliID&"'"
set rs00 = conn.execute(sql00)	
%>


<%
if (not isEmpty(Request.Form("atualizar"))) then

Dim erros

erros = 0
	
'------------ VERIFICA SE JÁ TEM O TELEFONE CADASTRADO ---------------------------


	Dim rs00x
	Dim sql00x
	set rs00x = server.CreateObject("adodb.recordset")
		sql00x = "SELECT * FROM tb_cliente WHERE cliTelefone = '"&cliTelefone&"' AND cliID <> '"&cliID&"' OR cliTelefone2 = '"&cliTelefone&"' AND cliID <> '"&cliID&"' OR cliTelefone3 = '"&cliTelefone&"' AND cliID <> '"&cliID&"' OR cliTelefone4 = '"&cliTelefone&"' AND cliID <> '"&cliID&"' OR cliTelefone5 = '"&cliTelefone&"' AND cliID <> '"&cliID&"' OR cliCpf = '"&cliCpf&"' AND cliID <> '"&cliID&"'"
	set rs00x = conn.Execute(sql00x)
	
		if (not rs00x.eof) then
		
		'Mensagem
		rs00x.close
		set rs00x = nothing
		msg = "ERRO: Telefone 1 ou CPF já Cadastrado!"
		erros = 1
		end if
	
	'-------------------- 2 --------------------------------
	
	if(cliTelefone2 <> "")Then
	
	Dim rs001
	Dim sql001
	set rs001 = server.CreateObject("adodb.recordset")
		sql001 = "SELECT * FROM tb_cliente WHERE cliTelefone = '"&cliTelefone2&"' AND cliID <> '"&cliID&"' OR cliTelefone2 = '"&cliTelefone2&"' AND cliID <> '"&cliID&"' OR cliTelefone3 = '"&cliTelefone2&"' AND cliID <> '"&cliID&"' OR cliTelefone4 = '"&cliTelefone2&"' AND cliID <> '"&cliID&"' OR cliTelefone5 = '"&cliTelefone2&"' AND cliID <> '"&cliID&"' OR cliCpf = '"&cliCpf&"' AND cliID <> '"&cliID&"'"
	set rs001 = conn.Execute(sql001)
	
		if(not rs001.EoF) then
		
		'Mensagem2
		rs001.close
		set rs001 = nothing
		msg = "ERRO: Telefone 2 ou CPF já Cadastrado!"
		erros = 1
		end if
	
	'--------------------- 3 -------------------------------
	
	elseif(cliTelefone3 <> "")Then
	
	Dim rs002
	Dim sql002
	set rs002 = server.CreateObject("adodb.recordset")
		sql002 = "SELECT * FROM tb_cliente WHERE cliTelefone = '"&cliTelefone3&"' AND cliID <> '"&cliID&"' OR cliTelefone2 = '"&cliTelefone3&"' AND cliID <> '"&cliID&"' OR cliTelefone3 = '"&cliTelefone3&"' AND cliID <> '"&cliID&"' OR cliTelefone4 = '"&cliTelefone3&"' AND cliID <> '"&cliID&"' OR cliTelefone5 = '"&cliTelefone3&"' AND cliID <> '"&cliID&"' OR cliCpf = '"&cliCpf&"' AND cliID <> '"&cliID&"'"
	set rs002 = conn.Execute(sql002)
	
		if(not rs002.EoF) then	
		 
		'Mensagem3
		rs002.close
		set rs002 = nothing
		msg = "ERRO: Telefone 3 ou CPF já Cadastrado!"
		erros = 1
		end if
	
	'-------------------- 4 --------------------------------
	
	elseif(cliTelefone4 <> "")Then
	
	Dim rs003
	Dim sql003
	set rs003 = server.CreateObject("adodb.recordset")
		sql003 = "SELECT * FROM tb_cliente WHERE cliTelefone = '"&cliTelefone4&"' AND cliID <> '"&cliID&"' OR cliTelefone2 = '"&cliTelefone4&"' AND cliID <> '"&cliID&"' OR cliTelefone3 = '"&cliTelefone4&"' AND cliID <> '"&cliID&"' OR cliTelefone4 = '"&cliTelefone4&"' AND cliID <> '"&cliID&"' OR cliTelefone5 = '"&cliTelefone4&"' AND cliID <> '"&cliID&"' OR cliCpf = '"&cliCpf&"' AND cliID <> '"&cliID&"'"
	set rs003 = conn.Execute(sql003)
	
		if(not rs003.EoF) then	
		 
		'Mensagem4
		rs003.close
		set rs003 = nothing
		msg = "ERRO: Telefone 4 ou CPF já Cadastrado!"
		erros = 1
		end if
	
	'-------------------- 5 ---------------------------------
	
	elseif(cliTelefone5 <> "")Then
	
	Dim rs004
	Dim sql004
	set rs004 = server.CreateObject("adodb.recordset")
		sql004 = "SELECT * FROM tb_cliente WHERE cliTelefone = '"&cliTelefone5&"' AND cliID <> '"&cliID&"' OR cliTelefone2 = '"&cliTelefone5&"' AND cliID <> '"&cliID&"' OR cliTelefone3 = '"&cliTelefone5&"' AND cliID <> '"&cliID&"' OR cliTelefone4 = '"&cliTelefone5&"' AND cliID <> '"&cliID&"' OR cliTelefone5 = '"&cliTelefone5&"' AND cliID <> '"&cliID&"' OR cliCpf = '"&cliCpf&"' AND cliID <> '"&cliID&"'"
	set rs004 = conn.Execute(sql004)
	
		if(not rs004.EoF) then
		 
		'Mensagem5
		rs004.close
		set rs004 = nothing
		msg = "ERRO: Telefone 5 ou CPF já Cadastrado!"
		erros = 1
		end if
	
     end if
	
'----------------------------- FIM DA VERIFICAÇÃO ---------------------------------------'
	

if(erros = 0)Then
	
		Dim rs01
		Dim sql01
		set rs01 = server.CreateObject("adodb.recordset")
			sql01 = "UPDATE tb_cliente SET "
			sql01=sql01&" usuLogin='"&usuLogin&"',"
			sql01=sql01&" cliNome='"&cliNome&"',"
			sql01=sql01&" cliEndereco='"&cliEndereco&"',"
			sql01=sql01&" baiID='"&baiID&"',"
			sql01=sql01&" cliEndRef='"&cliEndRef&"',"
			sql01=sql01&" cliCpf='"&cliCpf&"',"
			sql01=sql01&" cliTelefone='"&cliTelefone&"',"
			sql01=sql01&" cliTelefone2='"&cliTelefone2&"',"
			sql01=sql01&" cliTelefone3='"&cliTelefone3&"',"
			sql01=sql01&" cliTelefone4='"&cliTelefone4&"',"
			sql01=sql01&" cliTelefone5='"&cliTelefone5&"',"
			sql01=sql01&" cliEmail='"&cliEmail&"',"
			sql01=sql01&" cliSenha='"&cliSenha&"',"
			sql01=sql01&" cliPreferencias='"&cliPreferencias&"',"
			sql01=sql01&" cliAtivo='"&cliAtivo&"'"
			sql01=sql01&" WHERE cliID='"&cliID&"'"
		set rs01 = conn.execute(sql01)
		set rs01 = nothing
			
			response.redirect("pg_menu_pedidos.asp?tipo=telefone&cliID="&cliID&"&cliTelefone="&cliTelefone)
			'response.redirect("pg_menu_pedidos.asp?tipo=telefone")
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

<%
Dim rs03
Dim sql03
set rs03 = server.CreateObject("adodb.recordset")
	sql03 = "SELECT * FROM tb_bairro where baiAtivo = 'S' AND cidID='"&Request.QueryString("cidID")&"' ORDER BY baiNome"
set rs03 = conn.execute(sql03)	
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>SISTEM FORTE EM MÍDIA</title>
<link href="css/css1.css" rel="stylesheet" type="text/css" />

<script language="javascript" type="text/javascript">

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

window.location.href =  "pg_update_cliente.asp?cliNome="+cliNome+"&cliEndereco="+cliEndereco+"&cidID="+cidID+"&cidNome="+cidNome+"&baiID="+baiID+"&cliCpf="+cliCpf+"&cliTelefone="+cliTelefone+"&cliEmail="+cliEmail+"&cliEndRef="+cliEndRef+"&cliSenha="+cliSenha+"&cliCod="+cliCod+"&cliID="+cliID

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
	    <li><a href="pg_select_cliente.asp">Listar Clientes</a></li>        
	    <li><a href="pg_insert_bairro.asp">Bairro</a></li>
	    <li><a href="pg_insert_cidade.asp">Cidade</a></li>                
	  </ul>
	</div>
	<div id="areaPrincipal">
    <div style="height:25px; line-height:25px; background:#ccc"></div>
    <table width="96%" border="0" align="left" cellpadding="3" cellspacing="3">
  <tr>
    <td width="50" align="center"><img src="ico/ico_pessoafisica.gif" width="60" height="60" class="icone" /></td>
    <td width="751" align="center" class="titulo">ATUALIZAR CLIENTE</td>
    </tr>
  <tr>
    <td colspan="2" align="center"><form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?cliID=<%=cliID%>"
    onsubmit="return verForm(this)">
      <table width="777" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td height="25" colspan="4" align="center"><font color="#FF0000"><b><%=msg%></b></font></td>
        </tr>
        <tr>
          <td width="109" height="25" align="right">Nome do Cliente:</td>
          <td width="267" height="25" align="left"><label>
            <input name="cliNome" type="text" id="cliNome" value="<%=rs00.fields.item("cliNome").value%>" size="40" maxlength="30" />
          *</label></td>
          <td width="110" height="25" align="right">Código do Cliente</td>
          <td width="263" height="25" align="left"><input name="cliCod" type="text" id="cliCod" value="<%=rs00.fields.item("cliID").value%>" size="14" maxlength="14" readonly="readonly" /></td>
        </tr>
        <tr>
          <td height="25" align="right">Endereço:</td>
          <td height="25" align="left"><input name="cliEndereco" type="text" id="cliEndereco" value="<%=rs00.fields.item("cliEndereco").value%>" size="40" maxlength="50" />
            *</td>
          <td height="25" align="right">Documento CPF:</td>
          <td height="25" align="left"><input name="cliCpf" type="text" id="cliCpf" value="<%=rs00.fields.item("cliCpf").value%>" size="14" maxlength="11" /></td>
        </tr>
        <tr>
          <td height="25" align="right">Referência:</td>
          <td height="25" align="left"><input name="cliEndRef" type="text" id="cliEndRef" value="<%=rs00.fields.item("cliEndRef").value%>" size="40" maxlength="50" /></td>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="left">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" align="right">Cidade:</td>
          <td height="25" align="left">
          <select name="cidID" id="cidID" onchange="carregaBairro()">
            <option value="<%=rs00.fields.item("cidID").value%>"><%=rs00.fields.item("cidNome").value%></option>
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
          <td height="25" align="right">Bairro:</td>
          <td height="25" align="left"><select name="baiID" id="baiID">
            <option value="<%=rs00.fields.item("baiID").value%>"><%=rs00.fields.item("baiNome").value%></option>
            <%
		  if (not rs03.eof) then
          do while not rs03.eof
		  %>
            <option value="<%=rs03.fields.item("baiID").value%>"><%=rs03.fields.item("baiNome").value%></option>
            <%
		  rs03.moveNext 
		  Loop
		  end if
		  %>
          </select>
*</td>
        </tr>
        <tr>
          <td height="25" align="right">Telefone 1:</td>
          <td height="25" align="left"><input name="cliTelefone" type="text" id="cliTelefone" value="<%=rs00.fields.item("cliTelefone").value%>" size="9" maxlength="8" />
*</td>
          <td height="25" align="right">Telefone 2:</td>
          <td height="25" align="left"><input name="cliTelefone2" type="text" id="cliTelefone2" value="<%=rs00.fields.item("cliTelefone2").value%>" size="9" maxlength="8" /></td>
        </tr>
        <tr>
          <td height="25" align="right">Telefone 3:</td>
          <td height="25" align="left"><input name="cliTelefone3" type="text" id="cliTelefone3" value="<%=rs00.fields.item("cliTelefone3").value%>" size="9" maxlength="8" /></td>
          <td height="25" align="right">Telefone 4:</td>
          <td height="25" align="left"><input name="cliTelefone4" type="text" id="cliTelefone4" value="<%=rs00.fields.item("cliTelefone4").value%>" size="9" maxlength="8" /></td>
        </tr>
        <tr>
          <td height="25" align="right">Telefone 5:</td>
          <td height="25" align="left"><input name="cliTelefone5" type="text" id="cliTelefone5" value="<%=rs00.fields.item("cliTelefone5").value%>" size="9" maxlength="8" /></td>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="left">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" align="right">E-mail:</td>
          <td height="25" align="left"><input name="cliEmail" type="text" id="cliEmail" value="<%=rs00.fields.item("cliEmail").value%>" size="40" maxlength="50" style="text-transform:lowercase" /></td>
          <td height="25" align="right">Senha:</td>
          <td height="25" align="left"><input name="cliSenha" type="text" id="cliSenha" value="<%=rs00.fields.item("cliSenha").value%>" size="8" maxlength="8" /></td>
        </tr>
        <tr>
          <td height="25" align="right">Ativo:</td>
          <td height="25"><label>
            <select name="cliAtivo" id="cliAtivo">
              <option value="S" selected="selected">SIM</option>
              <option value="N">NÃO</option>
            </select>
          </label></td>
          <td height="25" align="right">&nbsp;</td>
          <td height="25">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" align="right" valign="top">Preferências:</td>
          <td height="25" colspan="3" valign="top"><textarea name="cliPreferencias" cols="93" rows="5" id="cliPreferencias"><%=rs00.fields.item("cliPreferencias").value%></textarea></td>
          </tr>
        <tr>
          <td height="25" align="right">&nbsp;</td>
          <td height="25"><input name="atualizar" type="submit" class="botao" id="atualizar" value="Atualizar" /></td>
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
rs02.Close
set rs02 = nothing
%>

<%
rs03.Close
set rs03 = nothing
%>

<%
call FechaConexao()
%>