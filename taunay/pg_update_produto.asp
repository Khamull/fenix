<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%option explicit%>

<!--#include file="inc/inc_conexao.inc"-->

<!--#include file="inc/inc_formato_data.inc"-->

<!--#include file="inc/inc_acesso.inc" -->

<%
call abreConexao()
%>

<%
msg = Request.QueryString("msg")

if (msg = "1") Then
 msg = "Os Dados foram Atualizados com Sucesso!"
end if
%>


<%
Dim msg
Dim usuLogin
Dim proDescricao
Dim proDescricao2
Dim tipID
Dim proUnidade
Dim proPrecoCusto
Dim proPrecoVenda
Dim proPrecoVendaM
Dim proCodEmpresa
Dim forID
Dim proCodFornecedor
Dim proMargemLucro
Dim proMargemLucroM
Dim proEstoqueMin
Dim proID
Dim proAtivo

proID = Request.QueryString("proID")
if(proID = "") then
response.redirect("pg_menu.asp")
end if

Dim rs00
Dim sql00
set rs00 = server.CreateObject("adodb.recordset")
	sql00 = "SELECT *, tb_fornecedor.forID, tb_fornecedor.forNome, tb_tipo.tipID, tb_tipo.tipDescr from tb_produto INNER JOIN tb_fornecedor ON tb_produto.forID = tb_fornecedor.forID INNER JOIN tb_tipo ON tb_tipo.tipID = tb_produto.tipID WHERE proID = '"&proID&"'"
set rs00 = conn.execute(sql00)	

usuLogin 				= 	UCase(Trim(Replace(Request.Form("usuLogin"),"'","")))
proDescricao			= 	UCase(Trim(Replace(Request.Form("proDescricao"),"'","")))
proDescricao2			= 	UCase(Trim(Replace(Request.Form("proDescricao2"),"'","")))
tipID					= 	UCase(Trim(Replace(Request.Form("tipID"),"'","")))
proUnidade				= 	UCase(Trim(Replace(Request.Form("proUnidade"),"'","")))
proPrecoCusto			= 	UCase(Trim(Replace(Request.Form("proPrecoCusto"),",",".")))
proPrecoVenda			= 	UCase(Trim(Replace(Request.Form("proPrecoVenda"),",",".")))
proPrecoVendaM			= 	UCase(Trim(Replace(Request.Form("proPrecoVendaM"),",",".")))
proCodEmpresa			= 	UCase(Trim(Replace(Request.Form("proCodEmpresa"),"'","")))
forID					= 	UCase(Trim(Replace(Request.Form("forID"),"'","")))
proCodFornecedor		= 	UCase(Trim(Replace(Request.Form("proCodFornecedor"),"'","")))
proMargemLucro			= 	UCase(Trim(Replace(Request.Form("proMargemLucro"),",",".")))
proMargemLucroM			= 	UCase(Trim(Replace(Request.Form("proMargemLucroM"),",",".")))
proEstoqueMin			= 	UCase(Trim(Replace(Request.Form("proEstoqueMin"),"'","")))
proAtivo				=	Request.Form("proAtivo")
%>


<%
if (not isEmpty(Request.Form("atualizar"))) then

if (not isNumeric(proMargemLucro)) then
proMargemLucro = 0.00
end if

if (not isNumeric(proMargemLucroM)) then
proMargemLucroM = 0.00
end if

if (not isNumeric(proPrecoCusto)) then
proPrecoCusto = 0.00
end if

if (not isNumeric(proPrecoVenda)) then
proPrecoVenda = 0.00
msg = "ERRO! O produto não pode ser cadastrado, preço de venda inválido!"

elseif (not isNumeric(proPrecoVendaM)) then
proPrecoVendaM = 0.00
msg = "ERRO! O produto não pode ser cadastrado, preço de venda inválido!"

end if

		Dim rs01
		Dim sql01
		set rs01 = server.CreateObject("adodb.recordset")
			sql01 = "UPDATE tb_produto SET"
			
			sql01=sql01&" usuLogin='"&usuLogin&"',"
			sql01=sql01&" proDescricao='"&proDescricao&"',"
			sql01=sql01&" proDescricao2='"&proDescricao2&"',"
			sql01=sql01&" tipID='"&tipID&"',"
			sql01=sql01&" proUnidade='"&proUnidade&"',"
			sql01=sql01&" proPrecoCusto='"&proPrecoCusto&"',"
			sql01=sql01&" proPrecoVenda='"&proPrecoVenda&"',"
			sql01=sql01&" proPrecoVendaM='"&proPrecoVendaM&"',"
			sql01=sql01&" proCodEmpresa='"&proCodEmpresa&"',"
			sql01=sql01&" forID='"&forID&"',"
			sql01=sql01&" proAtivo='"&proAtivo&"',"			
			sql01=sql01&" proCodFornecedor='"&proCodFornecedor&"',"
			sql01=sql01&" proMargemLucro='"&proMargemLucro&"',"
			sql01=sql01&" proMargemLucroM='"&proMargemLucroM&"',"
			sql01=sql01&" proEstoqueMin='"&proEstoqueMin&"'"
			sql01=sql01&" WHERE proID = '"&proID&"'"

		set rs01 = conn.execute(sql01)
		set rs01 = nothing
		
		'****************************

		Response.redirect("pg_update_produto.asp?msg=1&proID="&proID)
	
end if
%>


<%
Dim rs02
Dim sql02
set rs02 = server.CreateObject("adodb.recordset")
	sql02 = "SELECT * FROM tb_tipo where tipAtivo = 'S' ORDER BY tipDescr"
set rs02 = conn.execute(sql02)	
%>


<%
Dim rs03
Dim sql03
set rs03 = server.CreateObject("adodb.recordset")
	sql03 = "SELECT * FROM tb_fornecedor where forAtivo = 'S' ORDER BY forNome"
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

var proDescricao		= document.form1.proDescricao.value;

if (proDescricao.length < 3)
{
	alert("Informe o nome do produto!");
	document.form1.proDescricao.focus();
	return false;
	}

var proCodEmpresa				= document.form1.proCodEmpresa.value;

if (proCodEmpresa == "")
{
	alert("Informe o código do produto!");
	document.form1.proCodEmpresa.focus();
	return false;
}

var tipID	= document.form1.tipID.value;

if(tipID == ""){
	alert("Favor selecionar o tipo de produto!");
	document.form1.tipID.focus();
	return false;
	}

var forID = document.form1.forID.value;

if(forID == "")
{
	alert("Favor selecionar o fornecedor!");
	document.form1.forID.focus();
	return false;
	}

var proUnidade = document.form1.proUnidade.value;

if(proUnidade == "")
{	alert("Favor selecionar a unidade do produto!");
	document.form1.proUnidade.focus();
	return false;
	}

var proPrecoVenda = document.form1.proPrecoVenda.value;

if((proPrecoVenda == "0.00")||(proPrecoVenda == ""))
{
	alert("Favor informar um preço de venda para o produto!");
	document.form1.proPrecoVenda.focus();
	return false;
	}

var proEstoqueMin = document.form1.proEstoqueMin.value;

if(proEstoqueMin == "")
{	alert("Favor informar um estoque Mínimo!");
	document.form1.proEstoqueMin.focus();
	return false;
	}
	
var p = document.form1.proMargemLucro.value;
var v = document.form1.proPrecoVenda.value;
var c = document.form1.proPrecoCusto.value;
var e = document.form1.proEstoqueMin.value;	
	
	
	if (p.indexOf(".") == -1)
	{
		alert("Favor informar as casas decimais!\nPara a Margem de Lucro em % Exemplo: " + p +".00");
		document.form1.proMargemLucro.focus();
		return false;
	}
	
	if (v.indexOf(".") == -1)
	{
		alert("Favor informar as casas decimais!\nPara o Preço de Venda Exemplo: " + v +".00");
		document.form1.proPrecoVenda.focus();
		return false;
	}
		
	if (c.indexOf(".") == -1)
	{
		alert("Favor informar as casas decimais!\nPara o Preço de Custo Exemplo: " + c +".00");
		document.form1.proPrecoCusto.focus();
		return false;
	}

	if (e.indexOf(".") == -1)
	{
		alert("Favor informar as casas decimais!\nPara a quantidade de produto: " + e +".00");
		document.form1.proEstoqueMin.focus();
		return false;
	}
	
}
function verCodigo(){
	var proCodEmpresa = document.form1.proCodEmpresa.value;
	document.form1.proCodFornecedor.value = proCodEmpresa;
}
	
function verMargem(){
var p = document.form1.proMargemLucro.value;
var c = document.form1.proPrecoCusto.value;
p = p.replace(",",".");
c = c.replace(",",".");
var v = parseFloat(p*c)/100 + parseFloat(c);
v = parseFloat(v);
document.form1.proPrecoVenda.value = v;

}

function verMargemB(){
var c = document.form1.proPrecoCusto.value;
var i = document.form1.proPrecoVenda.value;

c = c.replace(",",".");
i = i.replace(",",".");

var v = parseFloat(i*100)/parseFloat(c) - 100;
v = parseFloat(v);
document.form1.proMargemLucro.value = v.toFixed(2);
}

function verMargemM(){
var c = document.form1.proPrecoCusto.value;
var i = document.form1.proPrecoVendaM.value;

c = c.replace(",",".");
i = i.replace(",",".");

var v = parseFloat(i*100)/parseFloat(c) - 100;
v = parseFloat(v);
document.form1.proMargemLucroM.value = v.toFixed(2);
}

function verPonto()
{
var proPrecoCusto = document.form1.proPrecoCusto.value;
var proPrecoVenda = document.form1.proPrecoVenda.value;
var proMargemLucro = document.form1.proMargemLucro.value;
var proEstoqueMin = document.form1.proEstoqueMin.value;

proPrecoCusto = proPrecoCusto.replace(",",".");
proPrecoVenda = proPrecoVenda.replace(",",".");
proEstoqueMin = proEstoqueMin.replace(",",".");

document.form1.proPrecoCusto.value = proPrecoCusto;
document.form1.proPrecoVenda.value = proPrecoVenda;
document.form1.proEstoqueMin.value = proEstoqueMin;

}	

function vendaBalcao(){
	var p1 = document.form1.proPrecoCusto.value;
	
	document.form1.proPrecoVendaM.value = p1;
	document.form1.proPrecoVenda.value = p1;
	
	//Coloca duas casas decimais com o toFixed(2)
	document.form1.proPrecoVendaM.value.toFixed(2);
	document.form1.proPrecoVenda.value.toFixed(2);
}

</script>

</head>
<body onload="verPonto()">
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
	    <li><a href="pg_insert_produto.asp">Novo Produto</a></li>    
	    <li><a href="pg_select_produto.asp">Listar Produtos</a></li>             
	    <li><a href="pg_insert_fornecedor.asp">Fornecedor</a></li>
	    <li><a href="pg_insert_tipoProduto.asp">Tipo de Produto</a></li>                
	  </ul>
	</div>
	<div id="areaPrincipal">
    <div style="height:25px; line-height:25px; background:#ccc"></div>
    <table width="777" border="0" align="left" cellpadding="3" cellspacing="3">
  <tr>
    <td width="100" align="center"><img src="ico/ico_pizza.gif" width="60" height="60" class="icone" /></td>
    <td width="1342" align="center" class="titulo">ATUALIZAR PRODUTO</td>
    </tr>
  <tr>
    <td colspan="2" align="center"><form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?proID=<%=proID%>" onsubmit="return verForm(this)">
      <table width="777" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td height="25" colspan="4" align="center"><font color="#FF0000"><b><%=msg%></b></font></td>
        </tr>
        <tr>
          <td width="137" height="25" align="right">Preço de Custo:</td>
          <td width="253" height="25" align="left"><input name="proPrecoCusto" type="text" id="proPrecoCusto" value="<%=Replace(FormatNumber(rs00.fields.item("proPrecoCusto").value),",",".")%>" size="10" maxlength="10" onkeypress="vendaBalcao()"  onblur="verPonto(); vendaBalcao()"/>
*</td>
          <td width="107" height="25" align="right">Cód. Produto:</td>
          <td width="267" height="25" align="left"><input name="proCodEmpresa" type="text" id="proCodEmpresa" onblur="verCodigo()" value="<%=rs00.fields.item("proCodEmpresa").value%>" size="20" maxlength="20"/></td>
        </tr>
        <tr>
          <td height="25" align="right">tipo de Produto:</td>
          <td height="25" align="left"><select name="tipID" id="tipID">
            <option value="<%=rs00.fields.item("tipID").value%>"><%=rs00.fields.item("tipDescr").value%></option>
            <%
		  if (not rs02.eof) then
          do while not rs02.eof
		  %>
            <option value="<%=rs02.fields.item("tipID").value%>"><%=rs02.fields.item("tipDescr").value%></option>
            <%
		  rs02.moveNext 
		  Loop
		  end if
		  %>
          </select>
*</td>
          <td height="25" align="right">Fornecedor:</td>
          <td height="25" align="left"><select name="forID" id="forID">
            <option value="<%=rs00.fields.item("forID").value%>"><%=rs00.fields.item("forNome").value%></option>
            <%
		  if (not rs03.eof) then
          do while not rs03.eof
		  %>
            <option value="<%=rs03.fields.item("forID").value%>"><%=rs03.fields.item("forNome").value%></option>
            <%
		  rs03.moveNext 
		  Loop
		  end if
		  %>
          </select>
*</td>
        </tr>
        <tr>
          <td height="25" align="right">Unidade:</td>
          <td height="25" align="left">
<select name="proUnidade" id="proUnidade">
            <option value="<%=rs00.fields.item("proUnidade").value%>" selected><%=rs00.fields.item("proUnidade").value%></option>
            <option value="UN">UN</option>
            <option value="KG">KG</option>
            <option value="PÇ">PÇ</option>
            <option value="LT">LT</option>
            <option value="ML">ML</option>
            <option value="CX">CX</option>
</select>
*</td>
          <td height="25" align="right">Cód. Fornecedor:</td>
          <td height="25" align="left"><input name="proCodFornecedor" type="text" id="proCodFornecedor" value="<%=Replace(rs00.fields.item("proCodFornecedor").value,",",".")%>" size="20" maxlength="20" /></td>
        </tr>
        <tr>
          <td height="25" align="right" valign="top">Título:</td>
          <td height="25" align="left"><input name="proDescricao" type="text" id="proDescricao" value="<%=rs00.fields.item("proDescricao").value%>" size="40"/></td>
          <td height="25" align="right" valign="top">&nbsp;</td>
          <td height="25" align="left" valign="top">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" align="right" valign="top">Descri&ccedil;&atilde;o:</td>
          <td height="25" align="left"><textarea name="proDescricao2" id="proDescricao2" cols="10" rows="5" style="width:300px;height:25;"><%=rs00.fields.item("proDescricao2").value%></textarea>
            </td>
          <td height="25" align="right" valign="top">Estoque Mínimo:</td>
          <td height="25" align="left" valign="top"><input name="proEstoqueMin" type="text" id="proEstoqueMin" value="<%=Replace(FormatNumber(rs00.fields.item("proEstoqueMin").value),",",".")%>" size="10" maxlength="10" onblur="verPonto()"/></td>
        </tr>
        <tr>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="left">&nbsp;</td>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="left">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" align="right" bgcolor="#99CC99">Preço Telefone/Balcão:</td>
          <td height="25" bgcolor="#99CC99"><input name="proPrecoVenda" type="text" id="proPrecoVenda" value="<%=Replace(FormatNumber(rs00.fields.item("proPrecoVenda").value),",",".")%>" size="10" maxlength="10" onkeypress="verMargemB()" onblur="verPonto(); verMargemB()"/></td>
          <td height="25" align="right" bgcolor="#66FF99">Preço Mesa:</td>
          <td height="25" bgcolor="#66FF99"><input type="text" name="proPrecoVendaM" id="proPrecoVendaM" value="<%=Replace(FormatNumber(rs00.fields.item("proPrecoVendaM").value),",",".")%>" size="10" maxlength="10" onkeypress="verMargemM()" onblur="verPonto(); verMargemM()"/></td>
        </tr>
        <tr>
          <td height="25" align="right" bgcolor="#99CC99">Lucro:</td>
          <td height="25" bgcolor="#99CC99"><input name="proMargemLucro" type="text" id="proMargemLucro"  onblur="verPonto()" onkeypress="verMargem()" value="<%=Replace(FormatNumber(rs00.fields.item("proMargemLucro").value),",",".")%>" size="6" maxlength="6" readonly="readonly"/>
%</td>
          <td height="25" align="right" bgcolor="#66FF99">Lucro:</td>
          <td height="25" bgcolor="#66FF99"><input name="proMargemLucroM" type="text" id="proMargemLucroM" onblur="verPonto()" value="<%=Replace(FormatNumber(rs00.fields.item("proMargemLucroM").value),",",".")%>" size="6" maxlength="6" readonly="readonly"/>
%</td>
        </tr>
        <tr>
          <td height="25" align="right">&nbsp;</td>
          <td height="25">&nbsp;</td>
          <td height="25" align="right">&nbsp;</td>
          <td height="25">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" align="right">Ativo:</td>
          <td height="25"><select name="proAtivo" id="proAtivo">
            <option value="S" selected="selected">SIM</option>
            <option value="N">NÃO</option>
          </select></td>
          <td height="25" align="right">&nbsp;</td>
          <td height="25">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" align="right">&nbsp;</td>
          <td height="25"><input name="atualizar" type="submit" class="botao" id="atualizar" value="Atualizar" onsubmit="return verForm()"/></td>
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
