<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%option explicit%>

<!--#include file="inc/inc_conexao.inc"-->

<!--#include file="inc/inc_formato_data.inc"-->

<!--#include file="inc/inc_acesso.inc" -->

<%
call abreConexao()
%>

<%
Dim acao
Dim id
dim idItem
id = Request.QueryString("id")
idItem = Request.QueryString("idItem")
acao = Request.QueryString("acao")

if (id = "") then
Response.redirect("pg_insert_compra.asp")
elseif (id <> "") then

	Dim rs06
	Dim sql06
	set rs06 = server.CreateObject("adodb.recordset")
		sql06 = "SELECT tb_compra.comID, tb_fornecedor.forNome, tb_fornecedor.forID FROM tb_compra INNER JOIN tb_fornecedor ON tb_fornecedor.forID = tb_compra.forID WHERE tb_compra.comID = '"&id&"'"
	set rs06 = conn.Execute(sql06)

end if
%>

<%
Dim proID
Dim iteQtde
Dim itePreco
Dim iteSubTotal
Dim comID
Dim proMargemLucro2
Dim proMargemLucro3
Dim proPrecoVenda2
Dim proPrecoVenda3

proID	 		=	Request.Form("proID")
iteQtde			=	Request.Form("iteQtde")
itePreco		=	Request.Form("itePreco")
iteSubTotal		=	Request.Form("iteSubTotal")
comID 			= 	Request.Form("comID")
proMargemLucro2 = Request.Form("proMargemLucro")
proPrecoVenda2 = Request.Form("proPrecoVenda")
proMargemLucro3 = Request.Form("proMargemLucroM")
proPrecoVenda3 = Request.Form("proPrecoVendaM")
%>

<%
if (not isEmpty(Request.Form("lancar"))) then
	
	
	'Grava
		Dim rs01
		Dim sql01
		set rs01 = server.CreateObject("adodb.recordset")
			sql01 = "INSERT INTO tb_itemcompra (proID,iteQtde,itePreco,iteSubTotal,proPrecoVenda,proPrecoVendaM,proMargemLucro,proMargemLucroM,comID)"
			sql01 = sql01&" VALUES ('"&proID&"','"&iteQtde&"','"&itePreco&"','"&iteSubTotal&"','"&proPrecoVenda2&"','"&proPrecoVenda3&"','"&proMargemLucro2&"','"&proMargemLucro3&"','"&comID&"')"
		set rs01 = conn.execute(sql01)
		set rs01 = nothing

	Response.redirect("pg_insert_itemCompra.asp?id="&Request.QueryString("id"))
	
	end if
	
%>

<%
Dim rs02
Dim sql02
set rs02 = server.CreateObject("adodb.recordset")
	sql02 = "SELECT * FROM tb_tipo WHERE tipAtivo = 'S' AND tipID <> '1' ORDER BY tipDescr"
set rs02 = conn.execute(sql02)	
%>

<%
Dim rs03
Dim sql03
set rs03 = server.CreateObject("adodb.recordset")
	sql03 = "SELECT * FROM tb_produto where NOT EXISTS(SELECT * FROM tb_itemcompra WHERE tb_itemcompra.proID = tb_produto.proID AND tb_itemcompra.staID = 9) AND proAtivo = 'S' AND tipID = '"&Request.QueryString("tipID")&"' AND forID='"&Request.QueryString("forID")&"' ORDER BY proDescricao"
set rs03 = conn.execute(sql03)	
%>

<%
Dim rs04
Dim sql04
set rs04 = server.CreateObject("adodb.recordset")
	sql04 = "SELECT * FROM tb_produto WHERE proID = '"&Request.QueryString("proID")&"'"
set rs04 = conn.execute(sql04)	

if (not rs04.eof)  then
Dim proPrecoCusto
Dim proMargemLucro
Dim proMargemLucroM
Dim proPrecoVenda
Dim proPrecoVendaM
proPrecoCusto = rs04.fields.item("proPrecoCusto").value
proMargemLucro = rs04.fields.item("proMargemLucro").value
proMargemLucroM = rs04.fields.item("proMargemLucroM").value
proPrecoVenda = rs04.fields.item("proPrecoVenda").value
proPrecoVendaM = rs04.fields.item("proPrecoVendaM").value
else
proPrecoCusto = "0.00"
end if
%>

<%
Dim rs05
Dim sql05
set rs05 = server.CreateObject("adodb.recordset")
	sql05 = "SELECT SUM(iteSubTotal) as total FROM tb_itemcompra WHERE comID = '"&Request.QueryString("id")&"'"
set rs05 = conn.execute(sql05)	
%>

<%

if (acao = "excluirItem") then

Dim rs08
Dim sql08
set rs08 = server.createObject("adodb.recordset")
sql08 = "DELETE FROM tb_itemcompra WHERE iteID = '"&idItem&"'"
set rs08 = conn.execute(sql08)

elseif (acao = "excluirPedido") then

set rs08 = server.createObject("adodb.recordset")
sql08 = "DELETE FROM tb_compra WHERE comID = '"&id&"'"
set rs08 = conn.execute(sql08)

response.redirect("pg_insert_itemCompra.asp")

'FECHA PEDIDO
elseif (acao = "fecharPedido") then

set rs08 = server.createObject("adodb.recordset")
sql08 = "UPDATE tb_compra SET staID = 8 WHERE comID = '"&id&"'"
set rs08 = conn.execute(sql08)

Dim rs09
Dim sql09
set rs09 = server.createObject("adodb.recordset")
sql09 = "UPDATE tb_itemcompra SET staID = 8 WHERE comID = '"&id&"'"
set rs09 = conn.execute(sql09)


	'ATUALIZA ESTOQUE DO PRODUTO COM OS NOVOS ITENS COMPRADOS
	Dim rs010
	Dim sql010
	
	Dim rs011
	Dim sql011
	
	Dim rs012
	Dim sql012
	
	
	set rs010 = Server.CreateObject("ADODB.Recordset")
	sql010 = "SELECT * FROM tb_itemcompra WHERE comID = '"&id&"'"
	set rs010 = conn.execute(sql010)
	
	Dim prodID
	Dim estoqueAtual
	Dim compra
	Dim estoqueFinal
	Dim proPrecoCusto1
	Dim proPrecoVenda1
	Dim proMargemLucro1
	Dim proPrecoVenda4
	Dim proMargemLucro4
	
	While Not rs010.EoF
	
	 prodID = rs010.fields.item("proID").value
	 compra = rs010.fields.item("iteQtde").value
	 proPrecoCusto1 = rs010.fields.item("itePreco").value
	 proPrecoVenda1 = rs010.fields.item("proPrecoVenda").value
	 proMargemLucro1 = rs010.fields.item("proMargemLucro").value
	 proPrecoVenda4 = rs010.fields.item("proPrecoVendaM").value
	 proMargemLucro4 = rs010.fields.item("proMargemLucroM").value
	 
	 'CONVERTENDO PARA FLOAT NO PADRÃO DO BANDO DE DADOS
	 proPrecoCusto1 = Replace(CDbl(proPrecoCusto1),",",".")
	 proPrecoVenda1 = Replace(CDbl(proPrecoVenda1),",",".")
	 proMargemLucro1 = Replace(CDbl(proMargemLucro1),",",".")
	 proPrecoVenda4 = Replace(CDbl(proPrecoVenda4),",",".")
	 proMargemLucro4 = Replace(CDbl(proMargemLucro4),",",".")
	 
	 set rs012 = Server.CreateObject("ADODB.Recordset")
	 sql012 = "SELECT * FROM tb_produto WHERE proID = '"&prodID&"'"
	 set rs012 = conn.execute(sql012)
	 
	 estoqueAtual = rs012.fields.item("proEstoque").value
	 
	 estoqueFinal = (estoqueAtual + compra)
	
	 set rs011 = Server.CreateObject("ADODB.Recordset")
	 sql011 = "UPDATE tb_produto SET proPrecoCusto = '"&proPrecoCusto1&"', proPrecoVenda = '"&proPrecoVenda1&"', proPrecoVendaM = '"&proPrecoVenda4&"', proMargemLucro = '"&proMargemLucro1&"', proMargemLucroM = '"&proMargemLucro4&"', proEstoque = '"&estoqueFinal&"' WHERE proID = '"&prodID&"' "
	 set rs011 = conn.execute(sql011)
	
	rs010.MoveNext
	Wend



response.redirect("pg_insert_itemCompra.asp")

end if
%>

<%
Dim rs07
Dim sql07

set rs07 = server.createObject("adodb.recordset")
sql07 = "SELECT tb_itemcompra.iteID, tb_itemcompra.proID, tb_itemcompra.iteQtde, tb_itemcompra.itePreco,tb_itemcompra.iteSubTotal, tb_produto.proDescricao, tb_produto.proUnidade, tb_produto.proCodFornecedor FROM tb_itemcompra INNER JOIN tb_produto ON tb_itemcompra.proID = tb_produto.proID WHERE tb_itemcompra.comID = '"&id&"'"
set rs07 = conn.execute(sql07)

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>SISTEM FORTE EM MÍDIA</title>
<link href="css/css1.css" rel="stylesheet" type="text/css" />

<script language="javascript" type="text/javascript">

function verForm(form1){


var tipID		= document.form1.tipID.value
var iteQtde		= document.form1.iteQtde.value;
var proID		= document.form1.proID.value;
var itePreco	= document.form1.itePreco.value;
var iteSubTotal = document.form1.iteSubTotal.value;

if (tipID == ""){
	alert("Favor selecionar o tipo de produto!");
	document.form1.tipID.focus();
	return false;
	}

if (iteQtde == ""){
	alert("Favor informar a quantidade de produto!");
	document.form1.iteQtde.focus();
	return false;
	}

if (proID == ""){
	alert("Favor selecionar o produto!");
	document.form1.proID.focus();
	return false;
	}	
if ((itePreco == "")||(itePreco == "0.00")){
	alert("Favor informar o valor do produto\nUtilize ponto para separar as casas decimais!");
	document.form1.itePreco.focus();
	return false;
	}	
if ((iteSubTotal == "")||(iteSubTotal == "0.00")){
	alert("Favor informar o valor do produto\nUtilize ponto para separar as casas decimais!");
	document.form1.itePreco.focus();
	return false;
	}	
	
}

function carregaProduto()
{
var tipDescr 	= document.form1.tipID.value;
tipDescr 		= document.getElementById('tipID');
tipDescr 		= tipDescr.options[tipDescr.selectedIndex].text;
tipID 			= document.getElementById('tipID').value;

var iteQtde = document.form1.iteQtde.value;
var itePreco = document.form1.itePreco.value;
var id = document.form1.id.value;
var forID = document.form1.forID.value;

window.location.href = "pg_insert_itemCompra.asp?id="+id+"&itePreco="+itePreco+"&iteQtde="+iteQtde+"&tipDescr="+tipDescr+"&tipID="+tipID+"&forID="+forID;

}

function carregaValor()
{
var proDescricao 	= document.form1.proID.value;
proDescricao 		= document.getElementById('proID');
proDescricao 		= proDescricao.options[proDescricao.selectedIndex].text;
proID 				= document.getElementById('proID').value;

var tipDescr 	= document.form1.tipID.value;
tipDescr 		= document.getElementById('tipID');
tipDescr 		= tipDescr.options[tipDescr.selectedIndex].text;
tipID 			= document.getElementById('tipID').value;

var iteQtde = document.form1.iteQtde.value;
var itePreco = document.form1.itePreco.value;
var id = document.form1.id.value;
var forID = document.form1.forID.value;

window.location.href = "pg_insert_itemCompra.asp?id="+id+"&itePreco="+itePreco+"&iteQtde="+iteQtde+"&tipDescr="+tipDescr+"&tipID="+tipID+"&forID="+forID+"&proDescricao="+proDescricao+"&proID="+proID;

}

function calcTotal()
{
var iteQtde =  document.form1.iteQtde.value;
var itePreco = document.form1.itePreco.value;
var iteSubTotal;

iteSubTotal = iteQtde * itePreco;

document.form1.iteSubTotal.value = iteSubTotal.toFixed(2);

}
function fecharPedido()
{
	if (confirm("Tem certeza que deseja fechar o pedido de compra?"))
	{
	window.location.href = "pg_insert_itemCompra.asp?acao=fecharPedido&id=<%=Request.QueryString("id")%>";
	}
}

function excluirPedido()
{
	if (confirm("Tem certeza que deseja excluir este pedido de compra?"))
	{
		window.location.href = "pg_insert_itemCompra.asp?acao=excluirPedido&id=<%=Request.QueryString("id")%>";
	}
}	

function verMargem(){
var p = document.form1.proMargemLucro.value;
var p1 = document.form1.proMargemLucroM.value;
var c = document.form1.itePreco.value;

p = p.replace(",",".");
p1 = p1.replace(",",".");
c = c.replace(",",".");

var v = parseFloat(p*c)/100 + parseFloat(c);
var v1 = parseFloat(p1*c)/100 + parseFloat(c);

v = parseFloat(v);
v1 = parseFloat(v1);

document.form1.proPrecoVenda.value = v.toFixed(2);
document.form1.proPrecoVendaM.value = v1.toFixed(2);
}	

function verMargemB(){
var c = document.form1.itePreco.value;
var i = document.form1.proPrecoVenda.value;

c = c.replace(",",".");
i = i.replace(",",".");

var v = parseFloat(i*100)/parseFloat(c) - 100;
v = parseFloat(v);
document.form1.proMargemLucro.value = v.toFixed(2);
}

function verMargemM(){
var c = document.form1.itePreco.value;
var i = document.form1.proPrecoVendaM.value;

c = c.replace(",",".");
i = i.replace(",",".");

var v = parseFloat(i*100)/parseFloat(c) - 100;
v = parseFloat(v);
document.form1.proMargemLucroM.value = v.toFixed(2);
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
        <li><a href="pg_insert_compra.asp">Comprar</a></li>           
	  </ul>
	</div>
	<div id="areaPrincipal">
    <div style="height:25px; line-height:25px; background:#ccc"></div>
    <table width="96%" border="0" align="left" cellpadding="3" cellspacing="3">
  <tr>
    <td width="50" align="center"><a href="pg_insert_compra.asp"><img src="ico/ico_calculadora.gif" alt="" width="60" height="60" border="0" class="icone" /></a></td>
    <td width="751" align="center" class="titulo">CADASTRO DE ITEM NO PEDIDO DE COMPRA</td>
    </tr>
  <tr>
    <td colspan="2" align="center">
    <form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?id=<%=Request.QueryString("id")%>" onsubmit="return verForm(this)">
      <table width="777" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td width="137" height="25" align="right">Nº do Pedido:</td>
          <td width="246" height="25" align="left"><input name="comID" type="text" id="comID" value=" <%=rs06.fields.item("comID").value%>" size="4" maxlength="4" readonly="readonly" />
            <input name="id" type="hidden" id="id" value="<%=request.QueryString("id")%>" />
            <input name="forID" type="hidden" id="forID" value="<%=rs06.fields.item("forID").value%>" /></td>
          <td width="112" height="25" align="right">Fornecedor:</td>
          <td width="269" height="25" align="left"><input name="forNome" type="text" id="forNome" value=" <%=rs06.fields.item("forNome").value%>" size="40" maxlength="50" readonly="readonly" /></td>
        </tr>
        <tr>
          <td height="25" align="right">Tipo de Produto:</td>
          <td height="25" align="left">
          
          <select name="tipID" id="tipID" onchange="carregaProduto()">
            <option value="<%=request.querystring("tipID")%>"><%=request.querystring("tipDescr")%></option>
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
            </select></td>
          <td height="25" align="right">Produto:</td>
          <td height="25" align="left">
          <select name="proID" id="proID" onChange="carregaValor()">
            <option value="<%=request.querystring("proID")%>" selected><%=request.querystring("proDescricao")%></option>
            <%
		  if (not rs03.eof) then
          do while not rs03.eof
		  %>
            <option value="<%=rs03.fields.item("proID").value%>"><%=rs03.fields.item("proDescricao").value%></option>
            <%
		  rs03.moveNext 
		  Loop
		  end if
		  %>
            </select></td>
        </tr>
        <tr>
          <td height="25" align="right">Qtde:</td>
          <td height="25" align="left"><input name="iteQtde" type="text" id="iteQtde" value="1.00" size="8" maxlength="8" onchange="calcTotal()" onkeydown="calcTotal()" onkeypress="calcTotal()" onblur="verMargem()"/></td>
          <td height="25" align="right">Preço de Custo:</td>
          <td height="25" align="left"><input name="itePreco" type="text" id="itePreco" value="<%=Replace(proPrecoCusto,",",".")%>" size="10" maxlength="10" onchange="calcTotal()" onkeydown="calcTotal()" onkeypress="calcTotal()" onblur="verMargem()" />
            &nbsp;&nbsp;Total: 
            <input name="iteSubTotal" type="text" id="iteSubTotal" value="<%=Replace(FormatNumber(proPrecoCusto),",",".")%>" size="10" maxlength="10" readonly="readonly" /></td>
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
          <td height="25" align="right" bgcolor="#99CC99">Preço Telefone/balcao:</td>
          <td height="25" align="left" bgcolor="#99CC99"><input name="proPrecoVenda" type="text" id="proPrecoVenda" value="<%=Replace(FormatNumber(proPrecoVenda),",",".")%>" size="10" maxlength="10" onkeypress="verMargemB()" onblur="verPonto(); verMargemB()"/></td>
          <td height="25" align="right" bgcolor="#66FF99">Preço Mesa:</td>
          <td height="25" align="left" bgcolor="#66FF99"><input type="text" name="proPrecoVendaM" value="<%=Replace(FormatNumber(proPrecoVendaM),",",".")%>" size="10" maxlength="10" onkeypress="verMargemM()" onblur="verPonto(); verMargemM()"/>
            *</td>
          </tr>
        <tr>
          <td height="25" align="right" bgcolor="#99CC99">Lucro:</td>
          <td height="25" align="left" bgcolor="#99CC99"><input name="proMargemLucro" type="text" id="proMargemLucro" value="<%=Replace(FormatNumber(proMargemLucro),",",".")%>" size="6" maxlength="6" onblur="verPonto()" readonly="readonly"/>
            %</td>
          <td height="25" align="right" bgcolor="#66FF99">Lucro:</td>
          <td height="25" bgcolor="#66FF99"><input name="proMargemLucroM" type="text" id="proMargemLucroM" onblur="verPonto()" value="<%=Replace(FormatNumber(proMargemLucroM),",",".")%>" size="6" maxlength="6" readonly="readonly"/>
            %</td>
          </tr>
        <tr>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="left">&nbsp;</td>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="left">&nbsp;</td>
        </tr>
        <tr>
          <td height="165" colspan="4" align="center" valign="top">
          <table width="740" border="0" cellspacing="2" cellpadding="2">
 <tr class="textoBranco">
    <td width="30" height="20" align="center" valign="middle" bgcolor="#990000">ID</td>
    <td width="110" height="20" valign="middle" bgcolor="#990000">&nbsp;Cód. Fornecedor</td>
    <td width="300" height="20" valign="middle" bgcolor="#990000">&nbsp;Descrição do Produto</td>
    <td width="30" height="20" align="center" valign="middle" bgcolor="#990000">UN</td>
    <td width="30" height="20" align="center" valign="middle" bgcolor="#990000">Qtde</td>
    <td width="80" height="20" valign="middle" bgcolor="#990000">&nbsp;Preço</td>
    <td width="115" height="20" valign="middle" bgcolor="#990000">&nbsp;Total</td>
    <td width="19" height="20" align="center" valign="middle" bgcolor="#990000">Ex</td>
</tr>
<%
if not rs07.eof then
do while not rs07.eof
%>
 <tr class="textoComum">
   <td align="left" valign="middle" bgcolor="#FFFFFF"><%=rs07.fields.item("proID").value%></td>
   <td align="left" valign="middle" bgcolor="#FFFFFF"><%=rs07.fields.item("proCodFornecedor").value%></td>
   <td align="left" valign="middle" bgcolor="#FFFFFF"><%=rs07.fields.item("proDescricao").value%></td>
   <td align="center" valign="middle" bgcolor="#FFFFFF"><%=rs07.fields.item("proUnidade").value%></td>
   <td align="left" valign="middle" bgcolor="#FFFFFF"><%=rs07.fields.item("iteQtde").value%></td>
   <td align="left" valign="middle" bgcolor="#FFFFFF"><%=rs07.fields.item("itePreco").value%></td>
   <td align="left" valign="middle" bgcolor="#FFFFFF"><%=Replace(rs07.fields.item("iteSubTotal").value,",",".")%></td>
   <td align="left" valign="middle" bgcolor="#FFFFFF"><a href="pg_insert_itemcompra.asp?id=<%=request.querystring("id")%>&acao=excluirItem&idItem=<%=rs07.fields.item("iteID").value%>"><img src="ico/ico_excluir.gif" width="15" height="15" border="0"/></a></td>
 </tr>
 <%
 rs07.moveNext
 loop
 end if
 %>
</table>
          </td>
        </tr>
        <tr>
          <td height="25" align="right">Valor Total:</td>
          <td height="25" align="left"><input name="iteTotal" type="text" id="iteTotal" value="<%=rs05.fields.item("total").value%>" size="10" maxlength="10" readonly="readonly" /></td>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="center"><input name="lancar" type="submit" class="botao" id="lancar" value="Lançar" />
            <input name="fechar" type="button" class="botao" id="fechar" value="Fechar" onClick="fecharPedido()"/>
            <input name="excluir" type="button" class="botao" id="excluir" value="Excluir" onClick="excluirPedido()"/></td>
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