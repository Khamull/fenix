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
Dim venID
dim idItem
venID = Request.QueryString("venID")
idItem = Request.QueryString("idItem")
acao = Request.QueryString("acao")

if (venID = "") then
Response.redirect("pg_menu_pedidos.asp")
elseif (venID <> "") then


	'------ GERA O NUMERO DA VENDA --------
	Dim rs00
	Dim sql00
	Dim numero
	
	set rs00 = Server.CreateObject("ADODB.Recordset")
	sql00 = "SELECT numerovenda FROM tb_numerovenda WHERE venID = '"&venID&"'"
	set rs00 = conn.execute(sql00)

		numero = rs00.fields.item("numerovenda").value

	'--------------------------------------

	Dim rs06
	Dim sql06
	set rs06 = server.CreateObject("adodb.recordset")
		sql06 = "SELECT tb_venda.venID, tb_venda.staID, tb_cliente.cliNome, tb_cliente.cliID FROM tb_venda INNER JOIN tb_cliente ON tb_cliente.cliID = tb_venda.cliID WHERE tb_venda.venID = '"&venID&"'"
	set rs06 = conn.Execute(sql06)

end if
%>

<%
Dim proID
Dim iteQtde
Dim itePreco
Dim iteSubTotal
Dim iteObs

proID	 		=	Request.Form("proID")
iteQtde			=	Request.Form("iteQtde")
itePreco		=	Request.Form("itePreco")
iteSubTotal		=	Request.Form("iteSubTotal")
iteObs			=	UCASE(Request.Form("iteObs"))

%>

<%
if (not isEmpty(Request.Form("lancar"))) then
	
	'Grava
		Dim rs01
		Dim sql01
		set rs01 = server.CreateObject("adodb.recordset")
			sql01 = "INSERT INTO tb_itemvenda (proID,iteQtde,itePreco,iteSubTotal,venID, iteObs)"
			sql01 = sql01&" VALUES ('"&proID&"','"&iteQtde&"','"&itePreco&"','"&iteSubTotal&"','"&venID&"','"&iteObs&"')"
		set rs01 = conn.execute(sql01)
		set rs01 = nothing

	Response.redirect("pg_insert_itemVendaMesa1.asp?venID="&Request.QueryString("venID"))
	
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
	sql03 = "SELECT * FROM tb_produto WHERE tipID = '"&Request.QueryString("tipID")&"' ORDER BY proDescricao"
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
Dim proPrecoVenda
proPrecoCusto = rs04.fields.item("proPrecoCusto").value
proPrecoVenda = rs04.fields.item("proPrecoVendaM").value
else
proPrecoCusto = "0.00"
proPrecoVenda = "0.00"
end if
%>

<%
Dim rs05
Dim sql05
set rs05 = server.CreateObject("adodb.recordset")
	sql05 = "SELECT SUM(iteSubTotal) as total FROM tb_itemvenda WHERE venID = '"&Request.QueryString("venID")&"'"
set rs05 = conn.execute(sql05)	
%>

<%

if (acao = "excluirItem") then

Dim rs08
Dim sql08
set rs08 = server.createObject("adodb.recordset")
sql08 = "DELETE FROM tb_itemvenda WHERE iteID = '"&idItem&"'"
set rs08 = conn.execute(sql08)

elseif (acao = "excluirPedido") then


 'Registra a Desistencia
  Dim rsx1
  Dim sqlx1
  Dim rsy1
  Dim sqly1
  Dim caixaID
  
  set rsx1 = Server.CreateObject("ADODB.Recordset")
  sqlx1 = "SELECT * FROM tb_caixa WHERE status = 'A'"
  set rsx1 = conn.execute(sqlx1)
  
   caixaID = rsx1.fields.item("caixaID").value
  
  set rsy1 = Server.CreateObject("ADODB.Recordset")
  sqly1 = "INSERT INTO tb_cancelados (caixaID, mesa) VALUES ('"&caixaID&"', '1')"
  set rsy1 = conn.execute(sqly1)
 '----------------------


set rs08 = server.createObject("adodb.recordset")
sql08 = "DELETE FROM tb_venda WHERE venID = '"&venID&"'"
set rs08 = conn.execute(sql08)

response.redirect("pg_select_pedidos_mesa.asp")

elseif (acao = "fecharPedido") then

set rs08 = server.createObject("adodb.recordset")
sql08 = "UPDATE tb_venda SET staID = 8 WHERE venID = '"&venID&"'"
set rs08 = conn.execute(sql08)

Dim rs09
Dim sql09
set rs09 = server.createObject("adodb.recordset")
sql09 = "UPDATE tb_itemvenda SET staID = 8 WHERE venID = '"&venID&"'"
set rs09 = conn.execute(sql09)

response.redirect("pg_select_pedidos_mesa.asp")

end if
%>

<%
Dim rs07
Dim sql07

set rs07 = server.createObject("adodb.recordset")
sql07 = "SELECT tb_itemvenda.iteID, tb_itemvenda.proID, tb_itemvenda.iteObs, tb_itemvenda.iteQtde, tb_itemvenda.itePreco,tb_itemvenda.iteSubTotal, tb_produto.proDescricao, tb_produto.proUnidade, tb_produto.proCodFornecedor FROM tb_itemvenda INNER JOIN tb_produto ON tb_itemvenda.proID = tb_produto.proID WHERE tb_itemvenda.venID = '"&venID&"'"
set rs07 = conn.execute(sql07)

%>





<%if Request.QueryString("venID") <> "" Then%>

<%
Dim rs044
Dim sql044
Dim quantidadeCompra
Dim pID

pID = Request.QueryString("proID")

set rs044 = Server.CreateObject("ADODB.Recordset")
sql044 = "SELECT SUM(tb_itemvenda.iteQtde) as soma, tb_itemvenda.iteID, tb_itemvenda.venID, tb_venda.venID, tb_venda.staID FROM tb_itemvenda INNER JOIN tb_venda ON tb_venda.venID = tb_itemvenda.venID WHERE tb_itemvenda.proID = '"&pID&"' AND tb_venda.staID = '1'"
set rs044 = conn.execute(sql044)

if (Not rs044.EoF) Then
 
		 if(rs044.fields.item("soma").value > 0)Then
		  quantidadeCompra = rs044.fields.item("soma").value
		 else
		  quantidadeCompra = 0
		 end if
 
else
quantidadeCompra = CInt("0")
end if

%>

<%
Dim estoque
Dim estoqueVisualizar
Dim estoqueMinimo
Dim msgEstoque
Dim x

'VERIFICA QUANTIDADE EM ESTOQUE
if(not rs04.EoF)Then
estoque = rs04.fields.item("proEstoque").value
estoqueVisualizar = (estoque - quantidadeCompra)
estoqueMinimo = rs04.fields.item("proEstoqueMin").value

else
estoque = 0
estoqueMinimo = 0
estoqueVisualizar = 0

end if

'VERIFICA SE ESTÁ ABAIXO DO ESTOQUE MÍNIMO
if(estoqueVisualizar < estoqueMinimo)Then
 if(rs04.fields.item("tipID").value <> "1") Then
	msgEstoque = "O estoque está abaixo do mínimo"
 end if

x = "1"
end if

%>


 <% end if %>



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
var venID = document.form1.venID.value;
var cliID = document.form1.cliID.value;

window.location.href = "pg_insert_itemVendaMesa1.asp?venID="+venID+"&itePreco="+itePreco+"&iteQtde="+iteQtde+"&tipDescr="+tipDescr+"&tipID="+tipID+"&cliID="+cliID;

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
var venID = document.form1.venID.value;
var cliID = document.form1.cliID.value;

window.location.href = "pg_insert_itemVendaMesa1.asp?venID="+venID+"&itePreco="+itePreco+"&iteQtde="+iteQtde+"&tipDescr="+tipDescr+"&tipID="+tipID+"&cliID="+cliID+"&proDescricao="+proDescricao+"&proID="+proID;

}

function calcTotal()
{
var iteQtde =  document.form1.iteQtde.value;
var itePreco = document.form1.itePreco.value;
var itePrecoCompara = document.form1.itePrecoCompara.value;
var iteSubTotal;

if(itePrecoCompara <= itePreco){

iteSubTotal = iteQtde * itePreco;

document.form1.iteSubTotal.value = iteSubTotal;

}else{
	alert("O produto não pode ser vendido \n por um valor mais baixo!");
	document.form1.itePreco.value = document.form1.itePrecoCompara.value;
	document.form1.itePreco.focus();
	return null;
	}
}

function fecharPedido()
{
	if (confirm("Tem certeza que deseja fechar a venda?"))
	{
	window.location.href = "pg_update_fechar_pedido_telefone.asp?venID=<%=Request.QueryString("venID")%>";
	}
}

function excluirPedido(staID)
{
	if (confirm("Tem certeza que deseja excluir este pedido?"))
	{
		if(staID != "10"){
			window.location.href = "pg_insert_itemVendaMesa1.asp?acao=excluirPedido&venID=<%=Request.QueryString("venID")%>";
		}else{
		window.open("pg_excluir_pedidos.asp?venID=<%=Request.QueryString("venID")%>&tipo=mesa", "Excluir", "width=300px, height=260px");
		}
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
	  </ul>
	</div>
	<div id="areaPrincipal">
    <div style="height:25px; line-height:25px; background:#ccc"></div>
    <table width="96%" border="0" align="left" cellpadding="3" cellspacing="3">
  <tr>
    <td width="50" align="center"><img src="ico/ico_mesa.gif" alt="" width="60" height="60" border="0" class="icone" /></td>
    <td width="751" align="center" class="titulo">CADASTRO ITENS - VENDA</td>
    </tr>
  <tr>
    <td colspan="2" align="center" valign="top">
    <form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?venID=<%=Request.QueryString("venID")%>" onsubmit="return verForm(this)">
      <table width="777" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td width="109" height="25" align="right">Nº do Pedido:</td>
          <td width="267" height="25" align="left">
           <input name="numerovenda" type="text" id="numerovenda" value="<%=numero%>" size="4" maxlength="4" readonly="readonly" />
           <input name="venID" type="hidden" id="venID" value=" <%=rs06.fields.item("venID").value%>" size="4" maxlength="4" readonly="readonly" />
           <input name="cliID" type="hidden" id="cliID" value="<%=rs06.fields.item("cliID").value%>" /></td>
          <td width="110" height="25" align="right">Cliente:</td>
          <td width="263" height="25" align="left"><input name="cliNome" type="text" id="cliNome" value=" <%=rs06.fields.item("cliNome").value%>" size="40" maxlength="50" readonly="readonly" /></td>
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
            <option value="<%=rs03.fields.item("proID").value%>">
			<%=rs03.fields.item("proCodEmpresa").value%> - <%=rs03.fields.item("proDescricao").value%>
            </option>
            <%
		  rs03.moveNext 
		  Loop
		  end if
		  %>
            </select></td>
        </tr>
        <tr>
          <td height="25" align="right">Qtde:</td>
          <td height="25" align="left"><input name="iteQtde" type="text" id="iteQtde" value="1.00" size="8" maxlength="8" onchange="calcTotal()" onkeydown="calcTotal()" onkeypress="calcTotal()"/></td>
          <td height="25" align="right">Preço:</td>
          <td height="25" align="left"><input name="itePreco" type="text" id="itePreco" onblur="calcTotal()" onchange="calcTotal()" value="<%=Replace(proPrecoVenda,",",".")%>" size="10" maxlength="10" />
&nbsp;&nbsp;Total:
<input name="iteSubTotal" type="text" id="iteSubTotal" value="<%=Replace(FormatNumber(proPrecoVenda),",",".")%>" size="10" maxlength="10" readonly="readonly" />
    
	<input name="itePrecoCompara" type="hidden" id="itePrecoCompara" onchange="calcTotal()" onkeypress="calcTotal()" onkeydown="calcTotal()" value="<%=Replace(proPrecoVenda,",",".")%>" size="10" maxlength="10" readonly="readonly" />
	 </td>
        </tr>
        <%
		 if Not rs04.EoF Then
		  if(rs04.fields.item("tipID").value <> "1") Then
		%>
        <tr>
          <td height="25" align="right">Qtde em Estoque</td>
          <%if (x = "1") Then%>
           <td height="25" align="left"><input style="color:#F00; background-color:#CCC" type="text" value="<%=estoqueVisualizar%>" size="8" maxlength="8"  readonly="readonly"/></td>
          <%else%>
           <td height="25" align="left"><input type="text" value="<%=estoqueVisualizar%>" size="8" maxlength="8"  readonly="readonly"/></td>
          <%end if%>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="left">&nbsp;</td>
        </tr>
        <%
		  end if
		 end if
		%>
        <tr>
          <td height="25" align="right">Obs:</td>
          <td height="25" align="left"><input name="iteObs" type="text" id="iteObs" size="40" maxlength="40" onchange="calcTotal()" onkeydown="calcTotal()" onkeypress="calcTotal()"/></td>
          <td height="25" align="right">&nbsp;</td>
          <td height="25" align="left"><font color="#FF0000"><b><%=msgEstoque%></b></font></td>
        </tr>
        <tr>
          <td height="112" colspan="4" align="center" valign="top">
          <table width="740" border="0" cellspacing="2" cellpadding="2">
 <tr class="textoBranco">
    <td width="50" height="20" align="center" valign="middle" bgcolor="#990000">COD.</td>
    <td width="404" height="20" valign="middle" bgcolor="#990000">&nbsp;Descrição do Produto - Obs:</td>
    <td width="30" height="20" align="center" valign="middle" bgcolor="#990000">UN</td>
    <td width="30" height="20" align="center" valign="middle" bgcolor="#990000">Qtde</td>
    <td width="80" height="20" valign="middle" bgcolor="#990000">&nbsp;Preço</td>
    <td width="80" height="20" valign="middle" bgcolor="#990000">&nbsp;Total</td>
    <td width="22" height="20" align="center" valign="middle" bgcolor="#990000">Ex</td>
</tr>
<%
if not rs07.eof then
do while not rs07.eof
%>
 <tr class="textoComum">
   <td align="left" valign="middle" bgcolor="#FFFFFF"><%=rs07.fields.item("proCodFornecedor").value%></td>
   <td align="left" valign="middle" bgcolor="#FFFFFF"><%=rs07.fields.item("proDescricao").value%> - <%=rs07.fields.item("iteObs").value%></td>
   <td align="center" valign="middle" bgcolor="#FFFFFF"><%=rs07.fields.item("proUnidade").value%></td>
   <td align="left" valign="middle" bgcolor="#FFFFFF"><%=rs07.fields.item("iteQtde").value%></td>
   <td align="left" valign="middle" bgcolor="#FFFFFF"><%=rs07.fields.item("itePreco").value%></td>
   <td align="left" valign="middle" bgcolor="#FFFFFF"><%=rs07.fields.item("iteSubTotal").value%></td>
   <td align="left" valign="middle" bgcolor="#FFFFFF"><a href="pg_insert_itemVendaMesa1.asp?venID=<%=request.querystring("venID")%>&acao=excluirItem&idItem=<%=rs07.fields.item("iteID").value%>"><img src="ico/ico_excluir.gif" width="15" height="15" border="0"/></a></td>
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
            <input name="excluir" type="button" class="botao" id="excluir" value="Excluir" onClick="excluirPedido(<%=rs06.fields.item("staID").value%>)"/></td>
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