<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include file="inc/inc_conexao.inc"-->

<%call abreConexao()%>

<%
Dim venID
venID = request.querystring("venID")

if venID = "" then
response.redirect("pg_select_pedidos_telefone.asp")
end if
%>

<%'Seleciona Numero da Venda
Dim rs00
Dim sql00

set rs00 = Server.CreateObject("ADODB.Recordset")
sql00 = "SELECT * FROM tb_numerovenda WHERE venID = '"&venID&"'"
set rs00 = conn.execute(sql00)
%>

<%
Dim rs01
Dim sql01
set rs01 = server.CreateObject("adodb.recordset") 
sql01 = "SELECT tb_venda.venValorT, tb_venda.venValorS, tb_venda.venValorF, tb_venda.venValorR, tb_venda.venValorTc, tb_venda.venValorA, tb_venda.venValorD,tb_venda.venLocalidade, tb_venda.venObs, tb_venda.cliID, tb_venda.tipVendaID, tb_venda.usuLogin, tb_venda.venData, tb_venda.staID as statusVenda, tb_venda.venID, tb_cliente.cliEndereco, tb_cliente.baiID, tb_cliente.cidID, tb_cliente.cliTelefone, tb_cliente.cliNome, tb_cliente.cliID, tb_bairro.baiID, tb_bairro.baiNome, tb_bairro.baiFrete, tb_cidade.cidID, tb_cidade.cidNome, tb_tipovenda.tipVendaID, tb_tipovenda.tipVendaDescricao FROM tb_venda INNER JOIN tb_cliente ON tb_cliente.cliID = tb_venda.cliID INNER JOIN tb_tipovenda ON tb_tipovenda.tipVendaID = tb_venda.tipVendaID INNER JOIN tb_bairro ON tb_bairro.baiID = tb_cliente.baiID INNER JOIN tb_cidade ON tb_cidade.cidID = tb_cliente.cidID WHERE tb_venda.venID = '"&venID&"'"
set rs01 = conn.execute(sql01)
%>

<%
Dim rs02
Dim sql02
set rs02 = server.CreateObject("adodb.recordset") 
sql02 = "SELECT tb_itemvenda.iteID, tb_itemvenda.proID, tb_itemvenda.iteObs, tb_itemvenda.iteQtde, tb_itemvenda.itePreco,tb_itemvenda.iteSubTotal, tb_produto.proDescricao, tb_produto.proUnidade, tb_produto.proCodFornecedor FROM tb_itemvenda INNER JOIN tb_produto ON tb_itemvenda.proID = tb_produto.proID WHERE tb_itemvenda.venID = '"&venID&"'"
set rs02 = conn.execute(sql02)
%>

<%
Dim rs03
Dim sql03
set rs03 = Server.CreateObject("ADODB.Recordset")
sql03 = "SELECT tb_mesa.mesID, tb_mesa.mesNumero, tb_venda.venID, tb_venda.mesID, tb_venda.tipVendaID FROM tb_venda INNER JOIN tb_mesa ON tb_mesa.mesID = tb_venda.mesID WHERE tb_venda.venID = '"&venID&"'"
set rs03 = conn.execute(sql03)
%>

<%
Dim rs04
Dim sql04
Dim venHoraF

venHoraF = time()

set rs04 = Server.CreateObject("ADODB.Recordset")
sql04 = "UPDATE  tb_venda SET venHoraF = '"&venHoraF&"' WHERE venID = '"&venID&"';"
set rs04 = conn.execute(sql04)
%>


<%'Seleciona Subtotal da Venda
Dim rs010
Dim sql010
set rs010 = server.CreateObject("adodb.recordset")
sql010 = "SELECT SUM(iteSubTotal) AS subTotal FROM tb_itemvenda WHERE venID='"&venID&"'"
set rs010 = conn.execute(sql010)
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Pedido</title>
</head>

<body>

<table width="200">
 <tr>
  <td width="69"><a href="pg_menu_pedidos.asp">FECHAR</a></td>
  <td width="290">
   <% if (Not rs03.EoF) Then %>
    <strong> <big> MESA: <%=rs03.fields.item("mesNumero").value%> </big> </strong>
   <% end if %>
  </td>
  <td width="25"><input type="image" src="ico/ico_printer.gif" border="0" title="Imprimir" onclick="window.print();" /></td>
 </tr>
</table>

<br />

<strong>PEDIDO</strong>
<br />
<br />
DATA..........................................<%=rs01.fields.item("venData").value%><br />
NUMERO....................................<%=rs00.fields.item("numerovenda").value%><br />
TIPO PEDIDO.............................<%=rs01.fields.item("tipVendaDescricao").value%><br />
ATENDENTE..............................<%=rs01.fields.item("usuLogin").value%><br />
COD CLIENTE...........................<%=rs01.fields.item("cliID").value%><br />
TELEFONE.................................<%=rs01.fields.item("cliTelefone").value%><br />
CLIENTE....................................<%=rs01.fields.item("cliNome").value%><br />
<br />
ENDERECO DE ENTREGA<br />
<%=rs01.fields.item("cliEndereco").value%><br />
<%=rs01.fields.item("baiNome").value%><br />
<%=rs01.fields.item("cidNome").value%><br />


<br />
<br />
ITENS DO PEDIDO<br />
<br />
DESCRICAO / PRECO / QTD / TOTAL <br />
<br />
<br />

<%While Not rs02.EoF%>
<%=rs02.fields.item("proDescricao").value%> / <%=Replace(FormatNumber(rs02.fields.item("itePreco").value),",",".")%> / <%=Replace(FormatNumber(rs02.fields.item("iteqtde").value),",",".")%> / <%=Replace(FormatNumber(rs02.fields.item("iteSubTotal").value),",",".")%><br /><br />
<%
  rs02.MoveNext
	Wend
%>
<br />
OBSERVACOES<br />
<%=rs01.fields.item("venObs").value%><br />
<br />
<br />
SUB TOTAL.................................<%=Replace(FormatNumber(rs010.fields.Item("subTotal").value),",",".")%><br />
DESCONTO.................................<%=Replace(FormatNumber(rs01.fields.item("venValorD").value),",",".")%><br />

<%
Dim acrescimo
acrescimo = 0
%>

<% if (Not rs03.EoF AND rs01.fields.item("statusVenda").value <> "10") Then %>

<%'Calcula Acrescimo
acrescimo = (rs010.fields.Item("subTotal").value/10)
%>
ACRESCIMO...............................<%=Replace(FormatNumber(acrescimo),",",".")%><br />
<%else%>
<%
acrescimo = rs01.fields.item("venValorA").value
%>
ACRESCIMO...............................<%=Replace(FormatNumber(acrescimo),",",".")%><br />
<%end if%>
VALOR FRETE............................<%=Replace(FormatNumber(rs01.fields.item("baiFrete").value),",",".")%><br />
TROCO.........................................<%=Replace(FormatNumber(rs01.fields.item("venValorTc").value),",",".")%><br />
<br />
<%'Calcula Total da Venda
Dim total
total = (acrescimo + rs01.fields.item("baiFrete").value + rs010.fields.Item("subTotal").value - rs01.fields.item("venValorD").value)
%>
<strong>TOTAL.........................................<%=Replace(FormatNumber(total),",",".")%></strong><br />
<br />
<br />
<br />
<font face="Arial, Helvetica, sans-serif" size="1"><b>NOME DO RESTAURANTE - Rua Endere&ccedil;o, 1<br />
Bairro - Cidade -- SP -

      CEP: <br />
Telefone: (11) <br />
Site: www.site.com.br</b></font><br />



</body>
</html>

<%call fechaConexao %>