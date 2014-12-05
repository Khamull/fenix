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

<div id="conteudo" style="width:500px; font-family:'Courier New', Courier, monospace; border:1px solid #666; background-color:#FFF;" >

<table cellpadding="2" cellspacing="2" border="0" align="center" width="490">
 <tr>
 <td colspan="4">

<table width="300">
 <tr>
  <td width="83"><a href="pg_menu_pedidos.asp">FECHAR</a></td>
  <td width="176">
   <% if (Not rs03.EoF) Then %>
    <strong> <big> MESA: <%=rs03.fields.item("mesNumero").value%> </big> </strong>
   <% end if %>
  </td>
  <td width="25"><input type="image" src="ico/ico_printer.gif" border="0" title="Imprimir" onclick="window.print();" /></td>
 </tr>
</table>

</td>
</tr>

<tr>
<td colspan="4"><strong>PEDIDO</strong></td>
</tr>
<tr>
<td height="10" colspan="4"></td>
</tr>
<tr>
<td>DATA</td>
<td colspan="3"><%=rs01.fields.item("venData").value%></td>
</tr>
<tr>
<td>NUMERO</td>
<td colspan="3"><%=rs00.fields.item("numerovenda").value%></td>
</tr>
<tr>
  <td>CODIGO PEDIDO</td>
  <td colspan="3"><%=venID%></td>
</tr>
<tr>
<td>TIPO PEDIDO</td>
<td colspan="3"><%=rs01.fields.item("tipVendaDescricao").value%></td>
</tr>
<tr>
<td>ATENDENTE</td>
<td colspan="3"><%=rs01.fields.item("usuLogin").value%></td>
</tr>
<tr>
<td>COD CLIENTE</td>
<td colspan="3"><%=rs01.fields.item("cliID").value%></td>
</tr>
<tr>
<td>TELEFONE</td>
<td colspan="3"><%=rs01.fields.item("cliTelefone").value%></td>
</tr>
<tr>
<td>CLIENTE</td>
<td colspan="3"></td>
</tr>
<tr>
<td colspan="4"><%=rs01.fields.item("cliNome").value%></td>
</tr>
<tr>
<td height="10" colspan="4"></td>
</tr>
<tr>
<td colspan="4">ENDERECO DE ENTREGA</td>
</tr>
<tr>
<td colspan="4"><%=rs01.fields.item("cliEndereco").value%></td>
</tr>
<tr>
<td colspan="4"><%=rs01.fields.item("baiNome").value%></td>
</tr>
<tr>
<td colspan="4"><%=rs01.fields.item("cidNome").value%></td>
</tr>
<tr>
<td height="10" colspan="4"></td>
</tr>
<tr>
<td height="10" colspan="4"></td>
</tr>
<tr>
<td colspan="4">ITENS DO PEDIDO</td>
</tr>
<tr>
<td width="296">DESCRICAO</td>
<td width="59">PRECO</td>
<td width="37">QTD</td>
<td width="72">TOTAL</td>
</tr>
<%While Not rs02.EoF%>
<tr>
<td height="25"><%=rs02.fields.item("proDescricao").value%></td>
<td height="25"><%=FormatNumber(rs02.fields.item("itePreco").value)%></td>
<td height="25"><%=rs02.fields.item("iteqtde").value%></td>
<td height="25"><%=FormatNumber(rs02.fields.item("iteSubTotal").value)%></td>
</tr>
<%
  rs02.MoveNext
	Wend
%>
<tr>
<td height="10" colspan="4"></td>
</tr>
<tr>
<td colspan="2">OBSERVACOES</td>
<td colspan="2"></td>
</tr>
<tr>
<td colspan="4"><%=rs01.fields.item("venObs").value%></td>
</tr>
<tr>
<td height="10" colspan="4"></td>
</tr>
<tr>
<td>SUB TOTAL</td>
<td colspan="3"><%=FormatNumber(rs010.fields.Item("subTotal").value)%></td>
</tr>
<tr>
<td>DESCONTO</td>
<td colspan="3"><%=FormatNumber(rs01.fields.item("venValorD").value)%></td>
</tr>
<%
Dim acrescimo
acrescimo = 0
%>
<% if (Not rs03.EoF AND rs01.fields.item("statusVenda").value <> "10") Then %>
<%'Calcula Acrescimo
acrescimo = (rs010.fields.Item("subTotal").value/10)
%>
<tr>
<td>ACRESCIMO</td>
<td colspan="3"><%=FormatNumber(acrescimo)%></td>
</tr>
<%else%>
<%
acrescimo = rs01.fields.item("venValorA").value
%>
<tr>
<td>ACRESCIMO</td>
<td colspan="3"><%=FormatNumber(acrescimo)%></td>
</tr>
<%end if%>
<tr>
<td>VALOR FRETE</td>
<td colspan="3"><%=FormatNumber(rs01.fields.item("baiFrete").value)%></td>
</tr>
<tr>
<td>TROCO</td>
<td colspan="3"><%=FormatNumber(rs01.fields.item("venValorTc").value)%></td>
</tr>
<tr>
<td height="10" colspan="4"></td>
</tr>
<%'Calcula Total da Venda
Dim total
total = (acrescimo + rs01.fields.item("baiFrete").value + rs010.fields.Item("subTotal").value - rs01.fields.item("venValorD").value)
%>
<tr>
  <td><strong>TOTAL</strong></td>
  <td colspan="3"><strong><%=FormatNumber(total)%></strong></td>
</tr>
<tr>
  <td height="10" colspan="4"></td>
</tr>
<tr>
<td height="10" colspan="4"></td>
</tr>
<tr>
<td height="10" colspan="4" align="center">RESTAURANTE TAUNAY</td>
</tr>
<tr>
<td height="10" colspan="4" align="center">Rua Visconde Taunay, 433</td>
</tr>
<tr>
<td height="10" colspan="4" align="center">Vila Arens - Jundiai - SP</td>
</tr>
<tr>
  <td height="10" colspan="4" align="center">Fone: 11 4587 - 5436 </td>
</tr>
<tr>
  <td height="10" colspan="4" align="center">www.restaurantetaunay.com.br</td>
</tr>
<tr>
<td height="100" colspan="4"></td>
</tr>
<tr>
<td colspan="4">_</td>
</tr>
</table>
</div>

</body>
</html>

<%call fechaConexao %>